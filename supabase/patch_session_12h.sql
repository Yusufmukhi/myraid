-- =============================================================================
-- patch_session_12h.sql: Shorten Active Session Expiry to 12 Hours
-- Safe to run on live data.
-- =============================================================================

begin;

-- 1. Update app_sessions table column default to 12 hours
alter table public.app_sessions 
  alter column expires_at set default (now() + interval '12 hours');

-- 2. Cap any existing active sessions so they do not exceed 12 hours from now
update public.app_sessions 
set expires_at = least(expires_at, now() + interval '12 hours')
where expires_at > now() + interval '12 hours';

-- 3. Clean up any expired sessions
delete from public.app_sessions 
where expires_at <= now();

-- 4. Update app_login function to issue 12-hour session tokens
create or replace function public.app_login(p_username text, p_password text)
returns json language plpgsql security definer set search_path = public, extensions as $$
declare
  u public.app_users;
  t uuid;
  user_org_id uuid;
  hdrs jsonb;
  client_ip text := '';
  login_fail_key text;
  uname text := btrim(p_username);
  t_hash text;
begin
  -- 1. Extract IP safely
  begin
    hdrs := current_setting('request.headers', true)::jsonb;
    if hdrs is not null then
      client_ip := coalesce(hdrs->>'cf-connecting-ip', hdrs->>'x-real-ip', split_part(hdrs->>'x-forwarded-for', ',', 1), '');
    end if;
  exception when others then client_ip := '';
  end;

  if client_ip is null or client_ip = '' or client_ip = 'local' then
    client_ip := 'conn_' || substr(encode(extensions.gen_random_bytes(4), 'hex'), 1, 8);
  end if;

  -- 2. Fail-Closed Rate Limiting: IP & IP-Account Combination
  login_fail_key := 'login_combo:' || client_ip || ':' || lower(uname);
  begin
    if not public.check_rate_limit('login_ip:' || client_ip, 20, 300) then
      raise exception 'Too many login attempts from this network. Please wait 5 minutes before trying again.' using errcode='P0001';
    end if;

    if not public.check_rate_limit(login_fail_key, 5, 300) then
      raise exception 'Too many failed login attempts from your connection. Please wait 5 minutes before trying again.' using errcode='P0001';
    end if;
  exception when others then
    if sqlstate = 'P0001' then raise; end if;
    raise exception 'Security validation failed. Please try again later.' using errcode='P0001';
  end;

  -- 3. Lookup user by username OR email
  select * into u from public.app_users
  where lower(username) = lower(uname)
     or (email is not null and lower(email) = lower(uname));

  -- Clean up expired sessions
  delete from public.app_sessions where expires_at <= now();

  -- 4. Verify password
  if u.id is null or u.password_hash <> extensions.crypt(p_password, u.password_hash) then
    if u.id is not null then
      update public.app_users
      set failed_attempts = failed_attempts + 1
      where id = u.id;

      select org_id into user_org_id from public.org_members where user_id = u.id limit 1;
      if user_org_id is not null then
        perform public._log_activity(
          user_org_id, null, u.id, 'login_failed',
          json_build_object('username', u.username, 'ip', client_ip)
        );
      end if;
    end if;
    raise exception 'Invalid username or password.' using errcode='P0001';
  end if;

  -- Reset failed attempts on success
  delete from public.rate_limits where client_key = login_fail_key;
  update public.app_users set failed_attempts = 0, locked_until = null where id = u.id;

  -- Single concurrent session: invalidate older sessions
  delete from public.app_sessions where user_id = u.id;

  -- Generate session with SHA-256 hash
  t := extensions.gen_random_uuid();
  t_hash := encode(extensions.digest(t::text, 'sha256'), 'hex');

  -- Session expiry strictly set to 12 hours (43200 seconds)
  insert into public.app_sessions (token_hash, user_id, expires_at)
  values (t_hash, u.id, now() + interval '12 hours');

  select org_id into user_org_id from public.org_members where user_id = u.id limit 1;
  if user_org_id is not null then
    perform public._log_activity(
      user_org_id, null, u.id, 'login_success',
      json_build_object('username', u.username, 'ip', client_ip)
    );
  end if;

  return json_build_object(
    'token', t,
    'user_id', u.id,
    'username', u.username,
    'must_change_password', coalesce(u.must_change_password, false)
  );
end $$;

commit;
