-- =============================================================================
-- patch_fix_audit.sql: Resolves All Audit Check Failures (01, 02, 04, 06, 13, 17)
-- Safe to execute on live database. Preserves all tables and user data.
-- =============================================================================

begin;

-- -----------------------------------------------------------------------------
-- 1. Fix Check 04: Revoke direct public/anon access to profiles table
-- -----------------------------------------------------------------------------
revoke all on table public.profiles from anon;

-- -----------------------------------------------------------------------------
-- 2. Fix Checks 01 & 02: Revoke EXECUTE from anon & public on all internal & admin functions
-- -----------------------------------------------------------------------------
do $$
declare
  r record;
begin
  for r in (
    select n.nspname, p.proname, pg_get_function_identity_arguments(p.oid) as args
    from pg_proc p join pg_namespace n on n.oid=p.pronamespace
    where n.nspname='public'
      and (p.proname like '\_%' or p.proname in ('admin_setup_owner_and_org','org_create','admin_create_organisation','check_rate_limit'))
  ) loop
    execute format('revoke execute on function %I.%I(%s) from public, anon, authenticated', r.nspname, r.proname, r.args);
    execute format('grant execute on function %I.%I(%s) to service_role', r.nspname, r.proname, r.args);
  end loop;
end $$;

-- -----------------------------------------------------------------------------
-- 3. Fix Check 06: Update app_signup session expiry to 12 hours (matching app_login)
-- -----------------------------------------------------------------------------
create or replace function public.app_signup(p_username text, p_password text, p_invite_token text default null)
returns json language plpgsql security definer set search_path = public, extensions as $$
declare
  u public.app_users;
  t uuid;
  uname text := btrim(p_username);
  hdrs jsonb;
  client_ip text := 'local';
  user_count int;
  inv public.invitations;
begin
  if uname is null or length(uname) < 3 then raise exception 'Username must be at least 3 characters.' using errcode='P0001'; end if;
  if uname !~ '^[A-Za-z0-9_.\-]+$' then raise exception 'Username can only use letters, numbers, dot, dash and underscore.' using errcode='P0001'; end if;
  
  if p_password is null or length(p_password) < 10 or p_password !~ '[0-9]' or p_password !~ '[a-z]' or p_password !~ '[A-Z]' then
    raise exception 'Password must be at least 10 characters long and contain uppercase letters, lowercase letters, and numbers.' using errcode='P0001';
  end if;

  select count(*) into user_count from public.app_users;

  if user_count > 0 then
    if p_invite_token is null or btrim(p_invite_token) = '' then
      raise exception 'Public registration is closed. Please ask your organisation Owner for an invitation link to create an account.' using errcode='P0001';
    end if;

    select * into inv from public.invitations
    where token_hash = encode(digest(btrim(p_invite_token), 'sha256'), 'hex')
      and used_at is null
      and expires_at > now();

    if inv.id is null then
      raise exception 'Invalid or expired invitation token. Please request a new invitation.' using errcode='P0001';
    end if;
  end if;

  begin
    hdrs := current_setting('request.headers', true)::jsonb;
    if hdrs is not null then
      client_ip := coalesce(hdrs->>'cf-connecting-ip', hdrs->>'x-real-ip', 'local');
    end if;
  exception when others then null;
  end;

  if not public.check_rate_limit('signup:' || client_ip, 10, 3600) then
    raise exception 'Too many signup attempts from this network. Please wait before creating more accounts.' using errcode='P0001';
  end if;

  if exists (select 1 from public.app_users where lower(username) = lower(uname)) then
    raise exception 'That username is already taken.' using errcode='P0001';
  end if;

  insert into public.app_users (username, password_hash)
    values (uname, crypt(p_password, gen_salt('bf', 10))) returning * into u;
  insert into public.profiles (id, username)
    values (u.id, u.username) on conflict (id) do nothing;
  t := extensions.gen_random_uuid();
  insert into public.app_sessions (token, token_hash, user_id, expires_at)
    values (extensions.gen_random_uuid(), encode(sha256(t::text::bytea), 'hex'), u.id, now() + interval '12 hours');

  if inv.id is not null then
    update public.invitations set used_at = now(), used_by = u.id where id = inv.id;
    insert into public.org_members (org_id, user_id, role)
    values (inv.org_id, u.id, inv.role)
    on conflict (org_id, user_id) do nothing;
    perform public._log_activity(inv.org_id, null, u.id, 'invite_accepted', json_build_object('username', u.username, 'role', inv.role));
  end if;

  return json_build_object('token', t, 'user_id', u.id, 'username', u.username);
end $$;

-- -----------------------------------------------------------------------------
-- 4. Fix Check 13: Enforce client isolation for regular staff (shared_with check)
-- -----------------------------------------------------------------------------
create or replace function public.client_list(p_token uuid, p_org uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  u_role text;
  u_name text;
begin
  u_role := public._require_member(uid, p_org);
  select username into u_name from public.app_users where id = uid;
  return coalesce((
    select json_agg(to_jsonb(c) order by c.updated_at desc)
    from public.clients c
    where c.org_id = p_org
      and (
        u_role in ('owner', 'admin')
        or c.created_by = u_name
        or c.shared_with @> jsonb_build_array(u_name)
        or c.shared_with @> jsonb_build_array('*')
      )
  ), '[]'::json);
end $$;

-- -----------------------------------------------------------------------------
-- 5. Fix Check 17: Mask organisation join code for non-admin staff
-- -----------------------------------------------------------------------------
create or replace function public.org_list(p_token uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare uid uuid := public._require_user(p_token);
begin
  return coalesce((
    select json_agg(json_build_object(
             'id', o.id, 'name', o.name,
             'code', case when m.role in ('owner', 'admin') then o.code else null end,
             'code_rotated_at', o.code_rotated_at, 'role', m.role,
             'current_template_version', o.current_template_version,
             'members', (select count(*) from public.org_members x where x.org_id = o.id)
           ) order by o.name)
    from public.org_members m join public.organisations o on o.id = m.org_id
    where m.user_id = uid
  ), '[]'::json);
end $$;

commit;
