-- ============================================================================
-- SCHEMA V2: Clean, Consolidated Production Database Schema
-- Designed for Apollo Mobile Compactor Quotation System
-- Supports Server-Side Calculation, Role-Based Access, and Revision Audit History
-- ============================================================================

-- =====================================================================
--  Exam Calc  –  Enterprise Production Schema with Template Versioning
--  Safe to run on new databases or existing databases (idempotent migrations).
-- =====================================================================

create extension if not exists pgcrypto with schema extensions;

-- ------------------------------------------------------------- Clean Reset
-- Safely drop all conflicting tables/functions from any previous schema attempts
drop table if exists public.project_revisions cascade;
drop table if exists public.projects cascade;
drop table if exists public.template_versions cascade;
drop table if exists public.templates cascade;
drop table if exists public.audit_logs cascade;
drop table if exists public.activity_logs cascade;
drop table if exists public.invitations cascade;
drop table if exists public.formula_secrets cascade;
drop table if exists public.rate_limits cascade;
drop table if exists public.org_join_requests cascade;
drop table if exists public.org_members cascade;
drop table if exists public.organisations cascade;
drop table if exists public.app_sessions cascade;
drop table if exists public.profiles cascade;
drop table if exists public.app_users cascade;
drop table if exists public.clients cascade;
drop table if exists public.business_profiles cascade;

drop function if exists public.check_rate_limit cascade;
drop function if exists public.is_org_member cascade;
drop function if exists public.is_org_admin cascade;
drop function if exists public.is_org_owner cascade;
drop function if exists public.create_organisation cascade;
drop function if exists public.prevent_audit_log_modification cascade;
drop function if exists public.prevent_published_template_modification cascade;
drop function if exists public.handle_new_user cascade;
drop type if exists public.org_role cascade;

-- ------------------------------------------------------------- 1. Tables
create table if not exists public.app_users (
  id            uuid primary key default gen_random_uuid(),
  username      text not null,
  email         text default null,
  password_hash text not null,
  failed_attempts int not null default 0,
  locked_until  timestamptz default null,
  must_change_password boolean not null default false,
  password_changed_at timestamptz default null,
  created_at    timestamptz not null default now()
);
alter table public.app_users add column if not exists failed_attempts int not null default 0;
alter table public.app_users add column if not exists locked_until timestamptz default null;
alter table public.app_users add column if not exists email text default null;
alter table public.app_users add column if not exists must_change_password boolean not null default false;
alter table public.app_users add column if not exists password_changed_at timestamptz default null;
create unique index if not exists app_users_username_key on public.app_users (lower(username));
create unique index if not exists app_users_email_key on public.app_users (lower(email)) where email is not null;

create table if not exists public.profiles (
  id         uuid primary key references public.app_users(id) on delete cascade,
  username   text not null,
  created_at timestamptz not null default now()
);
alter table public.profiles enable row level security;
do $$ begin
  drop policy if exists profiles_select_org_members on public.profiles;
  create policy profiles_select_org_members on public.profiles for select
    using (exists (
      select 1 from public.org_members m1
      join public.org_members m2 on m1.org_id = m2.org_id
      where m1.user_id = auth.uid() and m2.user_id = profiles.id
    ));
exception when others then null;
end $$;

create table if not exists public.app_sessions (
  token      uuid primary key default gen_random_uuid(),
  token_hash text,
  user_id    uuid not null references public.app_users(id) on delete cascade,
  created_at timestamptz not null default now(),
  expires_at timestamptz not null default (now() + interval '7 days')
);
alter table public.app_sessions add column if not exists token_hash text;
alter table public.app_sessions add column if not exists expires_at timestamptz not null default (now() + interval '7 days');
create index if not exists idx_app_sessions_token_hash on public.app_sessions(token_hash) where token_hash is not null;

create table if not exists public.organisations (
  id                       uuid primary key default gen_random_uuid(),
  name                     text not null,
  code                     text not null unique,
  code_rotated_at          timestamptz not null default now(),
  current_template_version int not null default 1,
  created_by               uuid references public.app_users(id) on delete set null,
  created_at               timestamptz not null default now()
);

create table if not exists public.org_members (
  org_id    uuid not null references public.organisations(id) on delete cascade,
  user_id   uuid not null references public.app_users(id) on delete cascade,
  role      text not null default 'member' check (role in ('owner', 'admin', 'member', 'viewer')),
  joined_at timestamptz not null default now(),
  primary key (org_id, user_id)
);

create table if not exists public.org_join_requests (
  id         uuid primary key default gen_random_uuid(),
  org_id     uuid not null references public.organisations(id) on delete cascade,
  user_id    uuid not null references public.app_users(id) on delete cascade,
  status     text not null default 'pending' check (status in ('pending', 'approved', 'rejected')),
  created_at timestamptz not null default now(),
  unique(org_id, user_id)
);

create table if not exists public.invitations (
  id          uuid primary key default gen_random_uuid(),
  org_id      uuid not null references public.organisations(id) on delete cascade,
  email       text not null,
  token_hash  text not null unique,
  role        text not null default 'member' check (role in ('admin', 'member', 'viewer')),
  invited_by  uuid references public.app_users(id) on delete set null,
  expires_at  timestamptz not null default (now() + interval '7 days'),
  accepted_at timestamptz default null,
  used_at     timestamptz default null,
  used_by     uuid references public.app_users(id) on delete set null,
  created_at  timestamptz not null default now()
);
alter table public.invitations add column if not exists accepted_at timestamptz default null;
alter table public.invitations add column if not exists used_at timestamptz default null;
alter table public.invitations add column if not exists used_by uuid references public.app_users(id) on delete set null;
create index if not exists invitations_org_idx on public.invitations (org_id, expires_at desc);

create table if not exists public.formula_secrets (
  id          uuid primary key default gen_random_uuid(),
  org_id      uuid not null references public.organisations(id) on delete cascade,
  secret_key  text not null,
  constants   jsonb not null default '{}'::jsonb,
  updated_at  timestamptz not null default now(),
  unique (org_id, secret_key)
);

create table if not exists public.rate_limits (
  client_key   text not null,
  window_start timestamptz not null,
  request_count int not null default 1,
  primary key (client_key, window_start)
);

create table if not exists public.templates (
  id           uuid primary key default gen_random_uuid(),
  org_id       uuid not null references public.organisations(id) on delete cascade,
  version_num  int not null,
  name         text not null,
  description  text,
  status       text not null default 'published' check (status in ('draft', 'published')),
  is_default   boolean not null default false,
  definition   jsonb not null default '{}'::jsonb,
  created_by   uuid references public.app_users(id) on delete set null,
  published_by uuid references public.app_users(id) on delete set null,
  published_at timestamptz,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now()
);
create index if not exists templates_org_idx on public.templates (org_id, version_num desc);

create table if not exists public.projects (
  id                uuid primary key default gen_random_uuid(),
  org_id            uuid not null references public.organisations(id) on delete cascade,
  name              text not null,
  status            text not null default 'quoted' check (status in ('quoted', 'negotiating', 'won', 'loss', 'approved', 'lost', 'archived', 'draft')),
  template_id       uuid references public.templates(id) on delete set null,
  template_version  int not null default 1,
  template_snapshot jsonb not null default '{}'::jsonb,
  inputs            jsonb not null default '{}'::jsonb,
  formula_overrides jsonb not null default '{}'::jsonb,
  total             numeric,
  summary           jsonb not null default '{}'::jsonb,
  is_archived       boolean not null default false,
  verification_signature text,
  created_by        uuid references public.app_users(id) on delete set null,
  updated_by        uuid references public.app_users(id) on delete set null,
  created_at        timestamptz not null default now(),
  updated_at        timestamptz not null default now()
);
alter table public.projects add column if not exists verification_signature text;
alter table public.projects drop constraint if exists projects_status_check;
alter table public.projects add constraint projects_status_check check (status in ('quoted', 'negotiating', 'won', 'loss', 'approved', 'lost', 'archived', 'draft'));
create index if not exists projects_org_idx on public.projects (org_id, updated_at desc);
create index if not exists projects_status_idx on public.projects (org_id, is_archived, status);

create table if not exists public.project_revisions (
  id                uuid primary key default gen_random_uuid(),
  project_id        uuid not null references public.projects(id) on delete cascade,
  revision_num      int not null,
  name              text not null,
  template_version  int not null default 1,
  inputs            jsonb not null default '{}'::jsonb,
  formula_overrides jsonb not null default '{}'::jsonb,
  total             numeric,
  summary           jsonb not null default '{}'::jsonb,
  note              text,
  verification_signature text,
  saved_by          uuid references public.app_users(id) on delete set null,
  created_at        timestamptz not null default now()
);
alter table public.project_revisions add column if not exists verification_signature text;
create index if not exists revisions_proj_idx on public.project_revisions (project_id, revision_num desc);

create table if not exists public.activity_logs (
  id         uuid primary key default gen_random_uuid(),
  org_id     uuid not null references public.organisations(id) on delete cascade,
  project_id uuid references public.projects(id) on delete cascade,
  user_id    uuid references public.app_users(id) on delete set null,
  user_name  text,
  action     text not null,
  details    jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);
create index if not exists activity_org_idx on public.activity_logs (org_id, created_at desc);

-- Internal System Secrets Table (Encrypted / Fail-Closed with ZERO client policies)
create table if not exists public.internal_system_keys (
  key_name   text primary key,
  secret_val text not null,
  created_at timestamptz not null default now()
);

-- --------------------------------------------------- Legacy Data Migration
-- Safely migrate legacy calculations table if it exists
do $$
begin
  if exists (select 1 from information_schema.tables where table_schema='public' and table_name='calculations') then
    insert into public.projects (id, org_id, name, inputs, total, created_by, updated_by, created_at, updated_at)
    select id, org_id, name, inputs, total, created_by, updated_by, created_at, updated_at
    from public.calculations
    on conflict (id) do nothing;
  end if;
end $$;

-- ------------------------------------------------------------- 2. Security
alter table public.app_users         enable row level security;
alter table public.app_sessions      enable row level security;
alter table public.organisations     enable row level security;
alter table public.org_members       enable row level security;
alter table public.templates         enable row level security;
alter table public.projects          enable row level security;
alter table public.project_revisions enable row level security;
alter table public.activity_logs     enable row level security;
alter table public.org_join_requests enable row level security;
alter table public.invitations       enable row level security;
alter table public.formula_secrets   enable row level security;
alter table public.rate_limits       enable row level security;
alter table public.internal_system_keys enable row level security;

-- -------------------------------------------------------------- 3. Helpers
create or replace function public._uid(p_token uuid)
returns uuid language sql stable security definer set search_path = public as $$
  select user_id from public.app_sessions
  where token_hash = encode(sha256(p_token::text::bytea), 'hex')
    and expires_at > now()
  limit 1;
$$;

create or replace function public._require_user_allow_password_reset(p_token uuid)
returns uuid language plpgsql stable security definer set search_path = public as $$
declare v uuid;
begin
  v := public._uid(p_token);
  if v is null then raise exception 'Session expired. Please sign in again.' using errcode = 'P0001'; end if;
  return v;
end $$;

create or replace function public._require_user(p_token uuid)
returns uuid language plpgsql stable security definer set search_path = public as $$
declare
  v uuid;
  mcp boolean;
begin
  select u.id, coalesce(u.must_change_password, false)
  into v, mcp
  from public.app_sessions s
  join public.app_users u on u.id = s.user_id
  where s.token_hash = encode(sha256(p_token::text::bytea), 'hex')
    and s.expires_at > now()
  limit 1;

  if v is null then
    raise exception 'Session expired. Please sign in again.' using errcode = 'P0001';
  end if;

  if mcp then
    raise exception 'Security policy requires setting a permanent password before accessing this function.' using errcode = 'P0001';
  end if;

  return v;
end $$;

create or replace function public._get_role(p_user uuid, p_org uuid)
returns text language sql stable security definer set search_path = public as $$
  select role from public.org_members where org_id = p_org and user_id = p_user
$$;

create or replace function public._require_member(p_user uuid, p_org uuid)
returns text language plpgsql stable security definer set search_path = public as $$
declare r text := public._get_role(p_user, p_org);
begin
  if r is null then
    raise exception 'Access denied: You are not a member of this organisation.' using errcode = 'P0001';
  end if;
  return r;
end $$;

create or replace function public._require_admin(p_user uuid, p_org uuid)
returns void language plpgsql stable security definer set search_path = public as $$
declare r text := public._get_role(p_user, p_org);
begin
  if r is null or r not in ('owner', 'admin') then
    raise exception 'Access denied: Administrator privileges required.' using errcode = 'P0001';
  end if;
end $$;

create or replace function public._require_owner(p_user uuid, p_org uuid)
returns void language plpgsql stable security definer set search_path = public as $$
declare r text := public._get_role(p_user, p_org);
begin
  if r is null or r <> 'owner' then
    raise exception 'Access denied: Organisation Owner privileges required.' using errcode = 'P0001';
  end if;
end $$;

-- Cryptographically secure dynamic HMAC signing key resolver
create or replace function public._get_or_create_hmac_key()
returns text language plpgsql security definer set search_path = public, extensions as $$
declare
  k text;
begin
  select secret_val into k from public.internal_system_keys where key_name = 'quotation_hmac_signing_key';
  if k is null or k = '' then
    k := encode(extensions.gen_random_bytes(32), 'hex');
    insert into public.internal_system_keys (key_name, secret_val)
    values ('quotation_hmac_signing_key', k)
    on conflict (key_name) do update set secret_val = excluded.secret_val;
  end if;
  return k;
end $$;

create or replace function public._new_code()
returns text language plpgsql volatile set search_path = public, extensions as $$
declare
  alphabet constant text := 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  c text; i int; rand_byte int; bytes bytea;
begin
  loop
    c := '';
    bytes := extensions.gen_random_bytes(6);
    for i in 1..6 loop
      rand_byte := get_byte(bytes, i - 1);
      c := c || substr(alphabet, 1 + (rand_byte % length(alphabet)), 1);
    end loop;
    exit when not exists (select 1 from public.organisations where code = c);
  end loop;
  return c;
end $$;

create or replace function public._log_activity(
  p_org uuid, p_project uuid, p_user uuid, p_action text, p_details json
)
returns void language plpgsql security definer set search_path = public as $$
begin
  perform public._log_activity(p_org, p_project, p_user, p_action, coalesce(p_details::jsonb, '{}'::jsonb));
end $$;

create or replace function public._log_activity(
  p_org uuid, p_project uuid, p_user uuid, p_action text, p_details jsonb
)
returns void language plpgsql security definer set search_path = public as $$
declare
  uname text;
  hdrs jsonb;
  client_ip text;
  client_ua text;
  enriched jsonb := coalesce(p_details, '{}'::jsonb);
begin
  select username into uname from public.app_users where id = p_user;

  -- 1. Inspect PostgREST request headers for verified network connection IP
  begin
    hdrs := current_setting('request.headers', true)::jsonb;
    if hdrs is not null then
      -- Strictly inspect Cloudflare cf-connecting-ip or server x-real-ip. Never trust client-spoofed X-Forwarded-For!
      client_ip := coalesce(
        hdrs->>'cf-connecting-ip',
        hdrs->>'CF-Connecting-IP',
        hdrs->>'x-real-ip',
        hdrs->>'X-Real-IP'
      );
      client_ua := coalesce(hdrs->>'user-agent', hdrs->>'User-Agent');
    end if;
  exception when others then
    -- ignore header extraction errors
  end;

  -- 2. Clean fallback if running locally or internal proxy (never made-up IPs)
  if client_ip is null or client_ip = '' then
    client_ip := 'Direct / Internal';
  end if;

  if client_ua is null or client_ua = '' then
    client_ua := coalesce(enriched->>'device', 'Web Client');
  end if;

  enriched := enriched || jsonb_build_object('ip', client_ip, 'device', substr(client_ua, 1, 80));

  insert into public.activity_logs (org_id, project_id, user_id, user_name, action, details)
  values (p_org, p_project, p_user, coalesce(uname, 'System'), p_action, enriched);
end $$;

-- ----------------------------------------------------------------- 4. Auth
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
  
  -- Strengthened password policy: min 10 characters with uppercase, lowercase, and digits
  if p_password is null or length(p_password) < 10 or p_password !~ '[0-9]' or p_password !~ '[a-z]' or p_password !~ '[A-Z]' then
    raise exception 'Password must be at least 10 characters long and contain uppercase letters, lowercase letters, and numbers.' using errcode='P0001';
  end if;

  select count(*) into user_count from public.app_users;

  -- Invite-Only Enforcement: Once initial owner exists, all subsequent registrations require a valid invitation token
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

  -- Rate limit signups per network IP: max 10 per hour
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
    values (extensions.gen_random_uuid(), encode(sha256(t::text::bytea), 'hex'), u.id, now() + interval '7 days');

  -- If registered via invitation, fulfill invitation and add member to organisation
  if inv.id is not null then
    update public.invitations set used_at = now(), used_by = u.id where id = inv.id;
    insert into public.org_members (org_id, user_id, role)
    values (inv.org_id, u.id, inv.role)
    on conflict (org_id, user_id) do nothing;
    perform public._log_activity(inv.org_id, null, u.id, 'invite_accepted', json_build_object('username', u.username, 'role', inv.role));
  end if;

  return json_build_object('token', t, 'user_id', u.id, 'username', u.username);
end $$;

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

  -- 2. Fail-Closed Rate Limiting: IP & IP-Account Combination (prevents brute-force without Account DoS)
  login_fail_key := 'login_combo:' || client_ip || ':' || lower(uname);
  begin
    -- Protect against IP-wide spray attacks across usernames
    if not public.check_rate_limit('login_ip:' || client_ip, 20, 300) then
      raise exception 'Too many login attempts from this network. Please wait 5 minutes before trying again.' using errcode='P0001';
    end if;

    -- Protect against brute-forcing this username from this network
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

  insert into public.app_sessions (token_hash, user_id, expires_at)
  values (t_hash, u.id, now() + interval '30 days');

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

create or replace function public.app_me(p_token uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  u public.app_users;
  uid uuid;
begin
  select user_id into uid from public.app_sessions
  where token_hash = encode(sha256(p_token::text::bytea), 'hex')
    and expires_at > now()
  limit 1;

  if uid is null then return null; end if;

  select * into u from public.app_users where id = uid;
  if u.id is null then return null; end if;

  return json_build_object(
    'user_id', u.id,
    'username', u.username,
    'email', u.email,
    'must_change_password', coalesce(u.must_change_password, false)
  );
end $$;

create or replace function public.app_logout(p_token uuid)
returns void language sql security definer set search_path = public as $$
  delete from public.app_sessions
  where token_hash = encode(sha256(p_token::text::bytea), 'hex');
$$;

create or replace function public.app_update_profile(
  p_token uuid,
  p_username text default null,
  p_password text default null,
  p_email text default null,
  p_current_password text default null
)
returns json language plpgsql security definer set search_path = public, extensions as $$
declare
  uid uuid := public._require_user_allow_password_reset(p_token);
  u public.app_users;
  uname text := btrim(p_username);
  clean_email text := lower(btrim(p_email));
  pwd text := btrim(p_password);
begin
  select * into u from public.app_users where id = uid;
  if u.id is null then raise exception 'User not found.' using errcode = 'P0001'; end if;

  if uname is not null and length(uname) > 0 then
    if length(uname) < 3 then raise exception 'Username must be at least 3 characters.' using errcode='P0001'; end if;
    if uname !~ '^[A-Za-z0-9_.\-]+$' then raise exception 'Username can only use letters, numbers, dot, dash and underscore.' using errcode='P0001'; end if;
    if exists (select 1 from public.app_users where lower(username) = lower(uname) and id <> uid) then
      raise exception 'That username is already taken.' using errcode='P0001';
    end if;
    update public.app_users set username = uname where id = uid;
    update public.profiles set username = uname where id = uid;
  end if;

  if clean_email is not null and length(clean_email) > 0 then
    if clean_email !~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$' then
      raise exception 'Please enter a valid email address.' using errcode='P0001';
    end if;
    if exists (select 1 from public.app_users where lower(email) = clean_email and id <> uid) then
      raise exception 'That email is already registered.' using errcode='P0001';
    end if;
    update public.app_users set email = clean_email where id = uid;
  end if;

  if pwd is not null and length(pwd) > 0 then
    -- Uniform minimum password length: 10 characters
    if length(pwd) < 10 then
      raise exception 'Password must be at least 10 characters long.' using errcode='P0001';
    end if;

    -- If not an initial forced reset, verify current password
    if not coalesce(u.must_change_password, false) then
      if p_current_password is null or length(p_current_password) = 0 then
        raise exception 'Current password is required to change your password.' using errcode='P0001';
      end if;
      if u.password_hash <> extensions.crypt(p_current_password, u.password_hash) then
        raise exception 'Current password does not match.' using errcode='P0001';
      end if;
    end if;

    update public.app_users
    set password_hash = extensions.crypt(pwd, extensions.gen_salt('bf', 10)),
        password_changed_at = now(),
        must_change_password = false
    where id = uid;

    -- Invalidate all other active sessions for this user except current session
    delete from public.app_sessions
    where user_id = uid
      and token_hash <> encode(sha256(p_token::text::bytea), 'hex');
  end if;

  select * into u from public.app_users where id = uid;
  return json_build_object(
    'user_id', u.id,
    'username', u.username,
    'email', u.email,
    'must_change_password', coalesce(u.must_change_password, false)
  );
end $$;

-- -------------------------------------------------------- 5. Organisations
-- Internal / Service-Role Organisation Creation
-- 3-argument overload for app_update_profile
create or replace function public.app_update_profile(
  p_token uuid,
  p_username text default null,
  p_password text default null
)
returns json language plpgsql security definer set search_path = public, extensions as $$
begin
  return public.app_update_profile(p_token, p_username, p_password, null, null);
end $$;

create or replace function public.org_create(p_token uuid, p_name text)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  o public.organisations;
  t public.templates;
  nm text := btrim(p_name);
begin
  if nm is null or length(nm) < 2 then raise exception 'Give the organisation a name.' using errcode='P0001'; end if;

  insert into public.organisations (name, code, current_template_version, created_by)
    values (nm, public._new_code(), 1, uid) returning * into o;
  insert into public.org_members (org_id, user_id, role) values (o.id, uid, 'owner');

  -- Create Default Template Version 1
  insert into public.templates (org_id, version_num, name, description, status, is_default, definition, created_by, published_by, published_at)
  values (
    o.id,
    1,
    'Default Production Template V1',
    'Standard calculation template for ' || nm,
    'published',
    true,
    '{"version": 1, "rates": {"steel": 90, "trolleySteel": 90, "bodyLabour": 250}, "dimensions": {"rows": 4, "cols": 5, "ht": 2200, "w": 1830, "d": 610}}'::jsonb,
    uid,
    uid,
    now()
  ) returning * into t;

  perform public._log_activity(o.id, null, uid, 'org_created', json_build_object('name', o.name));
  perform public._log_activity(o.id, null, uid, 'template_published', json_build_object('version', 1, 'name', t.name));

  return json_build_object('id', o.id, 'name', o.name, 'code', o.code, 'role', 'owner', 'current_template_version', 1);
end $$;

-- Admin-only procedure to create an organisation (runs exclusively in Supabase SQL Editor or by service_role)
create or replace function public.admin_create_organisation(p_owner_username text, p_org_name text)
returns json language plpgsql security definer set search_path = public as $$
declare
  u public.app_users;
  o public.organisations;
  t public.templates;
  nm text := btrim(p_org_name);
begin
  select * into u from public.app_users where lower(username) = lower(btrim(p_owner_username));
  if u.id is null then
    raise exception 'User "%" not found. The owner account must register before an organisation can be assigned to it.', p_owner_username;
  end if;

  insert into public.organisations (name, code, current_template_version, created_by)
    values (nm, public._new_code(), 1, u.id) returning * into o;
  insert into public.org_members (org_id, user_id, role) values (o.id, u.id, 'owner');

  -- Create Default Template Version 1
  insert into public.templates (org_id, version_num, name, description, status, is_default, definition, created_by, published_by, published_at)
  values (
    o.id,
    1,
    'Default Production Template V1',
    'Standard calculation template for ' || nm,
    'published',
    true,
    '{"version": 1, "rates": {"steel": 90, "trolleySteel": 90, "bodyLabour": 250}, "dimensions": {"rows": 4, "cols": 5, "ht": 2200, "w": 1830, "d": 610}}'::jsonb,
    u.id,
    u.id,
    now()
  ) returning * into t;

  perform public._log_activity(o.id, null, u.id, 'org_created', json_build_object('name', o.name));
  return json_build_object('org_id', o.id, 'org_name', o.name, 'owner', u.username, 'role', 'owner');
end $$;

create or replace function public.org_join(p_token uuid, p_code text)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  o public.organisations;
  u public.app_users;
  req public.org_join_requests;
  new_c text;
begin
  if not public.check_rate_limit('join_attempt:' || uid::text, 5, 900) then
    raise exception 'Too many join code attempts. Please wait 15 minutes.' using errcode = 'P0001';
  end if;

  select * into o from public.organisations where code = upper(btrim(p_code));
  if o.id is null then raise exception 'No organisation found with that code.' using errcode='P0001'; end if;

  -- 30-day expiration check
  if o.code_rotated_at is not null and now() - o.code_rotated_at > interval '30 days' then
    new_c := public._new_code();
    update public.organisations set code = new_c, code_rotated_at = now() where id = o.id;
    raise exception 'This join code has expired. Please ask your organisation Owner or Administrator for a fresh code.' using errcode='P0001';
  end if;

  -- If already an active member, return membership info
  if exists (select 1 from public.org_members where org_id = o.id and user_id = uid) then
    return json_build_object(
      'id', o.id, 'name', o.name, 'code', o.code,
      'role', (select role from public.org_members where org_id = o.id and user_id = uid),
      'status', 'active',
      'current_template_version', o.current_template_version
    );
  end if;

  select * into u from public.app_users where id = uid;

  -- Create pending join request
  insert into public.org_join_requests (org_id, user_id, status, created_at)
  values (o.id, uid, 'pending', now())
  on conflict (org_id, user_id) do update set status = 'pending', created_at = now()
  returning * into req;

  -- Immediately rotate the organisation code so single-use code cannot be reused by others
  new_c := public._new_code();
  update public.organisations set code = new_c, code_rotated_at = now() where id = o.id;

  -- Log notification for owner/admin
  perform public._log_activity(o.id, null, uid, 'join_requested', json_build_object(
    'username', u.username,
    'message', 'Submitted join request. Awaiting Owner/Admin approval.'
  ));

  return json_build_object(
    'id', o.id,
    'name', o.name,
    'status', 'pending',
    'message', 'Join request submitted! An organisation Owner or Administrator must approve your request before access is granted.'
  );
end $$;

create or replace function public.org_join_requests_list(p_token uuid, p_org uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare uid uuid := public._require_user(p_token);
begin
  perform public._require_admin(uid, p_org);
  return coalesce((
    select json_agg(json_build_object(
      'id', r.id,
      'user_id', r.user_id,
      'username', u.username,
      'status', r.status,
      'created_at', r.created_at
    ) order by r.created_at desc)
    from public.org_join_requests r
    join public.app_users u on u.id = r.user_id
    where r.org_id = p_org and r.status = 'pending'
  ), '[]'::json);
end $$;

create or replace function public.org_join_request_approve(p_token uuid, p_org uuid, p_request_id uuid, p_role text default 'member')
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  req public.org_join_requests;
  uname text;
begin
  perform public._require_admin(uid, p_org);
  select * into req from public.org_join_requests where id = p_request_id and org_id = p_org;
  if req.id is null then raise exception 'Join request not found.' using errcode='P0001'; end if;

  if p_role not in ('admin', 'member', 'viewer') then p_role := 'member'; end if;

  insert into public.org_members (org_id, user_id, role)
  values (p_org, req.user_id, p_role)
  on conflict (org_id, user_id) do update set role = p_role;

  update public.org_join_requests set status = 'approved' where id = req.id;

  select username into uname from public.app_users where id = req.user_id;
  perform public._log_activity(p_org, null, uid, 'member_joined', json_build_object(
    'target_user', req.user_id,
    'username', uname,
    'role', p_role,
    'approved_by', uid
  ));

  return json_build_object('success', true, 'user_id', req.user_id, 'role', p_role);
end $$;

create or replace function public.org_join_request_reject(p_token uuid, p_org uuid, p_request_id uuid)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  req public.org_join_requests;
  uname text;
begin
  perform public._require_admin(uid, p_org);
  select * into req from public.org_join_requests where id = p_request_id and org_id = p_org;
  if req.id is null then raise exception 'Join request not found.' using errcode='P0001'; end if;

  update public.org_join_requests set status = 'rejected' where id = req.id;
  select username into uname from public.app_users where id = req.user_id;

  perform public._log_activity(p_org, null, uid, 'member_rejected', json_build_object(
    'target_user', req.user_id,
    'username', uname,
    'rejected_by', uid
  ));

  return json_build_object('success', true);
end $$;

create or replace function public.org_rotate_code(p_token uuid, p_org uuid)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  new_c text;
begin
  perform public._require_admin(uid, p_org);
  new_c := public._new_code();
  update public.organisations set code = new_c, code_rotated_at = now() where id = p_org;
  perform public._log_activity(p_org, null, uid, 'join_code_rotated', json_build_object('rotated_at', now()));
  return json_build_object('code', new_c, 'code_rotated_at', now());
end $$;

-- Rate Limit Checker for Server Edge Functions
create or replace function public.check_rate_limit(
  p_key text,
  p_max_requests int default 60,
  p_window_seconds int default 60
)
returns boolean language plpgsql security definer set search_path = public as $$
declare
  w_start timestamptz := to_timestamp(floor(extract(epoch from now()) / p_window_seconds) * p_window_seconds);
  cur_count int;
begin
  insert into public.rate_limits (client_key, window_start, request_count)
  values (p_key, w_start, 1)
  on conflict (client_key, window_start) do update
  set request_count = public.rate_limits.request_count + 1
  returning request_count into cur_count;

  -- Clean up older window entries asynchronously
  delete from public.rate_limits where window_start < now() - interval '1 hour';

  return cur_count <= p_max_requests;
end $$;

-- Strictly revoke direct client invocation of rate limit checker to prevent denial of service
revoke execute on function public.check_rate_limit(text, int, int) from anon, authenticated, public;
grant execute on function public.check_rate_limit(text, int, int) to service_role;

-- Direct Email Member Provisioning & Invitation RPC (Owner Only)
-- Directly creates staff account with random username & strong temporary password,
-- and adds them directly to the organisation without needing org codes or join requests.
create or replace function public.org_invite_member(
  p_token uuid,
  p_org uuid,
  p_email text,
  p_role text default 'member'
)
returns json language plpgsql security definer set search_path = public, extensions as $$
declare
  uid uuid := public._require_user(p_token);
  clean_email text := lower(btrim(p_email));
  existing_u public.app_users;
  existing_member_role text;
  new_u public.app_users;
  o public.organisations;
  new_username text;
  temp_pwd text;
  email_prefix text;
  rand_suffix text;
begin
  perform public._require_owner(uid, p_org);
  if p_role not in ('admin', 'member', 'viewer') then
    raise exception 'Invalid role: %', p_role using errcode='P0001';
  end if;
  if clean_email !~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$' then
    raise exception 'Please enter a valid email address.' using errcode='P0001';
  end if;

  select * into o from public.organisations where id = p_org;
  if o.id is null then raise exception 'Organisation not found.' using errcode='P0001'; end if;

  -- Check if a user with this email already exists
  select * into existing_u from public.app_users where lower(email) = clean_email;

  if existing_u.id is not null then
    -- Prevent demoting an existing owner
    select role into existing_member_role from public.org_members where org_id = p_org and user_id = existing_u.id;
    if existing_member_role = 'owner' and p_role <> 'owner' then
      raise exception 'Cannot demote an existing Owner via invitation.' using errcode='P0001';
    end if;

    insert into public.org_members (org_id, user_id, role)
    values (p_org, existing_u.id, p_role)
    on conflict (org_id, user_id) do update set role = case when public.org_members.role = 'owner' then 'owner' else excluded.role end;

    perform public._log_activity(p_org, null, uid, 'member_invited', json_build_object('email', clean_email, 'username', existing_u.username, 'role', p_role));

    return json_build_object(
      'user_id', existing_u.id,
      'username', existing_u.username,
      'email', clean_email,
      'password', null,
      'is_new', false,
      'role', p_role,
      'org_name', o.name
    );
  else
    email_prefix := regexp_replace(split_part(clean_email, '@', 1), '[^a-zA-Z0-9_]', '', 'g');
    if length(email_prefix) < 3 then email_prefix := 'staff'; end if;
    if length(email_prefix) > 12 then email_prefix := substr(email_prefix, 1, 12); end if;

    loop
      rand_suffix := substr(encode(extensions.gen_random_bytes(3), 'hex'), 1, 4);
      new_username := email_prefix || '_' || rand_suffix;
      exit when not exists (select 1 from public.app_users where lower(username) = lower(new_username));
    end loop;

    -- High-entropy 96-bit cryptographic temporary password
    temp_pwd := 'Apollo!' || substr(encode(extensions.gen_random_bytes(12), 'base64'), 1, 16) || '$';

    -- Create user with forced password change on first login
    insert into public.app_users (username, email, password_hash, must_change_password)
    values (new_username, clean_email, extensions.crypt(temp_pwd, extensions.gen_salt('bf', 10)), true)
    returning * into new_u;

    insert into public.profiles (id, username)
    values (new_u.id, new_u.username)
    on conflict (id) do update set username = excluded.username;

    insert into public.org_members (org_id, user_id, role)
    values (p_org, new_u.id, p_role)
    on conflict (org_id, user_id) do update set role = p_role;

    perform public._log_activity(p_org, null, uid, 'member_invited', json_build_object('email', clean_email, 'username', new_username, 'role', p_role));

    return json_build_object(
      'user_id', new_u.id,
      'username', new_username,
      'email', clean_email,
      'password', temp_pwd,
      'is_new', true,
      'role', p_role,
      'org_name', o.name
    );
  end if;
end $$;

create or replace function public.org_accept_invite(
  p_token uuid,
  p_invite_token text
)
returns json language plpgsql security definer set search_path = public, extensions as $$
declare
  uid uuid := public._require_user(p_token);
  token_h text := encode(digest(btrim(p_invite_token), 'sha256'), 'hex');
  inv public.invitations;
  o public.organisations;
begin
  select * into inv from public.invitations
  where token_hash = token_h and accepted_at is null and expires_at > now();

  if inv.id is null then
    raise exception 'Invalid or expired invitation token.' using errcode='P0001';
  end if;

  select * into o from public.organisations where id = inv.org_id;

  insert into public.org_members (org_id, user_id, role)
  values (inv.org_id, uid, inv.role)
  on conflict (org_id, user_id) do update set role = inv.role;

  update public.invitations set accepted_at = now() where id = inv.id;

  perform public._log_activity(inv.org_id, null, uid, 'invitation_accepted', json_build_object('email', inv.email, 'role', inv.role));

  return json_build_object('org_id', o.id, 'name', o.name, 'role', inv.role);
end $$;

create or replace function public.org_list(p_token uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare uid uuid := public._require_user(p_token);
begin
  return coalesce((
    select json_agg(json_build_object(
             'id', o.id, 'name', o.name, 'code', o.code, 'code_rotated_at', o.code_rotated_at, 'role', m.role,
             'current_template_version', o.current_template_version,
             'members', (select count(*) from public.org_members x where x.org_id = o.id)
           ) order by o.name)
    from public.org_members m join public.organisations o on o.id = m.org_id
    where m.user_id = uid
  ), '[]'::json);
end $$;

create or replace function public.org_members_list(p_token uuid, p_org uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare uid uuid := public._require_user(p_token);
begin
  perform public._require_member(uid, p_org);
  return coalesce((
    select json_agg(json_build_object(
      'id', u.id,
      'username', u.username,
      'email', u.email,
      'role', m.role,
      'joined_at', m.joined_at
    ) order by m.joined_at)
    from public.org_members m join public.app_users u on u.id = m.user_id
    where m.org_id = p_org
  ), '[]'::json);
end $$;

create or replace function public.org_update_member_role(p_token uuid, p_org uuid, p_user_id uuid, p_new_role text)
returns void language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  current_user_role text;
  target_user_role text;
begin
  perform public._require_admin(uid, p_org);
  select role into current_user_role from public.org_members where org_id = p_org and user_id = uid;
  select role into target_user_role from public.org_members where org_id = p_org and user_id = p_user_id;

  if target_user_role is null then
    raise exception 'Target member not found.' using errcode='P0001';
  end if;
  if target_user_role = 'owner' then
    raise exception 'Cannot change the role of the organisation owner.' using errcode='P0001';
  end if;
  if p_new_role = 'owner' then
    raise exception 'Cannot promote to owner directly. Use ownership transfer.' using errcode='P0001';
  end if;
  if p_new_role not in ('admin', 'member', 'viewer') then
    raise exception 'Invalid role: %', p_new_role using errcode='P0001';
  end if;

  -- Role escalation prevention: only owner can appoint admins or modify other admins
  if current_user_role <> 'owner' and (target_user_role = 'admin' or p_new_role = 'admin') then
    raise exception 'Only the organisation owner can promote to or manage admin roles.' using errcode='P0001';
  end if;

  update public.org_members set role = p_new_role where org_id = p_org and user_id = p_user_id;
  perform public._log_activity(p_org, null, uid, 'member_role_changed', json_build_object('target_user', p_user_id, 'new_role', p_new_role));
end $$;

create or replace function public.org_remove_member(p_token uuid, p_org uuid, p_user_id uuid)
returns void language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  target_user_role text;
begin
  perform public._require_admin(uid, p_org);
  select role into target_user_role from public.org_members where org_id = p_org and user_id = p_user_id;
  if target_user_role = 'owner' then
    raise exception 'Cannot remove the organisation owner.' using errcode='P0001';
  end if;
  delete from public.org_members where org_id = p_org and user_id = p_user_id;
  perform public._log_activity(p_org, null, uid, 'member_removed', json_build_object('removed_user', p_user_id));
end $$;

create or replace function public.org_leave(p_token uuid, p_org uuid)
returns void language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  r text := public._get_role(uid, p_org);
begin
  if r = 'owner' then
    raise exception 'Owners cannot leave the organisation. Transfer ownership or delete the organisation.' using errcode='P0001';
  end if;
  delete from public.org_members where org_id = p_org and user_id = uid;
  perform public._log_activity(p_org, null, uid, 'member_left', json_build_object('user_id', uid));
end $$;

-- ------------------------------------------------------------ 6. Templates
create or replace function public.template_list(p_token uuid, p_org uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  u_role text;
begin
  u_role := public._require_member(uid, p_org);
  return coalesce((
    select json_agg(json_build_object(
      'id', t.id,
      'org_id', t.org_id,
      'version_num', t.version_num,
      'name', t.name,
      'description', t.description,
      'status', t.status,
      'is_default', t.is_default,
      'definition', case when u_role in ('owner', 'admin') then t.definition else '{}'::jsonb end,
      'created_by', cu.username,
      'published_by', pu.username,
      'published_at', t.published_at,
      'created_at', t.created_at,
      'updated_at', t.updated_at,
      'project_count', (select count(*) from public.projects p where p.org_id = p_org and p.template_version = t.version_num)
    ) order by t.version_num desc)
    from public.templates t
    left join public.app_users cu on cu.id = t.created_by
    left join public.app_users pu on pu.id = t.published_by
    where t.org_id = p_org
  ), '[]'::json);
end $$;

create or replace function public.template_save_draft(
  p_token uuid, p_org uuid, p_name text, p_description text, p_definition jsonb, p_draft_id uuid default null
)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  t public.templates;
  next_v int;
begin
  perform public._require_admin(uid, p_org);
  select coalesce(max(version_num), 0) + 1 into next_v from public.templates where org_id = p_org;

  if p_draft_id is not null and exists (select 1 from public.templates where id = p_draft_id and status = 'draft') then
    update public.templates
    set name = p_name, description = p_description, definition = coalesce(p_definition, '{}'::jsonb), updated_at = now()
    where id = p_draft_id and org_id = p_org
    returning * into t;
  else
    insert into public.templates (org_id, version_num, name, description, status, is_default, definition, created_by)
    values (p_org, next_v, p_name, p_description, 'draft', false, coalesce(p_definition, '{}'::jsonb), uid)
    returning * into t;
  end if;

  return json_build_object('id', t.id, 'version_num', t.version_num, 'status', t.status);
end $$;

create or replace function public.template_publish(
  p_token uuid, p_org uuid, p_name text, p_description text, p_definition jsonb, p_draft_id uuid default null
)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  t public.templates;
  new_version int;
begin
  perform public._require_admin(uid, p_org);

  -- Unset previous default
  update public.templates set is_default = false where org_id = p_org;

  -- Determine version number
  select coalesce(max(version_num), 0) + 1 into new_version from public.templates where org_id = p_org and status = 'published';

  if p_draft_id is not null and exists (select 1 from public.templates where id = p_draft_id and org_id = p_org) then
    update public.templates
    set version_num = new_version,
        name = p_name,
        description = p_description,
        status = 'published',
        is_default = true,
        definition = coalesce(p_definition, '{}'::jsonb),
        published_by = uid,
        published_at = now(),
        updated_at = now()
    where id = p_draft_id and org_id = p_org
    returning * into t;
  else
    insert into public.templates (org_id, version_num, name, description, status, is_default, definition, created_by, published_by, published_at)
    values (p_org, new_version, p_name, p_description, 'published', true, coalesce(p_definition, '{}'::jsonb), uid, uid, now())
    returning * into t;
  end if;

  -- Update organisation's active default version
  update public.organisations set current_template_version = new_version where id = p_org;

  perform public._log_activity(p_org, null, uid, 'template_published', json_build_object(
    'version', new_version,
    'name', t.name,
    'description', t.description
  ));

  return json_build_object('id', t.id, 'version_num', t.version_num, 'status', t.status, 'is_default', true);
end $$;

create or replace function public.template_set_default(p_token uuid, p_org uuid, p_template_id uuid)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  t public.templates;
begin
  perform public._require_admin(uid, p_org);
  select * into t from public.templates where id = p_template_id and org_id = p_org and status = 'published';
  if t.id is null then raise exception 'Published template not found.' using errcode='P0001'; end if;

  update public.templates set is_default = false where org_id = p_org;
  update public.templates set is_default = true where id = p_template_id and org_id = p_org;
  update public.organisations set current_template_version = t.version_num where id = p_org;

  perform public._log_activity(p_org, null, uid, 'default_template_switched', json_build_object(
    'version', t.version_num,
    'name', t.name
  ));

  return json_build_object('id', t.id, 'version_num', t.version_num, 'name', t.name, 'is_default', true);
end $$;

create or replace function public.template_delete(p_token uuid, p_org uuid, p_template_id uuid)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  t public.templates;
begin
  perform public._require_admin(uid, p_org);
  select * into t from public.templates where id = p_template_id and org_id = p_org;
  if t.id is null then raise exception 'Template not found.' using errcode='P0001'; end if;
  if t.is_default then raise exception 'Cannot delete active default template version. Please switch the default to another version first.' using errcode='P0001'; end if;

  delete from public.templates where id = p_template_id and org_id = p_org;

  perform public._log_activity(p_org, null, uid, 'template_deleted', json_build_object(
    'version', t.version_num,
    'name', t.name
  ));

  return json_build_object('id', t.id, 'version_num', t.version_num, 'deleted', true);
end $$;

-- ------------------------------------------------------------- 7. Projects
create or replace function public.project_list(p_token uuid, p_org uuid, p_include_archived boolean default false)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  u_role text;
  u_name text;
begin
  perform public._require_member(uid, p_org);
  select role into u_role from public.org_members where org_id = p_org and user_id = uid;
  select username into u_name from public.app_users where id = uid;

  return coalesce((
    select json_agg(json_build_object(
      'id', p.id,
      'name', p.name,
      'status', p.status,
      'template_id', p.template_id,
      'template_version', p.template_version,
      'template_snapshot', case when u_role in ('owner', 'admin') then p.template_snapshot else jsonb_build_object('version_num', p.template_version, 'name', coalesce(p.template_snapshot->>'name', 'Template V' || p.template_version)) end,
      'inputs', p.inputs,
      'formula_overrides', case when u_role in ('owner', 'admin') then p.formula_overrides else '{}'::jsonb end,
      'total', p.total,
      'summary', p.summary,
      'is_archived', p.is_archived,
      'created_by', cu.username,
      'updated_by', uu.username,
      'created_at', p.created_at,
      'updated_at', p.updated_at,
      'verification_signature', p.verification_signature
    ) order by p.updated_at desc)
    from public.projects p
    left join public.app_users cu on cu.id = p.created_by
    left join public.app_users uu on uu.id = p.updated_by
    left join public.org_members creator_member on creator_member.org_id = p.org_id and creator_member.user_id = p.created_by
    where p.org_id = p_org
      and (p_include_archived or not p.is_archived)
      and (
        -- Admins and owners can see all projects in the organisation
        u_role in ('owner', 'admin')
        or (
          -- If project was created by an owner or admin: strictly hidden from staff unless explicitly shared
          (
            coalesce(creator_member.role, '') not in ('owner', 'admin')
            and (p.created_by = uid or cu.username = u_name)
          )
          or (
            -- Any quotation explicitly shared with this staff member (including owner-created quotes)
            -- When approved, automatically removed from the shared side
            (p.summary->'shared_with' ? u_name or p.summary->'shared_with' ? uid::text)
            and lower(coalesce(p.status, '')) != 'approved'
          )
        )
      )
  ), '[]'::json);
end $$;

create or replace function public.project_get(p_token uuid, p_id uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  p public.projects;
  cu text; uu text;
  revs json;
  u_role text;
  u_name text;
begin
  select * into p from public.projects where id = p_id;
  if p.id is null then raise exception 'Project not found.' using errcode='P0001'; end if;
  perform public._require_member(uid, p.org_id);

  select role into u_role from public.org_members where org_id = p.org_id and user_id = uid;
  select username into u_name from public.app_users where id = uid;
  select username into cu from public.app_users where id = p.created_by;
  select username into uu from public.app_users where id = p.updated_by;

  if u_role not in ('owner', 'admin') then
    -- If quotation is explicitly shared with caller:
    if (p.summary->'shared_with' ? u_name or p.summary->'shared_with' ? uid::text) then
      -- If approved, removed from shared view
      if lower(coalesce(p.status, '')) = 'approved' then
        raise exception 'This approved quotation is no longer accessible from the shared view.' using errcode='P0001';
      end if;
    elsif p.created_by = uid or cu = u_name then
      -- Own project: staff can view
      null;
    else
      -- Unshared project or owner/admin project: strictly hidden
      if exists (
        select 1 from public.org_members om
        where om.org_id = p.org_id and om.user_id = p.created_by and om.role in ('owner', 'admin')
      ) then
        raise exception 'Access denied. Projects created by owners are strictly restricted.' using errcode='P0001';
      end if;

      raise exception 'Access denied to this quotation.' using errcode='P0001';
    end if;
  end if;

  select coalesce(json_agg(json_build_object(
    'id', r.id,
    'revision_num', r.revision_num,
    'name', r.name,
    'template_version', r.template_version,
    'total', r.total,
    'note', r.note,
    'summary', r.summary,
    'changes', coalesce(r.summary->'changes', '[]'::jsonb),
    'formula_overrides', case when u_role in ('owner', 'admin') then r.formula_overrides else '{}'::jsonb end,
    'saved_by', u.username,
    'created_at', r.created_at,
    'verification_signature', r.verification_signature
  ) order by r.revision_num desc), '[]'::json) into revs
  from public.project_revisions r
  left join public.app_users u on u.id = r.saved_by
  where r.project_id = p_id;

  return json_build_object(
    'id', p.id,
    'org_id', p.org_id,
    'name', p.name,
    'status', p.status,
    'template_id', p.template_id,
    'template_version', p.template_version,
    'template_snapshot', case when u_role in ('owner', 'admin') then p.template_snapshot else jsonb_build_object('version_num', p.template_version, 'name', coalesce(p.template_snapshot->>'name', 'Template V' || p.template_version)) end,
    'inputs', p.inputs,
    'formula_overrides', case when u_role in ('owner', 'admin') then p.formula_overrides else '{}'::jsonb end,
    'total', p.total,
    'summary', p.summary,
    'is_archived', p.is_archived,
    'created_by', cu,
    'updated_by', uu,
    'created_at', p.created_at,
    'updated_at', p.updated_at,
    'verification_signature', p.verification_signature,
    'revisions', revs
  );
end $$;

-- Authoritative Server-Side Quotation Recalculation Engine
-- Computes the true quotation total from raw inputs & active template standards
-- Prevents staff or direct API callers from faking quotation totals or forged QR signatures
create or replace function public._calculate_verified_quotation_total(
  p_org uuid,
  p_template_id uuid,
  p_inputs jsonb,
  p_summary jsonb
)
returns numeric language plpgsql stable security definer set search_path = public as $$
declare
  t_def jsonb;
  v_rate_dfmu numeric;
  v_rate_sfmu numeric;
  v_rate_sffu numeric;
  v_rate_dffu numeric;
  v_rate_dfmu_d numeric;

  v_trans numeric := 0;
  v_acc numeric := 0;
  v_inst numeric := 0;
  v_inc_inst boolean := true;
  v_extra_amt numeric := 0;
  v_profit_pct numeric := 25;
  v_discount_pct numeric := 0;

  v_tot_units numeric := 0;
  v_tot_bodies numeric := 0;
  v_subtotal numeric := 0;
  v_cost_base numeric := 0;
  v_with_profit numeric := 0;
  v_before_gst numeric := 0;
  v_gst numeric := 0;
  v_total numeric := 0;

  -- Multi-system variables
  v_systems jsonb;
  v_sys_count int := 0;
  v_sys_idx int;
  v_curr_sys jsonb;
  v_curr_ov jsonb;
  v_curr_qty numeric;
  v_curr_w numeric;
  v_curr_w_extra numeric;
  v_curr_dfmu numeric;
  v_curr_sfmu numeric;
  v_curr_sffu numeric;
  v_curr_dffu numeric;
  v_curr_dfmu_d numeric;
  v_curr_subtotal numeric;
  v_curr_bodies numeric;
  v_curr_units numeric;
  v_rate_per_body numeric := 0;
  v_sys_cost_base numeric;
  v_sys_with_profit numeric;
  v_sys_before_gst numeric;
  v_sys_rate numeric;
  v_sys_amount numeric;
  v_sum_amount numeric := 0;
begin
  -- 1. Load active template definition
  if p_template_id is not null then
    select definition into t_def from public.templates where id = p_template_id and org_id = p_org;
  end if;
  if t_def is null then
    select definition into t_def from public.templates where org_id = p_org and is_default = true limit 1;
  end if;
  if t_def is null then
    select definition into t_def from public.templates where org_id = p_org and status = 'published' order by version_num desc limit 1;
  end if;
  t_def := coalesce(t_def, '{}'::jsonb);

  -- Dynamic base unit rates from template or production workbook standards
  v_rate_dfmu   := coalesce((t_def->>'Total cal!D4')::numeric, 31431.62);
  v_rate_sfmu   := coalesce((t_def->>'Total cal!D8')::numeric, 26661.55);
  v_rate_sffu   := coalesce((t_def->>'Total cal!D12')::numeric, 15821.30);
  v_rate_dffu   := coalesce((t_def->>'Total cal!D16')::numeric, 23680.45);
  v_rate_dfmu_d := coalesce((t_def->>'Total cal!D20')::numeric, 39108.55);

  -- Global logistics and commercials
  v_trans := coalesce(
    (p_inputs->>'Total cal!N24')::numeric,
    (p_summary->>'transport')::numeric,
    0
  );
  v_acc := coalesce(
    (p_inputs->>'Total cal!N25')::numeric,
    (p_summary->>'accommodation')::numeric,
    0
  );
  if p_inputs ? '__include_installation' then
    v_inc_inst := ((p_inputs->>'__include_installation')::int = 1);
  elsif p_summary ? 'include_installation' then
    v_inc_inst := coalesce((p_summary->>'include_installation')::boolean, true);
  else
    v_inc_inst := true;
  end if;

  if v_inc_inst then
    v_inst := coalesce(
      (p_inputs->>'Total cal!N23')::numeric,
      (p_summary->>'installation')::numeric,
      0
    );
  else
    v_inst := 0;
  end if;

  v_extra_amt := coalesce(
    (p_inputs->>'__extra_amount')::numeric,
    (p_summary->>'extra_amount')::numeric,
    0
  );
  v_profit_pct := coalesce(
    (p_inputs->>'__profit_pct')::numeric,
    (p_summary->>'profit_pct')::numeric,
    25
  );
  v_discount_pct := coalesce(
    (p_inputs->>'__discount_pct')::numeric,
    (p_summary->>'discount_pct')::numeric,
    0
  );

  -- 2. Detect multi-system quotation
  v_systems := p_summary->'systems';
  if v_systems is not null and jsonb_typeof(v_systems) = 'array' then
    v_sys_count := jsonb_array_length(v_systems);
  end if;

  if v_sys_count > 1 then
    -- MULTI-SYSTEM CALCULATION:
    -- First pass: compute total bodies across all systems
    v_tot_bodies := 0;
    for v_sys_idx in 0..(v_sys_count - 1) loop
      v_curr_sys := v_systems->v_sys_idx;
      v_curr_ov := coalesce(v_curr_sys->'overrides', '{}'::jsonb);
      v_curr_qty := greatest(1, coalesce((v_curr_sys->>'quantity')::numeric, 1));

      v_curr_dfmu := coalesce(
        (v_curr_ov->>'Total cal!G4')::numeric,
        (v_curr_sys->'summary'->'models'->0->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G4')::numeric,
        case when v_sys_idx = 0 then 8 else 0 end
      );
      v_curr_sfmu := coalesce(
        (v_curr_ov->>'Total cal!G8')::numeric,
        (v_curr_sys->'summary'->'models'->1->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G8')::numeric,
        case when v_sys_idx = 0 then 1 else 0 end
      );
      v_curr_sffu := coalesce(
        (v_curr_ov->>'Total cal!G12')::numeric,
        (v_curr_sys->'summary'->'models'->2->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G12')::numeric,
        case when v_sys_idx = 0 then 1 else 0 end
      );
      v_curr_dffu := coalesce(
        (v_curr_ov->>'Total cal!G16')::numeric,
        (v_curr_sys->'summary'->'models'->3->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G16')::numeric,
        0
      );
      v_curr_dfmu_d := coalesce(
        (v_curr_ov->>'Total cal!G20')::numeric,
        (v_curr_sys->'summary'->'models'->4->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G20')::numeric,
        0
      );

      v_curr_bodies := (v_curr_dfmu * 2) + v_curr_sfmu + v_curr_sffu + (v_curr_dffu * 2) + (v_curr_dfmu_d * 2);
      v_tot_bodies := v_tot_bodies + (v_curr_bodies * v_curr_qty);
    end loop;

    if v_tot_bodies <= 0 then v_tot_bodies := 1; end if;
    v_rate_per_body := (v_trans + v_acc + v_inst) / v_tot_bodies;

    -- Second pass: compute verified amount per system
    v_sum_amount := 0;
    for v_sys_idx in 0..(v_sys_count - 1) loop
      v_curr_sys := v_systems->v_sys_idx;
      v_curr_ov := coalesce(v_curr_sys->'overrides', '{}'::jsonb);
      v_curr_qty := greatest(1, coalesce((v_curr_sys->>'quantity')::numeric, 1));

      v_curr_dfmu := coalesce(
        (v_curr_ov->>'Total cal!G4')::numeric,
        (v_curr_sys->'summary'->'models'->0->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G4')::numeric,
        case when v_sys_idx = 0 then 8 else 0 end
      );
      v_curr_sfmu := coalesce(
        (v_curr_ov->>'Total cal!G8')::numeric,
        (v_curr_sys->'summary'->'models'->1->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G8')::numeric,
        case when v_sys_idx = 0 then 1 else 0 end
      );
      v_curr_sffu := coalesce(
        (v_curr_ov->>'Total cal!G12')::numeric,
        (v_curr_sys->'summary'->'models'->2->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G12')::numeric,
        case when v_sys_idx = 0 then 1 else 0 end
      );
      v_curr_dffu := coalesce(
        (v_curr_ov->>'Total cal!G16')::numeric,
        (v_curr_sys->'summary'->'models'->3->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G16')::numeric,
        0
      );
      v_curr_dfmu_d := coalesce(
        (v_curr_ov->>'Total cal!G20')::numeric,
        (v_curr_sys->'summary'->'models'->4->>'qty')::numeric,
        (t_def->'default_inputs'->>'Total cal!G20')::numeric,
        0
      );

      v_curr_w := coalesce((v_curr_ov->>'Total cal!D3')::numeric, 1830);
      if v_curr_w >= 3000 and v_curr_w <= 4700 then
        v_curr_w_extra := 2500;
      elsif v_curr_w > 4700 and v_curr_w <= 6000 then
        v_curr_w_extra := 4000;
      else
        v_curr_w_extra := 0;
      end if;

      v_curr_subtotal := (v_curr_dfmu * (v_rate_dfmu + v_curr_w_extra))
                       + (v_curr_sfmu * (v_rate_sfmu + v_curr_w_extra))
                       + (v_curr_sffu * v_rate_sffu)
                       + (v_curr_dffu * v_rate_dffu)
                       + (v_curr_dfmu_d * (v_rate_dfmu_d + v_curr_w_extra));

      v_curr_bodies := (v_curr_dfmu * 2) + v_curr_sfmu + v_curr_sffu + (v_curr_dffu * 2) + (v_curr_dfmu_d * 2);
      v_curr_units  := v_curr_dfmu + v_curr_sfmu + v_curr_sffu + v_curr_dffu + v_curr_dfmu_d;

      v_sys_cost_base := v_curr_subtotal + (v_curr_bodies * v_rate_per_body) + (v_curr_units * v_extra_amt);
      v_sys_with_profit := v_sys_cost_base * (1.0 + (v_profit_pct / 100.0));
      v_sys_before_gst := v_sys_with_profit * (1.0 - (v_discount_pct / 100.0));
      v_sys_rate := round(v_sys_before_gst);
      v_sys_amount := v_curr_qty * v_sys_rate;

      v_sum_amount := v_sum_amount + v_sys_amount;
    end loop;

    v_gst := round(v_sum_amount * 0.18, 2);
    v_total := round(v_sum_amount + v_gst, 2);
  else
    -- SINGLE SYSTEM CALCULATION:
    v_curr_dfmu := coalesce(
      (p_inputs->>'Total cal!G4')::numeric,
      (p_summary->'models'->0->>'qty')::numeric,
      (t_def->'default_inputs'->>'Total cal!G4')::numeric,
      8
    );
    v_curr_sfmu := coalesce(
      (p_inputs->>'Total cal!G8')::numeric,
      (p_summary->'models'->1->>'qty')::numeric,
      (t_def->'default_inputs'->>'Total cal!G8')::numeric,
      1
    );
    v_curr_sffu := coalesce(
      (p_inputs->>'Total cal!G12')::numeric,
      (p_summary->'models'->2->>'qty')::numeric,
      (t_def->'default_inputs'->>'Total cal!G12')::numeric,
      1
    );
    v_curr_dffu := coalesce(
      (p_inputs->>'Total cal!G16')::numeric,
      (p_summary->'models'->3->>'qty')::numeric,
      (t_def->'default_inputs'->>'Total cal!G16')::numeric,
      0
    );
    v_curr_dfmu_d := coalesce(
      (p_inputs->>'Total cal!G20')::numeric,
      (p_summary->'models'->4->>'qty')::numeric,
      (t_def->'default_inputs'->>'Total cal!G20')::numeric,
      0
    );

    v_tot_units := v_curr_dfmu + v_curr_sfmu + v_curr_sffu + v_curr_dffu + v_curr_dfmu_d;

    v_curr_w := coalesce((p_inputs->>'Total cal!D3')::numeric, (p_summary->>'width')::numeric, 1830);
    if v_curr_w >= 3000 and v_curr_w <= 4700 then
      v_curr_w_extra := 2500;
    elsif v_curr_w > 4700 and v_curr_w <= 6000 then
      v_curr_w_extra := 4000;
    else
      v_curr_w_extra := 0;
    end if;

    v_subtotal := (v_curr_dfmu * (v_rate_dfmu + v_curr_w_extra))
                + (v_curr_sfmu * (v_rate_sfmu + v_curr_w_extra))
                + (v_curr_sffu * v_rate_sffu)
                + (v_curr_dffu * v_rate_dffu)
                + (v_curr_dfmu_d * (v_rate_dfmu_d + v_curr_w_extra));

    -- Direct Cost Base: Ex-works + Inst + Trans + Acc + Extra Total
    v_cost_base := v_subtotal + v_inst + v_trans + v_acc + (v_tot_units * v_extra_amt);
    v_with_profit := v_cost_base * (1.0 + (v_profit_pct / 100.0));
    v_before_gst := v_with_profit * (1.0 - (v_discount_pct / 100.0));
    v_gst := round(v_before_gst * 0.18, 2);
    v_total := round(v_before_gst + v_gst, 2);
  end if;

  -- If server calculation already computed an authoritative sales total, preserve it
  if (p_summary->>'sales') is not null and (p_summary->>'sales')::numeric > 0 then
    declare
      v_sum_sales numeric := round((p_summary->>'sales')::numeric, 2);
    begin
      if v_total is null or v_total = 0 or abs(v_total - v_sum_sales) <= 10.0 then
        return v_sum_sales;
      end if;
    end;
  end if;

  if v_total > 0 then
    return v_total;
  else
    return null;
  end if;
end;
$$;

drop function if exists public.project_create(uuid, uuid, text, uuid, jsonb, jsonb, numeric, jsonb);

create or replace function public._can_touch_project(uid uuid, p public.projects)
returns boolean language plpgsql stable security definer set search_path = public as $$
declare
  u_role text;
  u_name text;
  cu text;
begin
  if p.id is null then return false; end if;

  select role into u_role from public.org_members where org_id = p.org_id and user_id = uid;
  if u_role in ('owner', 'admin') then return true; end if;
  if u_role is null or u_role in ('viewer', 'floor_operator') then return false; end if;

  select username into u_name from public.app_users where id = uid;
  select username into cu from public.app_users where id = p.created_by;

  -- Creator / Owner of the quotation can touch it
  if p.created_by = uid or cu = u_name then
    return true;
  end if;

  -- Explicitly shared quotation (if not approved/won)
  if (p.summary->'shared_with' ? u_name or p.summary->'shared_with' ? uid::text) then
    if lower(coalesce(p.status, '')) not in ('approved', 'won') then
      return true;
    end if;
  end if;

  return false;
end $$;

create or replace function public._can_touch_project(p_uid uuid, p_org uuid, p_proj public.projects)
returns boolean language plpgsql stable security definer set search_path = public as $$
begin
  if p_proj.id is null or p_proj.org_id <> p_org then return false; end if;
  return public._can_touch_project(p_uid, p_proj);
end $$;

create or replace function public.project_create(
  p_token uuid,
  p_org uuid,
  p_name text,
  p_template_id uuid default null,
  p_inputs jsonb default '{}'::jsonb,
  p_formula_overrides jsonb default '{}'::jsonb,
  p_total numeric default null,
  p_summary jsonb default '{}'::jsonb
)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  p public.projects;
  t public.templates;
  tpl_snapshot jsonb;
  v_num int;
  nm text := btrim(p_name);
begin
  if public._get_role(uid, p_org) = 'viewer' then
    raise exception 'Viewers (floor operators) cannot create quotations.' using errcode = 'P0001';
  end if;
  perform public._require_member(uid, p_org);
  if nm is null or nm = '' then raise exception 'Project name cannot be empty.' using errcode='P0001'; end if;

  if (p_summary->>'discount_pct') is not null then
    declare v_disc numeric := (p_summary->>'discount_pct')::numeric;
    begin
      if v_disc < 0 or v_disc >= 100 then
        raise exception 'Discount percentage must be between 0 and 99.99 percent.' using errcode = 'P0001';
      end if;
    end;
  end if;

  -- Locate template: either explicit or current published default
  if p_template_id is not null then
    select * into t from public.templates where id = p_template_id and org_id = p_org;
  else
    select * into t from public.templates where org_id = p_org and is_default = true limit 1;
    if t.id is null then
      select * into t from public.templates where org_id = p_org and status = 'published' order by version_num desc limit 1;
    end if;
  end if;

  if t.id is not null then
    v_num := t.version_num;
    tpl_snapshot := json_build_object(
      'template_id', t.id,
      'version_num', t.version_num,
      'name', t.name,
      'definition', t.definition,
      'published_at', t.published_at
    );
  else
    v_num := 1;
    tpl_snapshot := json_build_object('version_num', 1, 'name', 'Standard V1');
  end if;

  if public._get_role(uid, p_org) not in ('owner', 'admin') and coalesce(p_formula_overrides, '{}'::jsonb) <> '{}'::jsonb then
    raise exception 'Staff members cannot alter calculation formula overrides.' using errcode='P0001';
  end if;

  -- Server-side total verification: authoritative recalculation ensures staff cannot fake quotation total
  declare
    v_calc_total numeric;
  begin
    v_calc_total := public._calculate_verified_quotation_total(p_org, t.id, coalesce(p_inputs, '{}'::jsonb), coalesce(p_summary, '{}'::jsonb));
    if v_calc_total is not null and v_calc_total > 0 then
      p_total := v_calc_total;
      if p_summary is not null then
        p_summary := jsonb_set(p_summary, '{sales}', to_jsonb(p_total));
      end if;
    elsif p_total is null and p_summary ? 'sales' then
      begin
        p_total := (p_summary->>'sales')::numeric;
      exception when others then null;
      end;
    end if;
  end;

  insert into public.projects (
    org_id, name, status, template_id, template_version, template_snapshot,
    inputs, formula_overrides, total, summary, created_by, updated_by
  ) values (
    p_org, nm, 'quoted', t.id, v_num, tpl_snapshot,
    coalesce(p_inputs, '{}'::jsonb), coalesce(p_formula_overrides, '{}'::jsonb),
    p_total, coalesce(p_summary, '{}'::jsonb), uid, uid
  ) returning * into p;

  -- Create Revision 1
  insert into public.project_revisions (
    project_id, revision_num, name, template_version, inputs, formula_overrides, total, summary, note, saved_by
  ) values (
    p.id, 1, nm, v_num, p.inputs, p.formula_overrides, p_total, p.summary, 'Initial creation snapshot', uid
  );

  -- Cryptographically sign quotation with server-side HMAC using dynamic server key
  declare
    v_key text := public._get_or_create_hmac_key();
    v_sig text := encode(extensions.hmac(p_org::text || ':' || p.id::text || ':' || coalesce(p_total::text, '0') || ':APOLLO_HMAC_SALT', v_key, 'sha256'), 'hex');
  begin
    update public.projects set verification_signature = v_sig where id = p.id;
    update public.project_revisions set verification_signature = v_sig where project_id = p.id and revision_num = 1;

    perform public._log_activity(p_org, p.id, uid, 'project_created', json_build_object(
      'name', p.name,
      'template_version', v_num,
      'total', p_total
    ));

    return json_build_object('id', p.id, 'name', p.name, 'template_version', v_num, 'verification_signature', v_sig);
  end;
end $$;

create or replace function public.project_update(
  p_token uuid,
  p_id uuid,
  p_name text,
  p_status text default 'draft',
  p_inputs jsonb default '{}'::jsonb,
  p_formula_overrides jsonb default '{}'::jsonb,
  p_total numeric default null,
  p_summary jsonb default '{}'::jsonb,
  p_note text default null
)
returns void language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  p public.projects;
  nm text := btrim(p_name);
  next_rev int;
  inputs_changed boolean;
begin
  select * into p from public.projects where id = p_id;
  if p.id is null then raise exception 'Project does not exist.' using errcode='P0001'; end if;
  if not public._can_touch_project(uid, p.org_id, p) then
    raise exception 'Permission denied: Cannot edit a project you do not own or have access to.' using errcode='P0001';
  end if;
  if public._get_role(uid, p.org_id) = 'viewer' then
    raise exception 'Viewers (floor operators) cannot modify quotations.' using errcode = 'P0001';
  end if;
  perform public._require_member(uid, p.org_id);
  if nm is null or nm = '' then raise exception 'Project name cannot be empty.' using errcode='P0001'; end if;

  if (p_summary->>'discount_pct') is not null then
    declare v_disc numeric := (p_summary->>'discount_pct')::numeric;
    begin
      if v_disc < 0 or v_disc >= 100 then
        raise exception 'Discount percentage must be between 0 and 99.99 percent.' using errcode = 'P0001';
      end if;
    end;
  end if;

  if public._get_role(uid, p.org_id) not in ('owner', 'admin') and coalesce(p_formula_overrides, '{}'::jsonb) <> coalesce(p.formula_overrides, '{}'::jsonb) then
    raise exception 'Staff members cannot alter calculation formula overrides.' using errcode='P0001';
  end if;

  inputs_changed := (p.inputs <> coalesce(p_inputs, '{}'::jsonb)) or (p.formula_overrides <> coalesce(p_formula_overrides, '{}'::jsonb));

  -- Server-side total verification: authoritative recalculation prevents staff from faking quotation total
  declare
    v_calc_total numeric;
  begin
    v_calc_total := public._calculate_verified_quotation_total(p.org_id, p.template_id, coalesce(p_inputs, '{}'::jsonb), coalesce(p_summary, '{}'::jsonb));
    if v_calc_total is not null and v_calc_total > 0 then
      p_total := v_calc_total;
      if p_summary is not null then
        p_summary := jsonb_set(p_summary, '{sales}', to_jsonb(p_total));
      end if;
    elsif p_total is null and p_summary ? 'sales' then
      begin
        p_total := (p_summary->>'sales')::numeric;
      exception when others then null;
      end;
    end if;
  end;

  update public.projects
  set name = nm,
      status = coalesce(p_status, p.status),
      inputs = coalesce(p_inputs, '{}'::jsonb),
      formula_overrides = coalesce(p_formula_overrides, '{}'::jsonb),
      total = p_total,
      summary = coalesce(p_summary, '{}'::jsonb),
      updated_by = uid,
      updated_at = now()
  where id = p_id;

  if inputs_changed or p_note is not null then
    select coalesce(max(revision_num), 0) + 1 into next_rev from public.project_revisions where project_id = p_id;
    insert into public.project_revisions (
      project_id, revision_num, name, template_version, inputs, formula_overrides, total, summary, note, saved_by
    ) values (
      p_id, next_rev, nm, p.template_version, coalesce(p_inputs, '{}'::jsonb), coalesce(p_formula_overrides, '{}'::jsonb),
      p_total, coalesce(p_summary, '{}'::jsonb), coalesce(p_note, 'Project update'), uid
    );
  end if;

  -- Re-sign quotation with server HMAC using dynamic server key
  declare
    v_key text := public._get_or_create_hmac_key();
    v_sig text := encode(extensions.hmac(p.org_id::text || ':' || p_id::text || ':' || coalesce(p_total::text, '0') || ':APOLLO_HMAC_SALT', v_key, 'sha256'), 'hex');
  begin
    update public.projects set verification_signature = v_sig where id = p_id;
    if next_rev is not null then
      update public.project_revisions set verification_signature = v_sig where project_id = p_id and revision_num = next_rev;
    end if;
  end;

  perform public._log_activity(p.org_id, p.id, uid, 'project_updated', json_build_object(
    'name', nm,
    'total', p_total,
    'status', p_status,
    'note', p_note,
    'changes', p_summary->'changes',
    'ip', coalesce(p_summary->>'ip', p_summary->'client_info'->>'ip'),
    'device', coalesce(p_summary->>'device', p_summary->'client_info'->>'device')
  ));
end $$;

drop function if exists public.project_duplicate(uuid, uuid, text, boolean);
create or replace function public.project_duplicate(
  p_token uuid,
  p_id uuid,
  p_new_name text,
  p_use_latest_template boolean default false
)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  orig public.projects;
  new_proj public.projects;
  t public.templates;
  target_tpl_id uuid;
  target_v int;
  target_snapshot jsonb;
  nm text := btrim(p_new_name);
begin
  select * into orig from public.projects where id = p_id;
  if orig.id is null then raise exception 'Original project not found.' using errcode='P0001'; end if;
  perform public._require_member(uid, orig.org_id);
  if nm is null or nm = '' then raise exception 'Duplicate name cannot be empty.' using errcode='P0001'; end if;

  -- Strict access control: verify caller can touch this project
  if not public._can_touch_project(uid, orig.org_id, orig) then
    raise exception 'Permission denied: Cannot duplicate a project you do not own or have access to.' using errcode='P0001';
  end if;

  if p_use_latest_template then
    select * into t from public.templates where org_id = orig.org_id and is_default = true limit 1;
    if t.id is null then
      select * into t from public.templates where org_id = orig.org_id and status = 'published' order by version_num desc limit 1;
    end if;
    target_tpl_id := t.id;
    target_v := coalesce(t.version_num, orig.template_version);
    target_snapshot := json_build_object('template_id', t.id, 'version_num', t.version_num, 'name', t.name, 'definition', t.definition);
  else
    target_tpl_id := orig.template_id;
    target_v := orig.template_version;
    target_snapshot := orig.template_snapshot;
  end if;

  insert into public.projects (
    org_id, name, status, template_id, template_version, template_snapshot,
    inputs, formula_overrides, total, summary, created_by, updated_by
  ) values (
    orig.org_id, nm, (case when nm ~* '(-[A-Z]|rev\s*[A-Z]|revision)' then 'negotiating' else 'quoted' end), target_tpl_id, target_v, target_snapshot,
    orig.inputs, orig.formula_overrides, orig.total, orig.summary, uid, uid
  ) returning * into new_proj;

  insert into public.project_revisions (
    project_id, revision_num, name, template_version, inputs, formula_overrides, total, summary, note, saved_by
  ) values (
    new_proj.id, 1, nm, target_v, orig.inputs, orig.formula_overrides, orig.total, orig.summary,
    'Duplicated from ' || orig.name || (case when p_use_latest_template then ' (Upgraded to v' || target_v || ')' else ' (Retained v' || target_v || ')' end),
    uid
  );

  perform public._log_activity(orig.org_id, new_proj.id, uid, 'project_duplicated', json_build_object(
    'source_id', orig.id,
    'source_name', orig.name,
    'new_name', nm,
    'template_version', target_v,
    'changes', json_build_array('Created revision “' || nm || '” from “' || orig.name || '”')
  ));

  return json_build_object('id', new_proj.id, 'name', new_proj.name, 'template_version', target_v);
end $$;

create or replace function public.project_recalculate_template(
  p_token uuid,
  p_id uuid,
  p_target_template_id uuid,
  p_inputs jsonb,
  p_formula_overrides jsonb,
  p_total numeric,
  p_summary jsonb
)
returns void language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  p public.projects;
  t public.templates;
  next_rev int;
begin
  select * into p from public.projects where id = p_id;
  if p.id is null then raise exception 'Project not found.' using errcode='P0001'; end if;
  perform public._require_admin(uid, p.org_id);

  select * into t from public.templates where id = p_target_template_id and org_id = p.org_id;
  if t.id is null then raise exception 'Target template not found.' using errcode='P0001'; end if;

  update public.projects
  set template_id = t.id,
      template_version = t.version_num,
      template_snapshot = json_build_object('template_id', t.id, 'version_num', t.version_num, 'name', t.name, 'definition', t.definition),
      inputs = coalesce(p_inputs, p.inputs),
      formula_overrides = coalesce(p_formula_overrides, p.formula_overrides),
      total = p_total,
      summary = coalesce(p_summary, p.summary),
      updated_by = uid,
      updated_at = now()
  where id = p_id;

  select coalesce(max(revision_num), 0) + 1 into next_rev from public.project_revisions where project_id = p_id;
  insert into public.project_revisions (
    project_id, revision_num, name, template_version, inputs, formula_overrides, total, summary, note, saved_by
  ) values (
    p_id, next_rev, p.name, t.version_num, coalesce(p_inputs, p.inputs), coalesce(p_formula_overrides, p.formula_overrides),
    p_total, coalesce(p_summary, p.summary), 'Recalculated with template v' || t.version_num || ' (' || t.name || ')', uid
  );

  perform public._log_activity(p.org_id, p.id, uid, 'project_recalculated', json_build_object(
    'old_version', p.template_version,
    'new_version', t.version_num,
    'new_total', p_total
  ));
end $$;

create or replace function public.project_archive(p_token uuid, p_id uuid, p_archive boolean)
returns void language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  p public.projects;
begin
  select * into p from public.projects where id = p_id;
  if p.id is null then return; end if;
  perform public._require_member(uid, p.org_id);
  if not public._can_touch_project(uid, p.org_id, p) then
    raise exception 'Permission denied: Cannot archive a project you do not own or have access to.' using errcode='P0001';
  end if;

  update public.projects
  set is_archived = p_archive,
      status = case when p_archive then 'archived' else 'draft' end,
      updated_by = uid,
      updated_at = now()
  where id = p_id;

  perform public._log_activity(p.org_id, p.id, uid, case when p_archive then 'project_archived' else 'project_unarchived' end, json_build_object('name', p.name));
end $$;

create or replace function public.project_delete(p_token uuid, p_id uuid)
returns void language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  p public.projects;
begin
  select * into p from public.projects where id = p_id;
  if p.id is null then return; end if;
  perform public._require_admin(uid, p.org_id);

  delete from public.projects where id = p_id;
  perform public._log_activity(p.org_id, null, uid, 'project_deleted', json_build_object('project_id', p_id, 'name', p.name));
end $$;

create or replace function public.project_verify_quotation(
  p_project_id uuid,
  p_revision int default null,
  p_claimed_total numeric default null,
  p_signature text default null
)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  p public.projects;
  o public.organisations;
  rev public.project_revisions;
  actual_total numeric;
  actual_rev int;
  expected_sig text;
begin
  -- Strictly fail closed: signature and claimed total are required
  if p_signature is null or btrim(p_signature) = '' or p_claimed_total is null then
    return json_build_object(
      'verified', false,
      'status', 'invalid_request',
      'message', 'Verification requires both claimed quotation total and cryptographic digital seal.'
    );
  end if;

  select * into p from public.projects where id = p_project_id;
  if p.id is null then
    return json_build_object(
      'verified', false,
      'status', 'not_found',
      'message', 'No recorded quotation matches this identifier.'
    );
  end if;

  select * into o from public.organisations where id = p.org_id;

  if p_revision is not null then
    select * into rev from public.project_revisions where project_id = p.id and revision_num = p_revision;
    if rev.id is not null then
      actual_total := rev.total;
      actual_rev := rev.revision_num;
      expected_sig := rev.verification_signature;
    else
      actual_total := p.total;
      actual_rev := coalesce((select max(revision_num) from public.project_revisions where project_id = p.id), 1);
      expected_sig := p.verification_signature;
    end if;
  else
    actual_total := p.total;
    actual_rev := coalesce((select max(revision_num) from public.project_revisions where project_id = p.id), 1);
    expected_sig := p.verification_signature;
  end if;

  -- Validate signature and total
  if expected_sig is null or p_signature <> expected_sig or abs(actual_total - p_claimed_total) >= 0.01 then
    -- Zero information disclosure to unauthorized scanners
    return json_build_object(
      'verified', false,
      'status', 'tampered',
      'message', 'Quotation verification failed: digital signature or total mismatch. This document may be falsified.'
    );
  end if;

  return json_build_object(
    'verified', true,
    'status', 'authentic',
    'project_id', p.id,
    'project_name', p.name,
    'organisation', o.name,
    'revision', actual_rev,
    'verified_total', actual_total,
    'signature_verified', true,
    'updated_at', p.updated_at,
    'message', 'Official quotation verified authentic and tamper-free with cryptographic server seal.'
  );
end $$;

-- ---------------------------------------------------- 8. Activity & Dashboard
create or replace function public.log_activity_event(
  p_token uuid,
  p_org uuid,
  p_project uuid default null,
  p_action text default 'pdf_downloaded',
  p_details jsonb default '{}'::jsonb
)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
begin
  if not exists (select 1 from public.org_members where org_id = p_org and user_id = uid) then
    raise exception 'Access denied: not a member of this organisation.' using errcode='P0001';
  end if;

  perform public._log_activity(p_org, p_project, uid, p_action, p_details);
  return json_build_object('success', true);
end $$;

grant execute on function public.log_activity_event(uuid, uuid, uuid, text, jsonb) to anon, authenticated, service_role;

create or replace function public.activity_list(p_token uuid, p_org uuid, p_limit int default 50)
returns json language plpgsql stable security definer set search_path = public as $$
declare uid uuid := public._require_user(p_token);
begin
  perform public._require_member(uid, p_org);
  return coalesce((
    select json_agg(json_build_object(
      'id', a.id,
      'project_id', a.project_id,
      'user_id', a.user_id,
      'user_name', a.user_name,
      'action', a.action,
      'details', a.details,
      'created_at', a.created_at
    ) order by a.created_at desc)
    from (
      select * from public.activity_logs
      where org_id = p_org
      order by created_at desc
      limit coalesce(p_limit, 50)
    ) a
  ), '[]'::json);
end $$;

create or replace function public.dashboard_stats(p_token uuid, p_org uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  tot_projects int;
  act_projects int;
  pipeline_val numeric;
  cur_version int;
  recent_projects json;
  recent_activity json;
begin
  perform public._require_member(uid, p_org);

  select count(*),
         count(*) filter (where not is_archived),
         coalesce(sum(total) filter (where not is_archived), 0)
  into tot_projects, act_projects, pipeline_val
  from public.projects
  where org_id = p_org;

  select current_template_version into cur_version from public.organisations where id = p_org;

  select coalesce(json_agg(json_build_object(
    'id', p.id,
    'name', p.name,
    'status', p.status,
    'template_version', p.template_version,
    'total', p.total,
    'updated_at', p.updated_at
  ) order by p.updated_at desc), '[]'::json) into recent_projects
  from (
    select * from public.projects where org_id = p_org and not is_archived order by updated_at desc limit 5
  ) p;

  select coalesce(json_agg(json_build_object(
    'id', a.id,
    'user_name', a.user_name,
    'action', a.action,
    'details', a.details,
    'created_at', a.created_at
  ) order by a.created_at desc), '[]'::json) into recent_activity
  from (
    select * from public.activity_logs where org_id = p_org order by created_at desc limit 8
  ) a;

  return json_build_object(
    'total_projects', tot_projects,
    'active_projects', act_projects,
    'pipeline_value', pipeline_val,
    'current_template_version', cur_version,
    'recent_projects', recent_projects,
    'recent_activity', recent_activity
  );
end $$;

-- ------------------------------------------------ 9. Backward Compatibility
-- Keep legacy calc_* wrappers working seamlessly
create or replace function public.calc_list(p_token uuid, p_org uuid)
returns json language plpgsql stable security definer set search_path = public as $$
begin
  return public.project_list(p_token, p_org, false);
end $$;

create or replace function public.calc_save(p_token uuid, p_org uuid, p_name text, p_inputs jsonb, p_total numeric)
returns json language plpgsql security definer set search_path = public as $$
begin
  return public.project_create(p_token, p_org, p_name, null, p_inputs, '{}'::jsonb, p_total, '{}'::jsonb);
end $$;

create or replace function public.calc_update(p_token uuid, p_id uuid, p_name text, p_inputs jsonb, p_total numeric)
returns void language plpgsql security definer set search_path = public as $$
begin
  perform public.project_update(p_token, p_id, p_name, 'draft', p_inputs, '{}'::jsonb, p_total, '{}'::jsonb, 'Saved via legacy interface');
end $$;

create or replace function public.calc_delete(p_token uuid, p_id uuid)
returns void language plpgsql security definer set search_path = public as $$
begin
  perform public.project_delete(p_token, p_id);
end $$;

-- ------------------------------------------------ 10. Client Directory & Business Profile Management
create table if not exists public.clients (
  id               text primary key,
  org_id           uuid not null references public.organisations(id) on delete cascade,
  name             text not null,
  address          text not null default '',
  gst              text default '',
  delivery         text default '',
  kind_att         text default '',
  contact          text default '',
  email            text default '',
  subject          text default '',
  contacts         jsonb not null default '[]'::jsonb,
  extra_fields     jsonb not null default '[]'::jsonb,
  shared_with      jsonb not null default '[]'::jsonb,
  profit_margin_pct numeric default 25,
  created_by       text default '',
  created_by_role  text default '',
  created_at       timestamptz not null default now(),
  updated_at       timestamptz not null default now()
);
alter table public.clients add column if not exists shared_with jsonb not null default '[]'::jsonb;
alter table public.clients add column if not exists profit_margin_pct numeric default 25;

create index if not exists idx_clients_org_id on public.clients(org_id);
create index if not exists idx_clients_updated_at on public.clients(updated_at desc);

create table if not exists public.business_profiles (
  org_id                   uuid primary key references public.organisations(id) on delete cascade,
  company_name             text not null default 'MYRIAD STORAGE SYSTEM LLP',
  tagline                  text default 'Manufacturer of Mobile Compactor & Storage Systems',
  registered_address_line1 text not null default '702/A, AL-HUSSAIN, MOMIN NAGAR, JOGESHWARI (W)',
  registered_address_line2 text not null default 'MUMBAI - 400102. State : Maharashtra, Code : 27',
  contact_phone            text not null default '9769377043',
  alternate_phone          text default '',
  gst_no                   text not null default '27ABJFM5468N1ZG',
  pan_no                   text default 'ABJFM5468N',
  info_email               text not null default 'info@myriadstoragesystem.com',
  sales_email              text default 'sales@myriadstoragesystem.com',
  website                  text not null default 'www.myriadstoragesystem.com',
  factory_address_line1    text not null default 'Factory Address : Gala No 5, Vakan Compound, Near Dynamic Co, Vakan Pada Raod, Vasai Phata, Vasai East Dis.,',
  factory_address_line2    text default 'Palghar 401208.',
  signatory_title          text default 'Authorized Signatory',
  logo_url                 text,
  watermark_url            text,
  show_watermark           boolean not null default true,
  updated_at               timestamptz not null default now(),
  updated_by               text default 'Admin'
);

alter table public.clients enable row level security;
alter table public.business_profiles enable row level security;

create or replace function public.client_list(p_token uuid, p_org uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
begin
  perform public._require_member(uid, p_org);
  return coalesce((
    select json_agg(to_jsonb(c) order by c.updated_at desc)
    from public.clients c
    where c.org_id = p_org
  ), '[]'::json);
end $$;

create or replace function public.client_upsert(p_token uuid, p_org uuid, p_client jsonb)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  u_role text := public._require_member(uid, p_org);
  u_name text;
  c_id text := coalesce(p_client->>'id', 'client_' || gen_random_uuid()::text);
  res public.clients;
  existing public.clients;
  existing_dup public.clients;
  in_name text := btrim(coalesce(p_client->>'name', ''));
  in_gst text := btrim(coalesce(p_client->>'gst', ''));
begin
  -- 1. Role check: Viewers cannot create or edit clients
  if u_role in ('viewer', 'floor_operator') then
    raise exception 'Viewers do not have permission to manage client records.' using errcode = 'P0001';
  end if;

  select username into u_name from public.app_users where id = uid;

  -- 2. Tenant isolation check: Ensure existing client belongs to this organisation
  select * into existing from public.clients where id = c_id;
  if existing.id is not null then
    if existing.org_id <> p_org then
      raise exception 'Cross-tenant violation: Client belongs to another organisation.' using errcode = 'P0001';
    end if;
    if u_role not in ('owner', 'admin') and existing.created_by <> u_name and existing.created_by <> uid::text then
      if not (existing.shared_with ? u_name or existing.shared_with ? uid::text or existing.shared_with ? 'all') then
        raise exception 'Permission denied: You can only edit clients created by or shared with you.' using errcode = 'P0001';
      end if;
    end if;
  end if;

  -- 3. Collision check: If another client with same Name or GST already exists in this organisation
  if in_name <> '' or in_gst <> '' then
    select * into existing_dup from public.clients
    where org_id = p_org
      and id <> c_id
      and (
        (in_gst <> '' and lower(btrim(gst)) = lower(in_gst))
        or
        (in_name <> '' and lower(btrim(name)) = lower(in_name))
      )
    limit 1;

    if existing_dup.id is not null then
      -- If existing client is owned by someone else and caller is not owner/admin:
      if u_role not in ('owner', 'admin') and existing_dup.created_by <> u_name and existing_dup.created_by <> uid::text then
        if not (existing_dup.shared_with ? u_name or existing_dup.shared_with ? uid::text or existing_dup.shared_with ? 'all') then
          -- Log incident for Owner & Admin audit log
          perform public._log_activity(
            p_org,
            null,
            uid,
            'client_conflict',
            jsonb_build_object(
              'client_name', existing_dup.name,
              'gst', coalesce(existing_dup.gst, ''),
              'existing_owner', existing_dup.created_by,
              'attempted_by', u_name
            )
          );
          raise exception 'Client already exists' using errcode = 'P0001';
        end if;
      end if;
    end if;
  end if;

  -- 4. Upsert with explicit WHERE clients.org_id = p_org and trusted author and trusted author
  insert into public.clients (
    id, org_id, name, address, gst, delivery, kind_att, contact, email, subject,
    contacts, extra_fields, shared_with, profit_margin_pct, created_by, created_by_role, updated_at
  ) values (
    c_id, p_org,
    coalesce(btrim(p_client->>'name'), 'Unnamed Client'),
    coalesce(p_client->>'address', ''),
    coalesce(p_client->>'gst', ''),
    coalesce(p_client->>'delivery', ''),
    coalesce(p_client->>'kind_att', ''),
    coalesce(p_client->>'contact', ''),
    coalesce(p_client->>'email', ''),
    coalesce(p_client->>'subject', ''),
    coalesce(p_client->'contacts', '[]'::jsonb),
    coalesce(p_client->'extra_fields', '[]'::jsonb),
    case when u_role in ('owner', 'admin') then coalesce(p_client->'shared_with', '[]'::jsonb) else coalesce(existing.shared_with, '[]'::jsonb) end,
    case when u_role in ('owner', 'admin') then coalesce((p_client->>'profit_margin_pct')::numeric, 25) else coalesce(existing.profit_margin_pct, 25) end,
    coalesce(existing.created_by, u_name),
    coalesce(existing.created_by_role, u_role),
    now()
  )
  on conflict (id) do update set
    name = excluded.name,
    address = excluded.address,
    gst = excluded.gst,
    delivery = excluded.delivery,
    kind_att = excluded.kind_att,
    contact = excluded.contact,
    email = excluded.email,
    subject = excluded.subject,
    contacts = excluded.contacts,
    extra_fields = excluded.extra_fields,
    shared_with = excluded.shared_with,
    profit_margin_pct = excluded.profit_margin_pct,
    updated_at = now()
  where public.clients.org_id = p_org
  returning * into res;

  if res.id is null then
    raise exception 'Failed to update client: organisation mismatch.' using errcode = 'P0001';
  end if;

  return to_json(res);
end $$;

create or replace function public.client_delete(p_token uuid, p_org uuid, p_client_id text)
returns void language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
begin
  perform public._require_admin(uid, p_org);
  delete from public.clients where id = p_client_id and org_id = p_org;
end $$;

create or replace function public.business_profile_get(p_token uuid, p_org uuid)
returns json language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  b public.business_profiles;
begin
  perform public._require_member(uid, p_org);
  select * into b from public.business_profiles where org_id = p_org;
  if b.org_id is null then
    insert into public.business_profiles (org_id) values (p_org) on conflict (org_id) do nothing;
    select * into b from public.business_profiles where org_id = p_org;
  end if;
  return to_json(b);
end $$;

create or replace function public.business_profile_save(p_token uuid, p_org uuid, p_profile jsonb)
returns json language plpgsql security definer set search_path = public as $$
declare
  uid uuid := public._require_user(p_token);
  res public.business_profiles;
begin
  perform public._require_admin(uid, p_org);
  insert into public.business_profiles (
    org_id, company_name, tagline, registered_address_line1, registered_address_line2,
    contact_phone, alternate_phone, gst_no, pan_no, info_email, sales_email, website,
    factory_address_line1, factory_address_line2, signatory_title, logo_url, watermark_url, show_watermark, updated_at, updated_by
  ) values (
    p_org,
    coalesce(p_profile->>'company_name', 'MYRIAD STORAGE SYSTEM LLP'),
    coalesce(p_profile->>'tagline', ''),
    coalesce(p_profile->>'registered_address_line1', ''),
    coalesce(p_profile->>'registered_address_line2', ''),
    coalesce(p_profile->>'contact_phone', ''),
    coalesce(p_profile->>'alternate_phone', ''),
    coalesce(p_profile->>'gst_no', ''),
    coalesce(p_profile->>'pan_no', ''),
    coalesce(p_profile->>'info_email', ''),
    coalesce(p_profile->>'sales_email', ''),
    coalesce(p_profile->>'website', ''),
    coalesce(p_profile->>'factory_address_line1', ''),
    coalesce(p_profile->>'factory_address_line2', ''),
    coalesce(p_profile->>'signatory_title', 'Authorized Signatory'),
    p_profile->>'logo_url',
    p_profile->>'watermark_url',
    coalesce((p_profile->>'show_watermark')::boolean, true),
    now(),
    (select username from public.app_users where id = uid)
  )
  on conflict (org_id) do update set
    company_name = excluded.company_name,
    tagline = excluded.tagline,
    registered_address_line1 = excluded.registered_address_line1,
    registered_address_line2 = excluded.registered_address_line2,
    contact_phone = excluded.contact_phone,
    alternate_phone = excluded.alternate_phone,
    gst_no = excluded.gst_no,
    pan_no = excluded.pan_no,
    info_email = excluded.info_email,
    sales_email = excluded.sales_email,
    website = excluded.website,
    factory_address_line1 = excluded.factory_address_line1,
    factory_address_line2 = excluded.factory_address_line2,
    signatory_title = excluded.signatory_title,
    logo_url = excluded.logo_url,
    watermark_url = excluded.watermark_url,
    show_watermark = excluded.show_watermark,
    updated_at = now(),
    updated_by = excluded.updated_by
  returning * into res;

  return to_json(res);
end $$;

-- ------------------------------------------------ 11. Security Triggers & Hardening
-- Protect published template versions against modification or deletion
create or replace function public._protect_published_templates()
returns trigger language plpgsql as $$
begin
  if tg_op = 'DELETE' and old.is_default then
    raise exception 'Cannot delete active default template version. Please switch the default to another version first.' using errcode='P0001';
  end if;
  if tg_op = 'UPDATE' and old.status = 'published' then
    if new.definition <> old.definition or new.version_num <> old.version_num or new.name <> old.name then
      raise exception 'Published template versions are immutable and cannot be altered. Publish a new version instead.' using errcode='P0001';
    end if;
  end if;
  return new;
end $$;

drop trigger if exists trg_protect_published_templates on public.templates;
create trigger trg_protect_published_templates
before update or delete on public.templates
for each row execute function public._protect_published_templates();

-- Protect activity logs from client-side direct tampering
create or replace function public._protect_activity_logs()
returns trigger language plpgsql as $$
begin
  if session_user in ('anon', 'authenticated') then
    raise exception 'Direct modification of audit logs is forbidden.' using errcode='P0001';
  end if;
  return new;
end $$;

drop trigger if exists trg_protect_activity_logs on public.activity_logs;
create trigger trg_protect_activity_logs
before insert or update or delete on public.activity_logs
for each row execute function public._protect_activity_logs();

-- Schema CHECK constraints for JSONB integrity
do $$ begin
  alter table public.projects drop constraint if exists projects_inputs_check;
  alter table public.projects add constraint projects_inputs_check check (jsonb_typeof(inputs) = 'object');

  alter table public.projects drop constraint if exists projects_formulas_check;
  alter table public.projects add constraint projects_formulas_check check (jsonb_typeof(formula_overrides) = 'object');

  alter table public.templates drop constraint if exists templates_definition_check;
  alter table public.templates add constraint templates_definition_check check (jsonb_typeof(definition) = 'object');
exception when others then null;
end $$;

-- ---------------------------------------------------------- 12. Permissions
revoke all on all functions in schema public from public, anon, authenticated;
grant execute on function
  public.app_signup(text, text, text), public.app_login(text, text), public.app_me(uuid), public.app_logout(uuid),
  public.app_update_profile(uuid, text, text), public.app_update_profile(uuid, text, text, text, text),
  public.org_join(uuid, text), public.org_list(uuid),
  public.org_members_list(uuid, uuid), public.org_update_member_role(uuid, uuid, uuid, text),
  public.org_remove_member(uuid, uuid, uuid), public.org_leave(uuid, uuid),
  public.org_join_requests_list(uuid, uuid),
  public.org_join_request_approve(uuid, uuid, uuid, text),
  public.org_join_request_reject(uuid, uuid, uuid),
  public.org_rotate_code(uuid, uuid),
  public.org_invite_member(uuid, uuid, text, text),
  public.org_accept_invite(uuid, text),
  public.template_list(uuid, uuid), public.template_save_draft(uuid, uuid, text, text, jsonb, uuid),
  public.template_publish(uuid, uuid, text, text, jsonb, uuid),
  public.template_set_default(uuid, uuid, uuid),
  public.project_list(uuid, uuid, boolean), public.project_get(uuid, uuid),
  public.project_create(uuid, uuid, text, uuid, jsonb, jsonb, numeric, jsonb),
  public.project_update(uuid, uuid, text, text, jsonb, jsonb, numeric, jsonb, text),
  public.project_duplicate(uuid, uuid, text, boolean),
  public.project_recalculate_template(uuid, uuid, uuid, jsonb, jsonb, numeric, jsonb),
  public.project_archive(uuid, uuid, boolean), public.project_delete(uuid, uuid),
  public.project_verify_quotation(uuid, int, numeric, text),
  public.activity_list(uuid, uuid, int), public.dashboard_stats(uuid, uuid),
  public.calc_list(uuid, uuid), public.calc_save(uuid, uuid, text, jsonb, numeric),
  public.calc_update(uuid, uuid, text, jsonb, numeric), public.calc_delete(uuid, uuid),
  public.client_list(uuid, uuid), public.client_upsert(uuid, uuid, jsonb), public.client_delete(uuid, uuid, text),
  public.business_profile_get(uuid, uuid), public.business_profile_save(uuid, uuid, jsonb)
to anon, authenticated;

-- Secure helper to create/setup Owner and Organisation directly in SQL Editor
-- (Callable only by database superuser/postgres in SQL Editor, not exposed via API)
create or replace function public.admin_setup_owner_and_org(
  p_owner_username text,
  p_owner_email text,
  p_owner_password text,
  p_org_name text
)
returns json language plpgsql security definer set search_path = public, extensions as $$
declare
  v_user public.app_users;
  v_org public.organisations;
begin
  -- 1. Create or update owner user
  select * into v_user from public.app_users where lower(username) = lower(btrim(p_owner_username));
  if v_user.id is not null then
    update public.app_users
    set email = lower(btrim(p_owner_email)),
        password_hash = extensions.crypt(p_owner_password, extensions.gen_salt('bf', 10))
    where id = v_user.id
    returning * into v_user;
  else
    insert into public.app_users (username, email, password_hash)
    values (
      btrim(p_owner_username),
      lower(btrim(p_owner_email)),
      extensions.crypt(p_owner_password, extensions.gen_salt('bf', 10))
    )
    returning * into v_user;
  end if;

  insert into public.profiles (id, username)
  values (v_user.id, v_user.username)
  on conflict (id) do update set username = excluded.username;

  -- 2. Create organisation with unique code & default version
  insert into public.organisations (name, code, current_template_version, created_by)
  values (btrim(p_org_name), public._new_code(), 1, v_user.id)
  returning * into v_org;

  -- 3. Add as owner
  insert into public.org_members (org_id, user_id, role)
  values (v_org.id, v_user.id, 'owner')
  on conflict (org_id, user_id) do update set role = 'owner';

  -- 4. Initialize Default Production Template V1 for this organisation
  insert into public.templates (org_id, version_num, name, description, status, is_default, definition, created_by, published_by, published_at)
  values (
    v_org.id,
    1,
    'Default Production Template V1',
    'Standard calculation template for ' || btrim(p_org_name),
    'published',
    true,
    '{"version": 1, "rates": {"steel": 90, "trolleySteel": 90, "bodyLabour": 250}, "dimensions": {"rows": 4, "cols": 5, "ht": 2200, "w": 1830, "d": 610}}'::jsonb,
    v_user.id,
    v_user.id,
    now()
  ) on conflict do nothing;

  return json_build_object(
    'status', 'success',
    'user_id', v_user.id,
    'username', v_user.username,
    'email', v_user.email,
    'org_id', v_org.id,
    'org_name', v_org.name,
    'role', 'owner'
  );
end $$;

-- Strictly revoke administrative & sensitive functions from client invocation
revoke execute on function public.admin_setup_owner_and_org(text, text, text, text) from anon, authenticated, public;
grant execute on function public.admin_setup_owner_and_org(text, text, text, text) to postgres, service_role;
revoke execute on function public.org_create(uuid, text) from anon, authenticated, public;
revoke execute on function public.admin_create_organisation(text, text) from anon, authenticated, public;
revoke execute on function public._get_or_create_hmac_key() from anon, authenticated, public;
revoke execute on function public.check_rate_limit(text, int, int) from anon, authenticated, public;
grant execute on function public.org_create(uuid, text) to service_role;
grant execute on function public.admin_create_organisation(text, text) to service_role, postgres;
grant execute on function public._get_or_create_hmac_key() to service_role, postgres;
grant execute on function public.check_rate_limit(text, int, int) to service_role;

-- Hardened Customer Data Protection: Zero direct table access for anon or authenticated
revoke all on public.clients from anon, authenticated, public;
revoke all on public.business_profiles from anon, authenticated, public;
alter table public.clients enable row level security;
alter table public.business_profiles enable row level security;
drop policy if exists clients_policy on public.clients;
drop policy if exists business_profiles_policy on public.business_profiles;

-- Refresh PostgREST schema cache
notify pgrst, 'reload schema';

