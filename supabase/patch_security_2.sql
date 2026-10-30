-- =============================================================================
-- patch_security_2.sql: Production Security Hardening Patch for Live Database
-- Safe to run on live data. Does NOT drop or recreate tables.
-- =============================================================================

begin;

-- 1. Ensure Row Level Security (RLS) is strictly enabled across all tables
alter table if exists public.app_users            enable row level security;
alter table if exists public.app_sessions         enable row level security;
alter table if exists public.organisations        enable row level security;
alter table if exists public.org_members          enable row level security;
alter table if exists public.templates            enable row level security;
alter table if exists public.projects             enable row level security;
alter table if exists public.project_revisions    enable row level security;
alter table if exists public.activity_logs        enable row level security;
alter table if exists public.org_join_requests    enable row level security;
alter table if exists public.invitations          enable row level security;
alter table if exists public.formula_secrets      enable row level security;
alter table if exists public.rate_limits          enable row level security;
alter table if exists public.internal_system_keys enable row level security;
alter table if exists public.clients              enable row level security;
alter table if exists public.business_profiles    enable row level security;

-- 2. Revoke direct table manipulation privileges from anon and authenticated roles.
-- All database operations must go through SECURITY DEFINER RPC functions or Edge Functions (service_role).
revoke all on public.formula_secrets      from anon, authenticated;
revoke all on public.internal_system_keys from anon, authenticated;
revoke all on public.rate_limits          from anon, authenticated;
revoke all on public.app_sessions         from anon, authenticated;
revoke all on public.app_users            from anon, authenticated;
revoke all on public.org_members          from anon, authenticated;
revoke all on public.organisations        from anon, authenticated;
revoke all on public.templates            from anon, authenticated;
revoke all on public.projects             from anon, authenticated;
revoke all on public.project_revisions    from anon, authenticated;
revoke all on public.activity_logs        from anon, authenticated;
revoke all on public.invitations          from anon, authenticated;
revoke all on public.org_join_requests    from anon, authenticated;

-- Ensure service_role maintains full administration access for Edge Functions
grant all on all tables in schema public to service_role;
grant all on all sequences in schema public to service_role;
grant all on all routines in schema public to service_role;

-- 3. Ensure RPC functions can be executed by clients (functions enforce authentication internally)
grant execute on all routines in schema public to anon, authenticated;

-- 4. Clean up any expired sessions or stale rate limits
delete from public.app_sessions where expires_at <= now();
delete from public.rate_limits where window_start < now() - interval '1 hour';

commit;
