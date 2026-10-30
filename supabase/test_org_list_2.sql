with sess as (
  insert into public.app_sessions (token_hash, user_id, expires_at)
  values (encode(sha256('a0000000-0000-0000-0000-000000000001'::bytea), 'hex'), 'f6419c0a-9727-4650-8411-da5662aa0a49', now() + interval '12 hours')
  returning token_hash
)
select public.org_list('a0000000-0000-0000-0000-000000000001'::uuid);
