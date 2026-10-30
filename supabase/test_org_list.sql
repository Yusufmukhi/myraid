do $$
declare
  t uuid := gen_random_uuid();
  t_hash text := encode(sha256(t::text::bytea), 'hex');
  res json;
begin
  insert into public.app_sessions (token_hash, user_id, expires_at)
  values (t_hash, 'f6419c0a-9727-4650-8411-da5662aa0a49', now() + interval '12 hours');

  res := public.org_list(t);
  raise notice 'ORG_LIST RESULT: %', res;
end $$;
