insert into public.app_sessions (token_hash, user_id, expires_at)
values ('4e18f9168ad3f1551865a9a48635d2c4376d464cb78cbce22f8ff90931140ec7', 'f6419c0a-9727-4650-8411-da5662aa0a49', now() + interval '12 hours');

select public.org_list('a0000000-0000-0000-0000-000000000001'::uuid);
