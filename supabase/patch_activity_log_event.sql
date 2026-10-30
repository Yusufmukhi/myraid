-- Create RPC function to log user activity events (like saving/downloading PDF)
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
