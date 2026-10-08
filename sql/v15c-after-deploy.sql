-- v15c: run in Supabase SQL editor AFTER Render has ADMIN_EMAIL + SUPABASE_SERVICE_ROLE and the new code is live.
-- 1) Old rows saved before v15c have user_email = NULL, so the per-user lists hide them.
--    Replace OWNER_LOGIN_EMAIL with the e-mail that should own the old rows (normally the owner's login).
--    Review first:  select 'site_measures' t, count(*) from public.site_measures where user_email is null
--                   union all select 'bills', count(*) from public.bills where user_email is null
--                   union all select 'estimates', count(*) from public.estimates where user_email is null
--                   union all select 'kachu_bills', count(*) from public.kachu_bills where user_email is null
--                   union all select 'media', count(*) from public.media where user_email is null;
update public.site_measures set user_email = lower('OWNER_LOGIN_EMAIL') where user_email is null;
update public.bills         set user_email = lower('OWNER_LOGIN_EMAIL') where user_email is null;
update public.estimates     set user_email = lower('OWNER_LOGIN_EMAIL') where user_email is null;
update public.kachu_bills   set user_email = lower('OWNER_LOGIN_EMAIL') where user_email is null;
update public.media         set user_email = lower('OWNER_LOGIN_EMAIL') where user_email is null;
-- 2) Optional: users that the old /api/me/profile reset from Active to Trial on login.
--    update public.profiles set subscription_status = 'Active' where email in ('<user1>', '<user2>');
