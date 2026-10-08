-- v14: close the anon REST door. Run ONLY after SUPABASE_SERVICE_ROLE is set on Render
-- and /api/db/health shows "service": true (the server then bypasses RLS with the service key).
-- v15 note (2026-10-08 23:45 IST): live already shows RLS on most tables (anon count 0), but public.bills is still readable with the anon key.
-- No policies are created for anon/authenticated: all reads/writes go through server.js.
alter table public.site_measures enable row level security;
alter table public.bills         enable row level security;
alter table public.kachu_bills   enable row level security;
alter table public.estimates     enable row level security;
alter table public.profiles      enable row level security;
alter table public.tour_days     enable row level security;
alter table public.tour_profile  enable row level security;
alter table public.media         enable row level security;
revoke all on public.site_measures, public.bills, public.kachu_bills, public.estimates,
              public.profiles, public.tour_days, public.tour_profile, public.media
  from anon, authenticated;
-- Optional, if a client ever reads its own profile directly:
-- create policy "own profile" on public.profiles for select to authenticated
--   using (lower(email) = lower(auth.jwt() ->> 'email'));
-- Check (should return 401/permission denied or an empty list with the anon key):
--   curl "$SUPABASE_URL/rest/v1/profiles?select=email&limit=1" -H "apikey: $ANON" -H "Authorization: Bearer $ANON"
