-- 003_private_media.sql  (SO+ step 3b, 9 Oct 2026)
-- Private, per-user site photos/sketches. Safe to run more than once (idempotent). Deletes NO data.
-- Run in Supabase > SQL Editor BEFORE the step-3b code goes live.
-- Bucket "site-media" is NOT changed: it still serves generated Excel/PDF (out/) and bill files (bills/) by public link.

-- 1) Private bucket for photos/sketches (max 8 MB, images only)
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('user-media', 'user-media', false, 8388608, array['image/jpeg','image/png','image/webp'])
on conflict (id) do update
  set public = false,
      file_size_limit = excluded.file_size_limit,
      allowed_mime_types = excluded.allowed_mime_types;

-- 2) Storage RLS: a logged-in user may touch ONLY objects under "<own user id>/..."
--    (the server uses the service role, which bypasses RLS; these policies are the safety net
--     if the browser ever talks to Storage directly)
drop policy if exists "user-media own select" on storage.objects;
create policy "user-media own select" on storage.objects for select to authenticated
  using (bucket_id = 'user-media' and (storage.foldername(name))[1] = auth.uid()::text);
drop policy if exists "user-media own insert" on storage.objects;
create policy "user-media own insert" on storage.objects for insert to authenticated
  with check (bucket_id = 'user-media' and (storage.foldername(name))[1] = auth.uid()::text);
drop policy if exists "user-media own update" on storage.objects;
create policy "user-media own update" on storage.objects for update to authenticated
  using (bucket_id = 'user-media' and (storage.foldername(name))[1] = auth.uid()::text)
  with check (bucket_id = 'user-media' and (storage.foldername(name))[1] = auth.uid()::text);
drop policy if exists "user-media own delete" on storage.objects;
create policy "user-media own delete" on storage.objects for delete to authenticated
  using (bucket_id = 'user-media' and (storage.foldername(name))[1] = auth.uid()::text);

-- 3) media table (already exists on live; created only if missing) + new columns
create table if not exists public.media (
  id text primary key default gen_random_uuid()::text,
  ts timestamptz not null default now(),
  kind text, work_name text, gps text, url text, user_email text
);
alter table public.media add column if not exists user_id uuid;
alter table public.media add column if not exists path text;
alter table public.media add column if not exists bucket text;
create index if not exists media_user_ts_idx on public.media (user_id, ts desc);

-- 4) media RLS: owner rows only. anon/authenticated get NO table grants (browser never reads it directly;
--    the server reads with the service role and filters by the verified user id).
alter table public.media enable row level security;
revoke all on public.media from anon, authenticated;
drop policy if exists "media own rows" on public.media;
create policy "media own rows" on public.media for all to authenticated
  using (user_id = auth.uid()) with check (user_id = auth.uid());

-- 5) Checks (read-only). Expected: user-media public = false; site-media public = true.
select id, public, file_size_limit from storage.buckets where id in ('site-media', 'user-media');
select policyname from pg_policies where schemaname = 'storage' and tablename = 'objects' and policyname like 'user-media%';
select column_name, data_type from information_schema.columns where table_schema = 'public' and table_name = 'media' order by ordinal_position;

-- 6) Old photos/sketches saved before 3b (public, in site-media root, rows without path). NOT deleted.
--    Review only; migrate later with an owner-run script (copy to user-media/<user_id>/..., then decide).
select name, created_at from storage.objects
 where bucket_id = 'site-media' and name !~ '^(out|bills)/' order by created_at desc;
select count(*) filter (where path is null) as legacy_rows, count(*) filter (where path is not null) as private_rows from public.media;
