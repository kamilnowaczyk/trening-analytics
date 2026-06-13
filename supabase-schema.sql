-- =====================================================================
--  Trening Analytics — schemat bazy Supabase
--  Uruchom RAZ w panelu Supabase: SQL Editor → New query → wklej → Run
--  Daje: tabelę stanu + prywatny magazyn plików + RLS (każdy widzi tylko
--  swoje dane) + realtime (synchronizacja na żywo).
-- =====================================================================

-- 1) Tabela stanu workbooka — jeden wiersz na użytkownika -------------
create table if not exists public.workbook_state (
  user_id    uuid primary key references auth.users(id) on delete cascade,
  updated_at timestamptz not null default now(),
  file_name  text,
  week_count int,
  device     text
);

alter table public.workbook_state enable row level security;

drop policy if exists "own row select" on public.workbook_state;
drop policy if exists "own row insert" on public.workbook_state;
drop policy if exists "own row update" on public.workbook_state;
drop policy if exists "own row delete" on public.workbook_state;

create policy "own row select" on public.workbook_state
  for select using (auth.uid() = user_id);
create policy "own row insert" on public.workbook_state
  for insert with check (auth.uid() = user_id);
create policy "own row update" on public.workbook_state
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "own row delete" on public.workbook_state
  for delete using (auth.uid() = user_id);

-- realtime (synchronizacja na żywo)
do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'workbook_state'
  ) then
    alter publication supabase_realtime add table public.workbook_state;
  end if;
end $$;

-- 2) Prywatny magazyn plików .xlsx -----------------------------------
insert into storage.buckets (id, name, public)
values ('workbooks', 'workbooks', false)
on conflict (id) do nothing;

drop policy if exists "own files select" on storage.objects;
drop policy if exists "own files insert" on storage.objects;
drop policy if exists "own files update" on storage.objects;
drop policy if exists "own files delete" on storage.objects;

-- użytkownik zarządza WYŁĄCZNIE plikami w swoim folderze {user_id}/...
create policy "own files select" on storage.objects
  for select using (bucket_id = 'workbooks' and (storage.foldername(name))[1] = auth.uid()::text);
create policy "own files insert" on storage.objects
  for insert with check (bucket_id = 'workbooks' and (storage.foldername(name))[1] = auth.uid()::text);
create policy "own files update" on storage.objects
  for update using (bucket_id = 'workbooks' and (storage.foldername(name))[1] = auth.uid()::text);
create policy "own files delete" on storage.objects
  for delete using (bucket_id = 'workbooks' and (storage.foldername(name))[1] = auth.uid()::text);
