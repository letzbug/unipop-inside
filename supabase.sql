-- UniPop Inside 1.1.0 – installation Supabase complète
-- À exécuter dans Supabase > SQL Editor > New query

create extension if not exists pgcrypto;

create table if not exists public.imports (
  id uuid primary key default gen_random_uuid(),
  original_filename text not null,
  generated_filename text,
  generated_file_path text,
  school_year text,
  course_count integer default 0,
  sheet_count integer default 0,
  is_active boolean default false,
  created_at timestamptz default now()
);

alter table public.imports
  add column if not exists generated_file_path text;

create unique index if not exists only_one_active_import
on public.imports (is_active)
where is_active = true;

create table if not exists public.courses (
  id uuid primary key default gen_random_uuid(),
  import_id uuid not null references public.imports(id) on delete cascade,
  source_sheet text,
  source_row integer,
  course_id text,
  school_year text,
  title text,
  level text,
  start_date date,
  end_date date,
  total_duration text,
  schedule text,
  places text,
  description text,
  additional_info text,
  location_name text,
  location_room text,
  trainer text,
  category text,
  subject text,
  link text,
  qr_data text,
  created_at timestamptz default now()
);

create index if not exists courses_import_idx on public.courses(import_id);
create index if not exists courses_search_idx on public.courses
using gin(to_tsvector('simple', coalesce(title,'') || ' ' || coalesce(course_id,'') || ' ' || coalesce(location_name,'')));

create table if not exists public.modifications (
  id uuid primary key default gen_random_uuid(),
  import_id uuid not null references public.imports(id) on delete cascade,
  course_id_ref uuid references public.courses(id) on delete cascade,
  course_code text,
  course_title text,
  trainer_name text not null,
  trainer_email text,
  field_key text not null,
  field_label text not null,
  original_value text,
  proposed_value text,
  created_at timestamptz default now()
);

create index if not exists modifications_import_idx on public.modifications(import_id);

alter table public.imports enable row level security;
alter table public.courses enable row level security;
alter table public.modifications enable row level security;

drop policy if exists "inside imports read" on public.imports;
create policy "inside imports read" on public.imports for select to anon using (true);
drop policy if exists "inside imports write" on public.imports;
create policy "inside imports write" on public.imports for all to anon using (true) with check (true);

drop policy if exists "inside courses read" on public.courses;
create policy "inside courses read" on public.courses for select to anon using (true);
drop policy if exists "inside courses insert" on public.courses;
create policy "inside courses insert" on public.courses for insert to anon with check (true);
drop policy if exists "inside courses write" on public.courses;

drop policy if exists "inside modifications read" on public.modifications;
create policy "inside modifications read" on public.modifications for select to anon using (true);
drop policy if exists "inside modifications insert" on public.modifications;
create policy "inside modifications insert" on public.modifications for insert to anon with check (true);
drop policy if exists "inside modifications write" on public.modifications;

-- Bucket privé pour les fichiers Excel générés.
insert into storage.buckets (id, name, public)
values ('unipop-files', 'unipop-files', false)
on conflict (id) do update set public = false;

drop policy if exists "unipop files read" on storage.objects;
create policy "unipop files read"
on storage.objects for select to anon
using (bucket_id = 'unipop-files');

drop policy if exists "unipop files insert" on storage.objects;
create policy "unipop files insert"
on storage.objects for insert to anon
with check (bucket_id = 'unipop-files');

drop policy if exists "unipop files update" on storage.objects;
create policy "unipop files update"
on storage.objects for update to anon
using (bucket_id = 'unipop-files')
with check (bucket_id = 'unipop-files');
