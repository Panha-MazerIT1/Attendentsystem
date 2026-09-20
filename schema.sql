-- =====================================================================
--  ប្រព័ន្ធគ្រប់គ្រងវត្តមាន · Supabase schema
--  ដំណើរការ៖ Supabase → SQL Editor → New query → Paste → Run
--  អាច Run ម្តងទៀតបានដោយសុវត្ថិភាព (មិនលុបទិន្នន័យ)
-- =====================================================================

-- 1) ការកំណត់ (ក្បាលអត្ថបទ, ចំណងជើងថ្នាក់, កាលវិភាគគំរូ ...) · មួយជួរក្នុងមួយគណនី
create table if not exists public.settings (
  user_id    uuid primary key default auth.uid() references auth.users (id) on delete cascade,
  data       jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

-- 2) បញ្ជីសិស្ស
create table if not exists public.students (
  id         uuid primary key default gen_random_uuid(),
  user_id    uuid not null default auth.uid() references auth.users (id) on delete cascade,
  seq        integer not null,                 -- លេខរៀង (ល.រ)
  name       text not null,
  sex        text,                             -- ប្រុស / ស្រី
  dob        text,                             -- ថ្ងៃខែឆ្នាំកំណើត (អត្ថបទ)
  active     boolean not null default true,    -- false = លាក់ (មិនបង្ហាញក្នុងរបាយការណ៍)
  created_at timestamptz not null default now()
);
create index if not exists students_user_seq on public.students (user_id, seq);

-- 3) សប្ដាហ៍ (ចាប់ផ្ដើមថ្ងៃច័ន្ទ) + កាលវិភាគរបស់សប្ដាហ៍នោះ
create table if not exists public.weeks (
  id         uuid primary key default gen_random_uuid(),
  user_id    uuid not null default auth.uid() references auth.users (id) on delete cascade,
  start_date date not null,
  timetable  jsonb not null default '{}'::jsonb,
  unique (user_id, start_date)
);

-- 4) អវត្តមាន (រក្សាទុកតែអ្នកអវត្តមាន · គ្មានជួរ = មក)
--    status: 1 = មានច្បាប់ · 0 = អត់ច្បាប់   (ដូចលេខក្នុង Excel ចាស់)
create table if not exists public.absences (
  user_id    uuid not null default auth.uid() references auth.users (id) on delete cascade,
  week_id    uuid not null references public.weeks (id) on delete cascade,
  student_id uuid not null references public.students (id) on delete cascade,
  day        smallint not null check (day between 0 and 4),          -- 0=ច័ន្ទ … 4=សុក្រ
  session    text     not null check (session in ('am', 'pm')),
  period     smallint not null check (period between 0 and 11),      -- លេខម៉ោងក្នុងពេលនោះ (ចាប់ពី 0)
  status     smallint not null check (status in (0, 1)),
  primary key (week_id, student_id, day, session, period)
);
create index if not exists absences_user on public.absences (user_id);

-- 5) សុវត្ថិភាព៖ គណនីនីមួយៗឃើញ/កែបានតែទិន្នន័យរបស់ខ្លួន
alter table public.settings enable row level security;
alter table public.students enable row level security;
alter table public.weeks    enable row level security;
alter table public.absences enable row level security;

drop policy if exists "own settings" on public.settings;
drop policy if exists "own students" on public.students;
drop policy if exists "own weeks"    on public.weeks;
drop policy if exists "own absences" on public.absences;

create policy "own settings" on public.settings for all to authenticated
  using ((select auth.uid()) = user_id) with check ((select auth.uid()) = user_id);
create policy "own students" on public.students for all to authenticated
  using ((select auth.uid()) = user_id) with check ((select auth.uid()) = user_id);
create policy "own weeks" on public.weeks for all to authenticated
  using ((select auth.uid()) = user_id) with check ((select auth.uid()) = user_id);
create policy "own absences" on public.absences for all to authenticated
  using ((select auth.uid()) = user_id) with check ((select auth.uid()) = user_id);

-- 6) សិទ្ធិចូលប្រើ៖ មានតែអ្នកដែល login (authenticated) ប៉ុណ្ណោះ · អ្នកមិនទាន់ login (anon) មិនបាន
grant usage on schema public to authenticated;
grant select, insert, update, delete on public.settings, public.students, public.weeks, public.absences to authenticated;
revoke all on public.settings, public.students, public.weeks, public.absences from anon;
