-- 오늘마감 KPI 테이블
-- supabase.com > SQL Editor 에 붙여넣고 Run

create table if not exists public.kpi_events (
  id bigint generated always as identity primary key,
  event_name text not null,
  item_type text,
  used_at timestamptz not null default now()
);

alter table public.kpi_events enable row level security;

drop policy if exists "anon_insert_only" on public.kpi_events;
create policy "anon_insert_only"
  on public.kpi_events
  for insert
  to anon
  with check (
    event_name in ('visit', 'add_item', 'complete_item', 'today_cleared', 'share_click', 'import_feed')
  );

-- 익명은 저장만 가능. 조회·수정·삭제 불가.
-- anon 키가 페이지에 노출돼도 데이터를 읽어갈 수 없다.
