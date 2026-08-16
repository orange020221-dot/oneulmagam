# 오늘마감

CHS7006 실습 2·3 산출물. 오늘과 이번 주 마감만 보는 스케이트보드 MVP.

## 로컬

`index.html`을 브라우저로 연다. 서버 없이 화면은 동작한다.

배포된 주소에서는 `/api/event`로 KPI가 쌓이고, GA가 있으면 방문·클릭도 잡힌다.

## 운영

- 배포: Vercel
- DB: Supabase `kpi_events` (익명 저장만)
- 측정: Google Analytics + `today_cleared` / `share_click`
- 검색: `robots.txt`, `sitemap.xml`

자세한 가설은 `실습2_기획.md`, 배포 절차는 `실습3_운영.md`.
