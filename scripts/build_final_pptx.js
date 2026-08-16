const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.defineLayout({ name: "WIDE16", width: 13.333, height: 7.5 });
pres.layout = "WIDE16";
pres.author = "이재훈";
pres.title = "오늘마감 기말 영상 초안";
pres.subject = "CHS7006 · 필수 항목 8장";

const INK = "1B1814";
const PAPER = "F3EEE4";
const CARD = "FFFCF6";
const MUTED = "6F675C";
const ACCENT = "B4451B";
const DONE = "2C5A4A";
const CREAM = "FFF8EE";

function shadow() {
  return { type: "outer", color: "1B1814", blur: 10, offset: 3, angle: 135, opacity: 0.08 };
}

function label(slide, text) {
  slide.addText(text, {
    x: 0.7, y: 0.38, w: 8, h: 0.28,
    fontFace: "Calibri", fontSize: 12, color: ACCENT, bold: true,
    charSpacing: 1.2, margin: 0
  });
}

function h1(slide, text, y) {
  slide.addText(text, {
    x: 0.7, y: y || 0.72, w: 12, h: 0.7,
    fontFace: "Georgia", fontSize: 30, color: INK, bold: true, margin: 0
  });
}

function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: fill || CARD },
    rectRadius: 0.08,
    shadow: shadow()
  });
}

function footer(slide, n) {
  slide.addText("오늘마감  ·  oneulmagam.vercel.app", {
    x: 0.7, y: 7.08, w: 8, h: 0.22,
    fontFace: "Calibri", fontSize: 11, color: MUTED, margin: 0
  });
  slide.addText(String(n) + " / 8", {
    x: 11.4, y: 7.08, w: 1.2, h: 0.22,
    fontFace: "Calibri", fontSize: 11, color: MUTED, align: "right", margin: 0
  });
}

// 1 표지
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("CHS7006  기말 프로젝트  ·  5분 영상", {
    x: 0.7, y: 1.55, w: 12, h: 0.32,
    fontFace: "Calibri", fontSize: 14, color: "C9B8A0", margin: 0
  });
  s.addText("오늘마감", {
    x: 0.7, y: 2.05, w: 12, h: 1.15,
    fontFace: "Georgia", fontSize: 60, color: CREAM, bold: true, margin: 0
  });
  s.addText("오늘과 이번 주 마감만 본다.", {
    x: 0.7, y: 3.3, w: 12, h: 0.45,
    fontFace: "Calibri", fontSize: 22, color: "E6D8C6", margin: 0
  });
  s.addText("이재훈  2021314509\noneulmagam.vercel.app", {
    x: 0.7, y: 5.85, w: 12, h: 0.7,
    fontFace: "Calibri", fontSize: 15, color: "A89884", margin: 0
  });
  s.addNotes("0:00~0:20. 안녕하세요, 오늘마감입니다. 아이캠퍼스를 매일 들어가지 않아도 마감을 놓치지 않게 하는 스케이트보드 MVP입니다. 지금부터 문제, 고객, 시나리오, As-Is To-Be, 솔루션, KPI, 그리고 한 번 고친 결과까지 다섯 분에 맞춰 보여드리겠습니다.");
}

// 2 문제
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  label(s, "필수 1  ·  문제");
  h1(s, "매일 들어가도, 가끔 놓친다.");
  const items = [
    ["흩어짐", "과제 · 강의 · 퀴즈 · 시험이 과목마다, 메뉴마다 따로 있다."],
    ["부정확", "공지 시각이 마감처럼 보이거나, 빈 날이 있다가 한꺼번에 쌓인다."],
    ["손해", "그래서 불안해서 매일 확인한다. 그래도 학점에 손해를 본다."]
  ];
  items.forEach((it, i) => {
    const x = 0.7 + i * 4.1;
    card(s, x, 1.7, 3.9, 4.7);
    s.addText(String(i + 1).padStart(2, "0"), {
      x: x + 0.3, y: 1.95, w: 3.3, h: 0.4,
      fontFace: "Georgia", fontSize: 20, color: ACCENT, margin: 0
    });
    s.addText(it[0], {
      x: x + 0.3, y: 2.55, w: 3.3, h: 0.55,
      fontFace: "Georgia", fontSize: 26, color: INK, bold: true, margin: 0
    });
    s.addText(it[1], {
      x: x + 0.3, y: 3.3, w: 3.3, h: 2.4,
      fontFace: "Calibri", fontSize: 16, color: MUTED, margin: 0
    });
  });
  footer(s, 2);
  s.addNotes("0:20~0:50. 문제는 아이캠퍼스에 마감이 흩어지고 가끔 부정확하다는 것입니다. 학생은 불안해서 매일 들어갑니다. 그런데도 가끔 마감을 놓치고 학점에 손해를 봅니다. 기능이 부족한 게 아니라, 매일 전체를 훑어야 하는 구조가 문제입니다.");
}

// 3 고객
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  label(s, "필수 2  ·  고객");
  h1(s, "페르소나는 한 명이다.");
  s.addText("여러 명이 나오면 문제가 덜 뾰족한 것이다.", {
    x: 0.7, y: 1.45, w: 12, h: 0.35,
    fontFace: "Calibri", fontSize: 16, color: MUTED, margin: 0
  });
  card(s, 0.7, 2.0, 12, 4.55);
  const rows = [
    ["5~7과목", "한 학기 과목이 많아서 일정이 한 화면에 안 모인다."],
    ["매일 확인", "불안해서 아이캠퍼스를 습관처럼 연다."],
    ["학점이 걸린다", "마감을 놓치면 점수가 깎인다고 강하게 느낀다."],
    ["관리 시간은 싫다", "캘린더 앱을 또 키고 알림을 설정하고 싶지 않다."]
  ];
  rows.forEach((r, i) => {
    const y = 2.25 + i * 1.02;
    s.addShape(pres.shapes.OVAL, {
      x: 1.05, y: y + 0.16, w: 0.22, h: 0.22,
      fill: { color: i === 3 ? ACCENT : DONE }
    });
    s.addText(r[0], {
      x: 1.5, y, w: 3.3, h: 0.55,
      fontFace: "Georgia", fontSize: 20, color: INK, bold: true, valign: "middle", margin: 0
    });
    s.addText(r[1], {
      x: 5.0, y, w: 7.2, h: 0.55,
      fontFace: "Calibri", fontSize: 16, color: MUTED, valign: "middle", margin: 0
    });
  });
  footer(s, 3);
  s.addNotes("0:50~1:15. 고객은 한 명으로 잡았습니다. 한 학기 5에서 7과목을 듣는 학부생. 아이캠퍼스를 매일 확인하고, 마감 놓침을 학점 손실로 느끼고, 일정 관리에 시간을 쓰기 싫어하는 사람입니다. 페르소나를 여러 명으로 나누지 않았습니다.");
}

// 4 시나리오
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  label(s, "필수 3  ·  시나리오");
  h1(s, "수업 사이, 30초.");
  const steps = [
    ["1", "아침 또는\n수업 사이", "가방을 내려놓고 폰을 연다."],
    ["2", "오늘마감만\n본다", "오늘 두 개, 이번 주 네 개."],
    ["3", "하나만\n고른다", "지금 할 과제를 고르고 닫는다."],
    ["4", "아이캠은\n나중에", "새 일정을 넣을 때만 다시 연다."]
  ];
  steps.forEach((st, i) => {
    const x = 0.7 + i * 3.15;
    card(s, x, 1.75, 3.0, 4.65);
    s.addText(st[0], {
      x: x + 0.25, y: 2.0, w: 2.5, h: 0.45,
      fontFace: "Georgia", fontSize: 22, color: ACCENT, margin: 0
    });
    s.addText(st[1], {
      x: x + 0.25, y: 2.6, w: 2.5, h: 1.35,
      fontFace: "Georgia", fontSize: 22, color: INK, bold: true, margin: 0
    });
    s.addText(st[2], {
      x: x + 0.25, y: 4.15, w: 2.5, h: 1.6,
      fontFace: "Calibri", fontSize: 16, color: MUTED, margin: 0
    });
  });
  footer(s, 4);
  s.addNotes("1:15~1:40. 쓰는 상황은 거창하지 않습니다. 아침이거나 수업과 수업 사이입니다. 오늘마감을 열어 오늘과 이번 주만 보고, 지금 할 일 하나만 고릅니다. 아이캠퍼스 전체를 훑지 않습니다.");
}

// 5 As-Is To-Be
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  label(s, "필수 4  ·  As-Is  /  To-Be");
  h1(s, "행동이 바뀌어야 가설이 산다.");
  card(s, 0.7, 1.7, 5.75, 4.75);
  card(s, 6.9, 1.7, 5.75, 4.75);
  s.addText("As-Is", {
    x: 1.0, y: 1.95, w: 5.2, h: 0.32,
    fontFace: "Calibri", fontSize: 13, color: ACCENT, bold: true, margin: 0
  });
  s.addText("매일 아이캠퍼스", {
    x: 1.0, y: 2.4, w: 5.2, h: 0.5,
    fontFace: "Georgia", fontSize: 24, color: INK, bold: true, margin: 0
  });
  s.addText("과목별로 과제·강의·시험을 훑는다.\n공지가 마감처럼 보이면 다시 확인한다.\n그래도 가끔 놓친다.", {
    x: 1.0, y: 3.15, w: 5.2, h: 2.6,
    fontFace: "Calibri", fontSize: 17, color: MUTED, margin: 0
  });
  s.addText("To-Be", {
    x: 7.2, y: 1.95, w: 5.2, h: 0.32,
    fontFace: "Calibri", fontSize: 13, color: DONE, bold: true, margin: 0
  });
  s.addText("아침 30초, 오늘 두 개", {
    x: 7.2, y: 2.4, w: 5.2, h: 0.5,
    fontFace: "Georgia", fontSize: 24, color: INK, bold: true, margin: 0
  });
  s.addText("오늘마감만 보고 할 일을 고른다.\n아이캠퍼스는 새 일정을 넣을 때만 연다.\n마감이 없으면 안 들어가도 된다고 듣는다.", {
    x: 7.2, y: 3.15, w: 5.2, h: 2.6,
    fontFace: "Calibri", fontSize: 17, color: MUTED, margin: 0
  });
  footer(s, 5);
  s.addNotes("1:40~2:05. 쓰기 전후입니다. 지금은 매일 아이캠퍼스를 과목별로 훑고도 놓칩니다. 쓰고 나면 아침 30초에 오늘 할 일만 보고, 아이캠퍼스는 새 일정이 생길 때만 엽니다.");
}

// 6 솔루션
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  label(s, "필수 5  ·  솔루션");
  h1(s, "스케이트보드 하나. 웹 주소 하나.");
  s.addText("oneulmagam.vercel.app", {
    x: 0.7, y: 1.45, w: 12, h: 0.32,
    fontFace: "Calibri", fontSize: 16, color: ACCENT, margin: 0
  });
  const feats = [
    ["오늘 / 이번 주만", "학기 캘린더는 없다. 매일 확인하는 습관을 복제하지 않기 위해서다."],
    ["직접 넣기 + 피드", "비밀번호 스크래핑은 하지 않는다. 캘린더 피드나 CSV만 가져온다."],
    ["완료하면 끝", "오늘 남은 마감이 0이면, 아이캠퍼스에 안 들어가도 된다고 말한다."],
    ["링크를 보낸다", "설문 대신 share_click. 같은 과 친구가 이 화면을 쓰는지가 지표다."]
  ];
  feats.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.7 + col * 6.3;
    const y = 1.95 + row * 2.3;
    card(s, x, y, 6.05, 2.1);
    s.addText(f[0], {
      x: x + 0.35, y: y + 0.28, w: 5.35, h: 0.42,
      fontFace: "Georgia", fontSize: 20, color: INK, bold: true, margin: 0
    });
    s.addText(f[1], {
      x: x + 0.35, y: y + 0.82, w: 5.35, h: 0.95,
      fontFace: "Calibri", fontSize: 15, color: MUTED, margin: 0
    });
  });
  footer(s, 6);
  s.addNotes("2:05~3:20. 여기서 화면을 전환합니다. 오늘마감 사이트를 엽니다. 오늘과 이번 주 탭, 일정 넣기, 아이캠퍼스에서 가져오기, 완료하면 빈 화면이 되는 것, 링크 복사까지 보여 줍니다. 알림과 학기 캘린더와 비밀번호 로그인은 일부러 안 만들었습니다.");
}

// 7 KPI
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  label(s, "필수 6  ·  KPI");
  h1(s, "설문하지 않는다. 행동을 센다.");
  const kpis = [
    ["0건", "1주 마감 놓침"],
    ["주 2회 이하", "아이캠퍼스 접속"],
    ["today_cleared", "아침 확인 후\n오늘을 비움"],
    ["share_click", "링크를 복사해\n보낸 횟수"]
  ];
  kpis.forEach((k, i) => {
    const x = 0.7 + i * 3.15;
    card(s, x, 1.7, 3.0, 3.35);
    s.addText(k[0], {
      x: x + 0.18, y: 2.05, w: 2.64, h: 1.25,
      fontFace: "Georgia", fontSize: i >= 2 ? 16 : 26, color: ACCENT, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(k[1], {
      x: x + 0.18, y: 3.4, w: 2.64, h: 1.2,
      fontFace: "Calibri", fontSize: 15, color: INK, align: "center", margin: 0
    });
  });
  s.addText("기록: Supabase kpi_events  ·  사이트 KPI 내려받기  ·  오늘마감_KPI.xlsx  ·  GA 실시간", {
    x: 0.7, y: 5.25, w: 12, h: 0.35,
    fontFace: "Calibri", fontSize: 14, color: MUTED, margin: 0
  });
  footer(s, 7);
  s.addNotes("3:20~4:15. KPI는 설문조사가 아닙니다. 일주일에 마감을 몇 번 놓쳤는지, 아이캠퍼스를 몇 번 열었는지, 오늘 할 일을 비웠는지, 링크를 보냈는지를 셉니다. 여기서 Supabase 표나 엑셀 이벤트로그를 보여 줍니다. 지금 쌓인 것은 사이트 사용 기록입니다. 일지 노란 칸은 앞으로 일주일 직접 적는 칸입니다.");
}

// 8 반복
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("필수 7  ·  반복과 결과", {
    x: 0.7, y: 0.45, w: 12, h: 0.3,
    fontFace: "Calibri", fontSize: 13, color: "C9B8A0", bold: true, margin: 0
  });
  s.addText("한 번 만들고 끝내지 않았다.", {
    x: 0.7, y: 0.9, w: 12, h: 0.55,
    fontFace: "Georgia", fontSize: 28, color: CREAM, bold: true, margin: 0
  });
  const loops = [
    ["가설을 나눴다", "알림과 한눈에 보기를 한 문장에 넣지 않았다. 화면만 먼저 검증한다."],
    ["가져오기를 바꿨다", "비밀번호 스크래핑 대신 공식 캘린더 피드와 CSV만 받는다."],
    ["측정부터 붙였다", "설문 대신 visit, today_cleared, share_click을 쌓는다."],
    ["다음 바퀴", "일주일 일지를 채운 뒤에야 알림을 붙일지 결정한다."]
  ];
  loops.forEach((L, i) => {
    const y = 1.7 + i * 1.15;
    s.addText(String(i + 1), {
      x: 0.7, y, w: 0.55, h: 0.9,
      fontFace: "Georgia", fontSize: 22, color: "C9B8A0", valign: "middle", margin: 0
    });
    s.addText(L[0], {
      x: 1.4, y, w: 3.4, h: 0.9,
      fontFace: "Georgia", fontSize: 18, color: CREAM, valign: "middle", margin: 0
    });
    s.addText(L[1], {
      x: 5.0, y, w: 7.6, h: 0.9,
      fontFace: "Calibri", fontSize: 16, color: "E6D8C6", valign: "middle", margin: 0
    });
  });
  s.addNotes("4:15~5:00. 한 번 만들고 끝내지 않았습니다. 처음 가설에는 알림이 들어 있었습니다. 검증이 두 개가 되어 알림을 뺐습니다. 아이캠퍼스 비밀번호 스크래핑 대신 공식 피드만 받도록 바꿨습니다. 설문 대신 행동 로그를 붙였습니다. 다음 바퀴는 일주일 일지를 채운 뒤입니다. 이상입니다.");
}

pres.writeFile({ fileName: "/Users/leejaehun/Desktop/오늘마감/오늘마감_기말영상_초안.pptx" })
  .then(() => console.log("ok"))
  .catch((e) => { console.error(e); process.exit(1); });
