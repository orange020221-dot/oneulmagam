const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.defineLayout({ name: "WIDE16", width: 13.333, height: 7.5 });
pres.layout = "WIDE16";
pres.author = "이재훈";
pres.title = "오늘마감";
pres.subject = "CHS7006 기말 5분 발표";

const INK = "1B1814";
const PAPER = "F3EEE4";
const CARD = "FFFCF6";
const MUTED = "6F675C";
const ACCENT = "B4451B";
const DONE = "2C5A4A";

function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: fill || CARD },
    rectRadius: 0.08,
    shadow: { type: "outer", color: "1B1814", blur: 10, offset: 3, angle: 135, opacity: 0.08 }
  });
}

// 1. Title
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("CHS7006  ·  신인류 AI 사피엔스 경험디자인", {
    x: 0.7, y: 1.6, w: 12, h: 0.35, fontFace: "Georgia", fontSize: 14, color: "C9B8A0", margin: 0
  });
  s.addText("오늘마감", {
    x: 0.7, y: 2.15, w: 12, h: 1.2, fontFace: "Georgia", fontSize: 60, color: "FFF8EE", bold: true, margin: 0
  });
  s.addText("오늘과 이번 주 마감만 본다.", {
    x: 0.7, y: 3.4, w: 12, h: 0.45, fontFace: "Arial", fontSize: 22, color: "E6D8C6", margin: 0
  });
  s.addText("이재훈  2021314509  ·  oneulmagam.vercel.app  ·  5분", {
    x: 0.7, y: 6.6, w: 12, h: 0.3, fontFace: "Arial", fontSize: 14, color: "A89884", margin: 0
  });
}

// 2. Problem
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addText("문제", { x: 0.7, y: 0.4, w: 12, h: 0.4, fontFace: "Georgia", fontSize: 16, color: ACCENT, margin: 0 });
  s.addText("매일 들어가도, 가끔 놓친다.", {
    x: 0.7, y: 0.85, w: 12, h: 0.7, fontFace: "Georgia", fontSize: 32, color: INK, bold: true, margin: 0
  });
  const probs = [
    ["흩어짐", "과제 · 강의 · 시험 · 퀴즈가 과목마다, 메뉴마다 따로 있다."],
    ["부정확", "공지 시각이 마감처럼 보이거나, 일정이 비었다가 한꺼번에 쌓인다."],
    ["불안", "그래서 매일 아이캠퍼스를 연다. 그래도 학점에 손해를 본다."]
  ];
  probs.forEach((p, i) => {
    const x = 0.7 + i * 4.1;
    card(s, x, 2.1, 3.85, 4.3);
    s.addText(String(i + 1).padStart(2, "0"), {
      x: x + 0.3, y: 2.35, w: 3.2, h: 0.4, fontFace: "Georgia", fontSize: 20, color: ACCENT, margin: 0
    });
    s.addText(p[0], {
      x: x + 0.3, y: 2.9, w: 3.2, h: 0.55, fontFace: "Georgia", fontSize: 24, color: INK, bold: true, margin: 0
    });
    s.addText(p[1], {
      x: x + 0.3, y: 3.6, w: 3.2, h: 2.2, fontFace: "Arial", fontSize: 16, color: MUTED, margin: 0
    });
  });
}

// 3. Persona
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addText("페르소나 하나", { x: 0.7, y: 0.4, w: 12, h: 0.35, fontFace: "Georgia", fontSize: 16, color: ACCENT, margin: 0 });
  s.addText("여러 명이 나오면 문제가 덜 뾰족한 것이다.", {
    x: 0.7, y: 0.85, w: 12, h: 0.55, fontFace: "Georgia", fontSize: 28, color: INK, bold: true, margin: 0
  });
  card(s, 0.7, 1.8, 12, 4.7);
  const traits = [
    ["5~7과목", "한 학기에 듣는 과목이 많아서 일정이 한 화면에 안 모인다."],
    ["매일 확인", "불안해서 아이캠퍼스를 습관처럼 연다."],
    ["학점이 걸린다", "마감을 놓치면 점수가 깎인다고 강하게 느낀다."],
    ["관리 시간은 싫다", "캘린더 앱을 또 키고 알림을 설정하고 싶지 않다."]
  ];
  traits.forEach((t, i) => {
    const y = 2.1 + i * 1.0;
    s.addShape(pres.shapes.OVAL, { x: 1.05, y: y + 0.12, w: 0.22, h: 0.22, fill: { color: i === 3 ? ACCENT : DONE } });
    s.addText(t[0], { x: 1.5, y, w: 3.2, h: 0.45, fontFace: "Georgia", fontSize: 20, color: INK, bold: true, margin: 0 });
    s.addText(t[1], { x: 4.8, y, w: 7.4, h: 0.45, fontFace: "Arial", fontSize: 16, color: MUTED, margin: 0 });
  });
}

// 4. Hypothesis
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("가설 하나", { x: 0.7, y: 1.5, w: 12, h: 0.35, fontFace: "Georgia", fontSize: 16, color: "C9B8A0", margin: 0 });
  s.addText("오늘 · 이번 주 마감만\n한 화면에서 볼 수 있으면,\n매일 아이캠퍼스 전체를 훑지 않아도\n마감을 놓치지 않을 것이다.", {
    x: 0.7, y: 2.05, w: 12, h: 2.8, fontFace: "Georgia", fontSize: 28, color: "FFF8EE", margin: 0
  });
  s.addText("알림은 다음 바퀴다. 한 문장에 검증을 두 개 넣지 않는다.", {
    x: 0.7, y: 5.4, w: 12, h: 0.4, fontFace: "Arial", fontSize: 16, color: "A89884", margin: 0
  });
}

// 5. As-Is / To-Be
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addText("As-Is  →  To-Be", { x: 0.7, y: 0.4, w: 12, h: 0.4, fontFace: "Georgia", fontSize: 16, color: ACCENT, margin: 0 });
  s.addText("행동이 바뀌어야 가설이 산다.", {
    x: 0.7, y: 0.85, w: 12, h: 0.5, fontFace: "Georgia", fontSize: 28, color: INK, bold: true, margin: 0
  });
  card(s, 0.7, 1.7, 5.7, 4.9);
  card(s, 6.95, 1.7, 5.7, 4.9);
  s.addText("지금", { x: 1.0, y: 1.95, w: 5.1, h: 0.35, fontFace: "Arial", fontSize: 13, color: ACCENT, margin: 0 });
  s.addText("매일 아이캠퍼스", { x: 1.0, y: 2.4, w: 5.1, h: 0.5, fontFace: "Georgia", fontSize: 24, color: INK, bold: true, margin: 0 });
  s.addText("과목별로 과제·강의·시험을 훑는다.\n공지가 마감처럼 보이면 다시 확인한다.\n그래도 가끔 놓친다.", {
    x: 1.0, y: 3.15, w: 5.1, h: 2.8, fontFace: "Arial", fontSize: 16, color: MUTED, margin: 0
  });
  s.addText("목표", { x: 7.25, y: 1.95, w: 5.1, h: 0.35, fontFace: "Arial", fontSize: 13, color: DONE, margin: 0 });
  s.addText("아침 30초, 오늘 두 개", { x: 7.25, y: 2.4, w: 5.1, h: 0.5, fontFace: "Georgia", fontSize: 24, color: INK, bold: true, margin: 0 });
  s.addText("오늘마감만 보고 할 일을 고른다.\n아이캠퍼스는 새 일정을 넣을 때만 연다.\n마감이 없으면 안 들어가도 된다고 말한다.", {
    x: 7.25, y: 3.15, w: 5.1, h: 2.8, fontFace: "Arial", fontSize: 16, color: MUTED, margin: 0
  });
}

// 6. Solution
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addText("솔루션", { x: 0.7, y: 0.4, w: 12, h: 0.35, fontFace: "Georgia", fontSize: 16, color: ACCENT, margin: 0 });
  s.addText("스케이트보드 하나. 웹 주소 하나.", {
    x: 0.7, y: 0.85, w: 12, h: 0.5, fontFace: "Georgia", fontSize: 28, color: INK, bold: true, margin: 0
  });
  const feats = [
    ["오늘 / 이번 주", "학기 캘린더는 만들지 않았다. 매일 확인하는 습관을 복제하지 않기 위해서다."],
    ["직접 넣기 + 피드", "비밀번호 스크래핑은 하지 않는다. 캘린더 피드나 CSV만 가져온다."],
    ["완료하면 끝", "오늘 남은 마감이 0이면, 아이캠퍼스에 안 들어가도 된다고 말한다."],
    ["링크를 보낸다", "설문 대신 share_click. 같은 과 친구가 이 화면을 쓰는지가 KPI다."]
  ];
  feats.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.7 + col * 6.3;
    const y = 1.65 + row * 2.55;
    card(s, x, y, 6.0, 2.35);
    s.addText(f[0], { x: x + 0.35, y: y + 0.3, w: 5.3, h: 0.45, fontFace: "Georgia", fontSize: 22, color: INK, bold: true, margin: 0 });
    s.addText(f[1], { x: x + 0.35, y: y + 0.9, w: 5.3, h: 1.1, fontFace: "Arial", fontSize: 15, color: MUTED, margin: 0 });
  });
}

// 7. KPI
{
  const s = pres.addSlide();
  s.background = { color: PAPER };
  s.addText("KPI", { x: 0.7, y: 0.4, w: 12, h: 0.35, fontFace: "Georgia", fontSize: 16, color: ACCENT, margin: 0 });
  s.addText("설문하지 않는다. 행동을 센다.", {
    x: 0.7, y: 0.85, w: 12, h: 0.5, fontFace: "Georgia", fontSize: 28, color: INK, bold: true, margin: 0
  });
  const kpis = [
    ["0", "1주 마감 놓침"],
    ["≤2", "아이캠퍼스 주간 접속"],
    ["today_cleared", "아침 확인 후 오늘을 비움"],
    ["share_click", "링크를 복사해 보낸 횟수"]
  ];
  kpis.forEach((k, i) => {
    const x = 0.7 + i * 3.15;
    card(s, x, 1.7, 3.0, 3.5);
    s.addText(k[0], {
      x: x + 0.2, y: 2.15, w: 2.6, h: 1.3, fontFace: "Georgia", fontSize: i === 2 || i === 3 ? 18 : 40, color: ACCENT, bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(k[1], {
      x: x + 0.2, y: 3.6, w: 2.6, h: 1.1, fontFace: "Arial", fontSize: 15, color: INK, align: "center", margin: 0
    });
  });
  s.addText("기록은 오늘마감_KPI.xlsx · 사이트에서 KPI 내려받기 · 이후 Supabase / GA", {
    x: 0.7, y: 5.5, w: 12, h: 0.35, fontFace: "Arial", fontSize: 14, color: MUTED, margin: 0
  });
}

// 8. Next
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("다음 바퀴", { x: 0.7, y: 0.55, w: 12, h: 0.35, fontFace: "Georgia", fontSize: 16, color: "C9B8A0", margin: 0 });
  s.addText("한 주를 써 보고, 틀린 곳을 고친다.", {
    x: 0.7, y: 1.05, w: 12, h: 0.55, fontFace: "Georgia", fontSize: 28, color: "FFF8EE", bold: true, margin: 0
  });
  const nexts = [
    ["지금", "한눈에 보기 + 피드 가져오기. 가설 1을 검증한다."],
    ["그다음", "화면이 아침 확인을 바꾸면, 그때 알림을 붙인다."],
    [" automations", "Supabase에 이벤트가 쌓이고, GA에서 방문이 잡히면 반복한다."]
  ];
  nexts[2][0] = "그 위";
  nexts.forEach((n, i) => {
    const y = 2.0 + i * 1.35;
    s.addText(n[0], { x: 0.7, y, w: 2.2, h: 0.9, fontFace: "Georgia", fontSize: 20, color: "C9B8A0", valign: "middle", margin: 0 });
    s.addText(n[1], { x: 3.1, y, w: 9.4, h: 0.9, fontFace: "Arial", fontSize: 20, color: "FFF8EE", valign: "middle", margin: 0 });
  });
  s.addText("oneulmagam.vercel.app", {
    x: 0.7, y: 6.55, w: 12, h: 0.3, fontFace: "Georgia", fontSize: 16, color: "C9B8A0", margin: 0
  });
}

pres.writeFile({ fileName: "/Users/leejaehun/Desktop/오늘마감/오늘마감_기말5분.pptx" })
  .then(() => console.log("pptx ok"))
  .catch((e) => { console.error(e); process.exit(1); });
