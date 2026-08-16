const ALLOWED = new Set(["canvas.skku.edu", "icampus.skku.edu"]);

export default async function handler(req, res) {
  const raw = String((req.query && req.query.url) || "");
  let parsed;
  try {
    parsed = new URL(raw);
  } catch {
    res.status(400).json({ ok: false, reason: "url" });
    return;
  }

  if (parsed.protocol !== "https:" || !ALLOWED.has(parsed.hostname)) {
    res.status(400).json({ ok: false, reason: "host" });
    return;
  }
  if (!parsed.pathname.includes("/feeds/calendars/") || !parsed.pathname.endsWith(".ics")) {
    res.status(400).json({ ok: false, reason: "path" });
    return;
  }

  const response = await fetch(parsed.toString(), {
    headers: { Accept: "text/calendar" }
  });
  if (!response.ok) {
    res.status(502).json({ ok: false, reason: "fetch" });
    return;
  }

  const text = await response.text();
  res.setHeader("Content-Type", "text/calendar; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.status(200).send(text);
}
