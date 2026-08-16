const ALLOWED = new Set([
  "visit",
  "add_item",
  "complete_item",
  "today_cleared",
  "share_click",
  "import_feed"
]);

function normalizeUrl(url) {
  return String(url || "").trim().replace(/\/+$/, "");
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ ok: false });
    return;
  }

  const url = normalizeUrl(process.env.SUPABASE_URL);
  const key = String(process.env.SUPABASE_ANON_KEY || "").trim();
  if (!url || !key) {
    res.status(501).json({ ok: false, reason: "db_not_configured" });
    return;
  }

  let body = req.body;
  if (typeof body === "string") {
    try { body = JSON.parse(body); } catch { body = {}; }
  }
  body = body || {};

  const event_name = String(body.event_name || "").slice(0, 40);
  if (!ALLOWED.has(event_name)) {
    res.status(400).json({ ok: false });
    return;
  }

  const payload = {
    event_name,
    item_type: body.item_type ? String(body.item_type).slice(0, 20) : null,
    used_at: new Date().toISOString()
  };

  const response = await fetch(`${url}/rest/v1/kpi_events`, {
    method: "POST",
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json",
      Prefer: "return=minimal"
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const text = await response.text();
    res.status(502).json({ ok: false, reason: text.slice(0, 200) });
    return;
  }

  res.status(204).end();
}
