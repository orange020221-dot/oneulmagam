export default function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  res.status(200).json({
    gaId: String(process.env.GA_MEASUREMENT_ID || "").trim(),
    siteUrl: String(process.env.SITE_URL || "").trim()
  });
}
