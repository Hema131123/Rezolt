export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  if (req.method === "OPTIONS") return res.status(200).end();

  const clean = (val) => (val || "").replace(/\\n/g, "").replace(/\n/g, "").replace(/\r/g, "").trim();

  const report = {
    anthropic: clean(process.env.ANTHROPIC_API_KEY).length > 10 ? "Ready" : "Missing",
    supabase_url: clean(process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL) ? "Ready" : "Missing",
    supabase_anon: clean(process.env.SUPABASE_ANON_KEY || process.env.VITE_SUPABASE_ANON_KEY) ? "Ready" : "Missing",
    service_role: clean(process.env.SUPABASE_SERVICE_ROLE_KEY) ? "Ready" : "Missing",
    status: "Healthy"
  };

  return res.status(200).json(report);
}
