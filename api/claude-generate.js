const cleanKey = (val) => (val || "").replace(/\\n/g, "").replace(/\n/g, "").replace(/\r/g, "").trim();

const anthropicKey = cleanKey(process.env.ANTHROPIC_API_KEY);

const allowedOrigins = new Set([
  "https://www.rezolt.in",
  "https://rezolt.in",
  "http://localhost:5173",
  "http://127.0.0.1:5173",
]);

function decodeJwt(token) {
  try {
    const payload = JSON.parse(Buffer.from(token.split(".")[1], "base64").toString());
    return payload;
  } catch {
    return null;
  }
}

export default async function handler(req, res) {
  const origin = req.headers.origin;
  if (origin && allowedOrigins.has(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
  }
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : "";

  if (!anthropicKey) {
    return res.status(500).json({ error: "Server keys missing. Check Vercel Dashboard." });
  }

  // Verify token via local JWT decode — no Supabase auth network call needed
  const payload = decodeJwt(token);
  if (!payload?.sub) {
    return res.status(401).json({ error: "Authentication failed. Please sign out and in." });
  }
  if (payload.exp && payload.exp < Math.floor(Date.now() / 1000)) {
    return res.status(401).json({ error: "Session expired. Please sign in again." });
  }

  const { prompt } = req.body || {};
  if (!prompt) return res.status(400).json({ error: "Empty prompt received." });

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": anthropicKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-3-5-sonnet-20241022",
        max_tokens: 3000,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      console.error(`Anthropic Error [${response.status}]:`, err);
      return res.status(502).json({ error: "AI Service Error", status: response.status, details: err });
    }

    const data = await response.json();
    return res.status(200).json({ text: data.content?.[0]?.text || "" });
  } catch (err) {
    console.error("API Crash:", err);
    return res.status(500).json({ error: "Server API Crash", details: err.message });
  }
}
