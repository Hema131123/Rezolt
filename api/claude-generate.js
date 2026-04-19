import { createClient } from "@supabase/supabase-js";

const cleanKey = (val) => (val || "").replace(/\\n/g, "").replace(/\n/g, "").replace(/\r/g, "").trim();

const supabaseUrl = cleanKey(process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL);
const supabaseAnonKey = cleanKey(process.env.SUPABASE_ANON_KEY || process.env.VITE_SUPABASE_ANON_KEY);
const anthropicKey = cleanKey(process.env.ANTHROPIC_API_KEY);

const allowedOrigins = new Set([
  "https://www.rezolt.in",
  "https://rezolt.in",
  "http://localhost:5173",
  "http://127.0.0.1:5173",
]);

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

  if (!anthropicKey || !supabaseUrl || !supabaseAnonKey) {
    return res.status(500).json({ error: "Server keys missing. Check Vercel Dashboard." });
  }

  // Auth check
  const supabase = createClient(supabaseUrl, supabaseAnonKey);
  const { data: authData, error: authError } = await supabase.auth.getUser(token);
  if (authError || !authData?.user) {
    return res.status(401).json({ error: "Authentication failed. Please sign out and in.", diagnostic: authError?.message });
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
        model: "claude-sonnet-4-5",
        max_tokens: 2500,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      console.error("Anthropic Error:", err);
      return res.status(502).json({ error: "AI Service Error", details: err });
    }

    const data = await response.json();
    return res.status(200).json({ text: data.content?.[0]?.text || "" });
  } catch (err) {
    console.error("API Crash:", err);
    return res.status(500).json({ error: "Server API Crash", details: err.message });
  }
}