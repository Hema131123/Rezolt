import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
const supabaseAnonKey = process.env.SUPABASE_ANON_KEY || process.env.VITE_SUPABASE_ANON_KEY;
const supabaseServiceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
const allowedOrigins = new Set([
  "https://www.rezolt.in",
  "https://rezolt.in",
  "http://localhost:5173",
  "http://127.0.0.1:5173",
]);

// Per-plan daily generation limits (per user, resets each UTC day)
const DAILY_LIMITS = {
  starter:   2,    // 1 free kit × 2 tabs each ≈ 2 calls
  hustler:   30,   // generous headroom for the credit pack
  unlimited: 15,   // cost-control hard cap for the ₹599 plan
};

async function checkDailyLimit(userId, plan) {
  if (!supabaseUrl || !supabaseServiceKey) return true; // allow if DB check impossible
  const supabase = createClient(supabaseUrl, supabaseServiceKey);
  const todayStart = new Date();
  todayStart.setUTCHours(0, 0, 0, 0);

  const { count, error } = await supabase
    .from("kits")
    .select("id", { count: "exact", head: true })
    .eq("user_id", userId)
    .gte("created_at", todayStart.toISOString());

  if (error) return true; // fail open, don't block on DB errors

  const limit = DAILY_LIMITS[plan] ?? DAILY_LIMITS.starter;
  return (count ?? 0) < limit;
}

// New-account free-kit abuse guard: check total kits ever created for free accounts
async function checkFreeAbuseGuard(userId, plan) {
  if (plan !== "starter") return true;
  if (!supabaseUrl || !supabaseServiceKey) return true;
  const supabase = createClient(supabaseUrl, supabaseServiceKey);
  const { count, error } = await supabase
    .from("kits")
    .select("id", { count: "exact", head: true })
    .eq("user_id", userId);
  if (error) return true;
  return (count ?? 0) < 1; // hard cap: 1 kit per free account lifetime
}

export default async function handler(req, res) {
  const origin = req.headers.origin;
  if (origin && allowedOrigins.has(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
  }
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : "";
  if (!token) {
    return res.status(401).json({ error: "Please sign in to use Rezolt generation." });
  }

  if (!supabaseUrl || !supabaseAnonKey) {
    console.error("Missing Supabase environment variables for API auth");
    return res.status(500).json({ error: "Server configuration error" });
  }

  const supabase = createClient(supabaseUrl, supabaseAnonKey);
  const { data: authData, error: authError } = await supabase.auth.getUser(token);
  const user = authData?.user;
  if (authError || !user) {
    console.warn("Blocked unauthenticated AI request", {
      ip: req.headers["x-forwarded-for"] || req.socket?.remoteAddress || "unknown",
    });
    return res.status(401).json({ error: "Your session expired. Please sign in again." });
  }

  // Fetch profile to know the user's current plan
  let userPlan = "starter";
  if (supabaseServiceKey) {
    const adminClient = createClient(supabaseUrl, supabaseServiceKey);
    const { data: profile } = await adminClient
      .from("profiles")
      .select("plan")
      .eq("id", user.id)
      .single();
    userPlan = profile?.plan ?? "starter";
  }

  // Free-account lifetime cap (anti-abuse: throwaway emails cycling free kits)
  if (!(await checkFreeAbuseGuard(user.id, userPlan))) {
    console.warn("Free-kit abuse guard triggered", { userId: user.id, email: user.email });
    return res.status(429).json({ error: "Free kit limit reached. Upgrade to continue." });
  }

  // Daily per-user generation cap
  if (!(await checkDailyLimit(user.id, userPlan))) {
    console.warn("Daily generation limit reached", { userId: user.id, plan: userPlan });
    return res.status(429).json({ error: "Daily generation limit reached. Come back tomorrow or upgrade your plan." });
  }

  const { prompt } = req.body || {};

  if (!prompt || typeof prompt !== "string" || prompt.length < 10 || prompt.length > 30000) {
    return res.status(400).json({ error: "Invalid prompt" });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    console.error("Missing ANTHROPIC_API_KEY environment variable");
    return res.status(500).json({ error: "Server configuration error" });
  }

  try {
    console.info("Anthropic generation request", {
      userId: user.id,
      email: user.email,
      promptLength: prompt.length,
      ip: req.headers["x-forwarded-for"] || req.socket?.remoteAddress || "unknown",
    });

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 2000,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      console.error("Anthropic API error:", err);
      return res.status(502).json({ error: "AI generation failed" });
    }

    const data = await response.json();
    const text = data.content?.[0]?.text || "";
    return res.status(200).json({ text });
  } catch (err) {
    console.error("Generate error:", err);
    return res.status(500).json({ error: "Internal server error" });
  }
}