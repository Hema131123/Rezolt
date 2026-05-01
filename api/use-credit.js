import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

const allowedOrigins = new Set([
  "https://www.rezolt.in",
  "https://rezolt.in",
  "http://localhost:5173",
  "http://localhost:5179",
  "http://127.0.0.1:5173",
]);

function decodeJwt(token) {
  try {
    return JSON.parse(Buffer.from(token.split(".")[1], "base64").toString());
  } catch {
    return null;
  }
}

export default async function handler(req, res) {
  const origin = req.headers.origin;
  if (origin && allowedOrigins.has(origin)) res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Content-Type", "application/json");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const token = (req.headers.authorization || "").replace("Bearer ", "").trim();
  if (!token) return res.status(401).json({ error: "Unauthorized" });

  const payload = decodeJwt(token);
  if (!payload?.sub) return res.status(401).json({ error: "Invalid session" });
  if (payload.exp && payload.exp < Math.floor(Date.now() / 1000)) {
    return res.status(401).json({ error: "Session expired" });
  }
  const userId = payload.sub;

  const serviceClient = createClient(supabaseUrl, serviceKey);
  const { data: profile, error: profileError } = await serviceClient
    .from("profiles")
    .select("credits, plan")
    .eq("id", userId)
    .single();

  if (profileError) return res.status(500).json({ error: profileError.message });

  const planLower = (profile?.plan || "").toLowerCase();
  if (planLower === "unlimited" || planLower === "free") {
    return res.status(200).json({ credits: null });
  }

  const current = profile?.credits ?? 0;
  if (current <= 0) {
    return res.status(200).json({ credits: 0 });
  }

  const next = current - 1;
  const { error: updateError } = await serviceClient
    .from("profiles")
    .update({ credits: next })
    .eq("id", userId);

  if (updateError) {
    console.error("Credit decrement failed:", updateError);
    return res.status(500).json({ error: updateError.message });
  }

  return res.status(200).json({ credits: next });
}
