import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
const supabaseAnonKey = process.env.SUPABASE_ANON_KEY || process.env.VITE_SUPABASE_ANON_KEY;
const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

const allowedOrigins = new Set([
  "https://www.rezolt.in",
  "https://rezolt.in",
  "http://localhost:5173",
  "http://localhost:5179",
  "http://127.0.0.1:5173",
]);

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

  const anonClient = createClient(supabaseUrl, supabaseAnonKey);
  const { data: authData, error: authError } = await anonClient.auth.getUser(token);
  if (authError || !authData?.user) return res.status(401).json({ error: "Invalid session" });

  const serviceClient = createClient(supabaseUrl, serviceKey);
  const { data: profile, error: profileError } = await serviceClient
    .from("profiles")
    .select("credits, plan")
    .eq("id", authData.user.id)
    .single();

  if (profileError) return res.status(500).json({ error: profileError.message });
  if (profile?.plan === "unlimited") return res.status(200).json({ credits: null });

  const current = profile?.credits ?? 0;
  const next = Math.max(0, current - 1);

  const { error: updateError } = await serviceClient
    .from("profiles")
    .update({ credits: next })
    .eq("id", authData.user.id);

  if (updateError) {
    console.error("Credit decrement failed:", updateError);
    return res.status(500).json({ error: updateError.message });
  }

  return res.status(200).json({ credits: next });
}
