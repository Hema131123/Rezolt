import Razorpay from "razorpay";

const allowedOrigins = new Set([
  "https://www.rezolt.in",
  "https://rezolt.in",
  "http://localhost:5173",
  "http://localhost:5179",
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

const PLAN_CATALOG = {
  starter_kit: { amountInr: 99, label: "Starter" },
  Pro_kit: { amountInr: 299, label: "Pro" },
  unlimited_monthly: { amountInr: 599, label: "Unlimited" },
};

export default async function handler(req, res) {
  const origin = req.headers.origin;
  if (origin && allowedOrigins.has(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
  }
  res.setHeader("Vary", "Origin");
  res.setHeader("Content-Type", "application/json");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : "";
  if (!token) {
    return res.status(401).json({ error: "Please sign in before starting a payment." });
  }

  const payload = decodeJwt(token);
  if (!payload?.sub) {
    return res.status(401).json({ error: "Your session expired. Please sign in again." });
  }
  if (payload.exp && payload.exp < Math.floor(Date.now() / 1000)) {
    return res.status(401).json({ error: "Your session expired. Please sign in again." });
  }
  const user = { id: payload.sub, email: payload.email || "" };

  const keyId = process.env.RAZORPAY_KEY_ID || process.env.VITE_RAZORPAY_KEY_ID;
  const keySecret = process.env.RAZORPAY_KEY_SECRET;

  if (!keyId || !keySecret) {
    console.error("Razorpay config missing:", {
      hasKeyId: Boolean(keyId),
      hasKeySecret: Boolean(keySecret),
    });
    return res.status(500).json({
      error: "Razorpay is not configured. Please add RAZORPAY_KEY_ID and RAZORPAY_KEY_SECRET in Vercel.",
    });
  }

  const razorpay = new Razorpay({
    key_id: keyId,
    key_secret: keySecret,
  });

  try {
    const { planId: directPlanId, receipt, notes } = req.body || {};
    const planId = directPlanId || notes?.plan;
    const selectedPlan = PLAN_CATALOG[planId];

    if (!selectedPlan) {
      return res.status(400).json({ error: "Invalid plan selected" });
    }

    const safeNotes = {
      plan: planId,
      plan_label: selectedPlan.label,
      user_id: user.id,
      user_email: user.email || "",
    };

    const order = await razorpay.orders.create({
      amount: Math.round(selectedPlan.amountInr * 100),
      currency: "INR",
      receipt: receipt || `rzp_${Date.now()}`,
      notes: safeNotes,
    });

    console.info("Created Razorpay order", {
      userId: user.id,
      orderId: order.id,
      amount: order.amount,
      planId,
    });

    return res.status(200).json({
      orderId: order.id,
      amount: order.amount,
      currency: order.currency,
      planId,
    });
  } catch (err) {
    const statusCode = err?.statusCode || 500;
    const description = err?.error?.description || err?.message || "Failed to create order";
    const code = err?.error?.code || err?.code || "UNKNOWN_ERROR";

    console.error("Razorpay order error:", { statusCode, code, description, userId: user?.id });

    return res.status(500).json({
      error: statusCode === 401
        ? "Razorpay authentication failed. Make sure your Key ID and Key Secret belong to the same Razorpay account and both are either test or live."
        : description,
    });
  }
}