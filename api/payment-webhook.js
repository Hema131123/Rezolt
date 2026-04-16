import crypto from "crypto";
import { createClient } from "@supabase/supabase-js";

export const config = { api: { bodyParser: false } };

const supabaseUrl = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
const PLAN_PRICING_PAISE = {
  starter_kit: 9900,
  hustler_kit: 29900,
  unlimited_monthly: 59900,
};

async function getRawBody(req) {
  return new Promise((resolve, reject) => {
    let data = "";
    req.on("data", chunk => { data += chunk; });
    req.on("end", () => resolve(data));
    req.on("error", reject);
  });
}

async function updateProfileWithRetry(userId, plan, maxAttempts = 4) {
  const supabase = createClient(supabaseUrl, serviceRoleKey);
  const creditMap = { starter_kit: 1, hustler_kit: 5 };

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const { data: profile, error: fetchErr } = await supabase
        .from("profiles")
        .select("credits, plan")
        .eq("id", userId)
        .single();

      if (fetchErr) throw fetchErr;

      const currentCredits = Math.max(0, profile?.credits ?? 0);
      const newCredits = plan === "unlimited_monthly" ? 9999 : currentCredits + (creditMap[plan] ?? 0);
      const newPlan = plan === "unlimited_monthly"
        ? "unlimited"
        : plan === "hustler_kit"
          ? (profile?.plan === "unlimited" ? "unlimited" : "hustler")
          : (profile?.plan ?? "starter");

      const { error: updateErr } = await supabase
        .from("profiles")
        .update({ credits: newCredits, plan: newPlan })
        .eq("id", userId);

      if (updateErr) throw updateErr;

      console.log(`[webhook] Updated user ${userId} to plan=${newPlan} credits=${newCredits} (attempt ${attempt})`);
      return true;
    } catch (err) {
      const delay = Math.min(500 * Math.pow(2, attempt - 1), 4000);
      console.error(`[webhook] Supabase update attempt ${attempt} failed for ${userId}: ${err?.message}. Retrying in ${delay}ms.`);
      if (attempt < maxAttempts) {
        await new Promise(resolve => setTimeout(resolve, delay));
      }
    }
  }

  // All retries exhausted — log for manual recovery
  console.error(`[webhook] CRITICAL: Failed to update profile for user ${userId} plan=${plan} after ${maxAttempts} attempts. Manual reconciliation required.`);
  return false;
}

function isMissingTableError(error) {
  const msg = `${error?.message || ""} ${error?.details || ""}`.toLowerCase();
  return msg.includes("does not exist") || msg.includes("relation") || msg.includes("42p01");
}

function isDuplicateError(error) {
  const msg = `${error?.message || ""} ${error?.details || ""}`.toLowerCase();
  return error?.code === "23505" || msg.includes("duplicate") || msg.includes("unique");
}

async function wasPaymentProcessed(paymentId) {
  const supabase = createClient(supabaseUrl, serviceRoleKey);
  const { data, error } = await supabase
    .from("payment_events")
    .select("payment_id")
    .eq("payment_id", paymentId)
    .maybeSingle();

  if (!error) return Boolean(data);
  if (isMissingTableError(error)) {
    console.warn("[webhook] payment_events table missing; idempotency check unavailable.");
    return false;
  }

  throw error;
}

async function markPaymentProcessed(payment) {
  const supabase = createClient(supabaseUrl, serviceRoleKey);
  const { error } = await supabase.from("payment_events").insert({
    payment_id: payment.id,
    order_id: payment.order_id || null,
    user_id: payment.notes?.user_id || null,
    plan_code: payment.notes?.plan || null,
    amount_paise: payment.amount || null,
    currency: payment.currency || null,
    status: payment.status || null,
  });

  if (!error) return true;
  if (isMissingTableError(error)) {
    console.warn("[webhook] payment_events table missing; idempotency write unavailable.");
    return false;
  }
  if (isDuplicateError(error)) {
    console.warn(`[webhook] Duplicate payment event ignored for payment_id=${payment.id}`);
    return false;
  }

  throw error;
}

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  try {
    const rawBody = await getRawBody(req);
    const signature = req.headers["x-razorpay-signature"];
    const secret = process.env.RAZORPAY_WEBHOOK_SECRET || process.env.RAZORPAY_KEY_SECRET;

    if (!signature || !secret) {
      console.error("Missing webhook signature or secret");
      return res.status(400).json({ error: "Invalid webhook configuration" });
    }

    const expected = crypto
      .createHmac("sha256", secret)
      .update(rawBody)
      .digest("hex");

    if (signature !== expected) {
      console.error("Webhook signature mismatch");
      return res.status(400).json({ error: "Invalid signature" });
    }

    const event = JSON.parse(rawBody);
    const { event: eventType, payload } = event;

    if (eventType === "payment.captured") {
      const payment = payload?.payment?.entity;
      const { plan, user_id } = payment?.notes || {};
      const expectedAmount = PLAN_PRICING_PAISE[plan];

      if (!payment?.id) {
        console.warn("[webhook] Missing payment id in captured event.");
        return res.status(200).json({ received: true });
      }

      if (expectedAmount == null) {
        console.warn(`[webhook] Unknown plan code in payment notes. payment_id=${payment.id} plan=${plan}`);
        return res.status(200).json({ received: true });
      }

      if (payment.currency !== "INR" || Number(payment.amount) !== Number(expectedAmount)) {
        console.error(`[webhook] Amount/currency mismatch. payment_id=${payment.id} plan=${plan} expected=${expectedAmount} got=${payment.amount} ${payment.currency}`);
        return res.status(200).json({ received: true });
      }

      if (await wasPaymentProcessed(payment.id)) {
        console.info(`[webhook] Already processed payment_id=${payment.id}; skipping.`);
        return res.status(200).json({ received: true });
      }

      if (user_id && plan && supabaseUrl && serviceRoleKey) {
        const success = await updateProfileWithRetry(user_id, plan);
        if (!success) {
          // Return 200 so Razorpay doesn't keep re-firing; we've already logged for manual recovery
          console.error(`[webhook] Manual recovery needed: payment_id=${payment?.id} user=${user_id} plan=${plan}`);
        } else {
          await markPaymentProcessed(payment);
        }
      } else {
        console.warn(`[webhook] Missing notes on payment. user_id=${user_id} plan=${plan} payment_id=${payment?.id}`);
      }
    }

    return res.status(200).json({ received: true });
  } catch (err) {
    console.error("Webhook error:", err);
    return res.status(500).json({ error: "Webhook processing failed" });
  }
}