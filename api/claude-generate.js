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

  const payload = decodeJwt(token);
  if (!payload?.sub) {
    return res.status(401).json({ error: "Authentication failed. Please sign out and in." });
  }
  if (payload.exp && payload.exp < Math.floor(Date.now() / 1000)) {
    return res.status(401).json({ error: "Session expired. Please sign in again." });
  }

  const { prompt } = req.body || {};
  if (!prompt) return res.status(400).json({ error: "Empty prompt received." });

  // AbortController gives us a hard 55-second ceiling under Vercel's 60-second limit
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 55000);

  try {
    // stream: true causes Anthropic to send HTTP headers immediately (200 OK),
    // which prevents Node.js undici's 30-second headersTimeout from firing.
    // The SSE body then flows over the full generation time (~35-50s for
    // Interview Prep), well within undici's 300-second bodyTimeout.
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": anthropicKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 4096,
        stream: true,
        messages: [{ role: "user", content: prompt }],
      }),
      signal: controller.signal,
    });

    if (!response.ok) {
      clearTimeout(timeoutId);
      const errText = await response.text();
      console.error(`Anthropic Error [${response.status}]:`, errText);
      let errMessage = "AI Service Error";
      try {
        const errJson = JSON.parse(errText);
        // Anthropic error shape: { type, error: { type, message } }
        if (errJson?.error?.message) errMessage = errJson.error.message;
        else if (typeof errJson?.error === "string") errMessage = errJson.error;
        else if (errJson?.message) errMessage = errJson.message;
      } catch {}
      return res.status(502).json({ error: errMessage, status: response.status });
    }

    // Accumulate the full SSE response body as a string.
    // response.text() is simpler and equally correct here — the stream: true
    // flag already triggered the immediate-headers behaviour we need.
    const streamBody = await response.text();
    clearTimeout(timeoutId);

    let fullText = "";
    for (const line of streamBody.split("\n")) {
      if (!line.startsWith("data: ")) continue;
      const data = line.slice(6).trim();
      if (data === "[DONE]") continue;
      try {
        const event = JSON.parse(data);
        if (event.type === "content_block_delta" && event.delta?.type === "text_delta") {
          fullText += event.delta.text;
        }
      } catch {}
    }

    if (!fullText) {
      console.warn("Empty text extracted from stream body. First 500 chars:", streamBody.slice(0, 500));
      return res.status(500).json({ error: "Empty response from AI. Please try again." });
    }

    return res.status(200).json({ text: fullText });
  } catch (err) {
    clearTimeout(timeoutId);
    console.error("API Error:", err.name, err.message);
    if (err.name === "AbortError") {
      return res.status(504).json({ error: "Request timed out after 55 seconds. Please try again." });
    }
    return res.status(500).json({ error: "Server error: " + err.message });
  }
}
