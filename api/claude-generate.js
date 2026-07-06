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

  try {
    // Use streaming so bytes flow immediately — prevents Node.js fetch from
    // timing out on long generations (interview prep can take 35–50 seconds).
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
    });

    if (!response.ok) {
      const errText = await response.text();
      console.error(`Anthropic Error [${response.status}]:`, errText);
      let errMessage = "AI Service Error";
      try {
        const errJson = JSON.parse(errText);
        if (errJson?.error?.message) errMessage = errJson.error.message;
      } catch {}
      return res.status(502).json({ error: errMessage, status: response.status, details: errText });
    }

    // Read the SSE stream and collect all text_delta chunks
    let fullText = "";
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });

      // Process complete SSE lines
      const lines = buffer.split("\n");
      buffer = lines.pop(); // keep the last incomplete line

      for (const line of lines) {
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
    }

    return res.status(200).json({ text: fullText });
  } catch (err) {
    console.error("API Crash:", err);
    return res.status(500).json({ error: "Server API Crash", details: err.message });
  }
}
