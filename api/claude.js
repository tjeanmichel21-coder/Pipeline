// Vercel serverless function — proxies AI requests so the API key
// stays server-side and is never exposed to the browser.
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "POST only" });
  }
  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({
      content: [{ type: "text", text: "FAILED — ANTHROPIC_API_KEY env var is not set in Vercel." }],
    });
  }
  try {
    const headers = {
      "content-type": "application/json",
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    };
    // MCP connector calls need the beta header (Outlook alerts use this)
    if (req.body?.mcp_servers) headers["anthropic-beta"] = "mcp-client-2025-04-04";

    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers,
      body: JSON.stringify(req.body),
    });
    const data = await upstream.json();
    return res.status(upstream.status).json(data);
  } catch (e) {
    return res.status(502).json({
      content: [{ type: "text", text: "FAILED — upstream error: " + e.message }],
    });
  }
}
