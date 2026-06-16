// Vercel serverless function — sends notification email (pings, labor-ready
// notice, task notifications, stale-bid alerts) WITHOUT using an LLM, so it
// never consumes Anthropic billing.
//
// It supports two free/cheap transports and picks whichever is configured:
//
//   1) Resend (recommended — free tier, no Microsoft admin needed)
//        RESEND_API_KEY   — your Resend API key (re_...)
//        RESEND_FROM      — verified sender, e.g. "Pipeline CRM <noreply@coastalplumbingswfl.com>"
//                           (optional; defaults to Resend's onboarding sender for testing)
//
//   2) Microsoft Graph (app-only / client credentials — needs Mail.Send admin consent)
//        MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER
//
// `to` may be a single address or a comma-separated list.
export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ ok: false, error: "POST only" });

  const { to, subject, body } = req.body || {};
  if (!to || !subject) return res.status(400).json({ ok: false, error: "Missing 'to' or 'subject'." });

  const addresses = String(to).split(",").map((a) => a.trim()).filter(Boolean);

  try {
    // ---- Transport 1: Resend ----
    if (process.env.RESEND_API_KEY) {
      const from = process.env.RESEND_FROM || "Pipeline CRM <onboarding@resend.dev>";
      const r = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: { authorization: `Bearer ${process.env.RESEND_API_KEY}`, "content-type": "application/json" },
        body: JSON.stringify({ from, to: addresses, subject, text: body || "" }),
      });
      if (r.ok) return res.status(200).json({ ok: true });
      const err = await r.json().catch(() => ({}));
      return res.status(502).json({ ok: false, error: "Resend error: " + (err.message || err.name || r.statusText) });
    }

    // ---- Transport 2: Microsoft Graph (app-only) ----
    const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER } = process.env;
    if (MS_TENANT_ID && MS_CLIENT_ID && MS_CLIENT_SECRET && MS_SENDER) {
      const tokenRes = await fetch(`https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "content-type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: MS_CLIENT_ID,
          client_secret: MS_CLIENT_SECRET,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
        }),
      });
      const tokenJson = await tokenRes.json();
      if (!tokenRes.ok || !tokenJson.access_token) {
        return res.status(502).json({ ok: false, error: "Auth failed: " + (tokenJson.error_description || tokenJson.error || "no access token returned") });
      }
      const toRecipients = addresses.map((address) => ({ emailAddress: { address } }));
      const graphRes = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(MS_SENDER)}/sendMail`, {
        method: "POST",
        headers: { authorization: `Bearer ${tokenJson.access_token}`, "content-type": "application/json" },
        body: JSON.stringify({
          message: { subject, body: { contentType: "Text", content: body || "" }, toRecipients },
          saveToSentItems: true,
        }),
      });
      if (graphRes.status === 202) return res.status(200).json({ ok: true });
      const errJson = await graphRes.json().catch(() => ({}));
      return res.status(502).json({ ok: false, error: "Graph send failed: " + (errJson.error?.message || graphRes.statusText) });
    }

    // ---- Nothing configured ----
    return res.status(500).json({
      ok: false,
      error: "Email isn't set up yet — add a RESEND_API_KEY (free) in Vercel, then redeploy.",
    });
  } catch (e) {
    return res.status(502).json({ ok: false, error: "Email service error: " + e.message });
  }
}
