// Vercel serverless function — sends email directly through Microsoft Graph
// (Outlook) using an app-only (client credentials) token. No LLM involved,
// so it does NOT consume Anthropic API billing.
//
// Required environment variables (set in Vercel → Settings → Environment Variables):
//   MS_TENANT_ID     — your Microsoft 365 / Entra directory (tenant) ID
//   MS_CLIENT_ID     — the Azure app registration's Application (client) ID
//   MS_CLIENT_SECRET — a client secret for that app registration
//   MS_SENDER        — the mailbox to send from (e.g. notifications@coastalplumbingswfl.com)
//
// The app registration needs the Microsoft Graph APPLICATION permission
// "Mail.Send" with admin consent granted.
export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ ok: false, error: "POST only" });

  const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER } = process.env;
  if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET || !MS_SENDER) {
    return res.status(500).json({
      ok: false,
      error: "Email isn't set up yet — an admin needs to add the Microsoft Graph env vars (MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER) in Vercel, then redeploy.",
    });
  }

  const { to, subject, body } = req.body || {};
  if (!to || !subject) return res.status(400).json({ ok: false, error: "Missing 'to' or 'subject'." });

  try {
    // 1) App-only access token via the client credentials flow.
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

    // 2) Send the mail. `to` may be a single address or a comma-separated list.
    const toRecipients = String(to)
      .split(",")
      .map((a) => a.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address } }));

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
  } catch (e) {
    return res.status(502).json({ ok: false, error: "Email service error: " + e.message });
  }
}
