// Vercel serverless function — emails a teammate a "create your login" invite
// using Supabase Auth's admin invite (sends Supabase's Invite email template
// with a magic link). Runs server-side because it needs the SERVICE ROLE key,
// which must never be exposed to the browser. No Anthropic / Microsoft needed.
//
// Required environment variables (Vercel → Settings → Environment Variables):
//   SUPABASE_SERVICE_ROLE_KEY — your Supabase project's service_role key
//   (the project URL is reused from VITE_SUPABASE_URL, already configured)
import { createClient } from "@supabase/supabase-js";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ ok: false, error: "POST only" });

  const SUPABASE_URL = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
  const SERVICE_ROLE = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!SUPABASE_URL || !SERVICE_ROLE) {
    return res.status(500).json({
      ok: false,
      error: "Invites aren't set up yet — add SUPABASE_SERVICE_ROLE_KEY in Vercel, then redeploy.",
    });
  }

  const { email, redirectTo } = req.body || {};
  if (!email) return res.status(400).json({ ok: false, error: "Missing 'email'." });

  try {
    const supabase = createClient(SUPABASE_URL, SERVICE_ROLE, {
      auth: { autoRefreshToken: false, persistSession: false },
    });
    const { error } = await supabase.auth.admin.inviteUserByEmail(
      String(email).toLowerCase().trim(),
      redirectTo ? { redirectTo } : undefined
    );
    if (error) return res.status(502).json({ ok: false, error: error.message });
    return res.status(200).json({ ok: true });
  } catch (e) {
    return res.status(502).json({ ok: false, error: e.message });
  }
}
