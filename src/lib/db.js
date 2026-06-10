// src/lib/db.js — shared data layer for Pipeline
// Uses Supabase (shared with ClearBid) when configured; falls back to
// this browser's localStorage when env vars are missing so the app
// always runs.
import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;
export const supabase = url && key ? createClient(url, key) : null;
export const usingSharedDb = !!supabase;

const LS_KEY = "pipeline9-crm-v1";
const today = () => new Date().toISOString().slice(0, 10);
const rid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

export async function loadAll() {
  if (!supabase) {
    const v = localStorage.getItem(LS_KEY);
    return v ? JSON.parse(v) : null;
  }
  const [i, j, a, c, t, s] = await Promise.all([
    supabase.from("itbs").select("id,data"),
    supabase.from("jobs").select("id,data"),
    supabase.from("accounts").select("id,data"),
    supabase.from("contacts").select("id,data"),
    supabase.from("tasks").select("id,data"),
    supabase.from("app_settings").select("data").eq("id", "default").maybeSingle(),
  ]);
  if (i.error) throw new Error(i.error.message);
  if (j.error) throw new Error(j.error.message);
  return {
    itbs: (i.data || []).map((r) => r.data),
    jobs: (j.data || []).map((r) => r.data),
    accounts: (a.data || []).map((r) => r.data),
    contacts: (c.data || []).map((r) => r.data),
    tasks: (t.data || []).map((r) => r.data),
    settings: s.data?.data || {},
  };
}

export async function saveAll(data) {
  if (!supabase) {
    localStorage.setItem(LS_KEY, JSON.stringify(data));
    return;
  }
  const now = new Date().toISOString();
  const rows = (arr) => (arr || []).map((d) => ({ id: d.id, data: d, updated_at: now }));
  const ops = [supabase.from("app_settings").upsert({ id: "default", data: data.settings || {} })];
  for (const [table, arr] of [["itbs", data.itbs], ["jobs", data.jobs], ["accounts", data.accounts], ["contacts", data.contacts], ["tasks", data.tasks]]) {
    const r = rows(arr);
    if (r.length) ops.push(supabase.from(table).upsert(r));
  }
  const results = await Promise.all(ops);
  const err = results.find((r) => r.error);
  if (err) throw new Error(err.error.message);
}

/* Pull ClearBid estimates and advance matching ITBs.
   - estimate status 'created' -> ITB moves to estimate_created
   - estimate status 'sent'    -> ITB moves to estimate_sent
   - line items are attached so winning the bid pre-fills the job's
     Estimate tab (and ARR) without re-reading any PDF. */
const STAGE_ORDER = ["new", "docs_requested", "docs_received", "estimate_created", "estimate_sent", "won", "lost"];

export async function syncClearBidEstimates(data) {
  if (!supabase) return data;
  const { data: ests, error } = await supabase.from("estimates").select("*");
  if (error || !ests?.length) return data;

  let advanced = false;

  // Estimates published from ClearBid for brand-new bids (no ITB yet)
  for (const est of ests.filter((e) => !e.itb_id && e.project?.projectName)) {
    const id = "cb-" + rid();
    const p = est.project;
    data.itbs = [{
      id, name: p.projectName, client: p.client || "", contact: p.contact || "", phone: p.phone || "",
      address: p.address || "", value: Number(est.total) || 0,
      stage: est.status === "sent" ? "estimate_sent" : "estimate_created",
      createdAt: today(), lastTouched: today(),
      clearbidEstimate: { items: est.items || [], total: est.total, pdfUrl: est.pdf_url || "" },
      notes: [{ id: rid(), date: today(), text: "📐 ITB created automatically from a ClearBid estimate." }],
      history: [{ stage: "new", date: today() }, { stage: est.status === "sent" ? "estimate_sent" : "estimate_created", date: today() }],
    }, ...(data.itbs || [])];
    est.itb_id = id;
    try { await supabase.from("estimates").update({ itb_id: id }).eq("id", est.id); } catch { /* retry next load */ }
    advanced = true;
  }

  data.itbs = (data.itbs || []).map((itb) => {
    const est = ests.find((e) => e.itb_id === itb.id);
    if (!est) return itb;
    const upd = {
      ...itb,
      clearbidEstimate: { items: est.items || [], total: est.total, pdfUrl: est.pdf_url || "" },
    };
    const target = est.status === "sent" ? "estimate_sent" : "estimate_created";
    if (STAGE_ORDER.indexOf(itb.stage) < STAGE_ORDER.indexOf(target)) {
      upd.stage = target;
      upd.value = Number(est.total) || itb.value;
      upd.lastTouched = today();
      upd.history = [...(itb.history || []), { stage: target, date: today() }];
      upd.notes = [
        { id: rid(), date: today(), text: `📐 ClearBid estimate attached automatically${est.total ? " — $" + Number(est.total).toLocaleString() : ""}. Stage advanced to ${target.replace("_", " ")}.` },
        ...(itb.notes || []),
      ];
      advanced = true;
    }
    return upd;
  });

  if (advanced) {
    try { await saveAll(data); } catch { /* will retry on next save */ }
  }
  return data;
}

/* ===================== AUTH + SEATS ===================== */
export async function getSession() {
  if (!supabase) return null;
  const { data } = await supabase.auth.getSession();
  return data.session;
}
export function onAuth(cb) {
  if (!supabase) return () => {};
  const { data } = supabase.auth.onAuthStateChange((_e, s) => cb(s));
  return () => data.subscription.unsubscribe();
}
export async function authSignIn(email, password) {
  const { error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) throw new Error(error.message);
}
export async function authSignUp(email, password) {
  const { data, error } = await supabase.auth.signUp({ email, password });
  if (error) throw new Error(error.message);
  return data;
}
export async function authSignOut() { await supabase?.auth.signOut(); }

export async function listSeats() {
  if (!supabase) return [];
  const { data, error } = await supabase.from("seats").select("*").order("created_at");
  if (error) throw new Error(error.message);
  return data || [];
}
export async function addSeat(email, role = "member") {
  const { error } = await supabase.from("seats").insert({ email: (email || "").toLowerCase().trim(), role });
  if (error) throw new Error(error.message);
}
export async function removeSeat(id) {
  const { error } = await supabase.from("seats").delete().eq("id", id);
  if (error) throw new Error(error.message);
}
