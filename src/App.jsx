import { useState, useEffect, useRef } from "react";
import { loadAll, saveAll, syncClearBidEstimates, usingSharedDb, getSession, onAuth, authSignIn, authSignUp, authSignOut, listSeats, addSeat, removeSeat } from "./lib/db";

/* ============================================================
   PIPELINE CRM — ITB Pipeline + Job Tracking for a plumbing co.
   Pipeline: New ITB → Docs Requested → Docs Received →
             Add Labor → Estimate Created → Estimate Sent → Won / Lost
   Won ITBs convert to Jobs (opportunities) with materials,
   labor, notes, and material invoices.
   ============================================================ */

const STAGES = [
  { id: "new", label: "New ITB", short: "NEW" },
  { id: "docs_requested", label: "Docs Requested", short: "DOC REQ" },
  { id: "docs_received", label: "Docs Received", short: "DOC RCV" },
  { id: "add_labor", label: "Add Labor", short: "LABOR" },
  { id: "estimate_created", label: "Estimate Created", short: "EST MADE" },
  { id: "estimate_sent", label: "Estimate Sent", short: "EST SENT" },
];

// When an ITB advances from "Docs Received" into "Add Labor", notify this address.
const LABOR_NOTIFY_EMAIL = "davidc@coastalplumbingswfl.com";

/* Turn raw email failure text into a short, human, actionable message. */
function friendlyEmailError(text) {
  const t = (text || "").toLowerCase();
  if (t.includes("isn't set up") || t.includes("ms_tenant") || t.includes("ms_client") || t.includes("graph env"))
    return "Email isn't set up yet — an admin needs to add the Microsoft Graph keys (MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER) in Vercel, then redeploy.";
  if (t.includes("auth failed") || t.includes("unauthorized") || t.includes("invalid_client") || t.includes("aadsts"))
    return "Email authorization failed — check the Azure app credentials and that Mail.Send admin consent was granted.";
  if (t.includes("graph send failed") || t.includes("forbidden") || t.includes("accessdenied") || t.includes("mailboxnotenabled"))
    return "Outlook rejected the send — verify the sender mailbox (MS_SENDER) exists and the app has Mail.Send permission.";
  if (t.includes("upstream") || t.includes("service error") || t.includes("502") || t.includes("failed to fetch"))
    return "Couldn't reach the email service — please try again in a moment.";
  return (text || "").trim().slice(0, 220) || "Send failed — check the email configuration.";
}

/* Send an email through the Microsoft Graph (Outlook) function at /api/send-email.
   No LLM / Anthropic billing. Reused by the labor-ready notice, team pings,
   task notifications, and stale-bid alerts. `to` may be a comma-separated list.
   Returns { ok, text }. */
async function sendOutlookEmail({ to, subject, body }) {
  try {
    const response = await fetch("/api/send-email", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ to, subject, body }),
    });
    const res = await response.json().catch(() => ({}));
    if (res.ok) return { ok: true, text: "SENT" };
    return { ok: false, text: friendlyEmailError(res.error || "") };
  } catch {
    return { ok: false, text: friendlyEmailError("failed to fetch") };
  }
}

/* Map raw Supabase invite errors to short, human guidance. */
function friendlyInviteError(text) {
  const t = (text || "").toLowerCase();
  if (t.includes("aren't set up") || t.includes("service_role") || t.includes("supabase_service"))
    return "Invites aren't set up yet — add SUPABASE_SERVICE_ROLE_KEY in Vercel, then redeploy.";
  if (t.includes("already") && (t.includes("registered") || t.includes("exist")))
    return "They already have a login — they can just sign in.";
  if (t.includes("rate") || t.includes("limit") || t.includes("too many"))
    return "Email rate limit hit — Supabase free tier sends only a few invites per hour. Wait a bit, or add custom SMTP in Supabase.";
  if (t.includes("redirect") || t.includes("not allowed"))
    return "Add your app URL under Supabase → Authentication → URL Configuration (Redirect URLs), then retry.";
  return (text || "").trim().slice(0, 200) || "Invite failed — check Supabase settings.";
}

/* Email a teammate a "create your login" invite via /api/invite (Supabase). */
async function sendInvite(email) {
  try {
    const response = await fetch("/api/invite", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, redirectTo: typeof window !== "undefined" ? window.location.origin : undefined }),
    });
    const res = await response.json().catch(() => ({}));
    return res.ok ? { ok: true } : { ok: false, error: friendlyInviteError(res.error || "") };
  } catch {
    return { ok: false, error: "Couldn't reach the invite service. Try again." };
  }
}

const STORAGE_KEY = "pipeline9-crm-v1";
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const today = () => new Date().toISOString().slice(0, 10);
const fmt = (n) =>
  "$" + (Number(n) || 0).toLocaleString("en-US", { maximumFractionDigits: 0 });
const fmtDate = (d) => {
  if (!d) return "—";
  const dt = new Date(d + (d.length === 10 ? "T12:00:00" : ""));
  return dt.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};

/* recurring-revenue helpers */
const ANNUAL_X = { one_time: 0, monthly: 12, quarterly: 4, annual: 1 };
const annualized = (it) => (Number(it.amount) || 0) * (ANNUAL_X[it.period] ?? 0);
const jobArr = (j) => (j.estimate?.items || []).reduce((s, it) => s + annualized(it), 0);
const jobOneTime = (j) =>
  (j.estimate?.items || [])
    .filter((it) => !it.period || it.period === "one_time")
    .reduce((s, it) => s + (Number(it.amount) || 0), 0);

const SEED = {
  itbs: [
    {
      id: "seed1", name: "Riverside Apartments Repipe", client: "Hargrove Property Mgmt",
      contact: "Dana Hargrove", phone: "(941) 555-0142", address: "1200 Riverside Dr, Bradenton FL",
      value: 84000, stage: "estimate_sent", createdAt: today(), bidDue: "",
      notes: [{ id: "n1", text: "GC wants estimate broken out by building.", date: today() }],
      history: [{ stage: "new", date: today() }],
    },
    {
      id: "seed2", name: "Marina Grill Grease Trap + Rough-In", client: "Coastal Restaurants LLC",
      contact: "Mike Torres", phone: "(941) 555-0188", address: "415 Bayfront Ave, Palmetto FL",
      value: 22500, stage: "docs_requested", createdAt: today(), bidDue: "",
      notes: [], history: [{ stage: "new", date: today() }],
    },
  ],
  jobs: [],
  accounts: [],
  contacts: [],
  tasks: [],
};

export default function App() {
  const [data, setData] = useState(null);
  const [view, setView] = useState("pipeline"); // pipeline | jobs
  const [selectedItb, setSelectedItb] = useState(null);
  const [selectedJob, setSelectedJob] = useState(null);
  const [showNewItb, setShowNewItb] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [session, setSession] = useState(null);
  const [authReady, setAuthReady] = useState(!usingSharedDb);
  const saveTimer = useRef(null);

  useEffect(() => {
    if (!usingSharedDb) return;
    getSession().then((s) => { setSession(s); setAuthReady(true); });
    return onAuth((s) => setSession(s));
  }, []);

  /* ---------- persistence ---------- */
  useEffect(() => {
    if (!authReady || (usingSharedDb && !session)) return;
    (async () => {
      let loaded = null;
      try {
        loaded = await loadAll();
      } catch {
        setError("Couldn't load the shared database. If you just signed up, your email may not have a seat yet — ask the owner to add you on the Team tab.");
        setTimeout(() => setError(""), 9000);
      }
      if (!loaded || (!loaded.itbs?.length && !loaded.jobs?.length)) loaded = { ...SEED, settings: loaded?.settings || {} };
      loaded.settings = { alertEmail: "", staleDays: 5, ...(loaded.settings || {}) };
      for (const k of ["accounts", "contacts", "tasks"]) loaded[k] = loaded[k] || [];
      try {
        loaded = await syncClearBidEstimates(loaded);
      } catch { /* estimates table not reachable — non-fatal */ }
      setData(loaded);
    })();
  }, [authReady, session]);

  const persist = (next) => {
    setData(next);
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        await saveAll(next);
      } catch {
        setError("Couldn't save to the shared database — changes may not persist.");
        setTimeout(() => setError(""), 4000);
      }
    }, 400);
  };

  /* ---------- ITB actions ---------- */
  const addItb = (form) => {
    const itb = {
      id: uid(), ...form, value: Number(form.value) || 0, stage: "new",
      createdAt: today(), lastTouched: today(), notes: [], history: [{ stage: "new", date: today() }],
    };
    persist({ ...data, itbs: [itb, ...data.itbs] });
    setShowNewItb(false);
  };

  const updateItb = (id, patch) =>
    persist({ ...data, itbs: data.itbs.map((i) => (i.id === id ? { ...i, ...patch, lastTouched: today() } : i)) });

  const moveStage = (itb, dir) => {
    const idx = STAGES.findIndex((s) => s.id === itb.stage);
    const next = STAGES[idx + dir];
    if (!next) return;
    updateItb(itb.id, {
      stage: next.id,
      history: [...itb.history, { stage: next.id, date: today() }],
    });
    // Notify when a bid crosses from Docs Received into Add Labor.
    if (itb.stage === "docs_received" && next.id === "add_labor") notifyLaborMove(itb);
  };

  // Best-effort email via the Outlook/Microsoft 365 MCP tools (same path as alerts).
  // Fires when an ITB enters the Add Labor stage; never blocks the stage move.
  const notifyLaborMove = async (itb) => {
    setNotice(`Notifying ${LABOR_NOTIFY_EMAIL} that "${itb.name}" is ready for labor…`);
    const body =
      `${itb.name} for ${itb.client} just moved from Docs Received to Add Labor in the Pipeline CRM and is ready for labor to be added.\n\n` +
      `Estimated value: ${fmt(itb.value)}\n` +
      (itb.address ? `Address: ${itb.address}\n` : "") +
      (itb.contact ? `Contact: ${itb.contact}${itb.phone ? " · " + itb.phone : ""}\n` : "") +
      `\n— Pipeline CRM`;
    try {
      const { ok, text } = await sendOutlookEmail({ to: LABOR_NOTIFY_EMAIL, subject: `Ready for Labor — ${itb.name} (${itb.client})`, body });
      if (ok) {
        setNotice(`✓ Labor-ready notice emailed to ${LABOR_NOTIFY_EMAIL} — ${fmtDate(today())}`);
      } else {
        setNotice("");
        setError(text.trim().slice(0, 220) || "Couldn't email the labor-ready notice — check the Microsoft 365 connection.");
      }
    } catch {
      setNotice("");
      setError("Couldn't reach the email service for the labor-ready notice.");
    }
  };

  const markLost = (itb) => {
    updateItb(itb.id, { stage: "lost", history: [...itb.history, { stage: "lost", date: today() }] });
    setSelectedItb(null);
  };

  const winItb = (itb) => {
    const job = {
      id: uid(), itbId: itb.id, name: itb.name, client: itb.client, contact: itb.contact,
      phone: itb.phone, address: itb.address, value: itb.value, status: "active",
      wonAt: today(), startDate: "", materials: [], labor: [], invoices: [],
      estimate: itb.clearbidEstimate?.items?.length
        ? {
            items: itb.clearbidEstimate.items.map((it) => ({
              id: uid(),
              description: it.description || it.name || "Line item",
              amount: Number(it.amount) || 0,
              period: ["monthly", "quarterly", "annual"].includes(it.period) ? it.period : "one_time",
            })),
            source: "clearbid", readAt: today(),
          }
        : undefined,
      notes: [
        ...(itb.clearbidEstimate?.items?.length
          ? [{ id: uid(), date: today(), text: "📐 Estimate tab pre-filled from the ClearBid estimate — ARR computed from its recurring line items." }]
          : []),
        ...(itb.notes || []),
      ],
    };
    persist({
      ...data,
      itbs: data.itbs.map((i) =>
        i.id === itb.id ? { ...i, stage: "won", history: [...i.history, { stage: "won", date: today() }] } : i
      ),
      jobs: [job, ...data.jobs],
    });
    setSelectedItb(null);
    setSelectedJob(job.id);
    setView("jobs");
  };

  const updateJob = (id, patch) =>
    persist({ ...data, jobs: data.jobs.map((j) => (j.id === id ? { ...j, ...patch } : j)) });

  const addRec = (kind, rec) => persist({ ...data, [kind]: [{ id: uid(), createdAt: today(), ...rec }, ...data[kind]] });
  const updateRec = (kind, id, patch) => persist({ ...data, [kind]: data[kind].map((r) => (r.id === id ? { ...r, ...patch } : r)) });
  const delRec = (kind, id) => persist({ ...data, [kind]: data[kind].filter((r) => r.id !== id) });

  if (!authReady)
    return (
      <Shell>
        <div style={{ padding: 60, textAlign: "center", color: "var(--ink-dim)" }}>Checking sign-in…</div>
      </Shell>
    );
  if (usingSharedDb && !session) return <Shell><Login /></Shell>;
  if (!data)
    return (
      <Shell>
        <div style={{ padding: 60, textAlign: "center", color: "var(--ink-dim)" }}>Loading pipeline…</div>
      </Shell>
    );

  const active = data.itbs.filter((i) => !["won", "lost"].includes(i.stage));
  const wonCount = data.itbs.filter((i) => i.stage === "won").length;
  const lostCount = data.itbs.filter((i) => i.stage === "lost").length;
  const pipelineValue = active.reduce((s, i) => s + (i.value || 0), 0);
  const winRate = wonCount + lostCount ? Math.round((wonCount / (wonCount + lostCount)) * 100) : null;
  const job = data.jobs.find((j) => j.id === selectedJob);
  const itb = data.itbs.find((i) => i.id === selectedItb);

  const thisYear = String(new Date().getFullYear());
  const totalArr = data.jobs.filter((j) => j.status !== "cancelled").reduce((s, j) => s + jobArr(j), 0);
  const newArrYtd = data.jobs.filter((j) => (j.wonAt || "").startsWith(thisYear) && j.status !== "cancelled").reduce((s, j) => s + jobArr(j), 0);
  const churnedArrYtd = data.jobs.filter((j) => j.status === "cancelled" && (j.cancelledAt || "").startsWith(thisYear)).reduce((s, j) => s + jobArr(j), 0);
  const narrYtd = newArrYtd - churnedArrYtd;

  const daysSince = (d) => Math.floor((Date.now() - new Date((d || today()) + "T12:00:00")) / 86400000);
  const staleItbs = active
    .map((i) => ({ ...i, idleDays: daysSince(i.lastTouched || i.createdAt) }))
    .filter((i) => i.idleDays >= (data.settings?.staleDays ?? 5))
    .sort((a, b) => b.idleDays - a.idleDays);

  return (
    <Shell>
      {/* ===== Header ===== */}
      <header>
        <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "9px 20px", background: "#ffffff", borderBottom: "1px solid var(--line)" }}>
          <PipeMark size={32} />
          <span style={{ fontFamily: "var(--display)", fontSize: 18, color: "var(--navy)", fontWeight: 700, letterSpacing: -0.2 }}>
            Pipeline
          </span>
          <span style={{ fontFamily: "var(--mono)", fontSize: 10, letterSpacing: 1.8, color: "var(--ink-dim)", textTransform: "uppercase", paddingTop: 2 }}>
            ITB · Estimate · Build{usingSharedDb ? " · SHARED DB" : " · LOCAL MODE"}
          </span>
          <div style={{ marginLeft: "auto", display: "flex", gap: 10, alignItems: "center" }}>
            {session && (
              <>
                <span style={{ fontSize: 12, color: "var(--ink-dim)" }}>{session.user?.email}</span>
                <button onClick={() => authSignOut()}
                  style={{ background: "#fff", border: "1px solid #c9c9c9", color: "var(--copper)", borderRadius: 4, padding: "6px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                  Sign out
                </button>
              </>
            )}
            <GlobalSearch data={data} go={(v, id, k) => {
              setView(v); setSelectedJob(null); setSelectedItb(null);
              if (k === "itb") setSelectedItb(id);
              if (k === "job") { setView("jobs"); setSelectedJob(id); }
            }} />
          </div>
        </div>
        <nav style={{ display: "flex", gap: 0, padding: "0 12px", background: "#ffffff", borderBottom: "1px solid var(--line)", boxShadow: "var(--shadow-sm)" }}>
          {[
            ["dashboard", "Home"],
            ["pipeline", `Pipeline (${active.length})`],
            ["jobs", `Opportunities (${data.jobs.length})`],
            ["accounts", `Accounts (${data.accounts.length})`],
            ["contacts", `Contacts (${data.contacts.length})`],
            ["tasks", `Tasks (${data.tasks.filter((t) => !t.done).length})`],
            ["team", "Team"],
          ].map(([v, label]) => (
            <button key={v} onClick={() => { setView(v); setSelectedJob(null); }}
              className={"p9-tab" + (view === v ? " on" : "")}
              style={{
                fontFamily: "var(--body)", fontSize: 13.5, padding: "12px 16px",
                background: "transparent",
                color: view === v ? "var(--copper)" : "#3e3e3c",
                border: "none",
                cursor: "pointer", fontWeight: view === v ? 700 : 400,
              }}>
              {label}
            </button>
          ))}
        </nav>
      </header>

      {/* ===== Stat strip ===== */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, margin: "20px 24px" }}>
        <Stat label="Open ITBs" value={active.length} />
        <Stat label="Pipeline value" value={fmt(pipelineValue)} />
        <Stat label="Win rate" value={winRate === null ? "—" : winRate + "%"} />
        <Stat label="Active opps" value={data.jobs.filter((j) => j.status === "active").length} accent />
        <Stat label="ARR" value={fmt(totalArr)} accent />
        <Stat label={"Net new ARR " + thisYear} value={(narrYtd < 0 ? "-" : "") + fmt(Math.abs(narrYtd))} accent={narrYtd >= 0} />
      </div>

      {error && (
        <div style={{ margin: "0 24px 12px", padding: "10px 14px", background: "#fdebe8", color: "#b03a2e", fontSize: 13, display: "flex", alignItems: "center", gap: 10, borderRadius: "var(--radius)", border: "1px solid #f4cdca" }}>
          <span style={{ flex: 1 }}>{error}</span>
          <button onClick={() => setError("")} aria-label="Dismiss"
            style={{ background: "none", border: "none", color: "inherit", cursor: "pointer", fontSize: 18, lineHeight: 1, padding: "0 2px" }}>×</button>
        </div>
      )}

      {notice && (
        <div style={{ margin: "0 24px 12px", padding: "10px 14px", background: "#e8f1fb", color: "#1f5fa6", fontSize: 13, display: "flex", alignItems: "center", gap: 10, borderRadius: "var(--radius)", border: "1px solid #c9def5" }}>
          <span style={{ flex: 1 }}>{notice}</span>
          <button onClick={() => setNotice("")} aria-label="Dismiss"
            style={{ background: "none", border: "none", color: "inherit", cursor: "pointer", fontSize: 18, lineHeight: 1, padding: "0 2px" }}>×</button>
        </div>
      )}

      <AlertsBar
        staleItbs={staleItbs} settings={data.settings}
        onSettings={(s) => persist({ ...data, settings: { ...data.settings, ...s } })}
        onOpenItb={(id) => setSelectedItb(id)}
      />

      {/* ===== Views ===== */}
      {view === "dashboard" && (
        <Dashboard data={data} active={active} totalArr={totalArr} narrYtd={narrYtd} thisYear={thisYear}
          winRate={winRate} pipelineValue={pipelineValue}
          go={(v) => { setView(v); setSelectedJob(null); }}
          onOpenItb={(id) => setSelectedItb(id)}
          onToggleTask={(id, done) => updateRec("tasks", id, { done })} />
      )}
      {view === "accounts" && (
        <Accounts data={data} onAdd={(r) => addRec("accounts", r)} onDel={(id) => delRec("accounts", id)}
          onOpenItb={(id) => setSelectedItb(id)} />
      )}
      {view === "contacts" && (
        <Contacts data={data} onAdd={(r) => addRec("contacts", r)} onDel={(id) => delRec("contacts", id)} />
      )}
      {view === "tasks" && (
        <Tasks data={data} onAdd={(r) => addRec("tasks", r)} onDel={(id) => delRec("tasks", id)}
          onToggle={(id, done) => updateRec("tasks", id, { done })}
          onAssign={(id, assignee) => updateRec("tasks", id, { assignee })} />
      )}
      {view === "team" && <Team session={session} />}
      {view === "pipeline" && (
        <Pipeline
          data={data} active={active} onSelect={(id) => setSelectedItb(id)}
          onMove={moveStage} onNew={() => setShowNewItb(true)}
        />
      )}
      {view === "jobs" && !job && (
        <JobsList jobs={data.jobs} onOpen={(id) => setSelectedJob(id)} />
      )}
      {view === "jobs" && job && (
        <JobDetail job={job} onBack={() => setSelectedJob(null)} onUpdate={(p) => updateJob(job.id, p)}
          settings={data.settings} onSettings={(s) => persist({ ...data, settings: { ...data.settings, ...s } })} />
      )}

      {/* ===== Overlays ===== */}
      {itb && (
        <ItbDrawer itb={itb} onClose={() => setSelectedItb(null)}
          onUpdate={(p) => updateItb(itb.id, p)} onMove={(d) => moveStage(itb, d)}
          onWin={() => winItb(itb)} onLose={() => markLost(itb)} />
      )}
      {showNewItb && <NewItbModal onSave={addItb} onClose={() => setShowNewItb(false)} />}
    </Shell>
  );
}

/* ============================================================ */
function Shell({ children }) {
  return (
    <div className="p9-root" style={{ minHeight: "100vh", color: "var(--ink)", fontFamily: "var(--body)" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Spline+Sans+Mono:wght@400;500;600&display=swap');
        /* ── Salesforce Lightning Design System (SLDS) theme tokens ── */
        :root {
          --bg: #f3f3f3;          /* SLDS app background */
          --panel: #ffffff;
          --panel2: #fafaf9;      /* SLDS neutral surface / table header */
          --line: #dddbda;        /* SLDS color border */
          --ink: #181818;         /* SLDS text default */
          --ink-dim: #706e6b;     /* SLDS text weak */
          --navy: #032d60;
          --navy2: #014486;
          --copper: #0176d3;      /* SLDS brand */
          --copper-hi: #1b96ff;   /* SLDS brand light / focus */
          --copper-dim: #aacbff;
          --green: #2e844a;       /* SLDS success */
          --red: #ba0517;         /* SLDS error / destructive */
          --display: 'Inter', 'Salesforce Sans', Arial, sans-serif;
          --body: 'Inter', 'Salesforce Sans', Arial, sans-serif;
          --mono: 'Spline Sans Mono', monospace;
          --radius: 0.25rem;      /* SLDS border-radius-medium */
          --shadow-sm: 0 2px 2px 0 rgba(0,0,0,.10);
          --shadow-md: 0 2px 3px 0 rgba(0,0,0,.16);
        }
        * { box-sizing: border-box; }
        .p9-root { background: var(--bg); }
        input, select, textarea {
          background: #ffffff; border: 1px solid var(--line); color: var(--ink);
          padding: 8px 11px; font-family: var(--body); font-size: 13px; width: 100%;
          border-radius: var(--radius); transition: border-color .1s linear, box-shadow .1s linear;
        }
        input:focus, select:focus, textarea:focus {
          outline: none; border-color: var(--copper-hi);
          box-shadow: 0 0 3px 0 var(--copper);
        }
        ::placeholder { color: #969492; }
        button { font-family: var(--body); }
        /* SLDS buttons are rectangular with a 0.25rem radius — never pills. */
        .p9-btn { border-radius: var(--radius) !important; }
        ::-webkit-scrollbar { height: 10px; width: 10px; }
        ::-webkit-scrollbar-thumb { background: #c9c7c5; border-radius: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }

        @keyframes p9rise { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        .p9-rise { animation: p9rise .2s ease both; }
        .p9-card { transition: box-shadow .1s linear, border-color .1s linear; border-radius: var(--radius); }
        .p9-card:hover { box-shadow: var(--shadow-md); border-color: #c9c7c5 !important; }
        .p9-btn { transition: background .1s linear, box-shadow .1s linear; }
        .p9-btn:hover:not(:disabled) { filter: brightness(0.96); }
        .p9-btn:active:not(:disabled) { transform: translateY(1px); }
        .p9-tab { position: relative; transition: color .1s linear; }
        .p9-tab::after {
          content: ""; position: absolute; left: 0; right: 0; bottom: -1px; height: 3px;
          background: var(--copper);
          transform: scaleX(0); transition: transform .1s linear;
        }
        .p9-tab.on::after { transform: scaleX(1); }
        table.p9-tbl tbody tr:nth-child(even) { background: #fafaf9; }
        table.p9-tbl tbody tr:hover { background: #f3f9fe; }
        table.p9-tbl thead th { background: #fafaf9; position: sticky; top: 0; }
        .p9-scroll::-webkit-scrollbar { height: 8px; }
        @media (prefers-reduced-motion: reduce) { .p9-rise, .p9-card, .p9-btn, .p9-tab::after { animation: none; transition: none; } }
      `}</style>
      {children}
    </div>
  );
}

const Stat = ({ label, value, accent }) => (
  <div className="p9-card p9-rise" style={{ padding: "14px 18px 13px", background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 4, boxShadow: "var(--shadow-sm)", borderTop: "3px solid " + (accent ? "var(--copper)" : "#c9d6e6") }}>
    <div style={{ fontFamily: "var(--mono)", letterSpacing: 1.5, fontSize: 10, color: "var(--ink-dim)", textTransform: "uppercase" }}>
      {label}
    </div>
    <div style={{ fontFamily: "var(--display)", fontWeight: 600, fontSize: 22, marginTop: 3, color: accent ? "var(--copper)" : "var(--ink)", fontVariantNumeric: "tabular-nums" }}>
      {value}
    </div>
  </div>
);

/* ============== ALERTS / OUTLOOK ============== */
function AlertsBar({ staleItbs, settings, onSettings, onOpenItb }) {
  const [open, setOpen] = useState(false);
  const [status, setStatus] = useState(null); // {type, msg}
  const count = staleItbs.length;

  const sendAlert = async () => {
    if (!settings.alertEmail?.trim()) {
      setStatus({ type: "error", msg: "Enter a recipient email first." });
      return;
    }
    setStatus({ type: "sending", msg: "Sending via Outlook…" });
    const lines = staleItbs
      .map((i) => {
        const st = STAGES.find((s) => s.id === i.stage)?.label || i.stage;
        return `• ${i.name} — ${i.client} | stage: ${st} | idle ${i.idleDays} day(s) | est. ${fmt(i.value)}`;
      })
      .join("\n");
    try {
      const { ok, text } = await sendOutlookEmail({
        to: settings.alertEmail,
        subject: `Pipeline Alert — ${count} ITB${count === 1 ? "" : "s"} need attention`,
        body: `These bids have had no activity for ${settings.staleDays}+ days:\n\n${lines}\n\nFollow up before they go cold.\n\n— Pipeline CRM`,
      });
      if (ok) {
        setStatus({ type: "ok", msg: "Alert emailed via Outlook ✓" });
      } else {
        setStatus({ type: "error", msg: text });
      }
    } catch {
      setStatus({ type: "error", msg: "Couldn't reach the email service. Try again." });
    }
  };

  return (
    <div style={{ margin: "0 24px 16px", border: "1px solid " + (count ? "#e6b94d" : "var(--line)"), background: count ? "#fdf6e3" : "var(--panel)", borderRadius: 4 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px" }}>
        <span style={{ fontFamily: "var(--display)", letterSpacing: 2, fontSize: 13, textTransform: "uppercase", color: count ? "#8a6d1a" : "var(--ink-dim)" }}>
          {count ? `⚠ ${count} ITB${count === 1 ? "" : "s"} untouched for ${settings.staleDays}+ days` : "✓ No stale ITBs — pipeline is moving"}
        </span>
        <button onClick={() => setOpen(!open)}
          style={{ marginLeft: "auto", background: "none", border: "1px solid var(--line)", color: "var(--ink-dim)", padding: "5px 14px", fontSize: 12, cursor: "pointer", fontFamily: "var(--display)", letterSpacing: 1.5 }}>
          {open ? "HIDE" : "ALERT SETTINGS"}
        </button>
      </div>

      {open && (
        <div style={{ padding: "0 16px 16px" }}>
          {staleItbs.map((i) => (
            <div key={i.id} onClick={() => onOpenItb(i.id)}
              style={{ display: "flex", gap: 14, padding: "8px 12px", background: "var(--panel)", border: "1px solid var(--line)", marginBottom: 6, cursor: "pointer", fontSize: 13, alignItems: "center" }}>
              <b>{i.name}</b>
              <span style={{ color: "var(--ink-dim)" }}>{i.client}</span>
              <span style={{ color: "var(--ink-dim)" }}>{STAGES.find((s) => s.id === i.stage)?.label}</span>
              <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", color: "#b07d0e" }}>{i.idleDays}d idle</span>
              <span style={{ fontFamily: "var(--mono)", color: "var(--copper)" }}>{fmt(i.value)}</span>
            </div>
          ))}
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginTop: 10 }}>
            <input style={{ flex: 1, minWidth: 220 }} type="email" placeholder="alerts@yourcompany.com"
              value={settings.alertEmail} onChange={(e) => onSettings({ alertEmail: e.target.value })} />
            <span style={{ fontSize: 12, color: "var(--ink-dim)", whiteSpace: "nowrap" }}>Stale after</span>
            <input style={{ width: 64 }} type="number" min="1" value={settings.staleDays}
              onChange={(e) => onSettings({ staleDays: Math.max(1, Number(e.target.value) || 5) })} />
            <span style={{ fontSize: 12, color: "var(--ink-dim)" }}>days</span>
            <Btn disabled={!count || status?.type === "sending"} onClick={sendAlert}>
              {status?.type === "sending" ? "Sending…" : "Send Outlook Alert"}
            </Btn>
          </div>
          {status && status.type !== "sending" && (
            <div style={{ marginTop: 8, fontSize: 13, color: status.type === "ok" ? "var(--green)" : "var(--red)" }}>
              {status.msg}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ============== PIPELINE BOARD ============== */
function Pipeline({ data, active, onSelect, onMove, onNew }) {
  return (
    <div style={{ padding: "0 24px 24px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, marginTop: 4 }}>
        <h2 style={{ fontFamily: "var(--display)", fontSize: 19, fontWeight: 700, letterSpacing: -0.2, margin: 0, color: "var(--ink)", display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ width: 4, height: 20, background: "var(--copper)", borderRadius: 2, display: "inline-block" }} />Bid pipeline
        </h2>
        <button onClick={onNew}
          className="p9-btn" style={{ background: "var(--copper)", color: "#ffffff", border: "none", padding: "11px 24px", fontFamily: "var(--display)", fontSize: 14.5, letterSpacing: 0.8, cursor: "pointer", fontWeight: 700, borderRadius: 499 }}>
          + NEW ITB
        </button>
      </div>
      <div style={{ display: "flex", gap: 12, overflowX: "auto", alignItems: "flex-start", paddingBottom: 8 }}>
        {STAGES.map((stage, si) => {
          const cards = active.filter((i) => i.stage === stage.id);
          const colValue = cards.reduce((s, c) => s + (c.value || 0), 0);
          return (
            <div key={stage.id} className="p9-rise" style={{ minWidth: 252, flex: 1, background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 4, overflow: "hidden", boxShadow: "var(--shadow-sm)", animationDelay: si * 60 + "ms" }}>
              <div style={{ padding: "11px 14px", borderBottom: "1px solid var(--line)", display: "flex", alignItems: "center", gap: 8, background: "var(--panel2)" }}>
                <span style={{ width: 20, height: 20, borderRadius: 499, background: "var(--copper)", color: "#fff", fontFamily: "var(--mono)", fontSize: 10, display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 600, flexShrink: 0 }}>{si + 1}</span>
                <span style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 14, letterSpacing: 0.3 }}>{stage.label}</span>
                <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 10.5, color: cards.length ? "var(--copper)" : "var(--ink-dim)", background: cards.length ? "#e5f0fb" : "var(--panel2)", padding: "2px 8px", borderRadius: 499 }}>{cards.length} · {fmt(colValue)}</span>
              </div>
              <div style={{ padding: 10, display: "flex", flexDirection: "column", gap: 8, minHeight: 80 }}>
                {cards.length === 0 && (
                  <div style={{ color: "#90a4b8", fontSize: 12, textAlign: "center", padding: "18px 0" }}>empty</div>
                )}
                {cards.map((c) => (
                  <div key={c.id} onClick={() => onSelect(c.id)} className="p9-card"
                    style={{ background: "var(--panel)", border: "1px solid var(--line)", borderLeft: "3px solid var(--copper)", padding: "11px 12px", cursor: "pointer", borderRadius: 4, boxShadow: "var(--shadow-sm)" }}>
                    <div style={{ fontWeight: 600, fontSize: 13.5, lineHeight: 1.35, fontFamily: "var(--body)" }}>{c.name}</div>
                    <div style={{ fontSize: 12, color: "var(--ink-dim)", marginTop: 4, display: "flex", alignItems: "center", gap: 6 }}>
                      <span style={{ width: 18, height: 18, borderRadius: 499, background: "var(--panel2)", border: "1px solid var(--line)", fontSize: 9, display: "inline-flex", alignItems: "center", justifyContent: "center", fontFamily: "var(--mono)", color: "var(--copper)", flexShrink: 0 }}>{(c.client || "?").trim().slice(0, 1).toUpperCase()}</span>
                      {c.client}
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8, alignItems: "center" }}>
                      <span style={{ fontFamily: "var(--mono)", fontSize: 13, color: "var(--copper)" }}>{fmt(c.value)}</span>
                      <span style={{ display: "flex", gap: 4 }}>
                        {si > 0 && (
                          <MiniBtn onClick={(e) => { e.stopPropagation(); onMove(c, -1); }}>‹</MiniBtn>
                        )}
                        {si < STAGES.length - 1 && (
                          <MiniBtn onClick={(e) => { e.stopPropagation(); onMove(c, 1); }}>›</MiniBtn>
                        )}
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

const MiniBtn = ({ children, onClick }) => (
  <button onClick={onClick}
    style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--ink-dim)", width: 24, height: 24, cursor: "pointer", fontSize: 14, lineHeight: 1 }}>
    {children}
  </button>
);

/* ============== ITB DRAWER ============== */
function ItbDrawer({ itb, onClose, onUpdate, onMove, onWin, onLose }) {
  const [note, setNote] = useState("");
  const [edit, setEdit] = useState(false);
  const [form, setForm] = useState(itb);
  const setField = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const startEdit = () => { setForm(itb); setEdit(true); };
  const saveEdit = () => {
    onUpdate({
      name: form.name, client: form.client, contact: form.contact,
      phone: form.phone, email: form.email || "", address: form.address,
      value: Number(form.value) || 0, estimateDue: form.estimateDue || "",
    });
    setEdit(false);
  };
  const stageIdx = STAGES.findIndex((s) => s.id === itb.stage);
  const atEnd = itb.stage === "estimate_sent";

  return (
    <Overlay onClose={onClose}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontFamily: "var(--display)", fontSize: 12, letterSpacing: 3, color: "var(--copper)", textTransform: "uppercase" }}>
            Intent to Bid
          </div>
          {edit ? (
            <div style={{ display: "grid", gap: 8, marginTop: 8 }}>
              <input value={form.name} onChange={setField("name")} placeholder="Project name" />
              <div style={{ display: "flex", gap: 8 }}>
                <input value={form.client} onChange={setField("client")} placeholder="Client / GC" />
                <input value={form.contact} onChange={setField("contact")} placeholder="Contact" />
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <input value={form.phone} onChange={setField("phone")} placeholder="Phone" />
                <input type="email" value={form.email || ""} onChange={setField("email")} placeholder="Contact email" />
              </div>
              <input value={form.address} onChange={setField("address")} placeholder="Job site address" />
            </div>
          ) : (
            <>
              <h2 style={{ margin: "4px 0 2px", fontFamily: "var(--display)", fontSize: 21, fontWeight: 700, letterSpacing: -0.2 }}>{itb.name}</h2>
              <div style={{ color: "var(--ink-dim)", fontSize: 13 }}>
                {itb.client} · {itb.contact} · {itb.phone}
              </div>
              {itb.email ? <div style={{ color: "var(--ink-dim)", fontSize: 13 }}>{itb.email}</div> : null}
              <div style={{ color: "var(--ink-dim)", fontSize: 13 }}>{itb.address}</div>
            </>
          )}
        </div>
        <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
          {!edit && (
            <button onClick={startEdit} title="Edit details"
              style={{ background: "none", border: "1px solid var(--line)", color: "var(--copper)", width: 30, height: 30, cursor: "pointer", fontSize: 14, borderRadius: "var(--radius)" }}>✎</button>
          )}
          <CloseBtn onClick={onClose} />
        </div>
      </div>

      {/* stage path (chevrons) */}
      <div style={{ display: "flex", gap: 3, margin: "18px 0" }}>
        {STAGES.map((s, i) => {
          const done = i < stageIdx, current = i === stageIdx;
          return (
            <div key={s.id}
              style={{
                flex: 1, textAlign: "center", padding: "9px 6px 9px 16px",
                background: done ? "var(--green)" : current ? "var(--copper)" : "#e2e8f0",
                color: done || current ? "#ffffff" : "var(--ink-dim)",
                fontFamily: "var(--display)", fontSize: 11, letterSpacing: 1.2, fontWeight: 600,
                clipPath: i === 0
                  ? "polygon(0 0, calc(100% - 10px) 0, 100% 50%, calc(100% - 10px) 100%, 0 100%)"
                  : "polygon(0 0, calc(100% - 10px) 0, 100% 50%, calc(100% - 10px) 100%, 0 100%, 10px 50%)",
              }}>
              {done ? "✓ " : ""}{s.short}
            </div>
          );
        })}
      </div>

      {edit ? (
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="Estimate value ($)" style={{ flex: 1, minWidth: 150, marginTop: 0 }}>
            <input type="number" value={form.value} onChange={setField("value")} />
          </Field>
          <Field label="Estimate due date" style={{ flex: 1, minWidth: 150, marginTop: 0 }}>
            <input type="date" value={form.estimateDue || ""} onChange={setField("estimateDue")} />
          </Field>
          <div style={{ display: "flex", gap: 8 }}>
            <Btn onClick={saveEdit}>Save</Btn>
            <Btn ghost onClick={() => setEdit(false)}>Cancel</Btn>
          </div>
        </div>
      ) : (
        <>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            {stageIdx > 0 && <Btn ghost onClick={() => onMove(-1)}>← Back a stage</Btn>}
            {!atEnd && <Btn onClick={() => onMove(1)}>Advance → {STAGES[stageIdx + 1].label}</Btn>}
            {atEnd && <Btn color="var(--green)" onClick={onWin}>✓ Closed Won — Create Opportunity</Btn>}
            <Btn color="var(--red)" ghost onClick={onLose}>Mark Lost</Btn>
          </div>

          <div style={{ display: "flex", gap: 30, marginTop: 16, flexWrap: "wrap" }}>
            <div>
              <div style={{ fontFamily: "var(--display)", fontSize: 11, letterSpacing: 2, color: "var(--ink-dim)", textTransform: "uppercase", marginBottom: 4 }}>Estimate value</div>
              <div style={{ fontFamily: "var(--mono)", fontSize: 18, color: "var(--copper)" }}>{fmt(itb.value)}</div>
            </div>
            <div>
              <div style={{ fontFamily: "var(--display)", fontSize: 11, letterSpacing: 2, color: "var(--ink-dim)", textTransform: "uppercase", marginBottom: 4 }}>Estimate due</div>
              <div style={{ fontFamily: "var(--mono)", fontSize: 18, color: itb.estimateDue && itb.estimateDue < today() ? "var(--red)" : "var(--ink)" }}>
                {itb.estimateDue ? fmtDate(itb.estimateDue) : "—"}{itb.estimateDue && itb.estimateDue < today() ? " · OVERDUE" : ""}
              </div>
            </div>
          </div>
        </>
      )}

      <SectionTitle>Notes</SectionTitle>
      <div style={{ display: "flex", gap: 8 }}>
        <input placeholder="Add a note — call summary, doc status, GC asks…" value={note}
          onChange={(e) => setNote(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && note.trim()) {
              onUpdate({ notes: [{ id: uid(), text: note.trim(), date: today() }, ...itb.notes] });
              setNote("");
            }
          }} />
        <Btn onClick={() => {
          if (!note.trim()) return;
          onUpdate({ notes: [{ id: uid(), text: note.trim(), date: today() }, ...itb.notes] });
          setNote("");
        }}>Add</Btn>
      </div>
      {itb.notes.map((n) => (
        <div key={n.id} style={{ borderLeft: "2px solid var(--line)", padding: "6px 12px", marginTop: 8, fontSize: 13 }}>
          <span style={{ color: "var(--ink-dim)", fontFamily: "var(--mono)", fontSize: 11, marginRight: 10 }}>{fmtDate(n.date)}</span>
          {n.text}
        </div>
      ))}

      <SectionTitle>Stage history</SectionTitle>
      {[...itb.history].reverse().map((h, i) => (
        <div key={i} style={{ fontSize: 12, color: "var(--ink-dim)", fontFamily: "var(--mono)", padding: "2px 0" }}>
          {fmtDate(h.date)} — {(STAGES.find((s) => s.id === h.stage) || { label: h.stage.toUpperCase() }).label}
        </div>
      ))}
    </Overlay>
  );
}

/* ============== NEW ITB MODAL ============== */
function NewItbModal({ onSave, onClose }) {
  const [f, setF] = useState({ name: "", client: "", contact: "", phone: "", email: "", address: "", value: "", estimateDue: "" });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const valid = f.name.trim() && f.client.trim();

  return (
    <Overlay onClose={onClose} narrow>
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <h2 style={{ fontFamily: "var(--display)", fontSize: 19, fontWeight: 700, letterSpacing: 0, margin: 0 }}>New Intent to Bid</h2>
        <CloseBtn onClick={onClose} />
      </div>
      <Field label="Project name *"><input value={f.name} onChange={set("name")} placeholder="Sunset Plaza — 3-story rough-in" /></Field>
      <Field label="Client / GC *"><input value={f.client} onChange={set("client")} /></Field>
      <div style={{ display: "flex", gap: 12 }}>
        <Field label="Contact" style={{ flex: 1 }}><input value={f.contact} onChange={set("contact")} /></Field>
        <Field label="Phone" style={{ flex: 1 }}><input value={f.phone} onChange={set("phone")} /></Field>
      </div>
      <Field label="Contact email"><input type="email" value={f.email} onChange={set("email")} placeholder="contact@company.com" /></Field>
      <Field label="Job site address"><input value={f.address} onChange={set("address")} /></Field>
      <div style={{ display: "flex", gap: 12 }}>
        <Field label="Estimated value ($)" style={{ flex: 1 }}><input type="number" value={f.value} onChange={set("value")} /></Field>
        <Field label="Estimate due date" style={{ flex: 1 }}><input type="date" value={f.estimateDue} onChange={set("estimateDue")} /></Field>
      </div>
      <div style={{ marginTop: 18 }}>
        <Btn disabled={!valid} onClick={() => valid && onSave(f)}>Create ITB</Btn>
      </div>
    </Overlay>
  );
}

/* ============== JOBS ============== */
function JobsList({ jobs, onOpen }) {
  if (!jobs.length)
    return (
      <div style={{ padding: 60, textAlign: "center", color: "var(--ink-dim)" }}>
        No opportunities yet — win an estimate in the pipeline to create one.
      </div>
    );
  return (
    <div style={{ padding: "0 24px 24px" }}>
      <h2 style={{ fontFamily: "var(--display)", fontSize: 18, fontWeight: 700, letterSpacing: 0.2, color: "var(--ink)" }}>
        Opportunities
      </h2>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 12 }}>
        {jobs.map((j) => {
          const matCost = j.materials.reduce((s, m) => s + (m.cost || 0), 0);
          const laborHrs = j.labor.reduce((s, l) => s + (l.hours || 0), 0);
          return (
            <div key={j.id} onClick={() => onOpen(j.id)} className="p9-card p9-rise"
              style={{ background: "var(--panel)", border: "1px solid var(--line)", borderTop: "3px solid " + (j.status === "active" ? "var(--green)" : j.status === "cancelled" ? "var(--red)" : "var(--ink-dim)"), padding: 16, cursor: "pointer", borderRadius: 4, boxShadow: "var(--shadow-sm)" }}>
              <div style={{ display: "flex", justifyContent: "space-between" }}>
                <span style={{ fontFamily: "var(--display)", fontSize: 11, letterSpacing: 2, color: j.status === "active" ? "var(--green)" : "var(--ink-dim)", textTransform: "uppercase" }}>
                  {j.status}
                </span>
                <span style={{ fontFamily: "var(--mono)", fontSize: 11, color: "var(--ink-dim)" }}>won {fmtDate(j.wonAt)}</span>
              </div>
              <div style={{ fontWeight: 600, fontSize: 16, margin: "6px 0 2px" }}>{j.name}</div>
              <div style={{ fontSize: 13, color: "var(--ink-dim)" }}>{j.client}</div>
              <div style={{ display: "flex", gap: 18, marginTop: 12, fontFamily: "var(--mono)", fontSize: 12 }}>
                <span style={{ color: "var(--copper)" }}>{fmt(j.value)}</span>
                <span style={{ color: "var(--ink-dim)" }}>{j.materials.length} materials · {fmt(matCost)}</span>
                <span style={{ color: "var(--ink-dim)" }}>{laborHrs} hrs</span>
                {jobArr(j) > 0 && <span style={{ color: "var(--green)" }}>{fmt(jobArr(j))}/yr ARR</span>}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function JobDetail({ job, onBack, onUpdate, settings, onSettings }) {
  const [tab, setTab] = useState("estimate");
  const matCost = job.materials.reduce((s, m) => s + (m.cost || 0), 0);
  const laborHrs = job.labor.reduce((s, l) => s + (l.hours || 0), 0);
  const invTotal = job.invoices.reduce((s, v) => s + (v.amount || 0), 0);

  return (
    <div style={{ padding: "0 24px 24px" }}>
      <button onClick={onBack} style={{ background: "none", border: "none", color: "var(--copper)", cursor: "pointer", fontSize: 13, padding: 0, marginBottom: 8 }}>
        ← All opportunities
      </button>
      <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderTop: "3px solid var(--copper)", padding: 22, borderRadius: 10, boxShadow: "var(--shadow-sm)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
          <div>
            <h2 style={{ fontFamily: "var(--display)", fontSize: 22, fontWeight: 700, margin: 0, letterSpacing: 0.5 }}>{job.name}</h2>
            <div style={{ color: "var(--ink-dim)", fontSize: 13, marginTop: 4 }}>
              {job.client} · {job.contact} · {job.phone}<br />{job.address}
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontFamily: "var(--mono)", fontSize: 24, color: "var(--copper)" }}>{fmt(job.value)}</div>
            <select value={job.status}
              onChange={(e) => onUpdate({ status: e.target.value, ...(e.target.value === "cancelled" ? { cancelledAt: today() } : {}) })}
              style={{ marginTop: 6, width: "auto" }}>
              <option value="active">Active</option>
              <option value="on_hold">On hold</option>
              <option value="complete">Complete</option>
              <option value="cancelled">Cancelled (churn)</option>
            </select>
          </div>
        </div>
        <div style={{ display: "flex", gap: 24, marginTop: 14, paddingTop: 14, borderTop: "1px solid var(--line)", fontFamily: "var(--mono)", fontSize: 13, flexWrap: "wrap" }}>
          <span>Materials: <b style={{ color: "var(--copper)" }}>{fmt(matCost)}</b></span>
          <span>Labor: <b style={{ color: "var(--copper)" }}>{laborHrs} hrs</b></span>
          <span>Invoices: <b style={{ color: "var(--copper)" }}>{fmt(invTotal)}</b></span>
          <span>Est. margin vs materials: <b style={{ color: job.value - matCost >= 0 ? "var(--green)" : "var(--red)" }}>{fmt(job.value - matCost)}</b></span>
          <span>ARR: <b style={{ color: "var(--green)" }}>{fmt(jobArr(job))}/yr</b></span>
        </div>
      </div>

      <div style={{ display: "flex", gap: 0, marginTop: 16, borderBottom: "1px solid var(--line)" }}>
        {["estimate", "materials", "labor", "invoices", "notes"].map((t) => (
          <button key={t} onClick={() => setTab(t)}
            className={"p9-tab" + (tab === t ? " on" : "")}
            style={{ background: "none", border: "none", color: tab === t ? "var(--copper)" : "var(--ink-dim)", fontFamily: "var(--display)", letterSpacing: 0.6, fontSize: 14.5, fontWeight: 700, textTransform: "capitalize", padding: "10px 18px", cursor: "pointer" }}>
            {t}
          </button>
        ))}
      </div>

      <div style={{ marginTop: 16 }}>
        {tab === "estimate" && <EstimateTab job={job} onUpdate={onUpdate} />}
        {tab === "materials" && <Materials job={job} onUpdate={onUpdate} />}
        {tab === "labor" && <Labor job={job} onUpdate={onUpdate} />}
        {tab === "invoices" && <Invoices job={job} onUpdate={onUpdate} settings={settings} onSettings={onSettings} />}
        {tab === "notes" && <JobNotes job={job} onUpdate={onUpdate} />}
      </div>
    </div>
  );
}

/* ---- Estimate tab (AI reader + ARR tracking) ---- */
function EstimateTab({ job, onUpdate }) {
  const [f, setF] = useState({ description: "", amount: "", period: "one_time" });
  const [ai, setAi] = useState({ busy: false, msg: "", err: false });
  const fileRef = useRef(null);
  const items = job.estimate?.items || [];
  const oneTime = jobOneTime(job);
  const arr = jobArr(job);

  const addItem = () => {
    if (!f.description.trim()) return;
    onUpdate({ estimate: { ...(job.estimate || {}), items: [...items, { id: uid(), ...f, amount: Number(f.amount) || 0 }] } });
    setF({ description: "", amount: "", period: "one_time" });
  };
  const setPeriod = (id, period) =>
    onUpdate({ estimate: { ...(job.estimate || {}), items: items.map((it) => (it.id === id ? { ...it, period } : it)) } });
  const remove = (id) =>
    onUpdate({ estimate: { ...(job.estimate || {}), items: items.filter((it) => it.id !== id) } });

  const readEstimateWithAI = async (file) => {
    if (!file) return;
    setAi({ busy: true, msg: "Reading estimate with AI…", err: false });
    try {
      const base64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("read failed"));
        r.readAsDataURL(file);
      });
      const isPdf = file.type === "application/pdf";
      const fileBlock = isPdf
        ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } }
        : { type: "image", source: { type: "base64", media_type: file.type || "image/jpeg", data: base64 } };

      const response = await fetch("/api/claude", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 2000,
          messages: [
            {
              role: "user",
              content: [
                fileBlock,
                {
                  type: "text",
                  text:
                    'This is a plumbing estimate/quote that was sent to a client and won. Extract every line item. Respond with ONLY a JSON object, no markdown fences, no extra text: {"items":[{"description":"","amount":0,"period":"one_time"}],"grand_total":0}. ' +
                    'For period, use "one_time" for installs, repairs, and labor; use "monthly", "quarterly", or "annual" ONLY for repeating charges like maintenance agreements, service plans, or monitoring contracts, with amount being the per-period price. If a field is unreadable use empty string or 0.',
                },
              ],
            },
          ],
        }),
      });
      const res = await response.json();
      const text = (res.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
      const parsed = JSON.parse(text.replace(/```json|```/g, "").trim());
      const newItems = (parsed.items || []).map((it) => ({
        id: uid(),
        description: it.description || "Line item",
        amount: Number(it.amount) || 0,
        period: ANNUAL_X[it.period] !== undefined ? it.period : "one_time",
      }));
      if (!newItems.length) throw new Error("no items");

      const ot = newItems.filter((i) => i.period === "one_time").reduce((s, i) => s + i.amount, 0);
      const ar = newItems.reduce((s, i) => s + annualized(i), 0);
      const note = {
        id: uid(), date: today(),
        text: `🤖 AI read the won estimate: ${newItems.length} line items — ${fmt(ot)} one-time work, ${fmt(ar)}/yr recurring (ARR).`,
      };
      onUpdate({
        estimate: { items: newItems, readAt: today() },
        value: ot + ar || job.value,
        notes: [note, ...job.notes],
      });
      setAi({ busy: false, msg: `Read ${newItems.length} line items ✓ — ${fmt(ot)} one-time, ${fmt(ar)}/yr ARR. Note added.`, err: false });
    } catch {
      setAi({ busy: false, msg: "Couldn't read that estimate — try a clearer scan or add line items manually below.", err: true });
    }
    if (fileRef.current) fileRef.current.value = "";
  };

  return (
    <div>
      {/* summary cards */}
      <div style={{ display: "flex", gap: 12, marginBottom: 14, flexWrap: "wrap" }}>
        {[
          ["One-time work", fmt(oneTime), "var(--copper)"],
          ["Recurring (ARR)", fmt(arr) + "/yr", "var(--green)"],
          ["First-year value", fmt(oneTime + arr), "var(--ink)"],
        ].map(([label, val, color]) => (
          <div key={label} style={{ flex: 1, minWidth: 160, background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 4, padding: "12px 16px", boxShadow: "0 1px 2px rgba(20,40,60,0.06)" }}>
            <div style={{ fontFamily: "var(--display)", letterSpacing: 2, fontSize: 11, color: "var(--ink-dim)", textTransform: "uppercase" }}>{label}</div>
            <div style={{ fontFamily: "var(--mono)", fontSize: 20, marginTop: 2, color }}>{val}</div>
          </div>
        ))}
      </div>

      {/* AI reader */}
      <div style={{ border: "1px dashed var(--copper-dim)", background: "var(--panel2)", padding: 14, marginBottom: 14, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", borderRadius: 4 }}>
        <span style={{ fontFamily: "var(--display)", letterSpacing: 2, fontSize: 13, color: "var(--copper)", textTransform: "uppercase" }}>
          🤖 AI Estimate Reader
        </span>
        <span style={{ fontSize: 13, color: "var(--ink-dim)" }}>
          Upload the estimate you sent — it'll pull line items, split one-time vs recurring, and set ARR.
        </span>
        <input ref={fileRef} type="file" accept="image/*,application/pdf" disabled={ai.busy}
          onChange={(e) => readEstimateWithAI(e.target.files?.[0])}
          style={{ width: "auto", border: "none", background: "none", padding: 0, fontSize: 13 }} />
        {ai.msg && (
          <span style={{ fontSize: 13, width: "100%", color: ai.busy ? "var(--copper)" : ai.err ? "var(--red)" : "var(--green)" }}>{ai.msg}</span>
        )}
      </div>

      {/* manual add */}
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <input style={{ flex: 2, minWidth: 200 }} placeholder="Line item — e.g. Annual backflow maintenance plan" value={f.description} onChange={(e) => setF({ ...f, description: e.target.value })} />
        <input style={{ width: 120 }} type="number" placeholder="Amount $" value={f.amount} onChange={(e) => setF({ ...f, amount: e.target.value })} />
        <select style={{ width: "auto" }} value={f.period} onChange={(e) => setF({ ...f, period: e.target.value })}>
          <option value="one_time">One-time</option>
          <option value="monthly">Monthly</option>
          <option value="quarterly">Quarterly</option>
          <option value="annual">Annual</option>
        </select>
        <Btn onClick={addItem}>Add</Btn>
      </div>

      <Table head={["Line item", "Amount", "Billing", "Annualized", ""]}>
        {items.map((it) => (
          <tr key={it.id}>
            <Td>{it.description}</Td>
            <Td mono>{fmt(it.amount)}{it.period !== "one_time" ? "/" + it.period.slice(0, 2) : ""}</Td>
            <Td>
              <select value={it.period} onChange={(e) => setPeriod(it.id, e.target.value)} style={{ width: "auto", fontSize: 12, padding: "4px 6px" }}>
                <option value="one_time">One-time</option>
                <option value="monthly">Monthly</option>
                <option value="quarterly">Quarterly</option>
                <option value="annual">Annual</option>
              </select>
            </Td>
            <Td mono dim>{it.period === "one_time" ? "—" : fmt(annualized(it)) + "/yr"}</Td>
            <Td><Del onClick={() => remove(it.id)} /></Td>
          </tr>
        ))}
      </Table>
      {!items.length && <Empty>No estimate line items yet — upload the sent estimate above or add them manually.</Empty>}
      {items.length > 0 && (
        <div style={{ fontSize: 12, color: "var(--ink-dim)", marginTop: 8 }}>
          Recurring line items roll up into the company ARR stat at the top. If this job is ever cancelled, {fmt(arr)}/yr is subtracted from this year's net new ARR.
        </div>
      )}
    </div>
  );
}

/* ---- Materials tab ---- */
function Materials({ job, onUpdate }) {
  const [f, setF] = useState({ item: "", vendor: "", cost: "" });
  const add = () => {
    if (!f.item.trim()) return;
    onUpdate({ materials: [...job.materials, { id: uid(), item: f.item, vendor: f.vendor, cost: Number(f.cost) || 0, status: "needed", date: today() }] });
    setF({ item: "", vendor: "", cost: "" });
  };
  const setStatus = (id, status) =>
    onUpdate({ materials: job.materials.map((m) => (m.id === id ? { ...m, status } : m)) });
  const remove = (id) => onUpdate({ materials: job.materials.filter((m) => m.id !== id) });

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <input style={{ flex: 2, minWidth: 180 }} placeholder='Item — e.g. 3/4" copper, 200ft' value={f.item} onChange={(e) => setF({ ...f, item: e.target.value })} />
        <input style={{ flex: 1, minWidth: 120 }} placeholder="Vendor" value={f.vendor} onChange={(e) => setF({ ...f, vendor: e.target.value })} />
        <input style={{ width: 110 }} type="number" placeholder="Cost $" value={f.cost} onChange={(e) => setF({ ...f, cost: e.target.value })} />
        <Btn onClick={add}>Add</Btn>
      </div>
      <Table head={["Item", "Vendor", "Cost", "Status", ""]}>
        {job.materials.map((m) => (
          <tr key={m.id}>
            <Td>{m.item}</Td>
            <Td dim>{m.vendor || "—"}</Td>
            <Td mono>{fmt(m.cost)}</Td>
            <Td>
              <select value={m.status} onChange={(e) => setStatus(m.id, e.target.value)} style={{ width: "auto", fontSize: 12, padding: "4px 6px" }}>
                <option value="needed">Needed</option>
                <option value="ordered">Ordered</option>
                <option value="received">Received</option>
                <option value="installed">Installed</option>
              </select>
            </Td>
            <Td><Del onClick={() => remove(m.id)} /></Td>
          </tr>
        ))}
      </Table>
      {!job.materials.length && <Empty>No materials yet — add what needs ordering above.</Empty>}
    </div>
  );
}

/* ---- Labor tab ---- */
function Labor({ job, onUpdate }) {
  const [f, setF] = useState({ name: "", role: "", date: today(), hours: "" });
  const add = () => {
    if (!f.name.trim()) return;
    onUpdate({ labor: [...job.labor, { id: uid(), ...f, hours: Number(f.hours) || 0 }] });
    setF({ name: "", role: "", date: today(), hours: "" });
  };
  const remove = (id) => onUpdate({ labor: job.labor.filter((l) => l.id !== id) });

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <input style={{ flex: 1, minWidth: 140 }} placeholder="Crew member" value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} />
        <input style={{ flex: 1, minWidth: 120 }} placeholder="Role — journeyman, apprentice…" value={f.role} onChange={(e) => setF({ ...f, role: e.target.value })} />
        <input style={{ width: 150 }} type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} />
        <input style={{ width: 90 }} type="number" placeholder="Hrs" value={f.hours} onChange={(e) => setF({ ...f, hours: e.target.value })} />
        <Btn onClick={add}>Add</Btn>
      </div>
      <Table head={["Crew", "Role", "Date", "Hours", ""]}>
        {[...job.labor].sort((a, b) => (a.date < b.date ? -1 : 1)).map((l) => (
          <tr key={l.id}>
            <Td>{l.name}</Td>
            <Td dim>{l.role || "—"}</Td>
            <Td mono>{fmtDate(l.date)}</Td>
            <Td mono>{l.hours}</Td>
            <Td><Del onClick={() => remove(l.id)} /></Td>
          </tr>
        ))}
      </Table>
      {!job.labor.length && <Empty>No labor scheduled — assign crew and dates above.</Empty>}
    </div>
  );
}

/* ---- Invoices tab ---- */
function Invoices({ job, onUpdate, settings, onSettings }) {
  const [f, setF] = useState({ vendor: "", ref: "", amount: "", date: today() });
  const [ai, setAi] = useState({ busy: false, msg: "", err: false });
  const fileRef = useRef(null);

  const add = () => {
    if (!f.vendor.trim()) return;
    onUpdate({ invoices: [...job.invoices, { id: uid(), ...f, amount: Number(f.amount) || 0, paid: false }] });
    setF({ vendor: "", ref: "", amount: "", date: today() });
  };
  const togglePaid = (id) =>
    onUpdate({ invoices: job.invoices.map((v) => (v.id === id ? { ...v, paid: !v.paid } : v)) });
  const remove = (id) => onUpdate({ invoices: job.invoices.filter((v) => v.id !== id) });

  const sendUrgentEmail = async (inv, summary) => {
    setAi({ busy: true, msg: "Invoice logged ✓ — sending urgent email…", err: false });
    try {
      const response = await fetch("/api/claude", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          messages: [
            {
              role: "user",
              content:
                `Use the Microsoft 365 / Outlook tools to send an email and mark it HIGH IMPORTANCE / urgent priority.\n` +
                `To: ${settings.alertEmail}\n` +
                `Subject: [URGENT] Invoice logged — ${inv.vendor} ${fmt(inv.amount)} (${job.name})\n` +
                `Body:\nAn invoice was just logged by the AI reader in Pipeline:\n\n` +
                `Opportunity: ${job.name} (${job.client})\nVendor: ${inv.vendor}\nInvoice #: ${inv.ref || "n/a"}\nAmount: ${fmt(inv.amount)}\nDate: ${inv.date}\n${summary ? "Details: " + summary : ""}\n\n` +
                `Status: UNPAID — review and schedule payment.\n\n` +
                `After sending, reply with exactly the word SENT if it succeeded, or FAILED plus the reason.`,
            },
          ],
          mcp_servers: [
            { type: "url", url: "https://microsoft365.mcp.claude.com/mcp", name: "microsoft365" },
          ],
        }),
      });
      const res = await response.json();
      const text = (res.content || []).filter((b) => b.type === "text").map((b) => b.text).join(" ");
      if (text.includes("SENT")) {
        setAi({ busy: false, msg: `Logged ${inv.vendor} — ${fmt(inv.amount)} ✓ · Urgent email sent to ${settings.alertEmail} ✓`, err: false });
      } else {
        setAi({ busy: false, msg: `Invoice logged ✓ but the email failed — check your Microsoft 365 connection.`, err: true });
      }
    } catch {
      setAi({ busy: false, msg: "Invoice logged ✓ but the email couldn't be sent.", err: true });
    }
  };

  const readInvoiceWithAI = async (file) => {
    if (!file) return;
    setAi({ busy: true, msg: "Reading invoice with AI…", err: false });
    try {
      const base64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("read failed"));
        r.readAsDataURL(file);
      });
      const isPdf = file.type === "application/pdf";
      const fileBlock = isPdf
        ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } }
        : { type: "image", source: { type: "base64", media_type: file.type || "image/jpeg", data: base64 } };

      const response = await fetch("/api/claude", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          messages: [
            {
              role: "user",
              content: [
                fileBlock,
                {
                  type: "text",
                  text:
                    'This is a supplier/material invoice for a plumbing job. Extract its details. Respond with ONLY a JSON object, no markdown fences, no extra text: {"vendor":"supply house name","invoice_number":"","amount":0,"date":"YYYY-MM-DD","summary":"one short sentence describing what was purchased"}. If a field is unreadable use an empty string or 0.',
                },
              ],
            },
          ],
        }),
      });
      const res = await response.json();
      const text = (res.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
      const parsed = JSON.parse(text.replace(/```json|```/g, "").trim());

      const inv = {
        id: uid(),
        vendor: parsed.vendor || "Unknown vendor",
        ref: parsed.invoice_number || "",
        amount: Number(parsed.amount) || 0,
        date: parsed.date || today(),
        paid: false,
      };
      const note = {
        id: uid(),
        date: today(),
        text: `🤖 AI logged invoice ${inv.ref ? "#" + inv.ref + " " : ""}from ${inv.vendor} — ${fmt(inv.amount)} (${fmtDate(inv.date)}). ${parsed.summary || ""}`.trim(),
      };
      onUpdate({ invoices: [...job.invoices, inv], notes: [note, ...job.notes] });
      if ((settings?.invoiceEmail ?? true) && settings?.alertEmail?.trim()) {
        await sendUrgentEmail(inv, parsed.summary || "");
      } else {
        setAi({ busy: false, msg: `Logged ${inv.vendor} — ${fmt(inv.amount)} ✓ (note added)`, err: false });
      }
    } catch {
      setAi({ busy: false, msg: "Couldn't read that file — try a clearer photo or enter it manually below.", err: true });
    }
    if (fileRef.current) fileRef.current.value = "";
  };

  return (
    <div>
      {/* AI reader */}
      <div style={{ border: "1px dashed var(--copper-dim)", background: "var(--panel2)", padding: 14, marginBottom: 14, display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
        <span style={{ fontFamily: "var(--display)", letterSpacing: 2, fontSize: 13, color: "var(--copper)", textTransform: "uppercase" }}>
          🤖 AI Invoice Reader
        </span>
        <span style={{ fontSize: 13, color: "var(--ink-dim)" }}>
          Upload a photo or PDF — it'll extract vendor, #, amount, date, and add a note.
        </span>
        <input ref={fileRef} type="file" accept="image/*,application/pdf" disabled={ai.busy}
          onChange={(e) => readInvoiceWithAI(e.target.files?.[0])}
          style={{ width: "auto", border: "none", background: "none", padding: 0, fontSize: 13 }} />
        {ai.busy && <span style={{ fontSize: 13, color: "var(--copper)" }}>{ai.msg}</span>}
        {!ai.busy && ai.msg && (
          <span style={{ fontSize: 13, color: ai.err ? "var(--red)" : "var(--green)", width: "100%" }}>{ai.msg}</span>
        )}
        <div style={{ width: "100%", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", paddingTop: 6, borderTop: "1px solid var(--line)" }}>
          <label style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 13, color: "var(--ink-dim)", cursor: "pointer", whiteSpace: "nowrap" }}>
            <input type="checkbox" style={{ width: "auto" }}
              checked={settings?.invoiceEmail ?? true}
              onChange={(e) => onSettings({ invoiceEmail: e.target.checked })} />
            Send urgent Outlook email when logged
          </label>
          <input type="email" placeholder="recipient@yourcompany.com" style={{ flex: 1, minWidth: 200 }}
            value={settings?.alertEmail || ""}
            onChange={(e) => onSettings({ alertEmail: e.target.value })} />
        </div>
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <input style={{ flex: 1, minWidth: 140 }} placeholder="Vendor / supply house" value={f.vendor} onChange={(e) => setF({ ...f, vendor: e.target.value })} />
        <input style={{ width: 140 }} placeholder="Invoice #" value={f.ref} onChange={(e) => setF({ ...f, ref: e.target.value })} />
        <input style={{ width: 120 }} type="number" placeholder="Amount $" value={f.amount} onChange={(e) => setF({ ...f, amount: e.target.value })} />
        <input style={{ width: 150 }} type="date" value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} />
        <Btn onClick={add}>Add</Btn>
      </div>
      <Table head={["Vendor", "Invoice #", "Amount", "Date", "Paid", ""]}>
        {job.invoices.map((v) => (
          <tr key={v.id}>
            <Td>{v.vendor}</Td>
            <Td mono dim>{v.ref || "—"}</Td>
            <Td mono>{fmt(v.amount)}</Td>
            <Td mono>{fmtDate(v.date)}</Td>
            <Td>
              <button onClick={() => togglePaid(v.id)}
                style={{ background: v.paid ? "var(--green)" : "transparent", color: v.paid ? "#ffffff" : "var(--ink-dim)", border: "1px solid " + (v.paid ? "var(--green)" : "var(--line)"), padding: "3px 10px", fontSize: 11, cursor: "pointer", fontFamily: "var(--display)", letterSpacing: 1 }}>
                {v.paid ? "PAID" : "UNPAID"}
              </button>
            </Td>
            <Td><Del onClick={() => remove(v.id)} /></Td>
          </tr>
        ))}
      </Table>
      {!job.invoices.length && <Empty>No material invoices logged.</Empty>}
    </div>
  );
}

/* ---- Notes tab ---- */
function JobNotes({ job, onUpdate }) {
  const [note, setNote] = useState("");
  const add = () => {
    if (!note.trim()) return;
    onUpdate({ notes: [{ id: uid(), text: note.trim(), date: today() }, ...job.notes] });
    setNote("");
  };
  return (
    <div>
      <div style={{ display: "flex", gap: 8 }}>
        <input placeholder="Add note — inspection scheduled, change order, punch list…" value={note}
          onChange={(e) => setNote(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
        <Btn onClick={add}>Add</Btn>
      </div>
      {job.notes.map((n) => (
        <div key={n.id} style={{ borderLeft: "2px solid var(--copper-dim)", padding: "8px 14px", marginTop: 10, fontSize: 14, background: "var(--panel)" }}>
          <div style={{ color: "var(--ink-dim)", fontFamily: "var(--mono)", fontSize: 11, marginBottom: 2 }}>{fmtDate(n.date)}</div>
          {n.text}
        </div>
      ))}
      {!job.notes.length && <Empty>No notes yet.</Empty>}
    </div>
  );
}



/* ============== LOGO ============== */
function PipeMark({ size = 36, light = false }) {
  const c1 = light ? "#ffffff" : "var(--copper)";
  const c2 = light ? "rgba(255,255,255,0.6)" : "var(--copper-hi)";
  const badge = light ? "rgba(255,255,255,0.12)" : "var(--navy)";
  return (
    <svg width={size} height={size} viewBox="0 0 48 48" fill="none" aria-hidden="true">
      <rect x="2" y="2" width="44" height="44" rx="11" fill={badge} />
      {/* Crossed combination wrench + screwdriver — a blue-collar trades emblem */}
      {/* Screwdriver (drawn first so the wrench sits on top at the crossing) */}
      <g transform="rotate(-45 24 24)">
        <rect x="20.5" y="8.5" width="7" height="12" rx="3" fill={c2} />
        <rect x="22" y="20" width="4" height="15" rx="1.5" fill={c2} />
        <path d="M22 35 H26 L24 40 Z" fill={c2} />
      </g>
      {/* Combination wrench: box-ring end + open jaw */}
      <g transform="rotate(45 24 24)">
        <rect x="20.75" y="10" width="6.5" height="28" rx="3.25" fill={c1} />
        <circle cx="24" cy="11" r="7" fill={c1} />
        <circle cx="24" cy="11" r="3.2" fill={badge} />
        <circle cx="24" cy="37" r="7" fill={c1} />
        <path d="M24 37 L17.4 44.6 L30.6 44.6 Z" fill={badge} />
      </g>
    </svg>
  );
}

/* ============== LOGIN ============== */
function Login() {
  const [mode, setMode] = useState("signin");
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [msg, setMsg] = useState(null); // {ok, text}
  const [busy, setBusy] = useState(false);

  const go = async () => {
    if (!email.trim() || !pw) { setMsg({ ok: false, text: "Enter your email and a password." }); return; }
    setBusy(true); setMsg(null);
    try {
      if (mode === "signin") {
        await authSignIn(email.trim(), pw);
      } else {
        const res = await authSignUp(email.trim(), pw);
        if (!res?.session) setMsg({ ok: true, text: "Account created — check your email for a confirmation link, then sign in." });
      }
    } catch (e) {
      setMsg({ ok: false, text: e.message || "Sign-in failed." });
    }
    setBusy(false);
  };

  return (
    <div style={{ minHeight: "100vh", display: "grid", gridTemplateColumns: "1.05fr 1fr", background: "var(--bg)" }} className="p9-login">
      <style>{`
        @media (max-width: 820px) { .p9-login { grid-template-columns: 1fr !important; } .p9-login-brand { display: none !important; } }
      `}</style>

      {/* brand panel */}
      <div className="p9-login-brand" style={{
        position: "relative", overflow: "hidden", color: "#fff",
        background: "linear-gradient(150deg, var(--navy) 0%, var(--navy2) 60%, #0a63b8 100%)",
        padding: "56px 56px", display: "flex", flexDirection: "column", justifyContent: "space-between",
      }}>
        {/* faint blueprint grid + glow */}
        <div style={{ position: "absolute", inset: 0, opacity: 0.5, background:
          "radial-gradient(600px 300px at 80% 10%, rgba(27,150,255,0.35), transparent 60%)," +
          "repeating-linear-gradient(0deg, rgba(255,255,255,0.05) 0 1px, transparent 1px 32px)," +
          "repeating-linear-gradient(90deg, rgba(255,255,255,0.05) 0 1px, transparent 1px 32px)" }} />
        <div style={{ position: "relative", display: "flex", alignItems: "center", gap: 12 }}>
          <PipeMark size={44} light />
          <span style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 26, letterSpacing: -0.3 }}>Pipeline</span>
        </div>
        <div style={{ position: "relative" }}>
          <div style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 34, lineHeight: 1.12, letterSpacing: -0.5, maxWidth: 420 }}>
            From intent to bid<br />to invoice paid.
          </div>
          <div style={{ fontSize: 15, color: "rgba(255,255,255,0.78)", marginTop: 16, maxWidth: 400, lineHeight: 1.5 }}>
            Track every ITB, estimate, and job in one place — built for plumbing &amp; MEP contractors.
          </div>
          <div style={{ display: "flex", gap: 26, marginTop: 30 }}>
            {[["Pipeline", "ITB → won"], ["Estimates", "ARR tracked"], ["Opportunities", "materials · labor"]].map(([a, b]) => (
              <div key={a}>
                <div style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 15 }}>{a}</div>
                <div style={{ fontSize: 12, color: "rgba(255,255,255,0.6)" }}>{b}</div>
              </div>
            ))}
          </div>
        </div>
        <div style={{ position: "relative", fontFamily: "var(--mono)", fontSize: 11, letterSpacing: 1.5, color: "rgba(255,255,255,0.45)", textTransform: "uppercase" }}>
          ITB · Estimate · Build
        </div>
      </div>

      {/* form panel */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 28 }}>
        <div className="p9-rise" style={{ width: "100%", maxWidth: 380 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }} className="p9-login-mark-sm">
            <PipeMark size={32} />
            <span style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 20, color: "var(--navy)" }}>Pipeline</span>
          </div>
          <h1 style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 24, margin: "10px 0 4px", color: "var(--ink)" }}>
            {mode === "signin" ? "Welcome back" : "Create your account"}
          </h1>
          <div style={{ fontSize: 13.5, color: "var(--ink-dim)", marginBottom: 22 }}>
            {mode === "signin" ? "Sign in to your team's workspace." : "Use the email your owner gave a seat to."}
          </div>
          <Field label="Email"><input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@yourcompany.com" /></Field>
          <Field label="Password"><input type="password" value={pw} onChange={(e) => setPw(e.target.value)} placeholder="••••••••" onKeyDown={(e) => e.key === "Enter" && go()} /></Field>
          {msg && (
            <div style={{ marginTop: 14, fontSize: 13, padding: "9px 12px", borderRadius: 6, background: msg.ok ? "#eaf6ef" : "#fdeceb", color: msg.ok ? "var(--green)" : "var(--red)", border: "1px solid " + (msg.ok ? "#bfe3cd" : "#f4cdca") }}>
              {msg.text}
            </div>
          )}
          <button onClick={go} disabled={busy} className="p9-btn"
            style={{ width: "100%", marginTop: 20, background: "var(--copper)", color: "#fff", border: "none", borderRadius: 6, padding: "12px", fontFamily: "var(--display)", fontWeight: 600, fontSize: 15, cursor: busy ? "default" : "pointer", opacity: busy ? 0.6 : 1 }}>
            {busy ? "Working…" : mode === "signin" ? "Sign in" : "Create account"}
          </button>
          <div style={{ textAlign: "center", marginTop: 18, fontSize: 13, color: "var(--ink-dim)" }}>
            {mode === "signin" ? "New to your team? " : "Already have an account? "}
            <button onClick={() => { setMode(mode === "signin" ? "signup" : "signin"); setMsg(null); }}
              style={{ background: "none", border: "none", color: "var(--copper)", fontSize: 13, cursor: "pointer", fontWeight: 600, padding: 0 }}>
              {mode === "signin" ? "Create an account" : "Sign in"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============== TEAM (SEATS) ============== */
function Team({ session }) {
  const [seats, setSeats] = useState(null);
  const [email, setEmail] = useState("");
  const [msg, setMsg] = useState(null);
  const [ping, setPing] = useState(null); // { to, text, status }
  const [inviteState, setInviteState] = useState({}); // { [email]: statusText }
  const me = (session?.user?.email || "").toLowerCase();
  const myRole = seats?.find((s) => s.email.toLowerCase() === me)?.role;

  const refresh = async () => {
    try { setSeats(await listSeats()); }
    catch (e) { setMsg({ ok: false, text: e.message }); setSeats([]); }
  };
  useEffect(() => { if (usingSharedDb) refresh(); }, []);

  if (!usingSharedDb)
    return (
      <ListPage title="Team" subtitle="Seats control who can sign in and see your data.">
        <Empty>Seats need the shared database. This preview runs in local mode — the deployed version (with Supabase connected) manages seats here.</Empty>
      </ListPage>
    );

  // Add the seat, then email them a "create your login" invite.
  const invite = async () => {
    const addr = email.trim().toLowerCase();
    if (!addr) return;
    setMsg(null);
    try {
      await addSeat(addr);
      setEmail(""); refresh();
      const r = await sendInvite(addr);
      setMsg(r.ok
        ? { ok: true, text: `Seat added — an invite to create their login was emailed to ${addr}.` }
        : { ok: false, text: `Seat added for ${addr}, but the invite didn't send: ${r.error} Use "Invite" to retry.` });
    } catch (e) {
      setMsg({ ok: false, text: e.message.includes("policy") || e.message.includes("permission") ? "Only owners can add seats." : e.message });
    }
  };
  const doInvite = async (addr) => {
    setInviteState((s) => ({ ...s, [addr]: "Sending…" }));
    const r = await sendInvite(addr);
    setInviteState((s) => ({ ...s, [addr]: r.ok ? "Invited ✓" : "✕ Failed" }));
    if (!r.ok) setMsg({ ok: false, text: r.error });
  };
  const drop = async (s) => {
    setMsg(null);
    try { await removeSeat(s.id); refresh(); }
    catch (e) { setMsg({ ok: false, text: e.message.includes("policy") || e.message.includes("permission") ? "Only owners can remove seats." : e.message }); }
  };

  const openPing = (to) => setPing({ to, text: "", status: null });
  const sendPing = async () => {
    if (!ping) return;
    setPing((p) => ({ ...p, status: { sending: true, text: "Sending…" } }));
    try {
      const { ok, text } = await sendOutlookEmail({
        to: ping.to,
        subject: `Pipeline CRM — a ping${me ? " from " + me : ""}`,
        body: (ping.text.trim() || "Just pinging you — please check Pipeline CRM when you get a chance.") + `\n\n— Sent from Pipeline CRM`,
      });
      setPing((p) => ({ ...p, status: { sending: false, ok, text: ok ? "Ping sent ✓" : (text.trim().slice(0, 200) || "Send failed — check the Microsoft 365 connection.") } }));
    } catch {
      setPing((p) => ({ ...p, status: { sending: false, ok: false, text: "Couldn't reach the email service." } }));
    }
  };

  return (
    <>
    <ListPage title="Team" subtitle="Add a seat and we'll email them a link to create their Pipeline login. Owners manage seats; members get full CRM access.">
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <input style={{ flex: 1, minWidth: 240 }} type="email" placeholder="teammate@yourcompany.com"
          value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={(e) => e.key === "Enter" && invite()} />
        <Btn onClick={invite}>Add seat &amp; invite</Btn>
        {seats?.length ? <Btn ghost color="var(--copper)" onClick={() => openPing(seats.map((s) => s.email).join(", "))}>✉ Ping all</Btn> : null}
      </div>
      {msg && <div style={{ marginBottom: 12, fontSize: 13, color: msg.ok ? "var(--green)" : "var(--red)" }}>{msg.text}</div>}
      {seats === null && <Empty>Loading seats…</Empty>}
      {seats !== null && (
        <Table head={["Email", "Role", "Added", ""]}>
          {seats.map((s) => (
            <tr key={s.id}>
              <Td><b>{s.email}</b>{s.email.toLowerCase() === me ? <span style={{ color: "var(--ink-dim)" }}> (you)</span> : null}</Td>
              <Td>
                <span style={{ fontSize: 11, fontWeight: 700, color: s.role === "owner" ? "var(--copper)" : "var(--ink-dim)", textTransform: "uppercase", letterSpacing: 1 }}>{s.role}</span>
              </Td>
              <Td mono dim>{fmtDate((s.created_at || "").slice(0, 10))}</Td>
              <Td>
                <span style={{ display: "flex", gap: 10, alignItems: "center", justifyContent: "flex-end" }}>
                  <button onClick={() => doInvite(s.email)} title={"Email " + s.email + " a link to create their login"}
                    style={{ background: "none", border: "1px solid var(--line)", borderRadius: "var(--radius)", color: inviteState[s.email] === "Invited ✓" ? "var(--green)" : "var(--copper)", cursor: "pointer", fontSize: 12, padding: "5px 10px", whiteSpace: "nowrap" }}>
                    {inviteState[s.email] || "✉ Invite"}
                  </button>
                  <button onClick={() => openPing(s.email)} title={"Ping " + s.email}
                    style={{ background: "none", border: "1px solid var(--line)", borderRadius: "var(--radius)", color: "var(--copper)", cursor: "pointer", fontSize: 12, padding: "5px 10px", whiteSpace: "nowrap" }}>
                    ✉ Ping
                  </button>
                  {s.role !== "owner" && myRole === "owner" ? <Del onClick={() => drop(s)} /> : null}
                </span>
              </Td>
            </tr>
          ))}
        </Table>
      )}
      <div style={{ fontSize: 12, color: "var(--ink-dim)", marginTop: 12 }}>
        How it works: add a seat → your teammate opens the app URL → "Create account" with that email → they're in. Removing a seat locks them out of all data on their next request.
      </div>
    </ListPage>
    {ping && (
      <Overlay narrow onClose={() => setPing(null)}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontFamily: "var(--display)", fontSize: 12, letterSpacing: 3, color: "var(--copper)", textTransform: "uppercase" }}>Ping team</div>
            <h2 style={{ margin: "4px 0 2px", fontFamily: "var(--display)", fontSize: 20, fontWeight: 700 }}>Send a ping</h2>
            <div style={{ color: "var(--ink-dim)", fontSize: 13, wordBreak: "break-word" }}>To: {ping.to}</div>
          </div>
          <CloseBtn onClick={() => setPing(null)} />
        </div>
        <Field label="Message">
          <textarea rows={5} placeholder="Optional — a quick note. Leave blank to send a generic check-in." value={ping.text}
            onChange={(e) => setPing((p) => ({ ...p, text: e.target.value }))} />
        </Field>
        {ping.status && (
          <div style={{ marginTop: 12, fontSize: 13, padding: "9px 12px", borderRadius: 6,
            background: ping.status.ok ? "#eaf6ef" : ping.status.sending ? "#e8f1fb" : "#fdeceb",
            color: ping.status.ok ? "var(--green)" : ping.status.sending ? "#1f5fa6" : "var(--red)",
            border: "1px solid " + (ping.status.ok ? "#bfe3cd" : ping.status.sending ? "#c9def5" : "#f4cdca") }}>
            {ping.status.text}
          </div>
        )}
        <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
          <Btn onClick={sendPing} disabled={ping.status?.sending}>{ping.status?.sending ? "Sending…" : "Send ping"}</Btn>
          <Btn ghost onClick={() => setPing(null)}>Close</Btn>
        </div>
      </Overlay>
    )}
    </>
  );
}

/* ============== GLOBAL SEARCH ============== */
function GlobalSearch({ data, go }) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    const key = (e) => {
      if (e.key === "/" && document.activeElement?.tagName !== "INPUT" && document.activeElement?.tagName !== "TEXTAREA") {
        e.preventDefault(); ref.current?.focus();
      }
      if (e.key === "Escape") { setOpen(false); ref.current?.blur(); }
    };
    window.addEventListener("keydown", key);
    return () => window.removeEventListener("keydown", key);
  }, []);

  const needle = q.trim().toLowerCase();
  const hit = (s) => (s || "").toLowerCase().includes(needle);
  const results = !needle ? [] : [
    ...data.itbs.filter((i) => hit(i.name) || hit(i.client) || hit(i.contact)).slice(0, 4)
      .map((i) => ({ k: "itb", id: i.id, view: "pipeline", title: i.name, sub: "ITB · " + i.client, icon: "◔" })),
    ...data.jobs.filter((j) => hit(j.name) || hit(j.client)).slice(0, 4)
      .map((j) => ({ k: "job", id: j.id, view: "jobs", title: j.name, sub: "Opportunity · " + j.client, icon: "▣" })),
    ...data.accounts.filter((a) => hit(a.name) || hit(a.industry)).slice(0, 3)
      .map((a) => ({ k: "account", id: a.id, view: "accounts", title: a.name, sub: "Account", icon: "◆" })),
    ...data.contacts.filter((c) => hit(c.name) || hit(c.accountName) || hit(c.email)).slice(0, 3)
      .map((c) => ({ k: "contact", id: c.id, view: "contacts", title: c.name, sub: "Contact · " + (c.accountName || "—"), icon: "●" })),
    ...data.tasks.filter((t) => hit(t.title) || hit(t.related)).slice(0, 3)
      .map((t) => ({ k: "task", id: t.id, view: "tasks", title: t.title, sub: "Task · due " + fmtDate(t.due), icon: "✓" })),
  ].slice(0, 10);

  return (
    <div style={{ position: "relative", width: 250 }}>
      <input ref={ref} value={q} placeholder="Search everything…  ( / )"
        onChange={(e) => { setQ(e.target.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
        style={{ background: "#ffffff", border: "1px solid #c9c9c9", color: "var(--ink)", borderRadius: 4, padding: "7px 14px", fontSize: 13 }} />
      {open && results.length > 0 && (
        <div className="p9-rise" style={{ position: "absolute", top: 44, left: 0, width: 330, maxHeight: 380, overflowY: "auto", background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 4, boxShadow: "var(--shadow-md)", zIndex: 60 }}>
          {results.map((r) => (
            <div key={r.k + r.id} onClick={() => { go(r.view, r.id, r.k); setQ(""); setOpen(false); }}
              style={{ display: "flex", gap: 10, alignItems: "center", padding: "10px 14px", cursor: "pointer", borderBottom: "1px solid var(--line)" }}>
              <span style={{ color: "var(--copper)", fontFamily: "var(--mono)", fontSize: 13 }}>{r.icon}</span>
              <div style={{ minWidth: 0 }}>
                <div style={{ fontSize: 13.5, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{r.title}</div>
                <div style={{ fontSize: 11.5, color: "var(--ink-dim)" }}>{r.sub}</div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ============== DASHBOARD ============== */
function Dashboard({ data, active, totalArr, narrYtd, thisYear, winRate, pipelineValue, go, onOpenItb, onToggleTask }) {
  const maxCount = Math.max(1, ...STAGES.map((s) => active.filter((i) => i.stage === s.id).length));
  const openTasks = data.tasks.filter((t) => !t.done).sort((a, b) => (a.due || "9999") < (b.due || "9999") ? -1 : 1).slice(0, 6);
  const overdue = (t) => t.due && t.due < today();
  const recent = [
    ...data.itbs.flatMap((i) => (i.notes || []).slice(0, 1).map((n) => ({ d: n.date, t: n.text, src: i.name }))),
    ...data.jobs.flatMap((j) => (j.notes || []).slice(0, 1).map((n) => ({ d: n.date, t: n.text, src: j.name }))),
  ].sort((a, b) => (a.d < b.d ? 1 : -1)).slice(0, 5);

  return (
    <div style={{ padding: "0 24px 24px", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(330px, 1fr))", gap: 14 }}>
      <Panel title="Pipeline funnel" action={["Open board →", () => go("pipeline")]}>
        {STAGES.map((s) => {
          const cards = active.filter((i) => i.stage === s.id);
          const val = cards.reduce((x, c) => x + (c.value || 0), 0);
          return (
            <div key={s.id} style={{ marginBottom: 9 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 3 }}>
                <span style={{ fontWeight: 600 }}>{s.label}</span>
                <span style={{ fontFamily: "var(--mono)", color: "var(--ink-dim)", fontSize: 11 }}>{cards.length} · {fmt(val)}</span>
              </div>
              <div style={{ height: 10, background: "var(--panel2)", borderRadius: 499, overflow: "hidden" }}>
                <div style={{ height: "100%", width: (cards.length / maxCount) * 100 + "%", minWidth: cards.length ? 10 : 0, background: "var(--copper)", borderRadius: 499, transition: "width .4s ease" }} />
              </div>
            </div>
          );
        })}
        <div style={{ fontSize: 12, color: "var(--ink-dim)", marginTop: 10, fontFamily: "var(--mono)" }}>
          {fmt(pipelineValue)} open · win rate {winRate === null ? "—" : winRate + "%"} · ARR {fmt(totalArr)} · NARR {thisYear}: {fmt(narrYtd)}
        </div>
      </Panel>

      <Panel title="Tasks due" action={["All tasks →", () => go("tasks")]}>
        {openTasks.length === 0 && <Empty>Nothing due — add tasks so follow-ups never slip.</Empty>}
        {openTasks.map((t) => (
          <label key={t.id} style={{ display: "flex", gap: 10, alignItems: "center", padding: "8px 4px", borderBottom: "1px solid var(--line)", cursor: "pointer" }}>
            <input type="checkbox" checked={!!t.done} onChange={(e) => onToggleTask(t.id, e.target.checked)} style={{ width: "auto" }} />
            <span style={{ fontSize: 13.5, flex: 1 }}>{t.title}{t.related ? <span style={{ color: "var(--ink-dim)" }}> — {t.related}</span> : null}</span>
            <span style={{ fontFamily: "var(--mono)", fontSize: 11, color: overdue(t) ? "var(--red)" : "var(--ink-dim)" }}>
              {overdue(t) ? "OVERDUE " : ""}{fmtDate(t.due)}
            </span>
          </label>
        ))}
      </Panel>

      <Panel title="Needs attention">
        {active.filter((i) => i.stage === "estimate_sent").slice(0, 5).map((i) => (
          <div key={i.id} onClick={() => onOpenItb(i.id)} style={{ display: "flex", justifyContent: "space-between", padding: "9px 4px", borderBottom: "1px solid var(--line)", cursor: "pointer", fontSize: 13.5 }}>
            <span><b>{i.name}</b> <span style={{ color: "var(--ink-dim)" }}>awaiting decision</span></span>
            <span style={{ fontFamily: "var(--mono)", color: "var(--copper)" }}>{fmt(i.value)}</span>
          </div>
        ))}
        {active.filter((i) => i.stage === "estimate_sent").length === 0 && <Empty>No estimates out for decision right now.</Empty>}
      </Panel>

      <Panel title="Recent activity">
        {recent.length === 0 && <Empty>Notes you add to ITBs and jobs show up here.</Empty>}
        {recent.map((r, i) => (
          <div key={i} style={{ padding: "8px 4px", borderBottom: "1px solid var(--line)", fontSize: 13 }}>
            <span style={{ fontFamily: "var(--mono)", fontSize: 10.5, color: "var(--ink-dim)", marginRight: 8 }}>{fmtDate(r.d)}</span>
            <b>{r.src}:</b> {r.t.length > 90 ? r.t.slice(0, 90) + "…" : r.t}
          </div>
        ))}
      </Panel>
    </div>
  );
}

const Panel = ({ title, action, children }) => (
  <div className="p9-rise" style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 4, padding: 18, boxShadow: "var(--shadow-sm)" }}>
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 12 }}>
      <span style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 14 }}>{title}</span>
      {action && (
        <button onClick={action[1]} style={{ background: "none", border: "none", color: "var(--copper)", fontSize: 12.5, cursor: "pointer", fontWeight: 600 }}>{action[0]}</button>
      )}
    </div>
    {children}
  </div>
);

/* ============== GC ESTIMATE SCOREBOARD ============== */
function GcScoreboard({ data, onOpenItb }) {
  // Auto-derived from every ITB's client — no setup needed.
  const norm = (s) => (s || "").trim().toLowerCase();
  const byGc = {};
  for (const i of data.itbs) {
    const k = norm(i.client) || "(no gc listed)";
    if (!byGc[k]) byGc[k] = { name: (i.client || "(no GC listed)").trim(), bids: 0, sent: 0, pending: 0, won: 0, lost: 0, valueSent: 0, valueWon: 0, pendingItbs: [] };
    const g = byGc[k];
    g.bids += 1;
    const wasSent = (i.history || []).some((h) => h.stage === "estimate_sent") || ["estimate_sent", "won"].includes(i.stage);
    if (wasSent) { g.sent += 1; g.valueSent += i.value || 0; }
    if (i.stage === "estimate_sent") { g.pending += 1; g.pendingItbs.push(i); }
    if (i.stage === "won") { g.won += 1; g.valueWon += i.value || 0; }
    if (i.stage === "lost") g.lost += 1;
  }
  const rows = Object.values(byGc).sort((a, b) => b.valueSent - a.valueSent);
  if (!rows.length) return null;

  const rate = (g) => (g.won + g.lost ? Math.round((g.won / (g.won + g.lost)) * 100) + "%" : "—");
  const rateColor = (g) => {
    if (!(g.won + g.lost)) return "var(--ink-dim)";
    const r = g.won / (g.won + g.lost);
    return r >= 0.4 ? "var(--green)" : r >= 0.2 ? "#b07d0e" : "var(--red)";
  };

  return (
    <div style={{ marginBottom: 22 }}>
      <div style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 15, marginBottom: 2 }}>Estimates by GC</div>
      <div style={{ fontSize: 12.5, color: "var(--ink-dim)", marginBottom: 10 }}>
        Auto-built from your ITBs — see who awards work and who just collects numbers.
      </div>
      <Table head={["GC / Client", "Bids", "Est. sent", "Awaiting", "Won", "Lost", "Win rate", "Value sent", "Value won"]}>
        {rows.map((g) => (
          <tr key={g.name}>
            <Td><b>{g.name}</b>
              {g.pendingItbs.slice(0, 3).map((i) => (
                <div key={i.id} onClick={() => onOpenItb(i.id)} style={{ color: "var(--copper)", cursor: "pointer", fontSize: 11.5 }}>
                  ↳ {i.name} ({fmt(i.value)})
                </div>
              ))}
            </Td>
            <Td mono>{g.bids}</Td>
            <Td mono>{g.sent}</Td>
            <Td mono>{g.pending || "—"}</Td>
            <Td mono><span style={{ color: g.won ? "var(--green)" : "var(--ink)" }}>{g.won}</span></Td>
            <Td mono>{g.lost || "—"}</Td>
            <Td mono><span style={{ color: rateColor(g), fontWeight: 600 }}>{rate(g)}</span></Td>
            <Td mono>{g.valueSent ? fmt(g.valueSent) : "—"}</Td>
            <Td mono><span style={{ color: g.valueWon ? "var(--green)" : "var(--ink-dim)" }}>{g.valueWon ? fmt(g.valueWon) : "—"}</span></Td>
          </tr>
        ))}
      </Table>
    </div>
  );
}

/* ============== ACCOUNTS ============== */
function Accounts({ data, onAdd, onDel, onOpenItb }) {
  const [f, setF] = useState({ name: "", industry: "", phone: "", website: "" });
  const linked = (a) => {
    const m = (s) => (s || "").toLowerCase().includes(a.name.toLowerCase());
    return {
      itbs: data.itbs.filter((i) => m(i.client) && !["won", "lost"].includes(i.stage)),
      jobs: data.jobs.filter((j) => m(j.client)),
    };
  };
  return (
    <ListPage title="Accounts & GCs" subtitle="Companies you bid for — with a live estimate scoreboard per GC.">
      <GcScoreboard data={data} onOpenItb={onOpenItb} />
      <div style={{ fontFamily: "var(--display)", fontWeight: 700, fontSize: 15, marginBottom: 10 }}>Account records</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <input style={{ flex: 2, minWidth: 180 }} placeholder="Company / GC name" value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} />
        <input style={{ flex: 1, minWidth: 130 }} placeholder="Industry" value={f.industry} onChange={(e) => setF({ ...f, industry: e.target.value })} />
        <input style={{ width: 140 }} placeholder="Phone" value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} />
        <input style={{ flex: 1, minWidth: 140 }} placeholder="Website" value={f.website} onChange={(e) => setF({ ...f, website: e.target.value })} />
        <Btn onClick={() => { if (f.name.trim()) { onAdd(f); setF({ name: "", industry: "", phone: "", website: "" }); } }}>Add</Btn>
      </div>
      <Table head={["Account", "Industry", "Phone", "Open bids", "Opps", "Open value", ""]}>
        {data.accounts.map((a) => {
          const l = linked(a);
          const v = l.itbs.reduce((s, i) => s + (i.value || 0), 0);
          return (
            <tr key={a.id}>
              <Td><b>{a.name}</b>{a.website ? <div style={{ fontSize: 11, color: "var(--ink-dim)" }}>{a.website}</div> : null}</Td>
              <Td dim>{a.industry || "—"}</Td>
              <Td mono>{a.phone || "—"}</Td>
              <Td>{l.itbs.length ? l.itbs.map((i) => (
                <div key={i.id} onClick={() => onOpenItb(i.id)} style={{ color: "var(--copper)", cursor: "pointer", fontSize: 12.5 }}>{i.name}</div>
              )) : <span style={{ color: "var(--ink-dim)" }}>—</span>}</Td>
              <Td mono>{l.jobs.length || "—"}</Td>
              <Td mono>{v ? fmt(v) : "—"}</Td>
              <Td><Del onClick={() => onDel(a.id)} /></Td>
            </tr>
          );
        })}
      </Table>
      {!data.accounts.length && <Empty>No accounts yet — add the GCs and property managers you bid for.</Empty>}
    </ListPage>
  );
}

/* ============== CONTACTS ============== */
function Contacts({ data, onAdd, onDel }) {
  const [f, setF] = useState({ name: "", accountName: "", role: "", phone: "", email: "" });
  return (
    <ListPage title="Contacts" subtitle="The people behind the bids.">
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <input style={{ flex: 1, minWidth: 150 }} placeholder="Full name" value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} />
        <input style={{ flex: 1, minWidth: 150 }} placeholder="Account / company" value={f.accountName} onChange={(e) => setF({ ...f, accountName: e.target.value })} list="p9-accounts" />
        <datalist id="p9-accounts">{data.accounts.map((a) => <option key={a.id} value={a.name} />)}</datalist>
        <input style={{ width: 140 }} placeholder="Role" value={f.role} onChange={(e) => setF({ ...f, role: e.target.value })} />
        <input style={{ width: 140 }} placeholder="Phone" value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} />
        <input style={{ flex: 1, minWidth: 160 }} placeholder="Email" value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} />
        <Btn onClick={() => { if (f.name.trim()) { onAdd(f); setF({ name: "", accountName: "", role: "", phone: "", email: "" }); } }}>Add</Btn>
      </div>
      <Table head={["Name", "Account", "Role", "Phone", "Email", ""]}>
        {data.contacts.map((c) => (
          <tr key={c.id}>
            <Td><b>{c.name}</b></Td>
            <Td dim>{c.accountName || "—"}</Td>
            <Td dim>{c.role || "—"}</Td>
            <Td mono>{c.phone || "—"}</Td>
            <Td mono>{c.email || "—"}</Td>
            <Td><Del onClick={() => onDel(c.id)} /></Td>
          </tr>
        ))}
      </Table>
      {!data.contacts.length && <Empty>No contacts yet.</Empty>}
    </ListPage>
  );
}

/* ============== TASKS ============== */
function Tasks({ data, onAdd, onDel, onToggle, onAssign }) {
  const [f, setF] = useState({ title: "", due: today(), related: "", assignee: "" });
  const [seats, setSeats] = useState([]);
  const [notify, setNotify] = useState({}); // { [taskId]: statusText }
  useEffect(() => { if (usingSharedDb) listSeats().then((s) => setSeats(s || [])).catch(() => setSeats([])); }, []);

  const overdue = (t) => !t.done && t.due && t.due < today();
  const sorted = [...data.tasks].sort((a, b) => (a.done !== b.done ? (a.done ? 1 : -1) : (a.due || "9999") < (b.due || "9999") ? -1 : 1));
  const reset = () => setF({ title: "", due: today(), related: "", assignee: "" });
  const add = () => { if (f.title.trim()) { onAdd(f); reset(); } };

  const notifyAssignee = async (t) => {
    if (!t.assignee) return;
    setNotify((n) => ({ ...n, [t.id]: "Sending…" }));
    try {
      const { ok } = await sendOutlookEmail({
        to: t.assignee,
        subject: `Task assigned to you — ${t.title}`,
        body:
          `You've been assigned a task in Pipeline CRM:\n\n` +
          `• ${t.title}\n` +
          `• Due: ${fmtDate(t.due)}\n` +
          (t.related ? `• Related to: ${t.related}\n` : "") +
          `\n— Pipeline CRM`,
      });
      setNotify((n) => ({ ...n, [t.id]: ok ? "Notified ✓" : "✕ Not sent" }));
    } catch {
      setNotify((n) => ({ ...n, [t.id]: "Email error" }));
    }
  };

  // Options for an assignee <select>; always keeps the current value selectable
  // even if that person no longer holds a seat.
  const seatOptions = (current) => (
    <>
      <option value="">Unassigned</option>
      {seats.map((s) => <option key={s.id} value={s.email}>{s.email}</option>)}
      {current && !seats.some((s) => s.email === current) && <option value={current}>{current}</option>}
    </>
  );

  return (
    <ListPage title="Tasks" subtitle="Follow-ups, call-backs, doc chases — assign them to a teammate so nothing slips.">
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <input style={{ flex: 2, minWidth: 200 }} placeholder="Task — e.g. Call Dana about Riverside docs" value={f.title}
          onChange={(e) => setF({ ...f, title: e.target.value })}
          onKeyDown={(e) => { if (e.key === "Enter") add(); }} />
        <input style={{ width: 150 }} type="date" value={f.due} onChange={(e) => setF({ ...f, due: e.target.value })} />
        <input style={{ flex: 1, minWidth: 140 }} placeholder="Related to (ITB, opp, account…)" value={f.related} onChange={(e) => setF({ ...f, related: e.target.value })} />
        <select style={{ width: 170 }} value={f.assignee} onChange={(e) => setF({ ...f, assignee: e.target.value })} title="Assign to a team member">
          {seatOptions(f.assignee)}
        </select>
        <Btn onClick={add}>Add</Btn>
      </div>
      {usingSharedDb && seats.length === 0 && (
        <div style={{ fontSize: 12, color: "var(--ink-dim)", marginBottom: 10 }}>Add teammates on the Team tab to assign tasks to them.</div>
      )}
      {sorted.map((t) => (
        <div key={t.id} className="p9-card" style={{ display: "flex", gap: 12, alignItems: "center", padding: "11px 14px", background: "var(--panel)", border: "1px solid " + (overdue(t) ? "var(--red)" : "var(--line)"), borderRadius: 4, marginBottom: 8, boxShadow: "var(--shadow-sm)", opacity: t.done ? 0.55 : 1 }}>
          <input type="checkbox" checked={!!t.done} onChange={(e) => onToggle(t.id, e.target.checked)} style={{ width: "auto" }} />
          <span style={{ flex: 1, fontSize: 14, textDecoration: t.done ? "line-through" : "none", minWidth: 120 }}>
            {t.title}{t.related ? <span style={{ color: "var(--ink-dim)" }}> — {t.related}</span> : null}
          </span>
          <select value={t.assignee || ""} title="Assignee"
            onChange={(e) => { onAssign(t.id, e.target.value); setNotify((n) => ({ ...n, [t.id]: undefined })); }}
            style={{ width: 160, fontSize: 12, padding: "6px 8px" }}>
            {seatOptions(t.assignee)}
          </select>
          {t.assignee && (
            <button onClick={() => notifyAssignee(t)} title={"Email " + t.assignee + " about this task"}
              style={{ background: "none", border: "1px solid var(--line)", borderRadius: "var(--radius)", color: notify[t.id] === "Notified ✓" ? "var(--green)" : "var(--copper)", cursor: "pointer", fontSize: 12, padding: "5px 9px", whiteSpace: "nowrap" }}>
              {notify[t.id] || "✉ Notify"}
            </button>
          )}
          <span style={{ fontFamily: "var(--mono)", fontSize: 11.5, color: overdue(t) ? "var(--red)" : "var(--ink-dim)", fontWeight: overdue(t) ? 600 : 400, whiteSpace: "nowrap" }}>
            {overdue(t) ? "OVERDUE · " : ""}{fmtDate(t.due)}
          </span>
          <Del onClick={() => onDel(t.id)} />
        </div>
      ))}
      {!data.tasks.length && <Empty>No tasks yet — press Enter in the box above to add fast.</Empty>}
    </ListPage>
  );
}

const ListPage = ({ title, subtitle, children }) => (
  <div style={{ padding: "0 24px 24px" }} className="p9-rise">
    <h2 style={{ fontFamily: "var(--display)", fontSize: 19, fontWeight: 700, letterSpacing: -0.2, margin: "6px 0 2px", color: "var(--ink)", display: "flex", alignItems: "center", gap: 10 }}>
      <span style={{ width: 4, height: 20, background: "var(--copper)", borderRadius: 2, display: "inline-block" }} />{title}</h2>
    <div style={{ fontSize: 13, color: "var(--ink-dim)", marginBottom: 16, marginLeft: 14 }}>{subtitle}</div>
    {children}
  </div>
);

/* ============== shared bits ============== */
const Overlay = ({ children, onClose, narrow }) => (
  <div onClick={onClose}
    style={{ position: "fixed", inset: 0, background: "rgba(15,35,60,0.35)", display: "flex", justifyContent: "flex-end", zIndex: 50 }}>
    <div onClick={(e) => e.stopPropagation()}
      className="p9-rise"
      style={{ width: narrow ? 470 : 560, maxWidth: "100%", height: "100%", overflowY: "auto", background: "#ffffff", borderLeft: "1px solid var(--line)", padding: 26, boxShadow: "-12px 0 40px rgba(12,37,67,.18)", borderRadius: 0 }}>
      {children}
    </div>
  </div>
);

const CloseBtn = ({ onClick }) => (
  <button onClick={onClick} style={{ background: "none", border: "1px solid var(--line)", color: "var(--ink-dim)", width: 30, height: 30, cursor: "pointer", fontSize: 15 }}>✕</button>
);

const Btn = ({ children, onClick, ghost, color, disabled }) => (
  <button onClick={onClick} disabled={disabled} className="p9-btn"
    style={{
      background: ghost ? "transparent" : color || "var(--copper)",
      color: ghost ? color || "var(--ink-dim)" : "#ffffff",
      border: "1px solid " + (ghost ? color || "var(--line)" : "transparent"),
      padding: "9px 18px", fontFamily: "var(--display)", fontSize: 13.5, letterSpacing: 0.8, borderRadius: 4,
      cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.45 : 1, fontWeight: 700,
      boxShadow: "none",
    }}>
    {children}
  </button>
);

const Field = ({ label, children, style }) => (
  <div style={{ marginTop: 14, ...style }}>
    <div style={{ fontFamily: "var(--display)", fontSize: 11, letterSpacing: 2, color: "var(--ink-dim)", textTransform: "uppercase", marginBottom: 5 }}>{label}</div>
    {children}
  </div>
);

const SectionTitle = ({ children }) => (
  <div style={{ fontFamily: "var(--display)", fontSize: 13, letterSpacing: 2.5, color: "var(--copper)", textTransform: "uppercase", margin: "22px 0 8px", borderBottom: "1px solid var(--line)", paddingBottom: 5 }}>
    {children}
  </div>
);

const Table = ({ head, children }) => (
  <div className="p9-scroll" style={{ overflowX: "auto", border: "1px solid var(--line)", borderRadius: 10, boxShadow: "var(--shadow-sm)" }}><table className="p9-tbl" style={{ width: "100%", borderCollapse: "collapse", background: "var(--panel)" }}>
    <thead>
      <tr>
        {head.map((h, i) => (
          <th key={i} style={{ textAlign: "left", padding: "10px 14px", fontFamily: "var(--display)", fontSize: 11, fontWeight: 700, letterSpacing: 0.6, color: "var(--ink-dim)", textTransform: "uppercase", borderBottom: "1px solid var(--line)", whiteSpace: "nowrap" }}>{h}</th>
        ))}
      </tr>
    </thead>
    <tbody>{children}</tbody>
  </table></div>
);

const Td = ({ children, dim, mono }) => (
  <td style={{ padding: "10px 14px", borderBottom: "1px solid var(--line)", fontSize: 13, color: dim ? "var(--ink-dim)" : "var(--ink)", fontFamily: mono ? "var(--mono)" : "var(--body)", verticalAlign: "top" }}>
    {children}
  </td>
);

const Del = ({ onClick }) => (
  <button onClick={onClick} style={{ background: "none", border: "none", color: "#8da1b5", cursor: "pointer", fontSize: 14 }}>✕</button>
);

const Empty = ({ children }) => (
  <div style={{ color: "var(--ink-dim)", fontSize: 13.5, padding: "28px 20px", textAlign: "center", background: "var(--panel2)", border: "1px dashed var(--line)", borderRadius: 10 }}>{children}</div>
);
