# Pipeline — Plumbing CRM

ITB pipeline → estimates → won jobs with materials, labor, invoices, notes, AI invoice/estimate reading, and ARR/NARR tracking.

## Launch on Vercel (step by step)

### 1. Put the code on GitHub
1. Go to github.com → **+** (top right) → **New repository**.
2. Name it `pipeline`, keep it **Private**, click **Create repository**.
3. Easiest path (no terminal): on the empty repo page click **uploading an existing file**, drag in everything from this folder (`api/`, `src/`, `index.html`, `package.json`, `vite.config.js`, `.gitignore`, `README.md`), and click **Commit changes**.

   Or with git:
   ```bash
   cd pipeline
   git init && git add . && git commit -m "Pipeline v1"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/pipeline.git
   git push -u origin main
   ```

### 2. Get an Anthropic API key (powers the AI readers)
1. Go to console.anthropic.com → sign in → **API Keys** → **Create Key**.
2. Copy it (starts with `sk-ant-`). You'll paste it into Vercel in the next step. API usage is pay-as-you-go; reading an invoice costs a fraction of a cent. Docs: https://docs.claude.com/en/api/overview

### 3. Deploy on Vercel
1. Go to vercel.com → **Sign Up** → **Continue with GitHub**.
2. **Add New… → Project** → find `pipeline` → **Import**.
3. Framework preset should auto-detect **Vite**. Leave build settings as-is.
4. Expand **Environment Variables** and add:
   - Name: `ANTHROPIC_API_KEY`  Value: your `sk-ant-...` key
5. Click **Deploy**. ~1 minute later you get a live URL like `pipeline.vercel.app`.

### 4. Test it
- Open the URL → you'll see the seeded pipeline. Add an ITB, advance it, win it.
- On the won job's **Estimate** tab, upload a PDF/photo of an estimate → AI splits one-time vs recurring and sets ARR.
- On **Invoices**, upload a supplier invoice photo → AI logs it + adds a note.

## What works where
| Feature | This deploy |
|---|---|
| Pipeline, jobs, ARR/NARR, all tracking | ✅ |
| AI invoice & estimate readers | ✅ (via your API key) |
| Data persistence | ✅ per-browser (localStorage) |
| Outlook alert emails | ⚠️ Needs Microsoft Graph setup — the in-Claude version uses Claude.ai's Microsoft 365 connector, which isn't available to a public website. Buttons will report a send failure until wired to Graph/SMTP. |

## Local development
```bash
npm install
npx vercel dev   # runs Vite + the /api/claude function locally
```
(`npm run dev` works too, but the AI features need `vercel dev` so the API route exists.)

## Launch checklist (GitHub → Supabase → Vercel)

### 1. GitHub
Create a private repo (e.g. `pipeline`), upload everything in this folder (or git push).

### 2. Supabase (database + logins + seats)
1. supabase.com → New project (free tier).
2. **SQL Editor** → paste `schema.sql` → **EDIT THE LINE marked YOUR_EMAIL_HERE with your real email** → Run.
3. Authentication → Sign In / Providers → Email: leave enabled. (Optional, fastest testing: turn OFF "Confirm email" so sign-ups work instantly.)
4. Project Settings → API → copy the Project URL and `anon` key.

### 3. Vercel
1. vercel.com → Add New → Project → import the repo (auto-detects Vite).
2. Environment Variables:
   - `ANTHROPIC_API_KEY` = sk-ant-... (console.anthropic.com — powers the AI readers)
   - `VITE_SUPABASE_URL` = your Supabase project URL
   - `VITE_SUPABASE_ANON_KEY` = your anon key
3. Deploy → you get `your-app.vercel.app`.

### 4. First sign-in + giving out seats
1. Open the URL → **Create account** using the SAME email you put in schema.sql → you're the owner.
2. Go to the **Team** tab → type a teammate's email → **Add seat**.
3. Send them the URL → they hit "Create account" with that exact email → full access.
4. Remove a seat anytime to revoke access instantly. Only owners can manage seats.

### ClearBid integration
Copy `clearbid-bridge.js` into the ClearBid repo (`npm install @supabase/supabase-js` there), add the same two Supabase env vars, and call `publishEstimateToPipeline({...})` when an estimate is created/sent. ClearBid writes to the open `estimates` mailbox table only; Pipeline creates and advances ITBs itself, so the locked-down CRM data stays seat-only.

## Notes & limits
- Keep invoice/estimate uploads under ~3 MB (serverless body limit is 4.5 MB and base64 adds ~33%).
- Data is stored in each browser's localStorage — fine for testing solo. For your team to share one database, the next step is Postgres (e.g. Vercel Postgres or Supabase) + auth.
