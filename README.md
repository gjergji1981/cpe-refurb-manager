# CPE Refurb Manager v2.0 — Vercel + Supabase

**Stack:** React 18 · Vite · Supabase (Postgres) · Vercel

---

## Quick Start (10 minutes total)

### Step 1 — Set up Supabase

1. Go to [supabase.com](https://supabase.com) → **New project** (free)
2. Choose a name, region closest to your users, and a database password
3. Wait ~2 minutes for provisioning
4. Go to **SQL Editor → New query**
5. Paste the entire contents of **`supabase_schema.sql`** → click **Run**
6. Go to **Project Settings → API** → copy:
   - **Project URL** (e.g. `https://abcdefgh.supabase.co`)
   - **anon public key** (long string starting with `eyJ…`)

### Step 2 — Deploy to Vercel

**Option A — Vercel CLI (fastest)**
```bash
unzip cpe-refurb-vercel.zip
cd cpe-refurb-vercel
npm install
npx vercel --prod
```
When prompted for environment variables, add both:
- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY`

**Option B — GitHub (auto-deploy on every push)**
```bash
git init && git add . && git commit -m "initial"
git remote add origin https://github.com/YOU/cpe-refurb.git
git push -u origin main
```
Then: vercel.com → **Add New Project** → import repo → add env vars → **Deploy**

### Step 3 — Test locally
```bash
cp .env.example .env.local
# Edit .env.local with your Supabase credentials
npm run dev
# → http://localhost:5173
```

---

## Demo credentials

| Role | Username | Password | Access |
|---|---|---|---|
| 🛡️ Administrator | `admin` | `admin123` | All views + User Management |
| 📦 Stock Management | `stock` | `stock123` | Dashboard, Intake, Stock, All Devices |
| 🔧 Refurb Partner | `partner` | `partner123` | Partner Portal only |

---

## Full device lifecycle

```
Intake & Triage
  → Register device, assign action (Refurb / Scrap / ECUS)
  → Execute queue

Refurbishment menu (internal)
  → Devices held in queue
  → Operator selects + dispatches to partner

Partner Portal — Refurb View
  → Partner marks Working / Not Working
  → Executes outcome queue

Refurbishment menu — Pending Confirmation
  → Internal confirms or rejects partner outcomes
  → Confirmed → stage "Confirmed"

Partner Portal — Ready to Return
  → Partner sees confirmed devices
  → Selects + releases to transit → stage "In Transit"

Stock → In Transit tab (Admin / Stock only)
  → Confirm physical receipt
  → Working → Ready (Stock)
  → Not Working → Scrap
```

---

## Project structure

```
cpe-refurb-vercel/
├── index.html              # HTML entry point
├── vite.config.js          # Vite config
├── vercel.json             # SPA routing
├── package.json
├── .env.example            # → copy to .env.local
├── supabase_schema.sql     # Run in Supabase SQL Editor
└── src/
    ├── main.jsx            # React root
    ├── App.jsx             # Full application (~4350 lines)
    ├── supabase.js         # Supabase client
    └── db.js               # All database operations
```

---

## Custom domain

1. Vercel project → **Settings → Domains** → Add domain
2. Add a CNAME record at your DNS: `your-subdomain → cname.vercel-dns.com`
3. SSL is automatic

---

## Security notes

- Passwords are stored plain text for this prototype
- Before production: enable Supabase Row Level Security and hash passwords
- The anon key is safe to expose in frontend code (it's public by design)
- Keep your service role key secret — never put it in frontend code
