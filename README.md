# CPE Refurb Manager — Vercel Deployment

## What's inside
- React 18 + Vite frontend
- SQLite-free — all data lives in React state (prototype mode)
- Login screen with 3 roles: Admin, Stock Management, Refurbishment Partner
- Works on desktop and mobile

---

## Option A — Deploy via GitHub (recommended)

### Step 1 — Push to GitHub

1. Create a new repo on github.com (name it e.g. `cpe-refurb-manager`)
2. In your terminal, from this folder:

```bash
npm install
git init
git add .
git commit -m "initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/cpe-refurb-manager.git
git push -u origin main
```

### Step 2 — Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) and sign in (free account is fine)
2. Click **Add New → Project**
3. Import your GitHub repo `cpe-refurb-manager`
4. Vercel auto-detects Vite — leave all settings as default
5. Click **Deploy**
6. Done — your app is live at `https://cpe-refurb-manager.vercel.app`

Every `git push` to `main` automatically redeploys.

---

## Option B — Deploy via Vercel CLI (no GitHub needed)

```bash
# Install Vercel CLI
npm install -g vercel

# Install project dependencies
npm install

# Deploy (follow the prompts)
vercel

# Deploy to production
vercel --prod
```

---

## Run locally

```bash
npm install
npm run dev
# Open http://localhost:5173
```

---

## Demo credentials

| Role | Username | Password |
|---|---|---|
| 🛡️ Administrator | `admin` | `admin123` |
| 📦 Stock Management | `stock` | `stock123` |
| 🔧 Refurbishment Partner | `partner` | `partner123` |

---

## Project structure

```
cpe-refurb-manager/
├── index.html          # HTML entry point
├── vite.config.js      # Vite configuration
├── vercel.json         # SPA routing for Vercel
├── package.json
├── .gitignore
└── src/
    ├── main.jsx        # React root
    └── App.jsx         # Full application (2900+ lines)
```

---

## Custom domain on Vercel

1. In your Vercel project → **Settings → Domains**
2. Add your domain (e.g. `cpe.yourdomain.com`)
3. Add a CNAME record at your DNS provider pointing to `cname.vercel-dns.com`
4. Vercel provisions SSL automatically

---

## Upgrading to a real database later

This prototype stores data in React state (resets on page refresh).
When ready to persist data, the recommended path is:

- **Supabase** — hosted Postgres + REST API, free tier generous
- **PlanetScale** — serverless MySQL, works great with Vercel
- Replace the `useState(SEED)` in `App.jsx` with `useEffect` API calls

See `cpe-refurb-production-guide.md` for the full backend setup.
