# Simple CRM – prospects, clients, funnel, weekly reminders

Flask app. Runs on **Vercel + Postgres** (recommended), or on your own PC/server with SQLite.

## Features
- **Secure login** (Login ID + password): hashed passwords, CSRF protection, 5-failure lockout (15 min), session timeout, forced password change at first login, security headers.
- **Admin controls everything**: add / disable users, reset passwords, unlock accounts, set reminder e-mail IDs and weekday, quarter convention, view the activity log, and delete records (single records, or everything via **Delete data** with a typed confirmation + password). Regular users can add/edit records, import/export and see reports.
- **Prospects**: stage (Lead → Qualified → Proposal → Negotiation, or Lost), probability %, likely closure date, expected value.
- **Clients (converted)**: closed date (quarter derived automatically) or just a closed quarter, deal value. One-click *Convert* from a prospect.
- **Excel upload** (`Prospects` and `Clients` sheets, or one sheet with a `Type` column): downloadable template, validation, dry-run option, row-by-row error report. **Excel export** of everything.
- **Funnel reporting**: dashboard + Reports page (funnel by stage with value and probability-weighted value, forecast by closure quarter, won by quarter, pipeline by owner, by lead source, overdue and upcoming closures).
- **Weekly e-mail reminder** listing every open prospect (overdue flagged) with a funnel summary, to the address(es) the admin sets.

---

## Deploy on Vercel

Vercel has no persistent disk and no always-on process, so this app uses a hosted **Postgres** database and **Vercel Cron** for the weekly e-mail. Both are already wired up.

### 1. Import the project
Vercel dashboard → **Add New… → Project** → import the GitHub repo. Framework preset is detected automatically (Flask); no build settings needed. Don't deploy yet – set up the database and variables first (or deploy, then redeploy after step 3).

### 2. Add a Postgres database
Project → **Storage** (or the **Marketplace**) → add **Neon** (free tier is plenty) and connect it to the project. This adds `DATABASE_URL` / `POSTGRES_URL` automatically. Tables are created on first request.

### 3. Set environment variables
Project → **Settings → Environment Variables** (Production). See `.env.example`.

| Variable | Value |
|---|---|
| `CRM_SECRET_KEY` | **Required.** Long random string: `python -c "import secrets; print(secrets.token_hex(32))"` |
| `CRON_SECRET` | **Required for reminders.** Another random string; Vercel sends it to the cron job automatically |
| `CRM_ADMIN_USER` / `CRM_ADMIN_PASSWORD` | First admin login (created on the first request). You are forced to change the password at first sign-in. Set these **before** the first visit |
| `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `SMTP_FROM` | Mail account for the reminder (Gmail: use an *App Password*, host `smtp.gmail.com`, port `587`) |

Then **redeploy** so the variables take effect.

### 4. Sign in and configure
Open the deployed URL → sign in as the admin → change the password → **Settings**: enter the reminder e-mail ID(s) and the weekday → **Send reminder now** to test.

### How the weekly reminder works on Vercel
`vercel.json` runs a cron job **daily at 03:30 UTC (09:00 IST)**. The app only sends on the weekday the admin picked in Settings, and at most once per day. To change the time of day, edit the `schedule` in `vercel.json` (UTC) and redeploy. Vercel only runs crons on **Production** deployments. A daily schedule is used because it is accepted on every plan (some plans limit how often, and how precisely, crons fire – check Vercel's current cron limits for yours). The app doesn't depend on the exact minute it fires.
Check **Settings → Last run** to see whether the last send succeeded.

### Limits to know about
- Excel uploads: **4 MB and 5,000 rows per file** (Vercel's request limit); split larger files and upload them one after another – existing records are matched on Company + E-mail, so nothing is duplicated.
- Requests time out after 60 s (`vercel.json`).
- Keep the **Neon database** as your source of truth: use *Export* regularly as an extra backup.

---

## Run locally (SQLite, no setup)
```bash
python -m venv venv && source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py                                         # http://127.0.0.1:5000
```
On the **first run** an admin account is created and its one-time password is printed in the console (or set `CRM_ADMIN_USER` / `CRM_ADMIN_PASSWORD` beforehand). Forgot it? `flask --app app reset-password`.
Data is stored in `instance/crm.db` (git-ignored – back it up). Set `DATABASE_URL` to use Postgres locally too.

Locally the reminder is sent by a built-in scheduler (needs the app running at that time; day/hour from Settings). Or disable it with `CRM_DISABLE_SCHEDULER=1` and run `flask --app app send-reminders` from cron.

## Other hosting (VPS / Docker)
```bash
gunicorn -c gunicorn.conf.py app:app
```
Keep **one worker** (set in `gunicorn.conf.py`) when using SQLite and the built-in scheduler; put it behind HTTPS and set `CRM_SECURE_COOKIES=1`.

## Excel format
Download the template from **Import / Export**.
| Sheet | Columns |
|---|---|
| Prospects | Company*, Contact Name, Email, Phone, Owner, Presales Lead, Region / Country, Product, Product / Service, Source, Stage, Probability (%), Expected Close Date, Expected Value (INR), Expected Value (USD), Notes |
| Clients | Company*, Contact Name, Email, Phone, Owner, Presales Lead, Region / Country, Product, Product / Service, Source, Closed Date, Closed Quarter, Deal Value (INR), Deal Value (USD), Notes |

Each amount goes in its own currency column (nothing is converted); region / country and presales lead are free text (older files without them still import). Dates are read as DD-MM-YYYY. Existing records are matched on **Company + E-mail** and updated; others are added. Bad rows are reported and skipped, good rows still import.

## Tests
```bash
python tests/smoke_test.py                                             # SQLite
TEST_DATABASE_URL=postgresql://user@host/emptydb python tests/smoke_test.py   # Postgres (wipes that DB!)
```

## Files
`app.py` (backend) · `templates/` · `public/` (CSS/JS) · `vercel.json` (function + cron) · `requirements.txt` · `.env.example` · `gunicorn.conf.py` (non-Vercel hosting)

## Analytics & look

- **Dashboard** – KPI tiles (open pipeline, weighted forecast, won this quarter, win rate, needs attention), funnel, forecast by month, region split, overdue / upcoming lists.
- **Analytics page** – pipeline by stage, probability and deal size, closed-won by quarter, win rate by region, pipeline by region / lead source / presales lead / sales owner, records added per month, top open deals, stale deals (30+ days idle), and a data-quality table (click a count to see the records missing that detail).
- INR and USD are always tracked separately (toggle at the top of each page); nothing is converted.
- Theme: sidebar SaaS dashboard style. Font: **Bai Jamjuree** (SIL Open Font License, see `public/fonts/LICENSE-BaiJamjuree-OFL.txt`). Fonts and Chart.js are self-hosted in `public/`, so no external requests are made.

## Products

- Every prospect and client has a **Product** (dropdown) and a **Product / Service** type.
- Admins manage the dropdown under **Product master** (add, rename, hide, delete if unused). It starts with mPower Wealth, mPower Asset Management, Pension admin and RTA.
- In an Excel import, a product that is not in the master yet is added to it automatically (spelling is matched case-insensitively to the master). Columns: `Product`, `Product / Service` (placed after Region / Country in the template).
- Analytics: pipeline by product, product vs service, closed-won by product, win rate by product, and a product scorecard; Reports has the same by product / product vs service.

## Assistant (chat)

A read-only chat page ("Assistant" in the sidebar): ask things like *"which deals are overdue?"*, *"compare products by pipeline and win rate"*, *"top 5 open deals by value"*.

- It can **only look things up** through a few fixed read-only functions (pipeline summary, breakdowns, record search, record details, product list). It cannot add, change or delete data, and it never converts between INR and USD.
- Uses any OpenAI-compatible AI service with a **free tier**. Default is Groq (model `openai/gpt-oss-120b`, with automatic fallback to other Groq models); Google Gemini and OpenRouter are built in.
- Set these environment variables (Vercel → Settings → Environment Variables), then redeploy:

| Variable | Meaning |
|---|---|
| `LLM_API_KEY` | API key from the provider (free, no card on Groq / Gemini). Without it the page says "not set up". |
| `LLM_PROVIDER` | `groq` (default), `gemini` or `openrouter` |
| `LLM_MODEL` | optional – override the model name if a provider retires one |
| `LLM_BASE_URL` | optional – any other OpenAI-compatible endpoint |
| `CHAT_PER_USER_HOUR` / `CHAT_PER_DAY` | optional limits (default 30 per user per hour, 300 per day in total) to stay inside free quotas |

- Privacy: the question and the CRM rows needed to answer it are sent to the AI provider. Free plans of some providers may use that data to improve their models – check the provider's terms before using real customer data. The app stores only a usage count, never the questions.
