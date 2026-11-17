# Simple CRM – prospects, clients, funnel, weekly reminders

Flask + SQLite. No external services needed except an SMTP mailbox for the weekly reminder e-mail.

## Features
- **Secure login** (Login ID + password). Passwords hashed, CSRF protection, 5-failure lockout (15 min), session timeout, forced password change at first login, security headers.
- **Admin controls everything**: add / disable users, reset passwords, unlock accounts, set reminder e-mail IDs and schedule, quarter convention, view the activity log, and delete records. Regular users can add/edit records, import/export and see reports.
- **Prospects**: stage (Lead → Qualified → Proposal → Negotiation, or Lost), probability %, likely closure date, expected value.
- **Clients (converted)**: closed date (quarter is derived automatically) or just a closed quarter, deal value. One-click *Convert* from a prospect.
- **Excel upload** (`Prospects` and `Clients` sheets, or one sheet with a `Type` column) with a downloadable template, validation, a dry-run option, and a row-by-row error report. **Excel export** of everything.
- **Front-end screens** for adding / editing / searching records.
- **Funnel reporting**: dashboard + Reports page (funnel by stage with value and probability-weighted value, forecast by closure quarter, won by quarter, pipeline by owner, by lead source, overdue and upcoming closures).
- **Weekly e-mail reminder** listing every open prospect (overdue flagged) with a funnel summary, to the address(es) the admin sets.

## Run locally
```bash
python -m venv venv && source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py                                         # http://127.0.0.1:5000
```
On the **first run** an admin account is created and its one-time password is printed in the console (or set `CRM_ADMIN_USER` / `CRM_ADMIN_PASSWORD` beforehand). You must change it at first login.
Forgot it? `flask --app app reset-password`

## Weekly reminder e-mail
1. Set the SMTP variables (see `.env.example`) in the environment where the app runs. For Gmail use an *App Password*.
2. Sign in as admin → **Settings** → enter the e-mail ID(s), day and hour (default Monday 09:00 IST) → Save → **Send reminder now** to test.

The scheduler runs inside the app, so the app must be running at that time. If the server was down for the whole scheduled day, that week's mail is skipped. Prefer cron? Disable the in-app scheduler with `CRM_DISABLE_SCHEDULER=1` and run `flask --app app send-reminders` weekly.

## Production
```bash
gunicorn -c gunicorn.conf.py app:app
```
- Keep **one worker** (already set) – SQLite and the in-app scheduler assume it.
- Put it behind HTTPS (nginx / Caddy / your host) and set `CRM_SECURE_COOKIES=1`.
- Data lives in `instance/` (`crm.db`, `secret.key`) – **back this folder up**; it is git-ignored. Set `CRM_DATA_DIR` to put it elsewhere (e.g. a persistent disk). Set `CRM_SECRET_KEY` if the host has an ephemeral filesystem.

## Excel format
Download the template from **Import / Export**. Key rules:
| Sheet | Columns |
|---|---|
| Prospects | Company*, Contact Name, Email, Phone, Owner, Source, Stage, Probability (%), Expected Close Date, Expected Value, Notes |
| Clients | Company*, Contact Name, Email, Phone, Owner, Source, Closed Date, Closed Quarter, Deal Value, Notes |

Dates are read as DD-MM-YYYY. Existing records are matched on **Company + E-mail** and updated; others are added. Bad rows are reported and skipped, good rows still import.

## Files
`app.py` (all backend logic) · `templates/` · `static/` · `gunicorn.conf.py` · `.env.example`
