"""Simple CRM: prospects & clients, Excel import/export, funnel reports,
weekly e-mail reminders, secure login with an admin who controls users/settings.

Run (dev):   python app.py
Run (prod):  gunicorn -c gunicorn.conf.py app:app
"""
import html as htmllib
import io
import logging
import os
import re
import secrets
import smtplib
import sqlite3
import sys
import threading
import time
from datetime import date, datetime, timedelta, timezone
from email.message import EmailMessage
from functools import wraps
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit
from zoneinfo import ZoneInfo

from flask import (Flask, abort, flash, g, redirect, render_template, request,
                   send_file, session, url_for)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from werkzeug.security import check_password_hash, generate_password_hash

log = logging.getLogger("crm")
logging.basicConfig(level=logging.INFO)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ON_VERCEL = bool(os.environ.get("VERCEL"))


def clean_db_url(url):
    """Drop connection-string options libpq does not understand (e.g. Supabase's `supa`)."""
    p = urlsplit(url)
    q = [(k, v) for k, v in parse_qsl(p.query, keep_blank_values=True) if k != "supa"]
    return urlunsplit(p._replace(query=urlencode(q)))


# DATABASE_URL (or POSTGRES_URL, set by Vercel's Neon/Supabase integrations) -> PostgreSQL.
# Without it the app uses a local SQLite file (fine for a PC or a single server).
DATABASE_URL = clean_db_url(os.environ.get("DATABASE_URL") or os.environ.get("POSTGRES_URL") or "")
USE_PG = bool(DATABASE_URL)
if USE_PG:
    import psycopg
if ON_VERCEL and not USE_PG:
    raise RuntimeError("DATABASE_URL is not set. Vercel needs a hosted Postgres database "
                       "(add Neon from the Vercel Marketplace) - see README.")
DATA_DIR = os.environ.get("CRM_DATA_DIR") or os.path.join(BASE_DIR, "instance")
DB_PATH = os.path.join(DATA_DIR, "crm.db")
MAX_IMPORT_ROWS = int(os.environ.get("CRM_IMPORT_MAX_ROWS", "5000"))
LIKE_OP = "ILIKE" if USE_PG else "LIKE"
TZ = ZoneInfo(os.environ.get("CRM_TIMEZONE", "Asia/Kolkata"))

STAGES = ["Lead", "Qualified", "Proposal", "Negotiation", "Lost"]
OPEN_STAGES = STAGES[:-1]
DEFAULT_PROB = {"Lead": 10, "Qualified": 25, "Proposal": 50, "Negotiation": 75, "Lost": 0}
STAGE_SYNONYMS = {
    "new": "Lead", "cold": "Lead", "qualification": "Qualified", "qualify": "Qualified",
    "proposal sent": "Proposal", "quote": "Qualified" if False else "Proposal",
    "quotation": "Proposal", "negotiating": "Negotiation", "closed lost": "Lost",
    "dropped": "Lost",
}
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
PAGE_SIZE = 25
MAX_LOCK_ATTEMPTS = 5
LOCK_MINUTES = 15

SETTING_DEFAULTS = {
    "company_name": "My CRM",
    "currency": "₹",
    "fy_start_month": "1",           # 1 = calendar-year quarters, 4 = April-March FY
    "reminder_enabled": "1",
    "reminder_emails": "",
    "reminder_day": "0",             # Monday
    "reminder_hour": "9",
    "reminder_last_attempt": "",
    "reminder_last_sent": "",
    "reminder_status": "",
}

def schema_statements():
    pk = "SERIAL PRIMARY KEY" if USE_PG else "INTEGER PRIMARY KEY AUTOINCREMENT"
    real = "DOUBLE PRECISION" if USE_PG else "REAL"
    uname = "TEXT NOT NULL" if USE_PG else "TEXT NOT NULL UNIQUE COLLATE NOCASE"
    stmts = [
        f"""CREATE TABLE IF NOT EXISTS users (
            id {pk},
            username {uname},
            email TEXT,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user' CHECK (role IN ('admin','user')),
            active INTEGER NOT NULL DEFAULT 1,
            must_change INTEGER NOT NULL DEFAULT 1,
            pw_version INTEGER NOT NULL DEFAULT 0,
            failed_attempts INTEGER NOT NULL DEFAULT 0,
            locked_until TEXT,
            created_at TEXT NOT NULL,
            last_login TEXT)""",
        f"""CREATE TABLE IF NOT EXISTS records (
            id {pk},
            type TEXT NOT NULL CHECK (type IN ('prospect','client')),
            company TEXT NOT NULL,
            contact_name TEXT, email TEXT, phone TEXT, owner TEXT, source TEXT,
            presales_lead TEXT, region TEXT,
            stage TEXT, probability INTEGER, expected_close TEXT,
            value {real}, value_usd {real},
            closed_date TEXT, closed_quarter TEXT,
            notes TEXT,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
            created_by INTEGER, updated_by INTEGER)""",
        "CREATE INDEX IF NOT EXISTS idx_records_type ON records(type)",
        "CREATE INDEX IF NOT EXISTS idx_records_company_lc ON records(lower(company))",
        "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)",
        f"""CREATE TABLE IF NOT EXISTS audit_log (
            id {pk},
            ts TEXT NOT NULL, username TEXT, action TEXT NOT NULL, detail TEXT)""",
    ]
    if USE_PG:
        stmts.append("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_lc ON users (lower(username))")
    return stmts

# --------------------------------------------------------------------------- helpers
def utcnow():
    return datetime.now(timezone.utc).replace(tzinfo=None, microsecond=0)


def iso(dt):
    return dt.isoformat(sep=" ")


def load_secret_key():
    key = os.environ.get("CRM_SECRET_KEY")
    if key:
        return key
    if ON_VERCEL:
        raise RuntimeError("CRM_SECRET_KEY environment variable must be set on Vercel "
                           "(any long random string).")
    os.makedirs(DATA_DIR, exist_ok=True)
    path = os.path.join(DATA_DIR, "secret.key")
    if not os.path.exists(path):
        try:
            fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, "w") as f:
                f.write(secrets.token_hex(32))
        except FileExistsError:
            pass
    with open(path) as f:
        return f.read().strip()


class Row:
    """Row that supports row["col"], row[0], row.keys() and dict(row) - like sqlite3.Row."""
    __slots__ = ("_cols", "_vals")

    def __init__(self, cols, vals):
        self._cols, self._vals = cols, tuple(vals)

    def __getitem__(self, k):
        if isinstance(k, str):
            try:
                return self._vals[self._cols.index(k)]
            except ValueError:
                raise KeyError(k) from None
        return self._vals[k]

    def keys(self):
        return list(self._cols)

    def __iter__(self):
        return iter(self._vals)

    def __len__(self):
        return len(self._vals)


def _pg_row_factory(cursor):
    cols = [c.name for c in (cursor.description or [])]
    return lambda values: Row(cols, values)


class DB:
    """Thin wrapper so the rest of the code can use `?` placeholders on SQLite and PostgreSQL."""

    def __init__(self, conn):
        self.conn = conn

    @staticmethod
    def _pg(sql):
        return sql.replace("%", "%%").replace("?", "%s")

    def execute(self, sql, params=()):
        if USE_PG:
            sql = self._pg(sql)
        return self.conn.execute(sql, params)

    def executemany(self, sql, rows):
        if USE_PG:
            with self.conn.cursor() as cur:
                cur.executemany(self._pg(sql), rows)
        else:
            self.conn.executemany(sql, rows)

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

    def close(self):
        self.conn.close()


def connect():
    if USE_PG:
        # prepare_threshold=None: works behind pgbouncer / Neon's pooled connection string.
        return DB(psycopg.connect(DATABASE_URL, row_factory=_pg_row_factory,
                                  prepare_threshold=None, connect_timeout=15))
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA journal_mode=WAL")
    return DB(conn)


def get_settings(conn):
    cfg = dict(SETTING_DEFAULTS)
    for r in conn.execute("SELECT key, value FROM settings"):
        cfg[r["key"]] = r["value"]
    return cfg


def set_setting(conn, key, value):
    conn.execute("INSERT INTO settings(key,value) VALUES(?,?) "
                 "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value)))


def audit(conn, username, action, detail=""):
    conn.execute("INSERT INTO audit_log(ts,username,action,detail) VALUES(?,?,?,?)",
                 (iso(utcnow()), username, action, detail[:500]))


def password_problem(pw, username=""):
    if len(pw) < 10:
        return "Password must be at least 10 characters."
    if not re.search(r"[A-Za-z]", pw) or not re.search(r"\d", pw):
        return "Password must contain at least one letter and one number."
    if username and pw.lower() == username.lower():
        return "Password must not be the same as the username."
    return None


NEW_RECORD_COLUMNS = {"presales_lead": "TEXT", "region": "TEXT",
                      "value_usd": "DOUBLE PRECISION" if USE_PG else "REAL"}


def migrate(conn):
    """Add columns introduced after the first release to databases created earlier."""
    if USE_PG:
        have = {r[0] for r in conn.execute(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_name='records' AND table_schema=current_schema()")}
    else:
        have = {r["name"] for r in conn.execute("PRAGMA table_info(records)")}
    for col, typ in NEW_RECORD_COLUMNS.items():
        if col not in have:
            conn.execute(f"ALTER TABLE records ADD COLUMN {col} {typ}")


def init_db():
    conn = connect()
    if USE_PG:   # serialise concurrent cold starts
        conn.execute("SELECT pg_advisory_xact_lock(7242026)")
    for stmt in schema_statements():
        conn.execute(stmt)
    migrate(conn)
    for k, v in SETTING_DEFAULTS.items():
        conn.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO NOTHING", (k, v))
    if not conn.execute("SELECT 1 FROM users").fetchone():
        username = os.environ.get("CRM_ADMIN_USER", "admin")
        pw = os.environ.get("CRM_ADMIN_PASSWORD")
        generated = pw is None
        if generated:
            pw = secrets.token_urlsafe(12)
        conn.execute(
            "INSERT INTO users(username,password_hash,role,must_change,created_at) VALUES(?,?,?,?,?)",
            (username, generate_password_hash(pw), "admin", 1, iso(utcnow())))
        audit(conn, "system", "admin_created", username)
        banner = "=" * 62
        print(f"\n{banner}\n  First run: admin account created\n  username: {username}\n"
              f"  password: {pw if generated else '(from CRM_ADMIN_PASSWORD)'}\n"
              f"  You must change the password at first login.\n{banner}\n",
              file=sys.stderr, flush=True)
    conn.commit()
    conn.close()


# --------------------------------------------------------------------------- dates / quarters
DATE_FORMATS = ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%d-%b-%Y", "%d %b %Y",
                "%d %B %Y", "%d/%m/%y", "%d-%m-%y", "%b %d, %Y"]


def parse_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    s = s.split(" 00:00:00")[0]
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"unrecognised date '{s}' (use DD-MM-YYYY)")


def quarter_label(d, fy_start=1):
    if fy_start == 1:
        return f"{d.year}-Q{(d.month - 1) // 3 + 1}"
    q = ((d.month - fy_start) % 12) // 3 + 1
    fy_end = d.year + (1 if d.month >= fy_start else 0)
    return f"FY{fy_end - 1}-{str(fy_end)[-2:]} Q{q}"


def normalise_quarter_text(text, fy_start):
    """Tidy hand-typed quarters like 'Q3 2026' -> '2026-Q3' (calendar mode only)."""
    text = (text or "").strip()
    if not text or fy_start != 1:
        return text[:30] or None
    m = re.fullmatch(r"(?i)q([1-4])\W*(\d{4})", text) or None
    if m:
        return f"{m.group(2)}-Q{m.group(1)}"
    m = re.fullmatch(r"(?i)(\d{4})\W*q([1-4])", text)
    if m:
        return f"{m.group(1)}-Q{m.group(2)}"
    return text[:30]


# --------------------------------------------------------------------------- record parsing
def clean_text(v, maxlen=500):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s[:maxlen] if s else None


def parse_stage(v):
    s = clean_text(v)
    if not s:
        return "Lead"
    low = s.lower()
    for st in STAGES:
        if st.lower() == low:
            return st
    if low in STAGE_SYNONYMS:
        return STAGE_SYNONYMS[low]
    raise ValueError(f"unknown stage '{s}' (use {', '.join(STAGES)})")


def parse_prob(v):
    if v in (None, ""):
        return None
    if isinstance(v, str):
        s = v.strip().rstrip("%").strip()
        if not s:
            return None
        n = float(s)
        pct = n
    else:
        n = float(v)
        pct = n * 100 if 0 < n <= 1 else n   # Excel % cells arrive as 0.6
    if not 0 <= pct <= 100:
        raise ValueError("probability must be between 0 and 100")
    return int(round(pct))


def parse_money(v):
    if v in (None, ""):
        return None
    if isinstance(v, str):
        raw = v.strip()
        if raw.lower() in ("", "-", "--", "—", "n/a", "na", "nil", "tbd", "none"):
            return None   # common ways of writing "no amount"
        s = re.sub(r"[^\d.\-]", "", raw)
        if not re.search(r"\d", s):
            raise ValueError("not a number")   # e.g. "abc" - don't silently drop it
        v = s
    n = float(v)
    if n < 0:
        raise ValueError("value cannot be negative")
    return round(n, 2)


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def normalize(raw, kind, fy_start=1):
    """Validate/clean one record. Returns (data, errors)."""
    errors = []
    d = {"type": kind}
    d["company"] = clean_text(raw.get("company"), 200)
    if not d["company"]:
        errors.append("Company is required")
    d["contact_name"] = clean_text(raw.get("contact_name"), 120)
    d["owner"] = clean_text(raw.get("owner"), 80)
    d["presales_lead"] = clean_text(raw.get("presales_lead"), 80)
    d["region"] = clean_text(raw.get("region"), 80)
    d["source"] = clean_text(raw.get("source"), 80)
    d["notes"] = clean_text(raw.get("notes"), 2000)
    email = clean_text(raw.get("email"), 200)
    if email and not EMAIL_RE.match(email):
        errors.append(f"invalid email '{email}'")
        email = None
    d["email"] = email.lower() if email else None
    phone = clean_text(raw.get("phone"), 30)
    if phone and len(re.findall(r"\d", phone)) < 5:
        errors.append(f"invalid phone '{phone}'")
        phone = None
    d["phone"] = phone

    def safe(fn, field, label):
        try:
            return fn(raw.get(field))
        except (ValueError, TypeError) as e:
            msg = str(e)
            errors.append(f"{label}: {msg}" if not msg.startswith("could not") else f"{label}: not a number")
            return None

    d["value"] = safe(parse_money, "value", "Value (INR)")
    d["value_usd"] = safe(parse_money, "value_usd", "Value (USD)")
    if kind == "prospect":
        d["stage"] = safe(parse_stage, "stage", "Stage") or "Lead"
        prob = safe(parse_prob, "probability", "Probability")
        d["probability"] = prob if prob is not None else DEFAULT_PROB[d["stage"]]
        d["expected_close"] = safe(parse_date, "expected_close", "Expected close date")
        d["closed_date"] = d["closed_quarter"] = None
    else:
        d["stage"] = d["probability"] = d["expected_close"] = None
        cd = safe(parse_date, "closed_date", "Closed date")
        d["closed_date"] = cd
        if cd:
            d["closed_quarter"] = quarter_label(cd, fy_start)
        else:
            d["closed_quarter"] = normalise_quarter_text(clean_text(raw.get("closed_quarter"), 30), fy_start)
            if not d["closed_quarter"] and "Closed date" not in " ".join(errors):
                errors.append("Closed date or closed quarter is required for a client")
    for k in ("expected_close", "closed_date"):
        if d.get(k):
            d[k] = d[k].isoformat()
    return d, errors


COLS = ["type", "company", "contact_name", "email", "phone", "owner", "presales_lead", "region",
        "source", "stage",
        "probability", "expected_close", "value", "value_usd", "closed_date", "closed_quarter", "notes"]


INSERT_SQL = (f"INSERT INTO records ({','.join(COLS)},created_at,updated_at,created_by,updated_by) "
              f"VALUES ({','.join('?' * len(COLS))},?,?,?,?)")
UPDATE_SQL = f"UPDATE records SET {','.join(c + '=?' for c in COLS)}, updated_at=?, updated_by=? WHERE id=?"


def insert_record(conn, d, uid):
    now = iso(utcnow())
    cur = conn.execute(INSERT_SQL + " RETURNING id", [d.get(c) for c in COLS] + [now, now, uid, uid])
    return cur.fetchone()[0]


def update_record(conn, rid, d, uid):
    conn.execute(UPDATE_SQL, [d.get(c) for c in COLS] + [iso(utcnow()), uid, rid])


# --------------------------------------------------------------------------- Excel import/export
HEADER_ALIASES = {
    "company": ["company", "company name", "customer", "customer name", "client", "client name",
                "organization", "organisation", "account", "account name", "name"],
    "contact_name": ["contact", "contact name", "contact person", "person"],
    "email": ["email", "email id", "e-mail", "email address", "mail id"],
    "phone": ["phone", "mobile", "phone number", "mobile number", "contact number"],
    "owner": ["owner", "sales owner", "account owner", "sales rep", "assigned to", "salesperson"],
    "source": ["source", "lead source"],
    "notes": ["notes", "remarks", "comments", "note"],
    "value": ["value", "deal value", "expected value", "amount", "deal size", "revenue",
              "order value", "deal amount",
              "value (inr)", "deal value (inr)", "expected value (inr)", "amount (inr)", "value inr",
              "deal value inr", "expected value inr", "amount inr", "inr", "inr value", "amount in inr",
              "value in inr", "expected value in inr", "deal value in inr", "₹"],
    "value_usd": ["value (usd)", "deal value (usd)", "expected value (usd)", "amount (usd)", "value usd",
                  "deal value usd", "expected value usd", "amount usd", "usd", "usd value", "amount in usd",
                  "value in usd", "expected value in usd", "deal value in usd", "$"],
    "type": ["type", "record type", "category"],
    "presales_lead": ["presales lead", "pre-sales lead", "pre sales lead", "presales", "pre-sales",
                      "presales owner", "presales engineer", "presales consultant", "presales resource",
                      "presales spoc", "pre-sales spoc", "presales name"],
    "region": ["region", "country", "region/country", "region / country", "region or country",
               "country/region", "country / region", "geo", "geography", "territory", "location"],
}
PROSPECT_ALIASES = {
    "stage": ["stage", "pipeline stage", "status"],
    "probability": ["probability", "probability (%)", "probability %", "prob", "prob %",
                    "win probability", "chance"],
    "expected_close": ["expected close date", "expected closure date", "likely closure date",
                       "likely close date", "expected close", "closure date", "close date",
                       "expected date", "target close date"],
}
CLIENT_ALIASES = {
    "closed_date": ["closed date", "closure date", "close date", "date closed", "won date",
                    "date of closure", "closing date", "converted date"],
    "closed_quarter": ["closed quarter", "closure quarter", "quarter", "quarter of closure",
                       "closing quarter"],
}


def norm_header(h):
    return re.sub(r"\s+", " ", re.sub(r"[*:]", " ", str(h or ""))).strip().lower()


def map_headers(header, kind):
    aliases = dict(HEADER_ALIASES)
    if kind == "prospect":
        aliases.update(PROSPECT_ALIASES)
    elif kind == "client":
        aliases.update(CLIENT_ALIASES)
    else:  # mixed sheet: unambiguous aliases only
        aliases["stage"] = ["stage", "pipeline stage"]
        aliases["probability"] = PROSPECT_ALIASES["probability"]
        aliases["expected_close"] = ["expected close date", "expected closure date",
                                     "likely closure date", "likely close date", "expected close"]
        aliases["closed_date"] = ["closed date", "date closed", "won date", "date of closure",
                                  "closing date", "converted date"]
        aliases["closed_quarter"] = CLIENT_ALIASES["closed_quarter"]
    lookup = {}
    for field, names in aliases.items():
        for n in names:
            lookup.setdefault(n, field)
    colmap, seen = {}, set()
    for i, h in enumerate(header):
        field = lookup.get(norm_header(h))
        if field and field not in seen:
            colmap[i] = field
            seen.add(field)
    return colmap


def row_kind(v):
    s = (clean_text(v) or "").lower()
    if s in ("prospect", "prospects", "lead", "leads"):
        return "prospect"
    if s in ("client", "clients", "customer", "customers", "converted", "won"):
        return "client"
    return None


def import_workbook(conn, fileobj, fy_start, update_existing, dry_run, uid):
    wb = load_workbook(fileobj, read_only=True, data_only=True)
    res = {"created": 0, "updated": 0, "skipped": 0, "errors": [], "sheets": []}
    # One lookup up front + batched writes at the end: keeps imports fast on a remote database.
    known = {}
    for r in conn.execute("SELECT id, type, company, email FROM records"):
        known[(r["type"], (r["company"] or "").lower(), (r["email"] or "").lower())] = r["id"]
    inserts, updates, total_rows, limit_hit = [], [], 0, False
    now = iso(utcnow())
    for ws in wb.worksheets:
        name, low = ws.title, ws.title.lower()
        if low in ("instructions", "readme", "help", "notes"):
            continue
        kind = ("client" if ("client" in low or "customer" in low)
                else "prospect" if any(k in low for k in ("prospect", "lead", "pipeline")) else None)
        colmap, header_seen, stats = None, False, {"sheet": name, "rows": 0, "ok": 0}
        for idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row is None or all(c in (None, "") for c in row):
                continue
            if not header_seen:
                header_seen = True
                colmap = map_headers(row, kind)
                fields = set(colmap.values())
                if "company" not in fields:
                    res["errors"].append((name, idx, "No 'Company' column found - sheet skipped"))
                    break
                if kind is None and "type" not in fields:
                    res["errors"].append((name, idx, "Sheet name should contain 'Prospects' or "
                                                     "'Clients', or add a 'Type' column - sheet skipped"))
                    break
                continue
            if total_rows >= MAX_IMPORT_ROWS:
                res["errors"].append((name, idx, f"Row limit ({MAX_IMPORT_ROWS:,} per upload) reached - "
                                                 "remaining rows ignored; upload them in a second file"))
                limit_hit = True
                break
            stats["rows"] += 1
            total_rows += 1
            raw = {colmap[i]: row[i] for i in colmap if i < len(row)}
            k = kind or row_kind(raw.get("type"))
            if not k:
                res["errors"].append((name, idx, "Type must be Prospect or Client"))
                continue
            d, errs = normalize(raw, k, fy_start)
            if errs:
                res["errors"].append((name, idx, "; ".join(errs)))
                continue
            key = (k, d["company"].lower(), (d["email"] or "").lower())
            if key in known:
                rid = known[key]
                if rid is None:
                    res["errors"].append((name, idx, "Duplicate of an earlier row in this file - skipped"))
                    continue
                if not update_existing:
                    res["skipped"] += 1
                    continue
                updates.append([d.get(c) for c in COLS] + [now, uid, rid])
                res["updated"] += 1
            else:
                inserts.append([d.get(c) for c in COLS] + [now, now, uid, uid])
                known[key] = None
                res["created"] += 1
            stats["ok"] += 1
        res["sheets"].append(stats)
        if limit_hit:
            break
    wb.close()
    if not dry_run:
        if inserts:
            conn.executemany(INSERT_SQL, inserts)
        if updates:
            conn.executemany(UPDATE_SQL, updates)
    return res


def xl_safe(v):
    """Neutralise spreadsheet formula injection on export."""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + v
    return v


HEAD_FILL = PatternFill("solid", fgColor="1F3A5F")
HEAD_FONT = Font(bold=True, color="FFFFFF")


def style_header(ws, widths):
    for c in ws[1]:
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "B2"


PROSPECT_HEADERS = ["Company*", "Contact Name", "Email", "Phone", "Owner", "Presales Lead",
                    "Region / Country", "Source", "Stage",
                    "Probability (%)", "Expected Close Date", "Expected Value (INR)", "Expected Value (USD)", "Notes"]
CLIENT_HEADERS = ["Company*", "Contact Name", "Email", "Phone", "Owner", "Presales Lead",
                  "Region / Country", "Source",
                  "Closed Date", "Closed Quarter", "Deal Value (INR)", "Deal Value (USD)", "Notes"]


def build_template():
    wb = Workbook()
    info = wb.active
    info.title = "Instructions"
    lines = [
        "CRM import template",
        "",
        "Fill the 'Prospects' sheet for open opportunities and the 'Clients' sheet for converted customers.",
        "Keep the header row. Columns marked * are required. Delete nothing but the data rows you do not need.",
        "Amounts: put the rupee amount in the (INR) column and the dollar amount in the (USD) column. Either can be left blank; nothing is converted.",
        "Region / Country and Presales Lead are free text - spell them the same way every time (India, UAE, USA...) so reports group correctly.",
        "",
        f"Prospects - Stage must be one of: {', '.join(STAGES)} (blank = Lead).",
        "Prospects - Probability is 0-100 (60 or 60% both work). Blank = default for the stage.",
        "Prospects - Expected Close Date: DD-MM-YYYY (e.g. 31-12-2026).",
        "Clients - give a Closed Date (the quarter is worked out automatically) OR a Closed Quarter like 2026-Q3.",
        "Existing records are matched on Company + Email and updated; others are added.",
        "Rows with problems are reported back and skipped; good rows are still imported.",
    ]
    for i, t in enumerate(lines, start=1):
        info.cell(row=i, column=1, value=t)
    info["A1"].font = Font(bold=True, size=14)
    info.column_dimensions["A"].width = 110
    ws = wb.create_sheet("Prospects")
    ws.append(PROSPECT_HEADERS)
    style_header(ws, [28, 22, 28, 16, 16, 18, 18, 16, 14, 14, 20, 20, 20, 40])
    dv = DataValidation(type="list", formula1='"' + ",".join(STAGES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("I2:I2000")
    for r in range(2, 2001):
        ws.cell(row=r, column=11).number_format = "DD-MM-YYYY"
    ws2 = wb.create_sheet("Clients")
    ws2.append(CLIENT_HEADERS)
    style_header(ws2, [28, 22, 28, 16, 16, 18, 18, 16, 16, 16, 18, 18, 40])
    for r in range(2, 2001):
        ws2.cell(row=r, column=9).number_format = "DD-MM-YYYY"
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def build_export(conn):
    wb = Workbook()
    ws = wb.active
    ws.title = "Prospects"
    ws.append(PROSPECT_HEADERS)
    for r in conn.execute("SELECT * FROM records WHERE type='prospect' ORDER BY lower(company)"):
        ws.append([xl_safe(x) for x in (
            r["company"], r["contact_name"], r["email"], r["phone"], r["owner"], r["presales_lead"],
            r["region"], r["source"], r["stage"], r["probability"], r["expected_close"], r["value"], r["value_usd"], r["notes"])])
    style_header(ws, [28, 22, 28, 16, 16, 18, 18, 16, 14, 14, 20, 20, 20, 40])
    ws2 = wb.create_sheet("Clients")
    ws2.append(CLIENT_HEADERS)
    for r in conn.execute("SELECT * FROM records WHERE type='client' ORDER BY lower(company)"):
        ws2.append([xl_safe(x) for x in (
            r["company"], r["contact_name"], r["email"], r["phone"], r["owner"], r["presales_lead"],
            r["region"], r["source"], r["closed_date"], r["closed_quarter"], r["value"], r["value_usd"], r["notes"])])
    style_header(ws2, [28, 22, 28, 16, 16, 18, 18, 16, 16, 16, 18, 18, 40])
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# --------------------------------------------------------------------------- reporting
def funnel(conn):
    rows = {r["stage"]: r for r in conn.execute(
        "SELECT stage, COUNT(*) c, COALESCE(SUM(value),0) v, "
        "COALESCE(SUM(value*probability/100.0),0) w, COALESCE(SUM(value_usd),0) vu, "
        "COALESCE(SUM(value_usd*probability/100.0),0) wu FROM records "
        "WHERE type='prospect' GROUP BY stage")}
    out = []
    for st in OPEN_STAGES:
        r = rows.get(st)
        out.append({"stage": st, "count": r["c"] if r else 0,
                    "value": r["v"] if r else 0, "weighted": r["w"] if r else 0,
                    "value_usd": r["vu"] if r else 0, "weighted_usd": r["wu"] if r else 0})
    lost = rows.get("Lost")
    return out, (lost["c"] if lost else 0)


def dashboard_data(conn, cfg):
    fy = int(cfg["fy_start_month"])
    stages, lost = funnel(conn)
    today = datetime.now(TZ).date()
    clients = conn.execute("SELECT COUNT(*) c, COALESCE(SUM(value),0) v, COALESCE(SUM(value_usd),0) vu "
                           "FROM records WHERE type='client'").fetchone()
    cur_q = quarter_label(today, fy)
    won_q = conn.execute("SELECT COUNT(*) c, COALESCE(SUM(value),0) v, COALESCE(SUM(value_usd),0) vu FROM records "
                         "WHERE type='client' AND closed_quarter=?", (cur_q,)).fetchone()
    open_n = sum(s["count"] for s in stages)
    total = clients["c"] + lost
    overdue = conn.execute(
        "SELECT * FROM records WHERE type='prospect' AND stage!='Lost' AND expected_close < ? "
        "ORDER BY expected_close", (today.isoformat(),)).fetchall()
    upcoming = conn.execute(
        "SELECT * FROM records WHERE type='prospect' AND stage!='Lost' AND expected_close BETWEEN ? AND ? "
        "ORDER BY expected_close", (today.isoformat(), (today + timedelta(days=30)).isoformat())).fetchall()
    no_date = conn.execute("SELECT COUNT(*) c FROM records WHERE type='prospect' AND stage!='Lost' "
                           "AND expected_close IS NULL").fetchone()["c"]
    return {
        "stages": stages, "lost": lost, "open_n": open_n,
        "pipeline": sum(s["value"] for s in stages), "weighted": sum(s["weighted"] for s in stages),
        "clients": clients["c"], "won_value": clients["v"], "won_value_usd": clients["vu"],
        "pipeline_usd": sum(s["value_usd"] for s in stages), "weighted_usd": sum(s["weighted_usd"] for s in stages),
        "win_rate": round(100 * clients["c"] / total) if total else None,
        "cur_q": cur_q, "won_q": won_q, "overdue": overdue, "upcoming": upcoming,
        "no_date": no_date, "max_count": max([s["count"] for s in stages] + [1]),
    }


def breakdown(conn, field):
    """Pipeline + client summary grouped by a free-text field, ignoring upper/lower-case differences."""
    assert field in ("region", "presales_lead")
    groups = {}
    for r in conn.execute(f"SELECT {field} f, type, stage, value, value_usd, probability FROM records").fetchall():
        label = (r["f"] or "").strip()
        grp = groups.setdefault(label.lower(), {
            "names": {}, "open_n": 0, "open_value": 0.0, "weighted": 0.0, "open_value_usd": 0.0,
            "weighted_usd": 0.0, "clients": 0, "won_value": 0.0, "won_value_usd": 0.0, "lost": 0})
        if label:
            grp["names"][label] = grp["names"].get(label, 0) + 1
        inr, usd, prob = r["value"] or 0, r["value_usd"] or 0, (r["probability"] or 0) / 100
        if r["type"] == "client":
            grp["clients"] += 1
            grp["won_value"] += inr
            grp["won_value_usd"] += usd
        elif r["stage"] == "Lost":
            grp["lost"] += 1
        else:
            grp["open_n"] += 1
            grp["open_value"] += inr
            grp["open_value_usd"] += usd
            grp["weighted"] += inr * prob
            grp["weighted_usd"] += usd * prob
    out = []
    for key, grp in groups.items():
        decided = grp["clients"] + grp["lost"]
        out.append({"name": max(grp["names"], key=grp["names"].get) if grp["names"] else "(not set)",
                    "open_n": grp["open_n"], "open_value": grp["open_value"], "weighted": grp["weighted"],
                    "open_value_usd": grp["open_value_usd"], "weighted_usd": grp["weighted_usd"],
                    "clients": grp["clients"], "won_value": grp["won_value"],
                    "won_value_usd": grp["won_value_usd"], "lost": grp["lost"],
                    "win_rate": round(100 * grp["clients"] / decided) if decided else None,
                    "unset": key == ""})
    out.sort(key=lambda x: (x["unset"], -(x["open_value"] + x["won_value"]), x["name"].lower()))
    return out


def distinct_values(field):
    """Existing values, used to suggest consistent spellings in the entry form."""
    assert field in ("region", "presales_lead", "owner", "source")
    return [r[0] for r in g.db.execute(
        f"SELECT DISTINCT {field} FROM records WHERE {field} IS NOT NULL AND {field}<>'' ORDER BY 1")]


def report_data(conn, cfg):
    fy = int(cfg["fy_start_month"])
    won = conn.execute(
        "SELECT COALESCE(closed_quarter,'(no quarter)') q, COUNT(*) c, COALESCE(SUM(value),0) v, "
        "COALESCE(SUM(value_usd),0) vu "
        "FROM records WHERE type='client' GROUP BY 1 ORDER BY 1 DESC").fetchall()
    forecast = {}
    for r in conn.execute("SELECT expected_close, value, value_usd, probability FROM records "
                          "WHERE type='prospect' AND stage!='Lost'"):
        if r["expected_close"]:
            q = quarter_label(date.fromisoformat(r["expected_close"]), fy)
        else:
            q = "(no date)"
        f = forecast.setdefault(q, {"q": q, "c": 0, "v": 0.0, "w": 0.0, "vu": 0.0, "wu": 0.0})
        f["c"] += 1
        f["v"] += r["value"] or 0
        f["w"] += (r["value"] or 0) * (r["probability"] or 0) / 100
        f["vu"] += r["value_usd"] or 0
        f["wu"] += (r["value_usd"] or 0) * (r["probability"] or 0) / 100
    by_owner = conn.execute(
        "SELECT COALESCE(owner,'(unassigned)') o, COUNT(*) c, COALESCE(SUM(value),0) v, "
        "COALESCE(SUM(value*probability/100.0),0) w, COALESCE(SUM(value_usd),0) vu, "
        "COALESCE(SUM(value_usd*probability/100.0),0) wu FROM records "
        "WHERE type='prospect' AND stage!='Lost' GROUP BY 1 ORDER BY 3 DESC").fetchall()
    by_source = conn.execute(
        "SELECT COALESCE(source,'(unknown)') s, "
        "SUM(CASE WHEN type='prospect' AND stage!='Lost' THEN 1 ELSE 0 END) p, "
        "SUM(CASE WHEN type='client' THEN 1 ELSE 0 END) c FROM records GROUP BY 1").fetchall()
    by_source = sorted(by_source, key=lambda r: (-(r["p"] + r["c"]), r["s"]))
    return {"won": won, "forecast": sorted(forecast.values(), key=lambda x: x["q"]),
            "by_owner": by_owner, "by_source": by_source,
            "by_region": breakdown(conn, "region"), "by_presales": breakdown(conn, "presales_lead")}


# --------------------------------------------------------------------------- e-mail reminders
def smtp_configured():
    return bool(os.environ.get("SMTP_HOST") and os.environ.get("SMTP_FROM"))


def parse_emails(text):
    out = []
    for part in re.split(r"[,;\s]+", text or ""):
        if part and part not in out:
            out.append(part)
    return out


def send_email(recipients, subject, text, html_body):
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = subject, os.environ["SMTP_FROM"], ", ".join(recipients)
    msg.set_content(text)
    msg.add_alternative(html_body, subtype="html")
    host, port = os.environ["SMTP_HOST"], int(os.environ.get("SMTP_PORT", "587"))
    if os.environ.get("SMTP_SSL") == "1":
        server = smtplib.SMTP_SSL(host, port, timeout=30)
    else:
        server = smtplib.SMTP(host, port, timeout=30)
        if os.environ.get("SMTP_STARTTLS", "1") != "0":
            server.starttls()
    try:
        if os.environ.get("SMTP_USER"):
            server.login(os.environ["SMTP_USER"], os.environ.get("SMTP_PASSWORD", ""))
        server.send_message(msg)
    finally:
        try:
            server.quit()
        except Exception:
            pass


def fmt_date(s):
    try:
        return date.fromisoformat(s).strftime("%d %b %Y")
    except (TypeError, ValueError):
        return "-"


def fmt_money(v, cur="₹"):
    return f"{cur}{(v or 0):,.0f}"


def fmt_amt(inr, usd):
    parts = []
    if inr:
        parts.append(fmt_money(inr, "₹"))
    if usd:
        parts.append(fmt_money(usd, "$"))
    return " / ".join(parts) or "-"


def build_reminder(conn, cfg):
    today = datetime.now(TZ).date()
    cur = "₹"
    rows = conn.execute(
        "SELECT * FROM records WHERE type='prospect' AND stage!='Lost' "
        "ORDER BY CASE WHEN expected_close IS NULL THEN 1 ELSE 0 END, expected_close, lower(company)"
    ).fetchall()
    stages, _ = funnel(conn)
    e = htmllib.escape
    subject = f"Weekly prospect reminder - {len(rows)} open prospect(s) - {today.strftime('%d %b %Y')}"
    t = [f"{cfg['company_name']} - weekly prospect reminder ({today.strftime('%d %b %Y')})", "",
         "FUNNEL"]
    for s in stages:
        t.append(f"  {s['stage']:<12} {s['count']:>3}   {fmt_money(s['value'], cur)} / {fmt_money(s['value_usd'], '$')}   "
                 f"(weighted {fmt_money(s['weighted'], cur)} / {fmt_money(s['weighted_usd'], '$')})")
    t += ["", f"OPEN PROSPECTS ({len(rows)})"]
    h = [f"<h2 style='margin:0 0 8px'>{e(cfg['company_name'])} - weekly prospect reminder</h2>",
         f"<p style='color:#555'>{today.strftime('%A, %d %B %Y')}</p>",
         "<table cellpadding='6' cellspacing='0' style='border-collapse:collapse;font:14px Arial'>"
         "<tr style='background:#1f3a5f;color:#fff'><th align='left'>Stage</th><th>Count</th>"
         "<th align='right'>Value (INR)</th><th align='right'>Weighted (INR)</th>"
         "<th align='right'>Value (USD)</th><th align='right'>Weighted (USD)</th></tr>"]
    for s in stages:
        h.append(f"<tr style='border-bottom:1px solid #ddd'><td>{s['stage']}</td><td align='center'>{s['count']}</td>"
                 f"<td align='right'>{fmt_money(s['value'], cur)}</td><td align='right'>{fmt_money(s['weighted'], cur)}</td>"
                 f"<td align='right'>{fmt_money(s['value_usd'], '$')}</td><td align='right'>{fmt_money(s['weighted_usd'], '$')}</td></tr>")
    h.append("</table><h3 style='margin:20px 0 6px'>Open prospects</h3>"
             "<table cellpadding='6' cellspacing='0' style='border-collapse:collapse;font:13px Arial;width:100%'>"
             "<tr style='background:#1f3a5f;color:#fff'><th align='left'>Company</th><th align='left'>Contact</th>"
             "<th align='left'>Stage</th><th>Prob.</th><th align='left'>Expected close</th>"
             "<th align='right'>Value (INR / USD)</th><th align='left'>Owner</th><th align='left'>Presales</th>"
             "<th align='left'>Region</th></tr>")
    for r in rows:
        flag = ""
        if r["expected_close"] and r["expected_close"] < today.isoformat():
            flag = " [OVERDUE]"
        elif not r["expected_close"]:
            flag = " [no date]"
        t.append(f"  - {r['company']} | {r['contact_name'] or '-'} | {r['stage']} | {r['probability']}% | "
                 f"{fmt_date(r['expected_close'])}{flag} | {fmt_amt(r['value'], r['value_usd'])} | {r['owner'] or '-'} | "
                 f"presales: {r['presales_lead'] or '-'} | {r['region'] or '-'}")
        color = "#b42318" if "OVERDUE" in flag else "#333"
        h.append(f"<tr style='border-bottom:1px solid #eee;color:{color}'><td>{e(r['company'])}</td>"
                 f"<td>{e(r['contact_name'] or '-')}</td><td>{e(r['stage'])}</td><td align='center'>{r['probability']}%</td>"
                 f"<td>{fmt_date(r['expected_close'])}{e(flag)}</td><td align='right'>{fmt_amt(r['value'], r['value_usd'])}</td>"
                 f"<td>{e(r['owner'] or '-')}</td><td>{e(r['presales_lead'] or '-')}</td>"
                 f"<td>{e(r['region'] or '-')}</td></tr>")
    if not rows:
        t.append("  No open prospects.")
        h.append("<tr><td colspan='9'>No open prospects.</td></tr>")
    h.append("</table><p style='color:#888;font-size:12px'>Sent automatically by your CRM.</p>")
    return subject, "\n".join(t), "".join(h)


def send_reminder_now(conn):
    cfg = get_settings(conn)
    recipients = parse_emails(cfg["reminder_emails"])
    if not recipients:
        raise RuntimeError("No reminder e-mail address is set.")
    if not smtp_configured():
        raise RuntimeError("SMTP is not configured on the server (see README).")
    subject, text, html_body = build_reminder(conn, cfg)
    send_email(recipients, subject, text, html_body)
    return len(recipients)


def maybe_send_weekly(conn, ignore_hour=False):
    cfg = get_settings(conn)
    if cfg["reminder_enabled"] != "1" or not parse_emails(cfg["reminder_emails"]):
        return
    now = datetime.now(TZ)
    if now.weekday() != int(cfg["reminder_day"]):
        return
    if not ignore_hour and now.hour < int(cfg["reminder_hour"]):
        return
    today = now.date().isoformat()
    if cfg["reminder_last_attempt"] == today:
        return
    set_setting(conn, "reminder_last_attempt", today)
    conn.commit()
    stamp = now.strftime("%d %b %Y %H:%M")
    try:
        n = send_reminder_now(conn)
        set_setting(conn, "reminder_last_sent", today)
        set_setting(conn, "reminder_status", f"Sent {stamp} to {n} recipient(s)")
        audit(conn, "system", "reminder_sent", f"{n} recipient(s)")
    except Exception as exc:  # network / auth problems must not kill the scheduler
        log.exception("weekly reminder failed")
        set_setting(conn, "reminder_status", f"FAILED {stamp}: {str(exc)[:200]}")
        audit(conn, "system", "reminder_failed", str(exc)[:200])
    conn.commit()


def scheduler_loop():
    while True:
        try:
            conn = connect()
            try:
                maybe_send_weekly(conn)
            finally:
                conn.close()
        except Exception:
            log.exception("scheduler error")
        time.sleep(60)


_scheduler_lock = threading.Lock()
_scheduler_started = False


def start_scheduler():
    global _scheduler_started
    if ON_VERCEL or os.environ.get("CRM_DISABLE_SCHEDULER") == "1":
        return   # serverless: Vercel Cron calls /api/cron/reminders instead
    with _scheduler_lock:
        if _scheduler_started:
            return
        _scheduler_started = True
        threading.Thread(target=scheduler_loop, daemon=True, name="reminders").start()
        log.info("Weekly reminder scheduler started")


# --------------------------------------------------------------------------- Flask app
app = Flask(__name__, static_folder="public", static_url_path="")
app.config.update(
    SECRET_KEY=load_secret_key(),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("CRM_SECURE_COOKIES") == "1" or ON_VERCEL,
    PERMANENT_SESSION_LIFETIME=timedelta(hours=int(os.environ.get("CRM_SESSION_HOURS", "8"))),
    MAX_CONTENT_LENGTH=4 * 1024 * 1024,   # Vercel rejects bodies over 4.5 MB
)
_inited = False


def ensure_init():
    """Create tables / first admin once per process (lazily on Vercel, at startup elsewhere)."""
    global _inited
    if not _inited:
        init_db()
        _inited = True


if not ON_VERCEL:
    ensure_init()


def csrf_token():
    if "_csrf" not in session:
        session["_csrf"] = secrets.token_urlsafe(32)
    return session["_csrf"]


app.jinja_env.globals.update(csrf_token=csrf_token, STAGES=STAGES, OPEN_STAGES=OPEN_STAGES,
                             DAYS=DAYS, DEFAULT_PROB=DEFAULT_PROB, ON_VERCEL=ON_VERCEL,
                             distinct=lambda f: distinct_values(f))


@app.template_filter("fdate")
def _fdate(s):
    return fmt_date(s)


@app.template_filter("money")
def _money(v):
    return fmt_money(v, "₹")


@app.template_filter("usd")
def _usd(v):
    return fmt_money(v, "$")


@app.before_request
def before():
    if request.endpoint == "static":
        return
    if request.method == "POST":
        sent = request.form.get("_csrf", "")
        if not sent or not secrets.compare_digest(sent, session.get("_csrf", "")):
            abort(400)
    ensure_init()
    g.db = connect()
    g.cfg = get_settings(g.db)
    g.user = None
    uid = session.get("uid")
    if uid:
        u = g.db.execute("SELECT * FROM users WHERE id=? AND active=1", (uid,)).fetchone()
        if u and session.get("pwv") == u["pw_version"]:
            g.user = u
        else:
            session.clear()


@app.teardown_request
def close_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


@app.after_request
def secure_headers(resp):
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "same-origin"
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; "
        "frame-ancestors 'none'; form-action 'self'; base-uri 'self'")
    if request.endpoint != "static":
        resp.headers["Cache-Control"] = "no-store"
    if app.config["SESSION_COOKIE_SECURE"]:
        resp.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return resp


def safe_next(target):
    if target and target.startswith("/") and not target.startswith("//") and "\\" not in target:
        return target.rstrip("?")
    return None


def login_required(f):
    @wraps(f)
    def wrapper(*a, **k):
        if not g.user:
            nxt = request.full_path if request.method == "GET" else None
            return redirect(url_for("login", next=nxt))
        if g.user["must_change"] and request.endpoint not in ("change_password", "logout"):
            return redirect(url_for("change_password"))
        return f(*a, **k)
    return wrapper


def admin_required(f):
    @wraps(f)
    @login_required
    def wrapper(*a, **k):
        if g.user["role"] != "admin":
            abort(403)
        return f(*a, **k)
    return wrapper


for code, msg in {400: "Bad request or expired form - please go back and retry.",
                  403: "You do not have permission to do that.",
                  404: "Page not found.", 413: "File too large (limit 4 MB).",
                  500: "Something went wrong."}.items():
    app.register_error_handler(code, lambda e, code=code, msg=msg: (
        render_template("error.html", code=code, msg=msg), code))


# ---- auth
@app.route("/login", methods=["GET", "POST"])
def login():
    if g.user:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        pw = request.form.get("password", "")
        user = g.db.execute("SELECT * FROM users WHERE lower(username)=lower(?)", (username,)).fetchone()
        now = utcnow()
        locked = bool(user and user["locked_until"] and datetime.fromisoformat(user["locked_until"]) > now)
        ok = bool(user) and check_password_hash(user["password_hash"] if user else
                                                generate_password_hash("x"), pw)
        if user and ok and user["active"] and not locked:
            g.db.execute("UPDATE users SET failed_attempts=0, locked_until=NULL, last_login=? WHERE id=?",
                         (iso(now), user["id"]))
            audit(g.db, user["username"], "login")
            g.db.commit()
            session.clear()
            session.permanent = True
            session["uid"], session["pwv"] = user["id"], user["pw_version"]
            return redirect(safe_next(request.args.get("next")) or url_for("dashboard"))
        if locked:
            flash(f"Account temporarily locked after repeated failures. Try again in {LOCK_MINUTES} minutes "
                  "or ask your admin to unlock it.", "error")
        else:
            if user:
                attempts = user["failed_attempts"] + 1
                until = iso(now + timedelta(minutes=LOCK_MINUTES)) if attempts >= MAX_LOCK_ATTEMPTS else None
                g.db.execute("UPDATE users SET failed_attempts=?, locked_until=? WHERE id=?",
                             (0 if until else attempts, until, user["id"]))
                audit(g.db, user["username"], "login_failed")
                g.db.commit()
            flash("Invalid username or password.", "error")
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/account/password", methods=["GET", "POST"])
@login_required
def change_password():
    if request.method == "POST":
        cur, new, new2 = (request.form.get(k, "") for k in ("current", "new", "new2"))
        problem = None
        if not check_password_hash(g.user["password_hash"], cur):
            problem = "Current password is incorrect."
        elif new != new2:
            problem = "New passwords do not match."
        elif new == cur:
            problem = "New password must be different from the current one."
        else:
            problem = password_problem(new, g.user["username"])
        if problem:
            flash(problem, "error")
        else:
            ver = g.user["pw_version"] + 1
            g.db.execute("UPDATE users SET password_hash=?, must_change=0, pw_version=? WHERE id=?",
                         (generate_password_hash(new), ver, g.user["id"]))
            audit(g.db, g.user["username"], "password_changed")
            g.db.commit()
            session["pwv"] = ver
            flash("Password updated.", "ok")
            return redirect(url_for("dashboard"))
    return render_template("change_password.html")


# ---- dashboard & reports
@app.route("/")
@login_required
def dashboard():
    return render_template("dashboard.html", d=dashboard_data(g.db, g.cfg))


@app.route("/reports")
@login_required
def reports():
    return render_template("reports.html", d=dashboard_data(g.db, g.cfg), r=report_data(g.db, g.cfg))


# ---- records
def form_values(rec=None):
    if rec is None:
        return {}
    return {k: rec[k] for k in rec.keys()}


@app.route("/records")
@login_required
def records():
    kind = "client" if request.args.get("type") == "client" else "prospect"
    q = request.args.get("q", "").strip()
    stage = request.args.get("stage", "")
    page = max(request.args.get("page", 1, type=int), 1)
    where, args = ["type=?"], [kind]
    if q:
        where.append("(" + " OR ".join(f"{c} {LIKE_OP} ?" for c in (
            "company", "contact_name", "email", "owner", "presales_lead", "region")) + ")")
        args += [f"%{q}%"] * 6
    if kind == "prospect" and stage in STAGES:
        where.append("stage=?")
        args.append(stage)
    elif kind == "prospect" and stage == "open":
        where.append("stage!='Lost'")
    w = " AND ".join(where)
    total = g.db.execute(f"SELECT COUNT(*) FROM records WHERE {w}", args).fetchone()[0]
    order = ("CASE WHEN expected_close IS NULL THEN 1 ELSE 0 END, expected_close, lower(company)"
             if kind == "prospect" else "closed_date DESC NULLS LAST, lower(company)")
    rows = g.db.execute(f"SELECT * FROM records WHERE {w} ORDER BY {order} LIMIT ? OFFSET ?",
                        args + [PAGE_SIZE, (page - 1) * PAGE_SIZE]).fetchall()
    pages = max((total + PAGE_SIZE - 1) // PAGE_SIZE, 1)
    return render_template("records.html", kind=kind, rows=rows, q=q, stage=stage, page=page,
                           pages=pages, total=total, today=datetime.now(TZ).date().isoformat())


def save_from_form(rid=None):
    kind = "client" if request.form.get("type") == "client" else "prospect"
    raw = {k: request.form.get(k) for k in
           ("company", "contact_name", "email", "phone", "owner", "presales_lead", "region", "source",
            "stage", "probability",
            "expected_close", "value", "value_usd", "closed_date", "closed_quarter", "notes")}
    d, errors = normalize(raw, kind, int(g.cfg["fy_start_month"]))
    if errors:
        for e in errors:
            flash(e, "error")
        return None, raw | {"type": kind}
    if rid:
        update_record(g.db, rid, d, g.user["id"])
        audit(g.db, g.user["username"], "record_updated", f"#{rid} {d['company']}")
    else:
        rid = insert_record(g.db, d, g.user["id"])
        audit(g.db, g.user["username"], "record_created", f"#{rid} {d['company']}")
    g.db.commit()
    return d, rid


@app.route("/records/new", methods=["GET", "POST"])
@login_required
def record_new():
    kind = "client" if request.values.get("type") == "client" else "prospect"
    if request.method == "POST":
        d, res = save_from_form()
        if d:
            flash(f"{d['company']} saved.", "ok")
            return redirect(url_for("records", type=d["type"]))
        return render_template("record_form.html", rec=res, kind=res["type"], is_new=True)
    return render_template("record_form.html", rec={"type": kind, "stage": "Lead", "probability": 10},
                           kind=kind, is_new=True)


@app.route("/records/<int:rid>/edit", methods=["GET", "POST"])
@login_required
def record_edit(rid):
    rec = g.db.execute("SELECT * FROM records WHERE id=?", (rid,)).fetchone()
    if not rec:
        abort(404)
    if request.method == "POST":
        d, res = save_from_form(rid)
        if d:
            flash(f"{d['company']} updated.", "ok")
            return redirect(url_for("records", type=d["type"]))
        return render_template("record_form.html", rec=res | {"id": rid}, kind=res["type"], is_new=False)
    data = form_values(rec)
    kind = rec["type"]
    if request.args.get("convert") == "1" and kind == "prospect":
        kind = data["type"] = "client"
        data["closed_date"] = datetime.now(TZ).date().isoformat()
    return render_template("record_form.html", rec=data, kind=kind, is_new=False,
                           converting=request.args.get("convert") == "1")


@app.route("/records/<int:rid>/delete", methods=["POST"])
@admin_required
def record_delete(rid):
    rec = g.db.execute("SELECT type, company FROM records WHERE id=?", (rid,)).fetchone()
    if not rec:
        abort(404)
    g.db.execute("DELETE FROM records WHERE id=?", (rid,))
    audit(g.db, g.user["username"], "record_deleted", f"#{rid} {rec['company']}")
    g.db.commit()
    flash(f"{rec['company']} deleted.", "ok")
    return redirect(url_for("records", type=rec["type"]))


# ---- import / export
@app.route("/data", methods=["GET", "POST"])
@login_required
def data_page():
    result = None
    if request.method == "POST":
        f = request.files.get("file")
        if not f or not f.filename:
            flash("Choose an Excel (.xlsx) file first.", "error")
        elif not f.filename.lower().endswith(".xlsx"):
            flash("Only .xlsx files are supported. In Excel use File > Save As > Excel Workbook.", "error")
        else:
            dry = request.form.get("dry_run") == "1"
            try:
                result = import_workbook(g.db, f.stream, int(g.cfg["fy_start_month"]),
                                         request.form.get("update_existing") == "1", dry, g.user["id"])
            except Exception as exc:
                g.db.rollback()
                log.warning("import failed: %s", exc)
                flash("Could not read that file - is it a valid .xlsx workbook?", "error")
            else:
                if dry:
                    g.db.rollback()
                else:
                    audit(g.db, g.user["username"], "import",
                          f"{f.filename[:80]}: +{result['created']} ~{result['updated']} !{len(result['errors'])}")
                    g.db.commit()
                result["dry"] = dry
    return render_template("data.html", result=result)


@app.route("/template.xlsx")
@login_required
def download_template():
    return send_file(build_template(), as_attachment=True, download_name="crm_import_template.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/export.xlsx")
@login_required
def download_export():
    audit(g.db, g.user["username"], "export")
    g.db.commit()
    name = f"crm_export_{datetime.now(TZ).strftime('%Y%m%d')}.xlsx"
    return send_file(build_export(g.db), as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---- admin: users
USERNAME_RE = re.compile(r"^[A-Za-z0-9_.@-]{3,40}$")


def active_admin_count(db, exclude=None):
    q = "SELECT COUNT(*) FROM users WHERE role='admin' AND active=1"
    a = []
    if exclude:
        q += " AND id!=?"
        a.append(exclude)
    return db.execute(q, a).fetchone()[0]


@app.route("/admin/users")
@admin_required
def users():
    rows = g.db.execute("SELECT * FROM users ORDER BY role, username").fetchall()
    return render_template("users.html", rows=rows, now=iso(utcnow()))


@app.route("/admin/users/new", methods=["GET", "POST"])
@admin_required
def user_new():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip()
        role = "admin" if request.form.get("role") == "admin" else "user"
        pw = request.form.get("password", "")
        err = None
        if not USERNAME_RE.match(username):
            err = "Username must be 3-40 characters: letters, numbers and . _ @ -"
        elif g.db.execute("SELECT 1 FROM users WHERE lower(username)=lower(?)", (username,)).fetchone():
            err = "That username already exists."
        elif email and not EMAIL_RE.match(email):
            err = "Invalid e-mail address."
        else:
            err = password_problem(pw, username)
        if err:
            flash(err, "error")
        else:
            g.db.execute("INSERT INTO users(username,email,password_hash,role,must_change,created_at) "
                         "VALUES(?,?,?,?,1,?)",
                         (username, email or None, generate_password_hash(pw), role, iso(utcnow())))
            audit(g.db, g.user["username"], "user_created", f"{username} ({role})")
            g.db.commit()
            flash(f"User {username} created. They must change the password at first login.", "ok")
            return redirect(url_for("users"))
    return render_template("user_form.html", u=None, suggestion=secrets.token_urlsafe(9) + "7a")


@app.route("/admin/users/<int:uid>", methods=["GET", "POST"])
@admin_required
def user_edit(uid):
    u = g.db.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if not u:
        abort(404)
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        role = "admin" if request.form.get("role") == "admin" else "user"
        active = 1 if request.form.get("active") == "1" else 0
        pw = request.form.get("password", "")
        err = None
        if email and not EMAIL_RE.match(email):
            err = "Invalid e-mail address."
        elif (role != "admin" or not active) and u["role"] == "admin" and u["active"] \
                and active_admin_count(g.db, exclude=uid) == 0:
            err = "There must be at least one active admin."
        elif uid == g.user["id"] and not active:
            err = "You cannot deactivate your own account."
        elif pw:
            err = password_problem(pw, u["username"])
        if err:
            flash(err, "error")
        else:
            g.db.execute("UPDATE users SET email=?, role=?, active=? WHERE id=?",
                         (email or None, role, active, uid))
            if request.form.get("unlock") == "1":
                g.db.execute("UPDATE users SET failed_attempts=0, locked_until=NULL WHERE id=?", (uid,))
            if pw:
                g.db.execute("UPDATE users SET password_hash=?, must_change=1, pw_version=pw_version+1 "
                             "WHERE id=?", (generate_password_hash(pw), uid))
                audit(g.db, g.user["username"], "password_reset", u["username"])
            audit(g.db, g.user["username"], "user_updated", f"{u['username']} role={role} active={active}")
            g.db.commit()
            if uid == g.user["id"] and pw:
                session["pwv"] = g.db.execute("SELECT pw_version FROM users WHERE id=?", (uid,)).fetchone()[0]
            flash("User updated.", "ok")
            return redirect(url_for("users"))
    return render_template("user_form.html", u=u, suggestion=secrets.token_urlsafe(9) + "7a")


# ---- admin: settings & audit
@app.route("/admin/settings", methods=["GET", "POST"])
@admin_required
def settings():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "test":
            try:
                n = send_reminder_now(g.db)
                flash(f"Reminder sent to {n} recipient(s).", "ok")
                audit(g.db, g.user["username"], "reminder_test", f"{n} recipient(s)")
            except Exception as exc:
                flash(f"Could not send: {str(exc)[:200]}", "error")
            g.db.commit()
            return redirect(url_for("settings"))
        emails = parse_emails(request.form.get("reminder_emails"))
        bad = [x for x in emails if not EMAIL_RE.match(x)]
        company = (request.form.get("company_name") or "My CRM").strip()[:60]
        currency = (request.form.get("currency") or "₹").strip()[:4]
        try:
            fy = int(request.form.get("fy_start_month", 1))
            day, hour = int(request.form.get("reminder_day", 0)), int(request.form.get("reminder_hour", 9))
            assert 1 <= fy <= 12 and 0 <= day <= 6 and 0 <= hour <= 23
        except (ValueError, AssertionError):
            flash("Invalid number in settings.", "error")
            return redirect(url_for("settings"))
        if bad:
            flash(f"Invalid e-mail address: {', '.join(bad)}", "error")
        else:
            old_fy = int(g.cfg["fy_start_month"])
            for k, v in {"company_name": company, "currency": currency, "fy_start_month": fy,
                         "reminder_emails": ", ".join(emails), "reminder_day": day, "reminder_hour": hour,
                         "reminder_enabled": "1" if request.form.get("reminder_enabled") == "1" else "0"}.items():
                set_setting(g.db, k, v)
            if fy != old_fy:   # re-label quarters of dated clients
                for r in g.db.execute("SELECT id, closed_date FROM records "
                                      "WHERE type='client' AND closed_date IS NOT NULL").fetchall():
                    g.db.execute("UPDATE records SET closed_quarter=? WHERE id=?",
                                 (quarter_label(date.fromisoformat(r["closed_date"]), fy), r["id"]))
            audit(g.db, g.user["username"], "settings_updated")
            g.db.commit()
            flash("Settings saved.", "ok")
            return redirect(url_for("settings"))
    return render_template("settings.html", smtp=smtp_configured(), tz=str(TZ))


@app.route("/admin/audit")
@admin_required
def audit_view():
    rows = g.db.execute("SELECT * FROM audit_log ORDER BY id DESC LIMIT 200").fetchall()
    return render_template("audit.html", rows=rows)


# ---- admin: delete data
DELETE_SCOPES = {"prospects": ("prospect",), "clients": ("client",), "all": ("prospect", "client")}


@app.route("/admin/delete-data", methods=["GET", "POST"])
@admin_required
def danger_zone():
    def counts():
        rows = {r["type"]: r["n"] for r in g.db.execute("SELECT type, COUNT(*) n FROM records GROUP BY type")}
        return {"prospects": rows.get("prospect", 0), "clients": rows.get("client", 0)}

    if request.method == "POST":
        scope = request.form.get("scope", "")
        if scope not in DELETE_SCOPES:
            abort(400)
        if request.form.get("confirm", "").strip() != "DELETE":
            flash("Nothing deleted - type DELETE (capital letters) to confirm.", "error")
        elif not check_password_hash(g.user["password_hash"], request.form.get("password", "")):
            audit(g.db, g.user["username"], "delete_data_denied", f"scope={scope}: wrong password")
            g.db.commit()
            flash("Nothing deleted - your admin password was incorrect.", "error")
        else:
            kinds = DELETE_SCOPES[scope]
            n = sum(counts()["prospects" if k == "prospect" else "clients"] for k in kinds)
            g.db.execute(f"DELETE FROM records WHERE type IN ({','.join('?' * len(kinds))})", kinds)
            audit(g.db, g.user["username"], "delete_data", f"scope={scope}: {n} record(s) deleted")
            g.db.commit()
            flash(f"Deleted {n} record(s) ({scope}).", "ok")
            return redirect(url_for("danger_zone"))
    return render_template("danger.html", n=counts())


# ---- Vercel Cron (runs daily; sends only on the admin's chosen weekday)
@app.route("/api/cron/reminders")
def cron_reminders():
    secret = os.environ.get("CRON_SECRET", "")
    sent = request.headers.get("Authorization", "")
    if not secret or not secrets.compare_digest(sent.encode(), f"Bearer {secret}".encode()):
        return "Unauthorized", 401
    maybe_send_weekly(g.db, ignore_hour=True)   # cron timing is only hour-accurate on some plans
    return {"ok": True, "status": get_settings(g.db)["reminder_status"]}


# ---- CLI helpers
@app.cli.command("send-reminders")
def cli_send_reminders():
    """Send the weekly prospect reminder immediately (for cron use)."""
    conn = connect()
    n = send_reminder_now(conn)
    print(f"Sent to {n} recipient(s)")


@app.cli.command("reset-password")
def cli_reset_password():
    """Reset a user's password: flask --app app reset-password USERNAME"""
    import click
    username = click.prompt("Username")
    pw = click.prompt("New password", hide_input=True, confirmation_prompt=True)
    problem = password_problem(pw, username)
    if problem:
        raise click.ClickException(problem)
    conn = connect()
    cur = conn.execute("UPDATE users SET password_hash=?, must_change=1, active=1, failed_attempts=0, "
                       "locked_until=NULL, pw_version=pw_version+1 WHERE lower(username)=lower(?)",
                       (generate_password_hash(pw), username))
    conn.commit()
    print("Password reset." if cur.rowcount else "No such user.")


if __name__ == "__main__":
    start_scheduler()
    app.run(host=os.environ.get("CRM_HOST", "127.0.0.1"), port=int(os.environ.get("PORT", 5000)))
