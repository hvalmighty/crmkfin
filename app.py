"""Simple CRM: prospects & clients, Excel import/export, funnel reports,
weekly e-mail reminders, secure login with an admin who controls users/settings.

Run (dev):   python app.py
Run (prod):  gunicorn -c gunicorn.conf.py app:app
"""
import html as htmllib
import io
import json
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

from flask import (Flask, abort, flash, g, has_request_context, jsonify, redirect, render_template, request,
                   send_file, session, url_for)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
import chat
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
    "company_name": "QUEST FT",
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

DEFAULT_PRODUCTS = ["mPower Wealth", "mPower Asset Management", "Pension admin", "RTA"]
OFFERING_TYPES = ["Product", "Service"]


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
            presales_lead TEXT, region TEXT, product TEXT, offering_type TEXT,
            stage TEXT, probability INTEGER, expected_close TEXT,
            value {real}, value_usd {real},
            closed_date TEXT, closed_quarter TEXT,
            notes TEXT,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
            created_by INTEGER, updated_by INTEGER)""",
        "CREATE INDEX IF NOT EXISTS idx_records_type ON records(type)",
        "CREATE INDEX IF NOT EXISTS idx_records_company_lc ON records(lower(company))",
        "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)",
        f"""CREATE TABLE IF NOT EXISTS products (
            id {pk},
            name {uname},
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL)""",
        f"""CREATE TABLE IF NOT EXISTS targets (
            id {pk},
            rep TEXT NOT NULL,
            fy INTEGER NOT NULL,
            quarter INTEGER NOT NULL,
            value_inr {real}, value_usd {real},
            updated_at TEXT NOT NULL)""",
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_targets_key ON targets (lower(rep), fy, quarter)",
        f"""CREATE TABLE IF NOT EXISTS audit_log (
            id {pk},
            ts TEXT NOT NULL, username TEXT, action TEXT NOT NULL, detail TEXT)""",
    ]
    if USE_PG:
        stmts.append("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_lc ON users (lower(username))")
        stmts.append("CREATE UNIQUE INDEX IF NOT EXISTS idx_products_name_lc ON products (lower(name))")
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

    def __init__(self, conn, pooled=False):
        self.conn = conn
        self.pooled = pooled

    @staticmethod
    def _pg(sql):
        return sql.replace("%", "%%").replace("?", "%s")

    def execute(self, sql, params=()):
        if has_request_context():
            g.nq = getattr(g, "nq", 0) + 1
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
        if not self.pooled:
            self.conn.close()
            return
        if self.conn.autocommit:  # read-only request: no transaction was opened, nothing to end (saves 2 round trips)
            return
        try:                      # keep the connection for the next request, but end this request's transaction
            self.conn.rollback()
        except Exception:
            _tls.conn = None
            try:
                self.conn.close()
            except Exception:
                pass


_tls = threading.local()      # one reusable PostgreSQL connection per thread (web requests only)
IDLE_PING_SECONDS = 20


def _new_pg():
    """Open a PostgreSQL connection. Retries: a sleeping serverless DB (e.g. Neon free tier) can refuse the
    very first connection while it wakes up - that used to show up as a page that only loaded after a refresh."""
    last = None
    for delay in (0, 1.0, 2.5):
        if delay:
            time.sleep(delay)
        try:
            # prepare_threshold=None: works behind pgbouncer / Neon's pooled connection string.
            return psycopg.connect(DATABASE_URL, row_factory=_pg_row_factory, prepare_threshold=None, connect_timeout=10)
        except psycopg.OperationalError as e:
            last = e
    raise last


def connect(reuse=False, autocommit=False):
    """reuse=True (web requests): keep the PostgreSQL connection open between requests handled by the same
    warm instance, which saves the TCP + TLS + login round-trips on every page."""
    if USE_PG:
        if not reuse:
            return DB(_new_pg())
        c = getattr(_tls, "conn", None)
        if c is not None and not c.closed:
            if time.time() - getattr(_tls, "used", 0) > IDLE_PING_SECONDS:
                try:                      # connection may have been dropped while the instance was idle
                    c.execute("SELECT 1")
                    c.rollback()
                except Exception:
                    c = None
        else:
            c = None
        if c is None:
            c = _tls.conn = _new_pg()
        _tls.used = time.time()
        if c.autocommit != autocommit:
            try:
                c.autocommit = autocommit
            except Exception:         # a transaction was left open: end it, then switch
                c.rollback(); c.autocommit = autocommit
        return DB(c, pooled=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA journal_mode=WAL")
    return DB(conn)


def get_settings(conn):
    cfg = dict(SETTING_DEFAULTS)
    for r in conn.execute("SELECT key, value FROM settings"):
        if not r["key"].startswith("_"):
            cfg[r["key"]] = r["value"]
    return cfg


_cfg_cache = {"t": 0.0, "v": None}
CFG_TTL = 15      # seconds a warm instance re-uses the settings (saves one database round-trip per page)


def cached_settings(conn):
    if _cfg_cache["v"] is None or time.time() - _cfg_cache["t"] > CFG_TTL:
        _cfg_cache.update(v=get_settings(conn), t=time.time())
    return dict(_cfg_cache["v"])


def set_setting(conn, key, value):
    _cfg_cache["v"] = None
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


NEW_RECORD_COLUMNS = {"presales_lead": "TEXT", "region": "TEXT", "product": "TEXT", "offering_type": "TEXT",
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


# Bump this whenever the schema / migrations / seed data change, so existing databases get upgraded once.
SCHEMA_VERSION = "2026-11-20.2"


def schema_is_current():
    """One cheap query on a cold start instead of ~20 setup statements."""
    conn = connect()
    try:
        r = conn.execute("SELECT value FROM settings WHERE key='_schema'").fetchone()
        return bool(r) and r["value"] == SCHEMA_VERSION
    except Exception:
        return False                    # table doesn't exist yet -> first run
    finally:
        try:
            conn.rollback()
        except Exception:
            pass
        conn.close()


def init_db():
    if schema_is_current():
        return
    conn = connect()
    if USE_PG:   # serialise concurrent cold starts
        conn.execute("SELECT pg_advisory_xact_lock(7242026)")
    for stmt in schema_statements():
        conn.execute(stmt)
    migrate(conn)
    for k, v in SETTING_DEFAULTS.items():
        conn.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO NOTHING", (k, v))
    if not conn.execute("SELECT 1 FROM products").fetchone():   # first run: seed the product master
        for nm in DEFAULT_PRODUCTS:
            conn.execute("INSERT INTO products(name,active,created_at) VALUES(?,1,?) ON CONFLICT DO NOTHING",
                         (nm, iso(utcnow())))
    # app was renamed from the original default name
    conn.execute("UPDATE settings SET value=? WHERE key='company_name' AND value='My CRM'", ("QUEST FT",))
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
    set_setting(conn, "_schema", SCHEMA_VERSION)
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
    d["product"] = clean_text(raw.get("product"), 120)
    ot = clean_text(raw.get("offering_type"), 40)
    if ot:
        ot_norm = {"product": "Product", "products": "Product", "service": "Service", "services": "Service"}.get(ot.lower())
        if not ot_norm:
            errors.append(f"Product / Service must be 'Product' or 'Service' (got '{ot}')")
        ot = ot_norm
    d["offering_type"] = ot
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
        "product", "offering_type", "source", "stage",
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


# --------------------------------------------------------------------------- product master
def product_map(conn):
    """lower-case name -> canonical name, for every product in the master (active or not)."""
    return {r["name"].lower(): r["name"] for r in conn.execute("SELECT name FROM products")}


def active_products(conn):
    return [r["name"] for r in conn.execute("SELECT name FROM products WHERE active=1 ORDER BY lower(name)")]


def add_product(conn, name):
    conn.execute("INSERT INTO products(name,active,created_at) VALUES(?,1,?) ON CONFLICT DO NOTHING",
                 (name, iso(utcnow())))


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
    "product": ["product", "product name", "products", "product interested", "product / solution",
                "product/solution", "solution", "product or service name", "service name", "offering"],
    "offering_type": ["product / service", "product/service", "product or service", "product / service type",
                      "offering type", "product type", "service type", "product service", "prod / service",
                      "product-service"],
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
    pmap = product_map(conn)
    res["new_products"] = []
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
            if key in known and (known[key] is None or not update_existing):
                pass
            elif d["product"]:      # use the master's spelling; unknown products are added to the master
                canon = pmap.get(d["product"].lower())
                if canon is None:
                    pmap[d["product"].lower()] = canon = d["product"]
                    res["new_products"].append(canon)
                d["product"] = canon
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
        for nm in res["new_products"]:
            add_product(conn, nm)
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
                    "Region / Country", "Product", "Product / Service", "Source", "Stage",
                    "Probability (%)", "Expected Close Date", "Expected Value (INR)", "Expected Value (USD)", "Notes"]
CLIENT_HEADERS = ["Company*", "Contact Name", "Email", "Phone", "Owner", "Presales Lead",
                  "Region / Country", "Product", "Product / Service", "Source",
                  "Closed Date", "Closed Quarter", "Deal Value (INR)", "Deal Value (USD)", "Notes"]


def build_template(products=()):
    wb = Workbook()
    info = wb.active
    info.title = "Instructions"
    lines = [
        "CRM import template",
        "",
        "Fill the 'Prospects' sheet for open opportunities and the 'Clients' sheet for converted customers.",
        "Keep the header row. Columns marked * are required. Delete nothing but the data rows you do not need.",
        "Amounts: put the rupee amount in the (INR) column and the dollar amount in the (USD) column. Either can be left blank; nothing is converted.",
        "Product: use one of the names in the Product master (admins manage it). A new name in the file is added to the master automatically.",
        "Product / Service: Product or Service.",
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
    style_header(ws, [28, 22, 28, 16, 16, 18, 18, 24, 16, 16, 14, 14, 20, 20, 20, 40])
    dv = DataValidation(type="list", formula1='"' + ",".join(STAGES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("K2:K2000")
    for r in range(2, 2001):
        ws.cell(row=r, column=13).number_format = "DD-MM-YYYY"
    ws2 = wb.create_sheet("Clients")
    ws2.append(CLIENT_HEADERS)
    style_header(ws2, [28, 22, 28, 16, 16, 18, 18, 24, 16, 16, 16, 16, 18, 18, 40])
    for r in range(2, 2001):
        ws2.cell(row=r, column=11).number_format = "DD-MM-YYYY"
    plist = ",".join(products)
    for sheet in (ws, ws2):
        dvt = DataValidation(type="list", formula1='"Product,Service"', allow_blank=True)
        sheet.add_data_validation(dvt)
        dvt.add("I2:I2000")
        if products and len(plist) < 250:     # Excel limits inline lists to 255 characters
            dvp = DataValidation(type="list", formula1='"' + plist.replace('"', "") + '"', allow_blank=True,
                                 showErrorMessage=False)    # warn-free: new products may be typed
            sheet.add_data_validation(dvp)
            dvp.add("H2:H2000")
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
            r["region"], r["product"], r["offering_type"], r["source"], r["stage"], r["probability"],
            r["expected_close"], r["value"], r["value_usd"], r["notes"])])
    style_header(ws, [28, 22, 28, 16, 16, 18, 18, 24, 16, 16, 14, 14, 20, 20, 20, 40])
    ws2 = wb.create_sheet("Clients")
    ws2.append(CLIENT_HEADERS)
    for r in conn.execute("SELECT * FROM records WHERE type='client' ORDER BY lower(company)"):
        ws2.append([xl_safe(x) for x in (
            r["company"], r["contact_name"], r["email"], r["phone"], r["owner"], r["presales_lead"],
            r["region"], r["product"], r["offering_type"], r["source"], r["closed_date"], r["closed_quarter"],
            r["value"], r["value_usd"], r["notes"])])
    style_header(ws2, [28, 22, 28, 16, 16, 18, 18, 24, 16, 16, 16, 16, 18, 18, 40])
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
    """Funnel + overdue / upcoming lists (client and won-quarter totals come from analytics_data)."""
    stages, lost = funnel(conn)
    today = datetime.now(TZ).date()
    open_n = sum(s["count"] for s in stages)
    # one query for both lists: everything not yet closed with a date up to 30 days ahead
    rows = conn.execute(
        "SELECT * FROM records WHERE type='prospect' AND stage!='Lost' AND expected_close <= ? "
        "ORDER BY expected_close", ((today + timedelta(days=30)).isoformat(),)).fetchall()
    overdue = [r for r in rows if r["expected_close"] < today.isoformat()]
    upcoming = [r for r in rows if r["expected_close"] >= today.isoformat()]
    return {
        "stages": stages, "lost": lost, "open_n": open_n,
        "pipeline": sum(s["value"] for s in stages), "weighted": sum(s["weighted"] for s in stages),
        "pipeline_usd": sum(s["value_usd"] for s in stages), "weighted_usd": sum(s["weighted_usd"] for s in stages),
        "overdue": overdue, "upcoming": upcoming, "max_count": max([s["count"] for s in stages] + [1]),
    }


BREAKDOWN_FIELDS = ("region", "presales_lead", "product", "offering_type", "owner", "source")


def breakdown_rows(conn):
    """Load once, then pass to breakdown() as many times as needed (one query instead of one per grouping)."""
    return conn.execute("SELECT type, stage, value, value_usd, probability, " + ", ".join(BREAKDOWN_FIELDS)
                        + " FROM records").fetchall()


def breakdown(conn, field, rows=None):
    """Pipeline + client summary grouped by a free-text field, ignoring upper/lower-case differences."""
    assert field in BREAKDOWN_FIELDS
    groups = {}
    for r in (rows if rows is not None else breakdown_rows(conn)):
        label = (r[field] or "").strip()
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


def _num(v):
    return float(v or 0)


def group_series(items, field, top=8):
    """Open-pipeline totals grouped by a free-text field (case-insensitive) as chart arrays."""
    groups = {}
    for r in items:
        label = (r[field] or "").strip()
        gp = groups.setdefault(label.lower(), {"names": {}, "n": 0, "inr": 0.0, "usd": 0.0, "w_inr": 0.0, "w_usd": 0.0})
        if label:
            gp["names"][label] = gp["names"].get(label, 0) + 1
        p = (r["probability"] or 0) / 100
        gp["n"] += 1
        gp["inr"] += _num(r["value"])
        gp["usd"] += _num(r["value_usd"])
        gp["w_inr"] += _num(r["value"]) * p
        gp["w_usd"] += _num(r["value_usd"]) * p
    rows = [dict(gp, label=(max(gp["names"], key=gp["names"].get) if gp["names"] else "(not set)"), unset=(k == ""))
            for k, gp in groups.items()]
    ti, tu = sum(x["inr"] for x in rows) or 1, sum(x["usd"] for x in rows) or 1
    rows.sort(key=lambda x: (x["unset"], -(x["inr"] / ti + x["usd"] / tu), x["label"].lower()))
    if len(rows) > top:
        rest = rows[top - 1:]
        other = {"label": "Other", "n": 0, "inr": 0.0, "usd": 0.0, "w_inr": 0.0, "w_usd": 0.0, "unset": False}
        for x in rest:
            for k in ("n", "inr", "usd", "w_inr", "w_usd"):
                other[k] += x[k]
        rows = rows[:top - 1] + [other]
    return {"labels": [x["label"] for x in rows], "n": [x["n"] for x in rows],
            "inr": [round(x["inr"]) for x in rows], "usd": [round(x["usd"]) for x in rows],
            "w_inr": [round(x["w_inr"]) for x in rows], "w_usd": [round(x["w_usd"]) for x in rows]}


DQ_FIELDS = [("expected_close", "Likely closure date", True, False), ("owner", "Sales owner", True, True),
             ("presales_lead", "Presales lead", True, True), ("region", "Region / Country", True, True),
             ("product", "Product", True, True), ("offering_type", "Product / Service", True, True),
             ("value", "Amount (INR)", True, True), ("value_usd", "Amount (USD)", True, True),
             ("source", "Lead source", True, True), ("closed_date", "Closed date", False, True)]
MISSING_FIELDS = tuple(f[0] for f in DQ_FIELDS)


def analytics_data(conn, cfg):
    fy = int(cfg["fy_start_month"])
    today = datetime.now(TZ).date()
    now = utcnow()
    rows = conn.execute(
        "SELECT id, type, company, stage, probability, expected_close, value, value_usd, closed_date, closed_quarter, "
        "owner, source, region, presales_lead, product, offering_type, created_at, updated_at FROM records").fetchall()
    prospects = [r for r in rows if r["type"] == "prospect"]
    opens = [r for r in prospects if r["stage"] != "Lost"]
    lost = [r for r in prospects if r["stage"] == "Lost"]
    clients = [r for r in rows if r["type"] == "client"]

    def p(r): return (r["probability"] or 0) / 100
    def inr(r): return _num(r["value"])
    def usd(r): return _num(r["value_usd"])

    def age_days(r):
        try:
            return max((now - datetime.fromisoformat(r["updated_at"])).days, 0)
        except (TypeError, ValueError):
            return 0

    sums = lambda items: {"inr": sum(inr(r) for r in items), "usd": sum(usd(r) for r in items),
                          "w_inr": sum(inr(r) * p(r) for r in items), "w_usd": sum(usd(r) * p(r) for r in items), "n": len(items)}
    pipe = sums(opens)
    today_s = today.isoformat()
    overdue = [r for r in opens if r["expected_close"] and r["expected_close"] < today_s]
    stale = sorted([r for r in opens if age_days(r) >= 30], key=age_days, reverse=True)
    nodate = [r for r in opens if not r["expected_close"]]
    soon = [r for r in opens if r["expected_close"] and today_s <= r["expected_close"] <= (today + timedelta(days=30)).isoformat()]
    cur_q = quarter_label(today, fy)
    won_q = [c for c in clients if c["closed_quarter"] == cur_q]
    decided = len(clients) + len(lost)
    win_rate = round(100 * len(clients) / decided) if decided else None
    n_inr, n_usd = sum(1 for r in opens if r["value"]), sum(1 for r in opens if r["value_usd"])

    # ---- chart data
    stages = {"labels": OPEN_STAGES, "count": [], "inr": [], "usd": [], "w_inr": [], "w_usd": []}
    for st in OPEN_STAGES:
        s = sums([r for r in opens if r["stage"] == st])
        stages["count"].append(s["n"])
        for k in ("inr", "usd", "w_inr", "w_usd"):
            stages[k].append(round(s[k]))

    months = []
    y, m = today.year, today.month
    for i in range(6):
        months.append(date(y + (m - 1 + i) // 12, (m - 1 + i) % 12 + 1, 1))
    labels = ["Overdue"] + [d.strftime("%b %Y") for d in months] + ["Later", "No date"]
    fc = {k: [0.0] * len(labels) for k in ("inr", "usd", "w_inr", "w_usd", "count")}
    for r in opens:
        if not r["expected_close"]:
            i = len(labels) - 1
        else:
            d = date.fromisoformat(r["expected_close"])
            diff = (d.year - today.year) * 12 + d.month - today.month
            i = 0 if d < today else (1 + diff if diff <= 5 else len(labels) - 2)
        fc["count"][i] += 1
        fc["inr"][i] += inr(r)
        fc["usd"][i] += usd(r)
        fc["w_inr"][i] += inr(r) * p(r)
        fc["w_usd"][i] += usd(r) * p(r)
    forecast = {"labels": labels, **{k: [round(v) for v in vs] for k, vs in fc.items()}}

    wq = {}
    for c in clients:
        if c["closed_quarter"]:
            g = wq.setdefault(c["closed_quarter"], {"count": 0, "inr": 0.0, "usd": 0.0})
            g["count"] += 1
            g["inr"] += inr(c)
            g["usd"] += usd(c)
    qs = sorted(wq)[-8:]
    won = {"labels": qs, "count": [wq[q]["count"] for q in qs], "inr": [round(wq[q]["inr"]) for q in qs],
           "usd": [round(wq[q]["usd"]) for q in qs]}

    bands = [("0–25%", 0, 25), ("26–50%", 26, 50), ("51–75%", 51, 75), ("76–100%", 76, 100)]
    prob = {"labels": [b[0] for b in bands], "count": [], "inr": [], "usd": []}
    for _, lo, hi in bands:
        items = [r for r in opens if lo <= (r["probability"] or 0) <= hi]
        prob["count"].append(len(items))
        prob["inr"].append(round(sum(inr(r) for r in items)))
        prob["usd"].append(round(sum(usd(r) for r in items)))

    def band_counts(vals, edges, names):
        out = [0] * len(names)
        for v in vals:
            out[next((i for i, e in enumerate(edges) if v < e), len(edges))] += 1
        return {"labels": names, "count": out}
    size = {"inr": band_counts([inr(r) for r in opens if r["value"]], [1e5, 5e5, 1e6, 5e6],
                               ["< ₹1 L", "₹1–5 L", "₹5–10 L", "₹10–50 L", "₹50 L +"]),
            "usd": band_counts([usd(r) for r in opens if r["value_usd"]], [1e3, 5e3, 1e4, 5e4],
                               ["< $1K", "$1–5K", "$5–10K", "$10–50K", "$50K +"])}

    ml = []
    for i in range(5, -1, -1):
        mm = today.month - i
        ml.append(date(today.year + (mm - 1) // 12, (mm - 1) % 12 + 1, 1))
    key = lambda r: (r["created_at"] or "")[:7]
    added = {"labels": [d.strftime("%b %y") for d in ml],
             "prospects": [sum(1 for r in prospects if key(r) == d.strftime("%Y-%m")) for d in ml],
             "clients": [sum(1 for r in clients if key(r) == d.strftime("%Y-%m")) for d in ml]}

    br = [x for x in breakdown(conn, "region", rows) if x["clients"] + x["lost"] > 0][:8]
    winrate = {"labels": [x["name"] for x in br], "rate": [x["win_rate"] or 0 for x in br],
               "won": [x["clients"] for x in br], "lost": [x["lost"] for x in br]}

    region_s, source_s = group_series(opens, "region"), group_series(opens, "source")
    by_product = breakdown(conn, "product", rows)
    pw = [x for x in by_product if x["clients"] + x["lost"] > 0 and not x["unset"]][:8]
    prodwin = {"labels": [x["name"] for x in pw], "rate": [x["win_rate"] or 0 for x in pw],
               "won": [x["clients"] for x in pw], "lost": [x["lost"] for x in pw]}
    product_s = group_series(opens, "product")
    charts = {"product": product_s, "offering": group_series(opens, "offering_type"),
              "prodwon": group_series(clients, "product"), "prodwin": prodwin,"stages": stages, "forecast": forecast, "won": won, "prob": prob, "size": size, "added": added,
              "winrate": winrate, "region": region_s, "source": source_s,
              "presales": group_series(opens, "presales_lead"), "owner": group_series(opens, "owner")}

    # ---- lists
    def row(r, extra=None):
        d = {k: r[k] for k in ("id", "company", "stage", "probability", "expected_close", "value", "value_usd",
                               "region", "presales_lead", "owner", "product", "offering_type")}
        d.update(extra or {})
        return d
    top_inr = [row(r) for r in sorted([r for r in opens if r["value"]], key=inr, reverse=True)[:8]]
    top_usd = [row(r) for r in sorted([r for r in opens if r["value_usd"]], key=usd, reverse=True)[:8]]
    stale_rows = [row(r, {"days": age_days(r)}) for r in stale[:8]]

    dq = []
    for field, label, for_p, for_c in DQ_FIELDS:
        dq.append({"key": field, "label": label,
                   "p": sum(1 for r in opens if r[field] is None) if for_p else None,
                   "c": sum(1 for r in clients if r[field] is None) if for_c else None})

    # ---- plain-language highlights
    hi = []
    pct = lambda a, b: round(100 * a / b) if b else 0
    if pipe["n"]:
        use = "inr" if pipe["inr"] else "usd"
        hi.append(("info", f"Weighted forecast is {pct(pipe['w_' + use], pipe[use])}% of the open pipeline "
                           f"({fmt_amt(pipe['w_inr'], pipe['w_usd'])} of {fmt_amt(pipe['inr'], pipe['usd'])})."))
        vals = sorted((inr(r) if pipe["inr"] else usd(r) for r in opens), reverse=True)
        if len(vals) >= 4 and sum(vals):
            share = pct(sum(vals[:3]), sum(vals))
            hi.append(("warn" if share >= 50 else "info",
                       f"The top 3 deals make up {share}% of the open pipeline"
                       + (" – a few slipping would hurt the forecast." if share >= 50 else ".")))
        pv = product_s["inr"] if pipe["inr"] else product_s["usd"]
        prod = [(l, v) for l, v in zip(product_s["labels"], pv) if l not in ("(not set)", "Other")]
        if prod and sum(pv):
            l, v = max(prod, key=lambda x: x[1])
            hi.append(("info", f"{l} leads the open pipeline by product – {pct(v, sum(pv))}% of the total."))
        rv = region_s["inr"] if pipe["inr"] else region_s["usd"]
        reg = [(l, v) for l, v in zip(region_s["labels"], rv) if l not in ("(not set)", "Other")]
        if reg and sum(rv):
            l, v = max(reg, key=lambda x: x[1])
            hi.append(("info", f"{l} is the largest region – {pct(v, sum(rv))}% of the open pipeline."))
    if soon:
        s = sums(soon)
        hi.append(("good", f"{len(soon)} prospect(s) are due to close in the next 30 days – weighted {fmt_amt(s['w_inr'], s['w_usd'])}."))
    if overdue:
        s = sums(overdue)
        hi.append(("bad", f"{len(overdue)} open prospect(s) worth {fmt_amt(s['inr'], s['usd'])} are past their expected closure date."))
    if stale:
        hi.append(("warn", f"{len(stale)} open prospect(s) haven't been updated in 30+ days."))
    if nodate:
        hi.append(("warn", f"{len(nodate)} open prospect(s) have no expected closure date, so they are missing from the forecast."))
    if win_rate is not None:
        hi.append(("good" if win_rate >= 50 else "info", f"Win rate is {win_rate}% ({len(clients)} won vs {len(lost)} lost)."))

    return {
        "pipe": pipe, "open_n": len(opens), "n_inr": n_inr, "n_usd": n_usd,
        "avg_inr": pipe["inr"] / n_inr if n_inr else 0, "avg_usd": pipe["usd"] / n_usd if n_usd else 0,
        "won_q": sums(won_q), "cur_q": cur_q, "clients": sums(clients), "lost_n": len(lost), "win_rate": win_rate,
        "overdue_n": len(overdue), "stale_n": len(stale), "nodate_n": len(nodate),
        "by_product": by_product, "by_offering": breakdown(conn, "offering_type", rows),
        "charts": charts, "top_inr": top_inr, "top_usd": top_usd, "stale": stale_rows, "dq": dq, "highlights": hi,
    }


def distinct_values(field):
    """Existing values, used to suggest consistent spellings in the entry form."""
    assert field in ("region", "presales_lead", "owner", "source", "product")
    return [r[0] for r in g.db.execute(
        f"SELECT DISTINCT {field} FROM records WHERE {field} IS NOT NULL AND {field}<>'' ORDER BY 1")]


def product_choices(current=None):
    """Active products for dropdowns, plus the record's current value if it was since deactivated."""
    names = active_products(g.db)
    if current and current.lower() not in {n.lower() for n in names}:
        names.append(current)
    return names


def report_data(conn, cfg):
    fy = int(cfg["fy_start_month"])
    rows = breakdown_rows(conn)
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
            "by_region": breakdown(conn, "region", rows), "by_presales": breakdown(conn, "presales_lead", rows),
            "by_product": breakdown(conn, "product", rows), "by_offering": breakdown(conn, "offering_type", rows)}


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


def indian_group(n):
    """12345678 -> '1,23,45,678' (lakh / crore grouping)."""
    s = str(abs(int(n)))
    if len(s) > 3:
        head, tail, parts = s[:-3], s[-3:], []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ",".join(parts + [tail])
    return ("-" if n < 0 else "") + s


def fmt_money(v, cur="₹"):
    n = int(round(v or 0))
    return f"₹{indian_group(n)}" if cur == "₹" else f"{cur}{n:,}"


def _trim(x):
    return f"{x:.2f}".rstrip("0").rstrip(".")


def fmt_compact(v, cur="₹"):
    """Short form for KPI tiles: ₹1.25 Cr, ₹45.3 L, $1.2M, $12.5K."""
    v = float(v or 0)
    if cur == "₹":
        if v >= 1e7:
            return f"₹{_trim(v / 1e7)} Cr"
        if v >= 1e5:
            return f"₹{_trim(v / 1e5)} L"
        return fmt_money(v, "₹")
    if v >= 1e9:
        return f"${_trim(v / 1e9)}B"
    if v >= 1e6:
        return f"${_trim(v / 1e6)}M"
    if v >= 1e4:
        return f"${_trim(v / 1e3)}K"
    return fmt_money(v, "$")


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
             "<th align='left'>Region</th><th align='left'>Product</th></tr>")
    for r in rows:
        flag = ""
        if r["expected_close"] and r["expected_close"] < today.isoformat():
            flag = " [OVERDUE]"
        elif not r["expected_close"]:
            flag = " [no date]"
        t.append(f"  - {r['company']} | {r['contact_name'] or '-'} | {r['stage']} | {r['probability']}% | "
                 f"{fmt_date(r['expected_close'])}{flag} | {fmt_amt(r['value'], r['value_usd'])} | {r['owner'] or '-'} | "
                 f"presales: {r['presales_lead'] or '-'} | {r['region'] or '-'} | "
                 f"{r['product'] or '-'}{' (' + r['offering_type'] + ')' if r['offering_type'] else ''}")
        color = "#b42318" if "OVERDUE" in flag else "#333"
        h.append(f"<tr style='border-bottom:1px solid #eee;color:{color}'><td>{e(r['company'])}</td>"
                 f"<td>{e(r['contact_name'] or '-')}</td><td>{e(r['stage'])}</td><td align='center'>{r['probability']}%</td>"
                 f"<td>{fmt_date(r['expected_close'])}{e(flag)}</td><td align='right'>{fmt_amt(r['value'], r['value_usd'])}</td>"
                 f"<td>{e(r['owner'] or '-')}</td><td>{e(r['presales_lead'] or '-')}</td>"
                 f"<td>{e(r['region'] or '-')}</td>"
                 f"<td>{e(r['product'] or '-')}{e(' (' + r['offering_type'] + ')') if r['offering_type'] else ''}</td></tr>")
    if not rows:
        t.append("  No open prospects.")
        h.append("<tr><td colspan='10'>No open prospects.</td></tr>")
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
                             ASSET_V=(os.environ.get("VERCEL_GIT_COMMIT_SHA") or "3")[:8],
                             distinct=lambda f: distinct_values(f),
                             product_choices=lambda cur=None: product_choices(cur),
                             OFFERING_TYPES=OFFERING_TYPES, chat_ready=lambda: chat.config() is not None)


@app.template_filter("fdate")
def _fdate(s):
    return fmt_date(s)


@app.template_filter("money")
def _money(v):
    return fmt_money(v, "₹")


@app.template_filter("usd")
def _usd(v):
    return fmt_money(v, "$")


@app.template_filter("inr_c")
def _inr_c(v):
    return fmt_compact(v, "₹")


@app.template_filter("usd_c")
def _usd_c(v):
    return fmt_compact(v, "$")


@app.before_request
def before():
    if request.endpoint == "static":
        return
    if request.method == "POST":
        sent = request.form.get("_csrf", "")
        if not sent or not secrets.compare_digest(sent, session.get("_csrf", "")):
            abort(400)
    g.t0 = time.perf_counter(); g.nq = 0
    ensure_init()
    # Plain page views run in autocommit mode on PostgreSQL: no BEGIN / ROLLBACK round trips to a remote database.
    g.db = connect(reuse=True, autocommit=request.method in ("GET", "HEAD"))
    g.cfg = cached_settings(g.db)
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
        t0 = getattr(g, "t0", None)
        if t0 is not None:        # visible in the browser's Network tab -> Timing; shows server time and DB round trips
            resp.headers["Server-Timing"] = f'app;dur={(time.perf_counter() - t0) * 1000:.0f}, db;desc="{getattr(g, "nq", 0)} queries"' 
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
    return render_template("dashboard.html", d=dashboard_data(g.db, g.cfg), a=analytics_data(g.db, g.cfg))


@app.route("/analytics")
@login_required
def analytics():
    a, tv = analytics_data(g.db, g.cfg), target_view()
    a["charts"]["targets"] = tv["chart"]
    return render_template("analytics.html", a=a, tv=tv)


@app.route("/reports")
@login_required
def reports():
    return render_template("reports.html", d=dashboard_data(g.db, g.cfg), r=report_data(g.db, g.cfg), tv=target_view())


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
            "company", "contact_name", "email", "owner", "presales_lead", "region", "product")) + ")")
        args += [f"%{q}%"] * 7
    if kind == "prospect" and stage in STAGES:
        where.append("stage=?")
        args.append(stage)
    elif kind == "prospect" and stage == "open":
        where.append("stage!='Lost'")
    missing = request.args.get("missing", "")
    if missing in MISSING_FIELDS:
        where.append(f"{missing} IS NULL")
    else:
        missing = ""
    w = " AND ".join(where)
    total = g.db.execute(f"SELECT COUNT(*) FROM records WHERE {w}", args).fetchone()[0]
    order = ("CASE WHEN expected_close IS NULL THEN 1 ELSE 0 END, expected_close, lower(company)"
             if kind == "prospect" else "closed_date DESC NULLS LAST, lower(company)")
    pages = max((total + PAGE_SIZE - 1) // PAGE_SIZE, 1)
    page = min(page, pages)           # e.g. the last row of the last page was just deleted
    rows = g.db.execute(f"SELECT * FROM records WHERE {w} ORDER BY {order} LIMIT ? OFFSET ?",
                        args + [PAGE_SIZE, (page - 1) * PAGE_SIZE]).fetchall()
    return render_template("records.html", kind=kind, rows=rows, q=q, stage=stage, page=page,
                           pages=pages, total=total, today=datetime.now(TZ).date().isoformat(), missing=missing,
                           missing_label=dict((f[0], f[1]) for f in DQ_FIELDS).get(missing, ""))


def save_from_form(rid=None):
    kind = "client" if request.form.get("type") == "client" else "prospect"
    raw = {k: request.form.get(k) for k in
           ("company", "contact_name", "email", "phone", "owner", "presales_lead", "region", "product",
            "offering_type", "source",
            "stage", "probability",
            "expected_close", "value", "value_usd", "closed_date", "closed_quarter", "notes")}
    d, errors = normalize(raw, kind, int(g.cfg["fy_start_month"]))
    if d.get("product"):
        pm = product_map(g.db)
        canon = pm.get(d["product"].lower())
        keep = None
        if rid:
            old = g.db.execute("SELECT product FROM records WHERE id=?", (rid,)).fetchone()
            keep = (old["product"] or "").lower() if old else None
        active = {n.lower() for n in active_products(g.db)}
        if canon is None or (canon.lower() not in active and canon.lower() != keep):
            errors.append("Product: pick one from the list (an admin can add products under Product master)")
        else:
            d["product"] = canon
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


def list_next():
    """Where to go back to after editing: the list page (with its filters and page number) the user came from.
    Only paths inside /records are accepted, so this cannot be used as an open redirect."""
    n = request.values.get("next", "")
    if n.startswith("/records") and not n.startswith("//") and "\\" not in n and not any(ch in n for ch in "\r\n"):
        return n
    return ""


@app.route("/records/new", methods=["GET", "POST"])
@login_required
def record_new():
    kind = "client" if request.values.get("type") == "client" else "prospect"
    if request.method == "POST":
        d, res = save_from_form()
        if d:
            flash(f"{d['company']} saved.", "ok")
            return redirect(url_for("records", type=d["type"]))
        return render_template("record_form.html", rec=res, kind=res["type"], is_new=True, next_url=list_next())
    return render_template("record_form.html", rec={"type": kind, "stage": "Lead", "probability": 10},
                           kind=kind, is_new=True, next_url=list_next())


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
            # stay on the same list page; if a prospect was converted it moves to the Clients list instead
            return redirect((list_next() if d["type"] == rec["type"] else "") or url_for("records", type=d["type"]))
        return render_template("record_form.html", rec=res | {"id": rid}, kind=res["type"], is_new=False,
                               next_url=list_next())
    data = form_values(rec)
    kind = rec["type"]
    if request.args.get("convert") == "1" and kind == "prospect":
        kind = data["type"] = "client"
        data["closed_date"] = datetime.now(TZ).date().isoformat()
    return render_template("record_form.html", rec=data, kind=kind, is_new=False, next_url=list_next(),
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
    return redirect(list_next() or url_for("records", type=rec["type"]))


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
    return send_file(build_template(active_products(g.db)), as_attachment=True, download_name="crm_import_template.xlsx",
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
        company = (request.form.get("company_name") or "QUEST FT").strip()[:60]
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
# ---- assistant (chat)
def chat_ctx():
    cfg = g.cfg
    return {"conn": g.db, "today": datetime.now(TZ).date(),
            "analytics": lambda: analytics_data(g.db, cfg), "funnel": lambda: funnel(g.db),
            "breakdown": lambda f: breakdown(g.db, f), "products": lambda: active_products(g.db),
            "targets": lambda fy, q: target_data(g.db, cfg, int(fy) if fy else current_fy(cfg), int(q or 0))}


@app.route("/assistant")
@login_required
def assistant():
    return redirect(url_for("dashboard"))      # IRIS now lives in the floating button on every page


@app.route("/api/chat", methods=["POST"])
@login_required
def api_chat():
    if chat.config() is None:
        return {"error": "The assistant is not set up yet. Ask the admin to add the LLM_API_KEY setting."}, 503
    try:
        history = chat.clean_history(json.loads(request.form.get("messages", "[]")))
    except ValueError:
        history = []
    if not history or history[-1]["role"] != "user":
        return {"error": "Type a question first."}, 400
    per_user, per_day = chat.limits()
    now = utcnow()
    n_user = g.db.execute("SELECT COUNT(*) FROM audit_log WHERE action='chat' AND username=? AND ts>?",
                          (g.user["username"], iso(now - timedelta(hours=1)))).fetchone()[0]
    n_day = g.db.execute("SELECT COUNT(*) FROM audit_log WHERE action='chat' AND ts>?",
                         (iso(now - timedelta(days=1)),)).fetchone()[0]
    if n_user >= per_user:
        return {"error": f"You've reached the limit of {per_user} questions per hour. Please try again later."}, 429
    if n_day >= per_day:
        return {"error": "The assistant's daily limit has been reached (it runs on a free plan). Try again tomorrow."}, 429
    audit(g.db, g.user["username"], "chat", "")       # counts usage only; questions are not stored
    g.db.commit()
    try:
        return {"reply": chat.answer(history, chat_ctx())}
    except chat.ChatError as e:
        return {"error": str(e)}, 502


# ---- sales targets (target vs achievement per sales rep)
def fy_of(d, fy_start):
    """(financial-year start year, quarter 1-4) that a date falls in."""
    return (d.year if d.month >= fy_start else d.year - 1), ((d.month - fy_start) % 12) // 3 + 1


def fy_name(y, fy_start):
    return str(y) if fy_start == 1 else f"FY{y}-{str(y + 1)[-2:]}"


def quarter_labels(y, fy_start):
    out = []
    for q in range(4):
        m = fy_start + 3 * q
        out.append(quarter_label(date(y + (m - 1) // 12, (m - 1) % 12 + 1, 1), fy_start))
    return out


def current_fy(cfg):
    return fy_of(datetime.now(TZ).date(), int(cfg["fy_start_month"]))[0]


def fy_options(conn, cfg):
    cur = current_fy(cfg)
    ys = {cur - 1, cur, cur + 1} | {r["fy"] for r in conn.execute("SELECT DISTINCT fy FROM targets")}
    return [{"y": y, "name": fy_name(y, int(cfg["fy_start_month"]))} for y in sorted(ys)]


def pct_of(a, t):
    return round(100 * a / t) if t else None


def target_data(conn, cfg, fy, quarter=0):
    """Target vs achieved per sales rep (owner) for one financial year and either one quarter or the whole year.
    Achieved = value of clients closed in the period; pipeline = probability-weighted open prospects expected
    to close in the period. INR and USD are never mixed."""
    fy_start = int(cfg["fy_start_month"])
    labels = quarter_labels(fy, fy_start)
    quarter = quarter if quarter in (1, 2, 3, 4) else 0
    reps = {}

    def slot(name):
        k = (name or "").strip().lower()
        if k not in reps:
            reps[k] = {"name": (name or "").strip() or "(unassigned)", "has_target": False,
                       "q": {i: dict(t_inr=0.0, t_usd=0.0, a_inr=0.0, a_usd=0.0, a_n=0, w_inr=0.0, w_usd=0.0) for i in range(1, 5)}}
        return reps[k]

    for r in conn.execute("SELECT rep, quarter, value_inr, value_usd FROM targets WHERE fy=?", (fy,)):
        x = slot(r["rep"]); x["name"] = r["rep"]; x["has_target"] = True
        x["q"][r["quarter"]]["t_inr"] = r["value_inr"] or 0.0
        x["q"][r["quarter"]]["t_usd"] = r["value_usd"] or 0.0
    for r in conn.execute("SELECT owner, value, value_usd, closed_date, closed_quarter FROM records WHERE type='client'"):
        if r["closed_date"]:
            y, q = fy_of(date.fromisoformat(r["closed_date"]), fy_start)
            if y != fy:
                continue
        elif r["closed_quarter"] in labels:
            q = labels.index(r["closed_quarter"]) + 1
        else:
            continue
        b = slot(r["owner"])["q"][q]
        b["a_inr"] += r["value"] or 0; b["a_usd"] += r["value_usd"] or 0; b["a_n"] += 1
    for r in conn.execute("SELECT owner, value, value_usd, probability, expected_close FROM records "
                          "WHERE type='prospect' AND stage!='Lost' AND expected_close IS NOT NULL"):
        y, q = fy_of(date.fromisoformat(r["expected_close"]), fy_start)
        if y == fy:
            b = slot(r["owner"])["q"][q]
            b["w_inr"] += (r["value"] or 0) * (r["probability"] or 0) / 100
            b["w_usd"] += (r["value_usd"] or 0) * (r["probability"] or 0) / 100

    qs = [quarter] if quarter else [1, 2, 3, 4]
    rows = []
    for x in reps.values():
        agg = {k: sum(x["q"][i][k] for i in qs) for k in x["q"][1]}
        row = {"name": x["name"], "has_target": x["has_target"], **agg,
               "pct_inr": pct_of(agg["a_inr"], agg["t_inr"]), "pct_usd": pct_of(agg["a_usd"], agg["t_usd"]),
               "gap_inr": max(agg["t_inr"] - agg["a_inr"], 0), "gap_usd": max(agg["t_usd"] - agg["a_usd"], 0),
               "proj_inr": pct_of(agg["a_inr"] + agg["w_inr"], agg["t_inr"]),
               "proj_usd": pct_of(agg["a_usd"] + agg["w_usd"], agg["t_usd"]),
               "quarters": [{**x["q"][i], "pct_inr": pct_of(x["q"][i]["a_inr"], x["q"][i]["t_inr"]),
                             "pct_usd": pct_of(x["q"][i]["a_usd"], x["q"][i]["t_usd"])} for i in range(1, 5)]}
        if any(agg[k] for k in ("t_inr", "t_usd", "a_inr", "a_usd", "w_inr", "w_usd")) or x["has_target"]:
            rows.append(row)
    rows.sort(key=lambda r: (r["name"] == "(unassigned)", -(r["t_inr"] + r["t_usd"]), -(r["a_inr"] + r["a_usd"]), r["name"].lower()))
    tot = {k: sum(r[k] for r in rows) for k in ("t_inr", "t_usd", "a_inr", "a_usd", "a_n", "w_inr", "w_usd", "gap_inr", "gap_usd")}
    tot.update(pct_inr=pct_of(tot["a_inr"], tot["t_inr"]), pct_usd=pct_of(tot["a_usd"], tot["t_usd"]),
               proj_inr=pct_of(tot["a_inr"] + tot["w_inr"], tot["t_inr"]), proj_usd=pct_of(tot["a_usd"] + tot["w_usd"], tot["t_usd"]),
               quarters=[{k: sum(r["quarters"][i][k] for r in rows) for k in ("t_inr", "t_usd", "a_inr", "a_usd")} for i in range(4)])
    for qd in tot["quarters"]:
        qd["pct_inr"], qd["pct_usd"] = pct_of(qd["a_inr"], qd["t_inr"]), pct_of(qd["a_usd"], qd["t_usd"])
    shown = rows[:15]
    chart = {"labels": [r["name"] for r in shown]}
    for k in ("t_inr", "a_inr", "w_inr", "t_usd", "a_usd", "w_usd"):
        chart[k] = [round(r[k], 2) for r in shown]
    return {"fy": fy, "fy_name": fy_name(fy, fy_start), "quarter": quarter, "q_labels": labels,
            "period": f"Q{quarter}" if quarter else "Full year", "rows": rows, "total": tot,
            "any_target": any(r["t_inr"] or r["t_usd"] for r in rows), "chart": chart, "years": fy_options(conn, cfg)}


def target_view():
    """Reads ?tfy=&tq= from the URL (used by Analytics and Reports)."""
    fy = request.args.get("tfy", type=int) or current_fy(g.cfg)
    q = request.args.get("tq", 0, type=int)
    return target_data(g.db, g.cfg, fy, q)


def upsert_targets(conn, items, now):
    """items: iterable of (rep, fy, quarter, inr, usd); a pair of None values removes that target."""
    existing = {}
    for r in conn.execute("SELECT id, lower(rep) k, fy, quarter FROM targets"):
        existing[(r["k"], r["fy"], r["quarter"])] = r["id"]
    n = 0
    for rep_name, fy, q, inr, usd in items:
        key = (rep_name.lower(), fy, q)
        if inr is None and usd is None:
            if key in existing:
                conn.execute("DELETE FROM targets WHERE id=?", (existing.pop(key),))
                n += 1
        elif key in existing:
            conn.execute("UPDATE targets SET rep=?, value_inr=?, value_usd=?, updated_at=? WHERE id=?",
                         (rep_name, inr, usd, now, existing[key]))
            n += 1
        else:
            conn.execute("INSERT INTO targets(rep,fy,quarter,value_inr,value_usd,updated_at) VALUES(?,?,?,?,?,?)",
                         (rep_name, fy, q, inr, usd, now))
            n += 1
    return n


def target_grid(fy):
    """Rows for the target master: every rep that has a target this year or appears as an owner on a record."""
    rows = {}
    for r in g.db.execute("SELECT rep, quarter, value_inr, value_usd FROM targets WHERE fy=?", (fy,)):
        x = rows.setdefault(r["rep"].lower(), {"name": r["rep"], "q": {}})
        x["name"] = r["rep"]; x["q"][r["quarter"]] = (r["value_inr"], r["value_usd"])
    for r in g.db.execute("SELECT MIN(owner) o, lower(owner) k FROM records WHERE owner IS NOT NULL AND owner<>'' GROUP BY lower(owner)"):
        rows.setdefault(r["k"], {"name": r["o"], "q": {}})
    return sorted(rows.values(), key=lambda x: x["name"].lower())


@app.route("/admin/targets", methods=["GET", "POST"])
@admin_required
def targets_admin():
    fy_start = int(g.cfg["fy_start_month"])
    fy = request.values.get("fy", type=int) or current_fy(g.cfg)
    if request.method == "POST":
        if request.form.get("copy_prev"):
            prev = g.db.execute("SELECT rep, quarter, value_inr, value_usd FROM targets WHERE fy=?", (fy - 1,)).fetchall()
            have = {(r["rep"].lower(), r["quarter"]) for r in g.db.execute("SELECT rep, quarter FROM targets WHERE fy=?", (fy,))}
            n = upsert_targets(g.db, [(r["rep"], fy, r["quarter"], r["value_inr"], r["value_usd"]) for r in prev
                                      if (r["rep"].lower(), r["quarter"]) not in have], iso(utcnow()))
            audit(g.db, g.user["username"], "targets_copied", f"{fy - 1} -> {fy}: {n} cell(s)")
            g.db.commit()
            flash(f"Copied {n} target(s) from {fy_name(fy - 1, fy_start)} into empty cells.", "ok")
            return redirect(url_for("targets_admin", fy=fy))
        items, errors, seen = [], [], set()
        for i, name in enumerate(request.form.getlist("rep")):
            name = clean_text(name, 80) or ""
            vals = []
            for q in range(1, 5):
                try:
                    vals.append((q, parse_money(request.form.get(f"t{i}_{q}_inr")), parse_money(request.form.get(f"t{i}_{q}_usd"))))
                except ValueError as e:
                    errors.append(f"{name or 'New row'} Q{q}: {e}")
            if not name:
                if any(v[1] is not None or v[2] is not None for v in vals):
                    errors.append("A row has targets but no sales rep name.")
                continue
            if name.lower() in seen:
                errors.append(f"'{name}' appears twice.")
                continue
            seen.add(name.lower())
            items += [(name, fy, q, a, b) for q, a, b in vals]
        if errors:
            for e in errors[:6]:
                flash(e, "error")
        else:
            n = upsert_targets(g.db, items, iso(utcnow()))
            audit(g.db, g.user["username"], "targets_saved", f"{fy_name(fy, fy_start)}: {n} change(s)")
            g.db.commit()
            flash(f"Targets for {fy_name(fy, fy_start)} saved.", "ok")
            return redirect(url_for("targets_admin", fy=fy))
    return render_template("targets.html", fy=fy, fy_label=fy_name(fy, fy_start), rows=target_grid(fy), years=fy_options(g.db, g.cfg),
                           qlabels=quarter_labels(fy, fy_start), fy_start=fy_start, fy_name=lambda y: fy_name(y, fy_start))


TARGET_HEADERS = ["Sales Rep*", "Financial Year*", "Quarter (Q1-Q4 or Year)*", "Target (INR)", "Target (USD)"]


@app.route("/admin/targets/template.xlsx")
@admin_required
def targets_template():
    wb = Workbook(); ws = wb.active; ws.title = "Targets"
    ws.append(TARGET_HEADERS)
    fy = current_fy(g.cfg)
    ws.append(["Example Rep", fy, "Q1", 2500000, 30000])
    ws.append(["Example Rep", fy, "Year", 10000000, 120000])
    style_header(ws, [26, 18, 26, 16, 16])
    ws2 = wb.create_sheet("Read me")
    for line in ("One row per sales rep and quarter. Quarter = Q1, Q2, Q3 or Q4 (of the financial year), or 'Year' to spread a yearly target evenly over the four quarters.",
                 "Financial Year = the year it starts in, e.g. 2026 (for an April-March year, 2026 means FY2026-27). Calendar-year companies just use the year.",
                 "Re-uploading overwrites the same rep + year + quarter. Leave both target cells empty to remove a target. INR and USD are separate."):
        ws2.append([line])
    ws2.column_dimensions["A"].width = 140
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return send_file(bio, as_attachment=True, download_name="targets_template.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/admin/targets/import", methods=["POST"])
@admin_required
def targets_import():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash("Choose an Excel (.xlsx) file.", "error")
        return redirect(url_for("targets_admin"))
    errors, items = [], []
    try:
        wb = load_workbook(f.stream, read_only=True, data_only=True)
        ws = wb["Targets"] if "Targets" in wb.sheetnames else wb.worksheets[0]
        for n, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if n > 2001:
                errors.append("Only the first 2,000 rows are read."); break
            row = list(row) + [None] * 5
            if all(v in (None, "") for v in row[:5]):
                continue
            name, fyv, qv, inr_v, usd_v = row[:5]
            name = clean_text(str(name) if name is not None else "", 80)
            m = re.search(r"(\d{4})", str(fyv or ""))
            try:
                if not name: raise ValueError("sales rep is missing")
                if not m: raise ValueError("financial year is missing (e.g. 2026)")
                inr, usd = parse_money(inr_v), parse_money(usd_v)
            except ValueError as e:
                errors.append(f"Row {n}: {e}"); continue
            qs = str(qv or "").strip().lower().replace("q", "")
            fy = int(m.group(1))
            if qs in ("1", "2", "3", "4"):
                items.append((name, fy, int(qs), inr, usd))
            elif qs in ("", "year", "fy", "annual", "full year", "yearly"):
                for q in range(1, 5):
                    items.append((name, fy, q, None if inr is None else round(inr / 4, 2), None if usd is None else round(usd / 4, 2)))
            else:
                errors.append(f"Row {n}: quarter must be Q1-Q4 or Year (got '{qv}')")
    except Exception as exc:
        g.db.rollback(); log.warning("targets import failed: %s", exc)
        flash("Could not read that file - is it a valid .xlsx workbook?", "error")
        return redirect(url_for("targets_admin"))
    n = upsert_targets(g.db, items, iso(utcnow()))
    audit(g.db, g.user["username"], "targets_import", f"{f.filename[:60]}: {n} change(s), {len(errors)} problem(s)")
    g.db.commit()
    flash(f"Imported targets: {n} saved." + (f" {len(errors)} row(s) skipped: " + "; ".join(errors[:4]) if errors else ""), "error" if errors else "ok")
    fy = items[0][1] if items else None
    return redirect(url_for("targets_admin", fy=fy) if fy else url_for("targets_admin"))




# ---- global search
SEARCH_COLS = ("company", "contact_name", "email", "owner", "presales_lead", "region", "product", "phone")


def like_escape(q):
    return q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def search_records(db, q, limit):
    pat = f"%{like_escape(q)}%"
    cond = " OR ".join(f"{c} {LIKE_OP} ? ESCAPE '\\'" for c in SEARCH_COLS)
    return db.execute(f"SELECT id, type, company, contact_name, owner, stage, product, closed_date FROM records WHERE {cond} "
                      f"ORDER BY CASE WHEN lower(company) LIKE ? THEN 0 ELSE 1 END, lower(company) LIMIT ?",
                      [pat] * len(SEARCH_COLS) + [like_escape(q.lower()) + "%", limit]).fetchall()


@app.route("/api/search")
@login_required
def api_search():
    q = request.args.get("q", "").strip()[:60]
    if len(q) < 2:
        return jsonify(items=[])
    pat = f"%{like_escape(q)}%"
    items = []
    for r in search_records(g.db, q, 7):
        sub = ("Client" if r["type"] == "client" else f"Prospect \u00b7 {r['stage']}") + (f" \u00b7 {r['owner']}" if r["owner"] else "")
        if r["contact_name"]:
            sub += f" \u00b7 {r['contact_name']}"
        items.append({"kind": "record", "label": r["company"], "sub": sub, "url": url_for("record_edit", rid=r["id"])})
    for r in g.db.execute(f"SELECT owner, COUNT(*) AS n FROM records WHERE owner {LIKE_OP} ? ESCAPE '\\' GROUP BY owner "
                          "ORDER BY n DESC LIMIT 3", (pat,)).fetchall():
        items.append({"kind": "rep", "label": r["owner"], "sub": f"Sales rep \u00b7 {r['n']} record(s)", "url": url_for("search_page", q=r["owner"])})
    for r in g.db.execute(f"SELECT name FROM products WHERE name {LIKE_OP} ? ESCAPE '\\' ORDER BY lower(name) LIMIT 2", (pat,)).fetchall():
        items.append({"kind": "product", "label": r["name"], "sub": "Product", "url": url_for("search_page", q=r["name"])})
    return jsonify(items=items)


@app.route("/search")
@login_required
def search_page():
    q = request.args.get("q", "").strip()[:60]
    rows = search_records(g.db, q, 200) if q else []
    return render_template("search.html", q=q, rows=rows)

# ---- admin: speed check
def db_location(url):
    """(host, cloud region guess, pooled?) from a Postgres URL, without exposing credentials."""
    from urllib.parse import urlparse
    host = urlparse(url).hostname or ""
    m = re.search(r"\b((?:us|eu|ap|sa|ca|me|af)-[a-z]+-\d)\b", host)
    return host, (m.group(1) if m else None), ("-pooler" in host or "pooler" in host)


# Vercel region -> the AWS region it sits in (used to compare with the database region)
VERCEL_TO_AWS = {"iad1": "us-east-1", "cle1": "us-east-2", "sfo1": "us-west-1", "pdx1": "us-west-2", "bom1": "ap-south-1",
                 "sin1": "ap-southeast-1", "syd1": "ap-southeast-2", "hnd1": "ap-northeast-1", "icn1": "ap-northeast-2",
                 "fra1": "eu-central-1", "dub1": "eu-west-1", "lhr1": "eu-west-2", "cdg1": "eu-west-3", "arn1": "eu-north-1",
                 "gru1": "sa-east-1", "yul1": "ca-central-1"}


@app.route("/admin/speed")
@admin_required
def speed_check():
    info = {"backend": "PostgreSQL" if USE_PG else "SQLite", "on_vercel": ON_VERCEL,
            "vercel_region": os.environ.get("VERCEL_REGION") or "", "verdicts": []}
    pings, connect_ms = [], None
    if USE_PG:
        host, region, pooled = db_location(DATABASE_URL)
        info.update(host=host, db_region=region, pooled=pooled)
        t = time.perf_counter(); tmp = _new_pg(); connect_ms = (time.perf_counter() - t) * 1000
        try:
            tmp.autocommit = True
            for _ in range(6):
                t = time.perf_counter(); tmp.execute("SELECT 1").fetchone(); pings.append((time.perf_counter() - t) * 1000)
        finally:
            tmp.close()
        pings = pings[1:]
    else:
        for _ in range(5):
            t = time.perf_counter(); g.db.execute("SELECT 1").fetchone(); pings.append((time.perf_counter() - t) * 1000)
    rtt = sum(pings) / len(pings)
    info.update(rtt=rtt, connect_ms=connect_ms, queries_per_page=5)
    est = rtt * 5
    info["est_page_ms"] = est
    v = info["verdicts"]
    if not USE_PG:
        v.append(("ok", "SQLite on local disk: no network latency."))
    elif rtt < 15:
        v.append(("ok", f"Database round trip is {rtt:.0f} ms - the database is close to the app. Slowness is not from the database distance."))
    else:
        v.append(("bad" if rtt > 60 else "warn", f"Each database call takes ~{rtt:.0f} ms. A page makes 3-12 calls, so expect roughly {est/1000:.1f}-{rtt*12/1000:.1f} s per page from distance alone."))
    if USE_PG and connect_ms and connect_ms > 400:
        v.append(("bad" if connect_ms > 1500 else "warn", f"Opening a new connection took {connect_ms:.0f} ms (paid on each cold start). "
                  + ("A value above ~1.5 s usually means the free database was asleep and woke up." if connect_ms > 1500 else "Distance and TLS setup add up here.")))
    if USE_PG and ON_VERCEL:
        aws = VERCEL_TO_AWS.get(info["vercel_region"])
        if info.get("db_region") and aws and aws != info["db_region"]:
            v.append(("bad", f"App runs in {info['vercel_region']} ({aws}) but the database is in {info['db_region']}. "
                             "Pin the app next to the database: set \"regions\" in vercel.json (see the README, Speed on Vercel) and redeploy."))
        elif not info.get("db_region"):
            v.append(("warn", "Could not tell the database region from its address. Compare it with the app region shown here."))
        if not info.get("pooled"):
            v.append(("warn", "The database address is not the pooled one (no '-pooler' in the host). Use Neon's pooled connection string for serverless."))
    return render_template("speed.html", i=info, pings=pings)

# ---- admin: product master
def product_rows():
    use = {}
    for r in g.db.execute("SELECT lower(product) p, type, COUNT(*) n FROM records "
                          "WHERE product IS NOT NULL AND product<>'' GROUP BY 1, 2"):
        use.setdefault(r["p"], {"prospect": 0, "client": 0})[r["type"]] = r["n"]
    rows = []
    for r in g.db.execute("SELECT * FROM products ORDER BY active DESC, lower(name)"):
        u = use.get(r["name"].lower(), {"prospect": 0, "client": 0})
        rows.append({"id": r["id"], "name": r["name"], "active": r["active"],
                     "prospects": u["prospect"], "clients": u["client"]})
    return rows


@app.route("/admin/products", methods=["GET", "POST"])
@admin_required
def products():
    if request.method == "POST":
        name = clean_text(request.form.get("name"), 120)
        if not name:
            flash("Enter a product name.", "error")
        elif name.lower() in product_map(g.db):
            flash(f"'{name}' is already in the list.", "error")
        else:
            add_product(g.db, name)
            audit(g.db, g.user["username"], "product_added", name)
            g.db.commit()
            flash(f"'{name}' added – it now appears in the product dropdown.", "ok")
            return redirect(url_for("products"))
    return render_template("products.html", rows=product_rows())


def _product_or_404(pid):
    row = g.db.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
    if not row:
        abort(404)
    return row


@app.route("/admin/products/<int:pid>/rename", methods=["POST"])
@admin_required
def product_rename(pid):
    row = _product_or_404(pid)
    new = clean_text(request.form.get("name"), 120)
    clash = product_map(g.db).get((new or "").lower())
    if not new:
        flash("Enter a product name.", "error")
    elif clash and clash != row["name"] and new.lower() != row["name"].lower():
        flash(f"'{new}' already exists.", "error")
    elif new != row["name"]:
        g.db.execute("UPDATE products SET name=? WHERE id=?", (new, pid))
        n = g.db.execute("UPDATE records SET product=? WHERE lower(product)=lower(?)",
                         (new, row["name"])).rowcount
        audit(g.db, g.user["username"], "product_renamed", f"{row['name']} -> {new} ({n} record(s))")
        g.db.commit()
        flash(f"Renamed to '{new}' ({n} record(s) updated).", "ok")
    return redirect(url_for("products"))


@app.route("/admin/products/<int:pid>/toggle", methods=["POST"])
@admin_required
def product_toggle(pid):
    row = _product_or_404(pid)
    g.db.execute("UPDATE products SET active=? WHERE id=?", (0 if row["active"] else 1, pid))
    audit(g.db, g.user["username"], "product_disabled" if row["active"] else "product_enabled", row["name"])
    g.db.commit()
    flash(f"'{row['name']}' {'hidden from' if row['active'] else 'back in'} the dropdown.", "ok")
    return redirect(url_for("products"))


@app.route("/admin/products/<int:pid>/delete", methods=["POST"])
@admin_required
def product_delete(pid):
    row = _product_or_404(pid)
    used = g.db.execute("SELECT COUNT(*) FROM records WHERE lower(product)=lower(?)", (row["name"],)).fetchone()[0]
    if used:
        flash(f"'{row['name']}' is used by {used} record(s) – hide it instead of deleting.", "error")
    else:
        g.db.execute("DELETE FROM products WHERE id=?", (pid,))
        audit(g.db, g.user["username"], "product_deleted", row["name"])
        g.db.commit()
        flash(f"'{row['name']}' deleted.", "ok")
    return redirect(url_for("products"))


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
            flash("Nothing deleted - type the word DELETE (capital letters) in the confirmation box.", "error")
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
