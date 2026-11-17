import io, os, re, sys, tempfile
"""Smoke test. SQLite by default; set TEST_DATABASE_URL=postgresql://... to run against Postgres
(WARNING: that database's public schema is wiped first)."""
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
tmp = tempfile.mkdtemp()
os.environ.update(CRM_DATA_DIR=tmp, CRM_DISABLE_SCHEDULER="1", CRM_ADMIN_PASSWORD="InitialPass123",
                  CRON_SECRET="cronsecret123", CRM_SECRET_KEY="test-secret-key")
os.environ.pop("POSTGRES_URL", None)
if os.environ.get("TEST_DATABASE_URL"):
    import psycopg
    os.environ["DATABASE_URL"] = os.environ["TEST_DATABASE_URL"]
    with psycopg.connect(os.environ["DATABASE_URL"], autocommit=True) as _c:
        _c.execute("DROP SCHEMA public CASCADE"); _c.execute("CREATE SCHEMA public")
else:
    os.environ.pop("DATABASE_URL", None)
import app as crm
from openpyxl import Workbook, load_workbook

c = crm.app.test_client()
fails = []
def check(name, cond, extra=""):
    print(("PASS " if cond else "FAIL ") + name, extra if not cond else "")
    if not cond: fails.append(name)

def token(client=None, path="/login"):
    r = (client or c).get(path)
    m = re.search(r'name="_csrf" value="([^"]+)"', r.get_data(as_text=True))
    return m.group(1) if m else None

# 1. unauthenticated access is blocked
for p in ["/", "/records", "/reports", "/data", "/export.xlsx", "/admin/users", "/admin/settings"]:
    r = c.get(p); check(f"anon redirected {p}", r.status_code == 302 and "/login" in r.headers["Location"])

# 2. login failures + CSRF
r = c.post("/login", data={"username": "admin", "password": "x"}); check("POST without csrf rejected", r.status_code == 400)
t = token()
r = c.post("/login", data={"_csrf": t, "username": "admin", "password": "wrong"}); check("bad password rejected", b"Invalid username" in r.data)
r = c.post("/login", data={"_csrf": t, "username": "admin", "password": "InitialPass123"})
check("login ok -> redirect", r.status_code == 302)
r = c.get("/", follow_redirects=False); check("must change password first", r.status_code == 302 and "password" in r.headers["Location"])

# 3. change password
t = token(path="/account/password")
r = c.post("/account/password", data={"_csrf": t, "current": "InitialPass123", "new": "short", "new2": "short"}); check("weak pw rejected", b"at least 10" in r.data)
r = c.post("/account/password", data={"_csrf": t, "current": "InitialPass123", "new": "NewSecret2026x", "new2": "NewSecret2026x"}, follow_redirects=True)
check("password changed, dashboard loads", r.status_code == 200 and b"Dashboard" in r.data, r.data[:300])
check("security headers", "frame-ancestors" in r.headers.get("Content-Security-Policy", "") and r.headers.get("X-Frame-Options") == "DENY")

# 4. data entry
t = token(path="/records/new")
r = c.post("/records/new", data={"_csrf": t, "type": "prospect", "company": "Acme Ltd", "email": "a@acme.com", "stage": "Proposal",
                                  "probability": "60", "expected_close": "2026-12-20", "value": "500000", "owner": "Ravi"}, follow_redirects=True)
check("prospect created", b"Acme Ltd" in r.data and b"60%" in r.data)
r = c.post("/records/new", data={"_csrf": t, "type": "prospect", "company": "Late Co", "stage": "Lead", "expected_close": "2026-01-10", "value": "100000"}, follow_redirects=True)
check("overdue prospect created", b"Late Co" in r.data)
r = c.post("/records/new", data={"_csrf": t, "type": "client", "company": "BigClient", "closed_date": "2026-08-15", "value": "900000"}, follow_redirects=True)
check("client created with quarter", b"2026-Q3" in r.data)
r = c.post("/records/new", data={"_csrf": t, "type": "client", "company": "NoDate"}); check("client without date/quarter rejected", b"required for a client" in r.data)
r = c.post("/records/new", data={"_csrf": t, "type": "prospect", "company": "X", "probability": "150"}); check("probability >100 rejected", b"between 0 and 100" in r.data)

# 5. convert
r = c.get("/records/1/edit?convert=1"); check("convert form preselects client", b'value="client" selected' in r.data)

# 6. Excel import
wb = Workbook(); ws = wb.active; ws.title = "Prospects"
ws.append(["Company*", "Contact Name", "Email", "Stage", "Probability (%)", "Expected Close Date", "Expected Value"])
ws.append(["Globex", "Jane", "j@globex.com", "Qualified", 0.4, "31-12-2026", "₹2,50,000"])
ws.append(["Initech", None, None, "Bogus", 10, None, None])
ws.append(["Acme Ltd", "Bob", "a@acme.com", "Negotiation", "80%", "15-01-2027", 600000])
ws2 = wb.create_sheet("Clients"); ws2.append(["Company", "Closed Date", "Closed Quarter", "Deal Value"])
ws2.append(["OldClient", "01-04-2026", None, 120000]); ws2.append(["QClient", None, "Q2 2026", 50000]); ws2.append(["Bad", None, None, 1])
bio = io.BytesIO(); wb.save(bio); bio.seek(0)
t = token(path="/data")
r = c.post("/data", data={"_csrf": t, "update_existing": "1", "dry_run": "1", "file": (io.BytesIO(bio.getvalue()), "x.xlsx")}, content_type="multipart/form-data")
check("dry run reports", b"Dry run" in r.data and b"unknown stage" in r.data)
with crm.app.app_context() if False else open(os.devnull) as _: pass
cnt = crm.connect().execute("SELECT COUNT(*) FROM records").fetchone()[0]
r = c.post("/data", data={"_csrf": t, "update_existing": "1", "file": (io.BytesIO(bio.getvalue()), "x.xlsx")}, content_type="multipart/form-data")
check("import done", b"Import complete" in r.data, r.data[-600:])
db = crm.connect()
check("dry run saved nothing", cnt == 3, cnt)
rows = {r["company"]: dict(r) for r in db.execute("SELECT * FROM records")}
check("Globex prob 40 and date", rows["Globex"]["probability"] == 40 and rows["Globex"]["expected_close"] == "2026-12-31" and rows["Globex"]["value"] == 250000.0, rows.get("Globex"))
check("Acme updated not duplicated", db.execute("SELECT COUNT(*) FROM records WHERE company='Acme Ltd'").fetchone()[0] == 1 and rows["Acme Ltd"]["stage"] == "Negotiation" and rows["Acme Ltd"]["probability"] == 80)
check("Initech (bad stage) skipped", "Initech" not in rows)
check("OldClient quarter 2026-Q2", rows["OldClient"]["closed_quarter"] == "2026-Q2")
check("QClient text quarter normalised", rows["QClient"]["closed_quarter"] == "2026-Q2")
check("Bad client skipped", "Bad" not in rows)

# 7. reports/dashboard/template/export
for p in ["/", "/reports", "/records?type=client", "/records?type=prospect&q=acme&stage=open", "/data", "/audit" if False else "/admin/audit"]:
    r = c.get(p); check(f"page ok {p}", r.status_code == 200, r.status_code)
r = c.get("/template.xlsx"); check("template downloads", r.status_code == 200 and load_workbook(io.BytesIO(r.data)).sheetnames == ["Instructions", "Prospects", "Clients"])
r = c.get("/export.xlsx"); wbx = load_workbook(io.BytesIO(r.data)); check("export has data", wbx["Prospects"].max_row > 1 and wbx["Clients"].max_row > 1)

# 8. reminder content
cfg = crm.get_settings(db); subj, text, html = crm.build_reminder(db, cfg)
check("reminder lists prospects + overdue", "Globex" in text and "OVERDUE" in text and "Late Co" in html, text)

# 9. users/admin: regular user cannot access admin; admin manages users
t = token(path="/admin/users/new")
r = c.post("/admin/users/new", data={"_csrf": t, "username": "priya", "role": "user", "password": "weak"}); check("weak temp pw rejected", b"at least 10" in r.data)
r = c.post("/admin/users/new", data={"_csrf": t, "username": "priya", "role": "user", "password": "TempPass12345"}, follow_redirects=True); check("user created", b"priya" in r.data)
u = crm.connect().execute("SELECT * FROM users WHERE username='priya'").fetchone()
u2 = crm.app.test_client()
t2 = token(u2); r = u2.post("/login", data={"_csrf": t2, "username": "priya", "password": "TempPass12345"}); check("user logs in", r.status_code == 302)
r = u2.get("/records"); check("user forced to change pw", r.status_code == 302 and "password" in r.headers["Location"])
t2 = token(u2, "/account/password"); u2.post("/account/password", data={"_csrf": t2, "current": "TempPass12345", "new": "PriyaPass2026", "new2": "PriyaPass2026"})
check("user sees records", u2.get("/records").status_code == 200)
for p in ["/admin/users", "/admin/settings", "/admin/audit"]:
    check(f"user blocked from {p}", u2.get(p).status_code == 403)
t2 = token(u2, "/records?type=prospect")
r = u2.post("/records/1/delete", data={"_csrf": t2}); check("user cannot delete record", r.status_code == 403)

# 10. admin: disable user -> kicked out; reset password -> session invalidated; last admin protected
t = token(path="/admin/users")
r = c.post(f"/admin/users/{u['id']}", data={"_csrf": t, "role": "user", "password": "ResetPass2026"}, follow_redirects=True)
check("admin reset password", b"User updated" in r.data)
check("user session invalidated after reset", u2.get("/records").status_code == 302)
r = c.post(f"/admin/users/{u['id']}", data={"_csrf": t, "role": "user"}, follow_redirects=True)  # active unchecked -> disable
t2 = token(u2); r = u2.post("/login", data={"_csrf": t2, "username": "priya", "password": "ResetPass2026"}); check("disabled user cannot log in", r.status_code == 200 and b"Invalid" in r.data)
r = c.post("/admin/users/1", data={"_csrf": t, "role": "user", "active": "1"}, follow_redirects=True); check("cannot demote last admin", b"at least one active admin" in r.data)
r = c.post("/admin/users/1", data={"_csrf": t, "role": "admin"}, follow_redirects=True); check("cannot deactivate self/last admin", b"at least one active admin" in r.data or b"cannot deactivate" in r.data)

# 11. admin delete record, settings
r = c.post("/records/2/delete", data={"_csrf": t}, follow_redirects=True); check("admin deletes", b"deleted" in r.data)
r = c.post("/admin/settings", data={"_csrf": t, "action": "save", "reminder_emails": "boss@x.com; me@y.com", "reminder_day": "0", "reminder_hour": "9",
                                     "reminder_enabled": "1", "fy_start_month": "4", "currency": "₹", "company_name": "Test Co"}, follow_redirects=True)
check("settings saved", b"Settings saved" in r.data)
check("quarters relabelled to FY", crm.connect().execute("SELECT closed_quarter FROM records WHERE company='BigClient'").fetchone()[0] == "FY2026-27 Q2")
r = c.post("/admin/settings", data={"_csrf": t, "action": "test"}, follow_redirects=True); check("test mail w/o SMTP gives clear error", b"SMTP is not configured" in r.data)

# 12. lockout
t3 = token(crm.app.test_client()); lc = crm.app.test_client(); t3 = token(lc)
for i in range(5): lc.post("/login", data={"_csrf": t3, "username": "admin", "password": "nope"})
r = lc.post("/login", data={"_csrf": t3, "username": "admin", "password": "NewSecret2026x"}); check("account locks after 5 failures (even with right pw)", b"locked" in r.data)

# 13. open redirect + formula injection
t4 = token(crm.app.test_client())
check("safe_next blocks //evil", crm.safe_next("//evil.com") is None and crm.safe_next("https://evil.com") is None and crm.safe_next("/records?x=1") == "/records?x=1")
check("xl_safe", crm.xl_safe("=HYPERLINK(1)") == "'=HYPERLINK(1)")

# 14. scheduler logic
import datetime as dt
db = crm.connect(); crm.set_setting(db, "reminder_last_attempt", ""); db.commit()
crm.maybe_send_weekly(db)  # day/hour may not match; must not raise

# 15. cron endpoint (Vercel)
check("cron without secret -> 401", c.get("/api/cron/reminders").status_code == 401)
check("cron wrong secret -> 401", c.get("/api/cron/reminders", headers={"Authorization": "Bearer nope"}).status_code == 401)
r = c.get("/api/cron/reminders", headers={"Authorization": "Bearer cronsecret123"})
check("cron with secret -> 200", r.status_code == 200 and r.get_json()["ok"] is True, r.data)

# 16. case-insensitive search + usernames
check("search is case-insensitive", b"Globex" in c.get("/records?type=prospect&q=GLOBEX").data)
t = token(path="/admin/users/new")
c.post("/admin/users/new", data={"_csrf": t, "username": "Sam", "role": "user", "password": "TempPass12345"})
r = c.post("/admin/users/new", data={"_csrf": t, "username": "SAM", "role": "user", "password": "TempPass12345"})
check("duplicate username (different case) rejected", b"already exists" in r.data)
sc = crm.app.test_client(); ts = token(sc)
check("login is case-insensitive on username", sc.post("/login", data={"_csrf": ts, "username": "SAM", "password": "TempPass12345"}).status_code == 302)

# 17. duplicate rows inside one upload + big batch
wb = Workbook(); ws = wb.active; ws.title = "Prospects"
ws.append(["Company", "Email", "Stage"])
ws.append(["Dup Co", "d@d.com", "Lead"]); ws.append(["dup co", "D@D.com", "Lead"])
for i in range(300): ws.append([f"Bulk {i}", f"b{i}@x.com", "Qualified"])
bio = io.BytesIO(); wb.save(bio)
t = token(path="/data")
r = c.post("/data", data={"_csrf": t, "update_existing": "1", "file": (io.BytesIO(bio.getvalue()), "dup.xlsx")}, content_type="multipart/form-data")
check("in-file duplicate reported", b"Duplicate of an earlier row" in r.data)
n = crm.connect().execute("SELECT COUNT(*) FROM records WHERE company LIKE 'Bulk %' OR lower(company)='dup co'").fetchone()[0]
check("bulk import inserted 301 rows", n == 301, n)
r = c.post("/data", data={"_csrf": t, "update_existing": "1", "file": (io.BytesIO(bio.getvalue()), "dup.xlsx")}, content_type="multipart/form-data")
n2 = crm.connect().execute("SELECT COUNT(*) FROM records WHERE company LIKE 'Bulk %' OR lower(company)='dup co'").fetchone()[0]
check("re-import is idempotent", n2 == 301, n2)
for p in ["/", "/reports", "/records?type=client", "/records?type=prospect&page=2"]:
    check(f"page ok after bulk {p}", c.get(p).status_code == 200)


# 18. region / presales lead
t = token(path="/records/new")
c.post("/records/new", data={"_csrf": t, "type": "prospect", "company": "Gulf Trade", "stage": "Proposal", "value": "200000",
                             "region": "UAE", "presales_lead": "Anil"}, follow_redirects=True)
row = crm.connect().execute("SELECT region, presales_lead FROM records WHERE company='Gulf Trade'").fetchone()
check("form saves region + presales lead", row["region"] == "UAE" and row["presales_lead"] == "Anil", dict(row))
check("form shows region suggestions", b'value="UAE"' in c.get("/records/new").data)
wb = Workbook(); ws = wb.active; ws.title = "Prospects"
ws.append(["Company", "Country", "Pre-sales Lead", "Stage", "Expected Value"])
ws.append(["RegionCo", "uae", "Priya", "Lead", 1000]); ws.append(["RegionCo2", "Uae ", "Priya", "Qualified", 3000])
ws.append(["RegionCo3", "India", "anil", "Lead", 500]); ws.append(["RegionCo4", None, None, "Lead", 100])
bio = io.BytesIO(); wb.save(bio)
t = token(path="/data")
r = c.post("/data", data={"_csrf": t, "update_existing": "1", "file": (io.BytesIO(bio.getvalue()), "r.xlsx")}, content_type="multipart/form-data")
check("import accepts Country + Pre-sales Lead headers", b"Import complete" in r.data and b"No 'Company'" not in r.data)
row = crm.connect().execute("SELECT region, presales_lead FROM records WHERE company='RegionCo'").fetchone()
check("imported region/presales values", row["region"] == "uae" and row["presales_lead"] == "Priya", dict(row))
bd = {x["name"].lower(): x for x in crm.breakdown(crm.connect(), "region")}
check("region report groups UAE/uae/'Uae ' together", bd["uae"]["open_n"] == 3 and bd["uae"]["open_value"] == 204000, bd.get("uae"))
check("unset region shown last as (not set)", crm.breakdown(crm.connect(), "region")[-1]["name"] == "(not set)")
pb = {x["name"].lower(): x for x in crm.breakdown(crm.connect(), "presales_lead")}
check("presales report groups Anil/anil", pb["anil"]["open_n"] >= 2, pb.get("anil"))
rep = c.get("/reports").get_data(as_text=True)
check("reports page has region + presales sections", "By region / country" in rep and "By presales lead" in rep and "UAE" in rep)
check("search finds by region", b"Gulf Trade" in c.get("/records?type=prospect&q=uae").data)
check("lists show new columns", b"Presales lead" in c.get("/records?type=prospect").data and b"Region" in c.get("/records?type=client").data)
tw = load_workbook(io.BytesIO(c.get("/template.xlsx").data))
check("template has new columns", "Region / Country" in [x.value for x in tw["Prospects"][1]] and "Presales Lead" in [x.value for x in tw["Clients"][1]])
ex = load_workbook(io.BytesIO(c.get("/export.xlsx").data))
hdr = [x.value for x in ex["Prospects"][1]]
vals = {r[0].value: r for r in ex["Prospects"].iter_rows(min_row=2)}
check("export has region/presales in right columns", vals["Gulf Trade"][hdr.index("Region / Country")].value == "UAE" and vals["Gulf Trade"][hdr.index("Presales Lead")].value == "Anil")
subj, text, html = crm.build_reminder(crm.connect(), crm.get_settings(crm.connect()))
check("reminder e-mail includes region + presales", "presales: Anil" in text and "Anil" in html and "UAE" in html)

print("backend:", "PostgreSQL" if crm.USE_PG else "SQLite")
print("\nFAILED:" if fails else "\nALL PASSED", fails)
