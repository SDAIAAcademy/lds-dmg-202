# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Generate the Sinwan Professional Licensing Authority file pack used by every lab (labs/files/).

    uv run labs/_tools/prepare_files.py

Deterministic: a fixed seed and fixed file times, so the same bytes on every run and every
answer key matches. Everything is synthetic: the Authority, its people, its practitioners and
its numbers. Identity numbers start with 9 and mobile numbers with 0599, so none is a real one;
e-mail addresses use the reserved domain example.com. The answer keys go to
labs/solutions/reference/ for the checker (labs/solutions/check.py).
"""
from __future__ import annotations

import csv
import io
import json
import random
import re
import sqlite3
import sys
import uuid
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import content as C  # noqa: E402

LABS = Path(__file__).resolve().parents[1]
OUT = LABS / "files"
SRC = OUT / "sources"
REF = LABS / "solutions" / "reference"
SEED = 2022
FIXED = datetime(2026, 11, 1, 9, 0, 0)
AS_OF = date(2026, 10, 31)

HEAD_FILL = PatternFill("solid", fgColor="0F766E")
FILL_FILL = PatternFill("solid", fgColor="FFF7E0")

rng = random.Random(SEED)

FIRST_M = ["عبدالله", "محمد", "فهد", "سلطان", "خالد", "ماجد", "تركي", "عبدالعزيز", "نايف", "بدر", "يوسف", "إبراهيم", "سعد", "حمد", "مشعل", "طلال", "زياد", "عمر", "أنس", "وليد"]
FIRST_F = ["نورة", "سارة", "ريم", "هند", "لطيفة", "أمل", "منى", "جواهر", "دانة", "شهد", "لمى", "هيا", "عبير", "رهف", "أسماء", "مها", "العنود", "بشاير", "وفاء", "روان"]
FATHER = ["سعيد", "علي", "أحمد", "صالح", "ناصر", "حسن", "عبدالرحمن", "سليمان", "منصور", "راشد"]
FAMILY = ["القحطاني", "الدوسري", "الغامدي", "الزهراني", "العتيبي", "الشمري", "الحربي", "المطيري", "العمري", "السبيعي", "الشهري", "البقمي", "الجهني", "العنزي", "الرشيدي", "الخالدي", "السلمي", "الثقفي", "اليامي", "الأحمدي"]
REGIONS = [("R01", "الرياض"), ("R02", "مكة المكرمة"), ("R03", "المدينة المنورة"), ("R04", "القصيم"), ("R05", "المنطقة الشرقية"), ("R06", "عسير"), ("R07", "تبوك"), ("R08", "حائل"), ("R09", "الحدود الشمالية"), ("R10", "جازان"), ("R11", "نجران"), ("R12", "الباحة"), ("R13", "الجوف")]
REGION_W = [26, 24, 8, 5, 15, 6, 3, 3, 1, 4, 2, 1, 2]
PROFESSIONS = [("TRN", "الترجمة المعتمدة", "Certified translation", 20), ("SRV", "المساحة", "Land surveying", 30),
               ("INT", "التصميم الداخلي", "Interior design", 15), ("OSH", "استشارات السلامة المهنية", "Occupational safety consulting", 25)]
PROF_W = [30, 20, 35, 15]
FEES = {"new": 600.0, "renewal": 300.0, "upgrade": 400.0}


# ------------------------------------------------------------------ writing helpers
def write_text(rel: str, text: str) -> None:
    path = OUT / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes((text.strip() + "\n").encode("utf-8"))


def write_csv(path: Path, header, rows) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    buf = io.StringIO(newline="")
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(header)
    w.writerows(rows)
    path.write_bytes(buf.getvalue().encode("utf-8"))


def sheet(wb, title, header, rows, fill_cols=(), widths=None, validations=None):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in rows:
        ws.append(list(r))
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if c.column - 1 in fill_cols:
                c.fill = FILL_FILL
    for i, w in enumerate(widths or [18] * len(header)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    ws.freeze_panes = "A2"
    for col, values in (validations or {}).items():
        dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{max(2, len(rows) + 1)}")
    return ws


def save(wb, path: Path) -> None:
    del wb[wb.sheetnames[0]]
    wb.properties.creator = C.AGENCY
    wb.properties.lastModifiedBy = C.AGENCY
    wb.properties.created = wb.properties.modified = FIXED
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    wb.save(tmp)
    with zipfile.ZipFile(tmp) as z:
        items = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    stamp = FIXED.strftime("%Y-%m-%dT%H:%M:%SZ").encode()
    items = [(n, re.sub(rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms)", rb"\g<1>" + stamp + rb"\g<2>", d))
             if n == "docProps/core.xml" else (n, d) for n, d in items]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items:
            info = zipfile.ZipInfo(name, date_time=FIXED.timetuple()[:6])
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    tmp.unlink()


def put_ref(name: str, data) -> None:
    REF.mkdir(parents=True, exist_ok=True)
    (REF / f"{name}.json").write_bytes((json.dumps(data, ensure_ascii=False, indent=1) + "\n").encode("utf-8"))


def d(x: date) -> str:
    return x.isoformat()


def rand_date(a: date, b: date) -> date:
    return a + timedelta(days=rng.randrange((b - a).days + 1))


def new_db(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()
    con = sqlite3.connect(path)
    con.execute("PRAGMA journal_mode=DELETE")
    return con


def close_db(con: sqlite3.Connection) -> None:
    con.commit()
    con.execute("VACUUM")
    con.close()


# ------------------------------------------------------------------ licensing domain
def person(i: int):
    female = rng.random() < 0.42
    first = rng.choice(FIRST_F if female else FIRST_M)
    name = f"{first} {rng.choice(FATHER)} {rng.choice(FAMILY)}"
    nid = "9" + "".join(str(rng.randrange(10)) for _ in range(9))
    mobile = "0599" + "".join(str(rng.randrange(10)) for _ in range(6))
    email = f"practitioner{i:05d}@example.com"
    birth = rand_date(date(1962, 1, 1), date(2001, 12, 31))
    region = rng.choices([r[0] for r in REGIONS], REGION_W)[0]
    reg = rand_date(date(2016, 1, 1), date(2026, 9, 30))
    return name, nid, mobile, email, birth, region, reg


def build_licensing():
    con = new_db(SRC / "licensing.db")
    con.executescript("""
    CREATE TABLE regions (region_code TEXT PRIMARY KEY, name_ar TEXT NOT NULL);
    CREATE TABLE professions (profession_code TEXT PRIMARY KEY, name_ar TEXT NOT NULL, name_en TEXT NOT NULL, cpd_hours_required INTEGER NOT NULL);
    CREATE TABLE fee_schedule (profession_code TEXT NOT NULL, kind TEXT NOT NULL, fee_sar REAL NOT NULL, effective_from TEXT NOT NULL);
    CREATE TABLE practitioners (practitioner_id TEXT PRIMARY KEY, full_name TEXT NOT NULL, national_id TEXT NOT NULL, mobile TEXT,
        email TEXT, birth_date TEXT, region_code TEXT, registered_on TEXT NOT NULL);
    CREATE TABLE licences (licence_no TEXT PRIMARY KEY, practitioner_id TEXT NOT NULL, profession_code TEXT NOT NULL, licence_class TEXT NOT NULL,
        status TEXT NOT NULL, issued_on TEXT NOT NULL, expiry_date TEXT NOT NULL, last_renewed_on TEXT);
    CREATE TABLE applications (application_id TEXT PRIMARY KEY, practitioner_id TEXT NOT NULL, profession_code TEXT NOT NULL, kind TEXT NOT NULL,
        status TEXT NOT NULL, submitted_on TEXT NOT NULL, decided_on TEXT, channel TEXT NOT NULL);
    CREATE TABLE renewals (renewal_id TEXT PRIMARY KEY, licence_no TEXT NOT NULL, requested_on TEXT NOT NULL, cpd_hours INTEGER,
        decision TEXT NOT NULL, decided_on TEXT);
    CREATE TABLE fee_payments (payment_id TEXT PRIMARY KEY, application_id TEXT NOT NULL, amount_sar REAL NOT NULL, paid_on TEXT NOT NULL, method TEXT NOT NULL);
    """)
    con.executemany("INSERT INTO regions VALUES (?,?)", REGIONS)
    con.executemany("INSERT INTO professions VALUES (?,?,?,?)", PROFESSIONS)
    con.executemany("INSERT INTO fee_schedule VALUES (?,?,?,?)",
                    [(p[0], k, v + (100.0 if p[0] == "SRV" else 0.0), "2025-01-01") for p in PROFESSIONS for k, v in FEES.items()])

    n = 2400
    practitioners, licences, applications, renewals, payments = [], [], [], [], []
    for i in range(1, n + 1):
        name, nid, mobile, email, birth, region, reg = person(i)
        pid = f"P-{i:05d}"
        if rng.random() < 0.03:
            mobile = None
        practitioners.append((pid, name, nid, mobile, email, d(birth), region, d(reg)))
        # most practitioners hold one licence; 6% hold none (registered, not licensed)
        if rng.random() < 0.06:
            applications.append((pid, None, reg))
            continue
        prof = rng.choices([p[0] for p in PROFESSIONS], PROF_W)[0]
        issued = max(reg + timedelta(days=rng.randrange(10, 60)), date(2016, 3, 1))
        if issued > date(2026, 10, 20):
            issued = date(2026, 10, 20) - timedelta(days=rng.randrange(1, 40))
        # three-year licences, renewed every three years
        expiry = issued + timedelta(days=365 * 3)
        last_ren = None
        while expiry < AS_OF and rng.random() < 0.82:
            last_ren = expiry - timedelta(days=rng.randrange(5, 45))
            expiry = expiry + timedelta(days=365 * 3)
        status = "active" if expiry >= AS_OF else "expired"
        r = rng.random()
        if r < 0.025:
            status = "suspended"
        elif r < 0.04:
            status = "cancelled"
        elif status == "expired" and rng.random() < 0.22:
            status = "active"  # planted: expired but not updated
        lic = f"L-{issued.year}-{len(licences) + 1:06d}"
        cls = "A" if rng.random() < 0.7 else "B"
        licences.append((lic, pid, prof, cls, status, d(issued), d(expiry), d(last_ren) if last_ren else None))
        applications.append((pid, prof, issued - timedelta(days=rng.randrange(3, 25))))
        if last_ren:
            applications.append((pid, prof, last_ren))
        if expiry >= AS_OF and (expiry - AS_OF).days < 120 and rng.random() < 0.5:
            renewals.append((lic, rand_date(AS_OF - timedelta(days=40), AS_OF)))
        if last_ren:
            renewals.append((lic, last_ren))
    # the two values the leaking drafts repeat: a name in the public register, a mobile in the reminders
    by_id = {p[0]: k for k, p in enumerate(practitioners)}
    in_register = next(l for l in licences if l[4] == "active")
    k = by_id[in_register[1]]
    practitioners[k] = (practitioners[k][0], "هند سعيد القحطاني") + practitioners[k][2:]
    reminded = next(l for l in licences if l[4] == "active" and "2026-11-30" <= l[6] < "2026-12-30" and practitioners[by_id[l[1]]][3])
    k = by_id[reminded[1]]
    practitioners[k] = practitioners[k][:3] + ("0599100234",) + practitioners[k][4:]
    con.executemany("INSERT INTO practitioners VALUES (?,?,?,?,?,?,?,?)", practitioners)
    con.executemany("INSERT INTO licences VALUES (?,?,?,?,?,?,?,?)", licences)

    app_rows = []
    for k, (pid, prof, sub) in enumerate(sorted(applications, key=lambda a: (a[2], a[0])), start=1):
        aid = f"A-{k:06d}"
        prof = prof or rng.choices([p[0] for p in PROFESSIONS], PROF_W)[0]
        kind = "new"
        if any(l[1] == pid and l[7] == d(sub) for l in licences):
            kind = "renewal"
        elif rng.random() < 0.05:
            kind = "upgrade"
        days = max(1, int(rng.gammavariate(2.2, 3.2)))
        decided = sub + timedelta(days=days)
        if decided > AS_OF:
            status, dec = rng.choice(["received", "under_review"]), None
        else:
            status = "approved" if rng.random() < 0.88 else rng.choice(["rejected", "withdrawn"])
            dec = d(decided)
        channel = "portal" if rng.random() < 0.83 else "service_centre"
        app_rows.append((aid, pid, prof, kind, status, d(sub), dec, channel))
        if status != "withdrawn":
            amount = FEES[kind] + (100.0 if prof == "SRV" else 0.0)
            method = rng.choices(["card", "sadad", "transfer"], [60, 35, 5])[0]
            payments.append((f"F-{len(payments) + 1:06d}", aid, amount, d(sub), method))
    con.executemany("INSERT INTO applications VALUES (?,?,?,?,?,?,?,?)", app_rows)
    con.executemany("INSERT INTO fee_payments VALUES (?,?,?,?,?)", payments)

    req = {p[0]: p[3] for p in PROFESSIONS}
    prof_of = {l[0]: l[2] for l in licences}
    ren_rows = []
    for k, (lic, when) in enumerate(sorted(renewals, key=lambda r: (r[1], r[0])), start=1):
        need = req[prof_of[lic]]
        hours = max(0, int(rng.gauss(need + 4, 6)))
        if when > AS_OF - timedelta(days=10):
            decision, dec = "pending", None
        else:
            decision = "approved" if hours >= need else rng.choice(["rejected", "approved"])
            dec = d(when + timedelta(days=rng.randrange(1, 12)))
        ren_rows.append((f"R-{k:06d}", lic, d(when), hours, decision, dec))
    con.executemany("INSERT INTO renewals VALUES (?,?,?,?,?,?)", ren_rows)

    # planted: two working tables that are not assets
    con.execute("CREATE TABLE tmp_import_0925 AS SELECT application_id, practitioner_id, submitted_on FROM applications WHERE submitted_on >= '2026-09-01' AND submitted_on < '2026-10-01'")
    con.execute("CREATE TABLE licences_backup_2024 AS SELECT * FROM licences WHERE issued_on < '2025-01-01'")
    con.execute("CREATE VIEW v_active_licences AS SELECT licence_no, practitioner_id, profession_code, licence_class, expiry_date "
                "FROM licences WHERE status = 'active' AND expiry_date >= date('now')")
    close_db(con)
    return practitioners, licences, app_rows, ren_rows, payments


def month_end(y: int, m: int) -> date:
    return (date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1))


def build_licensing_files(practitioners, licences, app_rows):
    fs = SRC / "filestore" / "licensing"
    mob = {p[0]: p[3] for p in practitioners}
    name = {p[0]: p[1] for p in practitioners}
    region = {p[0]: p[6] for p in practitioners}
    # renewal reminders: one file a month
    for (y, m) in [(2026, 9), (2026, 10)]:
        start = date(y, m, 1)
        rows = []
        for l in licences:
            exp = date.fromisoformat(l[6])
            if l[4] == "active" and start + timedelta(days=60) <= exp < start + timedelta(days=90) and mob[l[1]]:
                rows.append((l[0], l[1], mob[l[1]], l[6], d(start + timedelta(days=2))))
        write_csv(fs / "renewal_reminders" / f"{y}_{m:02d}.csv",
                  ["licence_no", "practitioner_id", "mobile", "expiry_date", "reminder_sent_on"], rows)
    # published licence stats, as reporting counts them (status only)
    for (y, m) in [(2026, 7), (2026, 8), (2026, 9)]:
        end = month_end(y, m)
        counts = {}
        for l in licences:
            if date.fromisoformat(l[5]) > end:
                continue
            key = (region[l[1]], l[2])
            c = counts.setdefault(key, [0, 0])
            if l[4] == "active":
                c[0] += 1
            if date.fromisoformat(l[5]).replace(day=1) == date(y, m, 1):
                c[1] += 1
        rows = [(f"{y}-{m:02d}", r, p, a, nw) for (r, p), (a, nw) in sorted(counts.items())]
        write_csv(fs / "published_licence_stats" / f"{y}_{m:02d}.csv",
                  ["month", "region_code", "profession_code", "active_licences", "new_licences"], rows)
    # public register: active and suspended licences, monthly snapshot
    rows = [(l[0], name[l[1]], l[2], l[4], l[6]) for l in licences if l[4] in ("active", "suspended")]
    write_csv(fs / "public_register" / "register_2026_10.csv",
              ["licence_no", "full_name", "profession_code", "status", "expiry_date"], rows)
    # published application stats and processing times
    for (y, m) in [(2026, 8), (2026, 9)]:
        pref = f"{y}-{m:02d}"
        counts, times = {}, {}
        for a in app_rows:
            if a[5].startswith(pref):
                counts[(a[3], a[4])] = counts.get((a[3], a[4]), 0) + 1
                if a[6]:
                    times.setdefault(a[2], []).append((date.fromisoformat(a[6]) - date.fromisoformat(a[5])).days)
        write_csv(fs / "published_application_stats" / f"{y}_{m:02d}.csv", ["month", "kind", "status", "applications", "channel_portal_share"],
                  [(pref, k, s, n, "") for (k, s), n in sorted(counts.items())])
        rows = []
        for p, ts in sorted(times.items()):
            ts.sort()
            rows.append((pref, p, ts[len(ts) // 2], len(ts)))
        write_csv(fs / "published_processing_times" / f"{y}_{m:02d}.csv", ["month", "profession_code", "median_days", "decided_applications"], rows)
    # staging: an unreviewed upload that must not be catalogued
    write_csv(fs / "_staging" / "upload_tmp.csv", ["col1", "col2"], [("x", "1"), ("y", "2")])


# ------------------------------------------------------------------ exams (teaching examples)
def build_exams():
    con = new_db(SRC / "exams.db")
    con.executescript("""
    CREATE TABLE exam_sessions (session_id TEXT PRIMARY KEY, profession_code TEXT NOT NULL, held_on TEXT NOT NULL, centre TEXT NOT NULL, seats INTEGER NOT NULL);
    CREATE TABLE candidates (candidate_id TEXT PRIMARY KEY, full_name TEXT NOT NULL, national_id TEXT NOT NULL, mobile TEXT);
    CREATE TABLE exam_results (result_id TEXT PRIMARY KEY, session_id TEXT NOT NULL, candidate_id TEXT NOT NULL, score INTEGER, passed INTEGER);
    """)
    sessions = []
    for k in range(1, 49):
        sessions.append((f"S-{k:03d}", rng.choices([p[0] for p in PROFESSIONS], PROF_W)[0],
                         d(rand_date(date(2025, 1, 1), date(2026, 10, 15))), rng.choice(["C1", "C2", "C3", "C4"]), rng.choice([30, 40, 60])))
    cands, results = [], []
    for k in range(1, 1201):
        nm, nid, mob, *_ = person(10000 + k)
        cands.append((f"C-{k:05d}", nm, nid, mob))
    for k in range(1, 1601):
        s = rng.choice(sessions)
        score = max(0, min(100, int(rng.gauss(68, 14))))
        results.append((f"X-{k:05d}", s[0], f"C-{rng.randrange(1, 1201):05d}", score if rng.random() > 0.02 else None, int(score >= 60)))
    con.executemany("INSERT INTO exam_sessions VALUES (?,?,?,?,?)", sessions)
    con.executemany("INSERT INTO candidates VALUES (?,?,?,?)", cands)
    con.executemany("INSERT INTO exam_results VALUES (?,?,?,?,?)", results)
    close_db(con)


# ------------------------------------------------------------------ accreditation (follow-up task)
def build_accreditation(practitioners):
    con = new_db(SRC / "accreditation.db")
    con.executescript("""
    CREATE TABLE providers (provider_id TEXT PRIMARY KEY, name_ar TEXT NOT NULL, city TEXT, contact_email TEXT, accredited_since TEXT);
    CREATE TABLE accreditations (accreditation_id TEXT PRIMARY KEY, provider_id TEXT NOT NULL, status TEXT NOT NULL, valid_from TEXT, valid_to TEXT);
    CREATE TABLE cpd_courses (course_id TEXT PRIMARY KEY, provider_id TEXT NOT NULL, profession_code TEXT NOT NULL, title_ar TEXT NOT NULL, hours INTEGER NOT NULL, delivery TEXT);
    CREATE TABLE cpd_records (record_id TEXT PRIMARY KEY, practitioner_id TEXT NOT NULL, course_id TEXT NOT NULL, completed_on TEXT, hours INTEGER, certificate_no TEXT);
    CREATE TABLE provider_visits (visit_id TEXT PRIMARY KEY, provider_id TEXT NOT NULL, visited_on TEXT, outcome TEXT, notes TEXT);
    CREATE TABLE tmp_cert_upload AS SELECT 1 AS x;
    """)
    words = ["أفق", "رواد", "مسار", "نماء", "إتقان", "خبرة", "مهارة", "ركيزة", "قدرات", "معرفة"]
    providers, accs, courses, records, visits = [], [], [], [], []
    for k in range(1, 41):
        pid = f"TP-{k:03d}"
        providers.append((pid, f"مركز {words[k % 10]} للتدريب {k}", rng.choice([r[1] for r in REGIONS[:6]]), f"provider{k:03d}@example.com", d(rand_date(date(2018, 1, 1), date(2026, 6, 1)))))
        start = rand_date(date(2024, 1, 1), date(2026, 6, 1))
        accs.append((f"AC-{k:03d}", pid, rng.choices(["valid", "expired", "suspended"], [80, 15, 5])[0], d(start), d(start + timedelta(days=730))))
        visits.append((f"V-{k:03d}", pid, d(rand_date(date(2025, 1, 1), date(2026, 10, 1))), rng.choice(["compliant", "follow_up"]), "ملاحظات الزيارة في نظام الاعتماد"))
    topics = ["أخلاقيات المهنة", "تحديثات الأنظمة", "أدوات القياس", "إدارة المشاريع", "السلامة في الموقع", "الترجمة القانونية", "التصميم المستدام", "تقييم المخاطر"]
    for k in range(1, 161):
        courses.append((f"CC-{k:04d}", f"TP-{rng.randrange(1, 41):03d}", rng.choices([p[0] for p in PROFESSIONS], PROF_W)[0], f"{rng.choice(topics)} {k}", rng.choice([3, 4, 6, 8, 12]), rng.choice(["onsite", "online"])))
    for k in range(1, 3001):
        c = rng.choice(courses)
        records.append((f"CR-{k:05d}", f"P-{rng.randrange(1, 2401):05d}", c[0], d(rand_date(date(2024, 1, 1), date(2026, 10, 25))), c[4], f"CERT-{k:06d}"))
    con.executemany("INSERT INTO providers VALUES (?,?,?,?,?)", providers)
    con.executemany("INSERT INTO accreditations VALUES (?,?,?,?,?)", accs)
    con.executemany("INSERT INTO cpd_courses VALUES (?,?,?,?,?,?)", courses)
    con.executemany("INSERT INTO cpd_records VALUES (?,?,?,?,?,?)", records)
    con.executemany("INSERT INTO provider_visits VALUES (?,?,?,?,?)", visits)
    close_db(con)
    fs = SRC / "filestore" / "accreditation"
    write_csv(fs / "published_provider_list" / "providers_2026_10.csv", ["provider_id", "name_ar", "city", "accreditation_status", "valid_to"],
              [(p[0], p[1], p[2], a[2], a[4]) for p, a in zip(providers, accs) if a[2] == "valid"])
    write_csv(fs / "cpd_course_catalogue" / "courses_2026_q4.csv", ["course_id", "provider_id", "profession_code", "title_ar", "hours", "delivery"], courses)
    by = {}
    for r in records:
        m = r[3][:7]
        by[(m, next(c[2] for c in courses if c[0] == r[2]))] = by.get((m, next(c[2] for c in courses if c[0] == r[2])), 0) + r[4]
    write_csv(fs / "published_cpd_hours" / "cpd_hours_2026.csv", ["month", "profession_code", "hours_recorded"],
              [(m, p, h) for (m, p), h in sorted(by.items()) if m.startswith("2026")])
    write_csv(fs / "certificate_uploads" / "uploads_2026_10.csv", ["record_id", "practitioner_id", "certificate_no", "uploaded_on"],
              [(r[0], r[1], r[5], r[3]) for r in records[-200:]])
    write_csv(fs / "_staging" / "provider_import.csv", ["a", "b"], [("1", "2")])


# ------------------------------------------------------------------ recipes (vendor example as supplied)
RECIPE_DB = """# Harvest recipe: Sinwan licensing database (SQLite)
# Copied from the vendor's example. Edit only the lines marked EDIT.
source:
  type: sqlalchemy
  config:
    platform: sqlite
    platform_instance: licensing
    connect_uri: "sqlite:///files/sources/licensing.db"
    include_views: true
    table_pattern:
      deny: []            # EDIT: patterns of tables that are not assets
    profiling:
      enabled: false      # EDIT: profile statistics
      include_field_sample_values: true   # EDIT: sample values
sink:
  type: file
  config:
    filename: out/licensing_db.json
"""

RECIPE_FILES = """# Harvest recipe: Sinwan file store, licensing folder (local files)
# Copied from the vendor's example. Edit only the lines marked EDIT.
source:
  type: s3
  config:
    platform: file
    path_specs:
      - include: "files/sources/filestore/licensing/{table}/*.csv"
        exclude: []       # EDIT: folders that are not assets
    profiling:
      enabled: false      # EDIT: profile statistics
      include_field_sample_values: true   # EDIT: sample values
sink:
  type: file
  config:
    filename: out/licensing_files.json
"""

SCHEDULE = """# Harvest schedule for the trainer's catalogue instance (datahub ingest deploy)
# cron: minute hour day-of-month month day-of-week
harvests:
  - recipe: recipes/licensing_db.yml
    name: licensing-db
    schedule: "* * * * *"      # EDIT
    timezone: UTC              # EDIT
  - recipe: recipes/licensing_files.yml
    name: licensing-files
    schedule: "* * * * *"      # EDIT
    timezone: UTC              # EDIT
"""


# ------------------------------------------------------------------ lineage events
NS = {"licensing": "sqlite://licensing", "warehouse": "sqlite://warehouse", "files": "file://filestore", "bi": "bi://dashboards"}
PRODUCER = "https://example.com/sinwan/pipelines"
SCHEMA_URL = "https://openlineage.io/spec/2-0-2/OpenLineage.json#/$defs/RunEvent"
CL_URL = "https://openlineage.io/spec/facets/1-2-0/ColumnLineageDatasetFacet.json#/$defs/ColumnLineageDatasetFacet"


def field(ns, name, col, ttype="DIRECT", sub="IDENTITY"):
    return {"namespace": NS[ns], "name": name, "field": col, "transformations": [{"type": ttype, "subtype": sub}]}


JOBS = [
    {"job": "nightly.load_licence_facts", "every": 1, "last": date(2026, 10, 30), "owner": "فارس الزاعق",
     "inputs": [("licensing", "licences"), ("licensing", "practitioners")],
     "output": ("warehouse", "fact_licence"),
     "fields": {"licence_no": [field("licensing", "licences", "licence_no")],
                "status": [field("licensing", "licences", "status")],
                "profession_code": [field("licensing", "licences", "profession_code")],
                "region_code": [field("licensing", "practitioners", "region_code")],
                "issued_on": [field("licensing", "licences", "issued_on")],
                "expiry_date": [field("licensing", "licences", "expiry_date")]},
     "dataset": []},
    {"job": "nightly.build_monthly_stats", "every": 1, "last": date(2026, 10, 30), "owner": "عمار الوشمي",
     "inputs": [("warehouse", "fact_licence")], "output": ("warehouse", "monthly_licence_stats"),
     "fields": {"active_licences": [field("warehouse", "fact_licence", "licence_no", "DIRECT", "AGGREGATION")],
                "new_licences": [field("warehouse", "fact_licence", "licence_no", "DIRECT", "AGGREGATION")],
                "region_code": [field("warehouse", "fact_licence", "region_code")],
                "profession_code": [field("warehouse", "fact_licence", "profession_code")]},
     "dataset": [field("warehouse", "fact_licence", "status", "INDIRECT", "FILTER"),
                 field("warehouse", "fact_licence", "issued_on", "INDIRECT", "GROUP_BY")]},
    {"job": "monthly.export_open_stats", "every": 30, "last": date(2026, 10, 1), "owner": "غيداء الحوشان",
     "inputs": [("warehouse", "monthly_licence_stats")], "output": ("files", "licensing/published_licence_stats"),
     "fields": {c: [field("warehouse", "monthly_licence_stats", c)] for c in ["active_licences", "new_licences", "region_code", "profession_code"]},
     "dataset": []},
    {"job": "nightly.export_public_register", "every": 1, "last": date(2026, 10, 30), "owner": "راكان التريكي",
     "inputs": [("licensing", "licences"), ("licensing", "practitioners")], "output": ("files", "licensing/public_register"),
     "fields": {"licence_no": [field("licensing", "licences", "licence_no")],
                "full_name": [field("licensing", "practitioners", "full_name")],
                "profession_code": [field("licensing", "licences", "profession_code")],
                "status": [field("licensing", "licences", "status")],
                "expiry_date": [field("licensing", "licences", "expiry_date")]},
     "dataset": [field("licensing", "licences", "status", "INDIRECT", "FILTER")]},
    {"job": "weekly.renewal_reminders", "every": 7, "last": date(2026, 10, 26), "owner": "فارس الزاعق",
     "inputs": [("licensing", "licences"), ("licensing", "practitioners")], "output": ("files", "licensing/renewal_reminders"),
     "fields": {"licence_no": [field("licensing", "licences", "licence_no")],
                "practitioner_id": [field("licensing", "licences", "practitioner_id")],
                "mobile": [field("licensing", "practitioners", "mobile")],
                "expiry_date": [field("licensing", "licences", "expiry_date")]},
     "dataset": [field("licensing", "licences", "status", "INDIRECT", "FILTER"),
                 field("licensing", "licences", "expiry_date", "INDIRECT", "FILTER")]},
    {"job": "daily.refresh_renewals_dashboard", "every": 1, "last": date(2026, 10, 30), "owner": "عمار الوشمي",
     "inputs": [("warehouse", "fact_licence"), ("licensing", "renewals")], "output": ("bi", "renewals_overview"),
     "fields": {"expiring_90_days": [field("warehouse", "fact_licence", "licence_no", "DIRECT", "AGGREGATION")],
                "renewals_requested": [field("licensing", "renewals", "renewal_id", "DIRECT", "AGGREGATION")]},
     "dataset": [field("warehouse", "fact_licence", "status", "INDIRECT", "FILTER"),
                 field("warehouse", "fact_licence", "expiry_date", "INDIRECT", "FILTER")]},
    {"job": "nightly.load_application_facts", "every": 1, "last": date(2026, 10, 30), "owner": "فارس الزاعق",
     "inputs": [("licensing", "applications")], "output": ("warehouse", "fact_application"),
     "fields": {c: [field("licensing", "applications", c)] for c in ["application_id", "kind", "status", "submitted_on", "decided_on"]},
     "dataset": []},
    {"job": "legacy.quarterly_extract", "every": 91, "last": date(2026, 4, 2), "owner": "",
     "inputs": [("licensing", "licences")], "output": ("files", "licensing/legacy_quarterly"),
     "fields": {"licence_no": [field("licensing", "licences", "licence_no")], "status": [field("licensing", "licences", "status")]},
     "dataset": []},
]

DATASET_OWNERS = [
    ("sqlite://licensing", "licences", "راكان التريكي", "منيرة العبيد"),
    ("sqlite://licensing", "practitioners", "راكان التريكي", "منيرة العبيد"),
    ("sqlite://licensing", "applications", "راكان التريكي", "منيرة العبيد"),
    ("sqlite://licensing", "renewals", "راكان التريكي", "منيرة العبيد"),
    ("sqlite://warehouse", "fact_licence", "عمار الوشمي", "فارس الزاعق"),
    ("sqlite://warehouse", "fact_application", "عمار الوشمي", "فارس الزاعق"),
    ("sqlite://warehouse", "monthly_licence_stats", "عمار الوشمي", "فارس الزاعق"),
    ("file://filestore", "licensing/published_licence_stats", "غيداء الحوشان", "منيرة العبيد"),
    ("file://filestore", "licensing/public_register", "راكان التريكي", "منيرة العبيد"),
    ("file://filestore", "licensing/renewal_reminders", "راكان التريكي", "فارس الزاعق"),
    ("file://filestore", "licensing/legacy_quarterly", "", ""),
    ("bi://dashboards", "renewals_overview", "عمار الوشمي", "فارس الزاعق"),
]


def build_lineage():
    events = []
    for j in JOBS:
        # the last three runs of each job (fewer for the legacy one)
        runs = [j["last"] - timedelta(days=j["every"] * k) for k in range(2, -1, -1)]
        for when in runs:
            run_id = str(uuid.uuid5(uuid.NAMESPACE_URL, f"{j['job']}/{when}"))
            base = {"run": {"runId": run_id}, "job": {"namespace": "sinwan", "name": j["job"]},
                    "producer": PRODUCER, "schemaURL": SCHEMA_URL}
            inputs = [{"namespace": NS[a], "name": b} for a, b in j["inputs"]]
            ons, oname = j["output"]
            out = {"namespace": NS[ons], "name": oname, "facets": {"columnLineage": {
                "_producer": PRODUCER, "_schemaURL": CL_URL,
                "fields": {k: {"inputFields": v} for k, v in j["fields"].items()}}}}
            if j["dataset"]:
                out["facets"]["columnLineage"]["dataset"] = j["dataset"]
            events.append({"eventType": "START", "eventTime": f"{when}T01:{len(events) % 50:02d}:00Z", **base, "inputs": inputs, "outputs": []})
            events.append({"eventType": "COMPLETE", "eventTime": f"{when}T02:{len(events) % 50:02d}:00Z", **base, "inputs": inputs, "outputs": [out]})
    events.sort(key=lambda e: (e["eventTime"], e["job"]["name"]))
    path = OUT / "lineage" / "events.jsonl"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(("\n".join(json.dumps(e, ensure_ascii=False, sort_keys=True) for e in events) + "\n").encode("utf-8"))
    write_csv(OUT / "lineage" / "owners.csv", ["namespace", "name", "owner", "steward"], DATASET_OWNERS)
    write_text("lineage/change_request.txt", C.CHANGE_REQUEST)
    write_text("lineage/notes_from_reporting.txt", C.LINEAGE_NOTES)
    write_text("lineage/impact_note_template.txt", C.IMPACT_TEMPLATE)
    return events


# ------------------------------------------------------------------ catalogue state (lab 2.2)
STATE = [
    # dataset, platform, columns, linked, description, owner, bsteward, itsteward, classification, schedule_days, last_ok
    ("licensing.practitioners", "sqlite", 8, 4, "reviewed", "راكان التريكي", "منيرة العبيد", "فارس الزاعق", "مقيد", 1, date(2026, 10, 31)),
    ("licensing.licences", "sqlite", 8, 8, "reviewed", "راكان التريكي", "منيرة العبيد", "فارس الزاعق", "مقيد", 1, date(2026, 10, 31)),
    ("licensing.applications", "sqlite", 8, 5, "reviewed", "راكان التريكي", "منيرة العبيد", "فارس الزاعق", "مقيد", 1, date(2026, 10, 31)),
    ("licensing.renewals", "sqlite", 6, 3, "reviewed", "راكان التريكي", "منيرة العبيد", "فارس الزاعق", "مقيد", 1, date(2026, 10, 31)),
    ("licensing.fee_payments", "sqlite", 5, 2, "machine_draft", "مدير الشؤون المالية", "", "فارس الزاعق", "مقيد", 1, date(2026, 10, 31)),
    ("licensing.fee_schedule", "sqlite", 4, 0, "machine_draft", "مدير الشؤون المالية", "", "فارس الزاعق", "عام", 1, date(2026, 10, 31)),
    ("licensing.professions", "sqlite", 4, 2, "reviewed", "راكان التريكي", "منيرة العبيد", "فارس الزاعق", "عام", 1, date(2026, 10, 31)),
    ("licensing.regions", "sqlite", 2, 1, "reviewed", "راكان التريكي", "منيرة العبيد", "فارس الزاعق", "عام", 1, date(2026, 10, 31)),
    ("licensing.v_active_licences", "sqlite", 5, 0, "none", "", "", "فارس الزاعق", "", 1, date(2026, 10, 31)),
    ("filestore.licensing.renewal_reminders", "file", 5, 2, "reviewed", "راكان التريكي", "", "فارس الزاعق", "مقيد", 7, date(2026, 10, 4)),
    ("filestore.licensing.published_licence_stats", "file", 5, 1, "machine_draft", "غيداء الحوشان", "منيرة العبيد", "فارس الزاعق", "عام", 7, date(2026, 10, 4)),
    ("filestore.licensing.public_register", "file", 5, 2, "reviewed", "راكان التريكي", "منيرة العبيد", "فارس الزاعق", "عام", 7, date(2026, 10, 4)),
    ("filestore.licensing.published_application_stats", "file", 5, 0, "machine_draft", "غيداء الحوشان", "", "فارس الزاعق", "", 7, date(2026, 10, 4)),
    ("filestore.licensing.published_processing_times", "file", 4, 0, "none", "", "", "فارس الزاعق", "", 7, date(2026, 10, 4)),
    ("warehouse.fact_licence", "sqlite", 6, 0, "reviewed", "عمار الوشمي", "", "فارس الزاعق", "مقيد", 1, date(2026, 10, 31)),
    ("warehouse.monthly_licence_stats", "sqlite", 4, 0, "none", "عمار الوشمي", "", "فارس الزاعق", "عام", 1, date(2026, 10, 31)),
]

HEALTH_CONFIG = """# Health measures for the licensing domain (kit/health.py)
# Which attributes must a dataset have to count as complete? true or false. EDIT
required:
  description_reviewed: false
  owner: true
  business_steward: false
  it_steward: false
  classification: false
  any_column_linked_to_term: false
# Does a machine-drafted description count as described? EDIT
count_machine_drafts_as_described: true
# A harvest is fresh when the last successful run is no older than this many schedule intervals. EDIT
fresh_within_intervals: 10
as_of: "2026-10-31"
"""


def build_state():
    rows = [(r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], d(r[10])) for r in STATE]
    wb = openpyxl.Workbook()
    sheet(wb, "datasets", ["dataset", "platform", "columns", "columns_linked_to_terms", "description_status", "owner",
                           "business_steward", "it_steward", "classification", "schedule_days", "last_successful_harvest"],
          rows, widths=[40, 10, 9, 12, 16, 20, 18, 16, 12, 10, 16])
    runs = []
    for k in range(28):
        day = AS_OF - timedelta(days=27 - k)
        runs.append(("licensing-db", d(day), "success", 11))
        if day.weekday() == 6:  # Sunday
            runs.append(("licensing-files", d(day), "success" if day <= date(2026, 10, 4) else "failed", 5 if day <= date(2026, 10, 4) else 0))
    for day in [date(2026, 10, 4)]:
        pass
    sheet(wb, "harvest_runs", ["pipeline", "run_date", "result", "datasets_written"], runs, widths=[18, 14, 10, 14])
    save(wb, OUT / "reference" / "catalogue_state.xlsx")
    (OUT / "reference" / "health_config.yml").write_bytes(HEALTH_CONFIG.encode("utf-8"))
    return rows, runs


# ------------------------------------------------------------------ dataset records (lab 2.3)
RECORDS_START = [
    # identifier, title_ar, title_en, description_ar, contact_name, contact_email, licence, keywords, theme, periodicity, modified, download_url, media_type
    ("sinwan-licence-stats", "إحصاءات الرخص المهنية الشهرية", "Monthly professional licence statistics",
     "عدد الرخص المهنية السارية والجديدة في آخر كل شهر، بحسب المهنة المنظمة والمنطقة الإدارية.",
     "مكتب البيانات المفتوحة", "opendata@example.com", "مجاني للجميع", "رخص، مهن، إحصاءات", "GOVE", "MONTHLY", "2026-10-01",
     "https://data.example.org/sinwan/licence-stats.csv", "text/csv"),
    ("sinwan-application-stats", "إحصاءات طلبات الترخيص الشهرية", "Monthly licensing application statistics",
     "عدد طلبات الترخيص في كل شهر بحسب نوع الطلب ومرحلته.",
     "مكتب البيانات المفتوحة", "", "CC-BY-4.0", "طلبات، ترخيص", "GOVE", "MONTHLY", "2026-10-01",
     "https://data.example.org/sinwan/application-stats.csv", "text/csv"),
    ("sinwan-processing-times", "مدة معالجة طلبات الترخيص", "Licensing application processing times",
     "الوسيط الشهري لمدة معالجة الطلب بالأيام من التقديم إلى القرار، لكل مهنة.",
     "مكتب البيانات المفتوحة", "opendata@example.com", "CC-BY-4.0", "مدة المعالجة، خدمات", "GOVE", "شهري", "2026-10-01",
     "https://data.example.org/sinwan/processing-times.csv", "text/csv"),
    ("sinwan-professions", "المهن المنظمة وساعات التطوير المهني", "",
     "المهن الأربع التي ترخصها الهيئة، وساعات التطوير المهني المطلوبة سنويًا لكل منها.",
     "مكتب البيانات المفتوحة", "opendata@example.com", "CC-BY-4.0", "مهن، تطوير مهني", "EDUC", "ANNUAL", "2025-01-01",
     "https://data.example.org/sinwan/professions.csv", "text/csv"),
    ("sinwan-fee-schedule", "جدول رسوم الترخيص", "Licensing fee schedule",
     "الرسوم المعتمدة بالريال لكل مهنة ونوع طلب، وتاريخ سريانها.",
     "مكتب البيانات المفتوحة", "opendata@example.com", "CC-BY-4.0", "رسوم", "ECON", "ANNUAL", "2025-01-01",
     "", "CSV"),
]

TARGET_PROFILE = """@prefix sh: <http://www.w3.org/ns/shacl#> .
@prefix dcat: <http://www.w3.org/ns/dcat#> .
@prefix dct: <http://purl.org/dc/terms/> .
@prefix vcard: <http://www.w3.org/2006/vcard/ns#> .
@prefix foaf: <http://xmlns.com/foaf/0.1/> .
@prefix xsd: <http://www.w3.org/2001/XMLSchema#> .
@prefix tp: <https://data.example.org/target-profile#> .

# Target portal profile of the Sinwan case (written for the course LDS-DMG-202; not a real portal's profile)

tp:DatasetShape a sh:NodeShape ;
  sh:targetClass dcat:Dataset ;
  sh:property [ sh:path dct:identifier ; sh:minCount 1 ; sh:maxCount 1 ; sh:message "identifier: one is required" ] ;
  sh:property [ sh:path dct:title ; sh:qualifiedValueShape [ sh:languageIn ( "ar" ) ] ; sh:qualifiedMinCount 1 ; sh:message "title: an Arabic title is required" ] ;
  sh:property [ sh:path dct:title ; sh:qualifiedValueShape [ sh:languageIn ( "en" ) ] ; sh:qualifiedMinCount 1 ; sh:message "title: an English title is required" ] ;
  sh:property [ sh:path dct:description ; sh:qualifiedValueShape [ sh:languageIn ( "ar" ) ] ; sh:qualifiedMinCount 1 ; sh:message "description: an Arabic description is required" ] ;
  sh:property [ sh:path dct:publisher ; sh:minCount 1 ; sh:nodeKind sh:IRI ; sh:message "publisher: required, as a link" ] ;
  sh:property [ sh:path dcat:contactPoint ; sh:minCount 1 ; sh:node tp:ContactShape ; sh:message "contact point: required, with an e-mail" ] ;
  sh:property [ sh:path dct:license ; sh:minCount 1 ; sh:in ( <https://creativecommons.org/licenses/by/4.0/> ) ; sh:message "licence: one from the portal's list is required" ] ;
  sh:property [ sh:path dcat:keyword ; sh:minCount 2 ; sh:message "keywords: at least two are required" ] ;
  sh:property [ sh:path dcat:theme ; sh:minCount 1 ; sh:in ( tp:ECON tp:GOVE tp:EDUC ) ; sh:message "theme: one from the portal's list is required" ] ;
  sh:property [ sh:path dct:accrualPeriodicity ; sh:minCount 1 ; sh:in (
      <http://publications.europa.eu/resource/authority/frequency/DAILY>
      <http://publications.europa.eu/resource/authority/frequency/WEEKLY>
      <http://publications.europa.eu/resource/authority/frequency/MONTHLY>
      <http://publications.europa.eu/resource/authority/frequency/QUARTERLY>
      <http://publications.europa.eu/resource/authority/frequency/ANNUAL> ) ; sh:message "periodicity: one from the portal's list is required" ] ;
  sh:property [ sh:path dct:modified ; sh:minCount 1 ; sh:datatype xsd:date ; sh:message "modified: a date (year-month-day) is required" ] ;
  sh:property [ sh:path dcat:distribution ; sh:minCount 1 ; sh:node tp:DistributionShape ; sh:message "distribution: at least one, with a download link and a media type" ] .

tp:ContactShape a sh:NodeShape ;
  sh:property [ sh:path vcard:hasEmail ; sh:minCount 1 ; sh:nodeKind sh:IRI ; sh:pattern "^mailto:[^@]+@[^@]+$" ] .

tp:DistributionShape a sh:NodeShape ;
  sh:property [ sh:path dcat:downloadURL ; sh:minCount 1 ; sh:nodeKind sh:IRI ; sh:pattern "^https://" ] ;
  sh:property [ sh:path dcat:mediaType ; sh:minCount 1 ; sh:nodeKind sh:IRI ; sh:pattern "^https://www.iana.org/assignments/media-types/" ] .
"""


def build_records():
    wb = openpyxl.Workbook()
    header = ["identifier", "title_ar", "title_en", "description_ar", "contact_name", "contact_email", "licence", "keywords",
              "theme", "periodicity", "modified", "download_url", "media_type"]
    sheet(wb, "records", header, RECORDS_START, fill_cols=range(1, 13), widths=[24, 30, 30, 50, 20, 24, 14, 24, 8, 12, 12, 40, 12])
    sheet(wb, "lists", ["licence", "theme", "periodicity", "media_type"],
          [("CC-BY-4.0", "ECON", "DAILY", "text/csv"), ("", "GOVE", "WEEKLY", "application/json"), ("", "EDUC", "MONTHLY", ""),
           ("", "", "QUARTERLY", ""), ("", "", "ANNUAL", "")], widths=[14, 10, 14, 18])
    save(wb, OUT / "records" / "dataset_records.xlsx")
    write_text("records/target_fields.txt", C.TARGET_FIELDS)
    (OUT / "records" / "target_profile.ttl").write_bytes(TARGET_PROFILE.encode("utf-8"))
    # the follow-up task starts from an empty workbook with the same columns
    wb = openpyxl.Workbook()
    sheet(wb, "records", header, [], widths=[24, 30, 30, 50, 20, 24, 14, 24, 8, 12, 12, 40, 12])
    save(wb, OUT / "capstone" / "accreditation_records.xlsx")


# ------------------------------------------------------------------ drafts and glossary workbooks
def build_drafts():
    wb = openpyxl.Workbook()
    rows = [(f"D{k:02d}", a, sent, draft, "", "", "", "") for k, (a, sent, draft, *_rest) in enumerate(C.DRAFTS, start=1)]
    sheet(wb, "drafts", ["draft_id", "asset", "sent_to_model", "drafted_description", "decision", "error_type", "corrected_text", "reviewer"],
          rows, fill_cols=(4, 5, 6, 7), widths=[8, 34, 30, 60, 10, 18, 50, 16],
          validations={"E": [C.ACCEPT, C.EDIT, C.REJECT], "F": C.ERROR_TYPES})
    save(wb, OUT / "drafts" / "drafts.xlsx")
    write_text("drafts/review_rubric.txt", C.REVIEW_RUBRIC)
    write_text("drafts/steward_notes.txt", C.GLOSSARY_NOTES)


def build_glossary():
    wb = openpyxl.Workbook()
    sheet(wb, "columns", ["dataset", "column", "type", "term_id", "reason_if_no_term"],
          [(a, b, t, "", "") for a, b, t in C.GLOSSARY_COLUMNS], fill_cols=(3, 4), widths=[34, 20, 10, 10, 30])
    sheet(wb, "terms", ["term_id", "name_ar", "name_en", "definition_ar", "owner", "steward", "classification", "status", "version", "approved_by", "written_without_ai"],
          [(f"T{k:02d}", "", "", "", "", "", "", "مسودة", "", "", "") for k in range(1, 16)],
          fill_cols=range(1, 11), widths=[8, 22, 26, 60, 18, 14, 12, 12, 8, 16, 12],
          validations={"G": C.LEVELS, "H": C.STATUSES, "K": ["نعم", "لا"]})
    save(wb, OUT / "glossary" / "glossary_workbook.xlsx")
    write_text("glossary/steward_notes.txt", C.GLOSSARY_NOTES)


RECORDS_FIXED = {
    "sinwan-licence-stats": {"licence": "CC-BY-4.0"},
    "sinwan-application-stats": {"contact_email": "opendata@example.com"},
    "sinwan-processing-times": {"periodicity": "MONTHLY"},
    "sinwan-professions": {"title_en": "Regulated professions and CPD hours"},
    "sinwan-fee-schedule": {"download_url": "https://data.example.org/sinwan/fee-schedule.csv", "media_type": "text/csv", "keywords": "رسوم، ترخيص"},
}

HEALTH_REFERENCE = HEALTH_CONFIG.replace("description_reviewed: false", "description_reviewed: true")     .replace("business_steward: false", "business_steward: true").replace("it_steward: false", "it_steward: true")     .replace("classification: false", "classification: true").replace("count_machine_drafts_as_described: true", "count_machine_drafts_as_described: false")     .replace("fresh_within_intervals: 10", "fresh_within_intervals: 2")


def build_references():
    """Reference solutions as filled-in copies of the participants' files (labs/solutions/reference/)."""
    wb = openpyxl.Workbook()
    rows = [(f"D{k:02d}", a, sent, draft, dec, et, ref, C.STEWARD)
            for k, (a, sent, draft, dec, et, ref) in enumerate(C.DRAFTS, start=1)]
    sheet(wb, "drafts", ["draft_id", "asset", "sent_to_model", "drafted_description", "decision", "error_type", "corrected_text", "reviewer"],
          rows, widths=[8, 34, 30, 60, 10, 18, 50, 16])
    save(wb, REF / "drafts_reviewed.xlsx")
    term_of = {c: t[0] for t in C.TERMS for c in t[5]}
    wb = openpyxl.Workbook()
    sheet(wb, "columns", ["dataset", "column", "type", "term_id", "reason_if_no_term"],
          [(a, b, ty, term_of[f"{a}.{b}"], "") for a, b, ty in C.GLOSSARY_COLUMNS], widths=[34, 20, 10, 10, 30])
    sheet(wb, "terms", ["term_id", "name_ar", "name_en", "definition_ar", "owner", "steward", "classification", "status", "version", "approved_by", "written_without_ai"],
          [(t[0], t[1], t[2], C.TERM_DEFS[t[0]], t[3], C.STEWARD, t[4], "معتمد", "1.0", t[3], "نعم" if k < 5 else "لا")
           for k, t in enumerate(C.TERMS)], widths=[8, 22, 26, 60, 18, 14, 12, 12, 8, 16, 12])
    save(wb, REF / "glossary_reference.xlsx")
    header = ["identifier", "title_ar", "title_en", "description_ar", "contact_name", "contact_email", "licence", "keywords",
              "theme", "periodicity", "modified", "download_url", "media_type"]
    fixed = []
    for r in RECORDS_START:
        row = dict(zip(header, r))
        row.update(RECORDS_FIXED[row["identifier"]])
        fixed.append([row[h] for h in header])
    wb = openpyxl.Workbook()
    sheet(wb, "records", header, fixed)
    save(wb, REF / "dataset_records_reference.xlsx")
    (REF / "health_config_reference.yml").write_bytes(HEALTH_REFERENCE.encode("utf-8"))


def main():
    if OUT.exists():
        for p in sorted(OUT.rglob("*"), reverse=True):
            if p.is_file():
                p.unlink()
    practitioners, licences, app_rows, ren_rows, payments = build_licensing()
    build_licensing_files(practitioners, licences, app_rows)
    build_exams()
    build_accreditation(practitioners)
    (OUT / "recipes").mkdir(parents=True, exist_ok=True)
    (OUT / "recipes" / "licensing_db.yml").write_bytes(RECIPE_DB.encode("utf-8"))
    (OUT / "recipes" / "licensing_files.yml").write_bytes(RECIPE_FILES.encode("utf-8"))
    (OUT / "recipes" / "schedule.yml").write_bytes(SCHEDULE.encode("utf-8"))
    write_text("recipes/harvest_notes.txt", C.HARVEST_NOTES)
    build_drafts()
    build_glossary()
    build_lineage()
    state_rows, runs = build_state()
    build_records()
    write_text("capstone/brief.txt", C.CAPSTONE_BRIEF)
    build_references()
    write_text("case_brief.txt", f"""
{C.AGENCY}
{C.NOTE}

الهيئة جهة حكومية افتراضية ترخص ممارسي أربع مهن منظمة (الترجمة المعتمدة، والمساحة، والتصميم الداخلي، واستشارات السلامة المهنية) وتعتمد مزودي التدريب الذين يقدمون ساعات التطوير المهني المستمر.
أنشأت الهيئة مكتبًا لإدارة البيانات، ووافقت رئيسته على تشغيل كتالوج بيانات مفتوح المصدر يبدأ بمجال الترخيص.

الأشخاص
- {C.PEOPLE['head']}
- {C.PEOPLE['owner']}
- {C.PEOPLE['bsteward']}
- {C.PEOPLE['itsteward']}
- {C.PEOPLE['opendata']}
- {C.PEOPLE['reporting']}
- أنت: محلل الكتالوج في مكتب إدارة البيانات.

المصادر في sources/
- licensing.db: قاعدة بيانات الترخيص (الممارسون، والرخص، والطلبات، والتجديد، والرسوم، والقوائم المرجعية).
- filestore/licensing/: مجلد الترخيص في مخزن الملفات (تذكيرات التجديد، والإحصاءات المنشورة، والسجل العام).
- exams.db: قاعدة الاختبارات المهنية (أمثلة الشرائح).
- accreditation.db و filestore/accreditation/: مجال الاعتماد والتطوير المهني (مهمة المتابعة).
""")

    # ---- answer keys
    active_rule = sum(1 for l in licences if l[4] == "active" and l[6] >= d(AS_OF))
    active_status = sum(1 for l in licences if l[4] == "active")
    put_ref("lab_1_1", {"db_datasets": ["applications", "fee_payments", "fee_schedule", "licences", "practitioners", "professions", "regions", "renewals", "v_active_licences"],
                        "file_datasets": ["public_register", "published_application_stats", "published_licence_stats", "published_processing_times", "renewal_reminders"],
                        "excluded": ["tmp_import_0925", "licences_backup_2024", "_staging"],
                        "rows": {"practitioners": len(practitioners), "licences": len(licences), "applications": len(app_rows), "renewals": len(ren_rows), "fee_payments": len(payments)},
                        "schedule": {"licensing-db": "0 3 * * *", "licensing-files": "0 4 * * 0", "timezone": "Asia/Riyadh"}})
    put_ref("lab_1_2", [{"id": f"D{k:02d}", "asset": a, "decision": dec, "error_type": et, "reference_text": ref}
                        for k, (a, _s, _d, dec, et, ref) in enumerate(C.DRAFTS, start=1)])
    put_ref("lab_1_3", {"terms": [{"id": t[0], "name_ar": t[1], "name_en": t[2], "owner": t[3], "classification": t[4], "columns": t[5]} for t in C.TERMS],
                        "active_licences": {"by_rule": active_rule, "by_status_only": active_status, "as_of": d(AS_OF)}})
    put_ref("lab_2_2", {"state": [list(r) for r in state_rows], "runs": [list(r) for r in runs]})
    print("files written to", OUT.relative_to(LABS.parent))
    print("practitioners", len(practitioners), "licences", len(licences), "applications", len(app_rows), "renewals", len(ren_rows), "payments", len(payments))
    print("active by rule", active_rule, "active by status only", active_status)


if __name__ == "__main__":
    main()
