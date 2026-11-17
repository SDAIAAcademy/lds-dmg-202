"""Lab 1.1 · Read what the harvest wrote and check it against the conditions of the harvest notes.

    uv run kit/harvest_check.py                       # licensing domain (lab 1.1)
    uv run kit/harvest_check.py --domain accreditation   # the follow-up task

It reads the catalogue files the two recipes wrote (out/<domain>_db.json, out/<domain>_files.json),
lists every dataset with its columns, rows and profile, and checks: every table, view and data
folder of the sources is catalogued; working tables and _staging are not; every table and file
has profile statistics; no value of a personal column is visible; the schedule is valid.
"""
from __future__ import annotations

import argparse
import re
import sqlite3

import yaml

from common import FILES, OUT, datasets, finish, fix, info, load_mcps, ok, short_name

PERSONAL = {
    "licensing": {"full_name", "national_id", "mobile", "email", "birth_date"},
    "accreditation": {"practitioner_id", "certificate_no", "contact_email"},
}
VALUE_KEYS = ("sampleValues", "distinctValueFrequencies", "histogram", "quantiles", "min", "max", "median", "mean")
NOT_ASSETS = re.compile(r"^(tmp_.*|.*_backup_\d{4})$")


def expected(domain: str):
    con = sqlite3.connect(FILES / "sources" / f"{domain}.db")
    objs = con.execute("SELECT name, type FROM sqlite_master WHERE type IN ('table','view') ORDER BY name").fetchall()
    con.close()
    tables = {n for n, t in objs if t == "table" and not NOT_ASSETS.match(n)}
    views = {n for n, t in objs if t == "view"}
    skipped = {n for n, t in objs if NOT_ASSETS.match(n)}
    root = FILES / "sources" / "filestore" / domain
    folders = {p.name for p in root.iterdir() if p.is_dir() and not p.name.startswith("_")}
    return tables, views, skipped, folders


def cron_ok(expr: str) -> bool:
    parts = expr.split()
    return len(parts) == 5 and all(re.fullmatch(r"[\d*/,-]+", p) for p in parts)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--domain", default="licensing")
    ap.add_argument("--schedule", default=None)
    a = ap.parse_args()
    dom = a.domain
    tables, views, skipped, folders = expected(dom)
    db = datasets(load_mcps(OUT / f"{dom}_db.json"))
    fl = datasets(load_mcps(OUT / f"{dom}_files.json"))
    if not db:
        fix(f"لم أجد out/{dom}_db.json: شغّل وصفة قاعدة البيانات أولًا")
    if not fl:
        fix(f"لم أجد out/{dom}_files.json: شغّل وصفة مخزن الملفات أولًا")

    print(f"{'dataset':34} {'kind':6} {'cols':>4} {'rows':>6}  profile  values")
    seen = {}
    for kind, group in (("db", db), ("files", fl)):
        for urn, asp in sorted(group.items(), key=lambda x: short_name(x[0])):
            name = short_name(urn)
            fields = (asp.get("schemaMetadata") or {}).get("fields", [])
            prof = asp.get("datasetProfile")
            leaks = []
            for fp in (prof or {}).get("fieldProfiles", []):
                if fp.get("fieldPath") in PERSONAL.get(dom, set()) and any(k in fp for k in VALUE_KEYS):
                    leaks.append(fp["fieldPath"])
            is_view = "view" in [s.lower() for s in (asp.get("subTypes") or {}).get("typeNames", [])]
            seen[name] = (kind, len(fields), prof, leaks, is_view)
            rows = prof.get("rowCount", "") if prof else ""
            print(f"{name:34} {kind:6} {len(fields):>4} {rows!s:>6}  {'yes' if prof else ('view' if is_view else 'NO'):7}  {'VISIBLE: ' + ','.join(leaks) if leaks else 'hidden'}")
    print()

    for t in sorted(tables | views):
        if t not in seen:
            fix(f"الجدول أو العرض {t} لم يُفهرس")
    for f in sorted(folders):
        if f not in seen:
            fix(f"مجلد البيانات {f} لم يُفهرس")
    for s in sorted(skipped | {"_staging"}):
        if s in seen:
            fix(f"{s} ليس أصلًا وقد فُهرس: أضف نمط استبعاد")
    for name, (kind, ncols, prof, leaks, is_view) in sorted(seen.items()):
        if ncols == 0:
            fix(f"{name}: لا مخطط بيانات")
        if not prof and not is_view and name not in skipped | {"_staging"}:
            fix(f"{name}: لا إحصاءات تنميط (profiling)")
        if leaks:
            fix(f"{name}: قيم ظاهرة في أعمدة شخصية ({', '.join(leaks)}): أطفئ القيم النموذجية وتكرار القيم")
    n_ok = len([1 for n in seen if n in tables | views | folders])
    if n_ok:
        ok(f"فُهرس {n_ok} من {len(tables | views | folders)} أصلًا ({len(tables)} جداول، {len(views)} عرض، {len(folders)} مجلدات ملفات)")

    sched_path = FILES / "recipes" / "schedule.yml" if a.schedule is None else FILES.parent / a.schedule
    if sched_path.exists():
        sched = yaml.safe_load(sched_path.read_text(encoding="utf-8")) or {}
        for h in sched.get("harvests", []):
            expr, tz = str(h.get("schedule", "")), str(h.get("timezone", ""))
            name = h.get("name", "?")
            if not cron_ok(expr):
                fix(f"الجدولة {name}: «{expr}» ليست تعبير cron من خمسة حقول")
            elif expr.split()[:2] == ["*", "*"] or expr.split()[1] == "*":
                fix(f"الجدولة {name}: «{expr}» تعمل كل دقيقة أو كل ساعة، والمطلوب مرة في اليوم أو الأسبوع")
            elif name.endswith("files") and expr.split()[4] == "*":
                fix(f"الجدولة {name}: مخزن الملفات أسبوعي، فحدد يوم الأسبوع في الحقل الخامس")
            else:
                ok(f"الجدولة {name}: {expr}")
            if tz != "Asia/Riyadh":
                fix(f"الجدولة {name}: المنطقة الزمنية {tz}، والمطلوب Asia/Riyadh")
            info(f"datahub ingest deploy --name {name} --schedule \"{expr}\" --time-zone {tz} -c files/{h.get('recipe')}")
    else:
        fix(f"لم أجد ملف الجدولة {sched_path.name}")
    return finish("الدليل: كل جدول وملف مفهرس بمخططه وإحصاءات تنميطه، بلا قيم شخصية ظاهرة، وبجدولة صحيحة")


if __name__ == "__main__":
    raise SystemExit(main())
