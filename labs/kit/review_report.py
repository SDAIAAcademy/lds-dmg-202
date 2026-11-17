"""Lab 1.2 · Summarise the review log of the 40 machine-drafted descriptions.

    uv run kit/review_report.py                                   # files/drafts/drafts.xlsx
    uv run kit/review_report.py --workbook path/to/my_drafts.xlsx

Counts accept, edit and reject, the error rate ((edit + reject) / drafts) and the error types,
and checks that every row has a decision, a reviewer, a corrected text where one is needed, and
that every leaked value was rejected.
"""
from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path

from common import FILES, finish, fix, info, ok, read_sheet

ACCEPT, EDIT, REJECT = "قبول", "تعديل", "رفض"
LEAK, NONE = "قيمة مسربة", "لا خطأ"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=str(FILES / "drafts" / "drafts.xlsx"))
    a = ap.parse_args()
    rows = read_sheet(Path(a.workbook), "drafts")
    n = len(rows)
    dec = Counter(r.get("decision", "") for r in rows)
    types = Counter(r.get("error_type", "") for r in rows if r.get("decision") in (EDIT, REJECT))
    for r in rows:
        rid = r.get("draft_id", "?")
        d, t = r.get("decision", ""), r.get("error_type", "")
        if d not in (ACCEPT, EDIT, REJECT):
            fix(f"{rid}: لا قرار")
            continue
        if not r.get("reviewer"):
            fix(f"{rid}: لا مراجِع مسمّى")
        if d == ACCEPT and t not in ("", NONE):
            fix(f"{rid}: قبول مع نوع خطأ «{t}»: إما أن الوصف صحيح وإما أنه يحتاج تعديلًا")
        if d in (EDIT, REJECT):
            if t in ("", NONE):
                fix(f"{rid}: {d} بلا نوع خطأ")
            if not r.get("corrected_text"):
                fix(f"{rid}: {d} بلا نص مصحح")
        if t == LEAK and d != REJECT:
            fix(f"{rid}: قيمة مسربة لم تُرفض: الرفض والإبلاغ إلزاميان")
    decided = dec[ACCEPT] + dec[EDIT] + dec[REJECT]
    print(f"drafts {n} · decided {decided} · accept {dec[ACCEPT]} · edit {dec[EDIT]} · reject {dec[REJECT]}")
    if decided:
        rate = (dec[EDIT] + dec[REJECT]) / decided
        print(f"error rate (edit + reject) / decided = {dec[EDIT] + dec[REJECT]}/{decided} = {rate:.1%}")
        print("error types, commonest first:")
        for t, c in types.most_common():
            info(f"{c:3}  {t or '(blank)'}")
    if n != 40:
        fix(f"في الملف {n} مسودة، والمتوقع 40")
    elif decided == 40:
        ok("قُرر في المسودات الأربعين كلها")
    return finish("الدليل: سجل قبول أو تعديل أو رفض، بمراجِع لكل سطر، وأكثر أنواع الخطأ شيوعًا")


if __name__ == "__main__":
    raise SystemExit(main())
