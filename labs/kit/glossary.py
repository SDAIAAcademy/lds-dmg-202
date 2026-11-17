"""Lab 1.3 · Check the glossary workbook, write it as the catalogue's glossary file and load it.

    uv run kit/glossary.py
    uv run kit/glossary.py --workbook path/to/my_glossary.xlsx

Checks 15 terms (names, definition, owner, steward, classification, status, version, approval by
the owner) and the links of the 30 columns, then writes out/glossary.yml in the business glossary
format of the open-source catalogue and runs its ingestion into out/glossary_catalogue.json, which
proves the catalogue accepts it. With a catalogue instance, the same file loads into the instance.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import yaml

from common import FILES, OUT, finish, fix, info, ok, read_sheet

LEVELS = {"عام", "مقيد", "سري", "سري للغاية"}
APPROVED = "معتمد"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=str(FILES / "glossary" / "glossary_workbook.xlsx"))
    ap.add_argument("--domain", default="الترخيص")
    a = ap.parse_args()
    wb = Path(a.workbook)
    terms = read_sheet(wb, "terms")
    cols = read_sheet(wb, "columns")
    ids = {t["term_id"] for t in terms if t.get("name_ar")}
    approved = 0
    for t in terms:
        tid = t.get("term_id", "?")
        if not t.get("name_ar"):
            fix(f"{tid}: لا اسم عربي")
            continue
        miss = [k for k in ("name_en", "definition_ar", "owner", "steward", "classification", "version") if not t.get(k)]
        if miss:
            fix(f"{tid} {t['name_ar']}: ينقصه {', '.join(miss)}")
        if t.get("classification") and t["classification"] not in LEVELS:
            fix(f"{tid}: مستوى التصنيف «{t['classification']}» ليس من المستويات الأربعة")
        if t.get("status") != APPROVED:
            fix(f"{tid}: الحالة «{t.get('status')}»، والمطلوب «{APPROVED}» بعد عرضه على المالك")
        elif t.get("approved_by") != t.get("owner"):
            fix(f"{tid}: اعتمده «{t.get('approved_by')}» والمالك «{t.get('owner')}»: يعتمد المالك وحده")
        else:
            approved += 1
    for t in terms[:5]:
        if t.get("name_ar") and t.get("written_without_ai") != "نعم":
            fix(f"{t['term_id']}: التعريفات الخمسة الأولى تُكتب دون أي أداة ذكاء اصطناعي")
    linked: dict[str, int] = {}
    for c in cols:
        key = f"{c['dataset']}.{c['column']}"
        tid = c.get("term_id", "")
        if tid:
            if tid not in ids:
                fix(f"{key}: مرتبط بـ {tid}، وهو مصطلح غير موجود")
            linked[tid] = linked.get(tid, 0) + 1
        elif not c.get("reason_if_no_term"):
            fix(f"{key}: لا مصطلح ولا سبب")
    for tid in sorted(ids):
        if not linked.get(tid):
            fix(f"{tid}: مصطلح غير مرتبط بأي عمود")
    n_linked = sum(linked.values())
    print(f"terms {len(ids)} · approved by their owner {approved} · columns linked {n_linked} of {len(cols)}")
    if len(ids) != 15:
        fix(f"في القاموس {len(ids)} مصطلحًا، والمطلوب 15")

    by_id = {t["term_id"]: t for t in terms if t.get("name_ar")}
    node = {"name": a.domain, "description": f"مصطلحات مجال {a.domain}", "terms": []}
    for tid, t in by_id.items():
        node["terms"].append({
            "name": t["name_ar"], "description": t.get("definition_ar") or "-", "term_source": "INTERNAL",
            "owners": {"users": [t.get("owner") or "unassigned"]},
            "custom_properties": {"term_id": tid, "name_en": t.get("name_en", ""), "classification": t.get("classification", ""),
                                  "status": t.get("status", ""), "version": t.get("version", ""), "steward": t.get("steward", ""),
                                  "approved_by": t.get("approved_by", ""),
                                  "columns": ", ".join(f"{c['dataset']}.{c['column']}" for c in cols if c.get("term_id") == tid)},
        })
    doc = {"version": "1", "source": "SinwanDMO", "owners": {"users": ["dmo"]}, "nodes": [node]}
    OUT.mkdir(exist_ok=True)
    gpath = OUT / "glossary.yml"
    gpath.write_text(yaml.safe_dump(doc, allow_unicode=True, sort_keys=False), encoding="utf-8")
    info(f"wrote {gpath.relative_to(OUT.parent)}")
    import os

    from datahub.ingestion.run.pipeline import Pipeline

    os.chdir(OUT.parent)  # the catalogue reads paths relative to labs/

    pipe = Pipeline.create({"source": {"type": "datahub-business-glossary", "config": {"file": "out/glossary.yml", "enable_auto_id": True}},
                            "sink": {"type": "file", "config": {"filename": "out/glossary_catalogue.json"}}})
    pipe.run()
    if pipe.source.get_report().failures or pipe.sink.get_report().failures:
        fix("الكتالوج رفض ملف القاموس: اقرأ رسالة الاستخلاص أعلاه")
    else:
        ok(f"قبل الكتالوج ملف القاموس: {len(node['terms'])} مصطلحًا في out/glossary_catalogue.json")
    return finish("الدليل: 15 مصطلحًا معتمدًا من مالكها، ومرتبطًا بالأعمدة، بمالك وتصنيف")


if __name__ == "__main__":
    raise SystemExit(main())
