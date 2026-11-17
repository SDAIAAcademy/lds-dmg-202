"""Run every lab's reference solution and assert its numbers. From the labs/ folder:

    uv run solutions/check.py            # every lab and the follow-up task
    uv run solutions/check.py 1.1 2.3    # some of them

Lab 1.1 and the follow-up task run the open-source catalogue's real harvest (no server: the
catalogue writes to a local file). The other labs run the kit on the reference files in
solutions/reference/. It also checks that the supplied starting files still hold their planted
problems. No lab calls a language model. Outputs go to out/, which you may delete afterwards.
"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

LABS = Path(__file__).resolve().parents[1]
REF = LABS / "solutions" / "reference"
OUT = LABS / "out"
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def run(*args: str) -> tuple[int, str]:
    p = subprocess.run([sys.executable, *args], cwd=LABS, capture_output=True, text=True, encoding="utf-8")
    return p.returncode, p.stdout + p.stderr


def ingest(recipe: str) -> None:
    code, text = run("-m", "datahub", "ingest", "-c", recipe)
    assert code == 0 and "Pipeline finished" in text, f"harvest failed: {recipe}\n{text[-1500:]}"


def ref(name: str):
    return json.loads((REF / f"{name}.json").read_text(encoding="utf-8"))


def lab_1_1() -> None:
    key = ref("lab_1_1")
    shutil.rmtree(OUT, ignore_errors=True)
    ingest("files/recipes/licensing_db.yml")
    ingest("files/recipes/licensing_files.yml")
    code, text = run("kit/harvest_check.py")
    assert code == 1 and "tmp_import_0925" in text and "licences_backup_2024" in text and "UTC" in text, "the starting recipes lost a planted problem"
    shutil.rmtree(OUT, ignore_errors=True)
    ingest("solutions/recipes/licensing_db.yml")
    ingest("solutions/recipes/licensing_files.yml")
    code, text = run("kit/harvest_check.py", "--schedule", "solutions/recipes/schedule.yml")
    print(text)
    assert code == 0 and "RESULT: PASS" in text, "reference harvest does not pass"
    n = len(key["db_datasets"]) + len(key["file_datasets"])
    assert f"فُهرس {n} من {n}" in text, n
    for name, rows in key["rows"].items():
        assert re.search(rf"^{name}\s+db\s+\d+\s+{rows}\s", text, re.M), (name, rows)


def lab_1_2() -> None:
    code, text = run("kit/review_report.py", "--workbook", "solutions/reference/drafts_reviewed.xlsx")
    print(text)
    key = ref("lab_1_2")
    acc = sum(k["decision"] == "قبول" for k in key)
    edi = sum(k["decision"] == "تعديل" for k in key)
    rej = sum(k["decision"] == "رفض" for k in key)
    assert code == 0 and f"accept {acc} · edit {edi} · reject {rej}" in text, (acc, edi, rej)
    assert (acc, edi, rej) == (16, 15, 9)
    code, text = run("kit/review_report.py")
    assert code == 1 and "decided 0" in text, "the starting workbook is not blank"


def lab_1_3() -> None:
    code, text = run("kit/glossary.py", "--workbook", "solutions/reference/glossary_reference.xlsx")
    print("\n".join(l for l in text.splitlines() if l.startswith(("[", "terms", "RESULT"))))
    assert code == 0 and "terms 15 · approved by their owner 15 · columns linked 30 of 30" in text
    k = ref("lab_1_3")["active_licences"]
    print(f"active licences on {k['as_of']}: {k['by_rule']} by the owner's rule, {k['by_status_only']} by status only")
    assert k["by_status_only"] > k["by_rule"]


def lab_2_1() -> None:
    code, text = run("kit/impact.py", "--column", "licensing.licences.status")
    print(text)
    assert code == 0 and "datasets affected: 7 · jobs: 7" in text
    for ds in ["warehouse.fact_licence", "warehouse.monthly_licence_stats", "files.licensing/published_licence_stats",
               "files.licensing/public_register", "files.licensing/renewal_reminders", "bi.renewals_overview", "files.licensing/legacy_quarterly"]:
        assert ds in text, ds
    assert "legacy.quarterly_extract · last run 2026-04-02 · STALE" in text
    assert "fact_application" not in text, "applications.status is a different column"


def lab_2_2() -> None:
    code, text = run("kit/health.py", "--config", "solutions/reference/health_config_reference.yml")
    print(text)
    assert "description_reviewed           9/16" in text
    assert "complete (all required attributes): 7/16" in text
    assert "fresh (last good harvest within 2 schedule intervals): 11/16" in text
    assert "failed 3" in text
    code, text = run("kit/health.py")
    assert "description_reviewed          13/16" in text and "16/16 = 100.0%" in text, "the starting config lost its traps"


def lab_2_3() -> None:
    code, text = run("kit/dcat.py", "--workbook", "solutions/reference/dataset_records_reference.xlsx")
    print(text)
    assert code == 0 and text.count("passes the target portal's validator") == 5
    code, text = run("kit/dcat.py")
    assert code == 1 and text.count("passes the target portal's validator") == 0, "the starting records lost a planted problem"


def capstone() -> None:
    shutil.rmtree(OUT, ignore_errors=True)
    ingest("solutions/recipes/accreditation_db.yml")
    ingest("solutions/recipes/accreditation_files.yml")
    code, text = run("kit/harvest_check.py", "--domain", "accreditation", "--schedule", "solutions/recipes/accreditation_schedule.yml")
    print(text)
    assert code == 0 and "RESULT: PASS" in text


LABS_RUN = {"1.1": lab_1_1, "1.2": lab_1_2, "1.3": lab_1_3, "2.1": lab_2_1, "2.2": lab_2_2, "2.3": lab_2_3, "capstone": capstone}

if __name__ == "__main__":
    wanted = sys.argv[1:] or list(LABS_RUN)
    for lab in wanted:
        print(f"===== lab {lab}")
        LABS_RUN[lab]()
        print(f"PASS {lab}\n")
