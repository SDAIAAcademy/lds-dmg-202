"""Shared helpers of the LDS-DMG-202 lab kit. Run every kit command from the labs/ folder."""
from __future__ import annotations

import json
import sys
from pathlib import Path

LABS = Path(__file__).resolve().parents[1]
FILES = LABS / "files"
OUT = LABS / "out"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

PROBLEMS: list[str] = []


def ok(text: str) -> None:
    print(f"[OK]  {text}")


def fix(text: str) -> None:
    PROBLEMS.append(text)
    print(f"[FIX] {text}")


def info(text: str) -> None:
    print(f"      {text}")


def finish(evidence: str) -> int:
    print()
    if PROBLEMS:
        print(f"RESULT: {len(PROBLEMS)} to fix. النتيجة: {len(PROBLEMS)} بنود تحتاج إصلاحًا، ولم يتحقق بعد: {evidence}")
        return 1
    print(f"RESULT: PASS. النتيجة: تحقق {evidence}")
    return 0


def load_mcps(path: Path) -> list[dict]:
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    return data if isinstance(data, list) else []


def short_name(urn: str) -> str:
    """urn:li:dataset:(urn:li:dataPlatform:sqlite,licensing.main.licences,PROD) -> licences"""
    inner = urn.split(",")[1] if "," in urn else urn
    return inner.replace("\\", "/").rstrip("/").split("/")[-1].split(".")[-1]


def datasets(mcps: list[dict]) -> dict[str, dict]:
    """One entry per dataset: its aspects by name (the last one wins)."""
    out: dict[str, dict] = {}
    for m in mcps:
        urn = m.get("entityUrn", "")
        if not urn.startswith("urn:li:dataset:"):
            continue
        out.setdefault(urn, {})[m.get("aspectName")] = (m.get("aspect") or {}).get("json", {})
    return out


def read_sheet(path: Path, name: str) -> list[dict]:
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h).strip() if h is not None else "" for h in rows[0]]
    result = []
    for r in rows[1:]:
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        result.append({h: ("" if v is None else str(v).strip()) for h, v in zip(head, r)})
    return result
