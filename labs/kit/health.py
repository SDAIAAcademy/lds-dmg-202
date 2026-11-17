"""Lab 2.2 · Coverage, completeness and freshness of a domain in the catalogue.

    uv run kit/health.py
    uv run kit/health.py --state files/reference/catalogue_state.xlsx --config files/reference/health_config.yml

Reads the catalogue state exported per dataset and the health configuration you set (which
attributes are required, whether machine drafts count, what "fresh" means), and prints the
coverage of each attribute, the datasets that are complete, the stale harvests and the
failed runs.
"""
from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import yaml

from common import FILES, info, read_sheet

ATTR = {
    "description_reviewed": "description",
    "owner": "owner",
    "business_steward": "business_steward",
    "it_steward": "it_steward",
    "classification": "classification",
    "any_column_linked_to_term": "columns_linked_to_terms",
}


def has(row: dict, key: str, cfg: dict) -> bool:
    if key == "description_reviewed":
        ok = {"reviewed", "machine_draft"} if cfg.get("count_machine_drafts_as_described") else {"reviewed"}
        return row["description_status"] in ok
    if key == "any_column_linked_to_term":
        return int(float(row["columns_linked_to_terms"] or 0)) > 0
    return bool(row[ATTR[key]])


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--state", default=str(FILES / "reference" / "catalogue_state.xlsx"))
    ap.add_argument("--config", default=str(FILES / "reference" / "health_config.yml"))
    a = ap.parse_args()
    cfg = yaml.safe_load(Path(a.config).read_text(encoding="utf-8"))
    rows = read_sheet(Path(a.state), "datasets")
    runs = read_sheet(Path(a.state), "harvest_runs")
    as_of = date.fromisoformat(str(cfg.get("as_of", "2026-10-31")))
    n = len(rows)
    print(f"datasets in the domain: {n}\n\ncoverage by attribute:")
    for key in ATTR:
        c = sum(has(r, key, cfg) for r in rows)
        req = " (required)" if cfg["required"].get(key) else ""
        info(f"{key:28} {c:3}/{n} = {c / n:6.1%}{req}")
    cols = sum(int(float(r["columns"])) for r in rows)
    linked = sum(int(float(r["columns_linked_to_terms"] or 0)) for r in rows)
    info(f"{'columns linked to a term':28} {linked:3}/{cols} = {linked / cols:6.1%}")
    req = [k for k, v in cfg["required"].items() if v]
    complete = [r for r in rows if all(has(r, k, cfg) for k in req)]
    print(f"\ncomplete (all required attributes): {len(complete)}/{n} = {len(complete) / n:.1%}")
    for r in rows:
        missing = [k for k in req if not has(r, k, cfg)]
        if missing:
            info(f"incomplete: {r['dataset']:44} missing {', '.join(missing)}")
    k = float(cfg.get("fresh_within_intervals", 2))
    stale = []
    for r in rows:
        age = (as_of - date.fromisoformat(r["last_successful_harvest"])).days
        if age > k * float(r["schedule_days"]):
            stale.append((r["dataset"], age))
    print(f"\nfresh (last good harvest within {k:g} schedule intervals): {n - len(stale)}/{n} = {(n - len(stale)) / n:.1%}")
    for ds, age in stale:
        info(f"stale: {ds:44} last good harvest {age} days ago")
    failed = [r for r in runs if r["result"] != "success"]
    print(f"\nharvest runs: {len(runs)} · failed {len(failed)}")
    for r in failed:
        info(f"failed: {r['pipeline']} on {r['run_date']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
