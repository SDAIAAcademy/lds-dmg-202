"""Lab 2.1 · Walk column-level lineage downstream from one column, using OpenLineage events.

    uv run kit/impact.py --column licensing.licences.status
    uv run kit/impact.py --column licensing.licences.expiry_date --as-of 2026-10-31
    uv run kit/impact.py --column accreditation.cpd_records.hours         --events files/capstone/accreditation_events.jsonl --owners files/capstone/accreditation_owners.csv

Reads files/lineage/events.jsonl (OpenLineage run events with the columnLineage facet) and
files/lineage/owners.csv (the follow-up task passes the accreditation domain's two files). Prints every dataset and column that depends on the column, how
(DIRECT: the value is copied or computed from it; INDIRECT: it filters or groups the rows), the
job that carries it, the job's last run and whether that run is stale, and the owner.
"""
from __future__ import annotations

import argparse
import csv
import json
from collections import defaultdict, deque
from datetime import date

from common import FILES, info

ALIAS = {"licensing": "sqlite://licensing", "warehouse": "sqlite://warehouse", "files": "file://filestore", "bi": "bi://dashboards",
         "accreditation": "sqlite://accreditation"}


def label(ns: str, name: str) -> str:
    rev = {v: k for k, v in ALIAS.items()}
    return f"{rev.get(ns, ns)}.{name}"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--column", required=True, help="namespace.dataset.column, e.g. licensing.licences.status")
    ap.add_argument("--as-of", default="2026-10-31")
    ap.add_argument("--events", default=str(FILES / "lineage" / "events.jsonl"))
    ap.add_argument("--owners", default=str(FILES / "lineage" / "owners.csv"))
    a = ap.parse_args()
    ns_alias, rest = a.column.split(".", 1)
    dataset, column = rest.rsplit(".", 1)
    start = (ALIAS.get(ns_alias, ns_alias), dataset, column)
    as_of = date.fromisoformat(a.as_of)

    edges = defaultdict(set)  # (ns, ds, col) -> {(ns, ds, col or '*', kind, job)}
    runs = defaultdict(list)
    outputs_cols = defaultdict(set)
    for line in open(a.events, encoding="utf-8"):
        e = json.loads(line)
        job = e["job"]["name"]
        if e["eventType"] != "COMPLETE":
            continue
        runs[job].append(date.fromisoformat(e["eventTime"][:10]))
        for out in e.get("outputs", []):
            cl = (out.get("facets") or {}).get("columnLineage") or {}
            tgt = (out["namespace"], out["name"])
            for col, spec in cl.get("fields", {}).items():
                outputs_cols[tgt].add(col)
                for f in spec.get("inputFields", []):
                    t = (f.get("transformations") or [{}])[0]
                    edges[(f["namespace"], f["name"], f["field"])].add((tgt[0], tgt[1], col, f"{t.get('type', 'DIRECT')}/{t.get('subtype', '')}", job))
            for f in cl.get("dataset", []):
                t = (f.get("transformations") or [{}])[0]
                edges[(f["namespace"], f["name"], f["field"])].add((tgt[0], tgt[1], "*", f"{t.get('type', 'INDIRECT')}/{t.get('subtype', '')}", job))

    owners = {}
    with open(a.owners, encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            owners[(r["namespace"], r["name"])] = (r["owner"] or "(no owner)", r["steward"] or "(no steward)")

    def job_state(job: str) -> str:
        r = sorted(runs[job])
        last = r[-1]
        gap = (r[-1] - r[-2]).days if len(r) > 1 else 1
        stale = (as_of - last).days > 2 * max(gap, 1)
        return f"last run {last}{' · STALE' if stale else ''}"

    print(f"downstream of {a.column} (as of {as_of}):\n")
    seen, queue, hits = {start}, deque([(start, 0)]), defaultdict(list)
    while queue:
        node, depth = queue.popleft()
        for (ns, ds, col, kind, job) in sorted(edges.get(node, ())):
            hits[(ns, ds)].append((col, kind, job, depth + 1, f"{node[1]}.{node[2]}"))
            nxt = [(ns, ds, c) for c in (outputs_cols[(ns, ds)] if col == "*" else [col])]
            for n in nxt:
                # an indirect dependency reaches every column of the output, but only DIRECT paths carry the value on
                if n not in seen:
                    seen.add(n)
                    queue.append((n, depth + 1))
    for (ns, ds), items in sorted(hits.items(), key=lambda x: min(i[3] for i in x[1])):
        own, stw = owners.get((ns, ds), ("(not in owners.csv)", ""))
        print(f"{label(ns, ds)}   hop {min(i[3] for i in items)} · owner {own} · steward {stw}")
        for col, kind, job, depth, src in sorted(set(items)):
            info(f"{'(all rows)' if col == '*' else col:22} {kind:22} from {src:28} job {job} · {job_state(job)}")
    print()
    print(f"datasets affected: {len(hits)} · jobs: {len({i[2] for v in hits.values() for i in v})}")
    print("not in the events: manual extracts, spreadsheets, scripts on a laptop. Ask the owners.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
