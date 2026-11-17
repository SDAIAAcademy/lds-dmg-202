"""Lab 2.3 · Turn dataset records into DCAT and validate them against the target portal's profile.

    uv run kit/dcat.py
    uv run kit/dcat.py --workbook files/capstone/accreditation_records.xlsx

Reads the records workbook, writes out/records.ttl (W3C DCAT in Turtle) and validates it with
SHACL against files/records/target_profile.ttl, the required fields of the case's target portal.
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

from rdflib import BNode, Graph, Literal, Namespace, URIRef
from rdflib.namespace import DCAT, DCTERMS, RDF, XSD

from common import FILES, OUT, finish, fix, info, ok, read_sheet

VCARD = Namespace("http://www.w3.org/2006/vcard/ns#")
TP = Namespace("https://data.example.org/target-profile#")
FREQ = Namespace("http://publications.europa.eu/resource/authority/frequency/")
BASE = "https://data.example.org/sinwan/dataset/"
PUBLISHER = URIRef("https://data.example.org/sinwan/")
LICENCES = {"CC-BY-4.0": "https://creativecommons.org/licenses/by/4.0/"}


def build(rows: list[dict]) -> Graph:
    g = Graph()
    g.bind("dcat", DCAT)
    g.bind("dct", DCTERMS)
    g.bind("vcard", VCARD)
    for r in rows:
        ds = URIRef(BASE + (r.get("identifier") or "missing"))
        g.add((ds, RDF.type, DCAT.Dataset))
        if r.get("identifier"):
            g.add((ds, DCTERMS.identifier, Literal(r["identifier"])))
        if r.get("title_ar"):
            g.add((ds, DCTERMS.title, Literal(r["title_ar"], lang="ar")))
        if r.get("title_en"):
            g.add((ds, DCTERMS.title, Literal(r["title_en"], lang="en")))
        if r.get("description_ar"):
            g.add((ds, DCTERMS.description, Literal(r["description_ar"], lang="ar")))
        g.add((ds, DCTERMS.publisher, PUBLISHER))
        cp = BNode()
        g.add((ds, DCAT.contactPoint, cp))
        g.add((cp, RDF.type, VCARD.Kind))
        if r.get("contact_name"):
            g.add((cp, VCARD.fn, Literal(r["contact_name"], lang="ar")))
        if r.get("contact_email"):
            g.add((cp, VCARD.hasEmail, URIRef("mailto:" + r["contact_email"])))
        lic = r.get("licence", "")
        g.add((ds, DCTERMS.license, URIRef(LICENCES[lic]) if lic in LICENCES else Literal(lic)))
        for kw in [k.strip() for k in re.split(r"[،,]", r.get("keywords", "")) if k.strip()]:
            g.add((ds, DCAT.keyword, Literal(kw, lang="ar")))
        if r.get("theme"):
            g.add((ds, DCAT.theme, TP[r["theme"]] if re.fullmatch(r"[A-Z]+", r["theme"]) else Literal(r["theme"])))
        per = r.get("periodicity", "")
        if per:
            g.add((ds, DCTERMS.accrualPeriodicity, FREQ[per] if re.fullmatch(r"[A-Z]+", per) else Literal(per)))
        mod = r.get("modified", "")[:10]
        if mod:
            g.add((ds, DCTERMS.modified, Literal(mod, datatype=XSD.date) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", mod) else Literal(mod)))
        dist = URIRef(str(ds) + "/csv")
        g.add((ds, DCAT.distribution, dist))
        g.add((dist, RDF.type, DCAT.Distribution))
        if r.get("download_url"):
            g.add((dist, DCAT.downloadURL, URIRef(r["download_url"])))
        mt = r.get("media_type", "")
        if mt:
            g.add((dist, DCAT.mediaType, URIRef("https://www.iana.org/assignments/media-types/" + mt) if "/" in mt else Literal(mt)))
    return g


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=str(FILES / "records" / "dataset_records.xlsx"))
    ap.add_argument("--profile", default=str(FILES / "records" / "target_profile.ttl"))
    a = ap.parse_args()
    from pyshacl import validate

    rows = read_sheet(Path(a.workbook), "records")
    g = build(rows)
    OUT.mkdir(exist_ok=True)
    out = OUT / "records.ttl"
    g.serialize(out, format="turtle", encoding="utf-8")
    info(f"wrote {out.relative_to(OUT.parent)}: {len(rows)} records, {len(g)} statements")
    shapes = Graph().parse(a.profile, format="turtle")
    conforms, report, _text = validate(g, shacl_graph=shapes, inference="none", allow_infos=True, allow_warnings=True)
    SH = Namespace("http://www.w3.org/ns/shacl#")
    problems: dict[str, list[str]] = {}
    for res in report.subjects(RDF.type, SH.ValidationResult):
        focus = str(report.value(res, SH.focusNode))
        msg = str(report.value(res, SH.resultMessage) or report.value(res, SH.resultPath))
        ident = focus.replace(BASE, "").split("/")[0]
        problems.setdefault(ident, []).append(msg)
    for r in rows:
        ident = r.get("identifier") or "missing"
        if ident in problems:
            for m in sorted(set(problems[ident])):
                fix(f"{ident}: {m}")
        else:
            ok(f"{ident}: passes the target portal's validator")
    if len(rows) != 5:
        fix(f"{len(rows)} records; five are required")
    return finish("الدليل: خمسة سجلات بالترخيص وجهة الاتصال تجتاز أداة التحقق في البوابة المستهدفة")


if __name__ == "__main__":
    raise SystemExit(main())
