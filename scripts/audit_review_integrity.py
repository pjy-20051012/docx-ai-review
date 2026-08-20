# -*- coding: utf-8 -*-
"""Audit a tracked-review DOCX against its clean baseline.

This is intentionally independent of python-docx: it audits the OOXML package
before Word has a chance to accept, normalize, or hide revisions.

Usage:
  python audit_review_integrity.py baseline.docx revised.docx \
      --target-index 12 --target-index 42
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def W(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _root(zf: zipfile.ZipFile, name: str):
    return etree.fromstring(zf.read(name))


def _paragraphs(root):
    return root.xpath("//w:body//w:p", namespaces=NS)


def _visible_text(el, state: str) -> str:
    """Return paragraph text in the original or accepted revision state."""
    tag = etree.QName(el).localname
    if tag == "ins" and state == "original":
        return ""
    if tag == "del" and state == "accepted":
        return ""
    if tag in {"t", "delText"}:
        return el.text or ""
    if tag in {"instrText", "fldChar"}:
        return ""
    if tag == "tab":
        return "\t"
    if tag in {"br", "cr"}:
        return "\n"
    return "".join(_visible_text(child, state) for child in el)


def _paragraph_texts(root, state: str) -> list[str]:
    return [_visible_text(p, state) for p in _paragraphs(root)]


def _revision_counts(root) -> Counter:
    return Counter(
        etree.QName(el).localname
        for el in root.xpath("//w:ins | //w:del", namespaces=NS)
    )


def _nested_revisions(root) -> int:
    count = 0
    for el in root.xpath("//w:ins | //w:del", namespaces=NS):
        if any(
            etree.QName(parent).localname in {"ins", "del"}
            for parent in el.iterancestors()
        ):
            count += 1
    return count


def _field_signature(root):
    return {
        "instructions": [
            el.text or "" for el in root.xpath("//w:instrText", namespaces=NS)
        ],
        "field_chars": [
            el.get(W("fldCharType"), "")
            for el in root.xpath("//w:fldChar", namespaces=NS)
        ],
    }


def _bookmark_signature(root):
    return sorted(
        (el.get(W("id"), ""), el.get(W("name"), ""))
        for el in root.xpath("//w:bookmarkStart", namespaces=NS)
    )


def _comment_signature(zf: zipfile.ZipFile, root):
    entities: set[str] = set()
    if "word/comments.xml" in zf.namelist():
        comments = _root(zf, "word/comments.xml")
        entities = {
            el.get(W("id"))
            for el in comments.xpath("//w:comment", namespaces=NS)
            if el.get(W("id")) is not None
        }
    markers = Counter(
        el.get(W("id"))
        for el in root.xpath(
            "//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference",
            namespaces=NS,
        )
        if el.get(W("id")) is not None
    )
    missing = sorted(cid for cid in markers if cid not in entities)
    return {
        "entity_count": len(entities),
        "marker_count": sum(markers.values()),
        "markers_by_id": dict(sorted(markers.items())),
        "missing_entities": missing,
    }


def audit(
    baseline: Path,
    revised: Path,
    targets: set[int],
    allow_baseline_revisions: bool,
    allow_non_target_revisions: bool,
) -> tuple[list[str], list[str], dict]:
    errors: list[str] = []
    warnings: list[str] = []
    with zipfile.ZipFile(baseline) as base_zip, zipfile.ZipFile(revised) as out_zip:
        required = "word/document.xml"
        if required not in base_zip.namelist() or required not in out_zip.namelist():
            return [f"missing required package part: {required}"], warnings, {}
        base_root = _root(base_zip, required)
        out_root = _root(out_zip, required)

        base_revs = _revision_counts(base_root)
        out_revs = _revision_counts(out_root)
        if base_revs and not allow_baseline_revisions:
            errors.append(
                "baseline is not clean: found "
                + ", ".join(f"w:{k}={v}" for k, v in sorted(base_revs.items()))
                + "; start from a clean/original DOCX"
            )
        if _nested_revisions(out_root):
            errors.append("revised document contains nested w:ins/w:del revisions")

        base_paras = _paragraphs(base_root)
        out_paras = _paragraphs(out_root)
        if len(base_paras) != len(out_paras):
            errors.append(
                f"body paragraph count changed: baseline={len(base_paras)}, revised={len(out_paras)}"
            )
        else:
            base_original = _paragraph_texts(base_root, "original")
            out_original = _paragraph_texts(out_root, "original")
            if base_original != out_original:
                differing = [
                    i for i, (a, b) in enumerate(zip(base_original, out_original)) if a != b
                ]
                errors.append(
                    "revised original-state text differs from baseline at paragraph(s): "
                    + ", ".join(map(str, differing[:20]))
                )
            if targets and not allow_non_target_revisions:
                for i, paragraph in enumerate(out_paras):
                    if i not in targets and paragraph.xpath(
                        ".//w:ins | .//w:del", namespaces=NS
                    ):
                        errors.append(f"non-target paragraph {i} contains tracked revisions")

        base_fields = _field_signature(base_root)
        out_fields = _field_signature(out_root)
        if base_fields != out_fields:
            errors.append(
                "native field structure changed: instrText/fldChar signature differs from baseline"
            )
        base_bookmarks = _bookmark_signature(base_root)
        out_bookmarks = _bookmark_signature(out_root)
        if base_bookmarks != out_bookmarks:
            errors.append(
                "bookmark structure changed: cross-reference/Zotero bookmark signature differs from baseline"
            )

        out_comments = _comment_signature(out_zip, out_root)
        if out_comments["missing_entities"]:
            errors.append(
                "comment markers reference missing comment entities: "
                + ", ".join(out_comments["missing_entities"])
            )
        if out_revs and not targets:
            warnings.append(
                "tracked revisions found but no --target-index was supplied; "
                "non-target isolation was not checked"
            )
        if not out_revs:
            warnings.append("revised document contains no w:ins/w:del revisions")

        details = {
            "baseline_paragraphs": len(base_paras),
            "revised_paragraphs": len(out_paras),
            "baseline_revisions": dict(base_revs),
            "revised_revisions": dict(out_revs),
            "field_signature_equal": base_fields == out_fields,
            "bookmark_signature_equal": base_bookmarks == out_bookmarks,
            "revised_comments": out_comments,
        }
    return errors, warnings, details


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Audit a tracked-review DOCX against a clean baseline"
    )
    parser.add_argument("baseline", type=Path)
    parser.add_argument("revised", type=Path)
    parser.add_argument(
        "--target-index",
        type=int,
        action="append",
        default=[],
        help="0-based paragraph containing an intended revision; repeatable",
    )
    parser.add_argument(
        "--allow-baseline-revisions",
        action="store_true",
        help="diagnostic escape hatch; do not use for a new review pass",
    )
    parser.add_argument("--allow-non-target-revisions", action="store_true")
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args(argv)
    try:
        errors, warnings, details = audit(
            args.baseline,
            args.revised,
            set(args.target_index),
            args.allow_baseline_revisions,
            args.allow_non_target_revisions,
        )
    except (OSError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        errors, warnings, details = [f"cannot audit DOCX package: {exc}"], [], {}
    payload = {
        "status": "failed" if errors else "ok",
        "errors": errors,
        "warnings": warnings,
        "details": details,
    }
    if args.as_json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print("AUDIT FAILED" if errors else "AUDIT OK")
        for message in errors:
            print(" - ERROR:", message)
        for message in warnings:
            print(" - WARNING:", message)
        if details:
            print(json.dumps(details, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
