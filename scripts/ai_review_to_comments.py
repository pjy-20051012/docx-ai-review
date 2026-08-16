# -*- coding: utf-8 -*-
"""Inject AI review comments into a .docx without changing the body text.

Commands:
  python ai_review_to_comments.py dump-text input.docx [--output text.json]
  python ai_review_to_comments.py apply input.docx reviews.json -o output.docx [--verify-only]
"""

from __future__ import annotations

import argparse
import copy
import json
import sys
import zipfile
from bisect import bisect_left
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.oxml.ns import qn
from docx.text.run import Run
from lxml import etree
from pydantic import BaseModel, Field, ValidationError

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CATEGORIES = ("学术规范", "语法修正", "逻辑表达", "用词精炼", "句式润色")
COMMENT_AUTHOR = "AI Reviewer"
COMMENT_INITIALS = "AI"


class ReviewItem(BaseModel):
    paragraph_index: int = Field(..., ge=0, description="0-based paragraph index")
    anchor_text: str = Field(..., min_length=1, description="exact substring in the paragraph")
    category: str = Field(..., description="one of " + "、".join(CATEGORIES))
    problem_analysis: str = Field(..., min_length=1, description="concise diagnosis, <= 50 chars")
    suggested_revision: str = Field(..., min_length=1, description="local revision for the anchored text")
    occurrence: int = Field(1, ge=1, description="1-based occurrence when the phrase repeats")
    paragraph_revision: str = Field(
        "",
        description="optional full rewritten paragraph text supplied by the reviewing model",
    )

    def model_post_init(self, __context) -> None:
        if self.category not in CATEGORIES:
            raise ValueError(f"category must be one of {CATEGORIES}, got {self.category!r}")
        if not self.anchor_text.strip():
            raise ValueError("anchor_text must contain non-whitespace characters")


class ReviewBatch(BaseModel):
    reviews: List[ReviewItem]


def _paragraph_text(paragraph) -> str:
    return paragraph.text


def _container_text(container) -> str:
    parts = []
    for child in container:
        tag = etree.QName(child.tag).localname
        if tag == "t":
            parts.append(child.text or "")
        elif tag in ("tab", "br", "cr", "noBreakHyphen"):
            parts.append("\t" if tag == "tab" else "\n")
    return "".join(parts)


def _char_map(paragraph_el):
    """Return (containers, cumulative_text_lengths) aligned with Paragraph.text."""
    containers = []
    lengths = []
    total = 0
    for child in paragraph_el:
        tag = etree.QName(child.tag).localname
        if tag in ("r", "hyperlink"):
            containers.append(child)
            total += len(_container_text(child))
            lengths.append(total)
    return containers, lengths


def _covering_index(offsets, pos: int) -> int:
    idx = bisect_left(offsets, pos + 1)
    if idx >= len(offsets):
        return len(offsets) - 1
    return idx


def _split_container(container, local_pos: int, paragraph_el) -> None:
    """Split container text at local_pos, inserting a deep-copied suffix container."""
    text = _container_text(container)
    if local_pos <= 0 or local_pos >= len(text):
        return
    suffix = copy.deepcopy(container)
    content = [c for c in suffix if etree.QName(c.tag).localname != "rPr"]
    seen = 0
    split_child = None
    prefix_children = []
    for child in content:
        tag = etree.QName(child.tag).localname
        if tag == "t":
            child_len = len(child.text or "")
        elif tag in ("tab", "br", "cr", "noBreakHyphen"):
            child_len = 1
        else:
            child_len = 0
        if child_len == 0:
            continue
        if local_pos < seen + child_len:
            if tag != "t":
                raise ValueError(
                    "anchor boundary falls inside a tab/line-break; choose an anchor that starts/ends on text"
                )
            rel = local_pos - seen
            if rel > 0:
                child.text = (child.text or "")[:rel]
                split_child = child
                suffix_text = copy.deepcopy(child)
                suffix_text.text = (copy.deepcopy(child).text or "")[rel:]
                child.addnext(suffix_text)
            else:
                prefix_children.append(child)
            break
        if seen + child_len == local_pos:
            prefix_children.append(child)
            break
        seen += child_len
        prefix_children.append(child)
    # The suffix keeps everything strictly after local_pos. Remove prefix elements.
    for child in prefix_children:
        if child.getparent() is suffix:
            suffix.remove(child)
    if split_child is not None and split_child.getparent() is suffix:
        suffix.remove(split_child)
    container.addnext(suffix)


def _split_at_global(paragraph_el, global_pos: int):
    containers, offsets = _char_map(paragraph_el)
    idx = _covering_index(offsets, global_pos)
    container = containers[idx]
    prev = offsets[idx - 1] if idx > 0 else 0
    local = global_pos - prev
    _split_container(container, local, paragraph_el)


def _isolate_runs(paragraph, start: int, end: int) -> List[Run]:
    """Split runs so [start, end) maps to exact Run objects within one paragraph."""
    if end <= start:
        raise ValueError(f"anchor range must be non-empty, got [{start}, {end})")
    p_el = paragraph._p
    _split_at_global(p_el, start)
    _split_at_global(p_el, end)
    containers, offsets = _char_map(p_el)
    si = _covering_index(offsets, start)
    ei = _covering_index(offsets, end - 1)
    runs = []
    for i in range(si, ei + 1):
        container = containers[i]
        if etree.QName(container.tag).localname == "r":
            runs.append(Run(container, paragraph))
        else:
            raise ValueError(
                "review anchor crosses a hyperlink boundary; choose a shorter anchor_text"
            )
    if not runs:
        raise ValueError("anchor container has no runs")
    return runs


def apply_reviews(doc: Document, reviews: List[ReviewItem]) -> int:
    paragraphs = doc.paragraphs
    count = 0
    for item in reviews:
        if item.paragraph_index >= len(paragraphs):
            raise IndexError(
                f"paragraph_index {item.paragraph_index} out of range (document has {len(paragraphs)} paragraphs)"
            )
        paragraph = paragraphs[item.paragraph_index]
        text = _paragraph_text(paragraph)
        anchor = item.anchor_text
        found = -1
        for _ in range(item.occurrence):
            found = text.find(anchor, found + 1)
            if found == -1:
                break
        if found == -1:
            raise ValueError(
                f"anchor_text {anchor!r} occurrence {item.occurrence} not found in paragraph {item.paragraph_index}"
            )
        runs = _isolate_runs(paragraph, found, found + len(anchor))
        comment_text = (
            f"【{item.category}】诊断：{item.problem_analysis}\n建议修改：{item.suggested_revision}"
        )
        if item.paragraph_revision:
            comment_text += "\n\n整段修改参考：\n" + item.paragraph_revision
        doc.add_comment(
            runs=runs,
            text=comment_text,
            author=COMMENT_AUTHOR,
            initials=COMMENT_INITIALS,
        )
        count += 1
    return count


class PolishEdit(BaseModel):
    """One atomic revision produced by a paper-polishing skill."""

    paragraph_index: int = Field(..., ge=0)
    original_text: str = Field(..., min_length=1, description="verbatim text from the manuscript")
    revision_type: str = Field(
        ...,
        description="one of 学术规范 / 语法修正 / 逻辑表达 / 用词精炼 / 句式润色",
    )
    reason: str = Field(..., min_length=1, description="concise reason, <= 60 chars")
    revised_text: str = Field(..., description="replacement for the anchored span")
    occurrence: int = Field(1, ge=1)
    whole_paragraph_revision: str = Field(
        "",
        description="optional full rewritten paragraph when the polish skill also rewrote it",
    )

    def model_post_init(self, __context) -> None:
        if self.revision_type not in CATEGORIES:
            raise ValueError(
                f"revision_type must be one of {CATEGORIES}, got {self.revision_type!r}"
            )


class PolishBatch(BaseModel):
    document: str = Field("", description="title or filename of the manuscript")
    journal_guidelines: str = Field("", description="journal name or short style notes")
    edits: List[PolishEdit]


def convert_polish_edits_to_reviews(paragraphs, edits: List[PolishEdit]) -> List[ReviewItem]:
    """Turn a polish skill's structured edit list into comment review items."""
    reviews: List[ReviewItem] = []
    for e in edits:
        if e.paragraph_index >= len(paragraphs):
            raise IndexError(
                f"paragraph_index {e.paragraph_index} out of range (document has {len(paragraphs)} paragraphs)"
            )
        text = paragraphs[e.paragraph_index].text
        if e.original_text not in text:
            raise ValueError(
                f"original_text {e.original_text!r} not found in paragraph {e.paragraph_index}"
            )
        reviews.append(
            ReviewItem(
                paragraph_index=e.paragraph_index,
                anchor_text=e.original_text,
                category=e.revision_type,
                problem_analysis=e.reason,
                suggested_revision=e.revised_text,
                occurrence=e.occurrence,
                paragraph_revision=e.whole_paragraph_revision,
            )
        )
    return reviews


def verify_comment_integrity(docx_path: Path) -> List[str]:
    problems: List[str] = []
    with zipfile.ZipFile(docx_path) as zf:
        names = zf.namelist()
        if "word/comments.xml" not in names:
            return ["missing word/comments.xml"]
        doc_xml = zf.read("word/document.xml")
        comments_xml = zf.read("word/comments.xml")
    doc_root = etree.fromstring(doc_xml)
    comments_root = etree.fromstring(comments_xml)
    ns = {"w": W_NS}
    comment_ids = {
        el.get(qn("w:id"))
        for el in comments_root.xpath("//w:comment", namespaces=ns)
        if el.get(qn("w:id")) is not None
    }
    refs = doc_root.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=ns)
    if not refs:
        problems.append("no comment markers in document.xml")
    for el in refs:
        cid = el.get(qn("w:id"))
        if cid not in comment_ids:
            problems.append(f"missing comment entity for id={cid} ({etree.QName(el.tag).localname})")
    return problems


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="AI review to Word comments")
    sub = parser.add_subparsers(dest="command", required=True)

    p_dump = sub.add_parser("dump-text", help="dump paragraphs with indices")
    p_dump.add_argument("input", type=Path)
    p_dump.add_argument("--output", type=Path, default=None)

    p_apply = sub.add_parser("apply", help="apply reviews JSON to a docx")
    p_apply.add_argument("input", type=Path)
    p_apply.add_argument("reviews", type=Path)
    p_apply.add_argument("-o", "--output", type=Path, default=None)
    p_apply.add_argument("--verify-only", action="store_true")

    p_convert = sub.add_parser(
        "convert",
        help="convert a paper-polishing skill's structured edit list into comments",
    )
    p_convert.add_argument("input", type=Path)
    p_convert.add_argument("polish_edits", type=Path)
    p_convert.add_argument("-o", "--output", type=Path, required=True)

    p_verify = sub.add_parser("verify", help="verify comment markers in an annotated docx")
    p_verify.add_argument("input", type=Path)

    args = parser.parse_args(argv)

    if args.command == "dump-text":
        doc = Document(str(args.input))
        entries = [
            {"paragraph_index": i, "text": p.text, "char_count": len(p.text)}
            for i, p in enumerate(doc.paragraphs)
        ]
        payload = {"document": str(args.input), "paragraphs": entries}
        output = json.dumps(payload, ensure_ascii=False, indent=2)
        if args.output:
            args.output.write_text(output, encoding="utf-8")
        else:
            print(output)
        return 0

    if args.command == "apply":
        if args.verify_only:
            problems = verify_comment_integrity(args.input)
            if problems:
                print("VERIFY FAILED")
                for p in problems:
                    print(" -", p)
                return 1
            print(f"VERIFY OK: {args.input}")
            return 0
        if not args.output:
            parser.error("apply requires -o/--output")
        raw = json.loads(args.reviews.read_text(encoding="utf-8"))
        try:
            batch = ReviewBatch.model_validate(raw)
        except ValidationError as exc:
            print("Invalid reviews JSON:")
            print(exc)
            return 2
        doc = Document(str(args.input))
        count = apply_reviews(doc, batch.reviews)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(args.output))
        problems = verify_comment_integrity(args.output)
        if problems:
            print("VERIFY FAILED after save")
            for p in problems:
                print(" -", p)
            return 3
        reopened = Document(str(args.output))
        print(
            f"OK: injected {count} comment(s); saved {args.output}; "
            f"reopened with {len(reopened.comments)} comment(s)."
        )
        return 0

    if args.command == "convert":
        raw = json.loads(args.polish_edits.read_text(encoding="utf-8"))
        try:
            batch = PolishBatch.model_validate(raw)
        except ValidationError as exc:
            print("Invalid polish edits JSON:")
            print(exc)
            return 2
        doc = Document(str(args.input))
        reviews = convert_polish_edits_to_reviews(doc.paragraphs, batch.edits)
        count = apply_reviews(doc, reviews)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(args.output))
        problems = verify_comment_integrity(args.output)
        if problems:
            print("VERIFY FAILED after save")
            for p in problems:
                print(" -", p)
            return 3
        print(
            f"OK: converted {len(reviews)} polish edit(s) into {count} comment(s); saved {args.output}"
        )
        return 0

    if args.command == "verify":
        problems = verify_comment_integrity(args.input)
        if problems:
            print("VERIFY FAILED")
            for p in problems:
                print(" -", p)
            return 1
        print(f"VERIFY OK: {args.input}")
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())