# Clean baselines, native fields, and formatting-only fixes

This reference records the safety rules for a Word review pass. It is especially
important for manuscripts containing Zotero citations, cross-references, and
pre-existing tracked changes.

## One review pass must have one clean baseline

Before running `tracked` or `rewrite`, identify the actual original DOCX. A file
that has already had revisions accepted, rejected, or partially left in place is
not an equivalent baseline: its visible text can hide `w:ins`/`w:del`, and a new
pass may create a mixed revision chain or duplicate an earlier correction.

The CLI therefore refuses an input containing existing `w:ins` or `w:del` by
default. If the source is not clean, stop and rebase the edit list against the
clean original. `--allow-existing-revisions` is only a deliberate diagnostic
escape hatch; it is not a safe production workflow.

After generation, compare the revised document's **original state** with the
clean baseline: normal text plus deleted text, excluding inserted text. The
accepted state is a separate view: normal text plus inserted text, excluding
deleted text. Never compare only the text Word currently displays.

## Preserve Zotero and cross-reference fields

Native citation fields are part of the document's OOXML, not ordinary prose.
Do not reconstruct or clear a paragraph that contains any of the following:

- `w:instrText`, `w:fldChar`, `w:fldSimple`;
- bookmark start/end nodes used by cross-references;
- Zotero citation fields or their field-result runs.

Edits should be anchored to ordinary text runs and should preserve the field
signature and bookmarks at the structural level. Never replace a Zotero field
with typed citation text. If Word and Zotero are available, refresh the fields
through the installed Zotero Word integration after the review pass; then
re-audit field codes and bookmarks. If Word/Zotero is unavailable, leave the
native fields intact and report that refresh remains pending.

## Sentence and word granularity

For a clear before/after pair, pair sentences first and then apply word-level
diffs inside a matched sentence. Use a full sentence deletion + insertion only
when the sentence is genuinely different or cannot be paired reliably. Do not
replace an entire paragraph merely because a paragraph-level rewrite was
provided as a reference.

## Superscript, subscript, and units

Inspect `w:vertAlign` before treating a token as a textual error. Unit exponents
and chemical notation are frequently already formatted as superscript or
subscript even when their plain-text extraction looks unusual. The markdown
refiner must skip protected vertical-alignment spans.

If the user explicitly asks for a pure formatting correction, apply it as a
direct formatting edit outside tracked linguistic changes, and never run a
broad regex over revision runs, field runs, or citation results. Re-run the
field/revision audit after such a direct edit.

## Required audit

```bash
python scripts/ai_review_to_comments.py verify revised.docx
python scripts/audit_review_integrity.py clean-original.docx revised.docx \
  --target-index 30 --target-index 42
```

The integrity audit checks:

- a clean baseline and paragraph-count stability;
- preservation of the baseline original-state text;
- no revisions in non-target paragraphs;
- no nested revisions;
- native field-code and bookmark signatures;
- comment markers resolving to comment entities.
