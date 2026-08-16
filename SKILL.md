---
name: docx-ai-review
description: "Use when the user wants AI review or polish feedback applied to a Word document as native Word comments instead of a rewritten draft. Triggers include: AI-modified manuscript to comments, teacher-style review comments on a .docx, character-level run splitting for comment anchoring, and workflows that must preserve the original document body while adding precise inline review notes."
---

# DOCX AI Review Comments

Turn an AI-generated revision into native Word comments anchored to the exact text spans, without changing the original body. The pipeline is: extract paragraph text -> ask the model for structured reviews -> inject comments with character-level anchors.

## Prerequisites

- python-docx >= 1.2.0, lxml, pydantic >= 2.0.
- Check quickly: `python -c "import docx, lxml, pydantic; print(docx.__version__)"`

## Workflow

1. Extract text so the model can see exact anchors:

```bash
python scripts/ai_review_to_comments.py dump-text input.docx --output text.json
```

2. Read `references/review_prompt_template.md`, pass the document context (or the relevant paragraphs from `text.json`) to the model, and ask it to return the JSON review batch defined in that template.

3. Validate the returned JSON against the schema by saving it as `reviews.json` and applying:

```bash
python scripts/ai_review_to_comments.py apply input.docx reviews.json -o annotated.docx
```

4. Inspect the result. The script reopens the file and verifies that every `commentRangeStart`/`commentRangeEnd`/`commentReference` id has a matching `comments.xml` entry. Optionally render to PDF and inspect visually.

## Tracked-changes workflow (recommended for clear before/after pairs)

When a review file explicitly provides an original span and a replacement (for example `alendaring` -> `calendaring`, or `W·m-1·K-1` -> `W·m⁻¹·K⁻¹`), use Word's review mode instead of a comment-only suggestion:

```bash
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx
```

This applies each edit as a real tracked change (`w:ins` / `w:del`) so Word shows the before/after side by side in the Review tab. Each change also carries a comment with the reason, so the author can confirm each revision and accept or reject it.

## Combined use with paper-polishing skills

When another paper-polishing skill supplies the revision content, require it to emit the structured edit list defined in `references/polish_skill_contract.md`, then convert directly:

```bash
python scripts/ai_review_to_comments.py convert input.docx polish_edits.json -o annotated.docx
```

The contract enforces verbatim anchors, atomic edits, paragraph-scoped changes, and optional whole-paragraph rewrites, so the polishing skill's judgments become granular, traceable comments.

## Rules enforced by the scripts

- Review items must contain `paragraph_index`, exact `anchor_text`, `category` (学术规范 / 语法修正 / 逻辑表达 / 用词精炼 / 句式润色), concise `problem_analysis`, and `suggested_revision`. `occurrence` disambiguates repeated phrases; optional `paragraph_revision` carries the model's full rewritten paragraph and is appended to the comment as reference.
- Anchors must be a substring of a single paragraph. The run isolator splits `w:r` elements at character offsets with `copy.deepcopy` of formatting and `xml:space` preservation, then hands the exact runs to `Document.add_comment()`. An anchor may span multiple runs, but it must not cross a hyperlink boundary.
- Comment text format: `【{category}】诊断：{problem_analysis}\n建议修改：{suggested_revision}`; author is `AI Reviewer`, initials `AI`.

## Troubleshooting

- "anchor crosses a hyperlink boundary": split the review into shorter anchors that stay inside the hyperlink or outside it.
- "anchor boundary falls inside a tab/line-break": choose an anchor that starts/ends on text.
- Word warns about a corrupted document: run `python scripts/ai_review_to_comments.py verify annotated.docx` to see the missing reference ids, and inspect the OOXML details in `references/ooxml_notes.md`.

## Testing

Run the skill test suite with the bundled runtime python:

```bash
python scripts/test_ai_review.py
```

It builds a formatted fixture document (bold/italic runs, hyperlink, repeated phrases), applies reviews, exercises the polishing-skill convert workflow, and checks text preservation plus comment reference integrity.