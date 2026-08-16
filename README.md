# docx-ai-review

**English** | [中文](./README.zh-CN.md)

Turn AI or mentor review feedback into native Word comments and tracked changes anchored to exact text spans, without rewriting the original body.

## Core idea

- Word **tracked changes** show before/after clearly: deletions in red, insertions in green.
- Every comment carries the **problem**, the **sentence translation** (when the review provides one), and the **suggestion**.
- Anchors are **character-precise**; sentence pairs are compared at the **sentence level** with word-level diffs.
- The original body and formatting are never destroyed; the author accepts or rejects each change in Word's Review tab.

## Features

- Native Word comments (`word/comments.xml`) with character-level anchors.
- Formatting-preserving run isolation: handles fragmented `<w:r>` runs while keeping `rPr`, `xml:space`, bold, italic, font and color.
- Tracked changes (`w:ins` / `w:del`) with red/green colors and reason comments.
- Automatic OOXML integrity verification.
- Review-markdown refinement: paragraph-by-paragraph markdown -> `polish_edits.json` -> tracked changes.
- Combined workflow with paper-polishing skills via a structured output contract.

## Two ways to use

### 1. Direct review to Word comments / tracked changes

```bash
# Extract paragraphs with indices
python scripts/ai_review_to_comments.py dump-text input.docx -o text.json

# Comment-only injection
python scripts/ai_review_to_comments.py apply input.docx reviews.json -o annotated.docx

# Tracked changes with reason comments (clear before/after pairs)
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx

# Full-paragraph rewrites split into sentence-level tracked changes
python scripts/ai_review_to_comments.py rewrite input.docx paragraph_rewrites.json -o revised.docx

# Verify comment/revision integrity
python scripts/ai_review_to_comments.py verify revised.docx
```

### 2. Refine a paragraph-by-paragraph review markdown

```bash
python scripts/refine_review_markdown.py input.docx review.md -o polish_edits.json
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx
```

The refiner keeps only edits that are review-listed, still present in the current document, and have an explicit replacement. It skips superscript unit exponents that are already correct.

### 3. Combined with paper-polishing skills

Require the polishing skill (e.g. `nature-polishing`, `academic-paper`) to emit the structured edit list defined in `references/polish_skill_contract.md`, then apply directly:

```bash
python scripts/ai_review_to_comments.py convert input.docx polish_edits.json -o annotated.docx

The polished result lands directly in the Word document as tracked changes and comments, so there is no need to compare two windows manually.

![Before/After: polishing skill working directly in Word](docs/images/before-after.png)
```

## Installation

```bash
git clone https://github.com/pjy-20051012/docx-ai-review.git
# copy the folder into your skills directory, e.g. E:\下载\Codex\home\skills\
```

Requirements: Python 3.10+, `python-docx>=1.2.0`, `lxml>=4.9.0`, `pydantic>=2.0.0`.

```bash
python -m pip install "python-docx>=1.2.0" "lxml>=4.9.0" "pydantic>=2.0.0"
```

## Repository layout

```text
docx-ai-review/
├── SKILL.md                          # Skill entry
├── README.md                         # English readme
├── README.zh-CN.md                   # Chinese readme
├── LICENSE
├── agents/openai.yaml                # UI metadata
├── references/
│   ├── ooxml_notes.md                # OOXML internals
│   ├── polish_skill_contract.md      # Contract for polishing skills
│   └── review_prompt_template.md     # Review JSON prompt
├── scripts/
│   ├── ai_review_to_comments.py      # Main CLI
│   ├── refine_review_markdown.py     # Review markdown refinement
│   └── test_ai_review.py             # Test suite
└── examples/
    ├── demo_source.docx
    ├── demo_reviews.json
    ├── demo_polish_edits.json
    └── demo_annotated.docx
```

## CLI

| Command | Description |
|---|---|
| `dump-text input.docx [-o text.json]` | Extract paragraphs with indices |
| `apply input.docx reviews.json -o out.docx` | Inject review comments |
| `convert input.docx polish_edits.json -o out.docx` | Convert polishing-skill edit list into comments |
| `tracked input.docx polish_edits.json -o out.docx` | Apply edits as tracked changes with reason comments |
| `rewrite input.docx paragraph_rewrites.json -o out.docx` | Apply full paragraph rewrites as sentence-level tracked changes |
| `verify annotated.docx` | Check comment reference integrity |

## Practical lessons

- Validate every anchor against the current manuscript; reviews often target an older draft.
- Check `superscript`/`subscript` formatting before treating plain text as an error; unit exponents are often already superscript in Word.
- Apply issue-driven edits only: review-listed issue + present in document + explicit replacement.
- Pair every tracked change with a comment containing the problem, sentence translation, and suggestion; deduplicate identical comments within a paragraph.
- Compare sentences with word-level diffs (character diffs mangle replaced words), preserve spaces, protect superscript units, and verify the accepted state.
- Split large rewrites sentence-by-sentence; never mark unchanged text.

## Testing

```bash
python scripts/test_ai_review.py
```

Covers run splitting, cross-run spans, Chinese characters, hyperlink anchors, repeated phrases, full-paragraph rewrites, tracked changes and the convert workflow.
