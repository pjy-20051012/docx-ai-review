# docx-ai-review / AI 润色批注技能

Turn AI or mentor feedback into native Word comments anchored to exact text spans, without changing the original document body.
将 AI / 导师意见以 Word 原生批注形式精确锚定到原文对应位置，正文与排版完全保持不变。

## Features / 功能

- Native Word comments (`word/comments.xml`) with character-level anchors, so every comment highlights the exact phrase it refers to.
- 原生 Word 批注 + 字符级精确锚定，每条批注只高亮对应短语。
- Formatting-preserving run isolation: Word fragments text across `<w:r>` elements; this skill splits runs with `copy.deepcopy` and preserves `xml:space`, bold, italic, font, color, etc.
- 保留格式的 Run 分裂：自动处理 Word 跨 Run 碎化文本，分裂后字体、加粗、斜体等格式原样保留。
- Two workflows (below) / 两种使用方式（见下）。
- Automatic OOXML integrity verification: every `commentRangeStart/End/Reference` id is checked against `comments.xml`.
- 自动校验 OOXML 结构完整性，批注引用 ID 一一对应。

## Two Ways to Use / 两种使用方式

### 1. Direct review-to-comments / 直接把意见稿转成批注

Upload an original `.docx` plus an AI or mentor opinion file (e.g. a Markdown paragraph-by-paragraph review). The skill reads the manuscript text, aligns each opinion to the matching paragraph, and injects comments with:
上传原稿 `.docx` 和 AI/导师意见稿（例如逐段分析的 Markdown），技能会读取原稿、把意见对齐到对应段落并注入批注：

- exact anchors copied verbatim from the manuscript / 从原文逐字复制的精确锚点
- granular categories: 学术规范 / 语法修正 / 逻辑表达 / 用词精炼 / 句式润色
- concise diagnosis plus a local suggested revision / 简短诊断 + 局部修改建议
- optional full-paragraph revision appended as reference / 可选整段修改参考

```bash
python scripts/ai_review_to_comments.py dump-text input.docx --output text.json
python scripts/ai_review_to_comments.py apply input.docx reviews.json -o annotated.docx
python scripts/ai_review_to_comments.py verify annotated.docx
```

See `examples/` for a working sample / 示例见 `examples/`。

For reviews that clearly provide before/after pairs, use Word review mode (tracked changes) with a reason comment on every change:
当意见稿明确给出“前后对照”时，使用 Word 审阅模式修改并逐条附批注理由：

```bash
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx
```

Open the result in Word, switch to the Review tab, and accept or reject each change.
用 Word 打开后，在“审阅”选项卡中即可看到前后对比并逐条接受/拒绝。

For full-paragraph rewrites, the skill splits both the original and the rewritten paragraph into sentences, then applies each pair as a tracked deletion followed by a tracked insertion, so every change has a visible before/after:
整段改写会自动按句拆分：原句删除、改写句紧跟插入，逐句成对显示前后对照：

```bash
python scripts/ai_review_to_comments.py rewrite input.docx paragraph_rewrites.json -o revised.docx
```

### 2. Combined with paper-polishing skills / 与论文润色技能联用

Use together with polishing skills such as `nature-polishing`, `academic-paper`, or any paper-language-polishing skill. The contract in `references/polish_skill_contract.md` constrains the polishing skill output so comments become more granular:
与 `nature-polishing`、`academic-paper` 等论文润色技能联用，通过 `references/polish_skill_contract.md` 中的输出契约约束润色技能产出，使批注更细化：

- polish skills must emit a structured edit list, not only a global rewrite
- 润色技能必须输出结构化修改清单，而不是只给全局重写稿
- each edit is atomic, scoped to one paragraph, with verbatim `original_text`
- 每个修改点原子化、限定单段、原文必须逐字真实
- each edit carries category, reason, occurrence, and optional whole-paragraph rewrite
- 每条修改带类别、理由、重复出现序号与可选整段改写

```bash
python scripts/ai_review_to_comments.py convert input.docx polish_edits.json -o annotated.docx
```

## Installation / 安装

Clone this repository and put the folder into your skills directory:
克隆本仓库，把文件夹放入你的 skills 目录：

```bash
git clone https://github.com/YOUR_NAME/docx-ai-review.git
# Windows / Codex: copy the folder to E:\下载\Codex\home\skills\
# or macOS/Linux: copy to $CODEX_HOME/skills/
```

Requirements / 依赖：

- Python 3.10+
- `python-docx>=1.2.0`
- `lxml>=4.9.0`
- `pydantic>=2.0.0`

```bash
python -m pip install "python-docx>=1.2.0" "lxml>=4.9.0" "pydantic>=2.0.0"
```

## Repository Layout / 仓库结构

```text
docx-ai-review/
├── SKILL.md                          # Skill entry / 技能入口
├── README.md                         # This file / 本说明
├── LICENSE                           # License / 许可证
├── agents/openai.yaml                # UI metadata / 界面元数据
├── references/
│   ├── ooxml_notes.md                # OOXML internals / 底层实现说明
│   ├── polish_skill_contract.md      # Contract for polishing skills / 润色技能输出契约
│   └── review_prompt_template.md     # Review JSON prompt / 批注生成提示词
├── scripts/
│   ├── ai_review_to_comments.py      # Main CLI / 主程序
│   └── test_ai_review.py             # Test suite / 测试套件
└── examples/
    ├── demo_source.docx              # Sample manuscript / 示例原稿
    ├── demo_reviews.json             # Sample reviews / 示例批注数据
    └── demo_annotated.docx           # Annotated result / 示例批注成品
```

## CLI / 命令

| Command / 命令 | Description / 说明 |
|---|---|
| `dump-text input.docx [-o text.json]` | Extract paragraphs with indices / 抽取正文与段落索引 |
| `apply input.docx reviews.json -o out.docx` | Inject review comments / 注入批注 |
| `convert input.docx polish_edits.json -o out.docx` | Convert polishing-skill edit list into comments / 把润色修改清单转为批注 |
| `tracked input.docx polish_edits.json -o out.docx` | Apply edits as tracked changes with reason comments / 审阅模式修订并附批注理由 |
| `rewrite input.docx paragraph_rewrites.json -o out.docx` | Apply full paragraph rewrites as sentence-level tracked changes / 整段改写按句修订 |
| `verify annotated.docx` | Check comment reference integrity / 校验批注引用完整性 |

## Refining a review markdown / 意见稿细化

When the review is a paragraph-by-paragraph markdown (original + Chinese translation + issue list + full rewrite), refine it into the structured edit list first, then apply tracked changes:
当意见稿是逐段 Markdown（原文 + 中文翻译 + 问题 + 整段改写）时，先细化为结构化修改清单，再应用审阅模式修改：

```bash
python scripts/refine_review_markdown.py input.docx review.md -o polish_edits.json
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx
```

## Practical lessons / 实战经验

- Validate every anchor against the current manuscript; reviews often target an older draft. / 所有锚点以当前原稿为准，意见稿常基于旧稿。
- Check `superscript`/`subscript` formatting before treating plain text as an error; unit exponents are often already superscript in Word. / 先检查上标/下标格式，单位指数常已是上标，不要误改。
- Apply issue-driven edits only: review-listed issue + present in document + explicit replacement. / 只应用“意见稿列出 + 原稿存在 + 有明确替换”的修改。
- Pair every tracked change with a reason comment; deduplicate identical comments within a paragraph. / 每条修改配批注理由，同段重复批注去重。
- Split large rewrites sentence-by-sentence with character-level diffs; never mark unchanged text. / 大段改写按句切分并做字符级 diff，不标记未变化文字。
- Use word-level diffs for sentence pairing (character diffs mangle replaced words), preserve spaces, protect superscript unit exponents, and verify the accepted state. / 逐句对照用词级 diff（字符级会拼坏单词）、保留空格、保护上标单位，并验证接受态。

## Testing / 测试

```bash
python scripts/test_ai_review.py
```

The test suite covers single-run splits, cross-run spans, Chinese characters, hyperlink anchors, repeated phrases, full-paragraph revisions, and the convert workflow.
测试覆盖单 Run 分裂、跨 Run 选区、中文字符、超链接锚点、重复短语、整段改写与 convert 联用流程。