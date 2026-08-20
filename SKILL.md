---
name: docx-ai-review
description: "Use when the user wants AI or mentor review/polish feedback applied to a Word document as native Word comments or Word tracked changes (审阅模式) without rewriting the original body. Triggers include: 意见稿、润色意见、AI 修改稿、导师意见、markdown 逐段意见、审阅模式修订、tracked changes、前后对照、逐句对照、批注、Word 批注、原文加批注、按句修改、高亮修改、接受/拒绝修改、docx 批注。Also use for character-level comment anchoring, sentence-level revision, review-markdown refinement, and preserving original formatting while adding precise review notes."
---

# DOCX AI Review Comments

把 AI / 导师 / 润色意见应用到 Word 原稿。核心是审阅模式修改：能明确前后对照的地方用 Word 审阅模式（红删绿插）真实修改，直观显示前后变化；每条批注标注该处存在问题、句子中文翻译与修改意见；标注位置必须字符级准确。无法明确替换或只是参考的整段改写只作批注。正文与格式不破坏，最终由作者在“审阅”视图逐条接受或拒绝。

## Trigger keywords / 触发场景

用户需求中出现以下关键词时使用本技能：

- 意见稿 / 润色意见 / AI 修改稿 / 导师意见 / 师兄意见 / DeepSeek 意见
- Word 批注 / 批注 / 原文加批注 / 问题批注 / 右侧批注
- 审阅模式 / 修订 / tracked changes / 删除线 / 红删绿插 / 接受或拒绝修改
- 前后对照 / 逐句对照 / 按句修改 / 句子级修改 / 一一对照
- docx / Word 文档修改意见 / markdown 意见稿 / 修改建议落到原文

## Core idea / 核心思路

1. **先锁定干净初稿**：`tracked` / `rewrite` 默认拒绝含既有 `w:ins` / `w:del` 的输入；不要在“全部接受修订”或混合修订文件上继续叠加。
2. **原稿优先**：所有锚点以当前干净初稿为准；意见稿常基于旧稿，原稿已修正的问题直接跳过。
3. **问题驱动**：只修改“意见稿明确列出 + 原稿仍存在 + 有明确替换”的内容，不用整段改写整体套用。
4. **审阅模式 + 批注**：有前后对照的修改写成 `w:ins` / `w:del`（红删绿插）；每条修改旁批注包含问题、句子中文翻译（若输入提供）与修改意见；段级问题单独一条批注。
5. **字段不破坏**：含 Zotero、交叉引用、`w:instrText`、`w:fldChar` 或书签的段落不得清空或重建；不得把字段转换成普通文本。
6. **格式感知**：检查 `superscript` / `subscript`，单位指数这类原稿已正确的格式不误改；纯格式修正按用户明确要求直接改，不伪装成语言修订。
7. **词级 diff**：逐句对照用词级差异（整词替换整词），避免字符级拼接产生乱词；保留词间空格。
8. **句级对比**：前后文对比以句为单位，句子精确配对后再标注句内差异；只有整句确实不同或无法可靠配对时才整句删插。

## Quick start / 快速上手

```bash
# 抽取正文，供模型/意见稿定位
python scripts/ai_review_to_comments.py dump-text input.docx -o text.json

# 结构化批注（只加批注，不改正文）
python scripts/ai_review_to_comments.py apply input.docx reviews.json -o annotated.docx

# 审阅模式修改（polish_edits.json：原文/修改后/理由）
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx

# 整段改写按句修订
python scripts/ai_review_to_comments.py rewrite input.docx paragraph_rewrites.json -o revised.docx

# 把逐段 markdown 意见稿细化为修改清单，再审阅修改
python scripts/refine_review_markdown.py input.docx review.md -o polish_edits.json
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx

# 校验批注/修订引用完整性
python scripts/ai_review_to_comments.py verify revised.docx

# 以干净初稿为基准审计原始态、非目标段落、字段/书签和批注引用
python scripts/audit_review_integrity.py input.docx revised.docx --target-index 30 --target-index 42
```

## Workflow A: structured reviews / polish edits

输入是结构化 JSON（`reviews.json` 或 `polish_edits.json`）时：

- `apply`：只注入原生批注，不修改正文。
- `tracked`：把“原文 → 修改后”逐条写成审阅模式修改，并附修改理由批注。

字段契约见 `references/polish_skill_contract.md`。

## Workflow B: paragraph-by-paragraph review markdown

输入是逐段 Markdown（每段含 **原文 / 中文翻译 / 英文问题 / 修改版本**）时：

1. 用 `refine_review_markdown.py` 解析并细化为 `polish_edits.json`，只保留“意见稿列出 + 原稿存在 + 有明确替换”的修改点。
2. 用 `tracked` 应用审阅修改。
3. 兼容两种标记：`英文问题 / 不妥点` 与 `英文不妥 / 错误`、`Composites Part B 修改版本` 与 `Composites Part B 规范改写`。

若意见稿提供完整“原稿 ↔ 新文”逐句对照，按句子配对后用词级 diff 标注差异；相似度过低的句子整句删除+插入；含上标单位的句子跳过。

修订前必须先确认输入没有旧修订。若文档含 Zotero 或交叉引用，先从无修订的初稿生成审阅版，再用 Word 中已安装的 Zotero 插件刷新字段；刷新后仍需核对字段代码和书签结构未变。详细规则见 `references/baseline-fields-and-formatting.md`。

## Workflow C: combined use with paper-polishing skills

与 `nature-polishing`、`academic-paper` 等润色技能联用时，**先给润色技能提要求**，再接收其输出。要求包括：按 `references/polish_skill_contract.md` 输出结构化修改清单；逐句给出原文与修改后文本；注明问题类别、理由与中文翻译；锚点必须逐字来自当前原稿。这样本技能可直接在文中以审阅模式落改：

```bash
python scripts/ai_review_to_comments.py convert input.docx polish_edits.json -o annotated.docx
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx
```

润色技能负责判断与改写，本技能负责精确落点与审阅展示，职责分离。

## Practical lessons / 实战经验

- 锚点一律以当前文档为准；意见稿基于旧稿时，原稿已修正的拼写/符号/空格跳过。
- 格式感知：`superscript`/`subscript` 必须检查；`W·m-1·K-1` 这类单位指数在 Word 中常已是上标，不要误改。
- 问题驱动优先于整段套用：只改“列出 + 存在 + 有替换”的问题；整段改写只作参考批注。
- 每条修改配批注，内容包含该处存在问题、句子中文翻译与修改意见；同段相同批注去重；段级问题单独一条批注。
- 逐句对照用词级 diff（字符级会拼坏单词）；替换文本保留空格；插入保留前导空白。
- 句子相似度过低时整句删除+插入；句子必须整句精确匹配才 diff，无法匹配的跳过并报告。
- 句内插入用零宽 span，先分裂 Run。
- Run 分裂必须字符精确；生成后模拟“接受全部修订”验证无乱词。
- 一个意见段落可能对应原稿多个物理段落，应逐句定位。
- 新审阅版的“原始态”（普通文本 + 删除文本，不含插入文本）必须与干净初稿一致；非目标段落不得出现修订。
- Zotero/交叉引用字段的 `instrText`、`fldChar` 与书签结构必须和初稿保持一致；不能只凭 Word 当前显示文本判断。
- 上下标检查应区分文本修改与格式修改；不要对修订、字段结果或引用段落做宽泛正则替换。

## Rules enforced by scripts

- Review items：`paragraph_index`、精确 `anchor_text`、`category`（学术规范/语法修正/逻辑表达/用词精炼/句式润色）、`problem_analysis`、`suggested_revision`；`occurrence` 消歧，`paragraph_revision` 附整段参考。
- 锚点必须是单段内子串，可跨 Run 不可跨超链接；分裂保留 `rPr` 与 `xml:space`。
- 批注格式：`【类别】诊断：…\n建议修改：…`；作者 `AI Reviewer`，缩写 `AI`。
- 修订必须带 author/date；引用 ID 与 `comments.xml` 一一对应。
- `audit_review_integrity.py` 可审计干净基线、原始态、目标段落隔离、嵌套修订、字段/书签和批注实体引用。

## Troubleshooting

- 锚点跨超链接/制表符：缩短锚点或调整提示词。
- Word 提示文档损坏：运行 `verify` 查看缺失引用 ID，参考 `references/ooxml_notes.md`。
- 单位被误改：检查是否为 `superscript` 格式，并在细化时跳过含上标单位的 span。
- 文档出现重复修订或初稿内容被替换：回到无修订初稿重新生成；不要使用 `--allow-existing-revisions` 作为常规方案。
- Zotero 引用变成静态文字或交叉引用失效：恢复干净初稿中的原生字段，重新落修订，再通过 Word/Zotero 刷新并运行完整审计。

## Testing

```bash
python scripts/test_ai_review.py
```

覆盖单 Run 分裂、跨 Run、中文、超链接、重复短语、整段改写、审阅模式与联用流程。
