# 与论文润色 Skill 联用的输出契约

本文件约束“论文润色类 Skill”（如学术写作、期刊润色、语言润色）的产出，使其能够直接被 docx-ai-review 转换为 Word 批注。润色 Skill 不应只输出全局重写稿，而应输出“结构化修改清单 + 可选整段改写”。

## 润色 Skill 的硬性约束

1. 逐段分析，不跨段修改：每个修改点必须落在单一段落内，禁止把跨段重写伪装成一个修改点。
2. 原文必须逐字真实：`original_text` 必须是从文档正文逐字复制的连续字符，禁止修正、补词、替换同义词后再当作原文。
3. 修改点原子化：一个修改点只处理一个具体问题。宁拆成 3 个更小修改点，也不要 1 个笼统修改点。
4. 定位信息齐全：必须给出段落索引、重复短语的第几次出现（occurrence），并说明问题类别与修改理由。
5. 每类问题只挑确实存在的硬伤：不把“风格偏好”写成“语法错误”；拿不准的意见不输出。
6. 期刊规范优先：如用户指定期刊（例如 Composites Part B: Engineering），应把符号、单位、术语、引文格式、Supporting Information 等规范作为独立修改点列出。
7. 整段改写单独给：若某段确实需要整体重写，把整段改写放在 `whole_paragraph_revision`，同时仍把其中的原子修改点逐一列出，不能只给整段。

## 输出 JSON 契约

```json
{
  "document": "复配0813英文_检查更新.docx",
  "journal_guidelines": "Composites Part B: Engineering",
  "edits": [
    {
      "paragraph_index": 30,
      "original_text": "calendering",
      "revision_type": "语法修正",
      "reason": "辊压工艺词拼写错误，标准写法为 calendaring",
      "revised_text": "calendaring",
      "occurrence": 1,
      "whole_paragraph_revision": ""
    }
  ]
}
```

`revision_type` 只能是：学术规范、语法修正、逻辑表达、用词精炼、句式润色。

## 转换流程

```bash
python scripts/ai_review_to_comments.py convert input.docx polish_edits.json -o annotated.docx
```

转换器会逐条核对 `original_text` 是否真实存在于对应段落，校验类别，然后把每条修改转换为格式为 `【类别】诊断：理由\n建议修改：修改后文本` 的 Word 批注；`whole_paragraph_revision` 非空时自动附为“整段修改参考”。

## 审阅模式（Track Changes）流程

当修改点具有明确“前后对照”（`original_text` → `revised_text`）时，优先使用 Word 审阅模式，而不是只加批注：

```bash
python scripts/ai_review_to_comments.py tracked input.docx polish_edits.json -o revised.docx
```

每条修改会变成真实的 `w:ins` / `w:del` 修订（Word “审阅”选项卡可直接看到前后对比并接受/拒绝），同时在该处附加一条“修改理由”批注。只有 `revised_text` 与原文不同才生成修订；纯插入、纯删除同样支持。

## 细化粒度建议

- 拼写与格式问题：一个词或一处符号一条。
- 术语问题：只锚定术语本身，不要连带整句。
- 语法问题：锚定错误的小句或搭配，避免整段。
- 逻辑问题：锚定转折/因果连接词或结论小句。
- 句式问题：锚定句式片段，并可在 `whole_paragraph_revision` 给整段示范。
- 重复表达：每一处重复分别用 occurrence 指定，避免只批第一处。

## 联合使用示例

1. 论文润色 Skill 先读取 `dump-text` 导出的段落 JSON，逐段产出 `polish_edits.json`。
2. docx-ai-review 运行 `convert`，把修改清单注入原稿。
3. 用 `verify` 检查批注引用完整性。

这样润色 Skill 负责“判断与改写”，docx-ai-review 负责“精确落点”，两者职责分离、可独立替换。