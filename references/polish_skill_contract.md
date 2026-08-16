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
## 实战约束（来自真实论文审校）

1. 以当前原稿为准：意见稿常基于旧稿，原稿已修正的问题（拼写、符号、空格）不要重复修改。
2. 格式优先：`superscript`/`subscript` 等格式必须检查。单位指数（如 `W·m-1·K-1`）在 Word 中常已是上标，纯文本误判为错误时不要改。
3. 只应用“意见稿明确列出 + 原稿仍存在 + 有明确替换”的修改点；整段改写只作参考，不整段套用。
4. 每条修改必须挂批注：包含该句中文翻译（如有）和对应问题；同一段落内相同批注去重。
5. 大段改写按句切分、逐句字符级 diff，未变化文字不标记。
6. 一个意见段落可能对应原稿多个物理段落，应逐句定位，不能按块整体套用。

## 逐句对照类意见稿的实战约束

1. 兼容两类段落标记：``英文问题 / 不妥点`` 与 ``英文不妥 / 错误``、``Composites Part B 修改版本`` 与 ``Composites Part B 规范改写`` 都要支持。
2. 逐句对照使用词级 diff：整词替换整词，禁止字符级拼接，否则 ``hardware`` 会被拼成 ``harare``。
3. 替换文本必须从原文提取并保留词间空格；纯插入词要保留前导空格，否则 ``performance requirements`` 会变成 ``performancerequirements``。
4. 句子相似度过低（<0.5）时整句删除 + 整句插入，避免碎片化错乱。
5. 上标/下标双重保护：既跳过带 ``vertAlign`` 的 Run 区间，也跳过插入文本含 Unicode 上标字符的修改，避免误改单位指数。
6. 句子必须整句精确匹配才做 diff；意见稿基于旧版、无法精确匹配的句子直接跳过并报告，不模糊定位。
7. 句内插入需要零宽 span，插入前先在该位置分裂 Run，保证空格与顺序正确。
8. 生成后模拟 Word“接受全部修订”验证最终文本，确认无乱词。
