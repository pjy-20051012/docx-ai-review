# OOXML 批注与 Run 分裂实现说明

## 一条合法批注需要的部件

- word/document.xml：`<w:commentRangeStart w:id="N"/>` 锚定选区起点，`<w:commentRangeEnd w:id="N"/>` 锚定选区终点，紧跟一个含 `<w:commentReference w:id="N"/>` 的 `<w:r>`。
- word/comments.xml：`<w:comment w:id="N" w:author="AI Reviewer" w:initials="AI" w:date="...">` 存放批注正文。
- word/_rels/document.xml.rels：声明 comments part 的关系。
- [Content_Types].xml：注册 comments part 的 content type。

python-docx 1.2.0 的 `Document.comments.add_comment()` 会自动创建 comments part、关系与 content type；`Document.add_comment(runs, ...)` 会自动插入 commentRangeStart/End 与 commentReference。本 skill 的核心工作是先把任意字符区间精确切到 Run 边界，再交给该 API。

## Run 分裂算法

1. 按 `w:r` 与 `w:hyperlink` 直接子元素建立字符偏移映射（与 python-docx 的 `Paragraph.text` 输出对齐，含 `w:t`、`w:tab`、`w:br`、`w:cr`、`w:noBreakHyphen`）。
2. 找到覆盖 anchor 起止字符的容器；起止可以落在同一段落的不同 Run 上，但不能跨越超链接边界。
3. 需要分裂时 `copy.deepcopy(container)` 克隆容器（含 `w:rPr` 等全部格式属性），按偏移把 `w:t` 文本一分为二，并显式保留 `xml:space`；之后用 `addnext()` 把后缀容器插回。
4. 重新建立映射，把起止区间内的 `w:r` 收集成 Run 列表，交给 `Document.add_comment()`。

## 不变量与边界约束

- 分裂前后段落可见文本必须逐字符一致；本脚本用 `Paragraph.text` 重新比对锚点子串。
- 锚点必须落在单一段落内；跨超链接的选区不支持，模型应改用更短的锚点。
- 起止偏移若落在 `w:tab`/`w:br` 内部会明确报错。
- 同一段落同一短语的重复匹配由 `occurrence` 字段消歧。
- 输出后用 `verify_comment_integrity()` 解包 docx，核对 document.xml 中的所有批注 id 都存在于 comments.xml。

## 已验证的边界用例

- 单个 Run 中间分裂（英文/中文/全角标点）。
- 三种不同格式 Run 拼接后的跨 Run 选区。
- 同段落内跨多个 Run 的长短语锚定。
- 超链接内部文本的选区锚定。
- 同段重复短语（occurrence 消歧）。
- 多批注叠加时偏移保持稳定。