# 风格文件目录说明（styles/）

本目录存放「MD → DOCX」的**独立风格文件**。每个文件是一个标准 md，供人阅读/维护，
其中唯一的 ` ```yaml ... ``` ` 代码块会被 `scripts/md_to_docx.py` 提取为排版规格。

风格规格字段一览（详见各风格文件注释）：

| 字段 | 说明 |
| --- | --- |
| `page.width_cm / height_cm` | 纸张尺寸（A4 = 21×29.7） |
| `page.margin_*.cm` | 页边距（公文：上3.7/下3.5/左2.8/右2.6） |
| `fonts.body_cn / body_latin` | 正文中/英文字体 |
| `fonts.head_cn` | 默认标题字体（一级标题可被 headings 覆盖） |
| `fonts.code / quote_cn` | 代码/引用字体 |
| `body.size_pt` | 正文字号（pt，1 号=26pt, 2号=22, 3号=16, 4号=14, 小4=12） |
| `body.line_spacing_multiple` | 倍行距（如 1.5） |
| `body.line_spacing_fixed_pt` | 固定行距磅值（公文："28"，优先级高于 multiple） |
| `body.first_line_indent_chars` | 正文首行缩进字符数（公文："2"） |
| `title.*` | 文档大标题（自动提取 frontmatter 或 `--title`） |
| `headings.h1..h6` | 各级标题 `{cn, size_pt, bold}` |
| `code_block.*` | 围栏代码块 `{size_pt, fill}` |
| `table.*` | 表格 `{size_pt, header_fill, header_align}` |
| `quote.*` | 引用 `{size_pt, color, border_color}` |
| `date.*` | 日期行（公文 `<div class="date">…</div>`） |
| `attachment.*` | 附件行（公文 `<div class="attachment">…</div>`） |
| `page_number.*` | 页脚页码（`format: "-{N}-"` 支持自定义前后缀） |

## 现有风格

| 文件名 | 风格名 | 适用场景 |
| --- | --- | --- |
| `default.md` | 默认（A4 宋体小四 黑体标题） | 通用 md → docx；与原版默认行为一致 |
| `gov-doc.md` | 政府公文（GB/T 9704） | 党政机关、企事业单位正式公文 |

## 新增风格（针对新场景）

1. 复制任一现有文件并改文件名（建议 `<场景>-<调性>.md`，如 `contract-dark.md`）。
2. 调整头部说明（适用场景/调性/排版要点）。
3. 在 ` ```yaml ``` ` 代码块内修改字段值即可；所有字段都有缺省，未列出的字段会回退到默认。
4. 字段可省略但不能写错（如 `first_line_indent_chars: "2"` 字符串会被 PyYAML 当字符串，需去引号）。
5. 公文/特殊行用 `<div class="date">...</div>` / `<div class="attachment">...</div>` 包裹自动套用 `date` / `attachment` 风格。

## 依赖

使用 `--style` 时需 `pip install pyyaml`；`python-docx` 始终必需。
不传 `--style` 时使用内置默认规格（不依赖 PyYAML）。

## 选择规则（供调用方判断）

- 默认无明确场景 → `default.md`
- 党政机关公文 → `gov-doc.md`
- 其他 → 从描述词（场景/调性/字号/字体）中挑最贴近的；都不贴合时按上文新增一个风格文件，而非硬套。