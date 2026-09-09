# 风格：默认（A4 宋体小四 黑体标题）

> **适用场景**：通用 md → docx 转换，沿用历史默认效果（A4 页面、宋体小四正文、黑体标题）。
> **特点**：页边距 2.54/2.54/3.17/3.17cm，正文 1.5 倍行距，标题黑体加粗。
> **使用方式**：`python md_to_docx.py 文档.md --style assets/styles/default.md`

## YAML（转换脚本提取本代码块）

```yaml
page:
  width_cm: 21.0
  height_cm: 29.7
  margin_top_cm: 2.54
  margin_bottom_cm: 2.54
  margin_left_cm: 3.17
  margin_right_cm: 3.17

fonts:
  body_cn: 宋体
  body_latin: Times New Roman
  head_cn: 黑体
  code: Consolas
  quote_cn: 楷体
  link_color: "0563C1"

body:
  size_pt: 12
  line_spacing_multiple: 1.5
  first_line_indent_chars: 0

title:
  cn: 黑体
  latin: Times New Roman
  size_pt: 18
  bold: true
  align: center
  line_spacing_multiple: 1.5

headings:
  h1: { cn: 黑体, size_pt: 16, bold: true }
  h2: { cn: 黑体, size_pt: 14, bold: true }
  h3: { cn: 黑体, size_pt: 12, bold: true }
  h4: { cn: 宋体, size_pt: 12, bold: true }
  h5: { cn: 宋体, size_pt: 10.5, bold: true }
  h6: { cn: 宋体, size_pt: 10.5, bold: true }

code_block:
  size_pt: 10.5
  fill: "F2F2F2"

table:
  size_pt: 10.5
  header_fill: "EDEDED"
  header_align: center

quote:
  size_pt: 12
  color: "595959"
  border_color: "9E9E9E"

date:
  align: right
  margin_top_lines: 0
  padding_right_chars: 0

attachment:
  indent_chars: 0

page_number:
  enabled: false
  format: "-{N}-"
  font_cn: 宋体
  size_pt: 14
  align: center
```