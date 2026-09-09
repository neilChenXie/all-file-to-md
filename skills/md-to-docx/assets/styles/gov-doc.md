# 风格：政府公文（GB/T 9704）

> **适用场景**：党政机关、企事业单位的正式公文（通知、通报、报告、请示、函等）。
> **标准依据**：GB/T 9704—2012《党政机关公文格式》。
> **排版要点**：方正小标宋_GBK 标题（2号居中）、仿宋_GB2312 正文（3号首行缩进 2 字）、黑体一级标题、楷体_GB2312 二级标题；
> A4 页边距 上3.7cm / 下3.5cm / 左2.8cm / 右2.6cm；行距固定 28pt（标题 32pt）；
> 页码 宋体 四号 "-N-" 格式居中。
> **使用方式**：
>   ```bash
>   python md_to_docx.py 公文.md -o 公文.docx \
>       --style assets/styles/gov-doc.md \
>       --title "关于XX的通知"
>   ```
> 公文末尾的日期/附件用 HTML 块包裹：

> **字体依赖**：正文 **仿宋_GB2312**（skill 自带 assets/fonts/ 中已内置该字体，
> 用 ensure_fonts.py 自动安装到系统后即生效）；标题依赖 **方正小标宋_GBK**。

>   ```html
>   <div class="date">2026 年 9 月 9 日</div>
>   <div class="attachment">附件：1. 任务分工表</div>
>   ```

## YAML（转换脚本提取本代码块）

```yaml
page:
  width_cm: 21.0
  height_cm: 29.7
  margin_top_cm: 3.7
  margin_bottom_cm: 3.5
  margin_left_cm: 2.8
  margin_right_cm: 2.6

fonts:
  body_cn: 仿宋_GB2312
  body_latin: Times New Roman
  head_cn: 黑体
  code: Consolas
  quote_cn: 楷体_GB2312
  link_color: "0563C1"

body:
  size_pt: 16                  # 3号
  line_spacing_fixed_pt: 28    # 固定 28 磅
  first_line_indent_chars: 2   # 首行缩进 2 字

title:
  cn: 方正小标宋_GBK
  latin: Times New Roman
  size_pt: 22                  # 2号
  bold: false
  align: center
  line_spacing_fixed_pt: 32    # 固定 32 磅

headings:
  h1: { cn: 黑体,        size_pt: 16, bold: false }   # 一级标题：黑体
  h2: { cn: 楷体_GB2312,  size_pt: 16, bold: false }   # 二级标题：楷体
  h3: { cn: 仿宋_GB2312,  size_pt: 16, bold: false }   # 三级标题：方正仿宋
  h4: { cn: 仿宋_GB2312,  size_pt: 16, bold: false }
  h5: { cn: 仿宋_GB2312,  size_pt: 16, bold: false }
  h6: { cn: 仿宋_GB2312,  size_pt: 16, bold: false }

code_block:
  size_pt: 14
  fill: "F2F2F2"

table:
  size_pt: 16
  header_fill: "EDEDED"
  header_align: center

quote:
  size_pt: 16
  color: "595959"
  border_color: "9E9E9E"

date:
  align: right                 # 右对齐
  margin_top_lines: 3          # 与正文空 3 行
  padding_right_chars: 4       # 右空 4 字

attachment:
  indent_chars: 2              # 左空 2 字

page_number:
  enabled: true                # 启用公文页码
  format: "-{N}-"              # "-1-" 格式
  font_cn: 宋体
  size_pt: 14                  # 四号
  align: center                # 居中
```