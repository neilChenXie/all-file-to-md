#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown → DOCX 转换器（md-to-docx 技能核心脚本）

将 Markdown(.md) 文件转换为排版规范的 Word(.docx) 文档：
  - A4 页面，正文宋体 12pt(小四)，标题黑体(一级 16pt / 二级 14pt / 三级 12pt)
  - 支持 ATX 标题(1-6 级)、表格、围栏代码块、无序/有序列表、引用、分隔线
  - 支持行内 **加粗** / *斜体* / ~~删除线~~ / `行内代码` / [链接](url) / <br> 换行
  - 支持本地图片 ![alt](相对或绝对路径)，宽度自动限制不超过 6 英寸
  - 文档开头 YAML frontmatter 自动忽略

依赖：python-docx（pip install python-docx）
用法：python md_to_docx.py 输入.md [-o 输出.docx]
"""

import argparse
import os
import re
import sys

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Inches, Pt, RGBColor
except ImportError:  # 依赖缺失时给出可操作的提示
    print("缺少依赖 python-docx，请先安装：pip install python-docx", file=sys.stderr)
    sys.exit(2)

# ----------------------------- 常量 ------------------------------------

FONT_CN_BODY = "宋体"
FONT_CN_HEAD = "黑体"
FONT_CN_QUOTE = "楷体"
FONT_LATIN = "Times New Roman"
FONT_CODE = "Consolas"

SIZE_BODY = 12          # 正文：小四
SIZE_CELL = 10.5        # 表格：五号
SIZE_CODE = 10.5        # 代码：五号
HEADING_SIZES = {1: 16, 2: 14, 3: 12, 4: 12, 5: 10.5, 6: 10.5}  # 三号/四号/小四/小四/五号/五号
HEADING_HEI = {1, 2, 3}  # 1-3 级标题用黑体，4-6 级用宋体加粗

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".emf", ".wmf", ".tif", ".tiff"}
MAX_IMAGE_W_IN = 6.0
MAX_IMAGE_H_IN = 8.5
CODE_FILL = "F2F2F2"     # 代码块底色
HEADER_FILL = "EDEDED"   # 表头底色
QUOTE_COLOR = RGBColor(0x59, 0x59, 0x59)
LINK_COLOR = "0563C1"

ATX_RE = re.compile(r"^(#{1,6})(?:\s+(.*?))?\s*$")
FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})\s*([\w+#.\-]*)\s*$")
HR_RE = re.compile(r"^\s*((?:-\s*){3,}|(?:\*\s*){3,}|(?:_\s*){3,})$")
UL_RE = re.compile(r"^\s*[-*+]\s+(.+)$")
OL_RE = re.compile(r"^\s*(\d{1,9})[.)]\s+(.+)$")
QUOTE_RE = re.compile(r"^(\s*>+)")
IMG_BLOCK_RE = re.compile(r"^!\[([^\]]*)\]\(\s*([^)\s]+)\s*\)\s*$")
# 行内元素：按顺序匹配 **粗体** / ~~删除线~~ / *斜体* / `行内代码` / ![图](url) / [文字](url) / <br>
INLINE_RE = re.compile(
    r"(\*\*[^*]+\*\*|~~[^~\n]+~~|`[^`\n]+`|\*[^*\n]+\*"
    r"|!\[[^\]\n]*\]\([^)\s]+\)|\[[^\]\n]+\]\([^)\s]+\)|<br\s*/?>|&nbsp;)"
)


def _set_run_font(run, east=FONT_CN_BODY, latin=FONT_LATIN, size=None,
                  bold=False, italic=False, color=None):
    """统一设置一个 run 的中西文字体与样式。"""
    run.font.name = latin
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east)
    if size:
        run.font.size = Pt(size)
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if color is not None:
        run.font.color.rgb = color


def _shade_paragraph(p, fill=CODE_FILL):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def _border_paragraph(p, edge="bottom", color="7F7F7F", sz="6"):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    b = OxmlElement(f"w:{edge}")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), sz)
    b.set(qn("w:space"), "4")
    b.set(qn("w:color"), color)
    pBdr.append(b)
    pPr.append(pBdr)


def _add_hyperlink(paragraph, url, text, size_pt=SIZE_BODY):
    """为段落添加一个可点击的超链接 run（字号跟随上下文）。"""
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                          is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), FONT_LATIN)
    rFonts.set(qn("w:hAnsi"), FONT_LATIN)
    rFonts.set(qn("w:eastAsia"), FONT_CN_BODY)
    rPr.append(rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))  # 半磅单位
    rPr.append(sz)
    c = OxmlElement("w:color")
    c.set(qn("w:val"), LINK_COLOR)
    rPr.append(c)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    r.append(rPr)

    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    hyperlink.append(r)
    paragraph._p.append(hyperlink)
    return hyperlink


def _strip_title_markers(text):
    """去掉标题文本中的行内标记，只保留纯文本。"""
    text = re.sub(r"!\[([^\]]*)\]\([^)\s]+\)", r"\1", text)   # 图片 -> alt
    text = re.sub(r"\[([^\]]+)\]\([^)\s]+\)", r"\1", text)    # 链接 -> 文字
    text = text.replace("**", "").replace("~~", "")
    text = re.sub(r"`([^`]*)`", r"\1", text)
    return re.sub(r"(?<!\*)\*(?!\*)([^*]*)\*", r"\1", text).strip() or "（空标题）"


def add_markdown_inline(paragraph, text, east=FONT_CN_BODY, latin=FONT_LATIN,
                        size=SIZE_BODY, bold=False, color=None):
    """把一段 Markdown 文本解析为行内格式并写入段落（支持加粗/斜体/删除线/行内代码/链接/<br>）。"""
    for part in INLINE_RE.split(text):
        if not part:
            continue
        if part == "<br>" or part == "<br/>" or part == "<br />":
            run = paragraph.add_run()
            run.add_break()
            continue
        if part == "&nbsp;":
            run = paragraph.add_run("\u00a0")
            _set_run_font(run, east=east, latin=latin, size=size, bold=bold, color=color)
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            _set_run_font(run, east=east, latin=latin, size=size, bold=True, color=color)
            continue
        if part.startswith("~~") and part.endswith("~~"):
            run = paragraph.add_run(part[2:-2])
            _set_run_font(run, east=east, latin=latin, size=size, bold=bold,
                          italic=True, color=color)
            run.font.strike = True
            continue
        if part.startswith("`") and part.endswith("`"):
            run = paragraph.add_run(part[1:-1])
            _set_run_font(run, east=FONT_CN_BODY, latin=FONT_CODE, size=size, bold=bold, color=color)
            continue
        if part.startswith("![") and "](http" in part:  # 行内网络图片：降级为文字
            m = re.match(r"!\[([^\]]*)\]\((http[^)\s]+)\)", part)
            run = paragraph.add_run(f"[图片:{m.group(1) or '链接'}]" if m else part)
            _set_run_font(run, east=east, latin=latin, size=size, bold=bold, color=color)
            continue
        m = re.match(r"\[([^\]]+)\]\(([^)\s]+)\)", part)   # [文字](url)
        if m:
            _add_hyperlink(paragraph, m.group(2), m.group(1), size_pt=size or SIZE_BODY)
            continue
        if part.startswith("*") and part.endswith("*") and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            _set_run_font(run, east=east, latin=latin, size=size, bold=bold,
                          italic=True, color=color)
            continue
        run = paragraph.add_run(part)
        _set_run_font(run, east=east, latin=latin, size=size, bold=bold, color=color)


# ----------------------------- 块级元素 ---------------------------------

def add_heading(doc, level, text):
    level = max(1, min(level, 6))
    p = doc.add_heading("", level=level)
    # Word 默认 Heading 样式带主题蓝色，这里改为黑色并设置中文字体
    add_markdown_inline(p, _strip_title_markers(text),
                        east=FONT_CN_HEAD if level in HEADING_HEI else FONT_CN_BODY,
                        latin=FONT_LATIN, size=HEADING_SIZES[level], bold=True,
                        color=RGBColor(0, 0, 0))
    return p


def add_hr(doc):
    p = doc.add_paragraph()
    _border_paragraph(p, edge="bottom", color="808080", sz="6")
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    return p


def add_code_block(doc, code_lines):
    """整块代码输出为单个带浅灰底纹、等宽字体的段落（保留空行与原缩进）。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.0
    pf.left_indent = Inches(0.1)
    _shade_paragraph(p, CODE_FILL)
    for idx, line in enumerate(code_lines):
        if idx > 0:
            br = p.add_run()
            br.add_break()
        run = p.add_run(line.replace("\t", "    "))
        _set_run_font(run, east=FONT_CN_BODY, latin=FONT_CODE, size=SIZE_CODE)
    return p


def add_quote_block(doc, lines):
    """连续引用行合并为一段：楷体、灰色、左侧竖线。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Inches(0.3)
    pf.space_before = Pt(2)
    pf.space_after = Pt(2)
    _border_paragraph(p, edge="left", color="9E9E9E", sz="16")
    for idx, (depth, content) in enumerate(lines):
        if idx > 0:
            br = p.add_run()
            br.add_break()
        add_markdown_inline(p, content, east=FONT_CN_QUOTE, latin=FONT_LATIN,
                            size=SIZE_BODY, color=QUOTE_COLOR)
    return p


def add_numbered_list(doc, items):
    """写入有序列表。手工编号以规避 python-docx 'List Number' 多个列表连号的已知问题。

    样式：编号从左侧 0.25 英寸起，文本在制表位 0.5 英寸处，回行与文本对齐。
    """
    for counter, text in enumerate(items, start=1):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Inches(0.5)
        pf.first_line_indent = Inches(-0.25)  # 悬挂缩进
        pf.tab_stops.add_tab_stop(Inches(0.5))
        add_markdown_inline(p, f"{counter}.\t", size=SIZE_BODY)
        add_markdown_inline(p, text, size=SIZE_BODY)


def add_image(doc, md_dir, alt, path):
    """插入本地图片：宽度自动限制在页宽内，居中显示；找不到时输出提示文字。"""
    path = path.strip().strip('"').strip("'")
    if not os.path.isabs(path):
        path = os.path.join(md_dir, path)
    if os.path.isfile(path) and os.path.splitext(path)[1].lower() in IMAGE_EXT:
        width_in = MAX_IMAGE_W_IN
        try:
            from PIL import Image as _PIL  # 可选依赖：读取原始像素尺寸避免小图被放大
            with _PIL.open(path) as im:
                w_px, h_px = im.size
                scale = 1.0
                if h_px / 96.0 > MAX_IMAGE_H_IN:
                    scale = min(scale, MAX_IMAGE_H_IN / (h_px / 96.0))
                if w_px / 96.0 > MAX_IMAGE_W_IN:
                    scale = min(scale, MAX_IMAGE_W_IN / (w_px / 96.0))
                width_in = (w_px / 96.0) * scale
        except Exception:
            pass
        doc.add_picture(path, width=Inches(width_in))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return True
    p = doc.add_paragraph()
    add_markdown_inline(p, f"[图片未能插入：{path or alt}]", color=RGBColor(0xC0, 0x00, 0x00))
    return False


# ----------------------------- 表格 -------------------------------------

SEP_RE = re.compile(r"^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$")


def is_separator_row(line):
    s = line.strip()
    if not s:
        return False
    return SEP_RE.match(s) is not None


def _cell_paragraph(cell, text, first=True):
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    if first:
        # 清空单元格默认产生的空 run，避免残留样式
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
    if text.strip():
        add_markdown_inline(p, text, east=FONT_CN_BODY, latin=FONT_LATIN, size=SIZE_CELL)
    return p


def convert_table(doc, table_lines):
    """把 Markdown 表格文本转换为 Word 表格。table_lines[1] 应为分隔行。"""
    if len(table_lines) < 2:
        return False
    if not is_separator_row(table_lines[1]):
        return False  # 第二行不是分隔行 -> 交给调用方按普通段落处理

    def split_row(line):
        s = line.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|") and not s.endswith("\\|"):
            s = s[:-1]
        # 简单处理单元格内的转义竖线
        cells, buf, esc = [], [], False
        for ch in s:
            if esc:
                buf.append(ch)
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == "|":
                cells.append("".join(buf).strip())
                buf = []
            else:
                buf.append(ch)
        cells.append("".join(buf).strip())
        return cells

    header = split_row(table_lines[0])
    data_rows = [split_row(line) for line in table_lines[2:] if line.strip()]
    ncols = max(len(header), *(len(r) for r in data_rows)) if data_rows else len(header)

    table = doc.add_table(rows=1 + len(data_rows), cols=ncols)
    table.style = "Table Grid"
    table.autofit = True

    for i in range(ncols):
        cell = table.rows[0].cells[i]
        text = header[i] if i < len(header) else ""
        _cell_paragraph(cell, text, first=True)
        # 表头格式：加粗、居中、浅灰底纹
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shade_paragraph(cell.paragraphs[0], HEADER_FILL)

    for r_idx, row in enumerate(data_rows):
        for c_idx in range(ncols):
            text = row[c_idx] if c_idx < len(row) else ""
            _cell_paragraph(table.rows[r_idx + 1].cells[c_idx], text, first=True)
    return True


# ----------------------------- 主转换流程 --------------------------------

def markdown_to_docx(input_file, output_file):
    content = _read_text(input_file)
    lines = content.splitlines()

    doc = Document()
    _setup_document(doc)

    md_dir = os.path.dirname(os.path.abspath(input_file))
    i = 0
    n = len(lines)
    quote_lines = []         # 连续引用行缓存
    ol_items = []            # 连续有序列表行缓存

    def flush_quote():
        nonlocal quote_lines
        if quote_lines:
            add_quote_block(doc, quote_lines)
            quote_lines = []

    def flush_ol():
        nonlocal ol_items
        if ol_items:
            add_numbered_list(doc, [t for _, t in ol_items])
            ol_items = []

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # 1) 空行：结束各类“连续块”
        if not stripped:
            flush_quote()
            flush_ol()
            i += 1
            continue

        # 2) 围栏代码块：从开 fence 起一直消费到收尾 fence，中间内容原样输出
        m = FENCE_RE.match(line)
        if m:
            flush_quote(); flush_ol()
            fence_char = m.group(1)[0]
            code_buffer = []
            i += 1
            while i < n:
                m2 = FENCE_RE.match(lines[i])
                if m2 and m2.group(1)[0] == fence_char:
                    i += 1
                    break
                code_buffer.append(lines[i])
                i += 1
            add_code_block(doc, code_buffer)
            continue

        # 3) 文档头 YAML frontmatter（--- 开始、--- 结束）
        if i == 0 and stripped == "---":
            i += 1
            while i < n and lines[i].strip() != "---":
                i += 1
            i += 1
            continue

        # 4) 标题
        m = ATX_RE.match(stripped)
        if m:
            flush_quote(); flush_ol()
            level = len(m.group(1))
            text = (m.group(2) or "").strip()
            if text:
                add_heading(doc, level, text)
            i += 1
            continue

        # 5) 分隔线
        if HR_RE.match(stripped):
            flush_quote(); flush_ol()
            add_hr(doc)
            i += 1
            continue

        # 6) 表格
        if stripped.startswith("|"):
            flush_quote(); flush_ol()
            table_lines = []
            while i < n and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            if not convert_table(doc, table_lines):
                # 不是合法表格(缺分隔行) -> 按原文段落输出
                for tl in table_lines:
                    p = doc.add_paragraph()
                    add_markdown_inline(p, tl, size=SIZE_BODY)
            continue

        # 7) 整行图片
        m = IMG_BLOCK_RE.match(stripped)
        if m:
            flush_quote(); flush_ol()
            add_image(doc, md_dir, m.group(1), m.group(2))
            i += 1
            continue

        # 8) 引用（连续引用行先缓存，遇到非引用行时统一输出）
        m = QUOTE_RE.match(line)
        if m:
            flush_ol()
            depth = m.group(1).count(">")
            content = line[m.end():]
            if content.startswith(" "):
                content = content[1:]
            quote_lines.append((depth, content))
            i += 1
            continue

        # 9) 无序列表
        m = UL_RE.match(line)
        if m:
            flush_quote()
            flush_ol()
            p = doc.add_paragraph(style="List Bullet")
            add_markdown_inline(p, m.group(1).strip(), size=SIZE_BODY)
            i += 1
            continue

        # 10) 有序列表（连续行收集后一次性渲染，保证手工编号连续）
        m = OL_RE.match(line)
        if m:
            flush_quote()
            ol_items.append((int(m.group(1)), m.group(2).strip()))
            i += 1
            continue

        # 11) 普通段落
        flush_quote(); flush_ol()
        p = doc.add_paragraph()
        add_markdown_inline(p, stripped, size=SIZE_BODY)
        i += 1

    flush_quote()
    flush_ol()

    # 输出目录不存在时自动创建
    out_dir = os.path.dirname(os.path.abspath(output_file))
    os.makedirs(out_dir, exist_ok=True)
    doc.save(output_file)
    print(f"转换完成：{output_file}")


def _read_text(input_file):
    """按 utf-8 读取，BOM 自动剥离；失败时回退 gb18030。"""
    for enc in ("utf-8-sig", "gb18030"):
        try:
            with open(input_file, "r", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(input_file, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def _setup_document(doc):
    """设置 A4 页面、Normal 样式（宋体小四、1.5 倍行距）。"""
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.17)
    sec.right_margin = Cm(3.17)

    normal = doc.styles["Normal"]
    normal.font.name = FONT_LATIN
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN_BODY)
    normal.font.size = Pt(SIZE_BODY)
    normal.paragraph_format.line_spacing = 1.5


def main():
    parser = argparse.ArgumentParser(
        description="将 Markdown(.md) 文件转换为排版规范的 Word(.docx) 文档",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例：
  python md_to_docx.py 文档.md                  # 输出同目录 文档.docx
  python md_to_docx.py 文档.md -o 输出.docx     # 指定输出文件
  python md_to_docx.py 文档.md --output "子目录/输出.docx"
        """,
    )
    parser.add_argument("input", help="输入的 Markdown 文件路径")
    parser.add_argument("-o", "--output",
                        help="输出的 DOCX 文件路径（默认与输入文件同名、同目录）")
    args = parser.parse_args()

    input_file = args.input
    if not os.path.isfile(input_file):
        print(f"错误：找不到输入文件 {input_file}", file=sys.stderr)
        sys.exit(1)

    if args.output:
        output_file = args.output
    else:
        output_file = os.path.splitext(input_file)[0] + ".docx"

    if not output_file.lower().endswith(".docx"):
        print(f"警告：输出文件建议使用 .docx 后缀：{output_file}")

    markdown_to_docx(input_file, output_file)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):  # Windows 控制台避免 GBK 编码报错
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    main()
