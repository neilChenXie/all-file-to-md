#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md_to_docx.py —— Markdown → DOCX（支持独立风格文件）

将 Markdown(.md) 文件转换为排版规范的 Word(.docx) 文档。

风格（字体/字号/行距/页边距/页码等）通过 ``--style <风格md文件>`` 加载独立风格文件，
风格 md 文件含 `````yaml`` 围栏块；缺省走内置 default 风格（兼容旧版默认效果）。

支持：
  - ATX 标题(1-6 级)、表格、围栏代码块、无序/有序列表、引用、分隔线
  - 行内 **加粗** / *斜体* / ~~删除线~~ / `行内代码` / [链接](url) / <br> 换行
  - 本地图片 ![alt](相对或绝对路径)
  - YAML frontmatter 自动忽略；``title:`` 字段可自动作为文档大标题（也可用 ``--title`` 覆盖）
  - 公文专属：``<div class="date">日期</div>`` / ``<div class="attachment">...</div>`` 自动套用风格

依赖：python-docx、PyYAML（使用 --style 时必需）
用法：python md_to_docx.py 输入.md [-o 输出.docx] [--style 风格md] [--title 文档标题]
"""

import argparse
import os
import re
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Inches, Pt, RGBColor
except ImportError:
    print("缺少依赖 python-docx，请先安装：pip install python-docx", file=sys.stderr)
    sys.exit(2)

try:
    import yaml
except ImportError:
    yaml = None  # 使用 --style 时会报错提示

# 脚本所在目录的上一级 = skill 根目录
SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_STYLE = SKILL_DIR / "assets" / "styles" / "default.md"

# ====================================================================
# 内置默认风格（缺省 / 无 --style 时使用）
# ====================================================================
DEFAULT_SPEC: dict = {
    "page": {
        "width_cm": 21.0, "height_cm": 29.7,
        "margin_top_cm": 2.54, "margin_bottom_cm": 2.54,
        "margin_left_cm": 3.17, "margin_right_cm": 3.17,
    },
    "fonts": {
        "body_cn": "宋体", "body_latin": "Times New Roman",
        "head_cn": "黑体", "code": "Consolas", "quote_cn": "楷体",
        "link_color": "0563C1",
    },
    "body": {
        "size_pt": 12, "line_spacing_multiple": 1.5, "first_line_indent_chars": 0,
    },
    "title": {
        "cn": "黑体", "latin": "Times New Roman",
        "size_pt": 18, "bold": True, "align": "center",
        "line_spacing_multiple": 1.5,
    },
    "headings": {
        "h1": {"cn": "黑体", "size_pt": 16, "bold": True},
        "h2": {"cn": "黑体", "size_pt": 14, "bold": True},
        "h3": {"cn": "黑体", "size_pt": 12, "bold": True},
        "h4": {"cn": "宋体", "size_pt": 12, "bold": True},
        "h5": {"cn": "宋体", "size_pt": 10.5, "bold": True},
        "h6": {"cn": "宋体", "size_pt": 10.5, "bold": True},
    },
    "code_block": {"size_pt": 10.5, "fill": "F2F2F2"},
    "table": {"size_pt": 10.5, "header_fill": "EDEDED", "header_align": "center"},
    "quote": {"size_pt": 12, "color": "595959", "border_color": "9E9E9E"},
    "date": {"align": "right", "margin_top_lines": 0, "padding_right_chars": 0},
    "attachment": {"indent_chars": 0},
    "page_number": {
        "enabled": False, "format": "-{N}-",
        "font_cn": "宋体", "size_pt": 14, "align": "center",
    },
}


def _merge_spec(base: dict, extra: dict) -> dict:
    """深度合并；extra 覆盖 base。"""
    out = {k: v for k, v in base.items()}
    for k, v in extra.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = _merge_spec(out[k], v)
        else:
            out[k] = v
    return out


def load_style(path: Path | None) -> dict:
    """加载风格 md 中的 ```yaml ... ``` 块并合并到默认 spec。"""
    if path is None:
        return DEFAULT_SPEC
    if not path.exists():
        raise RuntimeError(f"风格文件不存在: {path}")
    if yaml is None:
        raise RuntimeError(
            "使用 --style 需要 PyYAML，请先安装：pip install pyyaml\n"
            "（不传 --style 可用内置默认风格，无需 PyYAML）"
        )
    text = path.read_text(encoding="utf-8")
    m = re.search(r"```yaml\s*\n(.*?)```", text, re.S)
    if not m:
        raise RuntimeError(f"风格文件 {path} 中未找到 ```yaml ... ``` 代码块")
    spec = yaml.safe_load(m.group(1)) or {}
    return _merge_spec(DEFAULT_SPEC, spec)


def _parse_frontmatter_title(content: str) -> str | None:
    """从 YAML frontmatter 中提取 title: 字段（保留中文/引号）。"""
    if not content.startswith("---"):
        return None
    for line in content.splitlines()[1:]:
        s = line.strip()
        if s == "---":
            break
        m = re.match(r"^title\s*:\s*(.+?)\s*$", s)
        if m:
            v = m.group(1).strip()
            if len(v) >= 2 and v[0] == v[-1] and v[0] in ("'", '"'):
                v = v[1:-1]
            return v or None
    return None


# ====================================================================
# 字体/段落辅助
# ====================================================================
def _set_run_font(run, east, latin, size=None, bold=False, italic=False, color=None):
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


def _apply_line_spacing(pf, body_spec: dict):
    if "line_spacing_fixed_pt" in body_spec:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(body_spec["line_spacing_fixed_pt"])
    elif "line_spacing_multiple" in body_spec:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = body_spec["line_spacing_multiple"]


def _shade_paragraph(p, fill):
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


def _add_hyperlink(paragraph, url, text, size_pt):
    spec = _CURRENT_SPEC
    fonts = spec["fonts"]
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), fonts["body_latin"])
    rFonts.set(qn("w:hAnsi"), fonts["body_latin"])
    rFonts.set(qn("w:eastAsia"), fonts["body_cn"])
    rPr.append(rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    rPr.append(sz)
    c = OxmlElement("w:color")
    c.set(qn("w:val"), fonts["link_color"])
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


# 全局当前规格（在 markdown_to_docx 中设置，被辅助函数读取）
_CURRENT_SPEC: dict = {}


def _strip_title_markers(text):
    text = re.sub(r"!\[([^\]]*)\]\([^)\s]+\)", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)\s]+\)", r"\1", text)
    text = text.replace("**", "").replace("~~", "")
    text = re.sub(r"`([^`]*)`", r"\1", text)
    return re.sub(r"(?<!\*)\*(?!\*)([^*]*)\*", r"\1", text).strip() or "（空标题）"


INLINE_RE = re.compile(
    r"(\*\*[^*]+\*\*|~~[^~\n]+~~|`[^`\n]+`|\*[^*\n]+\*"
    r"|!\[[^\]\n]*\]\([^)\s]+\)|\[[^\]\n]+\]\([^)\s]+\)|<br\s*/?>|&nbsp;)"
)


def add_markdown_inline(paragraph, text, east=None, latin=None, size=None,
                        bold=False, color=None):
    """把一段 Markdown 文本解析为行内格式并写入段落。默认取自当前 spec。"""
    spec = _CURRENT_SPEC
    fonts = spec["fonts"]
    east = east or fonts["body_cn"]
    latin = latin or fonts["body_latin"]
    size = size or spec["body"]["size_pt"]
    for part in INLINE_RE.split(text):
        if not part:
            continue
        if part in ("<br>", "<br/>", "<br />"):
            paragraph.add_run().add_break()
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
            _set_run_font(run, east=east, latin=latin, size=size, bold=bold, italic=True, color=color)
            run.font.strike = True
            continue
        if part.startswith("`") and part.endswith("`"):
            run = paragraph.add_run(part[1:-1])
            _set_run_font(run, east=east, latin=fonts["code"], size=size, bold=bold, color=color)
            continue
        if part.startswith("![") and "](http" in part:
            m = re.match(r"!\[([^\]]*)\]\((http[^)\s]+)\)", part)
            run = paragraph.add_run(f"[图片:{m.group(1) or '链接'}]" if m else part)
            _set_run_font(run, east=east, latin=latin, size=size, bold=bold, color=color)
            continue
        m = re.match(r"\[([^\]]+)\]\(([^)\s]+)\)", part)
        if m:
            _add_hyperlink(paragraph, m.group(2), m.group(1), size_pt=size or spec["body"]["size_pt"])
            continue
        if part.startswith("*") and part.endswith("*") and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            _set_run_font(run, east=east, latin=latin, size=size, bold=bold, italic=True, color=color)
            continue
        run = paragraph.add_run(part)
        _set_run_font(run, east=east, latin=latin, size=size, bold=bold, color=color)


# ====================================================================
# 块级元素（均依赖 _CURRENT_SPEC）
# ====================================================================
def add_heading(doc, level, text):
    spec = _CURRENT_SPEC
    level = max(1, min(level, 6))
    h_spec = spec["headings"][f"h{level}"]
    p = doc.add_heading("", level=level)
    _apply_line_spacing(p.paragraph_format, spec["body"])
    add_markdown_inline(
        p, _strip_title_markers(text),
        east=h_spec["cn"], latin=spec["fonts"]["body_latin"],
        size=h_spec["size_pt"], bold=h_spec.get("bold", True),
        color=RGBColor(0, 0, 0),
    )
    return p


def add_title_paragraph(doc, text):
    """文档大标题（公文体：方正小标宋_GBK 2号 居中）。"""
    spec = _CURRENT_SPEC
    t = spec["title"]
    p = doc.add_paragraph()
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(t.get("align", "left"), WD_ALIGN_PARAGRAPH.LEFT)
    _apply_line_spacing(p.paragraph_format, t)
    add_markdown_inline(p, text,
                        east=t["cn"], latin=t.get("latin", spec["fonts"]["body_latin"]),
                        size=t["size_pt"], bold=t.get("bold", False))
    return p


def add_date_paragraph(doc, text):
    """日期行：右对齐、上空 N 行、右空 N 字（公文专属）。"""
    spec = _CURRENT_SPEC
    body = spec["body"]
    p = doc.add_paragraph()
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(spec["date"].get("align", "right"), WD_ALIGN_PARAGRAPH.RIGHT)
    _apply_line_spacing(p.paragraph_format, body)
    # 上空 N 行：用行距 × N 近似
    if spec["date"].get("margin_top_lines", 0) > 0:
        line_pt = body.get("line_spacing_fixed_pt") or body.get("size_pt", 12) * body.get("line_spacing_multiple", 1.5)
        p.paragraph_format.space_before = Pt(spec["date"]["margin_top_lines"] * line_pt)
    if spec["date"].get("padding_right_chars", 0) > 0:
        p.paragraph_format.right_indent = Pt(spec["date"]["padding_right_chars"] * body.get("size_pt", 12))
    add_markdown_inline(p, text, size=body["size_pt"])
    return p


def add_attachment_paragraph(doc, text):
    """附件行：左空 2 字（公文专属）。"""
    spec = _CURRENT_SPEC
    body = spec["body"]
    p = doc.add_paragraph()
    if spec["attachment"].get("indent_chars", 0) > 0:
        p.paragraph_format.first_line_indent = Pt(spec["attachment"]["indent_chars"] * body.get("size_pt", 12))
    _apply_line_spacing(p.paragraph_format, body)
    add_markdown_inline(p, text, size=body["size_pt"])
    return p


def add_hr(doc):
    spec = _CURRENT_SPEC
    p = doc.add_paragraph()
    _border_paragraph(p, edge="bottom", color="808080", sz="6")
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    return p


def add_code_block(doc, code_lines):
    spec = _CURRENT_SPEC
    cb = spec["code_block"]
    fonts = spec["fonts"]
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.0
    pf.left_indent = Inches(0.1)
    _shade_paragraph(p, cb["fill"])
    for idx, line in enumerate(code_lines):
        if idx > 0:
            br = p.add_run()
            br.add_break()
        run = p.add_run(line.replace("\t", "    "))
        _set_run_font(run, east=fonts["body_cn"], latin=fonts["code"], size=cb["size_pt"])
    return p


def add_quote_block(doc, lines):
    spec = _CURRENT_SPEC
    q = spec["quote"]
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Inches(0.3)
    pf.space_before = Pt(2)
    pf.space_after = Pt(2)
    _border_paragraph(p, edge="left", color=q["border_color"], sz="16")
    for idx, (depth, content) in enumerate(lines):
        if idx > 0:
            br = p.add_run()
            br.add_break()
        add_markdown_inline(
            p, content, east=spec["fonts"]["quote_cn"], latin=spec["fonts"]["body_latin"],
            size=q["size_pt"], color=RGBColor.from_string(q["color"]),
        )
    return p


def add_numbered_list(doc, items):
    spec = _CURRENT_SPEC
    body_size = spec["body"]["size_pt"]
    for counter, text in enumerate(items, start=1):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Inches(0.5)
        pf.first_line_indent = Inches(-0.25)
        pf.tab_stops.add_tab_stop(Inches(0.5))
        add_markdown_inline(p, f"{counter}.\t", size=body_size)
        add_markdown_inline(p, text, size=body_size)


def add_image(doc, md_dir, alt, path):
    spec = _CURRENT_SPEC
    IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".emf", ".wmf", ".tif", ".tiff"}
    MAX_W = 6.0
    MAX_H = 8.5
    path = path.strip().strip('"').strip("'")
    if not os.path.isabs(path):
        path = os.path.join(md_dir, path)
    if os.path.isfile(path) and os.path.splitext(path)[1].lower() in IMAGE_EXT:
        width_in = MAX_W
        try:
            from PIL import Image as _PIL
            with _PIL.open(path) as im:
                w_px, h_px = im.size
                scale = 1.0
                if h_px / 96.0 > MAX_H:
                    scale = min(scale, MAX_H / (h_px / 96.0))
                if w_px / 96.0 > MAX_W:
                    scale = min(scale, MAX_W / (w_px / 96.0))
                width_in = (w_px / 96.0) * scale
        except Exception:
            pass
        doc.add_picture(path, width=Inches(width_in))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return True
    p = doc.add_paragraph()
    add_markdown_inline(p, f"[图片未能插入：{path or alt}]", color=RGBColor(0xC0, 0x00, 0x00))
    return False


def add_page_number_footer(doc):
    """添加页码页脚。spec.page_number.format 支持 "{N}" 占位符，公文"-N-"即 prefix/suffix。"""
    spec = _CURRENT_SPEC
    pn = spec["page_number"]
    fmt = pn.get("format", "-{N}-")
    prefix, suffix = "", ""
    if "{N}" in fmt:
        prefix, suffix = fmt.split("{N}", 1)

    footer = doc.sections[0].footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(pn.get("align", "center"), WD_ALIGN_PARAGRAPH.CENTER)

    def _style_run(r):
        rPr = OxmlElement("w:rPr")
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:ascii"), spec["fonts"]["body_latin"])
        rFonts.set(qn("w:hAnsi"), spec["fonts"]["body_latin"])
        rFonts.set(qn("w:eastAsia"), pn["font_cn"])
        rPr.append(rFonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(pn["size_pt"] * 2)))
        rPr.append(sz)
        r.insert(0, rPr)

    def _literal(text):
        r = OxmlElement("w:r")
        _style_run(r)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        return r

    if prefix:
        p._p.append(_literal(prefix))
    r_field = OxmlElement("w:r")
    _style_run(r_field)
    fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE \\* MERGEFORMAT"
    fs = OxmlElement("w:fldChar"); fs.set(qn("w:fldCharType"), "separate")
    fe = OxmlElement("w:fldChar"); fe.set(qn("w:fldCharType"), "end")
    for el in (fb, instr, fs, fe):
        r_field.append(el)
    p._p.append(r_field)
    if suffix:
        p._p.append(_literal(suffix))


# ====================================================================
# 表格
# ====================================================================
SEP_RE = re.compile(r"^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$")


def is_separator_row(line):
    s = line.strip()
    if not s:
        return False
    return SEP_RE.match(s) is not None


def _cell_paragraph(cell, text, first=True):
    spec = _CURRENT_SPEC
    size = spec["table"]["size_pt"]
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    if first:
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
    if text.strip():
        add_markdown_inline(p, text, size=size)
    return p


def convert_table(doc, table_lines):
    spec = _CURRENT_SPEC
    t_spec = spec["table"]
    if len(table_lines) < 2:
        return False
    if not is_separator_row(table_lines[1]):
        return False

    def split_row(line):
        s = line.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|") and not s.endswith("\\|"):
            s = s[:-1]
        cells, buf, esc = [], [], False
        for ch in s:
            if esc:
                buf.append(ch); esc = False
            elif ch == "\\":
                esc = True
            elif ch == "|":
                cells.append("".join(buf).strip()); buf = []
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

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    header_align = align_map.get(t_spec["header_align"], WD_ALIGN_PARAGRAPH.CENTER)

    for i in range(ncols):
        cell = table.rows[0].cells[i]
        text = header[i] if i < len(header) else ""
        _cell_paragraph(cell, text, first=True)
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = header_align
        _shade_paragraph(cell.paragraphs[0], t_spec["header_fill"])

    for r_idx, row in enumerate(data_rows):
        for c_idx in range(ncols):
            text = row[c_idx] if c_idx < len(row) else ""
            _cell_paragraph(table.rows[r_idx + 1].cells[c_idx], text, first=True)
    return True


# ====================================================================
# 主转换流程
# ====================================================================
ATX_RE = re.compile(r"^(#{1,6})(?:\s+(.*?))?\s*$")
FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})\s*([\w+#.\-]*)\s*$")
HR_RE = re.compile(r"^\s*((?:-\s*){3,}|(?:\*\s*){3,}|(?:_\s*){3,})$")
UL_RE = re.compile(r"^\s*[-*+]\s+(.+)$")
OL_RE = re.compile(r"^\s*(\d{1,9})[.)]\s+(.+)$")
QUOTE_RE = re.compile(r"^(\s*>+)")
IMG_BLOCK_RE = re.compile(r"^!\[([^\]]*)\]\(\s*([^)\s]+)\s*\)\s*$")
DATE_BLOCK_RE = re.compile(r"^\s*<div\s+class=[\"']date[\"']>(.+?)</div>\s*$")
ATT_BLOCK_RE = re.compile(r"^\s*<div\s+class=[\"']attachment[\"']>(.+?)</div>\s*$")


def _read_text(input_file):
    for enc in ("utf-8-sig", "gb18030"):
        try:
            with open(input_file, "r", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(input_file, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def _setup_document(doc):
    """根据 spec 设置页面、Normal 样式。"""
    global _CURRENT_SPEC
    spec = _CURRENT_SPEC
    sec = doc.sections[0]
    sec.page_width = Cm(spec["page"]["width_cm"])
    sec.page_height = Cm(spec["page"]["height_cm"])
    sec.top_margin = Cm(spec["page"]["margin_top_cm"])
    sec.bottom_margin = Cm(spec["page"]["margin_bottom_cm"])
    sec.left_margin = Cm(spec["page"]["margin_left_cm"])
    sec.right_margin = Cm(spec["page"]["margin_right_cm"])

    normal = doc.styles["Normal"]
    fonts = spec["fonts"]
    normal.font.name = fonts["body_latin"]
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), fonts["body_cn"])
    normal.font.size = Pt(spec["body"]["size_pt"])
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    normal.paragraph_format.line_spacing = spec["body"].get("line_spacing_multiple", 1.5)
    chars = spec["body"].get("first_line_indent_chars", 0)
    if chars:
        normal.paragraph_format.first_line_indent = Pt(chars * spec["body"]["size_pt"])


def markdown_to_docx(input_file, output_file, style=None, title=None):
    global _CURRENT_SPEC
    _CURRENT_SPEC = load_style(style)
    spec = _CURRENT_SPEC

    content = _read_text(input_file)
    # frontmatter title 优先（用户 --title 显式传则覆盖）
    if title is None:
        title = _parse_frontmatter_title(content)
    if title is None:
        title = os.path.splitext(os.path.basename(input_file))[0]

    doc = Document()
    _setup_document(doc)
    add_title_paragraph(doc, title)

    lines = content.splitlines()
    # 跳过 YAML frontmatter
    if lines and lines[0].strip() == "---":
        i = 1
        while i < len(lines) and lines[i].strip() != "---":
            i += 1
        i += 1
    else:
        i = 0
    md_dir = os.path.dirname(os.path.abspath(input_file))
    n = len(lines)
    quote_lines = []
    ol_items = []

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

        if not stripped:
            flush_quote(); flush_ol(); i += 1; continue

        # 公文专属：日期 / 附件（HTML 块）
        md = DATE_BLOCK_RE.match(stripped)
        if md:
            flush_quote(); flush_ol()
            add_date_paragraph(doc, md.group(1))
            i += 1; continue
        ma = ATT_BLOCK_RE.match(stripped)
        if ma:
            flush_quote(); flush_ol()
            add_attachment_paragraph(doc, ma.group(1))
            i += 1; continue

        m = FENCE_RE.match(line)
        if m:
            flush_quote(); flush_ol()
            fence_char = m.group(1)[0]
            code_buffer = []
            i += 1
            while i < n:
                m2 = FENCE_RE.match(lines[i])
                if m2 and m2.group(1)[0] == fence_char:
                    i += 1; break
                code_buffer.append(lines[i]); i += 1
            add_code_block(doc, code_buffer); continue

        m = ATX_RE.match(stripped)
        if m:
            flush_quote(); flush_ol()
            level = len(m.group(1))
            text = (m.group(2) or "").strip()
            if text:
                add_heading(doc, level, text)
            i += 1; continue

        if HR_RE.match(stripped):
            flush_quote(); flush_ol(); add_hr(doc); i += 1; continue

        if stripped.startswith("|"):
            flush_quote(); flush_ol()
            table_lines = []
            while i < n and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip()); i += 1
            if not convert_table(doc, table_lines):
                for tl in table_lines:
                    p = doc.add_paragraph()
                    add_markdown_inline(p, tl, size=spec["body"]["size_pt"])
            continue

        m = IMG_BLOCK_RE.match(stripped)
        if m:
            flush_quote(); flush_ol()
            add_image(doc, md_dir, m.group(1), m.group(2)); i += 1; continue

        m = QUOTE_RE.match(line)
        if m:
            flush_ol()
            depth = m.group(1).count(">")
            content = line[m.end():]
            if content.startswith(" "):
                content = content[1:]
            quote_lines.append((depth, content)); i += 1; continue

        m = UL_RE.match(line)
        if m:
            flush_quote(); flush_ol()
            p = doc.add_paragraph(style="List Bullet")
            add_markdown_inline(p, m.group(1).strip(), size=spec["body"]["size_pt"])
            i += 1; continue

        m = OL_RE.match(line)
        if m:
            flush_quote()
            ol_items.append((int(m.group(1)), m.group(2).strip())); i += 1; continue

        flush_quote(); flush_ol()
        p = doc.add_paragraph()
        add_markdown_inline(p, stripped, size=spec["body"]["size_pt"])
        i += 1

    flush_quote(); flush_ol()

    if spec["page_number"]["enabled"]:
        add_page_number_footer(doc)

    out_dir = os.path.dirname(os.path.abspath(output_file))
    os.makedirs(out_dir, exist_ok=True)
    doc.save(output_file)
    print(f"转换完成：{output_file}")


def main():
    parser = argparse.ArgumentParser(
        description="将 Markdown(.md) 文件转换为排版规范的 Word(.docx) 文档",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例：
  python md_to_docx.py 文档.md                              # 默认风格（A4 宋体小四 黑体标题）
  python md_to_docx.py 文档.md -o 输出.docx --style styles/default.md
  python md_to_docx.py 公文.md -o 公文.docx --style styles/gov-doc.md --title "通知标题"
        """,
    )
    parser.add_argument("input", help="输入的 Markdown 文件路径")
    parser.add_argument("-o", "--output",
                        help="输出的 DOCX 文件路径（默认与输入文件同名、同目录）")
    parser.add_argument("--style", default=None,
                        help="风格 md 文件路径（含 ```yaml``` 块），缺省用内置默认风格")
    parser.add_argument("--title", default=None,
                        help="文档大标题（覆盖 frontmatter 与文件名）")
    args = parser.parse_args()

    input_file = args.input
    if not os.path.isfile(input_file):
        print(f"错误：找不到输入文件 {input_file}", file=sys.stderr)
        sys.exit(1)
    output_file = args.output or (os.path.splitext(input_file)[0] + ".docx")
    if not output_file.lower().endswith(".docx"):
        print(f"警告：输出文件建议使用 .docx 后缀：{output_file}")

    style_path = Path(args.style).resolve() if args.style else None
    markdown_to_docx(input_file, output_file, style=style_path, title=args.title)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    main()