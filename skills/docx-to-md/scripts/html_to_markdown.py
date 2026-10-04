#!/usr/bin/env python3
"""
HTML to Markdown converter.
Supports WPS/Word HTML exports (GB2312/GBK) and Pandoc HTML (UTF-8 standard HTML5).

方案C 编号回填：--docx <docx路径> 时，解析 docx 的 numbering.xml/document.xml
还原每段真实编号，按「<li> 文本 ↔ docx 段落文本」对齐回填，修复 pandoc
丢失的多级编号祖先路径（详见 docx_numbering.py）；未命中的 <li> 回退
<ol start> 还原逻辑并计入警告。

WPS/Word HTML 增强（test-case-3 揭露）：
- 条件注释（<!--[if ...]>）与声明节点整节点丢弃，杜绝域代码泄漏为正文；
- 目录区（p.MsoToc1..9）解析并结构化为 md 链接列表（层级缩进 + 页码后缀）；
- 编号反哺：目录条目按 _Toc 书签锚点对齐正文标题（未命中退化为归一化文本
  比对），用目录缓存的真实编号回填正文标题并剥离 mso-list:Ignore 错误缓存；
- 文本节点保留原有空白（&nbsp; 转空格），修复跨标签空格丢失；
- 表格 run 拆分不再以空格连接，消除假空格；
- 图片行下一行为图注（图N/表N）时回填语义化 alt。
"""

import argparse
import os
import re
import sys
from collections import deque

from bs4 import BeautifulSoup, NavigableString
from bs4.element import CData, Comment, Declaration, ProcessingInstruction
import chardet

# --docx 提供时由 main() 填充：段落归一化文本 -> 真实编号队列（按文档顺序消费）
NUMBER_INDEX = None
_NUMBER_MISS = []  # 编号 <li> 未命中对齐时的文本片段（警告清单）

_TOC_PLACEHOLDER = "@@TOC_PLACEHOLDER@@"
_TOC_P_RE = re.compile(r"^MsoToc([1-9])$")
_TOC_ANCHOR_RE = re.compile(r"^#?_Toc\d+$", re.IGNORECASE)
_TOC_NUM_RE = re.compile(r"^(第[0-9一二三四五六七八九十百]+章|\d+(?:\.\d+)*)\s+(.+)$")
_TO_ANCHOR_RE = re.compile(r"^_Toc\d+$", re.IGNORECASE)
_PANDOC_TOC_MIN = 3  # 连续命中段落达到该数量才判定为 pandoc TOC


def _norm_ws(text):
    """文本节点空白规范化，区分两类空白：

    - 语义空格（&nbsp;/全角空格、run 内普通空格）→ 保留（图注「表11 题名」分隔、
      「单 位」对齐空格均为原文内容）；
    - 排版空白（标签间换行/缩进构成的纯空白节点）→ 返回空串丢弃，
      节点内换行折叠为空格（WPS 跨行 HTML 不得引入假空格）。
    """
    text = text.replace("\xa0", "\x00").replace("\u3000", "\x00")
    text = re.sub(r"[ \t]*\r?\n[ \t\r\n]*", " ", text)
    text = re.sub(r"\t+", " ", text)
    has_semantic = "\x00" in text
    text = text.replace("\x00", " ")
    if not has_semantic and not text.strip():
        return ""
    return text


def _td_content(td):
    parts = []
    for child in td.descendants:
        if child.name == 'img':
            src = child.get('src', '')
            alt = child.get('alt', '').strip()
            if src:
                if not alt:
                    alt = os.path.basename(src)
                parts.append(f"![{alt}]({src})")
        elif child.name is None:
            # run 拆分（如 <font>C</font><font>har*</font>）直接连接，
            # 不以空格拼接（test-case-2 假空格根源）；节点间真实空白由文本节点自身携带
            parts.append(_norm_ws(str(child)))
    return re.sub(r" {2,}", " ", "".join(parts)).strip()


def convert_table(table):
    tbody = table.find('tbody') or table
    trs = tbody.find_all('tr')
    if not trs:
        return ""

    num_rows = len(trs)
    grid = {}

    for row_idx, tr in enumerate(trs):
        col_idx = 0
        for td in tr.find_all(['td', 'th']):
            while (row_idx, col_idx) in grid:
                col_idx += 1
            colspan = int(td.get('colspan', 1))
            rowspan = int(td.get('rowspan', 1))
            text = _td_content(td)
            cell = {'text': text, 'colspan': colspan}
            for r in range(rowspan):
                for c in range(colspan):
                    grid[(row_idx + r, col_idx + c)] = cell
            col_idx += colspan

    if not grid:
        return ""

    max_cols = max(c for (_, c) in grid.keys()) + 1

    result_rows = []
    for row_idx in range(num_rows):
        final_row = []
        col_idx = 0
        while col_idx < max_cols:
            cell = grid.get((row_idx, col_idx))
            if cell:
                colspan = cell['colspan']
                final_row.append(cell['text'])
                for _ in range(colspan - 1):
                    final_row.append('')
                col_idx += colspan
            else:
                final_row.append('')
                col_idx += 1
        result_rows.append(final_row[:max_cols])

    md = []
    md.append("| " + " | ".join(result_rows[0]) + " |")
    md.append("|" + "|".join(["---" for _ in range(max_cols)]) + "|")
    for row in result_rows[1:]:
        md.append("| " + " | ".join(row) + " |")

    return "\n".join(md) + "\n"


def _render_inline(node):
    """渲染节点的行内内容为 markdown 片段。

    与 get_text() 的区别：保留 <strong>/<em>/<img>/<code> 等行内语义，
    跳过嵌套的 <ol>/<ul>（由 convert_list 递归处理，避免文字粘连）。
    """
    parts = []
    for child in node.children:
        if child.name is None:
            text = _norm_ws(str(child)).strip()
            if text:
                parts.append(text)
        elif child.name in ('ol', 'ul'):
            continue
        elif child.name in ('strong', 'b'):
            text = child.get_text().strip()
            if text:
                parts.append(f"**{text}**")
        elif child.name in ('em', 'i'):
            text = child.get_text().strip()
            if text:
                parts.append(f"*{text}*")
        elif child.name == 'img':
            src = child.get('src', '')
            alt = child.get('alt', '').strip()
            if src:
                if not alt:
                    alt = os.path.basename(src)
                parts.append(f"![{alt}]({src})")
        elif child.name == 'br':
            parts.append(" ")
        elif child.name == 'p':
            text = _render_inline(child).strip()
            if text:
                parts.append(text)
        else:
            text = child.get_text().strip()
            if text:
                parts.append(text)
    return " ".join(parts)


def build_number_index(docx_path):
    """解析 docx 生成 {段落归一化文本: [真实编号队列]}，按文档顺序消费。

    仅收录渲染出编号文本的段落（bullet/无编号段落天然不参与回填）。
    依赖 docx_numbering 引擎（需 lxml），延迟导入以便方案A 路径不强制依赖。
    """
    from docx_numbering import DocxNumbering  # noqa: PLC0415 -- 延迟导入
    index = {}
    for p in DocxNumbering(docx_path).numbered_paragraphs():
        index.setdefault(p["key"], deque()).append(p["number"])
    return index


def _li_plain_text(node):
    """li 的纯文本：去除 md 标记影响，排除嵌套列表与图片，用于编号对齐。

    与 docx 段落文本（w:t + tab/br 转空格）同构：HTML 实体已由 soup 解码，
    br 转空格，剩余空白差异由 normalize_text 消除。
    """
    parts = []
    for child in node.children:
        if child.name is None:
            parts.append(str(child))
        elif child.name in ('ol', 'ul', 'img'):
            continue
        elif child.name == 'br':
            parts.append(' ')
        else:
            parts.append(_li_plain_text(child))
    return "".join(parts)


def _take_number(item):
    """按文本对齐取出 docx 真实编号；--docx 未提供或未命中返回 None。"""
    if NUMBER_INDEX is None:
        return None
    from docx_numbering import normalize_text  # noqa: PLC0415 -- 延迟导入
    dq = NUMBER_INDEX.get(normalize_text(_li_plain_text(item)))
    if dq:
        return dq.popleft()
    return None


def convert_list(lst, ordered=False, level=0, numbers=None):
    """转换 <ul>/<ol> 为 markdown。

    - --docx 场景：按 <li> 文本对齐 docx 真实编号直接回填（多级编号祖先路径
      不受列表被打断影响）；
    - 未提供 --docx 或未命中时回退：尊重 <ol start="N">，有序列表按
      「祖先路径 + 本级序号」还原多级编号（如 1.1、1.1.1），无序列表按层级缩进；
    - 列表项内的加粗/斜体/图片通过 _render_inline 保留（get_text() 会将其丢弃）。
    """
    result = []
    items = lst.find_all('li', recursive=False)
    if not items:
        return ""

    start = 1
    if ordered:
        try:
            start = int(lst.get('start', 1) or 1)
        except (TypeError, ValueError):
            start = 1

    parent_numbers = numbers or []

    for i, item in enumerate(items):
        text = _render_inline(item)
        if ordered:
            taken = _take_number(item)
            if taken is not None:
                numbers_now = []
                prefix = taken + " "
            else:
                if NUMBER_INDEX is not None:
                    _NUMBER_MISS.append(text[:40] or "(空文本)")
                numbers_now = parent_numbers + [start + i]
                if len(numbers_now) == 1:
                    prefix = f"{numbers_now[0]}. "
                else:
                    # 多级编号写成 1.1、1.1.1 形式（与 Word/WPS 多级编号一致，不加行尾点）
                    prefix = ".".join(str(n) for n in numbers_now) + " "
        else:
            numbers_now = []
            prefix = "  " * level + "- "

        if text:
            result.append(prefix + text)

        # 嵌套列表（pandoc 将多级编号导出为 li 内嵌 ol/ul）
        for sub in item.find_all(['ol', 'ul'], recursive=False):
            sub_md = convert_list(sub, sub.name == 'ol', level + 1, numbers_now)
            if sub_md.strip():
                result.append(sub_md.strip("\n"))

    return "\n\n".join(result) + "\n"


def convert_element(elem):
    result = []
    for child in elem.children:
        if child.name is None:
            # 保留语义空白（&nbsp;）与节点间原有空格（修复「1.4.1政策法规」类粘连）
            text = _norm_ws(str(child))
            if text:
                result.append(text)
        elif child.name == 'p':
            para_text = convert_element(child).strip()
            if para_text:
                result.append(para_text + "\n\n")
        elif child.name == 'br':
            result.append("\n")
        elif child.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            level = int(child.name[1])
            text = re.sub(r"\s+", " ", child.get_text()).strip()
            line = "#" * level + " " + text
            anchor = child.get("data-toc-anchor")
            if anchor:
                # 目录跳转锚（Typora/VSCode 预览可定位；GitHub 静默忽略）
                line += f' <a id="{anchor}"></a>'
            result.append(line + "\n\n")
        elif child.name in ['b', 'strong']:
            text = child.get_text().strip()
            if text:
                result.append(f"**{text}**")
        elif child.name in ['i', 'em']:
            text = child.get_text().strip()
            if text:
                result.append(f"*{text}*")
        elif child.name == 'table':
            result.append(convert_table(child) + "\n")
        elif child.name in ['ul', 'ol']:
            ordered = child.name == 'ol'
            result.append(convert_list(child, ordered) + "\n")
        elif child.name == 'div':
            result.append(convert_element(child))
        elif child.name == 'section':
            result.append(convert_element(child))
        elif child.name == 'pre':
            code_text = child.get_text().rstrip()
            if code_text:
                result.append(f"```\n{code_text}\n```\n\n")
        elif child.name == 'code':
            text = child.get_text().strip()
            if text:
                result.append(f"`{text}`")
        elif child.name == 'img':
            src = child.get('src', '')
            alt = child.get('alt', '').strip()
            if src:
                if not alt:
                    alt = os.path.basename(src)
                result.append(f"![{alt}]({src})\n")
        elif child.name == 'span':
            text = convert_element(child)
            if text.strip():
                result.append(text)
        else:
            result.append(convert_element(child))
    return "".join(result)


# ------------------------------------------------- WPS/Word HTML 预处理（test-case-3）


def _clean_soup_nodes(soup):
    """整节点删除条件注释/声明/处理指令（域代码块、<![if !supportLists]> 等）。"""
    for node in soup.find_all(
        string=lambda s: isinstance(
            s, (Comment, Declaration, ProcessingInstruction, CData)
        )
    ):
        node.extract()


def _extract_toc_page(p):
    """从目录段落的 PAGEREF 域块（条件注释内）提取缓存页码。"""
    for s in p.find_all(string=lambda t: isinstance(t, Comment)):
        m = re.search(r"field-separator.*?>(\d+)<.*?field-end", str(s), re.S)
        if m:
            return m.group(1)
    return None


def _parse_toc_entries(soup):
    """目录解析入口：优先 Word/WPS 的 MsoToc 段落，未命中时尝试 pandoc TOC。

    均须在 _clean_soup_nodes 之前调用（MsoToc 页码在条件注释域块内）。
    条目：{level, anchor, number, title, page}；首个段落原位替换为占位段，
    供 convert 输出后在后处理阶段生成结构化目录块。
    """
    entries = _parse_mso_toc(soup)
    if not entries:
        entries = _parse_pandoc_toc(soup)
    return entries


def _make_toc_entry(level, anchor, text, page):
    nm = _TOC_NUM_RE.match(text)
    if nm:
        number, title = nm.group(1), nm.group(2).strip()
    else:
        number, title = None, text
    return {
        "level": level,
        "anchor": anchor,
        "number": number,
        "title": title,
        "page": page,
        "backfill": None,
    }


def _toc_placeholder(soup, first_p, placeholder_done):
    """首个目录段原位替换为占位段，其余移除。"""
    if placeholder_done:
        first_p.decompose()
    else:
        placeholder = soup.new_tag("p")
        placeholder.string = _TOC_PLACEHOLDER
        first_p.replace_with(placeholder)
    return True


def _parse_mso_toc(soup):
    """解析 Word/WPS 目录区（p.MsoToc1..9），返回条目列表并从正文中移除。"""
    entries = []
    placeholder_done = False
    for p in list(soup.find_all("p")):
        classes = [c for c in (p.get("class") or []) if isinstance(c, str)]
        m = None
        for c in classes:
            m = _TOC_P_RE.match(c)
            if m:
                break
        if not m:
            continue
        level = int(m.group(1))
        a = p.find("a", href=_TOC_ANCHOR_RE)
        anchor = None
        if a is not None:
            anchor = a.get("href", "").lstrip("#") or None
            entry_text = a.get_text()
        else:
            entry_text = "".join(
                s for s in p.find_all(string=True) if not isinstance(s, Comment)
            )
        page = _extract_toc_page(p)
        text = _norm_ws(entry_text).strip()
        entries.append(_make_toc_entry(level, anchor, text, page))
        placeholder_done = _toc_placeholder(soup, p, placeholder_done)
    return entries


def _collect_heading_anchors(soup):
    """收集正文中可作为目录跳转目标的锚点（标题 id 与 <a name/_Toc> 书签）。"""
    ids = set()
    for h in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6"]):
        for key in ("id",):
            v = h.get(key)
            if v:
                ids.add(v)
    return ids


def _parse_pandoc_toc(soup):
    """识别 pandoc 从 docx 导出的 TOC（docx 目录域缓存文本，编号为渲染正确值）。

    形态：连续 `<p><a href="#标题id">编号 标题 <span>页码</span></a></p>`，
    且 href 目标必须是正文标题 id（硬条件，防误判正文交叉引用段落）；
    连续命中 ≥3 条才判定为 TOC。层级由编号深度推断。
    """
    heading_ids = _collect_heading_anchors(soup)
    if not heading_ids:
        return []
    runs, current = [], []
    for p in list(soup.find_all("p")):
        a = p.find("a")
        hit = None
        if a is not None and p.get_text().strip() == a.get_text().strip():
            href = a.get("href", "")
            anchor = href[1:] if href.startswith("#") else None
            if anchor and anchor in heading_ids:
                spans = a.find_all("span")
                page = None
                if spans and spans[-1].get_text().strip().isdigit():
                    page = spans[-1].get_text().strip()
                    spans[-1].extract()
                hit = (anchor, page)
        if hit:
            current.append((p, hit[0], hit[1]))
        else:
            if len(current) >= _PANDOC_TOC_MIN:
                runs.append(current)
            current = []
    if len(current) >= _PANDOC_TOC_MIN:
        runs.append(current)
    if not runs:
        return []
    run = max(runs, key=len)
    entries = []
    placeholder_done = False
    for p, anchor, page in run:
        a = p.find("a")
        text = _norm_ws(a.get_text()).strip()
        level = 1
        nm = _TOC_NUM_RE.match(text)
        if nm:
            level = nm.group(1).count(".") + 1 if not nm.group(1).startswith("第") else 1
        entries.append(_make_toc_entry(level, anchor, text, page))
        placeholder_done = _toc_placeholder(soup, p, placeholder_done)
    return entries


def _backfill_headings(soup, entries):
    """TOC 编号反哺正文标题：锚点对齐优先，归一化文本比对兜底。

    锚点来源：标题内 `<a name/_id="_TocXXX">` 书签（WPS/Word）与标题 `id`
    属性（pandoc，值为标题文本）。命中的标题剥离 mso-list:Ignore 错误缓存
    编号，插入目录缓存的真实编号，并记录 data-toc-anchor 供输出行尾锚。
    返回（命中数, 未命中条目标题列表）。
    """
    headings = soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6"])
    anchor_map = {}
    title_map = {}
    for h in headings:
        for a in h.find_all("a"):
            nm = a.get("name") or a.get("id")
            if nm and _TO_ANCHOR_RE.match(nm):
                anchor_map.setdefault(nm, h)
        hid = h.get("id")
        if hid and hid not in anchor_map:
            anchor_map[hid] = h
        t = re.sub(r"\s+", "", h.get_text())
        if t:
            title_map.setdefault(t, []).append(h)
    matched, misses = 0, []
    used = set()
    for e in entries:
        if not e["number"]:
            continue
        h = anchor_map.get(e["anchor"]) if e["anchor"] else None
        if h is None:
            cands = title_map.get(re.sub(r"\s+", "", e["title"]), [])
            if len(cands) == 1:
                h = cands[0]
            else:
                misses.append(e["number"] + " " + e["title"])
                continue
        if id(h) in used:
            continue
        for sp in h.find_all("span", style=re.compile(r"mso-list:Ignore", re.I)):
            sp.decompose()
        cur = re.sub(r"\s+", "", h.get_text())
        num = re.sub(r"\s+", "", e["number"])
        rest = cur[len(num):] if cur.startswith(num) else None
        if rest is not None and (not rest or (not rest[0].isdigit() and rest[0] != ".")):
            # 标题已携带同编号（如 test-case-2 正文标题自带编号缓存），仅记录锚点不重复插入
            if e["anchor"]:
                h["data-toc-anchor"] = e["anchor"]
            e["backfill"] = "already"
            used.add(id(h))
            matched += 1
            continue
        h.insert(0, NavigableString(e["number"] + " "))
        if e["anchor"]:
            h["data-toc-anchor"] = e["anchor"]
        e["backfill"] = "anchor" if anchor_map.get(e["anchor"]) is h else "text"
        used.add(id(h))
        matched += 1
    return matched, misses


def _render_toc_block(entries):
    """渲染结构化目录块：按 MsoToc 层级缩进的 md 链接列表 + 页码后缀。"""
    lines = []
    for e in entries:
        indent = "  " * (e["level"] - 1)
        label = f"{e['number']} {e['title']}" if e["number"] else e["title"]
        page = f" · {e['page']}" if e["page"] else ""
        if e["anchor"]:
            lines.append(f"{indent}- [{label}](#{e['anchor']}){page}")
        else:
            lines.append(f"{indent}- {label}{page}")
    return "\n".join(lines)


def _strip_leading_ws(content):
    """行级清理：代码块外清除行首空白（4 空格会被渲染为代码块），保留列表/引用/表格缩进。"""
    lines = content.split("\n")
    out = []
    in_code = False
    for line in lines:
        if line.strip().startswith("```"):
            in_code = not in_code
            out.append(line)
            continue
        if in_code:
            out.append(line)
        elif re.match(r"^\s*(?:[-*+]|\d+\.|>|\|)", line):
            out.append(line.rstrip())
        else:
            out.append(re.sub(r"^\s+", "", line).rstrip())
    return "\n".join(out)


def _backfill_image_alts(content):
    """图片行紧跟图注（图N/表N + 文字）时，用图注文本回填图片 alt。"""
    lines = content.split("\n")
    img_re = re.compile(r"^!\[([^]]*)\]\(([^)]+)\)\s*$")
    cap_re = re.compile(r"^((?:图|表)\s*\d+\s*\S.*)$")
    for i, line in enumerate(lines):
        m = img_re.match(line)
        if not m:
            continue
        for j in range(i + 1, len(lines)):
            nxt = lines[j].strip()
            if not nxt:
                continue
            if not nxt.startswith("```"):
                cm = cap_re.match(nxt)
                if cm and not nxt.startswith(("#", "|", "![", "[")):
                    alt = cm.group(1).strip().replace("[", "（").replace("]", "）")
                    lines[i] = f"![{alt}]({m.group(2)})"
            break
    return "\n".join(lines)


# ---------------------------------------------------------------- 疑点清单

_TOKEN_RE = re.compile(r"[A-Za-z0-9/+·\-]{8,}")
_GLUE_RE = re.compile(r"[a-z][A-Z]|[A-Z][A-Z][a-z]|[A-Za-z]\d|\d[A-Za-z]")
_TAIL_EN_RE = re.compile(r"[A-Za-z]{2,}$")
_HEAD_EN_RE = re.compile(r"^[A-Za-z]{2,}")
_SKIP_LINE_RE = re.compile(r"^\s*(\||#|```|!\[|\[.*\]\(|[-*+]\s|>\s|\d+\.\s)")
_RESIDUE_RE = re.compile(r"PAGEREF|mso-|\[if |field-(?:begin|separator|end)")
_NUMGLUE_RE = re.compile(r"^\d+(?:\.\d+)+[^\s.\d]")
_FIGGLUE_RE = re.compile(r"^(?:图|表)\d+[^ 0-9\s]")


def _checklist_items(content):
    """扫描转换产物，返回（疑似粘连, 疑似断行, 残留/粘连行）三组条目。

    只检测不改写：粘连正误无法自动判定（iPhone/IoT 等专有名词会被误拆），
    输出清单供人工按 SKILL.md 步骤5 逐一修复。
    """
    glued = []
    seen = set()
    lines = content.split("\n")
    for lineno, line in enumerate(lines, 1):
        if _SKIP_LINE_RE.match(line):
            continue
        for tok in _TOKEN_RE.findall(line):
            if _GLUE_RE.search(tok) and tok not in seen:
                seen.add(tok)
                glued.append((lineno, tok, line.strip()[:60]))
    broken = []
    for i in range(len(lines) - 1):
        cur, nxt = lines[i].rstrip(), lines[i + 1].lstrip()
        if _SKIP_LINE_RE.match(cur) or _SKIP_LINE_RE.match(nxt):
            continue
        if _TAIL_EN_RE.search(cur) and _HEAD_EN_RE.match(nxt):
            broken.append((i + 1, cur[-40:], nxt[:20]))
    issues = []
    in_code = False
    for lineno, line in enumerate(lines, 1):
        if line.strip().startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        if _RESIDUE_RE.search(line):
            issues.append((lineno, "域代码/样式残留", line.strip()[:60]))
        elif _NUMGLUE_RE.match(line.strip()) and not _SKIP_LINE_RE.match(line):
            issues.append((lineno, "编号与文字粘连", line.strip()[:60]))
        elif _FIGGLUE_RE.match(line.strip()):
            issues.append((lineno, "图注/表题粘连", line.strip()[:60]))
    return glued, broken, issues


def generate_checklist(content, md_path, backfill_miss=None):
    """生成疑点清单文件（与 md 同目录、同名 + .疑点清单.md），返回路径。"""
    glued, broken, issues = _checklist_items(content)
    backfill_miss = backfill_miss or []
    out_path = re.sub(r"\.md$", "", md_path) + ".疑点清单.md"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("# 疑点清单（自动检测，仅供人工核查定位，未自动改写）\n\n")
        f.write(f"来源：{os.path.basename(md_path)}\n\n")
        f.write(f"## 疑似英文粘连（{len(glued)} 处）\n\n")
        f.write("长 ASCII 词内含 小写→大写 / 大写串→小写 / 字母↔数字 边界，"
                "注意排除 iPhone、IoT 等正常专有名词。\n\n")
        if glued:
            f.write("| 行号 | 词元 | 所在行 |\n|---|---|---|\n")
            for lineno, tok, ctx in glued:
                f.write(f"| {lineno} | `{tok}` | {ctx} |\n")
        else:
            f.write("（无）\n")
        f.write(f"\n## 疑似软换行断行（{len(broken)} 处）\n\n")
        f.write("行尾与下一行行首均为英文词，可能为同一段被 pandoc 拆行"
                "（如 HTTP / API 被拆开），确认后合并。\n\n")
        if broken:
            f.write("| 行号 | 行尾 | 下行行首 |\n|---|---|---|\n")
            for lineno, tail, head in broken:
                f.write(f"| {lineno} | …{tail} | {head}… |\n")
        else:
            f.write("（无）\n")
        f.write(f"\n## 域代码残留 / 编号粘连（{len(issues)} 处）\n\n")
        f.write("域代码残留说明条件注释清理未覆盖；编号/图注粘连建议对照原文补空格。\n\n")
        if issues:
            f.write("| 行号 | 类型 | 所在行 |\n|---|---|---|\n")
            for lineno, kind, ctx in issues:
                f.write(f"| {lineno} | {kind} | {ctx} |\n")
        else:
            f.write("（无）\n")
        f.write(f"\n## 编号反哺未命中（{len(backfill_miss)} 条）\n\n")
        f.write("目录条目按锚点/文本均未对齐到正文标题，正文编号未被回填，"
                "建议对照目录手工修正。\n\n")
        if backfill_miss:
            for item in backfill_miss:
                f.write(f"- {item}\n")
        else:
            f.write("（无）\n")
    return out_path, len(glued), len(broken)


def html_to_markdown(html_path, md_path=None, docx_path=None, checklist=False):
    """
    Convert an HTML file exported from Word to Markdown.

    Args:
        html_path: Path to the HTML input file
        md_path:   Path to the Markdown output file. If None, defaults to
                   replacing .html/.htm extension with .md.

    Returns:
        md_path on success, None on failure
    """
    if not os.path.exists(html_path):
        print(f"Error: File not found: {html_path}")
        return None

    if md_path is None:
        md_path = html_path.replace('.html', '.md').replace('.htm', '.md')

    # --docx：装配编号索引（方案C 编号回填；不传时方案A 行为完全不变）
    global NUMBER_INDEX, _NUMBER_MISS
    NUMBER_INDEX = None
    _NUMBER_MISS = []
    if docx_path:
        if not os.path.exists(docx_path):
            print(f"Error: docx 文件不存在: {docx_path}")
            return None
        NUMBER_INDEX = build_number_index(docx_path)
        print(f"[OK] 编号索引: {sum(len(v) for v in NUMBER_INDEX.values())} 段 "
              f"（来自 {os.path.basename(docx_path)}）")

    html = None
    encodings = []
    with open(html_path, "rb") as f:
        raw = f.read()
    detected = chardet.detect(raw)
    if detected and detected['encoding']:
        encodings.append(detected['encoding'].lower())
    encodings.extend(['utf-8', 'gb2312', 'gbk', 'gb18030'])

    for enc in encodings:
        try:
            html = raw.decode(enc)
            print(f"Using encoding: {enc}")
            break
        except (UnicodeDecodeError, LookupError):
            continue

    if html is None:
        print("Error: Unable to read file, encoding not supported")
        return None

    soup = BeautifulSoup(html, 'html.parser')

    for tag in soup(["script", "style", "meta", "link"]):
        tag.decompose()

    # WPS/Word HTML 预处理：目录解析（须先于域代码清理，页码在注释域块内）
    # + 条件注释/域代码清理 + 编号反哺
    toc_entries = _parse_toc_entries(soup)
    _clean_soup_nodes(soup)
    backfill_ok, backfill_miss = _backfill_headings(soup, toc_entries)

    body = soup.find('body')
    content = convert_element(body) if body else convert_element(soup)

    content = re.sub(r'\[if !supportLists\]', '', content)
    content = re.sub(r'\[endif\]', '', content)
    content = re.sub(r'StartFragment|EndFragment', '', content)
    content = re.sub(r'^(\d+(?:\.\d+)*)(\*\*)', r'\1 \2', content, flags=re.MULTILINE)
    content = re.sub(r'\*\*\*\*', '', content)
    content = re.sub(r'\n{5,}', '\n\n\n\n', content)
    content = _strip_leading_ws(content)

    # 目录占位段 → 结构化目录块（md 链接列表 + 页码后缀）
    if toc_entries:
        toc_block = _render_toc_block(toc_entries)
        content = re.sub(
            rf"^{re.escape(_TOC_PLACEHOLDER)}\s*$", toc_block, content,
            count=1, flags=re.MULTILINE)
        content = content.replace(_TOC_PLACEHOLDER, "")
    # 图片 alt 图注回填
    content = _backfill_image_alts(content)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[OK] Conversion complete: {md_path}")
    print(f"[OK] File size: {len(content)} characters")

    if toc_entries:
        print(f"[OK] 目录结构化: {len(toc_entries)} 条")
        numbered = sum(1 for e in toc_entries if e["number"])
        if numbered:
            msg = f"[OK] 编号反哺: {backfill_ok}/{numbered} 条命中"
            if backfill_miss:
                msg += f"（未命中 {len(backfill_miss)} 条，详见疑点清单）"
            print(msg)

    # 编号回填未命中的 <li> 警告（回退 <ol start> 逻辑，此处仅提示核对）
    if NUMBER_INDEX is not None:
        if _NUMBER_MISS:
            print(f"[WARN] 编号 <li> 未命中 docx 对齐 {len(_NUMBER_MISS)} 处"
                  "（已回退 start 属性还原，建议对照原文核对）:")
            for frag in _NUMBER_MISS[:10]:
                print(f"       - {frag}")
            if len(_NUMBER_MISS) > 10:
                print(f"       ... 等共 {len(_NUMBER_MISS)} 处")
        else:
            print("[OK] 编号回填全部命中")

    if checklist:
        cl_path, n_glue, n_break = generate_checklist(
            content, md_path, backfill_miss=backfill_miss)
        print(f"[OK] 疑点清单: {cl_path}（粘连 {n_glue} / 断行 {n_break}，仅供人工核查）")

    return md_path


def main():
    parser = argparse.ArgumentParser(
        description="HTML to Markdown converter（兼容 WPS/Word HTML 与 pandoc HTML）")
    parser.add_argument("html_file", help="输入的 html 文件路径")
    parser.add_argument("output_md_file", nargs="?", default=None,
                        help="输出的 md 文件路径（默认与 html 同目录同名）")
    parser.add_argument("--docx", metavar="DOCX", default=None,
                        help="方案C 编号回填：提供源 docx 路径，按编号引擎还原真实编号")
    parser.add_argument("--checklist", action="store_true",
                        help="生成英文粘连/断行疑点清单（.疑点清单.md，仅供人工核查）")
    args = parser.parse_args()

    result = html_to_markdown(args.html_file, args.output_md_file,
                              docx_path=args.docx, checklist=args.checklist)
    sys.exit(0 if result else 1)


if __name__ == '__main__':
    main()
