#!/usr/bin/env python3
"""
HTML to Markdown converter.
Supports WPS/Word HTML exports (GB2312/GBK) and Pandoc HTML (UTF-8 standard HTML5).

方案C 编号回填：--docx <docx路径> 时，解析 docx 的 numbering.xml/document.xml
还原每段真实编号，按「<li> 文本 ↔ docx 段落文本」对齐回填，修复 pandoc
丢失的多级编号祖先路径（详见 docx_numbering.py）；未命中的 <li> 回退
<ol start> 还原逻辑并计入警告。
"""

import argparse
import os
import re
import sys
from collections import deque

from bs4 import BeautifulSoup
import chardet

# --docx 提供时由 main() 填充：段落归一化文本 -> 真实编号队列（按文档顺序消费）
NUMBER_INDEX = None
_NUMBER_MISS = []  # 编号 <li> 未命中对齐时的文本片段（警告清单）


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
            text = str(child).strip()
            if text:
                parts.append(text)
    return ' '.join(parts).strip()


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
            text = str(child).strip()
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
            text = str(child).strip()
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
            text = child.get_text().strip()
            result.append("#" * level + " " + text + "\n\n")
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


# ---------------------------------------------------------------- 疑点清单

_TOKEN_RE = re.compile(r"[A-Za-z0-9/+·\-]{8,}")
_GLUE_RE = re.compile(r"[a-z][A-Z]|[A-Z][A-Z][a-z]|[A-Za-z]\d|\d[A-Za-z]")
_TAIL_EN_RE = re.compile(r"[A-Za-z]{2,}$")
_HEAD_EN_RE = re.compile(r"^[A-Za-z]{2,}")
_SKIP_LINE_RE = re.compile(r"^\s*(\||#|```|!\[|\[.*\]\(|[-*+]\s|>\s|\d+\.\s)")


def _checklist_items(content):
    """扫描转换产物，返回（疑似粘连, 疑似断行）两组条目。

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
    return glued, broken


def generate_checklist(content, md_path):
    """生成疑点清单文件（与 md 同目录、同名 + .疑点清单.md），返回路径。"""
    glued, broken = _checklist_items(content)
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

    body = soup.find('body')
    content = convert_element(body) if body else convert_element(soup)

    content = re.sub(r'\[if !supportLists\]', '', content)
    content = re.sub(r'\[endif\]', '', content)
    content = re.sub(r'StartFragment|EndFragment', '', content)
    content = re.sub(r'^(\d+(?:\.\d+)*)(\*\*)', r'\1 \2', content, flags=re.MULTILINE)
    content = re.sub(r'\*\*\*\*', '', content)
    content = re.sub(r'\n{5,}', '\n\n\n\n', content)
    content = '\n'.join(line.rstrip() for line in content.split('\n'))

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[OK] Conversion complete: {md_path}")
    print(f"[OK] File size: {len(content)} characters")

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
        cl_path, n_glue, n_break = generate_checklist(content, md_path)
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
