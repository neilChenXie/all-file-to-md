#!/usr/bin/env python3
"""
HTML to Markdown converter.
Supports WPS/Word HTML exports (GB2312/GBK) and Pandoc HTML (UTF-8 standard HTML5).
"""

from bs4 import BeautifulSoup
import chardet
import re
import sys
import os


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


def convert_list(lst, ordered=False, level=0):
    result = []
    items = lst.find_all('li', recursive=False)

    for i, item in enumerate(items):
        indent = "  " * level
        if ordered:
            prefix = f"{indent}{i+1}. "
        else:
            prefix = f"{indent}- "

        text = item.get_text().strip()
        if text:
            result.append(prefix + text)

    return "\n".join(result) + "\n"


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


def html_to_markdown(html_path, md_path=None):
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
    return md_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python html_to_markdown.py <html_file> [output_md_file]")
        print("Example: python html_to_markdown.py document.html")
        print("      python html_to_markdown.py document.html output.md")
        sys.exit(1)

    html_path = sys.argv[1]
    md_path = sys.argv[2] if len(sys.argv) > 2 else None

    result = html_to_markdown(html_path, md_path)
    sys.exit(0 if result else 1)


if __name__ == '__main__':
    main()
