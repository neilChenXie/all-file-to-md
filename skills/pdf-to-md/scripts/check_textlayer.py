# -*- coding: utf-8 -*-
"""检测 PDF 是否带有文本层。

用法:
    python -X utf8 check_textlayer.py <pdf路径> [--json]

判定: 全文档可提取字符 > 100 且 非零文本页占比 >= 50% 视为有文本层。
退出码: 0 = 有文本层, 1 = 无文本层, 2 = 出错。
依赖: PyMuPDF (pip install pymupdf)
"""
import json
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    print('缺少依赖 PyMuPDF，请先安装: pip install pymupdf', file=sys.stderr)
    sys.exit(2)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    use_json = '--json' in sys.argv
    if not args:
        print('用法: python -X utf8 check_textlayer.py <pdf路径> [--json]', file=sys.stderr)
        sys.exit(2)
    pdf_path = args[0]
    doc = fitz.open(pdf_path)

    page_chars = []
    fonts = set()
    for page in doc:
        txt = page.get_text().strip()
        page_chars.append(len(txt))
    for i in range(min(5, doc.page_count)):
        for f in doc[i].get_fonts():
            fonts.add(f[3])

    total = sum(page_chars)
    nonempty = sum(1 for c in page_chars if c > 0)
    zero_pages = [i + 1 for i, c in enumerate(page_chars) if c == 0]
    has_textlayer = total > 100 and nonempty >= doc.page_count * 0.5

    result = {
        'pdf': pdf_path,
        'pages': doc.page_count,
        'total_chars': total,
        'nonempty_pages': nonempty,
        'zero_text_pages': zero_pages,
        'sample_fonts': sorted(fonts)[:8],
        'has_textlayer': has_textlayer,
    }
    if use_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print('PDF: %s' % pdf_path)
        print('页数: %d | 可提取字符总数: %d | 非零文本页: %d/%d' % (
            doc.page_count, total, nonempty, doc.page_count))
        if zero_pages:
            print('零文本页: %s' % zero_pages)
        print('抽样字体: %s' % ', '.join(sorted(fonts)[:8]) if fonts else '无内嵌字体')
        print('结论: %s' % ('有文本层，可直接提取，无需 OCR' if has_textlayer
                           else '无文本层（或文本极少），属扫描件，应走 OCR 方案'))
    sys.exit(0 if has_textlayer else 1)


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    main()
