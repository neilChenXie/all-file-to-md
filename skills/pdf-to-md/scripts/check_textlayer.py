# -*- coding: utf-8 -*-
"""检测 PDF 是否带有文本层，并检测文本层内容是否为乱码。

用法:
    python -X utf8 check_textlayer.py <pdf路径> [--json]

判定: 全文档可提取字符 > 100 且 非零文本页占比 >= 50% 视为有文本层。
      文本层存在但内容乱码（字体缺 ToUnicode 映射/编码错乱等）时不可直提，
      应放弃文本层提取、改走图片 OCR 路线。
退出码: 0 = 有文本层且可用, 1 = 无文本层, 2 = 出错, 3 = 有文本层但为乱码。
依赖: PyMuPDF (pip install pymupdf)
"""
import json
import re
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    print('缺少依赖 PyMuPDF，请先安装: pip install pymupdf', file=sys.stderr)
    sys.exit(2)

# 0x80~0xFF 区间在正常文档中也常见的合法符号（不计入乱码）
LAT1_KEEP = set('·×÷±°§µ¥£©®')
# 字体缺映射时 PyMuPDF 可能输出的 (cid:N) 占位符
CID_RE = re.compile(r'\(cid:\s*\d+\s*\)')
# 单一字符高重复判定时排除的中文常用标点（ASCII 标点已由 ord>127 条件排除）
TOP_CH_KEEP = set('，。、；：？！…—－·（）《》“”‘’【】')


def garbled_signals(text):
    """评估文本层文本的乱码特征，返回 (suspect_ratio, reasons)。

    乱码典型成因是字体缺 ToUnicode 映射或编码转换错误，特征：
    1. 无效字符占比高：替换符 U+FFFD、私用区字符、控制符、(cid:N) 占位符；
    2. Latin-1 补充区（U+0080~U+00FF）字符占比高（如 "ä¸æ–‡"），
       正常中英文文档该区间字符占比极低；
    3. 单一非 ASCII 字符高度重复：字体映射错乱时整段文本落到同一码位；
    4. 经典乱码串 "锟斤拷"。
    """
    s = ''.join(ch for ch in text if not ch.isspace())
    n = len(s)
    if n < 30:
        return 0.0, []
    reasons = []
    n_def = sum(1 for ch in s if ch == '\ufffd' or ord(ch) < 32
                or 0x80 <= ord(ch) <= 0x9F or 0xE000 <= ord(ch) <= 0xF8FF)
    n_def += sum(len(m) for m in CID_RE.findall(text))
    ratio_def = n_def / n
    if ratio_def >= 0.05:
        reasons.append('无效字符（替换符/私用区/(cid:)占位符）占比 %.0f%%' % (100 * ratio_def))
    n_lat1 = sum(1 for ch in s if 0x80 <= ord(ch) <= 0xFF and ch not in LAT1_KEEP)
    ratio_lat1 = n_lat1 / n
    if ratio_lat1 >= 0.25:
        reasons.append('Latin-1 补充区字符占比 %.0f%%（编码错乱特征）' % (100 * ratio_lat1))
    cnt = {}
    for ch in s:
        cnt[ch] = cnt.get(ch, 0) + 1
    top_ch, top_n = max(cnt.items(), key=lambda kv: kv[1])
    ratio_top = top_n / n
    if n >= 50 and ratio_top >= 0.40 and ord(top_ch) > 127 and top_ch not in TOP_CH_KEEP:
        reasons.append('单一字符 %s 占比 %.0f%%（字体映射错乱特征）' % (top_ch, 100 * ratio_top))
    if '锟斤拷' in text:
        reasons.append('出现经典乱码串“锟斤拷”')
    return max(ratio_def, ratio_lat1, ratio_top), reasons


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    use_json = '--json' in sys.argv
    if not args:
        print('用法: python -X utf8 check_textlayer.py <pdf路径> [--json]', file=sys.stderr)
        sys.exit(2)
    pdf_path = args[0]
    doc = fitz.open(pdf_path)

    page_chars = []
    page_texts = []
    fonts = set()
    for page in doc:
        txt = page.get_text().strip()
        page_chars.append(len(txt))
        page_texts.append(txt)
    for i in range(min(5, doc.page_count)):
        for f in doc[i].get_fonts():
            fonts.add(f[3])

    total = sum(page_chars)
    nonempty = sum(1 for c in page_chars if c > 0)
    zero_pages = [i + 1 for i, c in enumerate(page_chars) if c == 0]
    has_textlayer = total > 100 and nonempty >= doc.page_count * 0.5

    # --- 乱码检测：逐页评估，多数内容页乱码 -> 文本层不可直提（应走图片 OCR 路线）---
    garbled_pages = []
    evaluated = 0
    for i, txt in enumerate(page_texts):
        if len(''.join(txt.split())) < 30:
            continue
        evaluated += 1
        ratio, reasons = garbled_signals(txt)
        if reasons:
            garbled_pages.append({'page': i + 1,
                                  'suspect_ratio': round(ratio, 3),
                                  'reasons': reasons})
    garbled = has_textlayer and evaluated > 0 and len(garbled_pages) * 2 >= evaluated

    result = {
        'pdf': pdf_path,
        'pages': doc.page_count,
        'total_chars': total,
        'nonempty_pages': nonempty,
        'zero_text_pages': zero_pages,
        'sample_fonts': sorted(fonts)[:8],
        'has_textlayer': has_textlayer,
        'textlayer_garbled': garbled,
        'garbled_pages': [g['page'] for g in garbled_pages],
        'garbled_detail': garbled_pages[:8],
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
        if has_textlayer and garbled_pages:
            print('乱码页: %s（%d/%d 个有效内容页）' % (
                [g['page'] for g in garbled_pages], len(garbled_pages), evaluated))
            for g in garbled_pages[:3]:
                print('  - 第 %d 页: %s' % (g['page'], '；'.join(g['reasons'])))
        if garbled:
            print('结论: 有文本层但为乱码，不可直接提取 —— 放弃文本层提取，走图片 OCR 路线')
        elif has_textlayer:
            extra = ('；个别乱码页（%s）提取后需按 3.2 用图像识别回退'
                     % [g['page'] for g in garbled_pages]) if garbled_pages else ''
            print('结论: 有文本层，可直接提取，无需 OCR' + extra)
        else:
            print('结论: 无文本层（或文本极少），属扫描件，应走 OCR 方案')
    if garbled:
        sys.exit(3)
    sys.exit(0 if has_textlayer else 1)


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    main()
