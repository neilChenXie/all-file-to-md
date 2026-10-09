# -*- coding: utf-8 -*-
"""PDF 文本层 -> Markdown 全文提取（含表格还原）。

用法:
    python -X utf8 extract_pdf_to_md.py <pdf路径> [-o 输出.md] [--title-pats 正则 ...]

功能:
    - 逐页提取文本层（PyMuPDF，非 OCR），以 "<!-- 第 N 页 -->" 注释标注分页
    - 带框线表格还原为 Markdown 表格（find_tables + 伪表格过滤 + 覆盖校验）
    - 页脚（页码/打印日期/居中页码行）识别并移除，页码映射写入提取说明
    - 通用标题识别（不依赖特定领域词汇）：先按字号/加粗启发式，再按通用中文文档模式，
      章节标题从 h2 起逐级提升（##/###）
    - 版式特殊的文档可用 --title-pats 追加标题正则，无需修改代码
    - 乱码守卫：文本层内容为乱码（字体缺 ToUnicode 映射/编码错乱）时放弃直提——
      乱码页占有效内容页 ≥5% 时直接退出且不生成输出文件（加载 pdf-img-to-md skill 走图片 OCR 路线），
      个别乱码页（<5%）不入正文、仅留注释待按 SKILL 步骤3 用图像识别回退

退出码: 0 = 成功, 2 = 参数/依赖错误, 3 = 文本层为乱码已放弃直提。

依赖: PyMuPDF (pip install pymupdf)
"""
import argparse
import datetime
import os
import re
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    print('缺少依赖 PyMuPDF，请先安装: pip install pymupdf', file=sys.stderr)
    sys.exit(2)

BS = chr(92)   # 反斜杠
BR = '<br>'    # 单元格内换行

# --- 章节标题模式（整行、去空格归一化后匹配；通用中文文档模式，特殊文档用 --title-pats 追加）---
# 标题尾部字符类排除数字/点/逗号/句号，避免把目录条目（引导点+页码结尾）误判为标题
NO_TAIL = r'[^0-9.,，。]'
TITLE_SEC_PATS = [   # 提升为 ##
    re.compile(r'^第[一二三四五六七八九十百0-9]+[章节篇]' + NO_TAIL + r'{0,20}$'),
    re.compile(r'^(?:附录|附件)[0-9]{0,2}' + NO_TAIL + r'{0,20}$'),
    re.compile(r'^附表\d+[:：].{0,25}$'),
    re.compile(r'^(?:引言|前言|概述|总则|简介)$'),
]
TITLE_SUB_PATS = [   # 提升为 ###
    re.compile(r'^目录$'),
    re.compile(r'^第[一二三四五六七八九十百]+部分[^0-9]{0,15}$'),
    re.compile(r'^\d{1,2}．\S.{0,20}$'),
    re.compile(r'^\d{1,2}、\S.{0,20}$'),
]
# 正文行开头特征（命中则不做标题提升）
BODY_START_BAD = ('包括', '及', '和', '与', '指', '的', '或', '为', '按', '详见')
# 标题行不含句读：含逗号/句号/分号/问号/"详见"/"了"（正文叙述特征）的行视为正文
# （不含顿号——"1、xx" 编号小标题模式依赖它）
NO_TITLE_PUNCT = re.compile(r'[,，。；?!？!]|详见|了')
PAGE_NUM_RE = re.compile(r'第\s*(\d+)\s*/\s*(\d+)\s*页')


def norm(s):
    return s.replace(' ', '').replace('\u3000', '')


def is_bold_font(font):
    """按字体名判断加粗：兼容 latin 名（Bold/Black/Heavy）与 mojibake 还原后的中文名（黑体/粗体）。"""
    if not font:
        return False
    fn = font.lower()
    if 'bold' in fn or 'black' in fn or 'heavy' in fn:
        return True
    try:
        fixed = font.encode('latin-1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        return False
    return '黑' in fixed or '粗' in fixed


# --- 文本层乱码检测（与 check_textlayer.py 中同名实现保持一致）---
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


def heading_level(line, body_size):
    """通用标题启发式（不依赖领域词汇），返回 2/3/None（章节标题从 h2 起逐级安排）：
    - 行字号显著大于正文众数字号（>= +1pt）-> ##；
    - 整行加粗且不小于正文字号的短行 -> ###。"""
    if body_size and line['size'] >= body_size + 1:
        return 2
    if body_size and line['bold'] and line['size'] >= body_size and len(norm(line['txt'])) <= 30:
        return 3
    return None


def extract_page_lines(page):
    """提取一页文本行（含字号/加粗元数据），识别并移除页脚。
    返回 (kept_lines, docpage, doctotal)。"""
    d = page.get_text('dict')
    raw_lines = []
    for b in d.get('blocks', []):
        if b.get('type') != 0:
            continue
        for ln in b.get('lines', []):
            spans = [sp for sp in ln.get('spans', []) if sp.get('text', '').strip()]
            if not spans:
                continue
            txt = ''.join(sp['text'] for sp in ln.get('spans', [])).strip()
            if not txt:
                continue
            # 行主字号 = 字符数最多的 span 的字号；加粗占比 > 0.8 视为整行加粗
            main = max(spans, key=lambda sp: len(sp['text'].strip()))
            total_chars = sum(len(sp['text'].strip()) for sp in spans)
            bold_chars = sum(len(sp['text'].strip())
                             for sp in spans if is_bold_font(sp.get('font', '')))
            x0, y0, x1, y1 = ln['bbox']
            raw_lines.append({'y0': y0, 'x0': x0, 'x1': x1, 'y1': y1, 'txt': txt,
                              'size': main['size'],
                              'bold': total_chars > 0 and bold_chars / total_chars > 0.8})

    # --- 页脚移除（兼容"第x/y页...打印日期"与"打印日期...第x/y页"两种顺序，及居中纯页码行）---
    W, H = page.rect.width, page.rect.height
    docpage = None
    doctotal = None
    kept = []
    for L in raw_lines:
        txt = L['txt']
        m = PAGE_NUM_RE.search(txt)
        if m and ('打印日期' in txt):
            docpage = int(m.group(1)); doctotal = int(m.group(2))
            continue
        if re.match(r'^第\s*\d+\s*/\s*\d+\s*页$', txt):
            docpage = int(m.group(1)); doctotal = int(m.group(2))
            continue
        if txt.startswith('打印日期'):
            m2 = PAGE_NUM_RE.search(txt)
            if m2:
                docpage = int(m2.group(1)); doctotal = int(m2.group(2))
            continue
        # 居中纯页码行（如 "12"、"- 12 -"）：仅当位于页顶/页底 10% 边缘带且水平居中时移除
        if re.match(r'^[-—–.\s]*\d{1,4}[-—–.\s]*$', txt):
            cx, cy = (L['x0'] + L['x1']) / 2, (L['y0'] + L['y1']) / 2
            if 0.3 * W <= cx <= 0.7 * W and (cy < 0.10 * H or cy > 0.90 * H):
                docpage = int(re.search(r'\d{1,4}', txt).group())
                continue
        kept.append(L)
    return kept, docpage, doctotal


def escape_line(st):
    if re.match(r'^#{1,6}\s', st) or st.startswith('>') or st.startswith('|'):
        return BS + st
    core = st.strip()
    if len(core) >= 3 and re.match(r'^[-_*]+$', core):
        return BS + st
    return st


def clean_cell(c):
    if c is None:
        return ''
    c = c.strip()
    c = c.replace('|', BS + '|')
    c = c.replace('\n', BR)
    return c


def merge_continuation_rows(rows):
    """折行延续行合并：首列为空（序号列）的行是折行碎片。
    逐列按"前行该列非空且后行空 -> 碎片属于后行；反之属于前行"判定归属，
    同列多个碎片用 <br> 连接。"""
    ncol = len(rows[0])
    data = rows[1:]
    out = []
    i = 0
    while i < len(data):
        r = data[i]
        if r[0].strip():
            out.append(list(r))
            i += 1
            continue
        j = i
        while j < len(data) and not data[j][0].strip():
            j += 1
        prev_row = out[-1] if out else None
        next_row = data[j] if j < len(data) else None
        for ci in range(ncol):
            frags = [data[k][ci] for k in range(i, j) if data[k][ci].strip()]
            if not frags:
                continue
            frag_text = BR.join(frags)
            prev_has = prev_row is not None and ci < len(prev_row) and prev_row[ci].strip()
            next_has = next_row is not None and ci < len(next_row) and next_row[ci].strip()
            if prev_has and not next_has:
                target = next_row
            else:
                target = prev_row
            if target is not None and ci < len(target):
                target[ci] = (target[ci] + BR + frag_text) if target[ci] else frag_text
        i = j
    return [rows[0]] + out


def render_table(data):
    rows = [[clean_cell(c) for c in row] for row in data]
    rows = [r for r in rows if any(x for x in r)]
    if not rows:
        return None
    ncols = max(len(r) for r in rows)
    rows = [r + [''] * (ncols - len(r)) for r in rows]
    # 仅当首列在多数数据行非空（序号/标签列）时才启用延续行合并，避免误合并
    if len(rows) > 1:
        ratio = sum(1 for r in rows[1:] if r[0].strip()) / (len(rows) - 1)
        if ratio > 0.5:
            rows = merge_continuation_rows(rows)
    out = ['| ' + ' | '.join(rows[0]) + ' |']
    out.append('|' + '---|' * ncols)
    for r in rows[1:]:
        out.append('| ' + ' | '.join(r) + ' |')
    return out


def extract(pdf_path, out_path, extra_sec_pats=()):
    doc = fitz.open(pdf_path)
    n_titles = 0
    pages_out = []
    docpage_map = {}
    doc_total = None
    fallback_pages = []
    n_tables = 0

    # --- 乱码守卫：文本层为乱码时放弃直提（乱码多由字体缺 ToUnicode 映射/编码错乱造成）---
    garbled_pages = {}   # {pno: reasons}
    evaluated = 0
    for i, page in enumerate(doc):
        txt = page.get_text()
        if len(''.join(txt.split())) < 30:
            continue
        evaluated += 1
        _, reasons = garbled_signals(txt)
        if reasons:
            garbled_pages[i + 1] = reasons
    if evaluated and len(garbled_pages) / evaluated >= 0.05:
        print('检测到文本层为乱码（乱码页占有效内容页 ≥5%%）：%d/%d 个有效内容页命中乱码特征（如第 %s 页），'
              % (len(garbled_pages), evaluated,
                 '、'.join(map(str, sorted(garbled_pages)[:5]))), file=sys.stderr)
        for p in sorted(garbled_pages)[:3]:
            print('  第 %d 页: %s' % (p, '；'.join(garbled_pages[p])), file=sys.stderr)
        print('已放弃文本层提取，未生成输出文件。请加载 pdf-img-to-md skill 走图片 OCR 路线：'
              'python -X utf8 pdf_to_png.py <pdf> 导出 PNG 后交多模态子agent识别。', file=sys.stderr)
        sys.exit(3)

    # --- 预扫描：逐页提取文本行（含元数据、移除页脚），统计正文众数字号 ---
    page_rows = []   # (pno, page, kept_lines, docpage)
    size_weight = {}
    for i, page in enumerate(doc):
        kept, docpage, doctotal = extract_page_lines(page)
        if doctotal:
            doc_total = doctotal
        if docpage:
            docpage_map[i + 1] = docpage
        page_rows.append((i + 1, page, kept, docpage))
        if i + 1 in garbled_pages:
            continue   # 乱码页不参与正文字号统计
        for L in kept:
            key = round(L['size'], 1)
            size_weight[key] = size_weight.get(key, 0) + len(L['txt'])
    body_size = max(size_weight, key=size_weight.get) if size_weight else None

    for pno, page, kept, docpage in page_rows:
        if pno in garbled_pages:
            # 乱码页不直提：留注释，按 SKILL 步骤3 导出 PNG 后用图像识别回退
            pages_out.append('<!-- 第 ' + str(pno) + ' 页 -->')
            pages_out.append('')
            pages_out.append('<!-- 本页文本层乱码，已放弃直提（%s）：'
                             '请导出本页 PNG 后用多模态子agent识别，替换本注释 -->'
                             % '；'.join(garbled_pages[pno]))
            pages_out.append('')
            continue
        # --- 表格检测与过滤 ---
        # 条件: >=2行x2列且<=12列（超多列是无竖线正文被横线切割的伪表格）；
        #       非空列占比<=0.5且最长单元格<=60字（正文段落被框线围住的伪表格）；
        #       丢弃被更大表格包含的碎片
        tabs = page.find_tables()
        good = [t for t in tabs.tables if t.row_count >= 2 and 2 <= t.col_count <= 12]
        cache = {}
        tables = []
        for t in good:
            inside_bigger = any(
                t2 is not t and t2.bbox[0] <= t.bbox[0] and t2.bbox[1] <= t.bbox[1]
                and t2.bbox[2] >= t.bbox[2] and t2.bbox[3] >= t.bbox[3] for t2 in good)
            if inside_bigger:
                continue
            data = t.extract()
            if not data:
                continue
            ncols = max(len(r) for r in data)
            col_nonempty = [0] * ncols
            max_len = 0
            for r in data:
                for ci in range(min(ncols, len(r))):
                    if r[ci] and r[ci].strip():
                        col_nonempty[ci] += 1
                        max_len = max(max_len, len(r[ci].strip()))
            ratio = sum(1 for c in col_nonempty if c > 0) / ncols
            if ratio <= 0.5 and max_len <= 60:
                continue  # 伪表格：正文被框线/横线围住
            cache[id(t)] = data
            tables.append(t)

        # --- 覆盖校验：表格区域内的文本行必须都被某个单元格bbox覆盖
        #     （合并单元格延续格 extract() 返回 None 属正常，勿用字符数对比），否则整页回退 ---
        ok = True
        for t in tables:
            tx0, ty0, tx1, ty1 = t.bbox
            cell_boxes = t.cells
            for k in kept:
                cx, cy = (k['x0'] + k['x1']) / 2, (k['y0'] + k['y1']) / 2
                if tx0 - 1 <= cx <= tx1 + 1 and ty0 - 1 <= cy <= ty1 + 1:
                    if not any(c[0] - 1 <= cx <= c[2] + 1 and c[1] - 1 <= cy <= c[3] + 1
                               for c in cell_boxes):
                        ok = False
                        break
            if not ok:
                break
        if tables and not ok:
            fallback_pages.append(pno)
            tables = []

        # --- 过滤表格内文本行 ---
        kept2 = []
        for k in kept:
            in_tab = any(t.bbox[0] - 1 <= (k['x0'] + k['x1']) / 2 <= t.bbox[2] + 1
                         and t.bbox[1] - 1 <= (k['y0'] + k['y1']) / 2 <= t.bbox[3] + 1
                         for t in tables)
            if not in_tab:
                kept2.append(k)

        # --- 组装页面元素（按 y,x 排序，段间距>20pt 插空行）---
        elems = [(k['y0'], k['x0'], 0, k) for k in kept2]
        for t in tables:
            md = render_table(cache.get(id(t)))
            if md:
                elems.append((t.bbox[1], t.bbox[0], 1, md))
                n_tables += 1
        elems.sort(key=lambda e: (e[0], e[1]))

        out_lines = []
        prev_bottom = None
        for y, x, kind, payload in elems:
            if kind == 1:
                if out_lines and out_lines[-1] != '':
                    out_lines.append('')
                out_lines.extend(payload)
                out_lines.append('')
                prev_bottom = None
                continue
            k = payload
            if prev_bottom is not None and k['y0'] - prev_bottom > 20:
                if out_lines and out_lines[-1] != '':
                    out_lines.append('')
            prev_bottom = k['y1']
            txt = k['txt']
            n = norm(txt)
            promoted = False
            if any(pat.match(n) for pat in extra_sec_pats):
                # 用户显式指定的标题正则优先，且不受内置防误判约束
                out_lines.append('## ' + txt)
                n_titles += 1
                promoted = True
            elif (len(n) <= 45 and not n.startswith(BODY_START_BAD)
                    and not NO_TITLE_PUNCT.search(n)
                    and '..' not in n and '……' not in n):
                lvl = heading_level(k, body_size)
                if lvl is None:
                    if any(pat.match(n) for pat in TITLE_SEC_PATS):
                        lvl = 3
                    elif any(pat.match(n) for pat in TITLE_SUB_PATS):
                        lvl = 4
                if lvl == 2:
                    out_lines.append('## ' + txt)
                    n_titles += 1
                    promoted = True
                elif lvl == 3:
                    out_lines.append('### ' + txt)
                    promoted = True
            if not promoted:
                out_lines.append(escape_line(txt))
        while out_lines and out_lines[-1] == '':
            out_lines.pop()
        pages_out.append('<!-- 第 ' + str(pno) + ' 页 -->')
        pages_out.append('')
        pages_out.extend(out_lines)
        pages_out.append('')

    # --- 页码映射 ---
    offsets = set(dd - p for p, dd in docpage_map.items())
    uniform = len(offsets) == 1
    offset = list(offsets)[0] if offsets else None
    no_footer = [p for p in range(1, doc.page_count + 1) if p not in docpage_map]
    fname = os.path.basename(pdf_path)
    today = datetime.date.today().isoformat()

    header = []
    header.append('---')
    header.append('source: "' + fname + '"')
    header.append('extracted: ' + today)
    header.append('pdf_pages: ' + str(doc.page_count))
    header.append('doc_pages: ' + str(doc_total if doc_total else '?'))
    header.append('tables_restored: ' + str(n_tables))
    header.append('---')
    header.append('')
    header.append('# ' + os.path.splitext(fname)[0] + '（全文提取）')
    header.append('')
    header.append('> [!info] 提取说明')
    header.append('> - 来源：`' + fname + '`（' + str(doc.page_count) + ' 页，文本层直接提取，非 OCR）')
    header.append('> - 带完整框线的表格已还原为 Markdown 表格（共 ' + str(n_tables)
                  + ' 个，合并单元格以空单元格表示）；无框线版式仍为纯文本')
    header.append('> - 两栏版式按阅读顺序拼接，个别换行处文字顺序可能有细微出入')
    if uniform and offset:
        header.append('> - 页码映射：文档页码 = PDF 页码 ' + ('%+d' % offset)
                      + '（已用页脚逐页验证，' + str(len(docpage_map)) + '/' + str(doc.page_count)
                      + ' 页命中；无页脚页：' + ', '.join(map(str, no_footer)) + '）')
    header.append('> - 各页页脚（页码、打印日期等）已移除，其余正文内容全部保留')
    if garbled_pages:
        header.append('> - 文本层乱码回退页（未直提，需按步骤3 用图像识别补充）：'
                      + ', '.join(map(str, sorted(garbled_pages))))
    header.append('')
    content = '\n'.join(header) + '\n' + '\n'.join(pages_out)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(content)

    print('输出: %s | 大小 %.1f KB' % (out_path, len(content.encode('utf-8')) / 1024))
    print('还原表格数: %d | 校验回退页: %s | 乱码回退页: %s' % (
        n_tables, fallback_pages if fallback_pages else '无',
        sorted(garbled_pages) if garbled_pages else '无'))
    print('页脚命中: %d/%d | 章节标题: %d' % (len(docpage_map), doc.page_count, n_titles))


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description='PDF 文本层 -> Markdown 全文提取（含表格还原）')
    parser.add_argument('pdf', help='PDF 文件路径')
    parser.add_argument('-o', '--output', help='输出 .md 路径（默认与 PDF 同目录同名）')
    parser.add_argument('--title-pats', action='append', default=[], metavar='REGEX',
                        help='追加章节标题正则（可多次传入），命中的独立短行提升为 ##')
    args = parser.parse_args()
    out = args.output or os.path.splitext(args.pdf)[0] + '.md'
    extra = []
    for p in args.title_pats:
        try:
            extra.append(re.compile(p))
        except re.error as e:
            print('无效的 --title-pats 正则 %r: %s' % (p, e), file=sys.stderr)
            sys.exit(2)
    extract(args.pdf, out, extra)


if __name__ == '__main__':
    main()
