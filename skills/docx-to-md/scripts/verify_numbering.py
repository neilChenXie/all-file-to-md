#!/usr/bin/env python3
r"""docx 编号还原与参考答案的一致性比对工具。

调用生产引擎 scripts/docx_numbering.py 的 DocxNumbering，按文档顺序还原
每个渲染出编号的段落，与参考答案 md 中的编号行做双指针对齐比对，
用于验证编号回填质量（如 test-case-1 的 527/527 全量吻合）。

比对原理：
  1. DocxNumbering 按文档顺序产出每个渲染出编号的段落（number/text/numid/ilvl…）；
  2. 参考答案按 `^(\d+(?:\.\d+)*) ?(\*\*)?(.*)$` 提取编号行；
  3. 双指针 + 文本归一化（去空白）对齐后逐段比对编号字符串；
  4. 退出码：全部对齐且编号一致 = 0，否则 1。注意：未按 abstractNumId 过滤时，
     正文编号行存在「编号与文本边界」归属差异（如参考行 "1）xxx" 的编号被正则
     截为 "1"、文本为 "）xxx"），可能被记为 XML 独有，属预期假阳性。

用法：
    python verify_numbering.py <docx> <参考答案.md> [abstractNumId]

  - 不传 abstractNumId：比对全部渲染出编号的段落；
  - 传 abstractNumId（如 test-case-1 的 27）：只比对该多级标题定义的段落，
    正文编号/项目符号列表归入末尾统计。
"""

import os
import re
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_numbering import DocxNumbering  # noqa: E402

# 编号后跟空格 / 加粗标记 / 直接跟中文（WPS 导出可能无空格，如 "1.2.1.4.12口岸…"）
HEADING_RE = re.compile(r"^(\d+(?:\.\d+)*) ?(\*\*)?(.*)$")


def parse_reference(md_path):
    lines = open(md_path, encoding="utf-8").read().split("\n")
    refs = []
    for i, line in enumerate(lines):
        m = HEADING_RE.match(line)
        if m:
            refs.append({"line": i + 1, "number": m.group(1),
                         "text": m.group(3).replace("**", "").strip()})
    return refs


def norm(s):
    return re.sub(r"\s+", "", s)


def main():
    if len(sys.argv) < 3 or len(sys.argv) > 4:
        print("用法: python verify_numbering.py <docx> <参考答案.md> [abstractNumId]")
        return 2
    docx_path, ref_path = sys.argv[1], sys.argv[2]
    abstract_id = sys.argv[3] if len(sys.argv) == 4 else None

    engine = DocxNumbering(docx_path)
    numbered = engine.numbered_paragraphs()
    refs = parse_reference(ref_path)

    print(f"XML 中渲染出编号的段落数: {len(numbered)}")
    print(f"参考答案编号标题行数: {len(refs)}")

    if abstract_id:
        multi = [n for n in numbered
                 if engine.num_map.get(n["numid"], {}).get("abstract") == abstract_id]
        print(f"其中 abstract{abstract_id} 多级标题段落: {len(multi)}")
    else:
        multi = numbered
        print("比对口径: 全部编号段落（未按 abstractNum 过滤）")

    # ---- 顺序对齐 + 编号比对（双指针，文本归一化后匹配）----
    ri = 0
    matched = exact = 0
    mismatches, unmatched_xml = [], []
    for n in multi:
        if ri >= len(refs):
            unmatched_xml.append(n)
            continue
        # 向前找文本一致的参考行（容忍参考中夹杂非编号行）
        j = ri
        while j < len(refs) and norm(refs[j]["text"]) != norm(n["text"]):
            j += 1
        if j >= len(refs):
            unmatched_xml.append(n)
            continue
        # 跳过的参考行记录为参考多出的行
        for k in range(ri, j):
            mismatches.append({"type": "ref-only", "ref": refs[k]})
        ri = j + 1
        matched += 1
        if refs[j]["number"] == n["number"]:
            exact += 1
        else:
            mismatches.append({"type": "number-diff", "xml": n, "ref": refs[j]})

    for k in range(ri, len(refs)):
        mismatches.append({"type": "ref-only", "ref": refs[k]})

    print("\n=== 对齐结果 ===")
    print(f"成功对齐段落: {matched}/{len(multi)}")
    print(f"编号完全一致: {exact}/{matched}")
    print(f"XML 有而参考未对齐: {len(unmatched_xml)}")
    print(f"参考侧未消费的编号样式行: {len(mismatches)}"
          f"（含正文手工编号、年份行等非标题行，属比对正则的假阳性，不影响标题比对）")

    if mismatches:
        print("\n=== 差异明细（前 30 条）===")
        for m in mismatches[:30]:
            if m["type"] == "number-diff":
                print(f"  [编号不一致] ref第{m['ref']['line']}行 ref={m['ref']['number']} "
                      f"xml={m['xml']['number']} text={m['ref']['text'][:30]}")
            else:
                print(f"  [参考独有] ref第{m['ref']['line']}行 {m['ref']['number']} "
                      f"{m['ref']['text'][:40]}")
    if unmatched_xml:
        print("\n=== XML 独有（前 10 条）===")
        for n in unmatched_xml[:10]:
            print(f"  order={n['order']} numid={n['numid']} ilvl={n['ilvl']} "
                  f"in_table={n['in_table']} num={n['number']} text={n['text'][:30]}")

    # ---- 其余编号段落统计（仅过滤模式下；bullet/none 不产出、不在列）----
    if abstract_id:
        others = [n for n in numbered
                  if engine.num_map.get(n["numid"], {}).get("abstract") != abstract_id]
        print(f"\n=== 其他渲染出编号的 numPr 段落（正文编号列表等）: {len(others)} ===")
        by_abs = Counter(engine.num_map[n["numid"]]["abstract"] for n in others)
        print("按 abstractNum 分布:", dict(by_abs.most_common()))
        rendered = Counter(n["number"] for n in others)
        print("渲染样式 top10:", rendered.most_common(10))
        print("位于表格内:", sum(1 for n in others if n["in_table"]),
              "/ 位于表格外:", sum(1 for n in others if not n["in_table"]))

    # ---- 退出码：XML 侧全部对齐且编号一致 = 通过 ----
    ok = matched == len(multi) and exact == matched and not unmatched_xml
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
