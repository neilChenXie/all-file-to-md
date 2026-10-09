#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
pdf_to_png.py —— 将用户指定的 PDF 文件逐页导出为 PNG 图片（pdf-img-to-md skill 步骤2.2 / pdf-text-to-md skill 步骤3 用）

用法:
    python pdf_to_png.py <PDF文件路径> [--dpi 200] [--start-page N] [--end-page N] [--output-dir <目录>]

默认输出目录:
    <PDF所在目录>/tmp_<清洗后PDF文件名>/        （即 skill 的【图片目录】；每个 PDF 独占
    一份，同一目录下多份 PDF 并行处理时只清理/写入各自的临时目录，互不干扰）

命名规则:
    <清洗后PDF文件名>_01.png、<清洗后PDF文件名>_02.png ...（页码宽度按总页数自动补零，
    如 150 页则命名为 _001.png ~ _150.png，满足 pdf-img-to-md skill 步骤3.1 的前缀判断规则）
    文件名中的空格及其他非法字符（\ / : * ? " < > | 及控制字符）替换为下划线 _。

依赖:
    PyMuPDF (pymupdf)        安装命令: pip install pymupdf

示例:
    python pdf_to_png.py D:/docs/report.pdf --dpi 300
    python pdf_to_png.py D:/docs/scan.pdf --dpi 300 --start-page 3 --end-page 8
"""
import argparse
import os
import re
import sys

try:
    if sys.stdout and hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if sys.stderr and hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import pymupdf as fitz  # PyMuPDF 新版本（推荐）
except ImportError:
    try:
        import fitz  # PyMuPDF 旧版本兼容
    except ImportError:
        print("[ERROR] 未检测到 PyMuPDF 库，无法导出图片。", file=sys.stderr)
        print("        请先安装依赖后再运行本脚本：", file=sys.stderr)
        print("        pip install pymupdf", file=sys.stderr)
        sys.exit(1)

# 文件名清洗规则（与 skill 的【清洗后输入文件名】一致）：空格及非法字符 → 下划线
_ILLEGAL_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f\x7f]')


def sanitize_name(name: str) -> str:
    return _ILLEGAL_CHARS.sub("_", name).replace(" ", "_")


def main() -> None:
    parser = argparse.ArgumentParser(description="将 PDF 文件逐页导出为 PNG 图片")
    parser.add_argument("pdf_path", help="输入 PDF 文件路径")
    parser.add_argument("--dpi", type=int, default=200,
                        help="渲染分辨率 DPI，扫描件建议 300（默认 200）")
    parser.add_argument("--start-page", type=int, default=1,
                        help="起始页码（含，从 1 开始），默认 1")
    parser.add_argument("--end-page", type=int, default=0,
                        help="结束页码（含），0 表示导出到最后一页（默认 0）")
    parser.add_argument("--output-dir", default="",
                        help="自定义输出目录（默认：<PDF所在目录>/tmp_<清洗后PDF文件名>/）")
    args = parser.parse_args()

    pdf_path = os.path.abspath(args.pdf_path)
    if not os.path.isfile(pdf_path):
        print(f"[ERROR] 文件不存在: {pdf_path}", file=sys.stderr)
        sys.exit(2)
    if not pdf_path.lower().endswith(".pdf"):
        print(f"[ERROR] 不是 PDF 文件（请提供 .pdf 后缀的文件）: {pdf_path}", file=sys.stderr)
        sys.exit(2)

    prefix = sanitize_name(os.path.splitext(os.path.basename(pdf_path))[0])

    out_dir = os.path.abspath(args.output_dir) if args.output_dir else \
        os.path.join(os.path.dirname(pdf_path), "tmp_" + prefix)
    os.makedirs(out_dir, exist_ok=True)

    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        print(f"[ERROR] 打开 PDF 失败（文件可能已损坏或非有效 PDF）: {exc}", file=sys.stderr)
        sys.exit(3)

    total = doc.page_count
    if total == 0:
        print("[ERROR] PDF 没有任何页面。", file=sys.stderr)
        doc.close()
        sys.exit(3)

    start = max(1, args.start_page)
    end = min(total, args.end_page) if args.end_page > 0 else total
    if start > end:
        print("[ERROR] 起始页码大于结束页码。", file=sys.stderr)
        doc.close()
        sys.exit(3)

    # 页码补零宽度：至少 2 位，总页数超过 99 时按实际位数
    width = max(2, len(str(total)))
    zoom = args.dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    exported = 0
    try:
        for page_no in range(start, end + 1):
            page = doc.load_page(page_no - 1)  # 0-based
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out_name = f"{prefix}_{str(page_no).zfill(width)}.png"
            pix.save(os.path.join(out_dir, out_name))
            exported += 1
    except Exception as exc:
        print(f"[ERROR] 导出图片过程中出错: {exc}", file=sys.stderr)
        doc.close()
        sys.exit(4)
    finally:
        doc.close()

    print(f"[OK] 已导出 {exported} 页 PNG（第 {start}~{end} 页，PDF 共 {total} 页，DPI={args.dpi}）")
    print(f"[OK] 输出目录: {out_dir}")
    print(f"[OK] 命名规则: {prefix}_NN.png")


if __name__ == "__main__":
    main()
