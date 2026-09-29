#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
crop_images.py —— 从 PDF/PNG 页面中裁剪/提取图片（pdf-to-md skill 步骤4.8 用）

三种模式：

1) 坐标裁剪模式（默认，配合多模态定位子agent输出的百分比坐标）：
    python crop_images.py <PDF/PNG路径> --page <页码> --bbox "left,top,right,bottom" [--output <文件.png>]

2) 整页嵌入图提取模式（仅PDF，图片是独立对象时）：
    python crop_images.py <PDF路径> --page <页码> --extract-embedded [--output-dir <目录>]

3) 批量整页渲染模式（仅PDF，无嵌入图、无坐标，直接按页出图）：
    python crop_images.py <PDF路径> --pages "1,3,5-7" --render-all --dpi 300 [--output-dir <目录>]

坐标说明:
    --bbox 使用**百分比坐标**（0~1，左上角为原点、x向右、y向下），与多模态子agent的输出格式一致。
    例如: --bbox "0.115,0.195,0.910,0.785"

裁剪质量策略（关键）:
    - 优先提取该页的**原始嵌入扫描图**（若该页由一张覆盖全页的图片构成，如扫描件），得到最高分辨率。
    - 无嵌入图（如矢量文字页）时，回退为按页面渲染后裁剪，保证可用性。
    两种途径坐标含义一致（均相对页面全幅），因此百分比坐标对扫描件和电子版均适用。

依赖:
    PyMuPDF (pymupdf)、Pillow     安装命令: pip install pymupdf pillow

示例:
    # PDF 坐标裁剪
    python crop_images.py D:/docs/scan.pdf --page 10 --bbox "0.115,0.195,0.910,0.785" --output fig1.png
    
    # PNG 坐标裁剪（直接裁剪【图片目录】下的页面图，目录名形如 tmp_<清洗后PDF文件名>）
    python crop_images.py "D:/docs/tmp_report/report_12.png" --bbox "0.115,0.195,0.910,0.785" --output fig1.png
    
    # PDF 嵌入图提取
    python crop_images.py D:/docs/ebook.pdf --page 3 --extract-embedded --output-dir D:/imgs
    
    # PDF 整页渲染
    python crop_images.py D:/docs/scan.pdf --pages "1-5" --render-all --dpi 300 --output-dir D:/imgs
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
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        print("[ERROR] 未检测到 PyMuPDF 库。请先安装：pip install pymupdf", file=sys.stderr)
        sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("[ERROR] 未检测到 Pillow 库。请先安装：pip install pillow", file=sys.stderr)
    sys.exit(1)


def parse_pages(spec: str) -> list:
    """解析 '1,3,5-7' 形式的页码串为页码列表（1-based）。"""
    pages = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)(?:-(\d+))?$", part)
        if not m:
            raise ValueError(f"无法解析页码段: {part!r}（支持如 1,3,5-7）")
        start = int(m.group(1))
        end = int(m.group(2)) if m.group(2) else start
        if end < start:
            raise ValueError(f"页码段 {part!r} 结束页小于起始页")
        pages.extend(range(start, end + 1))
    return sorted(set(pages))


def get_page_full_image(page, min_cover=0.9):
    """
    判断页面是否由一张覆盖全页的扫描图构成。
    若命中，返回 (PIL.Image, xref)；否则返回 (None, None)。

    实现要点：get_image_info() 在新版 PyMuPDF 中不含 xref 字段，
    因此命中整页图后需用 get_images(full=True) 回查 xref 以提取原始资源。
    """
    pw, ph = page.rect.width, page.rect.height
    page_area = pw * ph if pw * ph > 0 else 1

    # 候选 xref：优先按覆盖面积匹配
    info_list = page.get_image_info()
    candidates = []
    for it in info_list:
        x0, y0, x1, y1 = it["bbox"]
        area = (x1 - x0) * (y1 - y0)
        cover = area / page_area
        if cover >= min_cover:
            candidates.append((cover, area, it))
    if not candidates:
        return None, None
    # 取覆盖最大的一个
    _, _, info = max(candidates, key=lambda c: (c[0], c[1]))

    # 从 get_images(full=True) 回查 xref（顺序无关，取第一个可用的）
    xref = 0
    for im in page.get_images(full=True):
        if im[0]:
            xref = im[0]
            break
    if xref:
        try:
            raw = page.parent.extract_image(xref)
            if raw:
                bio = __import__("io").BytesIO(raw["image"])
                return Image.open(bio), xref
        except Exception:
            pass
    return None, xref


def render_page(page, dpi=300):
    """渲染整页为 PIL.Image（作为无嵌入图时的回退途径）。"""
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def crop_by_bbox(img: Image.Image, bbox, page_label: str) -> Image.Image:
    """按百分比坐标裁剪。bbox = (left, top, right, bottom)，0~1。"""
    w, h = img.size
    l, t, r, b = (float(v) for v in bbox)
    if not (0.0 <= l < r <= 1.0 and 0.0 <= t < b <= 1.0):
        raise ValueError(
            f"[ERROR] {page_label}: 坐标越界或无效 {bbox}（应为 0<=left<right<=1, 0<=top<bottom<=1）"
        )
    box = (int(w * l), int(h * t), int(w * r), int(h * b))
    cropped = img.crop(box)
    print(f"  {page_label}: 从 {img.size} 裁剪区域 {box} -> 输出 {cropped.size}")
    return cropped


def ensure_output_path(output: str, default_dir: str, default_name: str) -> str:
    if output:
        return output
    os.makedirs(default_dir, exist_ok=True)
    return os.path.join(default_dir, default_name)


def process_pdf_crop(args):
    """处理PDF文件的坐标裁剪"""
    pdf_path = os.path.abspath(args.input_path)
    
    out_dir = os.path.abspath(args.output_dir) if args.output_dir else \
        os.path.join(os.path.dirname(pdf_path), "图片提取")

    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        print(f"[ERROR] 打开 PDF 失败: {exc}", file=sys.stderr)
        sys.exit(3)

    if args.bbox:
        if not args.page:
            print("[ERROR] 坐标裁剪模式需要 --page", file=sys.stderr)
            sys.exit(2)
        bbox = tuple(v.strip() for v in args.bbox.split(","))
        if len(bbox) != 4:
            print(f"[ERROR] --bbox 需为 4 个逗号分隔值: {args.bbox}", file=sys.stderr)
            sys.exit(2)
        page = doc[args.page - 1]
        label = f"PDF页{args.page}"
        img, xref = get_page_full_image(page)
        source = "原始嵌入扫描图" if img is not None else "页面渲染(300dpi)"
        if img is None:
            img = render_page(page, dpi=max(args.dpi, 300))
        try:
            cropped = crop_by_bbox(img, bbox, label)
        except ValueError as exc:
            print(exc, file=sys.stderr)
            sys.exit(2)
        # 图名取自输出文件名
        out_path = ensure_output_path(
            args.output, out_dir,
            f"page{args.page}_{bbox[0]}_{bbox[1]}_{bbox[2]}_{bbox[3]}.png")
        cropped.save(out_path)
        print(f"  裁剪来源: {source}（页面共 {img.size[0]}x{img.size[1]}）")
        print(f"[OK] 已输出: {out_path}")
        doc.close()
        return

    # 模式2：提取嵌入图片对象
    if args.extract_embedded:
        if not args.page:
            print("[ERROR] 嵌入提取模式需要 --page", file=sys.stderr)
            sys.exit(2)
        page = doc[args.page - 1]
        imgs = page.get_images(full=True)
        if not imgs:
            print(f"  PDF页{args.page}: 无嵌入图片对象（可能为扫描件或矢量页）")
            doc.close()
            return
        os.makedirs(out_dir, exist_ok=True)
        for i, im in enumerate(imgs, 1):
            xref = im[0]
            try:
                raw = doc.extract_image(xref)
                ext = raw.get("ext", "png")
                fname = os.path.join(out_dir, f"page{args.page}_img{i}.{ext}")
                with open(fname, "wb") as f:
                    f.write(raw["image"])
                print(f"[OK] 嵌入图片 {i}: xref={xref} {raw['width']}x{raw['height']} -> {fname}")
            except Exception as exc:
                print(f"[WARN] 提取嵌入图片 {i} (xref={xref}) 失败: {exc}", file=sys.stderr)
        doc.close()
        return

    # 模式3：整页渲染
    if args.render_all:
        if not args.pages:
            print("[ERROR] 整页渲染模式需要 --pages", file=sys.stderr)
            sys.exit(2)
        pages = parse_pages(args.pages)
        os.makedirs(out_dir, exist_ok=True)
        for pno in pages:
            if pno < 1 or pno > doc.page_count:
                print(f"[WARN] 页码 {pno} 超出范围(1-{doc.page_count})，跳过", file=sys.stderr)
                continue
            img = render_page(doc[pno - 1], dpi=args.dpi)
            fname = os.path.join(out_dir, f"{args.prefix}_{pno:03d}.png")
            img.save(fname)
            print(f"[OK] 渲染 PDF页{pno} ({args.dpi}dpi) {img.size} -> {fname}")
        doc.close()
        return

    doc.close()


def process_png_crop(args):
    """处理PNG图片文件的坐标裁剪"""
    png_path = os.path.abspath(args.input_path)
    
    out_dir = os.path.abspath(args.output_dir) if args.output_dir else \
        os.path.join(os.path.dirname(png_path), "图片提取")

    if args.bbox:
        bbox = tuple(v.strip() for v in args.bbox.split(","))
        if len(bbox) != 4:
            print(f"[ERROR] --bbox 需为 4 个逗号分隔值: {args.bbox}", file=sys.stderr)
            sys.exit(2)
        
        try:
            img = Image.open(png_path)
        except Exception as exc:
            print(f"[ERROR] 打开 PNG 失败: {exc}", file=sys.stderr)
            sys.exit(3)
        
        label = f"PNG图片"
        try:
            cropped = crop_by_bbox(img, bbox, label)
        except ValueError as exc:
            print(exc, file=sys.stderr)
            sys.exit(2)
        
        # 图名取自输出文件名
        out_path = ensure_output_path(
            args.output, out_dir,
            f"cropped_{bbox[0]}_{bbox[1]}_{bbox[2]}_{bbox[3]}.png")
        cropped.save(out_path)
        print(f"  裁剪来源: PNG图片（原始尺寸 {img.size[0]}x{img.size[1]}）")
        print(f"[OK] 已输出: {out_path}")
        return
    else:
        print("[ERROR] PNG 模式需要 --bbox 参数指定裁剪坐标", file=sys.stderr)
        sys.exit(2)


def main() -> None:
    parser = argparse.ArgumentParser(description="从 PDF/PNG 页面中裁剪/提取图片")
    parser.add_argument("input_path", help="输入 PDF 或 PNG 文件路径")
    parser.add_argument("--page", type=int, default=0, help="目标页码（1-based，PDF坐标裁剪/嵌入提取用）")
    parser.add_argument("--bbox", default="", help='百分比坐标 "left,top,right,bottom"（0~1）')
    parser.add_argument("--output", default="", help="输出图片路径（坐标裁剪模式）")
    parser.add_argument("--extract-embedded", action="store_true", help="提取该页所有嵌入图片对象（仅PDF）")
    parser.add_argument("--pages", default="", help='页码串 "1,3,5-7"（PDF整页渲染模式）')
    parser.add_argument("--render-all", action="store_true", help="整页渲染模式（仅PDF，每页输出一整张图）")
    parser.add_argument("--dpi", type=int, default=300, help="渲染分辨率 DPI（默认 300，仅PDF）")
    parser.add_argument("--output-dir", default="", help="输出目录（嵌入提取/整页渲染模式）")
    parser.add_argument("--prefix", default="page", help="整页渲染输出文件名前缀（默认 page，仅PDF）")
    args = parser.parse_args()

    input_path = os.path.abspath(args.input_path)
    if not os.path.isfile(input_path):
        print(f"[ERROR] 文件不存在: {input_path}", file=sys.stderr)
        sys.exit(2)

    # 根据文件后缀判断处理方式
    if input_path.lower().endswith(".pdf"):
        process_pdf_crop(args)
    elif input_path.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".webp")):
        process_png_crop(args)
    else:
        print(f"[ERROR] 不支持的文件格式（请提供 .pdf 或 .png/.jpg/.jpeg/.bmp/.tiff/.webp 文件）: {input_path}", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
