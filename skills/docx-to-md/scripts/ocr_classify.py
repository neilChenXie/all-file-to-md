# -*- coding: utf-8 -*-
"""
图片 OCR 分类工具 —— 自动识别图片类型：表格(table)、文字内容(text)、含文字的图像(image)。

用法：
  python ocr_classify.py <图片路径或目录> [--engine rapid|easy] [--json] [--text-out <dir>]

输出：
  --json       输出 JSON 格式的分类结果
  --text-out   将文字类图片的 OCR 文本写入指定目录（用于后续处理）

分类规则：
  - table:  多列多行对齐结构，列数≥2且行数≥2
  - text:   文字量充足（>50字符），非表格，含列表、段落等
  - image:  文字稀疏（≤50字符），或大量零散短文本（地图标注、图纸标签等）
"""
import os, sys, json, math

# ---------- 解析参数 ----------
TARGET = None
OUTPUT_JSON = False
TEXT_OUT_DIR = None
ENGINE = "rapid"

args = sys.argv[1:]
i = 0
positional = []
while i < len(args):
    a = args[i]
    if a == "--json":
        OUTPUT_JSON = True
    elif a == "--text-out":
        i += 1
        TEXT_OUT_DIR = args[i]
    elif a in ("--engine", "-e"):
        i += 1
        ENGINE = args[i]
    elif not a.startswith("--"):
        positional.append(a)
    i += 1

if positional:
    TARGET = positional[0]
else:
    print("用法: python ocr_classify.py <图片路径或目录> [--json] [--text-out <dir>]", file=sys.stderr)
    sys.exit(1)


def detect(path):
    """返回 [(x0, y0, x1, y1, text), ...] 按 y 排序"""
    boxes = []
    if ENGINE == "rapid":
        from rapidocr_onnxruntime import RapidOCR
        engine = RapidOCR()
        result, _ = engine(path)
        if result:
            for box, text, _conf in result:
                xs = [p[0] for p in box]
                ys = [p[1] for p in box]
                boxes.append((min(xs), min(ys), max(xs), max(ys), text))
    else:
        import easyocr
        reader = easyocr.Reader(["ch_sim", "en"], gpu=False, verbose=False)
        result = reader.readtext(path, detail=1)
        for box, text, _conf in result:
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            boxes.append((min(xs), min(ys), max(xs), max(ys), text))
    boxes.sort(key=lambda b: (b[1], b[0]))
    return boxes


def cluster_rows(boxes, tol=0.5):
    """按 y 中心聚类成行"""
    rows = []
    for b in boxes:
        cy = (b[1] + b[3]) / 2.0
        h = b[3] - b[1]
        placed = False
        for row in rows:
            if not row:
                continue
            row_cy = sum((bb[1] + bb[3]) / 2.0 for bb in row) / len(row)
            if abs(cy - row_cy) <= max(h, 6) * tol * 2 + 4:
                row.append(b)
                placed = True
                break
        if not placed:
            rows.append([b])
    rows.sort(key=lambda r: min(bb[1] for bb in r))
    for row in rows:
        row.sort(key=lambda bb: bb[0])
    return rows


def cluster_cols(rows, gap_tol=8):
    """x 投影聚类找列边界"""
    xs0 = []
    for row in rows:
        for b in row:
            xs0.append((b[0], b[2]))
    xs0.sort()
    clusters = []
    for left, right in xs0:
        if not clusters:
            clusters.append([left, right])
            continue
        if left - clusters[-1][1] <= gap_tol:
            clusters[-1][1] = max(clusters[-1][1], right)
        else:
            clusters.append([left, right])
    return clusters


def is_table_structure(rows, col_clusters):
    """判断是否具备表格结构：列≥2 且行≥2 且多列行占比≥60%"""
    if len(col_clusters) < 2 or len(rows) < 2:
        return False
    multi = sum(1 for row in rows if len(row) >= 2)
    return multi / len(rows) >= 0.6


def classify_image(path):
    """
    对单张图片分类。返回 dict:
      {
        "file": 文件名,
        "category": "table" | "text" | "image",
        "boxes": OCR 检测框总数,
        "char_count": 总字符数,
        "rows": 行数,
        "cols": 列数（仅 table 类别有效）,
        "avg_chars_per_box": 每框平均字符数,
      }
    """
    boxes = detect(path)
    if not boxes:
        return {"file": os.path.basename(path), "category": "image",
                "boxes": 0, "char_count": 0, "rows": 0, "cols": 0,
                "avg_chars_per_box": 0, "reason": "no_text_detected"}

    rows = cluster_rows(boxes)
    col_clusters = cluster_cols(rows)

    char_count = sum(len(b[4]) for b in boxes)
    avg_chars = char_count / len(boxes) if boxes else 0
    is_table = is_table_structure(rows, col_clusters)
    n_rows = len(rows)
    n_cols = len(col_clusters)

    # ---- 分类逻辑 ----
    def make_result(cat, reason):
        return {"file": os.path.basename(path), "category": cat,
                "boxes": len(boxes), "char_count": char_count,
                "rows": n_rows, "cols": n_cols,
                "avg_chars_per_box": round(avg_chars, 1),
                "reason": reason}

    # 1. 表格 —— 条件：多列多行对齐，且排除"地图标签"误判
    #    地图特征：大量短标签（avg_chars ≤ 5, boxes ≥ 30）→ 实为地图
    if is_table:
        if avg_chars <= 5 and len(boxes) >= 30:
            return make_result("image", "map_labels_mimic_table_rows")
        return make_result("table", "table_structure_detected")

    # 2. 文字（段落/列表等）
    #    条件：总字符 ≥ 30，且不是大量零散短标签（avg_chars ≥ 3）
    if char_count >= 30 and avg_chars >= 3:
        return make_result("text", "sufficient_text_content")

    # 2.5 单文本块（如标题横幅）—— 框数少但单框文字密集 → 判为文字
    if avg_chars >= 15 and char_count >= 15:
        return make_result("text", "single_dense_text_block")

    # 3. 图像（地图/图纸/照片）—— 字符稀疏或全是短标签
    return make_result("image", "sparse_or_label_text")


def format_text_output(boxes):
    """将 OCR 检测框按行输出纯文本（用于文字类图片的预提取）"""
    rows = cluster_rows(boxes)
    lines = []
    for row in rows:
        line_text = " ".join(b[4] for b in row)
        lines.append(line_text)
    return "\n".join(lines)


# ---------- 主逻辑 ----------
def main():
    if os.path.isdir(TARGET):
        image_files = sorted(
            f for f in os.listdir(TARGET)
            if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".webp"))
        )
        base_dir = TARGET
    else:
        image_files = [os.path.basename(TARGET)]
        base_dir = os.path.dirname(TARGET) or "."

    results = []
    for fname in image_files:
        path = os.path.join(base_dir, fname)
        r = classify_image(path)
        results.append(r)

        if TEXT_OUT_DIR and r["category"] in ("table", "text"):
            os.makedirs(TEXT_OUT_DIR, exist_ok=True)
            boxes = detect(path)
            text = format_text_output(boxes)
            name = os.path.splitext(fname)[0] + ".txt"
            with open(os.path.join(TEXT_OUT_DIR, name), "w", encoding="utf-8") as f:
                f.write(text)

    if OUTPUT_JSON:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        # 人类友好输出
        for r in results:
            cat_label = {"table": "[表格]", "text": "[文字]", "image": "[图像]"}.get(r["category"], "???")
            print(f"{cat_label} {r['file']}  "
                  f"chars={r['char_count']} boxes={r['boxes']} "
                  f"rows={r['rows']} cols={r['cols']}  "
                  f"({r['reason']})")

    # 统计摘要
    cats = {}
    for r in results:
        cats[r["category"]] = cats.get(r["category"], 0) + 1
    summary = " | ".join(f"{k}: {v}" for k, v in sorted(cats.items()))
    print(f"\n分类统计 → {summary}")


if __name__ == "__main__":
    main()
