---
name: xlsx-to-md
description: 将 xlsx 表格文件转存为结构化 Markdown。本技能用 openpyxl 直读单元格，把合并标题行转成 # 标题、分类小标题行转成 ##、明细行按表头生成 Markdown 表格，并正确处理数字格式与 HTML 特殊字符转义。当用户需要把 xlsx/xls（尤其是工程量清单、设备清单、报价表、数据表）转成 md 时使用。
---

# xlsx-to-md

将 xlsx 转存为保真 Markdown。

## 前置
- openpyxl（managed Python 已自带 3.1.5；若缺失则在 venv 里 `pip install openpyxl`）。
- 用 managed Python 运行：
  `C:\Users\fhdi2\.workbuddy\binaries\python\versions\3.13.12\python.exe`

## 用法
脚本位于本技能 `scripts/xlsx_to_md.py`，接受命令行参数（通用、不写死路径）：

```bash
# 单个文件（md 输出到 xlsx 同目录，同名 .md）
python "C:/Users/fhdi2/.workbuddy/skills/xlsx-to-md/scripts/xlsx_to_md.py" "路径/文件.xlsx"

# 多个文件
python ".../xlsx_to_md.py" "a.xlsx" "b.xlsx"

# 指定输出目录
python ".../xlsx_to_md.py" "a.xlsx" --out "D:/输出目录"
```

## 识别规则（关键）
BOQ/清单 xlsx 典型结构：
- **第 1 行**：合并标题（如 A1:F1），可能含 `\n` 说明 → 拆成 `# 标题` + 说明段落。
- **第 2 行**：表头（序号 / 建设项目 / 规格描述 / 单位 / 数量 / 依据条款…）。
- **分类小标题行**：A 列有内容、B 列（及之后）为空（如"一、中央SCADA控制中心"）→ 输出 `## 小标题`。
- **明细行**：A、B 均有内容 → 作为表格数据行。

> 合并单元格（如 A1:F1）用 `ws.cell` 读取时只在左上角有值，无需特殊处理。

## 转换要点
1. 每个分类小标题下独立生成一张 Markdown 表格（列名取第 2 行表头，在每个表里重复）。Markdown 不支持表内标题，按分类拆表是最保真可读的方式。
2. 单元格文本清洗：`str(v).replace("\n"," ").strip()`。
3. **数字格式保真**：float 且 number_format 含 `.00` → `f"{v:.2f}"`（保留 1.00）；整数值 float → `int(v)`；否则 `str(v)`。
4. **HTML 特殊字符转义**（重要）：单元格文本里的 `<` `>` 在 Markdown 中会被当 HTML 标签吞掉。源文件常见 `<RED>待定数量标注</RED>` 这类模板红字，必须转义为 `&lt;RED&gt;`。同时转义 `|` → `\|`。
5. 明细行跳过 A、B 均为空的行。

## 校验（每次转换后建议做）
- 用 openpyxl 重数：section 数（A 有值且 B 空）= md 中 `^## ` 数量；item 数（A、B 均有值）= md 中以 `| ` 开头且非表头/分隔符的数据行数。两者必须相等，确认无数据丢失。
- 抽查 数量 列是否保留 "1.00" 等两位小数。

## 注意事项
- **源文件笔误**：转换应当如实保留源内容。若发现标头/分类序号等明显笔误，先转存，再提示用户确认是否修正，不要擅自改动源数据含义。
