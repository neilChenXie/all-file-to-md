---
name: docx-to-md
description: 当用户需要将doc或docx文件转成md文件时，请严格按照本技能步骤处理。本技能步骤经过复杂排版的word文档验证，可以最大幅度保证md文件与原word文件的一致性。本技能支持将用户指定的docx文件转换为markdown格式的文档，支持表格、公式、图片等内容的识别和转换。
---

用户指定的docx文件转化为md文件。
- 在Windows系统，通过调用wps脚本导出html文件后，再利用python脚本将html文件转换为markdown格式的文档，支持表格、公式、图片等内容的识别和转换。
- 在mac系统，推荐用户手动通过wps导出html文件后，再利用python脚本将html文件转换为markdown格式的文档，支持表格、公式、图片等内容的识别和转换。

## 重要提醒！

当用户需要将doc或docx文件转成md文件时，请严格按照本技能步骤处理。本技能步骤及scripts脚本，经过复杂排版的word文档验证，可以最大幅度保证md文件与原word文件的一致性。

## 文件路径

* 过程临时文件：
  - 【输入文件路径】/tmp/【文件名】.html
  - 【输入文件路径】/tmp/【文件名】.md
  - 【输入文件路径】/tmp/【文件名】.files/not-transfer-img.md
  - 【输入文件路径】/tmp/【文件名】.files/【图片名】.md
* 最终文件：
  - 【输入文件路径】/【文件名】.md
  - 【输入文件路径】/【文件名】.files/【图片名】.png
  
## Workflow

### 步骤1：将docx文件导出为html文件

#### Windows系统

调用 `./scripts/docx_to_html.ps1` 脚本，将用户指定的docx文件导出为html文件。

**脚本参数说明：**

| 参数 | 必填 | 说明 |
|------|------|------|
| `-InputFile` | 是 | 输入的 docx 文件路径 |
| `-OutputFile` | 否 | 输出的 html 文件路径，默认为输入文件同目录下的同名 html 文件 |

**使用示例：**

```powershell
# 基本用法：指定输入文件，输出到默认位置
.\scripts\docx_to_html.ps1 -InputFile "C:\Users\xxx\document.docx"

# 指定输出路径（推荐输出到 【输入文件路径】/tmp/ 目录）
.\scripts\docx_to_html.ps1 -InputFile "C:\Users\xxx\document.docx" -OutputFile "【输入文件路径】\tmp\document.html"
```

导出后的html文件位于 `【输入文件路径】/tmp/` 目录下，图片文件会被导出到 `【输入文件路径】/tmp/【文件名】.files/` 目录下。

#### Mac/Linux系统

用户需要手动通过wps将docx文件导出为html文件，导出后的html文件位于 `【输入文件路径】/tmp/` 目录下。（帮助用户创建 `【输入文件路径】/tmp/` 目录，并提示用户将导出的html文件放入该目录下。）


### 步骤2：将html文件转换为md文件

> **依赖准备（重要）**：`html_to_markdown.py` 依赖 `beautifulsoup4` 与 `lxml`。若运行报 `ModuleNotFoundError: No module named 'bs4'`，需先在隔离的 Python 环境中安装：
> - 隔离环境（推荐）：`python -m venv <env>` 后，Windows 用 `<env>\Scripts\pip.exe install beautifulsoup4 lxml`（注意 Windows venv 的 pip 在 `Scripts\` 而非 `bin\`）；再用 `<env>\Scripts\python.exe` 运行脚本。
> - 之后所有 `python ./scripts/...py` 调用均改用该 venv 的 python 路径。

通过调用 `./scripts/html_to_markdown.py` python脚本，将导出的html文件转换为md文件。

**脚本参数说明：**

| 参数 | 必填 | 说明 |
|------|------|------|
| `html_file` | 是 | 输入的 html 文件路径 |
| `output_md_file` | 否 | 输出的 md 文件路径，默认为输入文件同目录下的同名 md 文件 |

**使用示例：**

```powershell
# 基本用法：只指定输入文件，输出到同目录
python .\scripts\html_to_markdown.py "【输入文件路径】\tmp\document.html"

# 指定输出路径（推荐输出到 【输入文件路径】/tmp/ 目录）
python .\scripts\html_to_markdown.py "【输入文件路径】\tmp\document.html" "【输入文件路径】\tmp\document.md"
```

转换后的md文件位于 `【输入文件路径】/tmp/` 目录下。

### 步骤3：图片分类与内容识别

对 `【输入文件路径】/tmp/【文件名】.files/` 下的所有图片，分两步处理：先用 OCR 快速分类，再分类型转换。

#### 3.1 快速分类（ocr_classify.py）

运行 OCR 分类脚本，自动将图片归为三类：

```bash
python ./scripts/ocr_classify.py "【输入文件路径】/tmp/【文件名】.files/" [--json]
```

**分类输出示例：**

```
[图像] image1.png  chars=692 boxes=168 rows=30 cols=2  (map_labels_mimic_table_rows)
[文字] image2.png  chars=476 boxes=35  rows=27 cols=1  (sufficient_text_content)
[表格] image3.png  chars=245 boxes=18  rows=6  cols=4  (table_structure_detected)
[文字] image4.png  chars=26  boxes=1   rows=1  cols=1  (single_dense_text_block)

分类统计 → image: 1 | table: 1 | text: 2
```

**三类含义：**
| 类别 | 标签 | 含义 | 后续处理 |
|------|------|------|----------|
| `[表格]` | table | 多列多行对齐的表格截图 | → 3.3 多模态模型转为 Markdown 表格 |
| `[文字]` | text | 段落、列表、标题等文字截图 | → 3.3 多模态模型转为正文 |
| `[图像]` | image | 地图、图纸、照片等非文字主导图 | → 3.2 存入 not-transfer-img.md |

**分类准确性：** 脚本基于文字密度和坐标结构自动判定，大宗情况下准确。如有疑虑，可将分类结果展示给用户确认后再继续。

#### 3.2 非文字图片处理（image 类）

对于分类为 `[图像]` 的图片，总结图片内容，按以下格式追加到 `【输入文件路径】/tmp/【文件名】.files/not-transfer-img.md`：

```md
![图片内容简短摘要](./【文件名】.files/图片名.png)
```

#### 3.3 文字类图片处理（table / text 类）

对于分类为 `[表格]` 或 `[文字]` 的图片，使用具有图像识别能力的大模型读取图片内容，转为 md 文件。

**处理方式：** 按每批不超过 10 张图片启动子 Agent，用 Read 工具读取图片文件（模型自动识别图片中的文字、表格结构），输出为格式化 md。

- 表格图片 → 输出 Markdown 表格格式（`| col1 | col2 |`），保留行列结构
- 列表/段落图片 → 输出正文段落或列表，保留层级缩进

子 Agent 输出的 md 内容存入 `【输入文件路径】/tmp/【文件名】.files/【图片名】.md`（每张图片对应一个 md 文件）。

**子 Agent prompt 模板：**
> 读取以下图片文件，将其中的表格或文字内容转换为 Markdown 格式：
> - 如果图片包含表格，输出标准 Markdown 表格（竖线分隔，表头加粗）
> - 如果图片包含列表/段落，保留原层级结构和缩进
> - 只输出内容，不要添加解释性文字
> - 输出存入 `【图片名】.md`

### 步骤4：整合图片识别内容

将 步骤3 识别生成的 `【图片名】.md` 文件 和 `not-transfer-img.md` 的内容整合进 步骤2 获取的网页内容md文件中，整合时注意保持原有的内容顺序。
- 用识别生成的md文件替换原有md文件中对应图片的引用。
- 对于 `not-transfer-img.md` 中的内容，保持原有md文件中对应图片的引用。

### 步骤5：验证输出结果

按以下步骤review输出的md文档，并修正发现的问题：
1. `【输入文件路径】/tmp/【文件名】.files/` 目录下的子agent输出的 `【图片名】.md` 和 `not-transfer-img.md` 文件的条目总和是否与图片文件数量一致。
2. 输出的md文件，表格是否有没有对齐的问题，是否有表格内容被拆分成多个表格的情况。
3. 输出的md文件，标题的编号是否出现错误，标题的层级是否正确。
4. 基于输出的md文件的结构，优化文章的标题等级和文章的格式，使其符合标准的markdown文档格式，比如标题等级、标题编号、段与段之间加空格等。

### 步骤6：正式输出md文件及图片文件

将 `【输入文件路径】/tmp/【md文档名】.md` 文件移动到`【输入文件路径】/【md文档名】.md`，并将 `【输入文件路径】/tmp/【文件名】.files/` 目录下的图片文件移动到 `【输入文件路径】/【文件名】.files/` 目录中。

