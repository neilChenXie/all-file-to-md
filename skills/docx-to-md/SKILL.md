---
name: docx-to-md
description: 当用户需要将doc或docx文件转成md文件时，请严格按照本技能步骤处理。本技能步骤经过复杂排版的word文档验证，可以最大幅度保证md文件与原word文件的一致性。本技能支持将用户指定的docx文件转换为markdown格式的文档，支持表格、公式、图片等内容的识别和转换。
---

用户指定的docx文件转化为md文件。

方案A / 方案C（HTML 链路）的整体链路：**docx → 导出 HTML → 解析为 Markdown → 校验修正 → 正式输出**。其中"docx → HTML"是决定保真度的关键一步，方案A / 方案C 的产物统一放在 `【输入文件路径】/tmp/【文件名】.html`，后续步骤完全一致：

| 方案 | 脚本 / 方式 | 原理 | 适用场景 | 保真度 |
|------|------|------|----------|--------|
| A（首选） | `docx_to_html.ps1` | 脚本内先试 WPS COM（`Kwps.Application` 等），不可用时自动回退 Microsoft Word COM（`Word.Application`），打开文档后另存为 HTML | Windows 且已安装 WPS 或 Microsoft Word | 高：合并单元格表格、公式、版式还原最好 |
| B（回退） | 提醒用户将 docx 转存为 PDF，再调用 `pdf-to-md` 技能 | 用户手动导出 PDF 后，由 `pdf-to-md` 将 PDF 逐页转 PNG，并用多模态模型识别为 md | 方案A 不可用或失败，且用户可将 docx 手动转存为 PDF（手动另存不受 COM 失败影响） | 高：基于 PDF 原版式逐页识别，图文、公式、表格版式还原好 |
| C（回退） | `docx_to_html_pandoc.py` | pandoc 转换（pypandoc-binary 自带 pandoc） | 用户无法将 docx 转存为 PDF（如未安装 WPS/Word，或不接受手动操作） | 中：表格结构保留，版式类信息有损，步骤5 需按专项核查清单修复 |

Mac 系统无 COM 自动化：优先采用方案B（由用户手动用 WPS / Word 将 docx 导出为 PDF，再交由 `pdf-to-md` 技能处理），或按方案A 的思路由用户手动用 WPS 将 docx 另存为 html 后放入 tmp 目录；以上均不可行时再用方案C（pandoc 跨平台）。

方案B（PDF 链路）的流程：**docx →（提醒用户手动转存）PDF → `pdf-to-md` 技能识别转换 → 正式输出**。该链路不执行本技能的步骤2~6，完整流程以 `pdf-to-md` 技能为准（详见步骤1 的方案B）。

## 重要提醒！

当用户需要将doc或docx文件转成md文件时，请严格按照本技能步骤处理。本技能步骤及scripts脚本，经过复杂排版的word文档验证，可以最大幅度保证md文件与原word文件的一致性。

## 文件路径

* 过程临时文件：
  - 【输入文件路径】/tmp/【文件名】.html
  - 【输入文件路径】/tmp/【文件名】.md
  - 【输入文件路径】/tmp/【文件名】.疑点清单.md（方案C 且 --checklist 时生成，步骤5 人工核查后可弃）
  - 【输入文件路径】/tmp/【文件名】.files/not-transfer-img.md
  - 【输入文件路径】/tmp/【文件名】.files/【图片名】.md
* 最终文件：
  - 【输入文件路径】/【文件名】.md
  - 【输入文件路径】/【文件名】.files/【图片名】.png

> 方案B（PDF 链路）的临时/最终文件沿用 `pdf-to-md` 技能的约定：临时文件为【输入文件路径】/tmp_【清洗后输入文件名】/ 下逐页导出的 PNG（及子agent 输出的【图片名】.md），最终产物同样是【输入文件路径】/【文件名】.md 及配套 `.files/` 图片目录。

## Workflow

> **链路说明**：方案A / 方案C 走 HTML 链路，依次执行步骤2~6；方案B 走 PDF 链路，由 `pdf-to-md` 技能完成全部转换，不执行步骤2~6。步骤1 决定选择哪种方案。

### 步骤1：选择转换方案并导出（HTML / PDF）

开始前先做一个判断：文档是否含图片——直接查 docx zip 包内是否存在 `word/media/` 条目（比转换后再 grep HTML 更可靠）。**无图片时，步骤3、步骤4 可整体跳过。**

然后按顺序选择方案：
1. Windows 且已安装 WPS 或 Microsoft Word → **方案A**（脚本内部先试 WPS，WPS 不可用时自动回退 Microsoft Word，无需手工选择）；
2. 方案A 不可用或失败 → **方案B**（提醒用户将 docx 转存为 PDF，再交由 `pdf-to-md` 技能处理）；
3. 用户无法转存 PDF → **方案C**（pandoc 导出）。

#### 方案A（首选）：WPS / Word COM 导出（Windows）

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

导出后的html文件位于 `【输入文件路径】/tmp/` 目录下（WPS 路径输出 charset=gb2312 的经典 HTML，Word 路径输出同样带 `【文件名】.files/` 图片目录的经典 HTML），图片文件会被导出到 `【输入文件路径】/tmp/【文件名】.files/` 目录下。

脚本内的执行顺序为 **WPS → Microsoft Word**：任一成功即返回，两个都失败才报错退出。

**故障排查（重要）：**

1. 若日志先出现「[WARN] WPS COM 连接失败」随后显示「[OK] Word conversion successful」→ 属正常回退，无需处理，继续后续步骤。
2. 若最终报「WPS 与 Microsoft Word 均无法完成导出」，再看脚本给出的诊断：
   - **提示 WPS / Word 程序打开着目标文件** → 这是最常见的失败原因：**WPS / Word 打开着目标 docx 文件时，COM 组件无法激活（`New-Object` 会静默失败）或无法打开该文件**。此时必须**提示用户关闭打开了目标文件的 WPS / Word 窗口（保险起见可关闭所有窗口），然后重试**。不要替用户强杀进程，避免丢失未保存的文档。
   - 用户关闭后重试仍失败，或提示未检测到 WPS / Word 进程 → 大概率未安装 WPS 与 Microsoft Word，或 COM 注册异常。提示用户安装其一后重试；若无法安装，转方案B（提醒用户用其他工具将 docx 转存为 PDF）或方案C（pandoc）。
3. 注意：部分执行环境（如安全策略）会拦截内联的 `New-Object -ComObject` 调用，此时应通过运行本脚本文件的方式触发 COM，而不是在命令行里内联实例化。

#### 方案B（回退）：转存 PDF + `pdf-to-md` 技能识别

方案A 不可用或失败时（未安装 WPS / Word、COM 被安全策略拦截、WPS/Word 打开着目标文件且关闭重试后仍失败等）采用本方案：**提醒用户手动将 docx 转存为 PDF，再调用 `pdf-to-md` 技能完成识别转换**。用户手动另存不依赖 COM 自动化，且最终 md 基于 PDF 原版式逐页识别，保真度高。

**操作步骤：**

1. **提醒用户转存 PDF**：请用户用 WPS / Word（或其他可用的 docx 转换工具）打开目标 docx 文件，选择"另存为 / 导出为 PDF"，保存到【输入文件路径】目录下，并**保持文件名与 docx 一致**（如 `document.docx` → `document.pdf`）。参考话术：
   > 本机无法自动完成 docx 转换，请手动协助一步：用 WPS / Word 打开该文件，将其"另存为 / 导出为 PDF"（注意保留原版式），文件名保持不变、保存到 `<【输入文件路径】>`，完成后告诉我。该方式转换的保真度更高。
2. **等待并核对**：用户确认后，检查 PDF 文件存在、可正常打开、内容完整（若 PDF 保存到了其他目录，则以该 PDF 所在目录作为后续的【输入文件路径】）。
3. **调用 `pdf-to-md` 技能**：以该 PDF 为输入，按 `pdf-to-md` 技能的完整流程执行（PDF 逐页导出 PNG → 分批多模态识别 → 合并 → 校验修正 → 按需提取插图），最终生成与 docx 同名的 md 文档及配套图片目录。
   - 两个技能使用相同的【输入文件路径】约定，PDF 与 docx 同目录即可无缝衔接；
   - `pdf-to-md` 只清理它自己的临时目录【输入文件路径】/tmp_【清洗后输入文件名】/（步骤4.0），**不会**动本技能方案A 遗留的【输入文件路径】/tmp/，如需清理请手动删除该目录；
   - 需当前模型具备图片识别能力（`pdf-to-md` 会自行校验）。
4. **本方案不执行本技能的步骤2~6**，最终产物以 `pdf-to-md` 的输出为准。
5. 若用户无法转存 PDF，转方案C。

#### 方案C（回退）：pandoc 导出（跨平台）

调用 `./scripts/docx_to_html_pandoc.py`，通过 pandoc 完成转换（经 pypandoc-binary 提供可执行文件，无需单独安装 pandoc）。

**依赖准备（一次性）**：在当前 Python 虚拟环境中安装 pypandoc-binary：

```bash
pip install pypandoc-binary -i https://pypi.tuna.tsinghua.edu.cn/simple
```

**脚本参数说明：**

| 参数 | 必填 | 说明 |
|------|------|------|
| `input.docx` | 是 | 输入的 docx 文件路径 |
| `output.html` | 否 | 输出的 html 文件路径，默认为输入文件同目录下的同名 html 文件 |

**使用示例：**

```bash
python ./scripts/docx_to_html_pandoc.py "【输入文件路径】/document.docx" "【输入文件路径】/tmp/document.html"
```

导出为 UTF-8 标准 HTML5，`html_to_markdown.py` 可直接处理。图片提取到 `【输入文件路径】/tmp/【文件名】.files/` 下（相对路径引用；pandoc 原生的 `media/` 子目录已由脚本自动平铺，与方案A 路径约定完全一致）。

### 步骤2：将html文件转换为md文件

> **依赖准备（重要）**：`html_to_markdown.py` 依赖 `beautifulsoup4` 与 `lxml`（`--docx` 编号回填同样需要 lxml）。若运行报 `ModuleNotFoundError: No module named 'bs4'`，需先在隔离的 Python 环境中安装：
> - 隔离环境（推荐）：`python -m venv <env>` 后，Windows 用 `<env>\Scripts\pip.exe install beautifulsoup4 lxml`（注意 Windows venv 的 pip 在 `Scripts\` 而非 `bin\`）；再用 `<env>\Scripts\python.exe` 运行脚本。
> - 之后所有 `python ./scripts/...py` 调用均改用该 venv 的 python 路径。

通过调用 `./scripts/html_to_markdown.py` python脚本，将导出的html文件转换为md文件。脚本自动探测编码，同时兼容方案A 的 WPS/Word HTML（GB2312/GBK）与方案C 的 pandoc HTML（UTF-8 HTML5）。

**脚本参数说明：**

| 参数 | 必填 | 说明 |
|------|------|------|
| `html_file` | 是 | 输入的 html 文件路径 |
| `output_md_file` | 否 | 输出的 md 文件路径，默认为输入文件同目录下的同名 md 文件 |
| `--docx <docx路径>` | 否 | 方案C 编号回填：解析源 docx 的 numbering.xml/document.xml 还原每段真实编号，按 `<li>` ↔ 段落文本对齐回填（修复 pandoc 丢失的多级编号祖先路径）。未命中的 `<li>` 回退 `<ol start>` 还原并打印警告。方案A 路径不传此参数，行为不变 |
| `--checklist` | 否 | 生成英文粘连/断行疑点清单 `.疑点清单.md`（仅供步骤5 人工核查定位，不自动改写） |

**使用示例：**

```powershell
# 方案A（WPS/Word HTML）：只指定输入文件，输出到同目录
python .\scripts\html_to_markdown.py "【输入文件路径】\tmp\document.html"

# 方案A 指定输出路径（推荐输出到 【输入文件路径】/tmp/ 目录）
python .\scripts\html_to_markdown.py "【输入文件路径】\tmp\document.html" "【输入文件路径】\tmp\document.md"

# 方案C（pandoc HTML）：带源 docx 做编号回填 + 疑点清单
python ./scripts/html_to_markdown.py "【输入文件路径】/tmp/document.html" "【输入文件路径】/tmp/document.md" --docx "【输入文件路径】/document.docx" --checklist
```

转换后的md文件位于 `【输入文件路径】/tmp/` 目录下。**方案C 建议始终带 `--docx`（源 docx 即输入文件）与 `--checklist`**；编号回填全部命中时打印 `[OK] 编号回填全部命中`，出现 `[WARN] 编号 <li> 未命中` 时需按警告清单对照原文核对相应列表。

### 步骤3：图片分类与内容识别

对 `【输入文件路径】/tmp/【文件名】.files/` 下的所有图片，分两步处理：先用 OCR 快速分类，再分类型转换。

> **路径差异**：方案A 与方案C 的图片现均直接位于 `【文件名】.files/` 下（方案C 的 `media/` 子目录已由 `docx_to_html_pandoc.py` 自动平铺），取图流程一致。

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

将 步骤3 识别生成的 【图片名】.md 文件 和 not-transfer-img.md 的内容整合进 步骤2 获取的网页内容md文件中，整合时注意保持原有的内容顺序。
- 用识别生成的md文件替换原有md文件中对应图片的引用。
- 对于 not-transfer-img.md 中的内容，保持原有md文件中对应图片的引用。

### 步骤5：验证输出结果

按以下步骤review输出的md文档，并修正发现的问题：

1. 【输入文件路径】/tmp/【文件名】.files/ 目录下的子agent输出的 【图片名】.md 和 not-transfer-img.md 文件的条目总和是否与图片文件数量一致。
2. 输出的md文件，表格是否有没有对齐的问题，是否有表格内容被拆分成多个表格的情况。
3. 输出的md文件，标题的编号是否出现错误，标题的层级是否正确。
4. 基于输出的md文件的结构，优化文章的标题等级和文章的格式，使其符合标准的markdown文档格式，比如标题等级、标题编号、段与段之间加空格等。
5. **pandoc 路径（方案C）专项核查**——方案A（WPS / Word）路径一般无此类问题。步骤2 已解决其中的系统性项，剩余项按工具辅助 + 人工判断处理：
   1) **自动编号丢失真实序号** → **已由步骤2 的 `--docx` 编号回填解决**（引擎语义与回归校验见 `test-case/verify_numbering_backfill.py` 与 `test-case/README.md`）。仅需抽查：确认转换日志为 `[OK] 编号回填全部命中`；若有 `[WARN]` 未命中项，按警告清单对照原文核对相应列表。
   2) **英文术语粘连或软换行断行** → 用 `--checklist` 生成的 `.疑点清单.md` 定位候选（长 ASCII 词内的 小写→大写 / 大写串→小写 / 字母↔数字 边界，及行尾↔行首英文词对），**人工确认后**按正确写法修复：拼接时 ASCII↔ASCII 边界补空格，CJK↔CJK 边界直接连接；注意排除 `iPhone`、`IoT` 等正常专有名词。
   3) **加粗丢失** → 列表项内加粗已由 `_render_inline` 保留；整行加粗的伪标题不会变成 `#` 标题，需按文档结构后处理（如 `**一、xxx**` → `## 一、xxx`），作为一般格式核查项。
   4) **段落合并/拆分要区分对待**：`<li>` 内多个 `<p>` 会被压成相邻行（段间无空行），而正文 `<p>` 之间有空行——"真软换行"（同一段被拆成多行）要合并，"相邻独立段落"要补空行，不能一刀切，保留人工判断。

### 步骤6：正式输出md文件及图片文件

将 【输入文件路径】/tmp/ 目录下的【md文档名】.md 文件移动到【输入文件路径】目录下，并将 【输入文件路径】/tmp/【文件名】.files/ 目录下的图片文件移动到【输入文件路径】/【文件名】.files/ 目录中。
