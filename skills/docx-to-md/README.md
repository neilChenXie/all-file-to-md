# docx-to-md

> 复杂排版 Word 文档（doc / docx）→ 高保真 Markdown 的技能（Agent Skill）。

本 README 是面向人的技能总览与能力说明；**完整操作流程（步骤1~6、参数表、故障排查、验收核查）以 [SKILL.md](./SKILL.md) 为准**，测试细节与实测数据见 [test-case/README.md](./test-case/README.md)。

## 技能简介

- **输入**：Word 文档（`.docx`；`.doc` 建议先另存为 `.docx`）
- **输出**：与输入同目录下的【文件名】`.md` + 配套【文件名】`.files/` 图片目录（过程临时文件位于【输入文件路径】`/tmp/`，约定详见 SKILL.md）
- **覆盖内容**：多级自动编号标题、复杂表格（合并单元格 / 不规则行列）、目录（TOC）与 Word 域、图片（含图内文字与表格的识别转换）、公式等
- **定位**：步骤与脚本经三类极端形态的复杂文档实测与回归验证（见「能力特点」），最大程度保证 md 与原 Word 文件的一致性

## 三种转换方案

docx → md 的核心是「docx → HTML / PDF」这一步，方案选择决定保真度上限：

| 方案 | 适用场景 | 原理（一句话） | 保真度 |
|------|---------|---------------|--------|
| **A（首选）** | Windows 且已安装 WPS 或 Microsoft Word | `docx_to_html.ps1` 经 WPS COM 导出 HTML（WPS 不可用时自动回退 Microsoft Word COM） | 高：合并单元格表格、公式、版式还原最好 |
| **B（回退）** | 方案A 不可用或失败，且用户可将 docx 手动转存 PDF | 用户手动另存 PDF → 交由 `pdf-to-md` 技能（逐页 PNG + 多模态识别） | 高：基于 PDF 原版式逐页识别 |
| **C（回退）** | 无法转存 PDF（如未安装 WPS/Word、Mac 无 COM 自动化） | `docx_to_html_pandoc.py` 经 pandoc 导出 HTML（pypandoc-binary 自带 pandoc） | 中：表格结构保留、版式类信息有损，由编号引擎与疑点清单补偿 |

**选择顺序**：

1. Windows 且已安装 WPS / Microsoft Word → **方案A**（脚本内部先试 WPS，失败自动回退 Word，无需手工选择）；
2. 方案A 不可用或失败 → **方案B**（提醒用户手动转存 PDF，不受 COM 失败影响）；
3. 用户无法转存 PDF → **方案C**（pandoc 跨平台）。

> 方案A / 方案C 的导出产物（HTML）走同一套解析与校验流程（步骤2~6，共用 `html_to_markdown.py`）；方案B 不执行步骤2~6，全流程以 `pdf-to-md` 技能为准。macOS 无 COM 自动化：优先方案B，或按方案A 的思路由用户手动另存 HTML 后走后续步骤；均不可行再用方案C。

## 能力特点

以下能力均基于 `test-case/` 中三个真实文档实测（数据出处见各节标注；完整口径与回归方法见 [test-case/README.md](./test-case/README.md)）：

| 用例 | 输入特点 | 主要考验 |
|------|---------|---------|
| test-case-1 | 约 28MB「投标施工方案」；524 个多级自动编号「伪标题」（编号不在正文文本中）；110 张内嵌图片；49 个表格 | 超大文档、编号还原、图片与表格 |
| test-case-2 | WPS 导出的 GB2312 HTML（无 charset 声明）；18 个不规则 API 函数说明表（328 处 colspan + 15 处 rowspan）；0 列表 0 图片 | 编码探测、复杂表格 |
| test-case-3 | 约 64MB「初步设计文件」；Word 域目录 169 条；多级编号标题且 docx 源头编号失步；142 处图片引用 | 域 / 目录、编号反哺、图片 alt |

### ① 超大文档端到端转换

- **能力**：28~64MB 输入、上万行产物的全流程（导出 → 解析 → 图片识别 → 校验 → 输出）不崩溃、不截断，正文与参考答案逐字比对。
- **实测**：test-case-1 参考答案约 9,200 行；test-case-3 参考答案 18,780 行。

（出处：test-case/README.md test-case-1 / test-case-3）

### ② 标题与多级自动编号还原（本技能处理得最深的一类问题）

- **标题层级**：真实 Heading 样式正确映射为 `#` ~ `######`（test-case-2 实测 2 个 h1 + 18 个 h2；test-case-3 多层标题全量转换）。
- **多级自动编号**：编号不在正文文本中、由 Word 列表引擎动态生成，两条链路各自还原：
  - **方案C**：`--docx` 编号引擎（`docx_numbering.py`）直接解析 docx 的 `numbering.xml` + `document.xml`，按 numId 隔离、按文档顺序统一计数后回填（修复 pandoc 丢弃编号祖先路径的问题）；
  - **方案A**：正文标题编号缓存失真 / 失步时，用目录（TOC）域缓存中的正确编号反哺正文标题（锚点对齐，文本比对兜底）。
- **实测**：
  - test-case-1：524 个多级编号标题与参考答案逐字一致，编号回填命中 771/771 段，编号引擎逐段核对 527/527 段吻合；列表项内加粗 536 处全部保留；
  - test-case-3：失步文档反哺 166/166 命中（`第2章` / `2.1` / `1.4.1` 等正确），方案C 与方案A 在该问题上完全拉平。
- **引擎保真要点**：编号计数器按 numId 隔离；中文数字（`一、二` / `壹、贰`）、圆圈数字（`①`）、字母（超 26 进位）等 numFmt 全量渲染；`lvlRestart`、`isLgl`、`lvlOverride` 等 OOXML 规则按规范处理；删除修订（`w:del`）与 `AlternateContent` 回退内容不参与计数。
- **一句话结论**：编号正确性取决于「编号由谁计算」——健康文档（test-case-1）靠编号引擎全量还原；失步文档（test-case-3）的正确值只存在于目录缓存，必须走反哺路线（引擎按 OOXML 规范计算的是失步值）。

（出处：test-case/README.md「文档自动编号」/ test-case-3）

### ③ 复杂表格还原

- **能力**：合并单元格（colspan / rowspan）网格展开、空单元格补位、同一表内各行列数不一致时不丢内容不错位；单元格内容按 run 直连，不引入假空格（如 `Char*` 被拆成多个文本 run 的写法）。
- **实测**：
  - test-case-2（核心）：18 个 API 函数说明表、328 处 colspan + 15 处 rowspan，190 表格行与参考答案一致；
  - test-case-1：49 个表格、847 表格行不丢失、不错位；
  - test-case-3：方案A 路径 3,136 表格行。

> 方案C（pandoc）对合并单元格的展开方式与 WPS 不同（test-case-3 实测 2,890 vs 3,136 行），属引擎层限制，见「已知边界」。

（出处：test-case/README.md test-case-2 / test-case-1 / test-case-3）

### ④ 目录结构化与域代码清理

- **能力**：Word 域（TOC / PAGEREF）与条件注释整节点清理，域代码不泄漏为正文；目录区从纯文本污染转换为可跳转的 md 列表（层级缩进 + 锚点链接 + 页码后缀）。
- **实测**：
  - test-case-3：域代码泄漏 169 处 → 0、`mso-` 残留 1,014 处 → 0；目录 169 条（MsoToc1/2/3 = 10/51/108）结构化，目录区约 1,200 行压缩为约 170 行；
  - test-case-2：MsoToc 目录 20 条结构化；
  - 方案C 兼容：pandoc 导出的连续 `<p><a href="#标题锚">编号 标题 页码</a>` 目录段同样识别（≥3 条判定）。

（出处：test-case/README.md test-case-3「发现的问题与普适化优化点」）

### ⑤ 图片提取、语义化与内容识别

- **能力**：
  - **提取与引用**：图片统一提取到【文件名】`.files/`，相对路径引用；方案C 的 `media/` 子目录自动平铺，与方案A 目录约定完全一致；
  - **alt 语义化**：图片行相邻图注（`图N` / `表N`）自动回填为 alt；
  - **内容识别**：OCR 快速分类为「表格 / 文字 / 图像」三类（判定依据：多列多行对齐结构 → 表格；文字量充足 → 文字；文字稀疏或大量零散短文本 → 图像），表格与文字类交多模态模型转 md 正文/表格，非文字图保留引用并汇总到 `not-transfer-img.md`。
- **实测**：
  - test-case-1：110 处图片引用全部提取一致（109 个唯一媒体文件，image82.png 被引用 2 次）；
  - test-case-3：142 处图片引用；87/142 处 alt 由图注回填。

（流程细节见 SKILL.md 步骤3；出处：test-case/README.md）

### ⑥ 编码与文本细节兼容

- **能力**：
  - **编码自动探测**：同一脚本兼容方案A 的 GB2312 / GBK 经典 HTML（可无 charset 声明）与方案C 的 UTF-8 HTML5；
  - **空格语义**：`&nbsp;` 语义空格保留、排版空白折叠；跨标签拆分的文本（WPS 常见）不粘连、不丢空格；
  - **疑点清单**：`--checklist` 输出英文粘连 / 软换行断行 / 域残留等候选（只列不改），供步骤5 人工定位。
- **实测**：
  - test-case-2：无 charset 声明的 GB2312 文档自动按 gb18030 正确解码；
  - test-case-3：图注 / 表题粘连 243 处 → 30 处（剩余为源 HTML 中空格已丢失、无法自动恢复，由清单兜底人工补）。

（出处：test-case/README.md test-case-2 / test-case-3）

## 已知边界与人工核查项

- **方案C 表格展开差异**：pandoc 对合并单元格的展开方式与 WPS 不同（test-case-3：2,890 vs 3,136 表格行），属 pandoc 引擎层限制，需按步骤5 人工核查。
- **失步文档勿用 `--docx` 覆盖反哺**：正文编号与目录编号矛盾的文档（如 test-case-3），正确值以**目录编号**为准；`--docx` 编号引擎按 OOXML 规范计算的是失步值。
- **剩余人工核查项**：
  - 图注 / 表题粘连 30 处（源 HTML 空格已丢失）——按疑点清单对照原文补空格；
  - 英文粘连 / 软换行断行候选——人工确认后修复（注意排除 `iPhone`、`IoT` 等正常专有名词）；
  - 编号回填出现 `[WARN] 编号 <li> 未命中` 时，按警告清单对照原文核对。
- **与参考答案的文本重合率**（验收基线，非逐字节相等）：test-case-1 ≥ 82.7%、test-case-3 方案A ≥ 96.8%、方案C ≥ 77.7%；差距主要来自 pandoc 引擎的表格展开与 run 空格风格（方案A 侧的少量差异为 WPS 重导出导致的页码缓存 / 图片文件名偏移，属预期）。

## 快速上手

**依赖（一次性）**：

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt    # Windows: .venv\Scripts\pip install -r requirements.txt
```

> `requirements.txt` 含方案C 的 `pypandoc-binary`（自带 pandoc 可执行文件，无需单独安装 pandoc）；方案B 改用 `pdf-to-md` 技能，不依赖本目录。

**方案A（Windows + WPS/Word，首选）**：

```powershell
.\scripts\docx_to_html.ps1 -InputFile "D:\docs\document.docx" -OutputFile "D:\docs\tmp\document.html"
python .\scripts\html_to_markdown.py "D:\docs\tmp\document.html" "D:\docs\tmp\document.md" --checklist
```

**方案C（跨平台回退，pandoc）**：

```bash
python ./scripts/docx_to_html_pandoc.py "docs/document.docx" "docs/tmp/document.html"
python ./scripts/html_to_markdown.py "docs/tmp/document.html" "docs/tmp/document.md" --docx "docs/document.docx" --checklist
```

**方案B（无法自动转换时）**：请用户用 WPS / Word 将 docx「另存为 / 导出为 PDF」，再交由 `pdf-to-md` 技能处理，完整流程以该技能为准。

> 以上仅为「docx → md 正文」的最小链路；图片分类识别、内容整合、校验修正、正式输出等步骤见 [SKILL.md](./SKILL.md) 步骤3~6。方案C 建议始终带 `--docx`（编号回填）与 `--checklist`（疑点清单）。

## 文档与脚本索引

| 文件 | 说明 |
|------|------|
| [SKILL.md](./SKILL.md) | 技能主文档（操作手册）：方案选择、步骤1~6、参数表、故障排查 |
| [test-case/README.md](./test-case/README.md) | 测试资产：三个用例的能力点、实测数据、回归命令与验收口径 |
| `scripts/docx_to_html.ps1` | 方案A：WPS / Microsoft Word COM 导出 HTML（WPS 优先，Word 自动回退） |
| `scripts/docx_to_html_pandoc.py` | 方案C：pandoc 导出 HTML（图片相对路径引用，`media/` 自动平铺） |
| `scripts/html_to_markdown.py` | HTML → md 主转换：编码探测、目录结构化、编号反哺、域清理、`--docx` 编号回填、`--checklist` 疑点清单 |
| `scripts/docx_numbering.py` | 编号引擎：解析 `numbering.xml` / `document.xml` 还原 Word 多级编号（供 `--docx` 回填） |
| `scripts/test_docx_numbering.py` | 编号引擎单元测试：内存构造最小 docx，19 用例覆盖 numFmt 渲染 / 级别语义 / 结构边界（`python skills/docx-to-md/scripts/test_docx_numbering.py`） |
| `scripts/verify_numbering.py` | 编号一致性比对：引擎还原编号 ↔ 参考答案 md 双指针对齐，可按 abstractNumId 过滤（如 test-case-1 传 27 → 527/527） |
| `scripts/ocr_classify.py` | 图片 OCR 分类：table / text / image 三类（步骤3.1） |
| [requirements.txt](./requirements.txt) | Python 依赖：beautifulsoup4、lxml、chardet、rapidocr-onnxruntime、Pillow、pypandoc-binary |

> 注：`test-case/` 为本地测试资产，未纳入 git 版本控制；本文引用的测试数据与回归口径均出自 `test-case/README.md`。
