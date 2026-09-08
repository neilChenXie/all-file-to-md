---
name: md-to-docx
description: 将 Markdown(.md/.markdown) 文件转换为排版规范的 Word(.docx) 文档。当用户提供 md 文件需要转成 Word/docx、或要求把 Markdown 内容按 Word 排版输出时使用。支持标题、表格、代码块、有序/无序列表、引用、分隔线、图片、加粗/斜体/删除线/行内代码/超链接/换行等常见语法；自动设置 A4 页面与中文字体（正文宋体、标题黑体），并忽略 YAML frontmatter。
agent_created: true
---

# md-to-docx 技能说明

将用户指定的 Markdown 文件转换为排版规范的 Word 文档（.docx）。脚本与调用步骤经多种 Markdown 结构（标题、表格、代码块、列表、引用、图片等）验证，转换结果可直接在 Word/WPS 中打开使用。

## 何时使用

- 用户给出 `.md` 文件，要求"转成 Word / docx / 文档"
- 需要把 Markdown 内容变成可直接编辑排版的 Word 文档

## 脚本位置与依赖

- 脚本：`./scripts/md_to_docx.py`
- 依赖：`python-docx`。若运行报 `ModuleNotFoundError: No module named 'docx'`，需在 Python 虚拟环境中安装：
  - Windows：`<venv>\Scripts\pip.exe install python-docx`
  - macOS/Linux：`<venv>/bin/pip install python-docx`
- 之后所有调用均使用该 venv 的 python：Windows 为 `<venv>\Scripts\python.exe`，macOS/Linux 为 `<venv>/bin/python`。禁止直接调用系统级 python。
- 运行 python 前设置 UTF-8 环境变量避免 Windows 控制台编码问题（可选）：PowerShell 执行 `$env:PYTHONUTF8 = '1'`。

## 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `input` | 是 | 输入的 Markdown 文件路径（utf-8 或 gb18030 均可自动识别） |
| `-o, --output` | 否 | 输出的 DOCX 路径；默认与输入文件同名同目录；输出目录不存在会自动创建 |

## 使用示例

基本用法（输出到输入文件同目录的 .docx）：

```bash
python ./scripts/md_to_docx.py "C:\path\to\文档.md"
```

指定输出路径：

```bash
python ./scripts/md_to_docx.py "C:\path\to\文档.md" -o "C:\path\to\输出.docx"
```

PowerShell 示例（先激活虚拟环境并设置 UTF-8）：

```powershell
$env:PYTHONUTF8 = '1'
<venv>\Scripts\python.exe .\scripts\md_to_docx.py "C:\path\to\文档.md" -o "C:\path\to\输出.docx"
```

转换完成后：将生成的 .docx 路径告知用户，并用文件呈现工具打开供预览/下载。

## 支持的语法

| Markdown 语法 | 转换行为 |
|------|------|
| `#` ~ `######` 标题（1-6 级） | Word 内置标题样式；1-3 级黑体（16/14/12pt），4-6 级宋体加粗，黑色 |
| 段落 | 正文宋体小四(12pt)、1.5 倍行距 |
| `\| 表头 \| 表头 \|` 表格 | Word 表格：表头加粗居中+浅灰底纹，支持单元格内 `<br>` 换行 |
| ` ``` ` / `~~~` 围栏代码块 | 等宽字体(Consolas)+浅灰底纹单段，保留空行与缩进，可带语言标注 |
| `- / * / +` 无序列表 | Word 项目符号列表 |
| `1.` 有序列表 | 手工连续编号（多段列表均各自从 1 开始，不会连号错乱） |
| `> ` 引用 | 楷体灰色、左侧竖线，连续引用行合并为一段 |
| `---` / `***` / `___` 分隔线 | 段落底部横线 |
| `![alt](本地路径)` 图片 | 插入图片并居中；相对路径以 md 文件所在目录解析；宽度自动限制 ≤6 英寸（高 ≤8.5 英寸），小图不放大；文件不存在/网络图片则输出红色提示文字 |
| `**加粗**` | 加粗 run |
| `*斜体*` / `~~删除线~~` | 斜体 / 删除线 run |
| `` `行内代码` `` | Consolas 等宽字体 |
| `[文字](https://…)` | Word 可点击超链接（蓝色下划线） |
| `<br>` / `<br/>` | 段落内换行 |
| YAML frontmatter（文档开头 `---` 包裹） | 自动忽略 |

## 降级与已知限制

以下情况会按纯文本处理或存在局限，如遇大段此类内容可提前告知用户：

- 嵌套列表层级缩进不保留（子项按同级项输出）；4 空格缩进代码块不支持（请用围栏代码块）
- 表格单元格内的图片、多级嵌套格式不支持
- 脚注、任务列表(`- [ ]`)、HTML 标签、行内数学公式不解析
- 转义符号：单元格内 `\|` 转义竖线可识别；普通文本中的 `\*` 等转义符不额外处理
- 页面固定 A4、正文宋体小四、行距 1.5；如需页边距/字号差异需在生成后于 Word/WPS 中调整
- 文档中部单独成行的 `---` 视为分隔线，不会误判为表格

## 转换后检查（可选）

确认以下事项，异常时检查源 md 并重转：

1. 表格行/列是否完整，单元格内容是否错位（常见原因：源 md 表格行内有多余 `|`）
2. 标题层级是否符合原文档结构（源文件标题是否规范从 1 级递增）
3. 图片是否全部插入（若有红色"图片未能插入"文字，说明路径或格式不受支持）
