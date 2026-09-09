---
name: md-to-docx
description: 将 Markdown(.md/.markdown) 文件转换为排版规范的 Word(.docx) 文档。当用户提供 md 文件需要转成 Word/docx、或要求把 Markdown 内容按 Word 排版输出时使用。支持标题、表格、代码块、有序/无序列表、引用、分隔线、图片、加粗/斜体/删除线/行内代码/超链接/换行等常见语法；通过 ``--style`` 指定独立风格文件（assets/styles/*.md 中的 ```yaml``` 规格），按场景选用"默认/政府公文/..."等排版，也可为新场景新增风格文件。
agent_created: true
---

# md-to-docx 技能说明

将用户指定的 Markdown 文件转换为排版规范的 Word 文档（.docx）。脚本与调用步骤经多种 Markdown 结构（标题、表格、代码块、列表、引用、图片等）验证，转换结果可直接在 Word/WPS 中打开使用。

**风格（字体/字号/行距/页边距/页码等）通过独立风格文件驱动**——
`assets/styles/<场景>.md` 里写 ```yaml``` 规格，脚本按需加载；后续遇到新场景只需新增风格 md 即可。

## 何时使用

- 用户给出 `.md` 文件，要求"转成 Word / docx / 文档"
- 需要把 Markdown 内容变成可直接编辑排版的 Word 文档
- 需按特定版式输出（如政府公文 GB/T 9704、正式报告等）

## 脚本位置与依赖

- 脚本：`./scripts/md_to_docx.py`
- 依赖（按需）：
  - `python-docx`（必需）：`<venv>\Scripts\pip.exe install python-docx`
  - `PyYAML`（仅当使用 `--style` 时必需）：`pip install pyyaml`
- 运行必须使用 venv python（按用户约定的虚拟环境查找规则），禁止直接调用系统级 python。
- Windows 运行前建议设置 `PYTHONUTF8=1` 避免 GBK 编码问题。

## 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `input` | 是 | 输入的 Markdown 文件路径（utf-8 或 gb18030 均可自动识别） |
| `-o, --output` | 否 | 输出 docx 路径；默认与输入同名同目录；目录不存在会自动创建 |
| `--style` | 否 | 风格 md 路径（含 ```yaml``` 规格），缺省走内置默认风格 |
| `--title` | 否 | 文档大标题，覆盖 frontmatter `title:` 与文件名 |

## 风格选择（先选风格，再转换）

| 场景 | 风格文件 | 说明 |
| --- | --- | --- |
| 默认 / 通用 | `assets/styles/default.md` | A4 宋体小四 / 黑体标题 / 1.5 倍行距 |
| 党政机关公文（GB/T 9704） | `assets/styles/gov-doc.md` | 方正小标宋标题 / 仿宋正文 / 28pt 固定行距 / "-N-" 页码 |
| 其他新场景 | 见 `assets/styles/README.md` | 复制现有风格 md 新建，只改 yaml 字段 |

选择规则：按用户描述的场景/调性挑最贴近的；无法判断时询问用户或默认。
若所有风格都不贴合，**新增一个风格文件**（引导用户给出字号/字体/调性即可，不要硬套）。

## 字体预检（依赖特殊字体的风格必须执行）

公文等风格依赖系统已安装的特定中文字体（如 **仿宋_GB2312、方正小标宋_GBK、楷体_GB2312**）。
字体缺失时 Word/WPS 会**静默回退**，表现为"字体没生效"，且不报任何错误。

转换前先执行以下两步：

1. **放字体**：把需要的 ttf/otf 放到本 skill 的 `assets/fonts/` 目录；字体文件的内部族名
   需与风格文件中的字体名一致（如公文体 → `仿宋_GB2312.ttf`，可先复制 assets 下已有的同名文件）。
2. **预检安装**（自动补装缺失字体到用户字体库，无需管理员权限）：

   ```bash
   python scripts/ensure_fonts.py               # 检查缺失并自动安装
   python scripts/ensure_fonts.py --check-only  # 只检查不安装
   ```

   脚本会扫描 `assets/fonts/` 下所有字体文件，逐一核对系统是否已注册同族名字体，
   缺失则复制到 `%LOCALAPPDATA%\Microsoft\Windows\Fonts` 并写入 HKCU 注册表
   （Windows）；macOS/Linux 分别装到 `~/Library/Fonts` 与 `~/.fonts`。

安装完成后，**新启动**的 Word/WPS/浏览器进程即可识别该字体；已打开的进程需重启。

## 使用示例

默认风格（与历史默认行为一致）：

```bash
python scripts/md_to_docx.py "C:\path\to\文档.md"
```

指定输出 + 显式风格：

```bash
python scripts/md_to_docx.py "C:\path\to\文档.md" \
    -o "C:\path\to\输出.docx" \
    --style "C:\Users\admin\.workbuddy\skills\md-to-docx\assets/styles/default.md"
```

政府公文（GB/T 9704）：

```bash
python scripts/md_to_docx.py "C:\path\to\公文.md" \
    -o "C:\path\to\公文.docx" \
    --style "C:\Users\admin\.workbuddy\skills\md-to-docx\assets/styles/gov-doc.md" \
    --title "关于XX事项的通知"
```

> 公文 md 末尾的日期 / 附件用 HTML 块包裹以触发对应样式：
> ```html
> <div class="date">2026 年 9 月 9 日</div>
> <div class="attachment">附件：1. 任务分工表</div>
> ```

PowerShell 示例（先激活 venv 并设置 UTF-8）：

```powershell
$env:PYTHONUTF8 = '1'
<venv>\Scripts\python.exe .\scripts\md_to_docx.py "C:\path\to\文档.md" `
    -o "C:\path\to\输出.docx" `
    --style "C:\Users\admin\.workbuddy\skills\md-to-docx\assets/styles/gov-doc.md"
```

## 支持的语法（与历史默认行为保持一致）

| Markdown 语法 | 转换行为 |
|------|------|
| `#` ~ `######` 标题（1-6 级） | 风格文件 `headings.h1`~`h6` 决定字体/字号/粗细 |
| 段落 | 风格文件 `body.*` 决定字体/字号/行距/首行缩进 |
| `\| 表头 \| 表头 \|` 表格 | `table.*` 决定单元格字号与表头底纹 |
| ` ``` ` / `~~~` 围栏代码块 | `code_block.*` 决定字号与底纹 |
| `- / * / +` 无序列表 | Word 项目符号列表 |
| `1.` 有序列表 | 手工连续编号（多段列表均各自从 1 开始） |
| `> ` 引用 | `quote.*` 决定字体/字号/灰色/左侧竖线 |
| `---` / `***` / `___` 分隔线 | 段落底部横线 |
| `![alt](本地路径)` 图片 | 居中插入；宽度自动限制 ≤6 英寸 |
| `**加粗**` | 加粗 run |
| `*斜体*` / `~~删除线~~` | 斜体 / 删除线 run |
| `` `行内代码` `` | `fonts.code` 等宽字体 |
| `[文字](https://…)` | Word 可点击超链接（蓝色下划线） |
| `<br>` / `<br/>` | 段落内换行 |
| YAML frontmatter（文档开头 `---` 包裹） | 自动忽略；`title:` 自动作为文档大标题 |
| `<div class="date">…</div>` / `<div class="attachment">…</div>` | 触发公文 `date` / `attachment` 风格 |

## 降级与已知限制

- 嵌套列表层级缩进不保留（子项按同级项输出）；4 空格缩进代码块不支持（请用围栏代码块）
- 表格单元格内的图片、多级嵌套格式不支持
- 脚注、任务列表(`- [ ]`)、HTML 标签（除 date/attachment 块）、行内数学公式不解析
- 转义符号：单元格内 `\|` 转义竖线可识别；普通文本中的 `\*` 等转义符不额外处理
- 公文页码（`-N-`）由 Word 域 `PAGE` 实现，打开时按 F9 可强制刷新
- 文档中部单独成行的 `---` 视为分隔线，不会误判为表格

## 转换后检查（可选）

确认以下事项，异常时检查源 md 并重转：

1. 表格行/列是否完整，单元格内容是否错位（常见原因：源 md 表格行内有多余 `|`）
2. 标题层级是否符合原文档结构（源文件标题是否规范从 1 级递增）
3. 图片是否全部插入（若有红色"图片未能插入"文字，说明路径或格式不受支持）
4. 公文页脚页码是否为"-1-"居中格式；如显示为"1"占位，按 `Ctrl+A` 全选后 `F9` 刷新域