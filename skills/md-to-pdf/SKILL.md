---
name: md-to-pdf
description: Convert Markdown files into nicely-styled PDF documents (Pandoc + headless Edge/Chrome). This skill should be used when the user asks to turn a .md file into PDF, e.g. "把 xx.md 转成 PDF"、"md 转 pdf"、"导出 PDF 版本"、"生成一份排好版的 PDF"。排版风格由 assets/styles/ 下独立风格 md 文件提供，可按场景选用或新增。
agent_created: true
---

# Md To Pdf

## Overview

将 Markdown(.md) 文件转换为**带排版样式**的 PDF：
pandoc 先把 md 转成 HTML 并注入所选风格的 CSS，再用无头 Edge/Chrome 打印为 PDF。
风格独立存放在 `assets/styles/*.md`，按文档场景挑选；后续可在该目录新增风格文件以覆盖更多场景。

典型触发语：`将 数字孪生-培训需求.md 转成 pdf` / `这份 md 帮我输出一份正式的 PDF 报告`。

## 字体预检（依赖特殊字体的风格必须执行）

公文等风格依赖系统已安装的特定中文字体（如 **仿宋_GB2312、方正小标宋_GBK、楷体_GB2312**）。
字体缺失时浏览器会**静默回退**到默认字体，表现为"字体没生效"。

转换前先执行以下两步：

1. **放字体**：把需要的 ttf/otf 放到本 skill 的 `assets/fonts/` 目录；字体文件的内部族名
   需与风格文件中的字体名一致（如公文体 → `仿宋_GB2312.ttf`）。
2. **预检安装**（自动补装缺失字体到用户字体库，无需管理员权限）：

   ```bash
   python scripts/ensure_fonts.py               # 检查缺失并自动安装
   python scripts/ensure_fonts.py --check-only  # 只检查不安装
   ```

   脚本会扫描 `assets/fonts/` 下所有字体文件，逐一核对系统是否已注册同族名字体，
   缺失则复制到 `%LOCALAPPDATA%\Microsoft\Windows\Fonts` 并写入 HKCU 注册表
   （Windows）；macOS/Linux 分别装到 `~/Library/Fonts` 与 `~/.fonts`。

安装完成后，**新启动**的 Word/WPS/浏览器进程即可识别该字体；已打开的进程需重启
（headless 打印每次都是新进程，装完即可直接用）。

## 风格选择（先选风格，再转换）

| 场景 | 风格文件 | 说明 |
| --- | --- | --- |
| 默认 / 会议纪要 / 培训需求 / 汇报提纲 | `assets/styles/default-blue.md` | 简洁商务蓝 |
| 技术方案 / 研究报告 / 申报材料 / 正式发文 | `assets/styles/formal-report.md` | 正式报告（深灰/学术） |
| 党政机关公文（GB/T 9704） | `assets/styles/gov-doc.md` | 含"-N-"页码模板与公文标准字号 |
| 其他新场景 | 见 `assets/styles/README.md` | 复制现有风格 md 新建，只改描述与 css 代码块 |

选择规则：按用户描述的场景/调性挑最贴近的；无法判断时询问用户或默认 blue。
若所有风格都不贴合，**新增一个风格文件**（引导用户给出主色/字体/调性即可，不要硬套），
新增后同步更新 `assets/styles/README.md` 的表格。

> 风格文件除 ```css 块外，可选追加 ```browser-extras 块（每行一条浏览器参数），
> 主要用于自定义页眉/页脚模板（如公文页码）。详见 `assets/styles/README.md`。

## 转换流程

1. **确认输入**：源 md 绝对路径、目标 pdf 输出路径、文档标题（默认取文件名；
   若 md 第一行已是 `# ` 一级标题则不额外注入标题）。
2. **运行脚本**（以 skill 根目录为基准，venv python 执行）：

   ```bash
   # 找到 skill 根目录（本文件上一级），用项目/用户 venv 的 python 运行
   python scripts/md2pdf.py "<输入.md>" "<输出.pdf>" \
       --style "assets/styles/default-blue.md" \
       [--title "文档标题"] [--browser "浏览器exe路径，可选"]
   ```

   示例（正式报告场景）：
   ```bash
   python "C:/Users/admin/.workbuddy/skills/md-to-pdf/scripts/md2pdf.py" \
       "需求.md" "需求-正式版.pdf" \
       --style "C:/Users/admin/.workbuddy/skills/md-to-pdf/assets/styles/formal-report.md" \
       --title "港口数字孪生平台培训需求"
   ```

3. **校验产物**：确认输出文件存在、非空且以 `%PDF-` 开头（脚本已自动校验并打印页数）；
   脚本退出码 0 表示成功。
4. **展示结果**：调用 present_files 呈现 PDF 文件；简要说明所用风格与输出路径。

## 常见问题与避坑（实测经验）

- **不要**尝试 pandoc 直接输出 PDF（本机无 LaTeX 引擎）；HTML 中转是无 LaTeX 环境的最稳路径。
- **中文/空格长路径**：无头浏览器的 `--print-to-pdf` 直接写含中文路径会报
  「系统找不到指定的路径」。脚本已内置：输出到 ASCII 临时文件后 `os.replace` 到最终路径，
  生成用 HTML 也放临时目录（文件名全 ASCII），从根源规避 file:// 编码问题。
- **浏览器被占用**：必须传独立 `--user-data-dir`，否则 headless 会复用已开实例而不打印。
- **页眉页脚**：新版用 `--no-pdf-header-footer`；旧版需 `--print-to-pdf-no-header`
  （脚本会依次尝试多种参数组合）。
- **headless 日志噪音**：Chrome/Edge 输出 QQBrowser importer 等无害 ERROR，属正常现象，不用理会。
- **中文字体**：样式依赖系统字体（雅黑/宋体），目标机器无该字体会回退，无需额外处理。
- **校验技巧**：`re.findall(rb"/Type\s*/Page[^s]", data)` 可粗略统计页数。

## 新增/维护风格

新增风格与文件约定详见 `assets/styles/README.md`。要点：
风格文件为普通 md，头部写「适用场景/调性/主色」，样式放唯一的 ` ```css ``` ` 代码块；
建议补齐 h1~h4、表格、代码块、引用样式，避免元素裸样式。

## 资源

- `scripts/md2pdf.py` — 核心转换脚本（pandoc 转 HTML → 注入 CSS → 无头浏览器打印 PDF）。
- `assets/styles/` — 独立风格 md 文件目录（场景模板，可扩展）；`README.md` 为目录说明与新增指南。
