# 风格文件目录说明（styles/）

本目录存放「MD → PDF」的**独立风格文件**。每个文件是一个标准 Markdown，供人阅读/维护，
其中唯一的 ` ```css ... ``` ` 代码块会被 `scripts/md2pdf.py` 提取为排版样式。

## 现有风格

| 文件名 | 风格名 | 适用场景 |
| --- | --- | --- |
| `default-blue.md` | 简洁商务蓝（默认） | 会议纪要、培训需求、汇报提纲、日常文档 |
| `formal-report.md` | 正式报告（深灰/学术） | 技术方案、研究报告、申报材料、正式发文 |
| `gov-doc.md` | 政府公文（GB/T 9704） | 党政机关公文、通知/通报/报告等正式发文 |

## 风格文件结构

每个风格文件是一个标准 md，至少包含一个 ```css 围栏代码块（必需），
可选 ```browser-extras 围栏块，用于追加浏览器参数。

### 1) 必需的 ```css 块

CSS 注入到 pandoc 生成的 HTML 的 `<style>` 中。`md2pdf.py` 会用 Chromium/Edge 打印，
因此需要写浏览器能识别的标准 CSS（不要用 LaTeX/Word 专有属性）。

### 2) 可选的 ```browser-extras 块

为风格追加额外的浏览器命令行参数，每行一条，**主要用于自定义页眉/页脚模板**
（例如政府公文要求的"-N-"页码）。脚本会原样追加到浏览器命令中。

支持的关键参数：
- `--print-to-pdf-header-template="<div>...</div>"` 自定义页眉
- `--print-to-pdf-footer-template="<div>...</div>"` 自定义页脚
- 占位符（仅在这两个模板中可用）：`pageNumber` / `totalPages` / `title` / `url` / `date`

⚠️ **一旦风格中包含上述任一参数，脚本不会再追加 `--no-pdf-header-footer`**，
由风格文件全权控制页眉页脚行为。

示例（政府公文的"-N-"页码）：

````
```browser-extras
--print-to-pdf-header-template=
--print-to-pdf-footer-template=<div style="font-size:14pt; font-family:'SimSun','宋体',serif; width:100%; text-align:center; -webkit-print-color-adjust:exact;">-<span class="pageNumber"></span>-</div>
```
````

## 新增风格（针对新场景）

遇到新场景时，复制任一现有文件并修改即可：

1. **复制**：`cp formal-report.md meeting-green.md`（命名建议 `<场景>-<调性>.md`，如 `contract-dark.md`）。
2. **改头部说明**：第 1 行标题、适用场景、调性、主色等描述字段。
3. **改 CSS 代码块**：只保留 `@page`、`body`、标题、列表、表格、代码等与排版相关规则；
   可调整的关键变量：
   - 页边距：`@page { margin: ... }`
   - 字体族：正文/标题字体（中文字体需系统已安装，否则回退）
   - 主色：h1/h2 标题颜色、表格表头底色
   - 行距：固定值用绝对单位 `line-height: 28pt`
4. **必要时追加 browser-extras**：需要页码、页眉页脚等浏览器层能力时添加。
5. **约定**：CSS 需放在 ```css 围栏内；建议同时定义 h1~h4、表格、代码块样式，
   避免文档出现元素"裸样式"。

## 选择规则（供调用方判断）

- 默认无明确场景 → `default-blue.md`
- 正式/严肃/对外 → `formal-report.md`
- 党政机关公文 → `gov-doc.md`
- 其他 → 从描述词（场景/调性/主色/字体）中挑最贴近的；都不贴合时按上文新增一个风格文件，而非硬套。
