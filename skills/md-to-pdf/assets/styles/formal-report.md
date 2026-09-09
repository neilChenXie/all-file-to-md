# 风格：正式报告（深灰 / 学术）

> **适用场景**：技术方案、研究报告、申报材料、结题文档、正式发文等严肃正式文档。
> **调性**：稳重、严谨、适合黑白打印。
> **主色**：深蓝灰 `#1f3864`；**正文**：宋体 12pt、行距 1.9；**标题**：黑体/雅黑加粗；**A4 页边距** 2.5cm/2.2cm。

## CSS（转换脚本提取本代码块）

```css
@page { size: A4; margin: 2.5cm 2.2cm; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body {
  font-family: "SimSun", "宋体", "Times New Roman", serif;
  font-size: 12pt; color: #1a1a1a; line-height: 1.9; margin: 0;
}
#title-block-header {
  border-bottom: 2.5px solid #1f3864; margin-bottom: 1.6em;
  padding: 0.2em 0 0.7em 0;
}
h1.title {
  font-size: 22pt; margin: 0 0 0.1em 0; color: #1f3864;
  font-family: "Microsoft YaHei", "微软雅黑", "SimHei", sans-serif;
  letter-spacing: 2px; font-weight: 700; text-align: center;
}
h1 {
  font-size: 17pt; color: #1f3864; margin: 1.6em 0 0.7em;
  font-family: "Microsoft YaHei", "微软雅黑", "SimHei", sans-serif;
  font-weight: 700; border-bottom: 1.5px solid #8fa8c8; padding-bottom: 4px;
}
h2 {
  font-size: 14.5pt; color: #1f3864; margin: 1.5em 0 0.5em 0;
  font-family: "Microsoft YaHei", "微软雅黑", "SimHei", sans-serif;
  font-weight: 700; padding-bottom: 3px; border-bottom: 1px solid #c3cede;
}
h3 {
  font-size: 13pt; color: #2c3e50; margin: 1.2em 0 0.4em 0;
  font-family: "Microsoft YaHei", "微软雅黑", "SimHei", sans-serif; font-weight: 700;
}
h4 { font-size: 12pt; color: #333333; margin: 1em 0 0.3em 0; font-weight: 700; }
p { margin: 0.5em 0; text-align: justify; }
ul, ol { margin: 0.2em 0 0.9em 0; padding-left: 2em; }
li { margin: 0.25em 0; }
li ul, li ol { margin: 0.15em 0 0.35em 0; }
blockquote {
  margin: 0.6em 0 1.2em 0; padding: 6px 16px; color: #444444;
  border-left: 3px solid #1f3864; background: #f4f6f9;
}
code {
  font-family: Consolas, "Courier New", monospace; font-size: 88%; color: #7a2c1e;
}
pre {
  background: #f6f6f4; border: 1px solid #d8d8d2; border-radius: 2px;
  padding: 10px 12px; overflow-x: auto; font-size: 10pt;
}
pre code { color: #2b2b2b; background: none; }
table { border-collapse: collapse; width: 100%; margin: 0.6em 0 1.2em 0; font-size: 11pt; }
th, td { border: 1px solid #999999; padding: 5px 10px; text-align: left; }
th { background: #eef1f6; color: #1f3864; font-weight: 700; }
hr { border: none; border-top: 1px solid #b0b0b0; margin: 1.4em 0; }
strong { font-weight: 700; }
```
