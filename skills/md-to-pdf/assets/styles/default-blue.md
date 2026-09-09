# 风格：简洁商务蓝（默认）

> **适用场景**：会议纪要、培训需求梳理、汇报提纲、日常备忘等通用文档。
> **调性**：清爽、现代、信息层级分明。
> **主色**：商务蓝 `#0b5cad`；**正文**：微软雅黑 11.5pt、行距 1.8；**A4 页边距** 2.3cm/2.1cm。

## CSS（转换脚本提取本代码块）

```css
@page { size: A4; margin: 2.3cm 2.1cm; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body {
  font-family: "Microsoft YaHei", "微软雅黑", "PingFang SC", "Segoe UI", sans-serif;
  font-size: 11.5pt; color: #222222; line-height: 1.8; margin: 0;
}
#title-block-header { border-bottom: 3px solid #0b5cad; margin-bottom: 1.4em; padding-bottom: 0.5em; }
h1.title { font-size: 20pt; margin: 0 0 0.1em 0; color: #0b3d6e; letter-spacing: 1px; font-weight: 700; }
h1 { font-size: 18pt; color: #0b3d6e; margin: 1.4em 0 0.6em; font-weight: 700; }
h2 {
  font-size: 14.5pt; color: #0b5cad; margin: 1.5em 0 0.6em 0;
  padding: 6px 12px; background: #eef4fb; border-left: 5px solid #0b5cad;
  border-radius: 2px; font-weight: 700;
}
h3 { font-size: 12.5pt; color: #0b3d6e; margin: 1.2em 0 0.4em 0; font-weight: 700; }
h4 { font-size: 12pt; color: #333333; margin: 1em 0 0.3em 0; font-weight: 700; }
p { margin: 0.4em 0; }
ul, ol { margin: 0.2em 0 0.8em 0; padding-left: 1.8em; }
li { margin: 0.22em 0; }
li ul, li ol { margin: 0.15em 0 0.3em 0; }
blockquote {
  margin: 0.5em 0 1em 0; padding: 4px 14px; color: #555555;
  border-left: 4px solid #b9d3ec; background: #f7fafd;
}
code { font-family: Consolas, "Courier New", monospace; font-size: 90%; color: #c0392b; }
pre {
  background: #f6f8fa; border: 1px solid #e1e4e8; border-radius: 4px;
  padding: 10px 12px; overflow-x: auto; font-size: 9.5pt;
}
pre code { color: #24292e; background: none; }
table { border-collapse: collapse; width: 100%; margin: 0.6em 0 1em 0; font-size: 10.5pt; }
th, td { border: 1px solid #ccd6e0; padding: 5px 10px; text-align: left; }
th { background: #eef4fb; color: #0b3d6e; font-weight: 700; }
tr:nth-child(even) td { background: #f8fafc; }
hr { border: none; border-top: 1px solid #d5dde5; margin: 1.2em 0; }
strong { color: #0b3d6e; }
```
