# 风格：政府公文（GB/T 9704）

> **适用场景**：党政机关、企事业单位的正式公文（通知、通报、报告、请示、函等）。
> **标准依据**：GB/T 9704—2012《党政机关公文格式》。
> **排版要点**：方正小标宋_GBK 标题、仿宋_GB2312 正文、黑体一级标题、楷体二级标题；
> A4 页边距 上3.7cm / 下3.5cm / 左2.8cm / 右2.6cm；行距固定 28pt（标题 32pt）；
> 页码 宋体 四号 "-N-" 格式居中。


> **字体依赖**：正文 **仿宋_GB2312**（skill 自带 assets/fonts/ 中已内置该字体，
> 用 ensure_fonts.py 自动安装到系统后即生效）；标题依赖 **方正小标宋_GBK**。

## 元素规范

| 角色 | 字体 | 字号 | 行距 | 其他 |
| --- | --- | --- | --- | --- |
| 文档标题 | 方正小标宋_GBK | 2号(22pt) | 固定32pt | 居中、不加粗 |
| 正文 / 列表 / 表格 | 仿宋_GB2312 | 3号(16pt) | 固定28pt | 首行缩进 2 字符 |
| 一级标题 | 黑体 | 3号(16pt) | 固定28pt | — |
| 二级标题 | 楷体_GB2312 | 3号(16pt) | 固定28pt | — |
| 三级及以下标题 | 仿宋_GB2312 | 3号(16pt) | 固定28pt | — |
| 数字/拉丁字母 | Times New Roman | 3号(16pt) | 固定28pt | 全局适用 |
| 页码 | 宋体 | 四号(14pt) | — | "-N-" 格式、居中 |
| 日期 | 仿宋_GB2312 | 3号(16pt) | 固定28pt | 右对齐（空 4 字），与正文空 3 行 |
| 附件 | 仿宋_GB2312 | 3号(16pt) | 固定28pt | 左空 2 字，后跟全角冒号"：" |

## CSS（转换脚本提取本代码块）

```css
@page { size: A4; margin: 3.7cm 2.6cm 3.5cm 2.8cm; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body {
  font-family: "Times New Roman", "仿宋_GB2312", "FangSong_GB2312", "FangSong", "STFangsong", serif;
  font-size: 16pt;
  color: #000000;
  line-height: 28pt;
  margin: 0;
  text-align: justify;
  word-break: break-all;
}
#title-block-header { margin: 0 0 0.3em 0; padding: 0; border: none; }
h1.title {
  font-family: "方正小标宋_GBK", "FZXiaoBiaoSong-B05S", "FZXBSJW--GB1-0", "STZhongsong", "SimSun", serif;
  font-size: 22pt;
  font-weight: normal;
  line-height: 32pt;
  text-align: center;
  margin: 0 0 0.6em 0;
  letter-spacing: 0.05em;
}
/* 一级标题：黑体 */
h1 {
  font-family: "黑体", "SimHei", "STHeiti", sans-serif;
  font-size: 16pt; font-weight: normal; line-height: 28pt;
  margin: 0.6em 0 0.2em 0; text-indent: 0;
}
/* 二级标题：楷体 */
h2 {
  font-family: "楷体_GB2312", "KaiTi_GB2312", "KaiTi", "STKaiti", serif;
  font-size: 16pt; font-weight: normal; line-height: 28pt;
  margin: 0.5em 0 0.2em 0; text-indent: 0;
}
/* 三级标题：方正仿宋 */
h3, h4, h5, h6 {
  font-family: "仿宋_GB2312", "FangSong_GB2312", "FangSong", "STFangsong", serif;
  font-size: 16pt; font-weight: normal; line-height: 28pt;
  margin: 0.5em 0 0.2em 0; text-indent: 0;
}
p { margin: 0; line-height: 28pt; text-indent: 2em; }
p:empty { display: none; }
ul, ol { margin: 0; padding-left: 2em; }
li { line-height: 28pt; margin: 0; }
/* 表格、引用、代码：方正仿宋同行 */
table { border-collapse: collapse; width: 100%; margin: 0.4em 0; font-size: 16pt; }
th, td { border: 1px solid #000000; padding: 4px 8px; line-height: 28pt; vertical-align: middle; }
th { font-weight: normal; text-align: center; }
blockquote {
  margin: 0; padding: 0 0 0 2em;
  font-family: "楷体_GB2312", "KaiTi_GB2312", "KaiTi", "STKaiti", serif;
  border: none;
}
code { font-family: "Times New Roman", "Consolas", "Courier New", monospace; font-size: 14pt; }
pre {
  font-family: "Times New Roman", "Consolas", "Courier New", monospace;
  font-size: 14pt; line-height: 24pt; padding: 4px 8px;
  background: #f4f4f4; border: 1px solid #cccccc;
}
hr { border: none; border-top: 1px solid #000; margin: 0.6em 0; }
strong { font-weight: normal; }
/* 公文专属：日期行（在文档末尾、与正文空三行、右对齐） */
.date {
  text-align: right;
  text-indent: 0;
  line-height: 28pt;
  margin-top: 84pt;       /* 3 行 × 28pt */
  padding-right: 4em;     /* 右空 4 字 */
}
/* 公文专属：附件 */
.attachment {
  text-indent: 2em;       /* 左空 2 字 */
  line-height: 28pt;
  margin: 0.4em 0;
}
```

## 浏览器附加参数（自定义页码 "-N-"）

```browser-extras
# 空页眉（避免浏览器自动添加 URL/标题/日期）
--print-to-pdf-header-template=<div></div>
# 自定义页码：宋体 四号 居中 "-N-" 格式
--print-to-pdf-footer-template=<div style="font-size:14pt; font-family:'SimSun','宋体',serif; width:100%; text-align:center; -webkit-print-color-adjust:exact;">-<span class="pageNumber"></span>-</div>
```

## 使用提示

- **文档标题**通过 `--title "..."` 或 md 内 YAML `title: ...` 提供；脚本会按 h1.title 渲染为方正小标宋。
- **日期行**：在 md 末尾用 HTML 包裹，例如：
  ```html
  <div class="date">2026 年 9 月 9 日</div>
  ```
- **附件**：用 `.attachment` 类（首行缩进 2 字），后接全角冒号"："：
  ```html
  <div class="attachment">附件：1. 任务分工表</div>
  ```
