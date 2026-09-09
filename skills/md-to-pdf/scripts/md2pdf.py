#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2pdf.py —— Markdown 转排版 PDF（Pandoc + 无头 Edge/Chrome）

把本会话沉淀的转换逻辑固化为可复用脚本：
  1. 从【风格 md 文件】中提取 ```css 代码块作为排版样式
  2. pandoc (gfm -> html5, -s, -H 注入 <style>) 生成 HTML
  3. 无头 Edge/Chrome --print-to-pdf 打印为 PDF

内置踩坑处理：
  - Windows 中文/空格路径：--print-to-pdf 直接写中文长路径会报
    "系统找不到指定的路径"，故先输出到 ASCII 临时路径再 os.replace
  - 浏览器打印输出的 HTML 同样放 ASCII 临时目录，规避 file URL 编码问题
  - --headless=new 需独立 --user-data-dir，避免与已打开的浏览器冲突
  - 兼容新旧 headless 参数（--no-pdf-header-footer / --print-to-pdf-no-header）

用法：
  python md2pdf.py <input.md> <output.pdf> [--style <style.md>] [--title <标题>] [--browser <exe>]

示例：
  python md2pdf.py 需求.md 需求.pdf --style styles/default-blue.md
  python md2pdf.py 方案.md 方案.pdf --style styles/formal-report.md --title "技术方案"
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# 脚本所在目录的上一级 = skill 根目录
SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_STYLE = SKILL_DIR / "assets" / "styles" / "default-blue.md"

# 常见浏览器路径（按序探测）
BROWSER_CANDIDATES = [
    # Windows Edge
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    # Windows Chrome
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    # macOS
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    # Linux
    "chromium-browser",
    "chromium",
    "google-chrome",
    "microsoft-edge",
]


def log(msg: str) -> None:
    print(msg, flush=True)


def find_pandoc() -> str:
    p = shutil.which("pandoc")
    if not p:
        raise RuntimeError("未找到 pandoc，请先安装：https://pandoc.org/installing.html")
    return p


def find_browser(explicit: str | None = None) -> str:
    if explicit:
        if os.path.isfile(explicit):
            return explicit
        raise RuntimeError(f"指定的浏览器不存在: {explicit}")
    for cand in BROWSER_CANDIDATES:
        if os.path.sep in cand or cand.startswith("/"):
            if os.path.isfile(cand):
                return cand
        else:
            p = shutil.which(cand)
            if p:
                return p
    raise RuntimeError("未找到 Edge/Chrome，无法打印 PDF；可用 --browser 指定浏览器路径")


def extract_css_and_extras(style_md: Path) -> tuple[str, list[str]]:
    """从风格 md 中同时提取 ```css 与可选 ```browser-extras 块。

    - 必需 ```css：注入到 HTML 的 <style>
    - 可选 ```browser-extras：每行一个浏览器参数（如 --print-to-pdf-footer-template="..."），
      主要用于自定义页眉/页脚模板（例如公文"-N-"格式页码）。空行与 # 注释行忽略。
    """
    if not style_md.exists():
        raise RuntimeError(f"风格文件不存在: {style_md}")
    text = style_md.read_text(encoding="utf-8")
    css_match = re.search(r"```css\s*\n(.*?)```", text, re.S)
    if not css_match:
        raise RuntimeError(
            f"风格文件 {style_md} 中未找到 ```css ... ``` 代码块，"
            "请按 assets/styles/README.md 的约定编写风格文件"
        )
    css = css_match.group(1).strip()
    if not css:
        raise RuntimeError(f"风格文件 {style_md} 的 css 代码块为空")

    extras: list[str] = []
    extras_match = re.search(r"```browser-extras\s*\n(.*?)```", text, re.S)
    if extras_match:
        for raw in extras_match.group(1).splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            extras.append(line)
    return css, extras


def detect_auto_title(md_path: Path) -> str | None:
    """自动选择文档标题：优先 YAML frontmatter `title:`；其次首行 # 标题（命中则不注入以免重复）；最后用文件名。"""
    try:
        text = md_path.read_text(encoding="utf-8")
        lines = text.splitlines()
    except Exception:
        return Path(md_path).stem
    # 1) 解析 YAML frontmatter
    if lines and lines[0].strip() == "---":
        for line in lines[1:]:
            s = line.strip()
            if s == "---":
                break
            m = re.match(r"^title\s*:\s*(.+?)\s*$", s)
            if m:
                v = m.group(1).strip()
                if v.startswith(("'", '"')) and v.endswith(("'", '"')) and len(v) >= 2:
                    v = v[1:-1]
                return v or None
    # 2) 正文首行 # 标题：命中则不注入标题
    for line in lines:
        s = line.strip()
        if not s:
            continue
        if s.startswith("# "):
            return None
        break
    # 3) 退化到文件名
    return Path(md_path).stem


def run_pandoc(pandoc: str, md: Path, header_html: Path, html_out: Path, title: str | None) -> None:
    cmd = [
        pandoc, str(md),
        "-f", "gfm",
        "-t", "html5",
        "-s",
        "-H", str(header_html),
        "-o", str(html_out),
    ]
    if title:
        cmd += ["--metadata=title=" + title]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if proc.returncode != 0:
        raise RuntimeError("pandoc 转换失败:\n" + (proc.stderr or proc.stdout))


def print_to_pdf(browser: str, html_file: Path, pdf_out: Path, workdir: Path,
                 extras: list[str] | None = None) -> None:
    """无头浏览器打印为 PDF。extras 来自风格文件的 browser-extras 块，可自定义页眉/页脚模板等。

    - 若 extras 含 --print-to-pdf-header/footer-template，则不追加 --no-pdf-header-footer
      （以免覆盖自定义模板），且按用户传入顺序直接组装参数
    - 否则走默认的多组参数重试，目标是兼容新旧 headless（--no-pdf-header-footer / --print-to-pdf-no-header）
    """
    url = html_file.as_uri()
    profile = workdir / "profile"
    extras = extras or []
    has_custom_tpl = any("header-template" in e or "footer-template" in e for e in extras)

    def _run(flags: list[str]) -> tuple[Path | None, str]:
        target = workdir / "out.pdf"
        if target.exists():
            target.unlink()
        cmd = [browser, *flags,
               f"--user-data-dir={profile}",
               *extras,
               f"--print-to-pdf={target}",
               url]
        try:
            proc = subprocess.run(cmd, capture_output=True, timeout=180)
        except subprocess.TimeoutExpired as e:
            return None, f"浏览器打印超时: {e}"
        err_text = (proc.stderr or proc.stdout or b"").decode("utf-8", errors="replace")[-500:]
        if target.exists() and target.stat().st_size > 0:
            return target, ""
        return None, err_text

    # 自定义模板场景：直接跑一次（覆盖默认的 --no-pdf-header-footer 行为）
    if has_custom_tpl:
        base = ["--headless=new", "--disable-gpu"]
        target, err = _run(base)
        if target is None:
            raise RuntimeError("无头浏览器打印 PDF 失败: " + (err or "未知错误"))
        os.replace(target, pdf_out)
        return

    # 默认：依次尝试多组参数（无自定义模板）
    attempts = [
        ["--headless=new", "--disable-gpu", "--no-pdf-header-footer"],
        ["--headless", "--disable-gpu", "--print-to-pdf-no-header"],
        ["--headless=new", "--disable-gpu"],
    ]
    last_err = "未知错误"
    for flags in attempts:
        target, err = _run(flags)
        if target is not None:
            os.replace(target, pdf_out)
            return
        last_err = err or "浏览器已退出但未产出 PDF"
    raise RuntimeError("无头浏览器打印 PDF 失败: " + last_err)


def verify_pdf(path: Path) -> int:
    """校验 PDF 文件头并统计页数（按 /Type /Page 计数，粗略）。"""
    data = path.read_bytes()
    if not data.startswith(b"%PDF-"):
        raise RuntimeError(f"输出文件不是有效 PDF: {path}")
    pages = len(re.findall(rb"/Type\s*/Page[^s]", data))
    return max(pages, 1)


def main() -> int:
    ap = argparse.ArgumentParser(description="Markdown 转排版 PDF")
    ap.add_argument("input", help="输入 .md 文件")
    ap.add_argument("output", help="输出 .pdf 文件路径")
    ap.add_argument("--style", default=str(DEFAULT_STYLE), help="风格 md 文件（含 css 代码块）")
    ap.add_argument("--title", default=None, help="文档标题；缺省取文件名，md 自带 # 一级标题时不注入")
    ap.add_argument("--browser", default=None, help="浏览器可执行文件路径（可选）")
    args = ap.parse_args()

    md = Path(args.input).resolve()
    out = Path(args.output).resolve()
    style = Path(args.style).resolve() if args.style else DEFAULT_STYLE
    if not md.exists():
        log(f"[错误] 输入文件不存在: {md}")
        return 1

    title = args.title if args.title is not None else detect_auto_title(md)
    log(f"[1/4] 读取风格: {style.name}")
    css, extras = extract_css_and_extras(style)
    if extras:
        log(f"      风格附带 {len(extras)} 条浏览器参数（如页码模板）")

    with tempfile.TemporaryDirectory(prefix="md2pdf_") as td:
        workdir = Path(td)
        header_html = workdir / "style_head.html"
        header_html.write_text(f"<style>\n{css}\n</style>\n", encoding="utf-8")
        html_out = workdir / "doc.html"

        pandoc = find_pandoc()
        log(f"[2/4] pandoc 生成 HTML（标题: {title or '无'}）")
        run_pandoc(pandoc, md, header_html, html_out, title)

        browser = find_browser(args.browser)
        log(f"[3/4] 无头浏览器打印: {Path(browser).name}")
        print_to_pdf(browser, html_out, workdir / "final.pdf", workdir, extras=extras)

        # 输出到 ASCII 临时路径后移动到最终位置，规避中文/空格路径问题。
        # 注意用 shutil.move 而非 os.replace：临时目录与目标可能跨盘符（如 C 盘 Temp → E 盘项目），
        # os.replace 跨盘会报 WinError 17。
        if not out.parent.exists():
            out.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(workdir / "final.pdf"), str(out))

    pages = verify_pdf(out)
    log(f"[4/4] 完成: {out}（{out.stat().st_size} 字节, {pages} 页）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
