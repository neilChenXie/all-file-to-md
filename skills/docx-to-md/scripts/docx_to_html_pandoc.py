#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx -> HTML 转换（pandoc 回退方案 C）。

适用场景（作为最后回退，方案A、方案B 均不可行时）：
- 用户无法将 docx 转存为 PDF（如未安装 WPS / Word，或不接受手动操作）；
- 本机 COM 自动化不可用（未安装 WPS / Microsoft Word、COM 注册异常、程序占用文件等）。

依赖：pypandoc-binary（自带 pandoc 可执行文件，无需单独安装 pandoc）：
  pip install pypandoc-binary -i https://pypi.tuna.tsinghua.edu.cn/simple

用法：
  python docx_to_html_pandoc.py <input.docx> [output.html]

行为：
- 未指定 output.html 时，默认输出到输入文件同目录同名 .html；
- 图片提取到 <output 同目录>/<文件名>.files/ 下（相对路径引用，与 WPS 导出的
  方案A 路径约定一致；pandoc 原生输出的 media/ 子目录已自动平铺，无需手工处理）；
- 输出为 UTF-8 标准 HTML5，html_to_markdown.py 可直接处理；
- 退出码：0 成功，1 失败，2 参数错误。
"""

import os
import subprocess
import sys


def _pandoc_fallback_paths():
    """pandoc 常见安装位置（pypandoc 与 PATH 均未找到时的兜底）。

    部分执行环境（IDE 拉起的子进程、cron、GUI 应用）的 PATH 可能缺少
    Homebrew 等目录，此处按平台补充常见安装位置。
    """
    home = os.path.expanduser("~")
    if os.name == "nt":
        appdata = os.environ.get("LOCALAPPDATA", "")
        progfiles = os.environ.get("ProgramFiles", r"C:\Program Files")
        progfiles86 = os.environ.get("ProgramFiles(x86)", "")
        return [
            os.path.join(progfiles, "Pandoc", "pandoc.exe"),
            os.path.join(progfiles86, "Pandoc", "pandoc.exe") if progfiles86 else "",
            os.path.join(appdata, "Pandoc", "pandoc.exe") if appdata else "",
            os.path.join(home, "AppData", "Local", "Pandoc", "pandoc.exe"),
        ]
    return [
        "/opt/homebrew/bin/pandoc",  # macOS Homebrew (Apple Silicon)
        "/usr/local/bin/pandoc",     # macOS Homebrew (Intel) / 手动安装
        "/usr/bin/pandoc",           # Linux 发行版包管理器
        "/snap/bin/pandoc",          # Linux snap
        os.path.join(home, ".local", "bin", "pandoc"),
    ]


def find_pandoc():
    """定位 pandoc 可执行文件：优先 pypandoc-binary 自带，其次 PATH，最后常见安装位置。

    注意：
    - Windows 上 pypandoc.get_pandoc_path() 返回的路径不带 .exe 后缀，
      需要尝试补 .exe 再判断存在性；
    - 只安装了 pypandoc（非 pypandoc-binary）且 PATH 中无 pandoc 时，
      get_pandoc_path() 会抛 OSError，必须捕获后继续回退，不能让脚本崩溃。
    """
    try:
        import pypandoc
        try:
            exe = pypandoc.get_pandoc_path()
        except Exception:  # noqa: BLE001 pypandoc 未找到 pandoc 时抛 OSError
            exe = None
        if exe:
            if os.path.exists(exe):
                return exe
            if os.name == "nt" and os.path.exists(exe + ".exe"):
                return exe + ".exe"
    except ImportError:
        pass
    from shutil import which
    exe = which("pandoc")
    if exe:
        return exe
    for cand in _pandoc_fallback_paths():
        if cand and os.path.isfile(cand):
            return cand
    return None


def _rewrite_html_refs(out_path, old, new):
    """把 HTML 中的引用前缀 old 替换为 new（存在才改写）。"""
    with open(out_path, "r", encoding="utf-8") as f:
        html = f.read()
    if old in html:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html.replace(old, new))
        return True
    return False


def _flatten_media(out_path):
    """把 pandoc 输出的 .files/media/ 子目录图片平铺到 .files/ 根。

    pandoc 原生把媒体放入 <extract-media 目录>/media/ 一层子目录，比方案A
    （WPS 导出）多一层；此处移动文件并同步改写 HTML 引用，使方案C 与方案A
    的图片目录约定完全一致。同一 pandoc 运行内媒体名天然唯一，同名即旧产物，
    直接覆盖。返回移动的文件数。
    """
    out_dir = os.path.dirname(out_path)
    base = os.path.splitext(os.path.basename(out_path))[0]
    files_dir = os.path.join(out_dir, base + ".files")
    media_dir = os.path.join(files_dir, "media")
    if not os.path.isdir(media_dir):
        return 0

    moved = 0
    for name in sorted(os.listdir(media_dir)):
        src = os.path.join(media_dir, name)
        if not os.path.isfile(src):
            continue
        os.replace(src, os.path.join(files_dir, name))
        moved += 1
    try:
        os.rmdir(media_dir)
    except OSError:
        pass

    _rewrite_html_refs(out_path, base + ".files/media/", base + ".files/")
    return moved


def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("用法: python docx_to_html_pandoc.py <input.docx> [output.html]")
        return 2

    in_path = os.path.abspath(sys.argv[1])
    if not os.path.isfile(in_path):
        print("[ERROR] 输入文件不存在: %s" % in_path)
        return 1

    if len(sys.argv) == 3:
        out_path = os.path.abspath(sys.argv[2])
    else:
        out_path = os.path.splitext(in_path)[0] + ".html"

    out_dir = os.path.dirname(out_path)
    os.makedirs(out_dir, exist_ok=True)
    media_dir = os.path.join(
        out_dir, os.path.splitext(os.path.basename(out_path))[0] + ".files")
    # .files/ 是本脚本专属输出目录：清理旧产物，避免残留文件与引用错位
    if os.path.isdir(media_dir):
        import shutil
        shutil.rmtree(media_dir)

    pandoc = find_pandoc()
    if not pandoc:
        print("[ERROR] 未找到 pandoc。可按以下任一方式安装后重试：")
        print("  1) 在当前 Python 环境安装 pypandoc-binary（自带 pandoc，推荐）：")
        print("     pip install pypandoc-binary -i https://pypi.tuna.tsinghua.edu.cn/simple")
        if os.name == "nt":
            print("  2) winget install --id JohnMacFarlane.Pandoc")
        else:
            print("  2) brew install pandoc（macOS）或用系统包管理器安装 pandoc（Linux）")
        return 1

    print("pandoc: %s" % pandoc)
    print("Converting: %s" % in_path)
    print("Output: %s" % out_path)

    cmd = [
        pandoc,
        "--standalone",
        "--from", "docx",
        "--to", "html",
        # 以 out_dir 为工作目录并传相对路径：HTML 内图片引用为相对路径
        # （传绝对路径时 pandoc 会把绝对路径写进引用，产物不可移植）
        "--extract-media=%s" % os.path.basename(media_dir),
        "--output=%s" % out_path,
        in_path,
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, cwd=out_dir)
    except Exception as e:  # noqa: BLE001
        print("[ERROR] 调用 pandoc 失败: %s" % e)
        return 1

    if proc.returncode != 0:
        err = proc.stderr.decode("utf-8", errors="replace").strip()
        print("[ERROR] pandoc 转换失败 (exit %s): %s" % (proc.returncode, err))
        return 1

    if not os.path.isfile(out_path):
        print("[ERROR] pandoc 执行完成但未生成输出文件")
        return 1

    moved = _flatten_media(out_path)

    size_kb = os.path.getsize(out_path) / 1024.0
    if moved:
        print("[OK] Conversion successful! (含 %d 张图片，已提取至 %s)"
              % (moved, media_dir))
    else:
        print("[OK] Conversion successful! (文档无图片)")
    print("[OK] Size: %.2f KB" % size_kb)
    return 0


if __name__ == "__main__":
    sys.exit(main())
