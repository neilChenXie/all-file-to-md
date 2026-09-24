#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx -> HTML 转换（pandoc 回退方案）。

适用场景：
- 本机未安装 WPS（docx_to_html.ps1 的 WPS COM 路径不可用）；
- WPS 程序打开着目标文件导致 COM 连接失败、提示用户关闭后重试仍失败。

依赖：pypandoc-binary（自带 pandoc 可执行文件，无需单独安装 pandoc）：
  pip install pypandoc-binary -i https://pypi.tuna.tsinghua.edu.cn/simple

用法：
  python docx_to_html_pandoc.py <input.docx> [output.html]

行为：
- 未指定 output.html 时，默认输出到输入文件同目录同名 .html；
- 图片提取到 <output 同目录>/<文件名>.files/media/ 子目录（注意：比 WPS 导出多一层
  media/，后续取图需按此路径，或先平铺到 .files/ 下）；
- 输出为 UTF-8 标准 HTML5，html_to_markdown.py 可直接处理；
- 退出码：0 成功，1 失败，2 参数错误。
"""

import os
import subprocess
import sys


def find_pandoc():
    """定位 pandoc 可执行文件：优先 pypandoc-binary 自带，其次 PATH。

    注意：Windows 上 pypandoc.get_pandoc_path() 返回的路径不带 .exe 后缀，
    需要尝试补 .exe 再判断存在性。
    """
    try:
        import pypandoc
        exe = pypandoc.get_pandoc_path()
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
    return None


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

    pandoc = find_pandoc()
    if not pandoc or not os.path.exists(pandoc):
        print("[ERROR] 未找到 pandoc。请先在当前 Python 环境安装 pypandoc-binary：")
        print("  pip install pypandoc-binary -i https://pypi.tuna.tsinghua.edu.cn/simple")
        print("说明：pypandoc-binary 自带 pandoc 可执行文件，无需单独安装 pandoc。")
        return 1

    print("Converting: %s" % in_path)
    print("Output: %s" % out_path)

    cmd = [
        pandoc,
        "--standalone",
        "--from", "docx",
        "--to", "html",
        "--extract-media=%s" % media_dir,
        "--output=%s" % out_path,
        in_path,
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True)
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

    size_kb = os.path.getsize(out_path) / 1024.0
    media_sub = os.path.join(media_dir, "media")
    if os.path.isdir(media_sub) and os.listdir(media_sub):
        print("[OK] Conversion successful! (含图片，位于 %s，注意 media/ 子目录)" % media_sub)
    else:
        print("[OK] Conversion successful! (文档无图片)")
    print("[OK] Size: %.2f KB" % size_kb)
    return 0


if __name__ == "__main__":
    sys.exit(main())
