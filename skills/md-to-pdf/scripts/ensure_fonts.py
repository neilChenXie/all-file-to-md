#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ensure_fonts.py —— 字体预检与按需安装（md-to-pdf / md-to-docx 通用）

公文等风格依赖系统已安装的特定中文字体（如 仿宋_GB2312、方正小标宋_GBK、楷体_GB2312）。
若缺失，Word/WPS/浏览器会静默回退（公文体"不生效"的常见原因）。本脚本：

  1. 扫描本 skill 的 assets/fonts/ 目录（也可用 --fonts-dir 指定）下的 ttf/otf/ttc；
  2. 读取每个字体文件的族名（含中文/英文本地化名）；
  3. 检查系统是否已安装同名族字体；
  4. 缺失则安装到【用户字体库】（无需管理员权限）：
     - Windows : 复制到 %LOCALAPPDATA%\\Microsoft\\Windows\\Fonts，
                 并写入 HKCU\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts
                 （新启动的 Word/WPS/Edge 进程即可识别）
     - macOS   : 复制到 ~/Library/Fonts
     - Linux   : 复制到 ~/.fonts（随后执行 fc-cache -f）

用法：
  python ensure_fonts.py                # 预检并安装缺失字体
  python ensure_fonts.py --check-only   # 只报告缺失，不安装
  python ensure_fonts.py --fonts-dir /path/to/fonts

说明：读取族名依赖 fontTools（可选）。未安装时退化为按文件名推断族名，
      若文件名与族名一致（如 仿宋_GB2312.ttf）仍可正常工作。
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

# 脚本所在目录的上一级 = skill 根目录
SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_FONTS_DIR = SKILL_DIR / "assets" / "fonts"

FONT_EXTS = {".ttf", ".otf", ".ttc"}

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


# ----------------------------- 族名读取 --------------------------------
def read_family_names(font_path: Path) -> list[str]:
    """返回字体的族名候选（优先中文 zh-CN 本地化名，其次英文/其他）。"""
    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        # 无 fontTools：按文件名推断
        return [font_path.stem, _ascii_of(font_path.stem)]

    names: list[str] = []
    try:
        font = TTFont(str(font_path), fontNumber=0, lazy=True)
        for rec in font["name"].names:
            if rec.nameID not in (1, 16):
                continue
            try:
                txt = rec.toUnicode().strip()
            except Exception:
                continue
            if txt and txt not in names:
                names.append(txt)
        font.close()
    except Exception as e:
        print(f"[警告] 读取字体 {font_path.name} 族名失败（{e}），按文件名推断")
        names = [font_path.stem, _ascii_of(font_path.stem)]
    if not names:
        names = [font_path.stem, _ascii_of(font_path.stem)]
    return names


def _ascii_of(s: str) -> str:
    # 文件名可能是中文，取相邻可打印 ASCII 片段（有则用）
    m = re.findall(r"[A-Za-z0-9_]+", s)
    return m[0] if m else s


# ----------------------------- 已安装检测 -------------------------------
def _expand_font_display(name: str) -> list[str]:
    """把注册表显示名拆成可匹配的族名候选。

    例如 'SimSun & NSimSun (TrueType)' -> ['SimSun', 'NSimSun']
        '仿宋_GB2312 (TrueType)'       -> ['仿宋_GB2312']
    """
    name = re.sub(r"\s*\((?:TrueType|OpenType|All res|Type 1)\)\s*$", "", name.strip())
    parts = re.split(r"\s*&\s*", name)
    return [p.strip() for p in parts if p.strip()]


def _decode_reg_output(data: bytes) -> str:
    """reg.exe 输出编码因环境而异（UTF-16LE / GBK / UTF-8），依次尝试并取最合理结果。"""
    if not data:
        return ""
    best, best_score = "", -1
    for enc in ("utf-16-le", "gbk", "utf-8"):
        try:
            text = data.decode(enc, errors="replace")
            # 以替换字符（乱码）数量为质量分，越少越好
            score = -text.count("\ufffd")
            if score > best_score:
                best, best_score = text, score
        except Exception:
            continue
    return best


def installed_font_families_win() -> set[str]:
    """Windows：收集系统（HKLM）与用户（HKCU）已注册字体族名。

    首选 `reg query`（兼容沙箱/受限环境对 winreg 读的隔离），
    辅以 winreg 枚举（正常环境下更全）。
    """
    found: set[str] = set()
    for hive in ("HKCU", "HKLM"):
        try:
            proc = subprocess.run(
                ["reg", "query", rf"{hive}\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"],
                capture_output=True, timeout=30,
            )
            text = _decode_reg_output(proc.stdout)
        except Exception:
            continue
        for raw in text.splitlines():
            line = raw.strip()
            if not line or line.startswith("HKEY_") or "REG_SZ" not in line:
                continue
            name = line.split("REG_SZ")[0].strip()
            for nm in _expand_font_display(name):
                found.add(nm)

    # 补充 winreg 枚举（reg query 已覆盖时基本为空操作）
    for hive in ("HKLM", "HKCU"):
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE if hive == "HKLM" else winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
            )
        except OSError:
            continue
        try:
            n = winreg.QueryInfoKey(key)[0]
            for i in range(n):
                try:
                    name, _, _ = winreg.EnumValue(key, i)
                    for nm in _expand_font_display(name):
                        found.add(nm)
                except OSError:
                    break
            key.Close()
        except OSError:
            pass
    return found


def installed_font_families_posix(extra_dirs: list[Path]) -> set[str]:
    """macOS/Linux：扫描系统字体目录中的字体文件并读取族名（仅当 fontTools 可用）。"""
    found: set[str] = set()
    search_dirs = []
    home = Path.home()
    if sys.platform == "darwin":
        search_dirs = [Path("/System/Library/Fonts"), Path("/Library/Fonts"),
                       home / "Library" / "Fonts"]
    else:
        search_dirs = [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
                       home / ".fonts", home / ".local" / "share" / "fonts"]
    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        return found  # 无 fontTools 时 posix 侧无法精确检测，视为未装（安装是幂等的）

    seen = set()
    for d in search_dirs + extra_dirs:
        if not d.exists():
            continue
        for fp in d.rglob("*"):
            if fp.suffix.lower() in FONT_EXTS and str(fp) not in seen:
                seen.add(str(fp))
                for fam in read_family_names(fp):
                    found.add(fam)
    return found


def normalize(fam: str) -> str:
    return fam.lower().replace(" ", "").replace("_", "")


def is_installed(families: list[str], installed: set[str]) -> bool:
    inst_norm = {normalize(x) for x in installed}
    return any(normalize(fam) in inst_norm for fam in families)


# ----------------------------- 安装 ------------------------------------
def install_windows(font_path: Path, families: list[str]) -> Path:
    import winreg
    dest_dir = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / font_path.name
    if not dest.exists() or dest.stat().st_size != font_path.stat().st_size:
        shutil.copy2(font_path, dest)
    # 注册显示名："中文族名 (TrueType)"（找不到中文名用英文名）
    display = next((f for f in families if not f.isascii()), families[0])
    value_name = f"{display} (TrueType)"
    reg_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    # 1) winreg 写入
    try:
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, str(dest))
        key.Close()
    except OSError:
        pass
    # 2) 交叉验证；受限环境下 winreg 读取不可靠时改用 reg.exe 兜底
    ok = False
    try:
        proc = subprocess.run(
            ["reg", "query", rf"HKCU\{reg_path}", "/v", value_name],
            capture_output=True, timeout=30,
        )
        ok = proc.returncode == 0
    except Exception:
        ok = False
    if not ok:
        subprocess.run(
            ["reg", "add", rf"HKCU\{reg_path}", "/v", value_name,
             "/t", "REG_SZ", "/d", str(dest), "/f"],
            capture_output=True, timeout=30,
        )
    return dest


def install_posix(font_path: Path, families: list[str]) -> Path:
    home = Path.home()
    if sys.platform == "darwin":
        dest_dir = home / "Library" / "Fonts"
    else:
        dest_dir = home / ".fonts"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / font_path.name
    if not dest.exists() or dest.stat().st_size != font_path.stat().st_size:
        shutil.copy2(font_path, dest)
    if sys.platform != "darwin":
        try:
            subprocess.run(["fc-cache", "-f", str(dest_dir)], capture_output=True, timeout=120)
        except Exception:
            pass
    return dest


def ensure_fonts(fonts_dir: Path, check_only: bool) -> int:
    fonts_dir = fonts_dir.resolve()
    if not fonts_dir.exists():
        print(f"[信息] 字体目录不存在，跳过: {fonts_dir}")
        return 0
    font_files = sorted(p for p in fonts_dir.iterdir() if p.suffix.lower() in FONT_EXTS)
    if not font_files:
        print(f"[信息] {fonts_dir} 下没有 ttf/otf/ttc 字体文件")
        return 0

    # 已安装族集合
    installed: set[str] = set()
    if sys.platform == "win32":
        installed = installed_font_families_win()
    else:
        installed = installed_font_families_posix([fonts_dir])

    changed = 0
    for fp in font_files:
        families = read_family_names(fp)
        if is_installed(families, installed):
            print(f"[已安装] {fp.name}（族名: {' / '.join(families)}）")
            continue
        if check_only:
            print(f"[缺失]   {fp.name}（族名: {' / '.join(families)}）→ 需安装")
            changed += 1
            continue
        try:
            if sys.platform == "win32":
                dest = install_windows(fp, families)
            else:
                dest = install_posix(fp, families)
            print(f"[已安装→] {fp.name} → {dest}")
            changed += 1
        except Exception as e:
            print(f"[失败]   安装 {fp.name}: {e}")
            return 2

    if check_only:
        print("完成（--check-only，未执行安装）" if changed == 0 else f"共 {changed} 个字体缺失")
    else:
        if changed == 0:
            print("全部字体已就绪")
        else:
            print(f"安装完成：{changed} 个字体；新启动的 Word/WPS/浏览器进程即可识别")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="字体预检与按需安装（扫描 assets/fonts）")
    ap.add_argument("--check-only", action="store_true", help="只报告缺失，不安装")
    ap.add_argument("--fonts-dir", default=str(DEFAULT_FONTS_DIR), help="字体所在目录")
    args = ap.parse_args()
    return ensure_fonts(Path(args.fonts_dir), args.check_only)


if __name__ == "__main__":
    sys.exit(main())
