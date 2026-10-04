#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx 编号引擎：还原每段真实编号（Word/WPS 多级编号语义）。

直接解析 docx 包内的 word/numbering.xml + word/document.xml（+ styles.xml
样式级编号），按文档顺序产出每个带编号段落的真实编号文本。供
`html_to_markdown.py --docx`（方案C 编号回填）与
`test-case/verify_numbering_backfill.py`（回归比对）共用。

语义要点（由 test-case-1 参考答案反推并 527/527 全量吻合，fixture 固化）：
  1. 计数器按 numId 隔离：每个 w:num 是独立列表实例，互不延续；
  2. 级别状态分「未启用」与「计数值」：未启用级渲染为该级 start 值
     （因此无 ilvl=0 段落时 %1 恒显示 start）；被更深段落清零的级计数值
     为 0，渲染为字面 0（如 1.2.1.3.4.0.1）；
  3. 段落使用 ilvl=i：本级取 start（未启用/清零）或 +1（已启用），
     更深层清零；
  4. numId 上的 startOverride：首次触达该 numId 且段落级别覆盖到该级时，
     计数器钉为覆盖值，且本级本次跳过递增；
  5. w:lvlRestart：0 = 不因更高级段落重置；k = 仅在遇到 ilvl=k 段落时
     重置；缺省 = 遇任何更高（更小 ilvl）段落重置；
  6. w:isLgl：该级被 %k 引用时强制按 decimal 渲染；
  7. w:lvlOverride 内的完整 w:lvl 重定义优先于 abstractNum 的同级定义。

边界处理：
  - mc:AlternateContent 的 Fallback 内容剔除（与 Choice 去重）；
  - 段落级删除修订（w:pPr/w:rPr/w:del）不计数、不产出；
  - run 级删除文本（w:delText）不进入段落文本；
  - numId=0 视为无编号；
  - bullet / none / 无 % 占位符的级别照常参与计数，但不产出编号文本。
"""

import re
import zipfile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006/main"


def _q(tag):
    return "{%s}%s" % (W_NS, tag)


# --------------------------------------------------------------- numFmt 渲染

_CN_DIGITS = "一二三四五六七八九"
_CN_LEGAL = "壹贰叁肆伍陆柒捌玖"
_CN_TENS = {"cn": ("十", "百", "千"), "legal": ("拾", "佰", "仟")}
_CN_ZERO = {"cn": "零", "legal": "零"}
_IDEO_DIGITS = "〇一二三四五六七八九"

_CIRCLE_RANGES = ((1, 20, 0x2460), (21, 35, 0x3251), (36, 50, 0x32B1))

_ROMAN = ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
          (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
          (5, "v"), (4, "iv"), (1, "i"))

_AE = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわゐゑをん"
_IROHA = "いろはにほへとちりぬるをわかよたれそつねならむうゐのおくやまけふこえてあさきゆめみしゑひもせす"

_ONES_EN = ["one", "two", "three", "four", "five", "six", "seven", "eight",
            "nine", "ten", "eleven", "twelve", "thirteen", "fourteen",
            "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]
_TENS_EN = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy",
            "eighty", "ninety"]


def _letters(n, base):
    """字母序号，超 26 进位：a..z, aa, ab…"""
    out = []
    while n > 0:
        n, r = divmod(n - 1, 26)
        out.append(chr(base + r))
    return "".join(reversed(out))


def _roman(n, upper):
    out = []
    for v, s in _ROMAN:
        while n >= v:
            out.append(s)
            n -= v
    s = "".join(out) or "i"
    return s.upper() if upper else s


def _chinese(n, digits, units, zero):
    """中文数字（含零规则）：105 -> 一百零五，110 -> 一百一十，11 -> 十一。"""
    if n <= 0 or n >= 10000:
        return str(n)
    result, pending_zero = "", False
    for i, ch in enumerate(str(n)):
        d = int(ch)
        pos = len(str(n)) - i - 1  # 0=个位 1=十位 2=百位 3=千位
        if d == 0:
            if result:
                pending_zero = True
            continue
        if pending_zero:
            result += zero
            pending_zero = False
        if pos == 0:
            result += digits[d - 1]
        elif pos == 1 and d == 1 and not result:
            result += units[0]  # 10 -> 十（十一，非一十）
        else:
            result += digits[d - 1] + units[pos - 1]
    return result


def _en_cardinal(n):
    if n < 20:
        return _ONES_EN[n - 1]
    ten, one = divmod(n, 10)
    if ten < 10:
        return _TENS_EN[ten] + ("-" + _ONES_EN[one - 1] if one else "")
    if n < 1000:
        hundred, rest = divmod(n, 100)
        base = _ONES_EN[hundred - 1] + " hundred"
        return base + (" and " + _en_cardinal(rest) if rest else "")
    if n < 1000000:
        thousand, rest = divmod(n, 1000)
        return _en_cardinal(thousand) + " thousand" + (
            " " + _en_cardinal(rest) if rest else "")
    return str(n)


def _en_ordinal_word(n):
    c = _en_cardinal(n)
    specials = {"one": "first", "two": "second", "three": "third",
                "five": "fifth", "eight": "eighth", "nine": "ninth",
                "twelve": "twelfth"}
    if c in specials:
        return specials[c]
    if c.endswith("y"):
        return c[:-1] + "ieth"
    return c + ("th" if c.endswith(("th", "four", "six", "seven",
                                    "ten", "eleven")) else "th")


def _en_ordinal_suffix(n):
    if 11 <= n % 100 <= 13:
        return "%dth" % n
    return "%d%s" % (n, {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th"))


def format_counter(value, fmt):
    """按 numFmt 把计数值渲染为编号片段。未识别格式兜底十进制。"""
    if fmt == "decimal":
        return str(value)
    if fmt == "lowerLetter":
        return _letters(value, ord("a"))
    if fmt == "upperLetter":
        return _letters(value, ord("A"))
    if fmt == "lowerRoman":
        return _roman(value, False)
    if fmt == "upperRoman":
        return _roman(value, True)
    if fmt in ("chineseCounting", "chineseCountingThousand"):
        return _chinese(value, _CN_DIGITS, _CN_TENS["cn"], _CN_ZERO["cn"])
    if fmt == "chineseLegalSimplified":
        return _chinese(value, _CN_LEGAL, _CN_TENS["legal"], _CN_ZERO["legal"])
    if fmt == "ideographDigital":
        return "".join(_IDEO_DIGITS[int(d)] for d in str(value)) \
            if value < 10000 else str(value)
    if fmt == "decimalEnclosedCircle":
        for lo, hi, base in _CIRCLE_RANGES:
            if lo <= value <= hi:
                return chr(base + value - lo)
        return str(value)
    if fmt == "aiueo":
        return _AE[(value - 1) % len(_AE)]
    if fmt == "iroha":
        return _IROHA[(value - 1) % len(_IROHA)]
    if fmt == "ordinal":
        return _en_ordinal_suffix(value)
    if fmt == "ordinalText":
        return _en_ordinal_word(value)
    if fmt == "cardinalText":
        return _en_cardinal(value)
    if fmt == "hex":
        return "%X" % value
    return str(value)  # bullet/none 等在调用方已拦截，其余兜底十进制


# --------------------------------------------------------------- 文本归一化

def normalize_text(s):
    """段落文本对齐用的归一化：去除所有空白字符。"""
    return re.sub(r"\s+", "", s or "")


# -------------------------------------------------------------------- 引擎

class DocxNumbering:
    UNSET = None  # 计数器缺失 = 该级未启用（渲染取 start）

    def __init__(self, docx_path):
        zf = zipfile.ZipFile(docx_path)
        self.num_map, self.abstract_map = self._parse_numbering(zf)
        self.styles = self._parse_styles(zf)
        self.doc = etree.fromstring(zf.read("word/document.xml"))
        self.counters = {}        # (numId, ilvl) -> int；缺失 = 未启用
        self.override_done = set()

    # ---------------- parsing ----------------

    @staticmethod
    def _parse_numbering(zf):
        num_map, abstract_map = {}, {}
        try:
            root = etree.fromstring(zf.read("word/numbering.xml"))
        except KeyError:
            return num_map, abstract_map

        for absnum in root.iter(_q("abstractNum")):
            aid = absnum.get(_q("abstractNumId"))
            lvls = {}
            for lvl in absnum.iter(_q("lvl")):
                ilvl = int(lvl.get(_q("ilvl")))
                start = lvl.find(_q("start"))
                fmt = lvl.find(_q("numFmt"))
                text = lvl.find(_q("lvlText"))
                restart = lvl.find(_q("lvlRestart"))
                lvls[ilvl] = {
                    "start": int(start.get(_q("val"))) if start is not None else 1,
                    "fmt": fmt.get(_q("val")) if fmt is not None else "decimal",
                    "text": text.get(_q("val")) if text is not None else "",
                    "restart": int(restart.get(_q("val"))) if restart is not None else None,
                    "is_lgl": lvl.find(_q("isLgl")) is not None,
                }
            abstract_map[aid] = lvls

        for num in root.iter(_q("num")):
            nid = num.get(_q("numId"))
            abs_el = num.find(_q("abstractNumId"))
            entry = {"abstract": abs_el.get(_q("val")) if abs_el is not None else None,
                     "override": {},   # ilvl -> startOverride 值
                     "redef": {}}      # ilvl -> 完整 lvl 重定义
            for ovr in num.iter(_q("lvlOverride")):
                ilvl = int(ovr.get(_q("ilvl")))
                so = ovr.find(_q("startOverride"))
                if so is not None:
                    entry["override"][ilvl] = int(so.get(_q("val")))
                full = ovr.find(_q("lvl"))
                if full is not None:
                    start = full.find(_q("start"))
                    fmt = full.find(_q("numFmt"))
                    text = full.find(_q("lvlText"))
                    restart = full.find(_q("lvlRestart"))
                    entry["redef"][ilvl] = {
                        "start": int(start.get(_q("val"))) if start is not None else 1,
                        "fmt": fmt.get(_q("val")) if fmt is not None else "decimal",
                        "text": text.get(_q("val")) if text is not None else "",
                        "restart": int(restart.get(_q("val"))) if restart is not None else None,
                        "is_lgl": full.find(_q("isLgl")) is not None,
                    }
            num_map[nid] = entry
        return num_map, abstract_map

    @staticmethod
    def _parse_styles(zf):
        """styleId -> {basedOn, numId, ilvl}（样式级编号定义）。"""
        try:
            root = etree.fromstring(zf.read("word/styles.xml"))
        except KeyError:
            return {}
        styles = {}
        for st in root.iter(_q("style")):
            sid = st.get(_q("styleId"))
            based = st.find(_q("basedOn"))
            numpr = st.find("%s/%s" % (_q("pPr"), _q("numPr")))
            numid = ilvl = None
            if numpr is not None:
                e1, e2 = numpr.find(_q("numId")), numpr.find(_q("ilvl"))
                numid = e1.get(_q("val")) if e1 is not None else None
                ilvl = int(e2.get(_q("val"))) if e2 is not None else None
            styles[sid] = {"basedOn": based.get(_q("val")) if based is not None else None,
                           "numId": numid, "ilvl": ilvl}
        return styles

    def _style_numbering(self, style_id):
        """沿 basedOn 链回溯样式编号定义（带环防护）。"""
        seen = set()
        while style_id and style_id in self.styles and style_id not in seen:
            seen.add(style_id)
            s = self.styles[style_id]
            if s["numId"] is not None:
                return s["numId"], s["ilvl"] if s["ilvl"] is not None else 0
            style_id = s["basedOn"]
        return None, None

    # ---------------- 段落遍历 ----------------

    def _fallback_paras(self):
        """mc:AlternateContent Fallback 内的段落集合（去重用）。"""
        excluded = set()
        for fb in self.doc.iter("{%s}Fallback" % MC_NS):
            for p in fb.iter(_q("p")):
                excluded.add(p)
        return excluded

    def _iter_paras(self, excluded):
        for p in self.doc.iter(_q("p")):
            if p in excluded:
                continue
            in_table = False
            node = p.getparent()
            while node is not None:
                if node.tag == _q("tbl"):
                    in_table = True
                    break
                node = node.getparent()
            yield p, in_table

    @staticmethod
    def para_text(p):
        parts = []
        for node in p.iter():
            if node.tag == _q("t"):
                parts.append(node.text or "")
            elif node.tag in (_q("tab"), _q("br"), _q("cr")):
                parts.append(" ")
        return "".join(parts).strip()

    @staticmethod
    def para_bold(p):
        for rpr in p.iter(_q("rPr")):
            b = rpr.find(_q("b"))
            if b is not None and b.get(_q("val")) not in ("0", "false", "none"):
                return True
        return False

    @staticmethod
    def _para_deleted(p):
        """段落整体被删除修订标记（接受修订后不存在）。"""
        ppr = p.find(_q("pPr"))
        if ppr is None:
            return False
        rpr = ppr.find(_q("rPr"))
        return rpr is not None and rpr.find(_q("del")) is not None

    def _para_numpr(self, p):
        """直接 numPr 优先；否则回溯 pStyle 样式链。"""
        ppr = p.find(_q("pPr"))
        if ppr is None:
            return None, None
        numpr = ppr.find(_q("numPr"))
        if numpr is not None:
            e1, e2 = numpr.find(_q("numId")), numpr.find(_q("ilvl"))
            numid = e1.get(_q("val")) if e1 is not None else None
            ilvl = int(e2.get(_q("val"))) if e2 is not None else 0
            return numid, ilvl
        pstyle = ppr.find(_q("pStyle"))
        if pstyle is not None:
            return self._style_numbering(pstyle.get(_q("val")))
        return None, None

    # ---------------- 计数与渲染 ----------------

    def _effective_levels(self, entry):
        """abstractNum 各级定义 + numId 级完整 lvl 重定义合并。"""
        eff = {i: dict(d) for i, d in self.abstract_map.get(entry["abstract"], {}).items()}
        for i, d in entry["redef"].items():
            eff[i] = dict(d)
        return eff

    def compute(self, numid, ilvl):
        """推进计数器并渲染编号文本；bullet/none 等无编号文本返回 None。"""
        entry = self.num_map.get(numid)
        if entry is None or entry["abstract"] is None:
            return None
        eff = self._effective_levels(entry)
        if ilvl not in eff:
            return None

        # 1) startOverride：首次触达且段落级别覆盖到该级 -> 钉值
        own_overridden = False
        for ovl, start in entry["override"].items():
            if ovl not in eff:
                continue
            if (numid, ovl) not in self.override_done and ilvl >= ovl:
                self.counters[(numid, ovl)] = start
                self.override_done.add((numid, ovl))
                if ovl == ilvl:
                    own_overridden = True

        # 2) 本级计数：未启用/清零 -> start；已启用 -> +1（被钉值则跳过）
        key = (numid, ilvl)
        cur = self.counters.get(key, self.UNSET)
        if not own_overridden:
            if cur is self.UNSET or cur == 0:
                self.counters[key] = eff[ilvl]["start"]
            else:
                self.counters[key] = cur + 1

        # 3) 更深层清零（受 lvlRestart 约束；未启用的级保持未启用态，
        #    仅曾启用的级才进入「清零为 0」状态——渲染语义见要点 2）
        for deeper in range(ilvl + 1, 10):
            if deeper not in eff:
                continue
            if (numid, deeper) not in self.counters:
                continue          # 从未启用：保持未启用（渲染取 start）
            r = eff[deeper].get("restart")
            if r == 0:
                continue          # 0 = 不因更高级重置
            if r is not None and r != ilvl:
                continue          # 仅在遇到指定级时重置
            self.counters[(numid, deeper)] = 0

        # 4) 按 lvlText 渲染（isLgl：本级渲染时所有 %k 强制 decimal）
        fmt, text = eff[ilvl]["fmt"], eff[ilvl]["text"]
        if fmt in ("bullet", "none") or "%" not in text:
            return None
        force_decimal = eff[ilvl].get("is_lgl")

        def repl(m):
            idx = int(m.group(1)) - 1
            if idx not in eff:
                return ""
            cnt = self.counters.get((numid, idx), self.UNSET)
            if cnt is self.UNSET:
                cnt = eff[idx]["start"]
            f = "decimal" if force_decimal else eff[idx]["fmt"]
            return format_counter(cnt, f)

        return re.sub(r"%(\d)", repl, text)

    # ---------------- 公共接口 ----------------

    def numbered_paragraphs(self):
        """按文档顺序产出有编号文本的段落（含表格内；bullet/none 不产出）。"""
        excluded = self._fallback_paras()
        result = []
        for order, (p, in_table) in enumerate(self._iter_paras(excluded), 1):
            if self._para_deleted(p):
                continue
            numid, ilvl = self._para_numpr(p)
            if numid is None or numid == "0" or ilvl is None:
                continue
            number = self.compute(numid, ilvl)
            if number is None:
                continue
            result.append({
                "order": order,
                "text": self.para_text(p),
                "key": normalize_text(self.para_text(p)),
                "number": number,
                "ilvl": ilvl,
                "numid": numid,
                "in_table": in_table,
                "bold": self.para_bold(p),
            })
        return result
