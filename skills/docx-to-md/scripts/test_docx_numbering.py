#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_numbering.py 编号引擎单元测试。

不依赖仓库内的真实 docx：在临时目录用 zipfile 内存构造最小 docx
（word/document.xml + word/numbering.xml [+ styles.xml]），逐用例断言
`docx_numbering.py` 引擎（DocxNumbering）还原的编号序列。
与被测引擎同目录，随引擎一起被 git 跟踪；引擎语义变更时先改这里。

覆盖（含 test-case-1 无真实样例的格式）：
  1.  chineseCounting / chineseLegalSimplified / decimalEnclosedCircle /
      ideographDigital 等 numFmt 渲染；
  2.  lowerLetter / upperLetter 超 26 进位（z -> aa）；
  3.  w:lvlRestart（0 = 不因更高级重置）；
  4.  w:isLgl（引用级强制 decimal 渲染）；
  5.  w:lvlOverride 完整 lvl 重定义；
  6.  既有语义回归：per-numId 计数隔离、未启用级渲染 start、被清零级渲染
      字面 0、startOverride 钉值跳递增；
  7.  mc:AlternateContent Choice/Fallback 去重、修订删除段落不计数、
      numId=0 无编号、表格内段落标记。

运行：python skills/docx-to-md/scripts/test_docx_numbering.py（任意工作目录均可）
"""

import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_numbering import DocxNumbering, normalize_text  # noqa: E402

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006/main"'

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/numbering.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    '</Types>'
)


# ---------------------------------------------------------------- 构造辅助

def para(text="", numid=None, ilvl=0, deleted=False, bold=False):
    """生成一个 w:p。deleted=True 时段落整体带删除修订标记。"""
    ppr = ""
    if numid is not None:
        del_mark = '<w:rPr><w:del w:id="9" w:author="t"/></w:rPr>' if deleted else ""
        ppr = (f'<w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/>'
               f'<w:numId w:val="{numid}"/></w:numPr>{del_mark}</w:pPr>')
    rpr = '<w:rPr><w:b/></w:rPr>' if bold else ""
    run = f'<w:r>{rpr}<w:t>{text}</w:t></w:r>' if text else ""
    return f"<w:p>{ppr}{run}</w:p>"


def lvl(ilvl, fmt, text, start=1, restart=None, is_lgl=False):
    extra = f'<w:lvlRestart w:val="{restart}"/>' if restart is not None else ""
    lgl = "<w:isLgl/>" if is_lgl else ""
    return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="{start}"/>'
            f'<w:numFmt w:val="{fmt}"/>{extra}{lgl}'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/></w:lvl>')


def numbering_xml(abstract_id, lvls, nums):
    body = f'<w:abstractNum w:abstractNumId="{abstract_id}">{lvls}</w:abstractNum>'
    body += "".join(
        f'<w:num w:numId="{nid}"><w:abstractNumId w:val="{aid}"/>{ovr}</w:num>'
        for nid, aid, ovr in nums)
    return f'<w:numbering {W_NS}>{body}</w:numbering>'


def document_xml(body):
    return f'<w:document {W_NS} {MC_NS}><w:body>{body}</w:body></w:document>'


def build_docx(document, numbering=None):
    tmpdir = tempfile.mkdtemp(prefix="numfix-")
    path = os.path.join(tmpdir, "fixture.docx")
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("word/document.xml", document)
        if numbering:
            z.writestr("word/numbering.xml", numbering)
    return path


def numbers_of(docx_path):
    return [p["number"] for p in DocxNumbering(docx_path).numbered_paragraphs()]


# ---------------------------------------------------------------- 用例

class NumFmtCase(unittest.TestCase):
    """numFmt 渲染覆盖。"""

    def test_chinese_counting(self):
        """一…十、十一、二十一；105 -> 一百零五。"""
        doc = document_xml(
            para("a", 1, 0) + para("b", 1, 0) + para("c", 1, 0)
            + para("d", 1, 0) + para("e", 1, 0) + para("f", 1, 0)
            + para("g", 1, 0) + para("h", 1, 0) + para("i", 1, 0)
            + para("j", 1, 0) + para("k", 1, 0) + para("l", 1, 0))
        num = numbering_xml("1", lvl(0, "chineseCounting", "%1、"),
                            [(1, "1", "")])
        self.assertEqual(numbers_of(build_docx(doc, num)),
                         ["一、", "二、", "三、", "四、", "五、", "六、",
                          "七、", "八、", "九、", "十、", "十一、", "十二、"])

    def test_chinese_counting_hundred_with_zero(self):
        doc = document_xml(para("a", 1, 0) * 110)
        num = numbering_xml("1", lvl(0, "chineseCounting", "%1、"), [(1, "1", "")])
        got = numbers_of(build_docx(doc, num))
        self.assertEqual(got[9], "十、")
        self.assertEqual(got[99], "一百、")
        self.assertEqual(got[104], "一百零五、")
        self.assertEqual(got[109], "一百一十、")

    def test_chinese_legal_simplified(self):
        doc = document_xml(para("a", 1, 0) * 4)
        num = numbering_xml("1", lvl(0, "chineseLegalSimplified", "第%1章"),
                            [(1, "1", "")])
        self.assertEqual(numbers_of(build_docx(doc, num)),
                         ["第壹章", "第贰章", "第叁章", "第肆章"])

    def test_enclosed_circle(self):
        """①…⑳、㉑、㊱。"""
        doc = document_xml(para("a", 1, 0) * 36)
        num = numbering_xml("1", lvl(0, "decimalEnclosedCircle", "%1"),
                            [(1, "1", "")])
        got = numbers_of(build_docx(doc, num))
        self.assertEqual(got[0], "①")
        self.assertEqual(got[19], "⑳")
        self.assertEqual(got[20], "㉑")
        self.assertEqual(got[35], "㊱")

    def test_ideograph_digital(self):
        """逐位直译：10 -> 一〇、12 -> 一二（数字 0 用 〇）。"""
        doc = document_xml(para("a", 1, 0) * 12)
        num = numbering_xml("1", lvl(0, "ideographDigital", "%1、"), [(1, "1", "")])
        got = numbers_of(build_docx(doc, num))
        self.assertEqual(got[0], "一、")
        self.assertEqual(got[9], "一〇、")
        self.assertEqual(got[11], "一二、")

    def test_lower_letter_wrap(self):
        """第 26 项 = z，第 27 项 = aa（进位，非回绕）。"""
        doc = document_xml(para("a", 1, 0) * 28)
        num = numbering_xml("1", lvl(0, "lowerLetter", "%1)"), [(1, "1", "")])
        got = numbers_of(build_docx(doc, num))
        self.assertEqual(got[0], "a)")
        self.assertEqual(got[25], "z)")
        self.assertEqual(got[26], "aa)")
        self.assertEqual(got[27], "ab)")

    def test_upper_letter_wrap(self):
        doc = document_xml(para("a", 1, 0) * 27)
        num = numbering_xml("1", lvl(0, "upperLetter", "%1."), [(1, "1", "")])
        got = numbers_of(build_docx(doc, num))
        self.assertEqual(got[25], "Z.")
        self.assertEqual(got[26], "AA.")

    def test_ordinal_variants(self):
        doc = document_xml(para("a", 1, 0) * 4)
        num = numbering_xml(
            "1",
            lvl(0, "ordinal", "%1") + lvl(1, "ordinalText", "%2")
            + lvl(2, "cardinalText", "%3"),
            [(1, "1", "")])
        doc2 = document_xml(
            para("a", 1, 0) + para("b", 1, 1) + para("c", 1, 2) + para("d", 1, 2))
        self.assertEqual(numbers_of(build_docx(doc2, num)),
                         ["1st", "first", "one", "two"])

    def test_hex_and_none(self):
        """hex 渲染；none 级段落参与本级计数但不产出编号文本。"""
        doc = document_xml(para("a", 1, 0) + para("b", 1, 0) + para("c", 1, 0))
        num = numbering_xml("1", lvl(0, "hex", "0x%1") + lvl(1, "none", "%2"),
                            [(1, "1", "")])
        doc2 = document_xml(para("a", 1, 0) + para("b", 1, 1) + para("c", 1, 0))
        self.assertEqual(numbers_of(build_docx(doc2, num)),
                         ["0x1", "0x2"])  # b 在 none 级不产出；c 是 hex 级第 2 次


class LevelSemanticsCase(unittest.TestCase):
    """级别控制属性与既有语义回归。"""

    def test_lvl_restart_zero_keeps_running(self):
        """lvlRestart=0（ECMA：0 = 本级永不重置）：更高级出现后深层计数继续。"""
        lvls = (lvl(0, "decimal", "%1.")
                + lvl(1, "decimal", "%1.%2", restart=0))
        num = numbering_xml("1", lvls, [(1, "1", "")])
        doc = document_xml(para("a", 1, 0) + para("b", 1, 1)
                           + para("c", 1, 0) + para("d", 1, 1))
        self.assertEqual(numbers_of(build_docx(doc, num)),
                         ["1.", "1.1", "2.", "2.2"])

    def test_default_restart_resets(self):
        """无 lvlRestart：更高级出现时深层重置。"""
        lvls = lvl(0, "decimal", "%1.") + lvl(1, "decimal", "%1.%2")
        num = numbering_xml("1", lvls, [(1, "1", "")])
        doc = document_xml(para("a", 1, 0) + para("b", 1, 1)
                           + para("c", 1, 0) + para("d", 1, 1))
        self.assertEqual(numbers_of(build_docx(doc, num)),
                         ["1.", "1.1", "2.", "2.1"])

    def test_is_lgl_forces_decimal(self):
        """isLgl：引用级即使定义为 chineseCounting 也按 decimal 渲染。"""
        lvls = (lvl(0, "chineseCounting", "%1、")
                + lvl(1, "decimal", "%1.%2", is_lgl=True))
        num = numbering_xml("1", lvls, [(1, "1", "")])
        doc = document_xml(para("a", 1, 0) + para("b", 1, 1) + para("c", 1, 1))
        self.assertEqual(numbers_of(build_docx(doc, num)),
                         ["一、", "1.1", "1.2"])

    def test_lvl_override_full_redefinition(self):
        """lvlOverride 内完整 lvl 重定义优先于 abstractNum 定义。"""
        lvls = lvl(0, "decimal", "%1.")
        num = numbering_xml(
            "1", lvls,
            [(1, "1", '<w:lvlOverride w:ilvl="0">'
                      '<w:lvl w:ilvl="0"><w:start w:val="5"/>'
                      '<w:numFmt w:val="decimal"/><w:lvlText w:val="[%1]"/>'
                      '<w:lvlJc w:val="left"/></w:lvl></w:lvlOverride>')])
        doc = document_xml(para("a", 1, 0) + para("b", 1, 0))
        self.assertEqual(numbers_of(build_docx(doc, num)), ["[5]", "[6]"])

    def test_per_numid_isolation_and_start_override(self):
        """test-case-1 语义迷你回归：numId 隔离 + startOverride 钉值跳递增。"""
        lvls = "".join(lvl(i, "decimal", "%" + "123456789"[0:i + 1] +
                           ("" if i == 0 else "." + "%" * 0))
                       for i in range(9))
        # 直接写全 9 级 lvlText：%1、%1.%2、…、%1.%2.%3.%4.%5.%6.%7.%8.%9
        lvls = "".join(
            lvl(i, "decimal", ".".join(f"%{k + 1}" for k in range(i + 1)))
            for i in range(9))
        ovr = "".join(f'<w:lvlOverride w:ilvl="{i}">'
                      f'<w:startOverride w:val="1"/></w:lvlOverride>'
                      for i in range(9))
        num = numbering_xml("27", lvls, [(1, "27", ""), (2, "27", ovr)])
        doc = document_xml(
            para("h1", 1, 1)      # 1.1（%1 未启用 -> start）
            + para("h2", 1, 3)    # 1.1.1.1（%2/%3 未启用 -> start）
            + para("h3", 1, 4)    # 1.1.1.1.1
            + para("h4", 1, 4)    # 1.1.1.1.2
            + para("h5", 1, 6)    # 1.1.1.1.2.1.1（%5/%6 未启用 -> 渲染 start，不启用计数器）
            + para("h6", 1, 4)    # 1.1.1.1.3（清零 5-8 后回到 ilvl4）
            + para("h7", 1, 6)    # 1.1.1.1.3.1.1（ilvl5 从未本级启用 -> 渲染 start 而非清零 0）
            + para("s1", 2, 2)    # numId2 独立计数 + 全级 override：1.1.1
            + para("s2", 2, 2))   # 1.1.2
        self.assertEqual(numbers_of(build_docx(doc, num)),
                         ["1.1", "1.1.1.1", "1.1.1.1.1", "1.1.1.1.2",
                          "1.1.1.1.2.1.1", "1.1.1.1.3", "1.1.1.1.3.1.1",
                          "1.1.1", "1.1.2"])


class DocumentStructureCase(unittest.TestCase):
    """文档结构边界：去重 / 修订 / numId=0 / 表格 / 加粗。"""

    AC = (f'<mc:AlternateContent>'
          f'<mc:Choice Requires="wps">'
          f'<w:txbxContent>{para("tb1", 1, 0)}</w:txbxContent></mc:Choice>'
          f'<mc:Fallback>'
          f'<w:txbxContent>{para("tb1", 1, 0)}</w:txbxContent></mc:Fallback>'
          f'</mc:AlternateContent>')

    def test_alternate_content_dedup(self):
        """Choice/Fallback 双份内容只计一次。"""
        num = numbering_xml("1", lvl(0, "decimal", "%1."), [(1, "1", "")])
        body = (self.AC.replace("tb1", "tb1")
                + para("a", 1, 0) + self.AC.replace("tb1", "tb2"))
        got = DocxNumbering(build_docx(document_xml(body), num)).numbered_paragraphs()
        self.assertEqual([p["number"] for p in got], ["1.", "2.", "3."])
        self.assertEqual([p["text"] for p in got], ["tb1", "a", "tb2"])

    def test_deleted_paragraph_skipped(self):
        num = numbering_xml("1", lvl(0, "decimal", "%1."), [(1, "1", "")])
        body = (para("a", 1, 0) + para("gone", 1, 0, deleted=True)
                + para("b", 1, 0))
        self.assertEqual(numbers_of(build_docx(document_xml(body), num)),
                         ["1.", "2."])

    def test_numid_zero_no_number(self):
        num = numbering_xml("1", lvl(0, "decimal", "%1."), [(1, "1", "")])
        body = para("a", 0, 0) + para("b", 1, 0)
        self.assertEqual(numbers_of(build_docx(document_xml(body), num)), ["1."])

    def test_table_marker_and_bold(self):
        num = numbering_xml("1", lvl(0, "decimal", "%1."), [(1, "1", "")])
        body = ('<w:tbl><w:tr><w:tc>' + para("in-tbl", 1, 0, bold=True)
                + '</w:tc></w:tr></w:tbl>' + para("top", 1, 0))
        got = DocxNumbering(build_docx(document_xml(body), num)).numbered_paragraphs()
        self.assertEqual([p["in_table"] for p in got], [True, False])
        self.assertTrue(got[0]["bold"])
        self.assertFalse(got[1]["bold"])

    def test_normalize_text(self):
        self.assertEqual(normalize_text("a b\tc\nd"), "abcd")


if __name__ == "__main__":
    unittest.main(verbosity=2)
