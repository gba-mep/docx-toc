#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_toc.py — 为 .docx 插入真正的 Word 目录（TOC）域。

零依赖：只用 Python 标准库（zipfile / re / argparse），不依赖任何编辑器插件或 MCP。

用法:
    python docx_toc.py build  in.docx -o out.docx [--title "目　录"] [--page-break before]
    python docx_toc.py verify out.docx
    python docx_toc.py selftest

原理（详见 references/OOXML-TOC.md）:
    Word 的目录是一个「域」（field）:
        begin → instrText(TOC \\o "1-3" \\h \\z \\u) → separate → 结果 → end
    本脚本直接构造域与结果条目，并给每个标题打上 _Toc 书签，
    使 HYPERLINK / PAGEREF 能正确指向。
    条目内页码为占位值；在 Word 中按 F9（或 Ctrl+A → F9 → 只更新页码）后即为真实页码。
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape
from xml.sax.saxutils import unescape as xml_unescape

# ── 常数 ────────────────────────────────────────────────────────────────────
FONT = "微软雅黑"
SZ_BODY = "21"      # 10.5pt = 21 half-points
SZ_TITLE = "32"     # 16pt
TAB_POS = "9072"    # 页码右对齐位置（twips）
TWIPS_PER_CHAR = 240

P_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.S)
PSTYLE_RE = re.compile(r'<w:pStyle\b[^>]*w:val="([^"]+)"')
OUTLINE_RE = re.compile(r'<w:outlineLvl\b[^>]*w:val="(\d+)"')
WT_RE = re.compile(r"<w:t\b[^>]*>(.*?)</w:t>", re.S)

PAGE_BREAK_P = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


# ── XML 辅助 ────────────────────────────────────────────────────────────────
def _rfonts() -> str:
    return (f'<w:rFonts w:hint="eastAsia" w:ascii="{FONT}" '
            f'w:hAnsi="{FONT}" w:eastAsia="{FONT}"/>')


def _rpr(level: int, size: str = SZ_BODY) -> str:
    """rPr 子元素必须有序: rStyle → rFonts → b → bCs → … → sz → szCs。"""
    bold = "<w:b/><w:bCs/>" if level == 1 else ""
    return f"<w:rPr>{_rfonts()}{bold}<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/></w:rPr>"


def _style_level(style: str) -> int | None:
    s = re.sub(r"[\s_\-]", "", style).lower()
    for n in range(1, 10):
        if s in (f"heading{n}", f"h{n}", str(n)):
            return n
    return None


def para_level(block: str) -> int | None:
    """回传段落的标题层级（1-9），非标题回传 None。"""
    m = PSTYLE_RE.search(block)
    if m:
        lvl = _style_level(m.group(1))
        if lvl:
            return lvl
    m = OUTLINE_RE.search(block)
    if m:
        return int(m.group(1)) + 1
    return None


def para_text(block: str) -> str:
    return xml_unescape("".join(WT_RE.findall(block))).strip()


def iter_paragraphs(xml: str):
    """产生 (start, end, block)；段落不会嵌套，可用非贪婪匹配。"""
    for m in P_RE.finditer(xml):
        yield m.start(), m.end(), m.group(0)


# ── 书签与目录区块 ──────────────────────────────────────────────────────────
def add_bookmark(block: str, bid: int, name: str) -> str:
    """在段落内容最前（pPr 之后）插 bookmarkStart，末尾插 bookmarkEnd。"""
    m = re.match(r"(<w:p\b[^>]*>)(.*)(</w:p>)\s*$", block, re.S)
    if not m:
        return block
    head, body, tail = m.group(1), m.group(2), m.group(3)
    start = f'<w:bookmarkStart w:id="{bid}" w:name="{name}"/>'
    end = f'<w:bookmarkEnd w:id="{bid}"/>'
    mppr = re.match(r"(\s*<w:pPr\b.*?</w:pPr>)(.*)", body, re.S)
    if mppr:
        body = mppr.group(1) + start + mppr.group(2)
    else:
        body = start + body
    return head + body + end + tail


def build_toc(title: str, entries: list[tuple[int, str, str]], max_level: int) -> str:
    """entries: [(level, text, bookmark_name), …]"""
    out: list[str] = []

    if title:
        out.append(
            '<w:p><w:pPr><w:spacing w:before="0" w:after="240"/>'
            '<w:jc w:val="center"/>'
            f"{_rpr(1, SZ_TITLE)}</w:pPr>"
            f'<w:r>{_rpr(1, SZ_TITLE)}<w:t xml:space="preserve">{xml_escape(title)}</w:t></w:r>'
            "</w:p>"
        )

    for i, (lvl, text, bm) in enumerate(entries):
        ind = 100 * (lvl - 1)
        left = TWIPS_PER_CHAR * (lvl - 1)
        p = [
            "<w:p><w:pPr>",
            f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{TAB_POS}"/></w:tabs>',
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>',
            f'<w:ind w:leftChars="{ind}" w:left="{left}"/>',
            _rpr(lvl),
            "</w:pPr>",
        ]
        if i == 0:  # 域开始 + 指令，放在第一个条目段落内（与 Word 一致）
            p.append('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
            p.append(
                '<w:r><w:instrText xml:space="preserve">'
                f' TOC \\o "1-{max_level}" \\h \\z \\u '
                "</w:instrText></w:r>"
            )
            p.append('<w:r><w:fldChar w:fldCharType="separate"/></w:r>')

        p.append(f'<w:hyperlink w:anchor="{bm}">')
        p.append(f'<w:r>{_rpr(lvl)}<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>')
        p.append("<w:r><w:tab/></w:r>")
        p.append('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
        p.append(
            '<w:r><w:instrText xml:space="preserve">'
            f" PAGEREF {bm} \\h "
            "</w:instrText></w:r>"
        )
        p.append('<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
        p.append("<w:r><w:t>1</w:t></w:r>")          # 占位页码，F9 后更新
        p.append('<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        p.append("</w:hyperlink>")
        p.append("</w:p>")
        out.append("".join(p))

    out.append('<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    return "".join(out)


def insert_toc(xml: str, levels=(1, 2, 3), title: str = "目　录",
               page_break: str = "both") -> tuple[str, int]:
    """在「第一个标题」之前插入目录，并为所有标题补书签。回传 (新 XML, 标题数)。"""
    headings = []
    for s, e, block in iter_paragraphs(xml):
        lvl = para_level(block)
        if lvl in levels:
            headings.append((s, e, block, lvl, para_text(block)))

    if not headings:
        raise ValueError("找不到任何标题段落（需要 Heading 1–%d 样式或 outlineLvl）"
                         % max(levels))

    entries = [(lvl, text, f"_Toc{90000001 + i}") for i, (_, _, _, lvl, text) in enumerate(headings)]

    pieces: list[str] = []
    cursor = 0
    inserted = False
    for i, (s, e, block, lvl, _text) in enumerate(headings):
        pieces.append(xml[cursor:s])
        if not inserted:
            if page_break in ("before", "both"):
                pieces.append(PAGE_BREAK_P)
            pieces.append(build_toc(title, entries, max(levels)))
            if page_break in ("after", "both"):
                pieces.append(PAGE_BREAK_P)
            inserted = True
        pieces.append(add_bookmark(block, 90000001 + i, entries[i][2]))
        cursor = e
    pieces.append(xml[cursor:])

    return "".join(pieces), len(headings)


# ── docx 读写 ───────────────────────────────────────────────────────────────
def read_document_xml(docx: Path) -> str:
    with zipfile.ZipFile(docx) as z:
        return z.read("word/document.xml").decode("utf-8")


def write_docx(src: Path, dst: Path, new_document_xml: str) -> None:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_document_xml.encode("utf-8")
            zout.writestr(item, data)


# ── 验证 ────────────────────────────────────────────────────────────────────
def verify(docx: Path) -> tuple[bool, list[str]]:
    xml = read_document_xml(docx)
    n_entries = xml.count('<w:hyperlink w:anchor="_Toc')
    n_pageref = xml.count("PAGEREF")
    n_bookmarks = xml.count("<w:bookmarkStart")
    n_ind2 = xml.count('w:leftChars="100"')
    n_ind3 = xml.count('w:leftChars="200"')
    has_field = 'TOC \\o "1-' in xml
    has_tab = 'w:leader="dot"' in xml

    lines = [
        f"TOC 域            : {'OK' if has_field else 'MISSING'}",
        f"条目 (hyperlink)  : {n_entries}",
        f"PAGEREF           : {n_pageref}",
        f"书签 bookmarkStart: {n_bookmarks}",
        f"二级缩进 (100)     : {n_ind2}",
        f"三级缩进 (200)     : {n_ind3}",
        f"点前导线 tab       : {'OK' if has_tab else 'MISSING'}",
    ]
    ok = (has_field and has_tab and n_entries > 0
          and n_pageref == n_entries and n_bookmarks == n_entries)
    return ok, lines


# ── 自测：从零造一份 docx，跑一遍并验证 ─────────────────────────────────────
_CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def _styles() -> str:
    def h(n: int) -> str:
        return (f'<w:style w:type="paragraph" w:styleId="Heading{n}">'
                f'<w:name w:val="heading {n}"/><w:basedOn w:val="Normal"/>'
                f'<w:pPr><w:outlineLvl w:val="{n - 1}"/></w:pPr>'
                f'<w:rPr><w:b/></w:rPr></w:style>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/></w:style>'
            + "".join(h(n) for n in (1, 2, 3)) + "</w:styles>")


def _selftest_docx(path: Path) -> int:
    """造一份含 3×H1 / 4×H2 / 2×H3 的文件，回传标题总数。"""
    spec = [("封面标题", 0)] + [
        ("第一章 总则", 1), ("1.1 适用范围", 2), ("1.2 引用标准", 2), ("1.2.1 澳门法规", 3),
        ("第二章 材料", 1), ("2.1 品牌库", 2), ("2.1.1 货期", 3), ("2.2 报批", 2),
        ("第三章 收尾", 1),
    ]
    body = []
    for text, lvl in spec:
        if lvl == 0:
            body.append(f'<w:p><w:r><w:t>{xml_escape(text)}</w:t></w:r></w:p>')
        else:
            body.append(f'<w:p><w:pPr><w:pStyle w:val="Heading{lvl}"/></w:pPr>'
                        f'<w:r><w:t>{xml_escape(text)}</w:t></w:r></w:p>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           "<w:body>" + "".join(body)
           + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT)
        z.writestr("_rels/.rels", _RELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        z.writestr("word/styles.xml", _styles())
    return sum(1 for _, lvl in spec if lvl)


def cmd_selftest(args) -> int:
    out_dir = Path(args.outdir or ".").resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    src = out_dir / "_selftest_in.docx"
    dst = out_dir / "_selftest_out.docx"

    n_headings = _selftest_docx(src)
    xml = read_document_xml(src)
    new_xml, n_found = insert_toc(xml, title="目　录")
    write_docx(src, dst, new_xml)

    ok, lines = verify(dst)
    print("=== selftest ===")
    print(f"来源标题数        : {n_headings}")
    print(f"脚本找到标题数    : {n_found}")
    for line in lines:
        print(line)
    consistent = ok and n_found == n_headings
    print("RESULT            :", "PASS" if consistent else "FAIL")
    print(f"输出              : {dst}")
    if args.keep:
        print(f"（来源保留：{src}）")
    else:
        src.unlink(missing_ok=True)
    return 0 if consistent else 1


def cmd_build(args) -> int:
    src = Path(args.input).resolve()
    dst = Path(args.output).resolve() if args.output else src.with_name(src.stem + "_TOC.docx")
    xml = read_document_xml(src)
    new_xml, n = insert_toc(xml, title=args.title, page_break=args.page_break)
    write_docx(src, dst, new_xml)
    ok, lines = verify(dst)
    for line in lines:
        print(line)
    print(f"标题数: {n}")
    print(f"输出  : {dst}")
    print("提醒  : 在 Word 中按 Ctrl+A → F9 →「只更新页码」以填入真实页码。")
    return 0 if ok else 1


def cmd_verify(args) -> int:
    ok, lines = verify(Path(args.input).resolve())
    for line in lines:
        print(line)
    print("RESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="为 .docx 插入真正的 Word 目录（TOC）域")
    sub = ap.add_subparsers(dest="cmd", required=True)

    b = sub.add_parser("build", help="插入目录")
    b.add_argument("input")
    b.add_argument("-o", "--output", default="")
    b.add_argument("--title", default="目　录")
    b.add_argument("--page-break", choices=["none", "before", "after", "both"], default="both")
    b.set_defaults(func=cmd_build)

    v = sub.add_parser("verify", help="验证目录结构")
    v.add_argument("input")
    v.set_defaults(func=cmd_verify)

    s = sub.add_parser("selftest", help="自测：造样本 → 插入 → 验证")
    s.add_argument("--outdir", default="")
    s.add_argument("--keep", action="store_true", help="保留样本档")
    s.set_defaults(func=cmd_selftest)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
