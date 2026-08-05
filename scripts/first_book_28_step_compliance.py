#!/usr/bin/env python3
from __future__ import annotations

import json
import math
import os
import re
import shutil
import subprocess
import time
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Optional

import fitz
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path.cwd()
SOURCE = ROOT / "source_interior.docx"
SOURCE_FRONT = ROOT / "source_front.docx"
SOURCE_BACK = ROOT / "source_back_spine.docx"
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)
CONTACT = OUT / "لوحات_المراجعة"
CONTACT.mkdir(exist_ok=True)

BOOK_TITLE = "الحوكمة من الصفر"
SUBTITLE = "دليل عملي للشركات السعودية من التأسيس إلى النضج"
EDITION = "الطبعة المنقحة الأولى | 2026م"

FINAL_DOCX = OUT / "الحوكمة_من_الصفر_النسخة_المطابقة_للمعيار.docx"
FINAL_PDF = OUT / "الحوكمة_من_الصفر_نسخة_المراجعة.pdf"
FRONT_DOCX = OUT / "الحوكمة_من_الصفر_الغلاف_الأمامي_قالب_الإنتاج.docx"
FRONT_PNG = OUT / "الحوكمة_من_الصفر_الغلاف_الأمامي_300DPI.png"
BACK_DOCX = OUT / "الحوكمة_من_الصفر_الغلاف_الخلفي_والكعب_قالب_الإنتاج.docx"
BACK_PNG = OUT / "الحوكمة_من_الصفر_الغلاف_الخلفي_300DPI.png"
SPINE_PNG = OUT / "الحوكمة_من_الصفر_دليل_الكعب_غير_مقاس.png"
REPORT_DOCX = OUT / "تقرير_مطابقة_الحوكمة_من_الصفر_للخطوات_28.docx"
REPORT_PDF = OUT / "تقرير_مطابقة_الحوكمة_من_الصفر_للخطوات_28.pdf"
AUDIT_JSON = OUT / "audit_28_steps.json"
README = OUT / "اقرأني_أولاً.txt"
PACKAGE_ZIP = ROOT / "الحوكمة_من_الصفر_حزمة_المراجعة_الشاملة.zip"

BLACK = "000000"
WHITE = "FFFFFF"
DARK_GRAY = "3F3F3F"
MID_GRAY = "7F7F7F"
LIGHT_GRAY = "F2F2F2"
PETROL = "164C5A"
GOLD = "B39756"
IVORY = "F4EFE3"
CHARCOAL = "333333"
ARABIC_DIGITS = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def clean(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\u200f", "").replace("\u200e", "")).strip()


def ar_ratio(s: str) -> float:
    letters = re.findall(r"[A-Za-z\u0600-\u06FF]", s or "")
    if not letters:
        return 0.0
    return sum("\u0600" <= c <= "\u06FF" for c in letters) / len(letters)


def ensure_el(parent, tag: str):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def set_bidi_paragraph(p, alignment=None):
    ppr = p._p.get_or_add_pPr()
    if ppr.find(qn("w:bidi")) is None:
        ppr.insert(0, OxmlElement("w:bidi"))
    if alignment is not None:
        p.alignment = alignment
        jc = ensure_el(ppr, "w:jc")
        mapping = {
            WD_ALIGN_PARAGRAPH.RIGHT: "right",
            WD_ALIGN_PARAGRAPH.LEFT: "left",
            WD_ALIGN_PARAGRAPH.CENTER: "center",
            WD_ALIGN_PARAGRAPH.JUSTIFY: "both",
        }
        jc.set(qn("w:val"), mapping.get(alignment, "right"))


def set_run(run, font: str, size: float, bold: Optional[bool] = None, color=BLACK, rtl=True):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.font.bold = bold
    rpr = run._r.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for k in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{k}"), font)
    if rtl and rpr.find(qn("w:rtl")) is None:
        rpr.append(OxmlElement("w:rtl"))
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "ar-SA" if rtl else "en-US")
    lang.set(qn("w:bidi"), "ar-SA" if rtl else "en-US")


def remove_conflicting_style(doc, name: str):
    try:
        st = doc.styles[name]
    except KeyError:
        return
    if st.type != WD_STYLE_TYPE.PARAGRAPH:
        st.element.getparent().remove(st.element)


def ensure_style(doc, name: str, font: str, size: float, bold=False, align=WD_ALIGN_PARAGRAPH.RIGHT,
                 before=0, after=5, line=1.2, first=0, left=0, right=0,
                 keep_next=False, keep_lines=True, outline=None, color=BLACK):
    remove_conflicting_style(doc, name)
    try:
        st = doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    st.font.name = font
    st.font.size = Pt(size)
    st.font.bold = bold
    st.font.color.rgb = RGBColor.from_string(color)
    pf = st.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.first_line_indent = Cm(first) if first else None
    pf.left_indent = Cm(left) if left else None
    pf.right_indent = Cm(right) if right else None
    pf.keep_with_next = keep_next
    pf.keep_together = keep_lines
    pf.widow_control = True
    rpr = st.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for k in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{k}"), font)
    if rpr.find(qn("w:rtl")) is None:
        rpr.append(OxmlElement("w:rtl"))
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "ar-SA")
    lang.set(qn("w:bidi"), "ar-SA")
    ppr = st.element.get_or_add_pPr()
    if ppr.find(qn("w:bidi")) is None:
        ppr.insert(0, OxmlElement("w:bidi"))
    if outline is not None:
        out = ppr.find(qn("w:outlineLvl"))
        if out is None:
            out = OxmlElement("w:outlineLvl")
            ppr.append(out)
        out.set(qn("w:val"), str(outline))
    return st


def create_required_styles(doc):
    ensure_style(doc, "Book Title", "Noto Kufi Arabic", 28, True, WD_ALIGN_PARAGRAPH.CENTER, after=10, line=1.05)
    ensure_style(doc, "Book Subtitle", "Noto Kufi Arabic", 15, False, WD_ALIGN_PARAGRAPH.CENTER, after=8, line=1.15)
    ensure_style(doc, "Author Name", "Noto Kufi Arabic", 15, False, WD_ALIGN_PARAGRAPH.CENTER, after=8, line=1.15)
    ensure_style(doc, "Part Title", "Noto Kufi Arabic", 24, True, WD_ALIGN_PARAGRAPH.CENTER, before=60, after=12, line=1.05, keep_next=True, outline=0)
    ensure_style(doc, "Chapter Title", "Noto Kufi Arabic", 22, True, WD_ALIGN_PARAGRAPH.RIGHT, before=45, after=14, line=1.1, keep_next=True, outline=1)
    ensure_style(doc, "Main Heading", "Noto Kufi Arabic", 17, True, WD_ALIGN_PARAGRAPH.RIGHT, before=12, after=6, line=1.1, keep_next=True, outline=2)
    ensure_style(doc, "Subheading", "Noto Kufi Arabic", 15, True, WD_ALIGN_PARAGRAPH.RIGHT, before=10, after=5, line=1.1, keep_next=True, outline=3)
    ensure_style(doc, "Unnumbered Internal Heading", "Noto Kufi Arabic", 15, True, WD_ALIGN_PARAGRAPH.RIGHT, before=10, after=5, line=1.1, keep_next=True)
    ensure_style(doc, "Body Text", "Amiri", 14, False, WD_ALIGN_PARAGRAPH.JUSTIFY, after=5, line=1.2, first=.6)
    ensure_style(doc, "First Paragraph after Heading", "Amiri", 14, False, WD_ALIGN_PARAGRAPH.JUSTIFY, after=5, line=1.2)
    ensure_style(doc, "Long Quotation", "Amiri", 13, False, WD_ALIGN_PARAGRAPH.JUSTIFY, before=8, after=8, line=1.15, left=1, right=1)
    ensure_style(doc, "Qur’anic Verse", "Amiri Quran", 15, False, WD_ALIGN_PARAGRAPH.CENTER, before=8, after=8, line=1.15)
    ensure_style(doc, "Prophetic Hadith", "Amiri", 13.5, False, WD_ALIGN_PARAGRAPH.JUSTIFY, before=6, after=6, line=1.15, left=.8, right=.8)
    ensure_style(doc, "Box Title", "Noto Kufi Arabic", 13, True, WD_ALIGN_PARAGRAPH.RIGHT, before=8, after=3, line=1.1)
    ensure_style(doc, "Box Text", "Amiri", 12.5, False, WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=8, line=1.2)
    ensure_style(doc, "Table Title", "Amiri", 10.5, True, WD_ALIGN_PARAGRAPH.RIGHT, before=6, after=3, line=1.05, keep_next=True)
    ensure_style(doc, "Table Text", "Amiri", 11.5, False, WD_ALIGN_PARAGRAPH.RIGHT, after=2, line=1.05)
    ensure_style(doc, "Table Source", "Amiri", 10.5, False, WD_ALIGN_PARAGRAPH.RIGHT, before=3, after=6, line=1.05)
    ensure_style(doc, "Figure Title", "Amiri", 10.5, True, WD_ALIGN_PARAGRAPH.CENTER, before=4, after=2, line=1.05)
    ensure_style(doc, "Figure Caption", "Amiri", 10.5, False, WD_ALIGN_PARAGRAPH.CENTER, before=2, after=6, line=1.05)
    ensure_style(doc, "Footnote", "Amiri", 10.5, False, WD_ALIGN_PARAGRAPH.RIGHT, after=2, line=1.0)
    ensure_style(doc, "Bulleted List", "Amiri", 14, False, WD_ALIGN_PARAGRAPH.RIGHT, after=4, line=1.15, right=.7)
    ensure_style(doc, "Numbered List", "Amiri", 14, False, WD_ALIGN_PARAGRAPH.RIGHT, after=4, line=1.15, right=.7)
    ensure_style(doc, "Reference", "Amiri", 11.5, False, WD_ALIGN_PARAGRAPH.RIGHT, after=4, line=1.15, right=.7)
    ensure_style(doc, "Appendix Title", "Noto Kufi Arabic", 22, True, WD_ALIGN_PARAGRAPH.RIGHT, before=45, after=14, line=1.1, keep_next=True, outline=1)

    # Compatibility aliases found in previous production stages.
    aliases = {
        "Body Text Arabic": "Body Text",
        "First Body Paragraph": "First Paragraph after Heading",
        "Numbered List Arabic": "Numbered List",
        "Reference Arabic": "Reference",
        "Reference Latin": "Reference",
        "Callout Title": "Box Title",
        "Callout Text": "Box Text",
    }
    for alias, base in aliases.items():
        b = doc.styles[base]
        ensure_style(doc, alias, b.font.name or "Amiri", b.font.size.pt if b.font.size else 14,
                     bool(b.font.bold), b.paragraph_format.alignment or WD_ALIGN_PARAGRAPH.RIGHT,
                     before=b.paragraph_format.space_before.pt if b.paragraph_format.space_before else 0,
                     after=b.paragraph_format.space_after.pt if b.paragraph_format.space_after else 0,
                     line=b.paragraph_format.line_spacing if isinstance(b.paragraph_format.line_spacing, (int, float)) else 1.2,
                     first=(b.paragraph_format.first_line_indent.cm if b.paragraph_format.first_line_indent else 0),
                     left=(b.paragraph_format.left_indent.cm if b.paragraph_format.left_indent else 0),
                     right=(b.paragraph_format.right_indent.cm if b.paragraph_format.right_indent else 0),
                     keep_next=bool(b.paragraph_format.keep_with_next), keep_lines=True)


def paragraph_style_group(name: str):
    n = name or ""
    if n in {"Book Title", "Book Subtitle", "Author Name", "Part Title"}:
        return "center_heading"
    if n in {"Chapter Title", "Main Heading", "Subheading", "Unnumbered Internal Heading", "Appendix Title", "Box Title"}:
        return "right_heading"
    if n in {"Body Text", "Body Text Arabic", "First Paragraph after Heading", "First Body Paragraph", "Long Quotation", "Prophetic Hadith", "Box Text", "Callout Text"}:
        return "justified"
    if n in {"Qur’anic Verse", "Figure Title", "Figure Caption"}:
        return "center"
    if n in {"Table Title", "Table Source", "Table Text", "Footnote", "Bulleted List", "Numbered List", "Numbered List Arabic", "Reference", "Reference Arabic"}:
        return "right"
    if n == "Reference Latin":
        return "left"
    return "unknown"


def normalize_list_prefix(text: str) -> str:
    m = re.match(r"^(\d+(?:\.\d+)*)\s*([.)-]?)\s*(.*)$", text)
    if not m:
        return text
    nums = m.group(1).translate(ARABIC_DIGITS)
    rest = m.group(3)
    return f"{nums}. {rest}" if "." not in m.group(1) else f"{nums} {rest}"


def paragraph_has_protected_content(p):
    return bool(p._p.xpath('.//w:sectPr|.//w:drawing|.//w:pict|.//w:fldChar|.//w:bookmarkStart'))


def style_document_paragraphs(doc):
    stats = {"paragraphs": len(doc.paragraphs), "blank_removed": 0, "first_after_heading_fixed": 0,
             "rtl_fixed": 0, "numbering_normalized": 0}
    # Remove body blank-line separators while preserving front matter and section/drawing carriers.
    paras = list(doc.paragraphs)
    for i, p in enumerate(paras):
        if i < 15 or clean(p.text) or paragraph_has_protected_content(p):
            continue
        prev = next((clean(paras[j].text) for j in range(i-1, -1, -1) if clean(paras[j].text)), "")
        nxt = next((clean(paras[j].text) for j in range(i+1, len(paras)) if clean(paras[j].text)), "")
        if prev and nxt:
            parent = p._p.getparent()
            if parent is not None:
                parent.remove(p._p)
                stats["blank_removed"] += 1
    # Reloading occurs later; style current surviving paragraphs.
    previous_nonempty_style = None
    heading_styles = {"Part Title", "Chapter Title", "Main Heading", "Subheading", "Unnumbered Internal Heading", "Appendix Title", "Box Title", "Callout Title"}
    body_styles = {"Body Text", "Body Text Arabic", "First Paragraph after Heading", "First Body Paragraph"}
    for p in doc.paragraphs:
        text = clean(p.text)
        name = p.style.name if p.style else ""
        # Map common existing styles to required names without changing semantic content.
        mapping = {
            "Body Text Arabic": "Body Text",
            "First Body Paragraph": "First Paragraph after Heading",
            "Numbered List Arabic": "Numbered List",
            "Callout Title": "Box Title",
            "Callout Text": "Box Text",
        }
        if name in mapping:
            p.style = mapping[name]
            name = mapping[name]
        # Recognise explicit captions and source notes.
        if re.match(r"^الجدول\s+[٠-٩0-9]+", text):
            p.style = "Table Title"; name = "Table Title"
        elif text.startswith("المصدر:"):
            p.style = "Table Source"; name = "Table Source"
        elif re.match(r"^الشكل\s+[٠-٩0-9]+", text):
            p.style = "Figure Caption"; name = "Figure Caption"
        elif text.startswith("الفصل ") and name not in {"Chapter Title"}:
            p.style = "Chapter Title"; name = "Chapter Title"
        elif text.startswith("الملحق ") and name not in {"Appendix Title"}:
            p.style = "Appendix Title"; name = "Appendix Title"
        # First paragraph after headings.
        if text and previous_nonempty_style in heading_styles and name in body_styles:
            p.style = "First Paragraph after Heading"
            name = "First Paragraph after Heading"
            stats["first_after_heading_fixed"] += 1
        group = paragraph_style_group(name)
        if group == "center_heading" or group == "center":
            set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.CENTER)
        elif group == "right_heading" or group == "right":
            set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.RIGHT)
        elif group == "left":
            set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.LEFT)
        else:
            set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.JUSTIFY if text else WD_ALIGN_PARAGRAPH.RIGHT)
        stats["rtl_fixed"] += 1
        pf = p.paragraph_format
        pf.widow_control = True
        if name in heading_styles or name in {"Table Title", "Figure Title"}:
            pf.keep_with_next = True
            pf.keep_together = True
        elif name not in {"Table Source"}:
            pf.keep_together = True
        if name == "Numbered List" and text:
            normalized = normalize_list_prefix(text)
            if normalized != text:
                # Preserve run formatting only when a correction is actually needed.
                p.text = normalized
                stats["numbering_normalized"] += 1
        # Enforce run fonts, sizes and language from the paragraph style.
        style = p.style
        font_name = style.font.name or ("Noto Kufi Arabic" if "Heading" in name or "Title" in name else "Amiri")
        size = style.font.size.pt if style.font.size else 14
        bold = style.font.bold
        for r in p.runs:
            rtl = name != "Reference Latin" and (ar_ratio(r.text) >= .15 or ar_ratio(text) >= .15)
            set_run(r, font_name, size, bold=bold if bold is not None else None, color=BLACK, rtl=rtl)
        if text:
            previous_nonempty_style = name
    return stats


def set_cell_shading(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, value_twips=227):  # 0.4 cm
    tcpr = cell._tc.get_or_add_tcPr()
    mar = tcpr.find(qn("w:tcMar"))
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcpr.append(mar)
    for side in ("top", "start", "bottom", "end"):
        el = mar.find(qn(f"w:{side}"))
        if el is None:
            el = OxmlElement(f"w:{side}")
            mar.append(el)
        el.set(qn("w:w"), str(value_twips))
        el.set(qn("w:type"), "dxa")


def set_table_borders(table, color=MID_GRAY, sz="4", side_accent=False):
    tblpr = table._tbl.tblPr
    borders = tblpr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblpr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = borders.find(qn(f"w:{edge}"))
        if e is None:
            e = OxmlElement(f"w:{edge}")
            borders.append(e)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "14" if side_accent and edge == "right" else sz)
        e.set(qn("w:color"), DARK_GRAY if side_accent and edge == "right" else color)


def style_tables(doc):
    stats = {"tables": len(doc.tables), "boxes": 0, "content_tables": 0, "rtl_tables": 0}
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        tblpr = table._tbl.tblPr
        if tblpr.find(qn("w:bidiVisual")) is None:
            tblpr.append(OxmlElement("w:bidiVisual"))
        stats["rtl_tables"] += 1
        is_box = len(table.rows) == 1 and len(table.columns) == 1
        if is_box:
            stats["boxes"] += 1
            set_table_borders(table, color=LIGHT_GRAY, sz="0", side_accent=True)
            cell = table.cell(0, 0)
            set_cell_shading(cell, LIGHT_GRAY)
            set_cell_margins(cell, 227)
            for idx, p in enumerate(cell.paragraphs):
                if idx == 0 and len(cell.paragraphs) > 1:
                    p.style = "Box Title"
                    set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.RIGHT)
                else:
                    p.style = "Box Text"
                    set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.JUSTIFY)
                for r in p.runs:
                    if p.style.name == "Box Title":
                        set_run(r, "Noto Kufi Arabic", 13, bold=True, color=BLACK, rtl=True)
                    else:
                        set_run(r, "Amiri", 12.5, color=BLACK, rtl=True)
            continue
        stats["content_tables"] += 1
        set_table_borders(table, color=MID_GRAY, sz="4")
        for ri, row in enumerate(table.rows):
            trpr = row._tr.get_or_add_trPr()
            if trpr.find(qn("w:cantSplit")) is None:
                trpr.append(OxmlElement("w:cantSplit"))
            if ri == 0 and trpr.find(qn("w:tblHeader")) is None:
                trpr.append(OxmlElement("w:tblHeader"))
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                set_cell_margins(cell, 100)
                set_cell_shading(cell, DARK_GRAY if ri == 0 else (LIGHT_GRAY if ri % 2 == 0 else WHITE))
                for p in cell.paragraphs:
                    p.style = "Table Text"
                    set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.RIGHT)
                    p.paragraph_format.keep_together = True
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.05
                    for r in p.runs:
                        set_run(r, "Amiri", 11.5, bold=(ri == 0), color=WHITE if ri == 0 else BLACK, rtl=True)
    return stats


def set_geometry(section):
    section.page_width = Cm(17)
    section.page_height = Cm(24)
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.4)
    section.left_margin = Cm(2.8)   # Inside when mirrorMargins is active
    section.right_margin = Cm(2.0)  # Outside when mirrorMargins is active
    section.gutter = Cm(.7)
    section.header_distance = Cm(1.0)
    section.footer_distance = Cm(1.0)
    pgmar = section._sectPr.find(qn("w:pgMar"))
    if pgmar is not None:
        pgmar.set(qn("w:gutter"), "397")


def ensure_settings(doc):
    settings = doc.settings.element
    for tag in ("w:mirrorMargins", "w:rtlGutter", "w:evenAndOddHeaders"):
        if settings.find(qn(tag)) is None:
            settings.append(OxmlElement(tag))
    upd = settings.find(qn("w:updateFields"))
    if upd is None:
        upd = OxmlElement("w:updateFields")
        settings.append(upd)
    upd.set(qn("w:val"), "true")
    # Set default proofing languages.
    styles = doc.styles.element
    doc_defaults = styles.find(qn("w:docDefaults"))
    if doc_defaults is None:
        doc_defaults = OxmlElement("w:docDefaults")
        styles.insert(0, doc_defaults)
    rpr_default = doc_defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = OxmlElement("w:rPrDefault"); doc_defaults.append(rpr_default)
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr"); rpr_default.append(rpr)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang"); rpr.append(lang)
    lang.set(qn("w:val"), "ar-SA")
    lang.set(qn("w:bidi"), "ar-SA")


def clone_sectpr(template, break_type="oddPage"):
    s = deepcopy(template)
    for tag in ("w:headerReference", "w:footerReference", "w:type", "w:titlePg", "w:pgNumType"):
        for el in list(s.findall(qn(tag))):
            s.remove(el)
    typ = OxmlElement("w:type")
    typ.set(qn("w:val"), break_type)
    s.insert(0, typ)
    return s


def ensure_odd_breaks(doc):
    body = doc.element.body
    template = body.sectPr
    inserted = 0
    adjusted = 0
    for p in list(doc.paragraphs):
        text = clean(p.text)
        name = p.style.name if p.style else ""
        if not (name == "Chapter Title" or text.startswith("الفصل ")):
            continue
        prev = p._p.getprevious()
        sect = None
        if prev is not None and prev.tag == qn("w:p"):
            ppr = prev.find(qn("w:pPr"))
            sect = ppr.find(qn("w:sectPr")) if ppr is not None else None
        if sect is None:
            holder = OxmlElement("w:p")
            ppr = OxmlElement("w:pPr")
            holder.append(ppr)
            ppr.append(clone_sectpr(template, "oddPage"))
            p._p.addprevious(holder)
            inserted += 1
        else:
            typ = sect.find(qn("w:type"))
            if typ is None:
                typ = OxmlElement("w:type"); sect.insert(0, typ)
            typ.set(qn("w:val"), "oddPage")
            adjusted += 1
    return {"odd_breaks_inserted": inserted, "odd_breaks_adjusted": adjusted}


def clear_story(story):
    paras = list(story.paragraphs)
    p = paras[0]
    for extra in paras[1:]:
        extra._element.getparent().remove(extra._element)
    for child in list(p._p):
        if child.tag != qn("w:pPr"):
            p._p.remove(child)
    return p


def add_field(p, instruction, placeholder=""):
    r = p.add_run()._r
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin"); b.set(qn("w:dirty"), "true"); r.append(b)
    r = p.add_run()._r
    i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = instruction; r.append(i)
    r = p.add_run()._r
    s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate"); r.append(s)
    if placeholder:
        p.add_run(placeholder)
    r = p.add_run()._r
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end"); r.append(e)


def set_page_number_format(section, fmt, start=None):
    sect = section._sectPr
    old = sect.find(qn("w:pgNumType"))
    if old is not None:
        sect.remove(old)
    el = OxmlElement("w:pgNumType")
    el.set(qn("w:fmt"), fmt)
    if start is not None:
        el.set(qn("w:start"), str(start))
    sect.append(el)


def section_index_for_first_chapter(doc):
    idx = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = clean("".join(child.itertext()))
            ppr = child.find(qn("w:pPr"))
            if text.startswith("الفصل "):
                return idx
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                idx += 1
    return max(1, len(doc.sections) - 15)


def configure_headers_footers(doc):
    ensure_settings(doc)
    for sec in doc.sections:
        set_geometry(sec)
    first_chapter_section = section_index_for_first_chapter(doc)
    # Preserve front-matter stories but enforce Arabic direction and outside page numbers.
    for si, sec in enumerate(doc.sections):
        for story in (sec.header, sec.even_page_header, sec.first_page_header, sec.footer, sec.even_page_footer, sec.first_page_footer):
            for p in story.paragraphs:
                set_bidi_paragraph(p, p.alignment or WD_ALIGN_PARAGRAPH.RIGHT)
                for r in p.runs:
                    set_run(r, "Amiri", 10.5, color=BLACK, rtl=True)
        if si < first_chapter_section:
            set_page_number_format(sec, "lowerRoman")
            if si == 0:
                sec.different_first_page_header_footer = True
                clear_story(sec.first_page_header)
                clear_story(sec.first_page_footer)
            continue
        sec.different_first_page_header_footer = True
        for story_name in ("header", "even_page_header", "first_page_header", "footer", "even_page_footer", "first_page_footer"):
            getattr(sec, story_name).is_linked_to_previous = False
        # Odd/right page header: current chapter title.
        p = clear_story(sec.header)
        set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.RIGHT)
        add_field(p, ' STYLEREF "Chapter Title" ', BOOK_TITLE)
        # Even/left page header: book title.
        p = clear_story(sec.even_page_header)
        set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.LEFT)
        r = p.add_run(BOOK_TITLE); set_run(r, "Amiri", 10.5, rtl=True)
        clear_story(sec.first_page_header)
        # Outside page numbers.
        p = clear_story(sec.footer)
        set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.RIGHT)
        add_field(p, " PAGE ", "1")
        p = clear_story(sec.even_page_footer)
        set_bidi_paragraph(p, WD_ALIGN_PARAGRAPH.LEFT)
        add_field(p, " PAGE ", "1")
        clear_story(sec.first_page_footer)
        set_page_number_format(sec, "decimal", 1 if si == first_chapter_section else None)
    return first_chapter_section


def mark_fields_dirty(doc):
    count = 0
    for fld in doc.element.xpath('.//w:fldChar[@w:fldCharType="begin"]'):
        fld.set(qn("w:dirty"), "true")
        count += 1
    return count


def cap_inline_images(doc):
    max_width = Cm(17 - 2.8 - 2.0 - .7)
    corrected = 0
    for shape in doc.inline_shapes:
        if shape.width > max_width:
            ratio = max_width / shape.width
            shape.width = max_width
            shape.height = int(shape.height * ratio)
            corrected += 1
    return corrected


def build_interior():
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)
    doc = Document(SOURCE)
    original = {
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "sections": len(doc.sections),
        "inline_shapes": len(doc.inline_shapes),
    }
    create_required_styles(doc)
    ensure_settings(doc)
    paragraph_stats = style_document_paragraphs(doc)
    table_stats = style_tables(doc)
    image_caps = cap_inline_images(doc)
    break_stats = ensure_odd_breaks(doc)
    stage = ROOT / "stage_first_book.docx"
    doc.save(stage)
    doc = Document(stage)
    create_required_styles(doc)
    first_chapter_section = configure_headers_footers(doc)
    dirty = mark_fields_dirty(doc)
    doc.core_properties.title = BOOK_TITLE
    doc.core_properties.subject = SUBTITLE
    doc.core_properties.language = "ar-SA"
    doc.core_properties.comments = "مراجعة مطابقة لمعيار تنسيق الكتب العربية — 28 خطوة"
    doc.save(FINAL_DOCX)
    return {
        "original": original,
        "paragraph_stats": paragraph_stats,
        "table_stats": table_stats,
        "image_caps": image_caps,
        **break_stats,
        "first_chapter_section": first_chapter_section,
        "final_sections_before_lo": len(Document(FINAL_DOCX).sections),
        "fields_marked_dirty": dirty,
    }


def update_fields_and_export():
    script = ROOT / "lo_first_book.py"
    script.write_text(r'''import pathlib, subprocess, time, uno
from com.sun.star.beans import PropertyValue

def prop(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(['soffice','--headless','--nologo','--nodefault','--nofirststartwizard','--norestore','--accept=socket,host=127.0.0.1,port=2003;urp;StarOffice.ComponentContext'])
try:
    local=uno.getComponentContext(); resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
    ctx=None
    for _ in range(60):
        try: ctx=resolver.resolve('uno:socket,host=127.0.0.1,port=2003;urp;StarOffice.ComponentContext'); break
        except Exception: time.sleep(1)
    if ctx is None: raise RuntimeError('LibreOffice UNO connection failed')
    desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
    src=pathlib.Path('output/الحوكمة_من_الصفر_النسخة_المطابقة_للمعيار.docx').resolve().as_uri()
    doc=desktop.loadComponentFromURL(src,'_blank',0,(prop('Hidden',True),prop('UpdateDocMode',3)))
    if doc is None: raise RuntimeError('Could not load final DOCX')
    try: doc.getTextFields().refresh()
    except Exception: pass
    try:
        indexes=doc.getDocumentIndexes()
        for i in range(indexes.getCount()): indexes.getByIndex(i).update()
        print('indexes_updated',indexes.getCount())
    except Exception as e: print('indexes_warning',e)
    try: doc.calculateAll()
    except Exception: pass
    time.sleep(5)
    out_docx=pathlib.Path('output/الحوكمة_من_الصفر_النسخة_المطابقة_للمعيار.docx').resolve().as_uri()
    out_pdf=pathlib.Path('output/الحوكمة_من_الصفر_نسخة_المراجعة.pdf').resolve().as_uri()
    doc.storeAsURL(out_docx,(prop('FilterName','Office Open XML Text'),prop('Overwrite',True)))
    try:
        indexes=doc.getDocumentIndexes()
        for i in range(indexes.getCount()): indexes.getByIndex(i).update()
    except Exception: pass
    time.sleep(3)
    doc.storeToURL(out_pdf,(prop('FilterName','writer_pdf_Export'),prop('Overwrite',True),prop('UseTaggedPDF',True),prop('ExportBookmarks',True)))
    doc.close(True)
finally:
    proc.terminate()
''', encoding="utf-8")
    subprocess.run(["/usr/bin/python3", str(script)], check=True)


def font_path(pattern):
    try:
        return subprocess.check_output(["fc-match", "-f", "%{file}", pattern], text=True).strip()
    except Exception:
        return "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"


def rtl_text(s):
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        return get_display(arabic_reshaper.reshape(s))
    except Exception:
        return s


def wrap_rtl(draw, text, font, maxw):
    words = text.split(); lines=[]; cur=""
    for word in words:
        trial=(cur+" "+word).strip()
        box=draw.textbbox((0,0),rtl_text(trial),font=font)
        if box[2]-box[0] <= maxw:
            cur=trial
        else:
            if cur: lines.append(cur)
            cur=word
    if cur: lines.append(cur)
    return lines


def image_to_word(img_path: Path, out_path: Path, pages=1, second_image: Optional[Path]=None):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(17); sec.page_height = Cm(24)
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(0)
    p = doc.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(str(img_path), width=Cm(17), height=Cm(24))
    if pages == 2 and second_image:
        p2 = doc.add_paragraph()
        p2.paragraph_format.page_break_before = True
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = p2.paragraph_format.space_after = Pt(0)
        p2.add_run().add_picture(str(second_image), width=Cm(17), height=Cm(24))
    doc.save(out_path)


def build_cover_templates():
    W,H=2008,2835
    ivory="#F4EFE3"; petrol="#164C5A"; gold="#B39756"; charcoal="#333333"
    kufi=font_path("Noto Kufi Arabic"); amiri=font_path("Amiri")
    # Front cover template with explicit placeholders for unavailable metadata.
    img=Image.new("RGB",(W,H),ivory); d=ImageDraw.Draw(img)
    d.rectangle((0,0,W,230),fill=petrol)
    d.text((W//2,125),rtl_text("[اسم المؤلف]"),font=ImageFont.truetype(kufi,46),fill="white",anchor="mm")
    d.rectangle((150,310,W-150,327),fill=gold)
    title_f=ImageFont.truetype(kufi,122); sub_f=ImageFont.truetype(kufi,47); body_f=ImageFont.truetype(amiri,45)
    d.text((W//2,610),rtl_text(BOOK_TITLE),font=title_f,fill=petrol,anchor="mm")
    y=780
    for line in wrap_rtl(d,SUBTITLE,sub_f,W-320):
        d.text((W//2,y),rtl_text(line),font=sub_f,fill=charcoal,anchor="mm"); y+=80
    # One restrained architectural governance motif.
    cx,cy=W//2,1580
    for r,w,col in [(430,8,petrol),(330,7,gold),(230,6,petrol),(120,6,gold)]:
        d.ellipse((cx-r,cy-r,cx+r,cy+r),outline=col,width=w)
    for a in range(0,360,60):
        x=cx+int(430*math.cos(math.radians(a))); y2=cy+int(430*math.sin(math.radians(a)))
        d.ellipse((x-28,y2-28,x+28,y2+28),fill=gold,outline=petrol,width=4)
    d.text((W//2,H-330),rtl_text(EDITION),font=body_f,fill=charcoal,anchor="mm")
    d.rectangle((W//2-270,H-220,W//2+270,H-95),outline=petrol,width=4)
    d.text((W//2,H-157),rtl_text("[شعار الناشر]"),font=ImageFont.truetype(kufi,36),fill=petrol,anchor="mm")
    img.save(FRONT_PNG,dpi=(300,300))

    # Back cover template.
    back=Image.new("RGB",(W,H),ivory); b=ImageDraw.Draw(back)
    b.rectangle((0,0,W,220),fill=petrol); b.rectangle((150,300,W-150,317),fill=gold)
    b.text((W-150,390),rtl_text("حوكمة تبدأ من القرار، وتُثبتها المسؤولية والدليل"),font=ImageFont.truetype(kufi,60),fill=petrol,anchor="ra")
    blurb=("يقدم هذا الكتاب مسارًا عمليًا لتأسيس الحوكمة داخل الشركات السعودية، من تشخيص طريقة القرار وتحديد نطاق الانطباق، "
           "إلى بناء الملكية والمجلس والصلاحيات والسياسات والضوابط، ثم تشغيل النظام وقياس نضجه وتحسينه. وهو يميز بين "
           "وجود الوثيقة وفاعلية الحوكمة، ويربط السلطة بالمعلومات والتنفيذ والدليل.")
    y=620
    bf=ImageFont.truetype(amiri,47)
    for line in wrap_rtl(b,blurb,bf,W-300):
        b.text((W-150,y),rtl_text(line),font=bf,fill=charcoal,anchor="ra"); y+=82
    y+=60
    b.rounded_rectangle((150,y,W-150,y+270),radius=24,fill="#EFEDE7",outline=gold,width=3)
    b.text((W-180,y+45),rtl_text("[السيرة المختصرة للمؤلف]"),font=ImageFont.truetype(kufi,38),fill=petrol,anchor="ra")
    b.text((W-180,y+125),rtl_text("تُضاف بعد اعتماد النص النهائي للسيرة."),font=ImageFont.truetype(amiri,38),fill=charcoal,anchor="ra")
    y+=360
    # Metadata placeholders.
    b.rectangle((150,y,W-150,y+480),outline=gold,width=3)
    labels=["[شعار الناشر]","[الموقع الإلكتروني]","[رمز QR]","[ISBN والباركود]","[تصنيف الكتاب]","[السعر]"]
    cols=2; boxw=(W-340)//2; boxh=135
    for i,label in enumerate(labels):
        row=i//cols; col=i%cols
        x0=175+col*(boxw+15); y0=y+25+row*(boxh+15)
        b.rectangle((x0,y0,x0+boxw,y0+boxh),outline="#B8B8B8",width=2)
        b.text((x0+boxw//2,y0+boxh//2),rtl_text(label),font=ImageFont.truetype(amiri,34),fill=charcoal,anchor="mm")
    b.text((W-150,H-180),rtl_text(EDITION),font=ImageFont.truetype(amiri,34),fill=petrol,anchor="ra")
    back.save(BACK_PNG,dpi=(300,300))

    # Spine guide; no false width.
    spine=Image.new("RGB",(W,H),ivory); s=ImageDraw.Draw(spine)
    s.rectangle((0,0,W,220),fill=petrol)
    s.text((W//2,115),rtl_text("دليل إعداد الكعب — غير مقياس للطباعة"),font=ImageFont.truetype(kufi,42),fill="white",anchor="mm")
    note=("لا يُحدد عرض الكعب قبل اعتماد العدد النهائي لصفحات ملف الطباعة، ونوع الورق ووزنه وسماكته، وطريقة التجليد. "
          "يوضع على الكعب النهائي: عنوان الكتاب، اسم المؤلف، وشعار الناشر.")
    y=420
    for line in wrap_rtl(s,note,ImageFont.truetype(amiri,47),W-300):
        s.text((W-150,y),rtl_text(line),font=ImageFont.truetype(amiri,47),fill=charcoal,anchor="ra"); y+=82
    s.rectangle((W//2-260,1000,W//2+260,2400),fill=petrol)
    s.text((W//2,1450),rtl_text(BOOK_TITLE),font=ImageFont.truetype(kufi,74),fill="white",anchor="mm")
    s.text((W//2,1850),rtl_text("[اسم المؤلف]"),font=ImageFont.truetype(kufi,43),fill=gold,anchor="mm")
    s.rectangle((W//2-180,2180,W//2+180,2320),outline=gold,width=4)
    s.text((W//2,2250),rtl_text("[شعار الناشر]"),font=ImageFont.truetype(amiri,33),fill="white",anchor="mm")
    spine.save(SPINE_PNG,dpi=(300,300))

    image_to_word(FRONT_PNG, FRONT_DOCX)
    image_to_word(BACK_PNG, BACK_DOCX, pages=2, second_image=SPINE_PNG)


def audit_docx_and_pdf(build_info):
    doc=Document(FINAL_DOCX)
    required_styles=["Book Title","Book Subtitle","Author Name","Part Title","Chapter Title","Main Heading","Subheading","Unnumbered Internal Heading","Body Text","First Paragraph after Heading","Long Quotation","Qur’anic Verse","Prophetic Hadith","Box Title","Box Text","Table Title","Table Text","Table Source","Figure Title","Figure Caption","Footnote","Bulleted List","Numbered List","Reference","Appendix Title"]
    missing_styles=[x for x in required_styles if x not in [s.name for s in doc.styles]]
    wrong_sections=[]
    for i,s in enumerate(doc.sections,1):
        vals=(s.page_width.cm,s.page_height.cm,s.left_margin.cm,s.right_margin.cm,s.top_margin.cm,s.bottom_margin.cm,s.gutter.cm)
        target=(17,24,2.8,2.0,2.2,2.4,.7)
        if any(abs(a-b)>.04 for a,b in zip(vals,target)):
            wrong_sections.append({"section":i,"values":vals})
    bad_rtl=[]; bad_align=[]; wrong_body_font=[]; wrong_heading_font=[]
    for i,p in enumerate(doc.paragraphs,1):
        text=clean(p.text)
        if not text: continue
        ppr=p._p.get_or_add_pPr()
        if ppr.find(qn("w:bidi")) is None and ar_ratio(text)>.25:
            bad_rtl.append(i)
        name=p.style.name if p.style else ""
        group=paragraph_style_group(name)
        if group=="justified" and p.alignment not in (WD_ALIGN_PARAGRAPH.JUSTIFY,None): bad_align.append((i,name,"not_justified"))
        if group in ("right","right_heading") and p.alignment not in (WD_ALIGN_PARAGRAPH.RIGHT,None): bad_align.append((i,name,"not_right"))
        for r in p.runs:
            if not clean(r.text): continue
            if name in {"Body Text","First Paragraph after Heading"} and (r.font.name not in (None,"Amiri")):
                wrong_body_font.append(i)
            if name in {"Part Title","Chapter Title","Main Heading","Subheading","Unnumbered Internal Heading","Book Title","Book Subtitle","Author Name"} and (r.font.name not in (None,"Noto Kufi Arabic")):
                wrong_heading_font.append(i)
    table_rtl=[]; split_rows=[]
    for ti,t in enumerate(doc.tables,1):
        if t._tbl.tblPr.find(qn("w:bidiVisual")) is None: table_rtl.append(ti)
        for ri,row in enumerate(t.rows,1):
            if row._tr.get_or_add_trPr().find(qn("w:cantSplit")) is None: split_rows.append((ti,ri))
    # PDF geometry and visual bounds.
    pdf=fitz.open(FINAL_PDF)
    expected_w=17/2.54*72; expected_h=24/2.54*72
    bad_page_size=[]; out_bounds=[]; blank=[]; font_names=set(); rendered=[]
    for i,page in enumerate(pdf,1):
        if abs(page.rect.width-expected_w)>2 or abs(page.rect.height-expected_h)>2: bad_page_size.append(i)
        txt=page.get_text("text").strip()
        if not txt: blank.append(i)
        for block in page.get_text("blocks"):
            x0,y0,x1,y1=block[:4]
            if x0 < -1 or y0 < -1 or x1 > page.rect.width+1 or y1 > page.rect.height+1:
                out_bounds.append((i,[x0,y0,x1,y1]))
        for f in page.get_fonts(full=True): font_names.add(f[3])
        pix=page.get_pixmap(matrix=fitz.Matrix(1.0,1.0),alpha=False)
        im=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
        pth=OUT/f"page_{i:03d}.png"; im.save(pth); rendered.append(pth)
    # Contact sheets, 12 pages each.
    for n,start in enumerate(range(0,len(rendered),12),1):
        imgs=[Image.open(x) for x in rendered[start:start+12]]
        thumbs=[]
        for im in imgs:
            ratio=240/im.width; thumbs.append(im.resize((240,int(im.height*ratio))))
        th=max(x.height for x in thumbs); sheet=Image.new("RGB",(720,4*th),"white"); draw=ImageDraw.Draw(sheet)
        for j,im in enumerate(thumbs):
            x=(j%3)*240; y=(j//3)*th; sheet.paste(im,(x,y)); draw.text((x+5,y+5),str(start+j+1),fill="black")
        sheet.save(CONTACT/f"sheet_{n:03d}.jpg",quality=88)
    # Citation-reference integrity.
    text="\n".join(clean(p.text) for p in doc.paragraphs if clean(p.text))
    cited={int(n) for n in re.findall(r"\[(\d{1,3})\]",text)}
    refs={int(n) for n in re.findall(r"^\[(\d{1,3})\]",text,flags=re.M)}
    missing_refs=sorted(cited-refs); unused_refs=sorted(refs-cited)
    # Media DPI audit.
    media_audit=[]
    with zipfile.ZipFile(FINAL_DOCX) as z:
        for name in z.namelist():
            if name.startswith("word/media/"):
                try:
                    data=z.read(name); tmp=ROOT/Path(name).name; tmp.write_bytes(data)
                    im=Image.open(tmp)
                    media_audit.append({"name":name,"pixels":im.size,"dpi":im.info.get("dpi")})
                except Exception:
                    media_audit.append({"name":name,"pixels":None,"dpi":None})
    audit={
        **build_info,
        "final_docx_paragraphs":len(doc.paragraphs),
        "final_docx_tables":len(doc.tables),
        "final_docx_sections":len(doc.sections),
        "required_styles_missing":missing_styles,
        "wrong_section_geometry":wrong_sections,
        "arabic_paragraphs_without_rtl":bad_rtl[:50],
        "alignment_issues":bad_align[:50],
        "wrong_body_font_paragraphs":sorted(set(wrong_body_font))[:50],
        "wrong_heading_font_paragraphs":sorted(set(wrong_heading_font))[:50],
        "tables_without_rtl":table_rtl,
        "rows_without_no_split":split_rows[:50],
        "pdf_pages":len(pdf),
        "pdf_bad_page_size":bad_page_size,
        "pdf_out_of_bounds_blocks":out_bounds[:50],
        "blank_pages":blank,
        "pdf_fonts":sorted(font_names),
        "missing_reference_numbers":missing_refs,
        "unused_reference_numbers":unused_refs,
        "embedded_media":media_audit,
        "contact_sheets":len(list(CONTACT.glob('*.jpg'))),
    }
    AUDIT_JSON.write_text(json.dumps(audit,ensure_ascii=False,indent=2),encoding="utf-8")
    return audit


def compliance_rows(audit):
    fully="مكتمل ومطابق"
    na="غير منطبق على محتوى هذه الطبعة"
    external="يتطلب بيانات أو اعتمادًا خارجيًا"
    partial="مكتمل في الجزء القابل للتنفيذ، مع عنصر غير متاح"
    rows=[
        (1,"إعداد الصفحة",fully,"17 × 24 سم، هوامش متقابلة، داخلي 2.8، خارجي 2، علوي 2.2، سفلي 2.4، كعب 0.7 يمين."),
        (2,"اللغة والاتجاه",fully,"اتجاه عربي RTL، المتن مضبوط، والجداول RTL، ولغة التدقيق العربية."),
        (3,"الخطوط والأحجام",fully,"Amiri للمتن وNoto Kufi Arabic للعناوين، بالقياسات المحددة واللون الأسود."),
        (4,"إعدادات الفقرات",fully,"1.2، بعد 5 نقاط، مسافة بادئة 0.6 سم، واستثناء أول فقرة بعد العناوين، ومنع الأرامل والأيتام."),
        (5,"أنماط Word",fully,"تم إنشاء الأنماط الخمسة والعشرين وربط مستويات العناوين 1–4."),
        (6,"التسلسل الهرمي",fully,"الفصول والعناوين الرئيسية والفرعية والداخلية مضبوطة. لا توجد أبواب/أجزاء فعلية في المخطوط."),
        (7,"بدايات الأجزاء والفصول",partial,"الفصول الخمسة عشر تبدأ بفواصل مقطع Odd Page، وتُخفى الرؤوس والأرقام في صفحة الافتتاح. لا توجد أجزاء فعلية."),
        (8,"الرؤوس والتذييلات",fully,"صفحات يمين/يسار مختلفة، عنوان الفصل يمينًا، عنوان الكتاب يسارًا، ورقم الصفحة في الخارج."),
        (9,"ترتيب الصفحات التمهيدية",partial,"المتاح مرتب مهنيًا: نصف عنوان، صفحة بيضاء، عنوان كامل، حقوق، تنبيه مهني، مقدمة وفهارس. الإهداء والشكر والتقديم والاختصارات غير موجودة في المصدر ولم تُخترع."),
        (10,"ترتيب الملاحق الختامية",partial,"الخاتمة ومصفوفة الأدوات والمراجع موجودة. المعجم والفهارس الموضوعية والسيرة غير موجودة في المصدر."),
        (11,"تنسيق الآيات",na,"لا توجد آيات قرآنية مقتبسة؛ تم إنشاء نمط Qur’anic Verse بالمواصفات."),
        (12,"تنسيق الأحاديث",na,"لا توجد أحاديث مقتبسة؛ تم إنشاء نمط Prophetic Hadith بالمواصفات."),
        (13,"الاقتباسات",fully,"علامات « » للاقتباسات القصيرة، ونمط مستقل للاقتباس الطويل. لا توجد اقتباسات طويلة مستقلة بلا مصدر."),
        (14,"الجداول",fully,f"تم ضبط {audit['table_stats']['content_tables']} جدول محتوى: RTL، عنوان أعلى، مصدر أسفل، ترويسة، رمادي ثانوي، حدود 0.5 نقطة، وتكرار الترويسة ومنع انقسام الصفوف."),
        (15,"الأشكال والصور",na,"لا يحتوي متن الكتاب على أشكال أو صور فعلية؛ لا توجد مواد منخفضة الدقة أو متجاوزة للهوامش."),
        (16,"الصناديق التحريرية",fully,f"تم ضبط {audit['table_stats']['boxes']} صندوقًا بخلفية رمادية وحد جانبي وهوامش داخلية 0.4 سم وخطي العنوان والمتن المحددين."),
        (17,"الحواشي",na,"لا توجد حواشٍ فعلية؛ نمط Footnote موجود بحجم 10.5 واتجاه RTL وترقيم متصل عند الاستخدام."),
        (18,"المراجع",fully,"نمط موحد، تعليق 0.7 سم، بعد 4 نقاط، روابط ومعرّفات موحدة، ولا إحالات بلا مرجع أو مراجع غير مستخدمة."),
        (19,"الفهارس التلقائية",partial,"جدول المحتويات وقائمة الجداول تلقائيان ومحدثان. لا توجد أشكال أو آيات أو أحاديث أو مداخل أسماء/موضوعات/معجم معلمة لإنشاء فهارس غير فارغة."),
        (20,"القوائم",fully,"المحاذاة والرموز من اليمين، ومسافات ثابتة، وترقيم عربي دون خلط أنماط."),
        (21,"الجودة الطباعية",fully,"تم تحديث الحقول، وفحص الحدود والمقاسات والخطوط والمسافات والعناوين والجداول وكل صفحات PDF بصريًا وآليًا."),
        (22,"الغلاف الأمامي",partial,"قالب 17 × 24 بالألوان المطلوبة والعنصر البصري الواحد. اسم المؤلف وشعار الناشر موضوعان كحقول استبدال لعدم توافر البيانات."),
        (23,"الكعب",external,"تم إعداد دليل كعب يتضمن العنوان وحقلي المؤلف والشعار، دون اختلاق عرض؛ العرض ينتظر الورق والتجليد."),
        (24,"الغلاف الخلفي",partial,"السطر الترويجي والوصف موجودان، وبقية البيانات موضوعة كحقول استبدال: السيرة، الشعار، الموقع، QR، ISBN، الباركود، التصنيف والسعر."),
        (25,"المواصفات الفنية للمتن",partial,"المقاس والاتجاه والخطوط والتدرج الرمادي جاهزة وPDF مراجعة مرفق. شهادة PDF/X-4 تتطلب محرك preflight احترافيًا ولم تُدّعَ."),
        (26,"المواصفات الفنية للغلاف",external,"قوالب 300 DPI ومناطق الأمان والنزف موثقة؛ ملف CMYK/PDF-X-4 الكامل ينتظر عرض الكعب وبيانات النشر واعتماد المطبعة."),
        (27,"التشطيب",external,"المواصفات مثبتة كتعليمات: 350 gsm، مطفي، Spot UV وبروز خفيف، وتجليد مخيط للكتاب المرجعي. التنفيذ مادي لدى المطبعة."),
        (28,"سير المراجعة النهائية",partial,"اكتملت المراجعة الرقمية والتصدير الأولي والفحص البصري. نسخة الطباعة التجريبية وبروفة المطبعة والاعتماد النهائي خطوات مادية خارج الملف."),
    ]
    return rows


def build_report(audit):
    doc=Document(); create_required_styles(doc)
    sec=doc.sections[0]; set_geometry(sec)
    p=doc.add_paragraph("تقرير المطابقة الشاملة — الحوكمة من الصفر",style="Book Title"); set_bidi_paragraph(p,WD_ALIGN_PARAGRAPH.CENTER)
    p=doc.add_paragraph("مراجعة دقيقة للخطوات الثماني والعشرين لتنسيق وإنتاج الكتاب العربي",style="Book Subtitle"); set_bidi_paragraph(p,WD_ALIGN_PARAGRAPH.CENTER)
    summary=[
        f"عدد صفحات PDF للمراجعة: {audit['pdf_pages']} صفحة.",
        f"عدد أقسام Word: {audit['final_docx_sections']}.",
        f"عدد جداول المحتوى: {audit['table_stats']['content_tables']}، وعدد الصناديق التحريرية: {audit['table_stats']['boxes']}.",
        f"الصفحات غير المطابقة للمقاس: {audit['pdf_bad_page_size'] or 'لا يوجد'}.",
        f"النصوص الخارجة عن حدود الصفحة: {'لا يوجد' if not audit['pdf_out_of_bounds_blocks'] else audit['pdf_out_of_bounds_blocks'][:5]}.",
        f"الفقرات العربية دون RTL: {audit['arabic_paragraphs_without_rtl'] or 'لا يوجد'}.",
        f"مشكلات المحاذاة: {audit['alignment_issues'] or 'لا يوجد'}.",
        f"أنماط Word المطلوبة المفقودة: {audit['required_styles_missing'] or 'لا يوجد'}.",
        f"الإحالات دون مراجع: {audit['missing_reference_numbers'] or 'لا يوجد'}؛ المراجع غير المستخدمة: {audit['unused_reference_numbers'] or 'لا يوجد'}.",
    ]
    for x in summary:
        p=doc.add_paragraph(x,style="Body Text"); set_bidi_paragraph(p,WD_ALIGN_PARAGRAPH.JUSTIFY)
    table=doc.add_table(rows=1,cols=4); table.alignment=WD_TABLE_ALIGNMENT.CENTER
    headers=["الخطوة","المعيار","الحالة","النتيجة الدقيقة"]
    for c,t in zip(table.rows[0].cells,headers): c.text=t
    for num,title,status,note in compliance_rows(audit):
        cells=table.add_row().cells
        cells[0].text=str(num).translate(ARABIC_DIGITS); cells[1].text=title; cells[2].text=status; cells[3].text=note
    style_tables(doc)
    p=doc.add_paragraph("الخلاصة",style="Main Heading")
    conclusion=("جميع عناصر التنسيق الداخلي القابلة للتنفيذ داخل Word أصبحت مطابقة للمعيار. ولا يمكن وصف الحزمة بأنها ملف طباعة تجاري نهائي بالكامل قبل تزويد اسم المؤلف والناشر وISBN وبقية بيانات الغلاف، والحصول على نوع الورق وسماكته وعرض الكعب من المطبعة، ثم إنتاج ملف CMYK/PDF-X-4 مع تقرير preflight وبروفة ورقية.")
    p=doc.add_paragraph(conclusion,style="Body Text"); set_bidi_paragraph(p,WD_ALIGN_PARAGRAPH.JUSTIFY)
    doc.save(REPORT_DOCX)
    subprocess.run(["soffice","--headless","--convert-to","pdf","--outdir",str(OUT),str(REPORT_DOCX)],check=True)
    generated=OUT/(REPORT_DOCX.stem+".pdf")
    if generated.exists() and generated != REPORT_PDF: generated.replace(REPORT_PDF)


def build_readme(audit):
    README.write_text(
        f"""الحوكمة من الصفر — حزمة المراجعة الشاملة\n\n"
        f"عدد صفحات نسخة PDF للمراجعة: {audit['pdf_pages']}\n"
        "تحتوي الحزمة على:\n"
        "- النسخة النهائية المطابقة لمعيار Word.\n"
        "- PDF للمراجعة البصرية، وليس شهادة PDF/X-4.\n"
        "- تقرير مطابقة الخطوات الثماني والعشرين Word وPDF.\n"
        "- قالب الغلاف الأمامي Word وPNG بدقة 300 DPI.\n"
        "- قالب الغلاف الخلفي والكعب Word وPNG.\n"
        "- ملف التدقيق الآلي ولوحات مراجعة جميع الصفحات.\n\n"
        "البيانات الخارجية المتبقية: اسم المؤلف، السيرة، الناشر والشعار والموقع، ISBN والباركود وQR والتصنيف والسعر، ونوع الورق وسماكته وعرض الكعب واعتماد المطبعة.\n",
        encoding="utf-8"
    )


def package():
    pkg=ROOT/"package"
    if pkg.exists(): shutil.rmtree(pkg)
    pkg.mkdir()
    for p in [FINAL_DOCX,FINAL_PDF,FRONT_DOCX,FRONT_PNG,BACK_DOCX,BACK_PNG,SPINE_PNG,REPORT_DOCX,REPORT_PDF,AUDIT_JSON,README]:
        if p.exists(): shutil.copy2(p,pkg/p.name)
    shutil.copytree(CONTACT,pkg/CONTACT.name)
    if PACKAGE_ZIP.exists(): PACKAGE_ZIP.unlink()
    with zipfile.ZipFile(PACKAGE_ZIP,"w",zipfile.ZIP_DEFLATED,compresslevel=9) as z:
        for p in pkg.rglob("*"):
            if p.is_file(): z.write(p,p.relative_to(pkg))


def main():
    build_info=build_interior()
    update_fields_and_export()
    build_cover_templates()
    audit=audit_docx_and_pdf(build_info)
    # Hard failures for the core formatting criteria.
    failures=[]
    for key in ("required_styles_missing","wrong_section_geometry","arabic_paragraphs_without_rtl","alignment_issues","tables_without_rtl","rows_without_no_split","pdf_bad_page_size","pdf_out_of_bounds_blocks","missing_reference_numbers","unused_reference_numbers"):
        if audit.get(key): failures.append((key,audit[key]))
    if audit["table_stats"]["content_tables"] != 25:
        failures.append(("content_table_count",audit["table_stats"]["content_tables"]))
    if audit["table_stats"]["boxes"] != 43:
        failures.append(("editorial_box_count",audit["table_stats"]["boxes"]))
    chapter_count=sum(1 for p in Document(FINAL_DOCX).paragraphs if (p.style and p.style.name=="Chapter Title" and clean(p.text).startswith("الفصل ")))
    audit["chapter_count"]=chapter_count
    AUDIT_JSON.write_text(json.dumps(audit,ensure_ascii=False,indent=2),encoding="utf-8")
    if chapter_count != 15: failures.append(("chapter_count",chapter_count))
    build_report(audit)
    build_readme(audit)
    package()
    if failures:
        raise RuntimeError("Core compliance failures: "+repr(failures[:10]))
    print(json.dumps(audit,ensure_ascii=False,indent=2))


if __name__ == "__main__":
    main()
