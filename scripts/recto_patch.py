#!/usr/bin/env python3
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


CHAPTER_RE = re.compile(
    r"^الفصل\s+(?:الأول|الثاني|الثالث|الرابع|الخامس|السادس|السابع|الثامن|التاسع|العاشر|"
    r"الحادي\s+عشر|الثاني\s+عشر|الثالث\s+عشر|الرابع\s+عشر|الخامس\s+عشر|[0-9٠-٩]+)"
)


def force_chapter_recto_breaks(docx_path: str | Path) -> dict:
    """Set existing section breaks before the 15 real chapter titles to odd-page.

    The function changes only w:type inside existing section properties. It does not
    add, remove, or reorder paragraphs, text, sections, headers, footers, or tables.
    """
    docx_path = Path(docx_path)
    with tempfile.TemporaryDirectory(prefix="recto_patch_") as tmp_name:
        tmp = Path(tmp_name)
        with zipfile.ZipFile(docx_path) as archive:
            archive.extractall(tmp)

        document_path = tmp / "word" / "document.xml"
        styles_path = tmp / "word" / "styles.xml"
        document_tree = etree.parse(str(document_path))
        styles_tree = etree.parse(str(styles_path))

        style_matches = styles_tree.xpath(
            '//w:style[w:name[@w:val="Chapter Title"]]', namespaces=NS
        )
        if not style_matches:
            raise RuntimeError("The paragraph style 'Chapter Title' was not found.")
        chapter_style_id = style_matches[0].get(w("styleId"))

        body = document_tree.find(f".//{w('body')}")
        if body is None:
            raise RuntimeError("The DOCX document body was not found.")
        children = list(body)
        chapter_paragraphs: list[tuple[etree._Element, str]] = []
        for element in children:
            if element.tag != w("p"):
                continue
            pstyle = element.find(f".//{w('pStyle')}")
            text = "".join(element.xpath(".//w:t/text()", namespaces=NS)).strip()
            if (
                pstyle is not None
                and pstyle.get(w("val")) == chapter_style_id
                and CHAPTER_RE.match(text)
            ):
                chapter_paragraphs.append((element, text))

        if len(chapter_paragraphs) != 15:
            raise RuntimeError(
                f"Expected 15 real chapter-title paragraphs, found {len(chapter_paragraphs)}."
            )

        updated: list[str] = []
        without_break: list[str] = []
        for chapter, text in chapter_paragraphs:
            index = children.index(chapter)
            section_properties = None
            # Existing section-break paragraphs are immediately before chapter titles,
            # with occasional empty paragraphs between them. Search a narrow window.
            for previous_index in range(index - 1, max(-1, index - 10), -1):
                candidate = children[previous_index]
                if candidate.tag != w("p"):
                    continue
                ppr = candidate.find(w("pPr"))
                found = ppr.find(w("sectPr")) if ppr is not None else None
                if found is not None:
                    section_properties = found
                    break
            if section_properties is None:
                without_break.append(text)
                continue

            break_type = section_properties.find(w("type"))
            if break_type is None:
                break_type = etree.Element(w("type"))
                section_properties.insert(0, break_type)
            break_type.set(w("val"), "oddPage")
            updated.append(text)

        # The first chapter can begin in an already-established section. Every later
        # chapter must have an existing break available for conversion.
        if len(updated) < 14:
            raise RuntimeError(
                f"Only {len(updated)} chapter breaks were found; missing: {without_break}"
            )

        document_tree.write(
            str(document_path),
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

        rebuilt = docx_path.with_name(docx_path.stem + "_recto_patched.docx")
        with zipfile.ZipFile(rebuilt, "w", zipfile.ZIP_DEFLATED) as archive:
            for item in tmp.rglob("*"):
                if item.is_file():
                    archive.write(item, item.relative_to(tmp))
        shutil.copy2(rebuilt, docx_path)
        rebuilt.unlink()

    return {
        "chapter_count": len(chapter_paragraphs),
        "updated_break_count": len(updated),
        "without_preceding_break": without_break,
    }


if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    args = parser.parse_args()
    print(json.dumps(force_chapter_recto_breaks(args.docx), ensure_ascii=False, indent=2))
