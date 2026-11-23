#!/usr/bin/env python3
import argparse
from pathlib import Path
import re

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


WHITESPACE = r"\t \u00a0\u3000"
CJK = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
BULLET_SPACE_PATTERN = re.compile(rf"•(?P<spaces>[{WHITESPACE}]+)")
OPENING_PUNCT = r"\(\[{<（【《〈“‘〔「『"
CLOSING_PUNCT = r"\)\]}>）】》〉”’〕」』"
INLINE_SPACE_PATTERNS = (
    re.compile(
        rf"(?<=[{CJK}])[{WHITESPACE}]+(?=[A-Za-z0-9])"
        rf"|(?<=[A-Za-z0-9])[{WHITESPACE}]+(?=[{CJK}])"
    ),
    re.compile(
        rf"(?<=[0-9])[{WHITESPACE}]+(?=[A-Za-z])"
        rf"|(?<=[A-Za-z])[{WHITESPACE}]+(?=[0-9])"
    ),
    re.compile(
        rf"(?<=[A-Za-z0-9])[{WHITESPACE}]+(?=[/／])"
        rf"|(?<=[/／])[{WHITESPACE}]+(?=[A-Za-z0-9])"
    ),
    re.compile(rf"(?<=[{CJK}])[{WHITESPACE}]+(?=[{OPENING_PUNCT}])"),
    re.compile(rf"(?<=[{CLOSING_PUNCT}])[{WHITESPACE}]+(?=[{CJK}])"),
)
AUTO_SPACE_TAGS = ("w:autoSpaceDE", "w:autoSpaceDN")
SUCCESSOR_TAGS = (
    "w:bidi",
    "w:adjustRightInd",
    "w:snapToGrid",
    "w:spacing",
    "w:ind",
    "w:contextualSpacing",
    "w:mirrorIndents",
    "w:suppressOverlap",
    "w:jc",
    "w:textDirection",
    "w:textAlignment",
    "w:textboxTightWrap",
    "w:outlineLvl",
    "w:divId",
    "w:cnfStyle",
    "w:rPr",
    "w:sectPr",
    "w:pPrChange",
)


def paragraph_roots(document):
    yield document.element.body
    for relationship in document.part.rels.values():
        if relationship.is_external:
            continue
        if relationship.reltype not in {RT.HEADER, RT.FOOTER}:
            continue
        element = getattr(relationship.target_part, "element", None)
        if element is not None:
            yield element


def owned_text_nodes(paragraph):
    for node in paragraph.iter(qn("w:t")):
        ancestor = node.getparent()
        while ancestor is not None and ancestor.tag != qn("w:p"):
            ancestor = ancestor.getparent()
        if ancestor is paragraph:
            yield node


def remove_text_at_offsets(text, offsets):
    if not offsets:
        return text
    return "".join(char for index, char in enumerate(text) if index not in offsets)


def clean_paragraph_text(paragraph):
    nodes = list(owned_text_nodes(paragraph))
    if not nodes:
        return 0, 0

    text = "".join(node.text or "" for node in nodes)
    bullet_offsets = set()
    inline_offsets = set()

    for match in BULLET_SPACE_PATTERN.finditer(text):
        bullet_offsets.update(range(*match.span("spaces")))
    for pattern in INLINE_SPACE_PATTERNS:
        for match in pattern.finditer(text):
            inline_offsets.update(range(*match.span()))

    offsets = bullet_offsets | inline_offsets
    if not offsets:
        return 0, 0

    cursor = 0
    for node in nodes:
        original = node.text or ""
        local_offsets = {
            index - cursor
            for index in offsets
            if cursor <= index < cursor + len(original)
        }
        updated = remove_text_at_offsets(original, local_offsets)
        if updated != original:
            node.text = updated
        cursor += len(original)

    return len(bullet_offsets), len(inline_offsets)


def disable_auto_spacing(paragraph):
    p_pr = paragraph.find(qn("w:pPr"))
    if p_pr is None:
        p_pr = OxmlElement("w:pPr")
        paragraph.insert(0, p_pr)

    changed = False
    for tag in AUTO_SPACE_TAGS:
        element = p_pr.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            p_pr.insert_element_before(element, *SUCCESSOR_TAGS)
            changed = True
        if element.get(qn("w:val")) != "0":
            element.set(qn("w:val"), "0")
            changed = True
    return changed


def clean_document(input_path, output_path):
    document = Document(input_path)
    paragraph_count = 0
    auto_spacing_changed = 0
    bullet_spaces_removed = 0
    inline_spaces_removed = 0

    for root in paragraph_roots(document):
        for paragraph in root.iter(qn("w:p")):
            paragraph_count += 1
            bullets, inline_spaces = clean_paragraph_text(paragraph)
            bullet_spaces_removed += bullets
            inline_spaces_removed += inline_spaces
            auto_spacing_changed += int(disable_auto_spacing(paragraph))

    document.save(output_path)
    return {
        "paragraphs": paragraph_count,
        "auto_spacing_changed": auto_spacing_changed,
        "bullet_spaces_removed": bullet_spaces_removed,
        "inline_spaces_removed": inline_spaces_removed,
    }


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Remove spaces after bullet markers and at CJK, digit, Latin, "
            "slash, and bracket boundaries in a DOCX."
        )
    )
    parser.add_argument("input", type=Path, help="Source .docx")
    parser.add_argument("output", type=Path, help="New output .docx")
    args = parser.parse_args()

    if args.input.suffix.lower() != ".docx" or args.output.suffix.lower() != ".docx":
        parser.error("input and output must both be .docx files")
    if args.input.resolve() == args.output.resolve():
        parser.error("output must differ from input; preserve the source document")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    result = clean_document(args.input, args.output)
    print(
        f"paragraphs={result['paragraphs']} "
        f"auto_spacing_changed={result['auto_spacing_changed']} "
        f"bullet_spaces_removed={result['bullet_spaces_removed']} "
        f"inline_spaces_removed={result['inline_spaces_removed']}"
    )


if __name__ == "__main__":
    main()
