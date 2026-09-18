#!/usr/bin/env python3
"""Structural checks for academic DOCX conversion.

This intentionally complements, rather than replaces, visual review in Word.
It requires only the Python standard library.
"""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}
TAG = lambda namespace, local: f"{{{namespace}}}{local}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def node_text(node: ET.Element) -> str:
    return "".join((item.text or "") for item in node.iter() if local_name(item.tag) == "t")


def has_math_operand_content(element: ET.Element | None) -> bool:
    if element is None:
        return False
    return any((item.text or "").strip() for item in element.iter() if item.tag == TAG(M, "t"))


def paragraphs(root: ET.Element) -> list[dict[str, str]]:
    output: list[dict[str, str]] = []
    for paragraph in root.findall(".//w:p", NS):
        style = paragraph.find("./w:pPr/w:pStyle", NS)
        output.append(
            {
                "text": node_text(paragraph),
                "style": "" if style is None else style.get(TAG(W, "val"), ""),
            }
        )
    return output


def bibliography_numbers(items: Iterable[dict[str, str]], title: str | None) -> list[int]:
    source = list(items)
    if title:
        title_positions = [index for index, item in enumerate(source) if item["text"].strip() == title]
        if not title_positions:
            return []
        source = source[title_positions[-1] + 1 :]
    numbers: list[int] = []
    for item in source:
        match = re.match(r"^\s*\[(\d+)\]\s+", item["text"])
        if match:
            numbers.append(int(match.group(1)))
    return numbers


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--require-valid-zip", action="store_true")
    parser.add_argument("--fail-on-empty-nary", action="store_true")
    parser.add_argument("--bibliography-title")
    parser.add_argument("--expected-references", type=int)
    parser.add_argument("--forbid-text", action="append", default=[])
    parser.add_argument("--forbid-marker", action="store_true")
    parser.add_argument("--require-heading-style", action="store_true")
    parser.add_argument("--require-a4", action="store_true")
    parser.add_argument("--require-no-header-footer", action="store_true")
    args = parser.parse_args()

    report: dict[str, object] = {"file": str(args.docx), "errors": [], "warnings": [], "checks": {}}
    errors: list[str] = report["errors"]  # type: ignore[assignment]
    warnings: list[str] = report["warnings"]  # type: ignore[assignment]

    if not args.docx.is_file():
        errors.append("DOCX file does not exist.")
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 2

    try:
        with zipfile.ZipFile(args.docx) as package:
            bad_member = package.testzip()
            report["checks"]["zip_valid"] = bad_member is None
            if args.require_valid_zip and bad_member:
                errors.append(f"Corrupt ZIP member: {bad_member}")
            if "word/document.xml" not in package.namelist():
                errors.append("word/document.xml is missing.")
                print(json.dumps(report, ensure_ascii=False, indent=2))
                return 2
            document = ET.fromstring(package.read("word/document.xml"))
            package_members = package.namelist()
    except (OSError, zipfile.BadZipFile, ET.ParseError) as error:
        errors.append(f"Cannot read DOCX: {error}")
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 2

    all_text = node_text(document)
    docs_paragraphs = paragraphs(document)
    nary_nodes = document.findall(".//m:nary", NS)
    empty_nary = [
        node
        for node in nary_nodes
        if not has_math_operand_content(node.find("./m:e", NS))
    ]
    report["checks"]["omml_objects"] = len(document.findall(".//m:oMath", NS)) + len(
        document.findall(".//m:oMathPara", NS)
    )
    report["checks"]["nary_operators"] = len(nary_nodes)
    report["checks"]["empty_nary_operands"] = len(empty_nary)
    report["checks"]["manual_line_breaks"] = len(document.findall(".//w:br", NS))
    report["checks"]["source_page_markers"] = len(re.findall(r"【原书第[^】]+页】", all_text))

    if args.fail_on_empty_nary and empty_nary:
        errors.append(f"{len(empty_nary)} OMML large operator(s) have an empty operand slot.")
    if args.forbid_marker and report["checks"]["source_page_markers"]:
        errors.append("Source-page markers remain in the document.")

    forbidden = {text: all_text.count(text) for text in args.forbid_text}
    report["checks"]["forbidden_text"] = forbidden
    for text, count in forbidden.items():
        if count:
            errors.append(f'Forbidden text appears {count} time(s): "{text}"')

    if args.bibliography_title or args.expected_references is not None:
        numbers = bibliography_numbers(docs_paragraphs, args.bibliography_title)
        actual_sequence = list(range(1, len(numbers) + 1))
        report["checks"]["bibliography_numbers"] = {
            "count": len(numbers),
            "first": numbers[0] if numbers else None,
            "last": numbers[-1] if numbers else None,
            "continuous": numbers == actual_sequence,
        }
        if args.bibliography_title and not numbers:
            errors.append(f'No numbered bibliography was found after "{args.bibliography_title}".')
        if args.expected_references is not None and len(numbers) != args.expected_references:
            errors.append(
                f"Bibliography count is {len(numbers)}, expected {args.expected_references}."
            )
        if numbers and numbers != actual_sequence:
            errors.append("Bibliography numbers are not a continuous sequence starting at 1.")

        if args.require_heading_style and args.bibliography_title:
            matches = [
                item
                for item in docs_paragraphs
                if item["text"].strip() == args.bibliography_title
            ]
            if len(matches) != 1:
                errors.append(
                    f'Expected exactly one bibliography title "{args.bibliography_title}", found {len(matches)}.'
                )
            elif not matches[0]["style"].lower().startswith("heading"):
                errors.append("Bibliography title is not a Word Heading style.")

    if args.require_a4:
        page_sizes = document.findall(".//w:sectPr/w:pgSz", NS)
        is_a4 = any(
            size.get(TAG(W, "w")) in {"11906", "11907"}
            and size.get(TAG(W, "h")) in {"16838", "16839"}
            for size in page_sizes
        )
        report["checks"]["a4_portrait"] = is_a4
        if not is_a4:
            errors.append("No A4 portrait page-size setting was found.")

    if args.require_no_header_footer:
        header_footer_parts = [
            name
            for name in package_members
            if re.fullmatch(r"word/(header|footer)\d+\.xml", name)
        ]
        report["checks"]["header_footer_parts"] = header_footer_parts
        if header_footer_parts:
            errors.append("Header or footer parts are present.")

    if not report["checks"]["omml_objects"]:
        warnings.append("No OMML math object was found; verify whether the source actually contains formulas.")

    report["ok"] = not errors
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if not errors else 2


if __name__ == "__main__":
    raise SystemExit(main())
