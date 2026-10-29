"""Validate formatting inputs and generated OOXML before publishing a report.

This is a targeted style/reference validator, not a desktop Excel test or a
complete ECMA-376 schema validator. In particular, XML well-formedness alone
is NOT sufficient: <family val="2.0"/> parses but violates Excel's integer type.
"""
from __future__ import annotations

from math import isfinite
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

try:
    from .utils import ReportError
except ImportError:
    from utils import ReportError

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_INTEGER = re.compile(r"[+-]?[0-9]+\Z")
_RGB = re.compile(r"[0-9A-Fa-f]{8}\Z")
# XlsxWriter format properties whose serialized values/enums require integers.
_INTEGER_PROPERTIES = {
    "font_family": (0, 5), "font_charset": (0, 255), "font_script": (0, 2),
    "pattern": (0, 18), "border": (0, 13), "left": (0, 13),
    "right": (0, 13), "top": (0, 13), "bottom": (0, 13),
    "diag_border": (0, 13), "diag_type": (0, 3), "indent": (0, 250),
    "reading_order": (0, 2), "rotation": (-90, 270),
}


def format_properties(properties: dict, name: str = "format") -> dict:
    """Normalize exact integral numbers; reject fractional/invalid style codes.

    Existing JSON from the previous migration contains font_family=2.0. This
    normalizes it to Python int(2), before XlsxWriter serializes styles.xml.
    Font sizes and column widths are deliberately NOT rounded to integers.
    """
    result = dict(properties)
    for key, (minimum, maximum) in _INTEGER_PROPERTIES.items():
        if key not in result:
            continue
        value = result[key]
        if (isinstance(value, bool) or not isinstance(value, (int, float))
                or not isfinite(value) or int(value) != value
                or not minimum <= value <= maximum):
            raise ReportError(f"Invalid Onboarding style {name}.{key}: expected an integer from {minimum} to {maximum}.")
        result[key] = int(value)
    if "underline" in result:
        value = result["underline"]
        if isinstance(value, bool) or value not in (0, 1, 2, 33, 34):
            raise ReportError(f"Invalid Onboarding style {name}.underline.")
        result["underline"] = int(value)
    for key, value in result.items():
        if key.endswith("_color") and (not isinstance(value, str) or not re.fullmatch(r"#[0-9A-Fa-f]{6}", value)):
            raise ReportError(f"Invalid Onboarding style {name}.{key}: expected #RRGGBB.")
    return result


def _integer(value: str | None, label: str, minimum: int = 0, maximum: int | None = None) -> int:
    if value is None or not _INTEGER.fullmatch(value):
        raise ReportError(f"Onboarding workbook validation failed: {label} must be an integer, not {value!r}.")
    number = int(value)
    if number < minimum or (maximum is not None and number > maximum):
        raise ReportError(f"Onboarding workbook validation failed: {label} is outside its allowed range.")
    return number


def _colors(root: ET.Element, part: str) -> None:
    for node in root.iter():
        if "rgb" in node.attrib and not _RGB.fullmatch(node.attrib["rgb"]):
            raise ReportError(f"Onboarding workbook validation failed: invalid RGB color in {part}.")


def validate_workbook(path: Path) -> dict:
    """Reject invalid styles/references before the temporary file is published.

    Stream worksheet XML to avoid retaining large Base sheets in memory. Return
    a small verification summary. This function does not rewrite the workbook.
    """
    try:
        with ZipFile(path) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ReportError("Onboarding workbook validation failed: duplicate ZIP parts.")
            styles = ET.fromstring(archive.read("xl/styles.xml"))
            if styles.tag != _NS + "styleSheet":
                raise ReportError("Onboarding workbook validation failed: missing stylesheet root.")
            sections = {}
            expected_child = {"fonts": "font", "fills": "fill", "borders": "border",
                              "cellStyleXfs": "xf", "cellXfs": "xf", "dxfs": "dxf",
                              "numFmts": "numFmt", "cellStyles": "cellStyle"}
            for section, child in expected_child.items():
                node = styles.find(_NS + section)
                children = [] if node is None else list(node)
                if node is not None:
                    if any(n.tag != _NS + child for n in children):
                        raise ReportError(f"Onboarding workbook validation failed: unexpected element in {section}.")
                    count = _integer(node.get("count"), f"styles.xml/{section}/@count")
                    if count != len(children):
                        raise ReportError(f"Onboarding workbook validation failed: incorrect {section} count.")
                sections[section] = children
            for section in ("fonts", "fills", "borders", "cellXfs", "cellStyleXfs"):
                if not sections[section]:
                    raise ReportError(f"Onboarding workbook validation failed: {section} is empty.")
            for node in styles.iter(_NS + "family"):
                _integer(node.get("val"), "styles.xml/font/family/@val", maximum=5)
            for node in styles.iter(_NS + "charset"):
                _integer(node.get("val"), "styles.xml/font/charset/@val", maximum=255)
            custom_formats = set()
            for node in sections["numFmts"]:
                value = _integer(node.get("numFmtId"), "styles.xml/numFmtId", minimum=164)
                if value in custom_formats or not node.get("formatCode"):
                    raise ReportError("Onboarding workbook validation failed: duplicate or empty number format.")
                custom_formats.add(value)
            for section in ("cellXfs", "cellStyleXfs"):
                for node in sections[section]:
                    for attribute, target in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                        _integer(node.get(attribute, "0"), f"styles.xml/xf/@{attribute}", maximum=len(sections[target])-1)
                    num_format = _integer(node.get("numFmtId", "0"), "styles.xml/xf/@numFmtId")
                    if num_format >= 164 and num_format not in custom_formats:
                        raise ReportError("Onboarding workbook validation failed: undefined number format.")
                    if "xfId" in node.attrib:
                        _integer(node.get("xfId"), "styles.xml/xf/@xfId", maximum=len(sections["cellStyleXfs"])-1)
                    alignment = node.find(_NS + "alignment")
                    if alignment is not None:
                        for attr, limit in (("indent", 250), ("readingOrder", 2), ("textRotation", 255)):
                            if attr in alignment.attrib:
                                _integer(alignment.get(attr), f"styles.xml/alignment/@{attr}", maximum=limit)
            for node in sections["cellStyles"]:
                _integer(node.get("xfId"), "styles.xml/cellStyle/@xfId", maximum=len(sections["cellStyleXfs"])-1)
            _colors(styles, "styles.xml")
            checked_cells, checked_rules, worksheet_count = 0, 0, 0
            for part in names:
                if not part.endswith((".xml", ".rels")) or part == "xl/styles.xml":
                    continue
                is_sheet = bool(re.fullmatch(r"xl/worksheets/sheet[0-9]+\.xml", part))
                if not is_sheet:
                    ET.fromstring(archive.read(part))
                    continue
                worksheet_count += 1
                with archive.open(part) as handle:
                    for _, node in ET.iterparse(handle, events=("end",)):
                        if node.tag == _NS + "c":
                            _integer(node.get("s", "0"), f"{part}/{node.get('r')}/@s", maximum=len(sections["cellXfs"])-1)
                            checked_cells += 1
                            node.clear()
                        elif node.tag == _NS + "row":
                            if "s" in node.attrib:
                                _integer(node.get("s"), f"{part}/row/@s", maximum=len(sections["cellXfs"])-1)
                            node.clear()
                        elif node.tag == _NS + "sheetData":
                            node.clear()
                        elif node.tag == _NS + "col" and "style" in node.attrib:
                            _integer(node.get("style"), f"{part}/col/@style", maximum=len(sections["cellXfs"])-1)
                        elif node.tag == _NS + "cfRule":
                            checked_rules += 1
                            if "dxfId" in node.attrib:
                                _integer(node.get("dxfId"), f"{part}/cfRule/@dxfId", maximum=len(sections["dxfs"])-1)
                            _integer(node.get("priority"), f"{part}/cfRule/@priority", minimum=1)
                            _colors(node, part)
                            scale = node.find(_NS + "colorScale")
                            if scale is not None:
                                stops, colors = scale.findall(_NS+"cfvo"), scale.findall(_NS+"color")
                                if len(stops) not in (2, 3) or len(stops) != len(colors):
                                    raise ReportError(f"Onboarding workbook validation failed: invalid color scale in {part}.")
                            node.clear()
            if worksheet_count != 7:
                raise ReportError("Onboarding workbook validation failed: expected seven worksheets.")
            return {"worksheets": worksheet_count, "cell_styles": len(sections["cellXfs"]),
                    "fonts": len(sections["fonts"]), "checked_cells": checked_cells,
                    "conditional_rules": checked_rules, "invalid_style_references": 0}
    except ReportError:
        raise
    except (BadZipFile, ET.ParseError, KeyError, ValueError, OSError) as exc:
        raise ReportError(f"Onboarding workbook validation failed: {exc}") from exc
