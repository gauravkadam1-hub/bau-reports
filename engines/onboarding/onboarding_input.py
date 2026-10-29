"""Input metadata validation shared by Polars ingestion and dependency-free tests."""
from __future__ import annotations
import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile
try:
    from .onboarding_common import DATE_COLUMN
    from .utils import ReportError, normalize_header, parse_date_value
except ImportError:
    from onboarding_common import DATE_COLUMN
    from utils import ReportError, normalize_header, parse_date_value
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

def validate_headers(headers):
    """Validate raw headers BEFORE a dataframe reader can rename duplicates."""
    mapping, duplicates = {}, []
    for index, raw in enumerate(headers):
        key = normalize_header(raw)
        if not key:
            continue
        if key in mapping:
            duplicates.append(str(raw))
        else:
            mapping[key] = index
    if duplicates:
        raise ReportError("The uploaded base contains duplicate headers: " + ", ".join(sorted(set(duplicates))))
    return mapping


def _column_number(reference):
    result = 0
    for character in re.match(r"[A-Za-z]+", reference).group(0).upper():
        result = result*26 + ord(character)-64
    return result-1


def _xlsx_headers(path):
    """Read only sheet names and the raw first header row, not report data.

    This small metadata check preserves duplicate-header validation when Calamine
    would otherwise make duplicate column names unique automatically.
    """
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in relationships if r.attrib.get("TargetMode") != "External"}
        sheets, needed_strings = [], set()
        for sheet in workbook.find(_NS+"sheets"):
            target = targets.get(sheet.attrib[_REL+"id"])
            if target is None:
                continue
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/"+target)
            if not part.startswith("xl/") or part not in archive.namelist():
                raise ReportError("The uploaded workbook contains an invalid worksheet relationship.")
            cells = {}
            with archive.open(part) as handle:
                for event, node in ET.iterparse(handle, events=("end",)):
                    if node.tag != _NS+"row":
                        continue
                    if int(node.attrib.get("r", "1")) == 1:
                        for cell in node.findall(_NS+"c"):
                            col = _column_number(cell.attrib["r"])
                            kind = cell.attrib.get("t")
                            if kind == "inlineStr":
                                cells[col] = ("text", "".join(t.text or "" for t in cell.findall(".//"+_NS+"t")))
                            else:
                                value = cell.find(_NS+"v")
                                text = value.text if value is not None else None
                                if kind == "s" and text is not None:
                                    needed_strings.add(int(text)); cells[col] = ("shared", int(text))
                                else:
                                    cells[col] = ("text", text)
                    break
            sheets.append({"name": sheet.attrib["name"], "part": part, "cells": cells})
        shared = {}
        if needed_strings:
            with archive.open("xl/sharedStrings.xml") as handle:
                index = 0
                for event, node in ET.iterparse(handle, events=("end",)):
                    if node.tag == _NS+"si":
                        if index in needed_strings:
                            shared[index] = "".join(t.text or "" for t in node.findall(".//"+_NS+"t"))
                        node.clear()
                        index += 1
                        if len(shared) == len(needed_strings):
                            break
        for sheet in sheets:
            cells = sheet.pop("cells")
            headers = [None] * (max(cells, default=-1)+1)
            for col, (kind, val) in cells.items():
                headers[col] = shared[val] if kind == "shared" else val
            sheet["headers"] = headers
        return sheets


def _mixed_numeric_date_overrides(path, sheet, source_values):
    """Disambiguate native Excel serial cells coerced to strings in a mixed column.

    A literal text '46000' stays invalid (legacy behavior); a numeric cell 46000
    is an Excel serial. Scan this column's XML only when this ambiguity exists.
    """
    candidates = {i for i, value in enumerate(source_values) if isinstance(value, str) and re.fullmatch(r"\s*[+]?[0-9]+(?:\.[0-9]+)?\s*", value)}
    if not candidates:
        return {}
    column = validate_headers(sheet["headers"])[normalize_header(DATE_COLUMN)]
    overrides = {}
    with ZipFile(path) as archive, archive.open(sheet["part"]) as handle:
        for event, node in ET.iterparse(handle, events=("end",)):
            if node.tag == _NS+"row":
                index = int(node.attrib.get("r", "1"))-2
                if index in candidates:
                    for cell in node.findall(_NS+"c"):
                        if _column_number(cell.attrib["r"]) == column and cell.attrib.get("t", "n") == "n":
                            value = cell.find(_NS+"v")
                            if value is not None and value.text is not None:
                                overrides[index] = parse_date_value(float(value.text))
                node.clear()
    return overrides
