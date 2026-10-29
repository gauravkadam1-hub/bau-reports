"""Data contracts and layout settings shared by the Polars pipeline and writer."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence
import json
import re

try:
    from .utils import ReportError
except ImportError:  # Existing portal loads report_generator.py by file path.
    from utils import ReportError

BASE_SHEET = "Onboarding Base"
AADHAR_DC, AADHAR_XD = "Aadhar - DC", "Aadhar - XD"
HYPERVERGE_DC, HYPERVERGE_XD = "Hyperverge - DC", "Hyperverge - XD"
MANUAL_HR, RED_HR = "Manual Onb - HR wise details", "RED Cases - HR wise details"
REQUIRED_TEMPLATE_SHEETS = (AADHAR_DC, AADHAR_XD, HYPERVERGE_DC, HYPERVERGE_XD, MANUAL_HR, RED_HR, BASE_SHEET)
DATE_COLUMN = "Onboarded Date"
FACILITY_COLUMN = "Registered Facility"
APPROVER_COLUMN = "Approved By"
AADHAAR_COLUMN = "Onboarded by Aadhaar (Y/N)"
HV_COLUMN = "Hyperverge Status"


def clean_facility(value):
    return "" if value is None else " ".join(str(value).strip().split())


def facility_type(value):
    value = clean_facility(value).casefold()
    return "DC" if value.endswith(" dc") else "XD" if value.endswith(" xd") else None


def clean_selection(values, kind):
    seen, result = set(), []
    for raw in values:
        name = clean_facility(raw)
        key = name.casefold()
        if facility_type(name) == kind and key not in seen:
            result.append(name)
            seen.add(key)
    return result


def profile_path(template_path=None):
    """Retain the portal's old template_path parameter as a settings-directory hint.

    An XLSX file is NOT opened or copied. A caller with a different/custom template
    must provide its corresponding onboarding_layout.json, avoiding silent fallback.
    """
    if template_path is None:
        return Path(__file__).with_name("onboarding_layout.json")
    p = Path(template_path).expanduser().resolve()
    return p if p.suffix.casefold() == ".json" else p.with_name("onboarding_layout.json")


def load_profile(template_path=None):
    p = profile_path(template_path)
    try:
        profile = json.loads(p.read_text(encoding="utf-8"))
        if profile.get("schema_version") != 1:
            raise ValueError("Unsupported layout schema version")
        if profile["sheet_order"] != list(REQUIRED_TEMPLATE_SHEETS):
            raise ValueError("The seven report sheets must be present in the original order")
        if len(profile["headers"]) != len(set(profile["headers"])):
            raise ValueError("Duplicate base headers in layout settings")
        # The preserved COUNTIFS formulas use these original base columns.
        required_positions = {7: FACILITY_COLUMN, 9: DATE_COLUMN, 14: APPROVER_COLUMN,
                              15: AADHAAR_COLUMN, 20: HV_COLUMN}
        for index, expected in required_positions.items():
            if len(profile["headers"]) <= index or profile["headers"][index] != expected:
                raise ValueError(f"Base column {index+1} must remain '{expected}' for the preserved formulas")
        for name in REQUIRED_TEMPLATE_SHEETS:
            if name not in profile["sheets"]:
                raise ValueError(f"Missing layout for {name}")
        defaults_file = p.with_name("onboarding_facility_defaults.json")
        if defaults_file.exists():
            additions = json.loads(defaults_file.read_text(encoding="utf-8"))
            for kind in ("DC", "XD"):
                profile["defaults"][kind] = clean_selection(
                    profile["defaults"][kind] + additions.get(kind, []), kind
                )
        return profile
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise ReportError(f"Could not read Onboarding layout settings at '{p.name}': {exc}") from exc


@dataclass(frozen=True)
class Counts:
    total: int = 0
    aadhar: int = 0
    manual: int = 0
    red: int = 0

    def __add__(self, other):
        return Counts(self.total+other.total, self.aadhar+other.aadhar,
                      self.manual+other.manual, self.red+other.red)

    @property
    def compliance(self):
        # Intentional legacy business rule: no onboarding => 100% compliance.
        return self.aadhar / self.total if self.total else 1.0


@dataclass
class PreparedReport:
    dates: list[date]
    dc_sites: list[str]
    xd_sites: list[str]
    base_rows: Iterable[Sequence[Any]]
    row_count: int
    # Keys are (displayed site, calendar date); cached values follow COUNTIFS.
    counts: dict[tuple[str, date], Counts]
    manual_names: list[str]
    red_names: list[str]
    manual_counts: dict[str, int]
    red_counts: dict[str, int]


def _wildcard_regex(text):
    parts, i = [], 0
    while i < len(text):
        ch = text[i]
        if ch == "~" and i+1 < len(text) and text[i+1] in "*?~":
            i += 1
            parts.append(re.escape(text[i]))
        elif ch == "*":
            parts.append(".*")
        elif ch == "?":
            parts.append(".")
        else:
            parts.append(re.escape(ch))
        i += 1
    return "^" + "".join(parts) + "$"


def excel_text_matches(value, criterion):
    """COUNTIFS text criteria are case-insensitive but do not trim whitespace.

    Preserve the template's wildcard behavior for facility/HR criteria. Do not
    accidentally reuse the more permissive trimming used for sorting/name lists.
    """
    value = "" if value is None else str(value)
    criterion = str(criterion)
    match = re.match(r"^(<>|>=|<=|>|<|=)(.*)$", criterion, re.S)
    op, target = (match.group(1), match.group(2)) if match else ("=", criterion)
    if op in ("=", "<>"):
        equal = bool(re.fullmatch(_wildcard_regex(target), value, re.I | re.S))
        return not equal if op == "<>" else equal
    left, right = value.casefold(), target.casefold()
    return {">": left > right, "<": left < right, ">=": left >= right, "<=": left <= right}[op]
