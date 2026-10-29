from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

from .models import DetectionResult

# Detection is intentionally filename-based because the export structure stays stable
# while the dates/timestamps inside the filenames change each day.
_PATTERNS = {
    "all_users": re.compile(r"^all_user_data_all_.*\.(csv|xlsx|xlsm|xls)$", re.I),
    "attendance": re.compile(r"^attendance_report_.*\.(csv|xlsx|xlsm)$", re.I),
    # Optional Attendance-only floater list. Examples: Floater.csv, Floaters.xlsx, Floater Base.xlsm.
    "floater": re.compile(r"^floaters?(?:[ _-].*)?\.(csv|xlsx|xlsm|xls)$", re.I),
    "roster": re.compile(r"^roster_all_facilities_.*\.(csv|xlsx|xlsm|xls)$", re.I),
    "onboarding": re.compile(r"^onboarding_funnel_all_.*\.(csv|xlsx|xlsm)$", re.I),
    # Real source is commonly flex.xlsx / Flex Base.xlsx. Keep the prefix strict.
    "flex": re.compile(r"^flex(?:[ _-].*|)\.(csv|xlsx|xlsm|xls)$", re.I),
}

# Explicitly ignored per current design discussion.
_IGNORED = {"all user data base.csv"}


def classify_filename(name: str) -> str:
    base = Path(name).name
    if base.casefold() in _IGNORED:
        return "ignored"
    for category in ("roster", "flex", "all_users", "attendance", "floater", "onboarding"):
        if _PATTERNS[category].match(base):
            return category
    return "unknown"


def detect_files(names: Iterable[str]) -> DetectionResult:
    result = DetectionResult()
    for name in names:
        category = classify_filename(name)
        getattr(result, category).append(Path(name).name)

    # Deterministic ordering matters because the two All Users exports are consumed
    # in chronological filename order by the report engines.
    for field_name in (
        "roster", "flex", "all_users", "attendance", "floater", "onboarding", "ignored", "unknown"
    ):
        getattr(result, field_name).sort(key=str.casefold)
    return result


def missing_requirements(result: DetectionResult, report_key: str) -> list[str]:
    missing: list[str] = []
    if report_key == "roster":
        if len(result.roster) != 1:
            missing.append("1 Roster Base")
        if len(result.attendance) != 1:
            missing.append("1 Attendance Base")
    elif report_key == "flex":
        if len(result.flex) != 1:
            missing.append("1 Flex Base")
        if len(result.all_users) != 2:
            missing.append("2 All Users files")
        if len(result.attendance) != 1:
            missing.append("1 Attendance Base")
    elif report_key == "attendance":
        if len(result.all_users) != 2:
            missing.append("2 All Users files")
        if len(result.attendance) != 1:
            missing.append("1 Attendance Base")
    elif report_key == "onboarding":
        if len(result.onboarding) != 1:
            missing.append("1 Onboarding Funnel file")
    return missing
