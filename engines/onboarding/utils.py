"""Shared Onboarding helpers. No pandas/openpyxl dependency in the new engine."""
from __future__ import annotations

from datetime import date, datetime, time, timedelta
from functools import lru_cache
from pathlib import Path


class ReportError(Exception):
    """A validation/generation error safe to show in the existing portal."""


def normalize_header(value):
    return "" if value is None else " ".join(str(value).strip().casefold().split())


@lru_cache(maxsize=8192, typed=True)
def parse_date_value(value):
    """Keep the original day-first/ISO/Excel-1900-serial date rules."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and value > 0:
        try:
            days, fraction = divmod(value, 1)
            # Match Excel/openpyxl's fictitious 29-Feb-1900 compatibility rule.
            if 0 <= value < 1:
                return None
            if 0 < value < 60:
                days += 1
            dt = datetime(1899, 12, 30) + timedelta(
                days=days, milliseconds=round(fraction * 86400000)
            )
            return dt.date()
        except (TypeError, ValueError, OverflowError):
            return None
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text or text in {"-", "--", "NA", "N/A"}:
        return None
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        pass
    for fmt in (
        "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%b-%Y", "%d %b %Y",
        "%d-%B-%Y", "%d %B %Y", "%m/%d/%Y", "%d-%m-%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def date_to_excel_datetime(value):
    return datetime.combine(value, time.min)


def inclusive_date_range(first_date, latest_date):
    if first_date > latest_date:
        raise ReportError("The selected From date cannot be after the To date.")
    return [first_date + timedelta(days=n) for n in range((latest_date-first_date).days+1)]


def ordinal_day(day):
    suffix = "th" if 10 <= day % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


def output_filename(report_date=None):
    report_date = report_date or date.today()
    months = ("Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec")
    return f"Onboarding Report - {ordinal_day(report_date.day)} {months[report_date.month-1]}.xlsx"


def clean_path(value):
    return Path(value).expanduser().resolve()
