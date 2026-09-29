from __future__ import annotations

import csv
import html
import json
import os
import re
import shutil
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

VALID_ROSTER_SHIFTS = {"P1", "P2", "P3", "WO"}
SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xlsm", ".xls"}

DEFAULT_DC_SITES = [
    "Ahmedabad DC", "Ajmer DC", "Bangalore DC", "Chennai DC", "Gorakhpur DC",
    "Guwahati DC", "Hyderabad DC", "Hyderabad REV DC", "Jaipur DC", "Kolkata DC",
    "Lucknow DC", "Ludhiana DC", "Mumbai DC", "NCR Bamnoli DC", "NCR Bilaspur DC",
    "NCR TAURU DC", "Nagpur DC", "Noida DC", "ONE NCR DC", "OneRJT DC",
    "OneSRT DC", "Patna DC", "Surat DC", "TwoNCR DC",
]

DEFAULT_XD_SITES = [
    "Akola XD", "Aurangabad XD", "BHUBANESWAR XD", "Bhopal XD", "Chandigarh XD",
    "Coimbatore XD", "Dibrugarh XD", "Hubli XD", "Indore XD", "JABALPUR XD",
    "Jammu XD", "Jodhpur XD", "Kanpur XD", "Kochi XD", "Kozhikode XD",
    "Madurai XD", "Malda XD", "Moradabad XD", "Mysore XD", "Pune XD",
    "Purnia XD", "Raipur XD", "Rajahmundry XD", "Ranchi XD", "SAMBALPUR XD",
    "SILCHAR XD", "Siliguri XD", "Srinagar XD", "Trichy XD", "Udaipur XD",
    "Vijayawada XD", "Visakhapatnam XD", "varanasi XD",
]


class ReportError(Exception):
    pass


@dataclass
class RosterData:
    source: pd.DataFrame
    user_col: str
    name_col: str
    role_col: str
    facility_col: str
    sub_status_col: str
    status_col: str
    onboarded_date_col: str | None
    last_working_day_col: str | None
    date_columns: list[tuple[date, str]]


@dataclass
class AttendanceData:
    raw: pd.DataFrame
    resolved: pd.DataFrame
    date_columns: list[date]


@dataclass
class ScanResult:
    roster_dates: list[date]
    attendance_dates: list[date]
    common_dates: list[date]
    dc_sites: list[str]
    xd_sites: list[str]
    roster_rows: int
    attendance_raw_rows: int

    @property
    def roster_start(self) -> date:
        return min(self.roster_dates)

    @property
    def roster_end(self) -> date:
        return max(self.roster_dates)

    @property
    def attendance_start(self) -> date:
        return min(self.attendance_dates)

    @property
    def attendance_end(self) -> date:
        return max(self.attendance_dates)


@dataclass
class ReportInputs:
    roster_path: str
    attendance_path: str
    output_folder: str
    report_from: date
    report_to: date
    selected_dc: list[str]
    selected_xd: list[str]


@dataclass
class Metrics:
    active: dict[tuple[str, date], int]
    rostered: dict[tuple[str, date], int]
    matched: dict[tuple[str, date], int]


ProgressCallback = Callable[[int, str], None]


def _progress(cb: ProgressCallback | None, pct: int, message: str) -> None:
    if cb:
        cb(pct, message)


def _clean_text_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()


def _norm_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def _find_col(columns: Iterable[object], aliases: Iterable[str], label: str) -> str:
    alias_set = {_norm_header(a) for a in aliases}
    for col in columns:
        if _norm_header(col) in alias_set:
            return str(col)
    raise ReportError(f"Required column not found: {label}")


def _find_optional_col(columns: Iterable[object], aliases: Iterable[str]) -> str | None:
    alias_set = {_norm_header(a) for a in aliases}
    for col in columns:
        if _norm_header(col) in alias_set:
            return str(col)
    return None


def _find_last_working_day_col(columns: Iterable[object]) -> str | None:
    """Locate the roster lifecycle end-date column without depending on its position.

    Export layouts occasionally rename the field slightly (for example `LWD` or
    `Last Working Day (LWD)`). Keep this matching deliberately narrow so unrelated
    attendance fields such as `last_present_day` are never treated as employment
    end dates.
    """
    direct = _find_optional_col(
        columns,
        [
            "Last Working Day", "last_working_day",
            "Last Working Date", "last_working_date",
            "LWD", "Last Work Day", "Last Work Date",
            "Employment End Date", "Exit Date", "Offboarding Date",
        ],
    )
    if direct:
        return direct

    for col in columns:
        raw = str(col).strip().lower()
        norm = _norm_header(col)
        # Handle labels such as "Last Working Day (LWD)" / "LWD Date".
        if norm.startswith("lwd") or norm.endswith("lwd"):
            return str(col)
        if "last" in raw and "work" in raw and ("day" in raw or "date" in raw):
            return str(col)
    return None


def _find_onboarded_date_col(columns: Iterable[object]) -> str | None:
    """Locate an onboarding/joining date field without depending on column position."""
    direct = _find_optional_col(
        columns,
        [
            "Onboarded Date", "Onboarding Date", "Onboard Date",
            "Joining Date", "Date of Joining", "DOJ",
            "Employment Start Date", "Start Date", "Activated Date",
        ],
    )
    if direct:
        return direct
    for col in columns:
        raw = str(col).strip().lower()
        norm = _norm_header(col)
        if norm in {"doj", "dateofjoining"}:
            return str(col)
        if ("onboard" in raw or "join" in raw or "employment start" in raw) and "date" in raw:
            return str(col)
    return None


def _parse_date_header(value: object) -> date | None:
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s or s.lower().startswith("unnamed"):
        return None
    formats = (
        "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d",
        "%d-%b-%Y", "%d %b %Y", "%d-%B-%Y", "%d %B %Y",
    )
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _read_flat(path: str) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ReportError(f"Unsupported file type: {ext}. Use CSV or Excel.")
    try:
        if ext == ".csv":
            return pd.read_csv(path, dtype=str, keep_default_na=False, low_memory=False, encoding="utf-8-sig")
        return pd.read_excel(path, dtype=str, keep_default_na=False)
    except Exception as exc:
        raise ReportError(f"Could not read {Path(path).name}: {exc}") from exc


def load_roster(path: str) -> RosterData:
    df = _read_flat(path)
    if df.empty:
        raise ReportError("Roster file is empty.")
    df.columns = [str(c).strip() for c in df.columns]

    user_col = _find_col(df.columns, ["User ID", "user_id", "Employee Code", "Employee ID"], "User ID")
    name_col = _find_col(df.columns, ["Name", "user_name", "Employee Name"], "Name")
    role_col = _find_col(df.columns, ["Role"], "Role")
    facility_col = _find_col(df.columns, ["Facility Name", "facility_name", "Facility"], "Facility Name")
    sub_status_col = _find_col(df.columns, ["Sub Status", "sub_status"], "Sub Status")
    status_col = _find_col(df.columns, ["Status"], "Status")
    onboarded_date_col = _find_onboarded_date_col(df.columns)
    last_working_day_col = _find_last_working_day_col(df.columns)

    date_columns: list[tuple[date, str]] = []
    for col in df.columns:
        parsed = _parse_date_header(col)
        if parsed:
            date_columns.append((parsed, col))
    date_columns.sort(key=lambda x: x[0])
    if not date_columns:
        raise ReportError("No roster date columns were detected.")
    if len({d for d, _ in date_columns}) != len(date_columns):
        raise ReportError("Roster contains duplicate date columns.")

    for col in df.columns:
        df[col] = _clean_text_series(df[col])

    ids = df[user_col]
    dupes = ids[ids.ne("") & ids.duplicated(keep=False)]
    if not dupes.empty:
        raise ReportError(
            f"Roster User ID must be unique. Found {dupes.nunique():,} duplicated employee IDs."
        )

    return RosterData(
        source=df,
        user_col=user_col,
        name_col=name_col,
        role_col=role_col,
        facility_col=facility_col,
        sub_status_col=sub_status_col,
        status_col=status_col,
        onboarded_date_col=onboarded_date_col,
        last_working_day_col=last_working_day_col,
        date_columns=date_columns,
    )


def _attendance_is_two_row(path: str) -> bool:
    ext = Path(path).suffix.lower()
    try:
        if ext == ".csv":
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.reader(f)
                first = next(reader, [])
                second = next(reader, [])
            return sum(str(x).strip().lower() == "shift" for x in second) >= 1 and any(
                _parse_date_header(x) for x in first
            )
        preview = pd.read_excel(path, header=None, nrows=2, dtype=str, keep_default_na=False)
        if preview.shape[0] < 2:
            return False
        first = preview.iloc[0].tolist()
        second = preview.iloc[1].tolist()
        return sum(str(x).strip().lower() == "shift" for x in second) >= 1 and any(
            _parse_date_header(x) for x in first
        )
    except Exception:
        return False


def _load_attendance_raw(path: str) -> tuple[pd.DataFrame, list[date]]:
    ext = Path(path).suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ReportError(f"Unsupported file type: {ext}. Use CSV or Excel.")

    two_row = _attendance_is_two_row(path)
    try:
        if two_row:
            if ext == ".csv":
                raw = pd.read_csv(path, header=[0, 1], dtype=str, keep_default_na=False, low_memory=False, encoding="utf-8-sig")
            else:
                raw = pd.read_excel(path, header=[0, 1], dtype=str, keep_default_na=False)

            cols = list(raw.columns)
            if len(cols) < 4:
                raise ReportError("Attendance file does not contain the expected employee columns.")

            out = pd.DataFrame({
                "Employee Code": _clean_text_series(raw.iloc[:, 0]),
                "Name": _clean_text_series(raw.iloc[:, 1]),
                "Role": _clean_text_series(raw.iloc[:, 2]),
                "Facility Name": _clean_text_series(raw.iloc[:, 3]),
            })
            dates: list[date] = []
            for idx, col in enumerate(cols[4:], start=4):
                top, sub = col if isinstance(col, tuple) else (col, "")
                d = _parse_date_header(top)
                if d and str(sub).strip().lower() == "shift":
                    if d in dates:
                        raise ReportError(f"Attendance contains duplicate shift columns for {d:%d-%b-%Y}.")
                    dates.append(d)
                    out[d] = _clean_text_series(raw.iloc[:, idx])
            if not dates:
                raise ReportError("No attendance 'shift' date columns were detected.")
            return out, dates

        flat = _read_flat(path)
        flat.columns = [str(c).strip() for c in flat.columns]
        user_col = _find_col(flat.columns, ["User ID", "user_id", "Employee Code", "Employee ID"], "Employee Code")
        name_col = _find_col(flat.columns, ["Name", "user_name", "Employee Name"], "Name")
        role_col = _find_col(flat.columns, ["Role"], "Role")
        facility_col = _find_col(flat.columns, ["Facility Name", "facility_name", "Facility"], "Facility Name")
        out = pd.DataFrame({
            "Employee Code": _clean_text_series(flat[user_col]),
            "Name": _clean_text_series(flat[name_col]),
            "Role": _clean_text_series(flat[role_col]),
            "Facility Name": _clean_text_series(flat[facility_col]),
        })
        dates: list[date] = []
        for col in flat.columns:
            d = _parse_date_header(col)
            if d:
                dates.append(d)
                out[d] = _clean_text_series(flat[col])
        dates = sorted(set(dates))
        if not dates:
            raise ReportError("No attendance date columns were detected.")
        return out, dates
    except ReportError:
        raise
    except Exception as exc:
        raise ReportError(f"Could not parse attendance file: {exc}") from exc


def resolve_attendance(raw: pd.DataFrame, roster: RosterData) -> pd.DataFrame:
    a = raw.copy()
    a["_row_order"] = range(len(a))

    roster_facility = roster.source[[roster.user_col, roster.facility_col]].copy()
    roster_facility.columns = ["Employee Code", "_roster_facility"]
    roster_facility["Employee Code"] = _clean_text_series(roster_facility["Employee Code"])
    roster_facility["_roster_facility"] = _clean_text_series(roster_facility["_roster_facility"])

    a = a.merge(roster_facility, on="Employee Code", how="left", sort=False)
    a["_facility_match"] = (
        a["Facility Name"].str.casefold() == a["_roster_facility"].fillna("").str.casefold()
    ).astype(int)
    a = a.sort_values(
        ["Employee Code", "_facility_match", "_row_order"],
        ascending=[True, False, True],
        kind="mergesort",
    )
    a = a.drop_duplicates(subset=["Employee Code"], keep="first")
    a = a.sort_values("_row_order", kind="mergesort")
    return a.drop(columns=["_row_order", "_roster_facility", "_facility_match"]).reset_index(drop=True)


def load_attendance(path: str, roster: RosterData) -> AttendanceData:
    raw, dates = _load_attendance_raw(path)
    resolved = resolve_attendance(raw, roster)
    return AttendanceData(raw=raw, resolved=resolved, date_columns=sorted(dates))


def _is_dc(site: str) -> bool:
    """Classify only facilities whose trimmed name ends exactly with the suffix " DC"."""
    return str(site).rstrip().casefold().endswith(" dc")


def _is_xd(site: str) -> bool:
    """Classify only facilities whose trimmed name ends exactly with the suffix " XD"."""
    return str(site).rstrip().casefold().endswith(" xd")


def scan_inputs(roster_path: str, attendance_path: str) -> ScanResult:
    roster = load_roster(roster_path)
    raw_att, attendance_dates = _load_attendance_raw(attendance_path)

    roster_sites = {x for x in _clean_text_series(roster.source[roster.facility_col]).tolist() if x}
    attendance_sites = {x for x in _clean_text_series(raw_att["Facility Name"]).tolist() if x}
    roster_by_key = {s.casefold(): s for s in roster_sites}
    attendance_keys = {s.casefold() for s in attendance_sites}
    intersection = sorted(
        [roster_by_key[k] for k in roster_by_key.keys() & attendance_keys],
        key=str.casefold,
    )

    dc_sites = [s for s in intersection if _is_dc(s)]
    xd_sites = [s for s in intersection if _is_xd(s)]
    roster_dates = [d for d, _ in roster.date_columns]
    common_dates = sorted(set(roster_dates) & set(attendance_dates))
    if not common_dates:
        raise ReportError("Roster and attendance files have no common dates.")

    return ScanResult(
        roster_dates=roster_dates,
        attendance_dates=sorted(attendance_dates),
        common_dates=common_dates,
        dc_sites=dc_sites,
        xd_sites=xd_sites,
        roster_rows=len(roster.source),
        attendance_raw_rows=len(raw_att),
    )


def ordinal_day(day: int) -> str:
    if 10 <= day % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


def default_output_filename(today: date | None = None) -> str:
    today = today or date.today()
    return f"Roster Report - {ordinal_day(today.day)} {today:%b}.xlsx"


def _daterange(start: date, end: date) -> list[date]:
    if end < start:
        raise ReportError("Report end date cannot be before start date.")
    days = (end - start).days
    return [start + timedelta(days=i) for i in range(days + 1)]


def _parse_date_series(values: pd.Series) -> pd.Series:
    """Parse mixed date text safely, including Excel serial dates when present."""
    text = _clean_text_series(values)
    try:
        parsed = pd.to_datetime(text.where(text.ne("")), errors="coerce", format="mixed", dayfirst=True)
    except TypeError:
        parsed = pd.to_datetime(text.where(text.ne("")), errors="coerce", dayfirst=True)

    missing = parsed.isna() & text.ne("")
    if missing.any():
        numeric = pd.to_numeric(text.where(missing), errors="coerce")
        serial_mask = numeric.between(20000, 80000, inclusive="both")
        if serial_mask.any():
            parsed.loc[serial_mask] = pd.to_datetime(
                numeric.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce"
            )
    return parsed


def _active_mask_for_date(roster: RosterData, report_date: date) -> pd.Series:
    """Return employees considered active on a specific historical report date.

    If an onboarding/joining date exists, employees are excluded before that date.
    If a Last Working Day exists, employees who are non-active today are restored
    for historical dates on/before their last working day. Blank lifecycle dates do
    not exclude otherwise eligible current-active employees.
    """
    sub_status = _clean_text_series(roster.source[roster.sub_status_col]).str.upper()
    current_active = sub_status.eq("ACTIVE")

    if roster.onboarded_date_col:
        start_dates = _parse_date_series(roster.source[roster.onboarded_date_col])
        started = start_dates.isna() | (start_dates <= pd.Timestamp(report_date))
    else:
        started = pd.Series(True, index=roster.source.index)

    current_active = current_active & started
    if not roster.last_working_day_col:
        return current_active

    lwd = _parse_date_series(roster.source[roster.last_working_day_col])
    historical_non_active = (
        sub_status.ne("ACTIVE") & started & lwd.notna() & (lwd >= pd.Timestamp(report_date))
    )
    return current_active | historical_non_active


def _validate_report_inputs(inputs: ReportInputs, scan: ScanResult) -> list[date]:
    if inputs.report_to >= date.today():
        raise ReportError(
            f"Report end date must be D-1 or earlier. Today is {date.today():%d-%b-%Y}."
        )
    report_dates = _daterange(inputs.report_from, inputs.report_to)
    common = set(scan.common_dates)
    missing = [d for d in report_dates if d not in common]
    if missing:
        missing_text = ", ".join(d.strftime("%d-%b-%Y") for d in missing[:8])
        if len(missing) > 8:
            missing_text += ", ..."
        raise ReportError(
            "Every report date must exist in both roster and attendance. Missing: " + missing_text
        )
    available_dc = set(scan.dc_sites)
    available_xd = set(scan.xd_sites)
    bad_sites = [s for s in inputs.selected_dc if s not in available_dc] + [
        s for s in inputs.selected_xd if s not in available_xd
    ]
    if bad_sites:
        raise ReportError("Selected sites are no longer present in both inputs: " + ", ".join(bad_sites))
    if not inputs.selected_dc and not inputs.selected_xd:
        raise ReportError("Select at least one DC or XD site.")
    return report_dates


def _compute_master_arrays(
    roster: RosterData,
    attendance: AttendanceData,
) -> tuple[dict[date, pd.Series], dict[date, pd.Series], dict[date, pd.Series]]:
    ids = roster.source[roster.user_col].astype(str)
    att = attendance.resolved.set_index("Employee Code", drop=False)

    att_by_date: dict[date, pd.Series] = {}
    roster_by_date: dict[date, pd.Series] = {}
    validation_by_date: dict[date, pd.Series] = {}
    attendance_dates = set(attendance.date_columns)

    for d, roster_col in roster.date_columns:
        rshift = _clean_text_series(roster.source[roster_col]).str.upper()
        roster_by_date[d] = rshift
        if d in attendance_dates:
            if d in att.columns:
                ashift = ids.map(att[d].to_dict()).fillna("").astype(str).str.strip().str.upper()
            else:
                ashift = pd.Series([""] * len(ids), index=ids.index, dtype=str)
            att_by_date[d] = ashift
            match = ((ashift.ne("") & ashift.eq(rshift)) | (rshift.eq("WO") & ashift.eq("-")))
            validation_by_date[d] = pd.Series(
                ["Match" if x else "No Match" for x in match], index=roster.source.index, dtype=str
            )
        else:
            att_by_date[d] = pd.Series([""] * len(ids), index=roster.source.index, dtype=str)
            validation_by_date[d] = pd.Series([""] * len(ids), index=roster.source.index, dtype=str)
    return att_by_date, roster_by_date, validation_by_date


def _compute_metrics(
    roster: RosterData,
    report_dates: list[date],
    selected_sites: list[str],
    roster_by_date: dict[date, pd.Series],
    validation_by_date: dict[date, pd.Series],
) -> Metrics:
    facilities = _clean_text_series(roster.source[roster.facility_col])
    current_active_mask = _clean_text_series(roster.source[roster.sub_status_col]).str.upper().eq("ACTIVE")
    day_active_masks: dict[date, pd.Series] = {d: _active_mask_for_date(roster, d) for d in report_dates}

    active: dict[tuple[str, date], int] = {}
    rostered: dict[tuple[str, date], int] = {}
    matched: dict[tuple[str, date], int] = {}

    for site in selected_sites:
        site_mask = facilities.eq(site)
        for d in report_dates:
            day_active_mask = day_active_masks[d]
            site_day_active = site_mask & day_active_mask
            active[(site, d)] = int(site_day_active.sum())

            # Rostered/Matched retain the established report logic based on current
            # Sub Status = ACTIVE. Only Active HC is being changed to a historical,
            # date-aware headcount in this patch.
            site_current_active = site_mask & current_active_mask
            rmask = roster_by_date[d].isin(VALID_ROSTER_SHIFTS)
            rostered[(site, d)] = int((site_current_active & rmask).sum())
            matched[(site, d)] = int((site_current_active & validation_by_date[d].eq("Match")).sum())
    return Metrics(active=active, rostered=rostered, matched=matched)


def _pct(num: int, den: int) -> float:
    if not den:
        return 0.0
    return min(100.0, num / den * 100.0)


def _write_value(ws, row: int, col: int, value: object, fmt=None) -> None:
    if value is None or value == "":
        ws.write_blank(row, col, None, fmt)
    elif isinstance(value, (pd.Timestamp, datetime)):
        ws.write_datetime(row, col, value.to_pydatetime() if isinstance(value, pd.Timestamp) else value, fmt)
    elif isinstance(value, date):
        ws.write_datetime(row, col, datetime.combine(value, datetime.min.time()), fmt)
    else:
        ws.write(row, col, value, fmt)


def _formats(workbook: xlsxwriter.Workbook) -> dict[str, object]:
    """Formats matched to the original 8-Sep reference workbook."""
    light_blue = "#DCE6F1"
    summary_total = "#D9D9D9"
    metric_total = "#BFBFBF"

    return {
        # Summary sheets (Calibri 10, light-blue two-row header, grey total row).
        "header": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": light_blue,
            "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True,
        }),
        "date_header": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": light_blue,
            "align": "center", "valign": "vcenter", "border": 1,
            "text_wrap": True, "num_format": "dd-mm-yyyy",
        }),
        "subheader": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": light_blue,
            "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True,
        }),
        "summary_site": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "font_color": "#000000",
            "border": 1, "align": "left", "valign": "vcenter", "text_wrap": True,
        }),
        "formula": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "font_color": "#000000",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        }),
        "percent": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "font_color": "#000000",
            "border": 1, "align": "center", "valign": "vcenter",
            "text_wrap": True, "num_format": '0"%"',
        }),
        "grand": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": summary_total,
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        }),
        "grand_site": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": summary_total,
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        }),
        "grand_pct": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": summary_total,
            "border": 1, "align": "center", "valign": "vcenter",
            "text_wrap": True, "num_format": '0"%"',
        }),

        # DC Adoption/Adherence (reference workbook uses Calibri 10).
        "dc_metric_header": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": light_blue,
            "border": 1, "align": "center", "valign": "vcenter",
        }),
        "dc_metric_date": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": light_blue,
            "border": 1, "align": "center", "valign": "vcenter", "num_format": "dd-mm-yyyy",
        }),
        "dc_metric_body": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "font_color": "#000000",
            "border": 1, "align": "center", "valign": "vcenter",
        }),
        "dc_metric_pct": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "font_color": "#000000",
            "border": 1, "align": "center", "valign": "vcenter", "num_format": '0"%"',
        }),
        "dc_metric_grand": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": metric_total,
            "border": 1, "align": "center", "valign": "vcenter",
        }),
        "dc_metric_grand_pct": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": metric_total,
            "border": 1, "align": "center", "valign": "vcenter", "num_format": '0"%"',
        }),

        # XD Adoption/Adherence (reference workbook uses Calibri 11).
        "xd_metric_header": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "font_color": "#000000", "bg_color": light_blue,
            "border": 1, "align": "center", "valign": "vcenter",
        }),
        "xd_metric_site_header": workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "font_color": "#000000", "bg_color": light_blue,
            "border": 1, "align": "center", "valign": "vcenter",
        }),
        "xd_metric_date": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "font_color": "#000000", "bg_color": light_blue,
            "border": 1, "align": "center", "valign": "vcenter", "num_format": "dd-mm-yyyy",
        }),
        "xd_metric_site": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000", "border": 1,
        }),
        "xd_metric_body": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000",
            "border": 1, "align": "center", "valign": "vcenter",
        }),
        "xd_metric_pct": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000",
            "border": 1, "align": "center", "valign": "vcenter", "num_format": '0"%"',
        }),
        "xd_metric_grand_site": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "font_color": "#000000", "bg_color": metric_total, "border": 1,
        }),
        "xd_metric_grand": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "font_color": "#000000", "bg_color": metric_total,
            "border": 1, "align": "center", "valign": "vcenter",
        }),
        "xd_metric_grand_pct": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "font_color": "#000000", "bg_color": metric_total,
            "border": 1, "align": "center", "valign": "vcenter", "num_format": '0"%"',
        }),

        # Roster Base / Attendance Base: plain worksheet grid, bold black headers, no fill.
        "base_header": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True, "font_color": "#000000",
        }),
        "base_date_header": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "font_color": "#000000", "num_format": "dd-mm-yyyy",
        }),
        "base_body": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000",
        }),
        "base_body_center": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000",
        }),
        "base_date_body": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000",
            "num_format": "dd-mm-yyyy",
        }),

        # Master Base: thin cell borders, plain black text, no header fills.
        "master_header": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "font_color": "#000000", "border": 1,
        }),
        "master_date": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "font_color": "#000000", "border": 1, "num_format": "dd-mm-yyyy",
        }),
        "master_sub": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000", "border": 1,
        }),
        "master_formula": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000", "border": 1,
        }),
        "master_imported": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000", "border": 1,
        }),

        # Methodology sheet: plain text like the reference workbook.
        "note_title": workbook.add_format({
            "font_name": "Cambria", "font_size": 12, "bold": True, "font_color": "#000000",
        }),
        "note_section": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000",
            "text_wrap": True, "valign": "top",
        }),
        "note": workbook.add_format({
            "font_name": "Calibri", "font_size": 11, "font_color": "#000000",
            "text_wrap": True, "valign": "top",
        }),
    }


def _configure_report_sheet(ws, total_cols: int) -> None:
    # The reference workbook keeps normal Excel gridlines visible and does not freeze report sheets.
    ws.set_column(0, 0, 16)
    if total_cols > 1:
        ws.set_column(1, total_cols - 1, 10)


def _active_hc_formula(
    excel_row: int,
    date_header_ref: str,
    roster_fac_col: int,
    roster_sub_col: int,
    roster_onboarded_col: int | None,
    roster_lwd_col: int | None,
) -> str:
    """Build the visible Excel formula for date-wise Active HC.

    Lifecycle columns are located by header name, so their physical positions may
    change in future roster exports without breaking the formula generation.
    """
    fac = xl_col_to_name(roster_fac_col)
    sub = xl_col_to_name(roster_sub_col)
    parts: list[str] = []

    if roster_onboarded_col is None:
        parts.append(
            f"COUNTIFS('Roster Base'!${sub}:${sub},\"ACTIVE\",'Roster Base'!${fac}:${fac},$A{excel_row})"
        )
    else:
        join = xl_col_to_name(roster_onboarded_col)
        parts.append(
            f"COUNTIFS('Roster Base'!${sub}:${sub},\"ACTIVE\",'Roster Base'!${fac}:${fac},$A{excel_row},'Roster Base'!${join}:${join},\"<=\"&{date_header_ref})"
        )
        parts.append(
            f"COUNTIFS('Roster Base'!${sub}:${sub},\"ACTIVE\",'Roster Base'!${fac}:${fac},$A{excel_row},'Roster Base'!${join}:${join},\"\")"
        )

    if roster_lwd_col is not None:
        lwd = xl_col_to_name(roster_lwd_col)
        if roster_onboarded_col is None:
            parts.append(
                f"COUNTIFS('Roster Base'!${sub}:${sub},\"<>ACTIVE\",'Roster Base'!${fac}:${fac},$A{excel_row},'Roster Base'!${lwd}:${lwd},\">=\"&{date_header_ref})"
            )
        else:
            join = xl_col_to_name(roster_onboarded_col)
            parts.append(
                f"COUNTIFS('Roster Base'!${sub}:${sub},\"<>ACTIVE\",'Roster Base'!${fac}:${fac},$A{excel_row},'Roster Base'!${lwd}:${lwd},\">=\"&{date_header_ref},'Roster Base'!${join}:${join},\"<=\"&{date_header_ref})"
            )
            parts.append(
                f"COUNTIFS('Roster Base'!${sub}:${sub},\"<>ACTIVE\",'Roster Base'!${fac}:${fac},$A{excel_row},'Roster Base'!${lwd}:${lwd},\">=\"&{date_header_ref},'Roster Base'!${join}:${join},\"\")"
            )

    return "=" + "+".join(parts)


def _sort_sites_by_latest_adherence(
    sites: list[str], metrics: Metrics, report_dates: list[date]
) -> list[str]:
    if not sites or not report_dates:
        return list(sites)
    latest = report_dates[-1]
    return sorted(
        sites,
        key=lambda site: (
            _pct(metrics.matched[(site, latest)], metrics.rostered[(site, latest)]),
            site.casefold(),
        ),
    )


def _write_summary_sheet(
    workbook, name: str, sites: list[str], report_dates: list[date], metrics: Metrics,
    roster_cols: dict[date, int], roster_fac_col: int, roster_sub_col: int,
    roster_onboarded_col: int | None, roster_lwd_col: int | None,
    master_validation_cols: dict[date, int], master_sub_col: int, fmts: dict[str, object]
) -> None:
    ws = workbook.get_worksheet_by_name(name) or workbook.add_worksheet(name)
    total_cols = 1 + 5 * len(report_dates)
    _configure_report_sheet(ws, total_cols)
    # Compact widths and wrapped labels mirror the original Summary tabs.
    ws.set_column(0, 0, 15.6 if name.startswith("DC") else 16.0)
    width_pattern = [7.7, 9.5, 10.9, 9.5, 11.2]
    # Keep the entire selected date range in Excel, but show only the final two
    # report-date blocks by default (D-2 and D-1 for the normal daily run).
    visible_dates = set(report_dates[-2:])
    for i, d in enumerate(report_dates):
        c = 1 + i * 5
        options = {"hidden": True} if d not in visible_dates else None
        for j, width in enumerate(width_pattern):
            ws.set_column(c + j, c + j, width, None, options)
    ws.set_row(1, 29)
    labels = ["Active HC", "Rostered HC", "Roster Adoption", "Roster Match", "Roster Adherence"]
    # Horizontal date blocks and the Site header are merged exactly like the reference workbook.
    for i, d in enumerate(report_dates):
        c = 1 + i * 5
        ws.merge_range(0, c, 0, c + 4, datetime.combine(d, datetime.min.time()), fmts["date_header"])
    ws.merge_range(0, 0, 1, 0, "Site", fmts["header"])
    for i, d in enumerate(report_dates):
        c = 1 + i * 5
        for j, label in enumerate(labels):
            ws.write(1, c + j, label, fmts["subheader"])

    first_data_row_excel = 3
    for r0, site in enumerate(sites, start=2):
        excel_row = r0 + 1
        ws.write(r0, 0, site, fmts["summary_site"])
        for i, d in enumerate(report_dates):
            c = 1 + i * 5
            roster_date_letter = xl_col_to_name(roster_cols[d])
            roster_fac_letter = xl_col_to_name(roster_fac_col)
            roster_sub_letter = xl_col_to_name(roster_sub_col)
            master_val_letter = xl_col_to_name(master_validation_cols[d])
            master_sub_letter = xl_col_to_name(master_sub_col)

            active_val = metrics.active[(site, d)]
            rostered_val = metrics.rostered[(site, d)]
            match_val = metrics.matched[(site, d)]
            adoption_val = _pct(rostered_val, active_val)
            adherence_val = _pct(match_val, rostered_val)

            f_active = _active_hc_formula(
                excel_row,
                f"{xl_col_to_name(c)}$1",
                roster_fac_col,
                roster_sub_col,
                roster_onboarded_col,
                roster_lwd_col,
            )
            f_rostered = (
                f'=COUNTIFS(\'Roster Base\'!${roster_fac_letter}:${roster_fac_letter},$A{excel_row},\'Roster Base\'!${roster_sub_letter}:${roster_sub_letter},"ACTIVE",\'Roster Base\'!${roster_date_letter}:${roster_date_letter},"P1")+'
                f'COUNTIFS(\'Roster Base\'!${roster_fac_letter}:${roster_fac_letter},$A{excel_row},\'Roster Base\'!${roster_sub_letter}:${roster_sub_letter},"ACTIVE",\'Roster Base\'!${roster_date_letter}:${roster_date_letter},"P2")+'
                f'COUNTIFS(\'Roster Base\'!${roster_fac_letter}:${roster_fac_letter},$A{excel_row},\'Roster Base\'!${roster_sub_letter}:${roster_sub_letter},"ACTIVE",\'Roster Base\'!${roster_date_letter}:${roster_date_letter},"P3")+'
                f'COUNTIFS(\'Roster Base\'!${roster_fac_letter}:${roster_fac_letter},$A{excel_row},\'Roster Base\'!${roster_sub_letter}:${roster_sub_letter},"ACTIVE",\'Roster Base\'!${roster_date_letter}:${roster_date_letter},"WO")'
            )
            f_adoption = f'=IFERROR(MIN(100,{xl_col_to_name(c+1)}{excel_row}/{xl_col_to_name(c)}{excel_row}*100),0)'
            f_match = f'=COUNTIFS(\'Master Base\'!$D:$D,$A{excel_row},\'Master Base\'!${master_sub_letter}:${master_sub_letter},"ACTIVE",\'Master Base\'!${master_val_letter}:${master_val_letter},"Match")'
            f_adherence = f'=IFERROR(MIN(100,{xl_col_to_name(c+3)}{excel_row}/{xl_col_to_name(c+1)}{excel_row}*100),0)'

            ws.write_formula(r0, c, f_active, fmts["formula"], active_val)
            ws.write_formula(r0, c + 1, f_rostered, fmts["formula"], rostered_val)
            ws.write_formula(r0, c + 2, f_adoption, fmts["percent"], adoption_val)
            ws.write_formula(r0, c + 3, f_match, fmts["formula"], match_val)
            ws.write_formula(r0, c + 4, f_adherence, fmts["percent"], adherence_val)

    grand_r0 = 2 + len(sites)
    grand_excel = grand_r0 + 1
    ws.write(grand_r0, 0, "Grand Total", fmts["grand_site"])
    last_site_excel = 2 + len(sites)
    for i, d in enumerate(report_dates):
        c = 1 + i * 5
        vals = [metrics.active[(s, d)] for s in sites]
        rosters = [metrics.rostered[(s, d)] for s in sites]
        matches = [metrics.matched[(s, d)] for s in sites]
        av, rv, mv = sum(vals), sum(rosters), sum(matches)
        ap, ad = _pct(rv, av), _pct(mv, rv)
        for offset, cached in [(0, av), (1, rv), (3, mv)]:
            letter = xl_col_to_name(c + offset)
            ws.write_formula(grand_r0, c + offset, f'=SUM({letter}{first_data_row_excel}:{letter}{last_site_excel})', fmts["grand"], cached)
        ws.write_formula(grand_r0, c + 2, f'=IFERROR(MIN(100,{xl_col_to_name(c+1)}{grand_excel}/{xl_col_to_name(c)}{grand_excel}*100),0)', fmts["grand_pct"], ap)
        ws.write_formula(grand_r0, c + 4, f'=IFERROR(MIN(100,{xl_col_to_name(c+3)}{grand_excel}/{xl_col_to_name(c+1)}{grand_excel}*100),0)', fmts["grand_pct"], ad)
        ws.conditional_format(2, c + 2, grand_r0 - 1, c + 2, {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})
        ws.conditional_format(2, c + 4, grand_r0 - 1, c + 4, {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})


def _write_three_metric_sheet(
    workbook, name: str, sites: list[str], report_dates: list[date], metrics: Metrics,
    mode: str, roster_cols: dict[date, int], roster_fac_col: int, roster_sub_col: int,
    roster_onboarded_col: int | None, roster_lwd_col: int | None,
    master_validation_cols: dict[date, int], master_sub_col: int, fmts: dict[str, object]
) -> None:
    ws = workbook.get_worksheet_by_name(name) or workbook.add_worksheet(name)
    total_cols = 1 + 3 * len(report_dates)
    _configure_report_sheet(ws, total_cols)
    is_xd = name.startswith("XD")
    if mode == "adoption":
        labels = ["Active HC", "Rostered HC", "Roster Adoption %"]
        widths = [9.0, 11.22, 16.66] if is_xd else [8.55, 10.89, 15.78]
    else:
        labels = ["Rostered HC", "Matched Count", "Roster Adherence %"]
        widths = [11.22, 14.22, 18.0] if is_xd else [10.89, 13.22, 17.22]
    ws.set_column(0, 0, 16.89 if is_xd else 15.55)
    for i in range(len(report_dates)):
        c = 1 + i * 3
        for j, width in enumerate(widths):
            ws.set_column(c + j, c + j, width)

    header_fmt = fmts["xd_metric_header"] if is_xd else fmts["dc_metric_header"]
    site_header_fmt = fmts["xd_metric_site_header"] if is_xd else fmts["dc_metric_header"]
    date_fmt = fmts["xd_metric_date"] if is_xd else fmts["dc_metric_date"]
    for i, d in enumerate(report_dates):
        c = 1 + i * 3
        ws.merge_range(0, c, 0, c + 2, datetime.combine(d, datetime.min.time()), date_fmt)
    ws.merge_range(0, 0, 1, 0, "Site", site_header_fmt)
    for i, d in enumerate(report_dates):
        c = 1 + i * 3
        for j, label in enumerate(labels):
            ws.write(1, c + j, label, header_fmt)

    body_fmt = fmts["xd_metric_body"] if is_xd else fmts["dc_metric_body"]
    pct_fmt = fmts["xd_metric_pct"] if is_xd else fmts["dc_metric_pct"]
    site_fmt = fmts["xd_metric_site"] if is_xd else fmts["dc_metric_body"]
    grand_site_fmt = fmts["xd_metric_grand_site"] if is_xd else fmts["dc_metric_grand"]
    grand_fmt = fmts["xd_metric_grand"] if is_xd else fmts["dc_metric_grand"]
    grand_pct_fmt = fmts["xd_metric_grand_pct"] if is_xd else fmts["dc_metric_grand_pct"]

    first_excel = 3
    for r0, site in enumerate(sites, start=2):
        excel_row = r0 + 1
        ws.write(r0, 0, site, site_fmt)
        for i, d in enumerate(report_dates):
            c = 1 + i * 3
            roster_date_letter = xl_col_to_name(roster_cols[d])
            roster_fac_letter = xl_col_to_name(roster_fac_col)
            roster_sub_letter = xl_col_to_name(roster_sub_col)
            master_val_letter = xl_col_to_name(master_validation_cols[d])
            master_sub_letter = xl_col_to_name(master_sub_col)

            av = metrics.active[(site, d)]
            rv = metrics.rostered[(site, d)]
            mv = metrics.matched[(site, d)]
            f_rostered = (
                f'=COUNTIFS(\'Roster Base\'!${roster_fac_letter}:${roster_fac_letter},$A{excel_row},\'Roster Base\'!${roster_sub_letter}:${roster_sub_letter},"ACTIVE",\'Roster Base\'!${roster_date_letter}:${roster_date_letter},"P1")+'
                f'COUNTIFS(\'Roster Base\'!${roster_fac_letter}:${roster_fac_letter},$A{excel_row},\'Roster Base\'!${roster_sub_letter}:${roster_sub_letter},"ACTIVE",\'Roster Base\'!${roster_date_letter}:${roster_date_letter},"P2")+'
                f'COUNTIFS(\'Roster Base\'!${roster_fac_letter}:${roster_fac_letter},$A{excel_row},\'Roster Base\'!${roster_sub_letter}:${roster_sub_letter},"ACTIVE",\'Roster Base\'!${roster_date_letter}:${roster_date_letter},"P3")+'
                f'COUNTIFS(\'Roster Base\'!${roster_fac_letter}:${roster_fac_letter},$A{excel_row},\'Roster Base\'!${roster_sub_letter}:${roster_sub_letter},"ACTIVE",\'Roster Base\'!${roster_date_letter}:${roster_date_letter},"WO")'
            )
            if mode == "adoption":
                f_active = _active_hc_formula(
                    excel_row,
                    f"{xl_col_to_name(c)}$1",
                    roster_fac_col,
                    roster_sub_col,
                    roster_onboarded_col,
                    roster_lwd_col,
                )
                ws.write_formula(r0, c, f_active, body_fmt, av)
                ws.write_formula(r0, c + 1, f_rostered, body_fmt, rv)
                ws.write_formula(r0, c + 2, f'=IFERROR(MIN(100,{xl_col_to_name(c+1)}{excel_row}/{xl_col_to_name(c)}{excel_row}*100),0)', pct_fmt, _pct(rv, av))
            else:
                f_match = f'=COUNTIFS(\'Master Base\'!$D:$D,$A{excel_row},\'Master Base\'!${master_sub_letter}:${master_sub_letter},"ACTIVE",\'Master Base\'!${master_val_letter}:${master_val_letter},"Match")'
                ws.write_formula(r0, c, f_rostered, body_fmt, rv)
                ws.write_formula(r0, c + 1, f_match, body_fmt, mv)
                ws.write_formula(r0, c + 2, f'=IFERROR(MIN(100,{xl_col_to_name(c+1)}{excel_row}/{xl_col_to_name(c)}{excel_row}*100),0)', pct_fmt, _pct(mv, rv))

    grand_r0 = 2 + len(sites)
    grand_excel = grand_r0 + 1
    last_excel = 2 + len(sites)
    ws.write(grand_r0, 0, "Grand Total", grand_site_fmt)
    for i, d in enumerate(report_dates):
        c = 1 + i * 3
        av = sum(metrics.active[(s, d)] for s in sites)
        rv = sum(metrics.rostered[(s, d)] for s in sites)
        mv = sum(metrics.matched[(s, d)] for s in sites)
        if mode == "adoption":
            for off, cached in [(0, av), (1, rv)]:
                letter = xl_col_to_name(c + off)
                ws.write_formula(grand_r0, c + off, f'=SUM({letter}{first_excel}:{letter}{last_excel})', grand_fmt, cached)
            ws.write_formula(grand_r0, c + 2, f'=IFERROR(MIN(100,{xl_col_to_name(c+1)}{grand_excel}/{xl_col_to_name(c)}{grand_excel}*100),0)', grand_pct_fmt, _pct(rv, av))
        else:
            for off, cached in [(0, rv), (1, mv)]:
                letter = xl_col_to_name(c + off)
                ws.write_formula(grand_r0, c + off, f'=SUM({letter}{first_excel}:{letter}{last_excel})', grand_fmt, cached)
            ws.write_formula(grand_r0, c + 2, f'=IFERROR(MIN(100,{xl_col_to_name(c+1)}{grand_excel}/{xl_col_to_name(c)}{grand_excel}*100),0)', grand_pct_fmt, _pct(mv, rv))
        ws.conditional_format(2, c + 2, grand_r0 - 1, c + 2, {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})


def _write_master_base(
    workbook, roster: RosterData, attendance: AttendanceData,
    att_by_date: dict[date, pd.Series], roster_by_date: dict[date, pd.Series],
    validation_by_date: dict[date, pd.Series], fmts: dict[str, object], cb: ProgressCallback | None
) -> tuple[dict[date, int], int]:
    """Write the resolved Master Base values used as formula caches.

    After the workbook is closed, _inject_master_base_formulas() replaces the three
    per-date fields with real Excel formulas while preserving these visible cached values.
    This keeps generation fast and reproduces the formula-driven reference workbook.
    """
    ws = workbook.get_worksheet_by_name("Master Base") or workbook.add_worksheet("Master Base")
    ws.freeze_panes(2, 4)
    ws.set_column(0, 0, 14.66, fmts["master_imported"])
    ws.set_column(1, 1, 22.66, fmts["master_imported"])
    ws.set_column(2, 2, 8.43, fmts["master_imported"])
    ws.set_column(3, 3, 18.66, fmts["master_imported"])

    dates = [d for d, _ in roster.date_columns]
    validation_cols: dict[date, int] = {}
    ws.write(0, 0, "Employee ID", fmts["master_header"])
    ws.write(0, 1, "Name", fmts["master_header"])
    ws.write(0, 2, "Role", fmts["master_header"])
    ws.write(0, 3, "Facility Name", fmts["master_header"])
    for i, d in enumerate(dates):
        c = 4 + i * 3
        validation_cols[d] = c + 2
        ws.write_datetime(0, c, datetime.combine(d, datetime.min.time()), fmts["master_date"])
        ws.write_blank(0, c + 1, None, fmts["master_sub"])
        ws.write_blank(0, c + 2, None, fmts["master_sub"])
        ws.set_column(c, c, 10.33, fmts["master_formula"])
        ws.set_column(c + 1, c + 2, 8.43, fmts["master_formula"])
    sub_col = 4 + 3 * len(dates)
    status_col = sub_col + 1
    ws.write(0, sub_col, "Sub Status", fmts["master_header"])
    ws.write(0, status_col, "Status", fmts["master_header"])

    for c in [0, 1, 2, 3]:
        ws.write_blank(1, c, None, fmts["master_sub"])
    for i, d in enumerate(dates):
        c = 4 + i * 3
        ws.write(1, c, "Attendance Shift", fmts["master_sub"])
        ws.write(1, c + 1, "Roster Shift", fmts["master_sub"])
        ws.write(1, c + 2, "Validation", fmts["master_sub"])
    ws.write_blank(1, sub_col, None, fmts["master_sub"])
    ws.write_blank(1, status_col, None, fmts["master_sub"])
    ws.set_column(sub_col, status_col, 8.43, fmts["master_formula"])

    attendance_dates = set(attendance.date_columns)
    user_values = roster.source[roster.user_col].astype(str).tolist()
    name_values = roster.source[roster.name_col].tolist()
    role_values = roster.source[roster.role_col].tolist()
    facility_values = roster.source[roster.facility_col].tolist()
    sub_values = roster.source[roster.sub_status_col].tolist()
    status_values = roster.source[roster.status_col].tolist()
    total = len(roster.source)

    for i in range(total):
        values: list[object] = [user_values[i], name_values[i], role_values[i], facility_values[i]]
        for d in dates:
            if d in attendance_dates:
                att_val = att_by_date[d].iat[i]
                val = validation_by_date[d].iat[i]
                values.extend([att_val or None, roster_by_date[d].iat[i] or None, val or None])
            else:
                values.extend([None, roster_by_date[d].iat[i] or None, None])
        values.extend([sub_values[i], status_values[i]])
        ws.write_row(i + 2, 0, values)
        if cb and (i + 1) % 4000 == 0:
            pct = 30 + int(30 * (i + 1) / max(total, 1))
            _progress(cb, pct, f"Building Master Base... {i+1:,}/{total:,} employees")
    return validation_cols, sub_col


_CELL_XML_RE = re.compile(rb'<c r="([A-Z]+)(\d+)"[^>]*(?:>.*?</c>|/>)')
_ROW_XML_RE = re.compile(rb'(<row r="(\d+)"[^>]*>)(.*?)(</row>)', re.S)
_INLINE_TEXT_RE = re.compile(rb'<t(?: [^>]*)?>(.*?)</t>', re.S)
_CELL_VALUE_RE = re.compile(rb'<v>(.*?)</v>', re.S)


def _excel_col_index(letters: str) -> int:
    value = 0
    for ch in letters:
        value = value * 26 + ord(ch) - 64
    return value - 1


def _cached_cell_text(cell_xml: bytes | None) -> str:
    if not cell_xml:
        return ""
    match = _INLINE_TEXT_RE.search(cell_xml)
    if match:
        return match.group(1).decode("utf-8")
    match = _CELL_VALUE_RE.search(cell_xml)
    if match:
        return match.group(1).decode("utf-8")
    return ""


def _formula_cell_xml(ref: str, style_id: str, formula: str, cached_xml_text: str) -> bytes:
    cached = xml_escape(html.unescape(cached_xml_text)) if cached_xml_text else ""
    formula_xml = xml_escape(formula)
    value_xml = f"<v>{cached}</v>" if cached else "<v/>"
    return (
        f'<c r="{ref}" s="{style_id}" t="str"><f>{formula_xml}</f>{value_xml}</c>'
    ).encode("utf-8")


def _worksheet_part_for_name(zf: zipfile.ZipFile, sheet_name: str) -> str:
    workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rel_id = None
    for sheet in workbook_root.findall(f"{{{main_ns}}}sheets/{{{main_ns}}}sheet"):
        if sheet.attrib.get("name") == sheet_name:
            rel_id = sheet.attrib.get(f"{{{rel_ns}}}id")
            break
    if not rel_id:
        raise ReportError(f"Worksheet not found while finalizing formulas: {sheet_name}")

    rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    for rel in rel_root.findall(f"{{{package_rel_ns}}}Relationship"):
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib.get("Target", "")
            if target.startswith("/"):
                return target.lstrip("/")
            return "xl/" + target.lstrip("./")
    raise ReportError(f"Worksheet relationship not found while finalizing formulas: {sheet_name}")


def _inject_master_base_formulas(
    workbook_path: Path,
    roster: RosterData,
    attendance: AttendanceData,
) -> None:
    """Inject real Master Base formulas into the finished XLSX, preserving cached values.

    Directly finalizing the worksheet XML is substantially faster than asking XlsxWriter
    to serialize more than a million unique formulas one Python cell-call at a time. The
    resulting cells are ordinary Excel formulas and recalculate normally in Excel.
    """
    replacement_path = workbook_path.with_name(workbook_path.stem + ".formula.xlsx")
    if replacement_path.exists():
        replacement_path.unlink()

    employee_ids = roster.source[roster.user_col].astype(str).tolist()
    attendance_row = {
        str(emp_id): idx + 2
        for idx, emp_id in enumerate(attendance.resolved["Employee Code"].astype(str).tolist())
        if str(emp_id)
    }
    attendance_dates = set(attendance.date_columns)
    attendance_col_letter = {
        d: xl_col_to_name(4 + idx) for idx, d in enumerate(attendance.date_columns)
    }
    roster_col_letter = {
        d: xl_col_to_name(roster.source.columns.get_loc(source_col))
        for d, source_col in roster.date_columns
    }
    date_meta = []
    for idx, (d, _source_col) in enumerate(roster.date_columns):
        master_att_col = 4 + idx * 3
        date_meta.append((
            d,
            master_att_col,
            xl_col_to_name(master_att_col),
            xl_col_to_name(master_att_col + 1),
            xl_col_to_name(master_att_col + 2),
            attendance_col_letter.get(d),
            roster_col_letter[d],
            d in attendance_dates,
        ))

    with zipfile.ZipFile(workbook_path, "r") as zin:
        master_part = _worksheet_part_for_name(zin, "Master Base")
        master_xml = zin.read(master_part)
        # Validation is populated for every roster row on an attendance-covered date, so
        # G3 reliably carries the Master Base body style in a non-empty report.
        style_match = re.search(rb'<c r="G3" s="(\d+)"', master_xml)
        if not style_match:
            style_match = re.search(rb'<c r="E3" s="(\d+)"', master_xml)
        formula_style = style_match.group(1).decode("ascii") if style_match else "0"

        sheet_data_open = master_xml.index(b"<sheetData>") + len(b"<sheetData>")
        sheet_data_close = master_xml.index(b"</sheetData>")

        with zipfile.ZipFile(replacement_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
            for info in zin.infolist():
                if info.filename != master_part:
                    with zin.open(info, "r") as src, zout.open(info, "w") as dst:
                        shutil.copyfileobj(src, dst, length=1024 * 1024)
                    continue

                with zout.open(info, "w") as dst:
                    dst.write(master_xml[:sheet_data_open])
                    cursor = sheet_data_open
                    for row_match in _ROW_XML_RE.finditer(master_xml, sheet_data_open, sheet_data_close):
                        dst.write(master_xml[cursor:row_match.start()])
                        row_number = int(row_match.group(2))
                        if row_number < 3 or row_number > len(employee_ids) + 2:
                            dst.write(row_match.group(0))
                            cursor = row_match.end()
                            continue

                        roster_idx = row_number - 3
                        employee_id = employee_ids[roster_idx]
                        source_att_row = attendance_row.get(employee_id)
                        existing_cells: dict[int, bytes] = {}
                        for cell_match in _CELL_XML_RE.finditer(row_match.group(3)):
                            col_idx = _excel_col_index(cell_match.group(1).decode("ascii"))
                            existing_cells[col_idx] = cell_match.group(0)

                        for (
                            _d, master_att_col, master_att_letter, master_roster_letter,
                            master_validation_letter, source_att_letter, source_roster_letter,
                            has_attendance,
                        ) in date_meta:
                            if has_attendance and source_att_row is not None:
                                old_xml = existing_cells.get(master_att_col)
                                cached = _cached_cell_text(old_xml)
                                formula = (
                                    f"IF('Attendance Base'!${source_att_letter}${source_att_row}=\"\",\"\","
                                    f"'Attendance Base'!${source_att_letter}${source_att_row})"
                                )
                                existing_cells[master_att_col] = _formula_cell_xml(
                                    f"{master_att_letter}{row_number}", formula_style, formula, cached
                                )

                            old_xml = existing_cells.get(master_att_col + 1)
                            cached = _cached_cell_text(old_xml)
                            source_roster_row = roster_idx + 2
                            formula = (
                                f"IF('Roster Base'!${source_roster_letter}${source_roster_row}=\"\",\"\","
                                f"'Roster Base'!${source_roster_letter}${source_roster_row})"
                            )
                            existing_cells[master_att_col + 1] = _formula_cell_xml(
                                f"{master_roster_letter}{row_number}", formula_style, formula, cached
                            )

                            if has_attendance:
                                old_xml = existing_cells.get(master_att_col + 2)
                                cached = _cached_cell_text(old_xml) or "No Match"
                                formula = (
                                    f"IF(OR(AND({master_att_letter}{row_number}<>\"\","
                                    f"{master_att_letter}{row_number}={master_roster_letter}{row_number}),"
                                    f"AND({master_roster_letter}{row_number}=\"WO\","
                                    f"{master_att_letter}{row_number}=\"-\")),\"Match\",\"No Match\")"
                                )
                                existing_cells[master_att_col + 2] = _formula_cell_xml(
                                    f"{master_validation_letter}{row_number}", formula_style, formula, cached
                                )

                        dst.write(row_match.group(1))
                        for col_idx in sorted(existing_cells):
                            dst.write(existing_cells[col_idx])
                        dst.write(row_match.group(4))
                        cursor = row_match.end()

                    dst.write(master_xml[cursor:sheet_data_close])
                    dst.write(master_xml[sheet_data_close:])

    os.replace(replacement_path, workbook_path)

def _write_roster_base(workbook, roster: RosterData, fmts: dict[str, object], cb: ProgressCallback | None) -> dict[str, object]:
    ws = workbook.get_worksheet_by_name("Roster Base") or workbook.add_worksheet("Roster Base")
    ws.freeze_panes(1, 0)
    ws.set_column(0, 0, 14.66, fmts["base_body"])
    ws.set_column(1, 1, 22.66, fmts["base_body"])
    ws.set_column(2, 2, 11.0, fmts["base_body"])
    ws.set_column(3, 6, 8.43, fmts["base_body"])
    ws.set_column(7, 7, 18.66, fmts["base_body"])

    date_by_col = {col: d for d, col in roster.date_columns}
    onboarded_idx = roster.source.columns.get_loc(roster.onboarded_date_col) if roster.onboarded_date_col else None
    onboarded_values = _parse_date_series(roster.source[roster.onboarded_date_col]) if roster.onboarded_date_col else None
    lwd_idx = roster.source.columns.get_loc(roster.last_working_day_col) if roster.last_working_day_col else None
    lwd_values = _parse_date_series(roster.source[roster.last_working_day_col]) if roster.last_working_day_col else None
    for c, col in enumerate(roster.source.columns):
        d = date_by_col.get(col)
        if d:
            ws.write_datetime(0, c, datetime.combine(d, datetime.min.time()), fmts["base_date_header"])
            ws.set_column(c, c, 10.33, fmts["base_body_center"])
        else:
            ws.write(0, c, col, fmts["base_header"])
            if (onboarded_idx is not None and c == onboarded_idx) or (lwd_idx is not None and c == lwd_idx):
                ws.set_column(c, c, 13.0, fmts["base_date_body"])

    total = len(roster.source)
    for r, row in enumerate(roster.source.itertuples(index=False, name=None), start=1):
        values = [val if val != "" else None for val in row]
        if onboarded_idx is not None and onboarded_values is not None:
            parsed_start = onboarded_values.iloc[r - 1]
            values[onboarded_idx] = parsed_start.to_pydatetime() if pd.notna(parsed_start) else None
        if lwd_idx is not None and lwd_values is not None:
            parsed_lwd = lwd_values.iloc[r - 1]
            values[lwd_idx] = parsed_lwd.to_pydatetime() if pd.notna(parsed_lwd) else None
        ws.write_row(r, 0, values)
        if cb and r % 5000 == 0:
            _progress(cb, 15 + int(10 * r / max(total, 1)), f"Writing Roster Base... {r:,}/{total:,}")
    return {
        "facility_col": roster.source.columns.get_loc(roster.facility_col),
        "sub_status_col": roster.source.columns.get_loc(roster.sub_status_col),
        "onboarded_date_col": onboarded_idx,
        "last_working_day_col": lwd_idx,
        "date_cols": {d: roster.source.columns.get_loc(col) for d, col in roster.date_columns},
    }


def _write_attendance_base(workbook, attendance: AttendanceData, fmts: dict[str, object], cb: ProgressCallback | None) -> None:
    ws = workbook.get_worksheet_by_name("Attendance Base") or workbook.add_worksheet("Attendance Base")
    ws.freeze_panes(1, 0)
    ws.set_column(0, 0, 14.66, fmts["base_body"])
    ws.set_column(1, 1, 22.66, fmts["base_body"])
    ws.set_column(2, 2, 8.43, fmts["base_body"])
    ws.set_column(3, 3, 18.66, fmts["base_body"])
    headers = ["Employee Code", "Name", "Role", "Facility Name"] + attendance.date_columns
    for c, h in enumerate(headers):
        if isinstance(h, date):
            ws.write_datetime(0, c, datetime.combine(h, datetime.min.time()), fmts["base_date_header"])
            ws.set_column(c, c, 10.33, fmts["base_body_center"])
        else:
            ws.write(0, c, h, fmts["base_header"])
    total = len(attendance.resolved)
    for r, row in enumerate(attendance.resolved.itertuples(index=False, name=None), start=1):
        values = [val if val != "" else None for val in row]
        ws.write_row(r, 0, values)
        if cb and r % 7000 == 0:
            _progress(cb, 25 + int(5 * r / max(total, 1)), f"Writing Attendance Base... {r:,}/{total:,}")


def _write_methodology(
    workbook, inputs: ReportInputs, scan: ScanResult, attendance: AttendanceData,
    report_dates: list[date], fmts: dict[str, object]
) -> None:
    ws = workbook.get_worksheet_by_name("Methodology & Data Quality") or workbook.add_worksheet("Methodology & Data Quality")
    ws.set_column(0, 0, 140.66)
    lines = [
        ("Roster Automation Report - Methodology & Data Quality", "title"),
        (f"Generated on {date.today():%d-%b-%Y}. Report view date range: {report_dates[0]:%d-%b-%Y} to {report_dates[-1]:%d-%b-%Y}.", "note"),
        ("1. Source files", "section"),
        (f"Roster: {Path(inputs.roster_path).name} ({scan.roster_rows:,} employee rows; roster dates {scan.roster_start:%d-%b-%Y} to {scan.roster_end:%d-%b-%Y}).", "note"),
        (f"Attendance: {Path(inputs.attendance_path).name} ({scan.attendance_raw_rows:,} raw rows; attendance dates {scan.attendance_start:%d-%b-%Y} to {scan.attendance_end:%d-%b-%Y}). Only each date's shift field is used; total hours, working hours and punch status are ignored.", "note"),
        ("2. Attendance duplicate resolution", "section"),
        ("For duplicated Employee Codes, the attendance row whose Facility Name matches the employee's current Facility Name in Roster Base is retained. If no duplicate row matches the roster facility, the first source occurrence is retained.", "note"),
        ("3. Master Base", "section"),
        ("One row is created for every roster employee, in roster order. For each roster date, Attendance Shift is formula-linked to the resolved employee/date cell in Attendance Base, Roster Shift is formula-linked to the corresponding Roster Base cell, and Validation is calculated by formula from those two Master Base cells. Dates not present in attendance keep Attendance/Validation blank while Roster Shift remains formula-linked.", "note"),
        ("Validation = Match when Attendance Shift equals Roster Shift and attendance is non-blank, OR when Roster Shift is WO and Attendance Shift is '-'. Otherwise Validation = No Match.", "note"),
        ("4. Metrics", "section"),
        ("Active HC is formula-driven by each date header using lifecycle dates available in Roster Base. Onboarded/Joining Date (when present) prevents future joiners from being counted before they started; Last Working Day/LWD (when present) restores employees who were active historically but are now non-active. These lifecycle fields are detected by header name, not position. If neither lifecycle date exists, the ACTIVE-only fallback remains because historical active status cannot be reconstructed reliably from roster shift columns alone.", "note"),
        ("Rostered HC = ACTIVE employees whose roster shift for the date is P1, P2, P3 or WO.", "note"),
        ("Roster Adoption % = MIN(100, Rostered HC / Active HC * 100), with 0 when the denominator is zero.", "note"),
        ("Matched Count = ACTIVE employees at the facility whose Master Base Validation is Match for the date.", "note"),
        ("Roster Adherence % = MIN(100, Matched Count / Rostered HC * 100), with 0 when the denominator is zero.", "note"),
        ("Grand Total percentages are recalculated from total counts; they are not simple averages of site percentages.", "note"),
        ("5. GUI controls applied", "section"),
        ("Only facilities present in BOTH roster and attendance are eligible for the DC/XD selector. The selected site list is applied consistently to Summary, Adoption and Adherence tabs. Summary and Adherence follow the selected report date range; Roster Adoption always spans the full roster month/date coverage. Source bases and Master Base retain their natural source coverage.", "note"),
        (f"Selected DC sites ({len(inputs.selected_dc)}): {', '.join(inputs.selected_dc) if inputs.selected_dc else 'None'}", "note"),
        (f"Selected XD sites ({len(inputs.selected_xd)}): {', '.join(inputs.selected_xd) if inputs.selected_xd else 'None'}", "note"),
    ]
    row = 0
    for text, kind in lines:
        if kind == "title":
            ws.write(row, 0, text, fmts["note_title"])
            ws.set_row(row, 15.6)
        elif kind == "section":
            ws.write(row, 0, text, fmts["note_section"])
        else:
            ws.write(row, 0, text, fmts["note"])
            ws.set_row(row, 34)
        row += 1


def generate_report(inputs: ReportInputs, progress: ProgressCallback | None = None) -> str:
    _progress(progress, 2, "Scanning and validating inputs...")
    scan = scan_inputs(inputs.roster_path, inputs.attendance_path)
    report_dates = _validate_report_inputs(inputs, scan)

    out_folder = Path(inputs.output_folder)
    out_folder.mkdir(parents=True, exist_ok=True)
    output_path = out_folder / default_output_filename()

    _progress(progress, 6, "Loading roster and attendance...")
    roster = load_roster(inputs.roster_path)
    attendance = load_attendance(inputs.attendance_path, roster)
    _progress(progress, 10, f"Attendance resolved to {len(attendance.resolved):,} unique employee codes.")

    att_by_date, roster_by_date, validation_by_date = _compute_master_arrays(roster, attendance)
    # Summary and Adherence follow the GUI-selected report range (normally through D-1).
    # Adoption is roster-only and must always span every date available in the roster month.
    adoption_dates = [d for d, _ in roster.date_columns]
    selected_sites = inputs.selected_dc + inputs.selected_xd
    metrics = _compute_metrics(roster, report_dates, selected_sites, roster_by_date, validation_by_date)

    temp_path = output_path.with_suffix(".tmp.xlsx")
    if temp_path.exists():
        temp_path.unlink()
    workbook = xlsxwriter.Workbook(str(temp_path), {"constant_memory": True})
    workbook.set_properties({
        "title": "Roster Report",
        "subject": "Roster adoption and adherence automation",
        "company": "",
        "comments": "Generated by Roster Automation",
    })
    try:
        workbook.set_calc_mode("auto")
    except Exception:
        pass
    fmts = _formats(workbook)

    # Create sheets in the same user-facing order as the reference workbook.
    for sheet_name in [
        "DC Summary", "XD Summary", "Master Base", "Roster Base", "Attendance Base",
        "DC Roster Adoption", "XD Roster Adoption", "DC Roster Adherence",
        "XD Roster Adherence", "Methodology & Data Quality",
    ]:
        workbook.add_worksheet(sheet_name)

    try:
        _progress(progress, 14, "Writing source bases...")
        roster_positions = _write_roster_base(workbook, roster, fmts, progress)
        _write_attendance_base(workbook, attendance, fmts, progress)

        _progress(progress, 30, "Building Master Base formulas...")
        master_validation_cols, master_sub_col = _write_master_base(
            workbook, roster, attendance, att_by_date, roster_by_date, validation_by_date, fmts, progress
        )

        _progress(progress, 62, "Building DC/XD summary views...")
        if inputs.selected_dc:
            dc_metrics = _compute_metrics(roster, report_dates, inputs.selected_dc, roster_by_date, validation_by_date)
            dc_summary_sites = _sort_sites_by_latest_adherence(inputs.selected_dc, dc_metrics, report_dates)
            _write_summary_sheet(
                workbook, "DC Summary", dc_summary_sites, report_dates, dc_metrics,
                roster_positions["date_cols"], roster_positions["facility_col"], roster_positions["sub_status_col"],
                roster_positions["onboarded_date_col"], roster_positions["last_working_day_col"],
                master_validation_cols, master_sub_col, fmts,
            )
        else:
            workbook.get_worksheet_by_name("DC Summary").write(0, 0, "No DC sites selected")

        if inputs.selected_xd:
            xd_metrics = _compute_metrics(roster, report_dates, inputs.selected_xd, roster_by_date, validation_by_date)
            xd_summary_sites = _sort_sites_by_latest_adherence(inputs.selected_xd, xd_metrics, report_dates)
            _write_summary_sheet(
                workbook, "XD Summary", xd_summary_sites, report_dates, xd_metrics,
                roster_positions["date_cols"], roster_positions["facility_col"], roster_positions["sub_status_col"],
                roster_positions["onboarded_date_col"], roster_positions["last_working_day_col"],
                master_validation_cols, master_sub_col, fmts,
            )
        else:
            workbook.get_worksheet_by_name("XD Summary").write(0, 0, "No XD sites selected")

        _progress(progress, 72, "Building adoption and adherence views...")
        if inputs.selected_dc:
            dc_adoption_metrics = _compute_metrics(roster, adoption_dates, inputs.selected_dc, roster_by_date, validation_by_date)
            dc_adherence_metrics = _compute_metrics(roster, report_dates, inputs.selected_dc, roster_by_date, validation_by_date)
            _write_three_metric_sheet(workbook, "DC Roster Adoption", inputs.selected_dc, adoption_dates, dc_adoption_metrics, "adoption", roster_positions["date_cols"], roster_positions["facility_col"], roster_positions["sub_status_col"], roster_positions["onboarded_date_col"], roster_positions["last_working_day_col"], master_validation_cols, master_sub_col, fmts)
            _write_three_metric_sheet(workbook, "DC Roster Adherence", inputs.selected_dc, report_dates, dc_adherence_metrics, "adherence", roster_positions["date_cols"], roster_positions["facility_col"], roster_positions["sub_status_col"], roster_positions["onboarded_date_col"], roster_positions["last_working_day_col"], master_validation_cols, master_sub_col, fmts)
        else:
            workbook.get_worksheet_by_name("DC Roster Adoption").write(0, 0, "No DC sites selected")
            workbook.get_worksheet_by_name("DC Roster Adherence").write(0, 0, "No DC sites selected")

        if inputs.selected_xd:
            xd_adoption_metrics = _compute_metrics(roster, adoption_dates, inputs.selected_xd, roster_by_date, validation_by_date)
            xd_adherence_metrics = _compute_metrics(roster, report_dates, inputs.selected_xd, roster_by_date, validation_by_date)
            _write_three_metric_sheet(workbook, "XD Roster Adoption", inputs.selected_xd, adoption_dates, xd_adoption_metrics, "adoption", roster_positions["date_cols"], roster_positions["facility_col"], roster_positions["sub_status_col"], roster_positions["onboarded_date_col"], roster_positions["last_working_day_col"], master_validation_cols, master_sub_col, fmts)
            _write_three_metric_sheet(workbook, "XD Roster Adherence", inputs.selected_xd, report_dates, xd_adherence_metrics, "adherence", roster_positions["date_cols"], roster_positions["facility_col"], roster_positions["sub_status_col"], roster_positions["onboarded_date_col"], roster_positions["last_working_day_col"], master_validation_cols, master_sub_col, fmts)
        else:
            workbook.get_worksheet_by_name("XD Roster Adoption").write(0, 0, "No XD sites selected")
            workbook.get_worksheet_by_name("XD Roster Adherence").write(0, 0, "No XD sites selected")

        _progress(progress, 84, "Writing methodology and final checks...")
        _write_methodology(workbook, inputs, scan, attendance, report_dates, fmts)
        workbook.close()

        _progress(progress, 90, "Finalizing Master Base formulas...")
        _inject_master_base_formulas(temp_path, roster, attendance)

        # Put sheets in the same logical order as the reference report by patching workbook sheet order
        # is unnecessary; XlsxWriter sheet order follows creation order. Here bases are intentionally first
        # during creation for streaming speed. Excel formulas work regardless of sheet order.

        if output_path.exists():
            output_path.unlink()
        os.replace(temp_path, output_path)
        _progress(progress, 100, f"Report created: {output_path.name}")
        return str(output_path)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass
        if temp_path.exists():
            temp_path.unlink()
        raise
