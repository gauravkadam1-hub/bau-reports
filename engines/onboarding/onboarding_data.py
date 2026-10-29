"""Polars ingestion, date filtering, aggregation and stable facility ordering."""
from __future__ import annotations

import csv
import io
from datetime import date, datetime
from pathlib import Path

import polars as pl
from engines.facility_scope import filter_polars_dcxd

try:
    from .onboarding_common import (
        AADHAAR_COLUMN, APPROVER_COLUMN, BASE_SHEET, DATE_COLUMN, FACILITY_COLUMN,
        HV_COLUMN, MANUAL_HR, RED_HR, Counts, PreparedReport, clean_facility,
        clean_selection, excel_text_matches, facility_type,
    )
    from .utils import ReportError, inclusive_date_range, normalize_header, parse_date_value
except ImportError:
    from onboarding_common import (
        AADHAAR_COLUMN, APPROVER_COLUMN, BASE_SHEET, DATE_COLUMN, FACILITY_COLUMN,
        HV_COLUMN, MANUAL_HR, RED_HR, Counts, PreparedReport, clean_facility,
        clean_selection, excel_text_matches, facility_type,
    )
    from utils import ReportError, inclusive_date_range, normalize_header, parse_date_value

try:
    from .onboarding_input import validate_headers, _xlsx_headers, _mixed_numeric_date_overrides
except ImportError:
    from onboarding_input import validate_headers, _xlsx_headers, _mixed_numeric_date_overrides


def _read_csv(path, headers):
    error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError as exc:
            error = exc
    else:
        raise ReportError(f"Could not read the selected CSV file.\n\n{error}")
    try:
        dialect = csv.Sniffer().sniff(text[:65536], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    raw_headers = next(csv.reader(io.StringIO(text), dialect), [])
    mapping = validate_headers(raw_headers)
    missing = [name for name in headers if normalize_header(name) not in mapping]
    if missing:
        raise ReportError("The uploaded file does not contain a valid Onboarding Base header row.\n\nMissing required headers: " + ", ".join(missing))
    if dialect.skipinitialspace or dialect.escapechar or not dialect.doublequote:
        # Keep unusual CSV dialect behavior identical to the old CSV reader.
        rows = list(csv.reader(io.StringIO(text), dialect))[1:]
        selected = [[row[mapping[normalize_header(h)]] if mapping[normalize_header(h)] < len(row) else None for h in headers] for row in rows]
        return pl.DataFrame(selected, schema={h: pl.String for h in headers}, orient="row")
    frame = pl.read_csv(io.StringIO(text), separator=dialect.delimiter, quote_char=dialect.quotechar,
                        has_header=True, infer_schema=False, try_parse_dates=False,
                        truncate_ragged_lines=True, raise_if_empty=False)
    # Source header spelling/case can differ; output retains the canonical order.
    names = {normalize_header(name): name for name in frame.columns}
    return frame.select([pl.col(names[normalize_header(h)]).alias(h) for h in headers])


def load_input(input_path, profile):
    """Read CSV or XLSX/XLSM, with Polars + Calamine/FastExcel for Excel data."""
    path = Path(input_path).expanduser().resolve()
    if not path.exists():
        raise ReportError("Please select an existing onboarding export.")
    suffix = path.suffix.casefold()
    headers = profile["headers"]
    sheet = None
    try:
        if suffix == ".csv":
            frame = _read_csv(path, headers)
        elif suffix in {".xlsx", ".xlsm"}:
            try:
                import fastexcel  # noqa: F401; explicit actionable error before parsing.
            except ImportError as exc:
                raise ReportError("Excel reading needs fastexcel. Install the updated requirements.txt and restart the app.") from exc
            sheets = _xlsx_headers(path)
            candidates = sorted(sheets, key=lambda item: item["name"] != BASE_SHEET)
            best_missing = headers
            for candidate in candidates:
                mapping = validate_headers(candidate["headers"])
                missing = [h for h in headers if normalize_header(h) not in mapping]
                if not missing:
                    sheet = candidate
                    break
                if len(missing) < len(best_missing):
                    best_missing = missing
            if sheet is None:
                raise ReportError("The uploaded Excel does not contain a valid Onboarding Base header row.\n\nMissing required headers: " + ", ".join(best_missing))
            frame = pl.read_excel(path, sheet_name=sheet["name"], engine="calamine", infer_schema_length=None,
                                  drop_empty_rows=False, drop_empty_cols=False, raise_if_empty=False)
            names = {normalize_header(name): name for name in frame.columns}
            frame = frame.select([pl.col(names[normalize_header(h)]).alias(h) for h in headers])
        elif suffix == ".xls":
            raise ReportError("Old .xls onboarding exports are not supported. Save/export as .xlsx or .csv and try again.")
        else:
            raise ReportError("Please select an .xlsx, .xlsm, or .csv onboarding export.")
        dtype = frame.schema[DATE_COLUMN]
        if dtype == pl.Date:
            dates = frame.get_column(DATE_COLUMN)
        elif dtype.base_type() == pl.Datetime:
            dates = frame.get_column(DATE_COLUMN).dt.date()
        else:
            values = frame.get_column(DATE_COLUMN).to_list()
            parsed = [parse_date_value(value) for value in values]
            if sheet is not None and dtype == pl.String:
                for index, value in _mixed_numeric_date_overrides(path, sheet, values).items():
                    if index < len(parsed):
                        parsed[index] = value
            dates = pl.Series(DATE_COLUMN, parsed, dtype=pl.Date)
        return frame.with_columns(dates.alias(DATE_COLUMN))
    except ReportError:
        raise
    except PermissionError as exc:
        raise ReportError("The selected file cannot be read. Close it in Excel and check file permissions.") from exc
    except Exception as exc:
        raise ReportError(f"Could not read the onboarding export.\n\n{exc}") from exc


def available_dates(frame):
    dates = frame.get_column(DATE_COLUMN).drop_nulls()
    if not len(dates):
        raise ReportError("No valid dates were found under 'Onboarded Date' in the uploaded base.")
    return dates.min(), dates.max()


def selected_dates(frame, first=None, last=None, enforce_bounds=False):
    available_first, available_last = available_dates(frame)
    first = first or available_first
    last = last or available_last
    if isinstance(first, datetime): first = first.date()
    if isinstance(last, datetime): last = last.date()
    if not isinstance(first, date) or not isinstance(last, date):
        raise ReportError("Please select a valid From and To date.")
    if first > last:
        raise ReportError("The selected From date cannot be after the To date.")
    if enforce_bounds and (first < available_first or last > available_last):
        raise ReportError("The selected date range must stay inside the dates available in the uploaded file.\n\n"
                          f"Available: {available_first.strftime('%d %b %Y')} to {available_last.strftime('%d %b %Y')}")
    return first, last


def filter_dates(frame, first, last):
    return frame.filter(pl.col(DATE_COLUMN).is_between(first, last, closed="both"))


def source_facilities(frame, first, last):
    filtered = filter_dates(frame, first, last)
    values = filtered.get_column(FACILITY_COLUMN).unique(maintain_order=True).to_list()
    return {kind: clean_selection(values, kind) for kind in ("DC", "XD")}


def _text(column):
    return pl.col(column).cast(pl.String).fill_null("")


def _ordered_names(names, existing):
    names = set(names)
    ordered = [name for name in existing if name in names]
    ordered.extend(sorted(names.difference(ordered), key=str.casefold))
    return ordered


def prepare_report(frame, profile, first, last, selected_dc=None, selected_xd=None):
    """Preserve the original business rules; calculate once in Polars."""
    count_days = (last-first).days+1
    if 1+4*count_days+min(2, count_days) > 16384:
        raise ReportError("The selected date range exceeds Excel's column limit.")
    # Scope before aggregations so the HR counts, cached values and Excel
    # COUNTIFS formulas all refer to the same smaller, DC/XD-only base.
    filtered = filter_polars_dcxd(filter_dates(frame, first, last), FACILITY_COLUMN)
    work = filtered.with_columns(
        _text(FACILITY_COLUMN).alias("_raw_site"),
        _text(FACILITY_COLUMN).map_elements(lambda value: clean_facility(value).casefold(), return_dtype=pl.String).alias("_sort_site"),
        _text(AADHAAR_COLUMN).str.to_lowercase().alias("_aadhar"),
        _text(HV_COLUMN).str.to_lowercase().alias("_hyper"),
        _text(APPROVER_COLUMN).alias("_raw_hr"),
        _text(APPROVER_COLUMN).str.strip_chars().alias("_name"),
    )
    grouped = work.group_by(["_raw_site", DATE_COLUMN]).agg(
        pl.len().alias("total"), (pl.col("_aadhar") == "y").sum().alias("aadhar"),
        (pl.col("_aadhar") == "n").sum().alias("manual"), (pl.col("_hyper") == "red").sum().alias("red"),
    )
    raw_counts = {(r["_raw_site"], r[DATE_COLUMN]): Counts(int(r["total"]), int(r["aadhar"]), int(r["manual"]), int(r["red"]))
                  for r in grouped.iter_rows(named=True)}
    # Sorting intentionally trims status values; COUNTIFS metric counts do not.
    latest = work.filter(pl.col(DATE_COLUMN) == last).group_by("_sort_site").agg(
        (pl.col("_aadhar").str.strip_chars() == "n").sum().alias("manual"),
        (pl.col("_hyper").str.strip_chars() == "red").sum().alias("red"),
    )
    sort_counts = {r["_sort_site"]: (int(r["red"]), int(r["manual"])) for r in latest.iter_rows(named=True)}
    chosen = {}
    for kind, supplied in (("DC", selected_dc), ("XD", selected_xd)):
        sites = clean_selection(profile["defaults"][kind] if supplied is None else supplied, kind)
        chosen[kind] = sorted(sites, key=lambda site: tuple(-n for n in sort_counts.get(site.casefold(), (0, 0))))
    dates = inclusive_date_range(first, last)
    counts = {}
    for site in chosen["DC"] + chosen["XD"]:
        matches = {raw for raw, _ in raw_counts if excel_text_matches(raw, site)}
        for day in dates:
            total = Counts()
            for raw in matches:
                total += raw_counts.get((raw, day), Counts())
            counts[(site, day)] = total
    hr = work.group_by("_raw_hr").agg(
        (pl.col("_aadhar") == "n").sum().alias("manual"), (pl.col("_hyper") == "red").sum().alias("red"),
    )
    hr_rows = list(hr.iter_rows(named=True))
    manual_names = _ordered_names(
        work.filter((pl.col("_name") != "") & (pl.col("_aadhar").str.strip_chars() == "n")).get_column("_name").to_list(),
        profile["sheets"][MANUAL_HR]["existing_names"],
    )
    red_names = _ordered_names(
        work.filter((pl.col("_name") != "") & (pl.col("_hyper").str.strip_chars() == "red")).get_column("_name").to_list(),
        profile["sheets"][RED_HR]["existing_names"],
    )
    manual_counts = {name: sum(int(r["manual"]) for r in hr_rows if excel_text_matches(r["_raw_hr"], name)) for name in manual_names}
    red_counts = {name: sum(int(r["red"]) for r in hr_rows if excel_text_matches(r["_raw_hr"], name)) for name in red_names}
    # Retain all DC/XD facilities in the selected dates, including those not
    # selected for summaries; do not deduplicate or narrow by the UI selection.
    base = filtered.select(profile["headers"]).with_columns(pl.col(DATE_COLUMN).cast(pl.Datetime("us")))
    return PreparedReport(dates, chosen["DC"], chosen["XD"], base.iter_rows(), base.height,
                          counts, manual_names, red_names, manual_counts, red_counts)
