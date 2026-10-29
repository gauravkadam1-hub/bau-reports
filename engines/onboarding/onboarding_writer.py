"""Fresh seven-sheet XLSX writer. Never opens or edits a template workbook."""
from __future__ import annotations

from datetime import date, datetime, time
from math import isfinite
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Callable

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

try:
    from .onboarding_common import (
        AADHAR_DC, AADHAR_XD, BASE_SHEET, HYPERVERGE_DC, HYPERVERGE_XD,
        MANUAL_HR, RED_HR, Counts, PreparedReport,
    )
    from .utils import ReportError, date_to_excel_datetime
    from .onboarding_styles import format_properties, validate_workbook
    from .onboarding_packaging import compact_workbook
except ImportError:
    from onboarding_common import (
        AADHAR_DC, AADHAR_XD, BASE_SHEET, HYPERVERGE_DC, HYPERVERGE_XD,
        MANUAL_HR, RED_HR, Counts, PreparedReport,
    )
    from utils import ReportError, date_to_excel_datetime
    from onboarding_styles import format_properties, validate_workbook
    from onboarding_packaging import compact_workbook


def _column_width(stored_width):
    # The template stores OOXML widths (including Excel's five-pixel padding).
    # XlsxWriter accepts character widths excluding that padding. Preserve the
    # visible width to the nearest pixel rather than adding the padding twice.
    pixels = round(stored_width * 7)
    return pixels / 12 if pixels <= 12 else (pixels - 5) / 7


def _excel_date_number(value):
    result = (value - date(1899, 12, 30)).days
    return result - 1 if value < date(1900, 3, 1) else result


class _Writer:
    def __init__(self, workbook, profile, report):
        self.workbook, self.profile, self.report = workbook, profile, report
        self.format_properties = {
            str(key): format_properties(props, str(key))
            for key, props in profile["formats"].items()
        }
        self.formats = {key: workbook.add_format(props) for key, props in self.format_properties.items()}
        self.date_formats = {}
        cf = profile["conditional"]
        props = {"bg_color": cf["positive_fill"]}
        if cf.get("positive_font"):
            props["font_color"] = cf["positive_font"]
        self.positive_format = workbook.add_format(props)
        self.sheets = {name: workbook.add_worksheet(name) for name in profile["sheet_order"]}

    def fmt(self, key):
        return self.formats[str(key)]

    def date_fmt(self, key, value):
        props = self.format_properties[str(key)] if key is not None else {}
        if props.get("num_format", "General") != "General":
            return self.fmt(key)
        kind = "time" if isinstance(value, time) else "datetime" if isinstance(value, datetime) else "date"
        cache_key = (str(key), kind)
        if cache_key not in self.date_formats:
            number_format = {"time": "h:mm:ss", "datetime": "dd-mmm-yyyy", "date": "dd-mmm-yyyy"}[kind]
            self.date_formats[cache_key] = self.workbook.add_format({**props, "num_format": number_format})
        return self.date_formats[cache_key]

    def value(self, ws, row, col, value, style=None):
        fmt = self.fmt(style) if style is not None else None
        if value is None or value == "":
            # Plain Base cells need no XML record for an empty value.
            if fmt is not None:
                ws.write_blank(row, col, None, fmt)
        elif isinstance(value, str):
            if len(value) > 32767:
                raise ReportError(f"Text exceeds Excel's cell limit at {ws.name}!{xl_col_to_name(col)}{row+1}.")
            # Source text is never executed as a formula or converted to a URL.
            ws.write_string(row, col, value, fmt)
        elif isinstance(value, bool):
            ws.write_boolean(row, col, value, fmt)
        elif isinstance(value, (datetime, date, time)):
            ws.write_datetime(row, col, value, self.date_fmt(style, value))
        elif isinstance(value, (int, float)):
            if not isfinite(value):
                raise ReportError(f"Non-finite numeric value at {ws.name}!{xl_col_to_name(col)}{row+1}.")
            ws.write_number(row, col, value, fmt)
        else:
            ws.write_string(row, col, str(value), fmt)

    def setup(self, name):
        ws, cfg = self.sheets[name], self.profile["sheets"][name]
        ws.set_default_row(cfg["default_row_height"])
        ws.hide_gridlines(0 if cfg["show_gridlines"] else 2)
        ws.set_zoom(cfg["zoom"])
        if cfg["orientation"] == "landscape":
            ws.set_landscape()
        else:
            ws.set_portrait()
        return ws, cfg

    def summary_setup(self, name, sites, block_width):
        ws, cfg = self.setup(name)
        ws.set_column(0, 0, _column_width(cfg["a_width"]))
        for row in (0, 1):
            ws.set_row(row, cfg["row_heights"].get(str(row+1), cfg["default_row_height"]))
        ws.merge_range(0, 0, 1, 0, "Site", self.fmt(cfg["site_header_format"]))
        for row, site in enumerate(sites, 2):
            ws.set_row(row, cfg["body_height"])
            ws.write_string(row, 0, site, self.fmt(cfg["site_body_format"]))
        total_row = len(sites) + 2
        ws.set_row(total_row, cfg["total_height"])
        ws.write_string(total_row, 0, "Grand Total", self.fmt(cfg["site_total_format"]))
        # Freeze two header rows and the Site column. Historical date columns
        # remain hidden but retain their formulas and can be unhidden by users.
        first_visible = 1 + max(0, len(self.report.dates)-2) * block_width
        ws.freeze_panes(2, 1)
        ws.set_selection(2, first_visible, 2, first_visible)
        return ws, cfg, total_row

    def conditional(self, ws, first_row, last_row, col, kind, maximum=None):
        if last_row < first_row:
            return
        cfg = self.profile["conditional"]
        if kind == "positive":
            rule = {"type": "cell", "criteria": ">", "value": 0, "format": self.positive_format}
        elif kind == "compliance":
            red, yellow, green = cfg["compliance_colors"]
            rule = {"type": "3_color_scale", "min_type": "min", "mid_type": "percentile", "mid_value": 50,
                    "max_type": "max", "min_color": red, "mid_color": yellow, "max_color": green}
        elif kind == "red_scale":
            # Do not create a scale for an all-zero/empty/negative-only column.
            # Totals never contribute to this maximum or to the formatted range.
            if maximum is None or maximum <= 0:
                return
            white, red = cfg["hyper_summary_colors"]
            # Numeric endpoints avoid Excel's degenerate zero-to-zero scale.
            # Zero stays white; only positive counts take on a red tint. Keep
            # the upper endpoint positive even if a user later edits all values
            # to zero. Each date uses its own generation-time maximum.
            rule = {"type": "2_color_scale", "min_type": "num", "min_value": 0,
                    "max_type": "num", "max_value": maximum,
                    "min_color": white, "max_color": red}
        else:
            raise ValueError(kind)
        ws.conditional_format(first_row, col, last_row, col, rule)

    @staticmethod
    def count_formula(row, start_col, extra=None):
        formula = f"=COUNTIFS('Onboarding Base'!$H:$H, $A{row+1}, 'Onboarding Base'!$J:$J, {xl_col_to_name(start_col)}$1"
        if extra is not None:
            column, criterion = extra
            formula += f', \'Onboarding Base\'!${column}:${column}, "{criterion}"'
        return formula + ")"

    def total(self, ws, row, col, style, cached):
        letter = xl_col_to_name(col)
        # An empty selection has a zero total, never SUM(B3:B2) (self-reference).
        formula = f"=SUM({letter}3:{letter}{row})" if row > 2 else "=0"
        ws.write_formula(row, col, formula, self.fmt(style), cached)

    def metric_block(self, ws, cfg, sites, total_row, day_index, width):
        block = cfg["block"]
        day = self.report.dates[day_index]
        start = 1 + day_index * width
        hidden = day_index < len(self.report.dates)-2
        ws.merge_range(0, start, 0, start+width-1, date_to_excel_datetime(day), self.fmt(block["date_formats"][0]))
        for offset in range(width):
            ws.set_column(start+offset, start+offset, _column_width(block["widths"][offset]), None, {"hidden": hidden})
            ws.write_string(1, start+offset, block["headers"][offset], self.fmt(block["header_formats"][offset]))
        total = Counts()
        for row, site in enumerate(sites, 2):
            counts = self.report.counts.get((site, day), Counts())
            total += counts
            if width == 2:
                formulas = [self.count_formula(row, start), self.count_formula(row, start, ("U", "red"))]
                values = [counts.total, counts.red]
            else:
                formulas = [self.count_formula(row, start), self.count_formula(row, start, ("P", "Y")),
                            self.count_formula(row, start, ("P", "N")),
                            f"=IFERROR({xl_col_to_name(start+1)}{row+1}/{xl_col_to_name(start)}{row+1},100%)"]
                values = [counts.total, counts.aadhar, counts.manual, counts.compliance]
            for offset, (formula, value) in enumerate(zip(formulas, values)):
                ws.write_formula(row, start+offset, formula, self.fmt(block["body_formats"][offset]), value)
        values = [total.total, total.red] if width == 2 else [total.total, total.aadhar, total.manual]
        for offset, value in enumerate(values):
            self.total(ws, total_row, start+offset, block["total_formats"][offset], value)
        if width == 4:
            rownum = total_row + 1
            ws.write_formula(total_row, start+3,
                             f"=IFERROR({xl_col_to_name(start+1)}{rownum}/{xl_col_to_name(start)}{rownum},100%)",
                             self.fmt(block["total_formats"][3]), total.compliance)
            self.conditional(ws, 2, total_row-1, start+2, "positive")
            self.conditional(ws, 2, total_row-1, start+3, "compliance")
        else:
            maximum = max((self.report.counts.get((site, day), Counts()).red for site in sites), default=0)
            self.conditional(ws, 2, total_row-1, start+1, "red_scale", maximum)

    def hyperverge(self, name, sites):
        ws, cfg, total_row = self.summary_setup(name, sites, 2)
        for i in range(len(self.report.dates)):
            self.metric_block(ws, cfg, sites, total_row, i, 2)

    def aadhar(self, name, hyper_name, sites):
        ws, cfg, total_row = self.summary_setup(name, sites, 4)
        for i in range(len(self.report.dates)):
            self.metric_block(ws, cfg, sites, total_row, i, 4)
        start = 1 + len(self.report.dates) * 4
        days = self.report.dates[-2:]
        block = cfg["hyper_summary"]
        if len(days) == 2:
            ws.merge_range(0, start, 0, start+1, "Hyperverge Red", self.fmt(block["date_formats"][0]))
        else:
            ws.write_string(0, start, "Hyperverge Red", self.fmt(block["date_formats"][0]))
        for offset, day in enumerate(days):
            col = start+offset
            ws.set_column(col, col, _column_width(block["widths"][offset]))
            day_index = len(self.report.dates)-len(days)+offset
            day_col = 1 + 4*day_index
            hyper_col = xl_col_to_name(2 + 2*day_index)
            ws.write_formula(1, col, f"={xl_col_to_name(day_col)}1", self.fmt(block["header_formats"][offset]), _excel_date_number(day))
            cached_total = 0
            maximum = 0
            for row, site in enumerate(sites, 2):
                value = self.report.counts.get((site, day), Counts()).red
                cached_total += value
                maximum = max(maximum, value)
                ws.write_formula(row, col, f"='{hyper_name}'!{hyper_col}{row+1}", self.fmt(block["body_formats"][offset]), value)
            self.total(ws, total_row, col, block["total_formats"][offset], cached_total)
            self.conditional(ws, 2, total_row-1, col, "red_scale", maximum)

    def hr(self, name, names, counts, criterion_col, criterion):
        ws, cfg = self.setup(name)
        ws.freeze_panes(1, 0)
        for col in range(2):
            ws.set_column(col, col, _column_width(cfg["widths"][col]))
            ws.write_string(0, col, cfg["headers"][col], self.fmt(cfg["header_formats"][col]))
        for row in range(len(names)+1):
            ws.set_row(row, cfg["row_heights"].get(str(row+1), cfg["default_row_height"]))
        for row, person in enumerate(names, 1):
            ws.write_string(row, 0, person, self.fmt(cfg["body_formats"][0]))
            formula = f'=COUNTIFS(\'Onboarding Base\'!$O:$O, $A{row+1}, \'Onboarding Base\'!${criterion_col}:${criterion_col}, "{criterion}")'
            ws.write_formula(row, 1, formula, self.fmt(cfg["body_formats"][1]), counts[person])
        if cfg.get("auto_filter"):
            # Preserve the template's filter setting without introducing tables.
            ws.autofilter(cfg["auto_filter"])

    def base(self):
        """Write raw values without decorative formatting or styled blanks.

        Date/time number formats are the only cell formats: Excel otherwise
        displays date serials as numbers. Types, order, and data are preserved.
        Column widths and the existing freeze pane are view conveniences only.
        """
        ws = self.sheets[BASE_SHEET]
        cfg = self.profile["sheets"][BASE_SHEET]
        ws.freeze_panes(1, 0)
        ws.hide_gridlines(0)
        for col, header in enumerate(self.profile["headers"]):
            ws.set_column(col, col, _column_width(cfg["widths"][col]))
            ws.write_string(0, col, header)
        observed = 0
        for row, values in enumerate(self.report.base_rows, 1):
            if len(values) != len(self.profile["headers"]):
                raise ReportError("An Onboarding Base row has an unexpected number of columns.")
            for col, value in enumerate(values):
                self.value(ws, row, col, value)
            observed += 1
        if observed != self.report.row_count:
            raise ReportError("The number of written Onboarding Base records does not match the filtered input.")


def write_report(output_path: Path, profile: dict, report: PreparedReport,
                 progress_callback: Callable | None = None) -> None:
    """Atomically create a fresh workbook; preserve an existing output on failure."""
    output_path = Path(output_path)
    if not report.dates:
        raise ReportError("The report must contain at least one calendar date.")
    if 1+4*len(report.dates)+min(2, len(report.dates)) > 16384:
        raise ReportError("The selected date range exceeds Excel's column limit.")
    if report.row_count > 1048575 or max(len(report.dc_sites), len(report.xd_sites)) > 1048573:
        raise ReportError("The report exceeds Excel's row limit. Select a smaller date range.")
    temp_path = None
    try:
        with NamedTemporaryFile(prefix=".onboarding-", suffix=".xlsx", dir=output_path.parent, delete=False) as handle:
            temp_path = Path(handle.name)
        with xlsxwriter.Workbook(temp_path, {"strings_to_formulas": False, "strings_to_urls": False}) as workbook:
            workbook.set_calc_mode("auto")
            workbook.set_properties({"comments": "Generated from a fresh workbook with XlsxWriter. Plain Base / positive-only RED scales / compact XLSX build 2026-10-24.3."})
            writer = _Writer(workbook, profile, report)
            def notify(percent, message):
                if progress_callback:
                    progress_callback(percent, message)
            notify(40, "Writing Hyperverge - DC...")
            writer.hyperverge(HYPERVERGE_DC, report.dc_sites)
            notify(50, "Writing Hyperverge - XD...")
            writer.hyperverge(HYPERVERGE_XD, report.xd_sites)
            notify(60, "Writing Aadhar - DC summary...")
            writer.aadhar(AADHAR_DC, HYPERVERGE_DC, report.dc_sites)
            notify(70, "Writing Aadhar - XD summary...")
            writer.aadhar(AADHAR_XD, HYPERVERGE_XD, report.xd_sites)
            notify(80, "Writing HR-wise detail sheets...")
            writer.hr(MANUAL_HR, report.manual_names, report.manual_counts, "P", "N")
            writer.hr(RED_HR, report.red_names, report.red_counts, "U", "red")
            notify(88, f"Writing Onboarding Base ({report.row_count:,} matching rows)...")
            writer.base()
            notify(94, "Saving the fresh Excel workbook...")
        notify(95, "Compressing the workbook without changing its data...")
        compact_workbook(temp_path)
        notify(97, "Checking Excel styles and worksheet references...")
        validate_workbook(temp_path)
        temp_path.replace(output_path)
    except (PermissionError, xlsxwriter.exceptions.FileCreateError) as exc:
        raise ReportError(f"Cannot save '{output_path.name}'. Close it in Excel and check the destination folder permissions.") from exc
    except ReportError:
        raise
    except Exception as exc:
        raise ReportError(f"Could not save the final report.\n\n{exc}") from exc
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()
