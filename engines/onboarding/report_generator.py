"""Onboarding report: Polars calculations + a fresh XlsxWriter workbook.

Public call signatures are kept compatible with the existing Streamlit portal.
The old master_template.xlsx is not opened, modified, or copied at runtime.
"""
from __future__ import annotations

import json
from importlib import import_module
from datetime import date
from pathlib import Path
from tempfile import NamedTemporaryFile

try:
    from .onboarding_common import (
        AADHAR_DC, AADHAR_XD, BASE_SHEET, HYPERVERGE_DC, HYPERVERGE_XD,
        MANUAL_HR, RED_HR, REQUIRED_TEMPLATE_SHEETS,
        clean_facility, clean_selection, facility_type, load_profile, profile_path,
    )
    from .onboarding_writer import write_report
    from .utils import ReportError, output_filename
except ImportError:
    from onboarding_common import (
        AADHAR_DC, AADHAR_XD, BASE_SHEET, HYPERVERGE_DC, HYPERVERGE_XD,
        MANUAL_HR, RED_HR, REQUIRED_TEMPLATE_SHEETS,
        clean_facility, clean_selection, facility_type, load_profile, profile_path,
    )
    from onboarding_writer import write_report
    from utils import ReportError, output_filename


# Backward-compatible names for callers inspecting facility suffixes.
_facility_type = facility_type
_clean_facility_name = clean_facility


def _data():
    try:
        module_name = f"{__package__}.onboarding_data" if __package__ else "onboarding_data"
        return import_module(module_name)
    except ModuleNotFoundError as exc:
        if exc.name == "polars":
            raise ReportError("Onboarding now requires Polars. Install the updated requirements.txt and restart the app.") from exc
        raise


def inspect_input_date_range(input_path, template_path=None):
    profile = load_profile(template_path)
    data = _data()
    return data.available_dates(data.load_input(input_path, profile))


def inspect_facilities(input_path, selected_start_date=None, selected_end_date=None, template_path=None):
    profile, data = load_profile(template_path), _data()
    frame = data.load_input(input_path, profile)
    first, last = data.selected_dates(frame, selected_start_date, selected_end_date)
    source = data.source_facilities(frame, first, last)
    defaults = profile["defaults"]
    return {
        "DC": clean_selection(defaults["DC"]+source["DC"], "DC"),
        "XD": clean_selection(defaults["XD"]+source["XD"], "XD"),
        "default_DC": list(defaults["DC"]), "default_XD": list(defaults["XD"]),
    }


def inspect_new_facilities(input_path, selected_start_date, selected_end_date, template_path=None):
    result = inspect_facilities(input_path, selected_start_date, selected_end_date, template_path)
    return {kind: sorted(
        [name for name in result[kind] if name.casefold() not in {s.casefold() for s in result['default_'+kind]}],
        key=str.casefold) for kind in ("DC", "XD")}


def add_new_facilities_to_master(facilities, template_path=None):
    """Compatibility API: persist approved defaults in JSON, never edit an XLSX.

    Existing portal facility-selection memory remains unchanged. This legacy API
    is retained for integrations which explicitly add new permanent defaults.
    """
    profile = load_profile(template_path)
    target = profile_path(template_path).with_name("onboarding_facility_defaults.json")
    values = {kind: clean_selection(profile["defaults"][kind]+list(facilities.get(kind, [])), kind) for kind in ("DC", "XD")}
    temporary = None
    try:
        with NamedTemporaryFile(mode="w", encoding="utf-8", dir=target.parent, prefix=".onboarding-defaults-", suffix=".json", delete=False) as handle:
            temporary = Path(handle.name)
            json.dump(values, handle, indent=2, ensure_ascii=False)
            handle.write("\n")
        temporary.replace(target)
    except OSError as exc:
        raise ReportError("The Onboarding facility defaults could not be saved. Check app-folder permissions.") from exc
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def generate_report(input_path, template_path=None, selected_start_date=None, selected_end_date=None,
                    selected_dc=None, selected_xd=None, output_dir=None, progress_callback=None,
                    generation_date=None):
    """Generate the same daily report using Polars and a fresh XlsxWriter file.

    Returns the existing (Path, selected start date, selected end date) tuple.
    """
    input_path = Path(input_path).expanduser().resolve()
    if not input_path.exists():
        raise ReportError("Please select an existing onboarding export.")
    destination = input_path.parent if output_dir is None else Path(output_dir).expanduser().resolve()
    if not destination.is_dir():
        raise ReportError("The selected destination folder does not exist.")
    output = destination / output_filename(generation_date or date.today())
    if output == input_path:
        raise ReportError("The output path is the same as the uploaded source file. Please choose a different destination folder.")
    def notify(percent, message):
        if progress_callback:
            progress_callback(int(percent), message)
    notify(3, "Loading Onboarding layout settings...")
    profile, data = load_profile(template_path), _data()
    notify(7, "Reading the onboarding export with Polars...")
    frame = data.load_input(input_path, profile)
    notify(10, "Validating date range...")
    first, last = data.selected_dates(frame, selected_start_date, selected_end_date, enforce_bounds=True)
    notify(12, "Filtering and calculating Onboarding counts with Polars...")
    report = data.prepare_report(frame, profile, first, last, selected_dc, selected_xd)
    notify(35, "Building a fresh Excel workbook...")
    write_report(output, profile, report, progress_callback)
    notify(100, "Report generated successfully.")
    return output, first, last
