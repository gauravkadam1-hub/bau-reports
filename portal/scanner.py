from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from .engine_loader import attendance_engine, flex_engine, onboarding_engine, roster_engine
from .models import DetectionResult, ReportScan


def _paths(saved: dict[str, Path], names: list[str]) -> list[str]:
    return [str(saved[name]) for name in names]


def scan_roster(d: DetectionResult, saved: dict[str, Path]) -> ReportScan:
    if not d.roster_ready:
        return ReportScan("roster", "Roster Report", False)
    try:
        engine = roster_engine()
        result = engine.scan_inputs(str(saved[d.roster[0]]), str(saved[d.attendance[0]]))
        yesterday = date.today() - timedelta(days=1)
        default_to = min(result.common_dates[-1], yesterday)
        if default_to < result.common_dates[0]:
            default_to = result.common_dates[-1]
        return ReportScan(
            "roster", "Roster Report", True,
            min_date=result.common_dates[0], max_date=result.common_dates[-1],
            default_from=result.common_dates[0], default_to=default_to,
            dc_sites=list(result.dc_sites), xd_sites=list(result.xd_sites),
            default_dc_sites=list(engine.DEFAULT_DC_SITES),
            default_xd_sites=list(engine.DEFAULT_XD_SITES),
        )
    except Exception as exc:
        return ReportScan("roster", "Roster Report", True, error=str(exc))


def scan_flex(d: DetectionResult, saved: dict[str, Path]) -> ReportScan:
    if not d.flex_ready:
        return ReportScan("flex", "Flex Adoption", False)
    try:
        engine = flex_engine()
        today_max = date.today() - timedelta(days=2)
        default_from = today_max.replace(day=1)
        default_to = today_max
        u1, u2 = _paths(saved, d.all_users)
        result = engine.scan_inputs(
            str(saved[d.flex[0]]), u1, u2, str(saved[d.attendance[0]]),
            default_from, default_to,
        )
        return ReportScan(
            "flex", "Flex Adoption", True,
            min_date=result.report_start, max_date=result.report_end,
            default_from=result.report_start, default_to=result.report_end,
            dc_sites=list(result.dc_sites), xd_sites=list(result.xd_sites),
            default_dc_sites=list(engine.DEFAULT_DC_SITES),
            default_xd_sites=list(engine.DEFAULT_XD_SITES),
        )
    except Exception as exc:
        return ReportScan("flex", "Flex Adoption", True, error=str(exc))


def scan_attendance(d: DetectionResult, saved: dict[str, Path]) -> ReportScan:
    if not d.attendance_ready:
        return ReportScan("attendance", "Attendance Adherence", False)
    try:
        engine = attendance_engine()
        end = date.today() - timedelta(days=1)
        start = end.replace(day=1)
        u1, u2 = _paths(saved, d.all_users)
        result = engine.scan_inputs(u1, u2, str(saved[d.attendance[0]]), start, end, None)
        return ReportScan(
            "attendance", "Attendance Adherence", True,
            min_date=result.attendance_start, max_date=result.attendance_end,
            default_from=start, default_to=min(end, result.attendance_end),
            dc_sites=list(result.dc_sites), xd_sites=list(result.xd_sites),
            default_dc_sites=list(engine.DEFAULT_DC_SITES),
            default_xd_sites=list(engine.DEFAULT_XD_SITES),
            roles=list(result.roles), dedicated_roles=list(result.dedicated_roles),
        )
    except Exception as exc:
        return ReportScan("attendance", "Attendance Adherence", True, error=str(exc))


def scan_onboarding(d: DetectionResult, saved: dict[str, Path]) -> ReportScan:
    if not d.onboarding_ready:
        return ReportScan("onboarding", "Onboarding Report", False)
    try:
        engine = onboarding_engine()
        first, latest = engine.inspect_input_date_range(str(saved[d.onboarding[0]]))
        facilities = engine.inspect_facilities(
            str(saved[d.onboarding[0]]),
            selected_start_date=first,
            selected_end_date=latest,
        )
        return ReportScan(
            "onboarding", "Onboarding Report", True,
            min_date=first, max_date=latest, default_from=first, default_to=latest,
            dc_sites=list(facilities["DC"]), xd_sites=list(facilities["XD"]),
            default_dc_sites=list(facilities["default_DC"]),
            default_xd_sites=list(facilities["default_XD"]),
        )
    except Exception as exc:
        return ReportScan("onboarding", "Onboarding Report", True, error=str(exc))
