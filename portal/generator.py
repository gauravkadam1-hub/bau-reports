from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Callable

from .engine_loader import attendance_engine, flex_engine, onboarding_engine, roster_engine
from .models import DetectionResult

ProgressCallback = Callable[[int, str], None]


@dataclass
class ReportRequest:
    key: str
    report_from: date
    report_to: date
    selected_dc: list[str]
    selected_xd: list[str]
    selected_reports: list[str] | None = None


def generate_one(
    request: ReportRequest,
    detected: DetectionResult,
    saved: dict[str, Path],
    output_folder: Path,
    progress: ProgressCallback,
) -> list[Path]:
    output_folder.mkdir(parents=True, exist_ok=True)
    key = request.key

    if key == "roster":
        e = roster_engine()
        inputs = e.ReportInputs(
            roster_path=str(saved[detected.roster[0]]),
            attendance_path=str(saved[detected.attendance[0]]),
            output_folder=str(output_folder),
            report_from=request.report_from,
            report_to=request.report_to,
            selected_dc=request.selected_dc,
            selected_xd=request.selected_xd,
        )
        return [Path(e.generate_report(inputs, progress))]

    if key == "flex":
        e = flex_engine()
        u1, u2 = [str(saved[n]) for n in detected.all_users]
        inputs = e.ReportInputs(
            flex_path=str(saved[detected.flex[0]]),
            users_path_1=u1,
            users_path_2=u2,
            attendance_path=str(saved[detected.attendance[0]]),
            output_folder=str(output_folder),
            report_start=request.report_from,
            report_end=request.report_to,
            selected_dc=request.selected_dc,
            selected_xd=request.selected_xd,
        )
        return [Path(e.generate_report(inputs, progress))]

    if key == "attendance":
        e = attendance_engine()
        u1, u2 = [str(saved[n]) for n in detected.all_users]
        if len(detected.floater) > 1:
            raise ValueError("Upload at most one Floater/Floaters file for Attendance.")
        floater_path = str(saved[detected.floater[0]]) if detected.floater else None
        inputs = e.ReportInputs(
            users_path_1=u1,
            users_path_2=u2,
            attendance_path=str(saved[detected.attendance[0]]),
            output_folder=str(output_folder),
            report_start=request.report_from,
            report_end=request.report_to,
            selected_dc=request.selected_dc,
            selected_xd=request.selected_xd,
            selected_reports=request.selected_reports,
            floater_path=floater_path,
        )
        return [Path(p) for p in e.generate_reports(inputs, progress)]

    if key == "onboarding":
        e = onboarding_engine()
        output, _, _ = e.generate_report(
            str(saved[detected.onboarding[0]]),
            template_path=str(Path(e.__file__).with_name("master_template.xlsx")),
            selected_start_date=request.report_from,
            selected_end_date=request.report_to,
            selected_dc=request.selected_dc,
            selected_xd=request.selected_xd,
            output_dir=str(output_folder),
            progress_callback=progress,
        )
        return [Path(output)]

    raise ValueError(f"Unknown report key: {key}")


def zip_outputs(files: list[Path]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        seen: set[str] = set()
        for path in files:
            arcname = path.name
            if arcname in seen:
                stem, suffix = path.stem, path.suffix
                i = 2
                while f"{stem} ({i}){suffix}" in seen:
                    i += 1
                arcname = f"{stem} ({i}){suffix}"
            seen.add(arcname)
            zf.write(path, arcname)
    return buf.getvalue()
