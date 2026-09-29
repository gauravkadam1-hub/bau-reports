from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class UploadedFileRef:
    """Filename-level information only; bytes remain owned by Streamlit."""

    name: str
    category: str


@dataclass
class DetectionResult:
    roster: list[str] = field(default_factory=list)
    flex: list[str] = field(default_factory=list)
    all_users: list[str] = field(default_factory=list)
    attendance: list[str] = field(default_factory=list)
    onboarding: list[str] = field(default_factory=list)
    ignored: list[str] = field(default_factory=list)
    unknown: list[str] = field(default_factory=list)

    @property
    def roster_ready(self) -> bool:
        return len(self.roster) == 1 and len(self.attendance) == 1

    @property
    def flex_ready(self) -> bool:
        return len(self.flex) == 1 and len(self.all_users) == 2 and len(self.attendance) == 1

    @property
    def attendance_ready(self) -> bool:
        return len(self.all_users) == 2 and len(self.attendance) == 1

    @property
    def onboarding_ready(self) -> bool:
        return len(self.onboarding) == 1

    def ready_report_keys(self) -> list[str]:
        out: list[str] = []
        if self.roster_ready:
            out.append("roster")
        if self.flex_ready:
            out.append("flex")
        if self.attendance_ready:
            out.append("attendance")
        if self.onboarding_ready:
            out.append("onboarding")
        return out

    def all_classified_names(self) -> set[str]:
        groups: Iterable[list[str]] = (
            self.roster,
            self.flex,
            self.all_users,
            self.attendance,
            self.onboarding,
            self.ignored,
            self.unknown,
        )
        return {name for group in groups for name in group}


@dataclass
class ReportScan:
    key: str
    display_name: str
    ready: bool
    error: str | None = None
    min_date: object | None = None
    max_date: object | None = None
    default_from: object | None = None
    default_to: object | None = None
    dc_sites: list[str] = field(default_factory=list)
    xd_sites: list[str] = field(default_factory=list)
    default_dc_sites: list[str] = field(default_factory=list)
    default_xd_sites: list[str] = field(default_factory=list)
    roles: list[str] = field(default_factory=list)
    dedicated_roles: list[str] = field(default_factory=list)
