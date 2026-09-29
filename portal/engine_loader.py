from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

ROOT = Path(__file__).resolve().parents[1]


def _load(unique_name: str, file_path: Path, prepend_path: Path | None = None) -> ModuleType:
    if unique_name in sys.modules:
        return sys.modules[unique_name]
    if prepend_path is not None:
        p = str(prepend_path)
        if p not in sys.path:
            sys.path.insert(0, p)
    spec = importlib.util.spec_from_file_location(unique_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not load {file_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[unique_name] = module
    spec.loader.exec_module(module)
    return module


def roster_engine() -> ModuleType:
    return _load("reporting_hub_roster_engine", ROOT / "engines/roster/engine.py")


def flex_engine() -> ModuleType:
    return _load("reporting_hub_flex_engine", ROOT / "engines/flex/engine.py")


def attendance_engine() -> ModuleType:
    return _load("reporting_hub_attendance_engine", ROOT / "engines/attendance/engine.py")


def onboarding_engine() -> ModuleType:
    folder = ROOT / "engines/onboarding"
    return _load("reporting_hub_onboarding_generator", folder / "report_generator.py", prepend_path=folder)
