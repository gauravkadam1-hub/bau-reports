from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Iterable

STATE_VERSION = 1


def _state_dir() -> Path:
    override = os.environ.get("REPORTING_HUB_STATE_DIR", "").strip()
    if override:
        base = Path(override).expanduser()
    elif os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home())) / "ReportingAutomationHub"
    else:
        base = Path.home() / ".reporting_automation_hub"
    base.mkdir(parents=True, exist_ok=True)
    return base


def settings_path() -> Path:
    return _state_dir() / "facility_preferences.json"


def _empty_store() -> dict:
    return {"version": STATE_VERSION, "reports": {}}


def _read_json(path: Path) -> dict | None:
    try:
        if path.exists():
            value = json.loads(path.read_text(encoding="utf-8"))
            return value if isinstance(value, dict) else None
    except Exception:
        pass
    return None


def _load_store() -> dict:
    data = _read_json(settings_path()) or _empty_store()
    if not isinstance(data.get("reports"), dict):
        data["reports"] = {}
    data["version"] = STATE_VERSION
    return data


def _save_store(data: dict) -> None:
    path = settings_path()
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


def normalize_site(report_key: str, value: object) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip())
    if report_key in {"flex", "attendance", "onboarding"}:
        return text.casefold()
    return text


def _default_site_keys(report_key: str, default_sites: Iterable[str]) -> set[str]:
    return {normalize_site(report_key, site) for site in default_sites if normalize_site(report_key, site)}


def _legacy_path(report_key: str) -> Path | None:
    if report_key == "roster":
        if os.name == "nt":
            return Path(os.environ.get("APPDATA", Path.home())) / "RosterAutomation" / "settings.json"
        return Path.home() / ".roster_automation" / "settings.json"
    if report_key == "flex":
        if os.name == "nt":
            return Path(os.environ.get("APPDATA", Path.home())) / "FlexAutomation" / "settings.json"
        return Path.home() / ".flex_automation" / "settings.json"
    if report_key == "attendance":
        return Path.home() / ".attendance_adherence_automation" / "settings.json"
    return None


def _import_legacy(report_key: str) -> dict | None:
    path = _legacy_path(report_key)
    saved = _read_json(path) if path else None
    if not saved:
        return None

    if report_key in {"roster", "flex"}:
        raw_states = saved.get("site_states", {})
        raw_known = saved.get("acknowledged_sites", [])
    else:
        raw_states = saved.get("selected_sites", {})
        raw_known = saved.get("known_sites", [])

    if not isinstance(raw_states, dict):
        raw_states = {}
    if not isinstance(raw_known, list):
        raw_known = []

    states = {
        normalize_site(report_key, site): bool(value)
        for site, value in raw_states.items()
        if normalize_site(report_key, site)
    }
    known = sorted(
        {
            normalize_site(report_key, site)
            for site in raw_known
            if normalize_site(report_key, site)
        }
    )
    return {"site_states": states, "known_sites": known, "legacy_imported": True}


def load_report_preferences(report_key: str, default_sites: Iterable[str] = ()) -> dict:
    store = _load_store()
    reports = store.setdefault("reports", {})
    pref = reports.get(report_key)
    if not isinstance(pref, dict):
        pref = _import_legacy(report_key)
        if pref is None:
            defaults = sorted(_default_site_keys(report_key, default_sites))
            pref = {
                "site_states": {site: True for site in defaults},
                "known_sites": defaults,
                "legacy_imported": False,
            }
        reports[report_key] = pref
        _save_store(store)

    if not isinstance(pref.get("site_states"), dict):
        pref["site_states"] = {}
    if not isinstance(pref.get("known_sites"), list):
        pref["known_sites"] = []
    return pref


def _save_report_preferences(report_key: str, pref: dict) -> None:
    store = _load_store()
    store.setdefault("reports", {})[report_key] = pref
    _save_store(store)


def site_preference(
    report_key: str,
    site: str,
    pref: dict | None = None,
    default_sites: Iterable[str] = (),
) -> tuple[bool, bool]:
    pref = pref or load_report_preferences(report_key, default_sites)
    key = normalize_site(report_key, site)
    states = pref.get("site_states", {})
    known = set(pref.get("known_sites", []))
    defaults = _default_site_keys(report_key, default_sites)

    if key in states:
        selected = bool(states[key])
    elif key in defaults:
        selected = True
    else:
        selected = False

    is_new = key not in known and key not in defaults
    return selected, is_new


def set_site_state(report_key: str, site: str, selected: bool) -> None:
    pref = load_report_preferences(report_key)
    key = normalize_site(report_key, site)
    pref.setdefault("site_states", {})[key] = bool(selected)

    # Roster's desktop behavior acknowledges a new site as soon as it is selected.
    if report_key == "roster" and selected:
        known = set(pref.setdefault("known_sites", []))
        known.add(key)
        pref["known_sites"] = sorted(known)

    _save_report_preferences(report_key, pref)


def set_many_site_states(report_key: str, sites: Iterable[str], selected: bool) -> None:
    pref = load_report_preferences(report_key)
    states = pref.setdefault("site_states", {})
    known = set(pref.setdefault("known_sites", []))
    for site in sites:
        key = normalize_site(report_key, site)
        states[key] = bool(selected)
        if report_key == "roster" and selected:
            known.add(key)
    pref["known_sites"] = sorted(known)
    _save_report_preferences(report_key, pref)


def acknowledge_sites(report_key: str, sites: Iterable[str]) -> None:
    pref = load_report_preferences(report_key)
    known = set(pref.setdefault("known_sites", []))
    known.update(normalize_site(report_key, site) for site in sites)
    pref["known_sites"] = sorted(site for site in known if site)
    _save_report_preferences(report_key, pref)
