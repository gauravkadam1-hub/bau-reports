from __future__ import annotations

import base64
import json
import os
import re
from pathlib import Path
from typing import Iterable

import requests

STATE_VERSION = 1
_REMOTE_LAST_ERROR = ""
_REMOTE_LAST_OK = False


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


def _normalize_store(data: dict | None) -> dict:
    data = data if isinstance(data, dict) else _empty_store()
    if not isinstance(data.get("reports"), dict):
        data["reports"] = {}
    data["version"] = STATE_VERSION
    return data


def _remote_config() -> dict | None:
    """Return GitHub state-repo settings from Streamlit secrets or environment variables.

    Streamlit Cloud should use secrets. Environment variables are also supported so the
    exact same backend can be tested locally without committing credentials.
    """
    repo = os.environ.get("REPORTING_HUB_STATE_REPO", "").strip()
    token = os.environ.get("REPORTING_HUB_STATE_TOKEN", "").strip()
    branch = os.environ.get("REPORTING_HUB_STATE_BRANCH", "main").strip() or "main"
    path = os.environ.get("REPORTING_HUB_STATE_FILE", "facility_preferences.json").strip() or "facility_preferences.json"

    if not (repo and token):
        try:
            import streamlit as st

            section = st.secrets.get("github_state", {})
            repo = str(section.get("repo", "")).strip()
            token = str(section.get("token", "")).strip()
            branch = str(section.get("branch", "main")).strip() or "main"
            path = str(section.get("path", "facility_preferences.json")).strip() or "facility_preferences.json"
        except Exception:
            return None

    if not (repo and token):
        return None
    if "/" not in repo:
        return None

    return {"repo": repo, "token": token, "branch": branch, "path": path}


def _github_headers(token: str) -> dict[str, str]:
    return {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "ReportingAutomationHub/1.0",
    }


def _remote_url(cfg: dict) -> str:
    # Repository and path are configured by the app owner, never by report uploads.
    repo = cfg["repo"].strip("/")
    path = cfg["path"].lstrip("/")
    return f"https://api.github.com/repos/{repo}/contents/{path}"


def _read_remote() -> tuple[dict | None, str | None, bool]:
    """Return (store, sha, exists). Raises on auth/network/API failures."""
    cfg = _remote_config()
    if not cfg:
        return None, None, False

    response = requests.get(
        _remote_url(cfg),
        headers=_github_headers(cfg["token"]),
        params={"ref": cfg["branch"]},
        timeout=15,
    )
    if response.status_code == 404:
        return None, None, False
    response.raise_for_status()
    payload = response.json()
    raw = base64.b64decode(payload.get("content", "")).decode("utf-8")
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("GitHub facility state file does not contain a JSON object.")
    return _normalize_store(data), payload.get("sha"), True


def _write_remote(data: dict) -> None:
    """Write the complete preference store to the private GitHub state repository.

    We refresh the current SHA before every write so normal edits from another browser
    session do not blindly overwrite an older repository version.
    """
    cfg = _remote_config()
    if not cfg:
        return

    _, sha, exists = _read_remote()
    encoded = base64.b64encode(
        json.dumps(_normalize_store(data), indent=2, ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    payload = {
        "message": "Update facility preferences",
        "content": encoded,
        "branch": cfg["branch"],
    }
    if exists and sha:
        payload["sha"] = sha

    response = requests.put(
        _remote_url(cfg),
        headers=_github_headers(cfg["token"]),
        json=payload,
        timeout=15,
    )
    response.raise_for_status()


def _load_store() -> dict:
    global _REMOTE_LAST_ERROR, _REMOTE_LAST_OK

    local = _normalize_store(_read_json(settings_path()))
    cfg = _remote_config()
    if not cfg:
        _REMOTE_LAST_ERROR = ""
        _REMOTE_LAST_OK = False
        return local

    try:
        remote, _, exists = _read_remote()
        if exists and remote is not None:
            _REMOTE_LAST_ERROR = ""
            _REMOTE_LAST_OK = True
            # Keep a local cache/fallback too. This is useful during transient GitHub outages.
            _save_local(remote)
            return remote

        # State repo is configured but the state file is absent. Seed it from local
        # state if available; otherwise the normal default-site initialization below
        # will populate it on the first report preference save.
        if local.get("reports"):
            _write_remote(local)
            _REMOTE_LAST_OK = True
            _REMOTE_LAST_ERROR = ""
        return local
    except Exception as exc:
        _REMOTE_LAST_OK = False
        _REMOTE_LAST_ERROR = str(exc)
        return local


def _save_local(data: dict) -> None:
    path = settings_path()
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(_normalize_store(data), indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


def _save_store(data: dict) -> None:
    global _REMOTE_LAST_ERROR, _REMOTE_LAST_OK

    data = _normalize_store(data)
    # Always maintain the local copy. Locally it is the durable backend; in cloud it is
    # a useful cache only and the GitHub state repo is authoritative across restarts.
    _save_local(data)

    if not _remote_config():
        _REMOTE_LAST_ERROR = ""
        _REMOTE_LAST_OK = False
        return

    try:
        _write_remote(data)
        _REMOTE_LAST_OK = True
        _REMOTE_LAST_ERROR = ""
    except Exception as exc:
        _REMOTE_LAST_OK = False
        _REMOTE_LAST_ERROR = str(exc)


def memory_backend_status() -> dict[str, object]:
    cfg = _remote_config()
    return {
        "mode": "github" if cfg else "local",
        "repo": cfg["repo"] if cfg else "",
        "path": cfg["path"] if cfg else str(settings_path()),
        "last_remote_ok": _REMOTE_LAST_OK,
        "last_error": _REMOTE_LAST_ERROR,
    }


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
