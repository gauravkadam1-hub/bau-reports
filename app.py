from __future__ import annotations

import tempfile
from datetime import date
from pathlib import Path

import streamlit as st

from portal.file_detector import detect_files, missing_requirements
from portal.facility_memory import (
    acknowledge_sites,
    load_report_preferences,
    memory_backend_status,
    set_many_site_states,
    set_site_state,
    site_preference,
)
from portal.generator import ReportRequest, generate_one, zip_outputs
from portal.scanner import scan_attendance, scan_flex, scan_onboarding, scan_roster
from portal.workspace import save_uploads

st.set_page_config(page_title="Reporting Automation Hub", page_icon="📊", layout="wide")

st.markdown(
    """
<style>
.block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 4rem;}
.hero {padding: 0.2rem 0 1.1rem 0;}
.hero h1 {margin-bottom: .15rem; font-size: 2.15rem;}
.hero p {color: #64748b; margin-top: 0; font-size: 1rem;}
.report-card {border: 1px solid #e2e8f0; border-radius: 14px; padding: 1rem 1.1rem; margin: .6rem 0; background: #fff;}
.ready {font-weight: 700; color: #15803d;}
.waiting {font-weight: 700; color: #64748b;}
.error {font-weight: 700; color: #b91c1c;}
.small {color:#64748b; font-size:.9rem;}
div.stButton > button[kind="primary"] {min-height: 3.15rem; font-size: 1.05rem; font-weight: 700; width: 100%;}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown('<div class="hero"><h1>Reporting Automation Hub</h1><p>Upload today\'s base files once. The system detects what can be generated and prepares the relevant controls.</p></div>', unsafe_allow_html=True)

_memory_status = memory_backend_status()
if _memory_status["mode"] == "github":
    st.caption("Facility memory: cloud-persistent storage configured ✓")
else:
    st.caption("Facility memory: local on this machine")

uploaded = st.file_uploader(
    "Upload base files",
    type=["csv", "xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
    help="Select all available base files in one go.",
)

REPORTS = [
    ("roster", "Roster Report"),
    ("flex", "Flex Adoption"),
    ("attendance", "Attendance Adherence"),
    ("onboarding", "Onboarding Report"),
]

if not uploaded:
    st.info("Select all available base files above. Report options will appear automatically.")
    st.stop()

names = [f.name for f in uploaded]
detected = detect_files(names)

with st.expander("Detected files", expanded=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Shared bases**")
        st.write(f"All Users: {len(detected.all_users)}")
        st.write(f"Attendance: {len(detected.attendance)}")
        st.write(f"Floater (optional): {len(detected.floater)}")
    with c2:
        st.markdown("**Report-specific bases**")
        st.write(f"Roster: {len(detected.roster)}")
        st.write(f"Flex: {len(detected.flex)}")
        st.write(f"Onboarding: {len(detected.onboarding)}")
    with c3:
        if detected.unknown:
            st.markdown("**Not recognised**")
            for n in detected.unknown:
                st.caption(n)
        if detected.ignored:
            st.markdown("**Ignored by design**")
            for n in detected.ignored:
                st.caption(n)

ready_keys = detected.ready_report_keys()
if not ready_keys:
    st.warning("No report has all of its required base files yet.")
    for key, label in REPORTS:
        miss = missing_requirements(detected, key)
        st.caption(f"{label}: missing {', '.join(miss)}")
    st.stop()

# Save uploads only for this Streamlit rerun while scanners inspect them.
# Generation gets its own isolated workspace later.
scan_holder = st.empty()
with tempfile.TemporaryDirectory(prefix="reporting_hub_scan_") as tmp:
    tmp_path = Path(tmp)
    saved = save_uploads(uploaded, tmp_path)
    scans = {
        "roster": scan_roster(detected, saved),
        "flex": scan_flex(detected, saved),
        "attendance": scan_attendance(detected, saved),
        "onboarding": scan_onboarding(detected, saved),
    }

st.subheader("Available reports & settings")

requests: dict[str, ReportRequest] = {}
progress_slots: dict[str, tuple[object, object]] = {}
valid_ready_count = 0


def _facility_widget_key(report_key: str, site: str) -> str:
    return f"facility::{report_key}::{site}"


def _remember_facility(report_key: str, site: str, widget_key: str) -> None:
    set_site_state(report_key, site, bool(st.session_state.get(widget_key, False)))


def _set_facility_group(report_key: str, sites: list[str], value: bool) -> None:
    set_many_site_states(report_key, sites, value)
    for site in sites:
        st.session_state[_facility_widget_key(report_key, site)] = value


def facility_picker(
    prefix: str,
    dc_sites: list[str],
    xd_sites: list[str],
    default_dc_sites: list[str],
    default_xd_sites: list[str],
) -> tuple[list[str], list[str]]:
    default_sites = default_dc_sites + default_xd_sites
    pref = load_report_preferences(prefix, default_sites)

    # Initialize Streamlit widget state from durable preferences only once per site/session.
    for site in dc_sites + xd_sites:
        widget_key = _facility_widget_key(prefix, site)
        if widget_key not in st.session_state:
            selected, _ = site_preference(prefix, site, pref, default_sites)
            st.session_state[widget_key] = selected

    selected_dc = [s for s in dc_sites if st.session_state.get(_facility_widget_key(prefix, s), False)]
    selected_xd = [s for s in xd_sites if st.session_state.get(_facility_widget_key(prefix, s), False)]

    col1, col2 = st.columns(2)
    for container, title, sites, selected_now in [
        (col1, "DC facilities", dc_sites, selected_dc),
        (col2, "XD facilities", xd_sites, selected_xd),
    ]:
        with container:
            st.markdown(f"**{title} — {len(selected_now)}/{len(sites)} selected**")
            b1, b2 = st.columns(2)
            b1.button(
                "Select all",
                key=f"{prefix}_{title}_select_all",
                use_container_width=True,
                on_click=_set_facility_group,
                args=(prefix, sites, True),
            )
            b2.button(
                "Clear all",
                key=f"{prefix}_{title}_clear_all",
                use_container_width=True,
                on_click=_set_facility_group,
                args=(prefix, sites, False),
            )

            with st.container(height=300, border=True):
                for site in sites:
                    widget_key = _facility_widget_key(prefix, site)
                    _, is_new = site_preference(prefix, site, pref, default_sites)
                    label = f"{site}  ·  NEW" if is_new else site
                    st.checkbox(
                        label,
                        key=widget_key,
                        on_change=_remember_facility,
                        args=(prefix, site, widget_key),
                    )

    # Read again after widget construction so the request always reflects current checkbox state.
    selected_dc = [s for s in dc_sites if st.session_state.get(_facility_widget_key(prefix, s), False)]
    selected_xd = [s for s in xd_sites if st.session_state.get(_facility_widget_key(prefix, s), False)]
    return selected_dc, selected_xd


for key, label in REPORTS:
    scan = scans[key]
    missing = missing_requirements(detected, key)
    with st.container(border=True):
        top_a, top_b = st.columns([4, 1])
        top_a.markdown(f"### {label}")
        if missing:
            top_b.markdown('<span class="waiting">NOT AVAILABLE</span>', unsafe_allow_html=True)
            st.caption("Missing: " + ", ".join(missing))
            continue
        if scan.error:
            top_b.markdown('<span class="error">CHECK INPUT</span>', unsafe_allow_html=True)
            st.error(scan.error)
            continue

        top_b.markdown('<span class="ready">READY</span>', unsafe_allow_html=True)
        valid_ready_count += 1

        d1, d2 = st.columns(2)
        with d1:
            report_from = st.date_input(
                "From date", value=scan.default_from, min_value=scan.min_date, max_value=scan.max_date,
                key=f"{key}_from",
            )
        with d2:
            report_to = st.date_input(
                "To date", value=scan.default_to, min_value=scan.min_date, max_value=scan.max_date,
                key=f"{key}_to",
            )

        selected_dc: list[str] = []
        selected_xd: list[str] = []
        selected_reports: list[str] | None = None

        if key in {"roster", "flex", "attendance", "onboarding"}:
            selected_dc, selected_xd = facility_picker(
                key, scan.dc_sites, scan.xd_sites, scan.default_dc_sites, scan.default_xd_sites
            )
            if key == "onboarding":
                st.caption(
                    "Existing master-template facilities keep their remembered state. "
                    "New source facilities appear as NEW and start unchecked until you select them."
                )

        if key == "attendance":
            labels = ["Main Attendance"] + scan.dedicated_roles
            selected_labels = st.multiselect(
                "Attendance reports",
                labels,
                default=labels,
                key="attendance_reports",
                help="Dedicated roles come from the deployed master architecture.",
            )
            # The existing Attendance engine expects its private token for Main, while
            # dedicated reports use the real role labels from Attendance/master data.
            selected_reports = [
                "__main__" if label == "Main Attendance" else label
                for label in selected_labels
            ]

        requests[key] = ReportRequest(
            key=key,
            report_from=report_from,
            report_to=report_to,
            selected_dc=selected_dc,
            selected_xd=selected_xd,
            selected_reports=selected_reports,
        )
        bar = st.progress(0, text="Waiting")
        status = st.empty()
        progress_slots[key] = (bar, status)

st.divider()
st.subheader("Generate all available reports")
overall_bar = st.progress(0, text="Overall progress: 0%")
overall_status = st.empty()

if "report_zip" not in st.session_state:
    st.session_state.report_zip = None
    st.session_state.report_files = []
    st.session_state.run_summary = []

can_generate = valid_ready_count > 0
if st.button("GENERATE ALL REPORTS", type="primary", disabled=not can_generate, use_container_width=True):
    st.session_state.report_zip = None
    st.session_state.report_files = []
    st.session_state.run_summary = []

    ordered = [key for key, _ in REPORTS if key in requests]
    completed_modules = 0
    output_paths: list[Path] = []

    with tempfile.TemporaryDirectory(prefix="reporting_hub_run_") as tmp:
        run_root = Path(tmp)
        saved = save_uploads(uploaded, run_root / "uploads")

        for key in ordered:
            label = dict(REPORTS)[key]
            bar, status = progress_slots[key]
            status.info(f"Generating {label}...")

            def update_module(pct: int, message: str, *, _key=key, _bar=bar, _status=status):
                pct = max(0, min(100, int(pct)))
                _bar.progress(pct, text=message)
                base = completed_modules / len(ordered)
                current = (pct / 100) / len(ordered)
                overall_pct = int(round((base + current) * 100))
                overall_bar.progress(min(overall_pct, 99), text=f"Overall progress: {min(overall_pct, 99)}%")

            try:
                if key == "attendance":
                    acknowledge_sites(key, scans[key].dc_sites + scans[key].xd_sites)

                files = generate_one(
                    requests[key], detected, saved, run_root / "outputs" / key, update_module
                )
                output_paths.extend(files)

                # Flex acknowledges all discovered sites after a successful run.
                if key == "flex":
                    acknowledge_sites(key, scans[key].dc_sites + scans[key].xd_sites)

                # Onboarding mirrors its desktop new-site behavior: only facilities
                # the user actually selected become known. Unticked new sites remain
                # NEW and are offered again on a future run.
                if key == "onboarding":
                    req = requests[key]
                    acknowledge_sites(key, req.selected_dc + req.selected_xd)
                bar.progress(100, text="Completed")
                status.success(f"{label} completed")
                st.session_state.run_summary.append((label, "success", ""))
            except Exception as exc:
                bar.progress(100, text="Processing finished with an error")
                status.error(str(exc))
                st.session_state.run_summary.append((label, "failed", str(exc)))
            finally:
                completed_modules += 1
                pct = int(round(completed_modules / len(ordered) * 100))
                overall_bar.progress(pct, text=f"Overall progress: {pct}%")

        if output_paths:
            st.session_state.report_zip = zip_outputs(output_paths)
            st.session_state.report_files = [p.name for p in output_paths]

    overall_status.success("Processing complete")

# Final download is deliberately visible from the start but disabled until a package exists.
download_disabled = st.session_state.report_zip is None
if download_disabled:
    st.button("DOWNLOAD ALL REPORTS", disabled=True, use_container_width=True)
else:
    st.download_button(
        "DOWNLOAD ALL REPORTS",
        data=st.session_state.report_zip,
        file_name=f"Reporting Package - {date.today():%d %b %Y}.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )
    if st.session_state.report_files:
        with st.expander("Files in package"):
            for name in st.session_state.report_files:
                st.write("✓", name)
