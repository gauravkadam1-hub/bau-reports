> Latest build: **2026-10-24.5 DC/XD-only base scope**. See [DCXD_BASE_SCOPE_V5.md](DCXD_BASE_SCOPE_V5.md) for the exact filter, installation, measurements and verification limits.

> Onboarding update: **build 2026-10-24.3**. See [ONBOARDING_COMPACT_V3.md](ONBOARDING_COMPACT_V3.md) for plain Base, zero-safe RED scales, measured file-size savings, and installation.

> Onboarding correction build **2026-10-24.2**: see [ONBOARDING_STYLES_FIX.md](ONBOARDING_STYLES_FIX.md) before upgrading.

# Reporting Automation Hub

**Onboarding update (24 October 2026):** Onboarding now uses Polars for processing and creates a fresh workbook with XlsxWriter. Read [ONBOARDING_MIGRATION.md](ONBOARDING_MIGRATION.md) for changes, installation, verification results, and the outstanding real-backend/Excel acceptance checks.

Phase-1 architecture for a single internal portal that:

1. accepts all daily base files in one multi-file upload;
2. identifies bases from stable filename patterns;
3. determines which existing reports have enough inputs;
4. calls each report engine's own scanner to expose its date/facility/role settings;
5. generates available reports serially using the existing engines;
6. shows per-report and overall progress; and
7. unlocks one ZIP download only after processing has completed.

## Filename recognition

- `roster_all_facilities_*` -> Roster Base
- `flex*` -> Flex Base
- `all_user_data_all_*` -> All Users (current portal build expects the existing two-file split for Flex/Attendance)
- `attendance_report_*` -> Attendance Base
- `onboarding_funnel_all_*` -> Onboarding Base
- `All user data base.csv` is explicitly ignored.

Dates/timestamps in filenames may change without affecting recognition.

## Existing engines

Streamlit adapters invoke each engine's public scan/generate functions. Attendance, Flex, and Roster retain the supplied engine implementations. Onboarding has been migrated to Polars + XlsxWriter while preserving its calculation rules, public API, and report layout; its reference XLSX is no longer loaded at runtime. See `ONBOARDING_MIGRATION.md` for the exact scope and verification limits.

## Local start

```bash
python -m venv .venv
# activate it
pip install -r requirements.txt
streamlit run app.py
```

For deployment, keep the Git repository private and deploy the single app from that repository.

## Facility selection memory

The portal keeps separate remembered facility selections for Roster, Flex, Attendance, and Onboarding.

- Checked/unchecked site choices survive normal local restarts.
- Existing desktop settings are imported automatically on first use when their settings files are available on the same Windows user account.
- Newly discovered non-reference sites start unchecked and are labelled `NEW`.
- Roster clears `NEW` when the new site is selected, matching the desktop app.
- Flex clears `NEW` after a successful Flex generation, matching the desktop app.
- Attendance records the current discovered site list when generation starts, matching the desktop app.
- Onboarding shows every current layout-default facility as a checkbox. New source-only facilities start unchecked and labelled `NEW`; selected new facilities become known after a successful generation, while unticked new facilities remain `NEW` for a future run.
- Onboarding applies chosen facilities to both Aadhar and Hyperverge DC/XD pairs, preserving paired ordering and existing sorting logic.

### Local backend

Without cloud secrets the portal stores preferences at:

- Windows: `%APPDATA%\\ReportingAutomationHub\\facility_preferences.json`
- macOS/Linux: `~/.reporting_automation_hub/facility_preferences.json`

### Streamlit Cloud persistent backend

For Community Cloud, use a second private GitHub repository only for the tiny preference JSON. No uploaded employee files or generated reports are written to GitHub.

1. Create a private repository, for example `bau-reports-state`.
2. Put `deployment/state_repo_seed/facility_preferences.json` at the repository root and commit it to `main`.
3. Create a fine-grained GitHub personal access token scoped only to that state repository, with **Contents: Read and write** permission.
4. In Streamlit Community Cloud open the app's **Settings -> Secrets** and add:

```toml
[github_state]
repo = "YOUR_GITHUB_USERNAME/bau-reports-state"
token = "YOUR_FINE_GRAINED_TOKEN"
branch = "main"
path = "facility_preferences.json"
```

5. Save the secrets and reboot/redeploy the app if Streamlit asks.

The real token must never be committed to either GitHub repository. The included `deployment/streamlit_secrets.example.toml` contains placeholders only.

With these secrets present, the app uses the private state repository as the durable source of truth and maintains a local cache only as a fallback. If the secrets are absent, local JSON behavior remains unchanged.
