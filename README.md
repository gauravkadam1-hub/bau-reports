# Reporting Automation Hub - Streamlit foundation

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
- `all_user_data_all_*` -> All Users (exactly two expected for Flex/Attendance)
- `attendance_report_*` -> Attendance Base
- `onboarding_funnel_all_*` -> Onboarding Base
- `All user data base.csv` is explicitly ignored.

Dates/timestamps in filenames may change without affecting recognition.

## Existing engines

The `engines/` folder contains copies of the current supplied generator engines/assets, not rewritten report calculations. Streamlit adapters invoke their public scan/generate functions.

## Local start

```bash
python -m venv .venv
# activate it
pip install -r requirements.txt
streamlit run app.py
```

For deployment, keep the Git repository private and deploy the single app from that repository.

## Facility selection memory

The Streamlit portal keeps separate remembered facility selections for Roster, Flex, Attendance, and Onboarding.

- Checked/unchecked site choices survive normal local restarts.
- Existing desktop settings are imported automatically on first use when their settings files are available on the same Windows user account.
- Newly discovered non-reference sites start unchecked and are labelled `NEW`.
- Roster clears `NEW` when the new site is selected, matching the desktop app.
- Flex clears `NEW` after a successful Flex generation, matching the desktop app.
- Attendance records the current discovered site list when generation starts, matching the desktop app.
- Onboarding shows every current template facility as a checkbox. New source-only facilities start unchecked and labelled `NEW`; selected new facilities become known after a successful generation, while unticked new facilities remain `NEW` for a future run.
- Onboarding applies the chosen facilities to both Aadhar and Hyperverge DC/XD pairs, preserving shared row order and latest-date RED-descending then Manual-descending sorting.

Local portal preferences are stored at:

- Windows: `%APPDATA%\\ReportingAutomationHub\\facility_preferences.json`
- macOS/Linux: `~/.reporting_automation_hub/facility_preferences.json`

For Streamlit Community Cloud, local file storage is not guaranteed to survive a cloud reboot. Before production deployment, configure a durable zero-cost preference backend if cross-reboot persistence is required.
