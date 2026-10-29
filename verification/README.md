CURRENT BUILD: 2026-10-24.3. Current evidence is under compact_v3/. Other folders and logs are historical.

> Historical migration notes. For the current stylesheet correction, see `../ONBOARDING_STYLES_FIX.md`. Current test evidence is in `verification/style_fix/`.

# Implementation-run evidence

These are actual results from the implementation environment, not a completed deployment sign-off.

- `all_tests.log` / `pytest_results.xml`: 56 passing tests, 24 real-Polars/FastExcel tests skipped.
- `strict_verification.log`: the local acceptance runner correctly blocked success without those dependencies.
- `execution_environment.json`: dependency availability for that run.
- `legacy_comparison.json`: original engine versus new writer, using the synthetic fixtures in `tests/onboarding_fixtures.py` and golden snapshots in `tests/fixtures/onboarding_legacy_golden.json`. The new writer received a standard-library reference model, NOT a mocked or real Polars pipeline.
- `recalculation_results.json`: independent spreadsheet recalculation of the new writer samples; all cached formula results agreed. Not a desktop Excel test.
- `source_integrity.json`: original-file comparison and unchanged report-engine/portal checks.

Run `python tools/verify_onboarding.py` after installing the real dependencies. That command writes a separate `verification/local_run/` folder and requires zero skipped tests. See `ONBOARDING_MIGRATION.md` for remaining manual acceptance.
