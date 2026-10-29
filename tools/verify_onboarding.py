#!/usr/bin/env python3
"""Run the real Onboarding migration tests; reject missing backends or skips.

Run from a project checkout after installing requirements.txt and pytest:
    python tools/verify_onboarding.py
No accounts, production inputs, or external services are used by these tests.
"""
from __future__ import annotations

import argparse
import importlib
import json
from pathlib import Path
import subprocess
import sys
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
REQUIRED = ("polars", "fastexcel", "xlsxwriter", "openpyxl", "pytest")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--results-dir", type=Path, default=ROOT / "verification" / "local_run")
    args = parser.parse_args()
    results = args.results_dir.expanduser().resolve()
    results.mkdir(parents=True, exist_ok=True)
    versions, missing = {}, []
    for name in REQUIRED:
        try:
            module = importlib.import_module(name)
            versions[name] = getattr(module, "__version__", "installed")
        except Exception as exc:
            missing.append(f"{name}: {exc}")
    status = {"python": sys.version, "dependencies": versions, "unavailable": missing}
    (results / "environment.json").write_text(json.dumps(status, indent=2) + "\n", encoding="utf-8")
    if missing:
        print("Verification BLOCKED: required dependencies are missing or failed to import.", file=sys.stderr)
        for message in missing:
            print(f"  {message}", file=sys.stderr)
        print("Install the project requirements and pytest, then rerun:", file=sys.stderr)
        print("  python -m pip install -r requirements.txt", file=sys.stderr)
        print('  python -m pip install "pytest>=8,<10"', file=sys.stderr)
        return 2
    report = results / "pytest_results.xml"
    command = [sys.executable, "-m", "pytest", "-q", "tests", "-ra", f"--junitxml={report}"]
    process = subprocess.run(command, cwd=ROOT, check=False)
    if process.returncode:
        return process.returncode
    try:
        root = ET.parse(report).getroot()
        suites = [root] if root.tag == "testsuite" else root.findall("testsuite")
        totals = {key: sum(int(suite.get(key, "0")) for suite in suites)
                  for key in ("tests", "failures", "errors", "skipped")}
    except (OSError, ET.ParseError, ValueError) as exc:
        print(f"Verification FAILED: cannot validate the test report: {exc}", file=sys.stderr)
        return 2
    (results / "summary.json").write_text(json.dumps(totals, indent=2) + "\n", encoding="utf-8")
    if not totals["tests"] or any(totals[key] for key in ("failures", "errors", "skipped")):
        print(f"Verification FAILED: a complete, non-skipped run is required: {totals}", file=sys.stderr)
        return 1
    print(f"Verification PASSED: {totals['tests']} tests; no failures, errors, or skips.")
    print("Final manual acceptance: generate a new report in the portal, open it in desktop Excel,")
    print("confirm no repair dialog, save with Ctrl+S, then close and reopen it.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
