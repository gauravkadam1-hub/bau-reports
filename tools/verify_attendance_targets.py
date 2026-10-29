#!/usr/bin/env python3
"""Verify Attendance targets using real backends; fail on missing modules/skips.

Run from the project folder:
    python tools/verify_attendance_targets.py
Tests use synthetic inputs and temporary masters, not live accounts or uploads.
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
REQUIRED = ('polars', 'pandas', 'xlsxwriter', 'openpyxl', 'pytest')


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--results-dir', type=Path, default=ROOT/'verification'/'attendance_targets_local')
    args = parser.parse_args()
    dest = args.results_dir.expanduser().resolve()
    dest.mkdir(parents=True, exist_ok=True)
    versions, missing = {}, []
    for name in REQUIRED:
        try:
            module = importlib.import_module(name)
            versions[name] = getattr(module, '__version__', 'available')
        except Exception as exc:
            missing.append(f'{name}: {exc}')
    (dest/'environment.json').write_text(json.dumps({'python':sys.version,'versions':versions,'missing':missing},indent=2)+'\n',encoding='utf-8')
    if missing:
        print('Verification BLOCKED. Required dependencies are missing:', file=sys.stderr)
        for message in missing: print('  '+message, file=sys.stderr)
        print('Install the existing requirements.txt and pytest, then rerun.', file=sys.stderr)
        return 2
    junit = dest/'pytest_results.xml'
    process = subprocess.run([sys.executable,'-m','pytest','-q','tests/test_attendance_targets.py','-ra',f'--junitxml={junit}'],cwd=ROOT,check=False)
    if process.returncode: return process.returncode
    root = ET.parse(junit).getroot()
    suites = [root] if root.tag=='testsuite' else root.findall('testsuite')
    totals = {name:sum(int(s.get(name,0)) for s in suites) for name in ('tests','failures','errors','skipped')}
    (dest/'summary.json').write_text(json.dumps(totals,indent=2)+'\n',encoding='utf-8')
    if not totals['tests'] or any(totals[name] for name in ('failures','errors','skipped')):
        print(f'Verification INCOMPLETE: no-skip run required. {totals}',file=sys.stderr)
        return 1
    print(f"Verification PASSED: {totals['tests']} tests, no failures, errors or skips.")
    print('Final desktop check: generate Housekeeping, Security and Main Attendance;')
    print('inspect targets for every report date, then open, save and reopen in Excel.')
    return 0


if __name__=='__main__':
    raise SystemExit(main())
