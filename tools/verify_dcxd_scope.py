#!/usr/bin/env python3
"""Strict local verification for the DC/XD-only reporting change.

Run from the project checkout:
    python tools/verify_dcxd_scope.py
Requires the existing requirements.txt and pytest. No real uploads, accounts,
cloud services, deployments, or production operations are used.
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
REQUIRED = ('polars', 'fastexcel', 'xlrd', 'pandas', 'xlsxwriter', 'openpyxl', 'pytest')


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--results-dir', type=Path, default=ROOT/'verification'/'dcxd_local_run')
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
        print('Verification BLOCKED. Required real backends are missing:', file=sys.stderr)
        for message in missing: print('  '+message, file=sys.stderr)
        print('Install requirements.txt and pytest, then rerun this command.', file=sys.stderr)
        return 2
    junit = dest/'pytest_results.xml'
    process = subprocess.run([sys.executable,'-m','pytest','-q','tests','-ra',f'--junitxml={junit}'],cwd=ROOT,check=False)
    if process.returncode: return process.returncode
    root = ET.parse(junit).getroot()
    suites = [root] if root.tag == 'testsuite' else root.findall('testsuite')
    totals = {name:sum(int(s.get(name,0)) for s in suites) for name in ('tests','failures','errors','skipped')}
    (dest/'summary.json').write_text(json.dumps(totals,indent=2)+'\n',encoding='utf-8')
    if not totals['tests'] or any(totals[name] for name in ('failures','errors','skipped')):
        print(f'Verification INCOMPLETE: a full no-skip run is required. {totals}',file=sys.stderr)
        return 1
    print(f"Verification PASSED: {totals['tests']} tests, no failures, errors or skips.")
    print('Final acceptance: regenerate reports, inspect every embedded base, open/save in desktop Excel,')
    print('and reconcile counts for any employees registered and attending in different facility types.')
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
