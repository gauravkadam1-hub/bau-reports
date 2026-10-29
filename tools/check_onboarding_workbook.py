#!/usr/bin/env python3
"""Check a generated seven-sheet Onboarding XLSX without changing it.

Usage:
    python tools/check_onboarding_workbook.py "Onboarding Report - 24th Oct.xlsx"
This is a targeted OOXML style/reference check, not a desktop Excel opening test.
"""
from pathlib import Path
import argparse
import json
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from engines.onboarding.onboarding_styles import validate_workbook
from engines.onboarding.utils import ReportError


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    args = parser.parse_args()
    try:
        result = validate_workbook(args.workbook)
        with ZipFile(args.workbook) as archive:
            core = ET.fromstring(archive.read('docProps/core.xml'))
            result['generator_description'] = core.findtext('{http://purl.org/dc/elements/1.1/}description', '')
        print(json.dumps({'status': 'PASS', **result}, indent=2))
        return 0
    except (ReportError, OSError, KeyError, ET.ParseError) as exc:
        print(f'FAIL: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
