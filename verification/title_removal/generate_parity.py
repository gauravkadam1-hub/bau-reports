"""Generate actual Onboarding writer/Roster outputs using either project copy.

The fixed creation timestamp is test-only. Onboarding uses the independent
PreparedReport fixture, not the unavailable Polars input pipeline.
"""
from __future__ import annotations
import csv
import json
import sys
from datetime import date, datetime
from pathlib import Path
import xlsxwriter

root, output = map(lambda p: Path(p).resolve(), sys.argv[1:3])
sys.path.insert(0, str(root))
output.mkdir(parents=True, exist_ok=True)
real_set_properties = xlsxwriter.Workbook.set_properties

def fixed_clock(self, properties):
    return real_set_properties(self, {**properties, 'created': datetime(2026, 10, 24)})
xlsxwriter.Workbook.set_properties = fixed_clock

from engines.onboarding.onboarding_common import load_profile
from engines.onboarding.onboarding_writer import write_report
from tests.onboarding_fixtures import DC, XD, reference_prepare, rows_for
from engines.roster.engine import ReportInputs, generate_report

result = {}
profile = load_profile()
for name, dc, xd in [('both', DC, XD), ('dc_only', DC, []), ('xd_only', [], XD), ('empty', [], [])]:
    destination = output / ('onboarding_' + name + '.xlsx')
    report = reference_prepare(profile, rows_for(profile), dc=dc, xd=xd)
    write_report(destination, profile, report)
    result['onboarding_' + name] = str(destination)

inputs = output / 'inputs'
inputs.mkdir(exist_ok=True)
roster, attendance = inputs / 'roster.csv', inputs / 'attendance.csv'
with roster.open('w', encoding='utf-8', newline='') as handle:
    writer = csv.writer(handle)
    writer.writerow(['User ID', 'Name', 'Role', 'Facility Name', 'Sub Status', 'Status', 'Onboarded Date', 'Last Working Day', '22-Oct-2026', '23-Oct-2026'])
    writer.writerows([
        ['001', 'Synthetic One', 'GC', 'Synthetic DC', 'ACTIVE', 'Onboarded', '01-Oct-2026', '', 'P1', 'P1'],
        ['002', 'Synthetic Two', 'GC', 'Synthetic DC', 'ACTIVE', 'Onboarded', '01-Oct-2026', '', 'P2', 'WO'],
        ['003', 'Synthetic Three', 'GC', 'Synthetic XD', 'ACTIVE', 'Onboarded', '01-Oct-2026', '', 'P3', 'P1'],
    ])
with attendance.open('w', encoding='utf-8', newline='') as handle:
    writer = csv.writer(handle)
    writer.writerow(['Employee Code', 'Name', 'Role', 'Facility Name', '22-Oct-2026', '23-Oct-2026'])
    writer.writerows([
        ['001', 'Synthetic One', 'GC', 'Synthetic DC', 'P1', 'P1'],
        ['002', 'Synthetic Two', 'GC', 'Synthetic DC', 'P2', '-'],
        ['003', 'Synthetic Three', 'GC', 'Synthetic XD', 'P3', 'P2'],
    ])
for name, xd in [('both', ['Synthetic XD']), ('dc_only', [])]:
    destination = output / ('roster_' + name)
    args = ReportInputs(str(roster), str(attendance), str(destination), date(2026, 10, 22), date(2026, 10, 23), ['Synthetic DC'], xd)
    result['roster_' + name] = generate_report(args)
(output / 'outputs.json').write_text(json.dumps(result, indent=2) + '\n')
print(json.dumps(result, indent=2))
