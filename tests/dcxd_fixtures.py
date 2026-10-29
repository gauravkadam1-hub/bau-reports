"""Synthetic cross-engine fixtures. No employee data or connected accounts."""
from datetime import date
from pathlib import Path
import csv

FIRST = date(2026, 10, 22)
LAST = date(2026, 10, 23)
DC = 'Synthetic DC'
XD = 'Synthetic XD'
EXTRA = 'Unselected XD'


def write_csv(path, headers, rows, second_header=None):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8-sig', newline='') as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        if second_header is not None:
            writer.writerow(second_header)
        writer.writerows(rows)
    return path


def mixed_inputs(folder, include_outside=True, cross_facility=False, additional_rows=0):
    folder = Path(folder)
    records = [
        ('v5-001', DC, 'GC'), ('v5-002', XD, 'GC'), ('v5-003', EXTRA, 'GC'),
        ('v5-004', DC, 'Security'), ('v5-005', XD, 'Housekeeping'),
        ('v5-006', '  Padded   dc  ', 'GC'),
    ]
    if include_outside:
        records += [('v5-007', 'Synthetic HUB', 'GC'), ('v5-008', 'Synthetic DC Annex', 'GC'),
                    ('v5-009', 'SyntheticDC', 'GC'), ('v5-010', '', 'GC')]
    for i in range(additional_rows):
        records.append((f'bulk-{i:06}', DC if i % 10 == 0 else 'Regional HUB', 'GC'))
    users = [[ident, 'Synthetic '+ident, role, 'No', 'Onboarded - Active', site,
              '2026-10-01', 'phone-'+ident, 'aad-'+ident] for ident, site, role in records]
    if cross_facility:
        users.append(['v5-orphan', 'Synthetic orphan', 'GC', 'Yes', 'Onboarded - Active',
                      'Registered HUB', '2026-10-01', 'phone-orphan', 'aad-orphan'])
    user_headers = ['ID', 'Name', 'Role', 'Floater', 'Status', 'Registered Facility',
                    'Onboarded Date', 'Phone Number', 'Aadhaar Number']
    midpoint = max(1, len(users)//2)
    u1 = write_csv(folder/'all_users_2026-10-01.csv', user_headers, users[:midpoint])
    u2 = write_csv(folder/'all_users_2026-10-02.csv', user_headers, users[midpoint:])
    tops = ['user_id', 'user_name', 'role', 'facility_name', 'stint_id', 'last_present_day']
    subs = ['']*len(tops)
    for day in (FIRST, LAST):
        tops += [day.isoformat(), '', '', '']
        subs += ['shift', 'total_hours', 'working_hours', 'punch_status']
    attendance_rows = [[ident, 'Synthetic '+ident, role, site, 'one', '',
                        'P1', '8', '8', 'P', 'P1', '8', '8', 'P'] for ident, site, role in records]
    if cross_facility:
        attendance_rows.append(['v5-orphan', 'Synthetic orphan', 'GC', DC, 'one', '',
                                'P1', '8', '8', 'P', 'P1', '8', '8', 'P'])
    att = write_csv(folder/'attendance.csv', tops, attendance_rows, subs)
    roster_headers = ['User ID', 'Name', 'Role', 'Facility Name', 'Sub Status', 'Status',
                      'Onboarded Date', 'Last Working Day', '22-Oct-2026', '23-Oct-2026']
    roster_rows = [[ident, 'Synthetic '+ident, role, site, 'ACTIVE', 'Onboarded',
                    '01-Oct-2026', '', 'P1', 'P1'] for ident, site, role in records]
    roster = write_csv(folder/'roster.csv', roster_headers, roster_rows)
    flex_headers = ['Name', 'Shift', 'Mobile Number', 'Aadhar Number', 'Vendor Name', 'Location',
                    'Date', 'Sub Cost Centre', 'HRBP Name', 'Flex rate', 'Employee Type',
                    'Type ( Sourcing / Internal)']
    flex_rows = [['Synthetic '+ident, 'P1', 'phone-'+ident, 'aad-'+ident, 'Synthetic vendor',
                  site, day.isoformat(), 'DC operations', 'Synthetic HR', '100', 'FT', 'Sourcing']
                 for ident, site, role in records for day in (FIRST, LAST)]
    flex = write_csv(folder/'flex.csv', flex_headers, flex_rows)
    return {'users1': u1, 'users2': u2, 'attendance': att, 'roster': roster, 'flex': flex}


def scoped_reference_rows(rows, headers):
    """Independent test oracle; deliberately does not call production helpers."""
    site_index = headers.index('Registered Facility')
    return [row for row in rows if ' '.join(str(row[site_index] or '').split()).casefold().endswith((' dc', ' xd'))]
