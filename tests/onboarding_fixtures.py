"""Synthetic, non-personal fixtures and a deliberately independent Python oracle."""
from datetime import date, datetime, timedelta
from engines.onboarding.onboarding_common import Counts, PreparedReport, clean_selection, clean_facility, excel_text_matches
from engines.onboarding.utils import parse_date_value, inclusive_date_range

START = date(2026, 9, 1)
END = date(2026, 9, 4)
DC = ["Alpha DC", "Beta DC", "Absent DC"]
XD = ["East XD", "West XD", "Absent XD"]


def rows_for(profile):
    records = []
    def add(ident, site, day, flag, status, person="Synthetic HR A", **extra):
        values = {h: None for h in profile['headers']}
        values.update({'ID': ident, 'Name': 'Synthetic Employee '+str(ident), 'Status': 'Onboarded - Active',
                       'Role': 'GC', 'Registered Facility': site, 'Onboarded Date': day,
                       'Approved By': person, 'Onboarded by Aadhaar (Y/N)': flag, 'Hyperverge Status': status,
                       'Phone Number': '001234567890', 'Present Count Since DOJ': 3})
        values.update(extra)
        records.append([values[h] for h in profile['headers']])
    add('0001', 'Alpha DC', START, 'Y', 'green')
    add('0002', 'Alpha DC', START, 'N', 'red', 'Synthetic HR B')
    add('0003', 'Beta DC', START+timedelta(days=1), 'Y', 'green')
    add('0004', 'West XD', START+timedelta(days=1), 'N', 'red', 'Synthetic HR C')
    add('0005', 'East XD', START, 'N', 'red', 'Synthetic HR C')
    add('0006', 'Alpha DC', END, 'Y', 'green')
    add('0007', 'Beta DC', END, 'N', 'red', 'Synthetic HR B')
    add('0008', 'Beta DC', END, 'n', 'RED', 'Synthetic HR B')
    add('0008', 'Beta DC', END, 'n', 'RED', 'Synthetic HR B')  # Duplicate preserved, not deduplicated.
    add('0009', 'West XD', END, 'Y', 'red', 'Synthetic HR C')
    add('0010', 'East XD', END, 'N', 'green', 'Synthetic HR A')
    add('0011', 'Unselected DC', END, 'N', 'red', 'Synthetic Unselected HR')
    add('0012', 'Other Facility', END, 'N', 'red', 'Synthetic Other HR')
    add('0013', 'Alpha DC', 'not a date', 'N', 'red')
    add('0014', 'Alpha DC', START-timedelta(days=1), 'N', 'red')
    add('0015', 'Alpha DC', END+timedelta(days=1), 'N', 'red')
    add('0016', '  Alpha   DC ', END, ' N ', ' red ', ' Synthetic Padded HR ')
    add('0017', 'ALPHA DC', END, 'N', 'Red', 'Synthetic HR A')
    return records


def reference_prepare(profile, rows, first=START, last=END, dc=DC, xd=XD):
    """Oracle only: this is NOT a fallback in production."""
    heads = profile['headers']
    position = {name: index for index, name in enumerate(heads)}
    filtered = []
    for source in rows:
        day = parse_date_value(source[position['Onboarded Date']])
        if day is not None and first <= day <= last:
            row = list(source)
            row[position['Onboarded Date']] = datetime.combine(day, datetime.min.time())
            filtered.append(row)
    def value(row, column): return row[position[column]]
    def rawtext(item): return '' if item is None else str(item)
    sort = {}
    manual, red = set(), set()
    for row in filtered:
        site = clean_facility(value(row,'Registered Facility')).casefold()
        flag = rawtext(value(row,'Onboarded by Aadhaar (Y/N)')).strip().casefold()
        hv = rawtext(value(row,'Hyperverge Status')).strip().casefold()
        hr = rawtext(value(row,'Approved By')).strip()
        if hr and flag=='n': manual.add(hr)
        if hr and hv=='red': red.add(hr)
        if value(row,'Onboarded Date').date()==last:
            entry = sort.setdefault(site, [0,0])
            entry[0] += hv=='red'; entry[1] += flag=='n'
    chosen = {}
    for kind, selection in [('DC',dc),('XD',xd)]:
        sites=clean_selection(profile['defaults'][kind] if selection is None else selection,kind)
        chosen[kind]=sorted(sites,key=lambda name: tuple(-v for v in sort.get(name.casefold(),[0,0])))
    dates=inclusive_date_range(first,last)
    counts={}
    for site in chosen['DC']+chosen['XD']:
        for day in dates:
            matching=[row for row in filtered if value(row,'Onboarded Date').date()==day and excel_text_matches(value(row,'Registered Facility'),site)]
            counts[site,day]=Counts(len(matching),
                sum(rawtext(value(row,'Onboarded by Aadhaar (Y/N)')).casefold()=='y' for row in matching),
                sum(rawtext(value(row,'Onboarded by Aadhaar (Y/N)')).casefold()=='n' for row in matching),
                sum(rawtext(value(row,'Hyperverge Status')).casefold()=='red' for row in matching))
    def hr_counts(names,column,criterion):
        return {name:sum(excel_text_matches(value(row,'Approved By'),name) and rawtext(value(row,column)).casefold()==criterion for row in filtered) for name in names}
    def names_order(names,sheet):
        old=profile['sheets'][sheet]['existing_names']
        ordered=[n for n in old if n in names]
        return ordered+sorted(names.difference(ordered),key=str.casefold)
    manual=names_order(manual,'Manual Onb - HR wise details')
    red=names_order(red,'RED Cases - HR wise details')
    return PreparedReport(dates,chosen['DC'],chosen['XD'],filtered,len(filtered),counts,
                          manual,red,hr_counts(manual,'Onboarded by Aadhaar (Y/N)','n'),hr_counts(red,'Hyperverge Status','red'))
