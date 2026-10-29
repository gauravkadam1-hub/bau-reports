#!/usr/bin/env python3
"""Read-only XLSX size breakdown; no employee names or cell values are printed.

Usage: python tools/inspect_onboarding_size.py old.xlsx new.xlsx
"""
from __future__ import annotations
import argparse
import json
import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def inspect(path: Path) -> dict:
    with ZipFile(path) as archive:
        result = {'file': path.name, 'bytes': path.stat().st_size,
                  'largest_parts': [{'part': p.filename, 'compressed_bytes': p.compress_size,
                                     'uncompressed_bytes': p.file_size}
                                    for p in sorted(archive.infolist(), key=lambda p: -p.compress_size)[:8]]}
        book = ET.fromstring(archive.read('xl/workbook.xml'))
        rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
        base = next((s for s in book.find(NS+'sheets') if s.get('name') == 'Onboarding Base'), None)
        if base is not None:
            target = rels[base.get(REL+'id')]
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
            cells = blanks = styled = data_rows = 0
            row_has_value = False
            with archive.open(part) as handle:
                for event, node in ET.iterparse(handle, events=('start', 'end')):
                    if event == 'start' and node.tag == NS+'row':
                        row_has_value = False
                    elif event == 'end' and node.tag == NS+'c':
                        cells += 1
                        styled += node.get('s', '0') != '0'
                        has_value = any(node.find(NS+tag) is not None for tag in ['v','is','f'])
                        blanks += not has_value
                        row_has_value |= has_value
                        node.clear()
                    elif event == 'end' and node.tag == NS+'row':
                        data_rows += row_has_value and int(node.get('r', '0')) > 1
                        node.clear()
            result['base'] = {'data_rows': data_rows, 'stored_cells': cells,
                              'stored_blank_cells': blanks, 'styled_cells': styled}
        return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('files', type=Path, nargs='+')
    args = parser.parse_args()
    try:
        print(json.dumps([inspect(path) for path in args.files], indent=2))
    except Exception as exc:
        parser.exit(1, f'Inspection failed: {exc}\n')


if __name__ == '__main__':
    main()
