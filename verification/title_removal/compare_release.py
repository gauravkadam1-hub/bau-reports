"""Compare project copies and parity workbooks. No production files are edited.

Usage: python compare_release.py BASELINE_PROJECT UPDATED_PROJECT VERIFICATION_DIR
Run generate_parity.py for each project into VERIFICATION_DIR/baseline_outputs
and VERIFICATION_DIR/updated_outputs first.
"""
from __future__ import annotations
import ast
import copy
import hashlib
import importlib.util
import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

base, updated, out = map(lambda p: Path(p).resolve(), sys.argv[1:4])
CORE = 'docProps/core.xml'
TITLE = '{http://purl.org/dc/elements/1.1/}title'
changed_expected = ['engines/flex/engine.py', 'engines/onboarding/onboarding_writer.py', 'engines/roster/engine.py']

class RemoveMetadataTitle(ast.NodeTransformer):
    def visit_Call(self, node):
        self.generic_visit(node)
        if isinstance(node.func, ast.Attribute) and node.func.attr == 'set_properties':
            assert len(node.args) == 1 and isinstance(node.args[0], ast.Dict)
            dictionary = node.args[0]
            pairs = [(k,v) for k,v in zip(dictionary.keys, dictionary.values)
                     if not (isinstance(k, ast.Constant) and k.value == 'title')]
            dictionary.keys = [k for k,v in pairs]
            dictionary.values = [v for k,v in pairs]
        return node

changed_runtime = []
unchanged_runtime = 0
for path in sorted(base.rglob('*')):
    if not path.is_file() or '__pycache__' in path.parts or '.pytest_cache' in path.parts:
        continue
    rel = path.relative_to(base)
    if not (rel.parts[0] in ('engines', 'portal', 'deployment', '.streamlit', '.devcontainer')
            or str(rel) in ('app.py', 'requirements.txt')):
        continue
    other = updated / rel
    assert other.is_file(), str(rel)
    if path.read_bytes() == other.read_bytes():
        unchanged_runtime += 1
    else:
        changed_runtime.append(rel.as_posix())
        before = RemoveMetadataTitle().visit(ast.parse(path.read_text()))
        after = ast.parse(other.read_text())
        assert ast.dump(before) == ast.dump(after), str(rel)
assert sorted(changed_runtime) == changed_expected, changed_runtime

checks = {}
for scenario, old_path in json.loads((out/'baseline_outputs/outputs.json').read_text()).items():
    new_path = json.loads((out/'updated_outputs/outputs.json').read_text())[scenario]
    with ZipFile(old_path) as before, ZipFile(new_path) as after:
        assert before.testzip() is None and after.testzip() is None
        assert set(before.namelist()) == set(after.namelist())
        changed = sorted(n for n in before.namelist() if before.read(n) != after.read(n))
        assert changed == [CORE], (scenario, changed)
        old_core, new_core = ET.fromstring(before.read(CORE)), ET.fromstring(after.read(CORE))
        old_title = old_core.find(TITLE)
        assert old_title is not None
        previous_title = old_title.text
        assert new_core.find(TITLE) is None
        old_core.remove(old_title)
        assert ET.tostring(old_core) == ET.tostring(new_core)
        checks[scenario] = {
            'scope': 'Actual Onboarding writer with independent prepared-data fixture' if scenario.startswith('onboarding') else 'Full Roster generation from synthetic CSV inputs',
            'previous_title': previous_title,
            'new_title': None,
            'changed_parts': changed,
            'byte_identical_other_parts': len(before.namelist()) - 1,
            'all_non_title_core_metadata_unchanged': True,
        }

# Use actual property dictionaries from all three generators, including Flex,
# in a real XlsxWriter serialization test (not a complete Flex pipeline run).
spec = importlib.util.spec_from_file_location('report_title_tests', updated/'tests/test_report_titles.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
for name, case in module.CASES.items():
    module.test_title_removal_changes_only_core_metadata(name)
    checks[name + '_metadata_serialization'] = {
        'scope': 'Actual source property dictionary; representative XlsxWriter workbook',
        'changed_parts': [CORE],
        'all_non_title_core_metadata_unchanged': True,
        'new_title': None,
    }

compiled = []
for path in sorted(updated.rglob('*.py')):
    if '__pycache__' not in path.parts:
        compile(path.read_bytes(), str(path), 'exec')
        compiled.append(path.relative_to(updated).as_posix())

result = {
    'build': '2026-10-24.4-no-workbook-titles',
    'baseline': 'BAU_Onboarding_Compact_v3.zip',
    'runtime_files_changed': changed_runtime,
    'unchanged_runtime_and_config_files': unchanged_runtime,
    'source_ast_identical_after_only_title_removal': True,
    'existing_tests_changed': False,
    'syntax_checks_passed': len(compiled),
    'comparisons': checks,
    'desktop_excel_clipboard_tested': False,
    'gmail_paste_tested': False,
    'full_flex_pipeline_tested': False,
    'full_polars_onboarding_pipeline_tested': False,
    'limitations': ['Polars and FastExcel are unavailable here; their 24 existing integration cases remain skipped.',
                    'No desktop Excel or Gmail clipboard automation was performed.'],
}
(out/'comparison_results.json').write_text(json.dumps(result, indent=2)+'\n')
print(json.dumps(result, indent=2))
