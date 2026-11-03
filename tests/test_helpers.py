"""Small, independent fixtures; no production logs or real dataset are included."""
from pathlib import Path
import csv
import hashlib
import importlib.util
import json
import os
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / 'skills' / 'online-data-eval-dataset' / 'scripts'


def load(name):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


inventory = load('inventory_sources')
validator = load('validate_dataset')
packager = load('package_skill')


def make_xlsx(path):
    """Minimal OOXML fixture: different cell types, multiline request and a cached formula."""
    s = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    files = {
        'xl/workbook.xml': f'<workbook xmlns="{s}" xmlns:r="{r}"><workbookPr date1904="0"/><sheets><sheet name="明细" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/sharedStrings.xml': f'<sst xmlns="{s}"><si><t>原始问题\n不要丢掉第二行</t></si></sst>',
        'xl/styles.xml': f'<styleSheet xmlns="{s}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>',
        'xl/worksheets/sheet1.xml': f'''<worksheet xmlns="{s}"><sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>TraceID</t></is></c><c r="B1" t="inlineStr"><is><t>提问内容</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>sample-trace-1</t></is></c><c r="B2" t="s"><v>0</v></c><c r="C2" s="1"><v>45291</v></c><c r="F2"><f t="shared" si="0" ref="F2:F3">SUM(1,2)</f><v>3</v></c></row>
<row r="3"><c r="B3" t="inlineStr"><is><t>【已有记忆】背景\n【新增对话】继续执行\n请输出整合后的完整记忆摘要。</t></is></c></row>
</sheetData></worksheet>''',
    }
    with zipfile.ZipFile(path, 'w') as z:
        for name, content in files.items():
            z.writestr(name, content)


def dataset(root):
    root.mkdir()
    for name in ['数据集说明.md', '来源与调整说明.md']:
        (root / name).write_text('管理资料\n', encoding='utf-8')
    cases = []
    for number in [1, 2]:
        ident = f'A{number:02d}'
        title = f'状态汇总{number}'
        relative = f'A_表格数据处理与分析/{ident}_{title}'
        folder = root / relative
        folder.mkdir(parents=True)
        (folder / '输入.csv').write_text('编号,状态\nX1,已解决\nX2,待解决\n', encoding='utf-8-sig')
        output = ident + '_状态汇总.xlsx'
        (folder / '问题描述.txt').write_text(f'帮我按输入.csv里的状态统计问题数量，给我“{output}”，保留明细和汇总。', encoding='utf-8-sig')
        cases.append({'id': ident, 'title': title, 'category': 'A_表格数据处理与分析', 'folder': relative,
                      'ability': '按给定状态统计并核对数量。',
                      'sources': [{'kind': 'summary', 'file': '汇总.md', 'locator': f'第{number}行', 'excerpt': '按状态统计。'}],
                      'inputs': [{'name': '输入.csv', 'origin': 'constructed', 'notes': '独立测试夹具。'}],
                      'outputs': [{'name': output, 'description': '明细与状态汇总'}]})
    with (root / '题目索引.csv').open('w', encoding='utf-8-sig', newline='') as stream:
        w = csv.writer(stream)
        w.writerow(['编号', '类别', '题目', '相对目录', '同目录输入文件', '正式交付', '联网', '素材性质'])
        for case in cases:
            w.writerow([case['id'], case['category'], case['title'], case['folder'], '输入.csv', case['outputs'][0]['name'], '否', '构造'])
    validator.write_manifest(root)
    return {'batch': 'fixture', 'expected_count': 2, 'cases': cases}


class InventoryTests(unittest.TestCase):
    def test_raw_excel_values_locations_and_internal_requests_are_preserved(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            p = root / '日志.xlsx'
            make_xlsx(p)
            output = root / 'out' / 'records.jsonl'
            result = inventory.inventory([p], output)
            values = [json.loads(line) for line in output.read_text(encoding='utf-8').splitlines()]
            self.assertTrue(result['complete'])
            self.assertEqual(values[1]['locator'], {'sheet': '明细', 'row': 2})
            self.assertEqual(values[1]['cells']['B2']['value'], '原始问题\n不要丢掉第二行')
            self.assertEqual(values[1]['cells']['C2']['value'], '45291')
            self.assertEqual(values[1]['cells']['C2']['number_format'], 'yyyy-mm-dd')
            self.assertEqual(values[1]['cells']['F2']['formula_attributes']['si'], '0')
            self.assertIn('请输出整合后的完整记忆摘要', values[2]['cells']['B3']['value'])

    def test_markdown_and_csv_multiline_and_same_names_have_distinct_provenance(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            a = root / 'one'; b = root / 'two'
            a.mkdir(); b.mkdir()
            (a / '日志.md').write_text('# 错误分类\n\n帮我更新通报。\nTraceID: example-only\n', encoding='utf-8')
            (b / '日志.md').write_text('继续执行\n', encoding='utf-8')
            (a / '记录.csv').write_text('id,text\n1,"第一行\n第二行"\n', encoding='utf-8')
            result = inventory.inventory([a, b], root / 'out' / 'records.jsonl')
            data = [json.loads(x) for x in (root / 'out' / 'records.jsonl').read_text(encoding='utf-8').splitlines()]
            self.assertTrue(result['complete'])
            question = next(x for x in data if x.get('raw_line') == '帮我更新通报。')
            self.assertEqual(question['locator']['line'], 3)
            self.assertEqual(question['headings'], ['错误分类'])
            self.assertEqual(next(x for x in data if x.get('values', [''])[0] == '1')['values'][1], '第一行\n第二行')
            self.assertEqual(len({x['file'] for x in data if x['source_type'] == 'md'}), 2)

    def test_empty_unsupported_and_corrupt_sources_are_not_claimed_complete(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            empty = root / 'empty'; empty.mkdir()
            self.assertFalse(inventory.inventory([empty], root / 'out' / 'empty.jsonl')['complete'])
            (root / 'old.xls').write_bytes(b'old format')
            (root / 'bad.xlsx').write_bytes(b'broken')
            result = inventory.inventory([root / 'old.xls', root / 'bad.xlsx'], root / 'out' / 'records.jsonl')
            self.assertFalse(result['complete'])
            self.assertEqual({x['status'] for x in result['files']}, {'unsupported', 'error'})

    def test_inventory_refuses_to_overwrite_a_source(self):
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / 'source.md'; p.write_text('original', encoding='utf-8')
            with self.assertRaises(FileExistsError):
                inventory.inventory([p], p)
            self.assertEqual(p.read_text(encoding='utf-8'), 'original')


class DatasetTests(unittest.TestCase):
    def test_complete_dataset_passes_without_claiming_semantic_quality(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td) / 'dataset'; catalog = dataset(root)
            result = validator.validate(root, catalog)
            self.assertTrue(result['structural_checks_passed'], result)
            self.assertTrue(result['semantic_review_required'])
            self.assertEqual(result['input_files'], 2)

    def test_duplicate_output_and_missing_attachment_are_detected(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td) / 'dataset'; catalog = dataset(root)
            catalog['cases'][1]['outputs'][0]['name'] = catalog['cases'][0]['outputs'][0]['name']
            (root / catalog['cases'][0]['folder'] / '输入.csv').unlink()
            result = validator.validate(root, catalog)
            codes = {x['code'] for x in result['errors']}
            self.assertIn('duplicate_output', codes)
            self.assertIn('case_files', codes)
            self.assertIn('index_content', codes)

    def test_subfolders_archives_and_case_path_traversal_are_detected(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td) / 'dataset'; catalog = dataset(root)
            folder = root / catalog['cases'][0]['folder']
            (folder / '问题源素材').mkdir()
            (folder / '素材.zip').write_bytes(b'not an allowed input')
            catalog['cases'][1]['folder'] = '../outside'
            codes = {x['code'] for x in validator.validate(root, catalog)['errors']}
            self.assertTrue({'nested_input', 'archive', 'case_path'} <= codes)

    def test_product_confirmation_is_distinguished_from_waiting_for_user(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td) / 'dataset'; catalog = dataset(root)
            p = root / catalog['cases'][0]['folder'] / '问题描述.txt'
            p.write_text('帮我做带确认抽选和清空提示的工具，交付A01_状态汇总.xlsx。', encoding='utf-8')
            warnings = validator.validate(root, catalog, False)['warnings']
            self.assertNotIn('confirmation_gate_review', {x['code'] for x in warnings})
            p.write_text('先给我方案，等我确认后再执行，交付A01_状态汇总.xlsx。', encoding='utf-8')
            warnings = validator.validate(root, catalog, False)['warnings']
            self.assertIn('confirmation_gate_review', {x['code'] for x in warnings})

    def test_manifest_detects_changed_material_and_error_only_source_is_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td) / 'dataset'; catalog = dataset(root)
            (root / catalog['cases'][0]['folder'] / '输入.csv').write_text('changed', encoding='utf-8')
            catalog['cases'][1]['sources'][0]['kind'] = 'error_evidence'
            codes = {x['code'] for x in validator.validate(root, catalog)['errors']}
            self.assertIn('manifest_hash', codes)
            self.assertIn('source', codes)

    def test_legitimate_simulation_is_a_review_warning_not_automatic_rejection(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td) / 'dataset'; catalog = dataset(root)
            p = root / catalog['cases'][0]['folder'] / '问题描述.txt'
            p.write_text('帮我核查仿真项目的模拟数据，交付A01_状态汇总.xlsx。', encoding='utf-8')
            validator.write_manifest(root)
            result = validator.validate(root, catalog)
            self.assertTrue(result['structural_checks_passed'])
            self.assertIn('input_disclosure_review', {x['code'] for x in result['warnings']})


class PackageTests(unittest.TestCase):
    def test_reproducible_archive_and_version_consistency(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td); skill = root / 'sample-skill'; skill.mkdir()
            (skill / 'SKILL.md').write_text('---\nname: sample-skill\ndescription: Example\nmetadata:\n  version: "1.0.0"\n---\nBody\n', encoding='utf-8')
            first = packager.package(skill, '1.0.0', root / 'one')
            os.utime(skill / 'SKILL.md', (1234567890, 1234567890))
            second = packager.package(skill, '1.0.0', root / 'two')
            self.assertEqual(first['sha256'], second['sha256'])
            self.assertEqual(hashlib.sha256(Path(first['file']).read_bytes()).hexdigest(), first['sha256'])
            with zipfile.ZipFile(first['file']) as z:
                self.assertEqual(z.namelist(), ['sample-skill/SKILL.md'])
            with self.assertRaises(ValueError):
                packager.package(skill, '1.0.1', root / 'three')
            with self.assertRaises(ValueError):
                packager.package(skill, '1.0.0', skill / 'dist')
            (skill / 'new.md').write_text('changed content', encoding='utf-8')
            with self.assertRaises(FileExistsError):
                packager.package(skill, '1.0.0', root / 'one')
            self.assertEqual(hashlib.sha256(Path(first['file']).read_bytes()).hexdigest(), first['sha256'])


if __name__ == '__main__':
    unittest.main()
