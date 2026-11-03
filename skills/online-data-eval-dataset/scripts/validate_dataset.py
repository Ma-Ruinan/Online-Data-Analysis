"""Read-only structural checks; never a substitute for business/content review."""
from __future__ import annotations
import argparse
import csv
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import xml.etree.ElementTree as ET
import zipfile

ADMIN = {'数据集说明.md', '题目索引.csv', '来源与调整说明.md', '文件清单.csv'}
ARCHIVES = {'.zip', '.rar', '.7z', '.gz', '.tar', '.tgz', '.bz2', '.xz'}
TEXT = {'.txt', '.md', '.csv', '.json', '.js', '.ts', '.html', '.css', '.log'}
NATIVE = {'.xlsx': 'xl/workbook.xml', '.docx': 'word/document.xml', '.pptx': 'ppt/presentation.xml'}
SOURCE_KINDS = {'user_quote', 'summary', 'embedded_dialogue', 'error_evidence'}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def filename(value) -> bool:
    if not isinstance(value, str) or not value or value in {'.', '..'}:
        return False
    if any(c in value for c in '/\\<>:"|?*\x00') or value.endswith((' ', '.')):
        return False
    return value.split('.')[0].upper() not in {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{x}' for x in range(1, 10)), *(f'LPT{x}' for x in range(1, 10))}


def safe_relative(value) -> bool:
    if not isinstance(value, str) or '\\' in value:
        return False
    p = PurePosixPath(value)
    return not p.is_absolute() and len(p.parts) == 2 and all(filename(x) for x in p.parts)


def write_manifest(root: Path):
    root = root.resolve()
    for p in root.rglob('*'):
        if p.is_symlink():
            raise ValueError('Do not hash symlinked dataset paths')
    with (root / '文件清单.csv').open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.writer(stream)
        writer.writerow(['相对路径', '字节数', 'SHA256'])
        for path in sorted(root.rglob('*')):
            if path.is_file() and path != root / '文件清单.csv':
                writer.writerow([path.relative_to(root).as_posix(), path.stat().st_size, sha256(path)])


def validate(root: Path, catalog: dict, verify_manifest=True) -> dict:
    root = root.resolve()
    errors, warnings = [], []
    def issue(code, message, case=None, warning=False):
        (warnings if warning else errors).append({'code': code, 'case': case, 'message': message})
    if not root.is_dir():
        return {'structural_checks_passed': False, 'errors': [{'code': 'missing_root', 'case': None, 'message': str(root)}], 'warnings': []}
    if not isinstance(catalog, dict):
        return {'structural_checks_passed': False, 'errors': [{'code': 'catalog_schema', 'case': None, 'message': 'Catalog must be an object'}], 'warnings': []}
    for name in ADMIN - ({'文件清单.csv'} if not verify_manifest else set()):
        if not (root / name).is_file():
            issue('missing_admin', name)
    for p in root.rglob('*'):
        if p.is_symlink():
            issue('symlink', str(p.relative_to(root)))
    cases = catalog.get('cases', [])
    if not isinstance(cases, list) or not cases:
        issue('empty_catalog', 'Catalog must contain a nonempty cases list')
        cases = []
    if 'expected_count' in catalog and catalog['expected_count'] != len(cases):
        issue('count', 'Expected count differs from catalog cases')
    seen_ids, seen_folders, deliveries = set(), set(), set()
    input_count = 0
    for case in cases:
        if not isinstance(case, dict):
            issue('case_schema', 'Case must be an object')
            continue
        ident, relative = case.get('id', ''), case.get('folder', '')
        ident = ident if isinstance(ident, str) else '<invalid>'
        if not isinstance(ident, str) or not re.fullmatch(r'[A-Z]\d{2,}', ident) or ident in seen_ids:
            issue('case_id', 'Missing, malformed or duplicate ID', str(ident))
        else:
            seen_ids.add(ident)
        if not safe_relative(relative):
            issue('case_path', 'Use category/case relative directory', str(ident))
            continue
        if relative.casefold() in seen_folders:
            issue('case_path', 'Duplicate case directory', ident)
        seen_folders.add(relative.casefold())
        folder = root / relative
        if not folder.is_dir() or folder.is_symlink() or not folder.resolve().is_relative_to(root):
            issue('case_folder', 'Missing or unsafe case directory', ident)
            continue
        for key in ['title', 'category', 'ability']:
            if not isinstance(case.get(key), str) or not case[key].strip():
                issue('case_schema', f'Missing {key}', ident)
        sources = case.get('sources', [])
        if not isinstance(sources, list) or not sources:
            issue('source', 'Missing business-intent source', ident)
            sources = []
        has_intent = False
        for source in sources:
            if not isinstance(source, dict):
                issue('source', 'Source must be an object', ident)
                continue
            kind = source.get('kind')
            kind = kind if isinstance(kind, str) else ''
            if kind not in SOURCE_KINDS or any(not isinstance(source.get(k), str) or not source[k].strip() for k in ['file', 'locator', 'excerpt']):
                issue('source', 'Source needs file, locator, excerpt and valid kind', ident)
            has_intent |= kind in SOURCE_KINDS - {'error_evidence'}
        if not has_intent:
            issue('source', 'An error log alone is not a business request', ident)
        inputs = case.get('inputs', [])
        if not isinstance(inputs, list):
            issue('input_schema', 'inputs must be a list', ident)
            inputs = []
        names = []
        for item in inputs:
            if not isinstance(item, dict) or not filename(item.get('name')):
                issue('input_name', 'Input needs an ordinary filename', ident)
                continue
            name = item['name']
            names.append(name)
            if item.get('origin') not in {'original', 'constructed', 'official'} or not str(item.get('notes', '')).strip():
                issue('input_origin', 'Input needs origin and provenance notes', ident)
            if name == '问题描述.txt' or Path(name).suffix.lower() in ARCHIVES:
                issue('input_name', 'Reserved prompt name or compressed input', ident)
        if len({n.casefold() for n in names}) != len(names):
            issue('input_name', 'Duplicate input name', ident)
        input_count += len(names)
        actual = {p.name for p in folder.iterdir() if p.is_file()}
        if actual != set(names) | {'问题描述.txt'}:
            issue('case_files', f'Expected input/prompt set differs: {sorted(actual.symmetric_difference(set(names) | {"问题描述.txt"}))}', ident)
        for p in folder.iterdir():
            if p.is_dir():
                issue('nested_input', p.name, ident)
                continue
            if p.suffix.lower() in ARCHIVES:
                issue('archive', p.name, ident)
            if p.name.startswith('~$') or p.name.endswith(('.inspect.ndjson', '.pyc')):
                issue('temporary_file', p.name, ident)
            if p.suffix.lower() in NATIVE:
                try:
                    with zipfile.ZipFile(p) as z:
                        if z.testzip() is not None:
                            raise ValueError('CRC failure')
                        ET.fromstring(z.read(NATIVE[p.suffix.lower()]))
                        if any(x.startswith('xl/externalLinks/') or 'vbaProject' in x for x in z.namelist()):
                            issue('native_dependency', p.name, ident)
                except (OSError, KeyError, ValueError, ET.ParseError, zipfile.BadZipFile) as exc:
                    issue('native_container', f'{p.name}: {exc}', ident)
            elif p.suffix.lower() == '.pdf':
                value = p.read_bytes()
                if not value.startswith(b'%PDF-'):
                    issue('pdf_header', p.name, ident)
            if p.suffix.lower() in TEXT:
                try:
                    text = p.read_text(encoding='utf-8-sig')
                    if '\ufffd' in text:
                        issue('text_corruption', p.name, ident)
                    if any(x in text for x in ['模拟数据', '仅供测试', '本题', '标准答案', '用于测评']):
                        issue('input_disclosure_review', f'{p.name}: review construction/evaluation wording in context', ident, True)
                except UnicodeError:
                    issue('text_encoding', f'{p.name}: expected UTF-8', ident)
        prompt_path = folder / '问题描述.txt'
        try:
            prompt = prompt_path.read_text(encoding='utf-8-sig').strip()
        except (OSError, UnicodeError):
            issue('prompt', 'Prompt missing or unreadable', ident)
            prompt = ''
        if not prompt:
            issue('prompt', 'Prompt is empty', ident)
        elif prompt.startswith(('#', ident + '_', ident + ' ', ident + '：', '问题描述.txt', str(case.get('title', '\x00')))):
            issue('prompt_header', 'Prompt should begin with the user request', ident)
        if '请将正式交付文件保存到当前工作目录' in prompt:
            issue('prompt_boilerplate', 'Remove generic save-directory ending', ident)
        if re.search(r'等我确认|我确认后|确认后再(?:执行|归档|制作)|先.{0,20}方案.{0,30}(?:确认|同意)', prompt, re.S):
            issue('confirmation_gate_review', 'Review whether task waits for an intermediate reply', ident, True)
        outs = case.get('outputs', [])
        if not isinstance(outs, list) or not outs:
            issue('output', 'Missing formal deliverable', ident)
            outs = []
        if len(outs) > 3:
            issue('output_count_review', 'Consider aggregating many deliverables', ident, True)
        for item in outs:
            if not isinstance(item, dict) or not filename(item.get('name')) or not item.get('description'):
                issue('output', 'Deliverable needs ordinary filename and content description', ident)
                continue
            name = item['name']
            if name.casefold() in deliveries:
                issue('duplicate_output', name, ident)
            deliveries.add(name.casefold())
            if name not in prompt:
                issue('output_reference', name, ident)
            if name.casefold() in {n.casefold() for n in names}:
                issue('output_overwrites_input', name, ident)
            if not name.startswith(ident + '_'):
                issue('output_prefix_review', name, ident, True)
            if Path(name).suffix.lower() in ARCHIVES:
                issue('compressed_output_review', name, ident, True)
    actual_cases = {p.parent.relative_to(root).as_posix().casefold() for p in root.glob('*/*/问题描述.txt')}
    if actual_cases != seen_folders:
        issue('case_coverage', 'Catalog directories do not cover exactly the prompt directories')
    if (root / '题目索引.csv').is_file():
        try:
            with (root / '题目索引.csv').open(encoding='utf-8-sig', newline='') as stream:
                index = list(csv.DictReader(stream))
            by_id = {row['编号']: row for row in index}
            if len(by_id) != len(index) or set(by_id) != seen_ids:
                issue('index_coverage', 'Index IDs differ from catalog')
            for case in cases:
                if not isinstance(case, dict) or not isinstance(case.get('id'), str) or case['id'] not in by_id:
                    continue
                row = by_id[case['id']]
                expected_inputs = sorted(x['name'] for x in case.get('inputs', []) if isinstance(x, dict) and isinstance(x.get('name'), str))
                expected_outputs = sorted(x['name'] for x in case.get('outputs', []) if isinstance(x, dict) and isinstance(x.get('name'), str))
                got_inputs = sorted(row['同目录输入文件'].split('；')) if row['同目录输入文件'] != '无附件' else []
                if row['相对目录'].replace('\\', '/') != case.get('folder') or row['题目'] != case.get('title') or got_inputs != expected_inputs or sorted(row['正式交付'].split('；')) != expected_outputs:
                    issue('index_content', 'Index folder/title/input/output differs from catalog', case['id'])
        except (OSError, ValueError, KeyError, TypeError, csv.Error) as exc:
            issue('index', str(exc))
    if verify_manifest and (root / '文件清单.csv').is_file():
        try:
            with (root / '文件清单.csv').open(encoding='utf-8-sig', newline='') as stream:
                manifest = list(csv.DictReader(stream))
            expected = {p.relative_to(root).as_posix() for p in root.rglob('*') if p.is_file() and p != root / '文件清单.csv'}
            listed = [r['相对路径'] for r in manifest]
            if len(set(listed)) != len(listed) or set(listed) != expected:
                issue('manifest_coverage', 'Manifest must cover each file exactly once, except itself')
            for item in manifest:
                rel = PurePosixPath(item['相对路径'])
                p = root.joinpath(*rel.parts)
                if rel.is_absolute() or '..' in rel.parts or '\\' in item['相对路径'] or not p.resolve().is_relative_to(root):
                    issue('manifest_path', item['相对路径'])
                elif not p.is_file() or p.stat().st_size != int(item['字节数']) or sha256(p) != item['SHA256']:
                    issue('manifest_hash', item['相对路径'])
        except (OSError, ValueError, KeyError, csv.Error) as exc:
            issue('manifest', str(exc))
    return {'structural_checks_passed': not errors, 'cases': len(cases), 'input_files': input_count,
            'deliverables': len(deliveries), 'errors': errors, 'warnings': warnings,
            'semantic_review_required': True, 'note': 'Does not certify business correctness, source quotations, layout or end-to-end execution.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', type=Path)
    parser.add_argument('--catalog', required=True, type=Path)
    parser.add_argument('--report', required=True, type=Path)
    parser.add_argument('--write-manifest', action='store_true', help='Explicitly rebuild root manifest after basic structural checks')
    args = parser.parse_args()
    if args.report.resolve().is_relative_to(args.root.resolve()) or args.report.resolve() == args.catalog.resolve():
        parser.error('Keep the check report outside the dataset and separate from the catalog')
    catalog = json.loads(args.catalog.read_text(encoding='utf-8-sig'))
    result = validate(args.root, catalog, not args.write_manifest)
    if args.write_manifest and not result['errors']:
        write_manifest(args.root)
        result = validate(args.root, catalog)
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: v for k, v in result.items() if k not in {'errors', 'warnings', 'note'}}, ensure_ascii=False))
    return 1 if result['errors'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
