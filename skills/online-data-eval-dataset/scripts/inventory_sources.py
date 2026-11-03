"""Lossless source inventory; deliberately does not select or classify tasks."""
from __future__ import annotations
import argparse
import csv
import hashlib
import io
import json
import posixpath
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

SUPPORTED = {'.md', '.txt', '.csv', '.jsonl', '.xlsx'}
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def xlsx_rows(path: Path):
    """Keep XML cell values, formulas and formats; do not guess Excel date meaning."""
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        if bad:
            raise ValueError(f'Corrupt ZIP member: {bad}')
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            for item in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si', NS):
                strings.append(''.join(t.text or '' for t in item.findall('.//s:t', NS)))
        formats = {}
        style_ids = []
        if 'xl/styles.xml' in z.namelist():
            styles = ET.fromstring(z.read('xl/styles.xml'))
            formats = {x.attrib['numFmtId']: x.attrib['formatCode'] for x in styles.findall('s:numFmts/s:numFmt', NS)}
            style_ids = [x.attrib.get('numFmtId', '0') for x in styles.findall('s:cellXfs/s:xf', NS)]
        rels = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        workbook = ET.fromstring(z.read('xl/workbook.xml'))
        props = workbook.find('s:workbookPr', NS)
        epoch = '1904' if props is not None and props.attrib.get('date1904') in {'1', 'true'} else '1900'
        for sheet in workbook.findall('s:sheets/s:sheet', NS):
            rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            target = rels[rid]
            member = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            if not member.startswith('xl/worksheets/'):
                raise ValueError(f'Non-worksheet sheet requires separate review: {sheet.attrib["name"]}')
            header = None
            with z.open(member) as stream:
                for event, element in ET.iterparse(stream, events=('end',)):
                    if element.tag != '{' + NS['s'] + '}row':
                        continue
                    cells = {}
                    for cell in element.findall('s:c', NS):
                        address = cell.attrib['r']
                        kind = cell.attrib.get('t', 'n')
                        value_el = cell.find('s:v', NS)
                        value = value_el.text if value_el is not None else None
                        if kind == 's' and value is not None:
                            value = strings[int(value)]
                        elif kind == 'inlineStr':
                            value = ''.join(t.text or '' for t in cell.findall('.//s:t', NS))
                        formula = cell.find('s:f', NS)
                        if value is None and formula is None:
                            continue
                        item = {'value': value, 'type': kind}
                        if formula is not None:
                            item['formula'] = formula.text or ''
                            item['formula_attributes'] = dict(formula.attrib)
                        style = int(cell.attrib.get('s', '0'))
                        if style < len(style_ids):
                            fmt_id = style_ids[style]
                            item['number_format_id'] = fmt_id
                            if fmt_id in formats:
                                item['number_format'] = formats[fmt_id]
                        cells[address] = item
                    if cells:
                        if header is None:
                            header = {re_column(a): c['value'] for a, c in cells.items()}
                        yield {'source_type': 'xlsx', 'locator': {'sheet': sheet.attrib['name'], 'row': int(element.attrib['r'])},
                               'sheet_state': sheet.attrib.get('state', 'visible'), 'date_system': epoch,
                               'header_cells': header, 'cells': cells}
                    element.clear()


def re_column(address: str) -> str:
    return ''.join(x for x in address if x.isalpha())


def records(path: Path, encoding: str = 'utf-8-sig'):
    if path.suffix.lower() == '.xlsx':
        yield from xlsx_rows(path)
        return
    text = path.read_text(encoding=encoding)
    suffix = path.suffix.lower()
    if suffix == '.csv':
        reader = csv.reader(io.StringIO(text, newline=''))
        for logical_row, values in enumerate(reader, 1):
            yield {'source_type': 'csv', 'locator': {'record': logical_row, 'line_end': reader.line_num}, 'values': values}
    elif suffix == '.jsonl':
        for line_number, line in enumerate(text.splitlines(), 1):
            if line.strip():
                yield {'source_type': 'jsonl', 'locator': {'line': line_number}, 'raw_line': line, 'value': json.loads(line)}
    else:
        headings = []
        for line_number, line in enumerate(text.splitlines(), 1):
            if suffix == '.md' and line.startswith('#'):
                level = len(line) - len(line.lstrip('#'))
                if level <= 6 and len(line) > level and line[level] == ' ':
                    headings = [(n, h) for n, h in headings if n < level] + [(level, line[level:].strip())]
            if line.strip():
                yield {'source_type': suffix[1:], 'locator': {'line': line_number},
                       'headings': [h for n, h in headings], 'raw_line': line}


def inventory(inputs, output: Path, recursive=False, encoding='utf-8-sig') -> dict:
    output = output.resolve()
    metadata = output.with_suffix('.inventory.json')
    if output.exists() or metadata.exists():
        raise FileExistsError('Choose a fresh output path; inventory does not overwrite source or prior files')
    candidates = set()
    for value in inputs:
        p = Path(value).resolve()
        if p.is_file():
            candidates.add(p)
        elif p.is_dir():
            candidates.update(x for x in (p.rglob('*') if recursive else p.iterdir()) if x.is_file())
        else:
            raise FileNotFoundError(p)
    output.parent.mkdir(parents=True, exist_ok=True)
    statuses = []
    total = 0
    with output.open('w', encoding='utf-8', newline='\n') as dest:
        for path in sorted(candidates):
            if path in {output, metadata} or path.name.startswith('~$') or any(x in {'.git', '__pycache__'} for x in path.parts):
                continue
            item = {'file': str(path), 'sha256': sha256(path), 'bytes': path.stat().st_size, 'records': 0}
            if path.is_symlink():
                item.update(status='unsupported', error='Symlink requires explicit source review')
            elif path.suffix.lower() not in SUPPORTED:
                item.update(status='unsupported', error='Use an appropriate parser; file was not covered')
            else:
                try:
                    for value in records(path, encoding):
                        value.update(file=str(path), sha256=item['sha256'])
                        dest.write(json.dumps(value, ensure_ascii=False) + '\n')
                        item['records'] += 1
                        total += 1
                    item['status'] = 'ok'
                except (OSError, ValueError, KeyError, IndexError, UnicodeError, ET.ParseError, zipfile.BadZipFile) as error:
                    item.update(status='error', error=f'{type(error).__name__}: {error}')
            statuses.append(item)
    result = {'files': statuses, 'record_count': total, 'complete': bool(statuses) and all(x['status'] == 'ok' for x in statuses),
              'note': 'Raw values and locators only. No semantic selection or task quality certification.'}
    metadata.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('inputs', nargs='+')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--recursive', action='store_true')
    parser.add_argument('--encoding', default='utf-8-sig')
    args = parser.parse_args()
    result = inventory(args.inputs, args.output, args.recursive, args.encoding)
    print(json.dumps({'files': len(result['files']), 'records': result['record_count'], 'complete': result['complete']}, ensure_ascii=False))
    return 0 if result['complete'] else 2


if __name__ == '__main__':
    raise SystemExit(main())
