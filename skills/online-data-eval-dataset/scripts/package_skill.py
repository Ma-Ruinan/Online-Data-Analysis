"""Build a deterministic .skill ZIP and SHA256 sidecar from one skill folder."""
from __future__ import annotations
import argparse
import hashlib
import io
from pathlib import Path
import re
import zipfile


def package(skill: Path, version: str, output: Path) -> dict:
    skill = skill.resolve()
    output = output.resolve()
    if not re.fullmatch(r'\d+\.\d+\.\d+(?:-[a-zA-Z0-9.-]+)?', version):
        raise ValueError('Use a semantic version such as 1.0.0')
    if not re.fullmatch(r'[a-z0-9-]{1,63}', skill.name) or not (skill / 'SKILL.md').is_file():
        raise ValueError('Not a named skill directory containing SKILL.md')
    header = (skill / 'SKILL.md').read_text(encoding='utf-8-sig').split('---', 2)
    declared = re.search(r'^\s+version:\s*[\'\"]?([^\s\'\"]+)', header[1] if len(header) == 3 else '', re.M)
    if declared and declared.group(1) != version:
        raise ValueError('Package version differs from skill metadata')
    if output.is_relative_to(skill):
        raise ValueError('Package output must be outside the skill folder')
    files = []
    for p in sorted(skill.rglob('*')):
        if p.is_symlink():
            raise ValueError(f'Refuse symlink: {p}')
        if not p.is_file() or any(x in {'.git', '__pycache__', '.venv'} for x in p.relative_to(skill).parts) or p.suffix in {'.pyc', '.pyo'}:
            continue
        files.append(p)
    output.mkdir(parents=True, exist_ok=True)
    dest = output / f'{skill.name}-{version}.skill'
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for p in files:
            entry = zipfile.ZipInfo(skill.name + '/' + p.relative_to(skill).as_posix(), date_time=(1980, 1, 1, 0, 0, 0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.create_system = 3
            entry.external_attr = 0o100644 << 16
            z.writestr(entry, p.read_bytes(), compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    payload = buffer.getvalue()
    digest = hashlib.sha256(payload).hexdigest()
    if dest.exists() and dest.read_bytes() != payload:
        raise FileExistsError('This version already has different bytes; bump the skill/package version')
    dest.write_bytes(payload)
    dest.with_suffix(dest.suffix + '.sha256').write_text(f'{digest}  {dest.name}\n', encoding='utf-8')
    with zipfile.ZipFile(dest) as z:
        assert z.testzip() is None and f'{skill.name}/SKILL.md' in z.namelist()
    return {'file': str(dest), 'sha256': digest, 'skill_files': len(files)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('skill', type=Path)
    parser.add_argument('--version', required=True)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    import json
    print(json.dumps(package(args.skill, args.version, args.output), ensure_ascii=False))


if __name__ == '__main__':
    main()
