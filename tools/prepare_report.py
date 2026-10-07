"""Check report inputs and fingerprint sources without running the application."""
import argparse
from datetime import date
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import sys

try:
    import yaml
except ModuleNotFoundError:
    raise SystemExit('Missing PyYAML. Use a Python environment with requirements.txt installed.')

EXCLUDED = {'.git', '.vs', '.idea', 'node_modules', 'bin', 'obj', 'dist', 'build',
            'coverage', '__pycache__', '.venv', 'venv', 'temp', 'output'}
DOCUMENTS = {'.docx', '.pdf', '.txt', '.md'}
REQUIRED = {
    'student': ('full_name', 'group', 'specialty_code', 'specialty_name'),
    'institution': ('name', 'city'),
    'practice': ('kind', 'module_code', 'practice_code', 'title', 'start_date', 'end_date'),
    'project': ('path', 'topic', 'assignment'),
}

def sha256(path):
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()

def documents(directory):
    return sorted(p for p in directory.rglob('*') if p.is_file() and p.suffix.lower() in DOCUMENTS)

def inside(root, value):
    path = (root / value).resolve()
    if not path.is_relative_to(root.resolve()):
        raise ValueError('Input must be inside the generator folder: ' + value)
    return path

def prepare(root, config_path):
    issues = []
    try:
        config = yaml.safe_load(config_path.read_text(encoding='utf-8-sig'))
        if not isinstance(config, dict):
            raise ValueError('Config must contain mappings.')
    except (OSError, ValueError, yaml.YAMLError) as error:
        return {'status': 'needs-input', 'issues': [str(error)], 'fingerprints': {}}
    for group, fields in REQUIRED.items():
        section = config.get(group)
        if not isinstance(section, dict):
            issues.append('Missing mapping: ' + group)
            continue
        for field in fields:
            if not isinstance(section.get(field), str) or not section[field].strip():
                issues.append('Missing value: ' + group + '.' + field)
    practice = config.get('practice') if isinstance(config.get('practice'), dict) else {}
    supervisors = practice.get('supervisors')
    if not isinstance(supervisors, list) or not supervisors or any(not isinstance(s, str) or not s.strip() for s in supervisors):
        issues.append('practice.supervisors must contain confirmed names.')
    dates = {}
    for field in ('start_date', 'end_date'):
        value = practice.get(field)
        if value:
            try:
                dates[field] = date.fromisoformat(str(value))
            except ValueError:
                issues.append('Invalid YYYY-MM-DD date: practice.' + field)
    if len(dates) == 2 and dates['end_date'] < dates['start_date']:
        issues.append('Practice end date precedes start date.')

    fingerprints = {config_path.relative_to(root).as_posix(): sha256(config_path)}
    def remember(path):
        fingerprints[path.relative_to(root).as_posix()] = sha256(path)
    for path in [root/'AGENTS.md', *documents(root/'docs'), *documents(root/'skills'), *documents(root/'references/official')]:
        if path.is_file():
            remember(path)
    if not documents(root/'references/official'):
        issues.append('No official requirements found.')

    refs = config.get('references') if isinstance(config.get('references'), dict) else {}
    selected = {}
    for key, folder, required in [('current_example', 'current-example', True), ('style_example', 'previous-reports', False)]:
        candidates = documents(root/'references'/folder)
        value = refs.get(key)
        if value:
            try:
                chosen = inside(root, value)
                if chosen not in candidates:
                    raise ValueError('Selected reference missing or outside its reference folder: ' + str(value))
            except (ValueError, TypeError) as error:
                issues.append(str(error))
                continue
        elif len(candidates) == 1:
            chosen = candidates[0]
        else:
            if required or len(candidates) > 1:
                issues.append('Select references.' + key + f' ({len(candidates)} candidates).')
            continue
        selected[key] = chosen.relative_to(root).as_posix()
        remember(chosen)

    project = config.get('project') if isinstance(config.get('project'), dict) else {}
    for key in ('path', 'assignment'):
        value = project.get(key)
        if not isinstance(value, str) or not value.strip():
            continue
        try:
            path = inside(root, value)
            if key == 'assignment':
                if not path.is_file():
                    raise ValueError('Assignment file missing: ' + value)
                remember(path)
            else:
                if not path.is_dir() or not path.is_relative_to(root/'project'):
                    raise ValueError('project.path must be a directory under project/.')
                count = 0
                for directory, subdirs, names in os.walk(path, followlinks=False):
                    subdirs[:] = sorted(s for s in subdirs if s not in EXCLUDED and not (Path(directory)/s).is_symlink())
                    for name in sorted(names):
                        file = Path(directory)/name
                        if file.is_symlink() or name == '.gitkeep' or name.startswith('.env'):
                            continue
                        remember(file)
                        count += 1
                if not count:
                    issues.append('Project directory has no source files.')
        except (ValueError, OSError, TypeError) as error:
            issues.append(str(error))
    environment = {'python': sys.version.split()[0],
        'libraries': {name: importlib.util.find_spec(name) is not None for name in
                      ('docx', 'lxml', 'yaml', 'pdfplumber', 'pypdf', 'pypdfium2')},
        'libreoffice': shutil.which('soffice') or shutil.which('libreoffice'),
        'word_com_candidate': sys.platform == 'win32'}
    signature = hashlib.sha256(json.dumps(fingerprints, sort_keys=True).encode()).hexdigest()
    return {'status': 'needs-input' if issues else 'ready', 'issues': issues,
            'selected_references': selected, 'environment': environment,
            'input_signature': signature, 'fingerprints': fingerprints,
            'execution_verified': False}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--config', type=Path, default=Path('config/student.yaml'))
    parser.add_argument('--json', type=Path)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    config_path = inside(root, str(args.config))
    result = prepare(root, config_path)
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(payload, encoding='utf-8')
    print(payload)
    return 0 if result['status'] == 'ready' else 2

if __name__ == '__main__':
    raise SystemExit(main())
