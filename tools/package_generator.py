"""Package only maintained generator files; exclude all personal report inputs."""
import argparse
from pathlib import Path
import zipfile

def package(root, destination):
    files = [root/name for name in ('AGENTS.md', 'README.md', 'requirements.txt', '.gitignore', 'config/student.example.yaml')]
    for folder, extensions in [('docs', {'.md'}), ('skills', {'.md', '.yaml'}), ('tools', {'.py', '.ps1'}), ('references/official', {'.pdf', '.docx', '.md', '.txt'})]:
        files.extend(p for p in (root/folder).rglob('*') if p.is_file() and p.suffix.lower() in extensions and '__pycache__' not in p.parts)
    for path in files:
        if path.is_symlink() or not path.resolve().is_relative_to(root.resolve()):
            raise ValueError('Refusing to package a linked/outside file: ' + str(path))
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(set(files)):
            archive.write(path, path.relative_to(root).as_posix())
        archive.write(root/'config/student.example.yaml', 'config/student.yaml')
        for folder in ('project', 'references/current-example', 'references/previous-reports',
                       'screenshots/raw', 'screenshots/selected', 'temp', 'output'):
            archive.writestr(folder+'/.gitkeep', '')
    return destination

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path, default=Path('output/generator.zip'))
    args = parser.parse_args()
    print(package(Path(__file__).resolve().parents[1], args.output).resolve())
