"""Locate the four DOCX tools in the current Python/Windows environment."""
import importlib.metadata
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path


def find_pandoc():
    override = os.environ.get('PANDOC_BINARY')
    if override:
        return str(Path(override)) if Path(override).is_file() else None
    found = shutil.which('pandoc')
    if found:
        return found
    candidates = []
    for variable in ('LOCALAPPDATA', 'ProgramFiles', 'ProgramFiles(x86)'):
        if os.environ.get(variable):
            candidates.append(Path(os.environ[variable]) / 'Pandoc' / 'pandoc.exe')
    return next((str(path) for path in candidates if path.is_file()), None)


def diagnose():
    tools = {}
    for name in ('python-docx', 'docxtpl', 'lxml', 'jinja2'):
        try:
            tools[name] = {'available': True, 'version': importlib.metadata.version(name)}
        except importlib.metadata.PackageNotFoundError:
            tools[name] = {'available': False, 'reason': 'Not installed in this Python interpreter'}
    executable = find_pandoc()
    tools['pandoc'] = {'available': False, 'path': executable}
    if executable:
        try:
            result = subprocess.run([executable, '--version'], capture_output=True, text=True, encoding='utf-8', timeout=15, check=True)
            tools['pandoc'].update(available=True, version=result.stdout.splitlines()[0])
        except (OSError, subprocess.SubprocessError) as exc:
            tools['pandoc']['reason'] = str(exc)
    node = shutil.which('node')
    tools['mammoth'] = {'available': False, 'runtime': 'Node.js', 'node': node}
    if node:
        try:
            result = subprocess.run([node, str(Path(__file__).with_name('mammoth_extract.cjs')), '--doctor'], capture_output=True, text=True, encoding='utf-8', timeout=20, check=True)
            tools['mammoth'].update(json.loads(result.stdout), available=True)
        except (OSError, ValueError, subprocess.SubprocessError) as exc:
            tools['mammoth']['reason'] = str(exc)
    else:
        tools['mammoth']['reason'] = 'Node.js not found on PATH'
    renderers = {'libreoffice': shutil.which('soffice'), 'word_registered': False}
    if sys.platform == 'win32':
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Word.Application\CLSID'):
                renderers['word_registered'] = True
        except OSError:
            pass
        for variable in ('ProgramFiles', 'ProgramFiles(x86)'):
            base = os.environ.get(variable)
            if base:
                path = Path(base) / 'LibreOffice' / 'program' / 'soffice.exe'
                if path.is_file():
                    renderers['libreoffice'] = str(path)
    return {'python': sys.executable, 'python_version': sys.version.split()[0], 'tools': tools, 'renderers': renderers, 'note': 'Installed tools do not establish Word/WPS layout fidelity; renderer discovery does not execute an office application'}


if __name__ == '__main__':
    report = diagnose()
    print(json.dumps(report, ensure_ascii=True, indent=2))
    sys.exit(0 if all(report['tools'][name]['available'] for name in ('python-docx', 'docxtpl', 'pandoc', 'mammoth')) else 1)
