"""Install the bundled skill for Claude Code, Codex, and DSH without dependencies."""
import argparse
import os
import re
import stat
import sys
import tempfile
from pathlib import Path


SKILL = 'docx-killer'
CACHES = {'__pycache__', '.pytest_cache', '.mypy_cache', '.ruff_cache'}


def is_link(path):
    try:
        attributes = path.lstat()
    except FileNotFoundError:
        return False
    return stat.S_ISLNK(attributes.st_mode) or bool(
        getattr(attributes, 'st_file_attributes', 0) & getattr(stat, 'FILE_ATTRIBUTE_REPARSE_POINT', 0x400))


def skill_files(root):
    def unreadable(error):
        raise error

    files = {}
    for directory, folders, names in os.walk(root, followlinks=False, onerror=unreadable):
        folders[:] = [name for name in folders if name not in CACHES]
        for name in folders + names:
            path = Path(directory) / name
            if is_link(path):
                raise ValueError('Refusing a symlink or junction: ' + str(path))
        for name in names:
            path = Path(directory) / name
            if path.suffix not in {'.pyc', '.pyo'}:
                files[path.relative_to(root)] = path.read_bytes()
    return files


def description_span(text):
    lines = text.splitlines(keepends=True)
    if not lines or lines[0].strip() != '---':
        return None
    start = None
    offset = len(lines[0])
    for line in lines[1:]:
        if line.rstrip('\r\n') == '---':
            return (start, offset) if start is not None else None
        if line.startswith('description:'):
            value = line[len('description:'):].strip()
            quoted = r"""(?:"(?:[^"\\]|\\.)*"|'(?:[^']|'')*')[ \t]*(?:#.*)?"""
            if value.startswith(('"', "'")) and not re.fullmatch(quoted, value):
                raise ValueError('Multiline quoted descriptions are unsupported; use an indented > or | block')
            start = offset
        elif start is not None and line.strip() and not line[0].isspace():
            return start, offset
        offset += len(line)
    return None


def preserve_dsh_description(source, previous):
    old = previous.decode('utf-8-sig')
    old_span = description_span(old)
    if old_span is None:
        return source
    new = source.decode('utf-8-sig')
    new_span = description_span(new)
    if new_span is None:
        raise ValueError('The bundled SKILL.md is missing its description')
    return (new[:new_span[0]] + old[old_span[0]:old_span[1]] + new[new_span[1]:]).encode('utf-8')


def install_skill(source, target, agent):
    source, target = Path(source).absolute(), Path(target).expanduser().absolute()
    if is_link(source) or not source.is_dir():
        raise ValueError('The bundled skill must be a regular directory: ' + str(source))
    if is_link(target) or (target.exists() and not target.is_dir()):
        raise ValueError('The target must be a regular directory: ' + str(target))
    resolved_source, resolved_target = source.resolve(), target.resolve()
    if resolved_source.is_relative_to(resolved_target) or resolved_target.is_relative_to(resolved_source):
        raise ValueError('Source and target must not contain each other')
    files = skill_files(source)
    if Path('SKILL.md') not in files:
        raise ValueError('The bundled skill is missing SKILL.md')
    previous = skill_files(target) if target.exists() else None
    if agent == 'dsh' and previous and Path('SKILL.md') in previous:
        files[Path('SKILL.md')] = preserve_dsh_description(files[Path('SKILL.md')], previous[Path('SKILL.md')])
    if previous == files:
        return 'unchanged', None

    target.parent.mkdir(parents=True, exist_ok=True)
    backup = None
    with tempfile.TemporaryDirectory(prefix='.docx-killer-install-', dir=target.parent.parent) as temporary:
        staged = Path(temporary) / SKILL
        staged.mkdir()
        for relative, content in files.items():
            file = staged / relative
            file.parent.mkdir(parents=True, exist_ok=True)
            file.write_bytes(content)
        if previous is not None:
            backup_root = target.parent.parent / 'skill-backups'
            if is_link(backup_root):
                raise ValueError('Refusing a redirected backup directory: ' + str(backup_root))
            backup_root.mkdir(exist_ok=True)
            backup = Path(tempfile.mkdtemp(prefix=SKILL + '-', dir=backup_root)) / SKILL
            target.rename(backup)
        try:
            staged.rename(target)
        except OSError as error:
            if backup is not None:
                try:
                    backup.rename(target)
                except OSError:
                    raise OSError('Installation failed; the previous skill is preserved at ' + str(backup)) from error
            raise
    return ('updated' if previous is not None else 'installed'), backup


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--agent', action='append', choices=('claude', 'codex', 'dsh'),
                        help='Install only for this agent; repeat to select several (default: all three)')
    args = parser.parse_args(argv)
    home = Path.home()
    configurations = {'claude': home / '.claude',
                      'codex': Path(os.environ.get('CODEX_HOME') or home / '.codex'),
                      'dsh': home / '.dsh'}
    source = Path(__file__).resolve().parent / 'skills' / SKILL
    failed = False
    for agent in dict.fromkeys(args.agent or configurations):
        target = configurations[agent].expanduser() / 'skills' / SKILL
        try:
            status, backup = install_skill(source, target, agent)
            print(f'{agent}: {status} -> {target}')
            if backup is not None:
                print(f'{agent}: backup -> {backup}')
        except (OSError, ValueError) as error:
            print(f'{agent}: {error}', file=sys.stderr)
            failed = True
    return int(failed)


if __name__ == '__main__':
    sys.exit(main())
