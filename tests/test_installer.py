import importlib.util
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


INSTALLER = Path(__file__).resolve().parents[1] / 'install.py'


class Installer(unittest.TestCase):
    def setUp(self):
        self.assertTrue(INSTALLER.is_file(), 'The repository needs a standalone install.py')
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.repo = self.root / 'checkout'
        self.source = self.repo / 'skills' / 'docx-killer'
        self.source.mkdir(parents=True)
        self.home = self.root / 'home'
        self.home.mkdir()
        self.script = self.repo / 'install.py'
        self.script.write_bytes(INSTALLER.read_bytes())
        self.files = {
            'SKILL.md': b'---\nname: docx-killer\ndescription: Standard description\n---\n\nNew instructions.\n',
            'requirements.txt': b'example>=1\n',
            'scripts/tool.py': b'print("new tool")\n',
            'references/guide.md': b'New reference.\n',
        }
        for name, content in self.files.items():
            file = self.source / name
            file.parent.mkdir(parents=True, exist_ok=True)
            file.write_bytes(content)
        self.environment = dict(os.environ, HOME=str(self.home), USERPROFILE=str(self.home))
        self.environment.pop('CODEX_HOME', None)
        specification = importlib.util.spec_from_file_location('skill_installer', INSTALLER)
        self.installer = importlib.util.module_from_spec(specification)
        specification.loader.exec_module(self.installer)

    def run_cli(self, *arguments, environment=None):
        return subprocess.run([sys.executable, '-B', str(self.script), *arguments],
                              cwd=self.root, env=environment or self.environment,
                              capture_output=True, text=True, encoding='utf-8')

    def target(self, agent):
        return self.home / ('.' + agent) / 'skills' / 'docx-killer'

    def assert_files(self, target):
        actual = {str(file.relative_to(target)).replace('\\', '/'): file.read_bytes()
                  for file in target.rglob('*') if file.is_file()}
        self.assertEqual(actual, self.files)

    def test_defaults_install_all_three_agents_from_any_working_directory(self):
        result = self.run_cli()
        self.assertEqual(result.returncode, 0, result.stderr)
        for agent in ('claude', 'codex', 'dsh'):
            self.assert_files(self.target(agent))
            self.assertIn(agent + ': installed', result.stdout)

    def test_agent_selection_is_repeatable_and_deduplicated(self):
        result = self.run_cli('--agent', 'claude', '--agent', 'claude', '--agent', 'dsh')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assert_files(self.target('claude'))
        self.assert_files(self.target('dsh'))
        self.assertFalse(self.target('codex').exists())
        self.assertEqual(result.stdout.count('claude: installed'), 1)

    def test_single_agent_and_codex_home_override(self):
        custom = self.root / 'custom codex'
        result = self.run_cli('--agent', 'codex', environment=dict(self.environment, CODEX_HOME=str(custom)))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assert_files(custom / 'skills' / 'docx-killer')
        for agent in ('claude', 'codex', 'dsh'):
            self.assertFalse(self.target(agent).exists())

    def test_unchanged_install_does_not_write_or_create_backups(self):
        self.assertEqual(self.run_cli('--agent', 'codex').returncode, 0)
        target = self.target('codex')
        before = {file: file.stat().st_mtime_ns for file in target.rglob('*') if file.is_file()}
        cache = target / '__pycache__'
        cache.mkdir()
        (cache / 'tool.pyc').write_bytes(b'local cache')
        result = self.run_cli('--agent', 'codex')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn('codex: unchanged', result.stdout)
        self.assertEqual({file: file.stat().st_mtime_ns for file in before}, before)
        self.assertFalse((self.home / '.codex' / 'skill-backups').exists())

    def test_update_backs_up_the_complete_previous_directory_outside_skills(self):
        target = self.target('claude')
        target.mkdir(parents=True)
        (target / 'SKILL.md').write_text('Old local instructions', encoding='utf-8')
        (target / 'local-only.txt').write_bytes(b'keep this in the backup')
        result = self.run_cli('--agent', 'claude')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assert_files(target)
        backups = list((self.home / '.claude' / 'skill-backups').rglob('local-only.txt'))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), b'keep this in the backup')
        self.assertEqual((backups[0].parent / 'SKILL.md').read_text(encoding='utf-8'), 'Old local instructions')
        self.assertIn('backup', result.stdout)

    def test_dsh_preserves_only_its_single_or_multiline_description(self):
        for description in ('description: "Use office for new documents"\n',
                            'description: "Use \\"office\\" and \\\\ paths" # keep comment\n',
                            "description: 'Use ''office'' and # hashtags' # keep comment\n",
                            'description: >-\n  Use office for new documents.\n  Preserve this second line.\n',
                            'description: |-\n  Use office for new documents.\n  Preserve this second line.\n'):
            with self.subTest(description=description):
                target = self.target('dsh')
                target.mkdir(parents=True, exist_ok=True)
                old = '---\nname: old-name\n' + description + 'extra: local\n---\nOld body.\n'
                (target / 'SKILL.md').write_text(old, encoding='utf-8')
                result = self.run_cli('--agent', 'dsh')
                self.assertEqual(result.returncode, 0, result.stderr)
                expected = self.files['SKILL.md'].decode().replace('description: Standard description\n', description)
                self.assertEqual((target / 'SKILL.md').read_text(encoding='utf-8'), expected)
                second = self.run_cli('--agent', 'dsh')
                self.assertEqual(second.returncode, 0, second.stderr)
                self.assertIn('dsh: unchanged', second.stdout)

    def test_multiline_quoted_dsh_descriptions_are_refused_without_changing_the_old_skill(self):
        for description in ('description: "First line\nsecond line"\n',
                            "description: 'First line\nsecond line'\n",
                            'description: "First line\n  second line"\n'):
            with self.subTest(description=description):
                target = self.target('dsh')
                target.mkdir(parents=True, exist_ok=True)
                old = ('---\nname: docx-killer\n' + description + '---\nOld body.\n').encode('utf-8')
                (target / 'SKILL.md').write_bytes(old)
                result = self.run_cli('--agent', 'dsh')
                self.assertNotEqual(result.returncode, 0)
                self.assertIn('indented > or | block', result.stderr)
                self.assertEqual((target / 'SKILL.md').read_bytes(), old)
                self.assertFalse((target / 'scripts').exists())
                self.assertFalse((self.home / '.dsh' / 'skill-backups').exists())

    def test_source_caches_are_not_distributed(self):
        for name in ('__pycache__/tool.pyc', 'scripts/__pycache__/tool.pyc',
                     'scripts/old.pyo', '.pytest_cache/cache'):
            file = self.source / name
            file.parent.mkdir(parents=True, exist_ok=True)
            file.write_bytes(b'cache')
        result = self.run_cli('--agent', 'claude')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assert_files(self.target('claude'))

    def test_unreadable_source_directory_does_not_install_an_incomplete_skill(self):
        scandir = os.scandir

        def unreadable(path):
            if Path(path) == self.source / 'references':
                raise PermissionError('source directory is unreadable')
            return scandir(path)

        with mock.patch.object(self.installer.os, 'scandir', unreadable):
            with self.assertRaisesRegex(PermissionError, 'source directory is unreadable'):
                self.installer.install_skill(self.source, self.target('claude'), 'claude')
        self.assertFalse(self.target('claude').exists())

    def test_failed_publication_restores_the_previous_directory(self):
        target = self.target('claude')
        target.mkdir(parents=True)
        (target / 'SKILL.md').write_bytes(b'previous instructions')
        (target / 'local.txt').write_bytes(b'previous local file')
        rename = Path.rename

        def fail_publication(path, destination):
            if path.parent.name.startswith('.docx-killer-install-'):
                raise PermissionError('simulated publication failure')
            return rename(path, destination)

        with mock.patch.object(Path, 'rename', fail_publication):
            with self.assertRaisesRegex(OSError, 'simulated publication failure'):
                self.installer.install_skill(self.source, target, 'claude')
        self.assertEqual((target / 'SKILL.md').read_bytes(), b'previous instructions')
        self.assertEqual((target / 'local.txt').read_bytes(), b'previous local file')
        self.assertFalse((target / 'scripts').exists())

    def test_file_target_is_refused_and_cli_reports_failure(self):
        target = self.target('codex')
        target.parent.mkdir(parents=True)
        target.write_bytes(b'user file')
        result = self.run_cli('--agent', 'codex')
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('codex:', result.stderr)
        self.assertEqual(target.read_bytes(), b'user file')

    def test_source_and_target_cannot_contain_each_other(self):
        for target in (self.source, self.source / 'nested', self.repo):
            with self.subTest(target=target):
                with self.assertRaises(ValueError):
                    self.installer.install_skill(self.source, target, 'codex')
        self.assertEqual((self.source / 'SKILL.md').read_bytes(), self.files['SKILL.md'])

    def test_symbolic_links_are_never_followed(self):
        outside = self.root / 'outside'
        outside.mkdir()
        (outside / 'keep.txt').write_bytes(b'outside data')
        target = self.target('codex')
        target.parent.mkdir(parents=True)
        try:
            target.symlink_to(outside, target_is_directory=True)
        except OSError as error:
            self.skipTest('Creating symlinks is unavailable: ' + str(error))
        with self.assertRaises(ValueError):
            self.installer.install_skill(self.source, target, 'codex')
        target.unlink()
        (self.source / 'linked').symlink_to(outside, target_is_directory=True)
        with self.assertRaises(ValueError):
            self.installer.install_skill(self.source, target, 'codex')
        self.assertFalse(target.exists())
        self.assertEqual((outside / 'keep.txt').read_bytes(), b'outside data')

    @unittest.skipUnless(os.name == 'nt', 'Junctions are Windows-only')
    def test_source_and_target_junctions_are_refused(self):
        outside = self.root / 'outside'
        outside.mkdir()
        (outside / 'keep.txt').write_bytes(b'outside data')
        target = self.target('codex')
        target.parent.mkdir(parents=True)
        result = subprocess.run(['cmd', '/c', 'mklink', '/J', str(target), str(outside)], capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        with mock.patch.object(Path, 'is_junction', return_value=False, create=True):
            with self.assertRaises(ValueError):
                self.installer.install_skill(self.source, target, 'codex')
        target.rmdir()
        result = subprocess.run(['cmd', '/c', 'mklink', '/J', str(self.source / 'linked'), str(outside)], capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        with self.assertRaises(ValueError):
            self.installer.install_skill(self.source, target, 'codex')
        self.assertFalse(target.exists())
        self.assertEqual((outside / 'keep.txt').read_bytes(), b'outside data')


if __name__ == '__main__':
    unittest.main()
