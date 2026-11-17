import contextlib
import importlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from docx import Document


SCRIPTS = Path(__file__).resolve().parents[1] / 'skills' / 'docx-killer' / 'scripts'
sys.path.insert(0, str(SCRIPTS))
ops = importlib.import_module('docx_ops')
doctor = importlib.import_module('doctor')
POWERSHELL = shutil.which('powershell')


class Rendering(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.source = self.root / '报告.docx'
        Document().save(self.source)
        self.output = self.root / 'rendered'
        self.output.mkdir()
        self.pdf = self.output / '报告.pdf'

    def libreoffice(self, runner):
        stack = contextlib.ExitStack()
        self.addCleanup(stack.close)
        stack.enter_context(mock.patch.dict(os.environ, {'ProgramFiles': '', 'ProgramFiles(x86)': ''}))
        stack.enter_context(mock.patch.object(ops.shutil, 'which', side_effect=lambda name: 'soffice' if name == 'soffice' else None))
        stack.enter_context(mock.patch.object(ops, '_office_progids', return_value=[]))
        stack.enter_context(mock.patch.object(ops.subprocess, 'run', side_effect=runner))

    def test_existing_pdf_is_refused_without_altering_it(self):
        self.pdf.write_bytes(b'previous PDF')
        self.libreoffice(lambda *args, **kwargs: subprocess.CompletedProcess(args[0], 0, '', ''))
        with self.assertRaisesRegex(FileExistsError, 'already exists'):
            ops._render_document(self.source, self.output, 'libreoffice')
        self.assertEqual(self.pdf.read_bytes(), b'previous PDF')

    def test_libreoffice_exports_in_empty_staging_directory_and_publishes(self):
        def export(command, **kwargs):
            destination = Path(command[command.index('--outdir') + 1])
            self.assertNotEqual(destination, self.output)
            self.assertEqual(list(destination.iterdir()), [])
            (destination / '报告.pdf').write_bytes(b'new PDF')
            return subprocess.CompletedProcess(command, 0, '', '')

        self.libreoffice(export)
        report = ops._render_document(self.source, self.output, 'libreoffice')
        self.assertEqual(report['status'], 'rendered')
        self.assertEqual(Path(report['pdf']), self.pdf)
        self.assertEqual(self.pdf.read_bytes(), b'new PDF')

    def test_failed_export_does_not_publish_partial_pdf(self):
        def export(command, **kwargs):
            destination = Path(command[command.index('--outdir') + 1])
            (destination / '报告.pdf').write_bytes(b'partial')
            return subprocess.CompletedProcess(command, 1, '', 'export failed')

        self.libreoffice(export)
        with self.assertRaisesRegex(RuntimeError, 'export failed'):
            ops._render_document(self.source, self.output, 'libreoffice')
        self.assertFalse(self.pdf.exists())

    def test_staging_cleanup_failure_retains_published_pdf(self):
        def export(command, **kwargs):
            destination = Path(command[command.index('--outdir') + 1])
            (destination / self.pdf.name).write_bytes(b'exported PDF')
            return subprocess.CompletedProcess(command, 0, '', '')

        temporary_directory = tempfile.TemporaryDirectory

        @contextlib.contextmanager
        def cleanup_failure(*args, **kwargs):
            with temporary_directory(*args, **kwargs) as staging:
                yield staging
            raise PermissionError('staging file still in use')

        self.libreoffice(export)
        with mock.patch.object(ops.tempfile, 'TemporaryDirectory', side_effect=cleanup_failure):
            report = ops._render_document(self.source, self.output, 'auto')
        self.assertEqual(report['status'], 'rendered')
        self.assertEqual(Path(report['pdf']).read_bytes(), b'exported PDF')
        self.assertIn('staging file still in use', report['cleanup_warning'])
        self.assertEqual(report['attempts'][0]['status'], 'rendered')

    def test_auto_preserves_export_failure_instead_of_reporting_unavailable(self):
        self.libreoffice(lambda *args, **kwargs: subprocess.CompletedProcess(args[0], 1, '', 'export failed'))
        report = ops.verify_document(self.source, output_dir=self.output)
        self.assertEqual(report['status'], 'passed')
        self.assertEqual(report['render']['status'], 'failed')
        self.assertIn('export failed', '; '.join(report['render']['errors']))
        self.assertEqual(report['render']['attempts'][0]['status'], 'failed')

    def test_auto_distinguishes_missing_renderer_from_failed_attempt(self):
        with mock.patch.dict(os.environ, {'ProgramFiles': '', 'ProgramFiles(x86)': ''}), mock.patch.object(ops.shutil, 'which', return_value=None):
            report = ops._render_document(self.source, self.output, 'auto')
        self.assertEqual(report['status'], 'unavailable')
        self.assertEqual(report.get('attempts'), [])
        self.assertEqual(report.get('errors'), [])

    def test_success_exit_without_a_new_pdf_is_a_failed_attempt(self):
        self.libreoffice(lambda *args, **kwargs: subprocess.CompletedProcess(args[0], 0, '', ''))
        report = ops._render_document(self.source, self.output, 'auto')
        self.assertEqual(report['status'], 'failed')
        self.assertIn('PDF', '; '.join(report['errors']))

    @unittest.skipUnless(os.name == 'nt', 'COM rendering is Windows-only')
    def test_auto_fallback_keeps_failed_word_attempt_when_libreoffice_succeeds(self):
        def run(command, **kwargs):
            if command[0] == 'powershell':
                return subprocess.CompletedProcess(command, 1, '', 'COM export failed')
            destination = Path(command[command.index('--outdir') + 1])
            (destination / '报告.pdf').write_bytes(b'PDF')
            return subprocess.CompletedProcess(command, 0, '', '')

        with mock.patch.dict(os.environ, {'ProgramFiles': '', 'ProgramFiles(x86)': ''}), mock.patch.object(ops.shutil, 'which', side_effect=lambda name: name if name in {'powershell', 'soffice'} else None), mock.patch.object(ops, '_office_progids', return_value=['KWPS.Application']), mock.patch.object(ops.subprocess, 'run', side_effect=run):
            report = ops._render_document(self.source, self.output, 'auto')
        self.assertEqual(report['status'], 'rendered')
        self.assertIn('COM export failed', '; '.join(report.get('errors', [])))
        self.assertEqual([item['status'] for item in report['attempts']], ['failed', 'rendered'])

    def test_auto_reports_timeout_as_a_failed_attempt(self):
        self.libreoffice(lambda *args, **kwargs: (_ for _ in ()).throw(subprocess.TimeoutExpired(args[0], 120)))
        try:
            report = ops._render_document(self.source, self.output, 'auto')
        except subprocess.TimeoutExpired:
            self.fail('auto rendering must retain timeout diagnostics in its report')
        self.assertEqual(report['status'], 'failed')
        self.assertIn('120', '; '.join(report['errors']))

    def test_cli_returns_failure_while_preserving_structure_report(self):
        self.libreoffice(lambda *args, **kwargs: subprocess.CompletedProcess(args[0], 1, '', 'export failed'))
        output = io.StringIO()
        with mock.patch.object(sys, 'argv', ['docx_ops.py', 'verify', str(self.source), '--output-dir', str(self.output)]), contextlib.redirect_stdout(output):
            code = ops.main()
        self.assertEqual(code, 1)
        self.assertEqual(json.loads(output.getvalue())['structure']['errors'], [])

    def test_failed_previews_preserve_existing_page_images(self):
        self.pdf.write_bytes(b'PDF')
        pages = self.output / 'pages'
        pages.mkdir()
        previous = pages / 'page-1.png'
        previous.write_bytes(b'previous preview')

        def run(command, **kwargs):
            if command[0] == 'pdfinfo':
                return subprocess.CompletedProcess(command, 0, b'Pages: 1\n', b'')
            Path(str(command[-1]) + '-1.png').write_bytes(b'partial preview')
            return subprocess.CompletedProcess(command, 1, b'', b'preview failed')

        with mock.patch.object(ops.shutil, 'which', side_effect=lambda name: name), mock.patch.object(ops.subprocess, 'run', side_effect=run):
            with self.assertRaisesRegex(RuntimeError, 'preview failed'):
                ops.pdf_artifacts(self.pdf, pages)
        self.assertEqual(previous.read_bytes(), b'previous preview')

    def test_preview_failure_retains_successful_pdf_export_and_diagnostics(self):
        def run(command, **kwargs):
            if command[0] == 'soffice':
                destination = Path(command[command.index('--outdir') + 1])
                (destination / self.pdf.name).write_bytes(b'exported PDF')
                return subprocess.CompletedProcess(command, 0, '', '')
            return subprocess.CompletedProcess(command, 1, b'', b'preview metadata failed')

        self.libreoffice(run)
        with mock.patch.object(ops.shutil, 'which', side_effect=lambda name: name if name in ('soffice', 'pdfinfo', 'pdftoppm') else None):
            report = ops.verify_document(self.source, output_dir=self.output)
        self.assertEqual(report['render']['status'], 'rendered')
        self.assertEqual(Path(report['render']['pdf']).read_bytes(), b'exported PDF')
        self.assertIn('preview metadata failed', report['render']['pages']['warning'])
        self.assertIn('preview metadata failed', '; '.join(report['limitations']))

    def test_doctor_and_renderer_share_standard_install_discovery(self):
        executable = self.root / 'program files' / 'LibreOffice' / 'program' / 'soffice.exe'
        executable.parent.mkdir(parents=True)
        executable.touch()
        self.assertTrue(hasattr(doctor, 'find_libreoffice'), 'LibreOffice discovery must be shared with rendering')
        with mock.patch.dict(os.environ, {'ProgramFiles': str(executable.parents[2]), 'ProgramFiles(x86)': ''}), mock.patch.object(doctor.shutil, 'which', return_value=None):
            self.assertEqual(doctor.find_libreoffice(), str(executable))
            report = doctor.diagnose()
            self.assertEqual(report['renderers']['libreoffice'], str(executable))

            def export(command, **kwargs):
                self.assertEqual(command[0], str(executable))
                destination = Path(command[command.index('--outdir') + 1])
                (destination / '报告.pdf').write_bytes(b'PDF')
                return subprocess.CompletedProcess(command, 0, '', '')

            with mock.patch.object(ops.subprocess, 'run', side_effect=export):
                rendered = ops._render_document(self.source, self.output, 'libreoffice')
            self.assertEqual(rendered['status'], 'rendered')

    @unittest.skipUnless(POWERSHELL, 'Windows PowerShell is not installed')
    def test_powershell_failed_com_creation_preserves_previous_pdf(self):
        self.pdf.write_bytes(b'previous PDF')
        result = self.run_powershell("function New-Object { throw 'COM unavailable' }")
        self.assertNotEqual(result.returncode, 0)
        self.assertTrue(self.pdf.is_file(), 'a failed COM attempt deleted the existing PDF')
        self.assertEqual(self.pdf.read_bytes(), b'previous PDF')

    @unittest.skipUnless(POWERSHELL, 'Windows PowerShell is not installed')
    def test_powershell_json_round_trips_unicode_without_office(self):
        result = self.run_powershell(self.mock_application())
        self.assertEqual(result.returncode, 0, result.stderr.decode('utf-8', errors='replace'))
        report = json.loads(result.stdout.decode('utf-8', errors='replace'))
        self.assertEqual(Path(report['pdf']), self.pdf)
        self.assertIn('文字', report['renderer'])
        self.assertEqual(self.pdf.read_bytes(), b'PDF')

    @unittest.skipUnless(POWERSHELL, 'Windows PowerShell is not installed')
    def test_powershell_failed_export_does_not_publish_partial_pdf(self):
        result = self.run_powershell(self.mock_application(fail_export=True))
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(self.pdf.exists(), 'failed export published a partial PDF')

    @unittest.skipUnless(POWERSHELL, 'Windows PowerShell is not installed')
    def test_powershell_cleanup_failure_keeps_success_report(self):
        setup = self.mock_application() + "\nfunction Remove-Item { throw 'staging file still in use' }"
        result = self.run_powershell(setup)
        self.assertEqual(result.returncode, 0, result.stderr.decode('utf-8', errors='replace'))
        report = json.loads(result.stdout.decode('utf-8'))
        self.assertEqual(report['status'], 'rendered')
        self.assertEqual(Path(report['pdf']).read_bytes(), b'PDF')
        self.assertIn('staging file still in use', report['cleanup_warning'])

    def run_powershell(self, setup):
        def quote(value):
            return "'" + str(value).replace("'", "''") + "'"

        wrapper = self.root / 'mock-render.ps1'
        command = f"& {quote(SCRIPTS / 'render_docx.ps1')} -InputPath {quote(self.source)} -OutputDirectory {quote(self.output)}"
        wrapper.write_text("$ErrorActionPreference = 'Stop'\n[Console]::OutputEncoding = [System.Text.Encoding]::ASCII\n" + setup + '\n' + command, encoding='utf-8-sig')
        return subprocess.run([POWERSHELL, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', str(wrapper)], capture_output=True, timeout=20)

    def mock_application(self, fail_export=False):
        fail = "throw 'export failed'" if fail_export else ''
        return """
$global:mockDocument = [pscustomobject]@{}
$global:mockDocument | Add-Member ScriptMethod ExportAsFixedFormat {
    param($path, $format)
    [System.IO.File]::WriteAllBytes($path, [byte[]](80,68,70))
    FAIL_EXPORT
}
$global:mockDocument | Add-Member ScriptMethod Close { param($save) }
$global:mockDocuments = [pscustomobject]@{}
$global:mockDocuments | Add-Member ScriptMethod Open { return $global:mockDocument }
function New-Object {
    param($ComObject)
    $application = [pscustomobject]@{
        Visible = $false; DisplayAlerts = 0; Name = 'WPS 文字'; Version = '1'
        Documents = $global:mockDocuments
    }
    $application | Add-Member ScriptMethod Quit {}
    return $application
}
""".replace('FAIL_EXPORT', fail)


if __name__ == '__main__':
    unittest.main()
