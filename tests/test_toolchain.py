import importlib
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills' / 'docx-killer' / 'scripts'
sys.path.insert(0, str(SCRIPTS))
try:
    doctor = importlib.import_module('doctor')
except ModuleNotFoundError:
    doctor = None


class InstalledToolchain(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / 'example.docx'
        doc = Document()
        doc.add_heading('Example report', 1)
        p = doc.add_paragraph('Customer: ')
        p.add_run('Acme & \u516c\u53f8').bold = True
        doc.add_table(rows=1, cols=1).cell(0, 0).text = 'Table value'
        doc.save(self.source)

    def run_command(self, *args, success=True):
        result = subprocess.run(args, capture_output=True, text=True, encoding='utf-8', timeout=45)
        if success:
            self.assertEqual(result.returncode, 0, result.stderr)
        else:
            self.assertNotEqual(result.returncode, 0)
        return result

    def test_doctor_locates_all_four_installed_tools(self):
        self.assertIsNotNone(doctor, 'Windows tool discovery is not implemented yet')
        report = doctor.diagnose()
        for name in ('python-docx', 'docxtpl', 'pandoc', 'mammoth'):
            self.assertTrue(report['tools'][name]['available'], report['tools'][name])
        self.assertTrue(Path(report['tools']['pandoc']['path']).is_file())

    def test_pandoc_extracts_and_rebuilds_semantic_content(self):
        self.assertIsNotNone(doctor, 'Windows tool discovery is not implemented yet')
        executable = doctor.find_pandoc()
        self.assertIsNotNone(executable, 'Pandoc required for the four-tool integration check')
        markdown = self.root / 'report.md'
        output = self.root / 'rebuilt.docx'
        self.run_command(executable, str(self.source), '-f', 'docx', '-t', 'markdown', '-o', str(markdown))
        text = markdown.read_text(encoding='utf-8')
        self.assertIn('Example report', text)
        self.assertIn('Acme', text)
        self.assertIn('Table value', text)
        self.run_command(executable, str(markdown), '-o', str(output))
        self.assertIn('Example report', '\n'.join(p.text for p in Document(output).paragraphs))

    def test_mammoth_html_and_text_with_global_npm_install(self):
        script = SCRIPTS / 'mammoth_extract.cjs'
        self.assertTrue(script.is_file(), 'Mammoth wrapper is not implemented yet')
        node = shutil.which('node')
        self.assertIsNotNone(node)
        output = self.root / 'report.html'
        result = self.run_command(node, str(script), str(self.source), str(output))
        report = json.loads(result.stdout)
        self.assertIn('messages', report)
        html = output.read_text(encoding='utf-8')
        self.assertIn('<meta charset="utf-8">', html)
        self.assertIn('<h1>Example report</h1>', html)
        self.assertIn('<strong>Acme &amp; \u516c\u53f8</strong>', html)
        plain = self.root / 'report.txt'
        self.run_command(node, str(script), str(self.source), str(plain), '--text')
        self.assertIn('Acme & \u516c\u53f8', plain.read_text(encoding='utf-8'))
        before = output.read_bytes()
        self.run_command(node, str(script), str(self.source), str(output), success=False)
        self.assertEqual(output.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
