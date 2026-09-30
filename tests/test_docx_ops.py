import base64
import hashlib
import importlib
import sys
import tempfile
import unittest
from unittest import mock
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills' / 'docx-killer' / 'scripts'
sys.path.insert(0, str(SCRIPTS))
try:
    ops = importlib.import_module('docx_ops')
except ModuleNotFoundError:
    ops = None


def payloads(path):
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def alter_package(path, name, data):
    entries = payloads(path)
    entries[name] = data
    with ZipFile(path, 'w') as archive:
        for key, value in entries.items():
            archive.writestr(key, value)


class DocumentOperations(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(ops, 'DOCX helper is not implemented yet')
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / 'source.docx'
        self.output = self.root / 'output.docx'

    def save_text(self, text='Acme Corporation'):
        doc = Document()
        doc.add_paragraph(text)
        doc.save(self.source)
        return doc

    def add_test_picture(self, doc):
        image = self.root / 'pixel.png'
        image.write_bytes(base64.b64decode(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
        ))
        doc.add_picture(str(image))

    def test_cross_run_replacement_preserves_unrelated_parts_and_formatting(self):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run('Before ').italic = True
        p.add_run('Acme ').bold = True
        p.add_run('Corporation').italic = True
        p.add_run(' after').underline = True
        doc.sections[0].header.paragraphs[0].text = 'Acme Corporation'
        doc.add_table(rows=1, cols=1).cell(0, 0).text = 'Acme Corporation'
        doc.save(self.source)
        original_hash = hashlib.sha256(self.source.read_bytes()).hexdigest()
        report = ops.replace_text(self.source, self.output, 'Acme Corporation', 'Northwind & Co', 2)
        self.assertEqual(report['replacements'], 2)
        out = Document(self.output)
        self.assertEqual(out.paragraphs[0].text, 'Before Northwind & Co after')
        self.assertTrue(out.paragraphs[0].runs[1].bold)
        self.assertTrue(out.paragraphs[0].runs[-1].underline)
        self.assertEqual(out.sections[0].header.paragraphs[0].text, 'Acme Corporation')
        self.assertEqual(out.tables[0].cell(0, 0).text, 'Northwind & Co')
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), original_hash)
        before, after = payloads(self.source), payloads(self.output)
        self.assertEqual(set(before), set(after))
        self.assertEqual([n for n in before if before[n] != after[n]], ['word/document.xml'])

    def test_multiple_matches_whitespace_and_chinese_survive(self):
        self.save_text('Acme Corporation / Acme Corporation / \u5b8c\u6210')
        ops.replace_text(self.source, self.output, 'Acme Corporation', ' \u65b0 & <\u540d> ', 2)
        self.assertEqual(Document(self.output).paragraphs[0].text, ' \u65b0 & <\u540d>  /  \u65b0 & <\u540d>  / \u5b8c\u6210')

    def test_replacement_is_not_recursively_replaced(self):
        self.save_text('aa aa')
        ops.replace_text(self.source, self.output, 'aa', 'aaa', 2)
        self.assertEqual(Document(self.output).paragraphs[0].text, 'aaa aaa')

    def test_explicit_header_scope(self):
        doc = self.save_text()
        doc.sections[0].header.paragraphs[0].text = 'Acme Corporation'
        doc.save(self.source)
        ops.replace_text(self.source, self.output, 'Acme Corporation', 'New', 1, ['word/header*.xml'])
        self.assertEqual(Document(self.output).paragraphs[0].text, 'Acme Corporation')
        self.assertEqual(Document(self.output).sections[0].header.paragraphs[0].text, 'New')

    def test_count_mismatch_leaves_no_output(self):
        self.save_text()
        with self.assertRaisesRegex(ValueError, 'count'):
            ops.replace_text(self.source, self.output, 'Acme Corporation', 'New', 2)
        self.assertFalse(self.output.exists())

    def test_existing_output_and_source_are_never_overwritten(self):
        self.save_text()
        self.output.write_bytes(b'keep me')
        with self.assertRaises((ValueError, FileExistsError)):
            ops.replace_text(self.source, self.output, 'Acme', 'New', 1)
        self.assertEqual(self.output.read_bytes(), b'keep me')
        with self.assertRaises(ValueError):
            ops.replace_text(self.source, self.source, 'Acme', 'New', 1)
        self.assertEqual(Document(self.source).paragraphs[0].text, 'Acme Corporation')

    def test_bookmark_boundary_is_refused_without_partial_output(self):
        doc = Document()
        doc.add_paragraph('Acme Corporation')
        p = doc.add_paragraph('Acme ')
        bookmark = OxmlElement('w:bookmarkStart')
        bookmark.set(qn('w:id'), '0')
        bookmark.set(qn('w:name'), 'customer')
        p._p.append(bookmark)
        p.add_run('Corporation')
        doc.save(self.source)
        with self.assertRaisesRegex(ValueError, 'protected'):
            ops.replace_text(self.source, self.output, 'Acme Corporation', 'New', 2)
        self.assertFalse(self.output.exists())

    def test_fields_across_paragraphs_are_refused(self):
        doc = Document()
        p = doc.add_paragraph()
        field = OxmlElement('w:fldChar')
        field.set(qn('w:fldCharType'), 'begin')
        p.add_run()._r.append(field)
        doc.add_paragraph('Acme Corporation')
        p = doc.add_paragraph()
        field = OxmlElement('w:fldChar')
        field.set(qn('w:fldCharType'), 'end')
        p.add_run()._r.append(field)
        doc.save(self.source)
        with self.assertRaisesRegex(ValueError, 'protected'):
            ops.replace_text(self.source, self.output, 'Acme Corporation', 'New', 1)
        self.assertFalse(self.output.exists())

    def test_hyperlinks_and_revisions_are_refused(self):
        for tag in ['w:hyperlink', 'w:ins', 'w:del', 'w:sdt']:
            with self.subTest(tag=tag):
                doc = Document()
                p = doc.add_paragraph()
                wrapper = OxmlElement(tag)
                run = OxmlElement('w:r')
                text = OxmlElement('w:t')
                text.text = 'Acme Corporation'
                run.append(text)
                wrapper.append(run)
                p._p.append(wrapper)
                doc.save(self.source)
                with self.assertRaisesRegex(ValueError, 'protected'):
                    ops.replace_text(self.source, self.output, 'Acme Corporation', 'New', 1)
                self.assertFalse(self.output.exists())

    def test_invalid_relationship_target_is_detected(self):
        self.save_text()
        rels = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rBad" Type="image" Target="media/missing.png"/></Relationships>'
        alter_package(self.source, 'word/_rels/document.xml.rels', rels)
        report = ops.inspect_package(self.source)
        self.assertTrue(any('missing.png' in error for error in report['errors']))
        with self.assertRaises(ValueError):
            ops.replace_text(self.source, self.output, 'Acme', 'New', 1)

    def test_inspection_and_comparison_report_actual_changes(self):
        self.save_text()
        ops.replace_text(self.source, self.output, 'Acme Corporation', 'New', 1)
        report = ops.inspect_package(self.source, include_text=True)
        self.assertEqual(report['errors'], [])
        self.assertIn('Acme Corporation', report['stories']['word/document.xml']['paragraphs'][0]['text'])
        diff = ops.compare_packages(self.source, self.output)
        self.assertEqual(diff['changed'], ['word/document.xml'])
        self.assertEqual(diff['added'], [])
        self.assertEqual(diff['removed'], [])

    def test_inspection_filters_large_reports_without_losing_counts(self):
        doc = Document()
        for index in range(6):
            doc.add_paragraph(f'Paragraph {index}')
        doc.add_table(rows=2, cols=2)
        self.add_test_picture(doc)
        doc.save(self.source)
        report = ops.inspect_document(
            self.source,
            include_text=True,
            summary=True,
            stories=['word/document.xml'],
            paragraph_range=(2, 4),
            include_tables=True,
            include_media=True,
        )
        self.assertNotIn('parts', report)
        story = report['stories']['word/document.xml']
        self.assertEqual(story['paragraph_count'], 11)
        self.assertEqual([item['index'] for item in story['paragraphs']], [2, 3])
        self.assertEqual(len(story['tables']), 1)
        self.assertEqual(report['media']['count'], 1)
        self.assertEqual(report['media']['items'][0]['part'], 'word/media/image1.png')

    def test_advanced_structure_warning_names_features_and_stories(self):
        doc = Document()
        doc.add_paragraph('Before')
        self.add_test_picture(doc)
        doc.save(self.source)
        report = ops.inspect_package(self.source)
        warning = next(item for item in report['warnings'] if item.startswith('Advanced structures present:'))
        self.assertIn('drawings=1', warning)
        self.assertIn('word/document.xml', warning)

    def test_verify_document_reports_reopen_statistics_without_rendering(self):
        doc = Document()
        doc.add_paragraph('One')
        doc.add_table(rows=1, cols=1).cell(0, 0).text = 'Two'
        doc.save(self.source)
        report = ops.verify_document(self.source, render='never')
        self.assertEqual(report['status'], 'passed')
        self.assertEqual(report['structure']['errors'], [])
        self.assertEqual(report['reopen']['paragraphs'], 1)
        self.assertEqual(report['reopen']['tables'], 1)
        self.assertEqual(report['render']['status'], 'skipped')
        self.assertIn('Layout was not rendered', report['limitations'])

    @unittest.skipUnless(__import__('shutil').which('pdfinfo') and __import__('shutil').which('pdftoppm'), 'Poppler is not installed')
    def test_pdf_artifacts_report_page_count_and_create_previews(self):
        from PIL import Image
        pdf = self.root / 'rendered.pdf'
        Image.new('RGB', (100, 100), 'white').save(pdf, 'PDF')
        pages = self.root / 'pages'
        pages.mkdir()
        (pages / 'page-99.png').write_bytes(b'stale')
        report = ops.pdf_artifacts(pdf, pages)
        self.assertEqual(report['page_count'], 1)
        self.assertEqual(len(report['previews']), 1)
        self.assertIn('creator', report)
        self.assertIn('producer', report)
        self.assertTrue(Path(report['previews'][0]).is_file())

    def test_subprocess_output_uses_windows_local_code_page(self):
        with mock.patch.object(ops.locale, 'getpreferredencoding', return_value='cp936'):
            self.assertEqual(ops.decode_output('WPS 文字'.encode('cp936')), 'WPS 文字')

    def test_explicit_unavailable_renderer_fails(self):
        self.save_text('One')
        with mock.patch.object(ops.shutil, 'which', return_value=None):
            with self.assertRaisesRegex(RuntimeError, 'renderer is unavailable'):
                ops._render_document(self.source, self.root / 'render', 'word')

    def test_auto_renderer_prefers_registered_wps_interface(self):
        with mock.patch.object(ops, '_registered_com_progid', side_effect=lambda progid: progid in {'KWPS.Application', 'Word.Application'}):
            self.assertEqual(ops._office_progids('auto'), ['KWPS.Application', 'Word.Application'])
            self.assertEqual(ops._office_progids('word'), ['Word.Application'])

    def test_cli_verify_and_filtered_inspect_emit_json(self):
        import json
        import subprocess
        self.save_text('One')
        cli = SCRIPTS / 'docx_ops.py'
        inspect = subprocess.run([sys.executable, str(cli), 'inspect', str(self.source), '--summary', '--text', '--paragraph-range', '0:1'], capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(inspect.returncode, 0, inspect.stderr)
        self.assertNotIn('"parts"', inspect.stdout)
        verify = subprocess.run([sys.executable, str(cli), 'verify', str(self.source), '--render', 'never'], capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(verify.returncode, 0, verify.stderr)
        self.assertEqual(json.loads(verify.stdout)['status'], 'passed')

    def test_cli_verify_returns_nonzero_for_failed_structure(self):
        import subprocess
        self.save_text('One')
        rels = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rBad" Type="image" Target="media/missing.png"/></Relationships>'
        alter_package(self.source, 'word/_rels/document.xml.rels', rels)
        cli = SCRIPTS / 'docx_ops.py'
        verify = subprocess.run([sys.executable, str(cli), 'verify', str(self.source), '--render', 'never'], capture_output=True, text=True, encoding='utf-8')
        self.assertNotEqual(verify.returncode, 0)

    def test_signed_package_refused(self):
        self.save_text()
        alter_package(self.source, '_xmlsignatures/sig1.xml', b'<signature/>')
        with self.assertRaisesRegex(ValueError, 'signed'):
            ops.replace_text(self.source, self.output, 'Acme', 'New', 1)
        self.assertFalse(self.output.exists())

    def test_template_escaping_and_repeated_rows(self):
        doc = Document()
        doc.add_paragraph('{{ customer.name }}')
        table = doc.add_table(rows=3, cols=1)
        table.cell(0, 0).text = '{%tr for item in items %}'
        table.cell(1, 0).text = '{{ item.label }}'
        table.cell(2, 0).text = '{%tr endfor %}'
        doc.save(self.source)
        ops.render_template(self.source, self.output, {'customer': {'name': 'A < B & \u516c\u53f8'}, 'items': [{'label': 'One'}, {'label': 'Two'}]})
        result = Document(self.output)
        self.assertEqual(result.paragraphs[0].text, 'A < B & \u516c\u53f8')
        self.assertEqual([row.cells[0].text for row in result.tables[0].rows], ['One', 'Two'])
        self.assertEqual(ops.inspect_package(self.output)['errors'], [])

    def test_template_missing_nested_value_fails_without_output(self):
        self.save_text('{{ customer.name }}')
        with self.assertRaises(Exception):
            ops.render_template(self.source, self.output, {'customer': {}})
        self.assertFalse(self.output.exists())


if __name__ == '__main__':
    unittest.main()
