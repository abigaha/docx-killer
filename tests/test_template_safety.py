import importlib
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from jinja2 import UndefinedError
from lxml import etree

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills' / 'docx-killer' / 'scripts'
sys.path.insert(0, str(SCRIPTS))
ops = importlib.import_module('docx_ops')
STORIES = ('document', 'header1', 'footer1', 'footnotes', 'endnotes', 'comments')


class TemplateSafety(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.source = Path(temp.name) / 'source.docx'
        self.output = Path(temp.name) / 'output.docx'

    def save_template(self, texts, sections=1):
        self.output.unlink(missing_ok=True)
        doc = Document()
        paragraph = doc.add_paragraph(texts.get('document', 'Body'))
        for _ in range(sections - 1):
            doc.add_section()
        for kind in ('header', 'footer'):
            if kind + '1' in texts:
                getattr(doc.sections[0], kind).paragraphs[0].text = texts[kind + '1']
        for kind in ('footnotes', 'endnotes', 'comments'):
            if kind not in texts:
                continue
            singular = kind[:-1]
            root = OxmlElement('w:' + kind)
            if kind != 'comments':
                for note_id, separator in ((-1, 'separator'), (0, 'continuationSeparator')):
                    note = OxmlElement('w:' + singular)
                    note.set(qn('w:id'), str(note_id))
                    note.set(qn('w:type'), separator)
                    p, run = OxmlElement('w:p'), OxmlElement('w:r')
                    run.append(OxmlElement('w:' + separator))
                    p.append(run)
                    note.append(p)
                    root.append(note)
            item = OxmlElement('w:' + singular)
            item.set(qn('w:id'), '1')
            if kind == 'comments':
                item.set(qn('w:author'), 'Test Author')
                for tag in ('commentRangeStart', 'commentRangeEnd'):
                    marker = OxmlElement('w:' + tag)
                    marker.set(qn('w:id'), '1')
                    if tag == 'commentRangeStart':
                        paragraph._p.insert(0, marker)
                    else:
                        paragraph._p.append(marker)
            p, run, text = OxmlElement('w:p'), OxmlElement('w:r'), OxmlElement('w:t')
            if kind != 'comments':
                run.append(OxmlElement('w:' + singular + 'Ref'))
            text.text = texts[kind]
            run.append(text)
            p.append(run)
            item.append(p)
            root.append(item)
            reference = OxmlElement('w:' + singular + 'Reference')
            reference.set(qn('w:id'), '1')
            paragraph.add_run()._r.append(reference)
            part = Part(PackURI('/word/' + kind + '.xml'), getattr(CT, 'WML_' + kind.upper()),
                        etree.tostring(root), doc.part.package)
            doc.part.relate_to(part, getattr(RT, kind.upper()))
        doc.save(self.source)
        self.assertEqual(ops.inspect_package(self.source)['errors'], [])

    def story_text(self, kind):
        with ZipFile(self.output) as archive:
            root = etree.fromstring(archive.read('word/' + kind + '.xml'))
        return ''.join(root.itertext())

    def test_two_section_footnote_data_is_only_rendered_once(self):
        self.save_template({'document': '{{ customer }}', 'footnotes': '{{ customer }}'}, sections=2)
        ops.render_template(self.source, self.output, {'customer': '{{ 6 * 7 }}'})
        self.assertEqual(self.story_text('document'), '{{ 6 * 7 }}')
        self.assertEqual(self.story_text('footnotes'), '{{ 6 * 7 }}')

    def test_endnotes_and_comments_render_variables(self):
        for kind in ('endnotes', 'comments'):
            with self.subTest(story=kind):
                self.save_template({kind: '{{ customer }}'})
                ops.render_template(self.source, self.output, {'customer': 'Northwind'})
                self.assertEqual(self.story_text(kind), 'Northwind')

    def test_missing_values_in_each_story_leave_no_output(self):
        for kind in STORIES:
            with self.subTest(story=kind):
                self.save_template({kind: '{{ customer.name }}'})
                with self.assertRaises(UndefinedError):
                    ops.render_template(self.source, self.output, {'customer': {}})
                self.assertFalse(self.output.exists())

    def test_xml_sensitive_values_are_escaped_in_each_story(self):
        for kind in STORIES:
            with self.subTest(story=kind):
                self.save_template({kind: '{{ customer }}'})
                ops.render_template(self.source, self.output, {'customer': 'A < B & 公司'})
                self.assertEqual(self.story_text(kind), 'A < B & 公司')
                self.assertEqual(ops.inspect_package(self.output)['errors'], [])


if __name__ == '__main__':
    unittest.main()
