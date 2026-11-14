# -*- coding: utf-8 -*-
"""Focused regressions for title blank lines and paragraph-spacing selection."""
import contextlib
import json
from io import StringIO, BytesIO
from unittest.mock import patch
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import wfp_attachment_tests
from wfp_attachment_tests import AttachmentTestCase
from wfp_config import DEFAULT_CONFIG
from wfp_core import WordProcessor, normalize_format_options


class TitleAfterBlankLineTests(AttachmentTestCase):
    def grouped(self):
        return wfp_attachment_tests.TitleAfterSpacingTests.grouped_document(self)

    @staticmethod
    def blank_after(document, text):
        paragraphs = document.paragraphs
        index = next(i for i, p in enumerate(paragraphs) if p.text == text)
        return paragraphs[index + 1].text == ''

    def test_disabled_defaults_and_legacy_spacing_are_preserved(self):
        self.assertFalse(DEFAULT_CONFIG['enable_title_after_blank_line'])
        self.assertFalse(DEFAULT_CONFIG['enable_attachment_title_after_blank_line'])
        doc = self.grouped()
        result = self.format(doc)
        self.assertEqual([p.text for p in result.paragraphs], [p.text for p in doc.paragraphs])
        result = self.format(doc, enable_title_after_spacing=True, title_after_spacing=15)
        self.assertEqual(len(result.paragraphs), len(doc.paragraphs))
        self.assertEqual(self.spacing(result.paragraphs[3]), '300')

    def test_one_blank_after_last_subtitle_in_main_and_attachment(self):
        result = self.format(self.grouped(), enable_title_after_blank_line=True,
                             enable_attachment_title_after_blank_line=True)
        self.assertEqual(len(result.paragraphs), 13)
        for text in ('副标题第二行', '附件副标题第二行'):
            self.assertTrue(self.blank_after(result, text))
        self.assertFalse(self.blank_after(result, '主标题第二行'))
        self.assertFalse(self.blank_after(result, '附件标题第二行'))
        blanks = [p for p in result.paragraphs if p.text == '']
        self.assertTrue(all(p.paragraph_format.line_spacing.pt == 28 for p in blanks))
        self.assertTrue(all(p._p.pPr.find(qn('w:rPr')).find(qn('w:sz')).get(qn('w:val')) == '32' for p in blanks))

    def test_existing_one_or_multiple_whitespace_paragraphs_are_kept(self):
        for existing in ([''], ['', ''], [' \t\u3000']):
            doc = Document()
            self.paragraph(doc, '文章标题', True, 22)
            for text in existing:
                doc.add_paragraph(text)
            self.paragraph(doc, '正文')
            result = self.format(doc, enable_title_after_blank_line=True)
            self.assertEqual([p.text for p in result.paragraphs], [p.text for p in doc.paragraphs])

    def test_repeated_formatting_and_reused_processor_do_not_accumulate(self):
        source = self.folder / 'source.docx'
        self.grouped().save(source)
        processor = WordProcessor({**DEFAULT_CONFIG, 'enable_title_after_blank_line': True,
                                   'enable_attachment_title_after_blank_line': True})
        try:
            with patch.object(processor, '_com_available', return_value=False):
                first = self.folder / 'first.docx'
                second = self.folder / 'second.docx'
                processor.format_document(str(source), str(first))
                processor.format_document(str(first), str(second))
                self.assertEqual([p.text for p in Document(first).paragraphs],
                                 [p.text for p in Document(second).paragraphs])
                self.assertEqual(len(Document(second).paragraphs), 13)
                other = Document()
                self.paragraph(other, '另一标题', True, 22)
                self.paragraph(other, '另一正文')
                other.save(source)
                processor.format_document(str(source), str(first))
                self.assertEqual([p.text for p in Document(first).paragraphs], ['另一标题', '', '另一正文'])
        finally:
            processor._cleanup_temp_files()

    def test_main_attachment_and_attachment_master_switch_are_independent(self):
        result = self.format(self.grouped(), enable_title_after_blank_line=True)
        self.assertTrue(self.blank_after(result, '副标题第二行'))
        self.assertFalse(self.blank_after(result, '附件副标题第二行'))
        result = self.format(self.grouped(), enable_attachment_title_after_blank_line=True)
        self.assertFalse(self.blank_after(result, '副标题第二行'))
        self.assertTrue(self.blank_after(result, '附件副标题第二行'))
        result = self.format(self.grouped(), enable_attachment_formatting=False,
                             enable_attachment_title_after_blank_line=True)
        self.assertEqual(len(result.paragraphs), 11)

    def test_blank_mode_does_not_stack_paragraph_spacing(self):
        doc = Document()
        title = self.paragraph(doc, '文章标题', True, 22)
        title.style.element.get_or_add_pPr().get_or_add_spacing().set(qn('w:afterLines'), '100')
        self.paragraph(doc, '正文')
        result = self.format(doc, enable_title_after_blank_line=True,
                             enable_title_after_spacing=True, title_after_spacing=30)
        self.assertEqual(len(result.paragraphs), 3)
        self.assertEqual(self.spacing(result.paragraphs[0]), '0')
        self.assertEqual(self.spacing(result.paragraphs[0], 'afterLines'), '0')

    def test_tables_and_images_are_not_mistaken_for_empty_paragraphs(self):
        from wfp_tests import InlineFormattingProtectionTests
        doc = Document()
        self.paragraph(doc, '文章标题', True, 22)
        doc.add_paragraph().add_run().add_picture(BytesIO(InlineFormattingProtectionTests.PIXEL))
        doc.add_table(1, 1).cell(0, 0).text = '表格内容'
        result = self.format(doc, enable_title_after_blank_line=True)
        self.assertEqual(len(result.paragraphs), 3)
        self.assertEqual(len(result.element.xpath('.//w:drawing')), 1)
        self.assertEqual(len(result.tables), 1)
        self.assertEqual(result.tables[0].cell(0, 0).text, '表格内容')
        with ZipFile(self.folder / 'source.docx') as before, ZipFile(self.folder / 'output.docx') as after:
            names = [name for name in before.namelist() if name.startswith('word/media/')]
            self.assertTrue(all(before.read(name) == after.read(name) for name in names))

    def test_trailing_break_and_section_boundary_are_not_modified(self):
        for section in (False, True):
            doc = Document()
            title = self.paragraph(doc, '文章标题', True, 22)
            if section:
                title._p.get_or_add_pPr().append(OxmlElement('w:sectPr'))
            else:
                title.runs[0].add_break()
            self.paragraph(doc, '正文')
            result = self.format(doc, enable_title_after_blank_line=True)
            self.assertEqual(len(result.paragraphs), 2)
            self.assertEqual(len(result.sections), len(doc.sections))

    def test_cli_config_roundtrip_and_blank_mode_precedence(self):
        import wfp_cli
        config_path = self.folder / 'config.json'
        with contextlib.redirect_stdout(StringIO()):
            code = wfp_cli.main(['save-config', '-o', str(config_path), '--set',
                                 'enable_title_after_blank_line=true', '--set',
                                 'enable_title_after_spacing=true'])
        self.assertEqual(code, 0)
        saved = json.loads(config_path.read_text(encoding='utf-8'))
        self.assertTrue(saved['enable_title_after_blank_line'])
        self.assertFalse(saved['enable_title_after_spacing'])
        self.assertEqual(normalize_format_options(saved), saved)
        with self.assertRaises(ValueError):
            normalize_format_options({'enable_title_after_blank_line': 'true'})
