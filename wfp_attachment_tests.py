"""Title-group spacing and conservative attachment recognition regression tests."""
import base64
from io import BytesIO
from copy import deepcopy
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt

from wfp_config import DEFAULT_CONFIG
from wfp_core import WordProcessor, normalize_format_options


class AttachmentTestCase(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix='wfp_attachment_test_')
        self.addCleanup(self.directory.cleanup)
        self.folder = Path(self.directory.name)
        self.logs = []

    def format(self, doc, **options):
        source = self.folder / 'source.docx'
        output = self.folder / 'output.docx'
        doc.save(source)
        before = source.read_bytes()
        processor = WordProcessor({**DEFAULT_CONFIG, **options}, self.logs.append)
        try:
            with patch.object(processor, '_com_available', return_value=False):
                processor.format_document(str(source), str(output))
            self.assertEqual(source.read_bytes(), before)
            return Document(output)
        finally:
            processor._cleanup_temp_files()

    @staticmethod
    def paragraph(doc, text, centered=False, size=16):
        para = doc.add_paragraph(text)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT
        para.runs[0].font.size = Pt(size)
        return para

    @staticmethod
    def spacing(para, attribute='after'):
        element = para._p.get_or_add_pPr().find(qn('w:spacing'))
        return element.get(qn('w:' + attribute)) if element is not None else None


class TitleAfterSpacingTests(AttachmentTestCase):
    def grouped_document(self):
        doc = Document()
        for text, centered, size in (
            ('主标题第一行', True, 22), ('主标题第二行', True, 22),
            ('副标题第一行', True, 16), ('副标题第二行', True, 16), ('正文', False, 16),
            ('附件1', False, 16), ('附件标题第一行', True, 22), ('附件标题第二行', True, 22),
            ('附件副标题第一行', True, 16), ('附件副标题第二行', True, 16), ('附件正文', False, 16),
        ):
            self.paragraph(doc, text, centered, size)
        return doc

    def test_defaults_do_not_add_spacing_or_paragraphs(self):
        doc = self.grouped_document()
        result = self.format(doc)
        self.assertEqual([p.text for p in result.paragraphs], [p.text for p in doc.paragraphs])
        self.assertTrue(all(self.spacing(p) == '0' for p in result.paragraphs))
        self.assertFalse(DEFAULT_CONFIG['enable_title_after_spacing'])
        self.assertFalse(DEFAULT_CONFIG['enable_attachment_title_after_spacing'])

    def test_spacing_only_follows_last_subtitle_in_each_group(self):
        result = self.format(self.grouped_document(), enable_title_after_spacing=True,
                             title_after_spacing=1.5, title_after_spacing_unit='lines',
                             enable_attachment_title_after_spacing=True,
                             attachment_title_after_spacing=12, attachment_title_after_spacing_unit='pt')
        self.assertEqual(len(result.paragraphs), 11)
        self.assertEqual(self.spacing(result.paragraphs[3], 'afterLines'), '150')
        self.assertEqual(self.spacing(result.paragraphs[3]), '0')
        self.assertEqual(self.spacing(result.paragraphs[9]), '240')
        self.assertEqual(self.spacing(result.paragraphs[9], 'afterLines'), '0')
        self.assertTrue(all(self.spacing(p) == '0' for i, p in enumerate(result.paragraphs) if i != 9))
        self.assertTrue(all(self.spacing(p, 'afterLines') in (None, '0')
                            for i, p in enumerate(result.paragraphs) if i != 3))

    def test_main_and_attachment_settings_are_independent(self):
        result = self.format(self.grouped_document(), enable_title_after_spacing=True, title_after_spacing=28)
        self.assertEqual(self.spacing(result.paragraphs[3]), '560')
        self.assertEqual(self.spacing(result.paragraphs[9]), '0')
        result = self.format(self.grouped_document(), enable_attachment_title_after_spacing=True,
                             attachment_title_after_spacing=1, attachment_title_after_spacing_unit='lines')
        self.assertEqual(self.spacing(result.paragraphs[3]), '0')
        self.assertEqual(self.spacing(result.paragraphs[9], 'afterLines'), '100')

    def test_last_title_without_subtitle_and_existing_blank_line(self):
        doc = Document()
        self.paragraph(doc, '主标题第一行', True, 22)
        self.paragraph(doc, '主标题第二行', True, 22)
        doc.add_paragraph()
        self.paragraph(doc, '正文')
        result = self.format(doc, enable_title_after_spacing=True, title_after_spacing=10)
        self.assertEqual(len(result.paragraphs), 4)
        self.assertEqual(result.paragraphs[2].text, '')
        self.assertEqual(self.spacing(result.paragraphs[0]), '0')
        self.assertEqual(self.spacing(result.paragraphs[1]), '200')

    def test_zero_overrides_inherited_spacing_and_contextual_suppression(self):
        doc = Document()
        para = self.paragraph(doc, '标题', True, 22)
        para.style.paragraph_format.space_after = Pt(18)
        para.style.element.get_or_add_pPr().get_or_add_spacing().set(qn('w:afterLines'), '100')
        contextual = OxmlElement('w:contextualSpacing')
        para._p.get_or_add_pPr().append(contextual)
        WordProcessor({**DEFAULT_CONFIG, 'enable_title_after_spacing': True})._apply_title_after_spacing(para)
        self.assertEqual(self.spacing(para), '0')
        self.assertEqual(self.spacing(para, 'afterLines'), '0')
        self.assertEqual(contextual.get(qn('w:val')), '0')

    def test_repeated_formatting_keeps_spacing_and_paragraph_count(self):
        first = self.format(self.grouped_document(), enable_title_after_spacing=True,
                            title_after_spacing=1, title_after_spacing_unit='lines')
        second = self.format(first, enable_title_after_spacing=True,
                             title_after_spacing=1, title_after_spacing_unit='lines')
        self.assertEqual([p.text for p in first.paragraphs], [p.text for p in second.paragraphs])
        self.assertEqual(self.spacing(second.paragraphs[3], 'afterLines'), '100')

    def test_invalid_spacing_is_rejected(self):
        for value in (-1, float('nan'), float('inf'), 'bad'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                normalize_format_options({'title_after_spacing': value})


class AttachmentRecognitionTests(AttachmentTestCase):
    def test_supported_markers_and_prose_boundaries(self):
        for text in ('附件', '附件：', '附件1', '附件一', '附件 2：实施方案', '附件一 实施方案'):
            with self.subTest(text=text):
                self.assertIsNotNone(WordProcessor._parse_attachment_line(Document().add_paragraph(text)))
        for text in ('附件1所列资料应报送。', '附件一中的规定', '附件说明', '详见附件1',
                     '附件1 所列资料应报送', '附件1：请按要求提交。'):
            with self.subTest(text=text):
                self.assertIsNone(WordProcessor._parse_attachment_line(Document().add_paragraph(text)))

    def test_inline_and_separate_lists_do_not_page_break_or_become_headings(self):
        for lines in (('附件：1.实施方案 2.统计表',), ('附件：', '1.实施方案', '2.统计表'),
                      ('附件：', '（1）实施方案', '（2）统计表'),
                      ('附件1：实施方案', '附件2：统计表')):
            with self.subTest(lines=lines):
                doc = Document()
                self.paragraph(doc, '正文内容。')
                for text in lines:
                    self.paragraph(doc, text)
                result = self.format(doc, enable_attachment_title_after_spacing=True, attachment_title_after_spacing=30)
                self.assertEqual([p.text for p in result.paragraphs], [p.text for p in doc.paragraphs])
                for para in result.paragraphs[1:]:
                    self.assertFalse(para.paragraph_format.page_break_before)
                    self.assertIsNone(para._p.find('./' + qn('w:pPr') + '/' + qn('w:outlineLvl')))
                    self.assertEqual(self.spacing(para), '0')

    def test_automatic_numbered_list_entries_are_classified_after_preprocessing(self):
        doc = Document()
        self.paragraph(doc, '正文内容。')
        self.paragraph(doc, '附件：')
        doc.add_paragraph('实施方案', style='List Number')
        doc.add_paragraph('统计表', style='List Number')
        result = self.format(doc)
        self.assertTrue(result.paragraphs[2].text.startswith('1.'))
        self.assertFalse(result.paragraphs[2].paragraph_format.page_break_before)
        self.assertIsNone(result.paragraphs[2]._p.find('./' + qn('w:pPr') + '/' + qn('w:outlineLvl')))

    def test_ambiguous_single_numbered_entry_stays_a_list_and_logs(self):
        doc = Document()
        self.paragraph(doc, '正文内容。')
        self.paragraph(doc, '附件1：实施方案')
        self.paragraph(doc, '其他说明。')
        result = self.format(doc)
        self.assertEqual(len(result.paragraphs), 3)
        self.assertFalse(result.paragraphs[1].paragraph_format.page_break_before)
        self.assertTrue(any('缺少附件正文线索' in text for text in self.logs))

    def test_front_of_document_adjacent_labels_still_form_a_list(self):
        doc = Document()
        self.paragraph(doc, '附件1：实施方案')
        self.paragraph(doc, '附件2：统计表')
        result = self.format(doc)
        self.assertEqual([p.text for p in result.paragraphs], ['附件1：实施方案', '附件2：统计表'])
        self.assertFalse(any(p.paragraph_format.page_break_before for p in result.paragraphs))

    def test_list_enumeration_does_not_swallow_next_body_heading(self):
        doc = Document()
        for text in ('正文内容。', '附件：', '1.实施方案', '2.统计表', '一、后续正文事项'):
            self.paragraph(doc, text)
        result = self.format(doc)
        self.assertEqual(result.paragraphs[-1].runs[0].font.name, DEFAULT_CONFIG['h1_font'])
        self.assertEqual(result.paragraphs[-1]._p.find('./' + qn('w:pPr') + '/' + qn('w:outlineLvl')).get(qn('w:val')), '0')

    def test_separate_body_labels_are_followed_by_real_titles(self):
        for marker in ('附件', '附件：', '附件1', '附件一'):
            with self.subTest(marker=marker):
                doc = Document()
                self.paragraph(doc, marker)
                self.paragraph(doc, '实施方案', True, 22)
                self.paragraph(doc, '附件正文。')
                result = self.format(doc, enable_attachment_title_after_spacing=True,
                                     attachment_title_after_spacing=8)
                self.assertTrue(result.paragraphs[0].paragraph_format.page_break_before)
                self.assertEqual(result.paragraphs[1].alignment, WD_ALIGN_PARAGRAPH.CENTER)
                self.assertEqual(self.spacing(result.paragraphs[1]), '160')
                self.assertEqual(self.spacing(result.paragraphs[2]), '0')

    def test_confirmed_inline_body_splits_and_preserves_hyperlink_image_and_bookmark(self):
        doc = Document()
        para = doc.add_paragraph()
        para.paragraph_format.page_break_before = True
        bookmark = OxmlElement('w:bookmarkStart')
        bookmark.set(qn('w:id'), '18')
        bookmark.set(qn('w:name'), 'attachment_name')
        para._p.append(bookmark)
        link = OxmlElement('w:hyperlink')
        relationship = para.part.relate_to('https://example.com/attachment', RT.HYPERLINK, is_external=True)
        link.set(qn('r:id'), relationship)
        run = OxmlElement('w:r')
        text = OxmlElement('w:t')
        text.text = '附件1：实施方案'
        run.append(text)
        link.append(run)
        para._p.append(link)
        pixel = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=')
        para.add_run().add_picture(BytesIO(pixel))
        end = OxmlElement('w:bookmarkEnd')
        end.set(qn('w:id'), '18')
        para._p.append(end)
        self.paragraph(doc, '一、工作安排')
        result = self.format(doc, enable_attachment_title_after_spacing=True,
                             attachment_title_after_spacing=1, attachment_title_after_spacing_unit='lines')
        texts = [WordProcessor._paragraph_text(p) for p in result.paragraphs]
        self.assertEqual(texts, ['附件1：', '实施方案', '一、工作安排'])
        self.assertEqual(len(list(result.element.iter(qn('w:drawing')))), 1)
        self.assertEqual(len(list(result.element.iter(qn('w:bookmarkStart')))), 1)
        self.assertEqual(len(list(result.element.iter(qn('w:bookmarkEnd')))), 1)
        links = list(result.element.iter(qn('w:hyperlink')))
        self.assertEqual([node.get(qn('r:id')) for node in links], [relationship, relationship])
        self.assertEqual(result.part.rels[relationship].target_ref, 'https://example.com/attachment')
        self.assertEqual(self.spacing(result.paragraphs[1], 'afterLines'), '100')
        with ZipFile(self.folder / 'source.docx') as before, ZipFile(self.folder / 'output.docx') as after:
            for name in before.namelist():
                if name.startswith('word/media/'):
                    self.assertEqual(before.read(name), after.read(name))
        repeated = self.format(result, enable_attachment_title_after_spacing=True,
                                attachment_title_after_spacing=1, attachment_title_after_spacing_unit='lines')
        self.assertEqual([WordProcessor._paragraph_text(p) for p in repeated.paragraphs], texts)

    def test_content_control_crossing_boundary_keeps_its_identity_and_same_paragraph(self):
        doc = Document()
        para = doc.add_paragraph()
        para.paragraph_format.page_break_before = True
        control = OxmlElement('w:sdt')
        properties = OxmlElement('w:sdtPr')
        identity = OxmlElement('w:id')
        identity.set(qn('w:val'), '12345')
        properties.append(identity)
        control.append(properties)
        content = OxmlElement('w:sdtContent')
        run = OxmlElement('w:r')
        text = OxmlElement('w:t')
        text.text = '附件1：实施方案'
        run.append(text)
        content.append(run)
        control.append(content)
        para._p.append(control)
        result = self.format(doc)
        self.assertEqual(len(result.paragraphs), 1)
        self.assertEqual(WordProcessor._paragraph_text(result.paragraphs[0]), '附件1：实施方案')
        self.assertEqual(len(list(result.element.iter(qn('w:sdt')))), 1)
        self.assertEqual(next(result.element.iter(qn('w:id'))).get(qn('w:val')), '12345')
        self.assertTrue(any('保留同段结构' in text for text in self.logs))

    def test_body_without_title_does_not_take_later_centered_text_as_title(self):
        doc = Document()
        self.paragraph(doc, '附件1')
        self.paragraph(doc, '一、工作安排')
        self.paragraph(doc, '正文内容。')
        self.paragraph(doc, '独立居中说明', True, 16)
        result = self.format(doc, enable_attachment_title_after_spacing=True, attachment_title_after_spacing=20)
        self.assertEqual(self.spacing(result.paragraphs[1]), '0')
        self.assertNotEqual(result.paragraphs[1].runs[0].font.size.pt, DEFAULT_CONFIG['title_size'])
        self.assertEqual(self.spacing(result.paragraphs[3]), '0')

    def test_split_moves_section_properties_once_and_keeps_property_order(self):
        doc = Document()
        para = self.paragraph(doc, '附件1：实施方案')
        section = deepcopy(doc.element.body.sectPr)
        para._p.get_or_add_pPr().append(section)
        processor = WordProcessor({**DEFAULT_CONFIG, 'enable_attachment_title_after_spacing': True,
                                   'attachment_title_after_spacing': 10})
        title_element = processor._split_attachment_paragraph(para, len('附件1：'))
        self.assertIsNotNone(title_element)
        self.assertIsNone(para._p.pPr.find(qn('w:sectPr')))
        self.assertIs(title_element.pPr.find(qn('w:sectPr')), section)
        self.assertEqual(len(list(doc.element.iter(qn('w:sectPr')))), 2)
        from docx.text.paragraph import Paragraph
        title = Paragraph(title_element, para._parent)
        processor._format_title_paragraph(title, False, attachment=True)
        processor._apply_title_after_spacing(title, attachment=True)
        tags = [child.tag for child in title_element.pPr]
        self.assertLess(tags.index(qn('w:contextualSpacing')), tags.index(qn('w:sectPr')))

    def test_text_and_markdown_inputs_support_lists_and_attachment_bodies(self):
        for extension in ('.txt', '.md'):
            source = self.folder / ('text' + extension)
            output = self.folder / 'text-result.docx'
            source.write_text('主标题\n正文内容。\n附件：\n1.实施方案\n2.统计表', encoding='utf-8')
            processor = WordProcessor(DEFAULT_CONFIG)
            try:
                processor.format_document(str(source), str(output))
                result = Document(output)
                self.assertEqual(len(result.paragraphs), 5)
                self.assertFalse(any(p.paragraph_format.page_break_before for p in result.paragraphs))
                source.write_text('主标题\n正文内容。\n附件1：实施方案\n一、工作安排', encoding='utf-8')
                processor.format_document(str(source), str(output))
                result = Document(output)
                self.assertEqual([p.text for p in result.paragraphs], ['主标题', '正文内容。', '附件1：', '实施方案', '一、工作安排'])
                self.assertTrue(result.paragraphs[2].paragraph_format.page_break_before)
            finally:
                processor._cleanup_temp_files()

    def test_disabled_attachment_formatting_does_not_split_or_add_spacing(self):
        doc = Document()
        self.paragraph(doc, '正文内容。')
        para = self.paragraph(doc, '附件1：实施方案')
        para.paragraph_format.page_break_before = True
        self.paragraph(doc, '一、工作安排')
        result = self.format(doc, enable_attachment_formatting=False,
                             enable_attachment_title_after_spacing=True, attachment_title_after_spacing=30)
        self.assertEqual(len(result.paragraphs), 3)
        self.assertEqual(result.paragraphs[1].text, '附件1：实施方案')
        self.assertTrue(all(self.spacing(p) == '0' for p in result.paragraphs))

    def test_list_preserves_existing_page_break_instead_of_adding_one(self):
        doc = Document()
        self.paragraph(doc, '正文内容。')
        para = self.paragraph(doc, '附件：1.实施方案')
        para.paragraph_format.page_break_before = True
        result = self.format(doc)
        self.assertTrue(result.paragraphs[1].paragraph_format.page_break_before)
        self.assertEqual(len(result.paragraphs), 2)

    def test_body_does_not_add_a_second_break_after_an_existing_page_break(self):
        from docx.enum.text import WD_BREAK
        doc = Document()
        self.paragraph(doc, '正文内容。').add_run().add_break(WD_BREAK.PAGE)
        self.paragraph(doc, '附件1：实施方案')
        self.paragraph(doc, '一、工作安排')
        result = self.format(doc)
        self.assertEqual([p.text for p in result.paragraphs[1:]], ['附件1：', '实施方案', '一、工作安排'])
        self.assertFalse(result.paragraphs[1].paragraph_format.page_break_before)
        self.assertEqual(len(list(result.element.body.iter(qn('w:br')))), 1)

    def test_page_break_in_middle_of_previous_paragraph_is_not_a_page_start(self):
        from docx.enum.text import WD_BREAK
        doc = Document()
        para = self.paragraph(doc, '前页正文')
        para.add_run().add_break(WD_BREAK.PAGE)
        para.add_run('后一页正文')
        self.paragraph(doc, '附件1：实施方案')
        processor = WordProcessor(DEFAULT_CONFIG)
        roles = processor._classify_attachments(list(processor._iter_block_items(doc)), False)
        self.assertEqual(roles[doc.paragraphs[1]._p]['kind'], 'list')


if __name__ == '__main__':
    unittest.main()
