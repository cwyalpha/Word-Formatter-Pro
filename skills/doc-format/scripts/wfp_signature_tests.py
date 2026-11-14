"""Signature placement, conservative detection and OOXML preservation regressions."""
import base64
import os
from datetime import date
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Cm, Pt

from wfp_config import DEFAULT_CONFIG
from wfp_core import WordProcessor, normalize_format_options
from wfp_signature import (FontMetrics, SignatureFormatter, _find_font, _font_advances,
                           _font_candidates, _font_faces, parse_signature_date)


def nominal_width(text, name, size, bold=False, italic=False):
    return sum(size * (0.5 if ch.isascii() else 1) for ch in text)


class SignatureDateTests(unittest.TestCase):
    def test_complete_dates_and_chinese_numerals(self):
        for text in ('2026年11月14日', '２０２６年１１月１４日', '二〇二六年十一月十四日',
                     '二零二六年一一月一四日', '2026-11-14', '2026/11/14', ' 2026年 11月 14日 '):
            with self.subTest(text=text):
                self.assertEqual(parse_signature_date(text), date(2026, 11, 14))

    def test_leading_zeros_and_leap_year(self):
        self.assertEqual(parse_signature_date('2024年02月29日'), date(2024, 2, 29))
        self.assertEqual(parse_signature_date('二〇二六年正月初一日'), None)
        self.assertEqual(parse_signature_date('2026/01/02'), date(2026, 1, 2))

    def test_invalid_incomplete_ambiguous_or_prose_dates(self):
        for text in ('2026年2月29日', '2026年13月1日', '2026年1月32日', '2026年0月1日',
                     '26年1月1日', '2026年1月', '2026年', '1月2日', '01/02/2026',
                     '2026-1/2', '2026.1.2', '自2026年1月2日起', '2026年1月2日至3日',
                     '2026年1月2日印发', '2026年1月2日。', '二千零二十六年一月二日'):
            with self.subTest(text=text):
                self.assertIsNone(parse_signature_date(text))

    def test_config_defaults_validation_and_no_mutation(self):
        self.assertFalse(DEFAULT_CONFIG['enable_signature_formatting'])
        self.assertFalse(DEFAULT_CONFIG['signature_include_attachments'])
        self.assertFalse(DEFAULT_CONFIG['signature_normalize_date'])
        self.assertEqual(normalize_format_options({}), {})
        data = {'signature_layout': 'unsealed', 'signature_before_spacing': '1.5'}
        self.assertEqual(normalize_format_options(data)['signature_before_spacing'], 1.5)
        self.assertEqual(data['signature_before_spacing'], '1.5')
        for key, values in (
            ('signature_layout', ['both', 1]), ('signature_before_spacing_mode', ['bad']),
            ('signature_before_spacing_unit', ['cm']),
            ('signature_before_spacing', [-1, 'nan', 'inf', 'text']),
            ('enable_signature_formatting', ['false', 1]),
        ):
            for value in values:
                with self.subTest(key=key, value=value), self.assertRaises(ValueError):
                    normalize_format_options({key: value})


class SignatureCase(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix='wfp_signature_test_')
        self.addCleanup(self.directory.cleanup)
        self.folder = Path(self.directory.name)
        self.logs = []

    @staticmethod
    def document(unit='某某市教育局', date_text='2026年11月14日'):
        doc = Document()
        p = doc.add_paragraph('关于开展工作的通知')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.size = Pt(22)
        doc.add_paragraph('现将有关工作安排如下，请认真组织实施。')
        doc.add_paragraph(unit)
        doc.add_paragraph(date_text)
        return doc

    def format(self, doc, **options):
        source, output = self.folder / 'source.docx', self.folder / 'output.docx'
        doc.save(source)
        self.source_bytes = source.read_bytes()
        processor = WordProcessor({**DEFAULT_CONFIG, 'convert_auto_numbering': False,
                                   'enable_signature_formatting': True, **options}, self.logs.append)
        try:
            with patch.object(processor, '_com_available', return_value=False), \
                 patch.object(FontMetrics, 'measure', side_effect=nominal_width):
                processor.format_document(str(source), str(output))
            self.assertEqual(source.read_bytes(), self.source_bytes)
            return Document(output)
        finally:
            processor._cleanup_temp_files()

    def roles(self, doc, **options):
        processor = WordProcessor({**DEFAULT_CONFIG, 'enable_signature_formatting': True, **options}, self.logs.append)
        blocks = list(processor._iter_block_items(doc))
        formatter = SignatureFormatter(processor)
        roles = formatter.classify(blocks, processor._classify_attachments(blocks, False))
        processor._cleanup_temp_files()
        return roles

    @staticmethod
    def attr(para, element, attribute):
        found = para._p.find('./' + qn('w:pPr') + '/' + qn('w:' + element))
        return found.get(qn('w:' + attribute)) if found is not None else None

    @staticmethod
    def image(para):
        png = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=')
        para.add_run().add_picture(BytesIO(png), width=Cm(0.4))


class SignatureDetectionTests(SignatureCase):
    def test_main_pair_at_end(self):
        roles = self.roles(self.document())
        self.assertEqual(len(roles), 2)
        self.assertEqual({role['kind'] for role in roles.values()}, {'unit', 'date'})

    def test_front_author_date_with_body_following_is_not_signature(self):
        doc = self.document('小恐龙公文排版团队')
        doc.add_paragraph('一、VBA脚本是什么')
        self.assertEqual(self.roles(doc), {})
        doc = self.document()
        doc.add_paragraph('接下来继续说明有关工作。')
        self.assertEqual(self.roles(doc), {})

    def test_person_sentence_heading_and_attachment_items_are_rejected(self):
        for unit in ('张三', '请联系某某市教育局', '某某市教育局。', '附件1：某某市教育局',
                     '一、某某市教育局', '某某市教育局、某某市财政局', '某某市教育局 某某市财政局'):
            with self.subTest(unit=unit):
                # “请联系...” is a sentence even without terminal punctuation.
                expected = 0
                self.assertEqual(len(self.roles(self.document(unit))), expected)

    def test_no_inference_for_date_alone_and_same_paragraph(self):
        doc = self.document()
        doc.paragraphs[-2]._p.getparent().remove(doc.paragraphs[-2]._p)
        self.assertEqual(self.roles(doc), {})
        doc = self.document()
        doc.paragraphs[-2].add_run(' 2026年11月14日')
        doc.paragraphs[-1]._p.getparent().remove(doc.paragraphs[-1]._p)
        self.assertEqual(self.roles(doc), {})

    def test_joint_signature_is_skipped(self):
        doc = self.document()
        doc.paragraphs[-2].insert_paragraph_before('某某市财政局')
        self.assertEqual(self.roles(doc), {})
        self.assertTrue(any('联合署名' in text for text in self.logs))

    def test_main_signature_before_attachment_and_optional_attachment_signature(self):
        doc = self.document()
        doc.add_paragraph('附件1')
        title = doc.add_paragraph('工作方案')
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph('一、附件的工作安排')
        doc.add_paragraph('某某市财政局')
        doc.add_paragraph('2026年12月1日')
        self.assertEqual(len(self.roles(doc)), 2)
        self.assertEqual(len(self.roles(doc, signature_include_attachments=True)), 4)
        result = self.format(doc, enable_attachment_formatting=False, signature_include_attachments=True)
        for text in ('2026年11月14日', '2026年12月1日'):
            para = next(p for p in result.paragraphs if p.text == text)
            self.assertEqual(para.paragraph_format.right_indent.pt, 64)
        self.assertFalse(next(p for p in result.paragraphs if p.text == '附件1').paragraph_format.page_break_before)

    def test_attachment_list_is_part_of_main_document(self):
        doc = self.document()
        doc.paragraphs[-2].insert_paragraph_before('附件：1.工作任务清单')
        self.assertEqual(len(self.roles(doc)), 2)

    def test_explicit_notes_copy_and_imprint_may_follow(self):
        doc = self.document()
        doc.add_paragraph('（此件主动公开）')
        doc.add_paragraph('抄送：有关单位')
        doc.add_paragraph('某某市教育局办公室 2026年11月15日印发')
        self.assertEqual(len(self.roles(doc)), 2)

    def test_table_after_pair_or_between_pair_is_a_barrier(self):
        doc = self.document()
        doc.add_table(rows=1, cols=1)
        self.assertEqual(self.roles(doc), {})
        doc = self.document()
        table = doc.add_table(rows=1, cols=1)
        doc.paragraphs[-1]._p.addprevious(table._tbl)
        self.assertEqual(self.roles(doc), {})

    def test_fields_revisions_and_content_controls_are_skipped(self):
        for tag in ('fldSimple', 'ins', 'del', 'sdt', 'object', 'pict'):
            with self.subTest(tag=tag):
                doc = self.document()
                node = OxmlElement('w:' + tag)
                if tag == 'fldSimple':
                    node.set(qn('w:instr'), 'DATE')
                doc.paragraphs[-1]._p.append(node)
                self.assertEqual(self.roles(doc), {})

    def test_image_in_pair_or_adjacent_is_preserved_and_not_repositioned(self):
        for where in ('unit', 'between', 'before', 'after'):
            with self.subTest(where=where):
                doc = self.document()
                if where == 'unit':
                    para = doc.paragraphs[-2]
                elif where == 'before':
                    para = doc.paragraphs[-2].insert_paragraph_before()
                elif where == 'between':
                    para = doc.paragraphs[-1].insert_paragraph_before()
                else:
                    para = doc.add_paragraph()
                self.image(para)
                self.assertEqual(self.roles(doc), {})
                result = self.format(doc)
                self.assertEqual(len(result.inline_shapes), 1)
                self.assertTrue(any('落款识别完成，确认 0' in text for text in self.logs))

    def test_special_spacing_scaling_or_soft_line_break_skips_pair(self):
        for tag, value in (('spacing', '20'), ('w', '80'), ('vertAlign', 'superscript')):
            doc = self.document()
            element = OxmlElement('w:' + tag)
            element.set(qn('w:val'), value)
            doc.paragraphs[-2].runs[0]._r.get_or_add_rPr().append(element)
            self.assertEqual(self.roles(doc), {})
        doc = self.document()
        doc.paragraphs[-2].runs[0].add_break()
        self.assertEqual(self.roles(doc), {})

    def test_tab_or_trailing_padding_is_conservatively_skipped(self):
        for text in ('2026年11月14日 ', '\t2026年11月14日', '2026年11月14日\u3000'):
            self.assertEqual(self.roles(self.document(date_text=text)), {})


class SignatureLayoutTests(SignatureCase):
    def test_disabled_uses_existing_format_flow(self):
        result = self.format(self.document(), enable_signature_formatting=False)
        self.assertNotEqual(result.paragraphs[-1].paragraph_format.right_indent, Pt(64))
        self.assertFalse(any('落款识别' in message for message in self.logs))

    def test_sealed_name_is_centered_over_date(self):
        doc = self.document()
        doc.paragraphs[-2].paragraph_format.space_before = Pt(12)
        result = self.format(doc)
        unit, date_para = result.paragraphs[-2:]
        self.assertEqual(date_para.paragraph_format.right_indent.pt, 64)
        self.assertEqual(unit.paragraph_format.right_indent.pt, 72)
        self.assertEqual(unit.paragraph_format.space_before.pt, 12)
        self.assertTrue(unit.paragraph_format.keep_with_next)
        self.assertTrue(unit.paragraph_format.keep_together)
        self.assertEqual(unit.paragraph_format.line_spacing.pt, 28)
        self.assertEqual(unit.paragraph_format.first_line_indent.pt, 0)
        for para in (unit, date_para):
            for tag in ('autoSpaceDE', 'autoSpaceDN', 'adjustRightInd', 'snapToGrid'):
                self.assertEqual(self.attr(para, tag, 'val'), '0')

    def test_sealed_different_name_lengths_share_date_center(self):
        for name in ('市教育局', '某某市教育局', '某某市人民政府办公室'):
            with self.subTest(name=name):
                result = self.format(self.document(name))
                unit, date_para = result.paragraphs[-2:]
                unit_center_from_right = unit.paragraph_format.right_indent.pt + len(name) * 8
                date_center_from_right = date_para.paragraph_format.right_indent.pt + 56
                self.assertAlmostEqual(unit_center_from_right, date_center_from_right, places=1)

    def test_unsealed_long_name_date_starts_two_chars_later(self):
        name = '某某市人民政府办公室'
        result = self.format(self.document(name), signature_layout='unsealed')
        unit, date_para = result.paragraphs[-2:]
        self.assertEqual(unit.paragraph_format.right_indent.pt, 32)
        self.assertEqual(unit.paragraph_format.space_before.pt, 30)
        name_left_from_right = unit.paragraph_format.right_indent.pt + len(name) * 16
        date_left_from_right = date_para.paragraph_format.right_indent.pt + 112
        self.assertAlmostEqual(name_left_from_right - date_left_from_right, 32)

    def test_unsealed_short_name_shifts_both_left(self):
        result = self.format(self.document('市教育局'), signature_layout='unsealed')
        unit, date_para = result.paragraphs[-2:]
        self.assertEqual(date_para.paragraph_format.right_indent.pt, 32)
        self.assertEqual(unit.paragraph_format.right_indent.pt, 112)

    def test_existing_blank_or_preserve_spacing_does_not_add_gap(self):
        doc = self.document()
        doc.paragraphs[-2].insert_paragraph_before()
        doc.paragraphs[-2].paragraph_format.space_before = Pt(7)
        result = self.format(doc, signature_layout='unsealed')
        self.assertEqual(len(result.paragraphs), 5)
        self.assertEqual(result.paragraphs[-2].paragraph_format.space_before.pt, 7)
        doc = self.document()
        doc.paragraphs[-2].paragraph_format.space_before = Pt(9)
        result = self.format(doc, signature_layout='unsealed', signature_before_spacing_mode='preserve')
        self.assertEqual(result.paragraphs[-2].paragraph_format.space_before.pt, 9)

    def test_inherited_character_indents_and_paragraph_mark_size_are_overridden(self):
        doc = self.document()
        style_ind = doc.styles['Normal'].element.get_or_add_pPr().get_or_add_ind()
        for key, value in (('leftChars', '100'), ('rightChars', '300'), ('firstLineChars', '200')):
            style_ind.set(qn('w:' + key), value)
        result = self.format(doc)
        for para, right in zip(result.paragraphs[-2:], (655, 582)):
            self.assertEqual(self.attr(para, 'ind', 'rightChars'), str(right))
            self.assertEqual(self.attr(para, 'ind', 'firstLineChars'), '0')
            self.assertEqual(self.attr(para, 'ind', 'leftChars'), '0')
            size = para._p.find('./' + qn('w:pPr') + '/' + qn('w:rPr') + '/' + qn('w:sz'))
            self.assertEqual(size.get(qn('w:val')), '32')
        self.assertEqual(style_ind.get(qn('w:rightChars')), '300')
        self.assertEqual(SignatureFormatter._indent_character_size(result), 11)
        result.styles['Normal'].font.size = Pt(20)
        self.assertEqual(SignatureFormatter._indent_character_size(result), 20)

    def test_character_grid_pitch_is_used_for_indents_but_not_glyph_spacing(self):
        doc = self.document()
        grid = doc.sections[0]._sectPr.find(qn('w:docGrid'))
        grid.set(qn('w:type'), 'linesAndChars')
        grid.set(qn('w:charSpace'), '4096')
        result = self.format(doc)
        unit, date_para = result.paragraphs[-2:]
        self.assertEqual(SignatureFormatter._indent_character_size(result, unit), 12)
        self.assertEqual(self.attr(date_para, 'ind', 'rightChars'), '533')
        for para in (unit, date_para):
            for run in para.runs:
                snap = run._r.rPr.find(qn('w:snapToGrid'))
                self.assertEqual(snap.get(qn('w:val')), '0')

    def test_custom_spacing_zero_points_and_lines_override_inherited_values(self):
        for value, unit, expected, expected_lines in ((0, 'pt', '0', '0'), (12, 'pt', '240', '0'),
                                                     (1.5, 'lines', '0', '150')):
            doc = self.document()
            result = self.format(doc, signature_before_spacing_mode='custom',
                                 signature_before_spacing=value, signature_before_spacing_unit=unit)
            para = result.paragraphs[-2]
            self.assertEqual(self.attr(para, 'spacing', 'before'), expected)
            self.assertEqual(self.attr(para, 'spacing', 'beforeLines'), expected_lines)
            self.assertEqual(self.attr(para, 'contextualSpacing', 'val'), '0')

    def test_bold_body_font_size_and_multiple_spacing(self):
        doc = self.document()
        doc.paragraphs[-2].runs[0].bold = True
        doc.paragraphs[-1].runs[0].bold = False
        result = self.format(doc, body_font='仿宋', body_size=14, line_spacing_unit='multiple', line_spacing_multiple=1.5)
        unit, date_para = result.paragraphs[-2:]
        self.assertTrue(unit.runs[0].bold)
        self.assertFalse(date_para.runs[0].bold)
        self.assertEqual(unit.runs[0].font.size.pt, 14)
        self.assertEqual(date_para.paragraph_format.right_indent.pt, 56)
        self.assertEqual(unit.paragraph_format.line_spacing, 1.5)

    def test_oversized_pair_and_multicolumn_preserve_original_pair(self):
        doc = self.document('某某市' + '特别长名称' * 8 + '教育局')
        doc.paragraphs[-2].alignment = WD_ALIGN_PARAGRAPH.LEFT
        result = self.format(doc)
        self.assertEqual(result.paragraphs[-2].alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self.assertTrue(any('超出版心' in text for text in self.logs))
        doc = self.document()
        cols = doc.sections[0]._sectPr.find(qn('w:cols'))
        cols.set(qn('w:num'), '2')
        result = self.format(doc)
        self.assertNotEqual(result.paragraphs[-1].paragraph_format.right_indent, Pt(64))

    def test_date_normalization_opt_in_and_text_node_markers_preserved(self):
        doc = self.document(date_text='二〇二六年十一月十四日')
        para = doc.paragraphs[-1]
        para.runs[0].text = '二〇二六年'
        start = OxmlElement('w:bookmarkStart')
        start.set(qn('w:id'), '50')
        start.set(qn('w:name'), 'signature_date')
        para._p.append(start)
        para.add_run('十一月十四日')
        end = OxmlElement('w:bookmarkEnd')
        end.set(qn('w:id'), '50')
        para._p.append(end)
        result = self.format(doc)
        self.assertEqual(result.paragraphs[-1].text, '二〇二六年十一月十四日')
        result = self.format(doc, signature_normalize_date=True)
        self.assertEqual(result.paragraphs[-1].text, '2026年11月14日')
        self.assertEqual(len(result.paragraphs[-1]._p.findall(qn('w:bookmarkStart'))), 1)
        self.assertEqual(len(result.paragraphs[-1]._p.findall(qn('w:bookmarkEnd'))), 1)
        self.assertEqual(len(result.paragraphs[-1].runs), 2)
        for text in ('2026年01月02日', '2026/01/02'):
            result = self.format(self.document(date_text=text), signature_normalize_date=True)
            self.assertEqual(result.paragraphs[-1].text, '2026年1月2日')

    def test_original_leading_spaces_and_date_text_remain_when_normalization_off(self):
        doc = self.document('  某某市教育局', '  2026年 01月 02日')
        result = self.format(doc)
        self.assertEqual(result.paragraphs[-2].text, '  某某市教育局')
        self.assertEqual(result.paragraphs[-1].text, '  2026年 01月 02日')
        self.assertEqual(result.paragraphs[-1].alignment, WD_ALIGN_PARAGRAPH.RIGHT)

    def test_hyperlink_and_relationship_survive_date_normalization(self):
        doc = self.document()
        para = doc.paragraphs[-1]
        para.runs[0].text = '2026年'
        relationship = para.part.relate_to('https://example.com/date', RT.HYPERLINK, is_external=True)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), relationship)
        run = OxmlElement('w:r')
        text = OxmlElement('w:t')
        text.text = '01月02日'
        run.append(text)
        hyperlink.append(run)
        para._p.append(hyperlink)
        result = self.format(doc, signature_normalize_date=True)
        para = result.paragraphs[-1]
        self.assertEqual(WordProcessor._paragraph_text(para), '2026年1月2日')
        self.assertEqual(len(para._p.findall(qn('w:hyperlink'))), 1)
        self.assertEqual(para.part.rels[relationship].target_ref, 'https://example.com/date')

    def test_original_breaks_blank_paragraphs_media_and_table_survive(self):
        doc = self.document()
        picture = doc.paragraphs[1].insert_paragraph_before()
        self.image(picture)
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = '表格内容'
        doc.paragraphs[1]._p.addprevious(table._tbl)
        doc.paragraphs[-2].paragraph_format.page_break_before = True
        doc.paragraphs[-1].insert_paragraph_before()
        before_count = len(doc.paragraphs)
        result = self.format(doc)
        self.assertEqual(len(result.paragraphs), before_count)
        self.assertTrue(result.paragraphs[-3].paragraph_format.page_break_before)
        self.assertEqual(len(result.tables), 1)
        self.assertEqual(result.tables[0].cell(0, 0).text, '表格内容')
        self.assertEqual(len(result.inline_shapes), 1)
        with ZipFile(BytesIO(self.source_bytes)) as source, ZipFile(self.folder / 'output.docx') as output:
            for name in source.namelist():
                if name.startswith('word/media/'):
                    self.assertEqual(source.read(name), output.read(name))

    def test_section_breaks_survive_and_width_uses_current_section(self):
        doc = self.document()
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        boundary = doc.paragraphs[-1]
        doc.paragraphs[2]._p.addprevious(boundary._p)
        doc.sections[-1].page_width = Cm(21)
        result = self.format(doc)
        self.assertEqual(len(result.sections), 2)
        self.assertEqual(result.sections[-1].start_type, WD_SECTION_START.NEW_PAGE)
        self.assertEqual(result.paragraphs[-1].paragraph_format.right_indent.pt, 64)

    def test_txt_and_markdown_use_signature_pass_after_conversion(self):
        for extension in ('.txt', '.md'):
            source = self.folder / ('input' + extension)
            content = '关于开展工作的通知\n\n现将有关工作安排如下。\n\n某某市教育局\n2026年01月02日\n'
            source.write_text(content, encoding='utf-8')
            output = self.folder / ('text-' + extension[1:] + '.docx')
            processor = WordProcessor({**DEFAULT_CONFIG, 'enable_signature_formatting': True,
                                       'signature_normalize_date': True, 'blank_line_mode': '不改动任何空行'},
                                      self.logs.append)
            try:
                with patch.object(processor, '_com_available', return_value=False), \
                     patch.object(FontMetrics, 'measure', side_effect=nominal_width):
                    processor.format_document(str(source), str(output))
                result = Document(output)
                para = next(p for p in result.paragraphs if p.text == '2026年1月2日')
                self.assertEqual(para.paragraph_format.right_indent.pt, 64)
                self.assertEqual(source.read_text(encoding='utf-8'), content)
            finally:
                processor._cleanup_temp_files()

    def test_repeat_formatting_is_stable(self):
        result = self.format(self.document(date_text='2026年01月02日'), signature_layout='unsealed',
                             signature_normalize_date=True, signature_before_spacing_mode='custom',
                             signature_before_spacing=1.5, signature_before_spacing_unit='lines')
        original = [p._p.xml for p in result.paragraphs[-2:]]
        result = self.format(result, signature_layout='unsealed', signature_normalize_date=True,
                             signature_before_spacing_mode='custom', signature_before_spacing=1.5,
                             signature_before_spacing_unit='lines')
        self.assertEqual([p._p.xml for p in result.paragraphs[-2:]], original)


class FontMetricTests(SignatureCase):
    def test_macos_and_linux_font_locations_without_desktop_tools(self):
        for platform in ('darwin', 'linux'):
            visited = []
            def visit(folder, pattern):
                visited.append(folder)
                return iter(())
            _font_candidates.cache_clear()
            try:
                with patch('wfp_signature.sys.platform', platform), \
                     patch('wfp_signature.Path.home', return_value=self.folder), \
                     patch.object(Path, 'rglob', visit), \
                     patch.dict(os.environ, {'XDG_DATA_HOME': str(self.folder / 'data'),
                                             'XDG_DATA_DIRS': str(self.folder / 'shared')}):
                    self.assertEqual(_font_candidates(), ())
                if platform == 'darwin':
                    self.assertIn(self.folder / 'Library/Fonts', visited)
                    self.assertIn(Path('/System/Library/Fonts'), visited)
                else:
                    self.assertIn(self.folder / 'data/fonts', visited)
                    self.assertIn(Path('/usr/share/fonts'), visited)
                    self.assertIn(self.folder / 'shared/fonts', visited)
            finally:
                _font_candidates.cache_clear()

    def font_file(self):
        from fontTools.fontBuilder import FontBuilder
        from fontTools.pens.ttGlyphPen import TTGlyphPen
        path = self.folder / 'metrics.ttf'
        chars = '10中'
        glyphs = ['.notdef'] + ['uni{:04X}'.format(ord(ch)) for ch in chars]
        builder = FontBuilder(1000, isTTF=True)
        builder.setupGlyphOrder(glyphs)
        builder.setupCharacterMap(dict(zip(map(ord, chars), glyphs[1:])))
        builder.setupGlyf({name: TTGlyphPen(None).glyph() for name in glyphs})
        builder.setupHorizontalMetrics(dict(zip(glyphs, [(500, 0), (350, 0), (550, 0), (1000, 0)])))
        builder.setupHorizontalHeader(ascent=800, descent=-200)
        builder.setupNameTable({'familyName': 'WFP Test Metrics', 'styleName': 'Regular',
                                'uniqueFontIdentifier': 'WFP-test', 'fullName': 'WFP Test Metrics',
                                'psName': 'WFPTestMetrics', 'version': 'Version 1.0'})
        builder.setupOS2(sTypoAscender=800, sTypoDescender=-200, usWinAscent=800, usWinDescent=200)
        builder.setupPost()
        builder.setupMaxp()
        builder.save(path)
        return path

    def test_real_font_advances_and_cached_resolution(self):
        path = self.font_file()
        _find_font.cache_clear()
        with patch('wfp_signature._font_candidates', return_value=((str(path), ()),)) as candidates:
            metrics = FontMetrics(self.logs.append)
            self.assertAlmostEqual(metrics.measure('1中0', 'WFP Test Metrics', 10), 19)
            self.assertAlmostEqual(metrics.measure('01', 'WFP Test Metrics', 20), 18)
            self.assertEqual(candidates.call_count, 1)
            self.assertEqual(self.logs, [])
        _find_font.cache_clear()

    def test_bad_font_file_is_skipped(self):
        bad = self.folder / 'bad.ttf'
        bad.write_bytes(b'not a font')
        good = self.font_file()
        _find_font.cache_clear()
        with patch('wfp_signature._font_candidates', return_value=((str(bad), ()), (str(good), ()))):
            self.assertEqual(_find_font('WFP Test Metrics', False, False), (str(good), 0))
        _find_font.cache_clear()

    def test_missing_font_fallback_warns_once(self):
        with patch('wfp_signature._find_font', return_value=None):
            metrics = FontMetrics(self.logs.append)
            self.assertEqual(metrics.measure('中文10', 'missing', 16), 48)
            self.assertEqual(metrics.measure('日期', 'missing', 14), 28)
            self.assertEqual(len(self.logs), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
