# -*- coding: utf-8 -*-
"""Built-in unit tests for Word Formatter Pro."""

from __future__ import annotations

import os
import base64
from io import BytesIO
from zipfile import ZipFile
from lxml.etree import tostring as xml_bytes
import subprocess
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt

from wfp_config import DEFAULT_CONFIG
from wfp_process import converter_profile_uri, owned_temp_dir, run_isolated
from wfp_core import (
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
    BLANK_LINE_MODE_PRESERVE,
    LegacyConversionUnavailable,
    SofficeConverter,
    WordProcessor,
    BOLD_KEYS,
    normalize_format_options,
)


class FormatOptionCompatibilityTests(unittest.TestCase):
    def test_legacy_bold_and_table_switches_keep_their_meaning(self):
        import wfp_cli
        for enabled in (False, True):
            legacy = {key: enabled for key in BOLD_KEYS}
            legacy.update(table_smart_align=enabled, table_unified_borders=enabled)
            config = wfp_cli.normalize_config(legacy)
            self.assertEqual(config['title_bold'], 'bold' if enabled else 'preserve')
            self.assertEqual(config['table_header_bold'], 'bold' if enabled else 'preserve')
            self.assertEqual(config['table_cell_alignment'], 'smart' if enabled else 'preserve')
            self.assertEqual(config['table_border_mode'], 'unified' if enabled else 'preserve')
            self.assertNotIn('table_smart_align', config)
            self.assertNotIn('table_unified_borders', config)
            self.assertIs(legacy['title_bold'], enabled)

    def test_explicit_modes_take_precedence_and_defaults_keep_kaiti(self):
        import wfp_cli
        config = wfp_cli.normalize_config({'table_cell_alignment': 'right', 'table_smart_align': True})
        self.assertEqual(config['table_cell_alignment'], 'right')
        self.assertEqual(config['subtitle_font'], '楷体_GB2312')
        self.assertEqual(config['subtitle_bold'], 'preserve')
        self.assertEqual(config['table_alignment'], 'preserve')
        self.assertFalse(config['table_remove_text_wrapping'])
        with self.assertRaises(ValueError):
            normalize_format_options({'h2_bold': 'invalid'})

    def test_cli_config_layers_and_old_set_aliases(self):
        import json
        import wfp_cli
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'config.json'
            source.write_text(json.dumps({'table_cell_alignment': 'right', 'h2_bold': True}), encoding='utf-8')
            config, _ = wfp_cli.load_config(source, '{"table_smart_align": true, "h2_bold": false}')
            self.assertEqual(config['table_cell_alignment'], 'smart')
            self.assertEqual(config['h2_bold'], 'preserve')
            wfp_cli.apply_set_overrides(config, ['h2_bold=normal', 'table_unified_borders=false',
                                                   'table_alignment=center', 'table_smart_align=false'])
            self.assertEqual(config['h2_bold'], 'normal')
            self.assertEqual(config['table_cell_alignment'], 'preserve')
            self.assertEqual(config['table_border_mode'], 'preserve')
            self.assertEqual(config['table_alignment'], 'center')


class InlineFormattingProtectionTests(unittest.TestCase):
    PIXEL = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=')

    def processor(self, **options):
        config = dict(DEFAULT_CONFIG)
        config.update(options)
        return WordProcessor(config)

    @staticmethod
    def hyperlink(para, text):
        element = OxmlElement('w:hyperlink')
        element.set(qn('r:id'), para.part.relate_to('https://example.com/target', RT.HYPERLINK, is_external=True))
        run = OxmlElement('w:r')
        node = OxmlElement('w:t')
        node.text = text
        run.append(node)
        element.append(run)
        para._p.append(element)
        return element

    def test_split_keeps_original_objects_relationships_and_full_run_properties(self):
        doc = Document()
        para = doc.add_paragraph()
        start = OxmlElement('w:bookmarkStart')
        start.set(qn('w:id'), '42')
        start.set(qn('w:name'), 'kept_bookmark')
        para._p.append(start)
        link = self.hyperlink(para, '（一）标题。正文链接')
        run = next(WordProcessor._iter_inline_runs(para))
        run.italic = True
        run.underline = True
        run.font.superscript = True
        run.font.bold = True
        for name in ('commentRangeStart', 'commentRangeEnd', 'bookmarkEnd'):
            marker = OxmlElement('w:' + name)
            marker.set(qn('w:id'), '42')
            para._p.append(marker)
        field = para.add_run()
        for name, value in (('fldChar', 'begin'), ('instrText', ' PAGE '),
                            ('fldChar', 'separate'), ('t', '1'), ('fldChar', 'end')):
            node = OxmlElement('w:' + name)
            if name == 'fldChar':
                node.set(qn('w:fldCharType'), value)
            else:
                node.text = value
            field._r.append(node)
        para.add_run().add_picture(BytesIO(self.PIXEL))
        processor = self.processor(h2_bold='normal')
        before_text = processor._paragraph_text(para)
        protected = {tag: [xml_bytes(node) for node in para._p.iter(qn('w:' + tag))]
                     for tag in ('drawing', 'fldChar', 'instrText', 'bookmarkStart',
                                 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd')}
        rels = dict(para.part.rels)
        processor._format_inline_heading(para, len('（一）标题。'), True)
        self.assertEqual(processor._paragraph_text(para), before_text)
        self.assertIs(link.getparent(), para._p)
        for tag, before in protected.items():
            self.assertEqual([xml_bytes(node) for node in para._p.iter(qn('w:' + tag))], before)
        self.assertEqual(dict(para.part.rels), rels)
        runs = list(processor._iter_inline_runs(para))
        self.assertEqual([processor._inline_run_text(r) for r in runs[:2]], ['（一）标题。', '正文链接'])
        self.assertFalse(runs[0].bold)
        self.assertTrue(runs[1].bold)
        for run in runs[:2]:
            self.assertTrue(run.italic)
            self.assertTrue(run.underline)
            self.assertTrue(run.font.superscript)
        self.assertEqual(runs[0].font.name, DEFAULT_CONFIG['h2_font'])
        self.assertEqual(runs[1].font.name, DEFAULT_CONFIG['body_font'])
        count = len(runs)
        processor._format_inline_heading(para, len('（一）标题。'), True)
        self.assertEqual(len(list(processor._iter_inline_runs(para))), count)

    def test_split_with_text_and_picture_in_one_run_does_not_duplicate_picture(self):
        doc = Document()
        para = doc.add_paragraph()
        run = para.add_run('（一）标题。正文')
        run.add_picture(BytesIO(self.PIXEL))
        run.add_tab()
        run.add_text('尾文')
        run.add_break()
        before = WordProcessor._paragraph_text(para)
        self.processor()._format_inline_heading(para, len('（一）标题。'), False)
        self.assertEqual(WordProcessor._paragraph_text(para), before)
        self.assertEqual(len(list(para._p.iter(qn('w:drawing')))), 1)
        self.assertEqual(len(list(para._p.iter(qn('w:tab')))), 1)
        self.assertEqual(len(list(para._p.iter(qn('w:br')))), 1)

    def test_trim_and_bracket_normalization_preserve_objects_in_same_run(self):
        doc = Document()
        para = doc.add_paragraph()
        run = para.add_run('  (一)标题')
        run.add_picture(BytesIO(self.PIXEL))
        run.add_text('正文(保留)')
        drawing = next(para._p.iter(qn('w:drawing')))
        original = drawing.xml
        processor = self.processor()
        processor._strip_leading_whitespace(para)
        processor._normalize_heading_brackets(para)
        self.assertEqual(processor._paragraph_text(para), '（一）标题正文(保留)')
        self.assertEqual(drawing.xml, original)

    def test_bold_three_modes_override_inherited_style_and_preserve_body(self):
        from docx.enum.style import WD_STYLE_TYPE
        doc = Document()
        style = doc.styles.add_style('Inherited bold', WD_STYLE_TYPE.CHARACTER)
        style.font.bold = True
        para = doc.add_paragraph('标题正文')
        run = para.runs[0]
        run.style = style
        WordProcessor._apply_bold_to_first_chars(para, 2, 'normal')
        self.assertEqual([r.text for r in para.runs], ['标题', '正文'])
        self.assertIs(para.runs[0].bold, False)
        self.assertIsNone(para.runs[1].bold)
        WordProcessor._apply_bold_to_runs(para, 'preserve')
        self.assertIs(para.runs[0].bold, False)
        WordProcessor._apply_bold_to_runs(para, 'bold')
        self.assertTrue(all(r.bold and r.font.cs_bold for r in para.runs))

    def test_nested_textboxes_and_deleted_text_are_not_counted_twice(self):
        doc = Document()
        para = doc.add_paragraph('正文')
        drawing = OxmlElement('w:pict')
        textbox = OxmlElement('w:txbxContent')
        inner = OxmlElement('w:p')
        inner_run = OxmlElement('w:r')
        inner_text = OxmlElement('w:t')
        inner_text.text = '文本框文字'
        inner_run.append(inner_text)
        inner.append(inner_run)
        textbox.append(inner)
        drawing.append(textbox)
        para.runs[0]._r.append(drawing)
        deleted = OxmlElement('w:del')
        deleted_run = OxmlElement('w:r')
        deleted_text = OxmlElement('w:t')
        deleted_text.text = '已删除'
        deleted_run.append(deleted_text)
        deleted.append(deleted_run)
        para._p.append(deleted)
        self.assertEqual(WordProcessor._paragraph_text(para), '正文')
        self.assertEqual(len(list(WordProcessor._iter_inline_runs(para))), 1)

    def test_end_to_end_preserves_picture_links_fields_and_source(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source.docx'
            output = Path(directory) / 'output.docx'
            doc = Document()
            para = doc.add_paragraph('（一）')
            self.hyperlink(para, '标题。正文链接')
            para.add_run().add_picture(BytesIO(self.PIXEL))
            field = para.add_run()
            for name, value in (('fldChar', 'begin'), ('instrText', ' PAGE '),
                                ('fldChar', 'separate'), ('t', '1'), ('fldChar', 'end')):
                node = OxmlElement('w:' + name)
                if name == 'fldChar':
                    node.set(qn('w:fldCharType'), value)
                else:
                    node.text = value
                field._r.append(node)
            doc.add_paragraph('（二）标题。普通正文').runs[0].bold = True
            doc.save(source)
            original = source.read_bytes()
            processor = self.processor(h2_bold='normal')
            try:
                with patch.object(processor, '_com_available', return_value=False):
                    processor.format_document(str(source), str(output))
                formatted = Document(output)
                self.assertEqual([processor._paragraph_text(p) for p in formatted.paragraphs],
                                 ['（一）标题。正文链接1', '（二）标题。普通正文'])
                self.assertEqual(len(list(formatted.element.iter(qn('w:drawing')))), 1)
                self.assertEqual(len(list(formatted.element.iter(qn('w:hyperlink')))), 1)
                # Check the original body field separately from generated footer fields.
                self.assertEqual(len(list(formatted.element.body.iter(qn('w:fldChar')))), 3)
                self.assertEqual(next(formatted.element.body.iter(qn('w:instrText'))).text, ' PAGE ')
                self.assertFalse(formatted.paragraphs[1].runs[0].bold)
                self.assertTrue(formatted.paragraphs[1].runs[1].bold)
                with ZipFile(source) as before, ZipFile(output) as after:
                    for name in before.namelist():
                        if name.startswith('word/media/'):
                            self.assertEqual(before.read(name), after.read(name))
                self.assertEqual(source.read_bytes(), original)
                repeated = Path(directory) / 'repeated.docx'
                with patch.object(processor, '_com_available', return_value=False):
                    processor.format_document(str(output), str(repeated))
                repeated_doc = Document(repeated)
                self.assertEqual([processor._paragraph_text(p) for p in repeated_doc.paragraphs],
                                 [processor._paragraph_text(p) for p in formatted.paragraphs])
                self.assertEqual(len(list(repeated_doc.element.iter(qn('w:drawing')))), 1)
                self.assertEqual(len(list(repeated_doc.element.iter(qn('w:hyperlink')))), 1)
                self.assertEqual(len(list(repeated_doc.element.body.iter(qn('w:fldChar')))), 3)
            finally:
                processor._cleanup_temp_files()

    def test_end_to_end_all_heading_bold_modes_include_subtitle_and_attachment(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'headings.docx'
            output = Path(directory) / 'headings-out.docx'
            doc = Document()
            for text, centered, size in (('主标题', True, 22), ('副标题', True, 16),
                                         ('正文', False, 16), ('一、一级标题', False, 16),
                                         ('（一）二级标题。正文保留加粗', False, 16),
                                         ('附件1', False, 16), ('附件主标题', True, 22),
                                         ('附件副标题', True, 16), ('附件正文', False, 16)):
                para = doc.add_paragraph(text)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT
                para.runs[0].font.size = Pt(size)
                para.runs[0].bold = True
            doc.save(source)
            for mode in ('normal', 'bold', 'preserve'):
                processor = self.processor(**{key: mode for key in BOLD_KEYS})
                try:
                    with patch.object(processor, '_com_available', return_value=False):
                        processor.format_document(str(source), str(output))
                    paragraphs = Document(output).paragraphs
                    for index in (0, 1, 3, 4, 5, 6, 7):
                        self.assertEqual(paragraphs[index].runs[0].bold, mode != 'normal', (mode, index))
                    self.assertTrue(paragraphs[2].runs[0].bold)
                    self.assertTrue(paragraphs[4].runs[1].bold)
                finally:
                    processor._cleanup_temp_files()


class TableFormattingOptionsTests(unittest.TestCase):
    def make_table(self):
        doc = Document()
        table = doc.add_table(rows=2, cols=3)
        for row, texts in zip(table.rows, (('序号', '金额', '说明'), ('1', '123.45', '需要保留的较长说明文字'))):
            for cell, text in zip(row.cells, texts):
                cell.text = text
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        WordProcessor(DEFAULT_CONFIG)._set_table_indent(table, 240)
        floating = OxmlElement('w:tblpPr')
        floating.set(qn('w:tblpX'), '120')
        table._tbl.tblPr.append(floating)
        shade = OxmlElement('w:shd')
        shade.set(qn('w:fill'), 'FFCC00')
        table.cell(0, 0)._tc.get_or_add_tcPr().append(shade)
        processor = WordProcessor(DEFAULT_CONFIG)
        processor._set_table_borders(table, size_pt=2, color='FF0000')
        processor._set_cell_borders(table.cell(0, 0), size_pt=2, color='FF0000')
        return doc, table

    def format(self, doc, **options):
        config = {**DEFAULT_CONFIG, 'enable_table_formatting': True,
                  'table_auto_col_width': False, **options}
        WordProcessor(config)._format_tables(doc)

    def test_table_position_is_independent_of_cell_alignment(self):
        for mode, expected in (('left', WD_ALIGN_PARAGRAPH.LEFT),
                               ('center', WD_ALIGN_PARAGRAPH.CENTER), ('right', WD_ALIGN_PARAGRAPH.RIGHT)):
            doc, table = self.make_table()
            self.format(doc, table_alignment=mode)
            self.assertEqual(table.alignment, expected)
            self.assertEqual(table._tbl.tblPr.find(qn('w:tblInd')).get(qn('w:w')), '0')
            self.assertEqual(table.cell(1, 2).paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
            self.assertIsNotNone(table._tbl.tblPr.find(qn('w:tblpPr')))

    def test_smart_and_explicit_cell_modes_preserve_table_position(self):
        for mode, expected in (('left', WD_ALIGN_PARAGRAPH.LEFT),
                               ('center', WD_ALIGN_PARAGRAPH.CENTER), ('right', WD_ALIGN_PARAGRAPH.RIGHT)):
            doc, table = self.make_table()
            self.format(doc, table_cell_alignment=mode)
            self.assertEqual(table.alignment, WD_ALIGN_PARAGRAPH.RIGHT)
            self.assertEqual(table._tbl.tblPr.find(qn('w:tblInd')).get(qn('w:w')), '240')
            self.assertTrue(all(p.alignment == expected for row in table.rows for cell in row.cells for p in cell.paragraphs))
        doc, table = self.make_table()
        self.format(doc, table_cell_alignment='smart')
        self.assertEqual(table.cell(1, 0).paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(table.cell(1, 1).paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.RIGHT)
        self.assertEqual(table.cell(1, 2).paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.LEFT)

    def test_border_modes_keep_shading_and_preserve_original_borders(self):
        for mode in ('preserve', 'unified'):
            doc, table = self.make_table()
            before_table = xml_bytes(table._tbl.tblPr.find(qn('w:tblBorders')))
            before_cell = xml_bytes(table.cell(0, 0)._tc.tcPr.find(qn('w:tcBorders')))
            shade = xml_bytes(table.cell(0, 0)._tc.tcPr.find(qn('w:shd')))
            self.format(doc, table_border_mode=mode)
            self.assertEqual(xml_bytes(table.cell(0, 0)._tc.tcPr.find(qn('w:shd'))), shade)
            if mode == 'preserve':
                self.assertEqual(xml_bytes(table._tbl.tblPr.find(qn('w:tblBorders'))), before_table)
                self.assertEqual(xml_bytes(table.cell(0, 0)._tc.tcPr.find(qn('w:tcBorders'))), before_cell)
            else:
                self.assertNotEqual(xml_bytes(table._tbl.tblPr.find(qn('w:tblBorders'))), before_table)

    def test_wrapping_is_optional_and_disabled_table_formatting_is_untouched(self):
        doc, table = self.make_table()
        before = table._tbl.xml
        WordProcessor({**DEFAULT_CONFIG, 'table_remove_text_wrapping': True})._format_tables(doc)
        self.assertEqual(table._tbl.xml, before)
        self.format(doc, table_remove_text_wrapping=True)
        self.assertIsNone(table._tbl.tblPr.find(qn('w:tblpPr')))

    def test_header_normal_clears_bold_without_affecting_body(self):
        doc, table = self.make_table()
        table.cell(0, 0).paragraphs[0].runs[0].bold = True
        table.cell(1, 0).paragraphs[0].runs[0].bold = True
        self.format(doc, table_header_bold='normal')
        self.assertIs(table.cell(0, 0).paragraphs[0].runs[0].bold, False)
        self.assertIs(table.cell(1, 0).paragraphs[0].runs[0].bold, True)


class TextNormalizationTests(unittest.TestCase):
    def test_symbol_normalization_keeps_decimal_numbers(self):
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("你好,世界."),
            "你好，世界。",
        )
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("3.14 是 pi"),
            "3.14 是 pi",
        )

    def test_ellipsis_and_quotes(self):
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text('他说"你好"...'),
            "他说“你好”……",
        )
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("version 1.2"),
            "version 1.2",
        )

    def test_markdown_cleaning(self):
        raw = "# 标题\n**粗体** 和 [链接](https://example.com)\n![图片](a.png)\n> 引用"
        cleaned = WordProcessor._clean_markdown(raw)
        self.assertEqual(cleaned, "标题\n粗体 和 链接\n图片\n引用")

    def test_markdown_cleaning_preserves_source_numeric_numbering(self):
        raw = (
            "一、登录\n\n"
            "1. 打开软件\n"
            "2. 完成登录\n\n"
            "（一）首页搜索\n\n"
            "1. 输入关键词\n"
            "2. 点击院校\n"
            "1.2.3 小节号保持原样"
        )
        cleaned = WordProcessor._clean_markdown(raw)
        self.assertEqual(cleaned, raw)


class BlankLineTests(unittest.TestCase):
    def test_delete_single_blank_line_and_compress_multiple(self):
        text = "a\n\nb\n\n\nc"
        self.assertEqual(
            WordProcessor._remove_blank_lines_from_text(text),
            "a\nb\n\nc",
        )

    def test_keep_single_blank_line_and_compress_multiple(self):
        text = "a\n\nb\n\n\nc"
        self.assertEqual(
            WordProcessor._remove_blank_lines_from_text(text, keep_single_blank_lines=True),
            "a\n\nb\n\nc",
        )

    def test_blank_line_mode_aliases(self):
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("preserve"),
            BLANK_LINE_MODE_PRESERVE,
        )
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("delete_single"),
            BLANK_LINE_MODE_DELETE_SINGLE,
        )
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("keep_single"),
            BLANK_LINE_MODE_KEEP_SINGLE,
        )

    def test_processor_preserve_mode_leaves_text_unchanged(self):
        processor = WordProcessor({"blank_line_mode": BLANK_LINE_MODE_PRESERVE})
        text = "a\n\nb\n\n\nc"
        self.assertEqual(processor._normalize_text_blank_lines(text), text)


class TableHelperTests(unittest.TestCase):
    def test_numeric_table_text(self):
        self.assertTrue(WordProcessor._is_numeric_table_text("1,234.56"))
        self.assertTrue(WordProcessor._is_numeric_table_text("¥100元"))
        self.assertTrue(WordProcessor._is_numeric_table_text("12.5%"))
        self.assertFalse(WordProcessor._is_numeric_table_text("abc"))

    def test_short_table_text(self):
        self.assertTrue(WordProcessor._is_short_table_text("合计", max_len=4))
        self.assertFalse(WordProcessor._is_short_table_text("较长文本内容", max_len=4))

    def test_table_percentage_normalization(self):
        pcts = WordProcessor._normalize_table_pcts([1, 9], 20, 80)
        self.assertAlmostEqual(sum(pcts), 100.0)
        self.assertEqual(pcts, [20.0, 80.0])


class OoxmlProtectionTests(unittest.TestCase):
    def test_ooxml_element_detection(self):
        doc = Document()
        para = doc.add_paragraph()
        run = para.add_run()
        run._r.append(OxmlElement("w:drawing"))
        self.assertTrue(WordProcessor._has_drawing_or_pict(para))
        self.assertFalse(WordProcessor._has_field_codes(para))

        para_field = doc.add_paragraph()
        field_run = para_field.add_run()
        field_run._r.append(OxmlElement("w:fldChar"))
        self.assertTrue(WordProcessor._has_field_codes(para_field))


class SpacingAndIndentUnitTests(unittest.TestCase):
    def test_default_pt_line_spacing_writes_exact_twips(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        processor = WordProcessor(DEFAULT_CONFIG.copy())

        processor._apply_line_spacing(
            para,
            'line_spacing',
            'line_spacing_unit',
            'line_spacing_multiple',
            28,
        )

        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        self.assertEqual(spacing.get(qn('w:line')), '560')
        self.assertEqual(spacing.get(qn('w:lineRule')), 'exact')

    def test_multiple_line_spacing_writes_auto_240_units(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        config = DEFAULT_CONFIG.copy()
        config.update({'line_spacing_unit': 'multiple', 'line_spacing_multiple': 1.5})
        processor = WordProcessor(config)

        processor._apply_line_spacing(
            para,
            'line_spacing',
            'line_spacing_unit',
            'line_spacing_multiple',
            28,
        )

        spacing = para._p.get_or_add_pPr().get_or_add_spacing()
        self.assertEqual(spacing.get(qn('w:line')), '360')
        self.assertEqual(spacing.get(qn('w:lineRule')), 'auto')

    def test_default_indent_keeps_centimeter_fields_and_two_char_first_line(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        processor = WordProcessor(DEFAULT_CONFIG.copy())

        processor._apply_text_indent_and_align(para)

        ind = para._p.get_or_add_pPr().get_or_add_ind()
        self.assertEqual(ind.get(qn('w:left')), '0')
        self.assertEqual(ind.get(qn('w:right')), '0')
        self.assertEqual(ind.get(qn('w:firstLineChars')), '200')
        self.assertIsNone(ind.get(qn('w:leftChars')))
        self.assertIsNone(ind.get(qn('w:rightChars')))

    def test_character_indent_writes_character_fields(self):
        doc = Document()
        para = doc.add_paragraph("正文")
        config = DEFAULT_CONFIG.copy()
        config.update({
            'paragraph_indent_unit': 'chars',
            'left_indent_chars': 2,
            'right_indent_chars': 1.5,
            'first_line_indent_chars': 2,
        })
        processor = WordProcessor(config)

        processor._apply_text_indent_and_align(para)

        ind = para._p.get_or_add_pPr().get_or_add_ind()
        self.assertEqual(ind.get(qn('w:leftChars')), '200')
        self.assertEqual(ind.get(qn('w:rightChars')), '150')
        self.assertEqual(ind.get(qn('w:firstLineChars')), '200')
        self.assertIsNone(ind.get(qn('w:left')))
        self.assertIsNone(ind.get(qn('w:right')))

    def test_strip_leading_whitespace_removes_plain_blank_run(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("   ")
        para.add_run("正文")
        WordProcessor({})._strip_leading_whitespace(para)
        self.assertEqual(para.text, "正文")

    def test_strip_leading_whitespace_preserves_special_run(self):
        doc = Document()
        para = doc.add_paragraph()
        special_run = para.add_run()
        special_run._r.append(OxmlElement("w:fldChar"))
        para.add_run("正文")

        WordProcessor({})._strip_leading_whitespace(para)

        self.assertTrue(WordProcessor._has_field_codes(para))
        self.assertEqual(len(para.runs), 2)
        self.assertEqual(para.text, "正文")

    def test_strip_leading_tabs_preserves_following_field_nodes(self):
        doc = Document()
        para = doc.add_paragraph('\t\t  附件1')
        field = OxmlElement('w:fldChar')
        field.set(qn('w:fldCharType'), 'begin')
        para.runs[0]._r.append(field)
        WordProcessor({})._strip_leading_whitespace(para)
        self.assertEqual(para.text, '附件1')
        self.assertIs(field.getparent(), para.runs[0]._r)


class TempAndConversionTests(unittest.TestCase):
    def test_temp_docx_path_uses_system_temp_and_safe_name(self):
        processor = WordProcessor(DEFAULT_CONFIG.copy())
        temp_path = processor._make_temp_docx_path("copy", 'bad<name>:"?')
        try:
            self.assertEqual(
                os.path.normcase(os.path.abspath(os.path.dirname(temp_path))),
                os.path.normcase(os.path.abspath(tempfile.gettempdir())),
            )
            self.assertRegex(os.path.basename(temp_path), r"^~temp_copy_bad_name_[0-9]+_[0-9a-f]{8}\.docx$")
            self.assertIn(temp_path, processor.temp_files)
        finally:
            processor._cleanup_temp_files()

    def test_txt_conversion_uses_blank_line_mode_and_cleans_temp_file(self):
        with tempfile.TemporaryDirectory(prefix="wfp273_test_") as tmpdir:
            source = Path(tmpdir) / "sample.txt"
            source.write_text("标题\n\n正文一\n\n\n正文二", encoding="utf-8")

            processor = WordProcessor(DEFAULT_CONFIG.copy())
            temp_docx, is_from_txt = processor.convert_to_docx(str(source))
            try:
                self.assertTrue(is_from_txt)
                self.assertTrue(os.path.exists(temp_docx))
                self.assertEqual(
                    os.path.normcase(os.path.abspath(os.path.dirname(temp_docx))),
                    os.path.normcase(os.path.abspath(tempfile.gettempdir())),
                )
                self.assertFalse(list(Path(tmpdir).glob("~temp_*.docx")))

                converted = Document(temp_docx)
                self.assertEqual(
                    [p.text for p in converted.paragraphs],
                    ["标题", "正文一", "", "正文二"],
                )
            finally:
                processor._cleanup_temp_files()
            self.assertFalse(os.path.exists(temp_docx))

    def test_missing_soffice_marks_legacy_conversion_skipped(self):
        class MissingConverter:
            available = False

        processor = WordProcessor(DEFAULT_CONFIG.copy())
        processor.soffice_converter = MissingConverter()
        with self.assertRaises(LegacyConversionUnavailable):
            processor._convert_legacy_with_soffice("legacy.doc", "unused.docx")

    def test_format_document_skips_wps_when_soffice_missing(self):
        class MissingConverter:
            available = False

        with tempfile.TemporaryDirectory(prefix="wfp_legacy_skip_test_") as tmpdir:
            source = Path(tmpdir) / "legacy.wps"
            source.write_bytes(b"not a real wps file")

            processor = WordProcessor(DEFAULT_CONFIG.copy())
            processor.soffice_converter = MissingConverter()
            with patch.object(processor, '_com_available', return_value=False), self.assertRaises(LegacyConversionUnavailable):
                processor.format_document(str(source), str(Path(tmpdir) / "legacy_formatted.docx"))


class LegacyFormatTests(unittest.TestCase):
    def test_format_document_handles_doc_via_soffice_when_available(self):
        converter = SofficeConverter()
        if not converter.available:
            self.skipTest("LibreOffice soffice is not available")

        with owned_temp_dir(prefix="wfp_legacy_doc_test_") as tmpdir:
            root = Path(tmpdir)
            docx_source = root / "legacy_source.docx"
            source_doc = Document()
            source_doc.add_paragraph("旧格式标题")
            source_doc.add_paragraph("第一段正文")
            source_doc.save(docx_source)

            legacy_dir = root / "legacy"
            legacy_dir.mkdir()
            profile = root / "lo_profile"
            profile.mkdir()
            proc = run_isolated(
                [
                    converter.soffice_path,
                    "--headless",
                    f"-env:UserInstallation={converter_profile_uri(profile)}",
                    "--convert-to",
                    "doc",
                    "--outdir",
                    str(legacy_dir),
                    str(docx_source),
                ],
                timeout=120,
            )
            generated = sorted(legacy_dir.glob("*.doc"))
            if proc.returncode != 0 or not generated:
                detail = (proc.stderr or proc.stdout or "LibreOffice did not generate a .doc file").strip()
                self.fail(f"LibreOffice .doc export failed (exit {proc.returncode}, {converter.soffice_path}): {detail}")

            output = root / "legacy_formatted.docx"
            processor = WordProcessor(DEFAULT_CONFIG.copy(), soffice_path=converter.soffice_path)
            try:
                with patch.object(processor, '_com_available', return_value=False):
                    processor.format_document(str(generated[0]), str(output))
                self.assertTrue(output.exists())
                formatted = Document(output)
                self.assertTrue(any(paragraph.text.strip() for paragraph in formatted.paragraphs))
            finally:
                processor._cleanup_temp_files()
                processor.quit_com_app()


class EndToEndFormatTests(unittest.TestCase):
    def test_format_document_handles_direct_formats(self):
        with tempfile.TemporaryDirectory(prefix="wfp_format_test_") as tmpdir:
            root = Path(tmpdir)
            sources = []

            txt_source = root / "sample.txt"
            txt_source.write_text("测试标题\n\n第一段正文\n1. 小标题\n第二段正文", encoding="utf-8")
            sources.append(txt_source)

            md_source = root / "sample.md"
            md_source.write_text("# Markdown标题\n\n**第一段**\n\n1. 步骤一", encoding="utf-8")
            sources.append(md_source)

            docx_source = root / "sample.docx"
            source_doc = Document()
            source_doc.add_paragraph("Word标题")
            source_doc.add_paragraph("第一段正文")
            source_doc.save(docx_source)
            sources.append(docx_source)

            for source in sources:
                processor = WordProcessor(DEFAULT_CONFIG.copy())
                output = root / f"{source.stem}_formatted.docx"
                try:
                    processor.format_document(str(source), str(output))
                    self.assertTrue(output.exists(), f"missing output for {source.suffix}")
                    formatted = Document(output)
                    texts = [paragraph.text for paragraph in formatted.paragraphs]
                    self.assertTrue(any(text.strip() for text in texts), f"empty output for {source.suffix}")
                finally:
                    processor._cleanup_temp_files()


def main(argv=None):
    if argv:
        suite = unittest.defaultTestLoader.loadTestsFromNames(argv)
        return 0 if unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful() else 1
    suite = unittest.defaultTestLoader.loadTestsFromModule(__import__(__name__))
    import wfp_numbering_tests
    import wfp_numbering_integration_tests
    import wfp_attachment_tests
    import wfp_signature_tests
    import wfp_safety_tests
    import wfp_title_after_tests
    for module in (wfp_numbering_tests, wfp_numbering_integration_tests, wfp_attachment_tests, wfp_signature_tests, wfp_safety_tests, wfp_title_after_tests):
        suite.addTests(unittest.defaultTestLoader.loadTestsFromModule(module))
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    if result.wasSuccessful():
        print("所有单元测试通过")
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
