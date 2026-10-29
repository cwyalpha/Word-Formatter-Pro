# -*- coding: utf-8 -*-
"""Shared configuration defaults for Word Formatter Pro."""

from wfp_core import DEFAULT_BLANK_LINE_MODE

FONT_SIZE_MAP = {
    '一号 (26pt)': 26, '小一 (24pt)': 24, '二号 (22pt)': 22, '小二 (18pt)': 18,
    '三号 (16pt)': 16, '小三 (15pt)': 15, '四号 (14pt)': 14, '小四 (12pt)': 12,
    '五号 (10.5pt)': 10.5, '小五 (9pt)': 9,
}

DEFAULT_CONFIG = {
    'convert_auto_numbering': True,  # Non-Windows option; Windows uses COM or Python fallback.
    'page_number_align': '奇偶分页', 'footer_distance': 2.5, 'line_spacing': 28,
    'margin_top': 3.7, 'margin_bottom': 3.5, 'margin_left': 2.8, 'margin_right': 2.6,
    'title_font': '方正小标宋简体', 'h1_font': '黑体', 'h2_font': '楷体_GB2312',
    'body_font': '仿宋_GB2312', 'page_number_font': '宋体',
    'table_caption_font': '黑体', 'figure_caption_font': '黑体', 'attachment_font': '黑体',
    'subtitle_font': '黑体',
    'title_size': 22, 'h1_size': 16, 'h2_size': 16, 'body_size': 16, 'page_number_size': 14,
    'table_caption_size': 14, 'figure_caption_size': 14, 'attachment_size': 16,
    'subtitle_size': 16, 'title_line_spacing': 33, 'subtitle_line_spacing': 33, 
    'signature_spacing': 2,
    'signature_space_lines': 0,      # 落款段前：按正文行数
    'signature_space_pt': 0.0,       # 落款段前：直接磅值，优先级更高
    'line_spacing_unit': 'pt', 'line_spacing_multiple': 1.0,
    'title_line_spacing_unit': 'pt', 'title_line_spacing_multiple': 1.0,
    'subtitle_line_spacing_unit': 'pt', 'subtitle_line_spacing_multiple': 1.0,
    'left_indent_cm': 0.0, 'right_indent_cm': 0.0,
    'paragraph_indent_unit': 'cm',
    'left_indent_chars': 0.0, 'right_indent_chars': 0.0, 'first_line_indent_chars': 2.0,
    'title_bold': False, 'h1_bold': False, 'h2_bold': False,
    'set_outline': True, 'enable_attachment_formatting': True,
    'attachment_bold': False,
    'force_a4': False, 'use_custom_english_font': False, 'english_font': 'Times New Roman',
    'blank_line_mode': DEFAULT_BLANK_LINE_MODE, 'normalize_punctuation': False,
    'enable_table_formatting': False, 'table_header_font': '仿宋_GB2312',
    'table_font': '仿宋_GB2312', 'table_size': 12, 'table_line_spacing': 22,
    'table_line_spacing_unit': 'pt', 'table_line_spacing_multiple': 1.0,
    'table_row_height_cm': 0.7, 'table_auto_col_width': True, 'table_width_percent': 100,
    'table_header_bold': True, 'table_smart_align': False,
    'table_unified_borders': True, 'table_border_size_pt': 0.5,
    'table_preserve_borders': False, # 新增：是否保留原表格边框与底纹
    'output_dir': '',                # 新增：记忆上一次的输出目录
}

PRESET_FONT_OPTIONS = {
    'title': ['方正小标宋简体', '方正小标宋_GBK', '华文中宋'],
    'h1': ['黑体', '方正黑体_GBK', '方正黑体简体', '华文黑体'],
    'h2': ['楷体_GB2312', '方正楷体_GBK', '楷体', '方正楷体简体', '华文楷体'],
    'body': ['仿宋_GB2312', '方正仿宋_GBK', '仿宋', '方正仿宋简体', '华文仿宋'],
    'page_number': ['宋体', 'Times New Roman'],
    'table_caption': ['黑体', '宋体', '仿宋_GB2312'],
    'figure_caption': ['黑体', '宋体', '仿宋_GB2312'],
    'attachment': ['黑体', '宋体', '仿宋_GB2312'],
    'subtitle': ['楷体_GB2312', '方正楷体_GBK', '楷体', '方正楷体简体', '华文楷体'],
    'table': ['仿宋_GB2312', '宋体', '黑体', '楷体_GB2312', '方正仿宋_GBK', '仿宋'],
    'english': ['Times New Roman'],
}
