# -*- coding: utf-8 -*-
"""Tkinter GUI for Word Formatter Pro v2.7.3."""

import json
import logging
import os
import queue
import tempfile
import threading

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, Menu, font as tkfont
from tkinterdnd2 import DND_FILES, TkinterDnD

from wfp_config import DEFAULT_CONFIG, FONT_SIZE_MAP, PRESET_FONT_OPTIONS
from wfp_core import (
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
    BLANK_LINE_MODE_OPTIONS,
    LARGE_FOLDER_FILE_CONFIRM_THRESHOLD,
    SUPPORTED_FILE_EXTENSIONS,
    WPSAppManager,
    WordProcessor,
    _initialize_com_for_thread,
    _uninitialize_com_for_thread,
)

class WordFormatterGUI:
    def __init__(self, master):
        self.master = master
        master.title("Word文档智能排版工具 v2.7.3")
        master.geometry("1200x860")
        master.minsize(1200, 860)
        self.log_queue = queue.Queue()
        self.is_processing = False

        self.font_size_map = FONT_SIZE_MAP.copy()
        self.font_size_map_rev = {v: k for k, v in self.font_size_map.items()}
        self.default_params = DEFAULT_CONFIG.copy()
        self.font_separator = '── 已安装字体 ──'
        self.installed_fonts = self._get_installed_fonts()
        self.font_options = {
            key: self._with_installed_fonts(options)
            for key, options in PRESET_FONT_OPTIONS.items()
        }
        self.set_outline_var = tk.BooleanVar(value=self.default_params['set_outline'])
        self.enable_attachment_var = tk.BooleanVar(value=self.default_params['enable_attachment_formatting'])
        self.force_a4_var = tk.BooleanVar(value=self.default_params['force_a4'])
        self.use_custom_english_font_var = tk.BooleanVar(value=self.default_params['use_custom_english_font'])
        self.normalize_punctuation_var = tk.BooleanVar(value=self.default_params['normalize_punctuation'])
        self.enable_table_var = tk.BooleanVar(value=self.default_params['enable_table_formatting'])
        self.table_auto_col_width_var = tk.BooleanVar(value=self.default_params['table_auto_col_width'])
        self.table_header_bold_var = tk.BooleanVar(value=self.default_params['table_header_bold'])
        self.table_smart_align_var = tk.BooleanVar(value=self.default_params['table_smart_align'])
        self.table_unified_borders_var = tk.BooleanVar(value=self.default_params['table_unified_borders'])
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_text_var = tk.StringVar(value="")
        self.entries = {}
        self.attachment_option_widgets = []
        self.table_option_widgets = []
        
        self.default_config_path = "default_config.json"
        
        self.create_menu()
        self.create_widgets()
        self.load_initial_config()

        self.master.protocol("WM_DELETE_WINDOW", self._on_close)
        self.master.after(250, self.set_initial_pane_position)
        self.master.after(100, self._check_log_queue)

    def _get_installed_fonts(self):
        try:
            fonts = tkfont.families(self.master)
        except tk.TclError:
            return []

        unique_fonts = {
            font.strip()
            for font in fonts
            if font and font.strip() and not font.strip().startswith('@')
        }
        return sorted(unique_fonts, key=str.casefold)

    def _with_installed_fonts(self, preset_fonts):
        options = []
        seen = set()
        for font in preset_fonts:
            normalized = font.casefold()
            if normalized not in seen:
                options.append(font)
                seen.add(normalized)

        installed_fonts = [
            font for font in self.installed_fonts
            if font.casefold() not in seen
        ]
        if installed_fonts:
            options.append(self.font_separator)
            options.extend(installed_fonts)
        return options

    def _update_english_font_state(self):
        combo = self.entries.get('english_font')
        if combo:
            combo.configure(state='normal' if self.use_custom_english_font_var.get() else 'disabled')

    def _set_widgets_enabled(self, widgets, enabled):
        for widget in widgets:
            if not hasattr(widget, '_enabled_state'):
                try:
                    enabled_state = widget.cget('state') or 'normal'
                    widget._enabled_state = 'normal' if enabled_state == 'disabled' else enabled_state
                except tk.TclError:
                    widget._enabled_state = 'normal'
            try:
                widget.configure(state=widget._enabled_state if enabled else 'disabled')
            except tk.TclError:
                pass

    def _update_attachment_state(self):
        self._set_widgets_enabled(self.attachment_option_widgets, self.enable_attachment_var.get())

    def _update_table_state(self):
        self._set_widgets_enabled(self.table_option_widgets, self.enable_table_var.get())

    def _enable_dependent_widgets_for_config_load(self):
        self._set_widgets_enabled(self.attachment_option_widgets, True)
        self._set_widgets_enabled(self.table_option_widgets, True)
        english_font_combo = self.entries.get('english_font')
        if english_font_combo:
            english_font_combo.configure(state='normal')

    def _set_widget_value(self, widget, value, is_size=False):
        previous_state = None
        try:
            previous_state = widget.cget('state')
            if previous_state == 'disabled':
                widget.configure(state=getattr(widget, '_enabled_state', 'normal'))
        except tk.TclError:
            previous_state = None

        try:
            if is_size and isinstance(widget, ttk.Combobox):
                display_val = self.font_size_map_rev.get(value, str(value))
                widget.set(display_val)
            elif isinstance(widget, ttk.Combobox):
                widget.set(value)
                if value != self.font_separator:
                    widget._last_valid_value = value
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(value))
        finally:
            if previous_state == 'disabled':
                try:
                    widget.configure(state='disabled')
                except tk.TclError:
                    pass

    def set_initial_pane_position(self):
        # 获取窗口总宽度，设置左侧占约30%
        total_width = self.master.winfo_width()
        if total_width > 100:  # 确保窗口已经渲染
            left_width = int(total_width * 0.3)  # 左侧占30%
            # 找到PanedWindow并设置位置
            for widget in self.master.winfo_children():
                if isinstance(widget, ttk.PanedWindow):
                    widget.sashpos(0, left_width)
                    break

    def create_menu(self):
        menubar = Menu(self.master)
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_help_window)
        menubar.add_cascade(label="帮助", menu=help_menu)
        self.master.config(menu=menubar)

    def _show_help_tooltip(self, title, message):
        messagebox.showinfo(title, message, parent=self.master)
        
    def _create_help_label(self, parent, text, row, col):
        help_label = ttk.Label(parent, text="(?)", foreground="blue", cursor="hand2")
        help_label.grid(row=row, column=col, sticky='W', padx=(0, 5))
        help_label.bind("<Button-1>", lambda e: self._show_help_tooltip("识别规则说明", text))

    def create_widgets(self):
        main_pane = ttk.PanedWindow(self.master, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left_frame = ttk.Frame(main_pane, padding=5)
        main_pane.add(left_frame, weight=2)

        notebook = ttk.Notebook(left_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        self.notebook = notebook

        file_tab = ttk.Frame(notebook)
        notebook.add(file_tab, text=' 文件批量处理 ')
        
        list_frame = ttk.LabelFrame(file_tab, text="待处理文件列表（可拖拽文件或文件夹）")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        v_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        h_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL)
        self.file_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            selectmode=tk.EXTENDED
        )
        v_scrollbar.config(command=self.file_listbox.yview)
        h_scrollbar.config(command=self.file_listbox.xview)
        self.file_listbox.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)
        
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.handle_drop)
        self.placeholder_label = ttk.Label(self.file_listbox, text="可以拖拽文件或文件夹到这里", foreground="grey")
        
        file_button_frame = ttk.Frame(file_tab)
        file_button_frame.pack(fill=tk.X, pady=5)
        ttk.Button(file_button_frame, text="添加文件", command=self.add_files).grid(row=0, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="添加文件夹", command=self.add_folder).grid(row=0, column=1, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="移除文件", command=self.remove_files).grid(row=1, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(file_button_frame, text="清空列表", command=self.clear_list).grid(row=1, column=1, sticky='ew', padx=2, pady=2)
        file_button_frame.columnconfigure(0, weight=1)
        file_button_frame.columnconfigure(1, weight=1)

        text_tab = ttk.Frame(notebook)
        notebook.add(text_tab, text=' 直接输入文本 ')
        text_frame = ttk.LabelFrame(text_tab, text="在此处输入或粘贴文本")
        text_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.direct_text_input = scrolledtext.ScrolledText(text_frame, height=10, wrap=tk.WORD)
        self.direct_text_input.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure('Success.TButton', font=('Helvetica', 10, 'bold'), foreground='green')

        left_action_frame = ttk.Frame(left_frame)
        left_action_frame.pack(fill=tk.X, pady=(5, 0))
        self.start_btn = ttk.Button(
            left_action_frame,
            text="开始排版",
            style='Success.TButton',
            command=self.start_processing
        )
        self.start_btn.pack(fill=tk.X, ipady=8)

        progress_frame = ttk.Frame(left_frame)
        progress_frame.pack(fill=tk.X, pady=(5, 0))
        self.progressbar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            variable=self.progress_var,
            maximum=100
        )
        self.progressbar.pack(fill=tk.X)
        ttk.Label(progress_frame, textvariable=self.progress_text_var, foreground="grey").pack(anchor=tk.W)

        log_frame = ttk.LabelFrame(left_frame, text="调试日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        self.debug_text = scrolledtext.ScrolledText(log_frame, height=10, state='disabled', wrap=tk.WORD)
        self.debug_text.pack(fill=tk.BOTH, expand=True)

        right_frame = ttk.Frame(main_pane, padding=5)
        main_pane.add(right_frame, weight=4)
        
        canvas = tk.Canvas(right_frame)
        v_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=v_scrollbar.set)
        
        params_container = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=params_container, anchor='nw')
        
        params_frame = ttk.LabelFrame(params_container, text="参数设置", padding=10)
        params_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        params_frame.columnconfigure(1, weight=1)
        params_frame.columnconfigure(3, weight=1)
        params_frame.columnconfigure(5, weight=1)

        # Helper functions for creating widgets
        def create_entry(label, var_name, r, c):
            ttk.Label(params_frame, text=label).grid(row=r, column=c, sticky=tk.W, padx=3, pady=2)
            entry = ttk.Entry(params_frame, width=12)
            entry.grid(row=r, column=c+1, sticky=tk.EW, padx=3, pady=2)
            self.entries[var_name] = entry
            return entry
        
        def create_combo(label, var_name, opts, r, c, readonly=True): 
            ttk.Label(params_frame, text=label).grid(row=r, column=c, sticky=tk.W, padx=3, pady=2)
            state = 'readonly' if readonly else 'normal'
            combo = ttk.Combobox(params_frame, values=opts, state=state, width=15)
            combo.grid(row=r, column=c+1, sticky=tk.EW, padx=3, pady=2)
            if self.font_separator in opts:
                combo._last_valid_value = ''

                def remember_font_value(event, combo=combo):
                    current = combo.get().strip()
                    if current and current != self.font_separator:
                        combo._last_valid_value = current

                def reject_font_separator(event, combo=combo):
                    if combo.get() == self.font_separator:
                        combo.set(getattr(combo, '_last_valid_value', ''))
                    else:
                        remember_font_value(event, combo)

                combo.bind("<FocusIn>", remember_font_value, add="+")
                combo.bind("<<ComboboxSelected>>", reject_font_separator, add="+")
            self.entries[var_name] = combo
            return combo

        def create_font_size_combo(label, var_name, r, c):
            ttk.Label(params_frame, text=label).grid(row=r, column=c, sticky=tk.W, padx=3, pady=2)
            combo = ttk.Combobox(params_frame, values=list(self.font_size_map.keys()), width=15)
            combo.grid(row=r, column=c+1, sticky=tk.EW, padx=3, pady=2)
            self.entries[var_name] = combo
            return combo

        def create_section_header(text, help_text, r):
            header_frame = ttk.Frame(params_frame)
            header_frame.grid(row=r, column=0, columnspan=6, sticky='ew', pady=(6, 2))
            ttk.Label(header_frame, text=text, font=('Helvetica', 9, 'bold')).pack(side=tk.LEFT)
            if help_text:
                help_label = ttk.Label(header_frame, text="(?)", foreground="blue", cursor="hand2")
                help_label.pack(side=tk.LEFT, padx=(2, 0))
                help_label.bind("<Button-1>", lambda e, t=text, m=help_text: self._show_help_tooltip(f"{t} - 识别规则", m))
            ttk.Separator(params_frame, orient='horizontal').grid(row=r+1, column=0, columnspan=6, sticky='ew')
            return r + 2

        row = 0
        
        # Section: Page Layout
        row = create_section_header("页面设置", None, row)
        create_entry("上边距(cm)", 'margin_top', row, 0)
        create_entry("下边距(cm)", 'margin_bottom', row, 2)
        create_entry("页脚距(cm)", 'footer_distance', row, 4)
        row += 1
        create_entry("左边距(cm)", 'margin_left', row, 0)
        create_entry("右边距(cm)", 'margin_right', row, 2)
        ttk.Checkbutton(params_frame, text="强制设置为A4纸张", variable=self.force_a4_var).grid(row=row, column=4, columnspan=2, sticky=tk.W, padx=3)
        row += 1
        create_combo("页码对齐", 'page_number_align', ['奇偶分页', '居中'], row, 0)
        create_combo("页码字体", 'page_number_font', self.font_options['page_number'], row, 2, readonly=False)
        create_font_size_combo("页码字号", 'page_number_size', row, 4)
        row += 1

        # Section: Document Title
        title_help = "• 主标题: 识别文档开头的连续【居中】且【字体字号相同】的段落。\n• 副标题: 主标题下方，同样【居中】但【字体字号与主标题不同】的段落。\n• TXT文件: 会将首个非层级标题的段落视为题目。"
        row = create_section_header("标题样式", title_help, row)
        create_combo("题目字体", 'title_font', self.font_options['title'], row, 0, readonly=False)
        create_font_size_combo("题目字号", 'title_size', row, 2)
        create_entry("题目行距(磅)", 'title_line_spacing', row, 4)
        row += 1
        create_combo("副标题字体", 'subtitle_font', self.font_options['subtitle'], row, 0, readonly=False)
        create_font_size_combo("副标题字号", 'subtitle_size', row, 2)
        create_entry("副标题行距(磅)", 'subtitle_line_spacing', row, 4)
        row += 1
        
        # Section: Body and Headings
        headings_help = '• 一级标题: "一、", "二、" ...\n• 二级标题: "（一）", "（二）" ...\n• 三级标题: "1.", "2." ...\n• 四级标题: "(1)", "(2)" ...\n\n注：正文、三级、四级标题共用一套字体字号。'
        row = create_section_header("正文与层级", headings_help, row)
        create_combo("一级标题字体", 'h1_font', self.font_options['h1'], row, 0, readonly=False)
        create_font_size_combo("一级标题字号", 'h1_size', row, 2)
        row += 1
        create_combo("二级标题字体", 'h2_font', self.font_options['h2'], row, 0, readonly=False)
        create_font_size_combo("二级标题字号", 'h2_size', row, 2)
        row += 1
        create_combo("正文/三四级字体", 'body_font', self.font_options['body'], row, 0, readonly=False)
        create_font_size_combo("正文/三四级字号", 'body_size', row, 2)
        create_entry("正文行距(磅)", 'line_spacing', row, 4)
        row += 1
        create_entry("段落左缩进(cm)", 'left_indent_cm', row, 0)
        create_entry("段落右缩进(cm)", 'right_indent_cm', row, 2)
        row += 1

        # Section: Table Content
        table_help = (
            "• 默认不启用表格自动调整，启用后才会调整表头/内容字体、字号、行距、行高、列宽和边框。\n"
            "• 默认保留单元格原始对齐方式；勾选智能对齐后，表头/序号/短文本居中，数字靠右，长文本靠左。"
        )
        row = create_section_header("表格内容（实验功能）", table_help, row)
        ttk.Checkbutton(params_frame, text="启用表格自动调整（总开关）", variable=self.enable_table_var, command=self._update_table_state).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=2)
        table_auto_col_width_check = ttk.Checkbutton(params_frame, text="自动调整列宽", variable=self.table_auto_col_width_var)
        table_auto_col_width_check.grid(row=row, column=2, columnspan=2, sticky=tk.W, padx=3, pady=2)
        table_unified_borders_check = ttk.Checkbutton(params_frame, text="统一表格边框", variable=self.table_unified_borders_var)
        table_unified_borders_check.grid(row=row, column=4, columnspan=2, sticky=tk.W, padx=3, pady=2)
        row += 1
        table_header_font_combo = create_combo("表头字体", 'table_header_font', self.font_options['table'], row, 0, readonly=False)
        table_font_combo = create_combo("表格字体", 'table_font', self.font_options['table'], row, 2, readonly=False)
        table_size_combo = create_font_size_combo("表格字号", 'table_size', row, 4)
        row += 1
        table_line_spacing_entry = create_entry("表格行距(磅)", 'table_line_spacing', row, 0)
        table_row_height_entry = create_entry("表格行高(cm)", 'table_row_height_cm', row, 2)
        table_width_percent_entry = create_entry("表格宽度(%)", 'table_width_percent', row, 4)
        row += 1
        table_border_size_entry = create_entry("边框粗细(pt)", 'table_border_size_pt', row, 0)
        table_header_bold_check = ttk.Checkbutton(params_frame, text="表头行加粗", variable=self.table_header_bold_var)
        table_header_bold_check.grid(row=row, column=2, columnspan=2, sticky=tk.W, padx=3, pady=2)
        table_smart_align_check = ttk.Checkbutton(params_frame, text="智能调整单元格对齐", variable=self.table_smart_align_var)
        table_smart_align_check.grid(row=row, column=4, columnspan=2, sticky=tk.W, padx=3, pady=2)
        self.table_option_widgets = [
            table_auto_col_width_check, table_unified_borders_check,
            table_header_font_combo, table_font_combo, table_size_combo,
            table_line_spacing_entry, table_row_height_entry, table_width_percent_entry,
            table_border_size_entry, table_header_bold_check, table_smart_align_check
        ]
        self._update_table_state()
        row += 1
        
        # Section: Other Elements
        other_help = '• 图/表标题: 自动查找图片或表格【上方或下方】最近的、居中的、以"图"或"表"开头的段落。\n• 附件标识: 识别"附件1"、"附件："等独立段落。启用后将自动【段前分页】并按主副标题规则识别其自身标题。'
        row = create_section_header("其他元素", other_help, row)
        create_combo("表格标题字体", 'table_caption_font', self.font_options['table_caption'], row, 0, readonly=False)
        create_font_size_combo("表格标题字号", 'table_caption_size', row, 2)
        row += 1
        create_combo("图形标题字体", 'figure_caption_font', self.font_options['figure_caption'], row, 0, readonly=False)
        create_font_size_combo("图形标题字号", 'figure_caption_size', row, 2)
        row += 1
        ttk.Checkbutton(params_frame, text="启用附件格式化", variable=self.enable_attachment_var, command=self._update_attachment_state).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=2)
        attachment_font_combo = create_combo("附件标识字体", 'attachment_font', self.font_options['attachment'], row, 2, readonly=False)
        attachment_size_combo = create_font_size_combo("附件标识字号", 'attachment_size', row, 4)
        self.attachment_option_widgets = [attachment_font_combo, attachment_size_combo]
        self._update_attachment_state()
        row += 1

        # Section: Global Options
        ttk.Separator(params_frame, orient='horizontal').grid(row=row, column=0, columnspan=6, sticky='ew', pady=5)
        row += 1
        ttk.Checkbutton(params_frame, text="自动设置大纲级别 (用于生成导航目录)", variable=self.set_outline_var).grid(row=row, columnspan=6, sticky=tk.W, padx=3)
        row += 1
        ttk.Checkbutton(
            params_frame,
            text="自定义数字和字母字体",
            variable=self.use_custom_english_font_var,
            command=self._update_english_font_state
        ).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3)
        create_combo("数字和字母字体", 'english_font', self.font_options['english'], row, 2, readonly=False)
        self._update_english_font_state()
        row += 1
        blank_line_combo = create_combo("TXT/MD空行处理", 'blank_line_mode', BLANK_LINE_MODE_OPTIONS, row, 0)
        blank_line_combo.configure(width=42)
        blank_line_combo.grid_configure(columnspan=5)
        row += 1

        # 按钮区域
        ttk.Checkbutton(params_frame, text="启用符号标准化（实验功能，保守修复中英文标点混用）", variable=self.normalize_punctuation_var).grid(row=row, columnspan=6, sticky=tk.W, padx=3)
        row += 1

        button_frame = ttk.Frame(params_container)
        button_frame.pack(fill=tk.X, pady=5)
        
        # 配置按钮 - 2x2布局
        config_buttons = ttk.Frame(button_frame)
        config_buttons.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(config_buttons, text="加载配置", command=self.load_config).grid(row=0, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(config_buttons, text="保存配置", command=self.save_config).grid(row=0, column=1, sticky='ew', padx=2, pady=2)
        ttk.Button(config_buttons, text="保存为默认", command=self.save_default_config).grid(row=1, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(config_buttons, text="恢复内置默认", command=self.load_defaults).grid(row=1, column=1, sticky='ew', padx=2, pady=2)
        config_buttons.columnconfigure(0, weight=1)
        config_buttons.columnconfigure(1, weight=1)

        # 配置Canvas滚动
        def on_canvas_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # 调整Canvas内容宽度以适应Canvas
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)

        canvas.bind('<Configure>', on_canvas_configure)
        
        # 添加鼠标滚轮支持
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # 布局Canvas和滚动条
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._update_listbox_placeholder()

    def _append_log_message(self, message):
        try:
            self.debug_text.config(state='normal')
            self.debug_text.insert(tk.END, message + '\n')
            self.debug_text.config(state='disabled')
            self.debug_text.see(tk.END)
        except tk.TclError:
            pass

    def _check_log_queue(self):
        try:
            while True:
                self._append_log_message(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        try:
            self.master.after(100, self._check_log_queue)
        except tk.TclError:
            pass

    def log_to_debug_window(self, message):
        self.log_queue.put(message)

    def _drain_log_queue(self):
        try:
            while True:
                self._append_log_message(self.log_queue.get_nowait())
        except queue.Empty:
            pass

    def _clear_debug_log(self):
        self._drain_log_queue()
        try:
            self.debug_text.config(state='normal')
            self.debug_text.delete('1.0', tk.END)
            self.debug_text.config(state='disabled')
        except tk.TclError:
            pass

    def _run_on_main(self, callback, *args):
        try:
            self.master.after(0, lambda: callback(*args))
        except tk.TclError:
            pass

    def _set_progress(self, value, text=""):
        def update():
            try:
                self.progress_var.set(value)
                self.progress_text_var.set(text)
            except tk.TclError:
                pass
        self._run_on_main(update)
    
    def load_initial_config(self):
        if os.path.exists(self.default_config_path):
            try:
                with open(self.default_config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self._apply_config(config)
                self.log_to_debug_window(f"已加载默认配置文件: {self.default_config_path}")
            except Exception as e:
                self.log_to_debug_window(f"加载默认配置 '{self.default_config_path}' 失败: {e}。将使用内置默认值。")
                self.load_defaults()
        else:
            self.log_to_debug_window("未找到默认配置文件，将使用内置默认值。")
            self.load_defaults()

    @staticmethod
    def _legacy_blank_line_mode(remove_blank_lines):
        return BLANK_LINE_MODE_DELETE_SINGLE if remove_blank_lines else BLANK_LINE_MODE_KEEP_SINGLE

    def _apply_config(self, loaded_config):
        loaded_config = dict(loaded_config)
        if 'use_custom_english_font' not in loaded_config and loaded_config.get('use_times_new_roman'):
            loaded_config['use_custom_english_font'] = True
            loaded_config.setdefault('english_font', 'Times New Roman')
        if 'blank_line_mode' not in loaded_config:
            loaded_config['blank_line_mode'] = self._legacy_blank_line_mode(
                loaded_config.get('remove_blank_lines', True)
            )
        else:
            loaded_config['blank_line_mode'] = WordProcessor._normalize_blank_line_mode(
                loaded_config.get('blank_line_mode'),
                remove_blank_lines=loaded_config.get('remove_blank_lines', True)
            )
        loaded_config = {**self.default_params, **loaded_config}
        for key, default_value in self.default_params.items():
            value = loaded_config.get(key)
            if value is None or (isinstance(value, str) and not value.strip()):
                loaded_config[key] = default_value
        self.set_outline_var.set(loaded_config.get('set_outline', True))
        self.enable_attachment_var.set(loaded_config.get('enable_attachment_formatting', True))
        self.force_a4_var.set(loaded_config.get('force_a4', False))
        self.use_custom_english_font_var.set(loaded_config.get('use_custom_english_font', False))
        self.normalize_punctuation_var.set(loaded_config.get('normalize_punctuation', False))
        self.enable_table_var.set(loaded_config.get('enable_table_formatting', False))
        self.table_auto_col_width_var.set(loaded_config.get('table_auto_col_width', True))
        self.table_header_bold_var.set(loaded_config.get('table_header_bold', True))
        self.table_smart_align_var.set(loaded_config.get('table_smart_align', False))
        self.table_unified_borders_var.set(loaded_config.get('table_unified_borders', True))
        boolean_keys = [
            'set_outline', 'enable_attachment_formatting', 'force_a4',
            'use_custom_english_font', 'use_times_new_roman',
            'remove_blank_lines', 'normalize_punctuation',
            'enable_table_formatting', 'table_auto_col_width', 'table_header_bold',
            'table_smart_align', 'table_unified_borders'
        ]
        self._enable_dependent_widgets_for_config_load()
        for key, value in loaded_config.items():
            if key in boolean_keys: continue
            widget = self.entries.get(key)
            if widget:
                self._set_widget_value(widget, value, is_size=("_size" in key))
        self._update_english_font_state()
        self._update_attachment_state()
        self._update_table_state()

    def load_defaults(self):
        self._apply_config(self.default_params)
    
    def collect_config(self):
        config = {}
        for key, widget in self.entries.items():
            value = widget.get().strip()
            if isinstance(widget, ttk.Combobox) and value == self.font_separator:
                value = getattr(widget, '_last_valid_value', '').strip()
            if value == '' and key in self.default_params:
                config[key] = self.default_params[key]
                continue
            if "_size" in key and isinstance(widget, ttk.Combobox):
                if value in self.font_size_map:
                    config[key] = self.font_size_map[value]
                else:
                    try: config[key] = float(value)
                    except (ValueError, TypeError):
                        self.log_to_debug_window(f"警告: 无效的字号值 '{value}' for '{key}'. 使用默认值 16pt。")
                        config[key] = 16
            else:
                try: config[key] = float(value) if '.' in value else int(value)
                except (ValueError, TypeError): config[key] = value
        config['set_outline'] = self.set_outline_var.get()
        config['enable_attachment_formatting'] = self.enable_attachment_var.get()
        config['force_a4'] = self.force_a4_var.get()
        config['use_custom_english_font'] = self.use_custom_english_font_var.get()
        config['normalize_punctuation'] = self.normalize_punctuation_var.get()
        config['enable_table_formatting'] = self.enable_table_var.get()
        config['table_auto_col_width'] = self.table_auto_col_width_var.get()
        config['table_header_bold'] = self.table_header_bold_var.get()
        config['table_smart_align'] = self.table_smart_align_var.get()
        config['table_unified_borders'] = self.table_unified_borders_var.get()
        return config

    def save_config(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f: json.dump(self.collect_config(), f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", f"配置已保存至 {file_path}")
    
    def save_default_config(self):
        try:
            with open(self.default_config_path, 'w', encoding='utf-8') as f:
                json.dump(self.collect_config(), f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", f"当前配置已保存为默认配置。\n下次启动软件时将自动加载。")
        except Exception as e:
            messagebox.showerror("错误", f"保存默认配置失败: {e}")

    def load_config(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                self._apply_config(loaded_config)
                messagebox.showinfo("成功", "配置已加载")
            except Exception as e:
                messagebox.showerror("错误", f"加载配置文件失败: {e}")

    def _update_listbox_placeholder(self):
        if self.file_listbox.size() == 0:
            self.placeholder_label.place(in_=self.file_listbox, relx=0.5, rely=0.5, anchor=tk.CENTER)
        else:
            self.placeholder_label.place_forget()

    def handle_drop(self, event):
        paths = self.master.tk.splitlist(event.data)
        self._add_paths_to_listbox(paths)

    def _should_scan_folder(self, folder_path):
        file_count = 0
        for _, _, files in os.walk(folder_path):
            file_count += len(files)
            if file_count > LARGE_FOLDER_FILE_CONFIRM_THRESHOLD:
                folder_name = os.path.basename(os.path.normpath(folder_path)) or folder_path
                return messagebox.askyesno(
                    "确认",
                    f"文件夹“{folder_name}”包含超过 {LARGE_FOLDER_FILE_CONFIRM_THRESHOLD} 个文件，继续扫描可能需要较长时间。\n\n确定继续扫描吗？",
                    parent=self.master
                )
        return True

    def _add_paths_to_listbox(self, paths):
        current_files = set(self.file_listbox.get(0, tk.END))
        added_count = 0
        skipped_dirs = 0
        
        for path in paths:
            if os.path.isdir(path):
                if not self._should_scan_folder(path):
                    skipped_dirs += 1
                    self.log_to_debug_window(f"已跳过文件夹: {path}")
                    continue

                for root, _, files in os.walk(path):
                    for f in files:
                        if f.lower().endswith(SUPPORTED_FILE_EXTENSIONS):
                            full_path = os.path.join(root, f)
                            if full_path not in current_files:
                                self.file_listbox.insert(tk.END, full_path)
                                current_files.add(full_path)
                                added_count += 1
            elif os.path.isfile(path):
                if path.lower().endswith(SUPPORTED_FILE_EXTENSIONS):
                    if path not in current_files:
                        self.file_listbox.insert(tk.END, path)
                        current_files.add(path)
                        added_count += 1
        
        if added_count > 0:
            self.log_to_debug_window(f"通过按钮或拖拽添加了 {added_count} 个新文件。")
        if skipped_dirs > 0:
            self.log_to_debug_window(f"已跳过 {skipped_dirs} 个大文件夹。")
        
        self._update_listbox_placeholder()

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("所有支持的文件", "*.docx;*.doc;*.wps;*.txt;*.md"), ("Word 文档", "*.docx;*.doc"), ("WPS 文档", "*.wps"), ("纯文本", "*.txt"), ("Markdown", "*.md")])
        if files:
            self._add_paths_to_listbox(files)
        
    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self._add_paths_to_listbox([folder])

    def remove_files(self):
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请先在列表中选择要移除的文件。")
            return
        for index in sorted(selected_indices, reverse=True):
            self.file_listbox.delete(index)
        self._update_listbox_placeholder()

    def clear_list(self): 
        self.file_listbox.delete(0, tk.END)
        self._update_listbox_placeholder()

    def show_help_window(self):
        help_win = tk.Toplevel(self.master); help_win.title("使用说明"); help_win.geometry("600x600")
        help_text_widget = scrolledtext.ScrolledText(help_win, wrap=tk.WORD, state='disabled')
        help_text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        help_content = """
Word文档智能排版工具 v2.7.3 - 使用说明

本工具旨在提供一键式的专业文档排版体验，支持批量处理和高度自定义。

【核心功能模式】
1. 文件批量处理：可拖拽或添加 .docx, .doc, .wps, .txt, .md 文件。
2. 直接输入文本：直接粘贴文本进行排版（自动强制使用A4纸张）。

【操作流程】
1. 选择模式并添加内容。
2. （可选）在"参数设置"区调整格式，可点击各分区旁的 (?) 图标查看具体识别规则。
3. 点击"开始排版"，并选择输出位置。

【智能识别规则详解】
- 主标题与副标题:
  • 主标题: 识别文档开头的连续【居中】且【字体字号相同】的段落。
  • 副标题: 主标题下方，同样【居中】但【字体字号与主标题不同】的段落。
  • TXT/MD文件: 会将首个非层级标题的段落视为题目。

- 正文与层级标题:
  • 一级标题: “一、”, “二、” ...
  • 二级标题: “（一）”, “（二）” ...
  • 三级标题: “1.”, “2.” ...
  • 四级标题: “(1)”, “(2)” ...
  • 注：正文、三级、四级标题默认共用一套字体字号。

- 其他元素:
  • 图/表标题: 自动查找图片或表格【上方或下方】最近的、居中的、以“图”或“表”开头的段落。
  • 附件标识: 识别“附件1”、“附件：”等独立段落。启用附件格式化后，将自动【段前分页】并按主副标题规则识别其自身标题。
  • 表格内容: 默认不启用表格自动调整。启用后可分别设置表头/内容字体，并统一字号、行距、行高、列宽、边框，可选智能对齐。

【其他特性】
- 纸张设置：直接输入文本默认使用A4纸。文件处理默认保持原样，可勾选“强制设置为A4纸张”进行修改。
- 保留原文格式：统一格式时，会保留【加粗、斜体、下划线、字体颜色】等。
- 二级标题智能拆分：若二级标题后紧跟正文（如"（一）标题。正文..."），会自动在【同一个段落内】为标题和正文应用不同格式。
- 豁免内容：图片、嵌入对象等内容会自动跳过格式化；表格仅在勾选“启用表格自动调整”后处理。
- 参数自定义：所有核心参数均可在界面调整。配置方案可【保存】和【加载】。
- Markdown 支持：.md 文件会自动清理 Markdown 标记（标题#、粗体**、链接[]()、图片![]()等）后转为纯文本进行排版。
- 空行处理：TXT/MD 支持三种模式：不改动任何空行；删除单个空行且多个空行保留至1个空行；保留单个空行且多个空行保留至1个空行。默认使用“删除单个空行，多个空行保留至1个空行”。
- 跨平台说明：非 Windows 或未安装 pywin32 时，可处理 .docx/.txt/.md；.doc/.wps 需先另存为 .docx，且自动编号转文本等 COM 预处理会跳过。

【安全提示】
本工具【绝对不会】修改您的任何原始文件。所有操作都在后台的临时副本上进行，确保源文件100%安全。
"""
        help_text_widget.config(state='normal')
        help_text_widget.insert('1.0', help_content.strip())
        help_text_widget.config(state='disabled')

    def start_processing(self):
        if self.is_processing:
            messagebox.showinfo("提示", "正在处理中，请稍候...", parent=self.master)
            return

        warning_title = "处理前重要提示"
        warning_message = (
            "为了防止数据丢失，请在继续前关闭所有已打开的Word和WPS文档（包括wps、表格、PPT等所有文档）。\n\n"
            "本程序在转换文件格式时可能需要调用Word/WPS程序，这可能会影响未保存的工作。\n\n"
            "您确定要继续吗？"
        )
        if not messagebox.askokcancel(warning_title, warning_message, parent=self.master):
            self.log_to_debug_window("用户已取消操作。")
            return

        active_tab_index = self.notebook.index(self.notebook.select())
        collected_config = self.collect_config()
        file_list = []
        text_content = ""
        output_dir = None
        output_path = None

        if active_tab_index == 0:
            file_list = list(self.file_listbox.get(0, tk.END))
            if not file_list:
                messagebox.showwarning("警告", "文件列表为空，请先添加文件！", parent=self.master)
                return
            output_dir = filedialog.askdirectory(title="请选择一个文件夹用于存放处理后的文件")
            if not output_dir:
                return
        elif active_tab_index == 1:
            text_content = self.direct_text_input.get('1.0', tk.END).strip()
            if not text_content:
                messagebox.showwarning("警告", "文本框内容为空！", parent=self.master)
                return
            output_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                initialfile="formatted_document.docx"
            )
            if not output_path:
                return

        self._clear_debug_log()
        self.is_processing = True
        self.start_btn.config(state='disabled', text="排版中，请稍候...")
        self._set_progress(0, "开始处理...")

        def worker():
            com_initialized = _initialize_com_for_thread(self.log_to_debug_window)
            try:
                with WPSAppManager(self.log_to_debug_window) as com_mgr:
                    processor = WordProcessor(
                        collected_config,
                        self.log_to_debug_window,
                        com_manager=com_mgr
                    )
                    if active_tab_index == 0:
                        self._process_files(processor, file_list, output_dir)
                    else:
                        self._process_text(processor, text_content, output_path)
            except Exception as e:
                logging.error(f"处理过程中发生严重错误: {e}", exc_info=True)
                self.log_to_debug_window(f"\n❌ 处理过程中发生严重错误：\n{e}")
                self._set_progress(100, "处理失败")

                def show_error(err=e):
                    try:
                        messagebox.showerror("错误", f"处理过程中发生错误：\n{err}", parent=self.master)
                    except tk.TclError:
                        pass

                self._run_on_main(show_error)
            finally:
                _uninitialize_com_for_thread(com_initialized, self.log_to_debug_window)
                self._run_on_main(self._restore_after_processing)

        threading.Thread(target=worker, daemon=True).start()

    def _process_files(self, processor, file_list, output_dir):
        success_count, fail_count = 0, 0
        total = len(file_list)
        for i, input_path in enumerate(file_list, start=1):
            base_name = os.path.basename(input_path)
            self._set_progress((i - 1) / total * 100, f"处理中 {i}/{total}: {base_name}")
            try:
                self.log_to_debug_window(f"\n--- 开始处理文件 {i}/{total}: {base_name} ---")
                output_name = os.path.splitext(base_name)[0]
                output_path = os.path.join(output_dir, f"{output_name}_formatted.docx")
                processor.format_document(input_path, output_path)
                self.log_to_debug_window(f"✅ 文件处理成功，已保存至: {output_path}")
                success_count += 1
            except Exception as e:
                logging.error(f"处理文件失败: {input_path}\n{e}", exc_info=True)
                self.log_to_debug_window(f"\n❌ 处理文件 {base_name} 时发生严重错误：\n{e}")
                fail_count += 1
            finally:
                processor._cleanup_temp_files()

        summary_message = f"批量处理完成！\n\n成功: {success_count}个\n失败: {fail_count}个"
        if fail_count > 0:
            summary_message += "\n\n失败详情请查看日志窗口。"
        self._set_progress(100, f"完成（成功 {success_count} / 失败 {fail_count}）")
        self.log_to_debug_window(f"\n🎉 {summary_message}")

        def show_summary(msg=summary_message):
            try:
                messagebox.showinfo("完成", msg, parent=self.master)
            except tk.TclError:
                pass

        self._run_on_main(show_summary)

    def _process_text(self, processor, text_content, output_path):
        self._set_progress(20, "处理文本...")
        temp_file_path = None
        try:
            fd, temp_file_path = tempfile.mkstemp(suffix=".txt", text=True)
            with os.fdopen(fd, 'w', encoding='utf-8') as tmp:
                tmp.write(text_content)

            self.log_to_debug_window("\n--- 开始处理输入的文本 ---")
            processor.format_document(temp_file_path, output_path)
            self._set_progress(100, "完成")
            self.log_to_debug_window("\n🎉 排版全部完成！")

            def show_done(path=output_path):
                try:
                    messagebox.showinfo("完成", f"文档排版成功！\n文件已保存至：\n{path}", parent=self.master)
                except tk.TclError:
                    pass

            self._run_on_main(show_done)
        finally:
            processor._cleanup_temp_files()
            if temp_file_path and os.path.exists(temp_file_path):
                try:
                    os.remove(temp_file_path)
                    self.log_to_debug_window("  > 输入文本的临时文件已删除")
                except OSError:
                    pass

    def _restore_after_processing(self):
        self.is_processing = False
        try:
            self.start_btn.config(state='normal', text="开始排版")
        except tk.TclError:
            pass

    def _on_close(self):
        if self.is_processing:
            if not messagebox.askyesno("确认", "任务仍在进行中，确定要退出吗？", parent=self.master):
                return
        self._drain_log_queue()
        self.master.destroy()


def main():
    root = TkinterDnD.Tk()
    app = WordFormatterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
