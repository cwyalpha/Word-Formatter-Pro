# -*- coding: utf-8 -*-
"""Tkinter GUI for Word Formatter Pro."""

import json
import logging
import os
import queue
import sys
import tempfile
import threading


def _configure_tcl_tk_paths():
    if getattr(sys, 'frozen', False):
        return
    prefix = getattr(sys, 'prefix', '')
    dll_dir = os.path.join(prefix, 'Library', 'bin')
    if os.path.isdir(dll_dir) and hasattr(os, 'add_dll_directory'):
        try:
            os.add_dll_directory(dll_dir)
        except OSError:
            pass
    preferred_libraries = None
    if os.name == 'nt':
        try:
            import ctypes
            library_tcl_dll = os.path.join(prefix, 'Library', 'bin', 'tcl86t.dll')
            library_tk_dll = os.path.join(prefix, 'Library', 'bin', 'tk86t.dll')
            dlls_tcl_dll = os.path.join(prefix, 'DLLs', 'tcl86t.dll')
            dlls_tk_dll = os.path.join(prefix, 'DLLs', 'tk86t.dll')
            if os.path.exists(library_tcl_dll):
                tcl_dll = ctypes.CDLL(library_tcl_dll)
                if os.path.exists(library_tk_dll):
                    ctypes.CDLL(library_tk_dll)
                preferred_libraries = (
                    os.path.join(prefix, 'Library', 'lib', 'tcl8.6'),
                    os.path.join(prefix, 'Library', 'lib', 'tk8.6'),
                )
            elif os.path.exists(dlls_tcl_dll):
                tcl_dll = ctypes.CDLL(dlls_tcl_dll)
                if os.path.exists(dlls_tk_dll):
                    ctypes.CDLL(dlls_tk_dll)
                preferred_libraries = (
                    os.path.join(prefix, 'tcl', 'tcl8.6'),
                    os.path.join(prefix, 'tcl', 'tk8.6'),
                )
            else:
                tcl_dll = None
            if tcl_dll is not None:
                tcl_dll.Tcl_FindExecutable.argtypes = [ctypes.c_char_p]
                executable = sys.executable.replace('\\', '/').encode('utf-8')
                tcl_dll.Tcl_FindExecutable(executable)
        except Exception:
            preferred_libraries = None
    library_candidates = []
    if preferred_libraries:
        library_candidates.append(preferred_libraries)
    library_candidates.extend([
        (
            os.path.join(prefix, 'Library', 'lib', 'tcl8.6'),
            os.path.join(prefix, 'Library', 'lib', 'tk8.6'),
        ),
        (
            os.path.join(prefix, 'tcl', 'tcl8.6'),
            os.path.join(prefix, 'tcl', 'tk8.6'),
        ),
        (
            os.path.join(prefix, 'Lib', 'tcl8.6'),
            os.path.join(prefix, 'Lib', 'tk8.6'),
        ),
    ])
    for tcl_library, tk_library in library_candidates:
        if os.path.exists(os.path.join(tcl_library, 'init.tcl')) and os.path.exists(os.path.join(tk_library, 'tk.tcl')):
            os.environ.setdefault('TCL_LIBRARY', tcl_library)
            os.environ.setdefault('TK_LIBRARY', tk_library)
            break


_configure_tcl_tk_paths()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, Menu, font as tkfont
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    TKDND_AVAILABLE = False

from wfp_config import DEFAULT_CONFIG, FONT_SIZE_MAP, PRESET_FONT_OPTIONS
from wfp_core import (
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
    BLANK_LINE_MODE_OPTIONS,
    IS_WINDOWS,
    LARGE_FOLDER_FILE_CONFIRM_THRESHOLD,
    LegacyConversionUnavailable,
    SUPPORTED_FILE_EXTENSIONS,
    WPSAppManager,
    WordProcessor,
    _initialize_com_for_thread,
    _uninitialize_com_for_thread,
)
from wfp_version import APP_TITLE, __version__

UNIT_DISPLAY_TO_CONFIG = {
    '磅值': 'pt',
    '倍数': 'multiple',
    '厘米': 'cm',
    '字符': 'chars',
}
UNIT_CONFIG_TO_DISPLAY = {
    'pt': '磅值',
    'multiple': '倍数',
    'cm': '厘米',
    'chars': '字符',
}
UNIT_VALUE_SUFFIX = {
    'pt': '磅',
    'multiple': '倍',
    'cm': '厘米',
    'chars': '字符',
}

class WordFormatterGUI:
    def __init__(self, master):
        self.master = master
        master.title(f"{APP_TITLE}")
        self.left_pane_min_width = 360
        self.right_pane_min_width = 520
        self.preferred_left_ratio = 0.38
        self.main_pane = None
        self._configure_window_geometry()
        self.log_queue = queue.Queue()
        self.is_processing = False
        self.output_dir_var = tk.StringVar()

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
        self.attachment_bold_var = tk.BooleanVar(value=self.default_params.get('attachment_bold', False))
        self.force_a4_var = tk.BooleanVar(value=self.default_params['force_a4'])
        self.use_custom_english_font_var = tk.BooleanVar(value=self.default_params['use_custom_english_font'])
        self.normalize_punctuation_var = tk.BooleanVar(value=self.default_params['normalize_punctuation'])
        self.convert_auto_numbering_var = tk.BooleanVar(value=self.default_params['convert_auto_numbering'])
        self.enable_table_var = tk.BooleanVar(value=self.default_params['enable_table_formatting'])
        self.table_auto_col_width_var = tk.BooleanVar(value=self.default_params['table_auto_col_width'])
        self.table_header_bold_var = tk.BooleanVar(value=self.default_params['table_header_bold'])
        self.table_smart_align_var = tk.BooleanVar(value=self.default_params['table_smart_align'])
        self.table_unified_borders_var = tk.BooleanVar(value=self.default_params['table_unified_borders'])
        self.table_preserve_borders_var = tk.BooleanVar(value=self.default_params.get('table_preserve_borders', False))
        self.title_bold_var = tk.BooleanVar(value=self.default_params.get('title_bold', False))
        self.h1_bold_var = tk.BooleanVar(value=self.default_params.get('h1_bold', False))
        self.h2_bold_var = tk.BooleanVar(value=self.default_params.get('h2_bold', False))
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_text_var = tk.StringVar(value="")
        self.entries = {}
        self.attachment_option_widgets = []
        self.table_option_widgets = []
        self.spacing_controls = []
        self.indent_controls = []
        self._active_config_canvas = None
        
        self.default_config_path = "default_config.json"
        
        self.create_menu()
        self.create_widgets()
        self.load_initial_config()

        self.master.protocol("WM_DELETE_WINDOW", self._on_close)
        self.master.after(250, self.set_initial_pane_position)
        self.master.after(100, self._check_log_queue)

    def _configure_window_geometry(self):
        screen_width = max(self.master.winfo_screenwidth(), 1024)
        screen_height = max(self.master.winfo_screenheight(), 720)
        width = min(1600, max(1200, screen_width - 80))
        height = min(860, max(640, screen_height - 100))
        min_width = min(1400, max(1200, screen_width - 120))
        min_height = min(720, max(600, screen_height - 160))
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        self.master.geometry(f"{width}x{height}+{x}+{y}")
        self.master.minsize(min_width, min_height)

    def _clamped_left_width(self, total_width, preferred=None):
        total_width = max(int(total_width), 1)
        min_left = self.left_pane_min_width
        min_right = self.right_pane_min_width
        if total_width < min_left + min_right:
            min_left = max(280, int(total_width * 0.42))
            min_right = max(320, total_width - min_left)
        preferred = int(preferred if preferred is not None else total_width * self.preferred_left_ratio)
        return max(min_left, min(preferred, max(min_left, total_width - min_right)))

    def _ensure_pane_widths(self, event=None):
        pane = self.main_pane
        if pane is None:
            return
        try:
            total_width = pane.winfo_width()
            if total_width <= 100:
                return
            current_left = pane.sashpos(0)
            desired_left = self._clamped_left_width(total_width, preferred=current_left)
            if desired_left != current_left:
                pane.sashpos(0, desired_left)
        except tk.TclError:
            pass

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
        pass

    def _set_widgets_enabled(self, widgets, enabled):
        pass

    def _update_attachment_state(self):
        pass

    def _update_table_state(self):
        pass

    def _config_unit_value(self, widget, default):
        value = widget.get().strip()
        return UNIT_DISPLAY_TO_CONFIG.get(value, value or default)

    def _refresh_spacing_controls(self):
        for control in self.spacing_controls:
            unit = self._config_unit_value(control['unit'], 'pt')
            show_multiple = unit == 'multiple'
            control['suffix'].configure(
                text=UNIT_VALUE_SUFFIX.get(unit, UNIT_VALUE_SUFFIX['pt'])
            )
            if show_multiple:
                if not control['multiple'].get().strip():
                    control['multiple'].insert(0, '1.0')
                control['pt'].grid_remove()
                control['multiple'].grid()
            else:
                control['multiple'].grid_remove()
                control['pt'].grid()

    def _refresh_indent_controls(self):
        for control in self.indent_controls:
            unit = self._config_unit_value(control['unit'], 'cm')
            show_chars = unit == 'chars'
            for widget in control['cm_widgets']:
                widget.grid_remove() if show_chars else widget.grid()
            for widget in control['char_widgets']:
                widget.grid() if show_chars else widget.grid_remove()
                
            suffix_text = UNIT_VALUE_SUFFIX.get(unit, UNIT_VALUE_SUFFIX['cm'])
            for label in control['suffix_labels']:
                label.configure(text=suffix_text)

    def _enable_dependent_widgets_for_config_load(self):
        pass

    def _set_widget_value(self, widget, value, is_size=False):
        previous_state = None
        try:
            previous_state = widget.cget('state')
            if previous_state == 'disabled':
                widget.configure(state='normal')
        except tk.TclError:
            previous_state = None

        try:
            if is_size and isinstance(widget, ttk.Combobox):
                display_val = self.font_size_map_rev.get(value, str(value))
                widget.set(display_val)
            elif isinstance(widget, ttk.Combobox):
                display_val = UNIT_CONFIG_TO_DISPLAY.get(str(value).strip(), value)
                widget.set(display_val)
                if display_val != self.font_separator:
                    widget._last_valid_value = display_val
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
        pane = self.main_pane
        if pane is None:
            return
        try:
            total_width = pane.winfo_width()
            if total_width > 100:
                left_width = self._clamped_left_width(total_width)
                pane.sashpos(0, left_width)
        except tk.TclError:
            pass

    def _select_output_dir(self):
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_dir_var.set(os.path.normpath(path))
            self.save_default_config()

    def _open_output_dir(self):
        path = self.output_dir_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择输出文件夹。", parent=self.master)
            return
        if not os.path.exists(path):
            messagebox.showwarning("提示", "该文件夹不存在，请重新选择。", parent=self.master)
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {e}", parent=self.master)

    def create_menu(self):
        menubar = Menu(self.master)
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_help_window)
        help_menu.add_command(label="重置界面布局", command=self.reset_layout)
        menubar.add_cascade(label="帮助", menu=help_menu)
        self.master.config(menu=menubar)

    def reset_layout(self):
        self.set_initial_pane_position()
        self.log_to_debug_window("已重置界面布局。")
        
    def _show_help_tooltip(self, title, message):
        messagebox.showinfo(title, message, parent=self.master)
        
    def _create_help_label(self, parent, text, row, col):
        help_label = ttk.Label(parent, text="(?)", foreground="blue", cursor="hand2")
        help_label.grid(row=row, column=col, sticky='W', padx=(0, 5))
        help_label.bind("<Button-1>", lambda e: self._show_help_tooltip("识别规则说明", text))

    def create_widgets(self):
        main_pane = ttk.PanedWindow(self.master, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.main_pane = main_pane
        main_pane.bind("<Configure>", self._ensure_pane_widths, add="+")

        left_frame = ttk.Frame(main_pane, padding=5, width=420)
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
        
        if TKDND_AVAILABLE and hasattr(self.file_listbox, 'drop_target_register'):
            try:
                self.file_listbox.drop_target_register(DND_FILES)
                self.file_listbox.dnd_bind('<<Drop>>', self.handle_drop)
            except Exception as e:
                self.log_to_debug_window(f"拖拽组件初始化失败，已改为按钮添加文件/文件夹：{e}")
        else:
            self.log_to_debug_window("拖拽添加不可用，已改为按钮添加文件/文件夹。")
        dnd_enabled = TKDND_AVAILABLE and hasattr(self.file_listbox, 'drop_target_register')
        placeholder_text = "可以拖拽文件或文件夹到这里" if dnd_enabled else "请使用下方按钮添加文件或文件夹"
        self.placeholder_label = ttk.Label(self.file_listbox, text=placeholder_text, foreground="grey")
        
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

        output_frame = ttk.LabelFrame(left_frame, text="输出设置")
        output_frame.pack(fill=tk.X, pady=(5, 0))
        
        out_inner_frame = ttk.Frame(output_frame)
        out_inner_frame.pack(fill=tk.X, padx=2, pady=2)
        
        entry_output = ttk.Entry(out_inner_frame, textvariable=self.output_dir_var)
        entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        
        btn_sel = ttk.Button(out_inner_frame, text="选择", command=self._select_output_dir, width=5)
        btn_sel.pack(side=tk.LEFT, padx=(0, 2))
        
        btn_open = ttk.Button(out_inner_frame, text="打开", command=self._open_output_dir, width=5)
        btn_open.pack(side=tk.LEFT)

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

        right_frame = ttk.Frame(main_pane, padding=4, width=660)
        main_pane.add(right_frame, weight=4)

        config_buttons = ttk.Frame(right_frame)
        config_buttons.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))
        ttk.Button(config_buttons, text="加载配置", command=self.load_config).grid(row=0, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(config_buttons, text="保存配置", command=self.save_config).grid(row=0, column=1, sticky='ew', padx=2, pady=2)
        ttk.Button(config_buttons, text="保存为默认", command=self.save_default_config).grid(row=1, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(config_buttons, text="恢复内置默认", command=self.load_defaults).grid(row=1, column=1, sticky='ew', padx=2, pady=2)
        config_buttons.columnconfigure(0, weight=1)
        config_buttons.columnconfigure(1, weight=1)

        config_area = ttk.Frame(right_frame)
        config_area.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        config_canvas = tk.Canvas(config_area, highlightthickness=0)
        config_scrollbar = ttk.Scrollbar(config_area, orient=tk.VERTICAL, command=config_canvas.yview)
        config_canvas.configure(yscrollcommand=config_scrollbar.set)
        
        config_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        config_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        config_frame = ttk.Frame(config_canvas, padding=(6, 4))
        config_window = config_canvas.create_window((0, 0), window=config_frame, anchor='nw')

        col_left_frame = ttk.Frame(config_frame)
        col_right_frame = ttk.Frame(config_frame)
        col_left_frame.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        col_right_frame.grid(row=0, column=1, sticky="nsew")
        config_frame.columnconfigure(0, weight=1)
        config_frame.columnconfigure(1, weight=1)
        for frm in (col_left_frame, col_right_frame):
            frm.columnconfigure(1, weight=1)
            frm.columnconfigure(3, weight=1)

        def refresh_config_scrollregion(_event=None):
            config_canvas.configure(scrollregion=config_canvas.bbox("all"))
        def fit_config_width(event):
            config_canvas.itemconfig(config_window, width=event.width)
        config_frame.bind('<Configure>', refresh_config_scrollregion)
        config_canvas.bind('<Configure>', fit_config_width)

        def create_section_header(parent, title, help_text, cur_row):
            if help_text:
                self._create_help_label(parent, help_text, cur_row, 1)
            lbl = ttk.Label(parent, text=title, font=("", 10, "bold"))
            lbl.grid(row=cur_row, column=0, columnspan=4, sticky="W", pady=(8, 2))
            return cur_row + 1

        def create_entry(parent, label_text, cfg_key, cur_row, col):
            ttk.Label(parent, text=label_text).grid(row=cur_row, column=col, sticky="W", padx=2, pady=1)
            ent = ttk.Entry(parent)
            ent.grid(row=cur_row, column=col + 1, sticky="ew", padx=2, pady=1)
            self.entries[cfg_key] = ent
            return ent

        def create_combo(parent, label_text, cfg_key, values, cur_row, col, readonly=True, width=None):
            ttk.Label(parent, text=label_text).grid(row=cur_row, column=col, sticky="W", padx=2, pady=1)
            state = "readonly" if readonly else "normal"
            cb = ttk.Combobox(parent, values=values, state=state, width=width if width else 22)
            cb.grid(row=cur_row, column=col + 1, sticky="ew", padx=2, pady=1)
            cb._is_unit_combo = False
            self.entries[cfg_key] = cb
            return cb

        def create_font_size_combo(parent, label_text, cfg_key, cur_row, col):
            return create_combo(parent, label_text, cfg_key, list(self.font_size_map.keys()), cur_row, col, readonly=True, width=14)

        def create_spacing_control(parent, label_text, key_pt, key_unit, key_mult, cur_row):
            ttk.Label(parent, text=label_text).grid(row=cur_row, column=0, sticky="W", padx=2, pady=1)
            pt_ent = ttk.Entry(parent, width=8)
            pt_ent.grid(row=cur_row, column=1, sticky="w", padx=2, pady=1)
            unit_cb = ttk.Combobox(parent, values=["磅值", "倍数"], state="readonly", width=10)
            unit_cb.grid(row=cur_row, column=2, sticky="w", padx=2, pady=1)
            suffix_lbl = ttk.Label(parent, text="磅")
            suffix_lbl.grid(row=cur_row, column=3, sticky="w")
            self.entries[key_pt] = pt_ent
            self.entries[key_unit] = unit_cb
            mult_ent = ttk.Entry(parent, width=8)
            self.entries[key_mult] = mult_ent
            control = {
                "pt": pt_ent, "multiple": mult_ent,
                "unit": unit_cb, "suffix": suffix_lbl
            }
            self.spacing_controls.append(control)
            return [pt_ent, mult_ent, unit_cb, suffix_lbl]

        def create_indent_controls(parent, cur_row):
            lbl_chars = ttk.Label(parent, text="缩进单位")
            lbl_chars.grid(row=cur_row, column=0, sticky="W", padx=2, pady=1)
            unit_cb = ttk.Combobox(parent, values=["厘米", "字符"], state="readonly", width=10)
            unit_cb.grid(row=cur_row, column=1, sticky="w", padx=2, pady=1)
            unit_cb._is_unit_combo = True
            self.entries["paragraph_indent_unit"] = unit_cb
            
            # 【新增的这行】：绑定选择事件，触发刷新
            unit_cb.bind("<<ComboboxSelected>>", lambda _event: self._refresh_indent_controls(), add="+")
            
            # 左缩进
            lbl_left = ttk.Label(parent, text="左缩进")
            lbl_left.grid(row=cur_row + 1, column=0, sticky="W", padx=2, pady=1)
            ent_left_cm = ttk.Entry(parent, width=8)
            ent_left_char = ttk.Entry(parent, width=8)
            ent_left_cm.grid(row=cur_row + 1, column=1, sticky="w", padx=2, pady=1)
            ent_left_char.grid(row=cur_row + 1, column=1, sticky="w", padx=2, pady=1)
            self.entries["left_indent_cm"] = ent_left_cm
            self.entries["left_indent_chars"] = ent_left_char
            suffix_left = ttk.Label(parent, text="厘米")
            suffix_left.grid(row=cur_row + 1, column=2, sticky="w")

            # 右缩进
            lbl_right = ttk.Label(parent, text="右缩进")
            lbl_right.grid(row=cur_row + 2, column=0, sticky="W", padx=2, pady=1)
            ent_right_cm = ttk.Entry(parent, width=8)
            ent_right_char = ttk.Entry(parent, width=8)
            ent_right_cm.grid(row=cur_row + 2, column=1, sticky="w", padx=2, pady=1)
            ent_right_char.grid(row=cur_row + 2, column=1, sticky="w", padx=2, pady=1)
            self.entries["right_indent_cm"] = ent_right_cm
            self.entries["right_indent_chars"] = ent_right_char
            suffix_right = ttk.Label(parent, text="厘米")
            suffix_right.grid(row=cur_row + 2, column=2, sticky="w")

            # 首行缩进 (固定字符)
            lbl_first = ttk.Label(parent, text="首行缩进")
            lbl_first.grid(row=cur_row + 3, column=0, sticky="W", padx=2, pady=1)
            ent_first_char = ttk.Entry(parent, width=8)
            ent_first_char.grid(row=cur_row + 3, column=1, sticky="w", padx=2, pady=1)
            self.entries["first_line_indent_chars"] = ent_first_char
            
            suffix_first_chr = ttk.Label(parent, text="字符")
            suffix_first_chr.grid(row=cur_row + 3, column=2, sticky="w")

            ctrl = {
                "unit": unit_cb,
                "cm_widgets": [ent_left_cm, ent_right_cm],
                "char_widgets": [ent_left_char, ent_right_char],
                "suffix_labels": [suffix_left, suffix_right]
            }
            self.indent_controls.append(ctrl)
            return cur_row + 4

        # ========== 左栏 ==========
        rowL = 0
        page_frame = col_left_frame
        title_frame = col_left_frame
        body_frame = col_left_frame

        rowL = create_section_header(page_frame, "页面设置", None, rowL)
        create_entry(page_frame, "上边距(cm)", 'margin_top', rowL, 0)
        create_entry(page_frame, "下边距(cm)", 'margin_bottom', rowL, 2)
        rowL += 1
        create_entry(page_frame, "左边距(cm)", 'margin_left', rowL, 0)
        create_entry(page_frame, "右边距(cm)", 'margin_right', rowL, 2)
        rowL += 1
        create_entry(page_frame, "页脚距(cm)", 'footer_distance', rowL, 0)
        ttk.Checkbutton(page_frame, text="强制设置为A4纸张", variable=self.force_a4_var).grid(row=rowL, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        rowL += 1
        create_combo(page_frame, "页码对齐", 'page_number_align', ['奇偶分页', '居中'], rowL, 0)
        rowL += 1
        create_combo(page_frame, "页码字体", 'page_number_font', self.font_options['page_number'], rowL, 0, readonly=False)
        create_font_size_combo(page_frame, "页码字号", 'page_number_size', rowL, 2)
        rowL += 1

        title_help = "• 主标题: 识别文档开头的连续【居中】且【字体字号相同】的段落。\n• 副标题: 主标题下方，同样【居中】但【字体字号与主标题不同】的段落。\n• TXT文件: 会将首个非层级标题的段落视为题目。"
        rowL = create_section_header(title_frame, "文章标题", title_help, rowL)
        create_combo(title_frame, "题目字体", 'title_font', self.font_options['title'], rowL, 0, readonly=False)
        create_font_size_combo(title_frame, "题目字号", 'title_size', rowL, 2)
        rowL += 1
        create_spacing_control(title_frame, "题目行距", 'title_line_spacing', 'title_line_spacing_unit', 'title_line_spacing_multiple', rowL)
        rowL += 1
        create_combo(title_frame, "副标题字体", 'subtitle_font', self.font_options['subtitle'], rowL, 0, readonly=False)
        create_font_size_combo(title_frame, "副标题字号", 'subtitle_size', rowL, 2)
        rowL += 1
        create_spacing_control(title_frame, "副标题行距", 'subtitle_line_spacing', 'subtitle_line_spacing_unit', 'subtitle_line_spacing_multiple', rowL)
        rowL += 1
        ttk.Checkbutton(title_frame, text="题目加粗", variable=self.title_bold_var).grid(row=rowL, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        rowL += 1

        headings_help = '• 一级标题: "一、", "二、" ...\n• 二级标题: "（一）", "（二）" ...\n• 三级标题: "1.", "2." ...\n• 四级标题: "(1)", "(2)" ...\n\n注：正文、三级、四级标题共用一套字体字号。'
        rowL = create_section_header(title_frame, "层级标题", headings_help, rowL)
        create_combo(title_frame, "一级标题字体", 'h1_font', self.font_options['h1'], rowL, 0, readonly=False)
        create_font_size_combo(title_frame, "一级标题字号", 'h1_size', rowL, 2)
        rowL += 1
        create_combo(title_frame, "二级标题字体", 'h2_font', self.font_options['h2'], rowL, 0, readonly=False)
        create_font_size_combo(title_frame, "二级标题字号", 'h2_size', rowL, 2)
        rowL += 1
        ttk.Checkbutton(title_frame, text="一级标题加粗", variable=self.h1_bold_var).grid(row=rowL, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        ttk.Checkbutton(title_frame, text="二级标题加粗", variable=self.h2_bold_var).grid(row=rowL, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        rowL += 1

        rowL = create_section_header(body_frame, "正文设置", None, rowL)
        create_combo(body_frame, "正文/三四级字体", 'body_font', self.font_options['body'], rowL, 0, readonly=False)
        create_font_size_combo(body_frame, "正文/三四级字号", 'body_size', rowL, 2)
        rowL += 1
        create_spacing_control(body_frame, "正文行距", 'line_spacing', 'line_spacing_unit', 'line_spacing_multiple', rowL)
        rowL += 1

        # 落款设置
        ttk.Separator(body_frame, orient='horizontal').grid(row=rowL, column=0, columnspan=4, sticky='ew', pady=(6, 2))
        rowL += 1
        create_entry(body_frame, "落款段前行数", 'signature_space_lines', rowL, 0)
        create_entry(body_frame, "落款段前磅值", 'signature_space_pt', rowL, 2)
        rowL += 1
        ttk.Label(body_frame, text="说明：磅值优先级更高；都填0，正文直接紧贴落款", foreground="#555555").grid(row=rowL, column=0, columnspan=4, sticky="W", padx=3)
        rowL += 1

        rowL = create_section_header(body_frame, "段落缩进", None, rowL)
        rowL = create_indent_controls(body_frame, rowL)
        rowL += 1

        # ========== 右栏 ==========
        rowR = 0
        table_frame = col_right_frame
        advanced_frame = col_right_frame

        table_help = (
            "• 默认不启用表格自动调整，启用后才会调整表头/内容字体、字号、行距、行高、列宽和边框。\n"
            "• 默认保留单元格原始对齐方式；勾选智能对齐后，表头/序号/短文本居中，数字靠右，长文本靠左。"
        )
        rowR = create_section_header(table_frame, "表格内容", table_help, rowR)
        ttk.Checkbutton(table_frame, text="启用表格自动调整", variable=self.enable_table_var).grid(row=rowR, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        table_auto_col_width_check = ttk.Checkbutton(table_frame, text="自动调整列宽", variable=self.table_auto_col_width_var)
        table_auto_col_width_check.grid(row=rowR, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        rowR += 1
        table_unified_borders_check = ttk.Checkbutton(table_frame, text="统一表格边框", variable=self.table_unified_borders_var)
        table_unified_borders_check.grid(row=rowR, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        table_header_bold_check = ttk.Checkbutton(table_frame, text="表头行加粗", variable=self.table_header_bold_var)
        table_header_bold_check.grid(row=rowR, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        rowR += 1
        
        # 新增：保留原表格边框与底纹
        table_preserve_borders_check = ttk.Checkbutton(table_frame, text="保留原边框与底纹", variable=self.table_preserve_borders_var)
        table_preserve_borders_check.grid(row=rowR, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        table_smart_align_check = ttk.Checkbutton(table_frame, text="智能调整单元格对齐", variable=self.table_smart_align_var)
        table_smart_align_check.grid(row=rowR, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        rowR += 1

        table_header_font_combo = create_combo(table_frame, "表头字体", 'table_header_font', self.font_options['table'], rowR, 0, readonly=False)
        table_font_combo = create_combo(table_frame, "表格字体", 'table_font', self.font_options['table'], rowR, 2, readonly=False)
        rowR += 1
        table_size_combo = create_font_size_combo(table_frame, "表格字号", 'table_size', rowR, 0)
        rowR += 1
        table_spacing_widgets = create_spacing_control(table_frame, "表格行距", 'table_line_spacing', 'table_line_spacing_unit', 'table_line_spacing_multiple', rowR)
        rowR += 1
        table_row_height_entry = create_entry(table_frame, "表格行高(cm)", 'table_row_height_cm', rowR, 0)
        table_width_percent_entry = create_entry(table_frame, "表格宽度(%)", 'table_width_percent', rowR, 2)
        rowR += 1
        table_border_size_entry = create_entry(table_frame, "边框粗细(pt)", 'table_border_size_pt', rowR, 0)
        self.table_option_widgets = []
        rowR += 1

        other_help = '• 图/表标题: 自动查找图片或表格【上方或下方】最近的、居中的、以"图"或"表"开头的段落。'
        rowR = create_section_header(table_frame, "图表标题", other_help, rowR)
        create_combo(table_frame, "表格标题字体", 'table_caption_font', self.font_options['table_caption'], rowR, 0, readonly=False)
        create_font_size_combo(table_frame, "表格标题字号", 'table_caption_size', rowR, 2)
        rowR += 1
        create_combo(table_frame, "图形标题字体", 'figure_caption_font', self.font_options['figure_caption'], rowR, 0, readonly=False)
        create_font_size_combo(table_frame, "图形标题字号", 'figure_caption_size', rowR, 2)
        rowR += 1

        attachment_help = '• 附件标识: 识别"附件1"、"附件："等独立段落。启用后将自动【段前分页】并按主副标题规则识别其自身标题。'
        rowR = create_section_header(advanced_frame, "附件与增强", attachment_help, rowR)
        ttk.Checkbutton(advanced_frame, text="启用附件格式化", variable=self.enable_attachment_var).grid(row=rowR, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        attachment_font_combo = create_combo(advanced_frame, "附件标识字体", 'attachment_font', self.font_options['attachment'], rowR, 2, readonly=False)
        rowR += 1
        attachment_size_combo = create_font_size_combo(advanced_frame, "附件标识字号", 'attachment_size', rowR, 0)
        attachment_bold_check = ttk.Checkbutton(advanced_frame, text="附件标识加粗", variable=self.attachment_bold_var)
        attachment_bold_check.grid(row=rowR, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        self.attachment_option_widgets = []
        rowR += 1

        rowR = create_section_header(advanced_frame, "全局选项", None, rowR)
        if not IS_WINDOWS:
            ttk.Checkbutton(
                advanced_frame, text="识别自动编号并转换为文字",
                variable=self.convert_auto_numbering_var,
            ).grid(row=rowR, column=0, columnspan=4, sticky=tk.W, padx=3, pady=1)
            rowR += 1
        ttk.Checkbutton(advanced_frame, text="自动设置大纲级别", variable=self.set_outline_var).grid(row=rowR, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        ttk.Checkbutton(advanced_frame, text="启用符号标准化", variable=self.normalize_punctuation_var).grid(row=rowR, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        rowR += 1
        ttk.Checkbutton(
            advanced_frame,
            text="自定义数字和字母字体",
            variable=self.use_custom_english_font_var,
        ).grid(row=rowR, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        create_combo(advanced_frame, "数字和字母字体", 'english_font', self.font_options['english'], rowR, 2, readonly=False)
        rowR += 1
        blank_line_combo = create_combo(advanced_frame, "TXT/MD空行处理", 'blank_line_mode', BLANK_LINE_MODE_OPTIONS, rowR, 0, width=34)
        blank_line_combo.grid_configure(columnspan=3)

        self._refresh_spacing_controls()
        self._refresh_indent_controls()

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
        self.attachment_bold_var.set(loaded_config.get('attachment_bold', False))
        self.force_a4_var.set(loaded_config.get('force_a4', False))
        self.use_custom_english_font_var.set(loaded_config.get('use_custom_english_font', False))
        self.normalize_punctuation_var.set(loaded_config.get('normalize_punctuation', False))
        self.convert_auto_numbering_var.set(loaded_config.get('convert_auto_numbering', True))
        self.enable_table_var.set(loaded_config.get('enable_table_formatting', False))
        self.table_auto_col_width_var.set(loaded_config.get('table_auto_col_width', True))
        self.table_header_bold_var.set(loaded_config.get('table_header_bold', True))
        self.table_smart_align_var.set(loaded_config.get('table_smart_align', False))
        self.table_unified_borders_var.set(loaded_config.get('table_unified_borders', True))
        self.table_preserve_borders_var.set(loaded_config.get('table_preserve_borders', False))
        self.title_bold_var.set(loaded_config.get('title_bold', False))
        self.h1_bold_var.set(loaded_config.get('h1_bold', False))
        self.h2_bold_var.set(loaded_config.get('h2_bold', False))
        self.output_dir_var.set(loaded_config.get('output_dir', ''))
        
        boolean_keys = [
            'convert_auto_numbering',
            'set_outline', 'enable_attachment_formatting', 'force_a4',
            'use_custom_english_font', 'use_times_new_roman',
            'remove_blank_lines', 'normalize_punctuation',
            'enable_table_formatting', 'table_auto_col_width', 'table_header_bold',
            'table_smart_align', 'table_unified_borders', 'table_preserve_borders',
            'title_bold', 'h1_bold', 'h2_bold'
        ]
        for key, value in loaded_config.items():
            if key in boolean_keys or key == 'output_dir': continue
            widget = self.entries.get(key)
            if widget:
                self._set_widget_value(widget, value, is_size=("_size" in key))
        self._refresh_spacing_controls()
        self._refresh_indent_controls()

    def load_defaults(self):
        self._apply_config(self.default_params)
    
    def collect_config(self):
        config = {}
        for key, widget in self.entries.items():
            value = widget.get().strip()
            if isinstance(widget, ttk.Combobox) and value == self.font_separator:
                value = getattr(widget, '_last_valid_value', '').strip()
            if isinstance(widget, ttk.Combobox) and getattr(widget, '_is_unit_combo', False):
                value = UNIT_DISPLAY_TO_CONFIG.get(value, value)
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
        config['attachment_bold'] = self.attachment_bold_var.get()
        config['force_a4'] = self.force_a4_var.get()
        config['use_custom_english_font'] = self.use_custom_english_font_var.get()
        config['normalize_punctuation'] = self.normalize_punctuation_var.get()
        config['convert_auto_numbering'] = self.convert_auto_numbering_var.get()
        config['enable_table_formatting'] = self.enable_table_var.get()
        config['table_auto_col_width'] = self.table_auto_col_width_var.get()
        config['table_header_bold'] = self.table_header_bold_var.get()
        config['table_smart_align'] = self.table_smart_align_var.get()
        config['table_unified_borders'] = self.table_unified_borders_var.get()
        config['table_preserve_borders'] = self.table_preserve_borders_var.get()
        config['title_bold'] = self.title_bold_var.get()
        config['h1_bold'] = self.h1_bold_var.get()
        config['h2_bold'] = self.h2_bold_var.get()
        config['output_dir'] = self.output_dir_var.get().strip()
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
        help_content = f"""
{APP_TITLE} v{__version__} - 使用说明

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
- 跨平台说明：Windows 优先使用 WPS/Word 自动编号预处理，不可用或失败时回退到 Python 转换；macOS/Kylin/UOS/Linux 默认识别自动编号并转为文字（包含表格），可在全局选项关闭。Python 转换对不支持的编号保留并记录警告，不自动接受修订，也不新增编号后空白；随后按当前配置排版。非 Windows 或 Windows COM 不可用时，.doc/.wps 仍需 LibreOffice 转为 DOCX；未安装 LibreOffice 时会跳过这些旧格式文件，继续处理 .docx/.txt/.md。

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
        if IS_WINDOWS:
            warning_message = (
                "为了防止数据丢失，请在继续前关闭所有已打开的Word和WPS文档（包括wps、表格、PPT等所有文档）。\n\n"
                "本程序在转换旧格式文件或预处理自动编号时可能需要调用Word/WPS程序，这可能会影响未保存的工作。\n\n"
                "您确定要继续吗？"
            )
        else:
            warning_message = (
                "macOS/Kylin/UOS/Linux 下不会调用 WPS/Word。\n"
                + ("已启用自动编号转文字：包含表格，不支持的编号保留并记录警告。\n\n"
                   if self.convert_auto_numbering_var.get() else "自动编号转文字已关闭，保留原自动编号。\n\n") +
                "如处理 .doc/.wps，程序会尝试使用 LibreOffice；未安装 LibreOffice 时会跳过这些旧格式文件，继续处理 .docx/.txt/.md。\n\n"
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
            output_dir = self.output_dir_var.get().strip()
            if not output_dir:
                messagebox.showwarning("警告", "请先在左侧选择输出文件夹！", parent=self.master)
                return
            if not os.path.exists(output_dir):
                messagebox.showwarning("警告", "指定的输出文件夹不存在，请重新选择！", parent=self.master)
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
        success_count, fail_count, skipped_count = 0, 0, 0
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
            except LegacyConversionUnavailable as e:
                self.log_to_debug_window(f"\n已跳过旧格式文件 {base_name}：\n{e}")
                skipped_count += 1
            except Exception as e:
                logging.error(f"处理文件失败: {input_path}\n{e}", exc_info=True)
                self.log_to_debug_window(f"\n❌ 处理文件 {base_name} 时发生严重错误：\n{e}")
                fail_count += 1
            finally:
                processor._cleanup_temp_files()

        summary_message = f"批量处理完成！\n\n成功: {success_count}个\n跳过: {skipped_count}个\n失败: {fail_count}个"
        if fail_count > 0:
            summary_message += "\n\n失败详情请查看日志窗口。"
        self._set_progress(100, f"完成（成功 {success_count} / 跳过 {skipped_count} / 失败 {fail_count}）")
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


def _create_root():
    if sys.platform == "darwin":
        return tk.Tk(), "macOS 下已禁用拖拽组件，请使用按钮添加文件/文件夹。"
    if TKDND_AVAILABLE:
        previous_default_root = getattr(tk, "_default_root", None)
        try:
            return TkinterDnD.Tk(), None
        except Exception as e:
            failed_root = getattr(tk, "_default_root", None)
            if failed_root is not None and failed_root is not previous_default_root:
                try:
                    failed_root.destroy()
                except Exception:
                    pass
                try:
                    tk._default_root = previous_default_root
                except Exception:
                    pass
            return tk.Tk(), f"拖拽组件启动失败，已回退为普通 Tk 窗口：{e}"
    return tk.Tk(), None


def main():
    root, dnd_error = _create_root()
    app = WordFormatterGUI(root)
    if dnd_error:
        app.log_to_debug_window(dnd_error)
    root.mainloop()


if __name__ == "__main__":
    main()
