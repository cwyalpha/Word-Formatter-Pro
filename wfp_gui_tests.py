"""GUI option visibility and configuration round-trip (requires a display)."""
import json
from pathlib import Path
import tempfile
import tkinter as tk
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import wfp_gui


class NumberingOptionTests(unittest.TestCase):
    def check_platform(self, windows):
        state_directory = tempfile.TemporaryDirectory(prefix='wfp_numbering_gui_')
        self.addCleanup(state_directory.cleanup)
        root=tk.Tk()
        root.withdraw()
        try:
            with patch('wfp_gui.IS_WINDOWS',windows), \
                 patch('wfp_gui.get_gui_state_path', return_value=Path(state_directory.name) / 'gui_state.json'), \
                 patch.object(wfp_gui.WordFormatterGUI,'load_initial_config'):
                gui=wfp_gui.WordFormatterGUI(root)
            gui.load_defaults()
            def widgets(parent):
                for child in parent.winfo_children():
                    yield child
                    yield from widgets(child)
            checks=[w for w in widgets(root) if isinstance(w,wfp_gui.ttk.Checkbutton)
                    and w.cget('text')=='识别自动编号并转换为文字']
            self.assertEqual(len(checks),0 if windows else 1)
            self.assertTrue(gui.convert_auto_numbering_var.get())
            if checks:
                checks[0].invoke()
            else:
                gui.convert_auto_numbering_var.set(False)
            self.assertFalse(gui.collect_config()['convert_auto_numbering'])
            with tempfile.TemporaryDirectory() as tmp:
                path=str(Path(tmp)/'settings.json')
                with patch('wfp_gui.filedialog.asksaveasfilename',return_value=path), \
                     patch('wfp_gui.messagebox.showinfo'):
                    gui.save_config()
                self.assertIs(json.loads(Path(path).read_text(encoding='utf-8'))['convert_auto_numbering'],False)
                gui.load_defaults()
                self.assertTrue(gui.convert_auto_numbering_var.get())
                with patch('wfp_gui.filedialog.askopenfilename',return_value=path), \
                     patch('wfp_gui.messagebox.showinfo'):
                    gui.load_config()
                self.assertFalse(gui.convert_auto_numbering_var.get())
                gui.default_config_path=str(Path(tmp)/'default.json')
                with patch('wfp_gui.messagebox.showinfo'):
                    gui.save_default_config()
                self.assertIs(json.loads(Path(gui.default_config_path).read_text(encoding='utf-8'))['convert_auto_numbering'],False)
            gui._apply_config({})  # old configurations default to enabled
            self.assertTrue(gui.convert_auto_numbering_var.get())
        finally:
            for identifier in root.tk.call('after','info'):
                root.after_cancel(identifier)
            root.destroy()

    def test_non_windows_checkbox_and_config_roundtrip(self):
        self.check_platform(False)

    def test_windows_hides_checkbox(self):
        self.check_platform(True)


class GuiWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix='wfp_gui_workflow_')
        self.addCleanup(self.directory.cleanup)
        self.folder = Path(self.directory.name)
        self.state_path = self.folder / 'gui_state.json'
        self.path_patch = patch('wfp_gui.get_gui_state_path', return_value=self.state_path)
        self.path_patch.start()
        self.addCleanup(self.path_patch.stop)
        self.root = tk.Tk()
        self.root.withdraw()
        self.callback_errors = []
        self.root.report_callback_exception = lambda *args: self.callback_errors.append(args)
        with patch.object(wfp_gui.WordFormatterGUI, 'load_initial_config'):
            self.gui = wfp_gui.WordFormatterGUI(self.root)
        self.gui.default_config_path = str(self.folder / 'default_config.json')
        self.gui.load_defaults()

    def tearDown(self):
        self.root.update_idletasks()
        for identifier in self.root.tk.call('after', 'info'):
            self.root.after_cancel(identifier)
        self.root.destroy()
        self.assertEqual(self.callback_errors, [])

    def set_entry(self, key, value):
        self.gui._set_widget_value(self.gui.entries[key], value)

    def show_small_window(self):
        self.root.minsize(860, 600)
        self.root.geometry('900x640+20+20')
        self.root.deiconify()
        self.root.update()
        self.gui.main_pane.sashpos(0, 360)
        self.root.update()

    def test_tabs_save_all_settings_and_buttons_stay_visible(self):
        gui = self.gui
        notebook = gui.config_notebook
        self.assertEqual([notebook.tab(tab, 'text') for tab in notebook.tabs()],
                         ['页面与正文', '标题与附件', '表格与图表'])
        self.set_entry('body_size', 14.5)
        self.set_entry('title_font', '自定义题目字体')
        self.set_entry('table_line_spacing_multiple', 1.35)
        expected = gui.collect_config()
        self.assertNotIn('output_dir', expected)
        self.show_small_window()
        for tab in notebook.tabs():
            notebook.select(tab)
            self.root.update()
            self.assertTrue(gui.config_buttons.winfo_ismapped())
            bottom = gui.config_buttons.winfo_rooty() + gui.config_buttons.winfo_height()
            self.assertLessEqual(bottom, self.root.winfo_rooty() + self.root.winfo_height())
            self.assertTrue(gui.debug_text.winfo_viewable())
            self.assertGreaterEqual(gui.debug_text.winfo_height(), 40)
            for widget in gui.entries.values():
                if widget.winfo_viewable():
                    self.assertGreaterEqual(widget.winfo_width(), 45)
        self.root.geometry('1200x800')
        self.root.update()
        gui.main_pane.sashpos(0, 440)
        self.root.update()
        for tab in notebook.tabs():
            notebook.select(tab)
            self.root.update()
            self.assertFalse(gui.config_pages[tab]['compact'])
        self.assertEqual(gui.collect_config(), expected)
        path = self.folder / 'format.json'
        with patch('wfp_gui.filedialog.asksaveasfilename', return_value=str(path)), \
             patch('wfp_gui.messagebox.showinfo'):
            gui.save_config()
        gui.load_defaults()
        with patch('wfp_gui.filedialog.askopenfilename', return_value=str(path)), \
             patch('wfp_gui.messagebox.showinfo'):
            gui.load_config()
        self.assertEqual(gui.collect_config(), expected)

    def test_select_and_restart_keep_directory_separate_from_format_defaults(self):
        gui = self.gui
        defaults = Path(gui.default_config_path)
        defaults.write_text('{"body_size": 18}', encoding='utf-8')
        before = defaults.read_bytes()
        output = self.folder / '中文 输出'
        output.mkdir()
        with patch('wfp_gui.filedialog.askdirectory', return_value=str(output)), \
             patch.object(gui, 'save_default_config') as save_defaults:
            self.assertEqual(gui._select_output_dir(), str(output.resolve()))
            save_defaults.assert_not_called()
        self.assertEqual(defaults.read_bytes(), before)
        gui.load_defaults()
        gui._apply_config({'output_dir': 'not a formatting setting'})
        self.assertEqual(gui.output_dir_var.get(), str(output.resolve()))
        with patch('wfp_gui.messagebox.showinfo'):
            gui.save_default_config()
        self.assertNotIn('output_dir', json.loads(defaults.read_text(encoding='utf-8')))
        self.root.update_idletasks()
        for identifier in self.root.tk.call('after', 'info'):
            self.root.after_cancel(identifier)
        self.root.destroy()
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.report_callback_exception = lambda *args: self.callback_errors.append(args)
        with patch.object(wfp_gui.WordFormatterGUI, 'load_initial_config'):
            self.gui = wfp_gui.WordFormatterGUI(self.root)
        self.assertEqual(self.gui.output_dir_var.get(), str(output.resolve()))

    def test_manual_entry_cancel_and_invalid_paths(self):
        gui = self.gui
        gui.output_dir_var.set(str(self.folder))
        self.assertEqual(gui._commit_output_dir(SimpleNamespace(keysym='Return')), 'break')
        self.assertEqual(json.loads(self.state_path.read_text(encoding='utf-8'))['output_dir'], str(self.folder))
        with patch('wfp_gui.filedialog.askdirectory', return_value=''):
            self.assertIsNone(gui._select_output_dir())
        self.assertEqual(gui.output_dir_var.get(), str(self.folder))
        original = self.state_path.read_bytes()
        for bad in (self.folder / 'missing', self.state_path):
            gui.output_dir_var.set(str(bad))
            gui._commit_output_dir()
            with patch('wfp_gui.messagebox.showwarning') as warning:
                self.assertIsNone(gui._remember_output_dir(show_error=True))
                warning.assert_called_once()
            self.assertEqual(self.state_path.read_bytes(), original)

    def test_failed_memory_save_keeps_session_directory(self):
        with patch('wfp_gui.save_gui_state', side_effect=OSError('read only')), \
             patch.object(self.gui, 'log_to_debug_window') as log:
            self.assertEqual(self.gui._remember_output_dir(str(self.folder)), str(self.folder))
            self.assertEqual(self.gui.output_dir_var.get(), str(self.folder))
            log.assert_called_once()

    def test_direct_text_save_dialog_uses_and_updates_shared_directory(self):
        gui = self.gui
        gui._remember_output_dir(str(self.folder))
        other = self.folder / 'other'
        other.mkdir()
        gui.notebook.select(1)
        gui.direct_text_input.insert('1.0', '测试文本')
        with patch('wfp_gui.messagebox.askokcancel', return_value=True), \
             patch('wfp_gui.filedialog.asksaveasfilename', return_value='') as save:
            gui.start_processing()
            self.assertEqual(save.call_args.kwargs['initialdir'], str(self.folder))
            self.assertEqual(gui.output_dir_var.get(), str(self.folder))
        with patch('wfp_gui.messagebox.askokcancel', return_value=True), \
             patch('wfp_gui.filedialog.asksaveasfilename', return_value=str(other / 'test.docx')) as save, \
             patch('wfp_gui.threading.Thread') as worker:
            gui.start_processing()
            self.assertEqual(save.call_args.kwargs['initialdir'], str(self.folder))
            worker.return_value.start.assert_called_once()
        self.assertEqual(gui.output_dir_var.get(), str(other))

    def test_batch_reuses_directory_and_first_use_prompts_once(self):
        gui = self.gui
        gui.file_listbox.insert(tk.END, str(self.folder / 'source.docx'))
        for first_use in (True, False):
            gui.is_processing = False
            gui.output_dir_var.set('' if first_use else str(self.folder))
            with patch('wfp_gui.messagebox.askokcancel', return_value=True), \
                 patch('wfp_gui.filedialog.askdirectory', return_value=str(self.folder)) as choose, \
                 patch('wfp_gui.threading.Thread') as worker:
                gui.start_processing()
                self.assertEqual(choose.call_count, 1 if first_use else 0)
                worker.return_value.start.assert_called_once()

    def test_open_directory_failure_is_shown(self):
        self.gui._remember_output_dir(str(self.folder))
        with patch('wfp_gui.open_output_directory', side_effect=OSError('desktop unavailable')), \
             patch('wfp_gui.threading.Thread') as thread, \
             patch('wfp_gui.messagebox.showerror') as error:
            thread.side_effect = lambda target, daemon: SimpleNamespace(start=target)
            self.gui._open_output_dir()
            self.root.update()
            self.assertIn('desktop unavailable', error.call_args.args[1])

    def test_units_custom_sizes_and_enable_controls_survive_tabs(self):
        gui = self.gui
        self.show_small_window()
        control = next(c for c in gui.spacing_controls if c['unit'] == gui.entries['line_spacing_unit'])
        control['unit'].set('倍数')
        control['unit'].event_generate('<<ComboboxSelected>>')
        self.root.update()
        self.assertTrue(control['multiple'].grid_info())
        self.assertFalse(control['pt'].grid_info())
        self.assertEqual(gui.collect_config()['line_spacing_unit'], 'multiple')
        gui.entries['paragraph_indent_unit'].set('字符')
        gui.entries['paragraph_indent_unit'].event_generate('<<ComboboxSelected>>')
        self.root.update()
        self.assertTrue(gui.entries['left_indent_chars'].grid_info())
        self.assertFalse(gui.entries['left_indent_cm'].grid_info())
        self.assertTrue(gui.entries['first_line_indent_chars'].grid_info())
        self.set_entry('body_size', 17.5)
        self.assertEqual(gui.collect_config()['body_size'], 17.5)
        self.assertEqual(str(gui.entries['table_font'].cget('state')), 'disabled')
        gui.enable_table_var.set(True)
        gui._update_table_state()
        self.assertEqual(str(gui.entries['table_font'].cget('state')), 'normal')
        gui.enable_attachment_var.set(False)
        gui._update_attachment_state()
        self.assertEqual(str(gui.entries['attachment_font'].cget('state')), 'disabled')
        gui.use_custom_english_font_var.set(True)
        gui._update_english_font_state()
        self.assertEqual(str(gui.entries['english_font'].cget('state')), 'normal')

    def test_format_modes_migrate_legacy_config_and_roundtrip_all_tabs(self):
        gui = self.gui
        gui._apply_config({'title_bold': False, 'h1_bold': True, 'h2_bold': False,
                           'table_header_bold': False, 'table_smart_align': True,
                           'table_unified_borders': False})
        self.assertEqual(gui.entries['title_bold'].get(), '保留原样')
        self.assertEqual(gui.entries['h1_bold'].get(), '加粗')
        self.assertEqual(gui.entries['table_cell_alignment'].get(), '智能对齐')
        self.assertEqual(gui.entries['table_border_mode'].get(), '保留原边框')
        for key in ('title_bold', 'subtitle_bold', 'h1_bold', 'h2_bold',
                    'attachment_bold', 'table_header_bold'):
            self.set_entry(key, 'normal')
        self.set_entry('table_alignment', 'center')
        self.set_entry('table_cell_alignment', 'right')
        gui.table_remove_wrapping_var.set(True)
        expected = gui.collect_config()
        self.assertEqual(expected['h2_bold'], 'normal')
        self.assertEqual(expected['table_alignment'], 'center')
        self.assertEqual(expected['table_cell_alignment'], 'right')
        self.assertNotIn('table_smart_align', expected)
        self.assertNotIn('table_unified_borders', expected)
        path = self.folder / 'format-modes.json'
        with patch('wfp_gui.filedialog.asksaveasfilename', return_value=str(path)), \
             patch('wfp_gui.messagebox.showinfo'):
            gui.save_config()
        gui.load_defaults()
        self.assertEqual(gui.entries['subtitle_font'].get(), '楷体_GB2312')
        self.assertEqual(gui.entries['subtitle_bold'].get(), '保留原样')
        with patch('wfp_gui.filedialog.askopenfilename', return_value=str(path)), \
             patch('wfp_gui.messagebox.showinfo'):
            gui.load_config()
        self.assertEqual(gui.collect_config(), expected)

    def test_table_mode_widgets_and_border_width_follow_enable_state(self):
        gui = self.gui
        for key in ('table_header_bold', 'table_alignment', 'table_cell_alignment', 'table_border_mode'):
            self.assertEqual(str(gui.entries[key].cget('state')), 'disabled')
        gui.enable_table_var.set(True)
        gui._update_table_state()
        for key in ('table_header_bold', 'table_alignment', 'table_cell_alignment', 'table_border_mode'):
            self.assertEqual(str(gui.entries[key].cget('state')), 'readonly')
        self.set_entry('table_border_mode', 'preserve')
        gui._update_table_state()
        self.assertEqual(str(gui.entries['table_border_size_pt'].cget('state')), 'disabled')
        self.set_entry('table_border_mode', 'unified')
        gui._update_table_state()
        self.assertEqual(str(gui.entries['table_border_size_pt'].cget('state')), 'normal')
        gui.enable_attachment_var.set(False)
        gui._update_attachment_state()
        self.assertEqual(str(gui.entries['attachment_bold'].cget('state')), 'disabled')
        gui.enable_attachment_var.set(True)
        gui._update_attachment_state()
        self.assertEqual(str(gui.entries['attachment_bold'].cget('state')), 'readonly')

    def test_signature_defaults_controls_and_config_roundtrip(self):
        gui = self.gui
        self.assertFalse(gui.enable_signature_var.get())
        self.assertFalse(gui.signature_include_attachments_var.get())
        self.assertFalse(gui.signature_normalize_date_var.get())
        self.assertEqual(str(gui.entries['signature_layout'].cget('state')), 'disabled')
        gui.enable_signature_var.set(True)
        gui._update_signature_state()
        self.assertEqual(str(gui.entries['signature_layout'].cget('state')), 'readonly')
        self.assertEqual(str(gui.entries['signature_before_spacing'].cget('state')), 'disabled')
        self.set_entry('signature_layout', 'unsealed')
        self.set_entry('signature_before_spacing_mode', 'custom')
        gui._update_signature_state()
        self.assertEqual(str(gui.entries['signature_before_spacing'].cget('state')), 'normal')
        self.set_entry('signature_before_spacing', 1.5)
        self.set_entry('signature_before_spacing_unit', 'lines')
        gui.signature_include_attachments_var.set(True)
        gui.signature_normalize_date_var.set(True)
        expected = gui.collect_config()
        path = self.folder / 'signature-options.json'
        with patch('wfp_gui.filedialog.asksaveasfilename', return_value=str(path)), \
             patch('wfp_gui.messagebox.showinfo'):
            gui.save_config()
        gui.load_defaults()
        self.assertFalse(gui.enable_signature_var.get())
        with patch('wfp_gui.filedialog.askopenfilename', return_value=str(path)), \
             patch('wfp_gui.messagebox.showinfo'):
            gui.load_config()
        self.assertEqual(gui.collect_config(), expected)
        with patch('wfp_gui.messagebox.showinfo'):
            gui.save_default_config()
        self.assertEqual(json.loads(Path(gui.default_config_path).read_text(encoding='utf-8')), expected)
        gui.enable_signature_var.set(False)
        gui._update_signature_state()
        self.assertEqual(str(gui.entries['signature_before_spacing'].cget('state')), 'disabled')
        self.assertEqual(gui.collect_config()['signature_before_spacing'], 1.5)
        gui._apply_config({})
        self.assertFalse(gui.enable_signature_var.get())
        self.assertFalse(gui.signature_include_attachments_var.get())
        self.assertFalse(gui.signature_normalize_date_var.get())

    def test_invalid_signature_spacing_does_not_save_or_start(self):
        gui = self.gui
        self.set_entry('signature_before_spacing', '-1')
        with patch('wfp_gui.messagebox.showerror') as error, \
             patch('wfp_gui.messagebox.askokcancel', return_value=True), \
             patch('wfp_gui.filedialog.asksaveasfilename') as save, \
             patch('wfp_gui.threading.Thread') as worker:
            gui.save_config()
            gui.start_processing()
        self.assertEqual(error.call_count, 2)
        save.assert_not_called()
        worker.assert_not_called()

    def test_title_after_spacing_defaults_units_and_roundtrip(self):
        gui = self.gui
        gui.title_after_mode_var.set('段落间距')
        gui.attachment_after_mode_var.set('段落间距')
        self.assertFalse(gui.enable_title_after_spacing_var.get())
        self.assertFalse(gui.enable_attachment_title_after_spacing_var.get())

        self.assertEqual(str(gui.entries['title_after_spacing'].cget('state')), 'disabled')
        gui.enable_title_after_spacing_var.set(True)
        gui.enable_attachment_title_after_spacing_var.set(True)
        gui._update_title_after_spacing_state()
        self.set_entry('title_after_spacing', 1.5)
        self.set_entry('title_after_spacing_unit', 'lines')
        self.set_entry('attachment_title_after_spacing', 12)
        self.set_entry('attachment_title_after_spacing_unit', 'pt')
        self.assertEqual(gui.entries['title_after_spacing_unit'].get(), '行')
        self.assertEqual(str(gui.entries['title_after_spacing_unit'].cget('state')), 'readonly')
        expected = gui.collect_config()
        path = self.folder / 'title-spacing.json'
        with patch('wfp_gui.filedialog.asksaveasfilename', return_value=str(path)), \
             patch('wfp_gui.messagebox.showinfo'):
            gui.save_config()
        gui.load_defaults()
        self.assertFalse(gui.enable_title_after_spacing_var.get())
        with patch('wfp_gui.filedialog.askopenfilename', return_value=str(path)), \
             patch('wfp_gui.messagebox.showinfo'):
            gui.load_config()
        self.assertEqual(gui.collect_config(), expected)
        gui.enable_attachment_var.set(False)
        gui._update_attachment_state()
        self.assertEqual(str(gui.entries['attachment_title_after_spacing'].cget('state')), 'disabled')
        self.assertEqual(str(gui.entries['title_after_spacing'].cget('state')), 'normal')
        gui.enable_attachment_var.set(True)
        gui._update_attachment_state()
        self.assertEqual(str(gui.entries['attachment_title_after_spacing'].cget('state')), 'normal')
        gui._apply_config({})
        self.assertFalse(gui.enable_title_after_spacing_var.get())
        self.assertFalse(gui.enable_attachment_title_after_spacing_var.get())

    def test_invalid_title_spacing_does_not_overwrite_configs_or_start_work(self):
        gui = self.gui
        gui.title_after_mode_var.set('段落间距')
        defaults = Path(gui.default_config_path)
        defaults.write_text('{"body_size": 18}', encoding='utf-8')
        before = defaults.read_bytes()
        gui.enable_title_after_spacing_var.set(True)
        gui._update_title_after_spacing_state()
        gui.entries['title_after_spacing'].delete(0, tk.END)
        gui.entries['title_after_spacing'].insert(0, '-1')
        with patch('wfp_gui.messagebox.showerror') as error, \
             patch('wfp_gui.filedialog.asksaveasfilename') as choose:
            gui.save_config()
            choose.assert_not_called()
            error.assert_called_once()
        with patch('wfp_gui.messagebox.showerror'):
            gui.save_default_config()
        self.assertEqual(defaults.read_bytes(), before)
        with patch('wfp_gui.messagebox.askokcancel', return_value=True), \
             patch('wfp_gui.messagebox.showerror') as error, patch('wfp_gui.threading.Thread') as worker:
            gui.start_processing()
            error.assert_called_once()
            worker.assert_not_called()

    def test_title_after_blank_line_default_choice_and_roundtrip(self):
        gui = self.gui
        self.assertEqual(gui.title_after_mode_var.get(), '空一行')
        self.assertEqual(gui.attachment_after_mode_var.get(), '空一行')
        self.assertFalse(gui.enable_title_after_spacing_var.get())
        gui.enable_title_after_spacing_var.set(True)
        gui.enable_attachment_title_after_spacing_var.set(True)
        gui._update_title_after_spacing_state()
        self.assertEqual(str(gui.entries['title_after_spacing'].cget('state')), 'disabled')
        config = gui.collect_config()
        self.assertTrue(config['enable_title_after_blank_line'])
        self.assertTrue(config['enable_attachment_title_after_blank_line'])
        self.assertFalse(config['enable_title_after_spacing'])
        path = self.folder / 'blank-line.json'
        with patch('wfp_gui.filedialog.asksaveasfilename', return_value=str(path)), patch('wfp_gui.messagebox.showinfo'):
            gui.save_config()
        gui.load_defaults()
        with patch('wfp_gui.filedialog.askopenfilename', return_value=str(path)), patch('wfp_gui.messagebox.showinfo'):
            gui.load_config()
        self.assertEqual(gui.collect_config(), config)
        self.assertEqual(gui.title_after_mode_var.get(), '空一行')
        gui.title_after_mode_var.set('段落间距')
        gui._update_title_after_spacing_state()
        self.assertEqual(str(gui.entries['title_after_spacing'].cget('state')), 'normal')
        self.assertTrue(gui.collect_config()['enable_title_after_spacing'])
        self.assertFalse(gui.collect_config()['enable_title_after_blank_line'])

    def test_title_after_blank_line_legacy_config_and_attachment_dependency(self):
        gui = self.gui
        gui._apply_config({'enable_title_after_spacing': True, 'title_after_spacing': 20})
        self.assertEqual(gui.title_after_mode_var.get(), '段落间距')
        self.assertTrue(gui.collect_config()['enable_title_after_spacing'])
        gui._apply_config({'enable_attachment_title_after_blank_line': True})
        self.assertEqual(gui.attachment_after_mode_var.get(), '空一行')
        gui.enable_attachment_var.set(False)
        gui._update_attachment_state()
        self.assertEqual(str(gui.attachment_after_mode_widgets[0].cget('state')), 'disabled')
        gui.enable_attachment_var.set(True)
        gui._update_attachment_state()
        self.assertEqual(str(gui.attachment_after_mode_widgets[0].cget('state')), 'readonly')
        gui.load_defaults()
        config = gui.collect_config()
        self.assertFalse(config['enable_title_after_blank_line'])
        self.assertFalse(config['enable_title_after_spacing'])
        self.assertEqual(gui.title_after_mode_var.get(), '空一行')

    def test_wheel_routes_only_to_selected_page_and_keeps_scroll_position(self):
        gui = self.gui
        # Larger desktop fonts make this page scroll at the supported window size.
        wfp_gui.tkfont.Font(root=self.root, name='TkDefaultFont', exists=True).configure(size=14)
        self.show_small_window()
        tabs = gui.config_notebook.tabs()
        canvas = gui.config_pages[tabs[0]]['canvas']
        canvas.yview_moveto(0)
        widget = gui.entries['margin_top']
        for event in (SimpleNamespace(widget=widget, num=5, delta=0),
                      SimpleNamespace(widget=widget, num=None, delta=-120),
                      SimpleNamespace(widget=widget, num=None, delta=-1)):
            canvas.yview_moveto(0)
            before = canvas.yview()[0]
            gui._on_config_mousewheel(event)
            self.assertGreater(canvas.yview()[0], before)
        saved = canvas.yview()
        for excluded in (gui.debug_text, gui.direct_text_input, gui.file_listbox, gui.entries['page_number_align']):
            gui._on_config_mousewheel(SimpleNamespace(widget=excluded, num=5, delta=0))
            self.assertEqual(canvas.yview(), saved)
        gui.config_notebook.select(tabs[1])
        self.root.update()
        gui._on_config_mousewheel(SimpleNamespace(widget=widget, num=5, delta=0))
        gui.config_notebook.select(tabs[0])
        self.root.update()
        self.assertEqual(canvas.yview(), saved)


if __name__=='__main__':
    unittest.main()
