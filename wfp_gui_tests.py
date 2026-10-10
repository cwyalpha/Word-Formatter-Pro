"""GUI option visibility and configuration round-trip (requires a display)."""
import json
from pathlib import Path
import tempfile
import tkinter as tk
import unittest
from unittest.mock import patch

import wfp_gui


class NumberingOptionTests(unittest.TestCase):
    def check_platform(self, windows):
        root=tk.Tk()
        root.withdraw()
        try:
            with patch('wfp_gui.IS_WINDOWS',windows), \
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


if __name__=='__main__':
    unittest.main()
