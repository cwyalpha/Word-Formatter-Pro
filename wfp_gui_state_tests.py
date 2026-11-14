"""GUI preference storage and platform integration without a display."""

import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import Mock, patch

import wfp_gui_state as state


class GuiStateTests(unittest.TestCase):
    def test_user_config_locations(self):
        home = Path(tempfile.gettempdir()) / 'wfp_home'
        local = home / 'local'
        xdg = home / 'xdg'
        self.assertEqual(state.get_gui_state_path('win32', {'LOCALAPPDATA': str(local)}, home),
                         local / 'Word-Formatter-Pro' / 'gui_state.json')
        self.assertEqual(state.get_gui_state_path('win32', {}, home),
                         home / 'AppData' / 'Local' / 'Word-Formatter-Pro' / 'gui_state.json')
        self.assertEqual(state.get_gui_state_path('darwin', {}, home),
                         home / 'Library' / 'Application Support' / 'Word-Formatter-Pro' / 'gui_state.json')
        self.assertEqual(state.get_gui_state_path('linux', {'XDG_CONFIG_HOME': str(xdg)}, home),
                         xdg / 'word-formatter-pro' / 'gui_state.json')
        for env in ({}, {'XDG_CONFIG_HOME': 'relative'}):
            self.assertEqual(state.get_gui_state_path('linux', env, home),
                             home / '.config' / 'word-formatter-pro' / 'gui_state.json')

    def test_roundtrip_and_failed_replace_preserves_previous_state(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'settings' / 'gui_state.json'
            output = str(Path(tmp) / '中文 输出')
            state.save_gui_state(path, output)
            original = path.read_bytes()
            self.assertEqual(state.load_gui_state(path), {'output_dir': output})
            with patch('wfp_gui_state.os.replace', side_effect=OSError('read only')):
                with self.assertRaises(OSError):
                    state.save_gui_state(path, 'new output')
            self.assertEqual(path.read_bytes(), original)
            self.assertEqual(list(path.parent.iterdir()), [path])

    def test_missing_corrupt_and_invalid_settings_are_nonfatal(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'gui_state.json'
            log = Mock()
            self.assertEqual(state.load_gui_state(path, log), {})
            log.assert_not_called()
            for content in ('invalid json', '[]', '{"output_dir": 1}'):
                path.write_text(content, encoding='utf-8')
                self.assertEqual(state.load_gui_state(path, log), {})
            self.assertEqual(log.call_count, 3)
            path.write_text(json.dumps({'output_dir': 'example', 'body_size': 99}), encoding='utf-8')
            self.assertEqual(state.load_gui_state(path), {'output_dir': 'example'})

    def test_directory_validation_handles_spaces_and_rejects_files(self):
        with tempfile.TemporaryDirectory(prefix='wfp path ') as tmp:
            directory = Path(tmp) / '中文 输出'
            directory.mkdir()
            self.assertEqual(state.validate_output_dir(' ' + str(directory) + ' '), str(directory.resolve()))
            file = directory / 'document.docx'
            file.touch()
            for value in ('', str(file), str(directory / 'missing')):
                with self.assertRaises(ValueError):
                    state.validate_output_dir(value)

    def test_open_command_for_each_platform(self):
        with tempfile.TemporaryDirectory(prefix='wfp output ') as tmp:
            directory = str(Path(tmp).resolve())
            with patch.object(os, 'startfile', create=True) as startfile, \
                 patch('wfp_gui_state.subprocess.run') as run:
                state.open_output_directory(directory, 'win32')
                startfile.assert_called_once_with(directory)
                run.assert_not_called()
                for platform, command in (('darwin', 'open'), ('linux', 'xdg-open')):
                    state.open_output_directory(directory, platform)
                    self.assertEqual(run.call_args.args[0], [command, directory])
                    self.assertNotIn('shell', run.call_args.kwargs)

    def test_open_errors_are_reportable(self):
        with tempfile.TemporaryDirectory() as tmp:
            error = subprocess.CalledProcessError(1, ['xdg-open'], stderr='desktop unavailable')
            with patch('wfp_gui_state.subprocess.run', side_effect=error):
                with self.assertRaisesRegex(OSError, 'desktop unavailable'):
                    state.open_output_directory(tmp, 'linux')
            with patch('wfp_gui_state.subprocess.run', side_effect=FileNotFoundError('xdg-open')):
                with self.assertRaises(FileNotFoundError):
                    state.open_output_directory(tmp, 'linux')


if __name__ == '__main__':
    unittest.main()
