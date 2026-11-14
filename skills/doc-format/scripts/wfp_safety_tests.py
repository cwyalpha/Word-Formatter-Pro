# -*- coding: utf-8 -*-
"""Regressions for text/object preservation and converter ownership."""

import base64
from io import BytesIO
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch
from lxml.etree import tostring

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from wfp_config import DEFAULT_CONFIG
from wfp_core import SofficeConverter, WordProcessor
from wfp_process import converter_profile_uri, external_program_environment, extended_path, owned_temp_dir, remove_owned_temp_dir, run_isolated


def worker_command(*args):
    if getattr(sys, 'frozen', False):
        return [sys.executable, '--test-process-worker'] + list(args)
    return [sys.executable, str(Path(__file__).resolve()), '--worker'] + list(args)


def worker_main(args):
    mode, marker, duration = args[:3]
    if mode == 'ignore-term' and os.name != 'nt':
        signal.signal(signal.SIGTERM, signal.SIG_IGN)
    Path(marker).write_text(str(os.getpid()), encoding='ascii')
    if mode.startswith('parent'):
        child_mode = 'ignore-term' if mode == 'parent-ignore-term' else 'sleep'
        child = subprocess.Popen(worker_command(child_mode, marker + '.child', duration))
        until = time.monotonic() + 10
        while not Path(marker + '.child').exists() and time.monotonic() < until:
            time.sleep(.01)
        if mode == 'parent-exit':
            return 0
        if mode == 'parent-error':
            return 7
    time.sleep(float(duration))
    Path(marker + '.done').write_text('finished', encoding='ascii')
    return 0


def running(pid):
    if os.name == 'nt':
        import ctypes
        from ctypes import wintypes as w
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.OpenProcess.argtypes = [w.DWORD, w.BOOL, w.DWORD]
        kernel.OpenProcess.restype = w.HANDLE
        kernel.GetExitCodeProcess.argtypes = [w.HANDLE, ctypes.POINTER(w.DWORD)]
        kernel.CloseHandle.argtypes = [w.HANDLE]
        handle = kernel.OpenProcess(0x1000, False, pid)
        if not handle:
            return False
        try:
            code = w.DWORD()
            return bool(kernel.GetExitCodeProcess(handle, ctypes.byref(code))) and code.value == 259
        finally:
            kernel.CloseHandle(handle)
    try:
        if sys.platform.startswith('linux'):
            fields = Path('/proc/{}/stat'.format(pid)).read_text().rsplit(')', 1)[1].split()
            return fields[0] != 'Z'
        os.kill(pid, 0)
        return True
    except (ProcessLookupError, FileNotFoundError):
        return False


class ConverterOwnershipTests(unittest.TestCase):
    def test_timeout_kills_descendants_and_preserves_unrelated_process(self):
        with owned_temp_dir(prefix='wfp_process_test_') as directory:
            marker = str(Path(directory) / 'owned.pid')
            sentinel = subprocess.Popen(worker_command('sleep', str(Path(directory) / 'other.pid'), '30'))
            try:
                with self.assertRaises(subprocess.TimeoutExpired):
                    run_isolated(worker_command('parent-ignore-term', marker, '30'), timeout=8, external=False)
                for path in (marker, marker + '.child'):
                    self.assertTrue(Path(path).exists(), path)
                    self.assertFalse(running(int(Path(path).read_text())), path)
                self.assertIsNone(sentinel.poll())
                result = run_isolated(worker_command('sleep', marker + '.next', '.01'), timeout=15, external=False)
                self.assertEqual(result.returncode, 0)
            finally:
                sentinel.terminate()
                sentinel.wait(timeout=10)

    def test_launcher_exit_does_not_leave_background_child(self):
        with owned_temp_dir(prefix='wfp_process_test_') as directory:
            marker = str(Path(directory) / 'early.pid')
            with self.assertRaises(subprocess.TimeoutExpired):
                run_isolated(worker_command('parent-exit', marker, '30'), timeout=8, external=False)
            self.assertFalse(running(int(Path(marker + '.child').read_text())))

    def test_normal_and_error_exit_wait_for_children(self):
        with owned_temp_dir(prefix='wfp_process_test_') as directory:
            for mode, expected in (('parent-exit', 0), ('parent-error', 7)):
                marker = str(Path(directory) / (mode + '.pid'))
                result = run_isolated(worker_command(mode, marker, '.2'), timeout=20, external=False)
                self.assertEqual(result.returncode, expected)
                self.assertTrue(Path(marker + '.child.done').exists())
                self.assertFalse(running(int(Path(marker + '.child').read_text())))

    def test_converter_cleans_failed_and_timed_out_workdirs(self):
        with owned_temp_dir(prefix='wfp_cleanup_test_') as directory:
            source = Path(directory) / '输入 文件.doc'
            source.write_bytes(b'test')
            for kind in ('error', 'timeout', 'success'):
                created = []

                def convert(command, timeout):
                    out = Path(command[command.index('--outdir') + 1])
                    created.append(out.parent)
                    cache = out.parent / 'profile' / ('缓存 ' + 'p' * 45) / ('q' * 100)
                    os.makedirs(extended_path(cache))
                    with open(extended_path(cache / ('r' * 90 + '.bin')), 'wb') as stream:
                        stream.write(b'cache')
                    if kind == 'timeout':
                        raise subprocess.TimeoutExpired(command, timeout)
                    if kind == 'success':
                        Document().save(out / 'converted.docx')
                    return subprocess.CompletedProcess(command, 1 if kind == 'error' else 0, '', '')

                with patch('wfp_core.run_isolated', side_effect=convert):
                    converter = SofficeConverter('test-soffice')
                    if kind == 'success':
                        result, work = converter.convert_to_docx(source)
                        self.assertTrue(result.exists())
                        self.assertTrue(remove_owned_temp_dir(work, strict=True))
                    else:
                        with self.assertRaises(RuntimeError):
                            converter.convert_to_docx(source)
                self.assertTrue(created)
                self.assertFalse(os.path.exists(extended_path(created[0])))

    def test_long_unicode_path_cleanup(self):
        with owned_temp_dir(prefix='wfp_long_path_') as directory:
            nested = Path(directory) / ('中文 空格' * 8) / ('a' * 100) / ('b' * 100)
            os.makedirs(extended_path(nested))
            with open(extended_path(nested / 'cache.bin'), 'wb') as stream:
                stream.write(b'cache')
            self.assertGreater(len(str(nested / 'cache.bin')), 260)
        self.assertFalse(os.path.exists(extended_path(directory)))

    def test_external_environment_preserves_user_paths_and_parent_environment(self):
        bundle = str(Path(tempfile.gettempdir()) / 'test_bundle')
        before = {'PATH': bundle + os.sep + 'bin' + os.pathsep + '/user/bin',
                  'LD_LIBRARY_PATH': bundle + os.pathsep + '/user/lib',
                  'LD_LIBRARY_PATH_ORIG': '/user/lib', 'OTHER_VALUE': 'unchanged'}
        with patch.dict(os.environ, before, clear=True), \
             patch.object(sys, 'frozen', True, create=True), \
             patch.object(sys, '_MEIPASS', bundle, create=True):
            actual = external_program_environment()
            self.assertEqual(dict(os.environ), before)
            self.assertEqual(actual['PATH'], '/user/bin')
            self.assertEqual(actual['LD_LIBRARY_PATH'], '/user/lib')
            self.assertEqual(actual['OTHER_VALUE'], 'unchanged')

    @unittest.skipUnless(os.name == 'nt', 'Windows path rules')
    def test_unc_and_existing_extended_path(self):
        self.assertEqual(extended_path('\\\\server\\share\\folder'), '\\\\?\\UNC\\server\\share\\folder')
        self.assertEqual(extended_path('\\\\?\\C:\\folder'), '\\\\?\\C:\\folder')

    @unittest.skipUnless(os.name == 'nt', 'Windows path rules')
    def test_profile_compatibility_uri_refers_to_same_deep_directory(self):
        from urllib.parse import urlsplit
        from urllib.request import url2pathname
        with owned_temp_dir(prefix='wfp_profile_alias_') as directory:
            original = Path(directory) / ('中文 空格' * 8) / ('p' * 80)
            os.makedirs(extended_path(original))
            uri = converter_profile_uri(original)
            alias = url2pathname(urlsplit(uri).path)
            self.assertEqual(os.stat(extended_path(original)).st_ino,
                             os.stat(extended_path(alias)).st_ino)
            cache = original / ('q' * 120 + '.bin')
            self.assertGreater(len(str(cache)), 260)
            with open(extended_path(cache), 'wb') as stream:
                stream.write(b'long physical path, same directory')
        self.assertFalse(os.path.exists(extended_path(directory)))


class SymbolObjectProtectionTests(unittest.TestCase):
    PIXEL = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=')

    def setUp(self):
        self.processor = WordProcessor(dict(DEFAULT_CONFIG, normalize_punctuation=True))

    def test_same_run_picture_object_breaks_tabs_and_styles_survive(self):
        document = Document()
        paragraph = document.add_paragraph()
        run = paragraph.add_run('第一步:开发工具')
        run.bold = True
        run.add_picture(BytesIO(self.PIXEL))
        run.add_tab()
        run.add_break()
        run.add_text('说明...后续正文')
        embedded = OxmlElement('w:object')
        run._r.append(embedded)
        protected = [(node, tostring(node)) for node in run._r if node.tag != qn('w:t')]
        relationships = dict(document.part.rels)
        self.assertTrue(self.processor._normalize_paragraph_symbols(paragraph))
        self.assertEqual(paragraph.text, '第一步：开发工具\t\n说明……后续正文')
        for node, original in protected:
            self.assertIs(node.getparent(), run._r)
            self.assertEqual(tostring(node), original)
        self.assertEqual(dict(document.part.rels), relationships)
        self.assertFalse(self.processor._normalize_paragraph_symbols(paragraph))

    def test_length_change_across_runs_keeps_surrounding_formatting(self):
        paragraph = Document().add_paragraph()
        first = paragraph.add_run('前文.')
        first.bold = True
        second = paragraph.add_run('..')
        second.italic = True
        tail = paragraph.add_run('后文,继续')
        tail.underline = True
        self.assertTrue(self.processor._normalize_paragraph_symbols(paragraph))
        self.assertEqual(paragraph.text, '前文……后文，继续')
        self.assertTrue(first.bold)
        self.assertTrue(second.italic)
        self.assertTrue(tail.underline)
        self.assertEqual(tail.text, '后文，继续')

    def test_mixed_existing_ellipsis_and_dots_preserve_exact_text_and_styles(self):
        for prefix, suffix in (('中文...', '…'), ('中文...', '……'),
                               ('中文…', '...'), ('中文……', '...')):
            paragraph = Document().add_paragraph()
            first = paragraph.add_run(prefix)
            first.bold = True
            second = paragraph.add_run(suffix)
            second.italic = True
            expected = self.processor._normalize_symbols_in_text(prefix + suffix)
            self.processor._normalize_paragraph_symbols(paragraph)
            self.assertEqual(paragraph.text, expected)
            self.assertEqual(first.text, self.processor._normalize_symbols_in_text(prefix))
            self.assertTrue(first.bold)
            self.assertTrue(second.italic)
            self.assertFalse(self.processor._normalize_paragraph_symbols(paragraph))

    def test_adjacent_mixed_dot_groups_match_original_normalization_rules(self):
        for text in ('中文...…尾文,', '中文...。…', '他说"中文...…"。',
                     '中文......……后文....', '中文．。.…。后文', '中文...…...……尾文'):
            paragraph = Document().add_paragraph(text)
            expected = self.processor._normalize_symbols_in_text(text)
            self.processor._normalize_paragraph_symbols(paragraph)
            self.assertEqual(paragraph.text, expected)

    def test_bookmark_and_hyperlink_boundaries_are_not_collapsed(self):
        from wfp_tests import InlineFormattingProtectionTests
        paragraph = Document().add_paragraph('前文.')
        bookmark = OxmlElement('w:bookmarkStart')
        bookmark.set(qn('w:id'), '1')
        bookmark.set(qn('w:name'), 'original')
        paragraph._p.append(bookmark)
        paragraph.add_run('..后文,')
        link = InlineFormattingProtectionTests.hyperlink(paragraph, '链接:内容')
        relation = link.get(qn('r:id'))
        self.processor._normalize_paragraph_symbols(paragraph)
        self.assertEqual(self.processor._paragraph_text(paragraph), '前文...后文，链接：内容')
        self.assertEqual(link.get(qn('r:id')), relation)
        self.assertIs(bookmark.getparent(), paragraph._p)

    def test_fields_revisions_and_nested_textboxes_are_preserved(self):
        for kind in ('fldSimple', 'fldChar', 'ins'):
            paragraph = Document().add_paragraph('正文,内容')
            node = OxmlElement('w:' + kind)
            paragraph._p.append(node)
            original = paragraph._p.xml
            self.assertFalse(self.processor._normalize_paragraph_symbols(paragraph))
            self.assertEqual(paragraph._p.xml, original)
        paragraph = Document().add_paragraph('正文,内容')
        drawing = OxmlElement('w:pict')
        box = OxmlElement('w:txbxContent')
        inner = OxmlElement('w:p')
        inner_run = OxmlElement('w:r')
        text = OxmlElement('w:t')
        text.text = '框内,原样'
        inner_run.append(text)
        inner.append(inner_run)
        box.append(inner)
        drawing.append(box)
        paragraph.runs[0]._r.append(drawing)
        original = tostring(drawing)
        self.processor._normalize_paragraph_symbols(paragraph)
        self.assertEqual(tostring(drawing), original)

    def test_table_paragraph_and_xml_space(self):
        document = Document()
        paragraph = document.add_table(1, 1).cell(0, 0).paragraphs[0]
        paragraph.add_run(' 中文,内容 ')
        self.processor._normalize_document_symbols(document)
        self.assertEqual(paragraph.text, ' 中文，内容 ')
        self.assertEqual(paragraph.runs[0]._r.find(qn('w:t')).get(qn('xml:space')), 'preserve')


if __name__ == '__main__':
    if sys.argv[1:2] == ['--worker']:
        raise SystemExit(worker_main(sys.argv[2:]))
    unittest.main()
