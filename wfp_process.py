# -*- coding: utf-8 -*-
"""Owned converter processes and temporary directories (no GUI dependencies)."""

import contextlib
import ctypes
import locale
import logging
import math
import os
from pathlib import Path
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time


_LAUNCH_LOCK = threading.RLock()


def windows_compatible_path(path):
    """An existing directory's DOS alias, without moving it or changing NTFS."""
    absolute = os.path.abspath(os.fspath(path))
    if os.name != 'nt':
        return absolute
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.GetShortPathNameW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint]
    kernel.GetShortPathNameW.restype = ctypes.c_uint
    buffer = ctypes.create_unicode_buffer(32768)
    count = kernel.GetShortPathNameW(extended_path(absolute), buffer, len(buffer))
    if not count or count >= len(buffer):
        return absolute
    result = buffer.value
    if result.startswith('\\\\?\\UNC\\'):
        return '\\\\' + result[8:]
    if result.startswith('\\\\?\\'):
        return result[4:]
    return result


def converter_profile_uri(path):
    # LO's native cache writers can abort on long profile paths even when the
    # same filesystem supports them. Cleanup still uses the original long path.
    return Path(windows_compatible_path(path)).as_uri()


def external_program_environment():
    """Do not pass PyInstaller's bundled library paths to system LibreOffice."""
    environment = os.environ.copy()
    bundle = getattr(sys, '_MEIPASS', None)
    if not getattr(sys, 'frozen', False) or not bundle:
        return environment
    if os.name == 'nt':
        for key in ('TEMP', 'TMP', 'TMPDIR'):
            if environment.get(key) and os.path.isdir(extended_path(environment[key])):
                environment[key] = windows_compatible_path(environment[key])
    if sys.platform != 'darwin' and os.name != 'nt':
        if 'LD_LIBRARY_PATH_ORIG' in environment:
            environment['LD_LIBRARY_PATH'] = environment['LD_LIBRARY_PATH_ORIG']
        else:
            environment.pop('LD_LIBRARY_PATH', None)
    root = os.path.normcase(os.path.abspath(bundle))
    for key in ('PATH', 'LD_LIBRARY_PATH', 'DYLD_LIBRARY_PATH'):
        if key not in environment:
            continue
        parts = []
        for part in environment[key].split(os.pathsep):
            absolute = os.path.normcase(os.path.abspath(part)) if part else ''
            if absolute != root and not absolute.startswith(root + os.sep):
                parts.append(part)
        environment[key] = os.pathsep.join(parts)
    return environment


@contextlib.contextmanager
def _external_dll_search(external):
    with _LAUNCH_LOCK:
        if not external or not getattr(sys, 'frozen', False):
            yield
            return
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.GetDllDirectoryW.argtypes = [ctypes.c_uint, ctypes.c_wchar_p]
        kernel.GetDllDirectoryW.restype = ctypes.c_uint
        kernel.SetDllDirectoryW.argtypes = [ctypes.c_wchar_p]
        kernel.SetDllDirectoryW.restype = ctypes.c_int
        size = kernel.GetDllDirectoryW(0, None)
        previous = ctypes.create_unicode_buffer(size + 1)
        kernel.GetDllDirectoryW(len(previous), previous)
        if not kernel.SetDllDirectoryW(None):
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            yield
        finally:
            if not kernel.SetDllDirectoryW(previous.value or None):
                raise ctypes.WinError(ctypes.get_last_error())


def extended_path(path):
    """Use extended Windows paths for internal filesystem operations only."""
    path = os.path.abspath(os.fspath(path))
    if os.name != 'nt' or path.startswith('\\\\?\\'):
        return path
    if path.startswith('\\\\'):
        return '\\\\?\\UNC\\' + path[2:]
    return '\\\\?\\' + path


def remove_owned_temp_dir(path, log=None, strict=False):
    """Remove a directory created by this application, after its users exit."""
    absolute = Path(os.path.abspath(os.fspath(path)))
    if absolute == Path(absolute.anchor):
        raise ValueError('Refusing to remove a filesystem root')
    target = extended_path(absolute)
    for attempt in range(5):
        try:
            shutil.rmtree(target)
            return True
        except FileNotFoundError:
            if not os.path.lexists(target):
                return True
            error = sys.exc_info()[1]
        except OSError as exc:
            error = exc
        if attempt < 4:
            time.sleep(0.1 * (attempt + 1))
    message = '警告：清理临时目录失败 {}: {}'.format(absolute, error)
    (log or logging.warning)(message)
    if strict:
        raise error
    return False


@contextlib.contextmanager
def owned_temp_dir(prefix='wfp_', dir=None):
    path = tempfile.mkdtemp(prefix=prefix, dir=dir)
    try:
        yield path
    finally:
        remove_owned_temp_dir(path, strict=True)


def _group_running(group):
    # Docker's PID 1 may leave killed grandchildren as zombies. They no longer
    # execute or hold profile files, and cannot be reaped by this process.
    if sys.platform.startswith('linux'):
        for entry in Path('/proc').iterdir():
            if not entry.name.isdigit():
                continue
            try:
                fields = (entry / 'stat').read_text().rsplit(')', 1)[1].split()
                if int(fields[2]) == group and fields[0] != 'Z':
                    return True
            except (OSError, ValueError, IndexError):
                continue
        return False
    try:
        os.killpg(group, 0)
        return True
    except ProcessLookupError:
        return False


def _signal_group(group, sig):
    try:
        os.killpg(group, sig)
    except ProcessLookupError:
        pass


def _run_posix(command, timeout, cwd, stdin, stdout, stderr, environment, external):
    process = subprocess.Popen(command, cwd=cwd, stdin=stdin, stdout=stdout,
                               stderr=stderr, start_new_session=True, env=environment)
    deadline = time.monotonic() + timeout
    try:
        while True:
            code = process.poll()
            if code is not None and not _group_running(process.pid):
                return code
            if time.monotonic() >= deadline:
                raise subprocess.TimeoutExpired(command, timeout)
            time.sleep(0.03)
    finally:
        process.poll()
        if _group_running(process.pid):
            _signal_group(process.pid, signal.SIGTERM)
            until = time.monotonic() + 3
            while _group_running(process.pid) and time.monotonic() < until:
                process.poll()
                time.sleep(0.03)
            if _group_running(process.pid):
                _signal_group(process.pid, signal.SIGKILL)
        try:
            process.wait(timeout=3)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=3)
        until = time.monotonic() + 3
        while _group_running(process.pid) and time.monotonic() < until:
            time.sleep(0.03)
        if _group_running(process.pid):
            raise RuntimeError('LibreOffice 进程组未能完全退出: {}'.format(process.pid))


def _run_windows(command, timeout, cwd, stdin, stdout, stderr, environment, external):
    # Use Win7-era APIs through ctypes; pywin32 remains optional for conversion.
    import msvcrt
    from ctypes import wintypes as w

    class STARTUPINFO(ctypes.Structure):
        _fields_ = [('cb', w.DWORD), ('lpReserved', w.LPWSTR),
                    ('lpDesktop', w.LPWSTR), ('lpTitle', w.LPWSTR),
                    ('dwX', w.DWORD), ('dwY', w.DWORD),
                    ('dwXSize', w.DWORD), ('dwYSize', w.DWORD),
                    ('dwXCountChars', w.DWORD), ('dwYCountChars', w.DWORD),
                    ('dwFillAttribute', w.DWORD), ('dwFlags', w.DWORD),
                    ('wShowWindow', w.WORD), ('cbReserved2', w.WORD),
                    ('lpReserved2', ctypes.POINTER(w.BYTE)),
                    ('hStdInput', w.HANDLE), ('hStdOutput', w.HANDLE),
                    ('hStdError', w.HANDLE)]

    class PROCESSINFO(ctypes.Structure):
        _fields_ = [('process', w.HANDLE), ('thread', w.HANDLE),
                    ('pid', w.DWORD), ('tid', w.DWORD)]

    class BASICLIMIT(ctypes.Structure):
        _fields_ = [('process_time', ctypes.c_longlong),
                    ('job_time', ctypes.c_longlong), ('flags', w.DWORD),
                    ('min_working_set', ctypes.c_size_t),
                    ('max_working_set', ctypes.c_size_t),
                    ('active_limit', w.DWORD), ('affinity', ctypes.c_size_t),
                    ('priority', w.DWORD), ('scheduling', w.DWORD)]

    class LIMIT(ctypes.Structure):
        _fields_ = [('basic', BASICLIMIT), ('io', ctypes.c_ulonglong * 6),
                    ('process_memory', ctypes.c_size_t),
                    ('job_memory', ctypes.c_size_t),
                    ('peak_process_memory', ctypes.c_size_t),
                    ('peak_job_memory', ctypes.c_size_t)]

    class ACCOUNTING(ctypes.Structure):
        _fields_ = [('times', ctypes.c_longlong * 4), ('faults', w.DWORD),
                    ('total', w.DWORD), ('active', w.DWORD),
                    ('terminated', w.DWORD)]

    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    signatures = {
        'CreateJobObjectW': ([ctypes.c_void_p, w.LPCWSTR], w.HANDLE),
        'SetInformationJobObject': ([w.HANDLE, ctypes.c_int, ctypes.c_void_p, w.DWORD], w.BOOL),
        'QueryInformationJobObject': ([w.HANDLE, ctypes.c_int, ctypes.c_void_p, w.DWORD, ctypes.c_void_p], w.BOOL),
        'AssignProcessToJobObject': ([w.HANDLE, w.HANDLE], w.BOOL),
        'TerminateJobObject': ([w.HANDLE, w.UINT], w.BOOL),
        'CreateProcessW': ([w.LPCWSTR, w.LPWSTR, ctypes.c_void_p, ctypes.c_void_p, w.BOOL,
                            w.DWORD, ctypes.c_void_p, w.LPCWSTR,
                            ctypes.POINTER(STARTUPINFO), ctypes.POINTER(PROCESSINFO)], w.BOOL),
        'ResumeThread': ([w.HANDLE], w.DWORD),
        'TerminateProcess': ([w.HANDLE, w.UINT], w.BOOL),
        'GetExitCodeProcess': ([w.HANDLE, ctypes.POINTER(w.DWORD)], w.BOOL),
        'WaitForSingleObject': ([w.HANDLE, w.DWORD], w.DWORD),
        'CloseHandle': ([w.HANDLE], w.BOOL),
    }
    for name, (args, result) in signatures.items():
        getattr(kernel, name).argtypes = args
        getattr(kernel, name).restype = result

    def checked(value):
        if not value:
            raise ctypes.WinError(ctypes.get_last_error())
        return value

    job = checked(kernel.CreateJobObjectW(None, None))
    info = PROCESSINFO()
    try:
        limits = LIMIT()
        limits.basic.flags = 0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        checked(kernel.SetInformationJobObject(job, 9, ctypes.byref(limits), ctypes.sizeof(limits)))
        startup = STARTUPINFO()
        startup.cb = ctypes.sizeof(startup)
        startup.dwFlags = 0x100  # STARTF_USESTDHANDLES
        handles = [msvcrt.get_osfhandle(stream.fileno()) for stream in (stdin, stdout, stderr)]
        startup.hStdInput, startup.hStdOutput, startup.hStdError = handles
        commandline = subprocess.list2cmdline([os.fspath(arg) for arg in command])
        environment_block = ctypes.create_unicode_buffer(
            '\0'.join(key + '=' + value for key, value in sorted(environment.items())) + '\0\0') if environment is not None else None
        # The child is suspended until it belongs to our job: no launch race.
        for breakaway in (False, True):
            with _LAUNCH_LOCK:
                try:
                    for handle in handles:
                        os.set_handle_inheritable(handle, True)
                    flags = 0x08000000 | 0x4 | 0x400  # NO_WINDOW | SUSPENDED | UNICODE_ENVIRONMENT
                    if breakaway:
                        flags |= 0x01000000  # retry for Win7 hosts already in a job
                    checked(kernel.CreateProcessW(None, ctypes.create_unicode_buffer(commandline),
                                                   None, None, True, flags, environment_block,
                                                   os.fspath(cwd) if cwd else None,
                                                   ctypes.byref(startup), ctypes.byref(info)))
                finally:
                    for handle in handles:
                        os.set_handle_inheritable(handle, False)
            if kernel.AssignProcessToJobObject(job, info.process):
                break
            error = ctypes.get_last_error()
            kernel.TerminateProcess(info.process, 1)
            kernel.WaitForSingleObject(info.process, 3000)
            kernel.CloseHandle(info.thread)
            kernel.CloseHandle(info.process)
            info = PROCESSINFO()
            if breakaway or error != 5:
                raise RuntimeError('无法隔离 LibreOffice 转换进程: {}'.format(ctypes.WinError(error)))
        if kernel.ResumeThread(info.thread) == 0xffffffff:
            raise ctypes.WinError(ctypes.get_last_error())
        deadline = time.monotonic() + timeout
        accounting = ACCOUNTING()
        while True:
            checked(kernel.QueryInformationJobObject(job, 1, ctypes.byref(accounting),
                                                     ctypes.sizeof(accounting), None))
            if accounting.active == 0:
                code = w.DWORD()
                checked(kernel.GetExitCodeProcess(info.process, ctypes.byref(code)))
                return code.value
            if time.monotonic() >= deadline:
                raise subprocess.TimeoutExpired(command, timeout)
            time.sleep(0.03)
    finally:
        # No other Office/LibreOffice instance can belong to this unnamed job.
        kernel.TerminateJobObject(job, 1)
        if info.process:
            kernel.WaitForSingleObject(info.process, 3000)
        accounting = ACCOUNTING()
        until = time.monotonic() + 3
        while kernel.QueryInformationJobObject(job, 1, ctypes.byref(accounting),
                                               ctypes.sizeof(accounting), None):
            if not accounting.active or time.monotonic() >= until:
                break
            time.sleep(0.03)
        remaining = accounting.active
        for handle in (info.thread, info.process, job):
            if handle:
                kernel.CloseHandle(handle)
        if remaining:
            raise RuntimeError('LibreOffice 转换作业未能完全退出')


def run_isolated(command, timeout=120, cwd=None, external=True):
    """Wait for this invocation AND its descendants; kill only them on timeout."""
    if not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('转换超时必须大于 0')
    encoding = locale.getpreferredencoding(False)
    with open(os.devnull, 'rb') as stdin, tempfile.TemporaryFile() as stdout, tempfile.TemporaryFile() as stderr:
        try:
            runner = _run_windows if os.name == 'nt' else _run_posix
            environment = external_program_environment() if external else None
            if os.name == 'nt':
                # Keep the system search order until the external process tree
                # has exited; delayed loader work must not inherit the bundle.
                with _external_dll_search(external):
                    code = runner(command, timeout, cwd, stdin, stdout, stderr, environment, external)
            else:
                code = runner(command, timeout, cwd, stdin, stdout, stderr, environment, external)
        except subprocess.TimeoutExpired as exc:
            stdout.seek(0)
            stderr.seek(0)
            exc.stdout = stdout.read().decode(encoding, errors='replace')
            exc.stderr = stderr.read().decode(encoding, errors='replace')
            raise
        stdout.seek(0)
        stderr.seek(0)
        return subprocess.CompletedProcess(command, code,
                                          stdout.read().decode(encoding, errors='replace'),
                                          stderr.read().decode(encoding, errors='replace'))
