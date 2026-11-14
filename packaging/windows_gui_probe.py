"""Launch the real frozen GUI and close only windows owned by that invocation."""

import ctypes
from ctypes import wintypes as w
import json
import subprocess
import sys
import time


def main():
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    user = ctypes.WinDLL('user32', use_last_error=True)

    class ENTRY(ctypes.Structure):
        _fields_ = [('size', w.DWORD), ('usage', w.DWORD), ('pid', w.DWORD),
                    ('heap', ctypes.c_size_t), ('module', w.DWORD),
                    ('threads', w.DWORD), ('parent', w.DWORD), ('priority', w.LONG),
                    ('flags', w.DWORD), ('exe', w.WCHAR * 260)]

    kernel.CreateToolhelp32Snapshot.argtypes = [w.DWORD, w.DWORD]
    kernel.CreateToolhelp32Snapshot.restype = w.HANDLE
    kernel.Process32FirstW.argtypes = [w.HANDLE, ctypes.POINTER(ENTRY)]
    kernel.Process32NextW.argtypes = [w.HANDLE, ctypes.POINTER(ENTRY)]
    kernel.CloseHandle.argtypes = [w.HANDLE]
    kernel.OpenProcess.argtypes = [w.DWORD, w.BOOL, w.DWORD]
    kernel.OpenProcess.restype = w.HANDLE
    kernel.TerminateProcess.argtypes = [w.HANDLE, w.UINT]
    user.IsWindow.argtypes = [w.HWND]
    user.IsWindowVisible.argtypes = [w.HWND]
    user.GetWindowThreadProcessId.argtypes = [w.HWND, ctypes.POINTER(w.DWORD)]
    user.GetWindowTextW.argtypes = [w.HWND, w.LPWSTR, ctypes.c_int]
    user.GetClassNameW.argtypes = [w.HWND, w.LPWSTR, ctypes.c_int]
    user.PostMessageW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
    callback_type = ctypes.WINFUNCTYPE(w.BOOL, w.HWND, w.LPARAM)
    user.EnumWindows.argtypes = [callback_type, w.LPARAM]

    def descendants(root):
        snapshot = kernel.CreateToolhelp32Snapshot(2, 0)
        if snapshot == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            parents = {}
            entry = ENTRY()
            entry.size = ctypes.sizeof(entry)
            more = kernel.Process32FirstW(snapshot, ctypes.byref(entry))
            while more:
                parents[entry.pid] = entry.parent
                more = kernel.Process32NextW(snapshot, ctypes.byref(entry))
            owned = {root}
            while True:
                added = {pid for pid, parent in parents.items() if parent in owned}
                if added <= owned:
                    return owned
                owned.update(added)
        finally:
            kernel.CloseHandle(snapshot)

    process = subprocess.Popen([sys.argv[1]], creationflags=subprocess.CREATE_NO_WINDOW)
    owned = {process.pid}
    windows = []
    result = None
    try:
        deadline = time.monotonic() + 45
        while time.monotonic() < deadline:
            owned.update(descendants(process.pid))
            windows = []

            @callback_type
            def visit(hwnd, _):
                pid = w.DWORD()
                user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value in owned and user.IsWindowVisible(hwnd):
                    title, kind = ctypes.create_unicode_buffer(1024), ctypes.create_unicode_buffer(128)
                    user.GetWindowTextW(hwnd, title, len(title))
                    user.GetClassNameW(hwnd, kind, len(kind))
                    windows.append((hwnd, kind.value, title.value))
                return True

            user.EnumWindows(visit, 0)
            ready = [item for item in windows if item[1] == 'TkTopLevel' and sys.argv[2] in item[2]]
            if ready:
                hwnd, kind, title = ready[0]
                time.sleep(1)
                if not user.IsWindow(hwnd) or process.poll() is not None:
                    raise RuntimeError('Frozen GUI exited during initialization')
                user.PostMessageW(hwnd, 0x10, 0, 0)
                code = process.wait(timeout=20)
                if code != 0:
                    raise RuntimeError('Frozen GUI exit code: {}'.format(code))
                result = {'startup': True, 'class': kind, 'title': title, 'exit_code': code}
                break
            if process.poll() is not None:
                raise RuntimeError('Frozen GUI exited: {}'.format(process.returncode))
            if any(kind == '#32770' and 'Unhandled' in title for _, kind, title in windows):
                raise RuntimeError('Frozen executable showed an exception dialog: {}'.format(windows))
            time.sleep(.15)
        if result is None:
            raise RuntimeError('Frozen GUI did not become ready: {}'.format(windows))
        print(json.dumps(result, ensure_ascii=False))
    finally:
        for hwnd, _, _ in windows:
            user.PostMessageW(hwnd, 0x10, 0, 0)
        if process.poll() is None:
            try:
                process.wait(timeout=3)
            except subprocess.TimeoutExpired:
                for pid in owned:
                    handle = kernel.OpenProcess(1, False, pid)
                    if handle:
                        kernel.TerminateProcess(handle, 1)
                        kernel.CloseHandle(handle)
                process.wait(timeout=5)


if __name__ == '__main__':
    main()
