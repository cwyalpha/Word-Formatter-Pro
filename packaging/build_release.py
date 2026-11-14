#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build release artifacts for Word Formatter Pro.

Run this script on the target operating system. PyInstaller does not
cross-compile Windows, macOS, and Linux binaries from one host.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import platform
import re
import shutil
import stat
import subprocess
import sys
import textwrap
import urllib.request
import venv


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from wfp_version import APP_BINARY_BASENAME, __version__  # noqa: E402


RELEASE_DIR = ROOT / "release"
BUILD_ROOT = ROOT / "build" / "release"
PYINSTALLER_NAME = "WordFormatterPro"
APPIMAGE_ICON_NAME = "word-formatter-pro"
REUSED_ASSET_URLS = {
    f"{APP_BINARY_BASENAME}.v{__version__}.exe": (
        "https://github.com/cwyalpha/Word-Formatter-Pro/releases/download/"
        f"v2.7.4/{APP_BINARY_BASENAME}.v2.7.4.exe"
    ),
    f"{APP_BINARY_BASENAME}.v{__version__}.Kylin-V10.x86_64.AppImage": (
        "https://github.com/cwyalpha/Word-Formatter-Pro/releases/download/"
        f"v2.7.4/{APP_BINARY_BASENAME}.v2.7.4.Kylin-V10.x86_64.AppImage"
    ),
}


def run(
    cmd: list[str],
    cwd: Path | None = None,
    env: dict[str, str] | None = None,
) -> None:
    printable = " ".join(str(part) for part in cmd)
    print(f"+ {printable}")
    subprocess.run(cmd, cwd=str(cwd or ROOT), check=True, env=env)


def python_stdout(py: Path, code: str) -> str:
    return subprocess.check_output([str(py), "-c", code], text=True).strip()


def module_file(py: Path, module: str) -> Path:
    return Path(
        python_stdout(
            py,
            (
                "import importlib.util, pathlib; "
                f"spec = importlib.util.find_spec({module!r}); "
                "print(pathlib.Path(spec.origin).resolve())"
            ),
        )
    )


def host_key() -> str:
    system = platform.system().lower() or "unknown"
    machine = platform.machine().lower() or "unknown"
    return f"{system}-{machine}"


def venv_python(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def prepare_clean_venv(target: str, reuse: bool = False) -> Path:
    venv_dir = ROOT / f".venv-build-{target}-{host_key()}"
    if venv_dir.exists() and not reuse:
        shutil.rmtree(venv_dir)
    if not venv_dir.exists():
        venv.EnvBuilder(with_pip=True).create(venv_dir)

    py = venv_python(venv_dir)
    run([str(py), "-m", "pip", "install", "--upgrade", "pip"])
    run([str(py), "-m", "pip", "install", "-r", str(ROOT / "requirements.txt")])
    return py


def ensure_tk_available(py: Path) -> None:
    probe = "import tkinter; print(tkinter.TkVersion)"
    try:
        subprocess.run([str(py), "-c", probe], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "").strip()
        raise SystemExit(
            "Tkinter is required for the GUI build, but this Python environment cannot import it. "
            "Install Tk support for the target Python, or run this script with a Python that includes Tk. "
            f"Details: {detail}"
        ) from exc


def windows_tcl_tk_runtime(py: Path) -> dict:
    """Probe the DLLs actually loaded by this Python, not DLL search guesses."""
    probe = textwrap.dedent('''
        import ctypes, json, tkinter
        from ctypes import wintypes as w
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.GetModuleHandleW.argtypes = [w.LPCWSTR]
        kernel.GetModuleHandleW.restype = w.HMODULE
        kernel.GetModuleFileNameW.argtypes = [w.HMODULE, w.LPWSTR, w.DWORD]
        kernel.GetModuleFileNameW.restype = w.DWORD
        root = tkinter.Tk()
        root.withdraw()
        try:
            result = {}
            for name in ('tcl', 'tk'):
                version = root.tk.eval('package provide ' + ('Tcl' if name == 'tcl' else 'Tk'))
                digits = ''.join(version.split('.')[:2])
                module = kernel.GetModuleHandleW(name + digits + 't.dll') or kernel.GetModuleHandleW(name + digits + '.dll')
                buffer = ctypes.create_unicode_buffer(32768)
                if not module or not kernel.GetModuleFileNameW(module, buffer, len(buffer)):
                    raise RuntimeError('Cannot locate loaded ' + name + ' DLL')
                result[name] = {'version': version, 'dll': buffer.value,
                                'scripts': root.tk.eval('info library' if name == 'tcl' else 'set tk_library')}
            print(json.dumps(result))
        finally:
            root.destroy()
    ''')
    runtime = json.loads(python_stdout(py, probe))
    for name, filename, package in (('tcl', 'init.tcl', 'Tcl'), ('tk', 'tk.tcl', 'Tk')):
        item = runtime[name]
        script = Path(item['scripts']) / filename
        if not Path(item['dll']).is_file() or not script.is_file():
            raise SystemExit(f'Missing matching {name} runtime: {item}')
        expected = r'package\s+require\s+-exact\s+' + package + r'\s+' + re.escape(item['version']) + r'(?:\s|$)'
        if not re.search(expected, script.read_text(encoding='utf-8'), re.MULTILINE):
            raise SystemExit(f'{script} does not match loaded {package} {item["version"]}')
    print('Windows loaded Tcl/Tk: ' + json.dumps(runtime, ensure_ascii=False))
    return runtime


def windows_tcl_tk_build_env(py: Path, runtime=None) -> dict[str, str]:
    runtime = runtime or windows_tcl_tk_runtime(py)
    env = os.environ.copy()
    env['TCL_LIBRARY'] = runtime['tcl']['scripts']
    env['TK_LIBRARY'] = runtime['tk']['scripts']
    return env


def windows_build_spec(py: Path, name: str, runtime: dict) -> Path:
    """Replace Analysis entries explicitly: --add-binary alone can be shadowed."""
    spec = BUILD_ROOT / 'windows' / 'spec' / (name + '.spec')
    spec.parent.mkdir(parents=True, exist_ok=True)
    script = f'''
from pathlib import Path
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules
runtime = {runtime!r}
data, binaries, hidden = collect_all('tkinterdnd2')
binaries += [(item['dll'], '.') for item in runtime.values()]
data += collect_data_files('docx')
data += [({str(module_file(py, 'docx.parts'))!r}, 'docx/parts')]
hidden += ['wfp_numbering', 'wfp_signature'] + collect_submodules('fontTools.ttLib')
a = Analysis([{str(ROOT / 'wfp.py')!r}], pathex=[{str(ROOT)!r}], binaries=binaries,
             datas=data, hiddenimports=hidden, hookspath=[], runtime_hooks=[], excludes=[])
dll_names = {{Path(item['dll']).name.lower() for item in runtime.values()}}
a.binaries = [entry for entry in a.binaries if Path(entry[0]).name.lower() not in dll_names]
for kind, item in runtime.items():
    a.binaries.append((Path(item['dll']).name, item['dll'], 'BINARY'))
    prefix = '_' + kind + '_data'
    a.datas = [entry for entry in a.datas if entry[0].replace('\\\\', '/').split('/')[0] != prefix]
    directory = Path(item['scripts'])
    for path in sorted(directory.rglob('*')):
        if path.is_file() and 'demos' not in path.relative_to(directory).parts:
            a.datas.append((prefix + '/' + path.relative_to(directory).as_posix(), str(path), 'DATA'))
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name={name!r}, debug=False,
          bootloader_ignore_signals=False, strip=False, upx=True, console=False,
          disable_windowed_traceback=False)
'''
    spec.write_text(textwrap.dedent(script), encoding='utf-8')
    return spec


def validate_windows_archive(py: Path, binary: Path, runtime: dict) -> None:
    probe = textwrap.dedent('''
        import hashlib, json, sys
        from pathlib import Path
        from PyInstaller.archive.readers import CArchiveReader
        archive = CArchiveReader(sys.argv[1])
        runtime = json.loads(sys.argv[2])
        for kind, item in runtime.items():
            pairs = [(Path(item['dll']).name, Path(item['dll'])),
                     ('_' + kind + '_data/' + ('init.tcl' if kind == 'tcl' else 'tk.tcl'),
                      Path(item['scripts']) / ('init.tcl' if kind == 'tcl' else 'tk.tcl'))]
            for name, original in pairs:
                matches = [key for key in archive.toc if key.replace('\\\\', '/').lower() == name.lower()]
                if len(matches) != 1:
                    raise RuntimeError('Missing/duplicate runtime entry: ' + name)
                if hashlib.sha256(archive.extract(matches[0])).digest() != hashlib.sha256(original.read_bytes()).digest():
                    raise RuntimeError('Frozen runtime does not match loaded Python runtime: ' + name)
        print('Frozen Tcl/Tk DLL and script hashes match the probed runtime.')
    ''')
    run([str(py), '-c', probe, str(binary), json.dumps(runtime)])


def pyinstaller_base(py: Path, target: str) -> list[str]:
    target_build = BUILD_ROOT / target
    return [
        str(py),
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--distpath",
        str(target_build / "dist"),
        "--workpath",
        str(target_build / "work"),
        "--specpath",
        str(target_build / "spec"),
        "--collect-all",
        "tkinterdnd2",
        "--collect-data",
        "docx",
        "--hidden-import",
        "wfp_numbering",
        "--hidden-import",
        "wfp_signature",
        "--collect-submodules",
        "fontTools.ttLib",
        "--add-data",
        f"{module_file(py, 'docx.parts')}{os.pathsep}docx/parts",
    ]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_checksums(paths: list[Path]) -> Path:
    RELEASE_DIR.mkdir(exist_ok=True)
    checksum_file = RELEASE_DIR / f"{APP_BINARY_BASENAME}.v{__version__}.SHA256SUMS.txt"
    paths = [path for path in paths if path.name != checksum_file.name]
    lines = [f"{sha256(path)}  {path.name}" for path in sorted(paths)]
    checksum_file.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"checksums: {checksum_file}")
    return checksum_file


def require_host(expected_system: str, force: bool) -> None:
    if force:
        return
    actual = platform.system()
    if actual != expected_system:
        raise SystemExit(
            f"This target must be built on {expected_system}; current host is {actual}. "
            "Use --force only when you know the environment is compatible."
        )


def copy_python_docx_templates_for_macos(py: Path, app_path: Path) -> None:
    source = Path(
        python_stdout(
            py,
            "from pathlib import Path; import docx; print(Path(docx.__file__).resolve().parent / 'templates')",
        )
    )
    if not source.is_dir():
        raise SystemExit(f"Missing python-docx templates directory: {source}")

    target = app_path / "Contents" / "Frameworks" / "docx" / "templates"
    if target.exists():
        shutil.rmtree(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, target)
    (app_path / "Contents" / "Frameworks" / "docx" / "parts").mkdir(parents=True, exist_ok=True)
    print(f"copied python-docx templates: {source} -> {target}")


def build_macos(args: argparse.Namespace) -> Path:
    require_host("Darwin", args.force)
    py = prepare_clean_venv("macos", reuse=args.reuse_venv)
    ensure_tk_available(py)
    dist_dir = BUILD_ROOT / "macos" / "dist"
    run(
        pyinstaller_base(py, "macos")
        + [
            "--windowed",
            "--name",
            PYINSTALLER_NAME,
            "--osx-bundle-identifier",
            "com.cwyalpha.wordformatterpro",
            str(ROOT / "wfp.py"),
        ]
    )

    app_path = dist_dir / f"{PYINSTALLER_NAME}.app"
    if not app_path.exists():
        raise SystemExit(f"Missing app bundle: {app_path}")
    copy_python_docx_templates_for_macos(py, app_path)

    arch = args.arch or platform.machine()
    artifact = RELEASE_DIR / f"{APP_BINARY_BASENAME}.v{__version__}.macOS-{arch}.app.zip"
    RELEASE_DIR.mkdir(exist_ok=True)
    if artifact.exists():
        artifact.unlink()

    ditto = shutil.which("ditto")
    if ditto:
        run([ditto, "-c", "-k", "--sequesterRsrc", "--keepParent", app_path.name, str(artifact)], cwd=dist_dir)
    else:
        shutil.make_archive(str(artifact.with_suffix("")), "zip", root_dir=dist_dir, base_dir=app_path.name)
    print(f"artifact: {artifact}")
    return artifact


def build_windows(args: argparse.Namespace) -> Path:
    require_host("Windows", args.force)
    py = Path(args.python).resolve() if getattr(args, 'python', None) else prepare_clean_venv("windows", reuse=args.reuse_venv)
    ensure_tk_available(py)
    runtime = windows_tcl_tk_runtime(py)
    build_env = windows_tcl_tk_build_env(py, runtime)
    exe_name = f"{APP_BINARY_BASENAME}.v{__version__}"
    dist_dir = BUILD_ROOT / "windows" / "dist"
    spec = windows_build_spec(py, exe_name, runtime)
    run(
        [str(py), '-m', 'PyInstaller', '--noconfirm', '--clean',
         '--distpath', str(dist_dir), '--workpath', str(BUILD_ROOT / 'windows' / 'work'), str(spec)],
        env=build_env,
    )

    built = dist_dir / f"{exe_name}.exe"
    validate_windows_archive(py, built, runtime)
    run([str(py), str(ROOT / 'packaging' / 'windows_gui_probe.py'), str(built), __version__], env=build_env)
    (BUILD_ROOT / 'windows' / 'tcl-tk-runtime.json').write_text(json.dumps(runtime, indent=2), encoding='utf-8')
    artifact = RELEASE_DIR / f"{exe_name}.exe"
    RELEASE_DIR.mkdir(exist_ok=True)
    shutil.copy2(built, artifact)
    print(f"artifact: {artifact}")
    return artifact


def appimage_svg() -> str:
    return textwrap.dedent(
        """\
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128">
          <rect x="16" y="8" width="96" height="112" rx="10" fill="#2563eb"/>
          <rect x="32" y="30" width="64" height="8" fill="#ffffff"/>
          <rect x="32" y="50" width="64" height="8" fill="#ffffff"/>
          <rect x="32" y="70" width="46" height="8" fill="#ffffff"/>
          <path d="M88 90l16 16 22-30" fill="none" stroke="#16a34a" stroke-width="10" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        """
    )


def prepare_appdir(binary: Path, appdir: Path) -> None:
    if appdir.exists():
        shutil.rmtree(appdir)
    (appdir / "usr" / "bin").mkdir(parents=True)
    (appdir / "usr" / "share" / "applications").mkdir(parents=True)
    (appdir / "usr" / "share" / "icons" / "hicolor" / "scalable" / "apps").mkdir(parents=True)

    target_binary = appdir / "usr" / "bin" / PYINSTALLER_NAME
    shutil.copy2(binary, target_binary)
    target_binary.chmod(target_binary.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    app_run = appdir / "AppRun"
    app_run.write_text(
        "#!/bin/sh\n"
        'HERE="$(dirname "$(readlink -f "$0")")"\n'
        f'exec "$HERE/usr/bin/{PYINSTALLER_NAME}" "$@"\n',
        encoding="utf-8",
    )
    app_run.chmod(app_run.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    desktop = appdir / f"{APPIMAGE_ICON_NAME}.desktop"
    desktop.write_text(
        "[Desktop Entry]\n"
        "Type=Application\n"
        "Name=Word Formatter Pro\n"
        f"Exec={PYINSTALLER_NAME}\n"
        f"Icon={APPIMAGE_ICON_NAME}\n"
        "Categories=Office;WordProcessor;\n"
        "Terminal=false\n",
        encoding="utf-8",
    )
    shutil.copy2(desktop, appdir / "usr" / "share" / "applications" / desktop.name)

    icon = appdir / "usr" / "share" / "icons" / "hicolor" / "scalable" / "apps" / f"{APPIMAGE_ICON_NAME}.svg"
    icon.write_text(appimage_svg(), encoding="utf-8")
    shutil.copy2(icon, appdir / f"{APPIMAGE_ICON_NAME}.svg")
    shutil.copy2(icon, appdir / ".DirIcon")


def build_linux_desktop(
    args: argparse.Namespace,
    *,
    target: str,
    artifact_label: str,
) -> Path:
    require_host("Linux", args.force)
    py = prepare_clean_venv(target, reuse=args.reuse_venv)
    ensure_tk_available(py)
    dist_dir = BUILD_ROOT / target / "dist"
    run(
        pyinstaller_base(py, target)
        + [
            "--onefile",
            "--name",
            PYINSTALLER_NAME,
            str(ROOT / "wfp.py"),
        ]
    )

    built = dist_dir / PYINSTALLER_NAME
    if not built.exists():
        raise SystemExit(f"Missing Linux binary: {built}")

    arch = args.arch or platform.machine()
    appdir = BUILD_ROOT / target / "AppDir"
    prepare_appdir(built, appdir)

    appimagetool = args.appimagetool or shutil.which("appimagetool")
    artifact_base = RELEASE_DIR / f"{APP_BINARY_BASENAME}.v{__version__}.{artifact_label}.{arch}"
    raw_artifact = artifact_base
    if args.no_appimage or args.also_raw:
        RELEASE_DIR.mkdir(exist_ok=True)
        shutil.copy2(built, raw_artifact)
        raw_artifact.chmod(raw_artifact.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
        print(f"artifact: {raw_artifact}")
    if args.no_appimage:
        return raw_artifact
    if not appimagetool:
        raise SystemExit("appimagetool not found. Install it or pass --appimagetool, or use --no-appimage.")

    artifact = artifact_base.with_name(f"{artifact_base.name}.AppImage")
    RELEASE_DIR.mkdir(exist_ok=True)
    env = os.environ.copy()
    env.setdefault("ARCH", arch)
    print(f"+ {appimagetool} {appdir} {artifact}")
    subprocess.run([appimagetool, str(appdir), str(artifact)], cwd=str(ROOT), env=env, check=True)
    artifact.chmod(artifact.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
    print(f"artifact: {artifact}")
    return artifact


def build_kylin(args: argparse.Namespace) -> Path:
    return build_linux_desktop(args, target="kylin", artifact_label="Kylin-V10")


def build_uos(args: argparse.Namespace) -> Path:
    return build_linux_desktop(args, target="uos", artifact_label="UOS-V20")


def download_reused_assets(args: argparse.Namespace) -> list[Path]:
    RELEASE_DIR.mkdir(exist_ok=True)
    artifacts = []
    for filename, url in REUSED_ASSET_URLS.items():
        destination = RELEASE_DIR / filename
        if destination.exists() and not args.overwrite:
            print(f"exists: {destination}")
        else:
            print(f"download: {url}")
            with urllib.request.urlopen(url) as response, destination.open("wb") as fh:
                shutil.copyfileobj(response, fh)
        artifacts.append(destination)
    return artifacts


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build Word Formatter Pro release artifacts.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    def add_build_options(build_parser: argparse.ArgumentParser) -> None:
        build_parser.add_argument("--reuse-venv", action="store_true", help="Reuse the clean build venv if it already exists.")
        build_parser.add_argument("--force", action="store_true", help="Skip host OS checks.")
        build_parser.add_argument("--arch", help="Override artifact architecture label.")

    mac = subparsers.add_parser("macos", help="Build macOS .app.zip on macOS.")
    add_build_options(mac)
    mac.set_defaults(func=build_macos)

    win = subparsers.add_parser("windows", help="Build Windows .exe on Windows.")
    add_build_options(win)
    win.add_argument('--python', help='Use this existing Windows Python without installing or upgrading dependencies.')
    win.set_defaults(func=build_windows)

    def add_linux_build_options(build_parser: argparse.ArgumentParser) -> None:
        add_build_options(build_parser)
        build_parser.add_argument("--appimagetool", help="Path to appimagetool.")
        build_parser.add_argument(
            "--no-appimage",
            action="store_true",
            help="Emit only the PyInstaller Linux binary without wrapping AppImage.",
        )
        build_parser.add_argument(
            "--also-raw",
            action="store_true",
            help="Emit the raw PyInstaller Linux binary in addition to the AppImage.",
        )

    kylin = subparsers.add_parser("kylin", help="Build Kylin/Linux AppImage on Linux/Kylin.")
    add_linux_build_options(kylin)
    kylin.set_defaults(func=build_kylin)

    uos = subparsers.add_parser("uos", help="Build UnionTech UOS AppImage on UOS/Linux.")
    add_linux_build_options(uos)
    uos.set_defaults(func=build_uos)

    reused = subparsers.add_parser("reused-assets", help="Download reusable Windows/Kylin assets and rename them for the current version.")
    reused.add_argument("--overwrite", action="store_true", help="Overwrite existing files in release/.")
    reused.set_defaults(func=download_reused_assets)

    checksums = subparsers.add_parser("checksums", help="Write SHA256SUMS for release artifacts.")
    checksums.set_defaults(func=lambda _args: list(RELEASE_DIR.glob(f"{APP_BINARY_BASENAME}.v{__version__}*")))

    args = parser.parse_args(argv)
    result = args.func(args)
    artifacts = result if isinstance(result, list) else [result]
    artifacts = [path for path in artifacts if isinstance(path, Path) and path.exists()]
    if artifacts and args.command != "checksums":
        write_checksums(list(RELEASE_DIR.glob(f"{APP_BINARY_BASENAME}.v{__version__}*")))
    elif args.command == "checksums":
        write_checksums(artifacts)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
