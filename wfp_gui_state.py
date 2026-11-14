"""GUI preferences and desktop integration, separate from document settings."""

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile


def get_gui_state_path(platform=None, environ=None, home=None):
    platform = platform or sys.platform
    environ = os.environ if environ is None else environ
    home = Path.home() if home is None else Path(home)
    if platform.startswith('win'):
        base = Path(environ.get('LOCALAPPDATA') or home / 'AppData' / 'Local')
        return base / 'Word-Formatter-Pro' / 'gui_state.json'
    if platform == 'darwin':
        return home / 'Library' / 'Application Support' / 'Word-Formatter-Pro' / 'gui_state.json'
    base = Path(environ.get('XDG_CONFIG_HOME') or home / '.config')
    if not base.is_absolute():
        base = home / '.config'
    return base / 'word-formatter-pro' / 'gui_state.json'


def load_gui_state(path, log=None):
    try:
        with Path(path).open(encoding='utf-8') as stream:
            state = json.load(stream)
        if not isinstance(state, dict) or not isinstance(state.get('output_dir', ''), str):
            raise ValueError('输出目录设置格式无效')
        return {'output_dir': state.get('output_dir', '')}
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as exc:
        if log:
            log(f'读取输出目录记忆失败，将使用空目录：{exc}')
        return {}


def save_gui_state(path, output_dir):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=str(path.parent),
                                         prefix='gui_state_', suffix='.tmp', delete=False) as stream:
            temporary = Path(stream.name)
            json.dump({'output_dir': output_dir}, stream, ensure_ascii=False, indent=2)
            stream.write('\n')
        os.replace(str(temporary), str(path))
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def validate_output_dir(value):
    if not value.strip():
        raise ValueError('请先选择输出文件夹。')
    directory = Path(value.strip()).expanduser().resolve()
    if not directory.is_dir():
        raise ValueError(f'输出目录不存在或不是文件夹：{directory}')
    return str(directory)


def open_output_directory(directory, platform=None):
    directory = validate_output_dir(directory)
    platform = platform or sys.platform
    if platform.startswith('win'):
        os.startfile(directory)
        return
    command = ['open' if platform == 'darwin' else 'xdg-open', directory]
    try:
        subprocess.run(command, check=True, stdout=subprocess.DEVNULL,
                       stderr=subprocess.PIPE, text=True)
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or '').strip() or str(exc)
        raise OSError(f'{command[0]} 打开目录失败：{detail}') from exc
