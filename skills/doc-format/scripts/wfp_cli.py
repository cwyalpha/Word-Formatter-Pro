# -*- coding: utf-8 -*-
"""Command line entry point for Word Formatter Pro."""

from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path

from wfp_config import DEFAULT_CONFIG, FONT_SIZE_MAP
from wfp_core import (
    BLANK_LINE_MODE_OPTIONS,
    LegacyConversionUnavailable,
    SUPPORTED_FILE_EXTENSIONS,
    WordProcessor,
    WPSAppManager,
    _initialize_com_for_thread,
    _uninitialize_com_for_thread,
)
from wfp_version import __version__


CONFIG_FILE_NAME = "wfp_config.json"
SUPPORTED_EXTENSIONS = set(SUPPORTED_FILE_EXTENSIONS)
FONT_SIZE_NAMES = {value: label.split(" ", 1)[0] for label, value in FONT_SIZE_MAP.items()}
INSTALL_HELP_TEXT = """LibreOffice 安装命令参考：
macOS (Homebrew):
  brew install --cask libreoffice

Debian/Ubuntu:
  sudo apt-get update
  sudo apt-get install libreoffice

Kylin/银河麒麟（apt 系发行版）:
  sudo apt-get update
  sudo apt-get install libreoffice libreoffice-writer

Fedora:
  sudo dnf install libreoffice

Arch Linux:
  sudo pacman -S libreoffice-fresh

安装后请确认 soffice 可用：
  soffice --version

说明：CLI 会在 macOS/Kylin/Linux 处理 .doc/.wps 时自动尝试调用 soffice 转换为 .docx。
如自动查找失败，可在 format 命令中使用 --soffice 指定 soffice 可执行文件路径。
如果 macOS/Kylin/Linux 未安装 LibreOffice，.doc/.wps 会被记录为跳过；.docx/.txt/.md 仍可正常处理。
"""

CONFIG_DESCRIPTIONS = {
    "page_number_align": ("页码对齐", "奇偶分页 / 居中"),
    "footer_distance": ("页脚距离", "厘米"),
    "line_spacing": ("正文行距", "磅"),
    "margin_top": ("上边距", "厘米"),
    "margin_bottom": ("下边距", "厘米"),
    "margin_left": ("左边距", "厘米"),
    "margin_right": ("右边距", "厘米"),
    "title_font": ("题目字体", "例如 方正小标宋简体"),
    "h1_font": ("一级标题字体", "一、二、三、"),
    "h2_font": ("二级标题字体", "（一）（二）"),
    "body_font": ("正文字体", "正文、三四级标题默认字体"),
    "page_number_font": ("页码字体", "页脚页码字体"),
    "table_caption_font": ("表格标题字体", "表 1 等标题"),
    "figure_caption_font": ("图形标题字体", "图 1 等标题"),
    "attachment_font": ("附件标识字体", "附件：等段落"),
    "subtitle_font": ("副标题字体", "题目下方副标题"),
    "title_size": ("题目字号", "pt"),
    "h1_size": ("一级标题字号", "pt"),
    "h2_size": ("二级标题字号", "pt"),
    "body_size": ("正文字号", "pt"),
    "page_number_size": ("页码字号", "pt"),
    "table_caption_size": ("表格标题字号", "pt"),
    "figure_caption_size": ("图形标题字号", "pt"),
    "attachment_size": ("附件标识字号", "pt"),
    "subtitle_size": ("副标题字号", "pt"),
    "title_line_spacing": ("题目行距", "磅"),
    "subtitle_line_spacing": ("副标题行距", "磅"),
    "left_indent_cm": ("段落左缩进", "厘米"),
    "right_indent_cm": ("段落右缩进", "厘米"),
    "set_outline": ("设置大纲级别", "true/false"),
    "enable_attachment_formatting": ("附件格式化", "true/false"),
    "force_a4": ("强制 A4", "true/false"),
    "use_custom_english_font": ("单独设置数字和字母字体", "true/false"),
    "english_font": ("数字和字母字体", "默认 Times New Roman"),
    "blank_line_mode": ("TXT/MD 空行模式", "三选一：不改动、删除单空行、保留单空行"),
    "normalize_punctuation": ("符号标准化", "true/false，保守修复中英文标点混用"),
    "enable_table_formatting": ("表格内容自动调整", "true/false"),
    "table_header_font": ("表头字体", "enable_table_formatting 为 true 时生效"),
    "table_font": ("表格正文字体", "enable_table_formatting 为 true 时生效"),
    "table_size": ("表格字号", "pt"),
    "table_line_spacing": ("表格行距", "磅"),
    "table_row_height_cm": ("表格行高", "厘米"),
    "table_auto_col_width": ("自动调整列宽", "true/false"),
    "table_width_percent": ("表格宽度", "百分比"),
    "table_header_bold": ("表头加粗", "true/false"),
    "table_smart_align": ("表格智能对齐", "true/false"),
    "table_unified_borders": ("统一表格边框", "true/false"),
    "table_border_size_pt": ("表格边框粗细", "pt"),
}


@dataclass
class InputRecord:
    source: Path
    relative: Path


@dataclass
class Job:
    source: Path
    output: Path


def _stderr_log(enabled):
    if not enabled:
        return None

    def log(message):
        print(message, file=sys.stderr)

    return log


def parse_value(raw):
    text = raw.strip()
    lowered = text.lower()
    if lowered in {"true", "yes", "on"}:
        return True
    if lowered in {"false", "no", "off"}:
        return False
    if lowered in {"null", "none"}:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return raw


def load_json_file(path):
    with Path(path).expanduser().open("r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError(f"配置文件必须是 JSON 对象: {path}")
    return data


def normalize_config(config):
    config = dict(config)
    if "blank_line_mode" not in config and "remove_blank_lines" in config:
        config["blank_line_mode"] = WordProcessor._normalize_blank_line_mode(
            None,
            remove_blank_lines=config.get("remove_blank_lines", True),
        )
    if "use_custom_english_font" not in config and config.get("use_times_new_roman"):
        config["use_custom_english_font"] = True
        config.setdefault("english_font", "Times New Roman")

    merged = dict(DEFAULT_CONFIG)
    merged.update(config)
    for key, default_value in DEFAULT_CONFIG.items():
        value = merged.get(key)
        if value is None or (isinstance(value, str) and not value.strip()):
            merged[key] = default_value
    return merged


def load_config(config_path=None, config_json=None):
    source = "内置默认配置"
    config = dict(DEFAULT_CONFIG)

    if config_path:
        path = Path(config_path).expanduser().resolve()
        config.update(load_json_file(path))
        source = str(path)
    else:
        cwd_config = Path.cwd() / CONFIG_FILE_NAME
        if cwd_config.exists():
            config.update(load_json_file(cwd_config))
            source = str(cwd_config.resolve())

    if config_json:
        inline_config = json.loads(config_json)
        if not isinstance(inline_config, dict):
            raise ValueError("--config-json 必须是 JSON 对象")
        config.update(inline_config)
        source += " + --config-json"

    return normalize_config(config), source


def apply_set_overrides(config, set_items):
    if not set_items:
        return
    for item in set_items:
        if "=" not in item:
            raise ValueError(f"--set 参数必须使用 key=value 形式: {item}")
        key, raw_value = item.split("=", 1)
        key = key.strip()
        if key not in DEFAULT_CONFIG:
            raise KeyError(f"未知配置项: {key}")
        config[key] = parse_value(raw_value)


def apply_convenience_overrides(config, args):
    if getattr(args, "enable_table_formatting", False):
        config["enable_table_formatting"] = True
    if getattr(args, "disable_table_formatting", False):
        config["enable_table_formatting"] = False
    if getattr(args, "enable_custom_english_font", False):
        config["use_custom_english_font"] = True
    if getattr(args, "disable_custom_english_font", False):
        config["use_custom_english_font"] = False
    if getattr(args, "english_font", None):
        config["english_font"] = args.english_font
        config["use_custom_english_font"] = True
    if getattr(args, "normalize_punctuation", False):
        config["normalize_punctuation"] = True
    if getattr(args, "disable_normalize_punctuation", False):
        config["normalize_punctuation"] = False
    if getattr(args, "blank_line_mode", None):
        config["blank_line_mode"] = args.blank_line_mode


def load_config_with_overrides(args):
    config, source = load_config(getattr(args, "config", None), getattr(args, "config_json", None))
    apply_set_overrides(config, getattr(args, "set", None))
    apply_convenience_overrides(config, args)
    return normalize_config(config), source


def is_supported_file(path):
    path = Path(path)
    return path.is_file() and not path.name.startswith("~") and path.suffix.lower() in SUPPORTED_EXTENSIONS


def collect_records(paths, recursive=True):
    if not paths:
        raise ValueError("请提供至少一个输入文件或目录。")

    resolved_inputs = [Path(path).expanduser().resolve() for path in paths]
    include_root_prefix = len(resolved_inputs) > 1
    records = []

    for input_path in resolved_inputs:
        if input_path.is_file():
            if not is_supported_file(input_path):
                raise ValueError(f"不支持的文件格式: {input_path}")
            records.append(InputRecord(input_path, Path(input_path.name)))
            continue

        if input_path.is_dir():
            iterator = input_path.rglob("*") if recursive else input_path.glob("*")
            for file_path in sorted(iterator):
                if not is_supported_file(file_path):
                    continue
                rel = file_path.relative_to(input_path)
                if include_root_prefix:
                    rel = Path(input_path.name) / rel
                records.append(InputRecord(file_path, rel))
            continue

        raise FileNotFoundError(f"输入路径不存在: {input_path}")

    if not records:
        raise FileNotFoundError("未找到可处理的文件。支持格式: doc, docx, wps, txt, md")
    return records


def formatted_relative_path(relative):
    return relative.with_name(f"{relative.stem}_formatted.docx")


def unique_path(path, seen):
    candidate = path
    counter = 2
    while candidate in seen:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        counter += 1
    seen.add(candidate)
    return candidate


def build_jobs(input_paths, output_arg, recursive=True):
    records = collect_records(input_paths, recursive=recursive)
    output_path = Path(output_arg).expanduser().resolve() if output_arg else None

    if len(records) == 1 and records[0].source.is_file():
        source = records[0].source
        if output_path is None:
            output = source.with_name(f"{source.stem}_formatted.docx")
        elif output_path.suffix.lower() == ".docx":
            output = output_path
        else:
            output = output_path / f"{source.stem}_formatted.docx"
        return [Job(source, output)]

    if output_path is None:
        first = Path(input_paths[0]).expanduser().resolve()
        if len(input_paths) == 1 and first.is_dir():
            output_dir = first.parent / f"{first.name}_formatted"
        else:
            output_dir = Path.cwd() / "wfp_formatted"
    else:
        if output_path.suffix.lower() == ".docx":
            raise ValueError("多文件或目录输入时，--output 必须是输出目录，不能是单个 .docx 文件。")
        output_dir = output_path

    jobs = []
    seen = set()
    for record in records:
        rel_output = unique_path(formatted_relative_path(record.relative), seen)
        jobs.append(Job(record.source, output_dir / rel_output))
    return jobs


def format_paths(args):
    input_paths = []
    if args.inputs:
        input_paths.extend(args.inputs)
    input_paths.extend(args.paths or [])

    config, config_source = load_config_with_overrides(args)
    log = _stderr_log(args.verbose)
    if log:
        log(f"使用配置: {config_source}")

    try:
        jobs = build_jobs(input_paths, args.output, recursive=not args.no_recursive)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1

    com_initialized = _initialize_com_for_thread(log)
    failures = []
    skipped = []
    try:
        with WPSAppManager(log) as com_mgr:
            processor = WordProcessor(
                config,
                log,
                com_manager=com_mgr,
                soffice_path=args.soffice,
                soffice_timeout=args.soffice_timeout,
            )
            for index, job in enumerate(jobs, start=1):
                try:
                    if log:
                        log(f"开始处理 {index}/{len(jobs)}: {job.source}")
                    job.output.parent.mkdir(parents=True, exist_ok=True)
                    processor.format_document(str(job.source), str(job.output))
                    print(str(job.output.resolve()))
                except LegacyConversionUnavailable as exc:
                    skipped.append(job.source)
                    print(f"已跳过: {job.source}: {exc}", file=sys.stderr)
                except Exception as exc:  # CLI should continue directory batches.
                    failures.append((job.source, exc))
                    print(f"处理失败: {job.source}: {exc}", file=sys.stderr)
                finally:
                    processor._cleanup_temp_files()
    finally:
        _uninitialize_com_for_thread(com_initialized, log)

    if skipped:
        print(f"已跳过 {len(skipped)} 个旧格式文件。", file=sys.stderr)
    if failures:
        print(f"完成，但有 {len(failures)} 个文件失败。", file=sys.stderr)
        return 1
    return 0


def show_config(args):
    config, source = load_config_with_overrides(args)
    payload = {
        "config_source": source,
        "version": __version__,
        "config_file_auto_load": str((Path.cwd() / CONFIG_FILE_NAME).resolve()),
        "supported_inputs": sorted(SUPPORTED_EXTENSIONS),
        "output": "单文件输出 .docx；多文件或目录输出到目录，目录输入会递归保留原目录结构。",
        "legacy_conversion": ".doc/.wps 在 Windows 优先使用 WPS/Word COM；macOS/Kylin/Linux 或 COM 失败时尝试 LibreOffice soffice；macOS/Kylin/Linux 未安装 LibreOffice 时会跳过旧格式文件。",
        "font_size_names": {str(key): value for key, value in FONT_SIZE_NAMES.items()},
        "blank_line_modes": BLANK_LINE_MODE_OPTIONS,
        "optional_features": [
            "enable_table_formatting=true 启用表格内容自动调整",
            "use_custom_english_font=true 并设置 english_font，可单独指定数字和字母字体，默认 Times New Roman",
            "normalize_punctuation=true 启用符号标准化",
        ],
        "config": {},
    }
    for key, value in config.items():
        name, note = CONFIG_DESCRIPTIONS.get(key, (key, ""))
        payload["config"][key] = {"value": value, "name": name, "note": note}
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def save_config(args):
    config, source = load_config_with_overrides(args)
    output = Path(args.output or CONFIG_FILE_NAME).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    with open(output, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(str(output))
    print(
        "提示：配置已保存。还可按需启用表格内容自动调整、自定义数字和字母字体"
        "（默认 Times New Roman）或符号标准化；修改后可再次运行 format 重新转换。",
        file=sys.stderr,
    )
    if getattr(args, "verbose", False):
        print(f"配置来源: {source}", file=sys.stderr)
    return 0


def install_help(_args):
    print(INSTALL_HELP_TEXT)
    return 0


def run_tests(_args):
    from wfp_tests import main as test_main

    return test_main([])


def add_config_override_args(parser):
    parser.add_argument("--config", help="JSON 配置文件路径；未指定时自动读取当前目录 wfp_config.json")
    parser.add_argument("--config-json", help="内联 JSON 配置对象，会覆盖配置文件中的同名字段")
    parser.add_argument("--set", action="append", help="覆盖单个配置项，格式 key=value，可重复使用")
    parser.add_argument("--enable-table-formatting", action="store_true", help="启用表格内容自动调整")
    parser.add_argument("--disable-table-formatting", action="store_true", help="关闭表格内容自动调整")
    parser.add_argument("--enable-custom-english-font", action="store_true", help="启用数字和字母单独字体")
    parser.add_argument("--disable-custom-english-font", action="store_true", help="关闭数字和字母单独字体")
    parser.add_argument("--english-font", help="数字和字母字体；设置后会自动启用 use_custom_english_font")
    parser.add_argument("--normalize-punctuation", action="store_true", help="启用符号标准化")
    parser.add_argument("--disable-normalize-punctuation", action="store_true", help="关闭符号标准化")


def build_parser():
    parser = argparse.ArgumentParser(
        description="Word Formatter Pro CLI：将 doc/docx/wps/txt/md 按公文排版规则格式化为 docx。"
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    fmt = subparsers.add_parser("format", help="格式化单文件、多文件或目录")
    fmt.add_argument("paths", nargs="*", help="输入文件或目录，可一次传入多个")
    fmt.add_argument("-i", "--input", dest="inputs", action="append", help="输入文件或目录，可重复")
    fmt.add_argument("-o", "--output", help="输出文件或目录；目录输入会保留原目录结构")
    add_config_override_args(fmt)
    fmt.add_argument(
        "--blank-line-mode",
        choices=BLANK_LINE_MODE_OPTIONS,
        help="覆盖 TXT/MD 空行处理模式",
    )
    fmt.add_argument("--no-recursive", action="store_true", help="目录输入时不递归扫描")
    fmt.add_argument("--soffice", help="LibreOffice soffice 可执行文件路径；未指定时自动查找")
    fmt.add_argument("--soffice-timeout", type=int, default=120, help="LibreOffice 单文件转换超时秒数")
    fmt.add_argument("-v", "--verbose", action="store_true", help="输出详细处理日志到 stderr")
    fmt.set_defaults(func=format_paths)

    show = subparsers.add_parser("show-config", help="显示当前配置、默认配置说明和可选增强项")
    add_config_override_args(show)
    show.set_defaults(func=show_config)

    save = subparsers.add_parser("save-config", help="保存配置到当前目录 wfp_config.json 或指定路径")
    add_config_override_args(save)
    save.add_argument("-o", "--output", help="输出 JSON 配置文件路径，默认 ./wfp_config.json")
    save.add_argument("-v", "--verbose", action="store_true", help="输出配置来源到 stderr")
    save.set_defaults(func=save_config)

    subparsers.add_parser("install-help", help="显示 LibreOffice 安装提示").set_defaults(func=install_help)
    subparsers.add_parser("test", help="运行内置单元测试").set_defaults(func=run_tests)

    return parser


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
