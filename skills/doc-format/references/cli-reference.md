# CLI Reference

## 命令格式

主命令：

```bash
python scripts/wfp_cli.py <子命令> [参数]
```

格式化：

```bash
python scripts/wfp_cli.py format [输入路径...] [-i <输入路径>] [-o <输出路径>] [配置参数] [其他参数]
```

查看配置：

```bash
python scripts/wfp_cli.py show-config
```

保存配置：

```bash
python scripts/wfp_cli.py save-config [配置参数] [-o <配置文件路径>]
```

安装提示：

```bash
python scripts/wfp_cli.py install-help
```

运行测试：

```bash
python scripts/wfp_cli.py test
```

## `format` 参数

| 参数 | 默认值 | 说明 |
|---|---:|---|
| `paths` | 可选 | 位置参数输入路径，可一次传入多个文件或目录 |
| `-i, --input` | 可选 | 输入文件或目录，可重复。与 `paths` 可同时使用 |
| `-o, --output` | 自动生成 | 输出文件或目录。多文件或目录输入时必须是目录 |
| `--config` | 自动读取当前目录 `wfp_config.json` | 指定 JSON 配置文件，会覆盖默认配置中的同名字段 |
| `--config-json` | 无 | 内联 JSON 对象，会覆盖配置文件中的同名字段 |
| `--set key=value` | 无 | 覆盖单个配置项，可重复使用 |
| `--enable-table-formatting` | 关闭 | 启用表格内容自动调整 |
| `--disable-table-formatting` | 关闭 | 关闭表格内容自动调整 |
| `--enable-custom-english-font` | 关闭 | 启用数字和字母单独字体 |
| `--disable-custom-english-font` | 关闭 | 关闭数字和字母单独字体 |
| `--english-font <字体>` | `Times New Roman` | 设置数字和字母字体；使用后自动启用 `use_custom_english_font` |
| `--normalize-punctuation` | 关闭 | 启用符号标准化 |
| `--disable-normalize-punctuation` | 关闭 | 关闭符号标准化 |
| `--blank-line-mode` | 删除单个空行，多个空行保留至1个空行 | 覆盖 TXT/MD 空行处理模式 |
| `--no-recursive` | 关闭 | 目录输入时不递归子目录；默认递归 |
| `--soffice <路径>` | 自动查找 | 指定 LibreOffice `soffice` 路径，用于 `.doc/.wps` 转 `.docx` |
| `--soffice-timeout <秒>` | `120` | LibreOffice 单文件转换超时秒数 |
| `-v, --verbose` | 关闭 | 显示详细日志 |

`--blank-line-mode` 可选值：

- `不改动任何空行`
- `删除单个空行，多个空行保留至1个空行`
- `保留单个空行，多个空行保留至1个空行`

## `show-config`

```bash
python scripts/wfp_cli.py show-config
```

输出 JSON，包含当前配置来源、当前目录自动读取的 `wfp_config.json` 路径、支持的输入类型、字号名称对照、空行模式、常用可选增强项、所有配置项的值、名称和说明。

`show-config` 也支持 `--config`、`--config-json`、`--set` 和便利开关，用来查看覆盖后的配置效果。

## `save-config`

```bash
python scripts/wfp_cli.py save-config
```

- 默认输出到当前目录 `wfp_config.json`。
- 会先读取 `--config` 或当前目录已有 `wfp_config.json`，再应用 `--config-json`、`--set` 和便利开关，最后保存合并后的配置。
- 保存后提醒用户还可启用表格内容自动调整、数字和字母字体、符号标准化。

## `test`

```bash
python scripts/wfp_cli.py test
```

运行内置单元测试，覆盖文本处理、Markdown 清理、空行模式、OOXML 保护、表格辅助判断、临时文件路径和 TXT 转换等纯函数或轻量流程。

## 使用示例

```bash
# DOCX 单文件
python scripts/wfp_cli.py format -i report.docx

# 指定输出文件
python scripts/wfp_cli.py format -i report.docx -o report_final.docx

# TXT/MD 输入
python scripts/wfp_cli.py format -i meeting_notes.txt
python scripts/wfp_cli.py format -i draft.md --blank-line-mode "不改动任何空行"

# 多文件
python scripts/wfp_cli.py format a.docx b.txt c.md -o ./formatted_output
python scripts/wfp_cli.py format -i a.docx -i b.txt c.md -o ./formatted_output

# 目录批量处理，默认递归
python scripts/wfp_cli.py format -i ./documents -o ./documents_formatted

# 不递归子目录
python scripts/wfp_cli.py format -i ./documents -o ./documents_formatted --no-recursive

# 使用配置文件
python scripts/wfp_cli.py format -i input.docx --config ./wfp_config.json

# 指定 LibreOffice soffice 转换旧格式
python scripts/wfp_cli.py format -i old.doc --soffice /Applications/LibreOffice.app/Contents/MacOS/soffice

# 启用增强功能后处理
python scripts/wfp_cli.py format -i input.docx --enable-table-formatting --english-font "Times New Roman" --normalize-punctuation

# 查看和导出配置
python scripts/wfp_cli.py show-config
python scripts/wfp_cli.py save-config --set body_font=宋体 --set body_size=12

# 显示安装提示
python scripts/wfp_cli.py install-help

# 查看详细日志
python scripts/wfp_cli.py format -i input.docx -v
```

## 平台行为

- Windows：`.doc/.wps` 转 `.docx` 优先使用 WPS/Word COM 和 pywin32；如果 COM 转换失败，会尝试 LibreOffice `soffice` 兜底。
- macOS/Linux：可处理 `.docx/.txt/.md`；`.doc/.wps` 会尝试调用 LibreOffice `soffice` 转为 `.docx`，失败时提示安装 LibreOffice、用 `--soffice` 指定路径，或先手动另存为 `.docx`。
- 非 Windows 或 COM 不可用时，自动编号转文本会跳过。完成后需人工检查自动编号字体字号。

## 输出行为

- 成功时 stdout 每行打印一个输出 `.docx` 的绝对路径。
- `-v/--verbose` 开启后，详细处理日志写入 stderr。
- 单文件默认输出到同目录 `*_formatted.docx`。
- 目录默认输出到 `<输入目录>_formatted/`，或用户指定的输出目录。
- 退出码 `0` 表示全部成功，非 `0` 表示没有找到可处理文件或至少一个文件失败。
