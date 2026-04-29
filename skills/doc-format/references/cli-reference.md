# CLI Reference

## 命令格式

主命令：

```bash
python scripts/wfp_cli.py <子命令> [参数]
```

格式化：

```bash
python scripts/wfp_cli.py format [输入路径...] [-i <输入路径>] [-o <输出路径>] [配置参数] [转换参数] [其他参数]
```

查看配置：

```bash
python scripts/wfp_cli.py show-config [配置覆盖参数]
```

保存配置：

```bash
python scripts/wfp_cli.py save-config [配置覆盖参数] [-o <配置文件路径>]
```

## `format` 参数

| 参数 | 默认值 | 说明 |
|---|---:|---|
| `paths` | 可选 | 位置参数输入路径，可一次传入多个文件或目录 |
| `-i, --input` | 可选 | 输入文件或目录，可重复。与 `paths` 可同时使用 |
| `-o, --output` | 自动生成 | 输出文件或目录。多文件或目录输入时必须是目录 |
| `--config` | 自动读取当前目录 `wfp_config.json` | 指定 JSON 配置文件 |
| `--config-json` | 无 | 内联 JSON 对象，会覆盖配置文件中的同名字段 |
| `--set key=value` | 无 | 覆盖单个配置项，可重复使用 |
| `--enable-table-formatting` | 关闭 | 启用表格内容自动调整 |
| `--disable-table-formatting` | 关闭 | 关闭表格内容自动调整 |
| `--enable-custom-english-font` | 关闭 | 启用数字和字母单独字体 |
| `--disable-custom-english-font` | 关闭 | 关闭数字和字母单独字体 |
| `--english-font <字体>` | `Times New Roman` | 设置数字和字母字体；使用后自动启用 `use_custom_english_font` |
| `--normalize-punctuation` | 关闭 | 启用符号标准化 |
| `--disable-normalize-punctuation` | 关闭 | 关闭符号标准化 |
| `--backend auto/office/libreoffice/none` | `auto` | 转换后端。Windows 默认 Office/WPS，macOS/Linux 默认 LibreOffice |
| `--office-app auto/wps/word` | `auto` | Windows Office 后端选择 |
| `--soffice <路径>` | 自动查找 | 指定 LibreOffice `soffice` 路径 |
| `--timeout <秒>` | `120` | LibreOffice 单文件转换超时秒数 |
| `--no-recursive` | 关闭 | 目录输入时不递归子目录；默认递归 |
| `--convert-numbering` | Windows 默认开，macOS/Linux 默认关 | 尝试将自动编号转文本。macOS/Linux 通常不建议强制启用 |
| `--skip-numbering` | 关闭 | 跳过自动编号转文本 |
| `--keep-blank-lines` | 关闭 | 保留 TXT/MD 原始空行；默认删除多余空行 |
| `-v, --verbose` | 关闭 | 显示详细日志 |

## `show-config`

```bash
python scripts/wfp_cli.py show-config
```

输出 JSON，包含当前配置来源、当前目录自动读取的 `wfp_config.json` 路径、支持的输入类型、字号名称对照、常用可选增强项、所有配置项的值、名称和说明。

`show-config` 也支持 `--config`、`--config-json`、`--set` 和便利开关，用来查看覆盖后的配置效果。

## `save-config`

```bash
python scripts/wfp_cli.py save-config -o ./wfp_config.json
```

- 默认输出到当前目录 `wfp_config.json`
- 会先读取 `--config` 或当前目录已有 `wfp_config.json`，再应用 `--config-json`、`--set` 和便利开关，最后保存合并后的配置
- 保存后提醒用户还可启用表格内容自动调整、数字和字母字体、符号标准化

## 使用示例

```bash
# TXT/MD 输入
python scripts/wfp_cli.py format -i meeting_notes.txt
python scripts/wfp_cli.py format -i draft.md --keep-blank-lines

# 多文件
python scripts/wfp_cli.py format a.docx b.txt c.md -o ./formatted_output

# 不递归子目录
python scripts/wfp_cli.py format -i ./documents -o ./documents_formatted --no-recursive

# 启用增强功能后处理
python scripts/wfp_cli.py format -i input.docx \
  --enable-table-formatting \
  --english-font "Times New Roman" \
  --normalize-punctuation
```
