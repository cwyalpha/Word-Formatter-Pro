---
name: doc-format
description: 公文/报告/常规 Word 文档自动排版为公文格式的技能。用于将 .docx、.doc、.wps、.txt、.md 单文件、多文件或目录批量转为符合公文规范的 .docx，目录输入递归保留原目录结构；支持题目/副标题识别、标题层级识别、附件、图表标题、表格内容自动调整、数字和字母字体、符号标准化。需要做文档排版、Word/WPS 转 docx 后排版、或批量规范化文档时应使用本技能。
---

# doc-format

使用 `scripts/wfp_cli.py` 将文档按公文排版规则格式化为 `.docx`。原始文件不原地修改。

## 支持范围

输入文件类型：

- `.docx`
- `.doc`
- `.wps`
- `.txt`
- `.md`

输入形式：

- 单个文件
- 多个文件
- 一个或多个目录
- 目录输入默认递归扫描子目录，并在输出目录中保留原目录结构

输出形式：

- 单文件输入默认输出到同目录：`原文件名_formatted.docx`
- 多文件或目录输入默认输出到新的目录；目录输入的子目录结构会保留
- 所有成功输出均为 `.docx`
- stdout 每行打印一个成功生成的 `.docx` 绝对路径

## 依赖

Python 包：

```bash
pip install -r skills/doc-format/requirements.txt
```

平台转换依赖：

- Windows 处理 `.doc/.wps` 和自动编号转文本：安装 WPS Office 或 Microsoft Word，并安装 `pywin32`
- macOS/Linux 处理 `.doc/.wps`：安装 LibreOffice，确保 `soffice` 可用
- macOS/Linux 不做自动编号转文本；LibreOffice 也不提供等价能力

LibreOffice 安装提示：

```bash
python skills/doc-format/scripts/wfp_cli.py install-help
```

## 平台行为

Windows 默认使用 WPS/Office COM 接口：

- `.doc/.wps` 会先尝试通过 WPS，再尝试 Microsoft Word 转为 `.docx`
- 如果转换失败，提示用户先手动转成 `.docx` 后再输入
- 自动编号会尝试通过 WPS/Word 转成普通可编辑文本
- 如果自动编号转换失败，继续执行后续格式化，并在处理完成后提醒用户人工检查编号文本和段落

macOS/Linux 默认使用 LibreOffice：

- `.doc/.wps` 会尝试通过 LibreOffice 转为 `.docx`
- 如果转换失败，提示用户先手动转成 `.docx` 后再输入
- 不执行自动编号转文本
- 处理完成后提醒用户自动编号文本、段落可能没有正常格式化，需要人工检查

## 常用命令

单文件：

```bash
python skills/doc-format/scripts/wfp_cli.py format -i input.docx
```

指定输出文件：

```bash
python skills/doc-format/scripts/wfp_cli.py format -i input.docx -o output.docx
```

多个文件：

```bash
python skills/doc-format/scripts/wfp_cli.py format -i a.docx -i b.txt -o formatted_output
```

目录递归批量处理并保留结构：

```bash
python skills/doc-format/scripts/wfp_cli.py format -i ./documents -o ./documents_formatted
```

指定 LibreOffice：

```bash
python skills/doc-format/scripts/wfp_cli.py format -i old.doc --backend libreoffice --soffice /Applications/LibreOffice.app/Contents/MacOS/soffice
```

查看详细日志：

```bash
python skills/doc-format/scripts/wfp_cli.py format -i input.docx -v
```

## JSON 配置工作流

默认使用脚本内置配置。脚本运行时会自动读取当前工作目录的 `wfp_config.json`；也可以用 `--config` 指定配置文件。

查看当前配置：

```bash
python skills/doc-format/scripts/wfp_cli.py show-config
```

保存当前配置到当前目录：

```bash
python skills/doc-format/scripts/wfp_cli.py save-config
```

按用户要求修改并保存配置：

```bash
python skills/doc-format/scripts/wfp_cli.py save-config --set body_font=仿宋_GB2312 --set body_size=16
```

一次性使用内联 JSON：

```bash
python skills/doc-format/scripts/wfp_cli.py format -i input.docx --config-json "{\"force_a4\": true}"
```

常用可选增强项：

```bash
# 启用表格内容自动调整
python skills/doc-format/scripts/wfp_cli.py save-config --enable-table-formatting

# 启用数字和字母单独字体，默认 Times New Roman
python skills/doc-format/scripts/wfp_cli.py save-config --enable-custom-english-font --english-font "Times New Roman"

# 启用符号标准化
python skills/doc-format/scripts/wfp_cli.py save-config --normalize-punctuation
```

当用户完成配置后，提醒用户还可以按需启用：

- 表格内容自动调整：`enable_table_formatting=true`
- 自定义数字和字母字体：`use_custom_english_font=true`，`english_font` 默认 `Times New Roman`
- 符号标准化：`normalize_punctuation=true`

如果用户提出新的配置修改需求，先保存 `wfp_config.json`，再重新运行 `format` 转换一次。

## Agent 使用准则

1. 先确认输入是文件、多文件还是目录；目录默认递归处理。
2. 如用户问“当前配置”，运行 `show-config` 并总结关键配置，不要手写猜测。
3. 如用户要求修改配置，运行 `save-config --set key=value` 或便利开关，将配置保存到当前目录 `wfp_config.json`。
4. 运行 `format` 时优先使用当前目录自动配置；用户给了配置文件时使用 `--config`。
5. Windows 上遇到 `.doc/.wps` 转换失败时，明确告诉用户需要先另存为 `.docx`。
6. macOS/Linux 上遇到 `.doc/.wps` 转换失败时，提示安装 LibreOffice 或先手动转为 `.docx`。
7. macOS/Linux 环境下，在最终回复中提醒：自动编号文本、段落可能没有正常格式化，需要人工检查。
