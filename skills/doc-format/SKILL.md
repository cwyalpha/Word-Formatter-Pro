---
name: doc-format
description: 公文格式排版工具。将doc/docx/wps/txt/md文档按照公文排版规范进行自动格式化，输出标准docx。支持智能识别标题层级（一、/（一）/1./（1））、题目/副标题、附件格式化、图表标题识别、页面边距/页码/行距设定。支持批量处理目录、自定义字体配置。当用户需要排版公文、格式化文档、统一文档格式、调整字体字号行距边距时使用此技能。关键词：公文排版、文档格式化、字体设置、标题层级、格式统一、排版工具。
---

# doc-format — 公文格式排版工具

使用 `scripts/wfp_cli.py` 将文档按公文排版规则格式化为标准 `.docx`。原始文件不原地修改，Word/WPS 类文档会先复制或转换到临时副本，再执行排版。

## 核心能力

- 支持 `.docx`、`.doc`、`.wps`、`.txt`、`.md`。
- 支持单文件、多文件、一个或多个目录；目录输入默认递归扫描，输出目录保留原结构。
- 自动识别题目、副标题、四级标题、二级标题段内正文、图/表标题、附件标识。
- 可选表格内容自动调整、数字和字母单独字体、符号标准化。
- 支持 `wfp_config.json`、`--config`、`--config-json`、`--set key=value` 和便利开关覆盖配置。

## 使用前检查

1. 确认脚本路径：脚本位于本 Skill 目录下的 `scripts/wfp_cli.py`。如果当前工作目录不是 Skill 目录，使用 `SKILL.md` 所在目录推导脚本绝对路径，或先 `cd` 到 Skill 目录。
2. 检查配置：CLI 会自动读取当前工作目录的 `wfp_config.json`。如果用户指定 `--config` 或明确给出配置文件，使用用户指定配置；否则说明将使用当前目录配置或内置默认配置。
3. 检查输入：如果用户没有给出文件或目录，请先要求提供输入路径；告知支持 `.docx/.doc/.wps/.txt/.md`，目录默认递归处理。
4. 检查转换条件：`.doc/.wps` 需要转换后端。Windows 通常用 WPS/Word；macOS/Linux 通常用 LibreOffice。转换失败时提示用户先另存为 `.docx`。

## 常用命令

```bash
# 单文件
python scripts/wfp_cli.py format -i report.docx

# 指定输出文件
python scripts/wfp_cli.py format -i report.docx -o report_final.docx

# 多文件
python scripts/wfp_cli.py format -i a.docx -i b.txt -i c.md -o ./formatted_output

# 目录批量处理，默认递归并保留目录结构
python scripts/wfp_cli.py format -i ./documents -o ./documents_formatted

# 查看当前配置
python scripts/wfp_cli.py show-config

# 保存配置到当前目录 wfp_config.json
python scripts/wfp_cli.py save-config --set body_font=宋体 --set body_size=12
```

## 配置处理流程

当用户要求调整字体、字号、边距、行距、表格、符号等配置时：

1. 运行 `show-config` 了解当前配置来源和值。
2. 将用户自然语言需求映射到配置字段。
3. 用 `save-config --set ...` 或便利开关保存到当前目录 `wfp_config.json`，或保存到用户指定路径。
4. 再运行 `format` 处理文档。

常用便利开关：

```bash
# 启用表格内容自动调整
python scripts/wfp_cli.py save-config --enable-table-formatting

# 数字和字母使用 Times New Roman
python scripts/wfp_cli.py save-config --enable-custom-english-font --english-font "Times New Roman"

# 启用符号标准化
python scripts/wfp_cli.py save-config --normalize-punctuation
```

## 输出行为

- 成功生成的 `.docx` 绝对路径写入 stdout。
- 日志写入 stderr。
- 退出码 `0` 表示全部成功，非 `0` 表示至少一个文件失败。
- 多文件处理中单个文件失败不会阻止后续文件继续处理。
- 处理失败时，最终回复应说明失败文件和错误原因。

## 完成后必读

每次完成格式化后，都要提醒用户打开输出文件抽查题目、副标题、标题层级、附件、图/表标题、页码和表格。

如果处理过 `.doc/.wps` 或 Word/WPS 类文档，尤其要检查自动编号。自动编号未能确认转换成功时，编号文本、段落缩进、字体字号可能没有正常格式化；如果编号无法单独选中，说明可能仍是 Word/WPS 自动编号，必要时请将其转为手写文本编号后重新运行工具。

表格内容自动调整、符号标准化属于更容易影响版面的增强项，重要文件建议在副本上检查效果。

## Reference Files

- 需要完整命令格式、参数说明、更多命令示例时，读取 `references/cli-reference.md`。
- 需要把用户配置需求映射到字段、查看字号对照时，读取 `references/config-reference.md`。
- 需要说明默认排版规则、标题识别、页面段落、TXT/MD、表格和附件规则时，读取 `references/formatting-rules.md`。

## Agent 使用准则

1. 没有输入时不要直接运行 `format`。
2. 用户询问配置时运行 `show-config`，不要凭记忆列配置。
3. 用户提出配置修改时，优先保存配置，再用该配置处理文档。
4. 用户只要求一次性临时调整时，可用 `--config-json` 或 `--set` 直接运行 `format`。
5. 目录输入默认递归；只有用户明确要求“只处理当前目录”时才加 `--no-recursive`。
6. 处理结束后必须提示用户检查自动编号所在标题和段落的转换效果。并提示用户，表格内容自动调整、数字和字母使用 Times New Roman、符号标准化功能未启用，如有需要可要求调整。
