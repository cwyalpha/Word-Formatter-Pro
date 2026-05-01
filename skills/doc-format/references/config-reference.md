# Config Reference

## 根据用户需求修改配置

当用户要求调整字体、字号、边距、行距、表格、符号等配置时：

1. 运行 `show-config` 了解默认配置字段和值。
2. 将用户自然语言需求映射到配置字段。
3. 使用 `save-config --set ...` 或便利开关保存到当前目录 `wfp_config.json`，或保存到用户指定路径。
4. 使用 `format` 处理文档；CLI 会自动读取当前目录的 `wfp_config.json`。

用户只要求一次性临时调整时，可用 `--config-json` 或 `--set` 直接运行 `format`。

## 常见自然语言到配置字段

| 用户需求 | 配置字段或命令 |
|---|---|
| 正文字体改为宋体 | `--set body_font=宋体` |
| 正文字号改为小四 | `--set body_size=12` |
| 题目字体改为华文中宋 | `--set title_font=华文中宋` |
| 题目字号改为小二 | `--set title_size=18` |
| 一级标题字体改为黑体 | `--set h1_font=黑体` |
| 二级标题字体改为楷体 | `--set h2_font=楷体` |
| 页码字体改为 Times New Roman | `--set page_number_font="Times New Roman"` |
| 正文行距改为 30 磅 | `--set line_spacing=30` |
| 题目行距改为 33 磅 | `--set title_line_spacing=33` |
| 副标题行距改为 33 磅 | `--set subtitle_line_spacing=33` |
| 强制 A4 | `--set force_a4=true` |
| 不设置大纲级别 | `--set set_outline=false` |
| 不启用附件格式化 | `--set enable_attachment_formatting=false` |
| 启用表格内容自动调整 | `--enable-table-formatting` |
| 启用表格智能对齐 | `--set table_smart_align=true` |
| 数字和字母使用 Times New Roman | `--enable-custom-english-font --english-font "Times New Roman"` |
| 启用符号标准化 | `--normalize-punctuation` |
| TXT/MD 不改动任何空行 | `--set blank_line_mode="不改动任何空行"` |
| TXT/MD 删除单个空行，多个空行保留至 1 个 | `--set blank_line_mode="删除单个空行，多个空行保留至1个空行"` |
| TXT/MD 保留单个空行，多个空行保留至 1 个 | `--set blank_line_mode="保留单个空行，多个空行保留至1个空行"` |

旧配置中的 `remove_blank_lines` 仍可读取；如未提供 `blank_line_mode`，CLI 会按旧字段映射到新的 TXT/MD 空行模式。

## 常用字号对照

| 字号名称 | pt 值 |
|---|---:|
| 一号 | 26 |
| 小一 | 24 |
| 二号 | 22 |
| 小二 | 18 |
| 三号 | 16 |
| 小三 | 15 |
| 四号 | 14 |
| 小四 | 12 |
| 五号 | 10.5 |
| 小五 | 9 |

## 配置文件示例

```json
{
    "body_font": "宋体",
    "body_size": 12,
    "line_spacing": 30,
    "force_a4": true,
    "enable_table_formatting": true,
    "use_custom_english_font": true,
    "english_font": "Times New Roman",
    "normalize_punctuation": true,
    "blank_line_mode": "删除单个空行，多个空行保留至1个空行"
}
```

配置文件不需要包含所有字段；只写需要覆盖默认值的字段即可。若使用 `save-config` 导出的完整模板，可直接在模板中修改对应值。

## 使用配置处理文档

```bash
# 保存配置到当前目录 wfp_config.json
python scripts/wfp_cli.py save-config \
  --set body_font=宋体 \
  --set body_size=12 \
  --set line_spacing=30 \
  --enable-table-formatting \
  --english-font "Times New Roman" \
  --normalize-punctuation

# 使用当前目录自动配置
python scripts/wfp_cli.py format -i ./documents -o ./documents_formatted

# 使用指定配置文件
python scripts/wfp_cli.py format -i input.docx --config ./wfp_config.json

# 使用内联 JSON 临时覆盖
python scripts/wfp_cli.py format -i input.docx --config-json "{\"force_a4\": true, \"line_spacing\": 30}"

# 使用 --set 临时覆盖
python scripts/wfp_cli.py format -i input.docx --set force_a4=true --set title_size=18
```
