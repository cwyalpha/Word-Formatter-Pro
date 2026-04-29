# Config Reference

## 根据用户需求修改配置

当用户要求调整字体、字号、边距、行距、表格、符号等配置时：

1. 运行 `show-config` 了解当前配置来源和值。
2. 将用户自然语言需求映射到配置字段。
3. 使用 `save-config` 保存到当前目录 `wfp_config.json`，或保存到用户指定路径。
4. 再运行 `format` 处理文档。

## 常见自然语言到配置字段

| 用户需求 | 配置字段或命令 |
|---|---|
| 正文字体改为宋体 | `--set body_font=宋体` |
| 正文字号改为小四 | `--set body_size=12` |
| 题目字号改为小二 | `--set title_size=18` |
| 行距改为 30 磅 | `--set line_spacing=30` |
| 强制 A4 | `--set force_a4=true` |
| 不设置大纲级别 | `--set set_outline=false` |
| 不启用附件格式化 | `--set enable_attachment_formatting=false` |
| 启用表格内容自动调整 | `--enable-table-formatting` |
| 启用表格智能对齐 | `--set table_smart_align=true` |
| 数字和字母使用 Times New Roman | `--enable-custom-english-font --english-font "Times New Roman"` |
| 启用符号标准化 | `--normalize-punctuation` |

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

## 保存配置示例

```bash
python scripts/wfp_cli.py save-config \
  --set body_font=宋体 \
  --set body_size=12 \
  --set line_spacing=30 \
  --enable-table-formatting \
  --english-font "Times New Roman" \
  --normalize-punctuation
```

## 使用配置处理文档

```bash
# 使用当前目录自动配置
python scripts/wfp_cli.py format -i ./documents -o ./documents_formatted

# 使用指定配置文件
python scripts/wfp_cli.py format -i input.docx --config ./project_wfp_config.json

# 使用内联 JSON 临时覆盖
python scripts/wfp_cli.py format -i input.docx --config-json "{\"force_a4\": true, \"line_spacing\": 30}"

# 使用 --set 临时覆盖
python scripts/wfp_cli.py format -i input.docx --set force_a4=true --set title_size=18
```
