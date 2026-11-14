# Word Formatter Pro v2.7.9

- 标题后留白支持“空一行”和“段落间距”。默认关闭；在 GUI 中启用时默认选择空一行，标题后已有空行时不再增加。主文和附件正文独立设置，只作用于标题及副标题组的末尾。
- 空行采用正文行距；段落间距继续支持磅和行。旧配置中已启用的段落间距保留原设置。CLI 和独立 Skill 同步支持新选项。
- 配置界面采用“页面与正文”“标题与附件”“表格与图表”三个标签页，配置按钮固定显示。输出目录单独记忆，支持跨平台打开目录。
- 加粗支持保留原样、加粗、取消加粗；完善表格对齐、边框及文字环绕选项。保护段内标题拆分中的图片、域、书签和超链接。
- 扩展附件识别，区分附件清单与附件正文。新增默认关闭的单一机关落款格式化，支持盖章／不盖章文字布局、日期规范化及附件正文开关。
- 修复符号标准化可能删除图片、破坏文字片段格式的问题，保留图片、书签、域和超链接，改进混排省略号的文字范围定位。
- 修复 Windows 包 Tcl/Tk DLL 与脚本版本混用导致的启动失败；构建时检查冻结包内容及真实界面启动。
- 完善 LibreOffice 转换进程树的超时回收、冻结程序调用外部转换器的运行环境，以及 Windows 长路径临时目录清理。
- 同步 Skill 的代码、默认配置、CLI、依赖和测试；Skill description、SKILL.md 及 Skill 说明文档不变。

## 使用新增选项

主文标题后空一行：`--set enable_title_after_blank_line=true`。

附件正文标题后空一行：`--set enable_attachment_title_after_blank_line=true`，同时启用附件格式化。

改用主文段落间距：`--set enable_title_after_blank_line=false --set enable_title_after_spacing=true --set title_after_spacing=12 --set title_after_spacing_unit=pt`。

空行与段落间距不叠加；配置同时启用时，以空行方式为准。已有空行保持原数量；不把图片、表格或域当作空行。

## 本地构建与验证

本次构建 Windows x86_64、Kylin V10 x86_64、UOS V20 x86_64。Windows 使用 Python 3.8 环境；Kylin/UOS 使用 Docker。新增留白功能做定向测试，并检查三个平台打包后的启动界面。Win7 和 macOS 未做本轮实机验证。
