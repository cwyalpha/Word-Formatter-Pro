# Word Formatter Pro v2.7.8

- macOS、Kylin、UOS 和其他 Linux 版本新增“识别自动编号并转换为文字”，默认开启，支持保存配置及 CLI 的 `convert_auto_numbering` 配置项。
- 支持正文、表格及模块已覆盖的其他文档区域，包含中文、日式、多级编号和自动/手写编号混排。转换后编号直接连接原文，不新增编号后的空格或制表符。
- 不支持的编号保留原状并记录警告，继续排版；Python 预处理不自动接受修订。源文件保持不变，旧格式仍沿用现有 DOC/WPS 转 DOCX 流程。
- Windows 继续优先使用 WPS/Office；缺少 COM、无法启动办公软件或 COM 预处理失败时，已有 Python 排版回退流程现在也执行编号转换。Windows 不显示新增设置项。
- 同步自包含 Skill 的代码、配置、依赖和测试；Skill description、SKILL.md 和说明文档保持不变。
- 编号预处理保留原缩进；后续排版仍应用所选排版配置。移除编号后间隔时正文会自然靠近编号。

本次本机构建目标：Windows x86_64、Kylin V10 x86_64、UOS V20 x86_64。macOS 已接入代码，不在本次本机打包范围内。
