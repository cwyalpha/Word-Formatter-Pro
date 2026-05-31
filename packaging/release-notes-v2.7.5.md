# v2.7.5

* **新增 macOS Apple Silicon 发布包**：提供 `Word-Formatter-Pro.v2.7.5.macOS-arm64.app.zip`，下载后解压运行 `.app`。由于未签名，首次运行如遇 Gatekeeper 提示，可在“系统设置 > 隐私与安全性”中允许打开。
* **三平台 Release 资产**：本版本 Release 包含 Windows、Kylin 和 macOS 三个二进制资产。Windows `.exe` 与 Kylin V10 `.AppImage` 沿用 v2.7.4 二进制并按 v2.7.5 资产名重新上传；macOS 为 v2.7.5 新增构建。
* **明确 macOS/Kylin/Linux 行为**：macOS 与 Kylin/Linux 一样，不调用 WPS/Word COM，不执行自动编号转文本；处理 `.doc/.wps` 时仅在检测到 LibreOffice `soffice` 后转换，未安装 LibreOffice 时会跳过旧格式文件并继续处理 `.docx/.txt/.md`。
* **修复 macOS 启动显示问题**：macOS 包改用带 Tk 9 的 Homebrew Python 构建，并在 macOS 下禁用不稳定的拖拽组件，避免空白 `tk` 窗口或界面文字渲染异常；请使用按钮添加文件/文件夹。
* **修复 macOS 打包转换问题**：macOS `.app` 现在会把 `python-docx` 模板放入冻结环境可解析的位置，修复处理 `.md/.txt` 等文件时找不到 `default-footer.xml` 的错误；同时兼容 LibreOffice 转换旧 `.doc/.wps` 后可能产生的 `w:jc="start"` 段落对齐值。
* **补充跨平台打包脚本**：新增干净 venv 打包流程，后续仍可在对应系统环境中生成 Windows `.exe`、Kylin/Linux AppImage 和 macOS `.app.zip`。
* **同步 Skill 代码和文档**：`skills/doc-format` 继续复用同一套核心逻辑，CLI、配置、核心排版、测试和跨平台说明与主程序保持一致；内置测试新增直接格式端到端、LibreOffice `.doc` 转换和旧 `.wps` 跳过路径覆盖。
