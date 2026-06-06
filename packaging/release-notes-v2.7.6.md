# v2.7.6

* **修复部分 Windows 系统界面显示异常**：优化 Tkinter 主界面左右分栏布局，给左侧上传/排版/日志区域增加最小宽度保护，避免在高 DPI、系统字体放大或主题字体度量差异下被右侧参数区挤没。
* **优化窗口自适应**：启动窗口会根据屏幕尺寸设置初始大小和最小大小，不再固定要求 1200×860，更适合不同分辨率和缩放比例的客户端。
* **新增布局恢复入口**：帮助菜单增加“重置界面布局”，可一键恢复左右分栏比例。
* **三平台 Release 资产**：本版本 Release 包含 Windows、Kylin 和 macOS 三个二进制资产。Windows `.exe` 与 Kylin V10 `.AppImage` 使用本地 v2.7.6 构建/准备的资产；macOS Apple Silicon `.app.zip` 沿用 v2.7.5 包并按 v2.7.6 资产名重新上传。
* **跨平台行为不变**：Windows 下仍优先调用 WPS/Word COM；macOS/Kylin/Linux 下仍不调用 WPS/Word COM，不执行自动编号转文本，旧格式 `.doc/.wps` 仅在检测到 LibreOffice `soffice` 后转换。
