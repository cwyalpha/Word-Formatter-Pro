# Word文档智能排版工具 (Word-Formatter-Pro)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

一款桌面应用程序，可将格式混乱的 Word 文档（及兼容格式）一键转换为符合规范的专业文档，实现排版工作自动化。

![软件截图](https://raw.githubusercontent.com/cwyalpha/Word-Formatter-Pro/main/screenshot.png)

## 主要功能

*   **基于国家标准**：软件的默认参数遵循《党政机关公文格式国家标准》(GB/T 9704-2012) 设置。
*   **一键式操作**：导入文件或文本，点击按钮即可完成排版。

### 智能识别与处理

*   **多级标题识别**：自动识别“一、”、“（一）”、“1.”、“(1)”等四级常规标题并应用格式。
*   **题目与图表标题定位**：自动查找并格式化居中主标题，以及图片和表格附近的“图X”、“表X”标题。
*   **段内格式化**：当二级标题与正文在同一段落时（如 `（一）标题。正文...`），程序能为标题和正文应用不同格式，且不拆分段落。
*   **保留原有格式**：在统一全文格式时，会保留已设置的**加粗、斜体、下划线、字体颜色**等。
*   **豁免特定内容**：自动跳过对表格、图片及嵌入的附件（如PDF、Excel）的格式化。

### 兼容性与易用性

*   **批量处理**：支持拖入单个文件、多个文件或整个文件夹。
*   **格式支持**：原生处理 `.docx`，并能自动将 `.doc`、`.wps`、`.txt` 转换为 `.docx` 进行处理。
*   **输入灵活**：支持文件处理和直接在软件内粘贴文本进行排版。
*   **安全无损**：所有操作均在副本上进行，原始文件不会被修改。
*   **参数自定义**：所有核心参数（页边距、字体、字号、行距等）均可在界面调整。配置方案可保存和加载。

## 如何使用

### 方式一：直接运行程序 (推荐)

1.  **环境依赖**：确保电脑上已安装 **Microsoft Office** 或 **WPS Office**。
2.  **下载**：访问项目的 [Releases](https://github.com/cwyalpha/Word-Formatter-Pro/releases) 页面，下载最新的 `.exe` 可执行文件。
3.  **运行**：双击 `.exe` 文件即可，无需安装。

### 方式二：从源码运行

1.  **环境依赖**：
    *   已安装 Python 3.x。
    *   已安装 **Microsoft Office** 或 **WPS Office**。

2.  **克隆仓库**：
    ```bash
    git clone https://github.com/cwyalpha/Word-Formatter-Pro.git
    ```

3.  **进入项目目录**：
    ```bash
    cd Word-Formatter-Pro
    ```

4.  **安装所需库**：
    ```bash
    pip install -r requirements.txt
    ```

5.  **运行程序**：
    ```bash
    python wfp.py 
    ```

## 操作流程

1.  **选择模式**：选择“文件批量处理”或“直接输入文本”。
2.  **添加内容**：
    *   **文件模式**：点击“添加文件”或“添加文件夹”导入。
    *   **文本模式**：在文本框中粘贴内容。
3.  **调整参数 (可选)**：在“参数设置”区进行调整，或加载已保存的配置方案。
4.  **开始排版**：点击“开始排版”按钮。
5.  **选择输出位置**：根据提示选择处理后文件的保存位置。
6.  **完成**：处理完毕后会弹出成功提示。

## 许可证 (License)

本项目采用 [MIT License](LICENSE) 授权。

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=cwyalpha/Word-Formatter-Pro&type=date&legend=top-left)](https://www.star-history.com/#cwyalpha/Word-Formatter-Pro&type=date&legend=top-left)
