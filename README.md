# Word文档智能排版工具 (Word-Formatter-Pro)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Word-Formatter-Pro 是一款面向公文、报告和常规 Word 文档的桌面排版工具。它可以将格式混乱的 Word/WPS 文档、纯文本或 Markdown 内容转换为规范、统一的 `.docx` 文档，适合批量处理和日常办公排版。

![软件截图](https://raw.githubusercontent.com/cwyalpha/Word-Formatter-Pro/main/screenshot.png)

## 主要功能

*   **基于公文标准**：内置默认参数参考常用公文排版规范，包含页边距、页码、字体、字号、行距、标题层级等设置。
*   **一键式排版**：添加文件、文件夹或直接粘贴文本后，点击“开始排版”即可生成处理后的 Word 文档。
*   **安全副本处理**：原始文件不会被直接修改，程序会在系统临时目录中的安全副本或转换后的 `.docx` 上执行排版，临时文件使用带进程 ID 和随机后缀的命名方式，降低重名覆盖风险。

### 智能识别与处理

*   **主标题与副标题识别**：自动识别文档开头连续居中、字体字号一致的主标题；主标题下方字体字号不同的居中段落会识别为副标题。TXT/MD 文本会将首个非层级标题段落作为题目。
*   **四级标题识别**：自动识别“一、”、“（一）”、“1.”、“(1)”等常见层级标题，并分别应用对应格式和大纲级别。
*   **二级标题段内格式化**：当二级标题与正文在同一段落时，例如 `（一）标题。正文...`，程序会在同一段落内为标题和正文分别应用格式。
*   **图表标题定位**：自动查找图片或表格上方/下方最近的、居中的“图...”或“表...”标题，并套用图形标题或表格标题格式。
*   **附件格式化**：可识别“附件”、“附件1”、“附件：”等附件标识，自动设置段前分页，并继续识别附件自身的主标题和副标题；当附件标识后没有附件标题/副标题时，不会跳过紧随其后的正文段落。
*   **保留原有强调格式**：统一字体字号时，会尽量保留原有的加粗、斜体、下划线、字体颜色等格式。
*   **图片和嵌入对象保护**：包含图片、图形或嵌入对象的段落不会破坏对象本身，只对段落中的文字进行必要格式处理，并避免修改可能导致图片显示不全的行距等段落属性。
*   **特殊内容保护**：清理段首空白时，会避免删除包含域代码、书签、批注引用等特殊内容的 run。

### 文档格式兼容

*   **批量处理**：支持添加单个文件、多个文件或整个文件夹，也支持从文件管理器拖拽文件/文件夹到列表。
*   **多格式输入**：支持 `.docx`、`.txt`、`.md`，并可按环境处理 `.doc/.wps`。Windows 下优先使用 WPS/Word COM；macOS/Kylin/Linux 下会尝试调用 LibreOffice `soffice` 转换旧格式，未安装 LibreOffice 时会跳过 `.doc/.wps`，继续处理其他文件。
*   **直接输入文本**：可在软件中直接粘贴文本并生成排版后的 Word 文档，直接文本默认使用 A4 纸张。
*   **Markdown 清理**：处理 `.md` 文件时，会自动清理标题标记、粗体/斜体标记、链接、图片语法、引用、分隔线等 Markdown 标记；数字编号（如 `1.`、`1.2.3`）按源文档保留，不再自动递增重排。
*   **空行整理**：TXT/MD 文件支持三种空行处理方式：不改动任何空行；删除单个空行并将连续多个空行合并为一个空行；保留单个空行并将连续多个空行合并为一个空行。默认使用“删除单个空行，多个空行保留至1个空行”。
*   **修订和自动编号预处理**：仅 Windows 且可用 WPS/Word COM 时，会尝试接受修订并将自动编号转换为文本；macOS/Kylin/Linux 等非 Windows 环境固定跳过该步骤。
*   **大文件夹拖入保护**：添加或拖入包含超过 1000 个文件的文件夹时，会先弹出确认提示，避免误拖深层目录导致长时间扫描。

### CLI、Skill 与可维护性

*   **核心逻辑拆分**：排版核心、默认配置、GUI、CLI 和内置测试已拆分为独立文件，便于后续维护、命令行调用和 Skill 复用。
*   **命令行处理**：`wfp_cli.py` 支持单文件、多文件、重复 `-i`、位置参数和目录批量处理；目录输入默认递归，输出目录会保留原结构。
*   **配置覆盖能力**：CLI 会自动读取当前目录的 `wfp_config.json`，也支持 `--config`、`--config-json`、`--set key=value` 和常用便利开关覆盖配置。
*   **跨平台旧格式转换**：GUI/CLI/Skill 可自动查找 LibreOffice `soffice`，用于在 macOS/Kylin/Linux 或 COM 不可用时将 `.doc/.wps` 转为 `.docx`；也可通过 CLI 的 `--soffice` 手动指定路径。未检测到 LibreOffice 时，旧格式文件会被记录为跳过。
*   **自包含 Skill**：`skills/doc-format` 的 `scripts` 目录包含完整 CLI、核心逻辑、默认配置和测试代码，单独安装 `doc-format` 文件夹后也可运行。
*   **内置单元测试**：可通过 `python wfp.py --test`、`python wfp_cli.py test` 或 Skill 中的 `python scripts/wfp_cli.py test` 运行内置测试。
*   **正则预编译**：常用文本识别、Markdown 清理、标题判断、附件判断和表格数字判断的正则表达式已提升为模块级预编译常量，减少循环中的重复编译。

### 可配置排版参数

*   **页面设置**：可设置上下左右页边距、页脚距、页码对齐方式和页码字体字号；可选择是否强制设置为 A4 纸张。
*   **标题与正文样式**：可分别设置题目、副标题、一级标题、二级标题、正文/三四级标题的字体、字号和行距。
*   **段落缩进修复**：正文和层级标题会统一清理残留的左缩进、悬挂缩进和固定首行缩进，再按规则设置标准首行缩进，避免出现缩进叠加。
*   **字体选择增强**：所有字体下拉框会在内置推荐字体下方显示系统已安装字体，并用“── 已安装字体 ──”分隔；字体输入框仍支持手动输入未列出的字体名称。
*   **数字和字母字体设置**：可勾选“自定义数字和字母字体”后单独选择西文/数字字体，默认值为 Times New Roman；未勾选时，数字和字母会跟随正文、标题等各自的字体设置。
*   **TXT/MD 空行处理**：可通过下拉框选择不改动空行、删除单个空行或保留单个空行，同时支持将连续多个空行压缩为一个空行。
*   **符号标准化**：可启用实验性的符号标准化，保守修复中文语境中的中英文标点混用，包括逗号、句号、分号、冒号、问号、感叹号、括号、引号和省略号等。
*   **表格内容自动调整**：可启用实验性的表格内容格式化，分别设置表头/表格字体，统一字号、行距、行高、宽度、边框，并可自动调整列宽；总开关未启用时，相关选项会保持显示当前值但置灰不可编辑。
*   **表格智能对齐**：启用后，表头、序号列和短文本居中，数字靠右，长文本靠左；默认会保留单元格原始对齐方式。
*   **依赖选项联动**：附件格式化、自定义数字和字母字体、表格自动调整等开关未启用时，对应细项会自动置灰，避免误解当前生效范围。
*   **配置方案管理**：支持加载配置、保存配置、保存为默认配置和恢复内置默认配置。
*   **界面布局优化**：左侧集中处理文件/文本输入、开始排版、进度条和日志；文件列表支持水平滚动，方便查看较长路径；右侧为可滚动参数区，配置按钮保持 2x2 布局。
*   **后台线程化处理**：排版任务会在后台线程中执行，批量处理期间可继续滚动日志和调整窗口；任务进行中关闭窗口会询问确认，并尽量刷新剩余日志。

## 效果演示

### Word 文件处理前后

![Word处理前后](https://raw.githubusercontent.com/cwyalpha/Word-Formatter-Pro/main/demo_word_before_after.png)

### TXT 文件处理前后

![TXT处理前后](https://raw.githubusercontent.com/cwyalpha/Word-Formatter-Pro/main/demo_txt_before_after.png)

## 如何使用

### 方式一：直接运行程序 (推荐)

1.  **下载程序**：访问项目的 [Github Releases](https://github.com/cwyalpha/Word-Formatter-Pro/releases) 或 [Gitee Releases](https://gitee.com/cwyalpha/Word-Formatter-Pro/releases) 页面下载对应系统的成品程序。Windows 用户下载 `.exe`；Kylin/Linux 用户下载 `.AppImage`；macOS Apple Silicon 用户下载 `macOS-arm64.app.zip`。
2.  **运行程序**：Windows 下双击 `.exe` 即可使用；macOS 下解压 `.zip` 后双击 `.app`，未签名首次运行如遇系统拦截，可在“系统设置 > 隐私与安全性”中允许打开；Kylin/Linux 下运行发布包中的 `.AppImage`。如果系统未自动赋予执行权限，可在程序目录执行：
    ```bash
    chmod +x Word-Formatter-Pro.v2.7.6.Kylin-V10.x86_64.AppImage
    ./Word-Formatter-Pro.v2.7.6.Kylin-V10.x86_64.AppImage
    ```
3.  **可选依赖**：`.docx/.txt/.md` 可直接处理。Windows 下如需完整处理 `.doc/.wps`、修订和自动编号转文本，请确保已安装 **Microsoft Office** 或 **WPS Office**；macOS/Kylin/Linux 下不会调用 WPS/Word，也不会执行自动编号转文本。如需在 macOS/Kylin/Linux 下处理 `.wps/.doc` 旧格式，请安装 LibreOffice：
    ```bash
    # macOS
    brew install --cask libreoffice

    # Kylin/Linux
    sudo apt-get update
    sudo apt-get install libreoffice libreoffice-writer
    ```
    未安装 LibreOffice 时，macOS/Kylin/Linux 版本会跳过 `.wps/.doc`，继续处理 `.docx/.txt/.md`。

### 方式二：从源码运行

1.  **环境依赖**：
    *   已安装 Python 3.x（Kylin 上已验证目标为 Python 3.12 venv 环境）。
    *   Windows 下如需处理 `.doc/.wps` 或转换自动编号，需安装 **Microsoft Office** 或 **WPS Office**，并安装 pywin32。
    *   macOS/Kylin/Linux 可处理 `.docx/.txt/.md`；如需处理 `.doc/.wps`，建议安装 LibreOffice 并确保 `soffice` 可用。未安装 LibreOffice 时旧格式文件会跳过，且不会执行 WPS/Word 自动编号转文本。

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
    python -m pip install -U pip
    pip install -r requirements.txt
    ```

5.  **运行程序**：
    ```bash
    python wfp.py
    ```

### 从源码打包发布程序

项目提供 `packaging/build_release.py`，打包时会创建干净的构建 venv，避免把开发机全局环境中的无关依赖一起打进程序。PyInstaller 不支持跨系统交叉打包，请在目标系统上运行对应命令：

```bash
# macOS Apple Silicon：生成 release/Word-Formatter-Pro.v2.7.6.macOS-arm64.app.zip
python3 packaging/build_release.py macos --arch arm64

# Windows：生成 release/Word-Formatter-Pro.v2.7.6.exe
python packaging/build_release.py windows

# Kylin/Linux：生成 release/Word-Formatter-Pro.v2.7.6.Kylin-V10.x86_64.AppImage
python packaging/build_release.py kylin --arch x86_64 --appimagetool /path/to/appimagetool

# 如需复用旧发布资产，可下载后重新生成校验文件
python packaging/build_release.py reused-assets --overwrite
```

打包 GUI 必须使用可 `import tkinter` 的 Python。macOS 上建议使用 Homebrew Python 并安装 Tk 支持：

```bash
brew install python-tk@3.14
python3 -c "import tkinter; print(tkinter.TkVersion)"
```

如果输出为 `9.0` 或更高版本，再运行 macOS 打包命令；系统 `/usr/bin/python3` 自带的旧 Tcl/Tk 在新 macOS 上可能出现界面文字渲染异常。

### 方式三：命令行 CLI 使用

源码版本提供 `wfp_cli.py`，适合无界面批量处理、脚本调用和跨平台 Agent 调用。

```bash
# 单文件
python wfp_cli.py format -i input.docx

# 多文件或目录
python wfp_cli.py format a.docx -i b.txt -i ./documents -o ./formatted_output

# 查看当前配置和字段说明
python wfp_cli.py show-config

# 保存配置到当前目录 wfp_config.json
python wfp_cli.py save-config --set body_font=宋体 --set body_size=12

# 使用 LibreOffice soffice 转换 .doc/.wps
python wfp_cli.py format -i old.doc --soffice /Applications/LibreOffice.app/Contents/MacOS/soffice

# macOS/Kylin/Linux 未安装 LibreOffice 时，.doc/.wps 会跳过，其他文件继续处理
python wfp_cli.py format -i old.doc -i normal.docx -o ./formatted_output -v

# 运行内置测试
python wfp_cli.py test
```

CLI 支持 `--config`、`--config-json`、`--set key=value`、`--enable-table-formatting`、`--english-font`、`--normalize-punctuation`、`--blank-line-mode` 等参数；可通过 `python wfp_cli.py format --help` 查看完整说明。

### 方式四：作为 Agent Skill 安装和使用

项目已提供 `doc-format` Skill，可供支持 Skills 的 Win、Linux、macOS 端 Agent 调用。该 Skill 将排版核心改造为无界面的 CLI 脚本，支持处理 `.docx`、`.doc`、`.wps`、`.txt`、`.md`，可输入单文件、多文件或目录；目录输入会递归处理支持的文件，并在输出目录中保留原目录结构，最终输出 `.docx` 文件。

安装 Skill：

```bash
npx skills add https://github.com/cwyalpha/Word-Formatter-Pro
```

安装后，Agent 可按 Skill 说明调用 `scripts/wfp_cli.py`。Skill 的 `scripts` 目录包含完整 CLI、核心排版逻辑、默认配置和内置测试，不依赖上级项目目录。CLI 支持 JSON 配置：默认读取当前目录的 `wfp_config.json`，也可通过 `--config` 指定配置文件；可使用 `show-config` 查看当前配置，使用 `save-config` 保存修改后的配置。配置完成后，还可按需启用表格内容自动调整、自定义数字和字母字体（默认 Times New Roman）以及符号标准化，再重新执行一次格式化。

跨平台转换说明：Windows 下处理 `.doc/.wps` 会优先尝试 WPS 或 Microsoft Word 接口，并尝试将自动编号转为可编辑文本；macOS/Kylin/Linux 下不会调用 WPS/Word，也不会执行自动编号转文本，`.doc/.wps` 仅在检测到 LibreOffice `soffice` 时转换为 `.docx`。如果未安装 LibreOffice，旧格式文件会跳过；如需处理，请安装 LibreOffice、通过 `--soffice` 指定路径，或先手动另存为 `.docx`。

## 操作流程

1.  **选择模式**：选择“文件批量处理”或“直接输入文本”。
2.  **添加内容**：
    *   文件模式：点击“添加文件”“添加文件夹”，或直接拖拽文件/文件夹到列表。
    *   文本模式：在文本框中粘贴需要排版的内容。
3.  **调整参数**：在右侧“参数设置”中调整页面、标题、正文、表格、附件、TXT/MD 空行处理等参数，也可以加载已保存的配置。
4.  **开始排版**：点击左侧“开始排版”按钮，处理进度会显示在进度条和状态文字中。
5.  **选择输出位置**：文件批量处理时选择输出文件夹，直接输入文本时选择输出 `.docx` 文件位置。
6.  **完成处理**：处理结束后会生成 `_formatted.docx` 或指定名称的 Word 文档。

## 常见问题 (FAQ)

**Q1：为什么排版后的字体显示不正确，或者不是公文标准字体（如方正小标宋）？**

A：通常是因为电脑中缺少对应字体文件，例如方正小标宋、方正仿宋_GBK 等。出于版权原因，本工具无法捆绑这些字体。请自行从正规渠道获取并安装到 Windows 系统字体文件夹中，然后重新运行本工具。

**Q2：我使用的是 32 位 Windows 系统，为什么无法运行 .exe 文件？**

A：Releases 页面提供的 `.exe` 可执行文件通常面向 64 位 Windows 系统打包。32 位系统无法直接运行该文件。

**Q3：我使用的是 32 位系统或国产操作系统（如 UOS、Kylin），如何使用本工具？**

A：优先参考“方式一：直接运行程序”，下载发布页提供的 Kylin/Linux 成品程序；也可以参考“方式二：从源码运行”。国产操作系统或其他非 Windows 环境可直接处理 `.docx/.txt/.md`；如果需要处理 `.doc/.wps`，请安装 LibreOffice 并确认 `soffice` 可用。未安装 LibreOffice 时，`.doc/.wps` 会被记录为跳过。

**Q4：点击“开始排版”后程序报错、卡住或无响应？**

A：排版任务已放入后台线程，并增加进度条和状态提示。如果仍然报错，请查看日志窗口中的错误信息。Windows 下处理 `.doc/.wps`、修订或自动编号转换时，请确认电脑已正确安装 **Microsoft Office** 或 **WPS Office**，并具备 pywin32 环境；macOS/Kylin/Linux 下不调用 WPS/Word，旧格式转换依赖 LibreOffice，未安装时会跳过 `.doc/.wps`。若 Windows 程序结束后无法打开文档，可以在任务管理器中结束相关 Word/WPS 进程后重新打开。

**Q5：下载的 .exe 文件被杀毒软件报毒？**

A：本程序使用 PyInstaller 打包，非签名可执行文件可能触发部分杀毒软件误报。项目代码开源，不含恶意逻辑。可以添加白名单，或使用源码方式运行。

**Q6：符号标准化会不会误改英文、数字或代码内容？**

A：符号标准化采用保守规则，只在检测到中文语境时处理常见标点，并尽量跳过数字和英文字母附近的英文符号。该功能仍属于实验功能，重要文档建议先在副本上检查效果。

**Q7：表格内容自动调整默认为什么没有开启？**

A：不同文档的表格结构差异很大，自动调整可能改变原有表格布局。因此该功能默认关闭，需要在“表格内容（实验功能）”中手动启用。

**Q8：非 Windows 环境下自动编号转文本是否会执行？**

A：不会。自动编号转文本依赖 WPS/Word COM，因此 macOS/Kylin/Linux 等非 Windows 环境固定跳过该步骤；即使安装 LibreOffice 用于转换 `.doc/.wps`，也不会把自动编号转换为普通文本。处理完成后请人工检查自动编号，尤其是无法单独选中的编号，其字体字号可能仍由编号样式控制。

## 版本更新记录

### v2.7.6

*   **修复部分系统界面显示异常**：优化 Tkinter 左右分栏布局，给左侧上传、排版和日志区域增加最小宽度保护，避免被右侧参数设置区域挤没。
*   **优化窗口自适应**：启动窗口根据屏幕尺寸设置初始大小和最小大小，更适合不同分辨率、DPI 缩放和系统文字大小。
*   **新增布局恢复入口**：帮助菜单增加“重置界面布局”，便于用户手动恢复左右栏比例。
*   **三平台 Release 资产**：v2.7.6 包含 Windows `.exe`、Kylin V10 `.AppImage` 和 macOS `.app.zip`；macOS 包沿用 v2.7.5 并按 v2.7.6 资产名上传。

### v2.7.5

*   **新增 macOS Apple Silicon 发布包**：提供 `macOS-arm64.app.zip`，macOS 与 Kylin/Linux 一样不调用 WPS/Word COM，不执行自动编号转文本。
*   **三平台 Release 资产**：v2.7.5 包含 Windows `.exe`、Kylin V10 `.AppImage` 和 macOS `.app.zip`；Windows/Kylin 二进制沿用 v2.7.4 并按 v2.7.5 资产名上传。
*   **完善跨平台打包流程**：新增干净 venv 的 PyInstaller 打包脚本，保留 Windows、Kylin/Linux 和 macOS 三端构建能力；macOS 打包建议使用带 Tk 9 的 Homebrew Python。
*   **修复 macOS 启动显示问题**：macOS 下禁用不稳定的拖拽组件，改用按钮添加文件/文件夹，避免出现空白 `tk` 窗口或界面文字渲染异常。
*   **修复 macOS 转换错误**：macOS `.app` 会额外补齐 `python-docx` 模板运行时路径，避免处理 `.md/.txt` 等文件并设置页脚页码时找不到 `default-footer.xml`；同时兼容 LibreOffice 转换旧格式后出现的 `w:jc="start"` 对齐值。
*   **同步 Skill**：`skills/doc-format` 的 CLI、核心逻辑、配置、测试和说明同步更新。

### v2.7.4

*   **统一 Windows/Linux 兼容路径**：GUI、CLI 和 Skill 共用核心旧格式转换逻辑；Windows 优先 WPS/Word COM，Linux/Kylin/macOS 优先 LibreOffice `soffice`。
*   **适配 Kylin 发布包**：更新 PyInstaller spec 和依赖文件，并在直接运行说明中补充 Kylin/Linux 成品程序入口。
*   **Linux 无 LibreOffice 时跳过旧格式**：未检测到 LibreOffice 时跳过 `.doc/.wps`，继续处理 `.docx/.txt/.md`。
*   **明确自动编号策略**：Linux/Kylin 不调用 WPS/Word，也不执行自动编号转文本。

### v2.7.3

*   **修复 Markdown 数字编号递增问题**：TXT/MD 中手写的数字编号按源文档保留，不再把不同小节中的 `1.` 连续递增成 `6.`、`7.` 等编号。
*   **同步 Skill 脚本与测试**：更新 `skills/doc-format` 内置脚本，并增加 Markdown 编号保持原样的回归测试。

### v2.7.2

*   **重构核心架构**：将旧单文件实现拆分为 GUI 入口、CLI 入口、核心排版逻辑、默认配置和内置测试，降低后续维护和 Skill 复用成本。
*   **新增命令行 CLI**：`wfp_cli.py` 支持单文件、多文件、重复 `-i`、位置参数和目录批量处理；支持 `show-config`、`save-config`、`test` 和 `install-help` 子命令。
*   **增强 CLI 配置能力**：CLI 会自动读取当前目录 `wfp_config.json`，支持 `--config`、`--config-json`、`--set key=value` 以及表格、数字/字母字体、符号标准化等便利开关。
*   **完善跨平台旧格式转换**：CLI/Skill 在 macOS/Linux 或 COM 不可用时会尝试调用 LibreOffice `soffice` 将 `.doc/.wps` 转为 `.docx`；Windows 下仍优先使用 WPS/Word COM，失败后再尝试 `soffice` 兜底。
*   **新增自包含 Agent Skill**：`skills/doc-format` 的 `scripts` 目录包含完整 CLI、核心逻辑、默认配置和测试代码，单独安装 `doc-format` 文件夹后即可使用。
*   **新增内置单元测试**：通过 `python wfp.py --test`、`python wfp_cli.py test` 或 Skill CLI 的 `test` 子命令，可验证文本处理、Markdown 清理、空行模式、OOXML 保护、表格辅助判断和 TXT 转换等逻辑。
*   **实现正则预编译**：将常用正则表达式提升为模块级预编译常量，覆盖 Markdown 清理、中文判断、标题识别、附件识别、表格数字判断和安全文件名处理等路径。

### v2.7.1

*   **修复附件后段落跳转问题**：附件标识后没有附件标题或副标题时，不再跳过紧随其后的正文段落。
*   **修复段首空白 run 删除副作用**：删除空白 run 前会检查是否只包含普通文本节点，避免破坏域代码、书签、批注引用等特殊内容。
*   **优化 TXT/MD 空行处理**：新增三种空行处理模式，默认删除单个空行，并将连续多个空行合并为一个空行。
*   **提升跨平台兼容性**：`win32com` 和 `pythoncom` 改为条件导入，非 Windows 环境可启动 GUI 并处理 `.docx/.txt/.md`。
*   **新增后台线程化处理和进度提示**：排版任务在后台线程中执行，批量处理时显示当前进度、文件名和成功/失败数量。
*   **新增关闭确认与日志保护**：任务进行中关闭窗口会询问确认，并在关闭前尽量刷新日志队列。
*   **优化文件列表体验**：增加水平滚动条，长文件名和长路径可横向查看。
*   **新增大文件夹拖入保护**：添加或拖入超过 1000 个文件的文件夹时会弹出确认提示，并按每个目录分别判断。
*   **优化图片和嵌入对象检测**：改用 OOXML 元素查找判断域代码、图片和嵌入对象，同时保留媒体段落行距保护逻辑。
*   **优化 COM 应用管理和临时文件处理**：通过 `WPSAppManager` 统一管理 WPS/Word 进程，临时 `.docx` 文件改为生成在系统临时目录并使用更安全的命名方式。

### v2.6.9

*   **修复自动编号标题编号字体错误的问题**：现在自动编号的标题，编号的字体可以随标题字体设置改变。

### v2.6.8

*   **增强所有字体下拉框**：在内置推荐字体下方追加系统已安装字体列表，并使用“── 已安装字体 ──”分隔；下拉框仍保持可编辑，支持手动输入字体名称。
*   **优化数字和字母字体设置**：将原固定 Times New Roman 的开关改为“自定义数字和字母字体”+ 字体选择框。未勾选时数字/字母跟随当前正文、标题等字体；勾选后才使用自定义字体，默认 Times New Roman。
*   **优化依赖选项的界面状态**：自定义数字和字母字体、附件格式化、表格自动调整未启用时，对应设置项会变灰；启用后自动恢复可编辑。
*   **修复表格设置置灰时数值不显示的问题**：表格行距、行高、宽度、边框粗细等数值在总开关未启用时仍会显示当前值，只是置灰不可编辑。
*   **增强配置兼容性**：兼容旧配置中的 `use_times_new_roman` 字段；配置加载或保存时遇到空值，会回退到内置默认值，避免空字符串覆盖默认表格参数。

### v2.6.7

*   **新增表格内容自动调整选项**：增加“表格内容（实验功能）”设置区，支持启用表格自动调整，并可设置表头字体、表格字体、字号、行距、行高、表格宽度、边框粗细等。
*   **新增表格智能处理能力**：可自动调整列宽、统一表格边框、设置表头加粗，并支持智能调整单元格对齐。
*   **优化按钮和参数布局**：配置按钮改为 2x2 布局，参数区保持滚动显示，减少界面拥挤。

### v2.6.6

*   **新增符号标准化功能**：增加“启用符号标准化（实验功能）”选项，可保守修复中文语境中的中英文标点混用。
*   **支持段落和表格单元格处理**：符号标准化会同时处理正文段落和表格单元格中的文本，并跳过域代码等特殊内容。

### v2.6.5

*   **修复段落缩进叠加问题**：清理段落中残留的左缩进、悬挂缩进和固定首行缩进，再设置标准首行缩进，避免出现缩进异常。
*   **新增数字和字母字体设置**：增加“数字和字母使用 Times New Roman 字体”选项，支持中文字体与西文/数字字体分别设置。

### v2.6.4

*   **修正直接文本输入的纸张大小**：现在使用“直接输入文本”功能生成的文档，会强制默认使用 **A4** 纸张（之前默认为 Letter 信纸）。
*   **新增“强制设置为A4纸张”选项**：在“页面设置”区域增加了一个复选框。
    *   如果不勾选（默认）：处理现有的 Word/WPS 文件时，保持原文档的纸张大小不变。
    *   如果勾选：处理所有文件时，强制将纸张大小修改为 A4。

### v2.6.3

*   **优化 UI界面**：部分用户反馈，操作系统设置了125%文本后，按钮显示不全，根据反馈修改了页面布局。如v2.6.2显示无问题可不用更新。

### v2.6.2

*   **新增 多行标题、副标题识别**：识别多行标题、副标题，并设置字体、字号、行距。
*   **优化 UI界面**：优化参数设置显示，增加程序对题目、各级标题、附件等不同内容的判断规则。

### v2.6.1

*   **优化 大纲级别设置**：修复了大纲级别设置不准的问题。

### v2.6.0

*   **新增 设置自定义字体功能**：支持在字体选框中自定义设置字体。
*   **新增 设置附件格式功能**：支持识别并设置附件标识（附件、附件1、附件一等）、附件标题。

### v2.5.9

*   **优化 UI界面**：调整了窗口布局和组件尺寸，更适应宽屏显示器。

### v2.5.8

*   **新增 拖拽功能**：支持从文件管理器将单个或多个文件、文件夹直接拖拽到文件列表框中，简化了文件添加流程。
*   **优化 图片/附件段落处理**：重构了对包含图片、图形或附件的段落的处理逻辑。现在，程序会为这些段落中的文字应用正确的标题或正文格式（字体、字号），同时保留段落原有的对齐和缩进等布局，确保图文格式的统一性。
*   **优化 列表框体验**：当文件列表为空时，会显示“可以拖拽...”的文字提示。

### v2.5.7

*   **新增 二级标题段内智能拆分**：当二级标题与正文在同一个段落时，能够自动为标题和正文部分应用不同格式，而无需手动拆分段落。
*   **新增 配置方案管理**：增加了“保存为默认配置”功能，用户可以将当前设置保存，软件下次启动时将自动加载。
*   **优化 UI界面**：增加了“移除文件”按钮和处理前的重要安全提示。

## 许可证 (License)

本项目采用 [MIT License](LICENSE) 授权。

## Star History

<a href="https://www.star-history.com/?repos=cwyalpha%2FWord-Formatter-Pro&type=date&legend=top-left">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/chart?repos=cwyalpha/Word-Formatter-Pro&type=date&theme=dark&legend=top-left&sealed_token=gn65Nc67QWlWJxKbBGxjwVv9dEmnvkpismHIH9aLunIrha9mNHRszm0XxVld_tL5SAvNL-bTcl78Br5WTykGxEWYgxbKKOQm2zheOza8O_oBRpHo3ShIntiKJH4Jp7uQAQ93g0SW5BZkBK_91_hBvHm8LlRKI5XKatndgF76iNMWLh2N_tmpHW2x5C5J" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/chart?repos=cwyalpha/Word-Formatter-Pro&type=date&legend=top-left&sealed_token=gn65Nc67QWlWJxKbBGxjwVv9dEmnvkpismHIH9aLunIrha9mNHRszm0XxVld_tL5SAvNL-bTcl78Br5WTykGxEWYgxbKKOQm2zheOza8O_oBRpHo3ShIntiKJH4Jp7uQAQ93g0SW5BZkBK_91_hBvHm8LlRKI5XKatndgF76iNMWLh2N_tmpHW2x5C5J" />
   <img alt="Star History Chart" src="https://api.star-history.com/chart?repos=cwyalpha/Word-Formatter-Pro&type=date&legend=top-left&sealed_token=gn65Nc67QWlWJxKbBGxjwVv9dEmnvkpismHIH9aLunIrha9mNHRszm0XxVld_tL5SAvNL-bTcl78Br5WTykGxEWYgxbKKOQm2zheOza8O_oBRpHo3ShIntiKJH4Jp7uQAQ93g0SW5BZkBK_91_hBvHm8LlRKI5XKatndgF76iNMWLh2N_tmpHW2x5C5J" />
 </picture>
</a>
