# PPT 批量加水印保护工具 (Windows专版)

本项目是基于 `win32com` 和 `Pillow` 的图形界面工具，可以将 PPT 批量导出为 1080p 图片，打上固定在右下角的可缩放水印，然后重新拼接为纯图片版（扁平化、不可编辑）的 PPT。

## ⚠️ 运行环境要求
由于本程序需要直接调用 Microsoft PowerPoint 的 COM 接口，**必须在 Windows 操作系统下运行，并且该系统必须安装有 Microsoft Office (包含 PowerPoint)**。

## 🚀 安装与运行指南

1. **迁移到 Windows：**
   将本文件夹 (`ppt_watermark_tool`) 完整拷贝或同步到您的 Windows 电脑上。

2. **安装 Python：**
   确保您的 Windows 电脑上已安装 Python 3.8 或更高版本。

3. **安装依赖库：**
   打开命令行 (CMD) 或 PowerShell，进入本文件夹目录，执行以下命令安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
   *(如果下载慢，可以加清华源: `pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple`)*

4. **启动程序：**
   在文件夹下双击运行 `main_gui.py`，或在命令行中执行：
   ```bash
   python main_gui.py
   ```

## 🛠️ 功能特性
- **自动适配比例：** 自动检测源文件比例，如果是 16:9 导出为 `1920x1080`，如果接近 4:3 则导出为 `1440x1080`。
- **透明水印：** 建议使用 `.png` 格式的透明背景水印，默认会有 20 像素的安全边距，防止紧贴死角。
- **防覆盖保护：** 如果您选择的输出目录和源文件目录一致，程序会自动在输出文件名后加上 `_wm` 的后缀（如 `test_wm.pptx`）。
- **后台异步处理：** 处理大文件时界面不会卡死，提供实时进度条和状态提示。
