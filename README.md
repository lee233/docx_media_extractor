
# Docx Media Extractor | DOCX 媒体提取器

A simple Python script to extract all media files (images, videos, etc.) from a `.docx` file and save them in the same directory with sequential naming.  
一个简单的 Python 脚本，用于从 `.docx` 文件中提取所有嵌入的媒体文件（如图片、视频等），并按顺序命名后保存到相同目录下。

---

## Features | 功能

- Extracts all media files embedded in a `.docx` file (e.g., `JPEG`, `PNG`, `GIF`).
- 从 `.docx` 文件中提取所有嵌入的媒体文件（例如 `JPEG`、`PNG`、`GIF` 等）。
  
- Saves extracted files in a new folder with the same name as the document.
- 将提取的文件保存在一个与文档同名的新文件夹中。
  
- Automatically names files as `1.jpeg`, `2.png`, etc., in the order they appear in the document.
- 按照文件在文档中出现的顺序，自动命名为 `1.jpeg`、`2.png` 等。

- Easy to use via drag-and-drop or command-line.
- 支持拖拽和命令行两种运行方式，使用简单方便。

---

## Requirements | 环境要求

- Python 3.x  

- Libraries: `xml.etree.ElementTree` (built-in), `shutil`, `zipfile`.  
- 所需库：`xml.etree.ElementTree`（内置），`shutil`，`zipfile`。

---

## Installation | 安装

### Step 1: Clone this repository | 第一步：克隆此项目

```bash
git clone https://github.com/your-username/docx-media-extractor.git
cd docx-media-extractor
```

### Step 2: Install required libraries | 第二步：安装依赖库（可选）

```bash
pip install -r requirements.txt
```

如果使用内置库（如默认代码），无需安装任何库。

---

## Usage | 使用方法

### Option 1: Drag and Drop | 方法一：拖拽运行

1. Drag a `.docx` file onto the `extract_docx_media.py` script.  
1. 将 `.docx` 文件拖拽到 `extract_docx_media.py` 脚本上运行。
   
2.The script will automatically extract media files to a folder named after the document.  
2.脚本会自动提取媒体文件并保存到一个与文档同名的文件夹中。

---

### Option 2: Command Line | 方法二：命令行运行

1.Open a terminal or command prompt.  
1.打开终端或命令提示符。

2.Run the script with the following command:  
2.运行以下命令：

```bash
python extract_docx_media.py "path_to_your_docx_file.docx"
```

3.Replace `path_to_your_docx_file.docx` with the full path to your `.docx` file.  
3.将 `path_to_your_docx_file.docx` 替换为您的 `.docx` 文件的完整路径。

---

### Option 3: Compiled `.exe` | 方法三：编译为 `.exe`

If you'd like to distribute this script without requiring Python, compile it using PyInstaller:  
如果希望分发该脚本且不需要用户安装 Python 环境，可以使用 PyInstaller 将其打包为 `.exe`：

```bash
pyinstaller --onefile --noconsole extract_docx_media.py
```

---

## Output Example | 输出示例

For a `.docx` file named `Report.docx`, the output folder structure will look like this:  
假设 `.docx` 文件名为 `Report.docx`，输出的文件夹结构如下：

```
Report/
├── 1.jpeg
├── 2.png
├── 3.gif
```

---


## License | 许可证

This project is licensed under the MIT License. See the `LICENSE` file for details.  
本项目基于 MIT 许可证开源，详情请查看 `LICENSE` 文件。


