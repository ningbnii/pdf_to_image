# PDF 转图片工具

一个使用 Python 开发的简单 GUI 工具，用于将 PDF 文件转换为图片。它提供了直观的操作界面，让用户能够轻松地选择 PDF 文件和保存位置，将 PDF 的每一页转换为高质量的 PNG 图像。

## 功能特点

- 简洁直观的图形用户界面
- 支持选择任意 PDF 文件进行转换
- 可自定义图片保存位置，也可使用默认位置
- 将 PDF 的每一页转换为单独的 PNG 图像
- 转换后的图片默认保存在以 PDF 文件名命名的文件夹中
- 转换过程中实时显示状态提示
- 一键式操作，简单易用

## 安装说明

### 方法一：直接使用可执行文件（推荐）

1. 从`dist`目录下载`PDFToImageTool.exe`文件
2. 双击运行即可，无需安装任何依赖

### 方法二：从源代码运行

如果您想要从源代码运行或二次开发，请按以下步骤操作：

1. 确保您已安装 Python 3.9+
2. 克隆或下载本仓库
3. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
4. 运行程序：
   ```bash
   python pdf_to_image_converter.py
   ```

## 使用指南

1. 启动程序
2. 点击"选择 PDF"按钮，浏览并选择要转换的 PDF 文件
3. （可选）点击"选择目录"按钮，指定图片保存位置
   - 如果不选择，默认将在 PDF 文件所在目录下创建文件夹
4. 点击"开始转换"按钮，开始转换过程
5. 转换完成后，会显示成功提示和图片保存位置
6. 打开保存位置查看转换后的图片

## 开发环境

- Python 3.13.2
- PyMuPDF 1.25.5（用于 PDF 处理）
- Tkinter（GUI 界面）
- Windows 10

## 打包说明

本项目使用 PyInstaller 将 Python 脚本打包为独立的 Windows 可执行文件。如果您想自行打包，请执行以下步骤：

1. 安装 PyInstaller：

   ```bash
   pip install pyinstaller
   ```

2. 执行打包命令：

   ```bash
   pyinstaller --name PDFToImageTool --onefile --windowed pdf_to_image_converter.py
   ```

3. 打包完成后，可执行文件将位于`dist`目录中

## 依赖库

- PyMuPDF (fitz)：用于 PDF 处理和转换
- Tkinter：用于创建图形用户界面

## 注意事项

- 转换大型 PDF 文件可能需要较长时间，请耐心等待
- 分辨率设置为 2.0 倍，可以获得较高质量的图像输出
- 所有图片以 PNG 格式保存，文件名格式为"page_1.png"、"page_2.png"等
