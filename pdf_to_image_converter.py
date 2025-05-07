import tkinter as tk
from tkinter import filedialog, messagebox
import os
# from pdf2image import convert_from_path # Remove this line
import fitz  # Add PyMuPDF (fitz)
from pathlib import Path


class PDFConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF 转图片工具")
        master.geometry("500x300")  # 增加窗口高度以容纳新控件

        self.pdf_path = ""
        self.save_dir = ""  # 添加保存目录变量

        # PDF选择部分
        self.pdf_frame = tk.Frame(master)
        self.pdf_frame.pack(pady=10, fill=tk.X, padx=20)

        self.label = tk.Label(self.pdf_frame, text="请选择一个 PDF 文件")
        self.label.pack(side=tk.LEFT, padx=(0, 10))

        self.select_button = tk.Button(
            self.pdf_frame, text="选择 PDF", command=self.select_pdf)
        self.select_button.pack(side=tk.RIGHT)

        # 保存目录选择部分
        self.save_frame = tk.Frame(master)
        self.save_frame.pack(pady=10, fill=tk.X, padx=20)

        self.save_label = tk.Label(
            self.save_frame, text="保存目录 (可选，默认为PDF所在目录)")
        self.save_label.pack(side=tk.LEFT, padx=(0, 10))

        self.save_button = tk.Button(
            self.save_frame, text="选择目录", command=self.select_save_dir)
        self.save_button.pack(side=tk.RIGHT)

        # 转换按钮
        self.convert_button = tk.Button(
            master, text="开始转换", command=self.convert_pdf, state=tk.DISABLED)
        self.convert_button.pack(pady=15)

        # 状态标签
        self.status_label = tk.Label(master, text="")
        self.status_label.pack(pady=10)

    def select_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=(("PDF 文件", "*.pdf"), ("所有文件", "*.*"))
        )
        if file_path:
            self.pdf_path = file_path
            self.label.config(text=f"已选择: {os.path.basename(file_path)}")
            self.convert_button.config(state=tk.NORMAL)
            self.status_label.config(text="")
        else:
            self.label.config(text="未选择 PDF 文件")
            self.convert_button.config(state=tk.DISABLED)

    def select_save_dir(self):
        dir_path = filedialog.askdirectory(title="选择保存目录")
        if dir_path:
            self.save_dir = dir_path
            self.save_label.config(text=f"保存到: {dir_path}")
        else:
            self.save_dir = ""
            self.save_label.config(text="保存目录 (可选，默认为PDF所在目录)")

    def convert_pdf(self):
        if not self.pdf_path:
            messagebox.showerror("错误", "请先选择一个 PDF 文件。")
            return

        self.status_label.config(text="转换中，请稍候...")
        self.master.update_idletasks()  # Update GUI before long task

        try:
            pdf_document_path = Path(self.pdf_path)
            pdf_name = pdf_document_path.stem  # PDF 文件名（不含扩展名）

            # 确定输出目录
            if self.save_dir:
                # 如果用户选择了保存目录，则在该目录下创建以PDF名称命名的文件夹
                output_dir = Path(self.save_dir) / pdf_name
            else:
                # 否则按照原来的逻辑，在PDF所在目录创建同名文件夹
                output_dir = pdf_document_path.parent / pdf_name

            if not output_dir.exists():
                output_dir.mkdir(parents=True)

            # images = convert_from_path(self.pdf_path) # Remove this line

            # --- New logic using PyMuPDF (fitz) ---
            doc = fitz.open(self.pdf_path)
            # zoom_x and zoom_y control the resolution. Default is 1.0 (72 dpi).
            # Using 2.0 for 144 dpi, similar to your reference for better quality.
            zoom_x = 2.0  # horizontal zoom
            zoom_y = 2.0  # vertical zoom
            # zoom factor 2 in each dimension
            mat = fitz.Matrix(zoom_x, zoom_y)

            for page_num in range(len(doc)):
                page = doc.load_page(page_num)  # load page
                pix = page.get_pixmap(matrix=mat)  # render page to an image
                image_path = output_dir / f"page_{page_num + 1}.png"
                pix.save(str(image_path))  # save image as png

            doc.close()
            # --- End of new logic ---

            # for i, image in enumerate(images): # Remove this loop
            #     image_path = output_dir / f"page_{i+1}.png"
            #     image.save(image_path, "PNG")

            messagebox.showinfo("成功", f"转换完成！图片已保存到: \n{output_dir}")
            self.status_label.config(text="转换完成！")

        except Exception as e:
            messagebox.showerror("转换失败", f"发生错误: {e}")
            self.status_label.config(text="转换失败。")
        finally:
            self.convert_button.config(state=tk.NORMAL)  # Re-enable button
            # Optionally reset the selection
            # self.pdf_path = ""
            # self.label.config(text="请选择一个 PDF 文件")
            # self.convert_button.config(state=tk.DISABLED)


def main():
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
