import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import json
from pathlib import Path

# 导入PyMuPDF，并处理可能的API差异
try:
    import fitz
except ImportError:
    messagebox.showerror("错误", "缺少 PyMuPDF 库。请安装: pip install PyMuPDF")
    sys.exit(1)

# 修改配置文件路径，使用用户的AppData目录


def get_config_path():
    """返回配置文件的路径，确保在打包后也能使用"""
    try:
        # 使用系统的应用数据目录
        app_data = os.environ.get('APPDATA')
        if app_data:
            # Windows系统
            app_dir = os.path.join(app_data, "PDFToImageTool")
            if not os.path.exists(app_dir):
                os.makedirs(app_dir)
            return os.path.join(app_dir, "config.json")
        else:
            # 非Windows系统或无法获取APPDATA
            return os.path.join(os.path.expanduser("~"), ".pdf_to_image_config.json")
    except Exception as e:
        print(f"无法确定配置文件路径: {e}")
        # 使用当前目录作为后备方案
        return "pdf_converter_config.json"


# 配置文件路径
CONFIG_FILE = get_config_path()


class PDFConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF 转图片工具")
        master.geometry("600x400")  # 增加窗口尺寸

        # 设置窗口最小尺寸，防止元素被挤压
        master.minsize(600, 400)

        # 配置样式
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("微软雅黑", 10))
        self.style.configure("TButton", font=("微软雅黑", 10))
        self.style.configure("TEntry", font=("微软雅黑", 10))

        # 默认配置
        self.config = {
            "save_dir": "",
            "zoom_factor": "2.0"
        }

        # 加载上次的配置
        self.load_config()

        self.pdf_path = ""
        self.save_dir = self.config["save_dir"]  # 使用加载的保存目录

        # 创建主框架
        self.main_frame = ttk.Frame(master, padding="20 20 20 10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 使用网格布局管理整个界面
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=3)
        self.main_frame.columnconfigure(2, weight=1)

        # PDF选择部分
        ttk.Label(self.main_frame, text="请选择PDF文件:", anchor="e", width=15).grid(
            row=0, column=0, sticky="e", padx=(0, 10), pady=10)

        # 文件名显示
        self.filename_var = tk.StringVar()
        self.filename_entry = ttk.Entry(self.main_frame, textvariable=self.filename_var,
                                        state="readonly")
        self.filename_entry.grid(row=0, column=1, sticky="ew", pady=10)

        self.select_button = ttk.Button(self.main_frame, text="选择PDF",
                                        command=self.select_pdf, width=10)
        self.select_button.grid(
            row=0, column=2, sticky="w", padx=(10, 0), pady=10)

        # 保存目录部分
        ttk.Label(self.main_frame, text="保存目录:", anchor="e", width=15).grid(
            row=1, column=0, sticky="e", padx=(0, 10), pady=10)

        # 保存目录显示
        self.save_dir_var = tk.StringVar()
        self.save_dir_var.set(
            "未选择，将使用PDF所在目录" if not self.save_dir else self.save_dir)
        self.save_dir_entry = ttk.Entry(self.main_frame, textvariable=self.save_dir_var,
                                        state="readonly")
        self.save_dir_entry.grid(row=1, column=1, sticky="ew", pady=10)

        self.save_button = ttk.Button(self.main_frame, text="选择目录",
                                      command=self.select_save_dir, width=10)
        self.save_button.grid(row=1, column=2, sticky="w",
                              padx=(10, 0), pady=10)

        # 缩放比例选择部分
        ttk.Label(self.main_frame, text="缩放比例:", anchor="e", width=15).grid(
            row=2, column=0, sticky="e", padx=(0, 10), pady=10)

        # 缩放控件框架
        zoom_frame = ttk.Frame(self.main_frame)
        zoom_frame.grid(row=2, column=1, sticky="w", pady=10)

        # 创建缩放比例选择
        self.zoom_value = tk.StringVar(master)
        self.zoom_value.set(self.config["zoom_factor"])  # 使用加载的缩放比例

        # 缩放比例下拉菜单
        zoom_options = ["1.0", "1.5", "2.0", "3.0", "4.0", "5.0"]
        self.zoom_menu = ttk.Combobox(zoom_frame, textvariable=self.zoom_value,
                                      values=zoom_options, width=5, state="readonly")
        self.zoom_menu.pack(side=tk.LEFT)

        # 添加说明标签
        ttk.Label(zoom_frame, text="(值越大图片越清晰，但文件也会相应增大)").pack(
            side=tk.LEFT, padx=(10, 0))

        # 转换按钮 - 居中放置
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20)

        self.convert_button = ttk.Button(button_frame, text="开始转换",
                                         command=self.convert_pdf,
                                         state=tk.DISABLED, width=15)
        self.convert_button.pack()

        # 状态栏
        status_frame = ttk.Frame(self.main_frame)
        status_frame.grid(row=4, column=0, columnspan=3,
                          sticky="ew", pady=(10, 0))

        # 添加进度条
        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(status_frame, orient="horizontal",
                                            length=100, mode="determinate",
                                            variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X)

        # 添加窗口关闭事件处理
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def show_config_path(self):
        """用于调试的函数，显示配置文件的路径"""
        config_exists = os.path.exists(CONFIG_FILE)
        msg = f"配置文件路径: {CONFIG_FILE}\n存在: {'是' if config_exists else '否'}"
        if config_exists:
            try:
                with open(CONFIG_FILE, 'r') as f:
                    content = f.read()
                msg += f"\n内容: {content}"
            except Exception as e:
                msg += f"\n无法读取内容: {e}"
        messagebox.showinfo("配置信息", msg)

    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    loaded_config = json.load(f)
                    # 更新配置，只使用有效的键
                    for key in self.config:
                        if key in loaded_config:
                            self.config[key] = loaded_config[key]
                print(f"成功加载配置: {self.config}")
            else:
                print(f"配置文件不存在: {CONFIG_FILE}")
        except Exception as e:
            print(f"加载配置失败: {e}")
            try:
                # 尝试保存默认配置来测试文件权限
                self.save_config()
                print("已创建新的配置文件")
            except Exception as e2:
                print(f"无法创建配置文件: {e2}")

    def save_config(self):
        """保存配置到文件"""
        try:
            # 更新配置
            self.config["save_dir"] = self.save_dir
            self.config["zoom_factor"] = self.zoom_value.get()

            # 确保目录存在
            config_dir = os.path.dirname(CONFIG_FILE)
            if config_dir and not os.path.exists(config_dir):
                os.makedirs(config_dir)

            # 保存到文件
            with open(CONFIG_FILE, 'w') as f:
                json.dump(self.config, f)
            print(f"配置已保存到: {CONFIG_FILE}")
        except Exception as e:
            print(f"保存配置失败: {e}")
            messagebox.showwarning("提示", f"无法保存配置: {e}")

    def on_closing(self):
        """窗口关闭时的处理"""
        self.save_config()
        self.master.destroy()

    def select_pdf(self):
        """选择PDF文件"""
        filepath = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if not filepath:
            return

        self.pdf_path = filepath
        filename = os.path.basename(filepath)
        self.filename_var.set(filename)  # 设置文件名显示
        self.convert_button.config(state=tk.NORMAL)

    def select_save_dir(self):
        """选择保存目录"""
        directory = filedialog.askdirectory(title="选择保存目录")
        if not directory:
            return

        self.save_dir = directory
        self.save_dir_var.set(directory)  # 更新目录显示

    def convert_pdf(self):
        """将PDF转换为图片"""
        if not self.pdf_path:
            messagebox.showerror("错误", "请先选择一个PDF文件")
            return

        try:
            # 获取缩放因子
            zoom_factor = float(self.zoom_value.get())

            # 打开PDF文件
            pdf_document = fitz.open(self.pdf_path)
            total_pages = pdf_document.page_count

            # 重置进度条
            self.progress_var.set(0)

            # 更新窗口标题显示状态
            pdf_basename = os.path.basename(self.pdf_path)
            self.master.title(f"PDF 转图片工具 - 准备转换 {pdf_basename}")

            # 保存原始选择的目录（用于配置保存）
            original_save_dir = self.save_dir

            # 确定实际的输出目录（包含PDF文件名的子文件夹）
            output_dir = ""
            pdf_name = os.path.splitext(os.path.basename(self.pdf_path))[0]

            if not self.save_dir:
                # 如果没有指定保存目录，使用PDF所在目录
                pdf_dir = os.path.dirname(self.pdf_path)
                output_dir = os.path.join(pdf_dir, f"{pdf_name}_images")
            else:
                # 在选定的目录中创建以PDF文件名命名的子文件夹
                output_dir = os.path.join(self.save_dir, f"{pdf_name}_images")

            # 创建保存目录（如果不存在）
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # 开始转换
            self.master.title(f"PDF 转图片工具 - 转换中")
            self.master.update()

            # 转换每一页
            for page_num in range(total_pages):
                # 获取页面
                page = pdf_document[page_num]

                try:
                    # 方法1：使用DPI参数（适用于较新版本的PyMuPDF）
                    dpi = 72 * zoom_factor  # 标准PDF是72 DPI
                    pix = page.get_pixmap(dpi=dpi)
                except TypeError:
                    # 方法2：如果DPI参数不可用，使用Matrix但计算正确的缩放比例
                    # 保持相同尺寸，但增加每英寸像素数
                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(zoom_factor, zoom_factor))

                # 保存图像
                output_file = os.path.join(
                    output_dir, f"page_{page_num+1}.png")
                pix.save(output_file)

                # 更新进度条
                progress = int((page_num + 1) / total_pages * 100)
                self.progress_var.set(progress)

                # 更新窗口标题显示进度
                self.master.title(
                    f"PDF 转图片工具 - 转换中 {page_num+1}/{total_pages} ({progress}%)")
                self.master.update()

            # 关闭PDF
            pdf_document.close()

            # 确保进度条显示100%
            self.progress_var.set(100)

            # 成功转换后更新标题
            self.master.title(f"PDF 转图片工具 - 转换完成")

            # 询问是否打开输出文件夹
            if messagebox.askyesno("完成", f"所有页面已转换为图片。\n是否打开输出文件夹？"):
                # 根据操作系统选择打开文件夹的方式
                try:
                    if sys.platform == 'win32':
                        os.startfile(output_dir)
                    elif sys.platform == 'darwin':  # macOS
                        os.system(f'open "{output_dir}"')
                    else:  # Linux
                        os.system(f'xdg-open "{output_dir}"')
                except Exception as e:
                    print(f"无法打开文件夹: {e}")
                    messagebox.showinfo("提示", f"请手动打开文件夹：{output_dir}")

            # 还原窗口标题
            self.master.title("PDF 转图片工具")

            # 还原原始保存目录（确保配置中保存的是用户选择的目录）
            self.save_dir = original_save_dir

            # 恢复界面上显示的原始保存目录
            if original_save_dir:
                self.save_dir_var.set(original_save_dir)
            else:
                self.save_dir_var.set("未选择，将使用PDF所在目录")

            # 保存配置以记住用户的首选项
            self.save_config()

        except Exception as e:
            # 重置进度条
            self.progress_var.set(0)

            # 还原窗口标题
            self.master.title("PDF 转图片工具")

            # 出错时显示详细的错误信息
            error_msg = f"转换失败: {type(e).__name__}: {e}"
            print(error_msg)
            messagebox.showerror("错误", error_msg)


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = PDFConverterApp(root)
        root.mainloop()
    except Exception as e:
        # 捕获任何未处理的异常
        error_msg = f"程序遇到未处理的异常: {type(e).__name__}: {e}"
        print(error_msg)
        messagebox.showerror("严重错误", error_msg)
        sys.exit(1)
