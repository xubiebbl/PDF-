import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import configparser
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
import pypdf
from PIL import Image
import fitz
import pdf_tools_ico
# 导入通用设置管理器
from pdf_tools_common import CommonSettingsManager
import tempfile

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("效率猿-PDFTools")
        self.root.geometry("800x600")
        # 图标
        #  """从二进制数据设置窗口图标"""
        # 创建临时文件
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, f"temp_icon_{os.getpid()}.ico")

        # 写入二进制数据
        with open(temp_file, 'wb') as f:
            f.write(pdf_tools_ico.ico())
        self.root.iconbitmap(temp_file)

        # # notebook 颜色
        # self.colors = {
        #     "bg": "#f0f2f5",  # 主背景色 (浅灰蓝)
        #     "frame_bg": "#ffffff",  # 内容框架背景色 (纯白)
        #     "tab_bg": "#e4e7eb",  # 未选中标签背景色
        #     "tab_active_bg": "#ffffff",  # 选中标签背景色
        #     "tab_fg": "#2c3e50",  # 标签文字颜色 (深蓝灰)
        #     "tab_border": "#d1d9e6",  # 边框色
        #     "accent": "#3498db",  # 强调色 (蓝色)
        # }
        # 设置默认字体为11号黑体
        self.default_font = ("宋体", 10)

        # 设置默认字体为仿宋,9号
        self.default_font_i = ("仿宋", 9)

        # 配置根窗口的默认字体
        self.root.option_add("*Font", self.default_font)

        # 存储选中的文件路径
        self.selected_files = []

        # 获取实际的桌面路径
        self.desktop_path = self.get_desktop_path()

        # 配置文件路径
        self.config_file = "config.ini"
        self.load_config()

        # 创建通用设置管理器
        self.settings_manager = CommonSettingsManager(self)

        # 创建样式以设置标签在左侧
        style = ttk.Style()
        style.configure(
            "LeftTab.TNotebook",
            tabposition="wn",
        )
        # 创建Notebook组件并应用样式
        self.notebook = ttk.Notebook(root, style="LeftTab.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建合并页面
        self.merge_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.merge_frame, text="合并")

        # 创建拆分页面
        self.split_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.split_frame, text="拆分")

        # 创建插入页面
        self.insert_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.insert_frame, text="插入")

        # 创建替换页面
        self.replace_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.replace_frame, text="替换")

        # 创建加密解密页面
        self.encrypt_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.encrypt_frame, text="加密解密")

        # 创建加水印页面
        self.watermark_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.watermark_frame, text="加水印")

        # 创建图片提取页面
        self.extract_image_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.extract_image_frame, text="图片提取")

        # 初始化合并页面内容
        self.setup_merge_page()

        # 初始化拆分页面内容
        self.setup_split_page()

        # 初始化插入页面内容
        self.setup_insert_page()

        # 初始化替换页面内容
        self.setup_replace_page()

        # 初始化加密解密页面内容
        self.setup_encrypt_page()

        # 初始化加水印页面内容
        self.setup_watermark_page()

        # 初始化图片提取页面内容
        self.setup_extract_image_page()

    def get_desktop_path(self):
        """获取系统实际的桌面路径"""
        try:
            # Windows系统获取桌面路径
            import winreg

            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
            )
            desktop_path = winreg.QueryValueEx(key, "Desktop")[0]
            winreg.CloseKey(key)
            return desktop_path
        except:
            # 如果无法通过注册表获取，则使用默认路径
            return os.path.join(os.path.expanduser("~"), "Desktop")

    def load_config(self):
        """加载配置文件"""
        self.config = configparser.ConfigParser()
        self.config.read(self.config_file, encoding="utf-8")

        # 如果配置文件不存在，创建默认配置
        if not self.config.has_section("MergeSettings"):
            self.config.add_section("MergeSettings")
            self.config.set(
                "MergeSettings", "save_location_option", "desktop"
            )  # desktop or custom
            self.config.set("MergeSettings", "custom_save_location", "")
            self.config.set(
                "MergeSettings", "filename_option", "default"
            )  # default or custom
            self.config.set("MergeSettings", "custom_filename", "")
            self.save_config()

    def save_config(self):
        """保存配置到文件"""
        with open(self.config_file, "w", encoding="utf-8") as configfile:
            self.config.write(configfile)

    def setup_merge_page(self):
        """设置合并页面的内容"""
        # 添加选择文件的按钮
        select_btn = ttk.Button(
            self.merge_frame, text="选择要合并的PDF文件", command=self.select_files
        )
        select_btn.pack(pady=10)

        # 创建文件列表显示区域
        list_frame = ttk.Frame(self.merge_frame)
        list_frame.pack(pady=5, padx=20, fill=tk.BOTH, expand=True)

        # 创建列表框和滚动条
        self.file_listbox = tk.Listbox(list_frame, height=6)
        scrollbar = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview
        )
        self.file_listbox.configure(yscrollcommand=scrollbar.set)

        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 创建存储位置设置
        self.merge_storage_widgets = self.settings_manager.create_storage_settings(
            self.merge_frame,
            "save_location_var",
            "folder_path_var",
            "select_folder_btn",
            "folder_path_entry",
            self.on_save_location_change,
        )

        # 配置选择文件夹按钮的命令
        self.merge_storage_widgets["select_btn"].config(
            command=lambda: self.settings_manager.select_save_folder(
                self.merge_storage_widgets["folder_path_var"]
            )
        )

        # 创建文件名设置
        self.merge_filename_widgets = self.settings_manager.create_filename_settings(
            self.merge_frame, "filename_var", "filename_entry", self.on_filename_change
        )

        # 设置默认文件名
        if self.merge_filename_widgets["filename_var"].get() == "default":
            self.merge_filename_widgets["entry"].config(state="normal")
            default_name = datetime.now().strftime("%Y%m%d%H%M%S")
            self.merge_filename_widgets["entry"].delete(0, tk.END)
            self.merge_filename_widgets["entry"].insert(0, default_name)
            self.merge_filename_widgets["entry"].config(state="readonly")

        # 添加开始合并按钮
        merge_btn = ttk.Button(
            self.merge_filename_widgets["frame"],
            text="开始合并",
            command=self.merge_pdfs,
        )
        merge_btn.pack(pady=5)

    def on_save_location_change(self, *args):
        """存储位置选项改变时的处理"""
        # 保存设置到配置文件
        self.config.set(
            "MergeSettings",
            "save_location_option",
            self.merge_storage_widgets["location_var"].get(),
        )
        self.save_config()

    def on_filename_change(self, *args):
        """文件名选项改变时的处理"""
        # 保存设置到配置文件
        self.config.set(
            "MergeSettings",
            "filename_option",
            self.merge_filename_widgets["filename_var"].get(),
        )
        self.save_config()

        # 根据选择设置默认文件名
        if self.merge_filename_widgets["filename_var"].get() == "default":
            self.merge_filename_widgets["entry"].config(state="normal")
            default_name = datetime.now().strftime("%Y%m%d%H%M%S")
            self.merge_filename_widgets["entry"].delete(0, tk.END)
            self.merge_filename_widgets["entry"].insert(0, default_name)
            self.merge_filename_widgets["entry"].config(state="readonly")

    def select_files(self):
        """选择多个PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择要合并的PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        )

        if files:
            # 清空之前的文件列表
            self.file_listbox.delete(0, tk.END)
            self.selected_files.clear()

            # 添加新选择的文件
            for file_path in files:
                self.selected_files.append(file_path)
                # 只显示文件名，不显示完整路径
                try:
                    # 获取PDF页数
                    reader = pypdf.PdfReader(file_path)
                    page_count = len(reader.pages)
                    file_name = f"{page_count}页 -> {os.path.basename(file_path)} "
                except Exception as e:
                    # 如果无法读取页数，只显示文件名
                    file_name = f"{os.path.basename(file_path)} (无法读取页数)"
                self.file_listbox.insert(tk.END, file_name)

            # 更新状态信息
            print(f"已选择 {len(self.selected_files)} 个文件")

    def merge_pdfs(self):
        """合并PDF文件"""
        # 检查是否有选择文件
        if not self.selected_files:
            messagebox.showwarning("警告", "请先选择要合并的PDF文件！")
            return

        # 确定保存路径
        save_directory = self.settings_manager.get_save_directory(
            self.merge_storage_widgets["location_var"],
            self.merge_storage_widgets["folder_path_var"],
        )
        if not save_directory:
            messagebox.showerror("错误", "请选择有效的保存路径！")
            return

        # 确定文件名
        def get_default_name():
            return self.merge_filename_widgets["entry"].get()

        filename = self.settings_manager.get_filename(
            self.merge_filename_widgets["filename_var"],
            self.merge_filename_widgets["entry"],
            get_default_name,
        )
        if not filename:
            messagebox.showerror("错误", "请输入自定义文件名！")
            return

        # 完整保存路径
        output_path = os.path.join(save_directory, filename + ".pdf")

        # 检查文件是否已存在
        if os.path.exists(output_path):
            result = messagebox.askyesno(
                "文件已存在", f"文件 {filename}.pdf 已存在，是否覆盖？"
            )
            if not result:
                return

        try:
            # 创建PDF写入器
            pdf_writer = pypdf.PdfWriter()

            # 逐个读取并合并PDF文件
            for file_path in self.selected_files:
                try:
                    with open(file_path, "rb") as pdf_file:
                        pdf_reader = pypdf.PdfReader(pdf_file)
                        # 将每一页添加到写入器中
                        for page_num in range(len(pdf_reader.pages)):
                            page = pdf_reader.pages[page_num]
                            pdf_writer.add_page(page)
                except Exception as e:
                    messagebox.showerror(
                        "错误", f"读取文件 {file_path} 时出错：{str(e)}"
                    )
                    return

            # 写入合并后的PDF文件
            with open(output_path, "wb") as output_file:
                pdf_writer.write(output_file)

            messagebox.showinfo("成功", f"PDF文件合并完成！\n保存位置：{output_path}")

        except Exception as e:
            messagebox.showerror("错误", f"合并PDF文件时出错：{str(e)}")

    def setup_split_page(self):
        """设置拆分页面的内容"""
        # 添加选择文件的按钮
        select_btn = ttk.Button(
            self.split_frame, text="选择要拆分的PDF文件", command=self.select_split_file
        )
        select_btn.pack(pady=10)

        # 显示选中的文件
        self.split_file_var = tk.StringVar()
        split_file_entry = ttk.Entry(
            self.split_frame, textvariable=self.split_file_var, state="readonly"
        )
        split_file_entry.pack(pady=5, padx=20, fill=tk.X)

        # 拆分方式选择框架
        split_method_frame = ttk.LabelFrame(self.split_frame, text="拆分方式")
        split_method_frame.pack(pady=10, padx=20, fill=tk.X)

        # 创建顶部框架用于放置单选按钮
        top_frame = ttk.Frame(split_method_frame)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        # 拆分方式变量
        self.split_method_var = tk.StringVar(value="single")

        # 单页拆分
        single_radio = ttk.Radiobutton(
            top_frame,
            text="单页拆分（每页一个文件）",
            variable=self.split_method_var,
            value="single",
        )
        single_radio.pack(side=tk.LEFT, padx=(0, 20))

        # 按页数拆分
        pages_radio = ttk.Radiobutton(
            top_frame, text="按页数拆分", variable=self.split_method_var, value="pages"
        )
        pages_radio.pack(side=tk.LEFT, padx=(0, 20))

        # 按范围拆分
        range_radio = ttk.Radiobutton(
            top_frame, text="按范围拆分", variable=self.split_method_var, value="range"
        )
        range_radio.pack(side=tk.LEFT)

        # 页数设置框架
        self.pages_setting_frame = ttk.Frame(split_method_frame)
        self.pages_setting_frame.pack(fill=tk.X, padx=10, pady=(0, 5))

        # 添加占位空间将控件推到右侧
        ttk.Frame(self.pages_setting_frame).pack(side=tk.LEFT, expand=True)

        pages_inner_frame = ttk.Frame(self.pages_setting_frame)
        pages_inner_frame.pack(side=tk.LEFT)

        ttk.Label(pages_inner_frame, text="每").pack(side=tk.LEFT)
        self.pages_per_file_var = tk.StringVar(value="2")
        pages_spinbox = ttk.Spinbox(
            pages_inner_frame,
            from_=1,
            to=1000,
            width=5,
            textvariable=self.pages_per_file_var,
        )
        pages_spinbox.pack(side=tk.LEFT, padx=5)
        ttk.Label(pages_inner_frame, text="页一个文件").pack(side=tk.LEFT)

        # 范围设置框架
        self.range_setting_frame = ttk.Frame(split_method_frame)
        self.range_setting_frame.pack(fill=tk.X, padx=10, pady=5)

        # 添加占位空间将控件推到右侧
        ttk.Frame(self.range_setting_frame).pack(side=tk.LEFT, expand=True)

        range_inner_frame = ttk.Frame(self.range_setting_frame)
        range_inner_frame.pack(side=tk.RIGHT)

        ttk.Label(range_inner_frame, text="页码范围(如：1-3,5):").pack(side=tk.LEFT)
        self.range_var = tk.StringVar()
        range_entry = ttk.Entry(
            range_inner_frame, textvariable=self.range_var, width=20
        )
        range_entry.pack(side=tk.LEFT, padx=(5, 0))

        # 默认隐藏页数设置和范围设置
        self.pages_setting_frame.pack_forget()
        self.range_setting_frame.pack_forget()

        # 绑定拆分方式变化事件
        self.split_method_var.trace("w", self.on_split_method_change)

        # 创建存储位置设置
        self.split_storage_widgets = self.settings_manager.create_storage_settings(
            self.split_frame,
            "split_save_location_var",
            "split_folder_path_var",
            "select_split_folder_btn",
            "split_folder_path_entry",
            self.on_split_save_location_change,
        )

        # 配置选择文件夹按钮的命令
        self.split_storage_widgets["select_btn"].config(
            command=lambda: self.settings_manager.select_save_folder(
                self.split_storage_widgets["folder_path_var"]
            )
        )

        # 创建文件名设置
        self.split_filename_widgets = self.settings_manager.create_filename_settings(
            self.split_frame,
            "split_filename_var",
            "split_filename_entry",
            self.on_split_filename_change,
        )

        # 添加开始拆分按钮
        split_btn = ttk.Button(
            self.split_filename_widgets["frame"],
            text="开始拆分",
            command=self.split_pdf,
        )
        split_btn.pack(pady=5)

    def select_split_file(self):
        """选择要拆分的PDF文件"""
        file = filedialog.askopenfilename(
            title="选择要拆分的PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        )

        if file:
            reader = pypdf.PdfReader(file)
            page_count = len(reader.pages)
            self.filename_a = file
            self.split_file_var.set(f"{page_count}页 -> {os.path.basename(file)}")

    def on_split_method_change(self, *args):
        """拆分方式改变时的处理"""
        method = self.split_method_var.get()

        # 根据选择显示相应的设置选项
        if method == "pages":
            self.pages_setting_frame.pack(fill=tk.X, padx=10, pady=5)
            self.range_setting_frame.pack_forget()
        elif method == "range":
            self.range_setting_frame.pack(fill=tk.X, padx=10, pady=5)
            self.pages_setting_frame.pack_forget()
        else:  # single
            self.pages_setting_frame.pack_forget()
            self.range_setting_frame.pack_forget()

    def on_split_save_location_change(self, *args):
        """拆分存储位置选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.split_storage_widgets["location_var"].get()
        self.config.set("MergeSettings", "save_location_option", current_value)
        self.save_config()

    def on_split_filename_change(self, *args):
        """拆分文件名选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.split_filename_widgets["filename_var"].get()
        self.config.set("MergeSettings", "filename_option", current_value)
        self.save_config()

    def split_pdf(self):
        """拆分PDF文件"""
        # 检查是否选择了文件
        input_file = getattr(self, "filename_a", None)
        if not input_file:
            messagebox.showwarning("警告", "请先选择要拆分的PDF文件！")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("错误", "选择的文件不存在！")
            return

        try:
            # 确定保存路径
            save_directory = self.settings_manager.get_save_directory(
                self.split_storage_widgets["location_var"],
                self.split_storage_widgets["folder_path_var"],
            )
            if not save_directory:
                messagebox.showerror("错误", "请选择有效的保存路径！")
                return

            # 获取输入文件名（不含扩展名）
            input_filename = os.path.splitext(os.path.basename(input_file))[0]

            # 读取PDF文件
            with open(input_file, "rb") as f:
                reader = pypdf.PdfReader(f)
                total_pages = len(reader.pages)

                # 根据拆分方式执行拆分
                method = self.split_method_var.get()

                if method == "single":
                    # 单页拆分
                    self._split_single_page(reader, save_directory, input_filename)
                elif method == "pages":
                    # 按页数拆分
                    pages_per_file = int(self.pages_per_file_var.get())
                    self._split_by_pages(
                        reader, save_directory, input_filename, pages_per_file
                    )
                elif method == "range":
                    # 按范围拆分
                    range_str = self.range_var.get()
                    if not range_str:
                        messagebox.showerror("错误", "请输入页码范围！")
                        return
                    self._split_by_range(
                        reader, save_directory, input_filename, range_str
                    )

            messagebox.showinfo("成功", f"PDF拆分完成！\n保存位置：{save_directory}")

        except Exception as e:
            messagebox.showerror("错误", f"拆分PDF时发生错误：{str(e)}")

    def _split_single_page(self, reader, save_directory, input_filename):
        """单页拆分"""
        total_pages = len(reader.pages)
        for i in range(total_pages):
            writer = pypdf.PdfWriter()
            writer.add_page(reader.pages[i])

            # 生成文件名
            if self.split_filename_widgets["filename_var"].get() == "default":
                output_filename = (
                    f"{datetime.now().strftime('%Y%m%d%H%M%S')}_page_{i+1}.pdf"
                )
            else:
                custom_name = self.split_filename_widgets["entry"].get()
                if custom_name:
                    output_filename = f"{custom_name}_page_{i+1}.pdf"
                else:
                    output_filename = (
                        f"{datetime.now().strftime('%Y%m%d%H%M%S')}_page_{i+1}.pdf"
                    )

            output_path = os.path.join(save_directory, output_filename)
            with open(output_path, "wb") as output_file:
                writer.write(output_file)

    def _split_by_pages(self, reader, save_directory, input_filename, pages_per_file):
        """按页数拆分"""
        total_pages = len(reader.pages)
        file_index = 1

        for start in range(0, total_pages, pages_per_file):
            end = min(start + pages_per_file, total_pages)
            writer = pypdf.PdfWriter()

            # 添加指定范围的页面
            for i in range(start, end):
                writer.add_page(reader.pages[i])

            # 生成文件名
            if self.split_filename_widgets["filename_var"].get() == "default":
                output_filename = (
                    f"{datetime.now().strftime('%Y%m%d%H%M%S')}_part_{file_index}.pdf"
                )
            else:
                custom_name = self.split_filename_widgets["entry"].get()
                if custom_name:
                    output_filename = f"{custom_name}_part_{file_index}.pdf"
                else:
                    output_filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_part_{file_index}.pdf"

            output_path = os.path.join(save_directory, output_filename)
            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            file_index += 1

    def _split_by_range(self, reader, save_directory, input_filename, range_str):
        """按范围拆分"""
        try:
            ranges = range_str.split(",")
            page_numbers = []

            for r in ranges:
                if "-" in r:
                    start, end = map(int, r.split("-"))
                    page_numbers.extend(range(start, end + 1))
                else:
                    page_numbers.append(int(r))

            # 去重并排序
            page_numbers = sorted(set(page_numbers))

            # 生成文件名
            if self.split_filename_widgets["filename_var"].get() == "default":
                output_filename = (
                    f"{datetime.now().strftime('%Y%m%d%H%M%S')}_custom_range.pdf"
                )
            else:
                custom_name = self.split_filename_widgets["entry"].get()
                if custom_name:
                    output_filename = f"{custom_name}.pdf"
                else:
                    output_filename = (
                        f"{datetime.now().strftime('%Y%m%d%H%M%S')}_custom_range.pdf"
                    )

            output_path = os.path.join(save_directory, output_filename)
            writer = pypdf.PdfWriter()

            # 添加指定页码的页面
            for page_num in page_numbers:
                writer.add_page(reader.pages[page_num - 1])

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

        except Exception as e:
            messagebox.showerror("错误", f"解析页码范围时发生错误：{str(e)}")

    def setup_watermark_page(self):
        """设置加水印页面的内容"""
        # 添加标题
        desc_label = ttk.Label(
            self.watermark_frame, text="为PDF文件添加水印", font=self.default_font
        )
        desc_label.pack(pady=5)

        # 主框架
        main_frame = ttk.Frame(self.watermark_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 选择PDF文件
        file_frame = ttk.LabelFrame(main_frame, text="选择PDF文件")
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        self.watermark_file_var = tk.StringVar()
        file_select_frame = ttk.Frame(file_frame)
        file_select_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            file_select_frame, text="选择PDF文件", command=self.select_watermark_file
        ).pack(side=tk.LEFT)
        ttk.Entry(
            file_select_frame, textvariable=self.watermark_file_var, state="readonly"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # 水印类型选择
        watermark_type_frame = ttk.LabelFrame(main_frame, text="水印类型")
        watermark_type_frame.pack(fill=tk.X, padx=5, pady=5)

        self.watermark_type_var = tk.StringVar(value="text")

        # 文字水印和图片水印单选框
        type_buttons_frame = ttk.Frame(watermark_type_frame)
        type_buttons_frame.pack(fill=tk.X, padx=10, pady=5)

        text_radio = ttk.Radiobutton(
            type_buttons_frame,
            text="文字水印",
            variable=self.watermark_type_var,
            value="text",
        )
        text_radio.pack(side=tk.LEFT, padx=10)

        image_radio = ttk.Radiobutton(
            type_buttons_frame,
            text="图片水印",
            variable=self.watermark_type_var,
            value="image",
        )
        image_radio.pack(side=tk.LEFT, padx=10)
        ############################################################################
        # 文字水印设置
        self.text_watermark_frame = ttk.Frame(watermark_type_frame)
        self.text_watermark_frame.pack(fill=tk.X, padx=5, pady=5)

        text_settings_frame = ttk.LabelFrame(self.text_watermark_frame)
        text_settings_frame.pack(fill=tk.X, padx=5, pady=5)

        # 水印文字输入
        text_input_frame = ttk.Frame(text_settings_frame)
        text_input_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(text_input_frame, text="水印文字:").pack(side=tk.LEFT)
        self.watermark_text_var = tk.StringVar(value="富强民主文明和谐")
        watermark_text_entry = ttk.Entry(
            text_input_frame, textvariable=self.watermark_text_var
        )
        watermark_text_entry.pack(side=tk.LEFT, padx=5)

        # 字体大小设置
        font_size_frame = ttk.Frame(text_settings_frame)
        font_size_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(font_size_frame, text="字体大小:").pack(side=tk.LEFT)
        self.watermark_font_size_var = tk.StringVar(value="30")
        font_size_spinbox = ttk.Spinbox(
            font_size_frame,
            from_=10,
            to=200,
            width=10,
            textvariable=self.watermark_font_size_var,
        )
        font_size_spinbox.pack(side=tk.LEFT, padx=5)

        # 图片水印设置
        self.image_watermark_frame = ttk.Frame(watermark_type_frame)
        self.image_watermark_frame.pack(fill=tk.X, padx=5, pady=5)
        self.image_watermark_frame.pack_forget()  # 默认隐藏

        image_settings_frame = ttk.LabelFrame(self.image_watermark_frame)
        image_settings_frame.pack(fill=tk.X, padx=5, pady=5)

        self.watermark_image_var = tk.StringVar()
        image_select_frame = ttk.Frame(image_settings_frame)
        image_select_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            image_select_frame, text="选择水印图片", command=self.select_watermark_image
        ).pack(side=tk.LEFT)
        ttk.Entry(
            image_select_frame, textvariable=self.watermark_image_var, state="readonly"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # 水印透明度设置
        opacity_frame = ttk.Frame(image_settings_frame)
        opacity_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(opacity_frame, text="透明度 (0-100%):").pack(side=tk.LEFT)
        self.watermark_opacity_var = tk.StringVar(value="20")
        opacity_spinbox = ttk.Spinbox(
            opacity_frame,
            from_=0,
            to=100,
            width=10,
            textvariable=self.watermark_opacity_var,
        )
        opacity_spinbox.pack(side=tk.LEFT, padx=5)
        ############################################################################
        # 绑定水印类型变化事件
        self.watermark_type_var.trace("w", self.on_watermark_type_change)

        # 创建存储位置设置
        self.watermark_storage_widgets = self.settings_manager.create_storage_settings(
            main_frame,
            "watermark_save_location_var",
            "watermark_folder_path_var",
            "select_watermark_folder_btn",
            "watermark_folder_path_entry",
            self.on_watermark_save_location_change,
        )

        # 配置选择文件夹按钮的命令
        self.watermark_storage_widgets["select_btn"].config(
            command=lambda: self.settings_manager.select_save_folder(
                self.watermark_storage_widgets["folder_path_var"]
            )
        )

        # 创建文件名设置
        self.watermark_filename_widgets = (
            self.settings_manager.create_filename_settings(
                main_frame,
                "watermark_filename_var",
                "watermark_filename_entry",
                self.on_watermark_filename_change,
            )
        )

        # 开始操作按钮
        watermark_btn = ttk.Button(
            self.watermark_filename_widgets["frame"],
            text="开始添加水印",
            command=self.process_watermark_enhanced,
        )
        watermark_btn.pack(pady=5)

    def select_watermark_file(self):
        """选择要添加水印的PDF文件"""
        file = filedialog.askopenfilename(
            title="选择PDF文件", filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )

        if file:
            reader = pypdf.PdfReader(file)
            page_count = len(reader.pages)
            self.filename_a = file
            self.watermark_file_var.set(f"{page_count}页 -> {os.path.basename(file)} ")

    def select_watermark_image(self):
        """选择水印图片"""
        file = filedialog.askopenfilename(
            title="选择水印图片",
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif"),
                ("所有文件", "*.*"),
            ],
        )

        if file:
            self.watermark_image_var.set(file)

    def on_watermark_type_change(self, *args):
        """水印类型改变时的处理"""
        watermark_type = self.watermark_type_var.get()

        # 根据选择显示或隐藏设置框
        if watermark_type == "text":
            self.text_watermark_frame.pack(fill=tk.X, padx=5, pady=5)
            self.image_watermark_frame.pack_forget()
        else:
            self.text_watermark_frame.pack_forget()
            self.image_watermark_frame.pack(fill=tk.X, padx=5, pady=5)

    def on_watermark_save_location_change(self, *args):
        """水印存储位置选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.watermark_storage_widgets["location_var"].get()
        self.config.set("MergeSettings", "save_location_option", current_value)
        self.save_config()

    def on_watermark_filename_change(self, *args):
        """水印文件名选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.watermark_filename_widgets["filename_var"].get()
        self.config.set("MergeSettings", "filename_option", current_value)
        self.save_config()

    def process_watermark_enhanced(self):
        """增强版水印功能 - 支持更多选项"""
        # 检查是否选择了文件
        input_file = getattr(self, "filename_a", None)
        if not input_file:
            messagebox.showwarning("警告", "请先选择PDF文件！")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("错误", "选择的文件不存在！")
            return

        try:
            # 确定保存路径
            save_directory = self.settings_manager.get_save_directory(
                self.watermark_storage_widgets["location_var"],
                self.watermark_storage_widgets["folder_path_var"],
            )
            if not save_directory:
                messagebox.showerror("错误", "请选择有效的保存路径！")
                return

            # 生成文件名
            input_filename = os.path.splitext(os.path.basename(input_file))[0]

            def get_default_name():
                return f"{input_filename}_watermarked"

            filename = self.settings_manager.get_filename(
                self.watermark_filename_widgets["filename_var"],
                self.watermark_filename_widgets["entry"],
                get_default_name,
            )
            if not filename:
                filename = f"{input_filename}_watermarked"

            output_path = os.path.join(save_directory, filename + ".pdf")

            # 检查文件是否已存在
            if os.path.exists(output_path):
                if not messagebox.askyesno(
                    "文件已存在", f"文件 {filename}.pdf 已存在，是否覆盖？"
                ):
                    return

            # 获取水印参数
            watermark_type = self.watermark_type_var.get()

            # 显示进度
            progress_window = tk.Toplevel(self.root)
            progress_window.title("正在添加水印...")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()

            tk.Label(progress_window, text="正在处理水印，请稍候...").pack(pady=10)
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(
                progress_window, variable=progress_var, maximum=100
            )
            progress_bar.pack(pady=10, padx=20, fill=tk.X)
            progress_window.update()

            if watermark_type == "text":
                # 文字水印
                self._add_enhanced_text_watermark(
                    input_file, output_path, progress_var, progress_window
                )
            else:
                # 图片水印
                self._add_enhanced_image_watermark(
                    input_file, output_path, progress_var, progress_window
                )

            progress_window.destroy()

            # 完成后询问是否打开文件
            result = messagebox.askyesno(
                "成功", f"PDF水印添加完成！\n保存位置：{output_path}\n\n是否打开文件？"
            )
            if result:
                import subprocess

                try:
                    if os.name == "nt":  # Windows
                        os.startfile(output_path)
                    elif os.name == "posix":  # macOS or Linux
                        if sys.platform == "darwin":  # macOS
                            subprocess.call(("open", output_path))
                        else:  # Linux
                            subprocess.call(("xdg-open", output_path))
                except Exception as e:
                    print(f"打开文件失败: {e}")

        except Exception as e:
            messagebox.showerror("错误", f"处理PDF文件时发生错误：{str(e)}")
            import traceback

            traceback.print_exc()

    def _add_enhanced_text_watermark(
        self, input_file, output_file, progress_var=None, progress_window=None
    ):
        # 读取PDF
        reader = pypdf.PdfReader(input_file)
        total_pages = len(reader.pages)
        writer = pypdf.PdfWriter()

        # 注册字体
        self._register_chinese_fonts()

        # 获取参数
        text = self.watermark_text_var.get()
        font_size = int(self.watermark_font_size_var.get())
        position = "center"
        watermark_style = "repeat"  # self.watermark_type_var.get()  # single 或 repeat

        for page_idx, page in enumerate(reader.pages):
            # 更新进度
            if progress_var and progress_window:
                progress = (page_idx + 1) / total_pages * 100
                progress_var.set(progress)
                progress_window.update()

            # 页面尺寸
            width = page.mediabox.width
            height = page.mediabox.height

            # 创建水印
            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=(width, height))

            # 设置字体
            can.setFont(
                (
                    "ChineseFont"
                    if hasattr(self, "_chinese_font_registered")
                    and self._chinese_font_registered
                    else "Helvetica"
                ),
                font_size,
            )

            # 设置颜色和透明度
            can.setFillColorRGB(0.2, 0.2, 0.2)  # 深灰色
            can.setFillAlpha(0.2)  # 20%透明度

            if watermark_style == "repeat":
                # 平铺水印
                self._draw_enhanced_tiled_watermark(can, text, width, height, font_size)
            else:
                # 单个水印
                self._draw_enhanced_single_watermark(
                    can, text, position, width, height, font_size
                )

            can.save()
            packet.seek(0)

            # 合并
            watermark = pypdf.PdfReader(packet)
            page.merge_page(watermark.pages[0])
            writer.add_page(page)

        # 保存
        with open(output_file, "wb") as f:
            writer.write(f)

    def _add_enhanced_image_watermark(
        self, input_file, output_file, progress_var=None, progress_window=None
    ):
        # 读取PDF
        reader = pypdf.PdfReader(input_file)
        total_pages = len(reader.pages)
        writer = pypdf.PdfWriter()

        # 获取图片路径
        image_path = self.watermark_image_var.get()
        if not image_path or not os.path.exists(image_path):
            messagebox.showerror("错误", "请选择有效的水印图片！")
            return False

        try:
            # 获取水印参数
            opacity = int(self.watermark_opacity_var.get()) / 100.0
            position = "center"
            watermark_style = (
                "repeat"  # self.watermark_type_var.get()  # single 或 repeat
            )

            # 处理图片
            img = Image.open(image_path)

            # 获取图片缩放比例
            scale_percent = (
                int(self.watermark_scale_var.get())
                if hasattr(self, "watermark_scale_var")
                else 50
            )
            scale_factor = scale_percent / 100.0

            # 如果是RGBA模式，转换为RGB
            if img.mode in ("RGBA", "LA", "P"):
                if img.mode == "P":
                    img = img.convert("RGBA")

                # 创建白色背景
                background = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "RGBA":
                    background.paste(img, mask=img.split()[-1])
                else:  # LA
                    background.paste(img, mask=img.split()[0])
                img = background

            # 调整图片大小
            if scale_percent != 100:
                new_width = int(img.width * scale_factor)
                new_height = int(img.height * scale_factor)
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

            # 保存为临时文件
            temp_img_path = "temp_watermark.png"
            img.save(temp_img_path, "PNG")

            # 获取旋转角度
            rotation_angle = (
                int(self.watermark_rotation_var.get())
                if hasattr(self, "watermark_rotation_var")
                else 0
            )

            for page_idx, page in enumerate(reader.pages):
                # 更新进度
                if progress_var and progress_window:
                    progress = (page_idx + 1) / total_pages * 100
                    progress_var.set(progress)
                    progress_window.update()

                # 页面尺寸
                width = page.mediabox.width
                height = page.mediabox.height

                # 创建水印PDF
                packet = BytesIO()
                can = canvas.Canvas(packet, pagesize=(width, height))

                # 读取图片
                try:
                    image_reader = ImageReader(temp_img_path)
                    img_width, img_height = image_reader.getSize()

                    # 设置透明度
                    can.setFillAlpha(opacity)

                    if watermark_style == "repeat":
                        # 平铺图片水印
                        self._draw_tiled_image_watermark(
                            can,
                            image_reader,
                            width,
                            height,
                            img_width,
                            img_height,
                            rotation_angle,
                        )
                    else:
                        # 单个图片水印
                        self._draw_single_image_watermark(
                            can,
                            image_reader,
                            position,
                            width,
                            height,
                            img_width,
                            img_height,
                            rotation_angle,
                        )

                except Exception as e:
                    print(f"绘制图片水印失败: {e}")
                    # 如果图片水印失败，回退到文字提示
                    can.setFont("Helvetica", 12)
                    can.setFillAlpha(0.5)
                    can.drawString(50, 50, "水印图片加载失败")

                can.save()
                packet.seek(0)

                # 合并水印
                watermark_reader = pypdf.PdfReader(packet)
                watermark_page = watermark_reader.pages[0]

                # 合并到原页面
                page.merge_page(watermark_page)
                writer.add_page(page)

            # 保存结果
            with open(output_file, "wb") as f:
                writer.write(f)

            return True

        except Exception as e:
            messagebox.showerror("错误", f"添加图片水印失败：{str(e)}")
            import traceback

            traceback.print_exc()
            return False

        finally:
            # 清理临时文件
            try:
                if os.path.exists(temp_img_path):
                    os.remove(temp_img_path)
            except:
                pass

    def _draw_tiled_image_watermark(
        self,
        canvas,
        image_reader,
        page_width,
        page_height,
        img_width,
        img_height,
        rotation_angle=0,
    ):
        """绘制平铺图片水印"""
        # 计算间距（图片宽度+20%的边距）
        x_spacing = int(img_width * 1.2)
        y_spacing = int(img_height * 1.2)

        for i in range(-x_spacing, int(page_width) + x_spacing, x_spacing):
            for j in range(-y_spacing, int(page_height) + y_spacing, y_spacing):
                canvas.saveState()

                # 移动到中心位置
                center_x = i + img_width / 2
                center_y = j + img_height / 2

                # 应用旋转
                if rotation_angle != 0:
                    canvas.translate(center_x, center_y)
                    canvas.rotate(rotation_angle)
                    canvas.translate(-center_x, -center_y)

                # 绘制图片
                canvas.drawImage(
                    image_reader, i, j, width=img_width, height=img_height, mask="auto"
                )

                canvas.restoreState()

    def _draw_single_image_watermark(
        self,
        canvas,
        image_reader,
        position,
        page_width,
        page_height,
        img_width,
        img_height,
        rotation_angle=0,
    ):
        """绘制单个图片水印"""
        # 根据位置计算坐标
        if position == "center":
            x = (page_width - img_width) / 2
            y = (page_height - img_height) / 2
        elif position == "top_left":
            x = page_width * 0.05
            y = page_height - img_height - page_height * 0.05
        elif position == "top_right":
            x = page_width - img_width - page_width * 0.05
            y = page_height - img_height - page_height * 0.05
        elif position == "bottom_left":
            x = page_width * 0.05
            y = page_height * 0.05
        elif position == "bottom_right":
            x = page_width - img_width - page_width * 0.05
            y = page_height * 0.05
        else:  # 默认居中
            x = (page_width - img_width) / 2
            y = (page_height - img_height) / 2

        canvas.saveState()

        # 如果需要旋转
        if rotation_angle != 0:
            # 移动到图片中心
            center_x = x + img_width / 2
            center_y = y + img_height / 2

            canvas.translate(center_x, center_y)
            canvas.rotate(rotation_angle)
            canvas.translate(-center_x, -center_y)

        # 绘制图片
        canvas.drawImage(
            image_reader, x, y, width=img_width, height=img_height, mask="auto"
        )

        canvas.restoreState()

    def _register_chinese_fonts(self):
        """注册中文字体"""
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont

        if hasattr(self, "_chinese_font_registered") and self._chinese_font_registered:
            return

        # 字体搜索路径
        font_search_paths = [
            (
                "C:/Windows/Fonts/",
                ["simhei.ttf", "simsun.ttc", "msyh.ttc", "simkai.ttf", "simfang.ttf"],
            ),
            ("/System/Library/Fonts/", ["PingFang.ttc", "Hiragino Sans GB.ttc"]),
            (
                "/usr/share/fonts/truetype/",
                ["wqy-microhei.ttc", "droid/DroidSansFallbackFull.ttf"],
            ),
        ]

        for base_path, font_files in font_search_paths:
            for font_file in font_files:
                font_path = os.path.join(base_path, font_file)
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont("ChineseFont", font_path))
                        self._chinese_font_registered = True
                        print(f"已注册字体: {font_path}")
                        return
                    except Exception as e:
                        print(f"注册字体失败 {font_path}: {e}")

        # 如果没有找到中文字体
        self._chinese_font_registered = False
        print("未找到中文字体，将使用英文字体")

    def _draw_enhanced_tiled_watermark(
        self, canvas, text, page_width, page_height, font_size
    ):
        """增强版平铺水印"""
        # 计算间距
        text_width = canvas.stringWidth(text, canvas._fontname, canvas._fontsize)
        x_spacing = text_width + 150
        y_spacing = font_size + 100

        # 绘制网格
        for i in range(
            -int(x_spacing), int(page_width) + int(x_spacing), int(x_spacing)
        ):
            for j in range(
                -int(y_spacing), int(page_height) + int(y_spacing), int(y_spacing)
            ):
                canvas.saveState()
                canvas.translate(i, j)
                canvas.rotate(30)  # 角度可调

                # 可选：添加阴影效果
                canvas.setFillColorRGB(0.1, 0.1, 0.1, alpha=0.1)
                # canvas.drawString(2, -2, text)  # 阴影

                canvas.setFillColorRGB(0.5, 0.5, 0.5, alpha=0.2)
                canvas.drawString(0, 0, text)  # 主文字

                canvas.restoreState()

    def _draw_enhanced_single_watermark(
        self, canvas, text, position, page_width, page_height, font_size
    ):
        """增强版单个水印"""
        # 位置映射
        positions = {
            "center": (page_width / 2, page_height / 2, 45, True),
            "top_left": (page_width * 0.2, page_height * 0.8, 45, False),
            "top_right": (page_width * 0.8, page_height * 0.8, 45, False),
            "bottom_left": (page_width * 0.2, page_height * 0.2, 45, False),
            "bottom_right": (page_width * 0.8, page_height * 0.2, 45, False),
        }

        x, y, rotation, centered = positions.get(
            position, (page_width / 2, page_height / 2, 45, True)
        )

        canvas.saveState()
        canvas.translate(x, y)
        canvas.rotate(rotation)

        # 文字居中
        if centered:
            text_width = canvas.stringWidth(text, canvas._fontname, canvas._fontsize)
            x_offset = -text_width / 2
        else:
            x_offset = 0

        # 可选：添加边框
        canvas.setStrokeColorRGB(0.3, 0.3, 0.3, alpha=0.1)
        canvas.setLineWidth(0.5)
        canvas.rect(
            x_offset - 5,
            -font_size / 2 - 5,
            text_width + 10,
            font_size + 10,
            stroke=1,
            fill=0,
        )

        # 绘制文字
        canvas.drawString(x_offset, 0, text)

        canvas.restoreState()

    def setup_insert_page(self):
        """设置插入页面的内容"""
        # 添加标题
        # title_label = ttk.Label(self.insert_frame, text="PDF插入功能", font=("仿宋", 16))
        # title_label.pack(pady=10)

        # 说明标签
        desc_label = ttk.Label(
            self.insert_frame,
            text="在指定位置插入另一个PDF的全部或部分页面",
            font=self.default_font,
        )
        desc_label.pack(pady=5)

        # 主框架
        main_frame = ttk.Frame(self.insert_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 第一步：选择目标PDF文件（要被插入的文件）
        target_frame = ttk.LabelFrame(main_frame, text="主文件(被插入的文件)")
        target_frame.pack(fill=tk.X, padx=5, pady=5)

        self.target_file_var = tk.StringVar()
        target_file_frame = ttk.Frame(target_frame)
        target_file_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            target_file_frame, text="选择目标文件", command=self.select_target_file
        ).pack(side=tk.LEFT)
        ttk.Entry(
            target_file_frame, textvariable=self.target_file_var, state="readonly"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # 第二步：选择要插入的PDF文件
        insert_frame_inner = ttk.LabelFrame(main_frame, text="插入的PDF文件")
        insert_frame_inner.pack(fill=tk.X, padx=5, pady=5)

        self.insert_file_var = tk.StringVar()
        insert_file_frame = ttk.Frame(insert_frame_inner)
        insert_file_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            insert_file_frame, text="选择插入文件", command=self.select_insert_file
        ).pack(side=tk.LEFT)
        ttk.Entry(
            insert_file_frame, textvariable=self.insert_file_var, state="readonly"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # 插入方式选择
        insert_method_frame = ttk.LabelFrame(main_frame, text="插入方式")
        insert_method_frame.pack(fill=tk.X, padx=5, pady=5)

        self.insert_method_var = tk.StringVar(value="position")

        # 在指定位置插入、在文档首部插入、在文档尾部插入 水平摆放
        method_buttons_frame = ttk.Frame(insert_method_frame)
        method_buttons_frame.pack(fill=tk.X, padx=10, pady=5)

        position_radio = ttk.Radiobutton(
            method_buttons_frame,
            text="在指定位置插入",
            variable=self.insert_method_var,
            value="position",
        )
        position_radio.pack(side=tk.LEFT, padx=10)

        head_radio = ttk.Radiobutton(
            method_buttons_frame,
            text="在文档首部插入",
            variable=self.insert_method_var,
            value="head",
        )
        head_radio.pack(side=tk.LEFT, padx=10)

        tail_radio = ttk.Radiobutton(
            method_buttons_frame,
            text="在文档尾部插入",
            variable=self.insert_method_var,
            value="tail",
        )
        tail_radio.pack(side=tk.LEFT, padx=10)

        # 插入位置设置
        self.position_frame = ttk.Frame(insert_method_frame)
        self.position_frame.pack(fill=tk.X, padx=30, pady=5)

        ttk.Label(self.position_frame, text="插入位置（页码）:").pack(side=tk.LEFT)
        self.insert_position_var = tk.StringVar(value="1")
        position_spinbox = ttk.Spinbox(
            self.position_frame,
            from_=1,
            to=1000,
            width=10,
            textvariable=self.insert_position_var,
        )
        position_spinbox.pack(side=tk.LEFT, padx=5)

        # 插入页码范围设置
        range_frame = ttk.Frame(insert_method_frame)
        range_frame.pack(fill=tk.X, padx=30, pady=5)

        ttk.Label(range_frame, text="插入页码范围（如：1-3,5，留空则全部插入）:").pack(
            anchor=tk.W
        )
        self.insert_range_var = tk.StringVar()
        ttk.Entry(range_frame, textvariable=self.insert_range_var).pack(
            fill=tk.X, pady=5
        )

        # 默认隐藏位置设置（因为默认是position，所以显示）
        # self.position_frame.pack_forget()

        # 绑定插入方式变化事件
        self.insert_method_var.trace("w", self.on_insert_method_change)

        # 创建存储位置设置
        self.insert_storage_widgets = self.settings_manager.create_storage_settings(
            main_frame,
            "insert_save_location_var",
            "insert_folder_path_var",
            "select_insert_folder_btn",
            "insert_folder_path_entry",
            self.on_insert_save_location_change,
        )

        # 配置选择文件夹按钮的命令
        self.insert_storage_widgets["select_btn"].config(
            command=lambda: self.settings_manager.select_save_folder(
                self.insert_storage_widgets["folder_path_var"]
            )
        )

        # 创建文件名设置
        self.insert_filename_widgets = self.settings_manager.create_filename_settings(
            main_frame,
            "insert_filename_var",
            "insert_filename_entry",
            self.on_insert_filename_change,
        )

        # 开始插入按钮
        insert_btn = ttk.Button(
            self.insert_filename_widgets["frame"],
            text="开始插入",
            command=self.insert_pdf,
        )
        insert_btn.pack(pady=5)

    def select_target_file(self):
        """选择被插PDF文件"""
        file = filedialog.askopenfilename(
            title="选择目标PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        )

        if file:
            reader = pypdf.PdfReader(file)
            page_count = len(reader.pages)
            self.filename_a = file
            self.target_file_var.set(f"{page_count}页 -> {os.path.basename(file)}")

    def select_insert_file(self):
        """选择要插入的PDF文件"""
        file = filedialog.askopenfilename(
            title="选择要插入的PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        )

        if file:
            reader = pypdf.PdfReader(file)
            page_count = len(reader.pages)
            self.filename_b = file
            self.insert_file_var.set(f"{page_count}页 -> {os.path.basename(file)}")

    def on_insert_method_change(self, *args):
        """插入方式改变时的处理"""
        method = self.insert_method_var.get()

        # 根据选择显示或隐藏位置设置
        if method == "position":
            self.position_frame.pack(fill=tk.X, padx=30, pady=5)
        else:
            self.position_frame.pack_forget()

    def on_insert_save_location_change(self, *args):
        """插入存储位置选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.insert_storage_widgets["location_var"].get()
        self.config.set("MergeSettings", "save_location_option", current_value)
        self.save_config()

    def on_insert_filename_change(self, *args):
        """插入文件名选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.insert_filename_widgets["filename_var"].get()
        self.config.set("MergeSettings", "filename_option", current_value)
        self.save_config()

    def _add_pages_from_reader(self, writer, reader, range_str):
        """根据页码范围添加页面到writer"""
        total_pages = len(reader.pages)

        if not range_str:
            # 如果没有指定范围，则添加所有页面
            for i in range(total_pages):
                writer.add_page(reader.pages[i])
        else:
            # 解析范围字符串，例如："1-3,5"
            ranges = []
            for part in range_str.split(","):
                part = part.strip()
                if "-" in part:
                    start, end = map(int, part.split("-"))
                    ranges.extend(range(start - 1, end))  # 转换为0基索引
                else:
                    ranges.append(int(part) - 1)  # 转换为0基索引

            # 去重并排序
            ranges = sorted(list(set(ranges)))

            # 检查范围有效性
            for page_num in ranges:
                if page_num < 0 or page_num >= total_pages:
                    raise ValueError(f"页码 {page_num+1} 超出范围（1-{total_pages}）！")

            # 添加指定页面
            for page_index in ranges:
                writer.add_page(reader.pages[page_index])

    def process_watermark_enhanced(self):
        """增强版水印功能 - 支持更多选项"""
        # 检查是否选择了文件
        input_file = getattr(self, "filename_a", None)
        if not input_file:
            messagebox.showwarning("警告", "请先选择PDF文件！")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("错误", "选择的文件不存在！")
            return

        try:
            # 确定保存路径
            save_directory = self.settings_manager.get_save_directory(
                self.watermark_storage_widgets["location_var"],
                self.watermark_storage_widgets["folder_path_var"],
            )
            if not save_directory:
                messagebox.showerror("错误", "请选择有效的保存路径！")
                return

            # 生成文件名
            input_filename = os.path.splitext(os.path.basename(input_file))[0]

            def get_default_name():
                return f"{datetime.now().strftime('%Y%m%d%H%M%S')}_watermarked"

            filename = self.settings_manager.get_filename(
                self.watermark_filename_widgets["filename_var"],
                self.watermark_filename_widgets["entry"],
                get_default_name,
            )
            if not filename:
                filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_watermarked"

            output_path = os.path.join(save_directory, filename + ".pdf")

            # 检查文件是否已存在
            if os.path.exists(output_path):
                if not messagebox.askyesno(
                    "文件已存在", f"文件 {filename}.pdf 已存在，是否覆盖？"
                ):
                    return

            # 获取水印参数
            watermark_type = self.watermark_type_var.get()

            # 显示进度
            progress_window = tk.Toplevel(self.root)
            progress_window.title("正在添加水印...")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()

            tk.Label(progress_window, text="正在处理水印，请稍候...").pack(pady=10)
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(
                progress_window, variable=progress_var, maximum=100
            )
            progress_bar.pack(pady=10, padx=20, fill=tk.X)
            progress_window.update()

            if watermark_type == "text":
                # 文字水印
                self._add_enhanced_text_watermark(
                    input_file, output_path, progress_var, progress_window
                )
            else:
                # 图片水印
                self._add_enhanced_image_watermark(
                    input_file, output_path, progress_var, progress_window
                )

            progress_window.destroy()

            # 完成后询问是否打开文件
            result = messagebox.askyesno(
                "成功", f"PDF水印添加完成！\n保存位置：{output_path}\n\n是否打开文件？"
            )
            if result:
                import subprocess

                try:
                    if os.name == "nt":  # Windows
                        os.startfile(output_path)
                    elif os.name == "posix":  # macOS or Linux
                        if sys.platform == "darwin":  # macOS
                            subprocess.call(("open", output_path))
                        else:  # Linux
                            subprocess.call(("xdg-open", output_path))
                except Exception as e:
                    print(f"打开文件失败: {e}")

        except Exception as e:
            messagebox.showerror("错误", f"处理PDF文件时发生错误：{str(e)}")
            import traceback

            traceback.print_exc()

    def insert_pdf(self):
        """插入PDF文件"""
        # 检查是否选择了文件
        target_file = getattr(self, "filename_a", None)
        insert_file = getattr(self, "filename_b", None)

        if not target_file:
            messagebox.showwarning("警告", "请先选择目标PDF文件！")
            return

        if not insert_file:
            messagebox.showwarning("警告", "请先选择要插入的PDF文件！")
            return

        if not os.path.exists(target_file):
            messagebox.showerror("错误", "目标文件不存在！")
            return

        if not os.path.exists(insert_file):
            messagebox.showerror("错误", "要插入的文件不存在！")
            return

        try:
            # 确定保存路径
            save_directory = self.settings_manager.get_save_directory(
                self.insert_storage_widgets["location_var"],
                self.insert_storage_widgets["folder_path_var"],
            )
            if not save_directory:
                messagebox.showerror("错误", "请选择有效的保存路径！")
                return

            # 获取目标文件名（不含扩展名）
            target_filename = os.path.splitext(os.path.basename(target_file))[0]

            # 生成文件名
            def get_default_name():
                return f"{datetime.now().strftime('%Y%m%d%H%M%S')}_inserted"

            filename = self.settings_manager.get_filename(
                self.insert_filename_widgets["filename_var"],
                self.insert_filename_widgets["entry"],
                get_default_name,
            )
            if not filename:
                filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_inserted"

            output_path = os.path.join(save_directory, filename + ".pdf")

            # 检查文件是否已存在
            if os.path.exists(output_path):
                result = messagebox.askyesno(
                    "文件已存在", f"文件 {filename}.pdf 已存在，是否覆盖？"
                )
                if not result:
                    return

            # 读取目标PDF文件
            target_reader = pypdf.PdfReader(target_file)
            target_total_pages = len(target_reader.pages)

            # 读取要插入的PDF文件
            insert_reader = pypdf.PdfReader(insert_file)
            insert_total_pages = len(insert_reader.pages)

            # 创建PDF写入器
            writer = pypdf.PdfWriter()

            # 根据插入方式执行插入操作
            method = self.insert_method_var.get()

            if method == "head":
                # 在文档首部插入
                self._add_pages_from_reader(
                    writer, insert_reader, self.insert_range_var.get()
                )
                self._add_pages_from_reader(
                    writer, target_reader, ""
                )  # 添加所有目标页面
            elif method == "tail":
                # 在文档尾部插入
                self._add_pages_from_reader(
                    writer, target_reader, ""
                )  # 添加所有目标页面
                self._add_pages_from_reader(
                    writer, insert_reader, self.insert_range_var.get()
                )
            else:  # position
                # 在指定位置插入
                position = int(self.insert_position_var.get()) - 1  # 转换为0基索引
                if position < 0 or position > target_total_pages:
                    messagebox.showerror(
                        "错误", f"插入位置超出范围（1-{target_total_pages+1}）！"
                    )
                    return

                # 先添加目标文件中位置之前的部分
                for i in range(position):
                    writer.add_page(target_reader.pages[i])

                # 添加要插入的页面
                self._add_pages_from_reader(
                    writer, insert_reader, self.insert_range_var.get()
                )

                # 添加目标文件中位置之后的部分
                for i in range(position, target_total_pages):
                    writer.add_page(target_reader.pages[i])

            # 写入文件
            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            messagebox.showinfo("成功", f"PDF插入完成！\n保存位置：{output_path}")

        except Exception as e:
            messagebox.showerror("错误", f"插入PDF时发生错误：{str(e)}")

    def replace_pdf(self):
        """替换PDF文件中的页面"""
        # 检查是否选择了文件
        target_file = getattr(self, "filename_a", None)
        replace_file = getattr(self, "filename_b", None)

        if not target_file:
            messagebox.showwarning("警告", "请先选择目标PDF文件！")
            return

        if not replace_file:
            messagebox.showwarning("警告", "请先选择用来替换的PDF文件！")
            return

        if not os.path.exists(target_file):
            messagebox.showerror("错误", "目标文件不存在！")
            return

        if not os.path.exists(replace_file):
            messagebox.showerror("错误", "替换文件不存在！")
            return

        try:
            # 确定保存路径
            save_directory = self.settings_manager.get_save_directory(
                self.replace_storage_widgets["location_var"],
                self.replace_storage_widgets["folder_path_var"],
            )
            if not save_directory:
                messagebox.showerror("错误", "请选择有效的保存路径！")
                return

            # 获取目标文件名（不含扩展名）
            target_filename = os.path.splitext(os.path.basename(target_file))[0]

            # 生成文件名
            if self.replace_filename_widgets["filename_var"].get() == "default":
                output_filename = (
                    f"{datetime.now().strftime('%Y%m%d%H%M%S')}_replaced.pdf"
                )
            else:
                custom_name = self.replace_filename_widgets["entry"].get()
                if custom_name:
                    output_filename = f"{custom_name}.pdf"
                else:
                    output_filename = (
                        f"{datetime.now().strftime('%Y%m%d%H%M%S')}_replaced.pdf"
                    )

            output_path = os.path.join(save_directory, output_filename)

            # 检查文件是否已存在
            if os.path.exists(output_path):
                result = messagebox.askyesno(
                    "文件已存在", f"文件 {output_filename} 已存在，是否覆盖？"
                )
                if not result:
                    return

            # 读取目标PDF文件
            target_reader = pypdf.PdfReader(target_file)
            target_total_pages = len(target_reader.pages)

            # 读取替换PDF文件
            replace_reader = pypdf.PdfReader(replace_file)
            replace_total_pages = len(replace_reader.pages)

            # 创建PDF写入器
            writer = pypdf.PdfWriter()

            # 根据替换方式执行替换操作
            method = self.replace_method_var.get()

            if method == "single":
                # 替换单个页面
                replace_position = int(self.replace_position_var.get())  # 1基索引

                # 检查替换位置有效性
                if replace_position < 1 or replace_position > target_total_pages:
                    messagebox.showerror(
                        "错误", f"替换位置超出范围（1-{target_total_pages}）！"
                    )
                    return

                # 解析替换源范围
                source_ranges = self._parse_page_ranges(
                    self.replace_source_range_var.get(), replace_total_pages
                )

                # 检查替换页数是否匹配
                if len(source_ranges) != 1:
                    messagebox.showerror(
                        "错误", "替换单个页面时，替换文件只能指定一个页面！"
                    )
                    return

                # 添加目标文件中替换位置之前的部分
                for i in range(replace_position - 1):
                    writer.add_page(target_reader.pages[i])

                # 添加替换文件中指定的页面
                source_page_index = source_ranges[0] - 1  # 转换为0基索引
                writer.add_page(replace_reader.pages[source_page_index])

                # 添加目标文件中替换位置之后的部分
                for i in range(replace_position, target_total_pages):
                    writer.add_page(target_reader.pages[i])

            else:  # multi
                # 替换多个页面
                range_str = self.replace_range_var.get()
                if not range_str:
                    messagebox.showerror("错误", "请输入被替换的页码范围！")
                    return

                # 解析被替换的页码范围
                target_ranges = self._parse_page_ranges(range_str, target_total_pages)

                # 解析替换源范围
                source_ranges = self._parse_page_ranges(
                    self.replace_source_range_var.get(), replace_total_pages
                )

                # 检查替换页数是否匹配
                if len(target_ranges) != len(source_ranges):
                    messagebox.showerror(
                        "错误",
                        f"被替换页数({len(target_ranges)})与替换页数({len(source_ranges)})不匹配！",
                    )
                    return

                # 构建替换映射 {被替换页码: 替换页码索引}
                replace_map = dict(
                    zip(target_ranges, [p - 1 for p in source_ranges])
                )  # 替换页码转为0基索引

                # 逐页处理
                for i in range(target_total_pages):
                    page_num = i + 1  # 1基索引
                    if page_num in replace_map:
                        # 使用替换文件中的页面
                        source_page_index = replace_map[page_num]
                        writer.add_page(replace_reader.pages[source_page_index])
                    else:
                        # 使用原始目标文件中的页面
                        writer.add_page(target_reader.pages[i])

            # 写入文件
            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            messagebox.showinfo("成功", f"PDF页面替换完成！\n保存位置：{output_path}")

        except Exception as e:
            messagebox.showerror("错误", f"替换PDF页面时发生错误：{str(e)}")

    def process_encrypt(self):
        """处理加密操作"""
        # 检查是否选择了文件
        input_file = getattr(self, "filename_a", None)
        if not input_file:
            messagebox.showwarning("警告", "请先选择PDF文件！")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("错误", "选择的文件不存在！")
            return

        password = self.encrypt_password_var.get()
        if not password:
            messagebox.showwarning("警告", "请输入密码！")
            return

        try:
            # 确定保存路径
            save_directory = self.settings_manager.get_save_directory(
                self.encrypt_storage_widgets["location_var"],
                self.encrypt_storage_widgets["folder_path_var"],
            )
            if not save_directory:
                messagebox.showerror("错误", "请选择有效的保存路径！")
                return

            # 获取输入文件名（不含扩展名）
            input_filename = os.path.splitext(os.path.basename(input_file))[0]

            # 生成文件名
            def get_default_name():
                return f"{datetime.now().strftime('%Y%m%d%H%M%S')}_encrypted"

            filename = self.settings_manager.get_filename(
                self.encrypt_filename_widgets["filename_var"],
                self.encrypt_filename_widgets["entry"],
                get_default_name,
            )
            if not filename:
                filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_encrypted"

            output_path = os.path.join(save_directory, filename + ".pdf")

            # 检查文件是否已存在
            if os.path.exists(output_path):
                result = messagebox.askyesno(
                    "文件已存在", f"文件 {filename}.pdf 已存在，是否覆盖？"
                )
                if not result:
                    return

            # 读取PDF文件并加密
            reader = pypdf.PdfReader(input_file)
            writer = pypdf.PdfWriter()

            # 复制所有页面
            for page in reader.pages:
                writer.add_page(page)

            # 设置密码
            writer.encrypt(password)

            # 写入文件
            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            messagebox.showinfo("成功", f"PDF文件加密完成！\n保存位置：{output_path}")

        except Exception as e:
            messagebox.showerror("错误", f"处理PDF文件时发生错误：{str(e)}")

    def setup_replace_page(self):
        """设置替换页面的内容"""
        # 添加标题
        desc_label = ttk.Label(
            self.replace_frame, text="替换PDF中的指定页面", font=self.default_font
        )
        desc_label.pack(pady=5)

        # 主框架
        main_frame = ttk.Frame(self.replace_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 第一步：选择目标PDF文件（要被替换页面的文件）
        target_frame = ttk.LabelFrame(
            main_frame, text="目标PDF文件（要被替换页面的文件）"
        )
        target_frame.pack(fill=tk.X, padx=5, pady=5)

        self.replace_target_file_var = tk.StringVar()
        target_file_frame = ttk.Frame(target_frame)
        target_file_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            target_file_frame,
            text="选择目标文件",
            command=self.select_replace_target_file,
        ).pack(side=tk.LEFT)
        ttk.Entry(
            target_file_frame,
            textvariable=self.replace_target_file_var,
            state="readonly",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # 第二步：选择用来替换的PDF文件
        replace_frame_inner = ttk.LabelFrame(main_frame, text="用来替换的PDF文件")
        replace_frame_inner.pack(fill=tk.X, padx=5, pady=5)

        self.replace_file_var = tk.StringVar()
        replace_file_frame = ttk.Frame(replace_frame_inner)
        replace_file_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            replace_file_frame, text="选择替换文件", command=self.select_replace_file
        ).pack(side=tk.LEFT)
        ttk.Entry(
            replace_file_frame, textvariable=self.replace_file_var, state="readonly"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # 替换方式选择
        replace_method_frame = ttk.LabelFrame(main_frame, text="替换方式")
        replace_method_frame.pack(fill=tk.X, padx=5, pady=5)

        self.replace_method_var = tk.StringVar(value="single")

        # 替换单个页面、替换多个页面 水平摆放
        method_buttons_frame = ttk.Frame(replace_method_frame)
        method_buttons_frame.pack(fill=tk.X, padx=10, pady=5)

        single_radio = ttk.Radiobutton(
            method_buttons_frame,
            text="替换单个页面",
            variable=self.replace_method_var,
            value="single",
        )
        single_radio.pack(side=tk.LEFT, padx=10)

        multi_radio = ttk.Radiobutton(
            method_buttons_frame,
            text="替换多个页面",
            variable=self.replace_method_var,
            value="multi",
        )
        multi_radio.pack(side=tk.LEFT, padx=10)

        # 替换位置设置
        self.single_replace_frame = ttk.Frame(replace_method_frame)
        self.single_replace_frame.pack(fill=tk.X, padx=30, pady=5)

        ttk.Label(self.single_replace_frame, text="被替换的页码:").pack(side=tk.LEFT)
        self.replace_position_var = tk.StringVar(value="1")
        position_spinbox = ttk.Spinbox(
            self.single_replace_frame,
            from_=1,
            to=1000,
            width=10,
            textvariable=self.replace_position_var,
        )
        position_spinbox.pack(side=tk.LEFT, padx=5)

        # 替换页码范围设置
        self.multi_replace_frame = ttk.Frame(replace_method_frame)
        self.multi_replace_frame.pack(fill=tk.X, padx=30, pady=5)
        self.multi_replace_frame.pack_forget()  # 默认隐藏

        # ttk.Label(self.multi_replace_frame, text="被替换的页码范围（如：1-3,5）:").pack(anchor=tk.W)
        ttk.Label(self.multi_replace_frame, text="被替换的页码范围（如：1-3,5）:").pack(
            side=tk.LEFT
        )
        self.replace_range_var = tk.StringVar()
        # ttk.Entry(self.multi_replace_frame, textvariable=self.replace_range_var).pack(fill=tk.X, pady=5)
        ttk.Entry(self.multi_replace_frame, textvariable=self.replace_range_var).pack(
            side=tk.LEFT, anchor=tk.W
        )

        # 插入页码范围设置（从替换文件中选择页面）
        range_frame = ttk.Frame(replace_method_frame)
        range_frame.pack(fill=tk.X, padx=30, pady=5)

        ttk.Label(
            range_frame, text="替换文件的页码范围（如：1-3,5，留空则使用全部页面）:"
        ).pack(anchor=tk.W)
        self.replace_source_range_var = tk.StringVar()
        ttk.Entry(range_frame, textvariable=self.replace_source_range_var).pack(
            fill=tk.X, pady=5
        )

        # 绑定替换方式变化事件
        self.replace_method_var.trace("w", self.on_replace_method_change)

        # 创建存储位置设置
        self.replace_storage_widgets = self.settings_manager.create_storage_settings(
            main_frame,
            "replace_save_location_var",
            "replace_folder_path_var",
            "select_replace_folder_btn",
            "replace_folder_path_entry",
            self.on_replace_save_location_change,
        )

        # 配置选择文件夹按钮的命令
        self.replace_storage_widgets["select_btn"].config(
            command=lambda: self.settings_manager.select_save_folder(
                self.replace_storage_widgets["folder_path_var"]
            )
        )

        # 创建文件名设置
        self.replace_filename_widgets = self.settings_manager.create_filename_settings(
            main_frame,
            "replace_filename_var",
            "replace_filename_entry",
            self.on_replace_filename_change,
        )

        # 开始替换按钮
        replace_btn = ttk.Button(
            self.replace_filename_widgets["frame"],
            text="开始替换",
            command=self.replace_pdf,
        )
        replace_btn.pack(pady=5)

    def select_replace_target_file(self):
        """选择被替换PDF文件"""
        file = filedialog.askopenfilename(
            title="选择目标PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        )

        if file:
            reader = pypdf.PdfReader(file)
            page_count = len(reader.pages)
            self.filename_a = file
            self.replace_target_file_var.set(
                f"{page_count}页 -> {os.path.basename(file)}"
            )

    def select_replace_file(self):
        """选择用来替换的PDF文件"""
        file = filedialog.askopenfilename(
            title="选择用来替换的PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
        )

        if file:
            reader = pypdf.PdfReader(file)
            page_count = len(reader.pages)
            self.filename_b = file
            self.replace_file_var.set(f"{page_count}页 -> {os.path.basename(file)}")

    def on_replace_method_change(self, *args):
        """替换方式改变时的处理"""
        method = self.replace_method_var.get()

        # 根据选择显示或隐藏位置设置
        if method == "single":
            self.single_replace_frame.pack(fill=tk.X, padx=30, pady=5)
            self.multi_replace_frame.pack_forget()
        else:
            self.single_replace_frame.pack_forget()
            self.multi_replace_frame.pack(fill=tk.X, padx=30, pady=5)

    def on_replace_save_location_change(self, *args):
        """替换存储位置选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.replace_storage_widgets["location_var"].get()
        self.config.set("MergeSettings", "save_location_option", current_value)
        self.save_config()

    def on_replace_filename_change(self, *args):
        """替换文件名选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.replace_filename_widgets["filename_var"].get()
        self.config.set("MergeSettings", "filename_option", current_value)
        self.save_config()

    # remove this method as it's now handled by settings_manager

    def _parse_page_ranges(self, range_str, max_page):
        """解析页码范围字符串，返回页码列表"""
        if not range_str:
            return list(range(1, max_page + 1))

        ranges = []
        for part in range_str.split(","):
            part = part.strip()
            if "-" in part:
                start, end = map(int, part.split("-"))
                ranges.extend(range(start, end + 1))
            else:
                ranges.append(int(part))

        # 验证页码范围
        for page_num in ranges:
            if page_num < 1 or page_num > max_page:
                raise ValueError(f"页码 {page_num} 超出范围（1-{max_page}）！")

        return sorted(list(set(ranges)))

    def setup_extract_image_page(self):
        """设置图片提取页面的内容"""
        # 添加标题
        desc_label = ttk.Label(
            self.extract_image_frame, text="从PDF文件中提取图片", font=self.default_font
        )
        desc_label.pack(pady=5)

        # 主框架
        main_frame = ttk.Frame(self.extract_image_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 选择PDF文件
        file_frame = ttk.LabelFrame(main_frame, text="选择PDF文件")
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        self.extract_image_file_var = tk.StringVar()
        file_select_frame = ttk.Frame(file_frame)
        file_select_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            file_select_frame,
            text="选择PDF文件",
            command=self.select_extract_image_file,
        ).pack(side=tk.LEFT)
        ttk.Entry(
            file_select_frame,
            textvariable=self.extract_image_file_var,
            state="readonly",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # 页面范围设置
        page_range_frame = ttk.LabelFrame(main_frame, text="页面范围设置")
        page_range_frame.pack(fill=tk.X, padx=5, pady=5)

        self.extract_all_pages_var = tk.BooleanVar(value=True)

        # 提取所有页面复选框
        all_pages_check = tk.Checkbutton(
            page_range_frame,
            text="提取所有页面的图片",
            variable=self.extract_all_pages_var,
            command=self.toggle_page_range,
        )
        all_pages_check.pack(anchor=tk.W, padx=10, pady=5)

        # 页面范围输入
        self.page_range_frame = ttk.Frame(page_range_frame)
        self.page_range_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(self.page_range_frame, text="页面范围（如：1-3,5）:").pack(
            side=tk.LEFT
        )
        self.extract_page_range_var = tk.StringVar()
        page_range_entry = ttk.Entry(
            self.page_range_frame, textvariable=self.extract_page_range_var
        )
        page_range_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 默认隐藏页面范围输入框
        self.page_range_frame.pack_forget()

        # 创建存储位置设置
        self.extract_image_storage_widgets = (
            self.settings_manager.create_storage_settings(
                main_frame,
                "extract_image_save_location_var",
                "extract_image_folder_path_var",
                "select_extract_image_folder_btn",
                "extract_image_folder_path_entry",
                self.on_extract_image_save_location_change,
            )
        )

        # 配置选择文件夹按钮的命令
        self.extract_image_storage_widgets["select_btn"].config(
            command=lambda: self.settings_manager.select_save_folder(
                self.extract_image_storage_widgets["folder_path_var"]
            )
        )

        # 开始操作按钮
        extract_btn = ttk.Button(
            self.extract_image_storage_widgets["frame"],
            text="开始提取图片",
            command=self.process_extract_images,
        )
        extract_btn.pack(pady=5)

    def select_extract_image_file(self):
        """选择要提取图片的PDF文件"""
        file = filedialog.askopenfilename(
            title="选择PDF文件", filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )

        if file:
            reader = pypdf.PdfReader(file)
            page_count = len(reader.pages)
            self.filename_a = file
            self.extract_image_file_var.set(
                f"{page_count}页 -> {os.path.basename(file)} "
            )

    def toggle_page_range(self):
        """切换页面范围输入框的显示状态"""
        if self.extract_all_pages_var.get():
            self.page_range_frame.pack_forget()
        else:
            self.page_range_frame.pack(fill=tk.X, padx=10, pady=5)

    def on_extract_image_save_location_change(self, *args):
        """图片提取存储位置选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.extract_image_storage_widgets["location_var"].get()
        self.config.set("MergeSettings", "save_location_option", current_value)
        self.save_config()

    def process_extract_images(self):
        """处理图片提取操作"""
        # 检查是否选择了文件
        input_file = getattr(self, "filename_a", None)
        if not input_file:
            messagebox.showwarning("警告", "请先选择PDF文件！")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("错误", "选择的文件不存在！")
            return

        try:
            # 确定保存路径
            save_directory = self.settings_manager.get_save_directory(
                self.extract_image_storage_widgets["location_var"],
                self.extract_image_storage_widgets["folder_path_var"],
            )
            if not save_directory:
                messagebox.showerror("错误", "请选择有效的保存路径！")
                return

            # 创建图片保存目录
            pdf_filename = os.path.splitext(os.path.basename(input_file))[0]
            images_dir = os.path.join(save_directory, f"{pdf_filename}_images")
            os.makedirs(images_dir, exist_ok=True)

            # 获取要处理的页面范围
            pdf_document = fitz.open(input_file)  # 打开PDF文件
            total_pages = len(pdf_document)

            if self.extract_all_pages_var.get():
                # 提取所有页面
                page_numbers = list(range(total_pages))
            else:
                # 解析页面范围
                range_str = self.extract_page_range_var.get()
                if not range_str:
                    messagebox.showwarning("警告", "请输入页面范围！")
                    pdf_document.close()
                    return

                page_numbers = []
                for part in range_str.split(","):
                    part = part.strip()
                    if "-" in part:
                        start, end = map(int, part.split("-"))
                        # 转换为0基索引，并确保范围有效
                        start = max(0, start - 1)
                        end = min(total_pages, end)
                        page_numbers.extend(range(start, end))
                    else:
                        page_num = int(part) - 1
                        if 0 <= page_num < total_pages:
                            page_numbers.append(page_num)

                # 检查页面范围有效性
                for page_num in page_numbers:
                    if page_num < 0 or page_num >= total_pages:
                        pdf_document.close()
                        messagebox.showerror(
                            "错误", f"页码 {page_num+1} 超出范围（1-{total_pages}）！"
                        )
                        return

            # 提取图片
            image_count = 0
            processed_pages = len(page_numbers)

            for i, page_num in enumerate(page_numbers):
                page = pdf_document[page_num]

                # 获取页面的图片列表
                image_list = page.get_images(full=True)

                for img_index, img in enumerate(image_list):
                    xref = img[0]  # 获取图片的xref引用

                    # 提取图片
                    base_image = pdf_document.extract_image(xref)
                    if base_image:
                        image_bytes = base_image["image"]
                        image_ext = base_image["ext"]

                        # 生成图片文件名
                        image_filename = (
                            f"page_{page_num+1:03d}_img_{img_index+1:03d}.{image_ext}"
                        )
                        image_path = os.path.join(images_dir, image_filename)

                        # 保存图片
                        with open(image_path, "wb") as img_file:
                            img_file.write(image_bytes)

                        image_count += 1

                # 可选：显示进度
                if processed_pages > 10:  # 如果页面较多，显示进度
                    if i % 5 == 0 or i == processed_pages - 1:
                        print(f"正在处理第 {i+1}/{processed_pages} 页...")

            pdf_document.close()

            # 显示结果
            if image_count > 0:
                messagebox.showinfo(
                    "完成",
                    f"✅ 图片提取完成！\n"
                    f"• 共处理 {len(page_numbers)} 页\n"
                    f"• 提取 {image_count} 张图片\n"
                    f"• 保存路径：{images_dir}",
                )
            else:
                messagebox.showinfo(
                    "提示",
                    f"处理完成，但在指定页面中未找到图片。\n"
                    f"处理了 {len(page_numbers)} 页，未找到图片。",
                )

        except Exception as e:
            if "pdf_document" in locals():
                pdf_document.close()
            messagebox.showerror("错误", f"处理PDF文件时发生错误：{str(e)}")

    def setup_encrypt_page(self):
        """设置加密解密页面的内容"""
        # 添加标题
        desc_label = ttk.Label(
            self.encrypt_frame, text="PDF文件加密解密", font=self.default_font
        )
        desc_label.pack(pady=5)

        # 主框架
        main_frame = ttk.Frame(self.encrypt_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 选择PDF文件
        file_frame = ttk.LabelFrame(main_frame, text="选择PDF文件")
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        self.encrypt_file_var = tk.StringVar()
        file_select_frame = ttk.Frame(file_frame)
        file_select_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(
            file_select_frame, text="选择PDF文件", command=self.select_encrypt_file
        ).pack(side=tk.LEFT)
        ttk.Entry(
            file_select_frame, textvariable=self.encrypt_file_var, state="readonly"
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # 加密设置
        encrypt_frame = ttk.LabelFrame(main_frame, text="加密设置")
        encrypt_frame.pack(fill=tk.X, padx=5, pady=5)

        # 密码输入
        password_frame = ttk.Frame(encrypt_frame)
        password_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(password_frame, text="密码:").pack(side=tk.LEFT)
        self.encrypt_password_var = tk.StringVar()
        password_entry = ttk.Entry(
            password_frame, textvariable=self.encrypt_password_var, show="*"
        )
        password_entry.pack(side=tk.LEFT, padx=5)

        # 创建存储位置设置
        self.encrypt_storage_widgets = self.settings_manager.create_storage_settings(
            main_frame,
            "encrypt_save_location_var",
            "encrypt_folder_path_var",
            "select_encrypt_folder_btn",
            "encrypt_folder_path_entry",
            self.on_encrypt_save_location_change,
        )

        # 配置选择文件夹按钮的命令
        self.encrypt_storage_widgets["select_btn"].config(
            command=lambda: self.settings_manager.select_save_folder(
                self.encrypt_storage_widgets["folder_path_var"]
            )
        )

        # 创建文件名设置
        self.encrypt_filename_widgets = self.settings_manager.create_filename_settings(
            main_frame,
            "encrypt_filename_var",
            "encrypt_filename_entry",
            self.on_encrypt_filename_change,
        )

        # 开始操作按钮
        encrypt_btn = ttk.Button(
            self.encrypt_filename_widgets["frame"],
            text="开始加密",
            command=self.process_encrypt,
        )
        encrypt_btn.pack(pady=5)

    def select_encrypt_file(self):
        """选择要加密的PDF文件"""
        file = filedialog.askopenfilename(
            title="选择PDF文件", filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )

        if file:
            reader = pypdf.PdfReader(file)
            page_count = len(reader.pages)
            self.filename_a = file
            self.encrypt_file_var.set(f"{page_count}页 -> {os.path.basename(file)}")

    def on_encrypt_save_location_change(self, *args):
        """加密存储位置选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.encrypt_storage_widgets["location_var"].get()
        self.config.set("MergeSettings", "save_location_option", current_value)
        self.save_config()

    def on_encrypt_filename_change(self, *args):
        """加密文件名选项改变时的处理"""
        # 保存设置到配置文件
        current_value = self.encrypt_filename_widgets["filename_var"].get()
        self.config.set("MergeSettings", "filename_option", current_value)
        self.save_config()

    def process_encrypt(self):
        """处理加密操作"""
        # 检查是否选择了文件
        input_file = getattr(self, "filename_a", None)
        if not input_file:
            messagebox.showwarning("警告", "请先选择PDF文件！")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("错误", "选择的文件不存在！")
            return

        password = self.encrypt_password_var.get()
        if not password:
            messagebox.showwarning("警告", "请输入密码！")
            return

        try:
            # 确定保存路径
            save_directory = self.settings_manager.get_save_directory(
                self.encrypt_storage_widgets["location_var"],
                self.encrypt_storage_widgets["folder_path_var"],
            )
            if not save_directory:
                messagebox.showerror("错误", "请选择有效的保存路径！")
                return

            # 获取输入文件名（不含扩展名）
            input_filename = os.path.splitext(os.path.basename(input_file))[0]

            # 生成文件名
            def get_default_name():
                return f"{input_filename}_encrypted"

            filename = self.settings_manager.get_filename(
                self.encrypt_filename_widgets["filename_var"],
                self.encrypt_filename_widgets["entry"],
                get_default_name,
            )
            if not filename:
                filename = f"{input_filename}_encrypted"

            output_path = os.path.join(save_directory, filename + ".pdf")

            # 检查文件是否已存在
            if os.path.exists(output_path):
                result = messagebox.askyesno(
                    "文件已存在", f"文件 {filename}.pdf 已存在，是否覆盖？"
                )
                if not result:
                    return

            # 读取PDF文件并加密
            reader = pypdf.PdfReader(input_file)
            writer = pypdf.PdfWriter()

            # 复制所有页面
            for page in reader.pages:
                writer.add_page(page)

            # 设置密码
            writer.encrypt(password)

            # 写入文件
            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            messagebox.showinfo("成功", f"PDF文件加密完成！\n保存位置：{output_path}")

        except Exception as e:
            messagebox.showerror("错误", f"处理PDF文件时发生错误：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolApp(root)
    root.mainloop()
