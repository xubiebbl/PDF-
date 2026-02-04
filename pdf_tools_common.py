import tkinter as tk
from tkinter import ttk, filedialog
import os
from datetime import datetime


class CommonSettingsManager:
    """
    通用设置管理器，用于处理存储位置和文件名设置
    """

    def __init__(self, app, config_section="MergeSettings"):
        self.app = app
        self.config_section = config_section

    def create_storage_settings(self, parent_frame, location_var_name, folder_path_var_name,
                               select_btn_name, entry_name, trace_callback=None):
        """
        创建存储位置设置
        
        Args:
            parent_frame: 父级框架
            location_var_name: 存储位置变量名称
            folder_path_var_name: 文件夹路径变量名称
            select_btn_name: 选择按钮名称
            entry_name: 路径输入框名称
            trace_callback: 位置选项变更回调函数
            
        Returns:
            dict: 包含所有控件的字典
        """
        # 存储设置框架
        storage_frame = tk.LabelFrame(
            parent_frame, text="存储位置设置", fg="blue", font=self.app.default_font
        )
        storage_frame.pack(fill=tk.X, padx=5, pady=5)

        # 存储位置选项变量
        location_var = getattr(self.app, location_var_name, tk.StringVar(value="desktop"))

        # 存储位置单选框 - 水平排列
        radio_frame = ttk.Frame(storage_frame)
        radio_frame.pack(fill=tk.X, padx=10, pady=5)

        desktop_radio = tk.Radiobutton(
            radio_frame,
            text="默认存到桌面",
            variable=location_var,
            value="desktop",
            font=self.app.default_font,
        )
        desktop_radio.pack(side=tk.LEFT, padx=(0, 20))

        custom_radio = tk.Radiobutton(
            radio_frame,
            text="选择存储位置",
            variable=location_var,
            value="custom",
            font=self.app.default_font,
        )
        custom_radio.pack(side=tk.LEFT)

        # 文件夹选择和显示
        folder_frame = ttk.Frame(storage_frame)
        folder_frame.pack(fill=tk.X, padx=10, pady=5)

        # 选择文件夹按钮
        select_folder_btn = ttk.Button(
            folder_frame,
            text="选择文件夹",
            state=tk.DISABLED,
        )
        select_folder_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 显示选择的文件夹路径
        folder_path_var = tk.StringVar()
        folder_path_entry = ttk.Entry(
            folder_frame,
            textvariable=folder_path_var,
            state="readonly",
            width=50,
        )
        folder_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 更新按钮状态
        def update_button_state(*args):
            if location_var.get() == "desktop":
                select_folder_btn.config(state=tk.DISABLED)
                folder_path_var.set("")
            else:
                select_folder_btn.config(state=tk.NORMAL)
                # 恢复之前保存的路径
                saved_path = self.app.config.get(
                    self.config_section, "custom_save_location", fallback=""
                )
                if saved_path:
                    folder_path_var.set(saved_path)

        # 绑定事件
        location_var.trace("w", update_button_state)
        if trace_callback:
            location_var.trace("w", trace_callback)

        # 初始状态更新
        update_button_state()

        return {
            "frame": storage_frame,
            "location_var": location_var,
            "select_btn": select_folder_btn,
            "folder_path_var": folder_path_var,
            "folder_path_entry": folder_path_entry
        }

    def create_filename_settings(self, parent_frame, filename_var_name, entry_name,
                                trace_callback=None):
        """
        创建文件名设置
        
        Args:
            parent_frame: 父级框架
            filename_var_name: 文件名变量名称
            entry_name: 输入框名称
            trace_callback: 文件名选项变更回调函数
            
        Returns:
            dict: 包含所有控件的字典
        """
        # 文件名设置框架
        filename_frame = tk.LabelFrame(
            parent_frame, text="文件名设置", fg="blue", font=self.app.default_font
        )
        filename_frame.pack(fill=tk.X, padx=5, pady=5)

        # 文件名选项变量
        filename_var = getattr(self.app, filename_var_name, tk.StringVar(value="default"))

        # 文件名单选框 - 水平排列
        name_radio_frame = ttk.Frame(filename_frame)
        name_radio_frame.pack(fill=tk.X, padx=10, pady=5)

        default_name_radio = tk.Radiobutton(
            name_radio_frame,
            text="默认文件名",
            variable=filename_var,
            value="default",
            font=self.app.default_font,
        )
        default_name_radio.pack(side=tk.LEFT, padx=(0, 20))

        custom_name_radio = tk.Radiobutton(
            name_radio_frame,
            text="自定义文件名",
            variable=filename_var,
            value="custom",
            font=self.app.default_font,
        )
        custom_name_radio.pack(side=tk.LEFT)

        # 自定义文件名输入框
        name_frame = ttk.Frame(filename_frame)
        name_frame.pack(fill=tk.X, padx=10, pady=5)

        filename_entry = ttk.Entry(name_frame, width=50, state="readonly")
        filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 提示标签
        tip_label = ttk.Label(
            filename_frame, text="提示：不需要写文件后缀名", font=self.app.default_font_i
        )
        tip_label.pack(anchor=tk.W, padx=10, pady=(0, 5))

        # 更新输入框状态
        def update_entry_state(*args):
            current_value = filename_var.get()
            if current_value == "default":
                # 保存当前自定义名称
                custom_name = filename_entry.get()
                if custom_name:
                    self.app.config.set(self.config_section, "custom_filename", custom_name)
                    self.app.save_config()
                # 切换为只读状态并设置默认文件名
                filename_entry.config(state="normal")
                default_name = datetime.now().strftime("%Y%m%d%H%M%S")
                filename_entry.delete(0, tk.END)
                filename_entry.insert(0, default_name)
                filename_entry.config(state="readonly")
            else:
                # 切换为可编辑状态
                filename_entry.config(state="normal")

        # 绑定事件
        filename_var.trace("w", update_entry_state)
        if trace_callback:
            filename_var.trace("w", trace_callback)

        # 初始状态更新
        update_entry_state()

        return {
            "frame": filename_frame,
            "filename_var": filename_var,
            "entry": filename_entry,
            "tip_label": tip_label
        }

    def select_save_folder(self, folder_path_var, config_key="custom_save_location"):
        """
        选择保存文件夹
        
        Args:
            folder_path_var: 文件夹路径变量
            config_key: 配置键名
        """
        folder = filedialog.askdirectory(title="选择保存文件夹")
        if folder:
            # 保存到配置文件
            self.app.config.set(self.config_section, config_key, folder)
            self.app.save_config()

            # 更新显示
            folder_path_var.set(folder)

    def get_save_directory(self, location_var, folder_path_var):
        """
        获取保存目录
        
        Args:
            location_var: 存储位置变量
            folder_path_var: 文件夹路径变量
            
        Returns:
            str: 保存目录路径
        """
        if location_var.get() == "desktop":
            return self.app.desktop_path
        else:
            save_directory = folder_path_var.get()
            if not save_directory or not os.path.exists(save_directory):
                return None
            return save_directory

    def get_filename(self, filename_var, filename_entry, default_name_func):
        """
        获取文件名
        
        Args:
            filename_var: 文件名变量
            filename_entry: 文件名输入框
            default_name_func: 默认文件名生成函数
            
        Returns:
            str: 文件名
        """
        if filename_var.get() == "default":
            return default_name_func()
        else:
            custom_name = filename_entry.get()
            if not custom_name:
                return None
            return custom_name