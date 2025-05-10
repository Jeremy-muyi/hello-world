import os
import time
import logging
import tkinter as tk
from tkinter import filedialog
from PIL import Image
import fitz  # PyMuPDF
import comtypes.client
import pythoncom
import customtkinter as ctk
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from logging.handlers import RotatingFileHandler

# 配置参数类
class Config:
    DISPLAY_SECONDS = 10
    FOLDER_PATH = r"\\aq3nas01.ap.lilly.com\aq3nas_mfg.grp\ELM_Workspaces\AutoDisplay"
    LOG_FOLDER = r"\\aq3nas01.ap.lilly.com\aq3nas_mfg.grp\ELM_Workspaces\AutoDisplay\Log"
    BACKGROUND = r"\\aq3nas01.ap.lilly.com\aq3nas_mfg.grp\ELM_Workspaces\AutoDisplay\background.png"
    LOG_MAX_SIZE = 5 * 1024 * 1024  # 5MB
    LOG_BACKUP_COUNT = 5
    RETRY_INTERVAL = 3600  # 1小时
    MAX_RETRIES = 3

    @classmethod
    def validate_paths(cls):
        """验证关键路径是否存在"""
        missing = []
        if not os.path.exists(cls.FOLDER_PATH):
            missing.append(f"文件夹路径: {cls.FOLDER_PATH}")
        if not os.path.exists(cls.BACKGROUND):
            missing.append(f"背景图片: {cls.BACKGROUND}")
        if missing:
            raise FileNotFoundError("以下路径不存在:\n" + "\n".join(missing))

class FileChangeHandler(FileSystemEventHandler):
    """文件系统变更处理器"""
    def __init__(self, app):
        super().__init__()
        self.app = app

    def on_modified(self, event):
        """当检测到文件修改时刷新文件列表"""
        if not event.is_directory:
            self.app.refresh_file_list()

class InfoDisplayApp:
    def __init__(self, root):
        self.root = root
        self._init_resources()  # 先初始化资源
        self._init_logging()    # 接着初始化日志
        self._init_ui()         # 然后初始化UI
        self._init_file_monitor()  # 最后初始化文件监控
        
        try:
            self.show_background()
        except Exception as e:
            self.logger.error(f"初始化背景失败: {str(e)}")
            self._show_error_message("无法加载背景图片")

    def _init_resources(self):
        """初始化资源（确保最先执行）"""
        # 初始化基础属性
        self.files = []
        self.index = 0
        self.current_pdf = None
        self.pdf_page_index = 0
        self.is_running = False
        self.error_files = set()
        self.recent_failed = {}
        
        # 强制初始化图片属性
        self.default_image = None
        self._load_background()

    def _init_logging(self):
        """初始化日志系统"""
        # 确保日志目录存在
        os.makedirs(Config.LOG_FOLDER, exist_ok=True)
        
        self.logger = logging.getLogger("DisplayApp")
        self.logger.setLevel(logging.INFO)
        
        # 防止重复添加handler
        if not self.logger.handlers:
            handler = RotatingFileHandler(
                self._get_log_path(),
                maxBytes=Config.LOG_MAX_SIZE,
                backupCount=Config.LOG_BACKUP_COUNT,
                encoding='utf-8'
            )
            formatter = logging.Formatter('[%(asctime)s] %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)

    def _init_ui(self):
        """初始化用户界面"""
        self.root.attributes('-fullscreen', True)
        self.root.tk.call('tk', 'scaling', 2.0)  # 高DPI适配
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # 控制面板
        self.control_frame = ctk.CTkFrame(self.root, fg_color="#344955", height=40)
        self.control_frame.place(relx=0.5, y=0, anchor="n")
        
        # 控制按钮
        controls = [
            ("开始投屏", self.start_display, "blue"),
            ("结束投屏", self.stop_display, "red"),
            ("选择文件夹", self.choose_folder, "green"),
            ("退出程序", self.quit_program, "#a83232")
        ]
        for text, cmd, color in controls:
            btn = ctk.CTkButton(
                self.control_frame, 
                text=text, 
                command=cmd,
                height=32,
                corner_radius=20,
                fg_color=color
            )
            btn.pack(side="left", padx=10, pady=5)

        # 状态显示
        self.status_label = ctk.CTkLabel(
            self.control_frame, 
            text="状态：未开始", 
            text_color="white"
        )
        self.status_label.pack(side="right", padx=10)

        # 主显示区域
        self.label = ctk.CTkLabel(self.root, text="")
        self.label.pack(expand=True, fill="both")

        self.root.after(1000, self._check_mouse_position)

    def _init_file_monitor(self):
        """初始化文件监控"""
        self.observer = Observer()
        self.observer.schedule(
            FileChangeHandler(self),
            Config.FOLDER_PATH,
            recursive=True
        )
        self.observer.start()

    def _load_background(self):
        """安全加载背景图片"""
        try:
            if not os.path.exists(Config.BACKGROUND):
                raise FileNotFoundError(f"背景图片路径不存在: {Config.BACKGROUND}")
            
            img = Image.open(Config.BACKGROUND)
            self.default_image = img.resize(
                (self.root.winfo_screenwidth(), 
                 self.root.winfo_screenheight()),
                Image.Resampling.LANCZOS
            )
        except Exception as e:
            self.logger.error(f"背景加载失败: {str(e)}")
            # 创建纯色备用背景
            self.default_image = Image.new(
                'RGB', 
                (self.root.winfo_screenwidth(), 
                 self.root.winfo_screenheight()),
                color=(34, 73, 85)  # 深蓝色背景
            )

    def _get_log_path(self):
        """获取当日日志文件路径"""
        date_str = datetime.now().strftime("%Y%m%d")
        return os.path.join(
            Config.LOG_FOLDER, 
            f"display_log_{date_str}.txt"
        )

    def show_background(self):
        """安全显示背景图片"""
        try:
            if self.default_image:
                img = ctk.CTkImage(
                    light_image=self.default_image,
                    size=self.default_image.size
                )
                self.label.configure(image=img)
                self.label.image = img
        except Exception as e:
            self.logger.error(f"显示背景失败: {str(e)}")
            self._show_error_message("无法显示背景")

    def _show_error_message(self, message):
        """显示错误信息"""
        error_label = ctk.CTkLabel(
            self.root,
            text=message,
            text_color="red",
            font=("Arial", 24)
        )
        error_label.place(relx=0.5, rely=0.5, anchor="center")

    # 其余方法保持不变（start_display, stop_display等）

if __name__ == "__main__":
    try:
        Config.validate_paths()
        root = ctk.CTk()
        app = InfoDisplayApp(root)
        root.mainloop()
    except Exception as e:
        error_msg = f"程序启动失败: {str(e)}"
        print(error_msg)
        # 显示GUI错误提示
        root = tk.Tk()
        root.withdraw()
        tk.messagebox.showerror("启动错误", error_msg)
