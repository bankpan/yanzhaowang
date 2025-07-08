#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
研究生招生信息爬虫 - 图形界面管理程序
提供可视化的进度监控和控制功能
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
import json
import os
from datetime import datetime
from yanzhao_scraper_fixed import YanZhaoScraperFixed, MAJOR_CONFIG

class ScraperGUIWrapper(YanZhaoScraperFixed):
    """爬虫GUI包装类，处理文件占用等GUI特定问题"""
    
    def __init__(self, gui_instance, *args, **kwargs):
        self.gui_instance = gui_instance
        super().__init__(*args, **kwargs)
    
    def wait_for_file_access_confirmation(self, filename, retry_count, max_retries):
        """GUI版本的文件访问确认"""
        # 在主线程中显示消息框
        def show_file_occupied_dialog():
            result = messagebox.askretrycancel(
                "文件被占用", 
                f"文件 {filename} 正在被其他程序使用（可能是Excel）。\n\n"
                f"请关闭该文件后点击\"重试\"继续保存，或点击\"取消\"跳过本次保存。\n\n"
                f"当前重试次数：{retry_count}/{max_retries}",
                icon='warning'
            )
            return result
        
        # 由于这可能在工作线程中调用，需要在主线程中显示对话框
        if self.gui_instance:
            self.gui_instance.root.after(0, lambda: self._handle_file_occupied_dialog(filename, retry_count, max_retries))
            
            # 等待用户响应（通过共享变量）
            timeout = 60  # 60秒超时
            wait_time = 0
            while not hasattr(self, '_file_dialog_response') and wait_time < timeout:
                time.sleep(0.5)
                wait_time += 0.5
            
            # 获取响应并清理
            if hasattr(self, '_file_fdadialog_response'):
                response = self._file_dialog_response
                delattr(self, '_file_dialog_response')
                if not response:  # 用户选择取消
                    raise Exception("用户取消文件保存")
        else:
            time.sleep(2)  # 后备方案
    
    def _handle_file_occupied_dialog(self, filename, retry_count, max_retries):
        """在主线程中处理文件占用对话框"""
        result = messagebox.askretrycancel(
            "文件被占用", 
            f"文件 {filename} 正在被其他程序使用（可能是Excel）。\n\n"
            f"请关闭该文件后点击\"重试\"继续保存，或点击\"取消\"跳过本次保存。\n\n"
            f"当前重试次数：{retry_count}/{max_retries}",
            icon='warning'
        )
        self._file_dialog_response = result

class ScraperGUI:
    def __init__(self, root):
        """初始化GUI界面"""
        self.root = root
        self.root.title("研究生招生信息爬虫管理器")
        self.root.geometry("900x800")  # 增加窗口大小以容纳更多内容
        self.root.minsize(800, 700)    # 增加最小窗口大小
        
        # 爬虫实例
        self.scraper = None
        self.scraper_thread = None
        
        # 状态变量
        self.is_running = False
        self.is_paused = False
        self.start_time = None  # 开始运行时间
        self.paused_time = 0    # 暂停累计时间
        self.pause_start_time = None  # 暂停开始时间
        self.current_major = "125300"  # 默认专业
        
        # 创建GUI界面
        self.create_widgets()
        
        # 定时更新状态
        self.update_display()
        
    def on_major_changed(self, event=None):
        """专业选择改变时的回调"""
        try:
            selected = self.major_var.get()
            if " - " in selected:
                major_code = selected.split(" - ")[0]
                self.current_major = major_code
                major_name = MAJOR_CONFIG[major_code]["name"]
                
                self.log_message(f"已切换到专业：{major_name} ({major_code})", "info")
                
                # 检查该专业是否有现有数据
                self.check_existing_data_for_major(major_code)
                
                # 更新页面范围显示
                self.update_page_range_for_major(major_code)
                
        except Exception as e:
            self.log_message(f"切换专业失败: {e}", "error")
    
    def update_page_range_for_major(self, major_code):
        """根据专业更新页面范围显示"""
        try:
            progress_file = f'progress_{major_code}.json'
            
            # 重置页面显示信息
            self.page_info.config(text="未开始")
            self.records_info.config(text="0 条")
            self.status_info.config(text="就绪", foreground="green")
            
            if os.path.exists(progress_file):
                # 如果有进度文件，使用保存的信息
                with open(progress_file, 'r', encoding='utf-8') as f:
                    progress = json.load(f)
                    current_page = progress.get('current_page', 1)
                    total_pages = progress.get('total_pages', 1)
                    records_count = progress.get('records_count', 0)
                    
                    # 更新起始页面为当前进度
                    self.start_page_var.set(str(current_page))
                    
                    # 更新显示信息
                    self.page_info.config(text=f"第{current_page}页 / 共{total_pages}页")
                    self.records_info.config(text=f"{records_count} 条")
                    
                    # 如果有进度，自动选择继续模式
                    if current_page > 1 or records_count > 0:
                        self.mode_var.set("continue")
                        self.status_info.config(text="有进度", foreground="orange")
                    
                    self.log_message(f"已加载该专业的进度：从第{current_page}页继续，共{total_pages}页", "info")
            else:
                # 如果没有进度文件，重置为默认值
                self.start_page_var.set("1")
                self.mode_var.set("restart")
                self.log_message(f"该专业无进度记录，将从第1页开始", "info")
                
        except Exception as e:
            self.log_message(f"更新页面范围失败: {e}", "error")
    
    def check_existing_data_for_major(self, major_code):
        """检查指定专业是否有现有数据"""
        try:
            major_name = MAJOR_CONFIG[major_code]["name"]
            excel_file = f"研究生招生信息_{major_name}.xlsx"
            progress_file = f'progress_{major_code}.json'
            
            has_excel = os.path.exists(excel_file)
            has_progress = os.path.exists(progress_file)
            
            if has_excel or has_progress:
                if has_excel:
                    self.log_message(f"发现该专业的数据文件：{excel_file}", "info")
                
                if has_progress:
                    with open(progress_file, 'r', encoding='utf-8') as f:
                        progress = json.load(f)
                        current_page = progress.get('current_page', 1)
                        total_pages = progress.get('total_pages', 1)
                        records_count = progress.get('records_count', 0)
                        
                        self.log_message(f"该专业进度：第{current_page}页/共{total_pages}页，已获取{records_count}条记录", "info")
            else:
                self.log_message(f"该专业暂无数据文件", "info")
                
        except Exception as e:
            self.log_message(f"检查专业数据失败: {e}", "error")
        
    def create_widgets(self):
        """创建GUI组件"""
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="研究生招生信息爬虫管理器", 
                               font=("微软雅黑", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 进度显示区域
        progress_frame = ttk.LabelFrame(main_frame, text="进度信息", padding="10")
        progress_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(1, weight=1)
        
        # 进度条
        ttk.Label(progress_frame, text="总体进度:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                          maximum=100, length=300)
        self.progress_bar.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.progress_text = ttk.Label(progress_frame, text="0%")
        self.progress_text.grid(row=0, column=2, sticky=tk.W)
        
        # 状态信息
        info_grid_row = 1
        ttk.Label(progress_frame, text="当前页面:").grid(row=info_grid_row, column=0, sticky=tk.W, pady=(5, 0))
        self.page_info = ttk.Label(progress_frame, text="未开始")
        self.page_info.grid(row=info_grid_row, column=1, sticky=tk.W, pady=(5, 0))
        
        info_grid_row += 1
        ttk.Label(progress_frame, text="已获取记录:").grid(row=info_grid_row, column=0, sticky=tk.W, pady=(5, 0))
        self.records_info = ttk.Label(progress_frame, text="0 条")
        self.records_info.grid(row=info_grid_row, column=1, sticky=tk.W, pady=(5, 0))
        
        info_grid_row += 1
        ttk.Label(progress_frame, text="当前状态:").grid(row=info_grid_row, column=0, sticky=tk.W, pady=(5, 0))
        self.status_info = ttk.Label(progress_frame, text="就绪", foreground="green")
        self.status_info.grid(row=info_grid_row, column=1, sticky=tk.W, pady=(5, 0))
        
        # 添加运行时间显示
        info_grid_row += 1
        ttk.Label(progress_frame, text="运行时间:").grid(row=info_grid_row, column=0, sticky=tk.W, pady=(5, 0))
        self.runtime_info = ttk.Label(progress_frame, text="00:00:00", foreground="blue")
        self.runtime_info.grid(row=info_grid_row, column=1, sticky=tk.W, pady=(5, 0))
        
        # 控制按钮区域
        control_frame = ttk.LabelFrame(main_frame, text="控制面板", padding="10")
        control_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 按钮网格
        self.start_button = ttk.Button(control_frame, text="开始爬取", command=self.start_scraping)
        self.start_button.grid(row=0, column=0, padx=(0, 10), pady=5)
        
        self.pause_button = ttk.Button(control_frame, text="暂停", command=self.pause_scraping, state=tk.DISABLED)
        self.pause_button.grid(row=0, column=1, padx=(0, 10), pady=5)
        
        self.stop_button = ttk.Button(control_frame, text="停止", command=self.stop_scraping, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=2, padx=(0, 10), pady=5)
        
        self.progress_button = ttk.Button(control_frame, text="查看进度", command=self.view_progress)
        self.progress_button.grid(row=0, column=3, padx=(0, 10), pady=5)
        
        # 设置选项区域
        settings_frame = ttk.LabelFrame(main_frame, text="运行设置", padding="10")
        settings_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 专业选择
        ttk.Label(settings_frame, text="选择专业:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.major_var = tk.StringVar(value="125300")
        major_frame = ttk.Frame(settings_frame)
        major_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.major_combobox = ttk.Combobox(major_frame, textvariable=self.major_var, width=30, state="readonly")
        major_options = [f"{code} - {info['name']}" for code, info in MAJOR_CONFIG.items()]
        self.major_combobox['values'] = major_options
        self.major_combobox.set("125300 - 会计专硕")  # 默认选择
        self.major_combobox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.major_combobox.bind('<<ComboboxSelected>>', self.on_major_changed)
        
        # 运行模式选择
        ttk.Label(settings_frame, text="运行模式:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.mode_var = tk.StringVar(value="continue")
        mode_frame = ttk.Frame(settings_frame)
        mode_frame.grid(row=1, column=1, sticky=tk.W)
        
        ttk.Radiobutton(mode_frame, text="继续之前的任务", variable=self.mode_var, 
                       value="continue").grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        ttk.Radiobutton(mode_frame, text="重新开始", variable=self.mode_var, 
                       value="restart").grid(row=0, column=1, sticky=tk.W, padx=(0, 15))
        ttk.Radiobutton(mode_frame, text="测试运行", variable=self.mode_var, 
                       value="test").grid(row=0, column=2, sticky=tk.W)
        
        # 页面范围设置
        range_frame = ttk.Frame(settings_frame)
        range_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(range_frame, text="页面范围:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        ttk.Label(range_frame, text="从第").grid(row=0, column=1, sticky=tk.W, padx=(0, 5))
        self.start_page_var = tk.StringVar(value="1")
        start_page_entry = ttk.Entry(range_frame, textvariable=self.start_page_var, width=5)
        start_page_entry.grid(row=0, column=2, padx=(0, 5))
        
        ttk.Label(range_frame, text="页到第").grid(row=0, column=3, sticky=tk.W, padx=(0, 5))
        self.end_page_var = tk.StringVar(value="")  # 默认为空，自动检测
        end_page_entry = ttk.Entry(range_frame, textvariable=self.end_page_var, width=5)
        end_page_entry.grid(row=0, column=4, padx=(0, 5))
        
        ttk.Label(range_frame, text="页 (空白=自动检测, 测试模式限制每页").grid(row=0, column=5, sticky=tk.W, padx=(0, 5))
        self.test_limit_var = tk.StringVar(value="2")
        test_limit_entry = ttk.Entry(range_frame, textvariable=self.test_limit_var, width=3)
        test_limit_entry.grid(row=0, column=6, padx=(0, 5))
        
        ttk.Label(range_frame, text="个院校)").grid(row=0, column=7, sticky=tk.W)
        
        # 无头模式选项
        self.headless_var = tk.BooleanVar(value=True)  # 默认选中无头模式
        headless_check = ttk.Checkbutton(range_frame, text="无头模式（后台运行）", variable=self.headless_var)
        headless_check.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))

        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="10")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # 日志文本框 - 调整高度并确保正确配置
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=85, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 清空日志按钮
        ttk.Button(log_frame, text="清空日志", command=self.clear_log).grid(row=1, column=0, pady=(5, 0))
        
        # 状态栏
        self.status_bar = ttk.Label(main_frame, text="就绪", relief=tk.SUNKEN)
        self.status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 检查已有数据
        self.check_existing_data()
        
        # 设置初始显示值，避免显示nan
        self.progress_var.set(0)
        self.progress_text.config(text="0%")
        if not hasattr(self, '_initial_page_set'):
            self.page_info.config(text="第1页 / 共33页")
            self._initial_page_set = True
            
    def check_existing_data(self):
        """检查当前专业的已有数据并更新界面"""
        try:
            # 检查当前选择的专业的数据
            self.check_existing_data_for_major(self.current_major)
                
        except Exception as e:
            self.log_message(f"检查已有数据时出错: {e}", "error")
            
    def log_message(self, message, level="info"):
        """在日志区域显示消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 根据级别设置颜色
        color = "black"
        if level == "error":
            color = "red"
        elif level == "warning":
            color = "orange"
        elif level == "success":
            color = "green"
            
        # 插入消息
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        
        # 设置颜色（如果需要的话）
        if color != "black":
            start_line = self.log_text.index("end-2c linestart")
            end_line = self.log_text.index("end-1c")
            self.log_text.tag_add(level, start_line, end_line)
            self.log_text.tag_config(level, foreground=color)
            
        # 自动滚动到底部
        self.log_text.see(tk.END)
        
        # 更新状态栏
        self.status_bar.config(text=message)
        
        # 刷新界面
        self.root.update()
        
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
        
    def get_runtime_seconds(self):
        """获取实际运行时间（秒）"""
        if not self.start_time:
            return 0
            
        # 计算当前运行时间
        if self.is_paused and self.pause_start_time:
            # 如果当前暂停中，计算到暂停开始时的时间
            current_runtime = self.pause_start_time - self.start_time
        else:
            # 如果正在运行，计算到当前时间
            current_runtime = time.time() - self.start_time
            
        # 减去总暂停时间
        total_runtime = current_runtime - self.paused_time
        return max(0, total_runtime)  # 确保不会是负数
        
    def format_runtime(self, seconds):
        """格式化运行时间为 HH:MM:SS 格式"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
    def update_runtime_display(self):
        """更新运行时间显示"""
        runtime_seconds = self.get_runtime_seconds()
        runtime_text = self.format_runtime(runtime_seconds)
        
        # 根据状态设置颜色
        if self.is_paused:
            color = "orange"
        elif self.is_running:
            color = "blue"
        else:
            color = "gray"
            
        self.runtime_info.config(text=runtime_text, foreground=color)
        
    def progress_callback(self, progress_data):
        """进度更新回调函数"""
        # 验证数据有效性，防止显示nan
        current_page = progress_data.get('current_page', 1)
        total_pages = progress_data.get('total_pages', 33)
        records_count = progress_data.get('records_count', 0)
        progress_percentage = progress_data.get('progress_percentage', 0)
        
        # 确保数值有效
        if not isinstance(current_page, (int, float)) or current_page != current_page:  # 检查NaN
            current_page = 1
        if not isinstance(total_pages, (int, float)) or total_pages != total_pages:
            total_pages = 33
        if not isinstance(records_count, (int, float)) or records_count != records_count:
            records_count = 0
        if not isinstance(progress_percentage, (int, float)) or progress_percentage != progress_percentage:
            progress_percentage = 0
        
        # 更新界面显示
        self.progress_var.set(progress_percentage)
        self.progress_text.config(text=f"{progress_percentage:.1f}%")
        self.page_info.config(text=f"第{int(current_page)}页 / 共{int(total_pages)}页")
        self.records_info.config(text=f"{int(records_count)} 条")
        
        # 根据状态更新状态信息颜色
        status = progress_data.get('status', '运行中')
        status_color = "blue"
        if status in ["完成", "成功"]:
            status_color = "green"
        elif status in ["错误", "失败"]:
            status_color = "red"
        elif status in ["暂停", "警告"]:
            status_color = "orange"
            
        self.status_info.config(text=status, foreground=status_color)
        
    def status_callback(self, message, level="info"):
        """状态更新回调函数"""
        self.log_message(message, level)
        
    def start_scraping(self):
        """开始爬取"""
        try:
            if self.is_running:
                messagebox.showwarning("警告", "爬虫正在运行中！")
                return
            
            # 获取专业代码
            selected = self.major_var.get()
            if " - " in selected:
                major_code = selected.split(" - ")[0]
            else:
                major_code = "125300"  # 默认值
                
            # 获取设置参数
            mode = self.mode_var.get()
            start_page = int(self.start_page_var.get())
            end_page = int(self.end_page_var.get()) if self.end_page_var.get() else None
            test_limit = int(self.test_limit_var.get()) if mode == "test" else None
            headless_mode = self.headless_var.get()  # 获取无头模式选项
            
            if start_page < 1:
                messagebox.showerror("错误", "起始页面必须大于0！")
                return
            
            if end_page and end_page < start_page:
                messagebox.showerror("错误", "结束页面不能小于起始页面！")
                return
                
            # 记录运行模式
            mode_text = "无头模式（后台运行）" if headless_mode else "可视模式（显示浏览器）"
            major_name = MAJOR_CONFIG[major_code]["name"]
            self.log_message(f"启动爬虫 - {mode_text} - 专业：{major_name}")
                
            # 创建爬虫实例
            self.scraper = ScraperGUIWrapper(
                gui_instance=self,
                progress_callback=self.progress_callback,
                status_callback=self.status_callback,
                headless=headless_mode,
                major_code=major_code
            )
            
            # 根据模式调整参数
            if mode == "restart":
                self.log_message("重新开始任务，清空已有数据...")
                self.scraper.data = []
                self.scraper.current_page = 1
                
                # 清空该专业的进度文件
                progress_file = f'progress_{major_code}.json'
                if os.path.exists(progress_file):
                    os.remove(progress_file)
                    self.log_message(f"已清空专业{major_code}的进度记录")
                
                # 删除该专业的Excel文件（重新开始时）
                excel_file = f"研究生招生信息_{MAJOR_CONFIG[major_code]['name']}.xlsx"
                if os.path.exists(excel_file):
                    try:
                        os.remove(excel_file)
                        self.log_message(f"已删除原有数据文件：{excel_file}")
                    except Exception as e:
                        self.log_message(f"删除数据文件失败：{e}，将在保存时重写", "warning")
                
                self.log_message("重新开始，将从第1页开始爬取")
                        
            elif mode == "test":
                self.log_message(f"测试模式：处理第{start_page}页，限制{test_limit}个院校")
                
            elif mode == "continue":
                self.log_message("继续之前的任务...")
                start_page = self.scraper.current_page
                self.start_page_var.set(str(start_page))
                
            # 更新界面状态
            self.is_running = True
            self.is_paused = False
            
            # 记录开始时间和重置暂停时间
            self.start_time = time.time()
            self.paused_time = 0
            self.pause_start_time = None
            
            self.start_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.NORMAL, text="暂停")
            self.stop_button.config(state=tk.NORMAL)
            
            # 启动爬虫线程
            self.scraper_thread = threading.Thread(
                target=self.run_scraper,
                args=(start_page, end_page, test_limit),
                daemon=True
            )
            self.scraper_thread.start()
            
            self.log_message("爬虫已启动", "success")
            
        except ValueError:
            messagebox.showerror("错误", "页面范围设置必须是数字！")
        except Exception as e:
            messagebox.showerror("错误", f"启动爬虫时出错: {e}")
            self.log_message(f"启动爬虫时出错: {e}", "error")
            
    def run_scraper(self, start_page, end_page, test_limit):
        """在线程中运行爬虫"""
        try:
            success = self.scraper.run(
                start_page=start_page, 
                end_page=end_page, 
                max_universities_per_page=test_limit
            )
            
            if success:
                self.log_message("爬虫运行完成！", "success")
            else:
                self.log_message("爬虫运行失败！", "error")
                
        except Exception as e:
            self.log_message(f"爬虫运行异常: {e}", "error")
            
        finally:
            # 重置界面状态
            self.is_running = False
            self.is_paused = False
            
            # 重置时间状态
            if self.pause_start_time:
                # 如果在暂停状态下结束，计算最后的暂停时间
                self.paused_time += time.time() - self.pause_start_time
                self.pause_start_time = None
            
            self.start_button.config(state=tk.NORMAL)
            self.pause_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)
            
    def pause_scraping(self):
        """暂停/继续爬取"""
        if not self.scraper or not self.is_running:
            return
            
        if not self.is_paused:
            # 暂停
            self.scraper.pause()
            self.is_paused = True
            self.pause_button.config(text="继续")
            self.log_message("爬虫已暂停", "warning")
            
            # 记录暂停开始时间
            self.pause_start_time = time.time()
            
        else:
            # 继续
            self.scraper.resume()
            self.is_paused = False
            self.pause_button.config(text="暂停")
            self.log_message("爬虫已继续", "success")
            
            # 计算暂停时间并累加
            if self.pause_start_time:
                self.paused_time += time.time() - self.pause_start_time
                self.pause_start_time = None # 重置暂停开始时间
            
    def stop_scraping(self):
        """停止爬取"""
        if not self.scraper or not self.is_running:
            return
            
        result = messagebox.askyesno("确认", "确定要停止爬虫吗？已获取的数据将会保存。")
        if result:
            self.scraper.stop()
            self.log_message("正在停止爬虫...", "warning")
            
    def view_progress(self):
        """查看详细进度"""
        try:
            if os.path.exists('progress_fixed.json'):
                with open('progress_fixed.json', 'r', encoding='utf-8') as f:
                    progress_data = json.load(f)
                    
                # 创建进度查看窗口
                progress_window = tk.Toplevel(self.root)
                progress_window.title("详细进度信息")
                progress_window.geometry("500x400")
                
                # 进度信息文本
                progress_text = scrolledtext.ScrolledText(progress_window, height=20, width=60)
                progress_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                # 显示进度信息
                progress_text.insert(tk.END, f"最后更新时间: {progress_data.get('last_update', '未知')}\n")
                progress_text.insert(tk.END, f"当前页面: {progress_data.get('current_page', 0)}\n")
                progress_text.insert(tk.END, f"总页面数: {progress_data.get('total_pages', 33)}\n")
                progress_text.insert(tk.END, f"总记录数: {progress_data.get('total_records', 0)}\n")
                progress_text.insert(tk.END, f"完成百分比: {progress_data.get('progress_percentage', 0):.2f}%\n")
                progress_text.insert(tk.END, f"状态: {progress_data.get('status', '未知')}\n\n")
                
                progress_text.insert(tk.END, "已完成页面:\n")
                completed_pages = progress_data.get('completed_pages', [])
                if completed_pages:
                    progress_text.insert(tk.END, f"{', '.join(map(str, completed_pages))}\n\n")
                else:
                    progress_text.insert(tk.END, "无\n\n")
                    
                progress_text.insert(tk.END, "各页面数据统计:\n")
                data_count_by_page = progress_data.get('data_count_by_page', {})
                for page, count in sorted(data_count_by_page.items(), key=lambda x: int(x[0])):
                    progress_text.insert(tk.END, f"第{page}页: {count}条记录\n")
                    
                progress_text.config(state=tk.DISABLED)
                
            else:
                messagebox.showinfo("信息", "未找到进度文件")
                
        except Exception as e:
            messagebox.showerror("错误", f"查看进度时出错: {e}")
            
    def update_display(self):
        """定时更新显示"""
        try:
            # 更新运行时间显示
            self.update_runtime_display()
            
            # 如果爬虫正在运行，更新状态
            if self.scraper and self.is_running:
                status = self.scraper.get_status()
                if not status['is_running'] and self.is_running:
                    # 爬虫已结束
                    self.is_running = False
                    self.is_paused = False
                    self.start_button.config(state=tk.NORMAL)
                    self.pause_button.config(state=tk.DISABLED)
                    self.stop_button.config(state=tk.DISABLED)
                    
        except Exception as e:
            pass  # 忽略更新错误
            
        # 继续定时更新
        self.root.after(1000, self.update_display)
        
    def on_closing(self):
        """窗口关闭事件"""
        if self.is_running:
            result = messagebox.askyesnocancel("确认", "爬虫正在运行，是否停止并退出？")
            if result is None:  # 取消
                return
            elif result:  # 是，停止并退出
                if self.scraper:
                    self.scraper.stop()
                # 等待一下让爬虫保存数据
                time.sleep(2)
                
        self.root.destroy()

def main():
    """主函数"""
    root = tk.Tk()
    app = ScraperGUI(root)
    
    # 设置窗口关闭事件
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # 启动GUI
    root.mainloop()

if __name__ == "__main__":
    main() 