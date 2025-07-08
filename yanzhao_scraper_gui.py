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
# import json  # 不再需要，已去掉进度文件依赖
import os
from datetime import datetime
from yanzhao_scraper_fixed import YanZhaoScraperFixed, MAJOR_CONFIG

class ScraperGUIWrapper(YanZhaoScraperFixed):
    """对爬虫类的包装，用于GUI集成"""
    
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
    
    def login_and_navigate(self):
        """重写登录方法，登录成功后更新界面显示页面范围"""
        # 调用父类的登录方法
        result = super().login_and_navigate()
        
        # 登录成功后，更新界面显示实际的页面范围
        if result is not False:  # 登录成功或无需登录
            self.gui_instance.root.after(0, self._update_gui_page_range)
        
        return result
    
    def _update_gui_page_range(self):
        """更新GUI界面的页面范围显示"""
        try:
            # 重新分析Excel文件来确定准确的起始页
            major_name = MAJOR_CONFIG[self.major_code]["name"]
            study_mode_name = "全日制" if self.study_mode == "1" else "非全日制"
            excel_file = f"研究生招生信息_{major_name}_{study_mode_name}.xlsx"
            
            if os.path.exists(excel_file):
                # 重新分析Excel文件（登录后的精确分析）
                start_page, records_count, status_msg = self.gui_instance.analyze_excel_data(excel_file)
                self.current_page = start_page
                self.gui_instance.log_message(f"登录后重新分析Excel文件: {status_msg}", "info")
            else:
                # 没有Excel文件，从第1页开始
                start_page = 1
                records_count = 0
                self.current_page = start_page
                self.gui_instance.log_message("没有发现历史数据文件，将从第1页开始爬取", "info")
            
            # 获取总页数
            total_pages = self.total_pages
            
            # 全面更新界面显示
            self.gui_instance.start_page_var.set(str(start_page))
            self.gui_instance.end_page_var.set(str(total_pages))  # 更新结束页
            self.gui_instance.page_info.config(text=f"第{start_page}页 / 共{total_pages}页")
            self.gui_instance.records_info.config(text=f"{records_count} 条")
            
            # 更新进度条（基于已完成的页数）
            if total_pages > 0:
                completed_pages = max(0, start_page - 1)  # 已完成的页数
                progress_percentage = (completed_pages / total_pages) * 100
                self.gui_instance.progress_var.set(progress_percentage)
                self.gui_instance.progress_text.config(text=f"{progress_percentage:.1f}%")
            else:
                self.gui_instance.progress_var.set(0)
                self.gui_instance.progress_text.config(text="0%")
            
            # 更新状态信息
            if records_count > 0:
                self.gui_instance.status_info.config(text="有进度数据", foreground="orange")
            else:
                self.gui_instance.status_info.config(text="准备就绪", foreground="green")
            
            # 更新状态栏
            self.gui_instance.status_bar.config(text=f"登录成功！页面范围：第{start_page}-{total_pages}页，已有{records_count}条记录")
            
            # 记录详细信息到日志
            self.gui_instance.log_message(f"✓ 登录成功！总页数：{total_pages}页", "success")
            self.gui_instance.log_message(f"✓ 起始页：第{start_page}页，已有记录：{records_count}条", "success")
            self.gui_instance.log_message(f"✓ 页面范围：第{start_page}页到第{total_pages}页", "success")
            
        except Exception as e:
            self.gui_instance.log_message(f"更新页面范围显示失败: {e}", "error")

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
        
        # 控件变量
        self.major_var = tk.StringVar(value="125300 - 会计专硕")
        self.study_mode_var = tk.StringVar(value="1")  # 学习方式：1=全日制，2=非全日制
        self.start_page_var = tk.StringVar(value="待确定")
        self.end_page_var = tk.StringVar(value="")
        self.test_limit_var = tk.StringVar(value="2")
        self.mode_var = tk.StringVar(value="continue")
        self.headless_var = tk.BooleanVar(value=False)
        
        # 内部检测到的数据（登录后使用）
        self.detected_start_page = 1
        self.detected_records_count = 0
        
        # 创建GUI界面
        self.create_widgets()
        
        # 定时更新状态
        self.update_display()
        
    def on_major_changed(self, event=None):
        """专业选择改变时的处理"""
        try:
            selected = self.major_var.get()
            if " - " in selected:
                major_code = selected.split(" - ")[0]
                self.current_major = major_code
                
                # 检查该专业的现有数据
                self.update_page_range_for_major(major_code)
                
                self.log_message(f"已切换到专业: {selected}", "info")
            else:
                self.log_message("专业选择格式错误", "error")
        except Exception as e:
            self.log_message(f"切换专业失败: {e}", "error")
    
    def on_study_mode_changed(self):
        """学习方式改变时的处理"""
        try:
            study_mode = self.study_mode_var.get()
            study_mode_name = "全日制" if study_mode == "1" else "非全日制"
            
            # 重新检查当前专业的数据（因为文件名包含学习方式）
            if hasattr(self, 'current_major'):
                self.update_page_range_for_major(self.current_major)
            
            self.log_message(f"已切换学习方式: {study_mode_name}", "info")
        except Exception as e:
            self.log_message(f"切换学习方式失败: {e}", "error")
    
    def update_page_range_for_major(self, major_code):
        """根据专业更新页面范围显示"""
        try:
            # 直接调用check_existing_data_for_major来更新所有显示
            self.check_existing_data_for_major(major_code)
                
        except Exception as e:
            self.log_message(f"更新页面范围失败: {e}", "error")
    
    def analyze_excel_data(self, excel_file):
        """分析Excel文件数据，返回起始页和记录数（基于院校数量判断页面完整性）"""
        try:
            import pandas as pd
            df = pd.read_excel(excel_file)
            data = df.to_dict('records')
            records_count = len(data)
            
            if not data:
                return 1, 0, "文件为空，将从第1页开始"
            
            # 获取最后一条记录的页码
            last_record = data[-1]  # 最后一条记录
            last_page = last_record.get('页码', 1)
            
            if last_page <= 0:
                return 1, records_count, f"最后一条记录页码无效，将从第1页开始"
            
            # 统计最后一页的院校数量（通过院校名称去重）
            last_page_records = [r for r in data if r.get('页码') == last_page]
            last_page_universities = set()
            
            for record in last_page_records:
                university_name = record.get('招生单位', '') or record.get('院校名称', '')
                if university_name:
                    last_page_universities.add(university_name)
            
            university_count = len(last_page_universities)
            
            # 判断页面完整性：如果院校数量=10个，说明该页完整，从下一页开始
            # 如果院校数量<10个，说明该页不完整，从当前页重新开始
            if university_count >= 10:
                current_page = last_page + 1
                status_msg = f"第{last_page}页有{university_count}个院校（完整），将从第{current_page}页继续爬取"
            else:
                current_page = last_page
                status_msg = f"第{last_page}页仅有{university_count}个院校（不完整），将从第{current_page}页重新爬取"
            
            return current_page, records_count, status_msg
            
        except Exception as e:
            return 1, 0, f"读取Excel文件失败: {e}"

    def check_existing_data_for_major(self, major_code):
        """检查指定专业是否有现有数据"""
        try:
            major_name = MAJOR_CONFIG[major_code]["name"]
            study_mode = self.study_mode_var.get()
            study_mode_name = "全日制" if study_mode == "1" else "非全日制"
            excel_file = f"研究生招生信息_{major_name}_{study_mode_name}.xlsx"
            
            # 检查Excel文件是否存在
            if os.path.exists(excel_file):
                # 分析Excel文件数据（仅用于内部记录，不在界面显示具体页码）
                current_page, records_count, status_msg = self.analyze_excel_data(excel_file)
                
                # 记录分析结果到内部变量，但界面暂不显示具体页码
                self.detected_start_page = current_page
                self.detected_records_count = records_count
                
                self.log_message(f"发现该专业({study_mode_name})的数据文件：{excel_file}，包含{records_count}条记录", "info")
                self.log_message(status_msg, "info")
                
                # 立即更新界面显示起始页（不需要等登录）
                self.detected_start_page = current_page
                self.detected_records_count = records_count
                
                # 立即更新界面显示
                self.start_page_var.set(str(current_page))
                self.page_info.config(text=f"第{current_page}页 / 总页数待检测")
                self.records_info.config(text=f"{records_count} 条")
                
                # 如果有数据，自动选择继续模式
                if current_page > 1 or records_count > 0:
                    self.mode_var.set("continue")
                    self.status_info.config(text="有进度", foreground="orange")
                else:
                    self.status_info.config(text="就绪", foreground="green")
            else:
                # 没有数据文件时
                self.detected_start_page = 1
                self.detected_records_count = 0
                self.status_info.config(text="就绪", foreground="green")
                self.log_message(f"该专业({study_mode_name})暂无数据文件", "info")
                
                # 立即更新界面显示
                self.start_page_var.set("1")
                self.page_info.config(text="第1页 / 总页数待检测")
                self.records_info.config(text="0 条")
                
        except Exception as e:
            self.log_message(f"检查专业数据失败: {e}", "error")
            # 出错时的默认显示
            self.detected_start_page = 1
            self.detected_records_count = 0
            self.records_info.config(text="0 条")
            self.page_info.config(text="待检测...")
            self.start_page_var.set("待确定")
            self.mode_var.set("restart")
            self.status_info.config(text="就绪", foreground="green")
        
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
        major_frame = ttk.Frame(settings_frame)
        major_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        
        major_combobox = ttk.Combobox(major_frame, textvariable=self.major_var, width=20, state="readonly")
        major_combobox['values'] = [f"{code} - {info['name']}" for code, info in MAJOR_CONFIG.items()]
        major_combobox.grid(row=0, column=0, sticky=tk.W)
        major_combobox.bind('<<ComboboxSelected>>', self.on_major_changed)
        
        # 学习方式选择
        ttk.Label(settings_frame, text="学习方式:").grid(row=0, column=2, sticky=tk.W, padx=(20, 10))
        study_mode_frame = ttk.Frame(settings_frame)
        study_mode_frame.grid(row=0, column=3, sticky=tk.W, pady=(0, 10))
        
        fulltime_radio = ttk.Radiobutton(study_mode_frame, text="全日制", variable=self.study_mode_var, 
                                        value="1", command=self.on_study_mode_changed)
        fulltime_radio.grid(row=0, column=0, padx=(0, 10))
        
        parttime_radio = ttk.Radiobutton(study_mode_frame, text="非全日制", variable=self.study_mode_var, 
                                        value="2", command=self.on_study_mode_changed)
        parttime_radio.grid(row=0, column=1)
        
        # 运行模式选择
        ttk.Label(settings_frame, text="运行模式:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
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
        start_page_entry = ttk.Entry(range_frame, textvariable=self.start_page_var, width=5)
        start_page_entry.grid(row=0, column=2, padx=(0, 5))
        
        ttk.Label(range_frame, text="页到第").grid(row=0, column=3, sticky=tk.W, padx=(0, 5))
        end_page_entry = ttk.Entry(range_frame, textvariable=self.end_page_var, width=5)
        end_page_entry.grid(row=0, column=4, padx=(0, 5))
        
        ttk.Label(range_frame, text="页 (空白=自动检测, 测试模式限制每页").grid(row=0, column=5, sticky=tk.W, padx=(0, 5))
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
        
        # 检查已有数据并立即确定起始页和进度
        self.check_existing_data()
        self.update_initial_progress_display()
        
        # 确保运行模式默认为"继续之前的任务"
        self.mode_var.set("continue")
        self.log_message("运行模式已设置为：继续之前的任务", "info")
            
    def check_existing_data(self):
        """检查当前专业的已有数据并更新界面"""
        try:
            # 检查当前选择的专业的数据
            self.check_existing_data_for_major(self.current_major)
                
        except Exception as e:
            self.log_message(f"检查已有数据时出错: {e}", "error")
    
    def update_initial_progress_display(self):
        """更新初始的进度显示（基于Excel文件，不需要登录）"""
        try:
            # 根据已检测到的数据更新界面
            start_page = self.detected_start_page
            records_count = self.detected_records_count
            
            # 设置起始页显示
            self.start_page_var.set(str(start_page))
            
            # 设置页面信息（总页数待登录后确定）
            self.page_info.config(text=f"第{start_page}页 / 总页数待确定")
            
            # 设置记录数
            self.records_info.config(text=f"{records_count} 条")
            
            # 进度需要总页数，等登录后再显示
            self.progress_var.set(0)
            self.progress_text.config(text="待登录确定")
            
            # 设置状态
            if records_count > 0:
                self.status_info.config(text="有历史数据", foreground="orange")
                self.log_message(f"✓ 检测到历史数据：第{start_page}页开始，已有{records_count}条记录", "info")
            else:
                self.status_info.config(text="新任务", foreground="green")
                self.log_message("✓ 新任务：将从第1页开始", "info")
                
        except Exception as e:
            # 设置默认值
            self.progress_var.set(0)
            self.progress_text.config(text="待登录确定")
            self.page_info.config(text="第1页 / 总页数待确定")
            self.records_info.config(text="0 条")
            self.status_info.config(text="就绪", foreground="green")
            self.log_message(f"更新初始进度显示失败: {e}", "error")
            
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
            
            # 获取当前选择的专业代码
            selected = self.major_var.get()
            if " - " in selected:
                major_code = selected.split(" - ")[0]
            else:
                major_code = "125300"  # 默认专业
            
            # 获取设置参数
            mode = self.mode_var.get()
            
            # 获取页面范围设置
            start_page_str = self.start_page_var.get()
            if start_page_str in ["待确定", "待检测...", ""]:
                # 如果还没有确定起始页，使用检测到的值
                start_page = self.detected_start_page
            else:
                # 使用界面上的设置（可能是登录后更新的准确值）
                try:
                    start_page = int(start_page_str)
                except:
                    start_page = self.detected_start_page
            
            end_page = int(self.end_page_var.get()) if self.end_page_var.get() else None
            test_limit = int(self.test_limit_var.get()) if mode == "test" else None
            headless_mode = self.headless_var.get()  # 获取无头模式选项
            
            if start_page < 1:
                messagebox.showerror("错误", "起始页面必须大于0！")
                return
            
            if end_page and end_page < start_page:
                messagebox.showerror("错误", "结束页面不能小于起始页面！")
                return
            
            # 获取学习方式
            study_mode = self.study_mode_var.get()
            study_mode_name = "全日制" if study_mode == "1" else "非全日制"
            
            # 创建爬虫实例，传递正确的headless参数
            self.scraper = ScraperGUIWrapper(
                self,
                progress_callback=self.progress_callback,
                status_callback=self.status_callback,
                headless=headless_mode,  # 使用界面选择的无头模式设置
                major_code=major_code,
                study_mode=study_mode
            )
                
            # 记录运行模式
            mode_text = "无头模式（后台运行）" if headless_mode else "可视模式（显示浏览器）"
            major_name = MAJOR_CONFIG[major_code]["name"]
            self.log_message(f"启动爬虫 - {mode_text} - 专业：{major_name} - 学习方式：{study_mode_name}")
            self.log_message(f"已创建爬虫实例 - 专业: {MAJOR_CONFIG[major_code]['name']}, 学习方式: {study_mode_name}, 模式: {mode_text}", "info")
                
            # 根据模式调整参数
            if mode == "restart":
                self.log_message("重新开始任务，清空已有数据...")
                self.scraper.data = []
                self.scraper.current_page = 1
                
                # 删除该专业和学习方式的Excel文件（重新开始时）
                major_name = MAJOR_CONFIG[major_code]['name']
                study_mode_name = "全日制" if study_mode == "1" else "非全日制"
                excel_file = f"研究生招生信息_{major_name}_{study_mode_name}.xlsx"
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
                # 在继续模式下，start_page已经在前面根据界面设置或检测值确定了
                # 这里更新爬虫的当前页面设置
                if hasattr(self.scraper, 'current_page'):
                    self.scraper.current_page = start_page
                self.log_message(f"将从第{start_page}页继续爬取", "info")
                
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
            # 传递确定的起始页参数给爬虫
            self.log_message(f"准备启动爬虫：起始页={start_page}, 结束页={end_page or '自动检测'}", "info")
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
            # 获取当前选择的专业
            selected = self.major_var.get()
            if " - " in selected:
                major_code = selected.split(" - ")[0]
                major_name = MAJOR_CONFIG[major_code]["name"]
            else:
                major_code = "125300"
                major_name = "会计专硕"
            
            # 获取学习方式
            study_mode = self.study_mode_var.get()
            study_mode_name = "全日制" if study_mode == "1" else "非全日制"
            excel_file = f"研究生招生信息_{major_name}_{study_mode_name}.xlsx"
            
            if os.path.exists(excel_file):
                # 读取Excel文件分析进度
                import pandas as pd
                df = pd.read_excel(excel_file)
                data = df.to_dict('records')
                
                # 创建进度查看窗口
                progress_window = tk.Toplevel(self.root)
                progress_window.title("详细进度信息")
                progress_window.geometry("600x500")
                
                # 进度信息文本
                progress_text = scrolledtext.ScrolledText(progress_window, height=25, width=70)
                progress_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                # 分析数据
                total_records = len(data)
                completed_pages = set(record.get('页码', 0) for record in data)
                max_completed_page = max(completed_pages) if completed_pages else 0
                
                # 统计各页面数据
                data_count_by_page = {}
                for record in data:
                    page = record.get('页码', 0)
                    data_count_by_page[page] = data_count_by_page.get(page, 0) + 1
                
                # 显示进度信息
                progress_text.insert(tk.END, f"专业信息: {major_name} ({major_code})\n")
                progress_text.insert(tk.END, f"数据文件: {excel_file}\n")
                progress_text.insert(tk.END, f"文件修改时间: {datetime.fromtimestamp(os.path.getmtime(excel_file)).strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                progress_text.insert(tk.END, f"总记录数: {total_records}\n")
                progress_text.insert(tk.END, f"已完成页面数: {len(completed_pages)}\n")
                progress_text.insert(tk.END, f"最大页码: {max_completed_page}\n")
                
                if max_completed_page > 0:
                    estimated_progress = (max_completed_page / 33) * 100
                    progress_text.insert(tk.END, f"估算进度: {estimated_progress:.1f}%\n\n")
                else:
                    progress_text.insert(tk.END, f"估算进度: 0%\n\n")
                
                progress_text.insert(tk.END, "已完成页面:\n")
                if completed_pages:
                    sorted_pages = sorted(completed_pages)
                    progress_text.insert(tk.END, f"{', '.join(map(str, sorted_pages))}\n\n")
                else:
                    progress_text.insert(tk.END, "无\n\n")
                    
                progress_text.insert(tk.END, "各页面数据统计:\n")
                for page in sorted(data_count_by_page.keys()):
                    count = data_count_by_page[page]
                    progress_text.insert(tk.END, f"第{page}页: {count}条记录\n")
                
                # 显示最新几条记录
                if data:
                    progress_text.insert(tk.END, "\n最新5条记录:\n")
                    for i, record in enumerate(data[-5:], 1):
                        school = record.get('学校名称', '未知')
                        page = record.get('页码', '未知')
                        progress_text.insert(tk.END, f"{i}. {school} (第{page}页)\n")
                    
                progress_text.config(state=tk.DISABLED)
                
            else:
                messagebox.showinfo("信息", f"未找到该专业的数据文件: {excel_file}")
                
        except Exception as e:
            messagebox.showerror("错误", f"查看进度时出错: {e}")
            self.log_message(f"查看进度失败: {e}", "error")
            
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