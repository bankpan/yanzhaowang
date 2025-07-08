#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
研究生招生信息爬虫 - 修复版本
基于调试程序的成功经验重写
"""

import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import logging
from datetime import datetime
import os
import json
import threading
import queue

# 专业配置映射
MAJOR_CONFIG = {
    "125300": {
        "name": "会计专硕",
        "category": "1253",
        "full_name": "会计"
    },
    "125700": {
        "name": "审计专硕", 
        "category": "1257",
        "full_name": "审计"
    },
    "125500": {
        "name": "图书情报",
        "category": "1255", 
        "full_name": "图书情报"
    },
    "125604": {
        "name": "物流工程与管理",
        "category": "1256",
        "full_name": "物流工程与管理"
    },
    "125603": {
        "name": "工业工程与管理",
        "category": "1256",
        "full_name": "工业工程与管理"
    }
}

class YanZhaoScraperFixed:
    def __init__(self, progress_callback=None, status_callback=None, headless=False, major_code="125300"):
        """初始化爬虫"""
        self.headless = headless  # 无头模式标志
        self.major_code = major_code  # 选择的专业代码
        self.major_info = MAJOR_CONFIG.get(major_code, MAJOR_CONFIG["125300"])  # 专业信息
        
        # 先设置回调函数，避免setup_driver调用时出错
        self.progress_callback = progress_callback  # 进度更新回调
        self.status_callback = status_callback      # 状态更新回调
        self.is_paused = False                      # 暂停标志
        self.is_stopped = False                     # 停止标志
        self.is_running = False                     # 运行标志
        self.status_queue = queue.Queue()           # 状态消息队列
        
        # 初始化组件
        self.setup_logging()
        self.setup_driver()
        
        # 基本属性
        self.data = []
        self.current_page = 1
        self.total_pages = 1  # 初始值，将动态检测
        self.target_url = None  # 动态获取的目标URL
        self.username = "18042003196"
        self.password = "421950abcABC"
        
        # 设置Excel文件名（固定文件名，不使用时间戳）
        self.excel_filename = f"研究生招生信息_{self.major_info['name']}.xlsx"
        
        # 尝试加载已有数据和进度
        self.load_existing_data()
        
    def get_major_options():
        """获取所有可用的专业选项"""
        return {code: info["name"] for code, info in MAJOR_CONFIG.items()}
    
    def set_major(self, major_code):
        """设置专业代码"""
        if major_code in MAJOR_CONFIG:
            self.major_code = major_code
            self.major_info = MAJOR_CONFIG[major_code]
            
            # 更新文件名（固定文件名，不使用时间戳）
            self.excel_filename = f"研究生招生信息_{self.major_info['name']}.xlsx"
            
            # 重新加载该专业的进度
            self.load_existing_data()
            
            self.update_status(f"已切换到专业：{self.major_info['name']} ({major_code})", "info")
            return True
        else:
            self.update_status(f"无效的专业代码：{major_code}", "error")
            return False
    
    def load_existing_data(self):
        """加载已有数据和进度，实现断点续传（按专业区分）"""
        try:
            # 使用固定的Excel文件名
            excel_file = self.excel_filename
            progress_file = f'progress_{self.major_code}.json'
            
            # 检查Excel文件是否存在
            if os.path.exists(excel_file):
                # 加载已有数据
                try:
                    import pandas as pd
                    df = pd.read_excel(excel_file)
                    self.data = df.to_dict('records')
                    print(f"发现已有数据文件 {excel_file}，加载了 {len(self.data)} 条记录")
                    
                    # 分析已完成的页面
                    if self.data:
                        completed_pages = set(record.get('页码', 0) for record in self.data)
                        max_completed_page = max(completed_pages) if completed_pages else 0
                        
                        # 检查最后一页是否完整
                        last_page_records = [r for r in self.data if r.get('页码') == max_completed_page]
                        
                        # 如果最后一页记录数少于预期，从该页重新开始
                        if len(last_page_records) < 10:  # 假设每页至少有10个院校
                            self.current_page = max_completed_page
                            # 移除不完整页面的数据
                            self.data = [r for r in self.data if r.get('页码') != max_completed_page]
                            print(f"检测到第{max_completed_page}页数据不完整，将从第{max_completed_page}页重新开始")
                        else:
                            self.current_page = max_completed_page + 1
                            print(f"将从第{self.current_page}页继续爬取")
                            
                    else:
                        self.current_page = 1
                except Exception as e:
                    print(f"读取Excel文件失败: {e}")
                    self.data = []
                    self.current_page = 1
            else:
                print(f"未找到数据文件 {excel_file}，将从头开始")
                self.data = []
                self.current_page = 1
                        
            # 加载进度信息
            if os.path.exists(progress_file):
                with open(progress_file, 'r', encoding='utf-8') as f:
                    progress = json.load(f)
                    saved_page = progress.get('current_page', 1)
                    saved_total = progress.get('total_pages', 1)
                    print(f"进度文件显示上次处理到第{saved_page}页，共{saved_total}页")
                    
                    # 如果有保存的总页数，使用它
                    if saved_total > 1:
                        self.total_pages = saved_total
                        
        except Exception as e:
            print(f"加载已有数据失败: {e}")
            self.data = []
            self.current_page = 1
    
    def save_progress(self, status="running"):
        """保存进度信息（按专业区分）"""
        try:
            progress_file = f'progress_{self.major_code}.json'
            progress_data = {
                'major_code': self.major_code,
                'major_name': self.major_info['name'],
                'current_page': self.current_page,
                'total_pages': self.total_pages,
                'records_count': len(self.data),
                'last_update': datetime.now().isoformat(),
                'status': status,
                'target_url': self.target_url
            }
            
            with open(progress_file, 'w', encoding='utf-8') as f:
                json.dump(progress_data, f, ensure_ascii=False, indent=2)
                
            self.logger.info(f"进度已保存到 {progress_file}")
        except Exception as e:
            self.logger.error(f"保存进度失败: {e}")
    
    def get_target_url_by_major(self):
        """根据专业代码动态获取目标URL - 优化版本"""
        try:
            self.update_status("正在获取专业对应的目标页面...", "info")
            
            # 访问专业库首页
            base_url = "https://yz.chsi.com.cn/zsml/"
            self.update_status(f"访问页面: {base_url}", "info")
            self.driver.get(base_url)
            
            # 等待页面基础加载
            time.sleep(3)
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # 选择专业学位 - 直接使用最有效的方法
            self.update_status("选择专业学位...", "info")
            try:
                # 等待元素出现并直接点击
                WebDriverWait(self.driver, 8).until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '专业学位')]"))
                )
                
                # 找到可点击的专业学位元素
                professional_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), '专业学位')]")
                for element in professional_elements:
                    if element.is_displayed() and element.is_enabled():
                        element.click()
                        break
                time.sleep(2)
                
            except Exception as e:
                self.update_status(f"选择专业学位失败: {e}", "error")
                return None
            
            # 选择全日制 - 直接查找并点击
            self.update_status("选择全日制...", "info")
            try:
                fulltime_element = self.driver.find_element(By.XPATH, "//*[text()='全日制']")
                self.driver.execute_script("arguments[0].click();", fulltime_element)
                time.sleep(2)
            except:
                # 全日制可选，如果找不到就跳过
                pass
            
            # 选择专业类别 - 直接查找并点击
            category = self.major_info["category"]
            self.update_status(f"选择专业类别：{category}...", "info")
            try:
                time.sleep(2)  # 等待类别列表加载
                category_element = self.driver.find_element(By.XPATH, f"//*[contains(text(), '({category})')]")
                self.driver.execute_script("arguments[0].click();", category_element)
                time.sleep(3)  # 等待专业列表加载
                
            except Exception as e:
                self.update_status(f"选择专业类别失败: {e}", "error")
                return None
            
            # 等待专业列表完全加载
            self.update_status("等待专业列表加载...", "info")
            time.sleep(5)
            
            # 直接查找开设院校链接 - 使用最有效的策略
            self.update_status("查找开设院校链接...", "info")
            try:
                # 基于MCP演示，我们知道这个策略是有效的
                open_schools_link = None
                
                # 策略1：直接查找包含zydetail的链接（最快）
                all_links = self.driver.find_elements(By.TAG_NAME, "a")
                for link in all_links:
                    try:
                        href = link.get_attribute('href')
                        if href and 'zydetail.do' in href and self.major_code in href:
                            open_schools_link = href
                            break
                    except:
                        continue
                
                # 策略2：如果策略1失败，查找文本为"开设院校"的链接
                if not open_schools_link:
                    try:
                        school_elements = self.driver.find_elements(By.XPATH, "//*[contains(., '开设院校')][@href]")
                        for element in school_elements:
                            href = element.get_attribute('href')
                            if href and 'zydetail.do' in href:
                                open_schools_link = href
                                break
                    except:
                        pass
                
                if open_schools_link:
                    self.update_status(f"成功获取目标URL", "success")
                    return open_schools_link
                else:
                    self.update_status("未找到开设院校链接，使用备用URL", "warning")
                    return None
                    
            except Exception as e:
                self.update_status(f"查找开设院校链接失败: {e}", "error")
                return None
                
        except Exception as e:
            self.update_status(f"获取目标URL失败: {e}", "error")
            
            # 备用方案：直接使用已知的URL
            self.update_status("使用备用URL方案...", "info")
            if self.major_code == "125300":
                backup_url = "https://yz.chsi.com.cn/zsml/zydetail.do?zydm=125300&zymc=%E4%BC%9A%E8%AE%A1&xwlx=zyxw&mldm=12&mlmc=%E7%AE%A1%E7%90%86%E5%AD%A6&yjxkdm=1253&yjxkmc=%E4%BC%9A%E8%AE%A1&xxfs=1&tydxs=&jsggjh=&sign=73f11afdfd7ae989f9112d36b83036c9"
                self.update_status(f"使用会计专硕备用URL", "info")
                return backup_url
            elif self.major_code == "125700":
                backup_url = "https://yz.chsi.com.cn/zsml/zydetail.do?zydm=125700&zymc=%E5%AE%A1%E8%AE%A1&xwlx=zyxw&mldm=12&mlmc=%E7%AE%A1%E7%90%86%E5%AD%A6&yjxkdm=1257&yjxkmc=%E5%AE%A1%E8%AE%A1&xxfs=1&tydxs=&jsggjh="
                self.update_status(f"使用审计专硕备用URL", "info")
                return backup_url
            else:
                self.update_status(f"未知专业代码，无法提供备用URL", "error")
                return None
    
    def detect_total_pages(self):
        """自动检测总页数"""
        try:
            self.update_status("正在检测总页数...", "info")
            
            # 查找分页信息的多种可能元素
            page_selectors = [
                "//li[contains(@class, 'last')]/a",  # 最后一页链接
                "//a[contains(text(), '末页')]",      # 末页文字
                "//span[contains(@class, 'page_index')]//strong[last()]",  # 页码范围
                "//div[contains(@class, 'pagination')]//a[last()-1]",  # 分页容器中倒数第二个链接
            ]
            
            total_pages = 1  # 默认值
            
            for selector in page_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    if elements:
                        # 尝试从元素文本或属性中提取页数
                        for element in elements:
                            text = element.text.strip()
                            if text.isdigit():
                                detected_pages = int(text)
                                if detected_pages > total_pages:
                                    total_pages = detected_pages
                            
                            # 尝试从href属性中提取页数
                            href = element.get_attribute('href')
                            if href:
                                import re
                                page_match = re.search(r'page[=_](\d+)', href)
                                if page_match:
                                    detected_pages = int(page_match.group(1))
                                    if detected_pages > total_pages:
                                        total_pages = detected_pages
                        break
                except Exception as e:
                    continue
            
            # 如果没有检测到分页信息，尝试查找页面上所有的数字链接
            if total_pages == 1:
                try:
                    page_links = self.driver.find_elements(By.XPATH, "//a[string-length(text()) <= 3 and number(text()) = number(text())]")
                    for link in page_links:
                        try:
                            page_num = int(link.text.strip())
                            if page_num > total_pages:
                                total_pages = page_num
                        except:
                            continue
                except Exception as e:
                    pass
            
            self.total_pages = max(total_pages, 1)
            self.update_status(f"检测到总页数：{self.total_pages}", "info")
            return self.total_pages
            
        except Exception as e:
            self.update_status(f"检测总页数失败，使用默认值1: {e}", "warning")
            self.total_pages = 1
            return 1
    
    def setup_logging(self):
        """设置日志"""
        log_filename = f'yanzhao_{self.major_code}.log'
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def update_status(self, message, level="info"):
        """更新状态信息"""
        try:
            # 记录日志
            if level == "error":
                self.logger.error(message)
            elif level == "warning":
                self.logger.warning(message)
            else:
                self.logger.info(message)
            
            # 如果有状态回调函数，调用它
            if self.status_callback:
                self.status_callback(message, level)
                
        except Exception as e:
            print(f"更新状态失败: {e}")
    
    def update_progress(self, current_page, total_pages, records_count, status="运行中"):
        """更新进度信息"""
        try:
            # 计算进度百分比
            progress_percentage = (current_page - 1) / total_pages * 100 if total_pages > 0 else 0
            
            progress_data = {
                'current_page': current_page,
                'total_pages': total_pages,
                'records_count': records_count,
                'progress_percentage': progress_percentage,
                'status': status
            }
            
            # 如果有进度回调函数，调用它
            if self.progress_callback:
                self.progress_callback(progress_data)
                
        except Exception as e:
            print(f"更新进度失败: {e}")
    
    def check_stop_signal(self):
        """检查停止信号"""
        return self.is_stopped
    
    def wait_if_paused(self):
        """如果暂停则等待"""
        while self.is_paused and not self.is_stopped:
            time.sleep(0.1)
    
    def pause(self):
        """暂停爬虫"""
        self.is_paused = True
        self.update_status("爬虫已暂停", "warning")
    
    def resume(self):
        """恢复爬虫"""
        self.is_paused = False
        self.update_status("爬虫已恢复", "info")
    
    def stop(self):
        """停止爬虫"""
        self.is_stopped = True
        self.update_status("正在停止爬虫...", "warning")
        
    def setup_driver(self):
        """设置Chrome驱动"""
        chrome_options = Options()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        # 根据headless参数决定是否启用无头模式
        if self.headless:
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--disable-gpu')  # 无头模式下禁用GPU
            chrome_options.add_argument('--no-first-run')  # 跳过首次运行设置
            if self.status_callback:
                self.update_status("启用无头模式，浏览器将在后台运行", "info")
        else:
            if self.status_callback:
                self.update_status("启用可视模式，将显示浏览器窗口", "info")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.wait = WebDriverWait(self.driver, 10)
            
            mode_text = "无头模式" if self.headless else "可视模式"
            self.logger.info(f"Chrome驱动初始化成功 ({mode_text})")
            if self.status_callback:
                self.update_status(f"Chrome驱动初始化成功 ({mode_text})", "success")
        except Exception as e:
            self.logger.error(f"Chrome驱动初始化失败: {e}")
            if self.status_callback:
                self.update_status(f"Chrome驱动初始化失败: {e}", "error")
            raise
            
    def login_and_navigate(self):
        """登录并导航到目标页面 - 改进版本"""
        try:
            # 动态获取目标页面
            if not self.target_url:
                self.update_status("正在获取目标URL...", "info")
                self.target_url = self.get_target_url_by_major()
                if not self.target_url:
                    raise Exception("无法获取目标URL，请检查网络连接或页面结构")
            
            self.update_status(f"访问目标页面: {self.target_url}", "info")
            self.driver.get(self.target_url)
            
            # 等待页面加载
            self.update_status("等待页面加载...", "info")
            time.sleep(5)
            
            # 检测总页数
            try:
                self.detect_total_pages()
            except Exception as e:
                self.update_status(f"检测总页数失败: {e}", "warning")
                self.total_pages = 1  # 默认设为1页
            
            # 检查页面是否加载成功
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
            except:
                raise Exception("页面加载超时")
            
            # 检查是否需要登录
            page_source = self.driver.page_source.lower()
            
            if "登录后" in page_source or "请登录" in page_source:
                self.update_status("检测到需要登录，开始登录流程...", "info")
                
                try:
                    # 查找登录按钮
                    login_selectors = [
                        "//a[contains(text(), '登录')]",
                        "//button[contains(text(), '登录')]",
                        "//*[contains(text(), '登录')][@href or @onclick]",
                        "//a[@href and contains(@href, 'login')]"
                    ]
                    
                    login_button = None
                    for selector in login_selectors:
                        try:
                            login_button = WebDriverWait(self.driver, 3).until(
                                EC.element_to_be_clickable((By.XPATH, selector))
                            )
                            break
                        except:
                            continue
                    
                    if not login_button:
                        self.update_status("未找到登录按钮，尝试继续...", "warning")
                        return True
                    
                    # 点击登录按钮
                    self.driver.execute_script("arguments[0].click();", login_button)
                    time.sleep(3)
                    
                    # 等待登录页面加载
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@type='text' or @type='email'] | //input[contains(@placeholder, '账号') or contains(@placeholder, '用户名')]"))
                    )
                    
                    # 输入用户名
                    username_selectors = [
                        "//input[@placeholder='账号' or @name='username' or contains(@class, 'username')]",
                        "//input[@type='text'][1]",
                        "//input[contains(@placeholder, '账号')]",
                        "//input[contains(@placeholder, '用户名')]",
                        "//input[@name='account']"
                    ]
                    
                    username_input = None
                    for selector in username_selectors:
                        try:
                            username_input = self.driver.find_element(By.XPATH, selector)
                            break
                        except:
                            continue
                    
                    if username_input:
                        username_input.clear()
                        username_input.send_keys(self.username)
                        self.update_status("输入用户名完成", "info")
                    else:
                        raise Exception("未找到用户名输入框")
                    
                    # 输入密码
                    password_selectors = [
                        "//input[@type='password']",
                        "//input[@placeholder='密码' or @name='password']",
                        "//input[contains(@placeholder, '密码')]"
                    ]
                    
                    password_input = None
                    for selector in password_selectors:
                        try:
                            password_input = self.driver.find_element(By.XPATH, selector)
                            break
                        except:
                            continue
                    
                    if password_input:
                        password_input.clear()
                        password_input.send_keys(self.password)
                        self.update_status("输入密码完成", "info")
                    else:
                        raise Exception("未找到密码输入框")
                    
                    # 点击登录按钮
                    login_submit_selectors = [
                        "//button[contains(text(), '登录')]",
                        "//input[@type='submit' and @value='登录']",
                        "//input[@type='submit']",
                        "//button[@type='submit']",
                        "//*[contains(text(), '登录') and (@type='submit' or @onclick)]"
                    ]
                    
                    login_submit = None
                    for selector in login_submit_selectors:
                        try:
                            login_submit = self.driver.find_element(By.XPATH, selector)
                            break
                        except:
                            continue
                    
                    if login_submit:
                        self.driver.execute_script("arguments[0].click();", login_submit)
                        self.update_status("点击登录按钮", "info")
                    else:
                        raise Exception("未找到登录提交按钮")
                    
                    # 等待登录完成
                    time.sleep(5)
                    
                    # 检查登录是否成功
                    if "登录失败" in self.driver.page_source or "用户名或密码错误" in self.driver.page_source:
                        raise Exception("登录失败：用户名或密码错误")
                    
                    self.update_status("登录流程完成", "success")
                    
                except Exception as login_error:
                    self.update_status(f"登录失败: {login_error}", "error")
                    # 不直接返回False，先尝试看看是否能继续
                    self.update_status("尝试无需登录继续执行...", "warning")
                    
            else:
                self.update_status("无需登录，直接访问", "info")
            
            # 等待页面完全加载
            time.sleep(3)
            
            # 验证页面是否包含期望的内容
            if "个相关招生单位" in self.driver.page_source or "开设专业" in self.driver.page_source:
                self.update_status("页面加载成功，找到招生单位信息", "success")
                return True
            else:
                self.update_status("页面内容异常，但尝试继续执行", "warning")
                return True
                
        except Exception as e:
            self.update_status(f"导航过程出错: {e}", "error")
            return False
            
    def navigate_to_page(self, page_num):
        """导航到指定页面"""
        try:
            if page_num == 1:
                self.logger.info("已在第1页")
                return True
            
            self.logger.info(f"导航到第{page_num}页")
            
            # 查找页码链接
            page_links = self.driver.find_elements(By.XPATH, f"//li/a[text()='{page_num}']")
            if page_links:
                page_links[0].click()
                time.sleep(3)
                self.current_page = page_num
                self.logger.info(f"成功导航到第{page_num}页")
                return True
            else:
                # 如果找不到直接的页码链接，尝试使用下一页按钮
                for _ in range(page_num - self.current_page):
                    next_buttons = self.driver.find_elements(By.XPATH, "//li[contains(@class, 'next')]/a | //a[contains(text(), '下一页')]")
                    if next_buttons:
                        next_buttons[0].click()
                        time.sleep(2)
                        self.current_page += 1
                    else:
                        break
                        
                self.logger.info(f"通过下一页按钮导航到第{self.current_page}页")
                return True
            
        except Exception as e:
            self.logger.error(f"导航到第{page_num}页失败: {e}")
            return False
            
    def get_universities_on_page(self):
        """获取当前页面的所有院校 - 基于调试程序的成功经验"""
        try:
            # 等待页面加载完成
            time.sleep(2)
            
            # 直接查找展开按钮 - 这是调试程序证明有效的方法
            expand_buttons = self.driver.find_elements(By.XPATH, "//a[contains(text(), '展开')]")
            
            if not expand_buttons:
                self.logger.error("未找到展开按钮")
                return []
                
            self.logger.info(f"找到{len(expand_buttons)}个展开按钮")
            
            # 获取院校名称 - 基于调试程序的成功经验
            university_names = []
            name_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), '大学') or contains(text(), '学院')]")
            
            for elem in name_elements:
                try:
                    text = elem.text.strip()
                    if text and ('大学' in text or '学院' in text) and '(' in text and text.startswith('('):
                        university_names.append(text)
                except:
                    continue
                    
            self.logger.info(f"找到{len(university_names)}个院校名称")
            
            # 构建院校列表
            universities = []
            for i, (button, name) in enumerate(zip(expand_buttons, university_names)):
                try:
                    # 为每个展开按钮找到对应的院校容器
                    # 使用简单的方法：展开按钮的父级元素就是院校容器
                    parent = button.find_element(By.XPATH, "./ancestor::*[.//img][1]")
                    
                    universities.append({
                        'name': name,
                        'element': parent,
                        'expand_button': button,
                        'index': i + 1
                    })
                    
                    self.logger.info(f"找到院校: {name}")
                    
                except Exception as e:
                    self.logger.warning(f"处理第{i+1}个院校时出错: {e}")
                    continue
                    
            self.logger.info(f"在第{self.current_page}页成功构建{len(universities)}个院校")
            return universities
            
        except Exception as e:
            self.logger.error(f"获取院校列表失败: {e}")
            return []
            
    def process_university(self, university):
        """处理单个院校的所有硕士点"""
        try:
            self.logger.info(f"开始处理院校: {university['name']}")
            
            # 点击展开按钮
            university['expand_button'].click()
            time.sleep(3)
            
            # 查找详情链接
            detail_links = self.driver.find_elements(By.XPATH, "//a[contains(text(), '详情')]")
            
            if not detail_links:
                self.logger.warning(f"院校 {university['name']} 没有找到详情链接")
                return []
                
            self.logger.info(f"院校 {university['name']} 找到{len(detail_links)}个详情链接")
            
            university_data = []
            
            # 处理每个详情链接
            for i, detail_link in enumerate(detail_links):
                try:
                    self.logger.info(f"处理 {university['name']} 的第{i+1}个硕士点")
                    
                    # 点击详情链接
                    detail_link.click()
                    time.sleep(3)
                    
                    # 切换到新窗口
                    original_window = self.driver.current_window_handle
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    
                    # 提取详情信息 - 基于调试程序的成功经验
                    details = self.extract_program_details(university['name'], i + 1)
                    
                    if details:
                        university_data.append(details)
                        
                    # 关闭详情窗口，返回主窗口
                    self.driver.close()
                    self.driver.switch_to.window(original_window)
                    time.sleep(1)
                    
                except Exception as e:
                    self.logger.error(f"处理 {university['name']} 第{i+1}个硕士点失败: {e}")
                    # 确保返回主窗口
                    try:
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                    except:
                        pass
                    continue
                    
            # 收起院校 - 查找收起按钮
            try:
                collapse_buttons = self.driver.find_elements(By.XPATH, "//a[contains(text(), '收起')]")
                if collapse_buttons:
                    collapse_buttons[0].click()
                    time.sleep(1)
                    self.logger.info(f"收起院校: {university['name']}")
            except Exception as e:
                self.logger.warning(f"收起院校 {university['name']} 失败: {e}")
                
            return university_data
            
        except Exception as e:
            self.logger.error(f"处理院校 {university['name']} 失败: {e}")
            return []
            
    def extract_program_details(self, university_name, program_index):
        """提取硕士点详情信息 - 基于调试程序的成功经验"""
        try:
            # 等待页面加载
            time.sleep(2)
            
            # 定义字段选择器 - 基于手动页面分析的正确结构
            field_selectors = {
                '招生单位': [
                    "//div[contains(text(), '招生单位：')]/following-sibling::div",
                    "//div[contains(text(), '招生单位')]/following-sibling::div[1]",
                    "//*[contains(text(), '招生单位：')]/following-sibling::*[1]"
                ],
                '考试方式': [
                    "//div[contains(text(), '考试方式：')]/following-sibling::div",
                    "//div[contains(text(), '考试方式')]/following-sibling::div[1]",
                    "//*[contains(text(), '考试方式：')]/following-sibling::*[1]"
                ],
                '院系所': [
                    "//div[contains(text(), '院系所：')]/following-sibling::div",
                    "//div[contains(text(), '院系所')]/following-sibling::div[1]",
                    "//*[contains(text(), '院系所：')]/following-sibling::*[1]"
                ],
                '学习方式': [
                    "//div[contains(text(), '学习方式：')]/following-sibling::div",
                    "//div[contains(text(), '学习方式')]/following-sibling::div[1]",
                    "//*[contains(text(), '学习方式：')]/following-sibling::*[1]"
                ],
                '研究方向': [
                    "//div[contains(text(), '研究方向：')]/following-sibling::div",
                    "//div[contains(text(), '研究方向')]/following-sibling::div[1]",
                    "//*[contains(text(), '研究方向：')]/following-sibling::*[1]"
                ],
                '拟招生人数': [
                    "//div[contains(text(), '拟招生人数：')]/following-sibling::div",
                    "//div[contains(text(), '拟招生人数')]/following-sibling::div[1]",
                    "//*[contains(text(), '拟招生人数：')]/following-sibling::*[1]"
                ]
            }
            
            details = {}
            
            # 提取每个字段
            for field_name, selectors in field_selectors.items():
                details[field_name] = ""
                for selector in selectors:
                    try:
                        element = self.driver.find_element(By.XPATH, selector)
                        value = element.text.strip()
                        if value:
                            details[field_name] = value
                            self.logger.info(f"找到 {field_name}: {value}")
                            break
                    except:
                        continue
                        
                if not details[field_name]:
                    self.logger.warning(f"未找到字段: {field_name}")
            
            # 添加额外信息
            details['页码'] = self.current_page
            details['院校名称'] = university_name
            details['硕士点序号'] = program_index
            details['爬取时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            return details
            
        except Exception as e:
            self.logger.error(f"提取详情失败: {e}")
            return None
            
    def save_data_to_excel(self, filename=None):
        """保存数据到Excel文件 - 带文件占用检测和重试机制"""
        if filename is None:
            filename = self.excel_filename
            
        if not self.data:
            self.logger.warning("没有数据可保存")
            return False
        
        max_retries = 5  # 最大重试次数
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                df = pd.DataFrame(self.data)
                df.to_excel(filename, index=False, engine='openpyxl')
                self.logger.info(f"数据已更新到 {filename}，共{len(self.data)}条记录")
                
                # 同时保存CSV备份
                csv_filename = filename.replace('.xlsx', '.csv')
                df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
                
                return True
                
            except PermissionError as e:
                retry_count += 1
                self.logger.warning(f"文件被占用，第{retry_count}次重试保存 {filename}")
                
                # 通过回调通知GUI显示文件占用提示
                if self.status_callback:
                    self.status_callback(f"文件 {filename} 被占用，请关闭Excel文件后点击确定（第{retry_count}/{max_retries}次尝试）", "warning")
                
                # 如果是在GUI环境中，需要等待用户操作
                if self.status_callback:
                    # 通过消息队列等待用户确认
                    self.wait_for_file_access_confirmation(filename, retry_count, max_retries)
                else:
                    # 命令行环境下直接提示并等待
                    input(f"文件 {filename} 被占用，请关闭Excel文件后按回车继续...")
                    
            except Exception as e:
                self.logger.error(f"保存数据失败: {e}")
                if self.status_callback:
                    self.status_callback(f"保存失败: {e}", "error")
                return False
        
        # 如果所有重试都失败了
        self.logger.error(f"文件 {filename} 保存失败，已重试{max_retries}次")
        if self.status_callback:
            self.status_callback(f"文件保存失败，已重试{max_retries}次，请检查文件权限", "error")
        return False
    
    def wait_for_file_access_confirmation(self, filename, retry_count, max_retries):
        """等待用户确认文件访问权限已解除"""
        # 这个方法将被GUI重写，命令行版本直接等待
        if not self.status_callback:
            time.sleep(2)  # 命令行版本简单等待2秒
        else:
            # GUI版本会通过回调处理用户交互
            pass
            
    def emergency_save(self, reason="unknown"):
        """紧急保存数据"""
        try:
            self.logger.info(f"执行紧急保存，原因: {reason}")
            
            # 紧急保存到主文件
            success = self.save_data_to_excel()
            
            # 保存进度
            self.save_progress(status=f"emergency_save_{reason}")
            
            # 保存运行日志摘要
            summary = {
                '保存时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                '保存原因': reason,
                '当前页码': self.current_page,
                '总记录数': len(self.data),
                '文件名': self.excel_filename
            }
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            with open(f'紧急保存摘要_{timestamp}.json', 'w', encoding='utf-8') as f:
                json.dump(summary, f, ensure_ascii=False, indent=2)
                
            self.logger.info(f"紧急保存完成: {self.excel_filename}")
            return success
            
        except Exception as e:
            self.logger.error(f"紧急保存失败: {e}")
            return False
            
    def run(self, start_page=None, end_page=None, max_universities_per_page=None):
        """运行爬虫"""
        try:
            self.is_running = True
            self.is_stopped = False
            self.update_status(f"开始运行爬虫 - 专业：{self.major_info['name']} ({self.major_code})")
            
            # 登录并导航（会自动检测总页数）
            if not self.login_and_navigate():
                self.update_status("登录失败", "error")
                return False
            
            # 使用断点续传的起始页面
            if start_page is None:
                start_page = self.current_page
                
            if end_page is None:
                end_page = self.total_pages
            elif end_page > self.total_pages:
                end_page = self.total_pages
                
            self.update_status(f"断点续传：从第{start_page}页开始，到第{end_page}页结束（共{self.total_pages}页）")
            self.update_status(f"当前已有数据：{len(self.data)}条记录")
            
            # 更新初始进度
            self.update_progress(start_page, self.total_pages, len(self.data), "初始化中")
                
            # 遍历页面
            for page_num in range(start_page, end_page + 1):
                # 检查停止信号
                if self.check_stop_signal():
                    self.update_status("用户停止了爬虫，正在保存数据...", "warning")
                    # 立即保存已获取的数据
                    if self.data:
                        self.save_data_to_excel()
                        self.save_progress("user_stopped")
                        self.update_status(f"用户停止，已保存{len(self.data)}条记录", "warning")
                    break
                    
                # 等待暂停状态
                self.wait_if_paused()
                
                self.update_status(f"开始处理第{page_num}页 (共{self.total_pages}页)")
                self.update_progress(page_num, self.total_pages, len(self.data), "正在爬取")
                
                try:
                    # 导航到页面
                    if not self.navigate_to_page(page_num):
                        self.update_status(f"导航到第{page_num}页失败", "error")
                        continue
                        
                    # 获取院校列表
                    universities = self.get_universities_on_page()
                    if not universities:
                        self.update_status(f"第{page_num}页没有找到院校", "warning")
                        continue
                        
                    # 限制每页处理的院校数量（用于测试）
                    if max_universities_per_page:
                        universities = universities[:max_universities_per_page]
                        
                    # 处理每个院校
                    for univ_index, university in enumerate(universities, 1):
                        # 检查停止信号
                        if self.check_stop_signal():
                            self.update_status("用户停止了爬虫，正在保存数据...", "warning")
                            # 立即保存已获取的数据
                            if self.data:
                                self.save_data_to_excel()
                                self.save_progress("user_stopped")
                                self.update_status(f"用户停止，已保存{len(self.data)}条记录", "warning")
                            break
                            
                        # 等待暂停状态
                        self.wait_if_paused()
                        
                        try:
                            self.update_status(f"处理院校 {univ_index}/{len(universities)}: {university['name']}")
                            
                            # 处理院校
                            university_data = self.process_university(university)
                            
                            # 添加到总数据
                            self.data.extend(university_data)
                            
                            self.update_status(f"院校 {university['name']} 完成，获得{len(university_data)}条记录")
                                
                            # 随机延时
                            time.sleep(random.uniform(2, 4))
                            
                        except Exception as e:
                            self.update_status(f"处理院校 {university['name']} 失败: {e}", "error")
                            continue
                    
                    # 如果用户停止了，退出页面循环        
                    if self.check_stop_signal():
                        # 已在院校循环中保存过，这里只需要退出
                        break
                        
                    # 每页完成后保存
                    self.save_data_to_excel()
                    self.save_progress(f"completed_page_{page_num}")
                    
                    self.update_status(f"第{page_num}页完成，当前总记录数: {len(self.data)}")
                    
                    # 页面间延时
                    time.sleep(random.uniform(3, 6))
                    
                except Exception as e:
                    self.update_status(f"处理第{page_num}页失败: {e}", "error")
                    continue
                    
            # 最终保存
            self.save_data_to_excel()
            
            if self.check_stop_signal():
                self.update_status(f"爬虫已停止，共获取{len(self.data)}条记录", "warning")
            else:
                self.update_status(f"爬虫运行完成，共获取{len(self.data)}条记录", "success")
            
            return True
            
        except Exception as e:
            self.logger.error(f"爬虫运行失败: {e}")
            # 即使运行失败也要保存已获取的数据
            self.emergency_save("运行异常")
            return False
            
        finally:
            # 确保在任何情况下都保存数据
            try:
                if self.data:
                    self.save_data_to_excel()  # 最终保存
                    self.save_progress("completed")
                    self.logger.info(f"程序结束，最终保存了{len(self.data)}条记录")
                else:
                    self.logger.info("程序结束，没有获取到数据")
            except Exception as e:
                self.logger.error(f"最终保存失败: {e}")
                
            # 关闭浏览器
            try:
                self.driver.quit()
                self.logger.info("浏览器已关闭")
            except:
                pass
                
    def __del__(self):
        """析构函数"""
        try:
            if hasattr(self, 'driver'):
                self.driver.quit()
        except:
            pass

    def test_url_access(self):
        """测试URL获取和页面访问功能"""
        try:
            self.update_status("开始测试URL获取和页面访问...", "info")
            
            # 测试URL获取
            self.update_status("测试第1步：获取目标URL", "info")
            target_url = self.get_target_url_by_major()
            
            if target_url:
                self.update_status(f"✓ URL获取成功: {target_url}", "success")
            else:
                self.update_status("✗ URL获取失败", "error")
                return False
            
            # 测试页面访问
            self.update_status("测试第2步：直接访问目标页面", "info")
            self.driver.get(target_url)
            time.sleep(5)
            
            # 检查页面内容
            if "个相关招生单位" in self.driver.page_source or "开设专业" in self.driver.page_source:
                self.update_status("✓ 页面访问成功，内容正常", "success")
                
                # 检查页面结构
                expand_buttons = self.driver.find_elements(By.XPATH, "//a[contains(text(), '展开')] | //*[contains(text(), '展开')]")
                if expand_buttons:
                    self.update_status(f"✓ 找到{len(expand_buttons)}个展开按钮", "success")
                else:
                    self.update_status("⚠ 未找到展开按钮，但页面基本正常", "warning")
                
                return True
            else:
                self.update_status("✗ 页面内容异常，可能需要登录", "warning")
                
                # 尝试检查是否需要登录
                if "登录" in self.driver.page_source.lower():
                    self.update_status("页面提示需要登录，这是正常的", "info")
                    return True
                else:
                    self.update_status("页面结构可能已变化", "error")
                    return False
                    
        except Exception as e:
            self.update_status(f"测试失败: {e}", "error")
            return False




def main():
    """主函数"""
    print("研究生招生信息爬虫 - 修复版（支持断点续传）")
    print("=" * 50)
    
    # 创建爬虫实例
    scraper = None
    scraper_full = None
    
    try:
        # 创建爬虫实例
        scraper = YanZhaoScraperFixed()
        
        # 检查是否是断点续传
        if scraper.current_page > 1 or len(scraper.data) > 0:
            print(f"检测到未完成的任务：")
            print(f"  - 已有数据：{len(scraper.data)}条记录")
            print(f"  - 下次将从第{scraper.current_page}页开始")
            print(f"  - 剩余页面：{33 - scraper.current_page + 1}页")
            
            user_choice = input("选择运行模式：\n1. 继续之前的任务（推荐）\n2. 重新开始（会覆盖已有数据）\n3. 仅测试运行\n请输入选择 (1/2/3): ")
            
            if user_choice == '1':
                # 继续之前的任务
                print("继续之前的任务...")
                scraper_full = scraper  # 直接使用已加载数据的实例
                try:
                    scraper_full.run()  # 使用默认的断点续传参数
                    print(f"任务完成！总共获取到{len(scraper_full.data)}条记录")
                    
                except KeyboardInterrupt:
                    print(f"\n用户中断，已获取{len(scraper_full.data)}条记录")
                    scraper_full.emergency_save("用户中断续传任务")
                    
                except Exception as e:
                    print(f"续传任务出错: {e}")
                    if scraper_full and scraper_full.data:
                        scraper_full.emergency_save("续传任务异常")
                return
                
            elif user_choice == '2':
                # 重新开始
                print("重新开始任务，将清空已有数据...")
                scraper.data = []
                scraper.current_page = 1
                
                # 清空进度文件
                progress_file = f'progress_{scraper.major_code}.json'
                if os.path.exists(progress_file):
                    os.remove(progress_file)
                    print("已清空进度记录")
                
                # 删除Excel文件（重新开始时）
                if os.path.exists(scraper.excel_filename):
                    try:
                        os.remove(scraper.excel_filename)
                        print(f"已删除原有数据文件：{scraper.excel_filename}")
                    except Exception as e:
                        print(f"删除数据文件失败：{e}，将在保存时重写")
                
                print("重新开始，将从第1页开始爬取")
                        
            elif user_choice == '3':
                # 测试运行
                print("执行测试运行...")
            else:
                print("无效选择，默认执行测试运行")
        
        # 运行爬虫 - 先测试第1页的前2个院校
        print("开始测试运行（第1页，前2个院校）...")
        success = scraper.run(start_page=1, end_page=1, max_universities_per_page=2)
        
        if success:
            print("测试运行成功！")
            print(f"测试获取到{len(scraper.data)}条记录")
            
            # 询问是否继续完整运行
            user_input = input("测试成功，是否继续完整运行所有页面？(y/n): ")
            if user_input.lower() == 'y':
                print("开始完整运行...")
                scraper_full = YanZhaoScraperFixed()
                
                try:
                    scraper_full.run(start_page=1, end_page=33)
                    print(f"完整运行完成！总共获取到{len(scraper_full.data)}条记录")
                    
                except KeyboardInterrupt:
                    print(f"\n用户中断完整运行，已获取{len(scraper_full.data)}条记录")
                    scraper_full.emergency_save("用户中断完整运行")
                    
                except Exception as e:
                    print(f"完整运行出错: {e}")
                    if scraper_full and scraper_full.data:
                        scraper_full.emergency_save("完整运行异常")
                        
            else:
                print("测试完成，程序结束")
        else:
            print("测试运行失败！")
            # 即使测试失败，也尝试保存可能获取的数据
            if scraper and scraper.data:
                scraper.emergency_save("测试失败")
                print(f"测试失败，但已保存{len(scraper.data)}条记录")
            
    except KeyboardInterrupt:
        print("\n用户中断程序")
        # 保存测试阶段的数据
        if scraper and scraper.data:
            scraper.emergency_save("用户中断测试")
            print(f"用户中断，已保存测试数据{len(scraper.data)}条记录")
        # 保存完整运行阶段的数据
        if scraper_full and scraper_full.data:
            scraper_full.emergency_save("用户中断完整运行")
            print(f"用户中断，已保存完整运行数据{len(scraper_full.data)}条记录")
        
    except Exception as e:
        print(f"程序运行出错: {e}")
        # 保存测试阶段的数据
        if scraper and scraper.data:
            scraper.emergency_save("程序异常")
            print(f"程序异常，已保存测试数据{len(scraper.data)}条记录")
        # 保存完整运行阶段的数据
        if scraper_full and scraper_full.data:
            scraper_full.emergency_save("程序异常")
            print(f"程序异常，已保存完整运行数据{len(scraper_full.data)}条记录")
        
    finally:
        print("\n" + "="*50)
        print("数据保存总结：")
        
        # 显示测试数据统计
        if scraper and scraper.data:
            print(f"测试阶段保存记录: {len(scraper.data)}条")
        else:
            print("测试阶段: 无数据")
            
        # 显示完整运行数据统计
        if scraper_full and scraper_full.data:
            print(f"完整运行保存记录: {len(scraper_full.data)}条")
        else:
            print("完整运行: 无数据")
            
        print("程序结束")


if __name__ == "__main__":
    main() 