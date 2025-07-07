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

class YanZhaoScraperFixed:
    def __init__(self, progress_callback=None, status_callback=None, headless=False):
        """初始化爬虫"""
        self.headless = headless  # 无头模式标志
        
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
        self.total_pages = 33
        self.username = "18042003196"
        self.password = "421950abcABC"
        
        # 尝试加载已有数据和进度
        self.load_existing_data()
        
    def set_callbacks(self, progress_callback=None, status_callback=None):
        """设置回调函数"""
        if progress_callback:
            self.progress_callback = progress_callback
        if status_callback:
            self.status_callback = status_callback
            
    def update_progress(self, current_page, total_pages, records_count, status="运行中"):
        """更新进度信息"""
        if self.progress_callback:
            progress_data = {
                'current_page': current_page,
                'total_pages': total_pages,
                'records_count': records_count,
                'progress_percentage': round((current_page / total_pages) * 100, 2),
                'status': status
            }
            self.progress_callback(progress_data)
            
    def update_status(self, message, level="info"):
        """更新状态信息"""
        if self.status_callback:
            self.status_callback(message, level)
        self.logger.info(message)
        
    def pause(self):
        """暂停爬虫"""
        self.is_paused = True
        self.update_status("爬虫已暂停，正在保存当前数据...", "warning")
        
        # 暂停时立即保存数据
        try:
            if self.data:
                self.save_data_to_excel()
                self.save_progress("paused")
                self.update_status(f"暂停保存完成，已保存{len(self.data)}条记录", "warning")
            else:
                self.update_status("爬虫已暂停（暂无数据）", "warning")
        except Exception as e:
            self.update_status(f"暂停保存失败: {e}", "error")
            self.logger.error(f"暂停保存失败: {e}")
            
    def resume(self):
        """继续爬虫"""
        self.is_paused = False
        self.update_status("爬虫已恢复运行", "info")
        
    def stop(self):
        """停止爬虫"""
        self.is_stopped = True
        self.is_paused = False
        self.update_status("正在停止爬虫...", "warning")
        
    def wait_if_paused(self):
        """检查暂停状态，如果暂停则等待"""
        while self.is_paused and not self.is_stopped:
            time.sleep(0.5)
            
    def check_stop_signal(self):
        """检查停止信号"""
        return self.is_stopped
        
    def get_status(self):
        """获取当前状态"""
        return {
            'is_running': self.is_running,
            'is_paused': self.is_paused,
            'is_stopped': self.is_stopped,
            'current_page': self.current_page,
            'total_pages': self.total_pages,
            'records_count': len(self.data),
            'progress_percentage': round((self.current_page / self.total_pages) * 100, 2)
        }
        
    def load_existing_data(self):
        """加载已有数据和进度，实现断点续传"""
        try:
            # 检查是否有已保存的Excel文件
            excel_file = '研究生招生信息.xlsx'
            progress_file = 'progress_fixed.json'
            
            if os.path.exists(excel_file):
                # 加载已有数据
                df = pd.read_excel(excel_file)
                self.data = df.to_dict('records')
                print(f"发现已有数据文件，加载了 {len(self.data)} 条记录")
                
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
                    
            # 加载进度信息
            if os.path.exists(progress_file):
                with open(progress_file, 'r', encoding='utf-8') as f:
                    progress = json.load(f)
                    saved_page = progress.get('current_page', 1)
                    print(f"进度文件显示上次处理到第{saved_page}页")
                    
        except Exception as e:
            print(f"加载已有数据失败: {e}")
            self.data = []
            self.current_page = 1
            
    def setup_logging(self):
        """设置日志"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('yanzhao_fixed.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
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
        """登录并导航到目标页面"""
        try:
            # 访问目标页面
            target_url = "https://yz.chsi.com.cn/zsml/zydetail.do?zydm=125300&zymc=%E4%BC%9A%E8%AE%A1&xwlx=zyxw&mldm=12&mlmc=%E7%AE%A1%E7%90%86%E5%AD%A6&yjxkdm=1253&yjxkmc=%E4%BC%9A%E8%AE%A1&xxfs=&tydxs=&jsggjh=&sign=73f11afdfd7ae989f9112d36b83036c9"
            self.driver.get(target_url)
            self.logger.info("访问目标页面")
            
            # 等待页面加载
            time.sleep(3)
            
            # 检查是否需要登录
            if "登录" in self.driver.page_source:
                self.logger.info("检测到需要登录，开始登录流程")
                
                # 点击登录按钮
                login_button = self.driver.find_element(By.XPATH, "//a[contains(text(), '登录')]")
                login_button.click()
                time.sleep(3)
                
                # 输入用户名
                username_input = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, "//input[@placeholder='账号' or @name='username' or contains(@class, 'username')]"))
                )
                username_input.clear()
                username_input.send_keys(self.username)
                self.logger.info("输入用户名完成")
                
                # 输入密码
                password_input = self.driver.find_element(By.XPATH, "//input[@type='password' or @placeholder='密码' or @name='password']")
                password_input.clear()
                password_input.send_keys(self.password)
                self.logger.info("输入密码完成")
                
                # 点击登录按钮
                login_submit = self.driver.find_element(By.XPATH, "//button[contains(text(), '登录')] | //input[@type='submit' and @value='登录']")
                login_submit.click()
                self.logger.info("点击登录按钮")
                
                # 等待登录完成
                time.sleep(5)
                
                self.logger.info("登录流程完成")
            else:
                self.logger.info("无需登录")
                
            return True
            
        except Exception as e:
            self.logger.error(f"登录过程出错: {e}")
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
            
    def save_data_to_excel(self, filename='研究生招生信息.xlsx'):
        """保存数据到Excel文件 - 简化版，只用一个文件"""
        try:
            if self.data:
                df = pd.DataFrame(self.data)
                df.to_excel(filename, index=False, engine='openpyxl')
                self.logger.info(f"数据已更新到 {filename}，共{len(self.data)}条记录")
                
                # 同时保存CSV备份
                csv_filename = filename.replace('.xlsx', '.csv')
                df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
                
                return True
            else:
                self.logger.warning("没有数据可保存")
                return False
                
        except Exception as e:
            self.logger.error(f"保存数据失败: {e}")
            return False
            
    def save_progress(self, status="running"):
        """保存进度信息"""
        try:
            progress_data = {
                'current_page': self.current_page,
                'total_records': len(self.data),
                'status': status,
                'last_update': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'data_count_by_page': {},
                'completed_pages': [],
                'total_pages': self.total_pages,
                'progress_percentage': 0
            }
            
            # 统计每页的数据量
            for record in self.data:
                page = record.get('页码', 0)
                if page not in progress_data['data_count_by_page']:
                    progress_data['data_count_by_page'][page] = 0
                progress_data['data_count_by_page'][page] += 1
            
            # 确定已完成的页面
            if self.data:
                completed_pages = list(set(record.get('页码', 0) for record in self.data))
                completed_pages.sort()
                progress_data['completed_pages'] = completed_pages
                
                # 计算进度百分比
                progress_data['progress_percentage'] = round((len(completed_pages) / self.total_pages) * 100, 2)
            
            with open('progress_fixed.json', 'w', encoding='utf-8') as f:
                json.dump(progress_data, f, ensure_ascii=False, indent=2)
                
            self.logger.info(f"进度已保存: 第{self.current_page}页，共{len(self.data)}条记录，进度{progress_data['progress_percentage']}%，状态: {status}")
            
        except Exception as e:
            self.logger.error(f"保存进度失败: {e}")
            
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
                '文件名': '研究生招生信息.xlsx'
            }
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            with open(f'紧急保存摘要_{timestamp}.json', 'w', encoding='utf-8') as f:
                json.dump(summary, f, ensure_ascii=False, indent=2)
                
            self.logger.info(f"紧急保存完成: 研究生招生信息.xlsx")
            return success
            
        except Exception as e:
            self.logger.error(f"紧急保存失败: {e}")
            return False
            
    def run(self, start_page=None, end_page=None, max_universities_per_page=None):
        """运行爬虫"""
        try:
            self.is_running = True
            self.is_stopped = False
            self.update_status("开始运行修复版爬虫")
            
            # 使用断点续传的起始页面
            if start_page is None:
                start_page = self.current_page
                
            if end_page is None:
                end_page = self.total_pages
                
            self.update_status(f"断点续传：从第{start_page}页开始，到第{end_page}页结束")
            self.update_status(f"当前已有数据：{len(self.data)}条记录")
            
            # 更新初始进度
            self.update_progress(start_page, end_page, len(self.data), "初始化中")
                
            # 登录并导航
            if not self.login_and_navigate():
                self.update_status("登录失败", "error")
                return False
                
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
                # 删除已有文件
                for file in ['研究生招生信息.xlsx', '研究生招生信息.csv', 'progress_fixed.json']:
                    if os.path.exists(file):
                        os.remove(file)
                        print(f"已删除文件: {file}")
                        
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