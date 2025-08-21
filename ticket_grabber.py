#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
抢票工具 - 精确时间点击器
支持Windows和macOS跨平台运行
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import time
import datetime
import requests
import pyautogui
import json
from typing import Optional

class TicketGrabber:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("抢票工具 v1.0")
        # 适配Windows高DPI与不同缩放，适当加宽默认宽度，并允许横向缩放
        self.root.geometry("520x500")
        self.root.resizable(True, False)
        
        # 变量
        self.target_time = None
        self.target_datetime = None  # 本次等待的绝对目标时间
        self.is_running = False
        self.mouse_position = None  # 启动后不再预设，改为到点取当前鼠标位置
        self.time_offset = 0  # 本地时间与网络时间的偏差
        
        # 初始化GUI
        self.init_gui()
        
        # 启动时间更新线程
        self.start_time_update()
        
    def init_gui(self):
        """初始化GUI界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 时间设置区域
        time_frame = ttk.LabelFrame(main_frame, text="时间设置", padding="10")
        time_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(time_frame, text="目标时间 (HH:MM:SS):").grid(row=0, column=0, padx=(0, 10))
        
        # 时间输入框
        self.time_var = tk.StringVar()
        self.time_entry = ttk.Entry(time_frame, textvariable=self.time_var, width=15)
        self.time_entry.grid(row=0, column=1, padx=(0, 10), sticky=(tk.W, tk.E))
        self.time_entry.insert(0, "14:00:00")
        
        # 启动按钮
        self.start_button = ttk.Button(time_frame, text="启动", command=self.toggle_grabber)
        self.start_button.grid(row=0, column=2, sticky=tk.E)
        
        # 状态显示区域
        status_frame = ttk.LabelFrame(main_frame, text="状态信息", padding="10")
        status_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 当前时间
        ttk.Label(status_frame, text="当前时间:").grid(row=0, column=0, sticky=tk.W)
        self.current_time_label = ttk.Label(status_frame, text="--:--:--.---", font=("Courier", 12))
        self.current_time_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        # 倒计时
        ttk.Label(status_frame, text="倒计时:").grid(row=1, column=0, sticky=tk.W)
        self.countdown_label = ttk.Label(status_frame, text="--:--:--.---", font=("Courier", 12))
        self.countdown_label.grid(row=1, column=1, sticky=tk.W, padx=(10, 0))
        
        # 状态
        ttk.Label(status_frame, text="状态:").grid(row=2, column=0, sticky=tk.W)
        self.status_label = ttk.Label(status_frame, text="等待中...", foreground="blue")
        self.status_label.grid(row=2, column=1, sticky=tk.W, padx=(10, 0))
        
        # 鼠标提示（不再需要预设，提示用户启动后悬停目标位置）
        ttk.Label(status_frame, text="鼠标提示:").grid(row=3, column=0, sticky=tk.W)
        self.mouse_label = ttk.Label(status_frame, text="启动后将点击当前鼠标所在位置")
        self.mouse_label.grid(row=3, column=1, sticky=tk.W, padx=(10, 0))
        
        # 网络时间同步
        sync_frame = ttk.Frame(main_frame)
        sync_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(sync_frame, text="同步网络时间", command=self.sync_network_time).grid(row=0, column=0)
        self.sync_status_label = ttk.Label(sync_frame, text="", foreground="green")
        self.sync_status_label.grid(row=0, column=1, padx=(10, 0))
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="点击日志", padding="10")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, width=45)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        # 让时间设置区域的中间列（输入框）可伸缩，避免高DPI下被挤压
        time_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        # Windows高DPI: 提升进程DPI感知，减少缩放导致的布局错位
        try:
            import platform
            if platform.system() == 'Windows':
                import ctypes
                try:
                    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware
                except Exception:
                    try:
                        ctypes.windll.user32.SetProcessDPIAware()
                    except Exception:
                        pass
        except Exception:
            pass
        
    def start_time_update(self):
        """启动时间更新线程"""
        def update_time():
            while True:
                try:
                    current_time = datetime.datetime.now() + datetime.timedelta(seconds=self.time_offset)
                    time_str = current_time.strftime("%H:%M:%S.%f")[:-3]
                    self.current_time_label.config(text=time_str)
                    
                    # 更新倒计时
                    if self.target_time and self.is_running:
                        target_dt = datetime.datetime.combine(current_time.date(), self.target_time)
                        if target_dt < current_time:
                            target_dt += datetime.timedelta(days=1)
                        
                        diff = target_dt - current_time
                        if diff.total_seconds() > 0:
                            hours, remainder = divmod(int(diff.total_seconds()), 3600)
                            minutes, seconds = divmod(remainder, 60)
                            microseconds = diff.microseconds // 1000
                            countdown_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}.{microseconds:03d}"
                            self.countdown_label.config(text=countdown_str)
                        else:
                            self.countdown_label.config(text="00:00:00.000")
                    else:
                        self.countdown_label.config(text="--:--:--.---")
                        
                except Exception as e:
                    self.log(f"时间更新错误: {e}")
                
                time.sleep(0.01)  # 10ms更新间隔
        
        thread = threading.Thread(target=update_time, daemon=True)
        thread.start()
        
    def sync_network_time(self):
        """同步网络时间"""
        def sync():
            try:
                # 尝试多个NTP服务器
                ntp_servers = [
                    "time.nist.gov",
                    "pool.ntp.org", 
                    "time.google.com",
                    "ntp.aliyun.com"
                ]
                
                for server in ntp_servers:
                    try:
                        import socket
                        import struct
                        
                        client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                        client.settimeout(3)
                        
                        # NTP请求
                        msg = '\x1b' + 47 * '\0'
                        client.sendto(msg.encode('utf-8'), (server, 123))
                        msg, address = client.recvfrom(1024)
                        client.close()
                        
                        # 解析NTP响应
                        unpacked = struct.unpack('!12I', msg)
                        ntp_time = unpacked[10] + float(unpacked[11]) / 2**32
                        ntp_time -= 2208988800  # NTP to Unix timestamp
                        
                        local_time = time.time()
                        self.time_offset = ntp_time - local_time
                        
                        self.sync_status_label.config(text=f"已同步 (偏差: {self.time_offset:.3f}s)")
                        self.log(f"网络时间同步成功，偏差: {self.time_offset:.3f}秒")
                        return
                        
                    except Exception as e:
                        continue
                
                # 如果NTP失败，尝试HTTP方式
                try:
                    response = requests.get('http://worldtimeapi.org/api/timezone/Asia/Shanghai', timeout=3)
                    if response.status_code == 200:
                        data = response.json()
                        network_time = datetime.datetime.fromisoformat(data['datetime'].replace('Z', '+00:00'))
                        local_time = datetime.datetime.now()
                        self.time_offset = (network_time - local_time).total_seconds()
                        self.sync_status_label.config(text=f"已同步 (HTTP, 偏差: {self.time_offset:.3f}s)")
                        self.log(f"HTTP时间同步成功，偏差: {self.time_offset:.3f}秒")
                        return
                except:
                    pass
                    
                self.sync_status_label.config(text="同步失败", foreground="red")
                self.log("网络时间同步失败，将使用本地时间")
                
            except Exception as e:
                self.sync_status_label.config(text="同步失败", foreground="red")
                self.log(f"时间同步错误: {e}")
        
        threading.Thread(target=sync, daemon=True).start()
        
    def set_mouse_position(self):
        """保留占位（兼容旧调用），现逻辑不再需要预设鼠标位置"""
        self.status_label.config(text="启动后将在目标时间点击当前鼠标位置", foreground="blue")
        
    def toggle_grabber(self):
        """启动/停止抢票"""
        if not self.is_running:
            # 验证输入
            try:
                time_str = self.time_var.get().strip()
                self.target_time = datetime.datetime.strptime(time_str, "%H:%M:%S").time()
            except ValueError:
                messagebox.showerror("错误", "请输入正确的时间格式 (HH:MM:SS)")
                return
            
            # 计算本次的绝对目标时间（跨天处理）
            now = datetime.datetime.now() + datetime.timedelta(seconds=self.time_offset)
            target_dt = datetime.datetime.combine(now.date(), self.target_time)
            if target_dt <= now:
                target_dt += datetime.timedelta(days=1)
            self.target_datetime = target_dt

            self.is_running = True
            self.start_button.config(text="停止")
            self.status_label.config(text="等待目标时间，请将鼠标移至目标位置...", foreground="blue")
            self.log(f"开始等待目标时间: {time_str} (绝对时间: {self.target_datetime.strftime('%Y-%m-%d %H:%M:%S')})")
            
            # 启动点击线程
            threading.Thread(target=self.click_worker, daemon=True).start()
            
        else:
            self.is_running = False
            self.start_button.config(text="启动")
            self.status_label.config(text="已停止", foreground="red")
            self.log("抢票已停止")
            
    def click_worker(self):
        """点击工作线程"""
        while self.is_running:
            try:
                current_time = datetime.datetime.now() + datetime.timedelta(seconds=self.time_offset)
                # 使用启动时计算好的绝对目标时间，避免错过瞬间后被推到第二天
                time_diff = (self.target_datetime - current_time).total_seconds()

                if time_diff <= 0:  # 到点或已过，立即执行点击
                    start_time = time.perf_counter()
                    
                    # 执行点击：使用当前鼠标所在位置
                    current_pos = pyautogui.position()
                    pyautogui.click(current_pos.x, current_pos.y)
                    
                    end_time = time.perf_counter()
                    click_duration = (end_time - start_time) * 1000  # 转换为毫秒
                    
                    click_time = datetime.datetime.now() + datetime.timedelta(seconds=self.time_offset)
                    time_str = click_time.strftime("%H:%M:%S.%f")[:-3]
                    
                    self.log(f"{time_str} - 点击执行 ({click_duration:.1f}ms)")
                    self.status_label.config(text=f"点击完成 ({click_duration:.1f}ms)", foreground="green")
                    
                    # 自动停止
                    self.is_running = False
                    self.start_button.config(text="启动")
                    break
                    
                time.sleep(0.005)  # 降低CPU占用；Windows计时分辨率约15ms，无需更小
                
            except Exception as e:
                self.log(f"点击执行错误: {e}")
                self.status_label.config(text="点击失败", foreground="red")
                self.is_running = False
                self.start_button.config(text="启动")
                break
                
    def log(self, message):
        """添加日志"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        
    def run(self):
        """运行程序"""
        # 启动时自动同步网络时间
        self.sync_network_time()
        self.log("抢票工具已启动")
        self.root.mainloop()

if __name__ == "__main__":
    # 设置pyautogui
    pyautogui.FAILSAFE = True  # 将鼠标移至屏幕左上角可紧急停止
    
    app = TicketGrabber()
    app.run()
