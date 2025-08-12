import tkinter as tk
from tkinter import ttk
import os
import sys
import re  # 新增：用于处理各种空格分隔符
from datetime import datetime
import pystray
from PIL import Image, ImageDraw
import win32gui
import win32con

class CustomizableScheduleApp:
    def __init__(self):
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("今日课程表")
        
        # 窗口基本设置
        self.window_width = 400     # 窗口宽度（像素）
        self.window_height = 800    # 窗口高度（像素）
        self.default_x = 1310      # 窗口默认X坐标
        self.default_y = 0          # 窗口默认Y坐标
        self.allow_dragging = False # 是否允许拖动
        
        self.root.geometry(f"{self.window_width}x{self.window_height}+{self.default_x}+{self.default_y}")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.hide_from_taskbar()
        
        # 字体设置
        self.font_config = ('SimHei', 12)
        self.title_font = ('SimHei', 16, 'bold')
        self.time_font = ('SimHei', 12, 'bold')
        
        # 星期映射
        self.weekday_map = {0: "周一", 1: "周二", 2: "周三", 3: "周四", 4: "周五", 5: "周六", 6: "周日"}
        self.today = self.weekday_map[datetime.now().weekday()]
        
        # 创建UI框架
        self.main_frame = ttk.Frame(self.root, padding="20 20 20 20", style="Main.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.title_label = ttk.Label(self.main_frame, text=f"{self.today} 课程表", font=self.title_font)
        self.title_label.pack(pady=(0, 20))
        
        self.content_container = ttk.Frame(self.main_frame)
        self.content_container.pack(fill=tk.BOTH, expand=True)
        
        self.canvas_frame = ttk.Frame(self.content_container)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(self.canvas_frame)
        self.vscrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.configure(yscrollcommand=self.vscrollbar.set)
        
        self.content_frame = ttk.Frame(self.canvas, width=self.window_width-60)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        
        # 绑定滚动事件
        self.content_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # 拖动设置
        self.x = 0
        self.y = 0
        self.setup_dragging()
        
        # 样式配置
        self.style = ttk.Style()
        self.style.configure("Main.TFrame", background="#f0f0f0")
        
        # 加载课程
        self.load_and_display_schedule()
        
        # 托盘图标
        self.create_tray_icon()
        
    def setup_dragging(self):
        if self.allow_dragging:
            self.main_frame.bind("<ButtonPress-1>", self.start_move)
            self.main_frame.bind("<B1-Motion>", self.on_move)
        else:
            self.main_frame.unbind("<ButtonPress-1>")
            self.main_frame.unbind("<B1-Motion>")
        
    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        
    def hide_from_taskbar(self):
        try:
            hwnd = self.root.winfo_id()
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, 
                                 win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | 
                                 win32con.WS_EX_TOOLWINDOW)
        except Exception as e:
            print(f"隐藏任务栏图标失败: {e}")
        
    def create_tray_icon(self):
        image = Image.open("schedule_icon.png")
        
        menu = pystray.Menu(
            pystray.MenuItem("显示课程表", self.show_window),
            pystray.MenuItem(f"拖动: {'开启' if self.allow_dragging else '关闭'}", self.toggle_dragging),
            pystray.MenuItem("退出", self.exit_app)
        )
        
        self.tray_icon = pystray.Icon("课程表", image, "课程表", menu)
        import threading
        threading.Thread(target=self.tray_icon.run, daemon=True).start()
        
    def toggle_dragging(self, icon=None, item=None):
        self.allow_dragging = not self.allow_dragging
        self.setup_dragging()
        if icon and item:
            icon.menu = pystray.Menu(
                pystray.MenuItem("显示课程表", self.show_window),
                pystray.MenuItem(f"拖动: {'开启' if self.allow_dragging else '关闭'}", self.toggle_dragging),
                pystray.MenuItem("退出", self.exit_app)
            )
        
    def show_window(self, icon=None, item=None):
        self.root.deiconify()
        self.root.attributes("-topmost", True)
        self.root.attributes("-topmost", False)
        
    def hide_window(self, icon=None, item=None):
        self.root.withdraw()
        
    def exit_app(self, icon=None, item=None):
        if icon:
            icon.stop()
        self.root.destroy()
        sys.exit(0)
        
    def start_move(self, event):
        self.x = event.x
        self.y = event.y
        
    def on_move(self, event):
        x = self.root.winfo_x() + event.x - self.x
        y = self.root.winfo_y() + event.y - self.y
        self.root.geometry(f"+{x}+{y}")
    
    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def read_today_courses(self):
        """读取今天的课程文件（修复：支持全角/半角空格分隔）"""
        filename = f"{self.today}.txt"
        courses = []
        
        if not os.path.exists(filename):
            print(f"未找到课程文件: {filename}")
            return courses
            
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                        
                    # 关键修复：使用正则匹配2个及以上任意空格（包括全角/半角）
                    if re.search(r'\s{2,}', line):  # 匹配至少2个空白字符
                        # 按2个及以上空格分割（保留课程名中的空格）
                        parts = re.split(r'\s{2,}', line, 1)  # 只分割一次
                        if len(parts) == 2:
                            time_part, name_part = parts
                            if '~' in time_part:
                                start_time, end_time = time_part.split('~')
                                courses.append({
                                    'start': start_time,
                                    'end': end_time,
                                    'name': name_part
                                })
                            else:
                                print(f"时间格式错误: {line}")
                        else:
                            print(f"分割失败: {line}")
                    else:
                        print(f"格式错误（需至少2个空格分隔）: {line}")
            
            courses.sort(key=lambda x: datetime.strptime(x['start'], '%H:%M'))
            return courses
            
        except Exception as e:
            print(f"读取{self.today}课程文件出错: {e}")
            return []
    
    def load_and_display_schedule(self):
        """加载并显示今日课程（确保滚动区域包含所有内容）"""
        courses = self.read_today_courses()
        
        # 清除现有内容
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        if not courses:
            no_course_label = ttk.Label(
                self.content_frame, 
                text="今天没有课程，好好休息吧！", 
                font=self.font_config
            )
            no_course_label.pack(pady=50)
            return
        
        # 创建表头
        time_header = ttk.Label(
            self.content_frame, 
            text="时间", 
            font=self.title_font,
            borderwidth=1, 
            relief="solid",
            padding=8,
            anchor="center",
            width=10
        )
        time_header.grid(row=0, column=0, sticky="nsew")
        
        course_header = ttk.Label(
            self.content_frame, 
            text="课程", 
            font=self.title_font,
            borderwidth=1, 
            relief="solid",
            padding=8,
            anchor="center",
            width=20
        )
        course_header.grid(row=0, column=1, sticky="nsew")
        
        # 显示所有课程
        for i, course in enumerate(courses, start=1):
            # 时间标签
            time_label = ttk.Label(
                self.content_frame, 
                text=f"{course['start']}~{course['end']}", 
                font=self.time_font,
                borderwidth=1, 
                relief="solid",
                padding=10,
                anchor="center",
                width=10
            )
            time_label.grid(row=i, column=0, sticky="nsew", pady=2)
            
            # 课程名称标签
            course_label = ttk.Label(
                self.content_frame, 
                text=course['name'], 
                font=self.font_config,
                borderwidth=1, 
                relief="solid",
                padding=10,
                anchor="w",
                width=20,
                wraplength=250  # 适配窗口宽度
            )
            course_label.grid(row=i, column=1, sticky="nsew", pady=2)
        
        # 确保行和列能自适应内容
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(1, weight=2)
        # 手动触发滚动区域更新（确保包含所有课程）
        self.root.after(100, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = CustomizableScheduleApp()
    app.run()