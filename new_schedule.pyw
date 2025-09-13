import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import re
import json
from datetime import datetime
import pystray
from PIL import Image

import win32gui
import win32con

class ScheduleApplication:
    def __init__(self):
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("今日课程表")
        
        # 窗口基本设置
        self.window_width = 400
        self.window_height = 1000
        self.default_x = 1310
        self.default_y = 0
        self.allow_dragging = False
        self.opacity = 1.0
        self.course_files_dir = os.getcwd()
        self.config_file = os.path.join(self.course_files_dir, "schedule_config.json")
        
        # 字体设置 - 默认字体大小，可在此处更改初始默认值
        # 注意：程序启动时会尝试从schedule_config.json加载已保存的字体大小配置
        self.time_font_size = 18        # 时间字体默认大小
        self.course_font_size = 18      # 课程字体默认大小
        self.title_font_size = 18       # 标题字体默认大小
        self.update_fonts()
        
        # 加载保存的位置配置
        self.load_window_position()
        
        self.root.geometry(f"{self.window_width}x{self.window_height}+{self.default_x}+{self.default_y}")
        self.root.overrideredirect(True)  # 无边框窗口
        self.root.attributes("-alpha", self.opacity)
        self.hide_from_taskbar()
        
        # 绑定窗口移动事件，在移动后保存位置
        self.root.bind("<Configure>", self.on_window_configure)
        
        # 星期映射
        self.weekday_map = {0: "周一", 1: "周二", 2: "周三", 3: "周四", 4: "周五", 5: "周六", 6: "周日"}
        self.today = self.weekday_map[datetime.now().weekday()]
        
        # 检查并创建缺失的课程文件
        self.check_and_create_course_files()
        self.today_date = datetime.now().date()
        
        # 设置浙江首考和高考日期（这里使用2024年的日期作为示例，实际使用时请更新）
        # 浙江首考时间通常在1月
        self.first_exam_date = datetime(2026, 1, 6).date()  
        # 高考时间通常在6月7日开始
        self.college_exam_date = datetime(2026, 6, 7).date()  
        
        # 创建UI组件
        self.setup_ui()
        
        # 加载课程
        self.load_and_display_schedule()
        
        # 创建托盘图标
        self.create_tray_icon()
        
        # 更新倒计时
        self.update_countdown()
    
    def update_fonts(self):
        """更新字体配置"""
        self.font_config = ('SimHei', self.course_font_size)
        self.title_font = ('SimHei', self.title_font_size, 'bold')
        self.time_font = ('SimHei', self.time_font_size, 'bold')
    
    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="20 20 20 20", style="Main.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 设置样式
        self.style = ttk.Style()
        self.style.configure("Main.TFrame", background="#f0f0f0")
        self.style.configure("OddRow.TLabel", background="#f8f9fa")
        self.style.configure("EvenRow.TLabel", background="#e9ecef")
        self.style.configure("Countdown.TFrame", background="#fff3cd")
        self.style.configure("Countdown.TLabel", background="#fff3cd", font=('SimHei', 12, 'bold'))
        
        # 添加倒计时显示区域
        self.countdown_frame = ttk.Frame(self.main_frame, padding="10 10 10 10", style="Countdown.TFrame")
        self.countdown_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 首考倒计时标签
        self.first_exam_label = ttk.Label(self.countdown_frame, text="浙江首考倒计时: ", style="Countdown.TLabel")
        self.first_exam_label.pack(fill=tk.X, pady=2)
        
        # 高考倒计时标签
        self.college_exam_label = ttk.Label(self.countdown_frame, text="高考倒计时: ", style="Countdown.TLabel")
        self.college_exam_label.pack(fill=tk.X, pady=2)
        
        # 标题标签
        self.title_label = ttk.Label(self.main_frame, text=f"{self.today} 课程表", font=self.title_font)
        self.title_label.pack(pady=(0, 20))
        
        # 创建可滚动内容区域
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
        
        # 初始化拖动功能
        self.setup_dragging()
        self.x = 0
        self.y = 0
    
    def setup_dragging(self):
        """设置窗口拖动功能，扩大到整张课表"""
        if self.allow_dragging:
            # 将拖动事件绑定到整个画布上，而不仅仅是主框架
            self.canvas.bind("<ButtonPress-1>", self.start_move)
            self.canvas.bind("<B1-Motion>", self.on_move)
            self.main_frame.bind("<ButtonPress-1>", self.start_move)
            self.main_frame.bind("<B1-Motion>", self.on_move)
        else:
            # 解除所有绑定
            self.canvas.unbind("<ButtonPress-1>")
            self.canvas.unbind("<B1-Motion>")
            self.main_frame.unbind("<ButtonPress-1>")
            self.main_frame.unbind("<B1-Motion>")
    
    def hide_from_taskbar(self):
        """隐藏任务栏图标"""
        try:
            hwnd = self.root.winfo_id()
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                                 win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) |
                                 win32con.WS_EX_TOOLWINDOW)
        except Exception as e:
            print(f"隐藏任务栏图标失败: {e}")
    
    def create_tray_icon(self):
        """创建系统托盘图标"""
        # 如果图标文件不存在，创建默认图标
        if not os.path.exists("schedule_icon.png"):
            self.create_default_icon()
        
        image = Image.open("schedule_icon.png")
        
        # 创建托盘菜单
        self.menu = pystray.Menu(
            pystray.MenuItem("显示课程表", self.show_window),
            pystray.MenuItem("隐藏课程表", self.hide_window),
            pystray.MenuItem(f"拖动: {'开启' if self.allow_dragging else '关闭'}", self.toggle_dragging),
            pystray.MenuItem("字体设置", pystray.Menu(
                pystray.MenuItem("时间字体增大", self.increase_time_font),
                pystray.MenuItem("时间字体减小", self.decrease_time_font),
                pystray.MenuItem("课程名字字体增大", self.increase_course_font),
                pystray.MenuItem("课程名字字体减小", self.decrease_course_font),
            )),
            pystray.MenuItem("编辑今日课程", self.open_edit_window),
            pystray.MenuItem("刷新课程表", self.refresh_schedule),
            pystray.MenuItem("退出", self.exit_app)
        )
        
        self.tray_icon = pystray.Icon("课程表", image, "课程表", self.menu)
        import threading
        threading.Thread(target=self.tray_icon.run, daemon=True).start()
    
    def create_default_icon(self):
        """创建默认图标"""
        image = Image.new('RGB', (64, 64), color=(73, 109, 137))
        draw = ImageDraw.Draw(image)
        draw.text((10, 20), "课", font=self.title_font, fill=(255, 255, 255))
        image.save("schedule_icon.png")
    
    def read_today_courses(self):
        """读取今天的课程数据"""
        filename = os.path.join(self.course_files_dir, f"{self.today}.txt")
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
                        
                    # 匹配至少2个空白字符（包括全角/半角空格）
                    if re.search(r'\s{2,}', line):
                        # 按2个及以上空格分割（保留课程名中的空格）
                        parts = re.split(r'\s{2,}', line, 1)  # 只分割一次
                        if len(parts) == 2:
                            time_part, name_part = parts
                            if '~' in time_part:
                                start_time, end_time = time_part.split('~')
                                # 转换为datetime对象便于比较
                                try:
                                    start_datetime = datetime.strptime(f"{self.today_date} {start_time}", 
                                                                    "%Y-%m-%d %H:%M")
                                    end_datetime = datetime.strptime(f"{self.today_date} {end_time}", 
                                                                  "%Y-%m-%d %H:%M")
                                    courses.append({
                                        'start': start_time,
                                        'end': end_time,
                                        'start_datetime': start_datetime,
                                        'end_datetime': end_datetime,
                                        'name': name_part,
                                        'id': f"{start_time}-{name_part}"  # 唯一标识
                                    })
                                except ValueError:
                                    print(f"时间格式错误，应为HH:MM: {line}")
                            else:
                                print(f"时间格式错误，需包含~: {line}")
                        else:
                            print(f"分割失败: {line}")
                    else:
                        # 尝试处理可能只有时间没有课程名称的情况
                        if '~' in line:
                            try:
                                start_time, end_time = line.split('~')
                                start_datetime = datetime.strptime(f"{self.today_date} {start_time.strip()}",
                                                                 "%Y-%m-%d %H:%M")
                                end_datetime = datetime.strptime(f"{self.today_date} {end_time.strip()}",
                                                               "%Y-%m-%d %H:%M")
                                courses.append({
                                    'start': start_time.strip(),
                                    'end': end_time.strip(),
                                    'start_datetime': start_datetime,
                                    'end_datetime': end_datetime,
                                    'name': "",
                                    'id': f"{start_time.strip()}-empty"
                                })
                            except ValueError:
                                print(f"时间格式错误: {line}")
                        else:
                            print(f"格式错误: {line}")
            
            # 按开始时间排序
            courses.sort(key=lambda x: x['start_datetime'])
            return courses
        except Exception as e:
            print(f"读取{self.today}课程文件出错: {e}")
            return []
    
    def load_and_display_schedule(self):
        """加载并显示今日课程表"""
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
        now = datetime.now()
        for i, course in enumerate(courses, start=1):
            # 判断是否是当前进行中的课程
            is_current = course['start_datetime'] <= now <= course['end_datetime']
            # 交替行样式
            style = "EvenRow.TLabel" if i % 2 == 0 else "OddRow.TLabel"
            
            # 时间标签
            time_label = ttk.Label(
                self.content_frame,
                text=f"{course['start']}~{course['end']}",
                font=self.time_font,
                borderwidth=1,
                relief="solid",
                padding=10,
                anchor="center",
                width=17,
                style=style if not is_current else ""
            )
            
            # 当前课程高亮显示
            if is_current:
                time_label.configure(background="#d1ecf1", font=('SimHei', self.time_font_size, 'bold', 'underline'))
                
            time_label.grid(row=i, column=0, sticky="nsew", pady=2)
            
            # 课程名称标签
            course_label = ttk.Label(
                self.content_frame,
                text=course['name'],
                font=self.font_config,
                borderwidth=1,
                relief="solid",
                padding=10,
                anchor="center",
                width=17,
                wraplength=250,
                style=style if not is_current else ""
            )
            
            # 当前课程高亮显示
            if is_current:
                course_label.configure(background="#d1ecf1", font=('SimHei', self.course_font_size, 'bold', 'underline'))
                
            course_label.grid(row=i, column=1, sticky="nsew", pady=2)
        
        # 确保行和列能自适应内容
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(1, weight=2)
        
        # 手动触发滚动区域更新
        self.root.after(100, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
    
    def open_edit_window(self, icon=None, item=None):
        """打开课程编辑窗口"""
        self.edit_window = tk.Toplevel(self.root)
        self.edit_window.title(f"编辑 {self.today} 课程")
        self.edit_window.geometry("600x400")
        self.edit_window.transient(self.root)
        self.edit_window.grab_set()  # 模态窗口
        
        # 创建主容器
        main_container = ttk.Frame(self.edit_window)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 顶部按钮区域
        top_buttons = ttk.Frame(main_container)
        top_buttons.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(top_buttons, text="添加课程", command=self.add_course_row).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(top_buttons, text="删除选中", command=self.delete_course_row).pack(
            side=tk.LEFT, padx=5)
        
        # 创建编辑表格
        columns = ("时间", "课程名称")
        self.tree = ttk.Treeview(main_container, columns=columns, show="headings")
        self.tree.heading("时间", text="时间 (格式: HH:MM~HH:MM)")
        self.tree.heading("课程名称", text="课程名称")
        self.tree.column("时间", width=200)
        self.tree.column("课程名称", width=350)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局表格和滚动条
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部按钮区域
        btn_frame = ttk.Frame(self.edit_window)
        btn_frame.pack(side=tk.BOTTOM, pady=10, fill=tk.X, padx=10)
        
        ttk.Button(btn_frame, text="保存修改", command=self.save_course_changes).pack(
            side=tk.RIGHT, padx=10)
        ttk.Button(btn_frame, text="取消修改", command=self.edit_window.destroy).pack(
            side=tk.RIGHT)
        
        # 加载现有课程到表格
        self.load_courses_to_tree()
        
        # 绑定双击事件实现单元格编辑
        self.tree.bind("<Double-1>", self.on_tree_double_click)
    
    def load_courses_to_tree(self):
        """将当前课程加载到编辑表格"""
        courses = self.read_today_courses()
        for course in courses:
            time_str = f"{course['start']}~{course['end']}"
            self.tree.insert("", tk.END, values=(time_str, course['name']))
    
    def on_tree_double_click(self, event):
        """双击表格单元格进行编辑"""
        # 获取双击位置的信息
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            # 获取行和列
            row = self.tree.identify_row(event.y)
            column = self.tree.identify_column(event.x)
            column_index = int(column.replace('#', '')) - 1  # 转换为0-based索引
            
            # 获取单元格位置和值
            x, y, width, height = self.tree.bbox(row, column)
            current_value = self.tree.item(row, "values")[column_index]
            
            # 创建编辑条目
            self.edit_entry = ttk.Entry(self.tree)
            self.edit_entry.place(x=x, y=y, width=width, height=height)
            self.edit_entry.insert(0, current_value)
            self.edit_entry.focus()
            
            # 保存修改的函数
            def save_edit(event=None):
                new_value = self.edit_entry.get()
                # 更新树视图中的值
                values = list(self.tree.item(row, "values"))
                values[column_index] = new_value
                self.tree.item(row, values=values)
                self.edit_entry.destroy()
            
            # 绑定回车键和焦点离开事件保存修改
            self.edit_entry.bind("<FocusOut>", save_edit)
            self.edit_entry.bind("<Return>", save_edit)
            self.edit_entry.bind("<Escape>", lambda e: self.edit_entry.destroy())
    
    def add_course_row(self):
        """添加空行用于输入新课程"""
        self.tree.insert("", tk.END, values=("00:00~00:00", ""))
    
    def delete_course_row(self):
        """删除选中的课程行"""
        selected = self.tree.selection()
        if selected:
            for item in selected:
                self.tree.delete(item)
    
    def save_course_changes(self):
        """保存编辑后的课程到文件"""
        filename = os.path.join(self.course_files_dir, f"{self.today}.txt")
        courses = []
        
        # 收集表格数据并验证
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            time_str, name = values[0].strip(), values[1].strip()
            
            # 验证时间格式
            if "~" not in time_str:
                messagebox.showerror("格式错误", f"时间格式不正确: {time_str}\n请使用 HH:MM~HH:MM 格式")
                return
            
            start, end = time_str.split("~", 1)
            try:
                # 验证时间格式是否正确
                datetime.strptime(start, "%H:%M")
                datetime.strptime(end, "%H:%M")
                courses.append({
                    "start": start,
                    "end": end,
                    "name": name
                })
            except ValueError:
                messagebox.showerror("格式错误", f"时间格式不正确: {time_str}\n请使用 HH:MM 格式")
                return
        
        # 按开始时间排序
        courses.sort(key=lambda x: datetime.strptime(x['start'], "%H:%M"))
        
        # 写入文件
        try:
            with open(filename, "w", encoding="utf-8") as f:
                for course in courses:
                    # 使用两个全角空格分隔
                    f.write(f"{course['start']}~{course['end']}  {course['name']}\n")
            
            messagebox.showinfo("保存成功", f"{self.today}课程已更新")
            self.edit_window.destroy()
            # 刷新主窗口显示
            self.load_and_display_schedule()
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存文件: {str(e)}")
    
    # 托盘菜单功能实现
    def toggle_dragging(self, icon=None, item=None):
        """切换窗口拖动功能"""
        self.allow_dragging = not self.allow_dragging
        self.setup_dragging()
        self.update_tray_menu()
        
        # 提示用户操作结果
        message = "已开启拖动功能，点击窗口任意位置拖动"
        if not self.allow_dragging:
            message = "已关闭拖动功能，窗口位置已固定"
            # 关闭拖动时保存位置
            self.save_window_position()
        self.show_tooltip(message)
    
    def update_tray_menu(self):
        """更新托盘菜单"""
        # 更新主托盘菜单
        self.tray_icon.menu = pystray.Menu(
            pystray.MenuItem("显示课程表", self.show_window),
            pystray.MenuItem("隐藏课程表", self.hide_window),
            pystray.MenuItem(f"拖动: {'开启' if self.allow_dragging else '关闭'}", self.toggle_dragging),
            pystray.MenuItem("字体设置", pystray.Menu(
                pystray.MenuItem("时间字体增大", self.increase_time_font),
                pystray.MenuItem("时间字体减小", self.decrease_time_font),
                pystray.MenuItem("课程名字字体增大", self.increase_course_font),
                pystray.MenuItem("课程名字字体减小", self.decrease_course_font),
            )),
            pystray.MenuItem("编辑今日课程", self.open_edit_window),
            pystray.MenuItem("刷新课程表", self.refresh_schedule),
            pystray.MenuItem("退出", self.exit_app)
        )
    
    def refresh_schedule(self, icon=None, item=None):
        """刷新课程表"""
        self.today = self.weekday_map[datetime.now().weekday()]
        self.title_label.config(text=f"{self.today} 课程表")
        self.load_and_display_schedule()
    
    def show_window(self, icon=None, item=None):
        """显示窗口"""
        self.root.deiconify()
    
    def hide_window(self, icon=None, item=None):
        """隐藏窗口"""
        self.root.withdraw()
    
    def exit_app(self, icon=None, item=None):
        """退出应用程序"""
        # 保存窗口位置
        self.save_window_position()
        if icon:
            icon.stop()
        self.root.destroy()
        sys.exit(0)
    
    # 窗口拖动功能
    def start_move(self, event):
        self.x = event.x
        self.y = event.y
    
    def on_move(self, event):
        x = self.root.winfo_x() + event.x - self.x
        y = self.root.winfo_y() + event.y - self.y
        self.root.geometry(f"+{x}+{y}")
        
    def on_window_configure(self, event):
        """窗口配置变化时的处理"""
        # 确保这是实际的移动事件而不是其他配置变化
        if hasattr(event, 'width') and hasattr(event, 'height'):
            return
        
    def save_window_position(self):
        """保存窗口位置到配置文件"""
        try:
            x = self.root.winfo_x()
            y = self.root.winfo_y()
            
            config = {
                "window_position": {
                    "x": x,
                    "y": y
                },
                "allow_dragging": self.allow_dragging,
                "font_sizes": {
                    "time_font_size": self.time_font_size,
                    "course_font_size": self.course_font_size,
                    "title_font_size": self.title_font_size
                }
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存窗口位置失败: {e}")
    
    def load_window_position(self):
        """从配置文件加载窗口位置和背景图片设置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 加载窗口位置
                    if "window_position" in config:
                        self.default_x = config["window_position"].get("x", self.default_x)
                        self.default_y = config["window_position"].get("y", self.default_y)
                    # 加载拖动设置
                    if "allow_dragging" in config:
                        self.allow_dragging = config["allow_dragging"]
                    # 加载字体大小设置
                    if "font_sizes" in config:
                        font_config = config["font_sizes"]
                        self.time_font_size = font_config.get("time_font_size", self.time_font_size)
                        self.course_font_size = font_config.get("course_font_size", self.course_font_size)
                        self.title_font_size = font_config.get("title_font_size", self.title_font_size)
        except Exception as e:
            print(f"加载窗口位置失败: {e}")
    
    def show_tooltip(self, message):
        """显示临时提示消息"""
        tooltip = tk.Toplevel(self.root)
        tooltip.overrideredirect(True)
        tooltip.attributes("-alpha", 0.9)
        tooltip.geometry(f"+{self.root.winfo_x() + 20}+{self.root.winfo_y() + 20}")
        
        label = ttk.Label(tooltip, text=message, padding=(10, 5), background="#333", foreground="white")
        label.pack()
        
        # 2秒后自动关闭提示
        tooltip.after(2000, tooltip.destroy)
    
    # 滚动相关功能
    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    # 字体调整功能
    def increase_time_font(self, icon=None, item=None):
        self.time_font_size += 2
        self.update_fonts()
        self.load_and_display_schedule()
        self.save_window_position()  # 保存字体大小设置
    
    def decrease_time_font(self, icon=None, item=None):
        if self.time_font_size > 8:
            self.time_font_size -= 2
            self.update_fonts()
            self.load_and_display_schedule()
            self.save_window_position()  # 保存字体大小设置
    
    def increase_course_font(self, icon=None, item=None):
        self.course_font_size += 2
        self.update_fonts()
        self.load_and_display_schedule()
        self.save_window_position()  # 保存字体大小设置
    
    def decrease_course_font(self, icon=None, item=None):
        if self.course_font_size > 8:
            self.course_font_size -= 2
            self.update_fonts()
            self.load_and_display_schedule()
            # 更新倒计时字体
            self.style.configure("Countdown.TLabel", font=('SimHei', min(12, self.course_font_size + 2), 'bold'))
            self.save_window_position()  # 保存字体大小设置
    
    def update_countdown(self):
        """更新倒计时显示"""
        now = datetime.now().date()
        
        # 计算首考倒计时
        if now < self.first_exam_date:
            days_until_first_exam = (self.first_exam_date - now).days
            self.first_exam_label.config(text=f"浙江首考倒计时: {days_until_first_exam} 天")
        else:
            # 如果已经过了首考日期，显示已结束
            days_past = (now - self.first_exam_date).days
            self.first_exam_label.config(text=f"浙江首考已结束: {days_past} 天前")
        
        # 计算高考倒计时
        if now < self.college_exam_date:
            days_until_college_exam = (self.college_exam_date - now).days
            self.college_exam_label.config(text=f"高考倒计时: {days_until_college_exam} 天")
        else:
            # 如果已经过了高考日期，显示已结束
            days_past = (now - self.college_exam_date).days
            self.college_exam_label.config(text=f"高考已结束: {days_past} 天前")
        
        # 每60秒更新一次倒计时
        self.root.after(60000, self.update_countdown)
    

    
    def check_and_create_course_files(self):
        """检查并创建缺失的课程文件"""
        # 周一文件路径
        monday_file = os.path.join(self.course_files_dir, "周一.txt")
        template_content = """
6:40~7:20  英语
7:30~8:10  英语
8:20~9:00  语文
9:25~10:05  语文
10:15~10:55  数学
11:05~11:45  语文
13:50~14:30  语文
14:40~15:20  语文
15:35~16:15  语文
16:25~16:55  语文
17:50~18:30  语文
18:40~19:20  语文
19:30~20:10  语文
20:30~21:10  语文
21:20~22:10  语文
"""
        
        # 如果周一.txt存在，读取其内容作为模板
        if os.path.exists(monday_file):
            try:
                with open(monday_file, 'r', encoding='utf-8') as f:
                    template_content = f.read()
            except Exception as e:
                print(f"读取周一文件作为模板失败: {e}")
        
        # 检查所有星期文件
        for weekday in self.weekday_map.values():
            file_path = os.path.join(self.course_files_dir, f"{weekday}.txt")
            if not os.path.exists(file_path):
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(template_content)
                    print(f"已创建缺失的课程文件: {weekday}.txt")
                except Exception as e:
                    print(f"创建{weekday}.txt文件失败: {e}")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = ScheduleApplication()
    app.run()