import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import re
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
        self.opacity = 1.0  # 窗口透明度（1.0为不透明）
        self.course_files_dir = os.getcwd()  # 课程文件目录
        
        self.root.geometry(f"{self.window_width}x{self.window_height}+{self.default_x}+{self.default_y}")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", self.opacity)  # 设置透明度
        self.hide_from_taskbar()
        
        # 字体设置
        self.font_config = ('SimHei', 12)
        self.title_font = ('SimHei', 16, 'bold')
        self.time_font = ('SimHei', 12, 'bold')
        self.reminder_font = ('SimHei', 14, 'bold')
        
        # 星期映射
        self.weekday_map = {0: "周一", 1: "周二", 2: "周三", 3: "周四", 4: "周五", 5: "周六", 6: "周日"}
        self.today = self.weekday_map[datetime.now().weekday()]
        self.today_date = datetime.now().date()  # 今天的日期
        
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
        self.style.configure("Reminder.TFrame", background="#fff3cd")  # 提醒框背景色
        self.style.configure("OddRow.TLabel", background="#f8f9fa")
        self.style.configure("EvenRow.TLabel", background="#e9ecef")
        
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
        # 如果图标文件不存在，创建一个简单的默认图标
        if not os.path.exists("schedule_icon.png"):
            self.create_default_icon()
        
        image = Image.open("schedule_icon.png")
        
        # 托盘菜单设置
        menu = pystray.Menu(
            pystray.MenuItem("显示课程表", self.show_window),
            pystray.MenuItem("隐藏课程表", self.hide_window),
            pystray.MenuItem(f"拖动: {'开启' if self.allow_dragging else '关闭'}", self.toggle_dragging),
            pystray.MenuItem("编辑今日课程", self.open_edit_window),
            pystray.MenuItem("刷新课程表", self.refresh_schedule),
            pystray.MenuItem("退出", self.exit_app)
        )
        
        self.tray_icon = pystray.Icon("课程表", image, "课程表", menu)
        import threading
        threading.Thread(target=self.tray_icon.run, daemon=True).start()
    
    def create_default_icon(self):
        """创建默认图标"""
        image = Image.new('RGB', (64, 64), color=(73, 109, 137))
        draw = ImageDraw.Draw(image)
        draw.text((10, 20), "课", font=self.title_font, fill=(255, 255, 255))
        image.save("schedule_icon.png")
    
    def open_edit_window(self, icon=None, item=None):
        """打开课程编辑窗口，支持直接在表格中修改内容"""
        self.edit_window = tk.Toplevel(self.root)
        self.edit_window.title(f"编辑 {self.today} 课程")
        self.edit_window.geometry("600x400")
        self.edit_window.transient(self.root)
        self.edit_window.grab_set()  # 模态窗口
        
        # 创建主容器，使用grid布局管理
        main_container = ttk.Frame(self.edit_window)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 顶部按钮区域
        top_buttons = ttk.Frame(main_container)
        top_buttons.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(top_buttons, text="添加课程", command=self.add_course_row).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(top_buttons, text="删除选中", command=self.delete_course_row).pack(
            side=tk.LEFT, padx=5)
        
        # 创建编辑表格（支持单元格编辑）
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
        
        # 底部按钮区域，只保留保存和取消
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
    
    def load_courses_to_tree(self):
        """将当前课程加载到编辑表格"""
        courses = self.read_today_courses()
        for course in courses:
            time_str = f"{course['start']}~{course['end']}"
            self.tree.insert("", tk.END, values=(time_str, course['name']))
    
    def add_course_row(self):
        """添加空行用于输入新课程"""
        self.tree.insert("", tk.END, values=("00:00~00:00", ""))
    
    def delete_course_row(self):
        """删除选中的课程行"""
        selected = self.tree.selection()
        if selected:
            self.tree.delete(selected)
    
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
                    # 使用两个全角空格分隔（兼容原格式）
                    f.write(f"{course['start']}~{course['end']}  {course['name']}\n")
            
            messagebox.showinfo("保存成功", f"{self.today}课程已更新")
            self.edit_window.destroy()
            # 刷新主窗口显示
            self.load_and_display_schedule()
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存文件: {str(e)}")
        
    def toggle_dragging(self, icon=None, item=None):
        self.allow_dragging = not self.allow_dragging
        self.setup_dragging()
        self.update_tray_menu()
    
    def refresh_schedule(self, icon=None, item=None):
        """刷新课程表"""
        self.today = self.weekday_map[datetime.now().weekday()]
        self.title_label.config(text=f"{self.today} 课程表")
        self.load_and_display_schedule()
    
    def update_tray_menu(self):
        """更新托盘菜单"""
        self.tray_icon.menu = pystray.Menu(
            pystray.MenuItem("显示课程表", self.show_window),
            pystray.MenuItem("隐藏课程表", self.hide_window),
            pystray.MenuItem(f"拖动: {'开启' if self.allow_dragging else '关闭'}", self.toggle_dragging),
            pystray.MenuItem("编辑今日课程", self.open_edit_window),
            pystray.MenuItem("刷新课程表", self.refresh_schedule),
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
        """读取今天的课程文件（支持全角/半角空格分隔）"""
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
                        print(f"格式错误（需至少2个空格分隔）: {line}")
            
            # 按开始时间排序
            courses.sort(key=lambda x: x['start_datetime'])
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
        
        # 显示所有课程，添加交替行颜色
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
                width=10,
                style=style if not is_current else ""
            )
            # 当前课程高亮显示
            if is_current:
                time_label.configure(background="#d1ecf1", font=('SimHei', 12, 'bold', 'underline'))
                
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
                wraplength=250,  # 适配窗口宽度
                style=style if not is_current else ""
            )
            # 当前课程高亮显示
            if is_current:
                course_label.configure(background="#d1ecf1", font=('SimHei', 12, 'bold', 'underline'))
                
            course_label.grid(row=i, column=1, sticky="nsew", pady=2)
        
        # 确保行和列能自适应内容
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(1, weight=2)
        # 手动触发滚动区域更新
        self.root.after(100, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = CustomizableScheduleApp()
    app.run()
    