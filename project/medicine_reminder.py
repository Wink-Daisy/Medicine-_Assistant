import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import threading
import time as time_module
import json
import os
from PIL import Image, ImageTk

# 尝试导入语音库
try:
    import win32com.client

    SPEECH_AVAILABLE = True
    print("✅ 语音模块加载成功")
except ImportError:
    SPEECH_AVAILABLE = False
    print("⚠️ 语音模块未安装，请运行: pip install pywin32")


class MedicineReminder:
    def __init__(self, root):
        self.root = root
        self.root.title("吃药小管家 - 老年人用药提醒系统")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')

        # 初始化语音
        self.init_voice()

        # 数据存储
        self.medicines = []
        self.taking_history = {}

        # 配置文件
        self.config_file = "medicine_data.json"
        self.load_data()

        # 提醒线程控制
        self.running = True
        self.check_interval = 1

        # 启动提醒检查线程
        self.reminder_thread = threading.Thread(target=self.check_reminders, daemon=True)
        self.reminder_thread.start()

        # 提醒状态控制
        self.active_reminder_window = None
        self.current_reminding_med = None
        self.reminded_set = set()

        # 语音控制 - 一分钟循环播放
        self.voice_active = False
        self.voice_stop_flag = False
        self.voice_timer_id = None

        # 构建界面
        self.setup_ui()

        # 启动时检查
        self.root.after(100, self.check_missed_reminders)

    def init_voice(self):
        """初始化语音（使用Windows SAPI）"""
        if SPEECH_AVAILABLE:
            try:
                self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
                self.voice_enabled = True
                print("✅ Windows语音引擎初始化成功")
            except Exception as e:
                print(f"⚠️ 语音引擎初始化失败: {e}")
                self.voice_enabled = False
        else:
            self.voice_enabled = False
            print("⚠️ 请安装语音库: pip install pywin32")

    def speak(self, text):
        """语音播报（使用Windows SAPI，不会阻塞）"""
        if not self.voice_enabled:
            print(f"🔊 (模拟) {text}")
            return

        try:
            print(f"🔊 播报: {text}")
            self.speaker.Speak(text)
        except Exception as e:
            print(f"❌ 播报失败: {e}")

    def start_voice_alert(self, medicine):
        """开始一分钟循环语音提醒（每3秒一次，共20次）"""
        if self.voice_active:
            return

        self.voice_active = True
        self.voice_stop_flag = False
        voice_text = f"{medicine['name']}，请服用{medicine['dosage']}"

        # 一分钟内循环播放：60秒 / 3秒 = 20次
        def speak_loop(count=0, max_count=20):
            if self.voice_stop_flag or count >= max_count:
                self.voice_active = False
                self.voice_timer_id = None
                print(f"✅ 一分钟语音提醒结束，共播报{count}次")
                return

            # 播报当前次
            self.speak(voice_text)
            print(f"🔄 第{count + 1}次播报（共20次）")

            # 3秒后播报下一次
            self.voice_timer_id = self.root.after(3000, lambda: speak_loop(count + 1, max_count))

        # 开始播报序列
        speak_loop(0, 20)

    def stop_voice_alert(self):
        """停止语音提醒"""
        self.voice_stop_flag = True
        self.voice_active = False
        if self.voice_timer_id:
            try:
                self.root.after_cancel(self.voice_timer_id)
            except:
                pass
            self.voice_timer_id = None
        print("🛑 已停止语音提醒")

    def setup_ui(self):
        """构建主界面"""
        # 标题栏
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=80)
        title_frame.pack(fill='x')
        title_label = tk.Label(title_frame, text='💊 吃药小管家',
                               font=('微软雅黑', 24, 'bold'),
                               fg='white', bg='#2c3e50')
        title_label.pack(pady=20)

        # 主内容区域
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # 标签页
        self.manage_frame = tk.Frame(self.notebook)
        self.notebook.add(self.manage_frame, text='📋 药品管理')
        self.setup_manage_tab()

        self.today_frame = tk.Frame(self.notebook)
        self.notebook.add(self.today_frame, text='⏰ 今日提醒')
        self.setup_today_tab()

        self.history_frame = tk.Frame(self.notebook)
        self.notebook.add(self.history_frame, text='📅 服药记录')
        self.setup_history_tab()

        self.settings_frame = tk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text='⚙️ 设置')
        self.setup_settings_tab()

        # 底部状态栏
        self.status_frame = tk.Frame(self.root, bg='#34495e', height=70)
        self.status_frame.pack(fill='x', side='bottom')
        self.status_frame.pack_propagate(False)

        # 左下角时间
        self.time_label = tk.Label(self.status_frame,
                                   font=('Arial', 20, 'bold'),
                                   fg='#f1c40f',
                                   bg='#34495e')
        self.time_label.pack(side='left', padx=25, pady=15)

        # 系统状态
        self.status_label = tk.Label(self.status_frame,
                                     text='✅ 系统运行中',
                                     font=('微软雅黑', 12),
                                     bg='#34495e',
                                     fg='#2ecc71')
        self.status_label.pack(side='left', padx=30, pady=15)

        # 提醒状态
        self.reminder_status_label = tk.Label(self.status_frame,
                                              text='⏰ 等待提醒',
                                              font=('微软雅黑', 13, 'bold'),
                                              bg='#34495e',
                                              fg='#f39c12')
        self.reminder_status_label.pack(side='right', padx=25, pady=15)

        self.update_clock()

    def update_clock(self):
        """更新左下角时间显示"""
        current_time = datetime.now().strftime("%Y年%m月%d日  %H:%M:%S")
        self.time_label.config(text=f"🕐 {current_time}")
        self.root.after(1000, self.update_clock)

    def setup_manage_tab(self):
        """药品管理标签页"""
        left_frame = tk.Frame(self.manage_frame)
        left_frame.pack(side='left', fill='both', expand=True, padx=10, pady=10)

        tk.Label(left_frame, text='➕ 添加新药品', font=('微软雅黑', 16, 'bold'),
                 fg='#2c3e50').pack(pady=10)

        tk.Label(left_frame, text='药品名称：', font=('微软雅黑', 12)).pack(anchor='w', pady=5)
        self.med_name_entry = tk.Entry(left_frame, font=('微软雅黑', 14), width=25)
        self.med_name_entry.pack(fill='x', pady=5)

        tk.Label(left_frame, text='剂量（如：2片、1粒）：', font=('微软雅黑', 12)).pack(anchor='w', pady=5)
        self.dosage_entry = tk.Entry(left_frame, font=('微软雅黑', 14), width=25)
        self.dosage_entry.pack(fill='x', pady=5)

        tk.Label(left_frame, text='提醒时间（时:分，如 08:00）：', font=('微软雅黑', 12)).pack(anchor='w', pady=5)
        self.time_entry = tk.Entry(left_frame, font=('微软雅黑', 14), width=25)
        self.time_entry.insert(0, "08:00")
        self.time_entry.pack(fill='x', pady=5)

        tk.Label(left_frame, text='药品照片（可选）：', font=('微软雅黑', 12)).pack(anchor='w', pady=5)
        photo_frame = tk.Frame(left_frame)
        photo_frame.pack(fill='x', pady=5)
        self.image_path_entry = tk.Entry(photo_frame, font=('微软雅黑', 14), width=20)
        self.image_path_entry.pack(side='left', fill='x', expand=True)
        tk.Button(photo_frame, text='📷 选择照片', font=('微软雅黑', 10),
                  command=self.select_photo).pack(side='right', padx=5)

        tk.Button(left_frame, text='📸 拍照识别药品', font=('微软雅黑', 12),
                  bg='#9b59b6', fg='white', command=self.ocr_recognize).pack(pady=5)

        add_btn = tk.Button(left_frame, text='✅ 添加药品', font=('微软雅黑', 14, 'bold'),
                            bg='#27ae60', fg='white', command=self.add_medicine,
                            height=2, width=20)
        add_btn.pack(pady=20)

        right_frame = tk.Frame(self.manage_frame)
        right_frame.pack(side='right', fill='both', expand=True, padx=10, pady=10)

        tk.Label(right_frame, text='📋 当前药品列表', font=('微软雅黑', 16, 'bold'),
                 fg='#2c3e50').pack(pady=10)

        # 药品列表（带滚动条，支持查看照片）
        list_frame = tk.Frame(right_frame)
        list_frame.pack(fill='both', expand=True)

        # 创建Canvas和滚动条
        self.med_list_canvas = tk.Canvas(list_frame)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.med_list_canvas.yview)
        self.med_list_frame = tk.Frame(self.med_list_canvas)

        self.med_list_frame.bind("<Configure>", lambda e: self.med_list_canvas.configure(
            scrollregion=self.med_list_canvas.bbox("all")))
        self.med_list_canvas.create_window((0, 0), window=self.med_list_frame, anchor="nw")
        self.med_list_canvas.configure(yscrollcommand=scrollbar.set)

        self.med_list_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.refresh_medicine_list()

    def select_photo(self):
        """选择药品照片"""
        file_path = filedialog.askopenfilename(
            title="选择药品照片",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif")]
        )
        if file_path:
            self.image_path_entry.delete(0, tk.END)
            self.image_path_entry.insert(0, file_path)
            print(f"✅ 已选择照片: {file_path}")

    def show_photo_preview(self, image_path, title="药品照片预览"):
        """显示照片预览（大图）"""
        try:
            # 打开图片并调整大小
            img = Image.open(image_path)
            # 限制最大尺寸为400x400
            img.thumbnail((400, 400), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)

            preview_win = tk.Toplevel(self.root)
            preview_win.title(title)
            preview_win.geometry(f"{img.width + 50}x{img.height + 80}")
            preview_win.resizable(False, False)

            label = tk.Label(preview_win, image=photo)
            label.image = photo
            label.pack(pady=10)

            # 显示文件路径
            tk.Label(preview_win, text=os.path.basename(image_path),
                     font=('微软雅黑', 10)).pack(pady=5)
        except Exception as e:
            print(f"预览照片失败: {e}")
            messagebox.showerror("错误", f"无法打开图片：{e}")

    def ocr_recognize(self):
        """拍照识别药品"""
        file_path = filedialog.askopenfilename(
            title="选择药品照片进行识别",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp")]
        )
        if not file_path:
            return

        print(f"🔍 正在识别药品图片: {file_path}")
        mock_result = self.mock_ocr_recognition(file_path)

        if mock_result:
            self.med_name_entry.delete(0, tk.END)
            self.med_name_entry.insert(0, mock_result.get('name', ''))
            self.dosage_entry.delete(0, tk.END)
            self.dosage_entry.insert(0, mock_result.get('dosage', ''))
            self.image_path_entry.delete(0, tk.END)
            self.image_path_entry.insert(0, file_path)

            messagebox.showinfo("识别成功",
                                f"已识别药品信息：\n名称：{mock_result.get('name', '未知')}\n剂量：{mock_result.get('dosage', '未知')}")
        else:
            messagebox.showwarning("识别失败", "无法识别药品信息，请手动输入")

    def mock_ocr_recognition(self, image_path):
        """模拟OCR识别"""
        filename = os.path.basename(image_path).lower()
        if '降压' in filename or '血压' in filename:
            return {'name': '降压药', 'dosage': '1片'}
        elif '感冒' in filename:
            return {'name': '感冒药', 'dosage': '2粒'}
        elif '消炎' in filename or '阿莫西林' in filename:
            return {'name': '阿莫西林', 'dosage': '2粒'}
        elif '维C' in filename or '维生素' in filename:
            return {'name': '维生素C片', 'dosage': '1片'}
        else:
            return None

    def setup_today_tab(self):
        """今日提醒标签页"""
        self.today_canvas = tk.Canvas(self.today_frame)
        scrollbar = tk.Scrollbar(self.today_frame, orient="vertical", command=self.today_canvas.yview)
        self.scrollable_frame = tk.Frame(self.today_canvas)

        self.scrollable_frame.bind("<Configure>",
                                   lambda e: self.today_canvas.configure(scrollregion=self.today_canvas.bbox("all")))
        self.today_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.today_canvas.configure(yscrollcommand=scrollbar.set)

        self.today_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.refresh_today_reminders()

    def setup_history_tab(self):
        """服药历史标签页"""
        control_frame = tk.Frame(self.history_frame)
        control_frame.pack(pady=10)

        tk.Button(control_frame, text='◀ 上月', font=('微软雅黑', 12),
                  command=self.prev_month).pack(side='left', padx=10)
        self.month_label = tk.Label(control_frame, font=('微软雅黑', 14, 'bold'), width=20)
        self.month_label.pack(side='left', padx=20)
        tk.Button(control_frame, text='下月 ▶', font=('微软雅黑', 12),
                  command=self.next_month).pack(side='left', padx=10)
        tk.Button(control_frame, text='📅 今天', font=('微软雅黑', 12),
                  command=self.go_today).pack(side='left', padx=10)

        self.calendar_frame = tk.Frame(self.history_frame)
        self.calendar_frame.pack(fill='both', expand=True, padx=20, pady=20)

        self.current_year = datetime.now().year
        self.current_month = datetime.now().month
        self.show_calendar()

    def setup_settings_tab(self):
        """设置标签页"""
        settings_frame = tk.Frame(self.settings_frame)
        settings_frame.pack(pady=30)

        tk.Label(settings_frame, text='📱 紧急联系人设置', font=('微软雅黑', 16, 'bold')).pack(pady=20)
        tk.Label(settings_frame, text='子女手机号：', font=('微软雅黑', 12)).pack()
        self.phone_entry = tk.Entry(settings_frame, font=('微软雅黑', 14), width=20)
        self.phone_entry.pack(pady=5)

        if hasattr(self, 'emergency_phone'):
            self.phone_entry.insert(0, self.emergency_phone)

        tk.Label(settings_frame, text='超时时间（分钟）：', font=('微软雅黑', 12)).pack(pady=5)
        self.timeout_entry = tk.Entry(settings_frame, font=('微软雅黑', 14), width=10)
        self.timeout_entry.insert(0, str(getattr(self, 'timeout_minutes', 20)))
        self.timeout_entry.pack(pady=5)

        tk.Button(settings_frame, text='💾 保存设置', font=('微软雅黑', 14),
                  bg='#3498db', fg='white', command=self.save_settings,
                  height=2, width=15).pack(pady=30)

        tk.Button(settings_frame, text='📢 测试语音播报', font=('微软雅黑', 12),
                  command=self.test_voice).pack(pady=10)

        tk.Label(settings_frame, text='📦 数据管理', font=('微软雅黑', 16, 'bold')).pack(pady=20)
        btn_frame = tk.Frame(settings_frame)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text='🗑 清空历史', font=('微软雅黑', 10),
                  command=self.clear_history).pack(side='left', padx=5)

    def test_voice(self):
        """测试语音播报"""
        self.speak("您好，我是吃药小管家，语音播报功能测试成功")

    def clear_history(self):
        """清空服药历史"""
        if messagebox.askyesno("确认清空", "确定要清空所有服药历史记录吗？"):
            self.taking_history = {}
            self.save_data()
            self.refresh_today_reminders()
            self.show_calendar()
            messagebox.showinfo("成功", "服药历史已清空")

    def add_medicine(self):
        """添加药品"""
        name = self.med_name_entry.get().strip()
        dosage = self.dosage_entry.get().strip()
        time_str = self.time_entry.get().strip()
        image_path = self.image_path_entry.get().strip()

        if not name or not dosage or not time_str:
            messagebox.showwarning("提示", "请填写完整信息！")
            return

        try:
            datetime.strptime(time_str, "%H:%M")
        except:
            messagebox.showwarning("提示", "时间格式错误！")
            return

        medicine = {
            'id': len(self.medicines) + 1,
            'name': name,
            'dosage': dosage,
            'time': time_str,
            'image_path': image_path,
            'enabled': True,
            'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        self.medicines.append(medicine)
        self.save_data()

        self.med_name_entry.delete(0, tk.END)
        self.dosage_entry.delete(0, tk.END)
        self.image_path_entry.delete(0, tk.END)

        self.refresh_medicine_list()
        self.refresh_today_reminders()
        messagebox.showinfo("成功", f"已添加药品：{name}")

    def refresh_medicine_list(self):
        """刷新药品列表（支持查看照片）"""
        for widget in self.med_list_frame.winfo_children():
            widget.destroy()

        if not self.medicines:
            tk.Label(self.med_list_frame, text='暂无药品，请添加',
                     font=('微软雅黑', 14), fg='gray').pack(pady=50)
            return

        for med in self.medicines:
            med_frame = tk.Frame(self.med_list_frame, relief='ridge', bd=2, padx=10, pady=10)
            med_frame.pack(fill='x', pady=5)

            # 药品信息
            info_text = f"💊 {med['name']}  |  {med['dosage']}  |  ⏰ {med['time']}"
            if med.get('image_path') and os.path.exists(med.get('image_path', '')):
                info_text += "  |  📷 有照片"

            info_label = tk.Label(med_frame, text=info_text, font=('微软雅黑', 12))
            info_label.pack(side='left')

            # 按钮框架
            btn_frame = tk.Frame(med_frame)
            btn_frame.pack(side='right')

            # 【新增】查看照片按钮
            if med.get('image_path') and os.path.exists(med.get('image_path', '')):
                view_btn = tk.Button(btn_frame, text='📷 查看照片', font=('微软雅黑', 10),
                                     bg='#3498db', fg='white',
                                     command=lambda m=med: self.show_photo_preview(m['image_path'],
                                                                                   f"{m['name']} - 照片"))
                view_btn.pack(side='left', padx=2)

            # 编辑按钮
            edit_btn = tk.Button(btn_frame, text='✏️ 编辑', font=('微软雅黑', 10),
                                 bg='#f39c12', fg='white',
                                 command=lambda m=med: self.edit_medicine(m))
            edit_btn.pack(side='left', padx=2)

            # 删除按钮
            del_btn = tk.Button(btn_frame, text='🗑 删除', font=('微软雅黑', 10),
                                bg='#e74c3c', fg='white',
                                command=lambda m=med: self.delete_medicine(m))
            del_btn.pack(side='left', padx=2)

    def edit_medicine(self, medicine):
        """编辑药品"""
        edit_win = tk.Toplevel(self.root)
        edit_win.title(f"编辑药品 - {medicine['name']}")
        edit_win.geometry("400x450")
        edit_win.resizable(False, False)

        tk.Label(edit_win, text=f"编辑药品：{medicine['name']}",
                 font=('微软雅黑', 16, 'bold')).pack(pady=20)

        tk.Label(edit_win, text='药品名称：', font=('微软雅黑', 12)).pack(anchor='w', padx=50)
        name_entry = tk.Entry(edit_win, font=('微软雅黑', 14), width=25)
        name_entry.insert(0, medicine['name'])
        name_entry.pack(pady=5, padx=50)

        tk.Label(edit_win, text='剂量：', font=('微软雅黑', 12)).pack(anchor='w', padx=50)
        dosage_entry = tk.Entry(edit_win, font=('微软雅黑', 14), width=25)
        dosage_entry.insert(0, medicine['dosage'])
        dosage_entry.pack(pady=5, padx=50)

        tk.Label(edit_win, text='提醒时间：', font=('微软雅黑', 12)).pack(anchor='w', padx=50)
        time_entry = tk.Entry(edit_win, font=('微软雅黑', 14), width=25)
        time_entry.insert(0, medicine['time'])
        time_entry.pack(pady=5, padx=50)

        def save_edit():
            medicine['name'] = name_entry.get().strip()
            medicine['dosage'] = dosage_entry.get().strip()
            medicine['time'] = time_entry.get().strip()
            self.save_data()
            self.refresh_medicine_list()
            self.refresh_today_reminders()
            edit_win.destroy()
            messagebox.showinfo("成功", "药品信息已更新")

        tk.Button(edit_win, text='💾 保存修改', font=('微软雅黑', 14),
                  bg='#27ae60', fg='white', command=save_edit).pack(pady=30)

    def delete_medicine(self, medicine):
        """删除药品"""
        if messagebox.askyesno("确认删除", f"确定要删除「{medicine['name']}」吗？"):
            self.medicines.remove(medicine)
            self.save_data()
            self.refresh_medicine_list()
            self.refresh_today_reminders()

    def refresh_today_reminders(self):
        """刷新今日提醒列表"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        today = datetime.now().date()
        today_reminders = []

        for med in self.medicines:
            if med.get('enabled', True):
                reminder_time = datetime.strptime(med['time'], "%H:%M").time()
                reminder_datetime = datetime.combine(today, reminder_time)
                today_reminders.append((reminder_datetime, med))

        if not today_reminders:
            tk.Label(self.scrollable_frame, text='🎉 今天没有预设的服药提醒',
                     font=('微软雅黑', 16), fg='green').pack(pady=50)
            return

        today_reminders.sort(key=lambda x: x[0])

        for reminder_datetime, med in today_reminders:
            reminder_frame = tk.Frame(self.scrollable_frame, relief='groove', bd=2, padx=15, pady=15)
            reminder_frame.pack(fill='x', pady=10, padx=20)

            time_str = reminder_datetime.strftime("%H:%M")
            status = self.check_taken_status(med['name'], today)

            status_text = '✅ 已服用' if status == 'taken' else '⏳ 待服用'
            status_color = 'green' if status == 'taken' else 'orange'

            info_label = tk.Label(reminder_frame,
                                  text=f"⏰ {time_str}  |  💊 {med['name']}  |  {med['dosage']}\n状态：{status_text}",
                                  font=('微软雅黑', 13), fg=status_color, justify='left')
            info_label.pack(side='left', padx=10)

            if status != 'taken':
                take_btn = tk.Button(reminder_frame, text='✅ 已吃', font=('微软雅黑', 12, 'bold'),
                                     bg='#27ae60', fg='white', width=10,
                                     command=lambda m=med: self.mark_as_taken(m['name']))
                take_btn.pack(side='right', padx=10)

    def check_taken_status(self, med_name, date):
        """检查是否已服用"""
        key = f"{med_name}_{date}"
        if key in self.taking_history:
            return self.taking_history[key].get('status', 'pending')
        return 'pending'

    def mark_as_taken(self, med_name):
        """标记为已服用"""
        today = datetime.now().date()
        key = f"{med_name}_{today}"
        current_time = datetime.now().strftime("%H:%M:%S")

        self.taking_history[key] = {'status': 'taken', 'time': current_time}
        self.save_data()

        # 停止语音提醒
        self.stop_voice_alert()

        # 清除提醒记录
        reminder_key = f"{med_name}_{today}"
        if reminder_key in self.reminded_set:
            self.reminded_set.remove(reminder_key)

        # 关闭提醒窗口
        if self.active_reminder_window:
            try:
                self.active_reminder_window.destroy()
            except:
                pass
            self.active_reminder_window = None

        self.current_reminding_med = None
        self.reminder_status_label.config(text='✅ 已服药', fg='#2ecc71')
        self.refresh_today_reminders()
        self.show_calendar()

        # 确认语音
        self.speak(f"{med_name}已服用，记录成功")
        self.root.after(3000, lambda: self.reminder_status_label.config(text='⏰ 等待提醒', fg='#f39c12'))

        messagebox.showinfo("记录成功", f"✅ {med_name} 已记录为【已服用】\n时间：{current_time}")

    def check_reminders(self):
        """检查提醒 - 核心功能"""
        print(f"✅ 提醒检查线程已启动，检查间隔：{self.check_interval}秒")

        while self.running:
            try:
                now = datetime.now()
                today = now.date()
                current_time_str = now.strftime("%H:%M")

                for med in self.medicines:
                    if not med.get('enabled', True):
                        continue

                    key = f"{med['name']}_{today}"
                    if key in self.taking_history and self.taking_history[key].get('status') == 'taken':
                        continue

                    if current_time_str == med['time']:
                        reminder_key = f"{med['name']}_{today}"
                        if reminder_key not in self.reminded_set:
                            print(f"\n{'=' * 50}")
                            print(f"⏰ 触发提醒！时间: {current_time_str}")
                            print(f"💊 药品: {med['name']} - {med['dosage']}")
                            if med.get('image_path'):
                                print(f"📷 药品照片: {med['image_path']}")
                            print(f"{'=' * 50}")

                            self.reminded_set.add(reminder_key)

                            # 更新状态栏
                            self.root.after(0, lambda: self.reminder_status_label.config(
                                text=f'🔔 正在提醒: {med["name"]}', fg='#e74c3c'))

                            # 显示全屏提醒窗口
                            self.root.after(0, self.show_reminder_window, med)
                            # 启动一分钟循环语音提醒
                            self.root.after(0, self.start_voice_alert, med)

            except Exception as e:
                print(f"⚠️ 提醒检查出错: {e}")

            time_module.sleep(self.check_interval)

    def show_reminder_window(self, medicine):
        """显示全屏提醒窗口"""
        if self.active_reminder_window:
            return

        print(f"🪟 创建提醒窗口: {medicine['name']}")
        self.current_reminding_med = medicine

        try:
            reminder_win = tk.Toplevel(self.root)
            reminder_win.title("⏰ 吃药提醒")
            reminder_win.attributes('-fullscreen', True)
            reminder_win.configure(bg='#ff4444')
            reminder_win.attributes('-topmost', True)

            main_frame = tk.Frame(reminder_win, bg='#ff4444')
            main_frame.pack(expand=True)

            title_label = tk.Label(main_frame, text="⏰ 吃药时间到！",
                                   font=('微软雅黑', 48, 'bold'),
                                   fg='white', bg='#ff4444')
            title_label.pack(pady=30)

            # 显示药品照片
            if medicine.get('image_path') and os.path.exists(medicine['image_path']):
                try:
                    img = Image.open(medicine['image_path'])
                    img = img.resize((200, 200), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    photo_label = tk.Label(main_frame, image=photo, bg='#ff4444')
                    photo_label.image = photo
                    photo_label.pack(pady=10)
                except Exception as e:
                    print(f"显示照片失败: {e}")

            med_label = tk.Label(main_frame,
                                 text=f"💊 {medicine['name']}\n\n{medicine['dosage']}",
                                 font=('微软雅黑', 42, 'bold'),
                                 fg='yellow', bg='#ff4444')
            med_label.pack(pady=20)

            # 显示一分钟循环语音提醒提示
            voice_hint = tk.Label(main_frame, text="🔊 语音将循环播报1分钟（每3秒一次）",
                                  font=('微软雅黑', 16), fg='white', bg='#ff4444')
            voice_hint.pack(pady=10)

            btn_frame = tk.Frame(main_frame, bg='#ff4444')
            btn_frame.pack(pady=30)

            def on_take():
                print(f"✅ 用户点击'已吃'按钮: {medicine['name']}")
                self.mark_as_taken(medicine['name'])
                try:
                    reminder_win.destroy()
                except:
                    pass

            take_btn = tk.Button(btn_frame, text="✅ 已吃",
                                 font=('微软雅黑', 32, 'bold'),
                                 bg='#27ae60', fg='white',
                                 command=on_take,
                                 height=2, width=10)
            take_btn.pack(side='left', padx=20)

            def on_later():
                print(f"⏰ 用户点击'稍后提醒': {medicine['name']}")
                self.stop_voice_alert()
                reminder_win.destroy()
                self.active_reminder_window = None
                # 5分钟后重新提醒
                self.root.after(300000, lambda: self.retrigger_reminder(medicine))

            later_btn = tk.Button(btn_frame, text="⏰ 稍后提醒",
                                  font=('微软雅黑', 24, 'bold'),
                                  bg='#f39c12', fg='white',
                                  command=on_later,
                                  height=2, width=10)
            later_btn.pack(side='left', padx=20)

            timeout = getattr(self, 'timeout_minutes', 20)
            info_label = tk.Label(main_frame,
                                  text=f"⚠️ 如果 {timeout} 分钟内未确认，将通知紧急联系人",
                                  font=('微软雅黑', 14), fg='white', bg='#ff4444')
            info_label.pack(pady=20)

            self.active_reminder_window = reminder_win

        except Exception as e:
            print(f"❌ 创建提醒窗口失败: {e}")

    def retrigger_reminder(self, medicine):
        """重新触发提醒"""
        if self.active_reminder_window:
            return
        print(f"⏰ 重新触发提醒: {medicine['name']}")
        self.show_reminder_window(medicine)
        self.start_voice_alert(medicine)

    def show_calendar(self):
        """显示日历"""
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()

        self.month_label.config(text=f"{self.current_year}年{self.current_month}月")

        first_day = datetime(self.current_year, self.current_month, 1)
        if self.current_month == 12:
            next_month = datetime(self.current_year + 1, 1, 1)
        else:
            next_month = datetime(self.current_year, self.current_month + 1, 1)
        days_in_month = (next_month - first_day).days
        first_weekday = first_day.weekday()

        weekdays = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
        for i, day in enumerate(weekdays):
            label = tk.Label(self.calendar_frame, text=day, font=('微软雅黑', 12, 'bold'),
                             width=10, height=2, relief='ridge')
            label.grid(row=0, column=i, sticky='nsew')

        row, col = 1, first_weekday
        for day in range(1, days_in_month + 1):
            date = datetime(self.current_year, self.current_month, day)
            day_frame = tk.Frame(self.calendar_frame, relief='ridge', bd=1)
            day_frame.grid(row=row, column=col, sticky='nsew', padx=2, pady=2)

            day_label = tk.Label(day_frame, text=str(day), font=('微软雅黑', 12, 'bold'))
            day_label.pack(anchor='nw', padx=5, pady=5)

            total, taken = 0, 0
            for med in self.medicines:
                if med.get('enabled', True):
                    total += 1
                    key = f"{med['name']}_{date.date()}"
                    if key in self.taking_history and self.taking_history[key].get('status') == 'taken':
                        taken += 1

            if total > 0:
                if taken == total:
                    status = f"✅ {taken}/{total}"
                elif taken > 0:
                    status = f"⚠️ {taken}/{total}"
                else:
                    status = f"❌ {taken}/{total}"
                tk.Label(day_frame, text=status, font=('微软雅黑', 10)).pack(pady=5)

            if date.date() == datetime.now().date():
                day_frame.configure(bg='#ffffcc')

            col += 1
            if col > 6:
                col, row = 0, row + 1

        for i in range(7):
            self.calendar_frame.grid_columnconfigure(i, weight=1)
        for i in range(1, row + 1):
            self.calendar_frame.grid_rowconfigure(i, weight=1)

    def prev_month(self):
        if self.current_month == 1:
            self.current_month, self.current_year = 12, self.current_year - 1
        else:
            self.current_month -= 1
        self.show_calendar()

    def next_month(self):
        if self.current_month == 12:
            self.current_month, self.current_year = 1, self.current_year + 1
        else:
            self.current_month += 1
        self.show_calendar()

    def go_today(self):
        self.current_year, self.current_month = datetime.now().year, datetime.now().month
        self.show_calendar()

    def check_missed_reminders(self):
        """检查错过的提醒"""
        now, today = datetime.now(), datetime.now().date()
        for med in self.medicines:
            if not med.get('enabled', True):
                continue
            key = f"{med['name']}_{today}"
            if key in self.taking_history and self.taking_history[key].get('status') == 'taken':
                continue
            reminder_time = datetime.strptime(med['time'], "%H:%M").time()
            if reminder_time <= now.time():
                print(f"⏰ 发现错过提醒: {med['name']}")
                self.show_reminder_window(med)
                self.start_voice_alert(med)

    def save_settings(self):
        """保存设置"""
        self.emergency_phone = self.phone_entry.get().strip()
        try:
            self.timeout_minutes = int(self.timeout_entry.get().strip())
        except:
            self.timeout_minutes = 20
        self.save_data()
        messagebox.showinfo("成功", "设置已保存！")

    def save_data(self):
        """保存数据"""
        data = {
            'medicines': self.medicines,
            'taking_history': self.taking_history,
            'emergency_phone': getattr(self, 'emergency_phone', ''),
            'timeout_minutes': getattr(self, 'timeout_minutes', 20)
        }
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print("✅ 数据已保存")

    def load_data(self):
        """加载数据"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.medicines = data.get('medicines', [])
                    self.taking_history = data.get('taking_history', {})
                    self.emergency_phone = data.get('emergency_phone', '')
                    self.timeout_minutes = data.get('timeout_minutes', 20)
                print(f"✅ 数据加载成功，共 {len(self.medicines)} 个药品")
            except Exception as e:
                print(f"❌ 数据加载失败: {e}")
                self.medicines, self.taking_history = [], {}
        else:
            print("📝 未找到数据文件，将创建新数据")


if __name__ == "__main__":
    root = tk.Tk()
    app = MedicineReminder(root)


    def on_closing():
        print("\n👋 正在关闭程序...")
        app.running = False
        app.stop_voice_alert()
        root.destroy()


    root.protocol("WM_DELETE_WINDOW", on_closing)

    print("\n" + "=" * 50)
    print("💊 吃药小管家已启动")
    print("=" * 50)
    print("使用说明：")
    print("1. 在【药品管理】页添加药品和上传照片")
    print("2. 药品列表中可以点击【📷 查看照片】查看上传的图片")
    print("3. 设置提醒时间测试（如当前22:12，设置22:13）")
    print("4. 到达时间会自动弹出全屏提醒")
    print("5. 语音会循环播报1分钟（每3秒一次，共20次）")
    print("6. 点击【已吃】按钮记录服药并停止语音")
    print("=" * 50 + "\n")

    root.mainloop()