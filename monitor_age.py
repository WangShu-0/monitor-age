import os
import json
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext

CONFIG_FILE = "monitor_config.json"
REMINDER_RECORD_FILE = "reminded.json"

class AgeMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("年龄预警监控器")
        self.root.geometry("760x650")
        self.root.minsize(680, 600)

        self.monitoring = False
        self.after_id = None
        self.available_sheets = []

        self.config = self.load_config()
        self.excel_path = self.config.get("excel_path", "")
        self.sheet_mode = self.config.get("sheet_mode", "all")
        self.sheet_name = self.config.get("sheet_name", "")
        self.header_row = self.config.get("header_row", 0)
        self.target_age = self.config.get("target_age", 80)
        self.cond_type = self.config.get("cond_type", "本月达到年龄")
        self.deadline_str = self.config.get("deadline", "")
        if self.deadline_str:
            try:
                self.deadline_date = datetime.strptime(self.deadline_str, "%Y-%m-%d").date()
            except:
                self.deadline_date = datetime.now().date() + timedelta(days=30)
        else:
            self.deadline_date = datetime.now().date() + timedelta(days=30)
            self.deadline_str = self.deadline_date.strftime("%Y-%m-%d")

        self.hour = self.config.get("hour", 9)
        self.minute = self.config.get("minute", 0)

        self.reminded_set = self.load_reminded_set()
        self.reminded_data = self.load_reminded_data()

        self.create_widgets()
        self.load_settings_to_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ---------- 配置保存/加载 ----------
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_config(self):
        config = {
            "excel_path": self.excel_path,
            "sheet_mode": self.sheet_mode,
            "sheet_name": self.sheet_name,
            "header_row": self.header_row,
            "target_age": self.target_age,
            "cond_type": self.cond_type,
            "deadline": self.deadline_str if self.cond_type == "截止日期前达到年龄" else "",
            "hour": self.hour,
            "minute": self.minute,
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)

    def load_reminded_set(self):
        if os.path.exists(REMINDER_RECORD_FILE):
            try:
                with open(REMINDER_RECORD_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                        return {item["unique_id"] for item in data}
                    else:
                        return set(str(item) for item in data)
            except:
                return set()
        return set()

    def load_reminded_data(self):
        if os.path.exists(REMINDER_RECORD_FILE):
            try:
                with open(REMINDER_RECORD_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                        return data
                    else:
                        return []
            except:
                return []
        return []

    def save_reminded_data(self):
        with open(REMINDER_RECORD_FILE, "w", encoding="utf-8") as f:
            json.dump(self.reminded_data, f, ensure_ascii=False, indent=2, default=str)

    # ---------- 界面构建 ----------
    def create_widgets(self):
        self.sheet_mode_var = tk.StringVar(value=self.sheet_mode)
        self.sheet_name_var = tk.StringVar(value=self.sheet_name)
        self.header_row_var = tk.IntVar(value=self.header_row)

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', font=('微软雅黑', 9))
        style.configure('TButton', font=('微软雅黑', 9))

        main_frame = ttk.Frame(self.root, padding="8")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择
        file_frame = ttk.LabelFrame(main_frame, text="Excel文件", padding="5")
        file_frame.pack(fill=tk.X, pady=(0,8))

        row_file = ttk.Frame(file_frame)
        row_file.pack(fill=tk.X, pady=2)
        self.file_path_var = tk.StringVar(value=self.excel_path)
        entry_file = ttk.Entry(row_file, textvariable=self.file_path_var)
        entry_file.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        btn_browse = ttk.Button(row_file, text="浏览...", command=self.select_excel_file)
        btn_browse.pack(side=tk.RIGHT)

        # 工作表选择
        row_sheet = ttk.Frame(file_frame)
        row_sheet.pack(fill=tk.X, pady=2)
        rb_all = ttk.Radiobutton(row_sheet, text="全部工作表", variable=self.sheet_mode_var, value="all",
                                 command=self.on_sheet_mode_changed)
        rb_all.pack(side=tk.LEFT, padx=(0,10))
        rb_single = ttk.Radiobutton(row_sheet, text="指定工作表:", variable=self.sheet_mode_var, value="single",
                                    command=self.on_sheet_mode_changed)
        rb_single.pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(row_sheet, textvariable=self.sheet_name_var, width=20, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        btn_refresh = ttk.Button(row_sheet, text="刷新列表", command=self.refresh_sheet_list)
        btn_refresh.pack(side=tk.LEFT)

        # 表头行设置
        header_frame = ttk.Frame(file_frame)
        header_frame.pack(fill=tk.X, pady=2)
        ttk.Label(header_frame, text="表头行号(0表示第一行):").pack(side=tk.LEFT)
        header_spin = ttk.Spinbox(header_frame, from_=0, to=100, width=5, textvariable=self.header_row_var)
        header_spin.pack(side=tk.LEFT, padx=5)
        btn_auto_header = ttk.Button(header_frame, text="自动查找表头行", command=self.auto_find_header_row)
        btn_auto_header.pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="(表头行以下为正式数据)", foreground="gray").pack(side=tk.LEFT, padx=5)

        # 监控参数
        param_frame = ttk.LabelFrame(main_frame, text="监控参数", padding="5")
        param_frame.pack(fill=tk.X, pady=(0,8))

        row1 = ttk.Frame(param_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="目标年龄:").pack(side=tk.LEFT)
        self.age_var = tk.IntVar(value=self.target_age)
        age_spin = ttk.Spinbox(row1, from_=1, to=120, textvariable=self.age_var, width=6)
        age_spin.pack(side=tk.LEFT, padx=5)
        ttk.Label(row1, text="监测条件:").pack(side=tk.LEFT, padx=(15,0))
        self.cond_var = tk.StringVar(value=self.cond_type)
        rb_cond1 = ttk.Radiobutton(row1, text="本月达到年龄", variable=self.cond_var, value="本月达到年龄",
                                   command=self.on_cond_changed)
        rb_cond1.pack(side=tk.LEFT, padx=2)
        rb_cond2 = ttk.Radiobutton(row1, text="下月达到年龄", variable=self.cond_var, value="下月达到年龄",
                                   command=self.on_cond_changed)
        rb_cond2.pack(side=tk.LEFT, padx=2)
        rb_cond3 = ttk.Radiobutton(row1, text="本年度达到年龄", variable=self.cond_var, value="本年度达到年龄",
                                   command=self.on_cond_changed)
        rb_cond3.pack(side=tk.LEFT, padx=2)
        rb_cond4 = ttk.Radiobutton(row1, text="截止日期前达到年龄", variable=self.cond_var, value="截止日期前达到年龄",
                                   command=self.on_cond_changed)
        rb_cond4.pack(side=tk.LEFT, padx=2)

        # 截止日期输入（年/月/日下拉框 + 应用按钮）
        row2 = ttk.Frame(param_frame)
        row2.pack(fill=tk.X, pady=2)
        self.deadline_frame = ttk.Frame(row2)
        self.deadline_frame.pack(side=tk.LEFT)

        ttk.Label(self.deadline_frame, text="截止日期:").pack(side=tk.LEFT)

        current_year = datetime.now().year
        self.year_var = tk.IntVar(value=self.deadline_date.year)
        self.year_spin = ttk.Spinbox(self.deadline_frame, from_=1900, to=current_year+10,
                                     textvariable=self.year_var, width=6)
        self.year_spin.pack(side=tk.LEFT, padx=2)
        self.year_spin.bind('<FocusOut>', lambda e: self.on_date_changed())
        self.year_spin.bind('<Return>', lambda e: self.on_date_changed())
        ttk.Label(self.deadline_frame, text="年").pack(side=tk.LEFT)

        self.month_var = tk.IntVar(value=self.deadline_date.month)
        self.month_spin = ttk.Spinbox(self.deadline_frame, from_=1, to=12,
                                      textvariable=self.month_var, width=4)
        self.month_spin.pack(side=tk.LEFT, padx=2)
        self.month_spin.bind('<FocusOut>', lambda e: self.on_date_changed())
        self.month_spin.bind('<Return>', lambda e: self.on_date_changed())
        ttk.Label(self.deadline_frame, text="月").pack(side=tk.LEFT)

        self.day_var = tk.IntVar(value=self.deadline_date.day)
        self.day_spin = ttk.Spinbox(self.deadline_frame, from_=1, to=31,
                                    textvariable=self.day_var, width=4)
        self.day_spin.pack(side=tk.LEFT, padx=2)
        self.day_spin.bind('<FocusOut>', lambda e: self.on_date_changed())
        self.day_spin.bind('<Return>', lambda e: self.on_date_changed())
        ttk.Label(self.deadline_frame, text="日").pack(side=tk.LEFT)

        # 应用按钮
        btn_apply = ttk.Button(self.deadline_frame, text="应用", width=4, command=self.on_date_changed)
        btn_apply.pack(side=tk.LEFT, padx=5)

        ttk.Label(self.deadline_frame, text="(满足: 达到年龄日期 ≤ 截止日期)", foreground="gray").pack(side=tk.LEFT, padx=5)

        # 每日定时
        row3 = ttk.Frame(param_frame)
        row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="每日定时:").pack(side=tk.LEFT)
        self.hour_var = tk.IntVar(value=self.hour)
        hour_spin = ttk.Spinbox(row3, from_=0, to=23, width=3, textvariable=self.hour_var)
        hour_spin.pack(side=tk.LEFT)
        ttk.Label(row3, text=":").pack(side=tk.LEFT)
        self.minute_var = tk.IntVar(value=self.minute)
        minute_spin = ttk.Spinbox(row3, from_=0, to=59, width=3, textvariable=self.minute_var)
        minute_spin.pack(side=tk.LEFT)

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        self.btn_start = ttk.Button(btn_frame, text="开始监控", command=self.start_monitor)
        self.btn_start.pack(side=tk.LEFT, padx=3)
        self.btn_stop = ttk.Button(btn_frame, text="停止监控", command=self.stop_monitor, state="disabled")
        self.btn_stop.pack(side=tk.LEFT, padx=3)
        self.btn_manual = ttk.Button(btn_frame, text="手动监测", command=self.manual_check)
        self.btn_manual.pack(side=tk.LEFT, padx=3)
        self.btn_show_records = ttk.Button(btn_frame, text="显示提醒", command=self.show_reminded_records)
        self.btn_show_records.pack(side=tk.LEFT, padx=3)
        self.btn_clear = ttk.Button(btn_frame, text="清除提醒记录", command=self.clear_reminded)
        self.btn_clear.pack(side=tk.LEFT, padx=3)

        # 状态栏
        self.status_var = tk.StringVar(value="状态：未监控")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(5,0))

        # 日志
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(8,0))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, font=('Consolas', 9), wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.update_deadline_state()
        self.on_sheet_mode_changed()
        self.refresh_sheet_list()

    # ---------- 辅助方法 ----------
    def update_day_range(self):
        year = self.year_var.get()
        month = self.month_var.get()
        if month in (1, 3, 5, 7, 8, 10, 12):
            max_day = 31
        elif month in (4, 6, 9, 11):
            max_day = 30
        else:
            if (year % 400 == 0) or (year % 4 == 0 and year % 100 != 0):
                max_day = 29
            else:
                max_day = 28
        self.day_spin.config(to=max_day)
        if self.day_var.get() > max_day:
            self.day_var.set(max_day)

    def on_date_changed(self, event=None):
        self.update_day_range()
        year = self.year_var.get()
        month = self.month_var.get()
        day = self.day_var.get()
        try:
            new_date = datetime(year, month, day).date()
            if new_date != self.deadline_date:
                self.deadline_date = new_date
                self.deadline_str = self.deadline_date.strftime("%Y-%m-%d")
                self.log(f"截止日期已更改为: {self.deadline_str}")
        except ValueError as e:
            self.log(f"截止日期无效: {year}-{month}-{day}")

    def update_deadline_state(self):
        if self.cond_var.get() == "截止日期前达到年龄":
            self.deadline_frame.pack(side=tk.LEFT)
            for child in self.deadline_frame.winfo_children():
                if isinstance(child, ttk.Spinbox):
                    child.config(state="normal")
            for child in self.deadline_frame.winfo_children():
                if isinstance(child, ttk.Button):
                    child.config(state="normal")
        else:
            self.deadline_frame.pack_forget()

    def on_cond_changed(self):
        self.update_deadline_state()
        self.log(f"监测条件已改为: {self.cond_var.get()}")

    def on_sheet_mode_changed(self):
        self.sheet_mode = self.sheet_mode_var.get()
        if self.sheet_mode == "single":
            self.sheet_combo.config(state="readonly")
        else:
            self.sheet_combo.config(state="disabled")
            self.sheet_name_var.set("")
        self.save_config()
        self.refresh_sheet_list()

    def load_settings_to_ui(self):
        self.file_path_var.set(self.excel_path)
        self.sheet_mode_var.set(self.sheet_mode)
        self.sheet_name_var.set(self.sheet_name)
        self.header_row_var.set(self.header_row)
        self.age_var.set(self.target_age)
        self.cond_var.set(self.cond_type)
        self.year_var.set(self.deadline_date.year)
        self.month_var.set(self.deadline_date.month)
        self.day_var.set(self.deadline_date.day)
        self.update_day_range()
        self.hour_var.set(self.hour)
        self.minute_var.set(self.minute)
        self.update_deadline_state()
        self.on_sheet_mode_changed()
        self.refresh_sheet_list()

    def log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    # ---------- 文件与表头处理 ----------
    def select_excel_file(self):
        filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="选择人员数据Excel文件", filetypes=filetypes)
        if filename:
            self.excel_path = filename
            self.file_path_var.set(filename)
            self.save_config()
            self.log(f"已选择Excel文件: {filename}")
            self.refresh_sheet_list()
            self.auto_find_header_row()

    def refresh_sheet_list(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            self.sheet_combo['values'] = []
            return
        try:
            xl = pd.ExcelFile(self.excel_path)
            self.available_sheets = xl.sheet_names
            self.sheet_combo['values'] = self.available_sheets
            if self.sheet_name in self.available_sheets:
                self.sheet_name_var.set(self.sheet_name)
            elif self.available_sheets:
                self.sheet_name_var.set(self.available_sheets[0])
        except Exception as e:
            self.log(f"读取工作表列表失败: {e}")
            self.sheet_combo['values'] = []

    def auto_find_header_row(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            messagebox.showwarning("未选择文件", "请先选择Excel文件。")
            return
        sheet = self.sheet_name_var.get() if self.sheet_mode == "single" else None
        try:
            if sheet:
                raw_df = pd.read_excel(self.excel_path, sheet_name=sheet, header=None, nrows=30)
            else:
                raw_df = pd.read_excel(self.excel_path, header=None, nrows=30)
        except Exception as e:
            self.log(f"读取前30行失败: {e}")
            messagebox.showerror("读取错误", f"无法读取Excel文件：{e}")
            return

        name_keywords = ["姓名", "名字", "名称", "name"]
        birth_keywords = ["出生日期", "生日", "出生年月", "birth", "birthday", "出生"]
        found_row = None
        for i in range(min(30, len(raw_df))):
            row_series = raw_df.iloc[i].fillna('').astype(str).str.lower()
            row_texts = row_series.tolist()
            has_name = any(any(kw in cell for kw in name_keywords) for cell in row_texts)
            has_birth = any(any(kw in cell for kw in birth_keywords) for cell in row_texts)
            if has_name and has_birth:
                found_row = i
                break
        if found_row is not None:
            self.header_row_var.set(found_row)
            self.header_row = found_row
            self.save_config()
            self.log(f"自动识别到表头在第 {found_row+1} 行")
            messagebox.showinfo("识别成功", f"表头行已自动设置为第 {found_row+1} 行。\n该行以下为正式数据。")
        else:
            self.log("未找到同时包含'姓名'和'出生日期'的行，请手动设置表头行号。")
            messagebox.showwarning("未找到", "未能自动识别表头行，请手动输入行号（0表示第一行）。")

    def read_data_with_header(self, file_path, sheet_name=None, rows=None):
        header_row = self.header_row_var.get()
        try:
            if sheet_name:
                if rows:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, nrows=rows)
                else:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
            else:
                if rows:
                    df = pd.read_excel(file_path, header=header_row, nrows=rows)
                else:
                    df = pd.read_excel(file_path, header=header_row)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            return df
        except Exception as e:
            self.log(f"读取数据失败: {e}")
            return None

    def ensure_excel_file(self):
        if not self.excel_path:
            messagebox.showwarning("未选择文件", "请先选择Excel文件。")
            return False
        if not os.path.exists(self.excel_path):
            messagebox.showerror("文件不存在", f"文件不存在：{self.excel_path}")
            return False
        return True

    def get_current_settings(self):
        self.target_age = self.age_var.get()
        self.cond_type = self.cond_var.get()
        self.hour = self.hour_var.get()
        self.minute = self.minute_var.get()
        self.sheet_mode = self.sheet_mode_var.get()
        if self.sheet_mode == "single":
            self.sheet_name = self.sheet_name_var.get()
        else:
            self.sheet_name = ""
        self.header_row = self.header_row_var.get()
        if self.cond_type == "截止日期前达到年龄":
            try:
                self.deadline_date = datetime(self.year_var.get(), self.month_var.get(), self.day_var.get()).date()
                self.deadline_str = self.deadline_date.strftime("%Y-%m-%d")
            except:
                pass
        self.save_config()

    # ---------- 核心逻辑 ----------
    def convert_birth_date(self, val):
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, (int, float)):
            s = str(int(val))
            if len(s) == 8 and s.isdigit():
                try:
                    return pd.to_datetime(s, format='%Y%m%d')
                except:
                    pass
        elif isinstance(val, str):
            s = val.strip()
            if len(s) == 8 and s.isdigit():
                try:
                    return pd.to_datetime(s, format='%Y%m%d')
                except:
                    pass
        return pd.to_datetime(val, errors='coerce')

    def get_target_date(self, birth_date, target_age):
        return birth_date + relativedelta(years=target_age)

    def compute_age(self, birth_date):
        today = datetime.now().date()
        age = today.year - birth_date.year
        if (today.month, today.day) < (birth_date.month, birth_date.day):
            age -= 1
        return age

    def is_condition_met(self, target_date, today):
        cond = self.cond_type
        if cond == "本月达到年龄":
            return (target_date.year == today.year and target_date.month == today.month)
        elif cond == "下月达到年龄":
            next_month_start = today.replace(day=1) + relativedelta(months=1)
            next_month_end = next_month_start + relativedelta(months=1, days=-1)
            return next_month_start <= target_date <= next_month_end
        elif cond == "本年度达到年龄":
            return target_date.year == today.year
        elif cond == "截止日期前达到年龄":
            return target_date <= self.deadline_date
        return False

    def check_new_targets(self):
        if not self.ensure_excel_file():
            return pd.DataFrame()

        self.get_current_settings()

        # 确保截止日期与界面完全同步
        if self.cond_type == "截止日期前达到年龄":
            try:
                self.deadline_date = datetime(self.year_var.get(), self.month_var.get(), self.day_var.get()).date()
                self.deadline_str = self.deadline_date.strftime("%Y-%m-%d")
            except:
                pass
            self.log(f"[检查] 截止日期: {self.deadline_str}")

        sheets_to_read = []
        if self.sheet_mode == "all":
            try:
                xl = pd.ExcelFile(self.excel_path)
                sheets_to_read = xl.sheet_names
            except Exception as e:
                self.log(f"获取工作表列表失败: {e}")
                return pd.DataFrame()
        else:
            if self.sheet_name:
                sheets_to_read = [self.sheet_name]
            else:
                self.log("未指定工作表")
                return pd.DataFrame()

        all_results = []
        today = datetime.now().date()

        for sheet in sheets_to_read:
            df = self.read_data_with_header(self.excel_path, sheet)
            if df is None or df.empty:
                self.log(f"工作表 {sheet} 读取失败或为空")
                continue

            # 识别列名
            name_keywords = ["姓名", "名字", "名称", "name"]
            birth_keywords = ["出生日期", "生日", "出生年月", "birth", "birthday", "出生"]
            unit_keywords = ["单位", "部门", "组织", "unit", "department"]
            name_col = None
            birth_col = None
            unit_col = None
            for col in df.columns:
                col_lower = str(col).lower()
                if any(kw in col_lower for kw in name_keywords):
                    name_col = col
                if any(kw in col_lower for kw in birth_keywords):
                    birth_col = col
                if any(kw in col_lower for kw in unit_keywords):
                    unit_col = col

            if name_col is None or birth_col is None:
                self.log(f"工作表 {sheet} 中未找到姓名或出生日期列")
                continue

            df[birth_col] = df[birth_col].apply(self.convert_birth_date)
            df = df.dropna(subset=[birth_col])
            if df.empty:
                continue
            df['birth_date'] = df[birth_col].dt.date

            target_age = self.target_age
            for idx, row in df.iterrows():
                birth = row['birth_date']
                target_date = self.get_target_date(birth, target_age)
                if self.is_condition_met(target_date, today):
                    name_val = row[name_col]
                    if pd.isna(name_val):
                        continue
                    name_str = str(name_val)
                    unit_str = str(row[unit_col]) if unit_col and not pd.isna(row[unit_col]) else ""
                    age_now = self.compute_age(birth)
                    unique_id = f"{sheet}_{name_str}_{birth}_{target_age}_{self.cond_type}"
                    if unique_id not in self.reminded_set:
                        # 核心过滤：截止日期≥今天时只提醒未来达标
                        if self.cond_type == "截止日期前达到年龄" and self.deadline_date >= today:
                            if target_date <= today:
                                # 跳过已达龄的，不输出日志
                                continue
                            else:
                                remind_type = "即将达标"
                        else:
                            remind_type = "已达标" if target_date <= today else "即将达标"

                        record = row.to_dict()
                        record["工作表"] = sheet
                        record["姓名"] = name_str
                        record["出生日期"] = birth
                        record["单位"] = unit_str
                        record["年龄"] = age_now
                        record["达到目标年龄日期"] = target_date
                        record["检测条件"] = self.cond_type
                        record["提醒类型"] = remind_type
                        record["unique_id"] = unique_id
                        all_results.append(record)

        if all_results:
            future_count = sum(1 for r in all_results if r.get("提醒类型") == "即将达标")
            self.log(f"发现 {len(all_results)} 位符合条件人员，其中即将达标: {future_count}")
            return pd.DataFrame(all_results)
        else:
            self.log("没有发现新的符合条件人员")
            return pd.DataFrame()

    # ---------- 闪烁窗口安全版本 ----------
    def blink_window(self, win, count=12, interval=400):
        if count <= 0:
            try:
                win.title("年龄预警提醒")
            except:
                pass
            return
        try:
            if not win.winfo_exists():
                return
            new_bg = "red" if (count % 2) == 0 else "lightgray"
            win.configure(bg=new_bg)
            self.root.after(interval, lambda: self.blink_window(win, count-1, interval))
        except tk.TclError:
            return

    # ---------- 实时提醒窗口 ----------
    def show_alert_window(self, new_df):
        if new_df.empty:
            return

        alert_win = tk.Toplevel(self.root)
        alert_win.title("年龄预警提醒")
        alert_win.geometry("1100x600")
        alert_win.transient(self.root)
        alert_win.grab_set()
        self.blink_window(alert_win)

        fixed_cols = ["工作表", "姓名", "出生日期", "目标年龄", "达到目标年龄日期", "提醒类型"]
        if "目标年龄" not in new_df.columns:
            new_df = new_df.copy()
            new_df["目标年龄"] = self.target_age

        exclude_cols = set(fixed_cols + ["unique_id", "检测条件", "年龄", "单位"])
        available_extra = [col for col in new_df.columns if col not in exclude_cols and not col.startswith("_")]
        default_extra = "单位" if "单位" in available_extra else (available_extra[0] if available_extra else "")

        control_frame = ttk.Frame(alert_win)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(control_frame, text="第六列:").pack(side=tk.LEFT)
        extra1_var = tk.StringVar(value=default_extra)
        extra1_combo = ttk.Combobox(control_frame, textvariable=extra1_var, values=available_extra, state="readonly", width=15)
        extra1_combo.pack(side=tk.LEFT, padx=5)

        ttk.Label(control_frame, text="第七列:").pack(side=tk.LEFT, padx=(10,0))
        extra2_var = tk.StringVar(value=default_extra if len(available_extra) > 1 else "")
        extra2_combo = ttk.Combobox(control_frame, textvariable=extra2_var, values=available_extra, state="readonly", width=15)
        extra2_combo.pack(side=tk.LEFT, padx=5)

        count_label = ttk.Label(control_frame, text=f"共 {len(new_df)} 人", foreground="blue")
        count_label.pack(side=tk.LEFT, padx=(10,0))

        frame = ttk.Frame(alert_win)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tree_ref = [None]
        scroll_ref = [None]

        def rebuild_table():
            if tree_ref[0]:
                tree_ref[0].destroy()
            if scroll_ref[0]:
                scroll_ref[0].destroy()
            extra1 = extra1_var.get()
            extra2 = extra2_var.get()
            display_cols = fixed_cols[:]
            if extra1:
                display_cols.append(extra1)
            if extra2 and extra2 != extra1:
                display_cols.append(extra2)
            new_tree = ttk.Treeview(frame, columns=display_cols, show="headings", height=15)
            for col in display_cols:
                new_tree.heading(col, text=col)
                new_tree.column(col, width=150, minwidth=100, anchor="center")
            new_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            new_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=new_tree.yview)
            new_tree.configure(yscrollcommand=new_scroll.set)
            new_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            tree_ref[0] = new_tree
            scroll_ref[0] = new_scroll

            for _, row in new_df.iterrows():
                values = []
                for col in display_cols:
                    v = row[col]
                    if isinstance(v, (datetime, pd.Timestamp)):
                        values.append(v.strftime("%Y-%m-%d"))
                    elif pd.isna(v):
                        values.append("")
                    else:
                        values.append(str(v))
                new_tree.insert("", tk.END, values=values)

        def on_extra_changed(*args):
            rebuild_table()

        extra1_combo.bind("<<ComboboxSelected>>", on_extra_changed)
        extra2_combo.bind("<<ComboboxSelected>>", on_extra_changed)

        def export_current():
            extra1 = extra1_var.get()
            extra2 = extra2_var.get()
            export_cols = fixed_cols[:]
            if extra1:
                export_cols.append(extra1)
            if extra2 and extra2 != extra1:
                export_cols.append(extra2)
            if not export_cols:
                return
            export_data = []
            for _, row in new_df.iterrows():
                record = {}
                for col in export_cols:
                    val = row[col]
                    if isinstance(val, (datetime, pd.Timestamp)):
                        record[col] = val.strftime("%Y-%m-%d")
                    else:
                        record[col] = val
                export_data.append(record)
            export_df = pd.DataFrame(export_data)

            date_str = datetime.now().strftime("%Y-%m-%d")
            cond = self.cond_var.get()
            if cond == "截止日期前达到年龄":
                if self.deadline_date >= datetime.now().date():
                    cond_display = f"截止日期({self.deadline_str})前 (未来提醒)"
                else:
                    cond_display = f"截止日期({self.deadline_str})前"
            else:
                cond_display = cond
            filename = f"{date_str}_{cond_display}.xlsx"
            filepath = filedialog.asksaveasfilename(
                parent=alert_win,
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=filename
            )
            if filepath:
                try:
                    export_df.to_excel(filepath, index=False)
                    messagebox.showinfo("导出成功", f"已导出到\n{filepath}", parent=alert_win)
                except Exception as e:
                    messagebox.showerror("导出失败", str(e), parent=alert_win)

        btn_export = ttk.Button(control_frame, text="导出当前数据", command=export_current)
        btn_export.pack(side=tk.RIGHT, padx=5)

        btn_frame = tk.Frame(alert_win)
        btn_frame.pack(pady=10)

        def on_confirm():
            for _, row in new_df.iterrows():
                record = row.to_dict()
                if "unique_id" not in record:
                    continue
                if any(r.get("unique_id") == record["unique_id"] for r in self.reminded_data):
                    continue
                self.reminded_data.append(record)
                self.reminded_set.add(record["unique_id"])
            self.save_reminded_data()
            alert_win.destroy()
            self.log(f"已记录 {len(new_df)} 条提醒，不再重复。")
            messagebox.showinfo("已记录", f"已将 {len(new_df)} 条提醒记录保存。", parent=self.root)

        def on_cancel():
            alert_win.destroy()
            self.log("用户取消提醒，下次将继续提醒。")

        btn_confirm = tk.Button(btn_frame, text="确认（不再提醒）", command=on_confirm, bg="lightgreen", width=15)
        btn_confirm.pack(side=tk.LEFT, padx=5)
        btn_cancel = tk.Button(btn_frame, text="取消（仍会提醒）", command=on_cancel, width=12)
        btn_cancel.pack(side=tk.LEFT, padx=5)

        rebuild_table()

    # ---------- 历史提醒窗口 ----------
    def show_reminded_records(self):
        if not self.reminded_data:
            messagebox.showinfo("无记录", "暂无已提醒的记录。")
            return

        df = pd.DataFrame(self.reminded_data)
        if df.empty:
            messagebox.showinfo("无记录", "暂无已提醒的记录。")
            return

        required_cols = ["工作表", "姓名", "出生日期", "年龄", "检测条件", "达到目标年龄日期"]
        for col in required_cols:
            if col not in df.columns:
                messagebox.showerror("数据错误", f"历史记录缺少必要字段：{col}")
                return

        df['达到目标年龄日期'] = pd.to_datetime(df['达到目标年龄日期'])
        df = df.sort_values(by='达到目标年龄日期', ascending=False)

        exclude = set(required_cols + ["unique_id"])
        available_extra = [col for col in df.columns if col not in exclude and not col.startswith("_")]
        if not available_extra:
            available_extra = [""]

        rec_win = tk.Toplevel(self.root)
        rec_win.title("历史提醒记录")
        rec_win.geometry("1200x600")
        rec_win.transient(self.root)

        control_frame = ttk.Frame(rec_win)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(control_frame, text="第六列:").pack(side=tk.LEFT)
        extra1_var = tk.StringVar(value=available_extra[0] if available_extra else "")
        extra1_combo = ttk.Combobox(control_frame, textvariable=extra1_var, values=available_extra, state="readonly", width=15)
        extra1_combo.pack(side=tk.LEFT, padx=5)

        ttk.Label(control_frame, text="第七列:").pack(side=tk.LEFT, padx=(10,0))
        extra2_var = tk.StringVar(value=available_extra[1] if len(available_extra) > 1 else (available_extra[0] if available_extra else ""))
        extra2_combo = ttk.Combobox(control_frame, textvariable=extra2_var, values=available_extra, state="readonly", width=15)
        extra2_combo.pack(side=tk.LEFT, padx=5)

        count_label = ttk.Label(control_frame, text=f"共 {len(df)} 人", foreground="blue")
        count_label.pack(side=tk.LEFT, padx=(10,0))

        def export_history():
            filepath = filedialog.asksaveasfilename(
                parent=rec_win,
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"{datetime.now().strftime('%Y-%m-%d')}_综合结果.xlsx"
            )
            if filepath:
                try:
                    export_df = df.drop(columns=['unique_id'], errors='ignore')
                    export_df.to_excel(filepath, index=False)
                    messagebox.showinfo("导出成功", f"已导出到\n{filepath}", parent=rec_win)
                except Exception as e:
                    messagebox.showerror("导出失败", str(e), parent=rec_win)

        btn_export = ttk.Button(control_frame, text="导出数据", command=export_history)
        btn_export.pack(side=tk.RIGHT, padx=5)

        frame = ttk.Frame(rec_win)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        fixed_cols = ["工作表", "姓名", "出生日期", "年龄", "检测条件", "达到目标年龄日期"]
        tree_ref = [None]
        scroll_ref = [None]

        def rebuild_table():
            if tree_ref[0]:
                tree_ref[0].destroy()
            if scroll_ref[0]:
                scroll_ref[0].destroy()
            extra1 = extra1_var.get()
            extra2 = extra2_var.get()
            display_cols = fixed_cols[:]
            if extra1:
                display_cols.append(extra1)
            if extra2 and extra2 != extra1:
                display_cols.append(extra2)
            new_tree = ttk.Treeview(frame, columns=display_cols, show="headings", height=20)
            for col in display_cols:
                new_tree.heading(col, text=col)
                new_tree.column(col, width=140, minwidth=80, anchor="center")
            new_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            new_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=new_tree.yview)
            new_tree.configure(yscrollcommand=new_scroll.set)
            new_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            tree_ref[0] = new_tree
            scroll_ref[0] = new_scroll

            for idx, row in df.iterrows():
                values = []
                for col in display_cols:
                    v = row[col]
                    if isinstance(v, (datetime, pd.Timestamp)):
                        values.append(v.strftime("%Y-%m-%d"))
                    elif pd.isna(v):
                        values.append("")
                    else:
                        values.append(str(v))
                new_tree.insert("", tk.END, values=values)

        def on_extra_changed(*args):
            rebuild_table()

        extra1_combo.bind("<<ComboboxSelected>>", on_extra_changed)
        extra2_combo.bind("<<ComboboxSelected>>", on_extra_changed)
        rebuild_table()

    # ---------- 监控流程 ----------
    def perform_check(self):
        self.log("开始检查...")
        try:
            new_people = self.check_new_targets()
            if new_people.empty:
                # 已经在 check_new_targets 中输出日志，不再重复
                pass
            else:
                self.log(f"发现 {len(new_people)} 位符合条件人员，弹出提醒窗口。")
                self.root.after(0, lambda: self.show_alert_window(new_people))
        except Exception as e:
            self.log(f"检查过程中出错: {e}")
            import traceback
            self.log(traceback.format_exc())

    def manual_check(self):
        if not self.ensure_excel_file():
            return
        self.log("手动触发检查")
        self.perform_check()

    def schedule_daily(self):
        if not self.monitoring:
            return
        now = datetime.now()
        target_time = now.replace(hour=self.hour, minute=self.minute, second=0, microsecond=0)
        if now >= target_time:
            target_time += timedelta(days=1)
        delay_seconds = (target_time - now).total_seconds()
        self.after_id = self.root.after(int(delay_seconds * 1000), self.daily_check)
        self.log(f"下次自动检查时间: {target_time.strftime('%Y-%m-%d %H:%M:%S')}")

    def daily_check(self):
        if not self.monitoring:
            return
        self.log("定时检查触发")
        self.perform_check()
        self.schedule_daily()

    def start_monitor(self):
        if self.monitoring:
            return
        if not self.ensure_excel_file():
            return
        self.get_current_settings()
        test_df = self.read_data_with_header(self.excel_path, self.sheet_name if self.sheet_mode=="single" else None, rows=5)
        if test_df is None or test_df.empty:
            messagebox.showerror("读取失败", "无法读取数据，请检查表头行设置是否正确。")
            return
        name_keywords = ["姓名", "名字", "名称", "name"]
        birth_keywords = ["出生日期", "生日", "出生年月", "birth", "birthday", "出生"]
        has_name = any(any(kw in str(col).lower() for kw in name_keywords) for col in test_df.columns)
        has_birth = any(any(kw in str(col).lower() for kw in birth_keywords) for col in test_df.columns)
        if not (has_name and has_birth):
            reply = messagebox.askyesno("列名警告", 
                                        "自动识别到的列名中不包含明确的“姓名”或“出生日期”关键词，继续监控可能无法正确匹配人员。\n是否仍然继续？")
            if not reply:
                return
        self.monitoring = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.status_var.set("状态：监控运行中")
        self.log("监控已启动")
        self.schedule_daily()
        self.perform_check()

    def stop_monitor(self):
        if not self.monitoring:
            return
        self.monitoring = False
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.status_var.set("状态：已停止")
        self.log("监控已停止")

    def clear_reminded(self):
        if messagebox.askyesno("确认清除", "确定要清除所有提醒记录吗？"):
            self.reminded_set.clear()
            self.reminded_data.clear()
            self.save_reminded_data()
            self.log("已清空提醒记录。")
            messagebox.showinfo("已清除", "提醒记录已清空。")

    def on_closing(self):
        self.stop_monitor()
        self.save_reminded_data()
        self.save_config()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AgeMonitorApp(root)
    root.mainloop()
