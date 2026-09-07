# -*- coding: utf-8 -*-
"""
STR2ASCII - 字符串转 ASCII 十六进制码工具 (tkinter 桌面版)

功能：
    输入：左侧文本框 4 位 ID（格式：大写字母 F/G/H/I + 三位数字，如 F001）
    输出：右侧文本框 ASCII 十六进制码（如 46 30 30 31）
    记录：每次转换追加写入本地 conversion_log.csv，重复 ID 警告且不重复写入

作者：zhanghao
"""
import csv
import os
import re
import sys
import ctypes
from datetime import datetime

import tkinter as tk
from tkinter import messagebox, ttk

# ---------- 单实例锁（Windows 命名互斥量） ----------
# 防止用户双击多次开启多个窗口，导致 CSV 并发写入损坏。
# 互斥量由系统管理，进程退出/崩溃时自动释放，不会残留死锁。
MUTEX_NAME = "Local\\STR2ASCII_SingleInstance"  # Local 前缀：仅本登录会话


def acquire_single_instance_lock():
    """尝试获取单实例锁，返回句柄；若已存在其他实例则返回 None"""
    if sys.platform != "win32":
        return "non_windows_ok"  # 非 Windows 平台跳过锁（本程序面向 Windows）
    # CreateMutexW: 返回 NULL 表示失败（权限等），ERROR_ALREADY_EXISTS 表示已有实例
    handle = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not handle:
        return None
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        # 已有实例在运行，需关闭刚创建的句柄
        ctypes.windll.kernel32.CloseHandle(handle)
        return None
    return handle

# ---------- 打包后 tcl/tk 库路径设置 ----------
# PyInstaller 打包时会把 tcl/tk 的 DLL 和库目录作为数据文件加进来，
# 运行时需要显式告诉 Tcl/Tk 到哪里找库文件，否则报 DLL load failed
if getattr(sys, "frozen", False):
    _base = sys._MEIPASS
    _tcl_lib = os.path.join(_base, "_tcl_data", "tcl8.6")
    _tk_lib = os.path.join(_base, "_tcl_data", "tk8.6")
    if os.path.isdir(_tcl_lib):
        os.environ["TCL_LIBRARY"] = _tcl_lib
    if os.path.isdir(_tk_lib):
        os.environ["TK_LIBRARY"] = _tk_lib

# ---------- 路径处理（兼容 PyInstaller 打包） ----------
if getattr(sys, "frozen", False):
    # 打包成 exe 后，CSV 放在 exe 同目录
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_FILE = os.path.join(BASE_DIR, "conversion_log.csv")
CSV_HEADER = ["id", "hex_code", "timestamp"]

# ID 格式：一位大写字母（F/G/H/I）+ 三位数字，共 4 位
ID_PATTERN = re.compile(r"^[FGHI]\d{3}$")


class CsvFileInUseError(Exception):
    """CSV 文件被 Excel 等程序占用时抛出的异常"""


def _is_file_in_use_error(exc):
    if isinstance(exc, PermissionError):
        return True
    if isinstance(exc, OSError):
        win_error = getattr(exc, "winerror", None)
        if win_error == 32:
            return True
        errno = getattr(exc, "errno", None)
        if errno == 13:
            return True
    return False


def _build_file_in_use_message():
    return "CSV 文件正在被 Excel 等程序占用，请先关闭 Excel，再稍后重试转换"


# ---------- CSV 记录 ----------
def load_existing_ids():
    """读取 CSV 中已记录的所有 ID"""
    ids = set()
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r", encoding="utf-8-sig", newline="") as f:
                for row in csv.DictReader(f):
                    ids.add(row.get("id", ""))
        except Exception as e:
            if _is_file_in_use_error(e):
                raise CsvFileInUseError(_build_file_in_use_message()) from e
            pass
    return ids


def append_record(rec_id, hex_code):
    """追加一条转换记录到 CSV"""
    new_file = not os.path.exists(LOG_FILE)
    try:
        with open(LOG_FILE, "a", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            if new_file:
                writer.writerow(CSV_HEADER)
            writer.writerow([
                rec_id,
                hex_code,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ])
    except Exception as e:
        if _is_file_in_use_error(e):
            raise CsvFileInUseError(_build_file_in_use_message()) from e
        raise


def do_convert(rec_id):
    """校验 -> 转换 -> 重复检查 -> 记录，返回 (success, result, duplicate, message)"""
    text = rec_id.strip()

    if not text:
        return False, "", False, "输入不能为空"

    if not ID_PATTERN.match(text):
        return False, "", False, "格式错误：请输入 4 位 ID（大写字母 F/G/H/I + 三位数字），如 F001"

    # 核心转换：每个字符 -> 大写两位十六进制，空格分隔
    hex_codes = [f"{ord(ch):02X}" for ch in text]
    result = " ".join(hex_codes)

    try:
        # 重复检查：重复则不写入记录，仅提示警告
        duplicate = text in load_existing_ids()
        if not duplicate:
            append_record(text, result)
    except CsvFileInUseError as e:
        return False, "", False, str(e)

    return True, result, duplicate, ""


# ---------- 界面 ----------
class Str2AsciiApp:
    def __init__(self, root):
        self.root = root
        root.title("STR2ASCII - 字符串转 ASCII 十六进制码")
        root.geometry("860x620")
        root.minsize(860, 620)
        root.configure(bg="#f0f4f8")

        self._build_ui()
        self._load_history_to_list()

    def _build_ui(self):
        # 标题
        title = tk.Label(
            self.root,
            text="STR2ASCII 转换工具",
            font=("Microsoft YaHei", 18, "bold"),
            bg="#f0f4f8",
            fg="#1e293b"
        )
        title.pack(pady=(18, 4))

        subtitle = tk.Label(
            self.root,
            text="输入 4 位 ID（大写 F/G/H/I + 三位数字），输入满即自动转换并记录历史",
            font=("Microsoft YaHei", 10),
            bg="#f0f4f8",
            fg="#64748b"
        )
        subtitle.pack(pady=(0, 14))

        # 设备 ID 解锁命令（左上角固定显示，方便复制）
        unlock_frame = tk.Frame(self.root, bg="#f0f4f8")
        unlock_frame.pack(padx=24, fill="x", pady=(0, 10))
        
        unlock_left = tk.Frame(unlock_frame, bg="#f0f4f8")
        unlock_left.pack(side="left", padx=8, pady=6)
        
        tk.Label(
            unlock_left, text="设备 ID 解锁命令",
            font=("Microsoft YaHei", 9, "bold"), bg="#f0f4f8", fg="#334155"
        ).pack(side="left")
        
        self.unlock_cmd = "45 0E 43 4D 44 49"
        self.unlock_cmd_var = tk.StringVar(value=self.unlock_cmd)
        self.unlock_entry = tk.Entry(
            unlock_left, textvariable=self.unlock_cmd_var,
            font=("Consolas", 10, "bold"), width=18,
            state="readonly"
        )
        self.unlock_entry.pack(side="left", padx=8)
        
        tk.Button(
            unlock_left, text="复制",
            font=("Microsoft YaHei", 9),
            bg="#e2e8f0", fg="#475569", activebackground="#cbd5e1",
            relief="flat", cursor="hand2",
            padx=8, pady=2, command=self._copy_unlock_cmd
        ).pack(side="left")

        tk.Label(
            unlock_left, text="设备 ID 查询命令",
            font=("Microsoft YaHei", 9, "bold"), bg="#f0f4f8", fg="#334155"
        ).pack(side="left", padx=(160, 0))
        
        self.query_cmd = "F0 01"
        self.query_cmd_var = tk.StringVar(value=self.query_cmd)
        self.query_entry = tk.Entry(
            unlock_left, textvariable=self.query_cmd_var,
            font=("Consolas", 10, "bold"), width=15,
            state="readonly"
        )
        self.query_entry.pack(side="left", padx=8)

        tk.Button(
            unlock_left, text="复制",
            font=("Microsoft YaHei", 9),
            bg="#e2e8f0", fg="#475569", activebackground="#cbd5e1",
            relief="flat", cursor="hand2",
            padx=8, pady=2, command=self._copy_query_cmd
        ).pack(side="left")

        # 主区域：左右两个文本框
        main = tk.Frame(self.root, bg="#f0f4f8")
        main.pack(padx=24, fill="x")

        # 左侧：输入
        left = tk.Frame(main, bg="#f0f4f8")
        left.pack(side="left", expand=True, fill="both")

        left_label = tk.Frame(left, bg="#f0f4f8")
        left_label.pack(anchor="w", pady=(0, 6))
        tk.Label(
            left_label, text="输入 ID",
            font=("Microsoft YaHei", 11, "bold"), bg="#f0f4f8", fg="#334155"
        ).pack(side="left")
        tk.Label(
            left_label, text="  F/G/H/I + 3位数字",
            font=("Microsoft YaHei", 9), bg="#e0f2fe", fg="#0284c7",
            padx=6, pady=2
        ).pack(side="left", padx=8)

        self.input_text = tk.Text(
            left, width=24, height=4,
            font=("Consolas", 18, "bold"),
            relief="solid", bd=1, wrap="none",
            insertwidth=2
        )
        self.input_text.pack(fill="x")
        self.input_text.bind("<KeyPress>", self._on_input_change)
        self.input_text.focus_set()

        # 中间箭头
        tk.Label(
            main, text="  →  ",
            font=("Microsoft YaHei", 22, "bold"),
            bg="#f0f4f8", fg="#3b82f6"
        ).pack(side="left", padx=16)

        # 右侧：输出
        right = tk.Frame(main, bg="#f0f4f8")
        right.pack(side="left", expand=True, fill="both")

        right_label = tk.Frame(right, bg="#f0f4f8")
        right_label.pack(anchor="w", pady=(0, 6))
        tk.Label(
            right_label, text="ASCII 十六进制码",
            font=("Microsoft YaHei", 11, "bold"), bg="#f0f4f8", fg="#334155"
        ).pack(side="left")
        self.prefix_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            right_label, text="前缀 85 07",
            font=("Microsoft YaHei", 9), bg="#f0f4f8", fg="#64748b",
            variable=self.prefix_var
        ).pack(side="left", padx=8)

        self.output_text = tk.Text(
            right, width=24, height=4,
            font=("Consolas", 18, "bold"),
            relief="solid", bd=1, wrap="none",
            bg="#f1f5f9", state="disabled",
            insertwidth=0
        )
        self.output_text.pack(fill="x")

        # 按钮行
        btn_row = tk.Frame(self.root, bg="#f0f4f8")
        btn_row.pack(pady=14)

        self.btn_convert = tk.Button(
            btn_row, text="转 换",
            font=("Microsoft YaHei", 12, "bold"),
            bg="#2563eb", fg="white", activebackground="#1d4ed8",
            activeforeground="white", relief="flat", cursor="hand2",
            padx=28, pady=8, command=self._on_convert
        )
        self.btn_convert.pack(side="left", padx=8)

        tk.Button(
            btn_row, text="复制结果",
            font=("Microsoft YaHei", 11),
            bg="#e2e8f0", fg="#475569", activebackground="#cbd5e1",
            relief="flat", cursor="hand2",
            padx=24, pady=8, command=self._on_copy
        ).pack(side="left", padx=8)

        tk.Button(
            btn_row, text="清空历史",
            font=("Microsoft YaHei", 11),
            bg="#fee2e2", fg="#dc2626", activebackground="#fecaca",
            relief="flat", cursor="hand2",
            padx=24, pady=8, command=self._on_clear_history
        ).pack(side="left", padx=8)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = tk.Label(
            self.root, textvariable=self.status_var,
            font=("Microsoft YaHei", 10), bg="#f0f4f8", fg="#64748b"
        )
        self.status_label.pack(pady=(0, 4))

        # 历史记录（可滚动列表）
        hist_frame = tk.Frame(self.root, bg="#f0f4f8")
        hist_frame.pack(fill="both", expand=True, padx=24, pady=(0, 15))

        tk.Label(
            hist_frame, text="转换历史",
            font=("Microsoft YaHei", 10, "bold"), bg="#f0f4f8", fg="#334155"
        ).pack(anchor="w", pady=(0, 4))

        # Listbox + 滚动条
        hist_box_frame = tk.Frame(hist_frame, bg="#f0f4f8")
        hist_box_frame.pack(fill="both", expand=True)

        self.history_list = tk.Listbox(
            hist_box_frame, height=8,
            font=("Consolas", 10), relief="solid", bd=1,
            selectbackground="#bfdbfe"
        )
        self.history_list.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(hist_box_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        
        self.history_list.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.history_list.yview)

        # 绑定快捷键：Enter 转换，Ctrl+C 复制
        self.root.bind("<Return>", lambda e: self._on_convert())
        #self.input_text.bind("<Control-Return>", lambda e: self._on_convert())

    def _on_input_change(self, event=None):
        """在按键插入前处理：
        - 输入框已有内容时，若键入大写 F/G/H/I，先清空（开始新 ID）
        - 限制 4 位"""
        if event is None:
            return

        # 只处理普通字符按键（排除 Ctrl/Shift/BackSpace 等控制键）
        if len(event.keysym) != 1:
            return

        new_char = event.char
        # 已有内容时，键入 F/G/H/I 开头 -> 自动清空，开始新 ID
        existing = self.input_text.get("1.0", "end-1c")
        if existing and new_char in ("F", "G", "H", "I"):
            self.input_text.delete("1.0", "end")

        # 限制 4 位，解除注释后符合格式则自动转换（用 after 保证在字符插入后执行截断）
        def _trim():
            current = self.input_text.get("1.0", "end-1c")
            if len(current) > 4:
                self.input_text.delete("1.0", "end")
                self.input_text.insert("1.0", current[:4])
            # # 输入满 4 位且符合格式时自动转换
            # val = self.input_text.get("1.0", "end-1c")
            # if len(val) == 4 and ID_PATTERN.match(val):
            #     self._on_convert()
        self.input_text.after_idle(_trim)

    def _on_convert(self):
        rec_id = self.input_text.get("1.0", "end-1c")
        success, result, duplicate, msg = do_convert(rec_id)

        # 更新输出框
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", "end")
        if success:
            display = ("85 07 " + result) if self.prefix_var.get() else result
            self.output_text.insert("1.0", display)
            self.output_text.config(fg="#16a34a")
        else:
            self.output_text.insert("1.0", "")
            self.output_text.config(fg="#dc2626")
        self.output_text.config(state="disabled")

        # 更新状态
        if success:
            # 每次转换成功都刷新历史列表（含重复，保证列表显示最新）
            self._load_history_to_list()
            if duplicate:
                self._set_status("[警告] 该 ID 已转换过（重复记录），未写入历史", "#d97706")
            else:
                self._set_status("[成功] 转换完成，已记录历史", "#16a34a")
        else:
            self._set_status("[错误] " + msg, "#dc2626")
            if "Excel" in msg:
                messagebox.showwarning("文件占用", msg)
            else:
                messagebox.showwarning("输入错误", msg)

    def _on_copy(self):
        result = self.output_text.get("1.0", "end-1c")
        if not result:
            messagebox.showinfo("提示", "没有可复制的内容")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(result)
        self._set_status("[成功] 已复制到剪贴板", "#16a34a")

    def _copy_unlock_cmd(self):
        """复制设备 ID 解锁命令到剪贴板"""
        self.root.clipboard_clear()
        self.root.clipboard_append(self.unlock_cmd)
        self._set_status("[成功] 已复制解锁命令: " + self.unlock_cmd, "#16a34a")

    def _copy_query_cmd(self):
        """复制设备 ID 查询命令到剪贴板"""
        self.root.clipboard_clear()
        self.root.clipboard_append(self.query_cmd)
        self._set_status("[成功] 已复制查询命令: " + self.query_cmd, "#16a34a")

    def _on_clear_history(self):
        if not os.path.exists(LOG_FILE):
            self._set_status("历史记录为空", "#64748b")
            return
        if messagebox.askyesno("确认", "确定清空全部转换历史记录？\n此操作不可撤销！"):
            try:
                with open(LOG_FILE, "w", encoding="utf-8-sig", newline="") as f:
                    csv.writer(f).writerow(CSV_HEADER)
                self._set_status("[成功] 历史记录已清空", "#16a34a")
                self._load_history_to_list()
            except Exception as e:
                if _is_file_in_use_error(e):
                    self._set_status("✘ " + _build_file_in_use_message(), "#dc2626")
                    messagebox.showwarning("文件占用", _build_file_in_use_message())
                else:
                    self._set_status("✘ 清空失败: " + str(e), "#dc2626")

    def _set_status(self, msg, color):
        self.status_var.set(msg)
        self.status_label.config(fg=color)

    def _load_history_to_list(self):
        """从 CSV 读取最近 50 条记录显示"""
        self.history_list.delete(0, "end")
        if not os.path.exists(LOG_FILE):
            return
        try:
            with open(LOG_FILE, "r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f))
            for row in rows[-50:]:
                self.history_list.insert(
                    "end",
                    f"{row.get('id',''):<6}  {row.get('hex_code',''):<14}  {row.get('timestamp','')}"
                )
            self.history_list.see("end")
        except CsvFileInUseError as e:
            self._set_status("✘ " + str(e), "#dc2626")
            messagebox.showwarning("文件占用", str(e))
        except Exception:
            pass


def main():
    # 先尝试获取单实例锁
    lock_handle = acquire_single_instance_lock()
    if lock_handle is None:
        try:
            # 尝试弹出提示（若没有可用显示器/会话，忽略）
            import tkinter.messagebox as mb
            mb.showwarning(
                "程序已在运行",
                "STR2ASCII 已经在运行中。\n\n为避免数据记录冲突，请切换到已打开的窗口，"
                "或关闭它后再启动。"
            )
        except Exception:
            pass
        sys.exit(0)  # 已有实例，退出

    root = tk.Tk()
    app = Str2AsciiApp(root)
    # 主程序运行中保持锁句柄引用，防止被 GC
    app._lock_handle = lock_handle
    root.mainloop()
    # 关闭互斥量句柄（进程退出前释放）
    try:
        if sys.platform == "win32":
            ctypes.windll.kernel32.CloseHandle(lock_handle)
    except Exception:
        pass


if __name__ == "__main__":
    main()
