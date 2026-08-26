# -*- coding: utf-8 -*-
"""
nRF54L15 自动烧录工具 v7.1
新增：补烧失败后自动进入单独烧录流程
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox, simpledialog
import subprocess
import os
import sys
import json
import re
import time
import threading
import openpyxl
from datetime import datetime

# Windows 下隐藏 subprocess 弹出的命令行窗口
_SW_FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


# ─────────────────────────────────────────────
#  ToolTip（路径悬浮提示）
# ─────────────────────────────────────────────
class ToolTip:
    """鼠标悬停时在控件旁显示全文提示"""
    def __init__(self, widget, textvariable=None, text=""):
        self._widget = widget
        self._textvar = textvariable
        self._text = text
        self._tw = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, _=None):
        tip = self._textvar.get() if self._textvar else self._text
        if not tip:
            return
        x = self._widget.winfo_rootx()
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        tw = tk.Toplevel(self._widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(tw, text=tip, justify=tk.LEFT,
                       background="#1e2733", foreground="#e6edf3",
                       relief=tk.SOLID, borderwidth=1,
                       font=("Consolas", 9), padx=6, pady=3)
        lbl.pack()
        self._tw = tw

    def _hide(self, _=None):
        if self._tw:
            self._tw.destroy()
            self._tw = None


# ─────────────────────────────────────────────
#  颜色 & 字体
# ─────────────────────────────────────────────
BG_DARK   = "#0d1117"
BG_PANEL  = "#161b22"
BG_CARD   = "#1c2128"
BG_INPUT  = "#21262d"
BORDER    = "#30363d"
ACCENT    = "#58a6ff"
ACCENT2   = "#1f6feb"
SUCCESS   = "#3fb950"
DANGER    = "#f85149"
WARNING   = "#d29922"
TEXT_PRI  = "#e6edf3"
TEXT_SEC  = "#8b949e"
TEXT_DIM  = "#484f58"

FONT_TITLE = ("Microsoft YaHei UI", 16, "bold")
FONT_HEAD  = ("Microsoft YaHei UI", 10, "bold")
FONT_BODY  = ("Microsoft YaHei UI", 9)
FONT_MONO  = ("Consolas", 9)
FONT_SMALL = ("Microsoft YaHei UI", 8)
FONT_BIG   = ("Microsoft YaHei UI", 36, "bold")
FONT_MID   = ("Microsoft YaHei UI", 14)

# 步骤名（固定宽度，全中文对齐）
STEP_NAMES = {
    1: "步骤一：芯片解除保护",
    2: "步骤二：擦除芯片内容",
    3: "步骤三：密钥上传    ",
    4: "步骤四：生成烧录文件",
    5: "步骤五：烧录芯片    ",
}


# ─────────────────────────────────────────────
#  工具函数
# ─────────────────────────────────────────────
def sn_to_hex32(sn: str) -> str:
    raw = sn.strip().encode("utf-8").hex().upper()
    return raw[:32] if len(raw) >= 32 else raw.ljust(32, "0")


def parse_bat_env(bat_path: str) -> dict:
    env = os.environ.copy()
    if not bat_path or not os.path.isfile(bat_path):
        return env
    try:
        with open(bat_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        for line in lines:
            line = line.strip()
            m = re.match(r"(?i)^set\s+([^=]+)=(.*)$", line)
            if m:
                key   = m.group(1).strip()
                value = m.group(2).strip()
                value = re.sub(
                    r"%([^%]+)%",
                    lambda x: env.get(x.group(1), x.group(0)),
                    value
                )
                env[key] = value
    except Exception:
        pass
    return env


def get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def extract_error_summary(full_err: str) -> str:
    if not full_err:
        return "未知错误"
    lines = [l.strip() for l in full_err.strip().splitlines() if l.strip()]
    for line in reversed(lines):
        if any(kw in line for kw in
               ["Error:", "Exception:", "error:", "Failed", "失败"]):
            return line[:120]
    return lines[-1][:120] if lines else full_err[:120]


# ─────────────────────────────────────────────
#  持久化状态
# ─────────────────────────────────────────────
class FlashState:
    def __init__(self):
        self.state_file = os.path.join(get_base_dir(), "flash_state.json")
        self.data = {
            "next_row":        2,
            "total_flashed":   0,
            "daily_counter":   {},
            "last_exit_clean": True,
            "cfg_hex":         "",
            "cfg_xlsx":        "",
            "cfg_bat":         "",
            "cfg_out_dir":     "./output",
            "cfg_base_addr":   "0x178000",
            "cfg_use_bat":     True,
            "cfg_start_row":   "2",
        }
        self.load()

    def load(self):
        if os.path.isfile(self.state_file):
            try:
                with open(self.state_file, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                self.data.update(saved)
            except Exception:
                pass

    def save(self):
        with open(self.state_file, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

    def save_config(self, hex_path, xlsx_path, bat_path,
                    out_dir, base_addr, use_bat, start_row="2"):
        self.data.update({
            "cfg_hex":       hex_path,
            "cfg_xlsx":      xlsx_path,
            "cfg_bat":       bat_path,
            "cfg_out_dir":   out_dir,
            "cfg_base_addr": base_addr,
            "cfg_use_bat":   use_bat,
            "cfg_start_row": start_row,
        })
        self.save()

    def mark_start(self):
        self.data["last_exit_clean"] = False
        self.save()

    def mark_end(self):
        self.data["last_exit_clean"] = True
        self.save()

    @property
    def last_exit_clean(self):
        return self.data.get("last_exit_clean", True)

    @property
    def next_row(self):
        return self.data["next_row"]

    @next_row.setter
    def next_row(self, v):
        self.data["next_row"] = v
        self.save()

    @property
    def total_flashed(self):
        return self.data["total_flashed"]

    @total_flashed.setter
    def total_flashed(self, v):
        self.data["total_flashed"] = v
        self.save()

    def next_provisioned_name(self) -> tuple:
        today = datetime.now().strftime("%Y%m%d")
        daily = self.data.setdefault("daily_counter", {})
        daily[today] = daily.get(today, 0) + 1
        idx = daily[today]
        self.save()
        return today, idx, f"provisioned_{today}_{idx:03d}"

    @property
    def provisioned_counter(self):
        today = datetime.now().strftime("%Y%m%d")
        return self.data.get("daily_counter", {}).get(today, 0)


# ─────────────────────────────────────────────
#  主应用
# ─────────────────────────────────────────────
class NordicFlashApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("nRF54L15 自动烧录工具  v7.1")
        self.root.geometry("1160x900")
        self.root.configure(bg=BG_DARK)
        self.root.resizable(True, True)

        self.state        = FlashState()
        self.jlink_ids    = []
        self.is_flashing  = False
        self.stop_flag    = False
        self.log_lines    = []
        self.log_full     = []
        self.btn_start    = None
        self.btn_single   = None   # 单独烧录按钮
        self.use_bat      = None

        # 单独烧录队列：[(jlink_id, reason), ...]
        # 补烧失败后自动填入，逐一串行处理
        self._single_queue       = []
        self._session_start_row  = 2   # 本次烧录的起始行

        self._build_ui()
        self._load_saved_config()
        self._detect_jlinks()
        self._check_last_exit()

    # ──────────────────────────────────────────
    #  启动检测
    # ──────────────────────────────────────────
    def _check_last_exit(self):
        if not self.state.last_exit_clean:
            self.root.after(600, lambda: messagebox.showwarning(
                "⚠️ 检测到上次异常退出",
                f"上次烧录程序可能未正常结束（断电或崩溃）。\n\n"
                f"当前 Excel 行指针：第 {self.state.next_row} 行\n\n"
                f"请人工核查第 {self.state.next_row - 1} 行数据\n"
                f"对应的板子是否已成功烧录，确认后再继续。"
            ))
            self.log("⚠️ 检测到上次异常退出，请核查上一行数据是否烧录成功",
                     "WARNING")

    # ──────────────────────────────────────────
    #  路径配置持久化
    # ──────────────────────────────────────────
    def _load_saved_config(self):
        d = self.state.data
        for var, key, default in [
            (self.hex_path,  "cfg_hex",       ""),
            (self.xlsx_path, "cfg_xlsx",      ""),
            (self.bat_path,  "cfg_bat",       ""),
            (self.out_dir,   "cfg_out_dir",   "./output"),
            (self.base_addr, "cfg_base_addr", "0x178000"),
            (self.start_row, "cfg_start_row", "2"),
        ]:
            var.set(d.get(key, default))
        self.use_bat.set(d.get("cfg_use_bat", True))
        self.root.after(200, self._validate_paths)

    def _validate_paths(self):
        for var, entry_widget in self._path_entries:
            path = var.get()
            if path and not os.path.exists(path):
                entry_widget.config(bg="#3a1010")
            else:
                entry_widget.config(bg=BG_INPUT)

    def _auto_save_config(self, *_):
        self.state.save_config(
            self.hex_path.get(), self.xlsx_path.get(),
            self.bat_path.get(), self.out_dir.get(),
            self.base_addr.get(), self.use_bat.get(),
            self.start_row.get()
        )
        self._validate_paths()

    # ──────────────────────────────────────────
    #  界面构建
    # ──────────────────────────────────────────
    def _build_ui(self):
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_rowconfigure(3, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        # 标题栏
        title_bar = tk.Frame(self.root, bg=BG_PANEL, height=60)
        title_bar.grid(row=0, column=0, sticky="ew")
        title_bar.grid_propagate(False)
        title_bar.grid_columnconfigure(0, weight=1)
        tk.Label(
            title_bar, text="⚡  nRF54L15  自动烧录工具",
            font=FONT_TITLE, bg=BG_PANEL, fg=ACCENT
        ).grid(row=0, column=0, sticky="w", padx=22, pady=12)
        self.status_badge = tk.Label(
            title_bar, text="● 就绪",
            font=FONT_HEAD, bg=BG_PANEL, fg=SUCCESS)
        self.status_badge.grid(row=0, column=2, sticky="e", padx=22)
        tk.Label(
            title_bar, text="v7.1",
            font=FONT_SMALL, bg=BG_PANEL, fg=TEXT_DIM
        ).grid(row=0, column=1, sticky="e", padx=4)

        # 主体
        body = tk.Frame(self.root, bg=BG_DARK)
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=8)
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=0, minsize=420)
        body.grid_columnconfigure(1, weight=1)

        left = tk.Frame(body, bg=BG_DARK)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        right = tk.Frame(body, bg=BG_DARK)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)

        self._path_entries = []
        self._build_left(left)
        self._build_right(right)

        # 进度条
        prog_frame = tk.Frame(self.root, bg=BG_PANEL, height=32)
        prog_frame.grid(row=2, column=0, sticky="ew")
        prog_frame.grid_propagate(False)
        prog_frame.grid_columnconfigure(1, weight=1)
        prog_frame.grid_columnconfigure(4, weight=1)

        tk.Label(prog_frame, text="整体进度", font=FONT_SMALL,
                 bg=BG_PANEL, fg=TEXT_SEC
                 ).grid(row=0, column=0, padx=(12, 6), pady=5)
        self.overall_bar = ttk.Progressbar(
            prog_frame, orient="horizontal", mode="determinate")
        self.overall_bar.grid(row=0, column=1, sticky="ew",
                              padx=(0, 6), pady=5)
        self.overall_lbl = tk.Label(
            prog_frame, text="等待开始", font=FONT_SMALL,
            bg=BG_PANEL, fg=TEXT_SEC, width=24, anchor="w")
        self.overall_lbl.grid(row=0, column=2, padx=(0, 20))

        tk.Label(prog_frame, text="步骤五进度", font=FONT_SMALL,
                 bg=BG_PANEL, fg=TEXT_SEC
                 ).grid(row=0, column=3, padx=(0, 6))
        self.flash_bar = ttk.Progressbar(
            prog_frame, orient="horizontal", mode="determinate")
        self.flash_bar.grid(row=0, column=4, sticky="ew",
                            padx=(0, 6), pady=5)
        self.flash_lbl = tk.Label(
            prog_frame, text="空闲", font=FONT_SMALL,
            bg=BG_PANEL, fg=TEXT_SEC, width=14, anchor="w")
        self.flash_lbl.grid(row=0, column=5, padx=(0, 12))

        # 按钮栏
        btn_bar = tk.Frame(self.root, bg=BG_PANEL, height=56)
        btn_bar.grid(row=3, column=0, sticky="ew")
        btn_bar.grid_propagate(False)

        for text, bg, fg, cmd in [
            ("▶  开始烧录", "#0d419d", ACCENT,   self._start_flash),
            ("■  停止烧录", "#4a1515", DANGER,   self._stop_flash),
            ("💾  保存日志", "#0f2d1e", SUCCESS,  self._save_log),
            ("🗑  清空日志", BG_INPUT,  TEXT_SEC, self._clear_log),
            ("↺  重置状态", "#3a2800", WARNING,  self._reset_state),
        ]:
            b = tk.Button(
                btn_bar, text=text, font=FONT_HEAD,
                bg=bg, fg=fg, relief=tk.FLAT,
                activebackground=BORDER, activeforeground=TEXT_PRI,
                cursor="hand2", padx=16, pady=10, command=cmd
            )
            b.pack(side=tk.LEFT, padx=8, pady=8)
            if "开始烧录" in text:
                self.btn_start = b

        # 单独烧录按钮（默认隐藏，有待处理设备时显示）
        self.btn_single = tk.Button(
            btn_bar,
            text="🔧  单独烧录失败设备",
            font=FONT_HEAD,
            bg="#4a2800", fg=WARNING,
            relief=tk.FLAT,
            activebackground=BORDER, activeforeground=TEXT_PRI,
            cursor="hand2", padx=16, pady=10,
            command=self._start_single_flash
        )
        # 初始隐藏
        self.btn_single.pack_forget()

    def _build_left(self, parent):
        parent.grid_columnconfigure(0, weight=1)

        # JLink 卡片
        jc_outer, jc = self._card(parent, "🔌  JLink 设备列表")
        jc_outer.grid(row=0, column=0, sticky="ew", pady=(0, 9))

        top_row = tk.Frame(jc, bg=BG_CARD)
        top_row.pack(fill=tk.X, padx=10, pady=(4, 6))
        self.jlink_count_lbl = tk.Label(
            top_row, text="未检测",
            font=FONT_HEAD, bg=BG_CARD, fg=TEXT_SEC)
        self.jlink_count_lbl.pack(side=tk.LEFT)
        tk.Button(
            top_row, text="刷新检测", font=FONT_SMALL,
            bg=ACCENT2, fg=TEXT_PRI, relief=tk.FLAT,
            padx=10, pady=4, cursor="hand2",
            command=self._detect_jlinks
        ).pack(side=tk.RIGHT)

        self.jlink_listbox = tk.Listbox(
            jc, height=4, font=FONT_MONO,
            bg=BG_INPUT, fg=SUCCESS,
            selectbackground=ACCENT2, selectforeground=TEXT_PRI,
            borderwidth=0, highlightthickness=1,
            highlightcolor=BORDER, highlightbackground=BORDER
        )
        self.jlink_listbox.pack(fill=tk.X, padx=10, pady=(0, 6))

        # 单独烧录队列状态标签
        self.single_queue_lbl = tk.Label(
            jc, text="", font=FONT_SMALL,
            bg=BG_CARD, fg=WARNING, anchor="w"
        )
        self.single_queue_lbl.pack(fill=tk.X, padx=10, pady=(0, 8))

        # 文件配置卡片
        fc_outer, fc = self._card(parent, "📁  文件配置")
        fc_outer.grid(row=1, column=0, sticky="ew", pady=(0, 9))

        self.hex_path   = tk.StringVar()
        self.xlsx_path  = tk.StringVar()
        self.bat_path   = tk.StringVar()
        self.out_dir    = tk.StringVar(value="./output")
        self.base_addr  = tk.StringVar(value="0x178000")
        self.start_row  = tk.StringVar(value="2")   # 自定义起始行

        # 支持任意hex文件名（merged.hex / merged_A.hex 等）
        self._file_row(fc, "输入 hex", self.hex_path,
                       "*.hex",  "选择输入 hex 文件")
        self._file_row(fc, "Excel 数据",  self.xlsx_path,
                       "*.xlsx", "选择 Excel 数据文件")
        self._file_row(fc, "BAT 文件",    self.bat_path,
                       "*.bat",  "选择 ncs BAT 文件")
        self._label_entry_row(fc, "输出目录", self.out_dir,  browse_dir=True)
        self._label_entry_row(fc, "基础地址", self.base_addr)
        self._label_entry_row(fc, "起始行",   self.start_row)

        for var in [self.hex_path, self.xlsx_path, self.bat_path,
                    self.out_dir, self.base_addr, self.start_row]:
            var.trace_add("write", self._auto_save_config)

        self.use_bat = tk.BooleanVar(value=True)
        self.use_bat.trace_add("write", self._auto_save_config)
        chk_frame = tk.Frame(fc, bg=BG_CARD)
        chk_frame.pack(fill=tk.X, padx=10, pady=(6, 4))
        tk.Checkbutton(
            chk_frame,
            text="执行密钥上传（west ncs-provision upload）",
            variable=self.use_bat,
            font=FONT_BODY, bg=BG_CARD, fg=TEXT_PRI,
            selectcolor=BG_INPUT,
            activebackground=BG_CARD, activeforeground=TEXT_PRI
        ).pack(side=tk.LEFT)
        tk.Frame(fc, bg=BG_CARD, height=4).pack()

        # 统计卡片
        sc_outer, sc = self._card(parent, "📊  烧录统计")
        sc_outer.grid(row=2, column=0, sticky="ew", pady=(0, 9))

        self.stat_success = tk.StringVar(value="0")
        self.stat_fail    = tk.StringVar(value="0")
        self.stat_total   = tk.StringVar(value=str(self.state.total_flashed))
        self.stat_row     = tk.StringVar(value=f"第 {self.state.next_row} 行")
        self.stat_prov    = tk.StringVar(value=str(self.state.provisioned_counter))

        for lbl, var, color in [
            ("本次成功",      self.stat_success, SUCCESS),
            ("本次失败",      self.stat_fail,    DANGER),
            ("历史总计成功",  self.stat_total,   ACCENT),
            ("下次 Excel 行", self.stat_row,     TEXT_PRI),
            ("今日已烧录数",  self.stat_prov,    WARNING),
        ]:
            f = tk.Frame(sc, bg=BG_CARD)
            f.pack(fill=tk.X, padx=12, pady=3)
            tk.Label(f, text=lbl, font=FONT_BODY, bg=BG_CARD,
                     fg=TEXT_SEC, width=16, anchor="w").pack(side=tk.LEFT)
            tk.Label(f, textvariable=var, font=FONT_HEAD,
                     bg=BG_CARD, fg=color).pack(side=tk.LEFT)
        tk.Frame(sc, bg=BG_CARD, height=6).pack()

    def _build_right(self, parent):
        lc_outer, lc = self._card(parent, "📋  实时烧录日志")
        lc_outer.grid(row=0, column=0, sticky="nsew")
        lc.grid_rowconfigure(0, weight=1)
        self.log_text = scrolledtext.ScrolledText(
            lc, font=FONT_MONO,
            bg="#080c10", fg=TEXT_PRI,
            insertbackground=ACCENT,
            selectbackground=ACCENT2,
            borderwidth=0, highlightthickness=0,
            wrap=tk.WORD, state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=6, pady=(4, 8))
        for tag, fg, font in [
            ("INFO",     TEXT_SEC,  None),
            ("SUCCESS",  SUCCESS,   None),
            ("ERROR",    DANGER,    None),
            ("WARNING",  WARNING,   None),
            ("HEAD",     ACCENT,    None),
            ("STEP",     "#a5d6ff", None),
            ("SINGLE",   "#ffa657", None),   # 单独烧录专用颜色
            ("CRITICAL", DANGER,    ("Consolas", 9, "bold")),
            # ── 新增标签 ──
            ("HIGHLIGHT", "#ffffff", ("Consolas", 13, "bold")),
            ("HL_OK",     SUCCESS,   ("Consolas", 13, "bold")),
            ("HL_FAIL",   DANGER,    ("Consolas", 13, "bold")),
        ]:
            kw = {"foreground": fg}
            if font:
                kw["font"] = font
            self.log_text.tag_config(tag, **kw)
        # 会话横幅：带背景色的整行色块
        self.log_text.tag_config(
            "BANNER",
            foreground="#ffffff",
            background="#0d3a6e",
            font=("Microsoft YaHei UI", 11, "bold"),
            spacing1=6, spacing3=6,     # 上下留白
        )
        self.log_text.tag_config(
            "BANNER_END",
            foreground="#ffffff",
            background="#1a3a1a",
            font=("Microsoft YaHei UI", 11, "bold"),
            spacing1=6, spacing3=6,
        )
        self.log_text.tag_config(
            "BANNER_FAIL",
            foreground="#ffffff",
            background="#4a1515",
            font=("Microsoft YaHei UI", 11, "bold"),
            spacing1=6, spacing3=6,
        )

    # ──────────────────────────────────────────
    #  UI 辅助
    # ──────────────────────────────────────────
    def _card(self, parent, title: str):
        outer = tk.Frame(parent, bg=BORDER, bd=0)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(0, weight=1)
        wrapper = tk.Frame(outer, bg=BG_CARD, bd=0)
        wrapper.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        wrapper.grid_columnconfigure(0, weight=1)
        wrapper.grid_rowconfigure(1, weight=1)
        tk.Label(wrapper, text=title, font=FONT_HEAD,
                 bg=BG_CARD, fg=ACCENT, anchor="w",
                 padx=12, pady=7).grid(row=0, column=0, sticky="ew")
        tk.Frame(wrapper, bg=BORDER, height=1
                 ).grid(row=0, column=0, sticky="sew")
        content = tk.Frame(wrapper, bg=BG_CARD)
        content.grid(row=1, column=0, sticky="nsew")
        content.grid_columnconfigure(0, weight=1)
        return outer, content

    def _file_row(self, parent, label, var, ftype, title):
        f = tk.Frame(parent, bg=BG_CARD)
        f.pack(fill=tk.X, padx=10, pady=3)
        tk.Label(f, text=label, font=FONT_BODY, bg=BG_CARD,
                 fg=TEXT_SEC, width=11, anchor="w").pack(side=tk.LEFT)
        ext_map = {"*.hex":  [("HEX",   "*.hex")],
                   "*.xlsx": [("Excel", "*.xlsx")],
                   "*.bat":  [("BAT",   "*.bat")]}
        e = tk.Entry(f, textvariable=var, font=FONT_MONO,
                     bg=BG_INPUT, fg=TEXT_PRI, insertbackground=ACCENT,
                     relief=tk.FLAT, bd=4)
        e.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._path_entries.append((var, e))
        ToolTip(e, textvariable=var)
        tk.Button(
            f, text="浏览", font=FONT_SMALL, bg=BG_INPUT, fg=TEXT_SEC,
            relief=tk.FLAT, padx=6, cursor="hand2",
            command=lambda v=var, t=title, ft=ftype: v.set(
                filedialog.askopenfilename(
                    title=t,
                    filetypes=ext_map.get(ft, []) + [("所有文件", "*.*")]
                ) or v.get()
            )
        ).pack(side=tk.LEFT, padx=(4, 0))

    def _label_entry_row(self, parent, label, var, browse_dir=False):
        f = tk.Frame(parent, bg=BG_CARD)
        f.pack(fill=tk.X, padx=10, pady=3)
        tk.Label(f, text=label, font=FONT_BODY, bg=BG_CARD,
                 fg=TEXT_SEC, width=11, anchor="w").pack(side=tk.LEFT)
        tk.Entry(f, textvariable=var, font=FONT_MONO,
                 bg=BG_INPUT, fg=TEXT_PRI, insertbackground=ACCENT,
                 relief=tk.FLAT, bd=4
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        if browse_dir:
            tk.Button(
                f, text="浏览", font=FONT_SMALL, bg=BG_INPUT, fg=TEXT_SEC,
                relief=tk.FLAT, padx=6, cursor="hand2",
                command=lambda: var.set(
                    filedialog.askdirectory(title="选择输出目录") or var.get())
            ).pack(side=tk.LEFT, padx=(4, 0))

    def _set_status(self, text: str, color: str = TEXT_SEC):
        self.status_badge.config(text=f"● {text}", fg=color)
        self.root.update_idletasks()

    def _set_overall_progress(self, current: int, total: int, label: str = ""):
        pct = int(current / total * 100) if total > 0 else 0
        self.overall_bar["value"] = pct
        self.overall_lbl.config(text=label or f"{current}/{total}  {pct}%")
        self.root.update_idletasks()

    def _set_flash_progress(self, value: int, label: str = "",
                            pulse: bool = False):
        if pulse:
            self.flash_bar.config(mode="indeterminate")
            self.flash_bar.start(12)
        else:
            self.flash_bar.stop()
            self.flash_bar.config(mode="determinate")
            self.flash_bar["value"] = value
        self.flash_lbl.config(text=label)
        self.root.update_idletasks()

    def _update_stats(self, success=None, fail=None):
        if success is not None:
            self.stat_success.set(str(success))
        if fail is not None:
            self.stat_fail.set(str(fail))
        self.stat_total.set(str(self.state.total_flashed))
        self.stat_row.set(f"第 {self.state.next_row} 行")
        self.stat_prov.set(str(self.state.provisioned_counter))

    def _update_single_queue_ui(self):
        """更新单独烧录队列的界面状态"""
        if self._single_queue:
            n   = len(self._single_queue)
            ids = ", ".join(j for j, _ in self._single_queue)
            self.single_queue_lbl.config(
                text=f"⚠️  待单独烧录：{n} 个  ({ids})"
            )
            self.btn_single.pack(side=tk.LEFT, padx=8, pady=8)
        else:
            self.single_queue_lbl.config(text="")
            self.btn_single.pack_forget()
        self.root.update_idletasks()

    # ──────────────────────────────────────────
    #  日志
    # ──────────────────────────────────────────
    def log(self, msg: str, level: str = "INFO", full_detail: str = ""):
        ts   = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        self.log_lines.append(line)
        self.log_full.append(line)
        if full_detail:
            for dl in full_detail.splitlines():
                self.log_full.append(f"         {dl}")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, line + "\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def _save_log(self):
        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = os.path.join(get_base_dir(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")],
            initialfile=os.path.join(log_dir, f"flash_log_{ts}.txt")
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(self.log_full))
            self.log(f"日志已保存：{path}", "SUCCESS")

    def _auto_save_log(self):
        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = os.path.join(get_base_dir(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        path = os.path.join(log_dir, f"flash_log_{ts}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.log_full))
        self.log(f"📁 日志已自动保存：{path}", "SUCCESS")

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log_lines.clear()
        self.log_full.clear()

    # ──────────────────────────────────────────
    #  JLink 检测
    # ──────────────────────────────────────────
    def _detect_jlinks(self, silent: bool = False) -> list:
        if not silent:
            self.log("正在检测 JLink 设备...", "INFO")
        try:
            r = subprocess.run(
                ["nrfjprog", "--ids"],
                capture_output=True, text=True, timeout=10,
                creationflags=_SW_FLAGS,
            )
            ids = [x.strip() for x in r.stdout.strip().splitlines()
                   if x.strip()]
            ids.sort()
            self.jlink_ids = ids
            self.jlink_listbox.delete(0, tk.END)
            for i, jid in enumerate(ids, 1):
                self.jlink_listbox.insert(tk.END, f"  #{i}  {jid}")
            if ids:
                self.jlink_count_lbl.config(
                    text=f"检测到 {len(ids)} 个设备", fg=SUCCESS)
                if not silent:
                    self.log(
                        f"✅ 检测到 {len(ids)} 个 JLink：{', '.join(ids)}",
                        "SUCCESS"
                    )
            else:
                self.jlink_count_lbl.config(text="未检测到设备", fg=DANGER)
                if not silent:
                    self.log("⚠️ 未检测到任何 JLink 设备", "WARNING")
            return ids
        except FileNotFoundError:
            if not silent:
                self.log("❌ 未找到 nrfjprog，请确认已安装并加入 PATH", "ERROR")
            return []
        except Exception as e:
            if not silent:
                self.log(f"❌ JLink 检测异常：{e}", "ERROR")
            return []

    # ──────────────────────────────────────────
    #  Excel 读取
    # ──────────────────────────────────────────
    def _read_excel_row(self, row_index: int):
        path = self.xlsx_path.get()
        if not path or not os.path.isfile(path):
            return None
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            if row_index > ws.max_row:
                wb.close()
                return None
            row_data = None
            for row in ws.iter_rows(
                    min_row=row_index, max_row=row_index, values_only=True):
                row_data = row
            wb.close()
            if row_data is None or not any(row_data):
                return None
            return {
                "sn":    str(row_data[0]).strip() if row_data[0] else "",
                "uuid":  str(row_data[1]).strip() if row_data[1] else "",
                "token": str(row_data[2]).strip() if row_data[2] else "",
                "row":   row_index,
            }
        except PermissionError:
            self.root.after(0, lambda: messagebox.showerror(
                "Excel 文件被占用",
                f"无法读取 Excel 文件，请关闭 Excel 后重试：\n{path}"
            ))
            self.log("❌ Excel 文件被占用，请关闭后重试", "ERROR")
            return None
        except Exception as e:
            self.log(f"❌ 读取 Excel 第 {row_index} 行失败：{e}", "ERROR")
            return None

    def _get_excel_remaining_rows(self) -> int:
        path = self.xlsx_path.get()
        if not path or not os.path.isfile(path):
            return 0
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            count = max(0, ws.max_row - self.state.next_row + 1)
            wb.close()
            return count
        except Exception:
            return 0

    # ──────────────────────────────────────────
    #  subprocess 封装
    # ──────────────────────────────────────────
    def _run(self, cmd: list, timeout: int = 60,
             env: dict = None, cwd: str = None) -> dict:
        try:
            r = subprocess.run(
                cmd, capture_output=True, text=True,
                timeout=timeout,
                env=env if env else os.environ.copy(),
                cwd=cwd,
                creationflags=_SW_FLAGS,
            )
            return {"code": r.returncode,
                    "out":  r.stdout.strip(),
                    "err":  r.stderr.strip()}
        except subprocess.TimeoutExpired:
            return {"code": -1, "out": "",
                    "err": f"命令超时（>{timeout}s）"}
        except FileNotFoundError as e:
            return {"code": -2, "out": "", "err": f"命令未找到：{e}"}
        except Exception as e:
            return {"code": -3, "out": "", "err": str(e)}

    def _run_with_progress(self, cmd: list,
                           timeout: int = 120) -> tuple:
        self._set_flash_progress(0, "烧录中...", pulse=True)
        output_lines = []
        got_real_pct = False
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                env=os.environ.copy(),
                creationflags=_SW_FLAGS,
            )
            start = time.time()
            while True:
                if time.time() - start > timeout:
                    proc.kill()
                    self._set_flash_progress(0, "超时 ❌")
                    return (False, f"命令超时（>{timeout}s）",
                            "\n".join(output_lines))
                if self.stop_flag:
                    proc.kill()
                    self._set_flash_progress(0, "已停止")
                    return False, "用户手动停止", "\n".join(output_lines)

                line = proc.stdout.readline()
                if line == "" and proc.poll() is not None:
                    break
                if line:
                    line = line.rstrip()
                    output_lines.append(line)
                    m = re.search(r"\[\s*(\d+)%\]", line)
                    if m:
                        pct = int(m.group(1))
                        got_real_pct = True
                        self._set_flash_progress(pct, f"{pct}%")
                    elif "Verifying" in line and got_real_pct:
                        self._set_flash_progress(90, "校验中...")
                    elif ("Applying system reset" in line
                          or "Reset" in line) and got_real_pct:
                        self._set_flash_progress(98, "复位中...")

            ret      = proc.wait()
            full_out = "\n".join(output_lines)
            if ret == 0:
                self._set_flash_progress(100, "完成 ✅")
                return True, "", full_out
            else:
                self._set_flash_progress(0, "失败 ❌")
                return False, extract_error_summary(full_out), full_out

        except FileNotFoundError as e:
            self._set_flash_progress(0, "失败 ❌")
            return False, f"命令未找到：{e}", ""
        except Exception as e:
            self._set_flash_progress(0, "失败 ❌")
            return False, str(e), ""

    # ──────────────────────────────────────────
    #  单块板子烧录（核心）
    # ──────────────────────────────────────────
    def _flash_one(self, jlink_id: str, data: dict,
                   prov_name: str, prov_file: str) -> tuple:
        """返回 (success, summary_reason, full_reason)"""
        sn_orig  = data["sn"]
        sn_hex   = sn_to_hex32(sn_orig)
        uuid     = data["uuid"]
        token    = data["token"]
        row      = data["row"]
        bat_path = self.bat_path.get()
        bat_dir  = (os.path.dirname(os.path.abspath(bat_path))
                    if bat_path else ".")
        env      = parse_bat_env(bat_path)

        self.log("", "INFO")
        self.log(
            f"  ▶  JLink {jlink_id}   SN: {sn_orig}   (第{row}行)",
            "HIGHLIGHT"
        )
        self.log(f"     UUID: {uuid}  →  {prov_name}.hex", "INFO")

        def step_ok(n, el):
            self.log(f"   {STEP_NAMES[n]}  ✅  {el:.1f}s", "SUCCESS")

        def step_skip(n):
            self.log(f"   {STEP_NAMES[n]}  ⏭  已跳过", "INFO")

        def step_fail(n, el, summary, full):
            self.log(f"   {STEP_NAMES[n]}  ❌  {el:.1f}s  →  {summary}",
                     "ERROR", full_detail=full)

        # 步骤一
        t = time.time()
        r = self._run(["nrfjprog", "--recover", "-s", jlink_id], timeout=60)
        el = time.time() - t
        if r["code"] != 0:
            full = r["err"] or r["out"]
            step_fail(1, el, extract_error_summary(full), full)
            return False, "步骤一失败：芯片解除保护", full
        step_ok(1, el)
        if self.stop_flag:
            return False, "用户手动停止", ""

        # 步骤二
        t = time.time()
        r = self._run(["nrfjprog", "-e", "-s", jlink_id], timeout=60)
        el = time.time() - t
        if r["code"] != 0:
            full = r["err"] or r["out"]
            step_fail(2, el, extract_error_summary(full), full)
            return False, "步骤二失败：擦除芯片内容", full
        step_ok(2, el)
        if self.stop_flag:
            return False, "用户手动停止", ""

        # 步骤三（可选）
        if self.use_bat.get():
            if not bat_path or not os.path.isfile(bat_path):
                msg = "未指定有效 BAT 文件"
                self.log(f"   {STEP_NAMES[3]}  ❌  {msg}", "ERROR")
                return False, f"步骤三失败：{msg}", msg

            pem_file = os.path.join(
                bat_dir, "configuration",
                "nrf54l15dk_nrf54l15_cpuapp",
                "boot_signature_key_file_ed25519.pem"
            )
            tool_chain_dir = env.get("TOOL_CHAIN_DIR", "")
            ncs_python = os.path.join(
                tool_chain_dir, "opt", "bin", "python.exe"
            ) if tool_chain_dir else ""

            if not ncs_python or not os.path.isfile(ncs_python):
                ncs_python = ""
                for p in env.get("PATH", "").split(os.pathsep):
                    c = os.path.join(p, "python.exe")
                    if os.path.isfile(c) and "ncs" in c.lower():
                        ncs_python = c
                        break

            if not ncs_python or not os.path.isfile(ncs_python):
                msg = (f"找不到 NCS python.exe "
                       f"(TOOL_CHAIN_DIR='{tool_chain_dir}')")
                self.log(f"   {STEP_NAMES[3]}  ❌  {msg}", "ERROR")
                return False, f"步骤三失败：{msg}", msg

            t = time.time()
            r = self._run(
                [ncs_python, "-m", "west", "ncs-provision", "upload",
                 "-k", pem_file, "--keyname", "UROT_PUBKEY",
                 "-s", "nrf54l15", "--dev-id", jlink_id],
                timeout=90, env=env, cwd=bat_dir
            )
            el = time.time() - t
            if r["code"] != 0:
                full = r["err"] or r["out"]
                step_fail(3, el, extract_error_summary(full), full)
                return False, "步骤三失败：密钥上传", full
            step_ok(3, el)
        else:
            step_skip(3)
        if self.stop_flag:
            return False, "用户手动停止", ""

        # 步骤四
        if os.path.isfile(prov_file):
            os.remove(prov_file)
        t = time.time()
        r = self._run([
            "ncsfmntools", "provision",
            "--device", "NRF54L15",
            "--serial-number", sn_hex,
            "--mfi-token", token,
            "--mfi-uuid", uuid,
            "--input-hex-file", os.path.abspath(self.hex_path.get()),
            "--settings-base", self.base_addr.get(),
            "--output-path", prov_file,
        ], timeout=60)
        el = time.time() - t
        if r["code"] != 0:
            full = r["err"] or r["out"]
            step_fail(4, el, extract_error_summary(full), full)
            return False, "步骤四失败：生成烧录文件", full
        if not os.path.isfile(prov_file):
            msg = f"{prov_name}.hex 未生成"
            self.log(f"   {STEP_NAMES[4]}  ❌  {msg}", "ERROR")
            return False, f"步骤四失败：{msg}", msg
        step_ok(4, el)
        if self.stop_flag:
            return False, "用户手动停止", ""

        # 步骤五
        t = time.time()
        ok, err_summary, full_out = self._run_with_progress(
            ["nrfjprog", "--program", prov_file,
             "--verify", "--reset", "-s", jlink_id],
            timeout=120
        )
        el = time.time() - t
        if not ok:
            step_fail(5, el, err_summary, full_out)
            return False, "步骤五失败：烧录芯片", full_out
        step_ok(5, el)
        return True, "", ""

    # ──────────────────────────────────────────
    #  烧录控制
    # ──────────────────────────────────────────
    def _start_flash(self):
        if self.is_flashing:
            return

        # 自动刷新 JLink
        self.log("正在检测 JLink 设备...", "INFO")
        fresh_ids = self._detect_jlinks(silent=True)
        if fresh_ids:
            self.log(
                f"✅ 检测到 {len(fresh_ids)} 个 JLink：{', '.join(fresh_ids)}",
                "SUCCESS"
            )
        else:
            messagebox.showerror("错误", "未检测到 JLink 设备，请检查连接！")
            return

        if not self.hex_path.get() or not os.path.isfile(self.hex_path.get()):
            messagebox.showerror("错误", "输入 hex 文件不存在，请重新选择！")
            return
        if not self.xlsx_path.get() or not os.path.isfile(self.xlsx_path.get()):
            messagebox.showerror("错误", "Excel 数据文件不存在，请重新选择！")
            return

        # 解析并应用自定义起始行
        try:
            start_row_val = int(self.start_row.get().strip())
            if start_row_val < 2:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误",
                                 "起始行必须是大于等于2的整数（第1行为标题行）！")
            return
        # 用起始行覆盖 next_row（仅本次烧录生效，不永久改变state）
        self._session_start_row = start_row_val
        self.state.next_row = start_row_val

        remaining = self._get_excel_remaining_rows()
        n_jlinks  = len(fresh_ids)
        if remaining == 0:
            messagebox.showerror(
                "错误",
                f"Excel 数据已用完（从第 {self.state.next_row} 行起无数据）。\n"
                f"请补充数据后再烧录。"
            )
            return
        if remaining < n_jlinks:
            if not messagebox.askyesno(
                "数据不足提醒",
                f"Excel 剩余 {remaining} 行数据，JLink 有 {n_jlinks} 个。\n\n"
                f"数据不足，本轮将只烧录前 {remaining} 块板子，\n"
                f"剩余 {n_jlinks - remaining} 个 JLink 将自动跳过。\n\n"
                f"是否继续？"
            ):
                return

        if not messagebox.askyesno(
            "确认烧录",
            f"即将串行烧录，共检测到 {n_jlinks} 个 JLink：\n"
            f"{chr(10).join(fresh_ids)}\n\n"
            f"从 Excel 第 {self.state.next_row} 行开始\n"
            f"输入 hex：{os.path.basename(self.hex_path.get())}\n\n"
            f"确认开始？"
        ):
            return

        self.is_flashing = True
        self.stop_flag   = False
        self._set_status("烧录中...", WARNING)
        if self.btn_start:
            self.btn_start.config(
                state=tk.DISABLED, bg="#1a1a2e", fg=TEXT_DIM)
        threading.Thread(target=self._flash_thread, daemon=True).start()

    def _stop_flash(self):
        if self.is_flashing:
            self.stop_flag   = True
            self.is_flashing = False
            self._set_status("已停止", DANGER)
            self.log("\n⛔ 已停止烧录", "CRITICAL")
            self._set_overall_progress(0, 1, "已停止")
            self._set_flash_progress(0, "已停止")
            if self.btn_start:
                self.btn_start.config(
                    state=tk.NORMAL, bg="#0d419d", fg=ACCENT)
        else:
            self.log("当前未在烧录中", "INFO")

    # ──────────────────────────────────────────
    #  烧录主线程
    # ──────────────────────────────────────────
    def _flash_thread(self):
        total_jlinks  = len(self.jlink_ids)
        success_count = 0
        fail_count    = 0
        failed_jobs   = []
        remaining_cap = self._get_excel_remaining_rows()

        self.state.mark_start()

        today       = datetime.now().strftime("%Y%m%d")
        base_output = os.path.abspath(self.out_dir.get() or "./output")
        day_output  = os.path.join(base_output, today)
        os.makedirs(day_output, exist_ok=True)

        self._set_overall_progress(0, total_jlinks, f"0/{total_jlinks} 块")

        ts_now = datetime.now().strftime("%H:%M:%S")
        self.log(
            f"  🚀  开始烧录  {total_jlinks} 块板子  "
            f"Excel 第 {self.state.next_row} 行起  ──  {ts_now}",
            "BANNER"
        )
        self.log(
            f"   JLink：{', '.join(self.jlink_ids)}  |  输出：{day_output}",
            "INFO"
        )
        for idx, jlink_id in enumerate(self.jlink_ids):
            if self.stop_flag:
                self.log("⛔ 已停止，以下 JLink 未烧录：", "CRITICAL")
                for j in self.jlink_ids[idx:]:
                    self.log(f"   🔴 {j}", "CRITICAL")
                break

            if idx >= remaining_cap:
                skipped = self.jlink_ids[idx:]
                self.log(
                    f"\n📋 Excel 数据已全部用完，"
                    f"跳过剩余 {len(skipped)} 个 JLink：",
                    "WARNING"
                )
                for j in skipped:
                    self.log(f"   ⏭  JLink {j}  已跳过", "INFO")
                break

            row_index = self.state.next_row
            data = self._read_excel_row(row_index)
            if data is None:
                self.log(f"❌ Excel 第 {row_index} 行无数据，停止烧录", "ERROR")
                for j in self.jlink_ids[idx:]:
                    failed_jobs.append((j, "Excel 数据耗尽"))
                    fail_count += 1
                break

            _, _, prov_stem = self.state.next_provisioned_name()
            prov_file = os.path.join(day_output, f"{prov_stem}.hex")

            self._set_overall_progress(
                idx, total_jlinks,
                f"第一轮 {idx+1}/{total_jlinks}  {jlink_id}"
            )
            self._update_stats(success_count, fail_count)

            ok, summary, full = self._flash_one(
                jlink_id, data, prov_stem, prov_file)

            if ok:
                success_count += 1
                self.state.next_row += 1
                self.state.total_flashed += 1
                self.log(
                    f"  ✅  JLink {jlink_id}  SN: {data['sn']}",
                    "HL_OK"
                )
            else:
                fail_count += 1
                self.log(
                    f"  ❌  JLink {jlink_id}  {summary}  → 进入补烧",
                    "HL_FAIL"
                )
                failed_jobs.append((jlink_id, summary))

            self._update_stats(success_count, fail_count)

        # ── 补烧轮 ──
        retry_success = []
        retry_fail    = []

        if failed_jobs and not self.stop_flag:
            fail_list = "\n".join(
                [f"  • JLink {jid}：{s}" for jid, s in failed_jobs])
            self.log(
                f"   ⏸  第一轮结束，{len(failed_jobs)} 块失败，等待确认补烧...",
                "WARNING"
            )

            confirm_event  = threading.Event()
            confirm_result = [False]

            def ask_confirm():
                result = messagebox.askyesno(
                    "确认开始补烧",
                    f"第一轮烧录结束，以下 {len(failed_jobs)} 块板子失败：\n\n"
                    f"{fail_list}\n\n"
                    f"请调整好板子后点击【是】开始补烧，\n"
                    f"点击【否】跳过补烧直接结束。"
                )
                confirm_result[0] = result
                confirm_event.set()

            self.root.after(0, ask_confirm)
            confirm_event.wait()

            if not confirm_result[0]:
                self.log("⏭  已跳过补烧轮，直接结束", "WARNING")
                retry_fail = list(failed_jobs)
            else:
                self.log(
                    f"  🔁  补烧轮开始  共 {len(failed_jobs)} 块",
                    "BANNER"
                )

                for r_idx, (jlink_id, first_summary) in enumerate(
                        failed_jobs):
                    if self.stop_flag:
                        done    = set(retry_success) | {
                            j for j, _ in retry_fail}
                        pending = [j for j, _ in failed_jobs
                                   if j not in done]
                        self.log("⛔ 补烧中途停止，以下未完成：", "CRITICAL")
                        for j in pending:
                            self.log(
                                f"   ⛔ JLink {j} ← 需人工介入", "CRITICAL")
                        break

                    row_index = self.state.next_row
                    data = self._read_excel_row(row_index)
                    if data is None:
                        self.log(
                            f"❌ 补烧 JLink {jlink_id}：Excel 数据耗尽",
                            "ERROR"
                        )
                        retry_fail.append((jlink_id, "Excel 数据耗尽"))
                        fail_count += 1
                        self._update_stats(success_count, fail_count)
                        continue

                    _, _, prov_stem = self.state.next_provisioned_name()
                    prov_file = os.path.join(
                        day_output, f"{prov_stem}.hex")

                    self._set_overall_progress(
                        total_jlinks + r_idx,
                        total_jlinks + len(failed_jobs),
                        f"补烧 {r_idx+1}/{len(failed_jobs)}  {jlink_id}"
                    )
                    self.log(
                        f"\n🔁 [{r_idx+1}/{len(failed_jobs)}]"
                        f" 补烧 JLink {jlink_id}"
                        f"  （首轮：{first_summary}）",
                        "WARNING"
                    )

                    ok, summary, full = self._flash_one(
                        jlink_id, data, prov_stem, prov_file)

                    if ok:
                        retry_success.append(jlink_id)
                        success_count += 1
                        self.state.next_row += 1
                        self.state.total_flashed += 1
                        self.log(
                            f"  ✅  补烧成功  JLink {jlink_id}"
                            f"  SN: {data['sn']}",
                            "HL_OK"
                        )
                    else:
                        retry_fail.append((jlink_id, summary))
                        fail_count += 1
                        self.log(
                            f"  ❌  补烧失败  JLink {jlink_id}"
                            f"  {summary}",
                            "HL_FAIL"
                        )

                    self._update_stats(success_count, fail_count)

                # 补烧汇总（结果已在每块板子的 HL_OK/HL_FAIL 中醒目显示）
                if retry_fail:
                    self.log(
                        f"   补烧仍有 {len(retry_fail)} 块失败"
                        f" → 自动进入单独烧录队列",
                        "WARNING"
                    )

        # ── 补烧失败 → 自动加入单独烧录队列 ──
        if retry_fail:
            self._single_queue = list(retry_fail)
            self.root.after(0, self._update_single_queue_ui)
            self.log(
                f"\n🔧 {len(retry_fail)} 个 JLink 补烧失败，"
                f"已加入单独烧录队列，请调整板子后点击【单独烧录失败设备】",
                "SINGLE"
            )

        # ── 结束横幅 ──
        ts_end = datetime.now().strftime("%H:%M:%S")
        banner_tag = "BANNER_END" if fail_count == 0 else "BANNER_FAIL"
        self.log(
            f"  🏁  烧录结束  成功 {success_count}  失败 {fail_count}"
            f"  累计 {self.state.total_flashed}"
            f"  下次第{self.state.next_row}行  ──  {ts_end}",
            banner_tag
        )
        # 同步 UI 起始行为当前 next_row，避免下次点击重复烧录
        self.root.after(0, lambda r=self.state.next_row: self.start_row.set(str(r)))
        self._update_stats(success_count, fail_count)

        self._set_overall_progress(
            total_jlinks, total_jlinks, "全部完成 ✅")
        self._set_flash_progress(0, "空闲")
        self.is_flashing = False
        self.state.mark_end()

        if self.btn_start:
            self.btn_start.config(
                state=tk.NORMAL, bg="#0d419d", fg=ACCENT)

        self._auto_save_log()

        # 更新状态栏（不再弹出全屏遮罩）
        if fail_count == 0:
            self._set_status("上次全部成功 ✅", SUCCESS)
        else:
            self._set_status(f"上次有 {fail_count} 块失败 ❌", DANGER)

    # ──────────────────────────────────────────
    #  单独烧录
    # ──────────────────────────────────────────
    def _start_single_flash(self):
        """点击单独烧录按钮，逐一串行处理队列中的失败设备"""
        if self.is_flashing:
            self.log("⚠️ 烧录进行中，请等待完成", "WARNING")
            return
        if not self._single_queue:
            return

        self.is_flashing = True
        self.stop_flag   = False
        if self.btn_start:
            self.btn_start.config(
                state=tk.DISABLED, bg="#1a1a2e", fg=TEXT_DIM)
        self.btn_single.config(
            state=tk.DISABLED, bg="#2a1800", fg=TEXT_DIM)

        threading.Thread(
            target=self._single_flash_thread, daemon=True).start()

    def _single_flash_thread(self):
        today       = datetime.now().strftime("%Y%m%d")
        base_output = os.path.abspath(self.out_dir.get() or "./output")
        day_output  = os.path.join(base_output, today)
        os.makedirs(day_output, exist_ok=True)

        total         = len(self._single_queue)
        single_success = []
        single_fail    = []

        ts_now = datetime.now().strftime("%H:%M:%S")
        self.log(
            f"  🔧  单独烧录开始  共 {total} 个设备  ──  {ts_now}",
            "BANNER"
        )

        # 逐一串行处理，每次都用同一个 next_row（失败不推进）
        queue_snapshot = list(self._single_queue)
        for s_idx, (jlink_id, prev_reason) in enumerate(queue_snapshot):
            if self.stop_flag:
                self.log("⛔ 单独烧录已停止", "CRITICAL")
                # 未处理的重新放回队列
                remaining = queue_snapshot[s_idx:]
                self._single_queue = remaining
                self.root.after(0, self._update_single_queue_ui)
                break

            self._set_overall_progress(
                s_idx, total,
                f"单独烧录 {s_idx+1}/{total}  {jlink_id}"
            )
            self._set_status(f"单独烧录 {jlink_id}...", WARNING)

            self.log(
                f"\n🔧 [{s_idx+1}/{total}] 单独烧录 JLink {jlink_id}"
                f"  （原因：{prev_reason}）",
                "SINGLE"
            )

            # 每次都用 next_row，失败不推进
            row_index = self.state.next_row
            data = self._read_excel_row(row_index)
            if data is None:
                self.log(
                    f"❌ 单独烧录 JLink {jlink_id}：Excel 数据耗尽",
                    "ERROR"
                )
                single_fail.append((jlink_id, "Excel 数据耗尽"))
                continue

            _, _, prov_stem = self.state.next_provisioned_name()
            prov_file = os.path.join(day_output, f"{prov_stem}.hex")

            ok, summary, full = self._flash_one(
                jlink_id, data, prov_stem, prov_file)

            if ok:
                single_success.append(jlink_id)
                self.state.next_row += 1
                self.state.total_flashed += 1
                self._update_stats()
                self.log(
                    f"  ✅  JLink {jlink_id}  SN: {data['sn']}",
                    "HL_OK"
                )
            else:
                single_fail.append((jlink_id, summary))
                self.log(
                    f"  ❌  JLink {jlink_id}  {summary}",
                    "HL_FAIL"
                )

        # 单独烧录汇总横幅
        ts_end = datetime.now().strftime("%H:%M:%S")
        s_ok  = len(single_success)
        s_nok = len(single_fail)
        banner_tag = "BANNER_END" if s_nok == 0 else "BANNER_FAIL"
        self.log(
            f"  🔧  单独烧录结束  成功 {s_ok}  失败 {s_nok}  ──  {ts_end}",
            banner_tag
        )

        # 更新队列：仍失败的留在队列里，成功的移除
        self._single_queue = list(single_fail)
        self.root.after(0, self._update_single_queue_ui)
        # 同步 UI 起始行为当前 next_row，避免下次点击重复烧录
        self.root.after(0, lambda r=self.state.next_row: self.start_row.set(str(r)))
        self.is_flashing = False

        self._set_flash_progress(0, "空闲")
        self._set_overall_progress(total, total, "单独烧录完成")
        self.state.mark_end()

        if self.btn_start:
            self.btn_start.config(
                state=tk.NORMAL, bg="#0d419d", fg=ACCENT)
        self.btn_single.config(
            state=tk.NORMAL, bg="#4a2800", fg=WARNING)

        self._auto_save_log()

        # 更新状态栏（不再弹出全屏遮罩）
        if not self._single_queue:
            self._set_status(
                "就绪" if s_nok == 0 else f"有 {s_nok} 块仍失败",
                SUCCESS if s_nok == 0 else DANGER
            )
        else:
            self._set_status(
                f"还有 {len(self._single_queue)} 个设备待处理", WARNING)

    # ──────────────────────────────────────────
    #  重置（需输入 RESET）
    # ──────────────────────────────────────────
    def _reset_state(self):
        if self.is_flashing:
            messagebox.showwarning("警告", "烧录进行中，无法重置！")
            return
        ans = simpledialog.askstring(
            "安全确认",
            "此操作将重置 Excel 行指针至第2行，\n"
            "清空今日计数及历史累计，路径配置保留。\n\n"
            "请输入  RESET  确认：",
            parent=self.root
        )
        if ans and ans.strip() == "RESET":
            cfg_keys = ["cfg_hex", "cfg_xlsx", "cfg_bat",
                        "cfg_out_dir", "cfg_base_addr", "cfg_use_bat",
                        "cfg_start_row"]
            saved_cfg = {k: self.state.data.get(k, "") for k in cfg_keys}
            self.state.data = {
                "next_row":        2,
                "total_flashed":   0,
                "daily_counter":   {},
                "last_exit_clean": True,
            }
            self.state.data.update(saved_cfg)
            self.state.save()
            self._single_queue.clear()
            self.root.after(0, self._update_single_queue_ui)
            self.start_row.set("2")
            self._update_stats(0, 0)
            self._set_overall_progress(0, 1, "等待开始")
            self._set_flash_progress(0, "空闲")
            self._set_status("就绪", SUCCESS)
            self.log("🔄 状态已重置，从 Excel 第 2 行重新开始", "WARNING")
        elif ans is not None:
            messagebox.showwarning("取消", "输入不正确，重置已取消。")


# ─────────────────────────────────────────────
#  程序入口
# ─────────────────────────────────────────────
def main():
    root = tk.Tk()
    NordicFlashApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()