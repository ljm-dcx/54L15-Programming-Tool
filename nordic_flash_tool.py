# -*- coding: utf-8 -*-
"""
nRF54L15 自动烧录工具 v8.3（并行烧录版）
在 v7.1 基础上新增：
  1. 可配置 JLink 烧录速度（nrfjprog --clockspeed，1-8MHz，界面输入）
  2. 六槽位并行烧录：每轮开始前一次性、原子性地给所有探针分配好 Excel 数据
     （行号"分配游标"在分配那一刻就前移，不再等烧录成功才前移）
  3. 全新界面：6 槽位实时状态网格 + 按探针上色的日志 + 整批进度条
     （有失败时进度条变红，不会显示成容易误读的满格绿色）
  4. 补烧改为两级：本轮失败 → 询问是否"原地立即重试" → 仍失败则进入持久化的
     "待补烧队列"（保存在 flash_state.json，重启不丢失），行号不会被后续批次复用
  5. 面向产线操作员简化实时日志：只显示"当前烧的是哪块板子"和"成功/失败"，
     每一小步的详细耗时/报错只写进自动保存的日志文件，供工程师排查
v8.1 修复：nrfjprog 速度参数名写错（应为 --clockspeed，不是 --speed），
    实测中会导致 100% "invalid argument" 报错，与并行无关
v8.2 修复：步骤三密钥文件路径只认定"BAT 文件放在项目根目录"一种情况，
    BAT 文件若放在 configuration/ 目录里会拼出多一层的错误路径导致
    FileNotFoundError；现在两种位置都会尝试，并给出更清楚的报错。
v8.3 改动：
  1. "开始并行烧录"的确认弹窗改为主界面内嵌确认条（不再弹出系统对话框），
     高亮显示"确认并开始"按钮 + 槽位/SN 预览，交互动作不变，去掉弹窗抢焦点。
  2. 生成的烧录用 hex 文件（provisioned_xxx.hex）不再保留：每片板子烧录
     （无论成功失败）结束后立即删除，只作为烧录当下的临时文件使用，不再
     占用 output 目录空间；每片的 SN/UUID/行号仍完整记录在自动保存的日志里。
  3. 左侧"烧录统计"改为紧凑的方块网格布局，并给整个左侧栏包了可滚动容器，
     小屏幕电脑也能完整看到所有卡片内容（鼠标滚轮/滚动条均可滚动）。
  4. 本轮烧录完成后，日志"本批烧录结束"横幅里增加本轮实际用时。
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
import queue
import openpyxl
from datetime import datetime

# Windows 下隐藏 subprocess 弹出的命令行窗口
_SW_FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

MAX_SLOTS = 6


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

# 每个槽位（探针）在日志里的固定标识色，方便在并行滚动日志里区分是哪片板子
SLOT_COLORS = ["#58a6ff", "#ffa657", "#3fb950",
               "#f778ba", "#d29922", "#79c0ff"]

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


def valid_speed(s: str):
    """校验速度字符串，合法返回 int(kHz)，否则返回 None"""
    try:
        v = int(str(s).strip())
    except (TypeError, ValueError):
        return None
    if 1000 <= v <= 8000:
        return v
    return None


# ─────────────────────────────────────────────
#  持久化状态
# ─────────────────────────────────────────────
class FlashState:
    """
    行号分配模型（v8 起变更）：
      next_row 现在表示"分配游标"——下一个从未被分配过的 Excel 行号。
      claim_rows(n) 在一轮烧录**开始前**一次性把 n 个行号占用掉，
      游标立即前移，不等这些板子是否烧录成功。
      这样保证同一行数据（同一个 SN/UUID/Token）永远只会被分配给一块板子一次，
      并行、乱序完成都不会产生"两片板子用同一行数据"的竞态。

      分配出去但烧录失败的行，进 retry_queue（持久化在 json 里），
      由"处理待补烧设备"单独找机会补烧，不会被后续新一批的 claim_rows 复用。
    """
    def __init__(self):
        self.state_file = os.path.join(get_base_dir(), "flash_state.json")
        self.lock = threading.RLock()
        self.data = {
            "next_row":        2,
            "total_flashed":   0,
            "daily_counter":   {},
            "last_exit_clean": True,
            "retry_queue":     [],   # [{row, reason, fail_count}]
            "cfg_hex":         "",
            "cfg_xlsx":        "",
            "cfg_bat":         "",
            "cfg_out_dir":     "./output",
            "cfg_base_addr":   "0x178000",
            "cfg_use_bat":     True,
            "cfg_start_row":   "2",
            "cfg_speed_khz":   "4000",
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
                    out_dir, base_addr, use_bat, start_row="2",
                    speed_khz="4000"):
        with self.lock:
            self.data.update({
                "cfg_hex":       hex_path,
                "cfg_xlsx":      xlsx_path,
                "cfg_bat":       bat_path,
                "cfg_out_dir":   out_dir,
                "cfg_base_addr": base_addr,
                "cfg_use_bat":   use_bat,
                "cfg_start_row": start_row,
                "cfg_speed_khz": speed_khz,
            })
            self.save()

    def mark_start(self):
        with self.lock:
            self.data["last_exit_clean"] = False
            self.save()

    def mark_end(self):
        with self.lock:
            self.data["last_exit_clean"] = True
            self.save()

    @property
    def last_exit_clean(self):
        return self.data.get("last_exit_clean", True)

    @property
    def next_row(self):
        with self.lock:
            return self.data["next_row"]

    def set_next_row(self, v: int):
        """手动跳转分配游标（对应界面上的"起始行"覆盖功能）"""
        with self.lock:
            self.data["next_row"] = v
            self.save()

    def claim_rows(self, n: int):
        """原子性占用接下来最多 n 个行号，游标立即前移。返回占用到的行号列表。"""
        with self.lock:
            start = self.data["next_row"]
            rows = list(range(start, start + n))
            self.data["next_row"] = start + n
            self.save()
            return rows

    @property
    def total_flashed(self):
        with self.lock:
            return self.data["total_flashed"]

    def add_flashed(self, n: int = 1):
        with self.lock:
            self.data["total_flashed"] += n
            self.save()

    def next_provisioned_name(self) -> tuple:
        with self.lock:
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

    # ── 待补烧队列（持久化，重启不丢失）──
    def add_retry(self, row: int, reason: str):
        with self.lock:
            q = self.data.setdefault("retry_queue", [])
            for item in q:
                if item["row"] == row:
                    item["reason"] = reason
                    item["fail_count"] = item.get("fail_count", 0) + 1
                    self.save()
                    return
            q.append({"row": row, "reason": reason, "fail_count": 1})
            self.save()

    def remove_retry(self, row: int):
        with self.lock:
            q = self.data.setdefault("retry_queue", [])
            self.data["retry_queue"] = [x for x in q if x["row"] != row]
            self.save()

    def get_retry_list(self) -> list:
        with self.lock:
            return [dict(x) for x in self.data.get("retry_queue", [])]

    def reset_all(self, keep_cfg_keys):
        with self.lock:
            saved_cfg = {k: self.data.get(k, "") for k in keep_cfg_keys}
            self.data = {
                "next_row":        2,
                "total_flashed":   0,
                "daily_counter":   {},
                "last_exit_clean": True,
                "retry_queue":     [],
            }
            self.data.update(saved_cfg)
            self.save()


# ─────────────────────────────────────────────
#  主应用
# ─────────────────────────────────────────────
class NordicFlashApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("nRF54L15 自动烧录工具  v8.3  (并行版)")
        self.root.geometry("1320x940")
        self.root.configure(bg=BG_DARK)
        self.root.resizable(True, True)

        self.state        = FlashState()
        self.jlink_ids    = []
        self.is_flashing  = False
        self.stop_flag    = False
        self.log_lines    = []
        self.log_full     = []
        self.btn_start    = None
        self.btn_single   = None

        # 并行状态：jlink_id -> {sn,row,state,pct,detail}
        self.slot_status     = {}
        self.status_lock     = threading.Lock()
        self.log_queue       = queue.Queue()
        self._round_slot_map = {}   # jlink_id -> 0..5，当前网格显示用的槽位映射
        self.slot_frames     = []

        # 内嵌确认条待确认的这一轮数据（点【开始并行烧录】校验通过后先存这里，
        # 等操作员点了确认条里的【确认并开始】才真正启动线程）
        self._pending_pairs  = None
        self._pending_cfg    = None

        self._build_ui()
        self._load_saved_config()
        self._detect_jlinks()
        self._check_last_exit()
        self._refresh_manual_queue_ui()

        self.root.after(150, self._ui_tick)

    # ──────────────────────────────────────────
    #  启动检测
    # ──────────────────────────────────────────
    def _check_last_exit(self):
        if not self.state.last_exit_clean:
            self.root.after(600, lambda: messagebox.showwarning(
                "⚠️ 检测到上次异常退出",
                f"上次烧录程序可能未正常结束（断电或崩溃）。\n\n"
                f"当前分配游标：第 {self.state.next_row} 行\n\n"
                f"请人工核查最近几行数据对应的板子是否已成功烧录，\n"
                f"待补烧队列里的行也请一并核实，确认后再继续。"
            ))
            self._do_log("⚠️ 检测到上次异常退出，请核查最近数据是否烧录成功",
                          "WARNING")

    # ──────────────────────────────────────────
    #  路径 / 参数配置持久化
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
            (self.speed_var, "cfg_speed_khz", "4000"),
        ]:
            var.set(d.get(key, default))
        self.use_bat.set(d.get("cfg_use_bat", True))
        self.root.after(200, self._validate_paths)
        self.root.after(200, self._validate_speed)

    def _validate_paths(self):
        for var, entry_widget in self._path_entries:
            path = var.get()
            if path and not os.path.exists(path):
                entry_widget.config(bg="#3a1010")
            else:
                entry_widget.config(bg=BG_INPUT)

    def _validate_speed(self, *_):
        ok = valid_speed(self.speed_var.get()) is not None
        self.speed_entry.config(bg=BG_INPUT if ok else "#3a1010")
        return ok

    def _auto_save_config(self, *_):
        self.state.save_config(
            self.hex_path.get(), self.xlsx_path.get(),
            self.bat_path.get(), self.out_dir.get(),
            self.base_addr.get(), self.use_bat.get(),
            self.start_row.get(), self.speed_var.get(),
        )
        self._validate_paths()
        self._validate_speed()

    def _capture_cfg(self):
        """在主线程一次性把界面上的所有配置读成一份普通 dict，
        之后并行的工作线程只用这份 dict，不再触碰任何 Tkinter 变量。"""
        speed = valid_speed(self.speed_var.get())
        if speed is None:
            messagebox.showerror(
                "错误", "烧录速度必须是 1000–8000 之间的整数（单位 kHz，即 1–8MHz）")
            return None
        return {
            "hex_path":  self.hex_path.get(),
            "bat_path":  self.bat_path.get(),
            "out_dir":   self.out_dir.get(),
            "base_addr": self.base_addr.get(),
            "use_bat":   self.use_bat.get(),
            "speed_khz": speed,
        }

    # ──────────────────────────────────────────
    #  界面构建
    # ──────────────────────────────────────────
    def _setup_progressbar_styles(self):
        """给进度条按状态上色（空闲/进行中/成功/失败），
        这样"整批进度"条不会在全部失败时还显示成一条容易被误读的绿条。"""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        for name, color in [
            ("Idle.Horizontal.TProgressbar",    BORDER),
            ("Running.Horizontal.TProgressbar", ACCENT),
            ("Success.Horizontal.TProgressbar", SUCCESS),
            ("Fail.Horizontal.TProgressbar",    DANGER),
        ]:
            style.configure(name, troughcolor=BG_INPUT, background=color,
                             borderwidth=0, lightcolor=color, darkcolor=color)

    def _build_ui(self):
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_rowconfigure(3, weight=0)
        self.root.grid_rowconfigure(4, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        title_bar = tk.Frame(self.root, bg=BG_PANEL, height=60)
        title_bar.grid(row=0, column=0, sticky="ew")
        title_bar.grid_propagate(False)
        title_bar.grid_columnconfigure(0, weight=1)
        tk.Label(
            title_bar, text="⚡  nRF54L15  自动烧录工具  (并行版)",
            font=FONT_TITLE, bg=BG_PANEL, fg=ACCENT
        ).grid(row=0, column=0, sticky="w", padx=22, pady=12)
        self.status_badge = tk.Label(
            title_bar, text="● 就绪",
            font=FONT_HEAD, bg=BG_PANEL, fg=SUCCESS)
        self.status_badge.grid(row=0, column=2, sticky="e", padx=22)
        tk.Label(
            title_bar, text="v8.3",
            font=FONT_SMALL, bg=BG_PANEL, fg=TEXT_DIM
        ).grid(row=0, column=1, sticky="e", padx=4)

        body = tk.Frame(self.root, bg=BG_DARK)
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=8)
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=0, minsize=420)
        body.grid_columnconfigure(1, weight=1)

        left = tk.Frame(body, bg=BG_DARK)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(0, weight=1)
        right = tk.Frame(body, bg=BG_DARK)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(0, weight=0)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        self._path_entries = []
        self._build_left(left)
        self._build_right(right)

        # 整批进度条
        prog_frame = tk.Frame(self.root, bg=BG_PANEL, height=32)
        prog_frame.grid(row=2, column=0, sticky="ew")
        prog_frame.grid_propagate(False)
        prog_frame.grid_columnconfigure(1, weight=1)

        tk.Label(prog_frame, text="本批进度", font=FONT_SMALL,
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

        # 内嵌确认条（原来"确认并行烧录"的系统弹窗改成这里，不弹独立对话框）
        # 默认不显示，点击【开始并行烧录】校验通过后才 grid() 出来
        self.confirm_bar = tk.Frame(self.root, bg="#2d2200", height=1)
        self.confirm_bar.grid(row=3, column=0, sticky="ew")
        self.confirm_bar.grid_columnconfigure(0, weight=1)
        self.confirm_bar.grid_remove()

        cb_inner = tk.Frame(self.confirm_bar, bg="#2d2200")
        cb_inner.grid(row=0, column=0, sticky="ew", padx=16, pady=10)
        cb_inner.grid_columnconfigure(0, weight=1)

        self.confirm_warn_lbl = tk.Label(
            cb_inner, text="", font=FONT_BODY,
            bg="#2d2200", fg=WARNING, justify=tk.LEFT, anchor="w")
        self.confirm_warn_lbl.grid(row=0, column=0, sticky="ew")

        self.confirm_preview_lbl = tk.Label(
            cb_inner, text="", font=FONT_MONO,
            bg="#2d2200", fg=TEXT_PRI, justify=tk.LEFT, anchor="w")
        self.confirm_preview_lbl.grid(row=1, column=0, sticky="ew", pady=(2, 0))

        cb_btns = tk.Frame(cb_inner, bg="#2d2200")
        cb_btns.grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.btn_confirm_start = tk.Button(
            cb_btns, text="✅  确认并开始", font=FONT_HEAD,
            bg=SUCCESS, fg="#0d1117", relief=tk.FLAT,
            activebackground=SUCCESS, activeforeground="#0d1117",
            cursor="hand2", padx=18, pady=8,
            command=self._confirm_start_flash)
        self.btn_confirm_start.pack(side=tk.LEFT, padx=(0, 10))
        tk.Button(
            cb_btns, text="✖  取消", font=FONT_HEAD,
            bg=BG_INPUT, fg=TEXT_SEC, relief=tk.FLAT,
            activebackground=BORDER, activeforeground=TEXT_PRI,
            cursor="hand2", padx=18, pady=8,
            command=self._cancel_start_flash
        ).pack(side=tk.LEFT)

        # 按钮栏
        btn_bar = tk.Frame(self.root, bg=BG_PANEL, height=56)
        btn_bar.grid(row=4, column=0, sticky="ew")
        btn_bar.grid_propagate(False)

        for text, bg, fg, cmd in [
            ("▶  开始并行烧录", "#0d419d", ACCENT,   self._start_flash),
            ("■  停止烧录",     "#4a1515", DANGER,   self._stop_flash),
            ("💾  保存日志",    "#0f2d1e", SUCCESS,  self._save_log),
            ("🗑  清空日志",    BG_INPUT,  TEXT_SEC, self._clear_log),
            ("↺  重置状态",    "#3a2800", WARNING,  self._reset_state),
        ]:
            b = tk.Button(
                btn_bar, text=text, font=FONT_HEAD,
                bg=bg, fg=fg, relief=tk.FLAT,
                activebackground=BORDER, activeforeground=TEXT_PRI,
                cursor="hand2", padx=16, pady=10, command=cmd
            )
            b.pack(side=tk.LEFT, padx=8, pady=8)
            if "开始并行烧录" in text:
                self.btn_start = b

        self.btn_single = tk.Button(
            btn_bar,
            text="🔧  处理待补烧设备",
            font=FONT_HEAD,
            bg="#4a2800", fg=WARNING,
            relief=tk.FLAT,
            activebackground=BORDER, activeforeground=TEXT_PRI,
            cursor="hand2", padx=16, pady=10,
            command=self._start_manual_retry
        )
        self.btn_single.pack_forget()

    def _build_left(self, parent):
        # 小屏幕电脑上，左侧这一整列（JLink列表+文件配置+烧录统计）纵向
        # 总高度可能超过屏幕可用高度，之前是直接截断、看不到"烧录统计"。
        # 这里包一层可滚动的 Canvas，不管屏幕多矮，鼠标滚轮/滚动条都能
        # 看到完整内容，不会丢信息。
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        left_canvas = tk.Canvas(parent, bg=BG_DARK, highlightthickness=0)
        left_vsb = ttk.Scrollbar(parent, orient="vertical",
                                  command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_vsb.set)
        left_canvas.grid(row=0, column=0, sticky="nsew")
        left_vsb.grid(row=0, column=1, sticky="ns")

        inner = tk.Frame(left_canvas, bg=BG_DARK)
        inner.grid_columnconfigure(0, weight=1)
        inner_win = left_canvas.create_window((0, 0), window=inner,
                                                anchor="nw")

        def _sync_scrollregion(event=None):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        inner.bind("<Configure>", _sync_scrollregion)

        def _sync_inner_width(event):
            left_canvas.itemconfig(inner_win, width=event.width)
        left_canvas.bind("<Configure>", _sync_inner_width)

        def _on_mousewheel(event):
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # 只在鼠标停在左侧栏上时接管滚轮，避免影响右侧日志区域的滚动
        left_canvas.bind(
            "<Enter>",
            lambda e: left_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        left_canvas.bind(
            "<Leave>", lambda e: left_canvas.unbind_all("<MouseWheel>"))

        parent = inner  # 下面原有代码不用改动，卡片继续往 parent 里挂

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

        self.single_queue_lbl = tk.Label(
            jc, text="", font=FONT_SMALL,
            bg=BG_CARD, fg=WARNING, anchor="w", justify=tk.LEFT
        )
        self.single_queue_lbl.pack(fill=tk.X, padx=10, pady=(0, 8))

        fc_outer, fc = self._card(parent, "📁  文件配置")
        fc_outer.grid(row=1, column=0, sticky="ew", pady=(0, 9))

        self.hex_path   = tk.StringVar()
        self.xlsx_path  = tk.StringVar()
        self.bat_path   = tk.StringVar()
        self.out_dir    = tk.StringVar(value="./output")
        self.base_addr  = tk.StringVar(value="0x178000")
        self.start_row  = tk.StringVar(value="2")
        self.speed_var  = tk.StringVar(value="4000")

        self._file_row(fc, "输入 hex", self.hex_path,
                       "*.hex",  "选择输入 hex 文件")
        self._file_row(fc, "Excel 数据",  self.xlsx_path,
                       "*.xlsx", "选择 Excel 数据文件")
        self._file_row(fc, "BAT 文件",    self.bat_path,
                       "*.bat",  "选择 ncs BAT 文件")
        self._label_entry_row(fc, "输出目录", self.out_dir,  browse_dir=True)
        self._label_entry_row(fc, "基础地址", self.base_addr)
        self._label_entry_row(fc, "起始行",   self.start_row)

        # 烧录速度：单独构建以拿到 Entry 引用做校验变色
        speed_f = tk.Frame(fc, bg=BG_CARD)
        speed_f.pack(fill=tk.X, padx=10, pady=3)
        tk.Label(speed_f, text="速度(kHz)", font=FONT_BODY, bg=BG_CARD,
                 fg=TEXT_SEC, width=11, anchor="w").pack(side=tk.LEFT)
        self.speed_entry = tk.Entry(
            speed_f, textvariable=self.speed_var, font=FONT_MONO,
            bg=BG_INPUT, fg=TEXT_PRI, insertbackground=ACCENT,
            relief=tk.FLAT, bd=4)
        self.speed_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ToolTip(self.speed_entry, text="JLink SWD 速度，范围 1000–8000 kHz"
                                        "（1–8MHz），默认 4000。太高可能导致"
                                        "接触不稳的探针出现新的偶发失败，"
                                        "调整后建议先跑几十片验证稳定性。")

        for var in [self.hex_path, self.xlsx_path, self.bat_path,
                    self.out_dir, self.base_addr, self.start_row,
                    self.speed_var]:
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

        sc_outer, sc = self._card(parent, "📊  烧录统计")
        sc_outer.grid(row=2, column=0, sticky="ew", pady=(0, 9))

        self.stat_success = tk.StringVar(value="0")
        self.stat_fail    = tk.StringVar(value="0")
        self.stat_total   = tk.StringVar(value=str(self.state.total_flashed))
        self.stat_row     = tk.StringVar(value=f"第 {self.state.next_row} 行")
        self.stat_prov    = tk.StringVar(value=str(self.state.provisioned_counter))

        # 紧凑方块网格（3列），比原来 5 行竖排省下不少纵向空间，
        # 配合上面的可滚动容器，小屏幕也能完整看到。
        stats_grid = tk.Frame(sc, bg=BG_CARD)
        stats_grid.pack(fill=tk.X, padx=10, pady=(8, 10))
        for c in range(3):
            stats_grid.grid_columnconfigure(c, weight=1)

        for i, (lbl, var, color) in enumerate([
            ("本轮成功",      self.stat_success, SUCCESS),
            ("本轮失败",      self.stat_fail,    DANGER),
            ("历史总计成功",  self.stat_total,   ACCENT),
            ("下批起始行",    self.stat_row,     TEXT_PRI),
            ("今日已烧录数",  self.stat_prov,    WARNING),
        ]):
            r, c = divmod(i, 3)
            tile = tk.Frame(stats_grid, bg=BG_INPUT, bd=1,
                             highlightbackground=BORDER,
                             highlightthickness=1)
            tile.grid(row=r, column=c, padx=4, pady=4, sticky="nsew")
            tk.Label(tile, text=lbl, font=FONT_SMALL, bg=BG_INPUT,
                     fg=TEXT_SEC, anchor="w"
                     ).pack(fill=tk.X, padx=8, pady=(6, 0))
            tk.Label(tile, textvariable=var, font=FONT_HEAD,
                     bg=BG_INPUT, fg=color, anchor="w"
                     ).pack(fill=tk.X, padx=8, pady=(0, 6))

    def _build_right(self, parent):
        # 上：6 槽位并行状态网格
        grid_outer, grid_content = self._card(parent, "🧩  并行烧录状态（6 槽位）")
        grid_outer.grid(row=0, column=0, sticky="ew", pady=(0, 9))
        grid = tk.Frame(grid_content, bg=BG_CARD)
        grid.pack(fill=tk.X, padx=8, pady=8)
        for c in range(3):
            grid.grid_columnconfigure(c, weight=1)

        for i in range(MAX_SLOTS):
            r, c = divmod(i, 3)
            tile = tk.Frame(grid, bg=BG_INPUT, bd=1,
                             highlightbackground=BORDER,
                             highlightthickness=1)
            tile.grid(row=r, column=c, padx=5, pady=5, sticky="nsew")

            lbl_slot = tk.Label(tile, text=f"槽位 {i+1}", font=FONT_HEAD,
                                 bg=BG_INPUT, fg=TEXT_SEC)
            lbl_slot.pack(anchor="w", padx=8, pady=(6, 0))
            lbl_id = tk.Label(tile, text="（空）", font=FONT_MONO,
                               bg=BG_INPUT, fg=TEXT_DIM)
            lbl_id.pack(anchor="w", padx=8)
            lbl_sn = tk.Label(tile, text="", font=FONT_BODY,
                               bg=BG_INPUT, fg=TEXT_PRI)
            lbl_sn.pack(anchor="w", padx=8)
            bar = ttk.Progressbar(tile, orient="horizontal",
                                   mode="determinate", length=140)
            bar.pack(fill=tk.X, padx=8, pady=(4, 2))
            lbl_step = tk.Label(tile, text="等待开始", font=FONT_SMALL,
                                 bg=BG_INPUT, fg=TEXT_SEC)
            lbl_step.pack(anchor="w", padx=8, pady=(0, 6))

            self.slot_frames.append({
                "tile": tile, "id": lbl_id, "sn": lbl_sn,
                "bar": bar, "step": lbl_step,
            })

        # 下：实时日志
        lc_outer, lc = self._card(parent, "📋  实时烧录日志")
        lc_outer.grid(row=1, column=0, sticky="nsew")
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
            ("SINGLE",   "#ffa657", None),
            ("CRITICAL", DANGER,    ("Consolas", 9, "bold")),
            ("HIGHLIGHT", "#ffffff", ("Consolas", 13, "bold")),
            ("HL_OK",     SUCCESS,   ("Consolas", 13, "bold")),
            ("HL_FAIL",   DANGER,    ("Consolas", 13, "bold")),
        ]:
            kw = {"foreground": fg}
            if font:
                kw["font"] = font
            self.log_text.tag_config(tag, **kw)
        self.log_text.tag_config(
            "BANNER", foreground="#ffffff", background="#0d3a6e",
            font=("Microsoft YaHei UI", 11, "bold"), spacing1=6, spacing3=6)
        self.log_text.tag_config(
            "BANNER_END", foreground="#ffffff", background="#1a3a1a",
            font=("Microsoft YaHei UI", 11, "bold"), spacing1=6, spacing3=6)
        self.log_text.tag_config(
            "BANNER_FAIL", foreground="#ffffff", background="#4a1515",
            font=("Microsoft YaHei UI", 11, "bold"), spacing1=6, spacing3=6)
        # 每个槽位一个身份色 tag，只用来给 "[槽N]" 前缀上色
        for i, c in enumerate(SLOT_COLORS):
            self.log_text.tag_config(f"SLOT{i}", foreground=c,
                                      font=("Consolas", 9, "bold"))
        # JLink 编号 / SN 单独放大、用两种反差明显的颜色，方便操作员一眼区分
        self.log_text.tag_config(
            "JLINKID", foreground="#79c0ff", font=("Consolas", 14, "bold"))
        self.log_text.tag_config(
            "SNVAL", foreground="#ffd166", font=("Consolas", 15, "bold"))

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
        e = tk.Entry(f, textvariable=var, font=FONT_MONO,
                     bg=BG_INPUT, fg=TEXT_PRI, insertbackground=ACCENT,
                     relief=tk.FLAT, bd=4)
        e.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if browse_dir:
            tk.Button(
                f, text="浏览", font=FONT_SMALL, bg=BG_INPUT, fg=TEXT_SEC,
                relief=tk.FLAT, padx=6, cursor="hand2",
                command=lambda: var.set(
                    filedialog.askdirectory(title="选择输出目录") or var.get())
            ).pack(side=tk.LEFT, padx=(4, 0))
        return e

    def _set_status(self, text: str, color: str = TEXT_SEC):
        self.status_badge.config(text=f"● {text}", fg=color)

    def _update_stats(self, success=None, fail=None):
        if success is not None:
            self.stat_success.set(str(success))
        if fail is not None:
            self.stat_fail.set(str(fail))
        self.stat_total.set(str(self.state.total_flashed))
        self.stat_row.set(f"第 {self.state.next_row} 行")
        self.stat_prov.set(str(self.state.provisioned_counter))

    def _refresh_manual_queue_ui(self):
        items = self.state.get_retry_list()
        if items:
            rows = ", ".join(str(x["row"]) for x in items)
            self.single_queue_lbl.config(
                text=f"⚠️  待补烧：{len(items)} 行\n({rows})")
            self.btn_single.pack(side=tk.LEFT, padx=8, pady=8)
        else:
            self.single_queue_lbl.config(text="")
            self.btn_single.pack_forget()

    # ──────────────────────────────────────────
    #  日志（线程安全：worker 只 enqueue，UI tick 里才真正写控件）
    # ──────────────────────────────────────────
    def _enqueue_log(self, msg, level: str = "INFO",
                      full_detail: str = "", jlink_id=None, visible: bool = True):
        """visible=False：只写进 log_full（保存到文件用），不刷进滚动日志窗口——
        用来把每一小步的技术细节留给工程师查文件，操作员看到的界面保持简洁。
        level="BOARD_START" 时 msg 应为 {"sn":..,"row":..}，用于渲染
        JLink 编号 / SN 分开上色放大的那一行。"""
        self.log_queue.put((msg, level, full_detail, jlink_id, visible))

    def _do_log(self, msg, level: str = "INFO",
                full_detail: str = "", jlink_id=None, visible: bool = True):
        ts = datetime.now().strftime("%H:%M:%S")
        slot_idx = self._round_slot_map.get(jlink_id) if jlink_id else None
        prefix = f"[槽{slot_idx + 1}] " if slot_idx is not None else ""

        if level == "BOARD_START":
            sn, row = msg.get("sn", ""), msg.get("row", "")
            plain_line = f"[{ts}] {prefix}▶ JLink {jlink_id}  SN:{sn}  (第{row}行)"
        else:
            plain_line = f"[{ts}] {prefix}{msg}"

        self.log_lines.append(plain_line)
        self.log_full.append(plain_line)
        if full_detail:
            for dl in full_detail.splitlines():
                self.log_full.append(f"         {dl}")

        if not visible:
            return

        self.log_text.config(state=tk.NORMAL)
        if level == "BOARD_START":
            sn, row = msg.get("sn", ""), msg.get("row", "")
            self.log_text.insert(tk.END, f"[{ts}] ", "INFO")
            if slot_idx is not None:
                self.log_text.insert(tk.END, prefix,
                                      f"SLOT{slot_idx % len(SLOT_COLORS)}")
            self.log_text.insert(tk.END, "▶ JLink ", "INFO")
            self.log_text.insert(tk.END, f"{jlink_id}", "JLINKID")
            self.log_text.insert(tk.END, "   SN ", "INFO")
            self.log_text.insert(tk.END, f"{sn}", "SNVAL")
            self.log_text.insert(tk.END, f"   (第{row}行)\n", "INFO")
        elif slot_idx is not None:
            self.log_text.insert(tk.END, f"[{ts}] ", "INFO")
            self.log_text.insert(tk.END, prefix, f"SLOT{slot_idx % len(SLOT_COLORS)}")
            self.log_text.insert(tk.END, f"{msg}\n", level)
        else:
            self.log_text.insert(tk.END, f"[{ts}] {msg}\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _ui_tick(self):
        # 1) 批量消费日志队列（限量，避免极端情况下界面被刷爆卡死）
        drained = 0
        while drained < 300:
            try:
                item = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._do_log(*item)
            drained += 1

        # 2) 刷新 6 槽位状态网格
        with self.status_lock:
            snapshot = {k: dict(v) for k, v in self.slot_status.items()}
        for jlink_id, idx in self._round_slot_map.items():
            if idx >= len(self.slot_frames):
                continue
            st = snapshot.get(jlink_id)
            frame = self.slot_frames[idx]
            if st is None:
                continue
            frame["id"].config(text=jlink_id, fg=TEXT_PRI)
            row = st.get("row")
            sn  = st.get("sn", "")
            frame["sn"].config(
                text=(f"第{row}行  {sn}" if row else sn))
            slot_state = st.get("state", "idle")
            color = {"running": ACCENT, "success": SUCCESS,
                     "fail": DANGER, "idle": TEXT_SEC}.get(slot_state, TEXT_SEC)
            bar_style = {"running": "Running.Horizontal.TProgressbar",
                         "success": "Success.Horizontal.TProgressbar",
                         "fail":    "Fail.Horizontal.TProgressbar",
                         "idle":    "Idle.Horizontal.TProgressbar"}.get(
                             slot_state, "Idle.Horizontal.TProgressbar")
            frame["bar"].configure(style=bar_style)
            frame["bar"]["value"] = st.get("pct", 0)
            frame["step"].config(text=st.get("detail", ""), fg=color)
            frame["tile"].config(highlightbackground=color)

        # 3) 刷新整批进度条 —— 有失败时显示红色，不会在"全失败"时还是一条容易
        #    被误读成"都成功了"的绿条
        if self._round_slot_map:
            total = len(self._round_slot_map)
            states = [snapshot.get(jid, {}).get("state")
                      for jid in self._round_slot_map]
            done = sum(1 for s in states if s in ("success", "fail"))
            any_fail = any(s == "fail" for s in states)
            pct = int(done / total * 100) if total else 0
            self.overall_bar.configure(
                style="Fail.Horizontal.TProgressbar" if any_fail
                else "Success.Horizontal.TProgressbar")
            self.overall_bar["value"] = pct
            label = f"{done}/{total}  {pct}%" + ("　⚠ 有失败" if any_fail else "")
            self.overall_lbl.config(text=label,
                                    fg=(DANGER if any_fail else TEXT_SEC))

        self.root.after(150, self._ui_tick)

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
            self._do_log(f"日志已保存：{path}", "SUCCESS")

    def _auto_save_log(self):
        ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = os.path.join(get_base_dir(), "logs")
        os.makedirs(log_dir, exist_ok=True)
        path = os.path.join(log_dir, f"flash_log_{ts}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.log_full))
        self._do_log(f"📁 日志已自动保存：{path}", "SUCCESS")

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
            self._do_log("正在检测 JLink 设备...", "INFO")
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
                    self._do_log(
                        f"✅ 检测到 {len(ids)} 个 JLink：{', '.join(ids)}",
                        "SUCCESS")
            else:
                self.jlink_count_lbl.config(text="未检测到设备", fg=DANGER)
                if not silent:
                    self._do_log("⚠️ 未检测到任何 JLink 设备", "WARNING")

            # 预览槽位映射（还没开始烧录时，也让网格显示"就绪"）
            if not self.is_flashing:
                self._round_slot_map = {jid: i for i, jid in enumerate(ids[:MAX_SLOTS])}
                with self.status_lock:
                    for jid in ids[:MAX_SLOTS]:
                        self.slot_status[jid] = {
                            "state": "idle", "detail": "就绪，等待开始",
                            "pct": 0, "sn": "", "row": None,
                        }
                for i in range(len(self.slot_frames)):
                    if i >= len(ids):
                        f = self.slot_frames[i]
                        f["id"].config(text="（空）", fg=TEXT_DIM)
                        f["sn"].config(text="")
                        f["bar"]["value"] = 0
                        f["step"].config(text="", fg=TEXT_SEC)
                        f["tile"].config(highlightbackground=BORDER)
            return ids
        except FileNotFoundError:
            if not silent:
                self._do_log("❌ 未找到 nrfjprog，请确认已安装并加入 PATH", "ERROR")
            return []
        except Exception as e:
            if not silent:
                self._do_log(f"❌ JLink 检测异常：{e}", "ERROR")
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
            self._enqueue_log("❌ Excel 文件被占用，请关闭后重试", "ERROR")
            return None
        except Exception as e:
            self._enqueue_log(f"❌ 读取 Excel 第 {row_index} 行失败：{e}", "ERROR")
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

    def _run_with_progress(self, cmd: list, timeout: int,
                           progress_cb, stop_flag_ref) -> tuple:
        if progress_cb:
            progress_cb(pct=0, detail="烧录中...", state="running")
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
                    if progress_cb:
                        progress_cb(pct=0, detail="超时 ❌", state="fail")
                    return (False, f"命令超时（>{timeout}s）",
                            "\n".join(output_lines))
                if stop_flag_ref():
                    proc.kill()
                    if progress_cb:
                        progress_cb(pct=0, detail="已停止", state="fail")
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
                        if progress_cb:
                            progress_cb(pct=pct, detail=f"烧录 {pct}%", state="running")
                    elif "Verifying" in line and got_real_pct:
                        if progress_cb:
                            progress_cb(pct=90, detail="校验中...", state="running")
                    elif ("Applying system reset" in line
                          or "Reset" in line) and got_real_pct:
                        if progress_cb:
                            progress_cb(pct=98, detail="复位中...", state="running")

            ret      = proc.wait()
            full_out = "\n".join(output_lines)
            if ret == 0:
                if progress_cb:
                    progress_cb(pct=100, detail="完成 ✅", state="running")
                return True, "", full_out
            else:
                if progress_cb:
                    progress_cb(pct=0, detail="失败 ❌", state="fail")
                return False, extract_error_summary(full_out), full_out

        except FileNotFoundError as e:
            if progress_cb:
                progress_cb(pct=0, detail="失败 ❌", state="fail")
            return False, f"命令未找到：{e}", ""
        except Exception as e:
            if progress_cb:
                progress_cb(pct=0, detail="失败 ❌", state="fail")
            return False, str(e), ""

    # ──────────────────────────────────────────
    #  单块板子烧录（核心，纯函数式：只用传进来的 cfg/log_cb/progress_cb）
    # ──────────────────────────────────────────
    def _flash_one(self, jlink_id: str, data: dict,
                   prov_name: str, prov_file: str,
                   cfg: dict, log_cb, progress_cb) -> tuple:
        sn_orig  = data["sn"]
        sn_hex   = sn_to_hex32(sn_orig)
        uuid     = data["uuid"]
        token    = data["token"]
        row      = data["row"]
        bat_path = cfg["bat_path"]
        bat_dir  = (os.path.dirname(os.path.abspath(bat_path))
                    if bat_path else ".")
        env      = parse_bat_env(bat_path)
        # nrfjprog 的速度参数叫 --clockspeed（不是 --speed），
        # 见 `nrfjprog --help`：-c --clockspeed <speed>，单位 kHz。
        speed_args = ["--clockspeed", str(cfg["speed_khz"])]

        # 操作员看板：只显示"这个槽位现在对应哪块板子"，不刷每一小步的技术细节
        log_cb({"sn": sn_orig, "row": row}, "BOARD_START", jlink_id=jlink_id)
        log_cb(f"UUID: {uuid}  →  {prov_name}.hex", "INFO",
               jlink_id=jlink_id, visible=False)
        if progress_cb:
            progress_cb(sn=sn_orig, row=row, state="running",
                        pct=2, detail="初始化中...")

        def step_ok(n, el):
            # 每一步成功的详细耗时只写进保存的日志文件，不刷进操作员看的滚动窗口
            log_cb(f"   {STEP_NAMES[n]}  ✅  {el:.1f}s", "SUCCESS",
                   jlink_id=jlink_id, visible=False)

        def step_skip(n):
            log_cb(f"   {STEP_NAMES[n]}  ⏭  已跳过", "INFO",
                   jlink_id=jlink_id, visible=False)

        def step_fail(n, el, summary, full):
            # 失败必须让操作员看到，所以保持可见
            log_cb(f"   {STEP_NAMES[n]}  ❌  {el:.1f}s  →  {summary}",
                   "ERROR", full_detail=full, jlink_id=jlink_id)

        def fail_progress(detail):
            if progress_cb:
                progress_cb(state="fail", detail=detail, pct=0)

        def cleanup_prov_file():
            # 生成出来的烧录用 hex 只是这一次烧录当下要用的临时文件，不再
            # 保留归档：每片板子（不论成功失败）走完步骤四之后都会被删掉，
            # 不再占用 output 目录空间。每片的 SN/UUID/行号/对应文件名
            # 仍然完整记录在自动保存的日志文件里，需要时按日志就能查得到。
            try:
                if os.path.isfile(prov_file):
                    os.remove(prov_file)
            except OSError as e:
                log_cb(f"   （提示）临时烧录文件删除失败：{e}", "INFO",
                       jlink_id=jlink_id, visible=False)

        # 步骤一
        t = time.time()
        r = self._run(["nrfjprog", "--recover", "-s", jlink_id, *speed_args],
                      timeout=60)
        el = time.time() - t
        if r["code"] != 0:
            full = r["err"] or r["out"]
            step_fail(1, el, extract_error_summary(full), full)
            fail_progress("步骤一失败")
            return False, "步骤一失败：芯片解除保护", full
        step_ok(1, el)
        if self.stop_flag:
            return False, "用户手动停止", ""
        if progress_cb:
            progress_cb(pct=15, detail="步骤二：擦除中")

        # 步骤二
        t = time.time()
        r = self._run(["nrfjprog", "-e", "-s", jlink_id, *speed_args],
                      timeout=60)
        el = time.time() - t
        if r["code"] != 0:
            full = r["err"] or r["out"]
            step_fail(2, el, extract_error_summary(full), full)
            fail_progress("步骤二失败")
            return False, "步骤二失败：擦除芯片内容", full
        step_ok(2, el)
        if self.stop_flag:
            return False, "用户手动停止", ""
        if progress_cb:
            progress_cb(pct=30, detail="步骤三：密钥上传")

        # 步骤三（可选）
        if cfg["use_bat"]:
            if not bat_path or not os.path.isfile(bat_path):
                msg = "未指定有效 BAT 文件"
                log_cb(f"   {STEP_NAMES[3]}  ❌  {msg}", "ERROR", jlink_id=jlink_id)
                fail_progress("步骤三失败")
                return False, f"步骤三失败：{msg}", msg

            # 密钥文件相对 BAT 文件的位置有两种可能：BAT 放在项目根目录（这时
            # 密钥在 "根目录/configuration/nrf54l15dk_nrf54l15_cpuapp/..." 下），
            # 或者 BAT 文件本身就放在 configuration/ 目录里（这时密钥就在
            # BAT 文件旁边的 "nrf54l15dk_nrf54l15_cpuapp/..." 下）。两种都试一下，
            # 而不是只认定一种、找不到就让 west 自己抛一个看不懂的报错。
            pem_candidates = [
                os.path.join(bat_dir, "configuration",
                             "nrf54l15dk_nrf54l15_cpuapp",
                             "boot_signature_key_file_ed25519.pem"),
                os.path.join(bat_dir, "nrf54l15dk_nrf54l15_cpuapp",
                             "boot_signature_key_file_ed25519.pem"),
            ]
            pem_file = next((p for p in pem_candidates if os.path.isfile(p)), None)
            if pem_file is None:
                msg = ("找不到密钥文件 boot_signature_key_file_ed25519.pem，"
                       "已尝试以下位置：\n"
                       + "\n".join(f"           - {p}" for p in pem_candidates))
                log_cb(f"   {STEP_NAMES[3]}  ❌  找不到密钥文件（已尝试2处，"
                       f"详情见保存的日志文件）", "ERROR",
                       full_detail=msg, jlink_id=jlink_id)
                fail_progress("步骤三失败：找不到密钥文件")
                return False, "步骤三失败：找不到密钥文件", msg

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
                log_cb(f"   {STEP_NAMES[3]}  ❌  {msg}", "ERROR", jlink_id=jlink_id)
                fail_progress("步骤三失败")
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
                fail_progress("步骤三失败")
                return False, "步骤三失败：密钥上传", full
            step_ok(3, el)
        else:
            step_skip(3)
        if self.stop_flag:
            return False, "用户手动停止", ""
        if progress_cb:
            progress_cb(pct=50, detail="步骤四：生成烧录文件")

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
            "--input-hex-file", os.path.abspath(cfg["hex_path"]),
            "--settings-base", cfg["base_addr"],
            "--output-path", prov_file,
        ], timeout=60)
        el = time.time() - t
        if r["code"] != 0:
            full = r["err"] or r["out"]
            step_fail(4, el, extract_error_summary(full), full)
            fail_progress("步骤四失败")
            cleanup_prov_file()
            return False, "步骤四失败：生成烧录文件", full
        if not os.path.isfile(prov_file):
            msg = f"{prov_name}.hex 未生成"
            log_cb(f"   {STEP_NAMES[4]}  ❌  {msg}", "ERROR", jlink_id=jlink_id)
            fail_progress("步骤四失败")
            return False, f"步骤四失败：{msg}", msg
        step_ok(4, el)
        if self.stop_flag:
            cleanup_prov_file()
            return False, "用户手动停止", ""
        if progress_cb:
            progress_cb(pct=60, detail="步骤五：烧录中")

        # 步骤五（耗时主体）
        t = time.time()
        ok, err_summary, full_out = self._run_with_progress(
            ["nrfjprog", "--program", prov_file,
             "--verify", "--reset", "-s", jlink_id, *speed_args],
            timeout=120,
            progress_cb=progress_cb,
            stop_flag_ref=lambda: self.stop_flag,
        )
        el = time.time() - t
        if not ok:
            step_fail(5, el, err_summary, full_out)
            fail_progress("步骤五失败")
            cleanup_prov_file()
            return False, "步骤五失败：烧录芯片", full_out
        step_ok(5, el)
        cleanup_prov_file()
        if progress_cb:
            progress_cb(state="success", detail="完成 ✅", pct=100)
        log_cb(f"✅  SN {sn_orig}  烧录成功", "HL_OK", jlink_id=jlink_id)
        return True, "", ""

    # ──────────────────────────────────────────
    #  并行调度
    # ──────────────────────────────────────────
    def _make_progress_cb(self, jlink_id):
        def cb(**kwargs):
            with self.status_lock:
                st = self.slot_status.setdefault(jlink_id, {})
                st.update(kwargs)
        return cb

    def _flash_worker(self, jlink_id, row_data, prov_stem, prov_file,
                      cfg, results):
        progress_cb = self._make_progress_cb(jlink_id)
        ok, summary, full = self._flash_one(
            jlink_id, row_data, prov_stem, prov_file,
            cfg, self._enqueue_log, progress_cb)
        results[jlink_id] = (ok, summary, full, row_data)

    def _run_round(self, pairs, cfg, day_output, round_label="本轮"):
        """pairs: [(jlink_id, row_data), ...]。并行跑完，返回 (success, fail)：
        success = [(jlink_id, row_data), ...]
        fail    = [(jlink_id, row_data, summary, full), ...]
        """
        results = {}
        threads = []
        for jlink_id, row_data in pairs:
            with self.status_lock:
                self.slot_status[jlink_id] = {
                    "sn": row_data["sn"], "row": row_data["row"],
                    "state": "running", "pct": 0, "detail": "排队中",
                }
            _, _, prov_stem = self.state.next_provisioned_name()
            prov_file = os.path.join(day_output, f"{prov_stem}.hex")
            th = threading.Thread(
                target=self._flash_worker,
                args=(jlink_id, row_data, prov_stem, prov_file, cfg, results),
                daemon=True,
            )
            threads.append(th)

        # 操作员只需要知道"这一轮几片"，速度这类工程参数不放进看板日志，
        # 仍会写进保存的日志文件（下面这行本身就会被存档）
        self._enqueue_log(
            f"  🚀  {round_label}开始  共 {len(pairs)} 片", "BANNER")
        for th in threads:
            th.start()
        for th in threads:
            th.join()

        success, fail = [], []
        for jlink_id, row_data in pairs:
            ok, summary, full, _ = results.get(
                jlink_id, (False, "未知错误：工作线程未返回结果", "", row_data))
            if ok:
                success.append((jlink_id, row_data))
            else:
                fail.append((jlink_id, row_data, summary, full))
        return success, fail

    def _start_flash(self):
        if self.is_flashing:
            return

        self._do_log("正在检测 JLink 设备...", "INFO")
        fresh_ids = self._detect_jlinks(silent=True)
        if fresh_ids:
            self._do_log(
                f"✅ 检测到 {len(fresh_ids)} 个 JLink：{', '.join(fresh_ids)}",
                "SUCCESS")
        else:
            messagebox.showerror("错误", "未检测到 JLink 设备，请检查连接！")
            return

        if not self.hex_path.get() or not os.path.isfile(self.hex_path.get()):
            messagebox.showerror("错误", "输入 hex 文件不存在，请重新选择！")
            return
        if not self.xlsx_path.get() or not os.path.isfile(self.xlsx_path.get()):
            messagebox.showerror("错误", "Excel 数据文件不存在，请重新选择！")
            return

        cfg = self._capture_cfg()
        if cfg is None:
            return

        try:
            start_row_val = int(self.start_row.get().strip())
            if start_row_val < 2:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误",
                                 "起始行必须是大于等于2的整数（第1行为标题行）！")
            return
        self.state.set_next_row(start_row_val)

        sorted_ids = sorted(fresh_ids)[:MAX_SLOTS]
        rows = self.state.claim_rows(len(sorted_ids))

        pairs, skipped_no_data = [], []
        for jlink_id, row in zip(sorted_ids, rows):
            data = self._read_excel_row(row)
            if data is None:
                skipped_no_data.append(row)
                continue
            pairs.append((jlink_id, data))

        if not pairs:
            messagebox.showerror(
                "错误",
                f"分配到的行号（{rows}）都读不到数据，Excel 数据可能已用完。\n"
                f"请补充数据后再烧录。"
            )
            return

        # 原来这里是两个系统弹窗（数据不足提醒 + 确认并行烧录），
        # 现在合并成主界面里的一条内嵌确认条，不再弹独立对话框。
        # 真正的启动动作放到 _confirm_start_flash()，这里只是把这一轮
        # 的数据存起来、把确认条内容填好、显示出来。
        self._pending_pairs = pairs
        self._pending_cfg   = cfg
        self._show_start_confirm_bar(pairs, cfg, skipped_no_data)

    def _show_start_confirm_bar(self, pairs, cfg, skipped_no_data):
        if skipped_no_data:
            self.confirm_warn_lbl.config(
                text=f"⚠ 以下行号已分配但读不到数据（不会被后续批次复用）："
                     f"{skipped_no_data}，本轮将只烧录另外 {len(pairs)} 片")
        else:
            self.confirm_warn_lbl.config(text="")

        preview = "\n".join(
            f"  槽位{i+1}  JLink {jid}  →  第{d['row']}行  SN:{d['sn']}"
            for i, (jid, d) in enumerate(pairs))
        self.confirm_preview_lbl.config(
            text=f"即将并行烧录 {len(pairs)} 片（速度 {cfg['speed_khz']} kHz），"
                 f"请确认板子已按槽位顺序插好：\n{preview}")

        if self.btn_start:
            self.btn_start.config(state=tk.DISABLED, bg="#1a1a2e", fg=TEXT_DIM)
        self.confirm_bar.grid()

    def _hide_start_confirm_bar(self):
        self.confirm_bar.grid_remove()
        if self.btn_start and not self.is_flashing:
            self.btn_start.config(state=tk.NORMAL, bg="#0d419d", fg=ACCENT)

    def _cancel_start_flash(self):
        self._pending_pairs = None
        self._pending_cfg   = None
        self._hide_start_confirm_bar()
        self._do_log("已取消本轮烧录（未开始）", "INFO")

    def _confirm_start_flash(self):
        pairs = self._pending_pairs
        cfg   = self._pending_cfg
        self._pending_pairs = None
        self._pending_cfg   = None
        self._hide_start_confirm_bar()
        if not pairs or cfg is None:
            return

        self.is_flashing = True
        self.stop_flag   = False
        self._set_status("烧录中...", WARNING)
        if self.btn_start:
            self.btn_start.config(state=tk.DISABLED, bg="#1a1a2e", fg=TEXT_DIM)
        self._round_slot_map = {jid: i for i, (jid, _) in enumerate(pairs)}
        threading.Thread(target=self._flash_round_thread,
                         args=(pairs, cfg), daemon=True).start()

    def _stop_flash(self):
        if self.is_flashing:
            self.stop_flag = True
            self._set_status("正在停止...", DANGER)
            self._enqueue_log("\n⛔ 已请求停止，等待当前步骤结束...", "CRITICAL")
        else:
            self._do_log("当前未在烧录中", "INFO")

    def _flash_round_thread(self, pairs, cfg):
        round_start_ts = time.time()
        self.state.mark_start()
        today       = datetime.now().strftime("%Y%m%d")
        base_output = os.path.abspath(cfg["out_dir"] or "./output")
        day_output  = os.path.join(base_output, today)
        os.makedirs(day_output, exist_ok=True)

        success, fail = self._run_round(pairs, cfg, day_output, "第一轮")
        for _jid, _row in success:
            self.state.add_flashed(1)

        if fail and not self.stop_flag:
            fail_txt = "\n".join(
                f"  • 槽位{self._round_slot_map.get(j, -1) + 1}  "
                f"JLink {j}  第{d['row']}行：{s}"
                for j, d, s, _f in fail)
            confirm_event  = threading.Event()
            confirm_result = [False]

            def ask_confirm():
                r = messagebox.askyesno(
                    "确认原地补烧",
                    f"本轮 {len(fail)} 片失败：\n\n{fail_txt}\n\n"
                    f"如果是接触不良/瞬时电压不稳等偶发问题，通常原地重试即可。\n"
                    f"点击【是】立即原地补烧这些槽位，\n"
                    f"点击【否】跳过，转入待补烧队列稍后人工处理。"
                )
                confirm_result[0] = r
                confirm_event.set()

            self.root.after(0, ask_confirm)
            confirm_event.wait()

            if confirm_result[0] and not self.stop_flag:
                retry_pairs = [(j, d) for j, d, s, f in fail]
                r_success, r_fail = self._run_round(
                    retry_pairs, cfg, day_output, "原地补烧轮")
                for _jid, _row in r_success:
                    self.state.add_flashed(1)
                fail = r_fail

        for jlink_id, row_data, summary, _full in fail:
            self.state.add_retry(row_data["row"], summary)
            self._enqueue_log(
                f"🔧 第{row_data['row']}行仍未成功（{summary}），"
                f"已加入待补烧队列", "SINGLE", jlink_id=jlink_id)

        n_ok = len(pairs) - len(fail)
        ts_end = datetime.now().strftime("%H:%M:%S")
        round_elapsed = time.time() - round_start_ts
        banner_tag = "BANNER_END" if not fail else "BANNER_FAIL"
        self._enqueue_log(
            f"  🏁  本批烧录结束  成功 {n_ok}  失败 {len(fail)}  "
            f"用时 {round_elapsed:.1f}s  "
            f"累计 {self.state.total_flashed}  "
            f"下批从第{self.state.next_row}行起  ──  {ts_end}",
            banner_tag)

        self.is_flashing = False
        self.state.mark_end()
        self.root.after(0, lambda: self._on_round_finished(n_ok, len(fail)))

    def _on_round_finished(self, n_ok, n_fail):
        self._update_stats(n_ok, n_fail)
        self._refresh_manual_queue_ui()
        self._auto_save_log()
        # 把界面上的"起始行"同步成当前游标，避免下次点【开始并行烧录】时
        # 把游标覆盖回这个输入框里的旧值（该输入框只用于手动跳转/断点续烧）
        self.start_row.set(str(self.state.next_row))
        if self.btn_start:
            self.btn_start.config(state=tk.NORMAL, bg="#0d419d", fg=ACCENT)
        if n_fail == 0:
            self._set_status("上次全部成功 ✅", SUCCESS)
        else:
            self._set_status(f"上次有 {n_fail} 片失败 ❌", DANGER)

    # ──────────────────────────────────────────
    #  待补烧队列处理（复用并行引擎，只是槽位数 = 待补烧行数）
    # ──────────────────────────────────────────
    def _start_manual_retry(self):
        if self.is_flashing:
            self._do_log("⚠️ 烧录进行中，请等待完成", "WARNING")
            return
        items = self.state.get_retry_list()
        if not items:
            return

        fresh_ids = self._detect_jlinks(silent=True)
        if not fresh_ids:
            messagebox.showerror("错误", "未检测到 JLink 设备，请检查连接！")
            return

        cfg = self._capture_cfg()
        if cfg is None:
            return

        sorted_ids = sorted(fresh_ids)[:MAX_SLOTS]
        n = min(len(items), len(sorted_ids))
        preview = "\n".join(
            f"  • 第{items[i]['row']}行（失败{items[i]['fail_count']}次，"
            f"最近原因：{items[i]['reason']}）"
            for i in range(n))
        if not messagebox.askyesno(
            "确认补烧",
            f"待补烧队列共 {len(items)} 行，本次用 {n} 个探针处理：\n\n{preview}\n\n"
            f"请确认对应芯片已分别接好对应探针，继续吗？"
        ):
            return

        pairs = []
        for jlink_id, item in zip(sorted_ids[:n], items[:n]):
            data = self._read_excel_row(item["row"])
            if data is None:
                self._do_log(
                    f"❌ 第{item['row']}行读取不到数据，本次跳过", "ERROR")
                continue
            pairs.append((jlink_id, data))

        if not pairs:
            return

        self.is_flashing = True
        self._round_slot_map = {jid: i for i, (jid, _) in enumerate(pairs)}
        if self.btn_start:
            self.btn_start.config(state=tk.DISABLED, bg="#1a1a2e", fg=TEXT_DIM)
        self.btn_single.config(state=tk.DISABLED, bg="#2a1800", fg=TEXT_DIM)
        threading.Thread(target=self._manual_retry_thread,
                         args=(pairs, cfg), daemon=True).start()

    def _manual_retry_thread(self, pairs, cfg):
        retry_start_ts = time.time()
        today       = datetime.now().strftime("%Y%m%d")
        base_output = os.path.abspath(cfg["out_dir"] or "./output")
        day_output  = os.path.join(base_output, today)
        os.makedirs(day_output, exist_ok=True)

        success, fail = self._run_round(pairs, cfg, day_output, "待补烧处理")
        for _jid, row_data in success:
            self.state.add_flashed(1)
            self.state.remove_retry(row_data["row"])
        for _jid, row_data, summary, _full in fail:
            self.state.add_retry(row_data["row"], summary)

        ts_end = datetime.now().strftime("%H:%M:%S")
        retry_elapsed = time.time() - retry_start_ts
        self._enqueue_log(
            f"  🔧  待补烧处理结束  成功 {len(success)}  仍失败 {len(fail)}  "
            f"用时 {retry_elapsed:.1f}s  "
            f"──  {ts_end}",
            "BANNER_END" if not fail else "BANNER_FAIL")

        self.is_flashing = False
        self.state.mark_end()
        self.root.after(0, lambda: self._on_manual_retry_finished(
            len(success), len(fail)))

    def _on_manual_retry_finished(self, n_ok, n_fail):
        self._update_stats(n_ok, n_fail)
        self._refresh_manual_queue_ui()
        self._auto_save_log()
        if self.btn_start:
            self.btn_start.config(state=tk.NORMAL, bg="#0d419d", fg=ACCENT)
        self.btn_single.config(state=tk.NORMAL, bg="#4a2800", fg=WARNING)
        if n_fail == 0:
            self._set_status("就绪", SUCCESS)
        else:
            self._set_status(f"仍有 {n_fail} 片待补烧", DANGER)

    # ──────────────────────────────────────────
    #  重置（需输入 RESET）
    # ──────────────────────────────────────────
    def _reset_state(self):
        if self.is_flashing:
            messagebox.showwarning("警告", "烧录进行中，无法重置！")
            return
        ans = simpledialog.askstring(
            "安全确认",
            "此操作将重置行分配游标至第2行，\n"
            "清空今日计数、历史累计、待补烧队列，路径/速度配置保留。\n\n"
            "请输入  RESET  确认：",
            parent=self.root
        )
        if ans and ans.strip() == "RESET":
            cfg_keys = ["cfg_hex", "cfg_xlsx", "cfg_bat",
                        "cfg_out_dir", "cfg_base_addr", "cfg_use_bat",
                        "cfg_start_row", "cfg_speed_khz"]
            self.state.reset_all(cfg_keys)
            self.start_row.set("2")
            self._update_stats(0, 0)
            self._refresh_manual_queue_ui()
            self._set_status("就绪", SUCCESS)
            self._do_log("🔄 状态已重置，分配游标回到第 2 行", "WARNING")
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