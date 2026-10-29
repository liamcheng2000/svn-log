"""主窗口：输入、校验、后台查询、进度与日志。"""

from __future__ import annotations

import os
import queue
import threading
import traceback
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import config, svn_client
from .excel_writer import write_report
from .help_window import show_help
from .models import LogEntry
from .svn_client import Cancelled, SvnError

WEEKDAYS = ("一", "二", "三", "四", "五", "六", "日")


def _clamp_popup_position(
    anchor_x: int,
    anchor_y: int,
    anchor_w: int,
    anchor_h: int,
    popup_w: int,
    popup_h: int,
    screen_w: int,
    screen_h: int,
    margin: int = 8,
) -> tuple[int, int]:
    """把弹窗放到锚点下方（放不下则上方），并保证不超出屏幕。"""
    x = anchor_x + (anchor_w - popup_w) // 2
    y = anchor_y + anchor_h + 4
    if y + popup_h > screen_h - margin:
        y = anchor_y - popup_h - 4
    x = max(margin, min(x, screen_w - popup_w - margin))
    y = max(margin, min(y, screen_h - popup_h - margin))
    return x, y


def _is_file_locked(path: str) -> bool:
    """Windows 上判断文件是否被其他程序（如 Excel）独占。"""
    if not os.path.exists(path):
        return False
    try:
        with open(path, "a+b"):
            pass
        return False
    except OSError:
        return True


def _is_locked_error(exc: OSError) -> bool:
    """PermissionError / ERROR_SHARING_VIOLATION 都视为文件被占用。"""
    if isinstance(exc, PermissionError):
        return True
    if getattr(exc, "winerror", None) in (5, 32, 33):
        return True
    return exc.errno in (11, 13)


class CalendarPopup(tk.Toplevel):
    """轻量日历弹窗，避免引入第三方控件。"""

    def __init__(self, master: tk.Misc, initial: date, on_pick, anchor: tk.Misc | None = None):
        super().__init__(master)
        self.on_pick = on_pick
        self.anchor = anchor or master
        self.year = initial.year
        self.month = initial.month
        self.title("选择日期")
        self.resizable(False, False)
        self.transient(master.winfo_toplevel())
        self.configure(padx=8, pady=8)
        self.grab_set()

        self.header = ttk.Frame(self)
        self.header.pack(fill="x", pady=(0, 6))
        ttk.Button(self.header, text="<", width=3, command=lambda: self._shift(-1)).pack(side="left")
        self.month_label = ttk.Label(self.header, anchor="center", font=("微软雅黑", 10, "bold"))
        self.month_label.pack(side="left", expand=True, fill="x")
        ttk.Button(self.header, text=">", width=3, command=lambda: self._shift(1)).pack(side="right")

        self.body = ttk.Frame(self)
        self.body.pack()

        for col, name in enumerate(WEEKDAYS):
            ttk.Label(self.body, text=name, anchor="center", width=4).grid(row=0, column=col, padx=1, pady=1)

        self._render()
        self.bind("<Escape>", lambda _event: self.destroy())

        self.update_idletasks()
        self._place_near_anchor()

    def _place_near_anchor(self) -> None:
        """紧贴日历按钮弹出，并限制在屏幕范围内。"""
        anchor = self.anchor
        try:
            anchor_x = anchor.winfo_rootx()
            anchor_y = anchor.winfo_rooty()
            anchor_w = anchor.winfo_width()
            anchor_h = anchor.winfo_height()
        except tk.TclError:
            anchor_x = anchor_y = anchor_w = anchor_h = 0

        x, y = _clamp_popup_position(
            anchor_x,
            anchor_y,
            anchor_w,
            anchor_h,
            self.winfo_reqwidth(),
            self.winfo_reqheight(),
            self.winfo_screenwidth(),
            self.winfo_screenheight(),
        )
        self.geometry(f"+{x}+{y}")

    def _shift(self, delta: int) -> None:
        month = self.month + delta
        year = self.year + (month - 1) // 12
        month = (month - 1) % 12 + 1
        self.year, self.month = year, month
        self._render()

    def _render(self) -> None:
        import calendar

        for child in self.body.grid_slaves(row=1):
            child.destroy()
        for child in self.body.grid_slaves(row=2):
            child.destroy()
        for child in self.body.grid_slaves(row=3):
            child.destroy()
        for child in self.body.grid_slaves(row=4):
            child.destroy()
        for child in self.body.grid_slaves(row=5):
            child.destroy()
        for child in self.body.grid_slaves(row=6):
            child.destroy()

        self.month_label.configure(text=f"{self.year} 年 {self.month} 月")
        today = date.today()
        weeks = calendar.Calendar(firstweekday=0).monthdayscalendar(self.year, self.month)
        for row, week in enumerate(weeks, start=1):
            for col, day in enumerate(week):
                if day == 0:
                    ttk.Label(self.body, text="", width=4).grid(row=row, column=col, padx=1, pady=1)
                    continue
                is_today = (self.year, self.month, day) == (today.year, today.month, today.day)
                button = tk.Button(
                    self.body,
                    text=str(day),
                    width=3,
                    relief="flat",
                    bd=0,
                    cursor="hand2",
                    bg="#DDEBF7" if is_today else "white",
                    activebackground="#BDD7EE",
                    command=lambda d=day: self._pick(d),
                )
                button.grid(row=row, column=col, padx=1, pady=1, ipady=2)

    def _pick(self, day: int) -> None:
        picked = date(self.year, self.month, day)
        self.on_pick(picked.strftime("%Y-%m-%d"))
        self.destroy()


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.cfg = config.load_config()
        self.svn_exe = svn_client.find_svn_exe(self.cfg.get("svn_exe", ""))
        self.queue: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self._running = False

        self.title(f"{config.APP_NAME} v{config.APP_VERSION}")

        # 按屏幕尺寸自适应，避免小屏幕上窗口超出可见区域
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(1040, max(760, screen_w - 100))
        height = min(780, max(560, screen_h - 130))
        pos_x = max((screen_w - width) // 2, 0)
        pos_y = max((screen_h - height) // 2 - 10, 0)
        self.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        self.minsize(min(880, width), min(620, height))

        try:
            ttk.Style(self).theme_use("vista")
        except tk.TclError:
            pass

        self._build_vars()
        self._build_ui()
        self._apply_config()
        self._sync_mode_state()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<F1>", lambda _event: show_help(self))
        self.after(100, self._poll_queue)
        if not self.svn_exe:
            self._log("未找到 svn.exe，请手动指定（通常在 TortoiseSVN 的 bin 目录下）。")
        else:
            self._log(f"已定位 svn.exe: {self.svn_exe}")

    # ------------------------------------------------------------------ UI --
    def _build_vars(self) -> None:
        self.var_url = tk.StringVar()
        self.var_svn_exe = tk.StringVar(value=self.svn_exe)
        self.var_mode = tk.StringVar(value="rev")
        self.var_date_from = tk.StringVar()
        self.var_date_to = tk.StringVar()
        self.var_rev_from = tk.StringVar()
        self.var_rev_to = tk.StringVar()
        self.var_filter = tk.StringVar(value=config.DEFAULT_FILE_FILTER)
        self.var_output = tk.StringVar()
        self.var_keep_empty = tk.BooleanVar(value=False)
        self.var_open_after = tk.BooleanVar(value=True)
        self.var_status = tk.StringVar(value="就绪")

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)

        self._build_svn_frame(root).grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self._build_range_frame(root).grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self._build_output_frame(root).grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self._build_result_frame(root).grid(row=3, column=0, sticky="nsew")

    def _build_svn_frame(self, parent: ttk.Widget) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text=" SVN 配置 ", padding=10)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="SVN 地址").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.entry_url = ttk.Entry(frame, textvariable=self.var_url)
        self.entry_url.grid(row=0, column=1, sticky="ew")
        self.btn_test = ttk.Button(frame, text="测试连接", command=self._on_test)
        self.btn_test.grid(row=0, column=2, padx=(8, 0))

        ttk.Label(frame, text="svn.exe").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        self.entry_svn = ttk.Entry(frame, textvariable=self.var_svn_exe)
        self.entry_svn.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        self.btn_svn = ttk.Button(frame, text="浏览", command=self._on_browse_svn)
        self.btn_svn.grid(row=1, column=2, padx=(8, 0), pady=(8, 0))

        ttk.Label(
            frame,
            text="凭据复用 SVN 自身的认证缓存；认证失败时请点击「打开认证终端」手动登录一次。",
            foreground="#7F7F7F",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
        return frame

    def _build_range_frame(self, parent: ttk.Widget) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text=" 查询范围（二选一） ", padding=10)
        frame.columnconfigure(2, weight=1)
        frame.columnconfigure(5, weight=1)

        ttk.Radiobutton(
            frame, text="按日期范围", value="date", variable=self.var_mode, command=self._sync_mode_state
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text="起始").grid(row=0, column=1, sticky="e", padx=(12, 6))
        self.entry_date_from = ttk.Entry(frame, textvariable=self.var_date_from, width=14)
        self.entry_date_from.grid(row=0, column=2, sticky="ew")
        self.btn_cal_from = ttk.Button(
            frame, text="日历", width=6, command=lambda: self._pick_date(self.var_date_from, self.btn_cal_from)
        )
        self.btn_cal_from.grid(row=0, column=3, padx=(4, 12))
        ttk.Label(frame, text="结束").grid(row=0, column=4, sticky="e", padx=(0, 6))
        self.entry_date_to = ttk.Entry(frame, textvariable=self.var_date_to, width=14)
        self.entry_date_to.grid(row=0, column=5, sticky="ew")
        self.btn_cal_to = ttk.Button(
            frame, text="日历", width=6, command=lambda: self._pick_date(self.var_date_to, self.btn_cal_to)
        )
        self.btn_cal_to.grid(row=0, column=6, padx=(4, 0))
        ttk.Label(frame, text="留空 = 不限（YYYY-MM-DD）", foreground="#7F7F7F").grid(
            row=0, column=7, sticky="w", padx=(12, 0)
        )

        ttk.Radiobutton(
            frame, text="按版本范围", value="rev", variable=self.var_mode, command=self._sync_mode_state
        ).grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Label(frame, text="起始").grid(row=1, column=1, sticky="e", padx=(12, 6), pady=(10, 0))
        self.entry_rev_from = ttk.Entry(frame, textvariable=self.var_rev_from, width=14)
        self.entry_rev_from.grid(row=1, column=2, sticky="ew", pady=(10, 0))
        ttk.Label(frame, text="结束").grid(row=1, column=4, sticky="e", padx=(0, 6), pady=(10, 0))
        self.entry_rev_to = ttk.Entry(frame, textvariable=self.var_rev_to, width=14)
        self.entry_rev_to.grid(row=1, column=5, sticky="ew", pady=(10, 0))
        ttk.Label(frame, text="留空 = 最近 10 条（HEAD-10 ~ HEAD）", foreground="#7F7F7F").grid(
            row=1, column=7, sticky="w", padx=(12, 0), pady=(10, 0)
        )
        return frame

    def _build_output_frame(self, parent: ttk.Widget) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text=" 输出设置 ", padding=10)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="文件过滤").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.entry_filter = ttk.Entry(frame, textvariable=self.var_filter)
        self.entry_filter.grid(row=0, column=1, sticky="ew")
        ttk.Label(frame, text="分号分隔，如 .cpp;.h，留空 = 全部文件", foreground="#7F7F7F").grid(
            row=0, column=2, sticky="w", padx=(8, 0)
        )

        ttk.Label(frame, text="导出文件").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        self.entry_output = ttk.Entry(frame, textvariable=self.var_output)
        self.entry_output.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        self.btn_output = ttk.Button(frame, text="浏览", command=self._on_browse_output)
        self.btn_output.grid(row=1, column=2, padx=(8, 0), pady=(8, 0))

        options = ttk.Frame(frame)
        options.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Checkbutton(options, text="保留无匹配文件的版本", variable=self.var_keep_empty).pack(side="left")
        ttk.Checkbutton(options, text="生成后自动打开 Excel", variable=self.var_open_after).pack(side="left", padx=(16, 0))
        return frame

    def _build_result_frame(self, parent: ttk.Widget) -> ttk.Frame:
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        buttons = ttk.Frame(frame)
        buttons.grid(row=0, column=0, sticky="ew")
        self.btn_start = ttk.Button(buttons, text="开始导出", command=self._on_start)
        self.btn_start.pack(side="left")
        self.btn_cancel = ttk.Button(buttons, text="取消", command=self._on_cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=(8, 0))
        self.btn_auth = ttk.Button(buttons, text="打开认证终端", command=self._on_auth_terminal)
        self.btn_auth.pack(side="left", padx=(8, 0))
        self.btn_folder = ttk.Button(buttons, text="打开输出目录", command=self._on_open_folder)
        self.btn_folder.pack(side="left", padx=(8, 0))
        self.btn_help = ttk.Button(buttons, text="帮助", command=lambda: show_help(self))
        self.btn_help.pack(side="right")

        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(8, 4))

        ttk.Label(frame, textvariable=self.var_status, foreground="#404040").grid(row=2, column=0, sticky="w")

        text_frame = ttk.Frame(frame)
        text_frame.grid(row=3, column=0, sticky="nsew", pady=(6, 0))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            text_frame,
            height=12,
            wrap="word",
            font=("Consolas", 9),
            bg="#FBFBFB",
            relief="solid",
            bd=1,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        return frame

    # ------------------------------------------------------------- helpers --
    def _apply_config(self) -> None:
        self.var_url.set(self.cfg.get("url", ""))
        self.var_mode.set(self.cfg.get("mode", "rev"))
        self.var_filter.set(self.cfg.get("file_filter", config.DEFAULT_FILE_FILTER))
        self.var_keep_empty.set(bool(self.cfg.get("keep_empty", False)))
        self.var_open_after.set(bool(self.cfg.get("open_after", True)))
        self.var_date_from.set(self.cfg.get("last_date_from", ""))
        self.var_date_to.set(self.cfg.get("last_date_to", ""))
        self.var_rev_from.set(self.cfg.get("last_rev_from", ""))
        self.var_rev_to.set(self.cfg.get("last_rev_to", ""))

        output_dir = self.cfg.get("output_dir") or str(Path.home() / "Desktop")
        if not os.path.isdir(output_dir):
            output_dir = os.getcwd()
        self.var_output.set(os.path.join(output_dir, self._default_filename()))

    def _default_filename(self) -> str:
        return f"svn_log_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

    def _sync_mode_state(self) -> None:
        date_mode = self.var_mode.get() == "date"
        date_state = "normal" if date_mode else "disabled"
        rev_state = "disabled" if date_mode else "normal"
        for widget in (self.entry_date_from, self.entry_date_to, self.btn_cal_from, self.btn_cal_to):
            widget.configure(state=date_state)
        for widget in (self.entry_rev_from, self.entry_rev_to):
            widget.configure(state=rev_state)

    def _pick_date(self, variable: tk.StringVar, anchor: tk.Misc | None = None) -> None:
        initial = date.today()
        text = variable.get().strip()
        if text:
            try:
                initial = datetime.strptime(self._normalize_date(text), "%Y-%m-%d").date()
            except ValueError:
                pass
        CalendarPopup(self, initial, variable.set, anchor=anchor)

    def _log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{stamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _post(self, message: tuple) -> None:
        self.queue.put(message)

    def _set_running(self, running: bool) -> None:
        self._running = running
        self.btn_start.configure(state="disabled" if running else "normal")
        self.btn_cancel.configure(state="normal" if running else "disabled")
        self.btn_test.configure(state="disabled" if running else "normal")
        self.btn_auth.configure(state="disabled" if running else "normal")

    def _set_status(self, text: str) -> None:
        self.var_status.set(text)

    def _poll_queue(self) -> None:
        try:
            while True:
                message = self.queue.get_nowait()
                self._handle_message(message)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _handle_message(self, message: tuple) -> None:
        kind = message[0]
        if kind == "log":
            self._log(message[1])
        elif kind == "count":
            self._set_status(f"已获取 {message[1]} 个版本…")
        elif kind == "progress_set":
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress["value"] = message[1]
        elif kind == "done":
            path, rev_count, file_count = message[1], message[2], message[3]
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress["value"] = 100
            self._set_running(False)
            self._set_status(f"导出完成：{rev_count} 个版本，{file_count} 个文件")
            self._log(f"已生成：{path}")
            messagebox.showinfo(
                "导出完成",
                f"共导出 {rev_count} 个版本、{file_count} 个文件。\n\n文件：{path}",
            )
            if self.var_open_after.get():
                try:
                    os.startfile(path)  # noqa: S606
                except OSError as exc:
                    self._log(f"自动打开失败：{exc}")
        elif kind == "error":
            title, detail = message[1], message[2]
            self.progress.stop()
            self.progress.configure(mode="determinate", value=0)
            self._set_running(False)
            self._set_status("执行失败")
            self._log(f"{title}：{detail}")
            messagebox.showerror(title, detail)
        elif kind == "cancelled":
            self.progress.stop()
            self.progress.configure(mode="determinate", value=0)
            self._set_running(False)
            self._set_status("已取消")
            self._log("用户取消操作。")
        elif kind == "test_ok":
            self._log(message[1])
            self._set_status("连接成功")
            messagebox.showinfo("测试连接", message[1])
        elif kind == "test_err":
            self._log(f"测试连接失败：{message[1]}")
            self._set_status("测试连接失败")
            messagebox.showerror("测试连接失败", message[1])

    # ---------------------------------------------------------- validation --
    @staticmethod
    def _normalize_date(text: str) -> str:
        value = text.strip().replace("/", "-").replace(".", "-")
        if not value:
            return ""
        try:
            return datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError as exc:
            raise ValueError(f"日期格式不正确：{text}（应为 YYYY-MM-DD）") from exc

    def _collect_params(self) -> dict:
        url = self.var_url.get().strip()
        if not url:
            raise ValueError("请填写 SVN 地址。")
        if not url.lower().startswith(("http://", "https://", "svn://", "svn+ssh://", "file://")):
            raise ValueError("SVN 地址应以 http://、https://、svn:// 或 file:// 开头。")

        svn_exe = self.var_svn_exe.get().strip() or self.svn_exe
        if not svn_exe or not os.path.isfile(svn_exe):
            raise ValueError("未找到 svn.exe，请点击「浏览」指定（通常在 TortoiseSVN\\bin 下）。")

        mode = self.var_mode.get()
        if mode == "date":
            date_from = self._normalize_date(self.var_date_from.get())
            date_to = self._normalize_date(self.var_date_to.get())
            if date_from and date_to and date_from > date_to:
                raise ValueError("起始日期不能晚于结束日期。")
            rev_from = rev_to = ""
        else:
            date_from = date_to = ""
            rev_from = self.var_rev_from.get().strip()
            rev_to = self.var_rev_to.get().strip()
            for value in (rev_from, rev_to):
                if value and value.upper() != "HEAD" and not value.isdigit():
                    raise ValueError(f"版本号必须是数字或 HEAD：{value}")
            if rev_from and rev_to and rev_from.upper() != "HEAD" and rev_to.upper() != "HEAD":
                if int(rev_from) > int(rev_to):
                    raise ValueError("起始版本不能大于结束版本。")

        patterns = svn_client.parse_patterns(self.var_filter.get())

        output = self.var_output.get().strip()
        if not output:
            raise ValueError("请填写导出文件路径。")
        if not output.lower().endswith(".xlsx"):
            output += ".xlsx"
        output = os.path.abspath(output)

        limit = 10 if (mode == "rev" and not rev_from and not rev_to) else None

        return {
            "url": url,
            "svn_exe": svn_exe,
            "mode": mode,
            "date_from": date_from,
            "date_to": date_to,
            "rev_from": rev_from,
            "rev_to": rev_to,
            "limit": limit,
            "rev_spec": svn_client.build_rev_spec(mode, date_from or rev_from, date_to or rev_to),
            "patterns": patterns,
            "output": output,
            "keep_empty": self.var_keep_empty.get(),
        }

    # -------------------------------------------------------------- actions --
    def _on_start(self) -> None:
        if self._running:
            return
        try:
            params = self._collect_params()
        except ValueError as exc:
            messagebox.showwarning("输入有误", str(exc))
            return

        if not params["patterns"]:
            if not messagebox.askyesno("确认", "未设置文件过滤，将导出全部类型的文件，是否继续？"):
                return
        if os.path.exists(params["output"]):
            if _is_file_locked(params["output"]):
                messagebox.showwarning(
                    "文件被占用",
                    f"{params['output']}\n\n"
                    "该文件正被 Excel / WPS 打开，无法写入。\n"
                    "请先关闭该文件，或修改导出文件名后重试。",
                )
                return
            if not messagebox.askyesno("文件已存在", f"{params['output']}\n\n是否覆盖？"):
                return

        self.var_output.set(params["output"])
        self.cancel_event.clear()
        self._set_running(True)
        self._set_status("正在获取 SVN 日志…")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self._log("=" * 68)
        self._log(f"目标：{params['url']}")
        self._log(f"范围：{self._describe_range(params)}")
        self._log(f"文件过滤：{', '.join(params['patterns']) or '不过滤'}")

        self.worker = threading.Thread(target=self._run_export, args=(params,), daemon=True)
        self.worker.start()

    def _run_export(self, params: dict) -> None:
        started = datetime.now()
        try:
            command = svn_client.build_log_command(
                params["svn_exe"], params["url"], params["rev_spec"], params["limit"]
            )
            self._post(("log", "执行命令：" + " ".join(command)))

            raw_entries = svn_client.fetch_log(
                params["svn_exe"],
                params["url"],
                params["rev_spec"],
                limit=params["limit"],
                on_progress=lambda count: self._post(("count", count)),
                cancel_event=self.cancel_event,
            )
            self._post(("log", f"SVN 返回 {len(raw_entries)} 个版本。"))

            entries = svn_client.filter_entries(raw_entries, params["patterns"], params["keep_empty"])
            file_count = sum(item.file_count for item in entries)
            self._post(("log", f"过滤后保留 {len(entries)} 个版本、{file_count} 个文件。"))

            if not entries:
                self._post(
                    (
                        "error",
                        "没有匹配的日志",
                        "指定范围内没有符合条件的提交记录。\n"
                        "请放宽日期/版本范围，或检查文件过滤规则。",
                    )
                )
                return

            info_items = self._build_info_items(params, entries, len(raw_entries), started)
            os.makedirs(os.path.dirname(params["output"]) or ".", exist_ok=True)
            self._post(("log", "正在生成 Excel…"))
            self._post(("progress_set", 0))
            try:
                write_report(
                    params["output"],
                    info_items,
                    entries,
                    on_progress=lambda pct: self._post(("progress_set", pct)),
                )
            except OSError as exc:
                if _is_locked_error(exc):
                    self._post(
                        (
                            "error",
                            "Excel 文件被占用",
                            f"无法写入文件：\n{params['output']}\n\n"
                            "该文件可能正被 Excel / WPS 打开，或当前账号没有写入权限。\n"
                            "请关闭 Excel 后重新导出，或修改导出文件名/目录。",
                        )
                    )
                    return
                raise
            elapsed = (datetime.now() - started).total_seconds()
            self._post(("log", f"耗时 {elapsed:.1f} 秒。"))
            self._post(("done", params["output"], len(entries), file_count))
        except Cancelled:
            self._post(("cancelled",))
        except SvnError as exc:
            if exc.raw:
                self._post(("log", "svn 原始输出：" + exc.raw.strip()[:2000]))
            self._post(("error", "SVN 查询失败", exc.message))
        except Exception:  # noqa: BLE001 - 兜底，避免线程静默崩溃
            self._post(("error", "导出失败", traceback.format_exc()))

    def _describe_range(self, params: dict) -> str:
        if params["mode"] == "date":
            start = params["date_from"] or "不限"
            end = params["date_to"] or "不限"
            return f"按日期 {start} ~ {end}"
        if params.get("limit"):
            return f"最近 {params['limit']} 条（HEAD-10 ~ HEAD）"
        start = params["rev_from"] or "1"
        end = params["rev_to"] or "HEAD"
        return f"按版本 r{start} ~ r{end}"

    def _build_info_items(
        self, params: dict, entries: list[LogEntry], raw_count: int, started: datetime
    ) -> list[tuple[str, str]]:
        file_count = sum(item.file_count for item in entries)
        added = sum(item.count_action("A") for item in entries)
        modified = sum(item.count_action("M") for item in entries)
        deleted = sum(item.count_action("D") for item in entries)
        authors = sorted({item.author for item in entries if item.author})
        revisions = [item.revision for item in entries]

        if params["mode"] == "date":
            mode_text = "按日期范围"
            range_text = f"{params['date_from'] or '不限'} ~ {params['date_to'] or '不限'}"
        else:
            mode_text = "按版本范围"
            if params.get("limit"):
                range_text = "最近 10 条（HEAD-10 ~ HEAD）"
            else:
                range_text = f"r{params['rev_from'] or 1} ~ r{params['rev_to'] or 'HEAD'}"

        command = svn_client.build_log_command(
            params["svn_exe"], params["url"], params["rev_spec"], params["limit"]
        )
        return [
            ("生成时间", started.strftime("%Y-%m-%d %H:%M:%S")),
            ("SVN 地址", params["url"]),
            ("查询模式", mode_text),
            ("查询范围", range_text),
            ("文件类型过滤", ", ".join(params["patterns"]) or "不过滤（全部文件）"),
            ("保留无匹配文件的版本", "是" if params["keep_empty"] else "否"),
            ("SVN 返回版本数", f"{raw_count}"),
            ("导出后版本数", f"{len(entries)}"),
            ("导出文件数", f"{file_count}"),
            ("操作统计", f"新增 {added} / 修改 {modified} / 删除 {deleted}"),
            ("提交人数", f"{len(authors)}（{'、'.join(authors[:10])}{' 等' if len(authors) > 10 else ''}）"),
            ("版本区间", f"r{min(revisions)} ~ r{max(revisions)}" if revisions else "-"),
            ("执行命令", " ".join(command)),
            ("svn.exe 路径", params["svn_exe"]),
            ("生成工具", f"{config.APP_NAME} v{config.APP_VERSION}"),
        ]

    def _on_cancel(self) -> None:
        if not self._running:
            return
        self.cancel_event.set()
        self._set_status("正在取消…")
        self._log("已请求取消，正在终止 svn 进程…")

    def _on_test(self) -> None:
        url = self.var_url.get().strip()
        svn_exe = self.var_svn_exe.get().strip() or self.svn_exe
        if not url:
            messagebox.showwarning("提示", "请先填写 SVN 地址。")
            return
        if not svn_exe or not os.path.isfile(svn_exe):
            messagebox.showwarning("提示", "未找到 svn.exe，请先指定路径。")
            return

        self._set_status("正在测试连接…")
        self._log("正在测试连接…")

        def run() -> None:
            try:
                result = svn_client.test_connection(svn_exe, url)
                self._post(("test_ok", result))
            except SvnError as exc:
                self._post(("test_err", exc.message))
            except Exception:  # noqa: BLE001
                self._post(("test_err", traceback.format_exc()))

        threading.Thread(target=run, daemon=True).start()

    def _on_auth_terminal(self) -> None:
        url = self.var_url.get().strip()
        svn_exe = self.var_svn_exe.get().strip() or self.svn_exe
        if not url:
            messagebox.showwarning("提示", "请先填写 SVN 地址。")
            return
        if not svn_exe or not os.path.isfile(svn_exe):
            messagebox.showwarning("提示", "未找到 svn.exe，请先指定路径。")
            return
        try:
            svn_client.open_auth_terminal(svn_exe, url)
            self._log("已打开认证终端，请在其中输入账号密码，完成后回到本工具重试。")
        except OSError as exc:
            messagebox.showerror("打开失败", str(exc))

    def _on_browse_svn(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 svn.exe",
            filetypes=[("svn.exe", "svn.exe"), ("可执行文件", "*.exe"), ("所有文件", "*.*")],
        )
        if path:
            self.var_svn_exe.set(path)
            self.svn_exe = path

    def _on_browse_output(self) -> None:
        current = self.var_output.get().strip()
        initial_dir = os.path.dirname(current) if current else os.getcwd()
        if not os.path.isdir(initial_dir):
            initial_dir = os.getcwd()
        initial_file = os.path.basename(current) if current else self._default_filename()
        path = filedialog.asksaveasfilename(
            title="选择导出位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialdir=initial_dir,
            initialfile=initial_file,
        )
        if path:
            self.var_output.set(path)

    def _on_open_folder(self) -> None:
        path = self.var_output.get().strip()
        folder = os.path.dirname(path) if path else ""
        if folder and os.path.isdir(folder):
            os.startfile(folder)  # noqa: S606
        else:
            messagebox.showinfo("提示", "输出目录尚不存在。")

    def _on_close(self) -> None:
        if self._running:
            if not messagebox.askyesno("确认退出", "导出仍在进行中，确定要退出吗？"):
                return
            self.cancel_event.set()
        self._save_config()
        self.destroy()

    def _save_config(self) -> None:
        output = self.var_output.get().strip()
        config.save_config(
            {
                "url": self.var_url.get().strip(),
                "mode": self.var_mode.get(),
                "file_filter": self.var_filter.get(),
                "output_dir": os.path.dirname(output) if output else "",
                "keep_empty": bool(self.var_keep_empty.get()),
                "open_after": bool(self.var_open_after.get()),
                "last_date_from": self.var_date_from.get().strip(),
                "last_date_to": self.var_date_to.get().strip(),
                "last_rev_from": self.var_rev_from.get().strip(),
                "last_rev_to": self.var_rev_to.get().strip(),
                "svn_exe": self.var_svn_exe.get().strip(),
            }
        )
