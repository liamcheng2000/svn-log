"""内置使用说明窗口（读取 docs/user_manual.md 并渲染）。"""

from __future__ import annotations

import os
import re
import sys
import tkinter as tk
from tkinter import ttk

from . import config

_INLINE = re.compile(r"(\*\*.+?\*\*|`[^`]+`)")
_current: "HelpWindow | None" = None


def manual_path() -> str:
    """按 exe 同级 docs → 打包资源 → 源码目录 的顺序查找说明文档。"""
    candidates: list[str] = []

    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    candidates.append(os.path.join(exe_dir, "docs", "user_manual.md"))

    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", "")
        if base:
            candidates.append(os.path.join(base, "docs", "user_manual.md"))

    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    candidates.append(os.path.join(project_root, "docs", "user_manual.md"))

    for path in candidates:
        if os.path.isfile(path):
            return path
    return ""


def _load_manual() -> tuple[str, str]:
    path = manual_path()
    if not path:
        return (
            "",
            "未找到使用说明文档。\n\n"
            "请确认安装目录下存在 docs\\user_manual.md；\n"
            "如果是源码运行，请确认项目根目录存在 docs\\user_manual.md。",
        )
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            with open(path, "r", encoding=encoding) as fp:
                return fp.read(), path
        except UnicodeDecodeError:
            continue
        except OSError as exc:
            return "", f"读取使用说明失败：{exc}"
    return "", "使用说明文件编码无法识别，请另存为 UTF-8。"


class HelpWindow(tk.Toplevel):
    """只读的使用说明窗口，支持简单的 Markdown 样式。"""

    def __init__(self, master: tk.Misc, content: str, path: str):
        super().__init__(master)
        self.title(f"使用说明 - {config.APP_NAME} v{config.APP_VERSION}")
        self.geometry("920x720")
        self.minsize(680, 480)
        self.transient(master.winfo_toplevel())
        self.configure(bg="white")

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        text_frame = tk.Frame(self, bg="white")
        text_frame.grid(row=0, column=0, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.text = tk.Text(
            text_frame,
            wrap="word",
            bg="white",
            relief="flat",
            padx=28,
            pady=22,
            font=("微软雅黑", 10),
            cursor="arrow",
            spacing1=0,
        )
        self.text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.text.configure(yscrollcommand=scrollbar.set)

        self._configure_tags()
        if content:
            self._render(content)

        bottom = ttk.Frame(self, padding=(16, 8))
        bottom.grid(row=1, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)

        if path:
            label = ttk.Label(
                bottom,
                text=f"文档位置：{path}",
                foreground="#808080",
                anchor="w",
            )
        else:
            label = ttk.Label(bottom, text="未找到使用说明文档", foreground="#C00000", anchor="w")
        label.grid(row=0, column=0, sticky="w", padx=(0, 12))

        if path:
            ttk.Button(bottom, text="打开文档目录", command=lambda: self._open_folder(path)).grid(
                row=0, column=1, padx=(0, 8)
            )
        ttk.Button(bottom, text="关闭", command=self._close).grid(row=0, column=2)

        self.text.configure(state="disabled")
        self.bind("<Escape>", lambda _event: self._close())

        self.update_idletasks()
        self._center_on(master)

    # ------------------------------------------------------------------ UI --
    def _configure_tags(self) -> None:
        self.text.tag_configure("body", font=("微软雅黑", 10), spacing1=2, spacing3=3, lmargin1=6, lmargin2=6)
        self.text.tag_configure(
            "h1",
            font=("微软雅黑", 18, "bold"),
            foreground="#1F4E78",
            spacing1=6,
            spacing3=12,
            lmargin1=0,
            lmargin2=0,
        )
        self.text.tag_configure(
            "h2",
            font=("微软雅黑", 13, "bold"),
            foreground="#1F4E78",
            spacing1=18,
            spacing3=6,
        )
        self.text.tag_configure(
            "h3",
            font=("微软雅黑", 11, "bold"),
            foreground="#305496",
            spacing1=12,
            spacing3=4,
        )
        self.text.tag_configure("bullet", font=("微软雅黑", 10), lmargin1=24, lmargin2=38, spacing1=1, spacing3=2)
        self.text.tag_configure(
            "code",
            font=("Consolas", 9),
            background="#F4F4F4",
            lmargin1=24,
            lmargin2=24,
            spacing1=2,
            spacing3=2,
        )
        self.text.tag_configure("code_inline", font=("Consolas", 9), background="#F4F4F4")
        self.text.tag_configure("bold", font=("微软雅黑", 10, "bold"))

    def _render(self, content: str) -> None:
        in_code = False
        for raw in content.splitlines():
            line = raw.rstrip()
            stripped = line.strip()

            if stripped.startswith("```"):
                in_code = not in_code
                continue
            if in_code:
                self.text.insert("end", line + "\n", ("code",))
                continue
            if not stripped:
                self.text.insert("end", "\n")
                continue
            if line.startswith("# "):
                self._insert_inline(line[2:].strip(), ("h1",))
                self.text.insert("end", "\n")
            elif line.startswith("## "):
                self._insert_inline(line[3:].strip(), ("h2",))
                self.text.insert("end", "\n")
            elif line.startswith("### "):
                self._insert_inline(line[4:].strip(), ("h3",))
                self.text.insert("end", "\n")
            elif stripped in ("---", "***"):
                self.text.insert("end", "─" * 60 + "\n", ("body",))
            elif stripped.startswith(("- ", "* ")):
                self.text.insert("end", "•  ", ("bullet",))
                self._insert_inline(stripped[2:], ("bullet",))
                self.text.insert("end", "\n")
            elif re.match(r"^\d+[.、]\s*", stripped):
                self.text.insert("end", "   ", ("bullet",))
                self._insert_inline(stripped, ("bullet",))
                self.text.insert("end", "\n")
            else:
                self._insert_inline(stripped, ("body",))
                self.text.insert("end", "\n")

    def _insert_inline(self, text: str, base: tuple[str, ...]) -> None:
        for part in _INLINE.split(text):
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                self.text.insert("end", part[2:-2], base + ("bold",))
            elif part.startswith("`") and part.endswith("`") and len(part) > 2:
                self.text.insert("end", part[1:-1], base + ("code_inline",))
            else:
                self.text.insert("end", part, base)

    def _center_on(self, master: tk.Misc) -> None:
        try:
            parent = master.winfo_toplevel()
            x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass

    def _open_folder(self, path: str) -> None:
        folder = os.path.dirname(path)
        if folder and os.path.isdir(folder) and os.name == "nt":
            os.startfile(folder)  # noqa: S606

    def _close(self) -> None:
        global _current
        _current = None
        self.destroy()


def show_help(master: tk.Misc) -> None:
    """打开（或复用）使用说明窗口。"""
    global _current
    if _current is not None:
        try:
            if _current.winfo_exists():
                _current.deiconify()
                _current.lift()
                _current.focus_set()
                return
        except tk.TclError:
            pass

    content, path = _load_manual()
    _current = HelpWindow(master, content, path)
