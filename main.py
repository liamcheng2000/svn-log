"""程序入口。

用法:
    python main.py
"""

from __future__ import annotations

import ctypes
import os
import sys
import traceback


def _enable_dpi_awareness() -> None:
    """避免高分屏下界面模糊。"""
    if os.name != "nt":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:  # noqa: BLE001
            pass


def main() -> int:
    _enable_dpi_awareness()

    base = os.path.dirname(os.path.abspath(__file__))
    if base not in sys.path:
        sys.path.insert(0, base)

    try:
        from app.ui import App
    except Exception:  # noqa: BLE001
        _fatal("启动失败：依赖模块导入出错", traceback.format_exc())
        return 1

    try:
        app = App()
        app.mainloop()
    except Exception:  # noqa: BLE001
        _fatal("运行出错", traceback.format_exc())
        return 1
    return 0


def _fatal(title: str, detail: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, f"{detail}\n\n请确认已安装依赖：python -m pip install -r requirements.txt")
        root.destroy()
    except Exception:  # noqa: BLE001
        print(f"{title}\n{detail}", file=sys.stderr)


if __name__ == "__main__":
    raise SystemExit(main())
