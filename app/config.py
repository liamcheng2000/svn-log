"""全局配置与用户偏好持久化（不保存任何凭据）。"""

from __future__ import annotations

import json
import os
from typing import Any

APP_NAME = "SVN 日志导出工具"
APP_VERSION = "1.1.2"

DEFAULT_FILE_FILTER = ".cpp;.h"

# 用户偏好固定放在 %APPDATA%\SVNLogApp\config.json
_APPDATA = os.environ.get("APPDATA") or os.path.expanduser("~")
CONFIG_DIR = os.path.join(_APPDATA, "SVNLogApp")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")

DEFAULTS: dict[str, Any] = {
    "url": "",
    "mode": "rev",              # "date" | "rev"
    "file_filter": DEFAULT_FILE_FILTER,
    "output_dir": "",
    "keep_empty": False,
    "open_after": True,
    "last_date_from": "",
    "last_date_to": "",
    "last_rev_from": "",
    "last_rev_to": "",
    "svn_exe": "",
}


def load_config() -> dict[str, Any]:
    """读取用户偏好，缺失或损坏时回退到默认值。"""
    data = dict(DEFAULTS)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as fp:
            saved = json.load(fp)
        if isinstance(saved, dict):
            for key in DEFAULTS:
                if key in saved and isinstance(saved[key], type(DEFAULTS[key])):
                    data[key] = saved[key]
    except (OSError, ValueError):
        pass
    return data


def save_config(data: dict[str, Any]) -> None:
    """尽力保存用户偏好，失败不影响主流程。"""
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        payload = {key: data.get(key, DEFAULTS[key]) for key in DEFAULTS}
        with open(CONFIG_PATH, "w", encoding="utf-8") as fp:
            json.dump(payload, fp, ensure_ascii=False, indent=2)
    except OSError:
        pass
