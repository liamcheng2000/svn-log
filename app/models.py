"""日志数据模型。"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

# SVN 的变更类型
ACTION_NAMES = {
    "A": "新增",
    "M": "修改",
    "D": "删除",
    "R": "替换",
}


@dataclass
class FileChange:
    """一次提交中涉及的单个文件。"""

    path: str
    action: str = "M"

    @property
    def action_cn(self) -> str:
        return ACTION_NAMES.get(self.action, self.action or "修改")

    @property
    def ext(self) -> str:
        name = self.path.replace("\\", "/").rsplit("/", 1)[-1]
        if "." not in name:
            return ""
        return "." + name.rsplit(".", 1)[-1].lower()


@dataclass
class LogEntry:
    """一个提交版本（revision）。"""

    revision: int
    author: str
    date: datetime
    message: str
    files: list[FileChange] = field(default_factory=list)

    @property
    def file_count(self) -> int:
        return len(self.files)

    def count_action(self, action: str) -> int:
        return sum(1 for f in self.files if f.action.upper() == action)

    def summary_message(self, limit: int = 0) -> str:
        text = " ".join(self.message.split())
        if limit and len(text) > limit:
            return text[: limit - 1] + "…"
        return text
