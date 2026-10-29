"""使用说明文档自测：python -m unittest discover -s tests -v"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import APP_VERSION  # noqa: E402
from app.help_window import manual_path  # noqa: E402


class ManualTests(unittest.TestCase):
    def test_manual_exists(self) -> None:
        path = manual_path()
        self.assertTrue(path, "未找到 docs/user_manual.md")
        self.assertTrue(os.path.isfile(path))
        self.assertTrue(path.replace("\\", "/").endswith("docs/user_manual.md"))

    def test_manual_sections(self) -> None:
        with open(manual_path(), "r", encoding="utf-8") as fp:
            text = fp.read()

        for section in (
            "软件简介",
            "运行环境",
            "快速上手",
            "查询范围说明",
            "文件过滤",
            "Excel 报表说明",
            "认证与凭据",
            "常见问题",
            "版本记录",
        ):
            self.assertIn(section, text)

        self.assertIn(f"v{APP_VERSION}", text)

    def test_manual_not_empty(self) -> None:
        with open(manual_path(), "r", encoding="utf-8") as fp:
            text = fp.read()
        self.assertGreater(len(text), 2000)
        self.assertNotIn("\ufffd", text)


if __name__ == "__main__":
    unittest.main()
