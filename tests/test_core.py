"""核心逻辑自测：python -m unittest discover -s tests -v"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import load_workbook  # noqa: E402
from openpyxl.cell.rich_text import CellRichText, TextBlock  # noqa: E402

from app import svn_client  # noqa: E402
from app.excel_writer import split_message_segments, write_report  # noqa: E402
from app.models import FileChange, LogEntry  # noqa: E402

SAMPLE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<log>
<logentry revision="12">
<author>zhangsan</author>
<date>2024-06-30T02:12:33.123456Z</date>
<paths>
<path action="M" kind="file">/trunk/src/main.cpp</path>
<path action="A" kind="file">/trunk/src/util.H</path>
<path action="M" kind="dir">/trunk/src</path>
<path action="D" kind="file">/trunk/readme.md</path>
</paths>
<msg>修复崩溃 &amp; 优化</msg>
</logentry>
<logentry revision="11">
<author>lisi</author>
<date>2024-06-28T09:00:00.000000Z</date>
<paths>
<path action="M" kind="file">/trunk/doc/guide.txt</path>
</paths>
<msg>只改文档</msg>
</logentry>
</log>""".encode("utf-8")


class PatternTests(unittest.TestCase):
    def test_parse_patterns(self) -> None:
        self.assertEqual(svn_client.parse_patterns(".cpp;.h"), [".cpp", ".h"])
        self.assertEqual(svn_client.parse_patterns("*.cpp, H"), [".cpp", ".h"])
        self.assertEqual(svn_client.parse_patterns("cpp; h;.h"), [".cpp", ".h"])
        self.assertEqual(svn_client.parse_patterns(""), [])
        self.assertEqual(svn_client.parse_patterns("   "), [])

    def test_rev_spec(self) -> None:
        self.assertIsNone(svn_client.build_rev_spec("rev", "", ""))
        self.assertEqual(svn_client.build_rev_spec("rev", "", ""), None)
        self.assertEqual(svn_client.build_rev_spec("rev", "10", ""), "10:HEAD")
        self.assertEqual(svn_client.build_rev_spec("rev", "", "20"), "1:20")
        self.assertEqual(
            svn_client.build_rev_spec("date", "2024-01-01", "2024-06-30"),
            "{2024-01-01T00:00:00}:{2024-06-30T23:59:59}",
        )
        self.assertEqual(svn_client.build_rev_spec("date", "", "2024-06-30"), "1:{2024-06-30T23:59:59}")
        self.assertIsNone(svn_client.build_rev_spec("date", "", ""))

    def test_limit_command(self) -> None:
        command = svn_client.build_log_command("svn.exe", "https://svn.example.com/repo", None, 10)
        self.assertIn("--limit", command)
        self.assertEqual(command[command.index("--limit") + 1], "10")
        self.assertNotIn("-r", command)

        plain = svn_client.build_log_command("svn.exe", "https://svn.example.com/repo", "1:HEAD")
        self.assertNotIn("--limit", plain)


class ParseTests(unittest.TestCase):
    def test_parse_and_filter(self) -> None:
        entries = svn_client.parse_log_xml(SAMPLE_XML)
        self.assertEqual(len(entries), 2)
        self.assertEqual(entries[0].revision, 12)
        self.assertEqual(entries[0].author, "zhangsan")
        self.assertEqual(len(entries[0].files), 3)  # 目录被忽略
        self.assertEqual(entries[0].files[0].ext, ".cpp")
        self.assertEqual(entries[0].files[1].action_cn, "新增")
        self.assertIn("修复崩溃", entries[0].message)

        filtered = svn_client.filter_entries(entries, [".cpp", ".h"])
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered[0].revision, 12)
        self.assertEqual(filtered[0].file_count, 2)

        kept = svn_client.filter_entries(entries, [".cpp"], keep_empty=True)
        self.assertEqual([e.revision for e in kept], [12, 11])
        self.assertEqual(kept[1].file_count, 0)


class MessageFormatTests(unittest.TestCase):
    def test_split_structured(self) -> None:
        message = (
            "更改原因：空指针导致崩溃\n"
            "问题单号：BUG-1234\n"
            "更改措施：增加判空\n"
            "影响范围：核心模块\n"
            "桌面验证：通过"
        )
        segments = split_message_segments(message)
        self.assertEqual([s[0] for s in segments], ["更改原因：", "问题单号：", "更改措施：", "影响范围：", "桌面验证："])
        self.assertEqual(segments[0][1], "空指针导致崩溃")

    def test_split_inline_keywords(self) -> None:
        segments = split_message_segments("改了逻辑 更改原因：崩溃 问题单号：BUG-1 更改措施：修复")
        labels = [s[0] for s in segments]
        self.assertEqual(labels, ["", "更改原因：", "问题单号：", "更改措施："])
        self.assertEqual(segments[0][1], "改了逻辑")

    def test_plain_message(self) -> None:
        self.assertEqual(split_message_segments("Fix a crash in parser"), [("", "Fix a crash in parser")])
        self.assertEqual(split_message_segments(""), [])


class ExcelTests(unittest.TestCase):
    def test_write_report(self) -> None:
        entries = svn_client.filter_entries(svn_client.parse_log_xml(SAMPLE_XML), [".cpp", ".h"])
        info = [
            ("生成时间", "2024-07-01 10:00:00"),
            ("SVN 地址", "https://svn.example.com/repo/trunk"),
            ("查询模式", "按版本范围"),
            ("导出文件数", "2"),
        ]
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "out.xlsx")
            write_report(path, info, entries)
            self.assertTrue(os.path.isfile(path))

            workbook = load_workbook(path)
            self.assertEqual(workbook.sheetnames, ["基本信息", "日志明细", "版本汇总"])
            basic, detail, summary = (workbook[name] for name in workbook.sheetnames)
            self.assertIn("基本信息", str(basic["A1"].value))
            self.assertEqual(detail["A1"].value, "序号")
            self.assertEqual(detail["B2"].value, 12)
            self.assertEqual(detail["G2"].value, "/trunk/src/main.cpp")
            self.assertEqual(detail.freeze_panes, "A2")
            self.assertEqual(summary["F2"].value, 2)  # 文件数
            self.assertEqual(summary["A2"].value, 1)

    def test_empty_message_entry(self) -> None:
        entry = LogEntry(revision=1, author="a", date=datetime.now(), message="", files=[])
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "empty.xlsx")
            write_report(path, [("x", "y")], [entry])
            self.assertTrue(os.path.isfile(path))

    def test_rich_message_alignment_and_height(self) -> None:
        message = (
            "更改原因：空指针导致崩溃\n"
            "问题单号：BUG-1234\n"
            "更改措施：增加判空处理，避免空指针解引用\n"
            "影响范围：核心模块\n"
            "桌面验证：通过"
        )
        entry = LogEntry(
            revision=9,
            author="wangwu",
            date=datetime.now(),
            message=message,
            files=[FileChange("/trunk/src/main.cpp", "M")],
        )
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "rich.xlsx")
            write_report(path, [("x", "y")], [entry])
            workbook = load_workbook(path, rich_text=True)
            detail = workbook["日志明细"]

            # 除说明/路径外全部居中
            for column in ("A2", "B2", "C2", "D2", "F2", "H2"):
                self.assertEqual(detail[column].alignment.horizontal, "center", column)
            self.assertEqual(detail["G2"].alignment.horizontal, "left")

            value = detail["E2"].value
            self.assertIsInstance(value, CellRichText)
            text = "".join(str(block) for block in value)
            self.assertIn("更改原因：空指针导致崩溃", text)
            self.assertIn("\n问题单号：BUG-1234", text)
            self.assertNotIn("更改原因：空指针导致崩溃 问题单号", text)

            bold_blocks = [b for b in value if isinstance(b, TextBlock) and b.font.b]
            self.assertEqual(len(bold_blocks), 5)
            self.assertGreaterEqual(detail.row_dimensions[2].height, 16)

            summary = workbook["版本汇总"]
            self.assertIsInstance(summary["E2"].value, CellRichText)
            self.assertEqual(summary["F2"].alignment.horizontal, "center")


if __name__ == "__main__":
    unittest.main()
