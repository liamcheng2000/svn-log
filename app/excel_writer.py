"""用 openpyxl 生成三页 Excel 报表。"""

from __future__ import annotations

from datetime import datetime
from typing import Callable, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .models import LogEntry

FONT_NAME = "微软雅黑"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True)
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="1F4E78")
ALT_FILL = PatternFill("solid", fgColor="F2F7FB")
LABEL_FILL = PatternFill("solid", fgColor="DDEBF7")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT_TOP = Alignment(horizontal="left", vertical="top")
WRAP_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)

DETAIL_HEADERS = ["序号", "版本号", "作者", "提交日期", "提交说明", "操作", "文件路径", "扩展名"]
DETAIL_WIDTHS = [6, 9, 15, 18, 42, 8, 66, 8]

SUMMARY_HEADERS = ["序号", "版本号", "作者", "提交日期", "提交说明", "文件数", "新增", "修改", "删除", "文件列表"]
SUMMARY_WIDTHS = [6, 9, 15, 18, 42, 8, 7, 7, 7, 78]


def _set_widths(ws: Worksheet, widths: Sequence[int]) -> None:
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width


def _write_header(ws: Worksheet, headers: Sequence[str]) -> None:
    for column, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=column, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"


def _build_basic_sheet(ws: Worksheet, info_items: Sequence[tuple[str, str]]) -> None:
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 96

    ws.merge_cells("A1:B1")
    title = ws["A1"]
    title.value = "SVN 日志导出报表 - 基本信息"
    title.font = TITLE_FONT
    title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30

    row = 3
    for label, value in info_items:
        label_cell = ws.cell(row=row, column=1, value=label)
        label_cell.font = BOLD_FONT
        label_cell.fill = LABEL_FILL
        label_cell.alignment = Alignment(horizontal="left", vertical="center")
        label_cell.border = BORDER

        value_cell = ws.cell(row=row, column=2, value=value)
        value_cell.font = BODY_FONT
        value_cell.alignment = WRAP_TOP
        value_cell.border = BORDER
        row += 1

    note_row = row + 1
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=2)
    note = ws.cell(row=note_row, column=1)
    note.value = (
        "说明：本报表仅统计扩展名过滤后的文件；被删除(D)的文件同样保留，便于追溯。"
        "「日志明细」按文件一行，「版本汇总」按版本一行。"
    )
    note.font = Font(name=FONT_NAME, size=9, italic=True, color="808080")
    note.alignment = WRAP_TOP

    ws.sheet_view.showGridLines = False


def _build_detail_sheet(ws: Worksheet, entries: Sequence[LogEntry]) -> int:
    _write_header(ws, DETAIL_HEADERS)
    _set_widths(ws, DETAIL_WIDTHS)

    row = 2
    index = 0
    for group, entry in enumerate(entries):
        shade = group % 2 == 1
        if not entry.files:
            # keep_empty 打开时，无匹配文件的版本也占一行，便于核对
            index += 1
            values = [index, entry.revision, entry.author, entry.date, entry.summary_message(), "", "", ""]
            _write_row(ws, row, values, shade)
            row += 1
            continue
        for change in entry.files:
            index += 1
            values = [
                index,
                entry.revision,
                entry.author,
                entry.date,
                entry.summary_message(),
                change.action_cn,
                change.path,
                change.ext,
            ]
            _write_row(ws, row, values, shade)
            row += 1

    last_row = max(row - 1, 1)
    ws.auto_filter.ref = f"A1:H{last_row}"
    return last_row


def _write_row(ws: Worksheet, row: int, values: Sequence[object], shade: bool) -> None:
    for column, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=column, value=value)
        cell.font = BODY_FONT
        cell.border = BORDER
        if column == 4:
            cell.number_format = "yyyy-mm-dd hh:mm"
            cell.alignment = CENTER
        elif column in (1, 2, 6, 8):
            cell.alignment = CENTER
        elif column in (5, 7):
            cell.alignment = WRAP_TOP
        else:
            cell.alignment = LEFT_TOP
        if shade:
            cell.fill = ALT_FILL


def _build_summary_sheet(ws: Worksheet, entries: Sequence[LogEntry]) -> None:
    _write_header(ws, SUMMARY_HEADERS)
    _set_widths(ws, SUMMARY_WIDTHS)

    row = 2
    for index, entry in enumerate(entries, start=1):
        file_list = "\n".join(f"{f.action_cn} {f.path}" for f in entry.files) or "(无匹配文件)"
        values = [
            index,
            entry.revision,
            entry.author,
            entry.date,
            entry.summary_message(),
            entry.file_count,
            entry.count_action("A"),
            entry.count_action("M"),
            entry.count_action("D"),
            file_list,
        ]
        _write_row(ws, row, values, index % 2 == 0)
        row += 1

    last_row = max(row - 1, 1)
    ws.auto_filter.ref = f"A1:J{last_row}"

    # 合计行
    total_row = row + 1
    ws.cell(row=total_row, column=1, value="合计").font = BOLD_FONT
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=1).alignment = CENTER
    ws.cell(row=total_row, column=1).fill = LABEL_FILL
    for column, value in ((6, sum(e.file_count for e in entries)),
                          (7, sum(e.count_action("A") for e in entries)),
                          (8, sum(e.count_action("M") for e in entries)),
                          (9, sum(e.count_action("D") for e in entries))):
        cell = ws.cell(row=total_row, column=column, value=value)
        cell.font = BOLD_FONT
        cell.alignment = CENTER
        cell.fill = LABEL_FILL
    for column in range(1, 11):
        ws.cell(row=total_row, column=column).border = BORDER


def write_report(
    path: str,
    info_items: Sequence[tuple[str, str]],
    entries: Sequence[LogEntry],
    on_progress: Callable[[int], None] | None = None,
) -> None:
    """生成三页报表并保存。"""
    workbook = Workbook()

    basic = workbook.active
    basic.title = "基本信息"
    _build_basic_sheet(basic, info_items)
    if on_progress:
        on_progress(20)

    detail = workbook.create_sheet("日志明细")
    _build_detail_sheet(detail, entries)
    if on_progress:
        on_progress(65)

    summary = workbook.create_sheet("版本汇总")
    _build_summary_sheet(summary, entries)
    if on_progress:
        on_progress(90)

    workbook.properties.title = "SVN 日志导出报表"
    workbook.properties.creator = "SVN 日志导出工具"
    workbook.properties.created = datetime.now()
    workbook.save(path)
    if on_progress:
        on_progress(100)
