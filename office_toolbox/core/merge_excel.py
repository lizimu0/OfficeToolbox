"""Excel 合并模块:多个 xlsx 合并为一个(按工作表或按行追加)。"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook


def merge_as_sheets(paths: list[str | Path], output: str | Path) -> int:
    """每个源文件作为一个工作表写入目标工作簿,返回工作表数量。"""
    wb = Workbook()
    wb.remove(wb.active)
    used_names: set[str] = set()
    for p in paths:
        src = load_workbook(str(p), data_only=True)
        for ws in src.worksheets:
            name = ws.title
            if name in used_names:  # 避免重名冲突
                name = f'{Path(p).stem}_{name}'[:31]
            used_names.add(name)
            new_ws = wb.create_sheet(title=name[:31])
            for row in ws.iter_rows(values_only=True):
                new_ws.append(list(row))
        src.close()
    if not wb.worksheets:
        wb.create_sheet(title='Sheet1')
    wb.save(str(output))
    return len(wb.worksheets)


def merge_append_rows(paths: list[str | Path], output: str | Path,
                      skip_header: bool = True) -> int:
    """按行追加合并(适用于表头相同的报表),返回合并的数据行数。"""
    wb = Workbook()
    ws = wb.active
    ws.title = '合并结果'
    total_rows = 0
    header_written = False
    for p in paths:
        src = load_workbook(str(p), data_only=True, read_only=True)
        sheet = src.worksheets[0]
        for idx, row in enumerate(sheet.iter_rows(values_only=True)):
            if idx == 0:
                if not header_written:
                    ws.append(list(row))
                    header_written = True
                elif not skip_header:
                    ws.append(list(row))
                    total_rows += 1
                continue
            ws.append(list(row))
            total_rows += 1
        src.close()
    wb.save(str(output))
    return total_rows
