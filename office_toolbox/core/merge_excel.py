"""Excel 合并模块:多个 xlsx 合并为一个(按工作表或按行追加)。"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook


def _unique_sheet_name(base: str, used: set[str]) -> str:
    """返回不与 used 冲突的工作表名(最长 31 字符)。"""
    name = base[:31]
    if name not in used:
        return name
    # 截断后添加数字后缀，直到不冲突
    n = 2
    while True:
        suffix = f'_{n}'
        candidate = base[:31 - len(suffix)] + suffix
        if candidate not in used:
            return candidate
        n += 1


def merge_as_sheets(paths: list[str | Path], output: str | Path) -> int:
    """每个源文件作为一个工作表写入目标工作簿,返回工作表数量。"""
    if not paths:
        raise ValueError('没有可合并的 Excel 文件')
    wb = Workbook()
    wb.remove(wb.active)
    used_names: set[str] = set()
    for p in paths:
        p = Path(p)
        if not p.exists():
            raise FileNotFoundError(f'文件不存在: {p.name}')
        if p.suffix.lower() not in ('.xlsx', '.xlsm'):
            raise ValueError(f'Excel 合并只接受 .xlsx/.xlsm 文件,但发现「{p.name}」')
        try:
            src = load_workbook(str(p), data_only=True)
        except Exception as exc:
            raise ValueError(f'无法读取「{p.name}」: {exc}') from exc
        for ws in src.worksheets:
            base = f'{Path(p).stem}_{ws.title}' if ws.title in used_names else ws.title
            name = _unique_sheet_name(base, used_names)
            used_names.add(name)
            new_ws = wb.create_sheet(title=name)
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
    if not paths:
        raise ValueError('没有可合并的 Excel 文件')
    wb = Workbook()
    ws = wb.active
    ws.title = '合并结果'
    total_rows = 0
    header_written = False
    for p in paths:
        p = Path(p)
        if not p.exists():
            raise FileNotFoundError(f'文件不存在: {p.name}')
        if p.suffix.lower() not in ('.xlsx', '.xlsm'):
            raise ValueError(f'Excel 合并只接受 .xlsx/.xlsm 文件,但发现「{p.name}」')
        try:
            src = load_workbook(str(p), data_only=True, read_only=True)
        except Exception as exc:
            raise ValueError(f'无法读取「{p.name}」: {exc}') from exc
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
