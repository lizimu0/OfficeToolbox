"""格式转换模块:通过本机 Office COM 接口高保真转换(docx/xlsx/pptx → PDF 等)。

需要本机安装 Microsoft Office。转换在后台线程执行时会自动初始化 COM。
"""
from __future__ import annotations

import os
from pathlib import Path

from openpyxl import Workbook

WD_FORMAT_PDF = 17
XL_TYPE_PDF = 0
PP_SAVE_AS_PDF = 32
WD_FORMAT_TEXT_UTF8 = 7

_OFFICE_NOT_FOUND = '未检测到 Microsoft Office,格式转换功能需要安装 Office 才能使用。'


def _new_com_app(prog_id: str):
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch(prog_id)
    except OSError:
        raise RuntimeError(_OFFICE_NOT_FOUND) from None
    except Exception as exc:  # COM 组件未注册等
        raise RuntimeError(f'启动 Office 组件失败({prog_id}): {exc}') from exc
    app.DisplayAlerts = False
    app.Visible = False
    return app


class ConvertSession:
    """一次批量转换会话,Office 实例只在需要时启动并复用。"""

    def __init__(self) -> None:
        self._word = None
        self._excel = None
        self._powerpoint = None

    # ---- 惰性启动 ----
    def _ensure_word(self):
        if self._word is None:
            self._word = _new_com_app('Word.Application')
        return self._word

    def _ensure_excel(self):
        if self._excel is None:
            self._excel = _new_com_app('Excel.Application')
        return self._excel

    def _ensure_ppt(self):
        if self._powerpoint is None:
            self._powerpoint = _new_com_app('PowerPoint.Application')
        return self._powerpoint

    # ---- 转换入口 ----
    def convert(self, src: str | Path, dst: str | Path) -> None:
        src = Path(os.path.abspath(src))
        dst = Path(os.path.abspath(dst))
        dst.parent.mkdir(parents=True, exist_ok=True)
        pair = (src.suffix.lower(), dst.suffix.lower())
        if pair in {('.docx', '.pdf'), ('.doc', '.pdf')}:
            app = self._ensure_word()
            doc = app.Documents.Open(str(src), ReadOnly=True)
            try:
                doc.SaveAs2(str(dst), FileFormat=WD_FORMAT_PDF)
            finally:
                doc.Close(False)
        elif pair in {('.docx', '.txt'), ('.doc', '.txt')}:
            app = self._ensure_word()
            doc = app.Documents.Open(str(src), ReadOnly=True)
            try:
                doc.SaveAs2(str(dst), FileFormat=WD_FORMAT_TEXT_UTF8)
            finally:
                doc.Close(False)
        elif pair in {('.xlsx', '.pdf'), ('.xls', '.pdf')}:
            app = self._ensure_excel()
            wb = app.Workbooks.Open(str(src), ReadOnly=True)
            try:
                wb.ExportAsFixedFormat(XL_TYPE_PDF, str(dst))
            finally:
                wb.Close(False)
        elif pair in {('.pptx', '.pdf'), ('.ppt', '.pdf')}:
            app = self._ensure_ppt()
            prs = app.Presentations.Open(str(src), WithWindow=False)
            try:
                prs.SaveAs(str(dst), PP_SAVE_AS_PDF)
            finally:
                prs.Close()
        elif pair == ('.csv', '.xlsx'):
            csv_to_xlsx(src, dst)
        else:
            raise ValueError(f'不支持的转换: {src.suffix} → {dst.suffix}')

    def close(self) -> None:
        for app in (self._word, self._excel, self._powerpoint):
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
        self._word = self._excel = self._powerpoint = None
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass


def csv_to_xlsx(src: str | Path, dst: str | Path) -> None:
    """纯 Python 实现 csv → xlsx(无需 Office)。"""
    import csv
    wb = Workbook()
    ws = wb.active
    with open(src, newline='', encoding='utf-8-sig') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save(str(dst))


def target_extensions(src_path: str | Path) -> list[str]:
    """返回该文件可转换的目标扩展名列表(供界面下拉框使用)。"""
    ext = Path(src_path).suffix.lower()
    mapping = {
        '.docx': ['.pdf', '.txt'],
        '.doc': ['.pdf', '.txt'],
        '.xlsx': ['.pdf'],
        '.xls': ['.pdf'],
        '.pptx': ['.pdf'],
        '.ppt': ['.pdf'],
        '.csv': ['.xlsx'],
    }
    return mapping.get(ext, [])
