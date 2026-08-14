"""批量查找替换模块:在 docx / xlsx / pptx 中替换文本。"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from pptx import Presentation


def _replace_docx_paragraph(paragraph, pairs: list[tuple[str, str]]) -> int:
    """替换段落中的文本。若关键字被拆分到多个 run,则合并 run 后替换。"""
    count = 0
    full = paragraph.text
    for old, new in pairs:
        if old and old in full:
            count += full.count(old)
            full = full.replace(old, new)
    if count == 0:
        return 0
    if paragraph.runs:
        paragraph.runs[0].text = full
        for run in paragraph.runs[1:]:
            run.text = ''
    else:
        paragraph.add_run(full)
    return count


def _replace_docx(path: Path, pairs: list[tuple[str, str]]) -> int:
    doc = Document(str(path))
    total = 0

    def walk_paragraphs(paragraphs):
        nonlocal total
        for p in paragraphs:
            total += _replace_docx_paragraph(p, pairs)

    walk_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                walk_paragraphs(cell.paragraphs)
    # 页眉页脚
    for section in doc.sections:
        for hf in (section.header, section.footer):
            walk_paragraphs(hf.paragraphs)
            for table in hf.tables:
                for row in table.rows:
                    for cell in row.cells:
                        walk_paragraphs(cell.paragraphs)
    doc.save(str(path))
    return total


def _replace_xlsx(path: Path, pairs: list[tuple[str, str]]) -> int:
    wb = load_workbook(str(path))
    total = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell) or not isinstance(cell.value, str):
                    continue
                new_val = cell.value
                for old, new in pairs:
                    if old and old in new_val:
                        total += new_val.count(old)
                        new_val = new_val.replace(old, new)
                if new_val != cell.value:
                    cell.value = new_val
    wb.save(str(path))
    return total


def _replace_pptx_textframe(tf, pairs: list[tuple[str, str]]) -> int:
    count = 0
    for para in tf.paragraphs:
        full = ''.join(run.text for run in para.runs)
        replaced = full
        hit = 0
        for old, new in pairs:
            if old and old in replaced:
                hit += replaced.count(old)
                replaced = replaced.replace(old, new)
        if hit:
            count += hit
            if para.runs:
                para.runs[0].text = replaced
                for run in para.runs[1:]:
                    run.text = ''
            else:
                para.text = replaced
    return count


def _replace_pptx(path: Path, pairs: list[tuple[str, str]]) -> int:
    prs = Presentation(str(path))
    total = 0

    def walk_shapes(shapes):
        nonlocal total
        for shape in shapes:
            if shape.has_text_frame:
                total += _replace_pptx_textframe(shape.text_frame, pairs)
            if getattr(shape, 'has_table', False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        total += _replace_pptx_textframe(cell.text_frame, pairs)
            if shape.shape_type == 6:
                walk_shapes(shape.shapes)

    for slide in prs.slides:
        walk_shapes(slide.shapes)
    prs.save(str(path))
    return total


def replace_in_file(path: str | Path, pairs: list[tuple[str, str]]) -> int:
    """在单个文件中执行替换,返回替换发生的次数(直接保存原文件)。"""
    path = Path(path)
    ext = path.suffix.lower()
    if ext == '.docx':
        return _replace_docx(path, pairs)
    if ext in ('.xlsx', '.xlsm'):
        return _replace_xlsx(path, pairs)
    if ext == '.pptx':
        return _replace_pptx(path, pairs)
    raise ValueError(f'不支持的文件类型: {path.name}(仅支持 docx/xlsx/pptx)')
