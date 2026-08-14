"""模板填充模块:以 Excel 为数据源,批量填充 Word 模板生成文档。

模板中用 {{字段名}} 作为占位符,字段名对应数据源 Excel 首行的表头。
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

PLACEHOLDER_RE = re.compile(r'\{\{(.+?)\}\}')


def load_data_rows(xlsx_path: str | Path) -> tuple[list[str], list[dict[str, str]]]:
    """读取数据源首个 Sheet,返回 (表头列表, 字典行列表)。"""
    wb = load_workbook(str(xlsx_path), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    headers: list[str] = []
    rows: list[dict[str, str]] = []
    for idx, row in enumerate(ws.iter_rows(values_only=True)):
        if idx == 0:
            headers = [str(c).strip() if c is not None else '' for c in row]
            continue
        if all(c is None or str(c).strip() == '' for c in row):
            continue
        record = {}
        for col, value in enumerate(row):
            if col < len(headers) and headers[col]:
                record[headers[col]] = '' if value is None else str(value)
        rows.append(record)
    wb.close()
    return headers, rows


def _fill_paragraph(paragraph, record: dict[str, str]) -> None:
    """替换段落中的占位符。占位符被拆分到多个 run 时先合并再替换。"""
    full = paragraph.text
    if not PLACEHOLDER_RE.search(full):
        return
    replaced = PLACEHOLDER_RE.sub(lambda m: record.get(m.group(1).strip(), m.group(0)), full)
    if replaced == full:
        return
    if paragraph.runs:
        paragraph.runs[0].text = replaced
        for run in paragraph.runs[1:]:
            run.text = ''
    else:
        paragraph.add_run(replaced)


def _fill_doc(doc: Document, record: dict[str, str]) -> None:
    def walk(paragraphs):
        for p in paragraphs:
            _fill_paragraph(p, record)

    walk(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                walk(cell.paragraphs)
    for section in doc.sections:
        for hf in (section.header, section.footer):
            walk(hf.paragraphs)


def generate_from_template(template_path: str | Path,
                           data_path: str | Path,
                           out_dir: str | Path,
                           name_field: str | None = None) -> list[str]:
    """按数据源每一行生成一个 docx,返回生成的文件名列表。

    name_field: 用作输出文件名的字段(该字段值必须非空),为空则用序号命名。
    """
    template_path = Path(template_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    headers, rows = load_data_rows(data_path)
    if not rows:
        raise ValueError('数据源中没有可用数据行')
    if name_field and name_field not in headers:
        raise ValueError(f'数据源中不存在字段: {name_field}')

    generated: list[str] = []
    seen: set[str] = set()
    for idx, record in enumerate(rows, 1):
        doc = Document(str(template_path))
        _fill_doc(doc, record)
        if name_field:
            stem = record.get(name_field, '').strip() or f'文档{idx}'
        else:
            stem = f'文档{idx}'
        stem = re.sub(r'[\\/:*?"<>|]', '_', stem)
        unique = stem
        n = 2
        while unique in seen:
            unique = f'{stem}_{n}'
            n += 1
        seen.add(unique)
        out_file = out_dir / f'{unique}.docx'
        doc.save(str(out_file))
        generated.append(out_file.name)
    return generated
