"""模板填充模块:以 Excel 为数据源,批量填充 Word 模板生成文档。

模板中用 {{字段名}} 作为占位符,字段名对应数据源 Excel 首行的表头;
用 {{图片:字段名}}(或 {{img:字段名}})插入图片,字段值为本地图片路径。
"""
from __future__ import annotations

import re
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu
from openpyxl import load_workbook
from PIL import Image

PLACEHOLDER_RE = re.compile(r'\{\{(.+?)\}\}')
IMAGE_PLACEHOLDER_RE = re.compile(r'\{\{\s*(?:img|图片)\s*[:：]\s*(.+?)\s*\}\}')

_MAX_IMAGE_PX = 420  # 插入图片的最大显示宽度(约 11cm @96dpi),更小的按原尺寸
_PX_TO_EMU = 9525


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
    """替换段落中的占位符。

    优先逐 run 替换以保留字体/颜色格式；若占位符跨 run 断裂，
    则合并所有 run 文本后替换，写回首 run（格式退化为首 run 样式）。
    """
    # 先尝试逐 run 替换
    changed = False
    for run in paragraph.runs:
        if PLACEHOLDER_RE.search(run.text):
            new_text = PLACEHOLDER_RE.sub(
                lambda m: record.get(m.group(1).strip(), m.group(0)), run.text)
            if new_text != run.text:
                run.text = new_text
                changed = True
    if changed:
        return

    # 占位符跨 run：合并后替换
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


def _fill_images_in_paragraph(paragraph, record: dict[str, str]) -> int:
    """处理段落中的图片占位符,返回成功插入的图片数。

    段落按占位符切分后重建:文字部分恢复为普通 run,图片部分插入
    内嵌图片(超宽图片等比缩到 _MAX_IMAGE_PX)。图片路径无效时该
    占位符替换为空,不影响其余内容生成。
    """
    full = paragraph.text
    if not IMAGE_PLACEHOLDER_RE.search(full):
        return 0

    parts: list[tuple[str, str]] = []  # ('text', 文本) / ('image', 路径)
    inserted = 0
    pos = 0
    for m in IMAGE_PLACEHOLDER_RE.finditer(full):
        if m.start() > pos:
            parts.append(('text', full[pos:m.start()]))
        field = m.group(1).strip()
        path = record.get(field, '').strip()
        if path and Path(path).is_file():
            parts.append(('image', path))
            inserted += 1
        # 路径无效:直接丢弃占位符
        pos = m.end()
    if pos < len(full):
        parts.append(('text', full[pos:]))

    # 记住原段落首个 run 的字符格式,重建文字时套用,避免丢失字体设置
    fmt = None
    if paragraph.runs:
        rpr = paragraph.runs[0]._r.find(qn('w:rPr'))
        if rpr is not None:
            fmt = deepcopy(rpr)
    for run in paragraph.runs:  # 清空原文本,准备重建
        run.text = ''
    for kind, value in parts:
        if kind == 'text':
            run = paragraph.add_run(value)
            if fmt is not None:
                run._r.insert(0, deepcopy(fmt))
        else:
            _add_image_run(paragraph, value)
    return inserted


def _add_image_run(paragraph, image_path: str) -> None:
    """向段落追加一个内嵌图片 run,过宽的图片等比缩小。"""
    run = paragraph.add_run()
    try:
        with Image.open(image_path) as im:
            width_px = im.size[0]
    except Exception:
        width_px = 0
    if width_px > _MAX_IMAGE_PX:
        run.add_picture(image_path, width=Emu(_MAX_IMAGE_PX * _PX_TO_EMU))
    else:
        run.add_picture(image_path)


def _fill_doc(doc: Document, record: dict[str, str]) -> None:
    def walk(paragraphs):
        for p in paragraphs:
            _fill_paragraph(p, record)
            _fill_images_in_paragraph(p, record)

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
    if not template_path.exists():
        raise FileNotFoundError(f'模板文件不存在: {template_path.name}')
    if not Path(data_path).exists():
        raise FileNotFoundError(f'数据源文件不存在: {Path(data_path).name}')
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    headers, rows = load_data_rows(data_path)
    if not headers or all(not h for h in headers):
        raise ValueError('数据源首行没有有效的表头')
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
