"""PDF 处理与 Word 合并模块。"""
from __future__ import annotations

from pathlib import Path

from docxcompose.composer import Composer
from docx import Document
from pypdf import PdfReader, PdfWriter


def merge_pdfs(paths: list[str | Path], output: str | Path) -> int:
    """按列表顺序合并多个 PDF,返回合并的页数。"""
    if not paths:
        raise ValueError('没有可合并的 PDF 文件')
    writer = PdfWriter()
    total = 0
    for p in paths:
        p = Path(p)
        if p.suffix.lower() != '.pdf':
            raise ValueError(f'PDF 合并只接受 .pdf 文件,但发现「{p.name}」')
        if not p.exists():
            raise FileNotFoundError(f'文件不存在: {p.name}')
        try:
            reader = PdfReader(str(p))
        except Exception as exc:
            raise ValueError(f'无法读取 PDF「{p.name}」: {exc}') from exc
        for page in reader.pages:
            writer.add_page(page)
            total += 1
    if total == 0:
        raise ValueError('没有可合并的 PDF 页面')
    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, 'wb') as f:
        writer.write(f)
    return total


def split_pdf(path: str | Path, out_dir: str | Path,
              pages_per_file: int = 1) -> list[str]:
    """拆分 PDF:每 pages_per_file 页一个文件,返回生成的文件名列表。"""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f'文件不存在: {path.name}')
    if pages_per_file < 1:
        raise ValueError('每份页数必须大于等于 1')
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        raise ValueError(f'无法读取 PDF「{path.name}」: {exc}') from exc
    pages = reader.pages
    generated: list[str] = []
    for start in range(0, len(pages), pages_per_file):
        writer = PdfWriter()
        for page in pages[start:start + pages_per_file]:
            writer.add_page(page)
        if pages_per_file == 1:
            name = f'{path.stem}_第{start + 1}页.pdf'
        else:
            name = f'{path.stem}_第{start + 1}-{start + len(writer.pages)}页.pdf'
        with open(out_dir / name, 'wb') as f:
            writer.write(f)
        generated.append(name)
    return generated


def merge_word(paths: list[str | Path], output: str | Path) -> int:
    """按列表顺序合并多个 docx 为一个文档,返回合并的文档数。"""
    if not paths:
        raise ValueError('没有可合并的文档')
    for p in paths:
        p = Path(p)
        if p.suffix.lower() != '.docx':
            raise ValueError(f'Word 合并只接受 .docx 文件,但发现「{p.name}」')
        if not p.exists():
            raise FileNotFoundError(f'文件不存在: {p.name}')
    master = Document(str(paths[0]))
    composer = Composer(master)
    for p in paths[1:]:
        composer.append(Document(str(p)))
    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    composer.save(str(out))
    return len(paths)
