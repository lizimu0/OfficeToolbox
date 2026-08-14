# -*- coding: utf-8 -*-
"""核心功能冒烟测试:生成样例文件并逐一验证各模块。"""
import sys
import tempfile
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from pptx import Presentation

from office_toolbox.core import extract, merge_excel, pdf_tools, rename, replace, template
from office_toolbox.core.convert import csv_to_xlsx

tmp = Path(tempfile.mkdtemp(prefix='office_tool_test_'))
print(f'测试目录: {tmp}')


def check(name, cond):
    print(('  [通过] ' if cond else '  [失败] ') + name)
    if not cond:
        sys.exit(1)


# ---- 生成样例文件 ----
doc = Document()
doc.add_paragraph('甲方:某某科技有限公司')
doc.add_paragraph('合同金额为人民币100万元。')
table = doc.add_table(rows=1, cols=2)
table.cell(0, 0).text = '项目'
table.cell(0, 1).text = '备注'
docx_path = tmp / '样例.docx'
doc.save(str(docx_path))

wb = Workbook()
ws = wb.active
ws.append(['姓名', '部门', '工资'])
ws.append(['张三', '研发部', 12000])
ws.append(['李四', '市场部', 9000])
xlsx_path = tmp / '样例.xlsx'
wb.save(str(xlsx_path))

wb2 = Workbook()
ws2 = wb2.active
ws2.append(['姓名', '部门', '工资'])
ws2.append(['王五', '财务部', 10000])
xlsx_path2 = tmp / '样例2.xlsx'
wb2.save(str(xlsx_path2))

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = '季度汇报'
slide.placeholders[1].text = '某某科技有限公司业绩'
pptx_path = tmp / '样例.pptx'
prs.save(str(pptx_path))

# ---- 1. 文本提取 ----
print('1. 文本提取')
check('docx 提取包含段落与表格', '某某科技有限公司' in extract.extract_text(docx_path)
      and '备注' in extract.extract_text(docx_path))
check('xlsx 提取包含表头与数据', '张三' in extract.extract_text(xlsx_path)
      and '部门' in extract.extract_text(xlsx_path))
check('pptx 提取包含标题与正文', '季度汇报' in extract.extract_text(pptx_path))

# ---- 2. 批量替换 ----
print('2. 批量替换')
n = replace.replace_in_file(docx_path, [('某某科技有限公司', '星河集团')])
check(f'docx 替换 {n} 处', n >= 1 and '星河集团' in extract.extract_text(docx_path))
n = replace.replace_in_file(xlsx_path, [('研发部', '研发中心')])
check(f'xlsx 替换 {n} 处', n == 1 and '研发中心' in extract.extract_text(xlsx_path))
n = replace.replace_in_file(pptx_path, [('某某科技', '星河')])
check(f'pptx 替换 {n} 处', n >= 1 and '星河' in extract.extract_text(pptx_path))

# ---- 3. Excel 合并 ----
print('3. Excel 合并')
out_sheets = tmp / '合并_按表.xlsx'
count = merge_excel.merge_as_sheets([xlsx_path, xlsx_path2], out_sheets)
check(f'按工作表合并({count} 个表)', count == 2 and out_sheets.exists())
out_rows = tmp / '合并_按行.xlsx'
rows = merge_excel.merge_append_rows([xlsx_path, xlsx_path2], out_rows)
text = extract.extract_text(out_rows)
check(f'按行追加合并({rows} 行)', rows == 3 and '张三' in text and '王五' in text)

# ---- 4. 模板批量生成 ----
print('4. 模板批量生成')
tpl = Document()
tpl.add_paragraph('聘用通知:{{姓名}}')
tpl.add_paragraph('部门:{{部门}},薪资:{{工资}}')
tpl_path = tmp / '模板.docx'
tpl.save(str(tpl_path))
out_dir = tmp / '生成结果'
files = template.generate_from_template(tpl_path, xlsx_path, out_dir, name_field='姓名')
check(f'生成 {len(files)} 个文档', len(files) == 2 and files == ['张三.docx', '李四.docx'])
check('文档内容已填充', '张三' in extract.extract_text(out_dir / '张三.docx')
      and '{{' not in extract.extract_text(out_dir / '张三.docx'))

# ---- 5. CSV → XLSX ----
print('5. CSV 转 XLSX')
csv_path = tmp / '样例.csv'
csv_path.write_text('编号,名称\n1,测试项\n', encoding='utf-8')
xlsx_out = tmp / 'csv转换.xlsx'
csv_to_xlsx(csv_path, xlsx_out)
check('csv 转 xlsx', '测试项' in extract.extract_text(xlsx_out))

# ---- 6. PDF 合并拆分与 Word 合并 ----
print('6. PDF 合并拆分 / Word 合并')
from pypdf import PdfReader, PdfWriter  # noqa: E402


def make_pdf(path, pages):
    w = PdfWriter()
    for _ in range(pages):
        w.add_blank_page(width=200, height=300)
    with open(path, 'wb') as f:
        w.write(f)


pdf1, pdf2 = tmp / 'a.pdf', tmp / 'b.pdf'
make_pdf(pdf1, 2)
make_pdf(pdf2, 3)
merged_pdf = tmp / '合并.pdf'
n = pdf_tools.merge_pdfs([pdf1, pdf2], merged_pdf)
check(f'PDF 合并({n} 页)', n == 5 and len(PdfReader(str(merged_pdf)).pages) == 5)
split_out = tmp / '拆分结果'
files = pdf_tools.split_pdf(merged_pdf, split_out, pages_per_file=2)
check(f'PDF 拆分(每 2 页一份, {len(files)} 个)', len(files) == 3
      and all((split_out / f).exists() for f in files))

d1 = Document()
d1.add_paragraph('第一篇内容')
d1_path = tmp / 'w1.docx'
d1.save(str(d1_path))
d2 = Document()
d2.add_paragraph('第二篇内容')
d2_path = tmp / 'w2.docx'
d2.save(str(d2_path))
word_out = tmp / '合并文档.docx'
pdf_tools.merge_word([d1_path, d2_path], word_out)
text = extract.extract_text(word_out)
check('Word 合并', '第一篇内容' in text and '第二篇内容' in text)

# ---- 7. 批量重命名 ----
print('7. 批量重命名')
rn_dir = tmp / '重命名'
rn_dir.mkdir()
f1 = rn_dir / '旧名称_报告.docx'
f2 = rn_dir / '旧名称_合同.docx'
f1.write_bytes(Path(d1_path).read_bytes())
f2.write_bytes(Path(d2_path).read_bytes())
res = rename.rename_by_replace([f1, f2], '旧名称', '新名称')
check('文件名文本替换', len(res) == 2 and (rn_dir / '新名称_报告.docx').exists())
res = rename.rename_by_sequence(sorted(rn_dir.glob('*.docx')), '文档_', start=1)
check('前缀序号重命名', (rn_dir / '文档_001.docx').exists() and (rn_dir / '文档_002.docx').exists())
map_wb = Workbook()
map_ws = map_wb.active
map_ws.append(['原文件名', '新文件名'])
map_ws.append(['文档_001', '第一季度报告'])
map_path = tmp / '映射表.xlsx'
map_wb.save(str(map_path))
renamed, skipped = rename.rename_by_mapping(sorted(rn_dir.glob('*.docx')), map_path)
check('Excel 映射重命名', len(renamed) == 1 and len(skipped) == 1
      and (rn_dir / '第一季度报告.docx').exists())

# ---- 8. GUI 可导入 ----
print('8. GUI 导入检查')
from office_toolbox.gui.main_window import MainWindow  # noqa: E402
check('主窗口模块导入成功', MainWindow is not None)

print('\n全部测试通过')
