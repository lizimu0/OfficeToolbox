# -*- coding: utf-8 -*-
"""生成界面图片资源:下拉框箭头、单选框(未选中/选中黑圆点)。"""
from pathlib import Path

from PIL import Image, ImageDraw

out = Path(__file__).parent / 'office_toolbox' / 'resources'
out.mkdir(exist_ok=True)

# ---- 下拉框箭头(深灰色 V 形) ----
img = Image.new('RGBA', (128, 128), (0, 0, 0, 0))
d = ImageDraw.Draw(img)
d.line([(22, 44), (64, 86), (106, 44)], fill=(90, 99, 116, 255), width=16, joint='curve')
img.save(out / 'arrow.png')

# ---- 单选框:圆圈 + 选中时内部黑色大圆点 ----
def make_radio(name: str, checked: bool) -> None:
    im = Image.new('RGBA', (128, 128), (0, 0, 0, 0))
    dr = ImageDraw.Draw(im)
    dr.ellipse([8, 8, 120, 120], fill=(255, 255, 255, 255),
               outline=(160, 168, 186, 255), width=8)
    if checked:
        dr.ellipse([32, 32, 96, 96], fill=(20, 24, 34, 255))
    im.save(out / name)


make_radio('radio_unchecked.png', False)
make_radio('radio_checked.png', True)
print('资源已生成:', out)
