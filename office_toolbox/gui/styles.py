"""全局样式主题:现代卡片化风格,支持浅色/深色两套配色。

深色模式下的下拉箭头、单选框图标由 PIL 动态生成并缓存到临时目录。
"""
from pathlib import Path

# 配色方案:各控件样式统一从这里取色,新增主题只需增加一份调色板
LIGHT = {
    'accent': '#4f6ef7',
    'accent_hover': '#3f5ce0',
    'accent_pressed': '#3650c4',
    'accent_disabled': '#b9c4f2',
    'accent_soft': '#eef1fb',
    'window': '#f4f6fb',
    'card': '#ffffff',
    'border': '#e1e5ee',
    'border_hover': '#c3cdf0',
    'border_disabled': '#e5e8f0',
    'text': '#3c4353',
    'text_weak': '#6b7484',
    'text_disabled': '#a3aabb',
    'selected_bg': '#e8eeff',
    'selected_fg': '#2f54eb',
    'header_bg': '#f0f2f9',
    'header_fg': '#5b6472',
    'hover_bg': '#e9edf8',
    'track': '#e7eaf3',
    'track_dark': '#dfe3ef',
    'scroll_handle': '#c6cddd',
    'scroll_hover': '#aeb8ce',
    'gradient_end': '#7aa0ff',
    'log_fg': '#5b6472',
}

DARK = {
    'accent': '#6b8afd',
    'accent_hover': '#7f9dfe',
    'accent_pressed': '#5a7be8',
    'accent_disabled': '#3a4152',
    'accent_soft': '#2a3040',
    'window': '#1e222a',
    'card': '#262b36',
    'border': '#363d4d',
    'border_hover': '#4a5468',
    'border_disabled': '#2e3442',
    'text': '#d7dce5',
    'text_weak': '#98a1b3',
    'text_disabled': '#5b6376',
    'selected_bg': '#32406b',
    'selected_fg': '#c3d2ff',
    'header_bg': '#2b3140',
    'header_fg': '#aab3c5',
    'hover_bg': '#2e3442',
    'track': '#313847',
    'track_dark': '#3a4254',
    'scroll_handle': '#4a5468',
    'scroll_hover': '#5d6880',
    'gradient_end': '#8fb0ff',
    'log_fg': '#aab3c5',
}

QSS_TEMPLATE = """
* {{
    font-family: 'Microsoft YaHei UI', 'Segoe UI', sans-serif;
    font-size: 13px;
    outline: none;
}}

QMainWindow, QWidget {{
    background-color: {window};
    color: {text};
}}

QLabel {{
    background: transparent;
}}

/* ---------- 菜单栏 ---------- */
QMenuBar {{
    background-color: {window};
    color: {text_weak};
    border-bottom: 1px solid {border};
}}
QMenuBar::item {{
    padding: 6px 12px;
    border-radius: 6px;
    background: transparent;
}}
QMenuBar::item:selected {{
    background-color: {hover_bg};
    color: {text};
}}
QMenu {{
    background-color: {card};
    border: 1px solid {border};
    border-radius: 10px;
    padding: 6px;
}}
QMenu::item {{
    padding: 7px 24px 7px 14px;
    border-radius: 6px;
    color: {text};
}}
QMenu::item:selected {{
    background-color: {selected_bg};
    color: {selected_fg};
}}
QMenu::separator {{
    height: 1px;
    background: {border};
    margin: 4px 8px;
}}

/* ---------- 标签页 ---------- */
QTabWidget::pane {{
    border: none;
    background: {window};
}}
QTabBar {{
    background: transparent;
}}
QTabBar::tab {{
    background: transparent;
    color: {text_weak};
    padding: 10px 20px;
    margin: 6px 3px 8px 3px;
    border-radius: 9px;
}}
QTabBar::tab:hover {{
    background: {hover_bg};
}}
QTabBar::tab:selected {{
    background: {card};
    color: {accent};
    font-weight: bold;
}}

/* ---------- 按钮 ---------- */
QPushButton {{
    background-color: {card};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 7px 16px;
    color: {text};
}}
QPushButton:hover {{
    border-color: {accent};
    color: {accent};
}}
QPushButton:pressed {{
    background-color: {accent_soft};
}}
QPushButton:disabled {{
    color: {text_disabled};
    border-color: {border_disabled};
}}

QPushButton[class="primary"] {{
    background-color: {accent};
    border: none;
    border-radius: 9px;
    color: white;
    font-weight: bold;
    font-size: 14px;
    padding: 11px;
}}
QPushButton[class="primary"]:hover {{
    background-color: {accent_hover};
    color: white;
}}
QPushButton[class="primary"]:pressed {{
    background-color: {accent_pressed};
}}
QPushButton[class="primary"]:disabled {{
    background-color: {accent_disabled};
    color: white;
}}

QPushButton[class="danger"] {{
    background-color: {card};
    border: 1px solid #e06666;
    color: #e06666;
    font-weight: bold;
}}
QPushButton[class="danger"]:hover {{
    background-color: #fdeeee;
    color: #c94d4d;
    border-color: #c94d4d;
}}
QPushButton[class="danger"]:pressed {{
    background-color: #f8dcdc;
}}
QPushButton[class="danger"]:disabled {{
    color: {text_disabled};
    border-color: {border_disabled};
    background-color: {card};
}}

/* ---------- 输入与列表 ---------- */
QListWidget, QTableWidget, QLineEdit, QComboBox, QSpinBox {{
    background-color: {card};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 4px 8px;
    selection-background-color: {selected_bg};
    selection-color: {selected_fg};
}}
QListWidget:hover, QLineEdit:hover, QComboBox:hover {{
    border-color: {border_hover};
}}
QListWidget:focus, QLineEdit:focus, QComboBox:focus {{
    border-color: {accent};
}}
QListWidget::item {{
    padding: 5px 6px;
    border-radius: 6px;
}}
QListWidget::item:selected {{
    background-color: {selected_bg};
    color: {selected_fg};
}}

QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox::down-arrow {{
    image: url(__ARROW__);
    width: 12px;
    height: 12px;
}}
QComboBox QAbstractItemView {{
    background-color: {card};
    border: 1px solid {border};
    border-radius: 8px;
    selection-background-color: {selected_bg};
    selection-color: {selected_fg};
}}

/* ---------- 表格 ---------- */
QHeaderView::section {{
    background-color: {header_bg};
    color: {header_fg};
    font-weight: bold;
    border: none;
    border-bottom: 1px solid {border};
    padding: 6px;
}}
QTableWidget::item {{
    padding: 4px;
}}
QTableWidget::item:selected {{
    background-color: {selected_bg};
    color: {selected_fg};
}}

/* ---------- 分组框 ---------- */
QGroupBox {{
    background-color: {card};
    border: 1px solid {border};
    border-radius: 10px;
    margin-top: 14px;
    padding-top: 8px;
    font-weight: bold;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 6px;
    color: {accent};
    background-color: {window};
}}

/* ---------- 进度条 ---------- */
QProgressBar {{
    background-color: {track};
    border: none;
    border-radius: 7px;
    min-height: 14px;
    max-height: 14px;
    text-align: center;
    color: transparent;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 {accent}, stop:1 {gradient_end});
    border-radius: 7px;
}}

/* ---------- 执行日志 ---------- */
QPlainTextEdit#logPanel {{
    background-color: {card};
    color: {log_fg};
    border: 1px solid {border};
    border-radius: 10px;
    padding: 8px;
    font-family: 'Microsoft YaHei UI', sans-serif;
    font-size: 12px;
}}

/* ---------- 帮助页 ---------- */
QTextBrowser {{
    background-color: {card};
    border: 1px solid {border};
    border-radius: 10px;
    padding: 12px;
}}

/* ---------- 滚动条 ---------- */
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px;
}}
QScrollBar::handle:vertical {{
    background: {scroll_handle};
    border-radius: 5px;
    min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{
    background: {scroll_hover};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 10px;
    margin: 2px;
}}
QScrollBar::handle:horizontal {{
    background: {scroll_handle};
    border-radius: 5px;
    min-width: 30px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0px;
}}

/* ---------- 单选框与复选框文字 ---------- */
QRadioButton, QCheckBox {{
    spacing: 6px;
    background: transparent;
}}
QRadioButton::indicator {{
    width: 18px;
    height: 18px;
    image: url(__RADIO_OFF__);
}}
QRadioButton::indicator:checked {{
    image: url(__RADIO_ON__);
}}
QMenu::indicator:checked {{
    background: {accent};
    border-radius: 4px;
    width: 10px;
    height: 10px;
    margin-left: 4px;
}}
"""


def _dark_icons_dir():
    """深色主题图标缓存目录(首次使用时用 PIL 生成)。"""
    import tempfile
    from PIL import Image, ImageDraw

    out = Path(tempfile.gettempdir()) / 'office_toolbox_dark'
    out.mkdir(exist_ok=True)
    arrow = out / 'arrow.png'
    radio_off = out / 'radio_unchecked.png'
    radio_on = out / 'radio_checked.png'
    if not (arrow.exists() and radio_off.exists() and radio_on.exists()):
        # 下拉箭头(浅灰 V 形,深色背景下清晰)
        img = Image.new('RGBA', (128, 128), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.line([(22, 44), (64, 86), (106, 44)],
               fill=(160, 170, 190, 255), width=16, joint='curve')
        img.save(arrow)
        # 单选框:深色底圆圈 + 选中时浅色圆点
        def make_radio(path: Path, checked: bool) -> None:
            im = Image.new('RGBA', (128, 128), (0, 0, 0, 0))
            dr = ImageDraw.Draw(im)
            dr.ellipse([8, 8, 120, 120], fill=(38, 43, 54, 255),
                       outline=(90, 100, 120, 255), width=8)
            if checked:
                dr.ellipse([32, 32, 96, 96], fill=(215, 220, 229, 255))
            im.save(path)

        make_radio(radio_off, False)
        make_radio(radio_on, True)
    return arrow, radio_off, radio_on


def build_qss(palette: dict) -> str:
    """按调色板渲染 QSS(不含图标路径替换)。"""
    return QSS_TEMPLATE.format(**palette)


def apply(widget, dark: bool = False) -> None:
    """应用主题样式。dark=True 时使用深色主题并切换配套图标。"""
    from .widgets import resources_dir

    qss = build_qss(DARK if dark else LIGHT)
    if dark:
        arrow, radio_off, radio_on = _dark_icons_dir()
    else:
        res = resources_dir()
        arrow, radio_off, radio_on = (res / 'arrow.png',
                                      res / 'radio_unchecked.png',
                                      res / 'radio_checked.png')
    qss = (qss
           .replace('__ARROW__', arrow.as_posix())
           .replace('__RADIO_OFF__', radio_off.as_posix())
           .replace('__RADIO_ON__', radio_on.as_posix()))
    widget.setStyleSheet(qss)


# 向后兼容:styles.QSS / styles.ACCENT 仍可访问(浅色主题)
QSS = build_qss(LIGHT)
ACCENT = LIGHT['accent']
