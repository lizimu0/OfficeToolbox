"""全局样式主题:现代卡片化风格。"""

ACCENT = '#4f6ef7'

QSS = """
* {
    font-family: 'Microsoft YaHei UI', 'Segoe UI', sans-serif;
    font-size: 13px;
    outline: none;
}

QMainWindow, QWidget {
    background-color: #f4f6fb;
    color: #3c4353;
}

QLabel {
    background: transparent;
}

/* ---------- 标签页 ---------- */
QTabWidget::pane {
    border: none;
    background: #f4f6fb;
}
QTabBar {
    background: transparent;
}
QTabBar::tab {
    background: transparent;
    color: #6b7484;
    padding: 10px 20px;
    margin: 6px 3px 8px 3px;
    border-radius: 9px;
}
QTabBar::tab:hover {
    background: #e9edf8;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: """ + ACCENT + """;
    font-weight: bold;
}

/* ---------- 按钮 ---------- */
QPushButton {
    background-color: #ffffff;
    border: 1px solid #d9dee8;
    border-radius: 8px;
    padding: 7px 16px;
    color: #3c4353;
}
QPushButton:hover {
    border-color: """ + ACCENT + """;
    color: """ + ACCENT + """;
}
QPushButton:pressed {
    background-color: #eef1fb;
}
QPushButton:disabled {
    color: #a3aabb;
    border-color: #e5e8f0;
}

QPushButton[class="primary"] {
    background-color: """ + ACCENT + """;
    border: none;
    border-radius: 9px;
    color: white;
    font-weight: bold;
    font-size: 14px;
    padding: 11px;
}
QPushButton[class="primary"]:hover {
    background-color: #3f5ce0;
    color: white;
}
QPushButton[class="primary"]:pressed {
    background-color: #3650c4;
}
QPushButton[class="primary"]:disabled {
    background-color: #b9c4f2;
    color: white;
}

/* ---------- 输入与列表 ---------- */
QListWidget, QTableWidget, QLineEdit, QComboBox, QSpinBox {
    background-color: #ffffff;
    border: 1px solid #e1e5ee;
    border-radius: 8px;
    padding: 4px 8px;
    selection-background-color: #dfe6ff;
    selection-color: #2f3d8c;
}
QListWidget:hover, QLineEdit:hover, QComboBox:hover {
    border-color: #c3cdf0;
}
QListWidget:focus, QLineEdit:focus, QComboBox:focus {
    border-color: """ + ACCENT + """;
}
QListWidget::item {
    padding: 5px 6px;
    border-radius: 6px;
}
QListWidget::item:selected {
    background-color: #e8eeff;
    color: #2f54eb;
}

QComboBox::drop-down {
    border: none;
    width: 24px;
}
QComboBox::down-arrow {
    image: url(__ARROW__);
    width: 12px;
    height: 12px;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
    border: 1px solid #e1e5ee;
    border-radius: 8px;
    selection-background-color: #e8eeff;
    selection-color: #2f54eb;
}

/* ---------- 表格 ---------- */
QHeaderView::section {
    background-color: #f0f2f9;
    color: #5b6472;
    font-weight: bold;
    border: none;
    border-bottom: 1px solid #e1e5ee;
    padding: 6px;
}
QTableWidget::item {
    padding: 4px;
}
QTableWidget::item:selected {
    background-color: #e8eeff;
    color: #2f54eb;
}

/* ---------- 分组框 ---------- */
QGroupBox {
    background-color: #ffffff;
    border: 1px solid #e1e5ee;
    border-radius: 10px;
    margin-top: 14px;
    padding-top: 8px;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 6px;
    color: """ + ACCENT + """;
    background-color: #f4f6fb;
}

/* ---------- 进度条 ---------- */
QProgressBar {
    background-color: #e7eaf3;
    border: none;
    border-radius: 7px;
    min-height: 14px;
    max-height: 14px;
    text-align: center;
    color: transparent;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 """ + ACCENT + """, stop:1 #7aa0ff);
    border-radius: 7px;
}

/* ---------- 执行日志(浅色,融入整体风格) ---------- */
QPlainTextEdit#logPanel {
    background-color: #ffffff;
    color: #5b6472;
    border: 1px solid #e1e5ee;
    border-radius: 10px;
    padding: 8px;
    font-family: 'Microsoft YaHei UI', sans-serif;
    font-size: 12px;
}

/* ---------- 滚动条 ---------- */
QScrollBar:vertical {
    background: transparent;
    width: 10px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background: #c6cddd;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background: #aeb8ce;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QScrollBar:horizontal {
    background: transparent;
    height: 10px;
    margin: 2px;
}
QScrollBar::handle:horizontal {
    background: #c6cddd;
    border-radius: 5px;
    min-width: 30px;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}

/* ---------- 单选框(圆圈 + 黑色大圆点) ---------- */
QRadioButton {
    spacing: 6px;
    background: transparent;
}
QRadioButton::indicator {
    width: 18px;
    height: 18px;
    image: url(__RADIO_OFF__);
}
QRadioButton::indicator:checked {
    image: url(__RADIO_ON__);
}
"""


def apply(widget) -> None:
    from .widgets import resources_dir
    res = resources_dir()
    qss = (QSS
           .replace('__ARROW__', (res / 'arrow.png').as_posix())
           .replace('__RADIO_OFF__', (res / 'radio_unchecked.png').as_posix())
           .replace('__RADIO_ON__', (res / 'radio_checked.png').as_posix()))
    widget.setStyleSheet(qss)
