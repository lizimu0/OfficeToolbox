"""GUI 公共组件:文件列表选择器、后台任务线程、日志面板、系统通知。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (QFileDialog, QHBoxLayout, QListWidget,
                               QListWidgetItem, QPlainTextEdit, QPushButton,
                               QSystemTrayIcon, QVBoxLayout, QWidget)

OFFICE_FILTER = 'Office 文档 (*.docx *.xlsx *.xlsm *.pptx);;所有文件 (*.*)'
SCAN_EXTS = {'.docx', '.doc', '.xlsx', '.xlsm', '.xls', '.pptx', '.ppt', '.pdf', '.csv'}


def resources_dir() -> Path:
    """资源目录,兼容源码运行与 PyInstaller 打包两种模式。"""
    import sys
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent / 'resources'
    return Path(__file__).resolve().parent.parent / 'resources'


def icon_path() -> str:
    """程序图标路径。"""
    return str(resources_dir() / 'app.ico')


_tray: QSystemTrayIcon | None = None


def notify(title: str, message: str) -> None:
    """任务完成/失败时弹出系统托盘通知。"""
    global _tray
    try:
        if _tray is None:
            _tray = QSystemTrayIcon(QIcon(icon_path()))
            _tray.show()
        _tray.showMessage(title, message, QSystemTrayIcon.MessageIcon.Information, 5000)
    except Exception:
        pass  # 通知失败不影响主流程


class FileListWidget(QWidget):
    """可添加/移除/清空的文件列表,支持拖拽添加。"""

    def __init__(self, file_filter: str = OFFICE_FILTER, parent=None):
        super().__init__(parent)
        self.file_filter = file_filter
        self.setAcceptDrops(True)

        self.list_widget = QListWidget()
        self.list_widget.setAcceptDrops(True)

        btn_add = QPushButton('添加文件')
        btn_folder = QPushButton('添加文件夹')
        btn_remove = QPushButton('移除选中')
        btn_clear = QPushButton('清空')
        btn_add.clicked.connect(self._add_files)
        btn_folder.clicked.connect(self._add_folder)
        btn_remove.clicked.connect(self._remove_selected)
        btn_clear.clicked.connect(self.list_widget.clear)

        btn_box = QVBoxLayout()
        btn_box.addWidget(btn_add)
        btn_box.addWidget(btn_folder)
        btn_box.addWidget(btn_remove)
        btn_box.addWidget(btn_clear)
        btn_box.addStretch(1)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.list_widget, 1)
        layout.addLayout(btn_box)

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, '选择文件', '', self.file_filter)
        self.add_paths(paths)

    def _add_folder(self):
        d = QFileDialog.getExistingDirectory(self, '选择文件夹')
        if d:
            self.add_paths([d])

    def _remove_selected(self):
        for item in self.list_widget.selectedItems():
            self.list_widget.takeItem(self.list_widget.row(item))

    def add_paths(self, paths) -> None:
        existing = {self.list_widget.item(i).data(0) for i in range(self.list_widget.count())}
        for p in paths:
            p = Path(str(p))
            if p.is_dir():  # 文件夹:递归扫描所有支持的 Office 文件
                for f in sorted(p.rglob('*')):
                    if f.is_file() and f.suffix.lower() in SCAN_EXTS:
                        self._add_one(f, existing)
            else:
                self._add_one(p, existing)

    def _add_one(self, p: Path, existing: set) -> None:
        key = str(p)
        if key not in existing:
            item = QListWidgetItem(p.name)
            item.setData(0, key)
            item.setToolTip(key)
            self.list_widget.addItem(item)
            existing.add(key)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        self.add_paths(u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile())

    def count(self) -> int:
        return self.list_widget.count()

    def paths(self) -> list[str]:
        return [self.list_widget.item(i).data(0) for i in range(self.list_widget.count())]


class LogPanel(QPlainTextEdit):
    """只读日志输出面板。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('logPanel')
        self.setReadOnly(True)
        self.setMaximumBlockCount(2000)
        self.setFixedHeight(110)
        self.setPlaceholderText('执行记录将显示在这里')

    def log(self, message: str) -> None:
        self.appendPlainText(message)


class TaskThread(QThread):
    """通用后台任务线程:执行耗时函数并通过信号回报进度。

    func 接受一个 progress 回调:progress(current, total, message)。
    """

    progress = Signal(int, int, str)   # current, total, message
    finished_ok = Signal(object)        # 任务结果
    failed = Signal(str)                # 错误信息

    def __init__(self, func, *args, parent=None):
        super().__init__(parent)
        self._func = func
        self._args = args

    def run(self):
        try:
            result = self._func(self._report, *self._args)
            self.finished_ok.emit(result)
        except Exception as exc:  # 统一兜底,避免线程异常退出无提示
            self.failed.emit(str(exc))

    def _report(self, current: int, total: int, message: str) -> None:
        self.progress.emit(current, total, message)
