"""主窗口:多标签页工具箱,含使用习惯记忆与浅色/深色主题切换。"""
from PySide6.QtCore import QSettings
from PySide6.QtGui import QAction, QActionGroup, QIcon
from PySide6.QtWidgets import (QCheckBox, QComboBox, QLineEdit, QMainWindow,
                               QMenu, QSpinBox, QTabWidget)

from . import styles
from .tabs import (ConvertTab, ExtractTab, HelpTab, MergeExcelTab, RenameTab,
                   ReplaceTab, SplitMergeTab, TemplateTab)
from .widgets import icon_path


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Office 自动化工具箱')
        self.setWindowIcon(QIcon(icon_path()))
        self.resize(980, 700)
        self.setMinimumSize(860, 600)

        self.settings = QSettings('OfficeToolbox', 'settings')
        self._dark = str(self.settings.value('theme/dark', 'false')) == 'true'
        styles.apply(self, self._dark)

        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        tabs.addTab(ExtractTab(), '文本提取')
        tabs.addTab(ReplaceTab(), '批量替换')
        tabs.addTab(MergeExcelTab(), 'Excel 合并')
        tabs.addTab(TemplateTab(), '模板生成')
        tabs.addTab(ConvertTab(), '格式转换')
        tabs.addTab(SplitMergeTab(), '合并拆分')
        tabs.addTab(RenameTab(), '批量重命名')
        tabs.addTab(HelpTab(), '使用帮助')
        self.setCentralWidget(tabs)

        self._build_menu()
        self._restore_settings()

    # ---------- 主题菜单 ----------
    def _build_menu(self):
        view_menu: QMenu = self.menuBar().addMenu('视图')
        group = QActionGroup(self)
        group.setExclusive(True)
        self._act_light = QAction('浅色主题', self, checkable=True, checked=not self._dark)
        self._act_dark = QAction('深色主题', self, checkable=True, checked=self._dark)
        group.addAction(self._act_light)
        group.addAction(self._act_dark)
        self._act_light.triggered.connect(lambda: self._set_theme(False))
        self._act_dark.triggered.connect(lambda: self._set_theme(True))
        view_menu.addAction(self._act_light)
        view_menu.addAction(self._act_dark)

    def _set_theme(self, dark: bool):
        if dark == self._dark:
            return
        self._dark = dark
        self.settings.setValue('theme/dark', dark)
        styles.apply(self, dark)

    # ---------- 使用习惯记忆 ----------
    def _persistable_widgets(self):
        widgets = []
        for cls in (QLineEdit, QComboBox, QSpinBox, QCheckBox):
            widgets.extend(w for w in self.findChildren(cls) if w.objectName())
        return widgets

    def _restore_settings(self):
        geo = self.settings.value('geometry')
        if geo is not None:
            self.restoreGeometry(geo)
        for w in self._persistable_widgets():
            value = self.settings.value(f'field/{w.objectName()}')
            if value is None:
                continue
            try:
                if isinstance(w, QLineEdit):
                    w.setText(str(value))
                elif isinstance(w, QComboBox):
                    idx = w.findText(str(value))
                    w.setCurrentIndex(idx if idx >= 0 else 0)
                elif isinstance(w, QSpinBox):
                    w.setValue(int(value))
                elif isinstance(w, QCheckBox):
                    w.setChecked(str(value) == 'true')
            except Exception:
                pass

    def closeEvent(self, event):
        self.settings.setValue('geometry', self.saveGeometry())
        for w in self._persistable_widgets():
            if isinstance(w, QLineEdit):
                value = w.text()
            elif isinstance(w, QComboBox):
                value = w.currentText()
            elif isinstance(w, QSpinBox):
                value = w.value()
            else:
                value = w.isChecked()
            self.settings.setValue(f'field/{w.objectName()}', value)
        event.accept()
