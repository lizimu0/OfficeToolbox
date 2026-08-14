"""主窗口:多标签页工具箱,含使用习惯记忆。"""
from PySide6.QtCore import QSettings
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (QCheckBox, QComboBox, QLineEdit, QMainWindow,
                               QSpinBox, QTabWidget)

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
        styles.apply(self)

        self.settings = QSettings('OfficeToolbox', 'settings')

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

        self._restore_settings()

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
