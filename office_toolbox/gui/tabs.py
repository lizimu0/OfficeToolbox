"""各功能标签页:文本提取、批量替换、Excel 合并、模板生成、格式转换、
PDF/Word 合并拆分、批量重命名、使用帮助。"""
from __future__ import annotations

import shutil
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QComboBox, QFileDialog, QFormLayout,
                               QGroupBox, QHBoxLayout, QHeaderView, QLabel,
                               QLineEdit, QMessageBox, QProgressBar,
                               QPushButton, QRadioButton, QSpinBox,
                               QTableWidget, QTableWidgetItem, QTextBrowser,
                               QVBoxLayout, QWidget)

from ..core import convert, extract, merge_excel, pdf_tools, rename, replace, template
from .widgets import FileListWidget, LogPanel, TaskThread, notify


class BaseTab(QWidget):
    """标签页基类:统一持有日志面板、进度条、取消按钮和任务线程管理。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._thread: TaskThread | None = None
        self.log_panel = LogPanel()
        self.progress = QProgressBar()
        self.progress.setTextVisible(True)
        self.btn_run = QPushButton('开始执行')
        self.btn_run.setProperty('class', 'primary')
        self.btn_run.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_cancel = QPushButton('取消')
        self.btn_cancel.setProperty('class', 'danger')
        self.btn_cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_cancel.setVisible(False)
        self.btn_cancel.clicked.connect(self._on_cancel_clicked)

    def run_row(self) -> QHBoxLayout:
        """“开始执行 + 取消”按钮行,各标签页布局统一使用。"""
        row = QHBoxLayout()
        row.addWidget(self.btn_run, 1)
        row.addWidget(self.btn_cancel)
        return row

    def start_task(self, func, *args):
        if self._thread is not None and self._thread.isRunning():
            QMessageBox.information(self, '提示', '当前有任务正在执行,请稍候。')
            return False
        self.btn_run.setVisible(False)
        self.btn_cancel.setEnabled(True)
        self.btn_cancel.setVisible(True)
        self.progress.setValue(0)
        self._thread = TaskThread(func, *args, parent=self)
        self._thread.progress.connect(self._on_progress)
        self._thread.finished_ok.connect(self._on_success)
        self._thread.failed.connect(self._on_failed)
        self._thread.cancelled.connect(self._on_cancelled)
        self._thread.start()
        return True

    def _task_done(self):
        self.btn_cancel.setVisible(False)
        self.btn_run.setVisible(True)
        self.btn_run.setEnabled(True)

    def _on_cancel_clicked(self):
        if self._thread is not None and self._thread.isRunning():
            self.btn_cancel.setEnabled(False)
            self.log_panel.log('正在取消,等待当前文件处理完成...')
            self._thread.request_cancel()

    def _on_progress(self, current: int, total: int, message: str):
        if total > 0:
            self.progress.setMaximum(total)
            self.progress.setValue(current)
        self.log_panel.log(message)

    def _on_success(self, result):
        self._task_done()
        self.progress.setValue(self.progress.maximum())
        self.log_panel.log(f'任务完成。{result or ""}')
        notify('任务完成', str(result or ''))

    def _on_failed(self, error: str):
        self._task_done()
        self.log_panel.log(f'任务失败: {error}')
        notify('任务失败', error)
        QMessageBox.critical(self, '执行失败', error)

    def _on_cancelled(self):
        self._task_done()
        self.log_panel.log('任务已取消。')
        notify('任务已取消', '用户取消了本次任务')


class ExtractTab(BaseTab):
    """从 Office 文件中批量提取文字或图片。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_list = FileListWidget()
        self.edit_out = QLineEdit()
        self.edit_out.setObjectName('extract_out')
        btn_browse = QPushButton('浏览...')
        btn_browse.clicked.connect(self._choose_dir)
        self.combo_mode = QComboBox()
        self.combo_mode.addItems(['提取文字(每个文件一个 txt)',
                                  '提取文字(全部合并为一个 txt)',
                                  '提取图片'])

        out_row = QHBoxLayout()
        out_row.addWidget(QLabel('输出目录:'))
        out_row.addWidget(self.edit_out, 1)
        out_row.addWidget(btn_browse)
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel('提取方式:'))
        mode_row.addWidget(self.combo_mode, 1)
        mode_row.addStretch(1)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 6, 16, 14)
        layout.setSpacing(10)
        layout.addWidget(self.file_list, 3)
        layout.addLayout(out_row)
        layout.addLayout(mode_row)
        layout.addLayout(self.run_row())
        layout.addWidget(self.progress)
        layout.addWidget(self.log_panel, 2)
        self.btn_run.clicked.connect(self.run_task)

    def _choose_dir(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self.edit_out.setText(d)

    def run_task(self):
        paths = self.file_list.paths()
        out_dir = self.edit_out.text().strip()
        if not paths:
            QMessageBox.warning(self, '提示', '请先添加文件。')
            return
        if not out_dir:
            QMessageBox.warning(self, '提示', '请选择输出目录。')
            return
        mode = self.combo_mode.currentIndex()

        def task(progress, paths, out_dir, mode):
            out_dir = Path(out_dir)
            out_dir.mkdir(parents=True, exist_ok=True)
            total = len(paths)
            ok, fail = 0, 0
            merged: list[str] = []
            for i, p in enumerate(paths, 1):
                name = Path(p).name
                progress(i, total, f'[{i}/{total}] 处理 {name}')
                try:
                    if mode == 2:
                        saved = extract.extract_images(p, out_dir)
                        progress(i, total, f'    导出图片 {len(saved)} 张')
                    else:
                        text = extract.extract_text(p)
                        if mode == 0:
                            (out_dir / (Path(p).stem + '.txt')).write_text(text, encoding='utf-8')
                        else:
                            merged.append(f'========== {name} ==========\n{text}')
                    ok += 1
                except Exception as exc:
                    progress(i, total, f'    失败: {exc}')
                    fail += 1
            if mode == 1 and merged:
                (out_dir / '合并提取.txt').write_text('\n\n'.join(merged), encoding='utf-8')
            return f'成功 {ok} 个,失败 {fail} 个,结果已保存到 {out_dir}'

        self.start_task(task, paths, out_dir, mode)


class ReplaceTab(BaseTab):
    """批量查找替换。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_list = FileListWidget()

        self.table = QTableWidget(1, 2)
        self.table.setHorizontalHeaderLabels(['查找内容', '替换为'])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        btn_add_pair = QPushButton('添加一行')
        btn_del_pair = QPushButton('删除行')
        btn_add_pair.clicked.connect(lambda: self.table.insertRow(self.table.rowCount()))
        btn_del_pair.clicked.connect(self._del_row)

        self.chk_backup = QCheckBox('替换前自动备份原文件到:')
        self.chk_backup.setObjectName('replace_backup')
        self.chk_backup.setChecked(True)
        self.edit_backup_dir = QLineEdit()
        self.edit_backup_dir.setObjectName('replace_backup_dir')
        self.edit_backup_dir.setPlaceholderText('选择备份目录')
        btn_backup_dir = QPushButton('浏览...')
        btn_backup_dir.clicked.connect(self._choose_backup_dir)
        backup_row = QHBoxLayout()
        backup_row.addWidget(self.chk_backup)
        backup_row.addWidget(self.edit_backup_dir, 1)
        backup_row.addWidget(btn_backup_dir)

        self.chk_regex = QCheckBox('查找内容使用正则表达式(替换为中可用 \\1 引用分组)')
        self.chk_regex.setObjectName('replace_regex')
        self.chk_regex.setToolTip('勾选后“查找内容”按正则表达式解释,例如 \\d{4}年 或 (\\w+)@(\\w+)')

        tip = QLabel('提示:替换将直接保存到原文件。支持 docx / xlsx / pptx。')
        tip.setStyleSheet('color: #888;')

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 6, 16, 14)
        layout.setSpacing(10)
        layout.addWidget(self.file_list, 3)
        pair_box = QGroupBox('替换规则')
        pair_layout = QVBoxLayout(pair_box)
        pair_layout.addWidget(self.table)
        row = QHBoxLayout()
        row.addWidget(btn_add_pair)
        row.addWidget(btn_del_pair)
        row.addStretch(1)
        pair_layout.addLayout(row)
        pair_layout.addWidget(self.chk_regex)
        layout.addWidget(pair_box, 2)
        layout.addLayout(backup_row)
        layout.addWidget(tip)
        layout.addLayout(self.run_row())
        layout.addWidget(self.progress)
        layout.addWidget(self.log_panel, 2)
        self.btn_run.clicked.connect(self.run_task)

    def _del_row(self):
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def _choose_backup_dir(self):
        d = QFileDialog.getExistingDirectory(self, '选择备份目录')
        if d:
            self.edit_backup_dir.setText(d)

    def _pairs(self) -> list[tuple[str, str]]:
        pairs = []
        for r in range(self.table.rowCount()):
            old_item = self.table.item(r, 0)
            new_item = self.table.item(r, 1)
            old = old_item.text() if old_item else ''
            new = new_item.text() if new_item else ''
            if old:
                pairs.append((old, new))
        return pairs

    def run_task(self):
        paths = self.file_list.paths()
        pairs = self._pairs()
        if not paths:
            QMessageBox.warning(self, '提示', '请先添加文件。')
            return
        if not pairs:
            QMessageBox.warning(self, '提示', '请至少填写一条替换规则。')
            return
        backup_dir = self.edit_backup_dir.text().strip() if self.chk_backup.isChecked() else ''
        if self.chk_backup.isChecked() and not backup_dir:
            QMessageBox.warning(self, '提示', '请选择备份目录,或取消勾选备份选项。')
            return
        use_regex = self.chk_regex.isChecked()

        def task(progress, paths, pairs, backup_dir, use_regex):
            import datetime
            total = len(paths)
            ok, fail = 0, 0
            if backup_dir:
                backup_root = Path(backup_dir) / datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_root.mkdir(parents=True, exist_ok=True)
                progress(0, total, f'备份目录: {backup_root}')
            for i, p in enumerate(paths, 1):
                name = Path(p).name
                progress(i, total, f'[{i}/{total}] 处理 {name}')
                try:
                    if backup_dir:
                        shutil.copy2(p, backup_root / name)
                    count = replace.replace_in_file(p, pairs, use_regex=use_regex)
                    progress(i, total, f'    替换 {count} 处')
                    ok += 1
                except Exception as exc:
                    progress(i, total, f'    失败: {exc}')
                    fail += 1
            return f'成功 {ok} 个,失败 {fail} 个'

        self.start_task(task, paths, pairs, backup_dir, use_regex)


class MergeExcelTab(BaseTab):
    """多个 Excel 合并为一个。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_list = FileListWidget('Excel 文件 (*.xlsx *.xlsm)')
        self.radio_sheets = QRadioButton('每个文件作为一个工作表')
        self.radio_rows = QRadioButton('按行追加(表头相同时)')
        self.radio_rows.setChecked(True)
        self.edit_out = QLineEdit()
        self.edit_out.setObjectName('merge_out')
        btn_browse = QPushButton('浏览...')
        btn_browse.clicked.connect(self._choose_save)

        out_row = QHBoxLayout()
        out_row.addWidget(QLabel('保存为:'))
        out_row.addWidget(self.edit_out, 1)
        out_row.addWidget(btn_browse)
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel('合并方式:'))
        mode_row.addWidget(self.radio_sheets)
        mode_row.addWidget(self.radio_rows)
        mode_row.addStretch(1)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 6, 16, 14)
        layout.setSpacing(10)
        layout.addWidget(self.file_list, 3)
        layout.addLayout(mode_row)
        layout.addLayout(out_row)
        layout.addLayout(self.run_row())
        layout.addWidget(self.progress)
        layout.addWidget(self.log_panel, 2)
        self.btn_run.clicked.connect(self.run_task)

    def _choose_save(self):
        path, _ = QFileDialog.getSaveFileName(self, '保存合并结果', '合并结果.xlsx',
                                              'Excel 文件 (*.xlsx)')
        if path:
            self.edit_out.setText(path)

    def run_task(self):
        paths = self.file_list.paths()
        output = self.edit_out.text().strip()
        if not paths:
            QMessageBox.warning(self, '提示', '请先添加 Excel 文件。')
            return
        if not output:
            QMessageBox.warning(self, '提示', '请选择保存位置。')
            return
        as_sheets = self.radio_sheets.isChecked()

        def task(progress, paths, output, as_sheets):
            progress(0, 1, '正在合并...')
            if as_sheets:
                n = merge_excel.merge_as_sheets(paths, output)
                return f'已合并 {n} 个工作表到 {output}'
            n = merge_excel.merge_append_rows(paths, output)
            return f'已按行追加合并 {n} 行数据到 {output}'

        self.start_task(task, paths, output, as_sheets)


class TemplateTab(BaseTab):
    """Word 模板 + Excel 数据源批量生成文档。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.edit_template = QLineEdit()
        self.edit_template.setObjectName('tpl_path')
        self.edit_data = QLineEdit()
        self.edit_data.setObjectName('tpl_data')
        self.edit_out = QLineEdit()
        self.edit_out.setObjectName('tpl_out')
        self.combo_name_field = QComboBox()

        form = QFormLayout()
        form.addRow('Word 模板:', self._pick_row(self.edit_template, self._choose_template))
        form.addRow('Excel 数据源:', self._pick_row(self.edit_data, self._choose_data))
        form.addRow('输出目录:', self._pick_row(self.edit_out, self._choose_dir))
        name_row = QHBoxLayout()
        name_row.addWidget(self.combo_name_field, 1)
        name_row.addWidget(QLabel('(用该字段作为生成文件的文件名)'))
        form.addRow('命名方式:', name_row)

        tip = QLabel('模板中用 {{字段名}} 作为占位符,字段名需与数据源首行表头一致;'
                     '用 {{图片:字段名}} 插入图片,字段值为图片文件路径。')
        tip.setStyleSheet('color: #888;')

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 6, 16, 14)
        layout.setSpacing(10)
        layout.addLayout(form)
        layout.addWidget(tip)
        layout.addLayout(self.run_row())
        layout.addWidget(self.progress)
        layout.addWidget(self.log_panel, 2)
        self.btn_run.clicked.connect(self.run_task)
        self.edit_data.editingFinished.connect(self._refresh_fields)

    def _pick_row(self, edit: QLineEdit, slot) -> QWidget:
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(edit, 1)
        btn = QPushButton('浏览...')
        btn.clicked.connect(slot)
        h.addWidget(btn)
        return row

    def _choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择 Word 模板', '', 'Word 文档 (*.docx)')
        if path:
            self.edit_template.setText(path)

    def _choose_data(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择 Excel 数据源', '', 'Excel 文件 (*.xlsx)')
        if path:
            self.edit_data.setText(path)
            self._refresh_fields()

    def _choose_dir(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self.edit_out.setText(d)

    def _refresh_fields(self):
        data_path = self.edit_data.text().strip()
        self.combo_name_field.clear()
        self.combo_name_field.addItem('序号(文档1、文档2...)')
        if data_path and Path(data_path).exists():
            try:
                headers, _ = template.load_data_rows(data_path)
                self.combo_name_field.addItems(headers)
            except Exception:
                pass

    def run_task(self):
        tpl = self.edit_template.text().strip()
        data = self.edit_data.text().strip()
        out_dir = self.edit_out.text().strip()
        if not tpl or not data or not out_dir:
            QMessageBox.warning(self, '提示', '请先完整选择模板、数据源和输出目录。')
            return
        name_field = None
        if self.combo_name_field.currentIndex() > 0:
            name_field = self.combo_name_field.currentText()

        def task(progress, tpl, data, out_dir, name_field):
            progress(0, 1, '正在读取数据并生成文档...')
            files = template.generate_from_template(tpl, data, out_dir, name_field)
            return f'共生成 {len(files)} 个文档到 {out_dir}'

        self.start_task(task, tpl, data, out_dir, name_field)


class ConvertTab(BaseTab):
    """批量格式转换(docx/xlsx/pptx → PDF 等,需要本机 Office)。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_list = FileListWidget(
            'Office 文档 (*.docx *.doc *.xlsx *.xls *.pptx *.ppt *.csv)')
        self.combo_target = QComboBox()
        self.combo_target.addItems(['.pdf', '.txt', '.xlsx'])
        self.combo_target.setCurrentIndex(0)

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel('目标格式:'))
        mode_row.addWidget(self.combo_target)
        mode_row.addStretch(1)

        tip = QLabel('转 PDF/TXT 需要本机安装 Microsoft Office;csv → xlsx 无需 Office。')
        tip.setStyleSheet('color: #888;')

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 6, 16, 14)
        layout.setSpacing(10)
        layout.addWidget(self.file_list, 3)
        layout.addLayout(mode_row)
        layout.addWidget(tip)
        layout.addLayout(self.run_row())
        layout.addWidget(self.progress)
        layout.addWidget(self.log_panel, 2)
        self.btn_run.clicked.connect(self.run_task)

    def run_task(self):
        paths = self.file_list.paths()
        target = self.combo_target.currentText()
        if not paths:
            QMessageBox.warning(self, '提示', '请先添加文件。')
            return

        def task(progress, paths, target):
            total = len(paths)
            ok, skip, fail = 0, 0, 0
            session = convert.ConvertSession()
            try:
                for i, p in enumerate(paths, 1):
                    name = Path(p).name
                    progress(i, total, f'[{i}/{total}] 转换 {name}')
                    if target not in convert.target_extensions(p):
                        progress(i, total, f'    跳过:不支持 {Path(p).suffix} → {target}')
                        skip += 1
                        continue
                    dst = str(Path(p).with_suffix(target))
                    try:
                        session.convert(p, dst)
                        progress(i, total, f'    已生成 {Path(dst).name}')
                        ok += 1
                    except Exception as exc:
                        progress(i, total, f'    失败: {exc}')
                        fail += 1
            finally:
                session.close()
            return f'成功 {ok} 个,跳过 {skip} 个,失败 {fail} 个'

        self.start_task(task, paths, target)


class SplitMergeTab(BaseTab):
    """PDF 合并/拆分与 Word 合并。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.combo_mode = QComboBox()
        self.combo_mode.setObjectName('splitmerge_mode')
        self.combo_mode.addItems(['PDF 合并', 'PDF 拆分', 'Word 合并'])
        self.combo_mode.currentIndexChanged.connect(self._mode_changed)

        self.file_list = FileListWidget('PDF / Word (*.pdf *.docx)')

        self.spin_pages = QSpinBox()
        self.spin_pages.setRange(1, 9999)
        self.spin_pages.setValue(1)
        self.lbl_pages = QLabel('每份页数:')

        self.edit_src = QLineEdit()
        self.edit_src.setObjectName('splitmerge_src')
        self.edit_src.setPlaceholderText('选择要拆分的 PDF 文件')
        btn_src = QPushButton('浏览...')
        btn_src.clicked.connect(self._choose_src)

        self.edit_out = QLineEdit()
        self.edit_out.setObjectName('splitmerge_out')
        btn_browse = QPushButton('浏览...')
        btn_browse.clicked.connect(self._choose_out)

        self.lbl_out = QLabel('保存为:')

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel('操作类型:'))
        mode_row.addWidget(self.combo_mode)
        mode_row.addWidget(self.lbl_pages)
        mode_row.addWidget(self.spin_pages)
        mode_row.addStretch(1)
        src_row = QHBoxLayout()
        src_row.addWidget(QLabel('源文件:'))
        src_row.addWidget(self.edit_src, 1)
        src_row.addWidget(btn_src)
        out_row = QHBoxLayout()
        out_row.addWidget(self.lbl_out)
        out_row.addWidget(self.edit_out, 1)
        out_row.addWidget(btn_browse)
        self.src_row = src_row

        self.tip = QLabel()
        self.tip.setStyleSheet('color: #888;')

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 6, 16, 14)
        layout.setSpacing(10)
        layout.addLayout(mode_row)
        layout.addWidget(self.file_list, 3)
        layout.addLayout(src_row)
        layout.addLayout(out_row)
        layout.addWidget(self.tip)
        layout.addLayout(self.run_row())
        layout.addWidget(self.progress)
        layout.addWidget(self.log_panel, 2)
        self.btn_run.clicked.connect(self.run_task)
        self._mode_changed()

    def _mode_changed(self, _=None):
        split = self.combo_mode.currentIndex() == 1
        self.spin_pages.setVisible(split)
        self.lbl_pages.setVisible(split)
        self.file_list.setVisible(not split)
        self.edit_src.setVisible(split)
        self._set_row_visible(self.src_row, split)
        if split:
            self.lbl_out.setText('输出目录:')
            self.tip.setText('选择要拆分的 PDF 和输出目录,按“每份页数”拆成多个文件。')
        else:
            self.lbl_out.setText('保存为:')
            self.tip.setText('按列表顺序合并;PDF 合并请只添加 PDF,Word 合并请只添加 docx。')

    def _set_row_visible(self, row, visible: bool):
        for i in range(row.count()):
            w = row.itemAt(i).widget()
            if w is not None:
                w.setVisible(visible)

    def _choose_src(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择要拆分的 PDF', '', 'PDF 文件 (*.pdf)')
        if path:
            self.edit_src.setText(path)

    def _choose_out(self):
        if self.combo_mode.currentIndex() == 1:
            d = QFileDialog.getExistingDirectory(self, '选择输出目录')
            if d:
                self.edit_out.setText(d)
        else:
            if self.combo_mode.currentIndex() == 0:
                flt, default = 'PDF 文件 (*.pdf)', '合并.pdf'
            else:
                flt, default = 'Word 文档 (*.docx)', '合并.docx'
            path, _ = QFileDialog.getSaveFileName(self, '保存合并结果', default, flt)
            if path:
                self.edit_out.setText(path)

    def run_task(self):
        mode = self.combo_mode.currentIndex()
        out = self.edit_out.text().strip()
        if not out:
            QMessageBox.warning(self, '提示', '请先选择保存位置/输出目录。')
            return
        if mode == 1:
            src = self.edit_src.text().strip()
            if not src:
                QMessageBox.warning(self, '提示', '请选择要拆分的 PDF。')
                return
            pages = self.spin_pages.value()

            def task_split(progress, src, out, pages):
                progress(0, 1, '正在拆分...')
                files = pdf_tools.split_pdf(src, out, pages)
                return f'已拆分为 {len(files)} 个文件到 {out}'

            self.start_task(task_split, src, out, pages)
            return

        paths = self.file_list.paths()
        if len(paths) < 2:
            QMessageBox.warning(self, '提示', '请至少添加两个文件。')
            return

        def task_merge(progress, paths, out, mode):
            progress(0, 1, '正在合并...')
            if mode == 0:
                n = pdf_tools.merge_pdfs(paths, out)
                return f'已合并 {len(paths)} 个 PDF(共 {n} 页)到 {out}'
            n = pdf_tools.merge_word(paths, out)
            return f'已合并 {n} 个 Word 文档到 {out}'

        self.start_task(task_merge, paths, out, mode)


class RenameTab(BaseTab):
    """批量重命名:文本替换、前缀序号、Excel 映射。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_list = FileListWidget('所有文件 (*.*)')

        self.radio_replace = QRadioButton('文件名文本替换')
        self.radio_seq = QRadioButton('前缀 + 序号')
        self.radio_map = QRadioButton('Excel 映射表')
        self.radio_replace.setChecked(True)

        self.edit_old = QLineEdit()
        self.edit_old.setObjectName('rename_old')
        self.edit_new = QLineEdit()
        self.edit_new.setObjectName('rename_new')
        self.edit_prefix = QLineEdit()
        self.edit_prefix.setObjectName('rename_prefix')
        self.edit_prefix.setPlaceholderText('例如:合同_')
        self.spin_start = QSpinBox()
        self.spin_start.setRange(1, 99999)
        self.spin_start.setValue(1)
        self.edit_mapping = QLineEdit()
        self.edit_mapping.setObjectName('rename_mapping')
        btn_mapping = QPushButton('浏览...')
        btn_mapping.clicked.connect(self._choose_mapping)

        row_replace = QHBoxLayout()
        row_replace.addWidget(QLabel('把'))
        row_replace.addWidget(self.edit_old)
        row_replace.addWidget(QLabel('替换为'))
        row_replace.addWidget(self.edit_new)
        row_seq = QHBoxLayout()
        row_seq.addWidget(QLabel('前缀:'))
        row_seq.addWidget(self.edit_prefix)
        row_seq.addWidget(QLabel('起始序号:'))
        row_seq.addWidget(self.spin_start)
        row_seq.addStretch(1)
        row_map = QHBoxLayout()
        row_map.addWidget(self.edit_mapping, 1)
        row_map.addWidget(btn_mapping)

        rule_box = QGroupBox('重命名规则')
        rule_layout = QVBoxLayout(rule_box)
        rule_layout.addWidget(self.radio_replace)
        rule_layout.addLayout(row_replace)
        rule_layout.addWidget(self.radio_seq)
        rule_layout.addLayout(row_seq)
        rule_layout.addWidget(self.radio_map)
        rule_layout.addLayout(row_map)

        tip = QLabel('映射表格式:首行表头,第一列原文件名,第二列新文件名。执行前会弹出预览确认;重命名在原目录进行,重名自动加后缀。')
        tip.setStyleSheet('color: #888;')

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 6, 16, 14)
        layout.setSpacing(10)
        layout.addWidget(self.file_list, 3)
        layout.addWidget(rule_box, 2)
        layout.addWidget(tip)
        layout.addLayout(self.run_row())
        layout.addWidget(self.progress)
        layout.addWidget(self.log_panel, 2)
        self.btn_run.clicked.connect(self.run_task)

    def _choose_mapping(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择映射表', '', 'Excel 文件 (*.xlsx)')
        if path:
            self.edit_mapping.setText(path)

    def _confirm_rename(self, preview_pairs: list[tuple[str, str]],
                        skipped: list[str] | None = None) -> bool:
        """弹出预演确认对话框,返回用户是否同意执行。"""
        if not preview_pairs:
            QMessageBox.information(self, '提示', '没有需要重命名的文件。')
            return False
        lines = [f'{old}  →  {new}' for old, new in preview_pairs[:20]]
        if len(preview_pairs) > 20:
            lines.append(f'... 以及其余 {len(preview_pairs) - 20} 个')
        text = f'将重命名 {len(preview_pairs)} 个文件:\n\n' + '\n'.join(lines)
        if skipped:
            text += f'\n\n未匹配(将跳过){len(skipped)} 个。'
        text += '\n\n确定执行吗?'
        reply = QMessageBox.question(self, '重命名预览', text,
                                     QMessageBox.StandardButton.Yes
                                     | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        return reply == QMessageBox.StandardButton.Yes

    def run_task(self):
        paths = self.file_list.paths()
        if not paths:
            QMessageBox.warning(self, '提示', '请先添加文件。')
            return

        if self.radio_replace.isChecked():
            old, new = self.edit_old.text(), self.edit_new.text()
            if not old:
                QMessageBox.warning(self, '提示', '请填写要替换的文本。')
                return
            preview = rename.rename_by_replace(paths, old, new, preview=True)
            if not self._confirm_rename(preview):
                return

            def task_replace(progress, paths, old, new):
                progress(0, 1, '正在重命名...')
                results = rename.rename_by_replace(paths, old, new)
                return f'已重命名 {len(results)} 个文件'

            self.start_task(task_replace, paths, old, new)
        elif self.radio_seq.isChecked():
            prefix, start = self.edit_prefix.text(), self.spin_start.value()
            preview = rename.rename_by_sequence(paths, prefix, start, preview=True)
            if not self._confirm_rename(preview):
                return

            def task_seq(progress, paths, prefix, start):
                progress(0, 1, '正在重命名...')
                results = rename.rename_by_sequence(paths, prefix, start)
                return f'已重命名 {len(results)} 个文件'

            self.start_task(task_seq, paths, prefix, start)
        else:
            mapping = self.edit_mapping.text().strip()
            if not mapping:
                QMessageBox.warning(self, '提示', '请选择映射表。')
                return
            try:
                preview, skipped = rename.rename_by_mapping(paths, mapping, preview=True)
            except Exception as exc:
                QMessageBox.warning(self, '提示', str(exc))
                return
            if not self._confirm_rename(preview, skipped):
                return

            def task_map(progress, paths, mapping):
                progress(0, 1, '正在重命名...')
                renamed, skipped = rename.rename_by_mapping(paths, mapping)
                msg = f'已重命名 {len(renamed)} 个文件'
                if skipped:
                    msg += f',未匹配 {len(skipped)} 个'
                return msg

            self.start_task(task_map, paths, mapping)


HELP_HTML = """
<h2>Office 自动化工具箱 使用说明</h2>
<p>所有标签页都支持<b>拖拽文件或文件夹</b>到列表区,文件夹会自动递归扫描其中的 Office 文件;任务执行中可随时点击<b>取消</b>按钮中止。单文件出错不会中断整批任务,详见日志。</p>
<h3>文本提取</h3>
<p>从 docx/xlsx/pptx 批量提取文字或导出嵌入图片,可选每个文件一个 txt 或全部合并。</p>
<h3>批量替换</h3>
<p>在多个文件中按规则替换文本,默认先把原文件备份到指定目录(每次执行生成一个时间戳子文件夹)。勾选<b>正则表达式</b>后,查找内容按正则解释,替换为中可用 <code>\\1</code> 引用分组,例如把 <code>(\\d+)年</code> 替换为 <code>\\1年度</code>。</p>
<h3>Excel 合并</h3>
<p>“按行追加”适合表头相同的报表;“每个文件一个工作表”适合汇总不同文件。</p>
<h3>模板生成</h3>
<p>Word 模板中用 <code>{{字段名}}</code> 占位,字段名与数据源 Excel 首行表头一致,每行数据生成一份文档;用 <code>{{图片:字段名}}</code>(或 <code>{{img:字段名}}</code>)插入图片,该字段的值为本地图片文件路径,过宽图片会自动等比缩小。</p>
<h3>格式转换</h3>
<p>docx/xlsx/pptx 转 PDF 需本机安装 Microsoft Office,输出在源文件同目录;csv 转 xlsx 无需 Office。</p>
<h3>合并拆分</h3>
<p>PDF 合并按列表顺序拼接;PDF 拆分可按每 N 页一份;Word 合并把多个 docx 顺序拼接为一篇。</p>
<h3>批量重命名</h3>
<p>三种规则:文件名文本替换、前缀+序号、Excel 映射表(第一列原名、第二列新名)。重名自动加后缀避免覆盖。</p>
<h3>主题</h3>
<p>可在菜单栏“视图”中切换浅色/深色主题,选择会被自动记住。</p>
"""


class HelpTab(QWidget):
    """内置使用帮助。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setHtml(HELP_HTML)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 6, 16, 14)
        layout.addWidget(browser)
