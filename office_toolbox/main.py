"""Office 自动化工具箱 - 程序入口。

运行方式:python -m office_toolbox.main
"""
import sys

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from office_toolbox.gui.main_window import MainWindow
from office_toolbox.gui.widgets import icon_path


def main() -> None:
    # 设置 Windows 应用标识,让任务栏显示自定义图标而非 Python 图标
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            'OfficeToolbox.App')
    except Exception:
        pass
    # 保证日志中的特殊符号在 GBK 控制台下不报错
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(errors='replace')
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setApplicationName('Office 自动化工具箱')
    app.setWindowIcon(QIcon(icon_path()))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
