# -*- coding: utf-8 -*-
import os
import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon

from lh_exporter.main_window import MainWindow


def get_icon_path() -> str:
    """获取图标文件路径"""
    # 尝试多个可能的路径
    possible_paths = [
        "生成数据处理软件图片.png",
        "../生成数据处理软件图片.png",
        os.path.join(os.path.dirname(__file__), "生成数据处理软件图片.png"),
        os.path.join(os.path.dirname(__file__), "..", "生成数据处理软件图片.png"),
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return ""


def main():
    os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "1")
    app = QApplication(sys.argv)
    try:
        app.setStyle("Fusion")
    except Exception:
        pass
    try:
        icon_path = get_icon_path()
        if icon_path:
            app.setWindowIcon(QIcon(icon_path))
    except Exception:
        pass
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
