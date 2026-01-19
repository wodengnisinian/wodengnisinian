# -*- coding: utf-8 -*-
import sys
from PySide6.QtCore import Qt, QEvent, QObject
from PySide6.QtGui import QAction, QColor, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit,
    QListWidget, QListWidgetItem, QStackedWidget,
    QStatusBar, QFrame, QProgressBar, QStyle
)

APP_NAME = "宿舍普查数据清洗软件"
APP_VERSION = "v1.1.0.14"


class NiceProgressBar(QProgressBar):
    def __init__(self):
        super().__init__()
        self.setObjectName("NiceProgress")
        self.setRange(0, 100)
        self.setValue(20)
        self.setFixedHeight(14)
        self.setTextVisible(True)
        self.setFormat("%p%")
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("""
            QProgressBar#NiceProgress{
                border:none;
                background:#eef0f3;
                border-radius:7px;
                text-align:center;
                color:#2b2f36;
                font-weight:600;
            }
            QProgressBar#NiceProgress::chunk{
                border-radius:7px;
                background:qlineargradient(x1:0,y1:0,x2:1,y2:0,
                                           stop:0 #2f6fed, stop:1 #0b57d0);
            }
        """)


class ClickResetFilter(QObject):
    """
    全局事件过滤器：
    - 点击在“空白区域”(目标不是按钮/列表项等可交互控件)时，清除导入/导出选中态
    """
    def __init__(self, reset_fn, interactive_widgets: tuple[type, ...]):
        super().__init__()
        self.reset_fn = reset_fn
        self.interactive_widgets = interactive_widgets

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
            w = QApplication.widgetAt(event.globalPosition().toPoint())
            if w is None:
                self.reset_fn()
                return False

            for t in self.interactive_widgets:
                if isinstance(w, t) or w.inherits(t.__name__):
                    return False

            self.reset_fn()
        return False


def make_card_vertical(title: str, buttons: list[str]) -> QWidget:
    w = QFrame()
    w.setObjectName("Card")
    w.setFrameShape(QFrame.NoFrame)

    lay = QVBoxLayout(w)
    lay.setContentsMargins(12, 12, 12, 12)
    lay.setSpacing(10)

    t = QLabel(title)
    t.setObjectName("CardTitle")
    lay.addWidget(t)

    for text in buttons:
        b = QPushButton(text)
        b.setObjectName("SoftBtn")
        b.setMinimumHeight(40)
        b.setMinimumWidth(170)
        lay.addWidget(b)

    lay.addStretch(1)
    return w


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  {APP_VERSION}")
        self.setMinimumSize(1180, 760)

        self._build_menu()
        self._build_statusbar()
        self._build_central()
        self._apply_style()
        self._wire()

        self._set_io_active(None)
        self._install_blank_click_reset()
        self._apply_icons()

        self.current_export_mode = None  # 当前导出模式
        self.setWindowIcon(QIcon("生成数据处理软件图片.png"))

    def _install_blank_click_reset(self):
        interactive = (QPushButton, QListWidget, QTextEdit, QProgressBar, QStackedWidget)
        self._blank_filter = ClickResetFilter(
            reset_fn=lambda: self._set_io_active(None),
            interactive_widgets=interactive
        )
        QApplication.instance().installEventFilter(self._blank_filter)

    def _build_menu(self):
        m_file = self.menuBar().addMenu("文件")
        m_help = self.menuBar().addMenu("帮助")
        m_log = self.menuBar().addMenu("日志")
        m_about = self.menuBar().addMenu("关于")

        self.act_import = QAction("数据导入", self)
        self.act_export = QAction("数据导出", self)
        self.act_exit = QAction("退出", self)
        self.act_exit.triggered.connect(self.close)

        m_file.addAction(self.act_import)
        m_file.addAction(self.act_export)
        m_file.addSeparator()
        m_file.addAction(self.act_exit)

        self.act_open_log = QAction("打开日志面板", self)
        m_log.addAction(self.act_open_log)

        self.act_about = QAction("关于本软件", self)
        m_about.addAction(self.act_about)

        self.act_help = QAction("帮助文档", self)
        m_help.addAction(self.act_help)

    def _build_statusbar(self):
        sb = QStatusBar(self)
        sb.setObjectName("StatusBar")
        self.setStatusBar(sb)

        self.lbl_status_left = QLabel("就绪")
        self.lbl_status_left.setObjectName("StatusText")
        sb.addWidget(self.lbl_status_left, 1)

        self.lbl_status_right = QLabel(f"{APP_NAME}  {APP_VERSION}")
        self.lbl_status_right.setObjectName("StatusTextRight")
        sb.addPermanentWidget(self.lbl_status_right)

    def _build_central(self):
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        # ✅ 固定化三列布局：不用 QSplitter（不能拖动）
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(12)  # 只留间距，不画分割线

        # 左：导航（固定宽度）
        self.left_list = QListWidget()
        self.left_list.setObjectName("NavList")
        for t in ["考勤检查", "卫生普查", "院系词典", "详细设置"]:
            QListWidgetItem(t, self.left_list)
        self.left_list.setCurrentRow(0)

        self.left_list.setFixedWidth(150)
        self.left_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.left_list.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.left_list.setTextElideMode(Qt.ElideRight)

        # 中：快捷入口 + 导入导出（固定宽度）
        self.mid_stack = QStackedWidget()
        self.mid_stack.setObjectName("MidStack")

        self.card_att = make_card_vertical("考勤检查 · 快捷入口", ["上课考勤", "晚自习考勤", "考勤汇总"])
        self.card_hyg = make_card_vertical("卫生普查 · 快捷入口", ["卫生分表", "卫生比率"])
        self.card_dict = make_card_vertical("院系词典 · 快捷入口", ["导入词典", "编辑词典", "导出词典", "匹配规则"])
        self.card_set = make_card_vertical("详细设置 · 快捷入口", ["基础设置", "导出设置", "过滤规则", "日志设置"])
        for w in [self.card_att, self.card_hyg, self.card_dict, self.card_set]:
            self.mid_stack.addWidget(w)

        mid_wrap = QWidget()
        mid_wrap.setObjectName("MidWrap")
        mid_wrap.setFixedWidth(290)
        mid_layout = QVBoxLayout(mid_wrap)
        mid_layout.setContentsMargins(0, 0, 0, 0)
        mid_layout.setSpacing(10)

        self.btn_import = QPushButton("数据导入")
        self.btn_import.setObjectName("IOBtn")
        self.btn_import.setCheckable(True)
        self.btn_import.setFixedHeight(42)

        self.btn_export = QPushButton("数据导出")
        self.btn_export.setObjectName("IOBtn")
        self.btn_export.setCheckable(True)
        self.btn_export.setFixedHeight(42)

        mid_layout.addWidget(self.mid_stack, 1)
        mid_layout.addWidget(self.btn_import)
        mid_layout.addWidget(self.btn_export)

        # 右：工作区（自适应伸展）
        right_wrap = QWidget()
        right_wrap.setObjectName("RightWrap")
        rlay = QVBoxLayout(right_wrap)
        rlay.setContentsMargins(0, 0, 0, 0)
        rlay.setSpacing(10)

        self.header_bar = QFrame()
        self.header_bar.setObjectName("HeaderBar")
        hbl = QHBoxLayout(self.header_bar)
        hbl.setContentsMargins(12, 10, 12, 10)
        hbl.setSpacing(10)

        self.header_title = QLabel("工作区 · 考勤检查")
        self.header_title.setObjectName("HeaderTitle")

        self.header_pill = QLabel("就绪")
        self.header_pill.setObjectName("HeaderPill")
        self.header_pill.setAlignment(Qt.AlignCenter)

        hbl.addWidget(self.header_title, 1)
        hbl.addWidget(self.header_pill, 0)

        self.work_area = QTextEdit()
        self.work_area.setObjectName("WorkArea")
        self.work_area.setPlaceholderText("这里是工作区：点击中间的功能按钮后，在此显示详细信息/日志/结果预览…")

        self.progress = NiceProgressBar()

        rlay.addWidget(self.header_bar)
        rlay.addWidget(self.work_area, 1)
        rlay.addWidget(self.progress)

        # 组合三列
        row.addWidget(self.left_list)
        row.addWidget(mid_wrap)
        row.addWidget(right_wrap, 1)

        root.addLayout(row, 1)
        self.setCentralWidget(central)

    def _wire(self):
        self.left_list.currentRowChanged.connect(self._on_nav_changed)

        self.act_import.triggered.connect(lambda: self._on_io_clicked("import"))
        self.act_export.triggered.connect(lambda: self._on_io_clicked("export"))

        self.btn_import.clicked.connect(lambda: self._on_io_clicked("import"))
        self.btn_export.clicked.connect(lambda: self._on_io_clicked("export"))

        for i in range(self.mid_stack.count()):
            card = self.mid_stack.widget(i)
            for b in card.findChildren(QPushButton):
                b.clicked.connect(lambda checked=False, t=b.text(): self._append(f"功能：{t}"))

    def _on_nav_changed(self, idx: int):
        self.mid_stack.setCurrentIndex(idx)
        name_map = {0: "考勤检查", 1: "卫生普查", 2: "院系词典", 3: "详细设置"}
        page = name_map.get(idx, "—")
        self.header_title.setText(f"工作区 · {page}")
        self.statusBar().showMessage(f"切换到：{page}", 1200)

    def _on_io_clicked(self, which: str):
        if which == "import":
            self._set_io_active("import")
            self._append("按钮/菜单：数据导入")
        elif which == "export":
            self._set_io_active("export")
            self._append("按钮/菜单：数据导出")

    def _set_io_active(self, which):
        if which == "import":
            self.btn_import.setChecked(True)
            self.btn_export.setChecked(False)
        elif which == "export":
            self.btn_import.setChecked(False)
            self.btn_export.setChecked(True)
        else:
            self.btn_import.setChecked(False)
            self.btn_export.setChecked(False)

    def _append(self, text: str):
        self.work_area.append(f"【操作】{text}")
        v = self.progress.value() + 12
        self.progress.setValue(0 if v > 100 else v)
        self.header_pill.setText("处理中" if self.progress.value() not in (0, 100) else "就绪")

    def _apply_icons(self):
        self.btn_import.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.btn_export.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.act_import.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.act_export.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))

        icons = [
            QStyle.SP_FileDialogDetailedView,
            QStyle.SP_DialogApplyButton,
            QStyle.SP_FileDialogContentsView,
            QStyle.SP_FileDialogListView
        ]
        for i in range(self.left_list.count()):
            self.left_list.item(i).setIcon(self.style().standardIcon(icons[i]))

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow{ background:#f6f7f9; }

            /* ---------- Menu ---------- */
            QMenuBar{
                background:#ffffff;
                border-bottom:1px solid #e8eaee;
                padding:6px 6px;
            }
            QMenuBar::item{
                padding:6px 10px;
                border-radius:10px;
                color:#1f2328;
            }
            QMenuBar::item:selected{ background:#f2f4f7; }
            QMenu{
                background:#ffffff;
                border:1px solid #e8eaee;
                border-radius:10px;
                padding:6px;
            }
            QMenu::item{
                padding:8px 10px;
                border-radius:8px;
                color:#1f2328;
            }
            QMenu::item:selected{
                background:#eef4ff;
                color:#0b57d0;
            }

            /* ---------- StatusBar ---------- */
            QStatusBar#StatusBar{
                background:#ffffff;
                border-top:1px solid #e8eaee;
                padding:4px 6px;
            }
            QLabel#StatusText{ color:#667085; }
            QLabel#StatusTextRight{ color:#667085; }

            /* ---------- Left Nav ---------- */
            QListWidget#NavList{
                background:#ffffff;
                border:1px solid #e8eaee;
                border-radius:14px;
                padding:6px;
                outline:0;
            }
            QListWidget#NavList::item{
                padding:10px 12px;
                margin:4px 6px;
                border-radius:12px;
                color:#1f2328;
            }
            QListWidget#NavList::item:hover{ background:#f2f4f7; }
            QListWidget#NavList::item:selected{
                background:#eaf2ff;
                color:#0b57d0;
                font-weight:800;
            }

            /* ---------- Cards ---------- */
            QFrame#Card{
                background:#ffffff;
                border:1px solid #e8eaee;
                border-radius:14px;
            }
            QLabel#CardTitle{
                color:#111827;
                font-weight:900;
            }

            QPushButton#SoftBtn{
                background:#f8fafc;
                border:1px solid #e8eaee;
                border-radius:12px;
                padding:10px 12px;
                text-align:left;
                color:#1f2328;
                font-weight:650;
            }
            QPushButton#SoftBtn:hover{ background:#f2f4f7; }
            QPushButton#SoftBtn:pressed{ background:#e9eef6; }
            QPushButton#SoftBtn:focus{ border:1px solid #9ec2ff; }

            QPushButton#IOBtn{
                background:#ffffff;
                color:#1f2328;
                border:1px solid #dfe3ea;
                border-radius:14px;
                padding:10px 14px;
                font-weight:900;
            }
            QPushButton#IOBtn:hover{ background:#f4f6f9; }
            QPushButton#IOBtn:pressed{ background:#eef2f7; }
            QPushButton#IOBtn:checked{
                background:#111827;
                color:#ffffff;
                border:none;
            }

            /* ---------- HeaderBar ---------- */
            QFrame#HeaderBar{
                background:#ffffff;
                border:1px solid #e8eaee;
                border-radius:14px;
            }
            QLabel#HeaderTitle{
                color:#111827;
                font-weight:900;
            }
            QLabel#HeaderPill{
                background:#eef4ff;
                color:#0b57d0;
                border:1px solid #d7e6ff;
                border-radius:12px;
                padding:4px 10px;
                font-weight:800;
                min-width:56px;
            }

            /* ---------- WorkArea ---------- */
            QTextEdit#WorkArea{
                background:#ffffff;
                border:1px solid #e8eaee;
                border-radius:14px;
                padding:10px 12px;
                color:#111827;
                selection-background-color:#cfe1ff;
                selection-color:#0b57d0;
            }
            QTextEdit#WorkArea:focus{ border:1px solid #9ec2ff; }

            QScrollBar:vertical{
                background:transparent;
                width:10px;
                margin:8px 4px 8px 0px;
            }
            QScrollBar::handle:vertical{
                background:#d7dbe3;
                border-radius:5px;
                min-height:30px;
            }
            QScrollBar::handle:vertical:hover{ background:#c7ccdb; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical{ height:0px; }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical{ background:transparent; }

            QScrollBar:horizontal{
                background:transparent;
                height:10px;
                margin:0px 8px 4px 8px;
            }
            QScrollBar::handle:horizontal{
                background:#d7dbe3;
                border-radius:5px;
                min-width:30px;
            }
            QScrollBar::handle:horizontal:hover{ background:#c7ccdb; }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal{ width:0px; }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal{ background:transparent; }

            QToolTip{
                background:#111827;
                color:#ffffff;
                border:none;
                border-radius:8px;
                padding:6px 8px;
            }
        """)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    app.setWindowIcon(QIcon("生成数据处理软件图片.png"))

    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
