# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
from copy import deepcopy
from typing import Dict, Tuple

from PySide6.QtCore import Qt, QSettings
from PySide6.QtGui import QAction, QIcon, QIntValidator
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QListWidget, QListWidgetItem,
    QStackedWidget, QFrame, QStyle, QLineEdit, QMessageBox,
    QTabWidget, QFormLayout, QSpinBox, QCheckBox, QFileDialog,
    QDialog, QDialogButtonBox, QScrollArea, QSizePolicy, QGridLayout
)

from .ui import NiceProgressBar, ClickResetFilter, make_card_vertical, DropLineEdit
from .dialogs import SettingsDialog, DictionaryDialog, ImportDialog, HelpDialog, AboutDialog
from .workers import Worker, WorkerIface2, WorkerBriefing, SettingsRun, BriefingSettings
from .processing import __version__, __build_note__, parse_interval_text, parse_single_text, load_detail_dataframe
from .settings_utils import append_runtime_log, get_bool_setting, set_bool_setting, read_dictionary_setting, save_dictionary_setting

APP_NAME = "宿舍普查数据清洗软件"

class MainWindow(QMainWindow):
    ORG = "DormHealth"
    APP = "LHExporter"

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{__version__}")
        self.setMinimumSize(1180, 760)

        # Icon from PNG file
        self._load_window_icon()

        # State variables for navigation
        self.nav_state = "left"  # left, mid, right

        # Initialize settings and UI
        self._init_settings()

    def _load_window_icon(self):
        """从 PNG 文件加载窗口图标"""
        import os
        possible_paths = [
            "生成数据处理软件图片.png",
            "../生成数据处理软件图片.png",
            os.path.join(os.path.dirname(__file__), "..", "生成数据处理软件图片.png"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "生成数据处理软件图片.png"),
        ]
        for path in possible_paths:
            if os.path.exists(path):
                try:
                    self.setWindowIcon(QIcon(path))
                    return
                except Exception:
                    pass

    def _init_settings(self):
        self.qs = QSettings(self.ORG, self.APP)
        self.dept_dictionary = read_dictionary_setting(self.qs)

        # Data path edits (kept for logic)
        self.input_edit = DropLineEdit(self)
        self.rect_input_edit = DropLineEdit(self)
        self.brief_report_edit = DropLineEdit(self)

        self.worker = None
        self.last_progress_logged = 0

        self._build_menu()
        self._build_central()
        self.load_basic_settings()
        self.apply_theme()

        self._wire()
        self._install_blank_click_reset()
        self._apply_icons()

        # default nav
        self.left_list.setCurrentRow(0)
        self._on_left_nav_changed(0)

    # ---------- UI: Menu ----------
    def _build_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("文件")
        act_import = QAction("数据导入", self)
        act_import.triggered.connect(self.open_import_dialog)
        file_menu.addAction(act_import)

        help_menu = menubar.addMenu("帮助")
        act_help = QAction("使用说明", self)
        act_help.triggered.connect(self.open_help)
        help_menu.addAction(act_help)

        log_menu = menubar.addMenu("日志")
        act_log = QAction("查看运行日志", self)
        act_log.triggered.connect(self.open_logs_dialog)
        log_menu.addAction(act_log)

        about_menu = menubar.addMenu("关于")
        act_about = QAction("关于本软件", self)
        act_about.triggered.connect(self.show_about)
        about_menu.addAction(act_about)

    # ---------- UI: Central ----------
    def _build_central(self):
        central = QWidget(self)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(12)

        # Left nav list
        self.left_list = QListWidget(self)
        self.left_list.setObjectName("NavList")
        for t in ["卫生普查", "考勤检查", "院系词典", "详细设置"]:
            QListWidgetItem(t, self.left_list)
        self.left_list.setFixedWidth(160)
        self.left_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.left_list.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.left_list.setTextElideMode(Qt.ElideRight)

        # Middle stack cards + import/export
        self.mid_stack = QStackedWidget(self)
        self.mid_stack.setObjectName("MidStack")

        self.card_hyg = make_card_vertical("卫生普查 · 快捷入口", ["卫生分表", "卫生比率"])
        self.card_brief = make_card_vertical("考勤检查 · 快捷入口", ["上课考勤", "晚自习考勤", "考勤汇总"])
        self.card_dict = make_card_vertical("院系词典 · 快捷入口", ["打开院系词典"])
        self.card_set = make_card_vertical("详细设置 · 快捷入口", ["数据导入", "运行日志", "设置"])
        for w in [self.card_hyg, self.card_brief, self.card_dict, self.card_set]:
            self.mid_stack.addWidget(w)

        mid_wrap = QWidget(self)
        mid_wrap.setObjectName("MidWrap")
        mid_wrap.setFixedWidth(250)
        mid_layout = QVBoxLayout(mid_wrap)
        mid_layout.setContentsMargins(0, 0, 0, 0)
        mid_layout.setSpacing(10)

        self.btn_import = QPushButton("数据导入", self)
        self.btn_import.setObjectName("IOBtn")
        self.btn_import.setCheckable(True)
        self.btn_import.setFixedHeight(42)

        self.btn_export = QPushButton("数据导出", self)
        self.btn_export.setObjectName("IOBtn")
        self.btn_export.setCheckable(True)
        self.btn_export.setFixedHeight(42)

        mid_layout.addWidget(self.mid_stack, 1)
        mid_layout.addWidget(self.btn_import)
        mid_layout.addWidget(self.btn_export)

        # Right: header + tabs + progress
        right_wrap = QWidget(self)
        right_wrap.setObjectName("RightWrap")
        rlay = QVBoxLayout(right_wrap)
        rlay.setContentsMargins(0, 0, 0, 0)
        rlay.setSpacing(10)

        self.header_bar = QFrame(self)
        self.header_bar.setObjectName("HeaderBar")
        hbl = QHBoxLayout(self.header_bar)
        hbl.setContentsMargins(12, 10, 12, 10)
        hbl.setSpacing(10)

        self.header_title = QLabel("工作区 · 卫生普查", self)
        self.header_title.setObjectName("HeaderTitle")

        self.header_pill = QLabel("就绪", self)
        self.header_pill.setObjectName("HeaderPill")
        self.header_pill.setAlignment(Qt.AlignCenter)

        hbl.addWidget(self.header_title, 1)
        hbl.addWidget(self.header_pill, 0)

        self.tabs = QTabWidget(self)
        self.tabs.setObjectName("WorkTabs")

        self._build_tabs()
        self.tabs.tabBar().hide()  # 隐藏标签栏

        self.progress = NiceProgressBar()

        rlay.addWidget(self.header_bar)
        rlay.addWidget(self.tabs, 1)
        rlay.addWidget(self.progress)

        row.addWidget(self.left_list)
        row.addWidget(mid_wrap)
        row.addWidget(right_wrap, 1)

        root.addLayout(row, 1)
        self.setCentralWidget(central)

        # 保存引用
        self.mid_wrap = mid_wrap
        self.right_wrap = right_wrap

    def _build_tabs(self):
        # Tab1: 分表导出
        tab1 = QWidget(self); self.tabs.addTab(tab1, "分表导出")
        t1_layout = QVBoxLayout(tab1); t1_layout.setContentsMargins(14, 14, 14, 14); t1_layout.setSpacing(12)

        card1 = QFrame(self); card1.setObjectName("Card")
        c1 = QVBoxLayout(card1); c1.setContentsMargins(16, 16, 16, 16); c1.setSpacing(12)
        form1 = QFormLayout(); form1.setLabelAlignment(Qt.AlignRight)

        # 使用标签和按钮替代 QSpinBox
        self.max_sheets_value = 6  # 默认值
        max_sheets_layout = QHBoxLayout()
        
        # 显示值的输入框（左侧）
        self.input_max_sheets = QLineEdit(self)
        self.input_max_sheets.setText(str(self.max_sheets_value))
        self.input_max_sheets.setMinimumWidth(150)
        self.input_max_sheets.setValidator(QIntValidator(1, 50, self))
        self.input_max_sheets.textChanged.connect(self.on_max_sheets_changed)
        
        # 减少按钮
        self.btn_minus = QPushButton("-1", self)
        self.btn_minus.setMinimumWidth(40)
        self.btn_minus.clicked.connect(self.on_minus_clicked)
        
        # 增加按钮
        self.btn_plus = QPushButton("+1", self)
        self.btn_plus.setMinimumWidth(40)
        self.btn_plus.clicked.connect(self.on_plus_clicked)
        
        max_sheets_layout.addWidget(self.input_max_sheets)
        max_sheets_layout.addStretch(1)  # 中间空间
        max_sheets_layout.addWidget(self.btn_minus)
        max_sheets_layout.addWidget(self.btn_plus)
        
        self.chk_mark_num_zero = QCheckBox("标红：数字 0", self)
        self.chk_mark_text_zero = QCheckBox('标红：文本 "0"', self)
        self.chk_mark_text_zero_dot = QCheckBox('标红：文本 "0.0"', self)
        self.chk_use_majority_dept_ui1 = QCheckBox("寝室院系按人数占比判定（平局取首个院系）", self)

        self.output_edit1 = QLineEdit(self); self.output_edit1.setPlaceholderText("保存为 .xlsx …")
        btn_out1 = QPushButton("保存到…", self)
        btn_out1.clicked.connect(self.choose_output1)
        row1 = QHBoxLayout(); row1.addWidget(self.output_edit1); row1.addWidget(btn_out1)

        mark_row = QHBoxLayout(); mark_row.setSpacing(10)
        for cb in (self.chk_mark_num_zero, self.chk_mark_text_zero, self.chk_mark_text_zero_dot):
            mark_row.addWidget(cb)
        mark_row.addStretch(1)

        # 保持原有的 form1 布局方式，直接添加标签和控件组
        form1.addRow("每个工作表最多楼栋数：", self._wrap_layout(max_sheets_layout))
        form1.addRow("标红选项：", mark_row)
        form1.addRow("院系判定：", self.chk_use_majority_dept_ui1)
        form1.addRow("输出文件：", QWidget().setLayout(row1) if False else self._wrap_layout(row1))
        c1.addLayout(form1)

        self.btn_run1 = QPushButton("导出", self); self.btn_run1.setObjectName("PrimaryButton")
        self.btn_run1.setMinimumHeight(44)
        self.btn_run1.clicked.connect(self.start_run1)
        c1.addWidget(self.btn_run1)

        t1_layout.addWidget(card1)

        # Tab2: 楼栋遍历与排除
        tab2 = QWidget(self); self.tabs.addTab(tab2, "楼栋遍历与排除")
        t2_layout = QVBoxLayout(tab2); t2_layout.setContentsMargins(14, 14, 14, 14); t2_layout.setSpacing(12)

        card2 = QFrame(self); card2.setObjectName("Card")
        c2 = QVBoxLayout(card2); c2.setContentsMargins(16, 16, 16, 16); c2.setSpacing(12)
        form2 = QFormLayout(); form2.setLabelAlignment(Qt.AlignRight)

        self.exclude_buildings = list(range(1, 4))
        self.chk_excl_enable_ui2 = QCheckBox("启用“寝室区间排除”")
        self.chk_excl_enable_ui2.stateChanged.connect(lambda _: self._update_exclusion_summary())
        self.exclusion_cfg: Dict[str, Dict[int, Dict[str, str]]] = {"lan": {}, "mei": {}}

        self.btn_config_exclusion = QPushButton("配置区间/单间…", self)
        self.btn_config_exclusion.clicked.connect(self.open_exclusion_dialog)
        self.lbl_exclusion_summary = QLabel("未配置"); self.lbl_exclusion_summary.setObjectName("HeaderPill")

        ex_head_row = QHBoxLayout(); ex_head_row.setSpacing(10)
        ex_head_row.addWidget(self.chk_excl_enable_ui2)
        ex_head_row.addWidget(self.btn_config_exclusion)
        ex_head_row.addWidget(self.lbl_exclusion_summary, 1)

        self.chk_use_majority_dept = QCheckBox("寝室院系按人数占比判定（平局取首个院系）")
        self.chk_drop_zero_text_ui2 = QCheckBox("排除文本 0.0（含 0.00/0.000分 等写法）")
        self.chk_drop_zero_numeric_ui2 = QCheckBox("排除数值 0")
        self.chk_drop_zero_text_ui2.setChecked(True)
        self.chk_drop_zero_numeric_ui2.setChecked(True)

        self.output_edit2 = QLineEdit(self); self.output_edit2.setPlaceholderText("保存为 .xlsx …")
        btn_out2 = QPushButton("保存到…", self)
        btn_out2.clicked.connect(self.choose_output2)
        row2 = QHBoxLayout(); row2.addWidget(self.output_edit2); row2.addWidget(btn_out2)

        opt_row = QHBoxLayout(); opt_row.setSpacing(12)
        for cb in (self.chk_use_majority_dept, self.chk_drop_zero_text_ui2, self.chk_drop_zero_numeric_ui2):
            opt_row.addWidget(cb)
        opt_row.addStretch(1)

        c2.addLayout(ex_head_row)
        form2.addRow("统计选项：", self._wrap_layout(opt_row))
        form2.addRow("输出文件：", self._wrap_layout(row2))
        c2.addLayout(form2)

        self.btn_run2 = QPushButton("导出", self); self.btn_run2.setObjectName("PrimaryButton")
        self.btn_run2.setMinimumHeight(44)
        self.btn_run2.clicked.connect(self.start_run2)
        c2.addWidget(self.btn_run2)

        t2_layout.addWidget(card2)

        # Tab3: 考勤检查（空白）
        tab3 = QWidget(self); self.tabs.addTab(tab3, "考勤检查")
        t3_layout = QVBoxLayout(tab3); t3_layout.setContentsMargins(14, 14, 14, 14); t3_layout.setSpacing(12)
        
        # 添加卡片式大框
        card3 = QFrame(self); card3.setObjectName("Card")
        c3 = QVBoxLayout(card3); c3.setContentsMargins(16, 16, 16, 16); c3.setSpacing(12)
        
        # 功能状态显示
        status_row = QHBoxLayout()
        status_row.setSpacing(10)
        status_label = QLabel("功能状态：")
        status_label.setAlignment(Qt.AlignRight)
        self.lbl_attendance_status = QLabel("开发中")
        self.lbl_attendance_status.setObjectName("HeaderPill")
        self.lbl_attendance_status.setAlignment(Qt.AlignCenter)
        status_row.addWidget(status_label)
        status_row.addWidget(self.lbl_attendance_status)
        status_row.addStretch(1)
        
        # 空白内容
        blank_label = QLabel("考勤检查功能开发中...")
        blank_label.setAlignment(Qt.AlignCenter)
        blank_label.setStyleSheet("font-size: 16px; color: #667085;")
        
        c3.addLayout(status_row)
        c3.addWidget(blank_label, 1)
        t3_layout.addWidget(card3)

    def _wrap_layout(self, layout: QHBoxLayout) -> QWidget:
        w = QWidget(self)
        w.setLayout(layout)
        return w

    def _section_header(self, title: str, subtitle: str) -> QWidget:
        box = QFrame(self)
        v = QVBoxLayout(box); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(4)
        t = QLabel(title); t.setObjectName("SectionTitle")
        s = QLabel(subtitle); s.setObjectName("SectionSub"); s.setWordWrap(True)
        v.addWidget(t)
        v.addWidget(s)
        return box

    # ---------- Wiring ----------
    def _wire(self):
        self.left_list.currentRowChanged.connect(self._on_left_nav_changed)
        self.btn_import.clicked.connect(lambda: self._on_io_clicked("import"))
        self.btn_export.clicked.connect(lambda: self._on_io_clicked("export"))
        self.tabs.currentChanged.connect(lambda _: self._sync_header_title())

        # mid card buttons
        for b in self.card_hyg.findChildren(QPushButton):
            if "分表" in b.text():
                b.clicked.connect(lambda: (self.tabs.setCurrentIndex(0), setattr(self, "nav_state", "right")))
            else:
                b.clicked.connect(lambda: (self.tabs.setCurrentIndex(1), setattr(self, "nav_state", "right")))
        for b in self.card_brief.findChildren(QPushButton):
            if "上课考勤" in b.text():
                b.clicked.connect(lambda: (self.tabs.setCurrentIndex(2), setattr(self, "nav_state", "right")))
            elif "晚自习考勤" in b.text():
                b.clicked.connect(lambda: (self.tabs.setCurrentIndex(2), setattr(self, "nav_state", "right")))
            elif "考勤汇总" in b.text():
                b.clicked.connect(lambda: (self.tabs.setCurrentIndex(2), setattr(self, "nav_state", "right")))
        for b in self.card_dict.findChildren(QPushButton):
            b.clicked.connect(lambda: (self.open_dictionary_dialog(), setattr(self, "nav_state", "right")))
        for b in self.card_set.findChildren(QPushButton):
            if "数据导入" in b.text():
                b.clicked.connect(lambda: (self.open_import_dialog(), setattr(self, "nav_state", "right")))
            elif "运行日志" in b.text():
                b.clicked.connect(lambda: (self.open_logs_dialog(), setattr(self, "nav_state", "right")))
            else:
                b.clicked.connect(lambda: (self.open_settings(), setattr(self, "nav_state", "right")))

    def _install_blank_click_reset(self):
        interactive = (QPushButton, QListWidget, QTextEdit, QStackedWidget, QLineEdit, QCheckBox, QTabWidget, QSpinBox)
        self._blank_filter = ClickResetFilter(reset_fn=lambda: self._set_io_active(None), interactive_widgets=interactive)
        QApplication.instance().installEventFilter(self._blank_filter)

    def _apply_icons(self):
        self.btn_import.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.btn_export.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        icons = [QStyle.SP_DialogApplyButton, QStyle.SP_FileDialogInfoView, QStyle.SP_FileDialogDetailedView, QStyle.SP_FileDialogContentsView]
        for i in range(self.left_list.count()):
            self.left_list.item(i).setIcon(self.style().standardIcon(icons[i]))

    # ---------- Theme/QSS ----------
    def _qss(self, dark: bool) -> str:
        if not dark:
            bg = "#f6f7f9"; card="#ffffff"; border="#e8eaee"
            text="#1f2328"; muted="#667085"; accent="#0b57d0"
            hover="#f2f4f7"
        else:
            bg="#0b0b0c"; card="#141416"; border="#2a2a2e"
            text="#f2f2f7"; muted="#a1a1aa"; accent="#3b82f6"
            hover="#1f2937"

        return f"""
        QMainWindow{{ background:{bg}; color:{text}; }}
        QMenuBar{{ background:{card}; border-bottom:1px solid {border}; padding:6px 6px; }}
        QMenuBar::item{{ padding:6px 10px; border-radius:10px; color:{text}; }}
        QMenuBar::item:selected{{ background:{hover}; }}
        QMenu{{ background:{card}; border:1px solid {border}; border-radius:10px; padding:6px; }}
        QMenu::item{{ padding:8px 10px; border-radius:8px; color:{text}; }}
        QMenu::item:selected{{ background:{('#1f2937' if dark else '#eef4ff')}; color:{accent}; }}

        QListWidget#NavList{{
            background:{card};
            border:1px solid {border};
            border-radius:14px;
            padding:6px;
            outline:0;
        }}
        QListWidget#NavList::item{{
            padding:10px 12px;
            margin:4px 6px;
            border-radius:12px;
            color:{text};
        }}
        QListWidget#NavList::item:hover{{ background:{hover}; }}
        QListWidget#NavList::item:selected{{
            background:{('#1f2937' if dark else '#eaf2ff')};
            color:{accent};
            font-weight:800;
        }}

        QFrame#Card{{ background:{card}; border:1px solid {border}; border-radius:14px; }}
        QLabel#CardTitle{{ color:{text}; font-weight:900; }}

        QPushButton#SoftBtn{{
            background:{('#1f2937' if dark else '#f8fafc')};
            border:1px solid {border};
            border-radius:12px;
            padding:10px 12px;
            text-align:left;
            color:{text};
            font-weight:650;
        }}
        QPushButton#SoftBtn:hover{{ background:{hover}; }}

        QPushButton#IOBtn{{
            background:{card};
            color:{text};
            border:1px solid {border};
            border-radius:14px;
            padding:10px 14px;
            font-weight:900;
        }}
        QPushButton#IOBtn:hover{{ background:{hover}; }}
        QPushButton#IOBtn:checked{{
            background:{text};
            color:{card};
            border:none;
        }}

        QFrame#HeaderBar{{ background:{card}; border:1px solid {border}; border-radius:14px; }}
        QLabel#HeaderTitle{{ color:{text}; font-weight:900; }}
        QLabel#HeaderPill{{
            background:{('#1f2937' if dark else '#eef4ff')};
            color:{accent};
            border:1px solid {('#374151' if dark else '#d7e6ff')};
            border-radius:12px;
            padding:4px 10px;
            font-weight:800;
            min-width:56px;
        }}

        QLabel#SectionTitle{{ font-weight:900; }}
        QLabel#SectionSub{{ color:{muted}; }}

        QLineEdit {{
            background:{card};
            border:1px solid {border};
            border-radius:12px;
            padding:6px 10px;
            height:34px;
            color:{text};
        }}
        QLineEdit:focus {{ border:1px solid {accent}; }}
        
        QSpinBox {{
            background:{card};
            border:1px solid {border};
            border-radius:12px;
            padding:6px 30px 6px 10px; /* 增加右侧内边距，为箭头按钮留出空间 */
            height:34px;
            color:{text};
            min-width: 80px; /* 增加最小宽度 */
        }}
        QSpinBox:focus {{ border:1px solid {accent}; }}
        QSpinBox::up-button, QSpinBox::down-button {{
            subcontrol-origin: border;
            subcontrol-position: right;
            width: 24px; /* 增加按钮宽度 */
            height: 17px; /* 设置按钮高度 */
            border-left: 1px solid {border};
        }}
        QSpinBox::up-button {{ top: 0; }}
        QSpinBox::down-button {{ top: 17px; }}
        QSpinBox::up-arrow, QSpinBox::down-arrow {{
            width: 8px;
            height: 8px;
            background: {text};
        }}

        QPushButton {{
            height:36px; padding:6px 18px; border-radius:12px;
            background:{card}; border:1px solid {border}; color:{text};
        }}
        QPushButton:hover {{ background:{hover}; }}
        QPushButton#PrimaryButton{{ background:{accent}; color:white; border:none; font-weight:700; }}
        QPushButton#PrimaryButton:hover{{ background:{('#2563eb' if dark else '#094bbb')}; }}

        QTabWidget#WorkTabs::pane{{ border:0px; }}

        QToolTip{{ background:#111827; color:#ffffff; border:none; border-radius:8px; padding:6px 8px; }}
        """

    def apply_theme(self):
        theme = str(self.qs.value("theme", "浅色")).strip() or "浅色"
        dark = (theme == "深色")
        self.setStyleSheet(self._qss(dark))

    # ---------- Nav / Header ----------
    def _sync_header_title(self):
        idx = self.tabs.currentIndex()
        if idx == 0:
            self.header_title.setText("工作区 · 卫生分表")
        elif self.tabs.currentIndex() == 1:
            self.header_title.setText("工作区 · 卫生比率")
        else:
            self.header_title.setText("工作区 · 考勤检查")

    def _on_left_nav_changed(self, idx: int):
        # 切换中间卡片内容
        self.mid_stack.setCurrentIndex(idx)
        
        # 切换右侧标签页内容
        if idx == 0:
            self.tabs.setCurrentIndex(0)
        elif idx == 1:
            self.tabs.setCurrentIndex(2)
        
        # 更新标题
        self._sync_header_title()
        
        # 确保中间和右侧内容都可见
        self.mid_wrap.setVisible(True)
        self.right_wrap.setVisible(True)

    def _set_io_active(self, which: str | None):
        if which == "import":
            self.btn_import.setChecked(True); self.btn_export.setChecked(False)
        elif which == "export":
            self.btn_import.setChecked(False); self.btn_export.setChecked(True)
        else:
            self.btn_import.setChecked(False); self.btn_export.setChecked(False)

    def _on_io_clicked(self, which: str):
        if which == "import":
            self._set_io_active("import")
            self.open_import_dialog()
        else:
            self._set_io_active("export")
            self._trigger_export_for_current_tab()

    def _trigger_export_for_current_tab(self):
        idx = self.tabs.currentIndex()
        if idx == 0:
            self.start_run1()
        elif idx == 1:
            self.start_run2()
        else:
            self.start_run3()

    # ---------- Import/Dialogs ----------
    def open_import_dialog(self):
        dlg = ImportDialog(self, input_edit=self.input_edit, rectify_edit=self.rect_input_edit, brief_edit=self.brief_report_edit)
        if dlg.exec() == QDialog.Accepted:
            self.save_basic_settings()
            # auto suggest output names if empty
            in_path = self.input_edit.text().strip()
            if in_path and os.path.isfile(in_path):
                base = os.path.splitext(os.path.basename(in_path))[0]
                out_dir = self._default_output_dir(in_path)
                if not self.output_edit1.text().strip():
                    self.output_edit1.setText(os.path.join(out_dir, f"{base}_表一.xlsx"))
                if not self.output_edit2.text().strip():
                    self.output_edit2.setText(os.path.join(out_dir, f"{base}_表二.xlsx"))
            # 考勤数据相关逻辑
            attendance_path = self.rect_input_edit.text().strip()
            if attendance_path and os.path.isfile(attendance_path) and hasattr(self, 'attendance_output_edit'):
                base = os.path.splitext(os.path.basename(attendance_path))[0]
                out_dir = self._default_output_dir(attendance_path)
                self.attendance_output_edit.setText(os.path.join(out_dir, f"{base}_考勤结果.xlsx"))
            self.save_basic_settings()

    def open_help(self):
        HelpDialog(self).exec()

    def open_settings(self, show_logs: bool = False):
        dlg = SettingsDialog(self, self.qs, show_logs=show_logs)
        if dlg.exec() == QDialog.Accepted:
            self.apply_theme()
            QMessageBox.information(self, "提示", "设置已保存并应用。")

    def open_logs_dialog(self):
        self.open_settings(show_logs=True)

    def open_dictionary_dialog(self):
        dlg = DictionaryDialog(self, self.dept_dictionary)
        if dlg.exec() == QDialog.Accepted:
            self.dept_dictionary = dlg.result_dict
            save_dictionary_setting(self.qs, self.dept_dictionary)
            QMessageBox.information(self, "提示", "院系词典已更新。")

    def show_about(self):
        dlg = AboutDialog(self)
        dlg.exec()

    # ---------- Settings persistence ----------
    def load_basic_settings(self):
        self.input_edit.setText(str(self.qs.value("input_path", "")))
        self.rect_input_edit.setText(str(self.qs.value("rectify/path", "")))
        self.brief_report_edit.setText(str(self.qs.value("brief/report", "")))

        self.output_edit1.setText(str(self.qs.value("ui1/out", "")))
        self.output_edit2.setText(str(self.qs.value("ui2/out", "")))
        # 考勤检查功能开发中，暂时不需要加载输出路径

        self.max_sheets_value = int(self.qs.value("ui1/max", 6))
        self.input_max_sheets.setText(str(self.max_sheets_value))
        self.chk_mark_num_zero.setChecked(get_bool_setting(self.qs, "ui1/mark_num", True))
        self.chk_mark_text_zero.setChecked(get_bool_setting(self.qs, "ui1/mark_txt0", False))
        self.chk_mark_text_zero_dot.setChecked(get_bool_setting(self.qs, "ui1/mark_txt0dot", False))
        self.chk_use_majority_dept_ui1.setChecked(get_bool_setting(self.qs, "ui1/majority_dept", False))

        self.chk_excl_enable_ui2.setChecked(get_bool_setting(self.qs, "ui2/ex/enabled", False))
        self._load_exclusion_cfg()
        self._update_exclusion_summary()
        self.chk_use_majority_dept.setChecked(get_bool_setting(self.qs, "ui2/majority_dept", False))
        legacy = get_bool_setting(self.qs, "ui2/drop_zero", True)
        self.chk_drop_zero_text_ui2.setChecked(get_bool_setting(self.qs, "ui2/drop_zero_text", legacy))
        self.chk_drop_zero_numeric_ui2.setChecked(get_bool_setting(self.qs, "ui2/drop_zero_numeric", legacy))

        # 考勤检查功能开发中，暂时不需要加载相关设置

    def save_basic_settings(self):
        self.qs.setValue("input_path", self.input_edit.text().strip())
        self.qs.setValue("rectify/path", self.rect_input_edit.text().strip())
        self.qs.setValue("brief/report", self.brief_report_edit.text().strip())

        self.qs.setValue("ui1/out", self.output_edit1.text().strip())
        self.qs.setValue("ui2/out", self.output_edit2.text().strip())
        # 考勤检查功能开发中，暂时不需要保存输出路径

        self.qs.setValue("ui1/max", self.max_sheets_value)
        set_bool_setting(self.qs, "ui1/mark_num", self.chk_mark_num_zero.isChecked())
        set_bool_setting(self.qs, "ui1/mark_txt0", self.chk_mark_text_zero.isChecked())
        set_bool_setting(self.qs, "ui1/mark_txt0dot", self.chk_mark_text_zero_dot.isChecked())
        set_bool_setting(self.qs, "ui1/majority_dept", self.chk_use_majority_dept_ui1.isChecked())

        set_bool_setting(self.qs, "ui2/ex/enabled", self.chk_excl_enable_ui2.isChecked())
        self._save_exclusion_cfg()
        set_bool_setting(self.qs, "ui2/majority_dept", self.chk_use_majority_dept.isChecked())
        set_bool_setting(self.qs, "ui2/drop_zero_text", self.chk_drop_zero_text_ui2.isChecked())
        set_bool_setting(self.qs, "ui2/drop_zero_numeric", self.chk_drop_zero_numeric_ui2.isChecked())

        # 考勤检查功能开发中，暂时不需要保存相关设置

    def _default_output_dir(self, input_path: str) -> str:
        d = str(self.qs.value("default_dir", "")).strip()
        if d and os.path.isdir(d):
            return d
        return os.path.dirname(input_path) if input_path else os.getcwd()

    def _excel_font_name(self) -> str:
        return str(self.qs.value("font_name", "仿宋_GB2312")).strip() or "仿宋_GB2312"

    # ---------- Exclusion config (UI2) ----------
    def _load_exclusion_cfg(self):
        cfg = {"lan": {}, "mei": {}}
        for garden, prefix in (("lan", "ui2/ex/lan"), ("mei", "ui2/ex/mei")):
            for no in self.exclude_buildings:
                rng = str(self.qs.value(f"{prefix}/{no}/range", "") or "").strip()
                sgl = str(self.qs.value(f"{prefix}/{no}/single", "") or "").strip()
                if rng or sgl:
                    cfg[garden][no] = {"range": rng, "single": sgl}
        self.exclusion_cfg = cfg

    def _save_exclusion_cfg(self):
        for garden, prefix in (("lan", "ui2/ex/lan"), ("mei", "ui2/ex/mei")):
            garden_cfg = self.exclusion_cfg.get(garden, {})
            for no in self.exclude_buildings:
                vals = garden_cfg.get(no, {})
                self.qs.setValue(f"{prefix}/{no}/range", vals.get("range", ""))
                self.qs.setValue(f"{prefix}/{no}/single", vals.get("single", ""))

    def _update_exclusion_summary(self):
        if not self.chk_excl_enable_ui2.isChecked():
            self.lbl_exclusion_summary.setText("未启用")
            return
        lan_cnt = len(self.exclusion_cfg.get("lan", {}))
        mei_cnt = len(self.exclusion_cfg.get("mei", {}))
        if lan_cnt == 0 and mei_cnt == 0:
            self.lbl_exclusion_summary.setText("未配置")
        else:
            self.lbl_exclusion_summary.setText(f"兰苑{lan_cnt}栋，梅苑{mei_cnt}栋")

    def open_exclusion_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("区间/单间排除设置")
        dlg.setMinimumSize(840, 520)
        lay = QVBoxLayout(dlg); lay.setContentsMargins(16, 16, 16, 16); lay.setSpacing(12)
        tip = QLabel("按楼栋填写区间与单间，留空表示不排除；兰/梅苑均提供 1-3 栋输入。")
        tip.setWordWrap(True)
        lay.addWidget(tip)

        scroll = QScrollArea(dlg); scroll.setWidgetResizable(True)
        wrap = QWidget()
        grid_wrap = QGridLayout(wrap)
        grid_wrap.setContentsMargins(0, 0, 0, 0)
        grid_wrap.setHorizontalSpacing(16)
        grid_wrap.setVerticalSpacing(12)

        edit_refs: Dict[Tuple[str, int], Tuple[QLineEdit, QLineEdit]] = {}

        def build_panel(title: str, garden_key: str, col: int):
            card = QFrame(); card.setObjectName("Card")
            card_lay = QVBoxLayout(card); card_lay.setContentsMargins(10, 10, 10, 10); card_lay.setSpacing(8)
            caption = QLabel(title); caption.setObjectName("SectionTitle")
            card_lay.addWidget(caption)

            grid = QGridLayout(); grid.setHorizontalSpacing(8); grid.setVerticalSpacing(6)
            grid.setColumnStretch(1, 1); grid.setColumnStretch(2, 1)
            for idx, no in enumerate(self.exclude_buildings):
                grid.addWidget(QLabel(f"{no}栋："), idx, 0)
                rng = QLineEdit(); rng.setPlaceholderText("区间：101-120,201-220")
                sgl = QLineEdit(); sgl.setPlaceholderText("单间：101,203,305")
                vals = self.exclusion_cfg.get(garden_key, {}).get(no, {})
                rng.setText(vals.get("range", ""))
                sgl.setText(vals.get("single", ""))
                edit_refs[(garden_key, no)] = (rng, sgl)
                grid.addWidget(rng, idx, 1)
                grid.addWidget(sgl, idx, 2)
            card_lay.addLayout(grid)
            grid_wrap.addWidget(card, 0, col)

        build_panel("兰苑", "lan", 0)
        build_panel("梅苑", "mei", 1)
        scroll.setWidget(wrap)
        lay.addWidget(scroll, 1)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dlg)
        lay.addWidget(btns)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)

        if dlg.exec() == QDialog.Accepted:
            for (garden, no), (rng, sgl) in edit_refs.items():
                rng_text = rng.text().strip()
                sgl_text = sgl.text().strip()
                if rng_text or sgl_text:
                    self.exclusion_cfg.setdefault(garden, {})[no] = {"range": rng_text, "single": sgl_text}
                else:
                    self.exclusion_cfg.setdefault(garden, {}).pop(no, None)
            self._save_exclusion_cfg()
            self._update_exclusion_summary()
            self.save_basic_settings()

    def _build_ex_params(self) -> Dict:
        enabled = self.chk_excl_enable_ui2.isChecked()
        def collect(garden_key: str):
            cfg: Dict[int, Dict[str, object]] = {}
            for no in self.exclude_buildings:
                vals = self.exclusion_cfg.get(garden_key, {}).get(no, {})
                rng_text = str(vals.get("range", "") or "").strip()
                sgl_text = str(vals.get("single", "") or "").strip()
                ranges = parse_interval_text(rng_text)
                singles = parse_single_text(sgl_text)
                if ranges or singles:
                    cfg[no] = {"ranges": ranges, "singles": singles}
            return cfg
        return {"enabled": enabled, "lan": collect("lan"), "mei": collect("mei")}

    # ---------- Export (Threads) ----------
    def on_progress(self, value: int):
        self.progress.setValue(value)
        if value <= 0:
            self.header_pill.setText("就绪")
        elif value < 100:
            self.header_pill.setText("处理中")
        else:
            self.header_pill.setText("完成")

        if value >= 100 or value - self.last_progress_logged >= 20:
            self._log_runtime(f"当前进度：{value}%")
            self.last_progress_logged = value

    def _log_runtime(self, message: str):
        append_runtime_log(self.qs, message)

    def _open_folder(self, path: str):
        try:
            folder = os.path.dirname(path)
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception as e:
            self._log_runtime(f"打开文件夹失败: {e}")
            QMessageBox.warning(self, "错误", f"无法打开文件夹：\n{e}")

    def choose_output1(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存 分表导出 Excel", self.output_edit1.text().strip(), "Excel 文件 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_edit1.setText(path)
            self.save_basic_settings()

    def choose_output2(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存 楼栋遍历与排除 Excel", self.output_edit2.text().strip(), "Excel 文件 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_edit2.setText(path)
            self.save_basic_settings()

    def choose_brief_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存删行简报", self.brief_output_edit.text().strip(), "Excel 文件 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.brief_output_edit.setText(path)
            self.save_basic_settings()

    def _disable_run_buttons(self, disabled: bool):
        self.btn_run1.setEnabled(not disabled)
        self.btn_run2.setEnabled(not disabled)
        self.btn_export.setEnabled(not disabled)

    def on_minus_clicked(self):
        if self.max_sheets_value > 1:
            self.max_sheets_value -= 1
            self.input_max_sheets.setText(str(self.max_sheets_value))

    def on_plus_clicked(self):
        if self.max_sheets_value < 50:
            self.max_sheets_value += 1
            self.input_max_sheets.setText(str(self.max_sheets_value))

    def on_max_sheets_changed(self, text):
        try:
            value = int(text)
            if 1 <= value <= 50:
                self.max_sheets_value = value
        except ValueError:
            pass

    def start_run1(self):
        in_path = self.input_edit.text().strip()
        out_path = self.output_edit1.text().strip()
        if not in_path or not os.path.isfile(in_path):
            QMessageBox.warning(self, "提示", "请先“数据导入”选择有效的原始 Excel（.xlsx/.xls）。")
            self.open_import_dialog()
            return
        if not out_path:
            QMessageBox.warning(self, "提示", "请指定 分表导出 输出文件路径（.xlsx）。")
            return

        s = SettingsRun(
            input_path=in_path,
            output_path=out_path,
            max_sheets=int(self.max_sheets_value),
            font_name=self._excel_font_name(),
            mark_num_zero=bool(self.chk_mark_num_zero.isChecked()),
            mark_text_zero=bool(self.chk_mark_text_zero.isChecked()),
            mark_text_zero_dot=bool(self.chk_mark_text_zero_dot.isChecked()),
            dept_dictionary=deepcopy(self.dept_dictionary),
            use_majority_dept=bool(self.chk_use_majority_dept_ui1.isChecked()),
        )

        self._disable_run_buttons(True)
        self.progress.setValue(5)
        self.last_progress_logged = 0
        self.header_pill.setText("处理中")
        self._log_runtime(f"开始生成分表导出：{os.path.basename(out_path)}")

        self.worker = Worker(s)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def start_run2(self):
        in_path = self.input_edit.text().strip()
        out_path = self.output_edit2.text().strip()
        if not in_path or not os.path.isfile(in_path):
            QMessageBox.warning(self, "提示", "请先“数据导入”选择有效的原始 Excel（.xlsx/.xls）。")
            self.open_import_dialog()
            return
        if not out_path:
            QMessageBox.warning(self, "提示", "请指定 楼栋遍历与排除 输出文件路径（.xlsx）。")
            return

        self._disable_run_buttons(True)
        self.progress.setValue(5)
        self.last_progress_logged = 0
        self.header_pill.setText("处理中")
        self._log_runtime(f"开始生成楼栋遍历与排除：{os.path.basename(out_path)}")

        dept_dict = deepcopy(self.dept_dictionary)
        fallback_df = None
        try:
            fallback_df, _ = load_detail_dataframe(in_path, dept_dict)
        except Exception:
            fallback_df = None

        self.worker = WorkerIface2(
            in_path,
            out_path,
            self._build_ex_params(),
            bool(self.chk_drop_zero_text_ui2.isChecked()),
            bool(self.chk_drop_zero_numeric_ui2.isChecked()),
            bool(self.chk_use_majority_dept.isChecked()),
            dept_dict,
            fallback_df,
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def start_run3(self):
        # 考勤检查功能开发中
        QMessageBox.information(self, "提示", "考勤检查功能正在开发中，敬请期待！")

    def on_finished(self, out_path: str, summary: dict):
        self._disable_run_buttons(False)
        self.progress.setValue(100)
        self.header_pill.setText("完成")
        self.save_basic_settings()

        summary_msg = None
        if isinstance(self.worker, Worker):
            meta = summary or {}
            summary_msg = (
                f"分表导出 完成：{os.path.basename(out_path)}；明细表={meta.get('sheet','')}，有效行={meta.get('valid_rows',0)}，"
                f"楼栋表={meta.get('sheet_count',0)}，不及格条数={meta.get('fail_rows',0)}，0.0条数={meta.get('zero_rows',0)}。"
            )
        elif isinstance(self.worker, WorkerIface2):
            stat = summary or {}
            summary_msg = (
                f"楼栋遍历与排除 完成：{os.path.basename(out_path)}；原始行={stat.get('rows_raw',0)}，结构化={stat.get('rows_after_structure',0)}，"
                f"去零后={stat.get('rows_after_zero',0)}，排除后={stat.get('rows_after_exclusion',0)}，最终计数={stat.get('rows_final',0)}。"
            )
        elif isinstance(self.worker, WorkerBriefing):
            stat = summary or {}
            summary_msg = (
                f"考勤检查 完成：{os.path.basename(out_path)}；数据行={stat.get('report_rows',0)}，处理行={stat.get('rect_rows',0)}，"
                f"结果行={stat.get('remaining_rows',0)}。"
            )

        if summary_msg:
            self._log_runtime(summary_msg)
        else:
            self._log_runtime(f"已生成：{out_path}")

        if get_bool_setting(self.qs, "open_after", True):
            self._open_folder(out_path)

        QMessageBox.information(self, "完成", f"已生成：\n{out_path}\n版本：{__version__}")

    def on_error(self, msg: str):
        self._disable_run_buttons(False)
        self.progress.setValue(0)
        self.header_pill.setText("失败")
        self._log_runtime(f"处理失败：{msg}")
        QMessageBox.critical(self, "错误", f"处理失败：\n{msg}")
