# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from copy import deepcopy
from typing import Dict, List, Optional, Tuple

from PySide6.QtCore import Qt, QSettings
from PySide6.QtGui import QFontDatabase
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QMessageBox, QTabWidget, QFormLayout, QCheckBox, QDialogButtonBox,
    QPlainTextEdit, QScrollArea, QWidget, QFrame, QMenu, QTextEdit, QComboBox
)

from .processing import __version__, __build_note__, __history__, PRESET_DEPTS_IF2, normalize_dictionary, default_department_dictionary
from .settings_utils import get_bool_setting, set_bool_setting, read_dictionary_setting, save_dictionary_setting
from .ui import DropLineEdit

class ImportDialog(QDialog):
    """Data import dialog: raw excel, rectify list, briefing draft."""
    def __init__(self, parent=None, *, input_edit: DropLineEdit, rectify_edit: DropLineEdit, brief_edit: DropLineEdit):
        super().__init__(parent)
        self.setWindowTitle("数据导入")
        self.setMinimumSize(720, 260)
        self.input_edit = input_edit
        self.rectify_edit = rectify_edit
        self.brief_edit = brief_edit

        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(10)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)

        def row(edit: DropLineEdit, title: str, flt: str):
            btn = QPushButton("浏览…")
            def pick():
                path, _ = QFileDialog.getOpenFileName(self, title, edit.text().strip(), flt)
                if path:
                    edit.setText(path)
            btn.clicked.connect(pick)
            h = QHBoxLayout()
            h.addWidget(edit, 1)
            h.addWidget(btn)
            w = QWidget()
            w.setLayout(h)
            return w

        self.input_edit.setPlaceholderText("原始Excel（分表导出/楼栋遍历与排除 共用，.xlsx / .xls）")
        self.rectify_edit.setPlaceholderText("考勤数据（Excel / Word / PDF）")
        # 简报初稿不再使用，保持为空白

        form.addRow("原始数据：", row(self.input_edit, "选择原始 Excel", "Excel 文件 (*.xlsx *.xls)"))
        form.addRow("考勤数据：", row(self.rectify_edit, "选择考勤数据", "支持文件 (*.xlsx *.xls *.doc *.docx *.pdf)"))
        lay.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

class SettingsDialog(QDialog):
    """通用 + 版本 + 运行日志"""
    def __init__(self, parent=None, qsettings: Optional[QSettings] = None, show_logs: bool = False):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setMinimumWidth(760)
        self.qs = qsettings or QSettings("DormHealth", "LHExporter")
        self._show_logs = show_logs

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        tabs = QTabWidget(self)
        self.tabs = tabs

        # 通用
        pg_general = QWidget(); tabs.addTab(pg_general, "通用")
        g1 = QFormLayout(pg_general)
        self.cmb_theme = QComboBox()
        self.cmb_theme.addItems(["浅色", "深色"])

        self.chk_open_after = QCheckBox("导出完成后自动打开文件夹")
        self.edt_default_dir = QLineEdit(); self.edt_default_dir.setPlaceholderText("默认输出目录（留空=跟随输入文件所在目录）")
        btn_pick_dir = QPushButton("选择…")
        def pick_dir():
            d = QFileDialog.getExistingDirectory(self, "选择默认输出目录", "")
            if d:
                self.edt_default_dir.setText(d)
        btn_pick_dir.clicked.connect(pick_dir)
        row = QHBoxLayout(); row.addWidget(self.edt_default_dir, 1); row.addWidget(btn_pick_dir)
        w_row = QWidget(); w_row.setLayout(row)

        self.font_combo = QLineEdit()
        self.font_combo.setPlaceholderText("Excel 字体名（例如 仿宋_GB2312 / 宋体 / 微软雅黑）")
        try:
            fams = list(QFontDatabase.families())
        except Exception:
            fams = []
        self.font_combo.setToolTip("可用字体示例：" + "、".join(fams[:12]) + ("…" if len(fams) > 12 else ""))

        g1.addRow("主题：", self.cmb_theme)
        g1.addRow("", self.chk_open_after)
        g1.addRow("默认输出目录：", w_row)
        g1.addRow("Excel 字体：", self.font_combo)

        # 版本
        pg_version = QWidget(); tabs.addTab(pg_version, "版本信息")
        v1 = QVBoxLayout(pg_version)
        lbl_version = QLabel(f"版本：{__version__}\n\n更新说明：\n{__build_note__}\n\n历史版本：\n{__history__}")
        lbl_version.setWordWrap(True)
        v1.addWidget(lbl_version)

        # 日志
        pg_logs = QWidget(); tabs.addTab(pg_logs, "运行日志")
        v2 = QVBoxLayout(pg_logs); v2.setSpacing(10)
        hint_logs = QLabel("记录错误与工作进程，可复制/导出，便于追踪处理过程。")
        hint_logs.setStyleSheet("color:#6B7280;"); hint_logs.setWordWrap(True)
        self.txt_logs = QTextEdit(); self.txt_logs.setReadOnly(True)
        self.btn_copy_logs = QPushButton("复制全部")
        self.btn_export_logs = QPushButton("导出日志…")
        self.btn_clear_logs = QPushButton("清空日志")
        btn_row_logs = QHBoxLayout(); btn_row_logs.setSpacing(8)
        btn_row_logs.addWidget(self.btn_copy_logs)
        btn_row_logs.addWidget(self.btn_export_logs)
        btn_row_logs.addStretch(1)
        btn_row_logs.addWidget(self.btn_clear_logs)
        v2.addWidget(hint_logs)
        v2.addWidget(self.txt_logs, 1)
        v2.addLayout(btn_row_logs)
        self.btn_clear_logs.clicked.connect(self.clear_logs)
        self.btn_copy_logs.clicked.connect(self.copy_logs)
        self.btn_export_logs.clicked.connect(self.export_logs)

        layout.addWidget(tabs)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.RestoreDefaults)
        layout.addWidget(btns)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        btns.button(QDialogButtonBox.RestoreDefaults).clicked.connect(self.restore_defaults)

        self.load()
        if self._show_logs:
            tabs.setCurrentWidget(pg_logs)

    @staticmethod
    def defaults():
        return {
            "theme": "浅色",
            "open_after": True,
            "default_dir": "",
            "font": "仿宋_GB2312",
        }

    def load(self):
        d = self.defaults()
        theme = str(self.qs.value("theme", d["theme"]))
        index = self.cmb_theme.findText(theme)
        if index >= 0:
            self.cmb_theme.setCurrentIndex(index)
        else:
            self.cmb_theme.setCurrentIndex(0)
        self.chk_open_after.setChecked(get_bool_setting(self.qs, "open_after", d["open_after"]))
        self.edt_default_dir.setText(str(self.qs.value("default_dir", d["default_dir"])))
        self.font_combo.setText(str(self.qs.value("font_name", d["font"])))
        logs = str(self.qs.value("runtime/logs", ""))
        self.txt_logs.setPlainText(logs)

    def restore_defaults(self):
        d = self.defaults()
        index = self.cmb_theme.findText(d["theme"])
        if index >= 0:
            self.cmb_theme.setCurrentIndex(index)
        self.chk_open_after.setChecked(d["open_after"])
        self.edt_default_dir.setText(d["default_dir"])
        self.font_combo.setText(d["font"])

    def accept(self):
        theme = self.cmb_theme.currentText()
        self.qs.setValue("theme", theme)
        set_bool_setting(self.qs, "open_after", self.chk_open_after.isChecked())
        self.qs.setValue("default_dir", self.edt_default_dir.text().strip())
        self.qs.setValue("font_name", self.font_combo.text().strip())
        super().accept()

    def clear_logs(self):
        self.qs.setValue("runtime/logs", "")
        self.txt_logs.setPlainText("")

    def copy_logs(self):
        QApplication.clipboard().setText(self.txt_logs.toPlainText())

    def export_logs(self):
        logs = self.txt_logs.toPlainText()
        if not logs.strip():
            QMessageBox.information(self, "提示", "暂无可导出的日志内容。")
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出运行日志", "运行日志.txt", "文本文件 (*.txt)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(logs)
            QMessageBox.information(self, "完成", f"日志已导出到：\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{e}")

class DictionaryDialog(QDialog):
    def __init__(self, parent=None, data: Optional[Dict[str, Dict[str, List[str]]]] = None):
        super().__init__(parent)
        self.setWindowTitle("院系词典中心")
        self.setMinimumSize(720, 520)
        self._data = normalize_dictionary(data)
        self.result_dict: Dict[str, Dict[str, List[str]]] = deepcopy(self._data)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        tip = QLabel("维护“班级 → 院系”的映射。左侧切换院系，右侧编辑精确班级与关键字。")
        tip.setWordWrap(True)
        layout.addWidget(tip)

        body = QFrame(self)
        body_layout = QHBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(12)

        # left: dept buttons
        left_col = QVBoxLayout()
        left_col.setSpacing(8)
        left_col.addWidget(QLabel("院系快捷：点击切换，右键删除"))

        btn_scroll = QScrollArea()
        btn_scroll.setWidgetResizable(True)
        btn_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        btn_container = QWidget()
        btn_layout = QVBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)
        btn_layout.setAlignment(Qt.AlignTop)
        btn_scroll.setWidget(btn_container)

        self.dept_btns: Dict[str, QPushButton] = {}
        self.dept_btn_layout = btn_layout
        left_col.addWidget(btn_scroll, 1)

        add_row = QHBoxLayout()
        add_row.setSpacing(8)
        add_row.addWidget(QLabel("新增院系："))
        self.new_dept_input = QLineEdit()
        self.new_dept_input.setPlaceholderText("输入院系名称后点击新增")
        btn_add_dept = QPushButton("新增")
        btn_add_dept.clicked.connect(self.on_add_department)
        add_row.addWidget(self.new_dept_input, 1)
        add_row.addWidget(btn_add_dept)
        left_col.addLayout(add_row)

        body_layout.addLayout(left_col, 1)

        # right: tabs
        right_col = QVBoxLayout()
        right_col.setSpacing(8)
        right_col.addWidget(QLabel("词典编辑"))
        self.tabs = QTabWidget(self)
        self.edits: Dict[str, Tuple[QPlainTextEdit, QPlainTextEdit]] = {}
        for dept in sorted(self._data.keys()):
            self._add_department_tab(dept, self._data.get(dept, {}))
        self._refresh_dept_buttons()
        self.tabs.currentChanged.connect(lambda _: self._sync_dept_buttons())
        right_col.addWidget(self.tabs, 1)
        body_layout.addLayout(right_col, 2)

        layout.addWidget(body, 1)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        layout.addWidget(btns)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

    def _refresh_dept_buttons(self):
        self.dept_btns = {}
        while self.dept_btn_layout.count():
            item = self.dept_btn_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        all_depts = list(set(self._data.keys()) | set(self.edits.keys()))
        ordered = [d for d in PRESET_DEPTS_IF2 if d in all_depts]
        extras = sorted([d for d in all_depts if d not in PRESET_DEPTS_IF2])

        for dept in ordered + extras:
            btn = QPushButton(dept)
            btn.setObjectName("DeptButton")
            btn.setCheckable(True)
            btn.clicked.connect(lambda _=False, d=dept: self._switch_to_department(d))
            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(lambda pos, d=dept, b=btn: self._on_dept_btn_context_menu(pos, d, b))
            self.dept_btn_layout.addWidget(btn)
            self.dept_btns[dept] = btn

        self.dept_btn_layout.addStretch(1)
        self._sync_dept_buttons()

    def _sync_dept_buttons(self):
        current = self.tabs.tabText(self.tabs.currentIndex()) if self.tabs.count() else None
        for dept, btn in list(self.dept_btns.items()):
            btn.setChecked(dept == current)

    def _switch_to_department(self, dept: str):
        if dept not in self.edits:
            self._add_department_tab(dept, {"精确": [], "关键字": []})
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == dept:
                self.tabs.setCurrentIndex(i)
                break
        self._sync_dept_buttons()

    def _add_department_tab(self, dept: str, rules: Optional[Dict[str, List[str]]] = None):
        if not dept or dept in self.edits:
            return
        page = QWidget()
        form = QFormLayout(page)
        form.setLabelAlignment(Qt.AlignRight)
        txt_exact = QPlainTextEdit()
        txt_exact.setPlaceholderText("精确匹配：每行一个班级名称。")
        txt_kw = QPlainTextEdit()
        txt_kw.setPlaceholderText("关键字匹配：每行一个关键字，按长度优先。")
        if rules:
            txt_exact.setPlainText("\n".join(rules.get("精确", [])))
            txt_kw.setPlainText("\n".join(rules.get("关键字", [])))
        form.addRow("精确匹配", txt_exact)
        form.addRow("关键字", txt_kw)
        self.tabs.addTab(page, dept)
        self.edits[dept] = (txt_exact, txt_kw)
        self._refresh_dept_buttons()

    def _on_dept_btn_context_menu(self, pos, dept: str, btn: QPushButton):
        menu = QMenu(self)
        act_del = menu.addAction(f"删除院系：{dept}")
        chosen = menu.exec(btn.mapToGlobal(pos))
        if chosen == act_del:
            ret = QMessageBox.question(self, "确认删除", f"确定删除院系“{dept}”及其全部班级/关键字吗？",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret == QMessageBox.Yes:
                self._remove_department(dept)

    def _remove_department(self, dept: str):
        if dept in self.edits:
            for i in range(self.tabs.count()):
                if self.tabs.tabText(i) == dept:
                    self.tabs.removeTab(i)
                    break
            self.edits.pop(dept, None)
        self._data.pop(dept, None)
        self.result_dict.pop(dept, None)
        self._refresh_dept_buttons()
        if self.tabs.count():
            self.tabs.setCurrentIndex(0)

    def on_add_department(self):
        name = self.new_dept_input.text().strip()
        if not name:
            QMessageBox.information(self, "提示", "请输入院系名称后再新增。")
            return
        if name in self.edits:
            QMessageBox.information(self, "提示", "该院系已存在。")
            return
        self._add_department_tab(name, {"精确": [], "关键字": []})
        self.tabs.setCurrentIndex(self.tabs.count() - 1)
        self.new_dept_input.clear()

    @staticmethod
    def _parse_text(text: str) -> List[str]:
        parts: List[str] = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            for piece in re.split(r'[;,，；\s]+', line):
                piece = piece.strip()
                if piece:
                    parts.append(piece)
        return parts

    def accept(self):
        updated = {}
        for dept, (txt_exact, txt_kw) in self.edits.items():
            updated[dept] = {
                "精确": self._parse_text(txt_exact.toPlainText()),
                "关键字": self._parse_text(txt_kw.toPlainText()),
            }
        self.result_dict = normalize_dictionary(updated)
        super().accept()

def help_text() -> str:
    return """
【软件功能说明】

一、卫生普查
1. 卫生分表：
   - 按楼栋拆分明细
   - 自动排序宿舍号
   - 0分/0.0分标红
   - 生成目录与汇总表

2. 卫生比率：
   - 楼栋遍历统计
   - 支持区间/单间排除
   - 自动计算优秀率、合格率、不合格率

二、学风简报中心
- 导入“简报初稿”与“整改清单”，自动删掉已整改寝室

三、拖拽导入
- 支持 Excel / Word / PDF 拖入
"""

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("使用说明")
        self.setMinimumSize(720, 520)
        lay = QVBoxLayout(self)
        txt = QTextEdit()
        txt.setReadOnly(True)
        txt.setPlainText(help_text())
        lay.addWidget(txt)
        btn = QPushButton("关闭")
        btn.clicked.connect(self.accept)
        lay.addWidget(btn, 0, Qt.AlignRight)

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于本软件")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 标题区域
        title_frame = QFrame()
        title_frame.setObjectName("AboutTitleFrame")
        title_layout = QVBoxLayout(title_frame)
        title_layout.setContentsMargins(0, 0, 0, 0)
        
        app_name = QLabel("宿舍普查数据清洗软件")
        app_name.setObjectName("AboutAppName")
        app_name.setAlignment(Qt.AlignCenter)
        
        version = QLabel(f"版本：{__version__}")
        version.setObjectName("AboutVersion")
        version.setAlignment(Qt.AlignCenter)
        
        title_layout.addWidget(app_name)
        title_layout.addWidget(version)
        layout.addWidget(title_frame)
        
        # 内容区域
        content_frame = QFrame()
        content_layout = QVBoxLayout(content_frame)
        
        # 软件简介
        intro_group = QFrame()
        intro_layout = QVBoxLayout(intro_group)
        intro_title = QLabel("【软件简介】")
        intro_title.setObjectName("AboutSectionTitle")
        intro_text = QLabel(
            "本软件是一款专为宿舍卫生普查数据清洗与处理而设计的桌面应用程序，\n"
            "采用现代化卡片式界面设计，提供直观友好的用户体验。\n"
            "支持Excel、Word、PDF等多种格式的数据导入与导出，\n"
            "提供分表导出、楼栋遍历统计、考勤检查等核心功能。"
        )
        intro_text.setObjectName("AboutText")
        intro_text.setWordWrap(True)
        intro_layout.addWidget(intro_title)
        intro_layout.addWidget(intro_text)
        content_layout.addWidget(intro_group)
        
        # 主要功能
        features_group = QFrame()
        features_layout = QVBoxLayout(features_group)
        features_title = QLabel("【主要功能】")
        features_title.setObjectName("AboutSectionTitle")
        features_text = QLabel(
            "• 卫生分表导出：按楼栋拆分明细，自动排序宿舍号，支持0分标红\n"
            "• 卫生比率统计：楼栋遍历统计，支持区间/单间排除，自动计算优秀率/合格率\n"
            "• 考勤检查：考勤数据处理与汇总，包括上课考勤、晚自习考勤等（开发中）\n"
            "• 院系词典管理：维护班级与院系的映射关系，支持智能匹配\n"
            "• 数据导入导出：支持拖拽导入，自动建议输出文件名"
        )
        features_text.setObjectName("AboutText")
        features_text.setWordWrap(True)
        features_layout.addWidget(features_title)
        features_layout.addWidget(features_text)
        content_layout.addWidget(features_group)
        
        # 技术特点
        tech_group = QFrame()
        tech_layout = QVBoxLayout(tech_group)
        tech_title = QLabel("【技术特点】")
        tech_title.setObjectName("AboutSectionTitle")
        tech_text = QLabel(
            "• 采用PySide6构建现代化UI界面\n"
            "• 支持浅色/深色主题切换\n"
            "• 多线程处理，避免界面卡顿\n"
            "• 自动记录运行日志，便于追踪问题\n"
            "• 配置持久化，自动保存用户设置"
        )
        tech_text.setObjectName("AboutText")
        tech_text.setWordWrap(True)
        tech_layout.addWidget(tech_title)
        tech_layout.addWidget(tech_text)
        content_layout.addWidget(tech_group)
        
        # 更新说明
        update_group = QFrame()
        update_layout = QVBoxLayout(update_group)
        update_title = QLabel("【更新说明】")
        update_title.setObjectName("AboutSectionTitle")
        
        # 处理 build_note 元组
        build_note_str = "\n".join([f"• {item}" for item in __build_note__])
        update_text = QLabel(build_note_str)
        update_text.setObjectName("AboutText")
        update_text.setWordWrap(True)
        
        update_layout.addWidget(update_title)
        update_layout.addWidget(update_text)
        content_layout.addWidget(update_group)
        
        # 版权信息
        copyright_group = QFrame()
        copyright_layout = QVBoxLayout(copyright_group)
        copyright_title = QLabel("【版权信息】")
        copyright_title.setObjectName("AboutSectionTitle")
        copyright_text = QLabel("© 2025 DormHealth Team. All rights reserved.")
        copyright_text.setObjectName("AboutText")
        copyright_text.setAlignment(Qt.AlignCenter)
        copyright_layout.addWidget(copyright_title)
        copyright_layout.addWidget(copyright_text)
        content_layout.addWidget(copyright_group)
        
        # 滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(content_frame)
        layout.addWidget(scroll, 1)
        
        # 按钮
        btn_close = QPushButton("关闭")
        btn_close.setObjectName("PrimaryButton")
        btn_close.setMinimumHeight(36)
        btn_close.clicked.connect(self.accept)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch(1)
        btn_layout.addWidget(btn_close)
        btn_layout.addStretch(1)
        layout.addLayout(btn_layout)
        
        # 设置样式
        self.setStyleSheet("""
            QDialog {
                background: #f6f7f9;
            }
            
            QFrame#AboutTitleFrame {
                background: #ffffff;
                border: 1px solid #e8eaee;
                border-radius: 14px;
                padding: 20px;
                margin-bottom: 10px;
            }
            
            QLabel#AboutAppName {
                font-size: 24px;
                font-weight: 900;
                color: #1f2328;
                margin-bottom: 8px;
            }
            
            QLabel#AboutVersion {
                font-size: 14px;
                color: #667085;
            }
            
            QFrame {
                background: #ffffff;
                border: 1px solid #e8eaee;
                border-radius: 10px;
                padding: 15px;
                margin-bottom: 10px;
            }
            
            QLabel#AboutSectionTitle {
                font-size: 16px;
                font-weight: 800;
                color: #1f2328;
                margin-bottom: 8px;
            }
            
            QLabel#AboutText {
                font-size: 14px;
                color: #4b5563;
                line-height: 1.5;
            }
            
            QPushButton#PrimaryButton {
                background: #0b57d0;
                color: white;
                border: none;
                border-radius: 12px;
                padding: 8px 24px;
                font-weight: 700;
            }
            
            QPushButton#PrimaryButton:hover {
                background: #094bbb;
            }
        """)
