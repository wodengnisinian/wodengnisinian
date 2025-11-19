# -*- coding: utf-8 -*-
"""
校园宿舍数据处理平台 - 总入口 UI 界面

功能定位：
- 作为“启动器 / 中控台”，统一打开各个功能窗口：
  1）按楼栋拆分导出（界面一 / 界面二）
  2）学风简报中心
  3）院系词典中心
  4）寝室区间排除（按楼栋）
- 不改动各子模块原有 UI 和逻辑，只负责“打开对应窗口”。

命名规则（函数名用中文拼音首字母缩写）：
- 初始化界面 → chu shi hua jie mian → cshjm
- 打开按楼栋拆分导出窗口 → da kai an lou dong chai fen dao chu chuang kou → dkaldcfdck
- 打开学风简报中心窗口 → da kai xue feng jian bao zhong xin chuang kou → dkxfjbzxck
- 打开院系词典中心窗口 → da kai yuan xi ci dian zhong xin chuang kou → dkyxcdzxck
- 打开寝室区间排除窗口 → da kai qin shi qu jian pai chu chuang kou → dknsqjpcck
"""

import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame, QSizePolicy, QMessageBox
)
from PySide6.QtCore import Qt, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QIcon

# === 这里按你实际的文件名来导入 ===
# 1. 按楼栋拆分导出：就是你那份包含“界面一 / 界面二”的大代码文件
from 按楼栋拆分导出 import MainWindow as AldcMainWindow

# 2. 学风简报中心：页面_学风简报中心.py 里的 MainWindow
from 页面_学风简报中心 import MainWindow as XfjbMainWindow

# 3. 院系词典中心：子界面_院系词典中心.py 里的 QWidget
from 子界面_院系词典中心 import 子界面_院系词典中心 as YxCidianWidget

# 4. 寝室区间排除（按楼栋）：子界面_寝室区间排除_按楼栋.py 里的 QWidget
from 子界面_寝室区间排除_按楼栋 import 子界面_寝室区间排除_按楼栋 as QjPaichuWidget


# ---------------- 动画按钮（按下回弹） ----------------
class DongHuaAnNiu(QPushButton):
    """
    动画按钮：轻微“按下缩小 → 回弹”效果
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._press_anim = QPropertyAnimation(self, b"geometry")
        self._release_anim = QPropertyAnimation(self, b"geometry")
        self._orig_geo = None

    def mousePressEvent(self, event):
        self._orig_geo = self.geometry()
        shrink = self._orig_geo.adjusted(2, 2, -2, -2)

        self._press_anim.stop()
        self._press_anim.setDuration(90)
        self._press_anim.setEasingCurve(QEasingCurve.OutQuad)
        self._press_anim.setStartValue(self._orig_geo)
        self._press_anim.setEndValue(shrink)
        self._press_anim.start()

        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if self._orig_geo is None:
            super().mouseReleaseEvent(event)
            return

        self._release_anim.stop()
        self._release_anim.setDuration(120)
        self._release_anim.setEasingCurve(QEasingCurve.OutBack)
        self._release_anim.setStartValue(self.geometry())
        self._release_anim.setEndValue(self._orig_geo)
        self._release_anim.start()

        super().mouseReleaseEvent(event)


# ---------------- 主窗口：总入口 ----------------
class XiaoYuanSuShePingTai(QMainWindow):
    """
    校园宿舍数据处理平台 - 启动器主窗口
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("校园宿舍数据处理平台 - 总入口")
        self.resize(960, 600)

        # 子窗口引用，防止被 Python 回收
        self.ck_aldcf = None         # 按楼栋拆分导出窗口
        self.ck_xfjb = None          # 学风简报中心窗口
        self.ck_yxcd = None          # 院系词典中心窗口（外包 QMainWindow）
        self.ck_qjpc = None          # 寝室区间排除窗口（外包 QMainWindow）

        self.cshjm()
        self.szyzt()

    # ---------- 初始化界面：chu shi hua jie mian ----------
    def cshjm(self):
        """
        初始化总入口界面：大标题 + 功能列表卡片 + 底部说明
        """
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(16)

        # 顶部标题区
        title = QLabel("校园宿舍数据处理平台")
        title.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        title.setStyleSheet("font-size: 26px; font-weight: 700;")

        sub = QLabel(
            "面向周一普查 + 学风简报 + 日常统计的一站式工具：\n"
            "• 按楼栋拆分导出（界面一 / 界面二）\n"
            "• 学风简报中心（上周普查 × 本周整改）\n"
            "• 院系词典中心（班级/专业 → 院系映射）\n"
            "• 寝室区间排除（按楼栋维护排除区间规则）"
        )
        sub.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        sub.setWordWrap(True)
        sub.setStyleSheet("color: #6B7280; font-size: 13px;")

        root.addWidget(title)
        root.addWidget(sub)

        # 中间卡片：放 4 个大按钮
        card = QFrame()
        card.setObjectName("MainCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(24, 24, 24, 24)
        card_layout.setSpacing(12)

        # 每个功能一行按钮 + 描述
        card_layout.addWidget(self._scg_gnxx(
            "① 按楼栋拆分导出",
            "分楼栋分表、总分为 0 明细、目录、表一 / 表二统计，包含“寝室区间排除 + 去零”新口径。",
            self.dkaldcfdck
        ))

        card_layout.addWidget(self._scg_gnxx(
            "② 学风简报中心",
            "导入上周普查数据生成初稿，再导入本周整改结果，一键删除已整改寝室，导出本周学风简报。",
            self.dkxfjbzxck
        ))

        card_layout.addWidget(self._scg_gnxx(
            "③ 院系词典中心",
            "维护“班级/关键字 → 院系”的映射词典，为后续统计提供统一口径。",
            self.dkyxcdzxck
        ))

        card_layout.addWidget(self._scg_gnxx(
            "④ 寝室区间排除（按楼栋）",
            "为不同楼栋维护排除区间（整段/单间），给界面二统计或其他模块复用。",
            self.dknsqjpcck
        ))

        root.addWidget(card, 1)

        # 底部版权 / 版本信息（简单写死，你也可以自己改）
        bottom = QLabel("© 校园宿舍数据处理平台   当前为开发版（用于内部使用）")
        bottom.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        bottom.setStyleSheet("color: #9CA3AF; font-size: 11px;")
        root.addWidget(bottom)

        self.setCentralWidget(central)

    # ---------- 生成功能项行：shang cheng gong neng xin xi ----------
    def _scg_gnxx(self, bt_text: str, ms_text: str, slot_func):
        """
        生成一行“按钮 + 说明文字”的小卡片
        bt_text：按钮标题
        ms_text：功能说明
        slot_func：点击按钮后要调用的函数
        """
        row = QFrame()
        row.setObjectName("FuncRow")
        lay = QHBoxLayout(row)
        lay.setContentsMargins(12, 12, 12, 12)
        lay.setSpacing(12)

        btn = DongHuaAnNiu(bt_text)
        btn.setObjectName("FuncButton")
        btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn.setMinimumWidth(260)
        btn.setMinimumHeight(42)
        btn.clicked.connect(slot_func)

        lab = QLabel(ms_text)
        lab.setWordWrap(True)
        lab.setStyleSheet("color: #4B5563; font-size: 13px;")

        lay.addWidget(btn, 0, Qt.AlignTop)
        lay.addWidget(lab, 1)

        return row

    # ---------- 主题样式设置：she zhi yang shi ----------
    def szyzt(self):
        """
        设置统一的浅色主题样式（可以后面再扩展深色）
        """
        accent = "#2563EB"
        accent_hover = "#1D4ED8"
        accent_press = "#1E40AF"
        bg = "#F9FAFB"
        card = "#FFFFFF"
        border = "#E5E7EB"
        text = "#111827"

        self.setStyleSheet(f"""
        QMainWindow {{
            background: {bg};
            color: {text};
        }}
        #MainCard {{
            background: {card};
            border-radius: 18px;
            border: 1px solid {border};
        }}
        #FuncRow {{
            background: transparent;
            border-radius: 12px;
        }}
        #FuncRow:hover {{
            background: #F3F4F6;
        }}
        QPushButton#FuncButton {{
            background: {accent};
            color: white;
            border-radius: 999px;
            padding: 8px 18px;
            font-size: 14px;
            font-weight: 600;
            border: none;
        }}
        QPushButton#FuncButton:hover {{
            background: {accent_hover};
        }}
        QPushButton#FuncButton:pressed {{
            background: {accent_press};
        }}
        QPushButton {{
            font-family: "Microsoft YaHei", "微软雅黑", system-ui;
        }}
        QLabel {{
            font-family: "Microsoft YaHei", "微软雅黑", system-ui;
        }}
        """)

    # ---------- 打开按楼栋拆分导出窗口：da kai an lou dong chai fen dao chu chuang kou ----------
    def dkaldcfdck(self):
        """
        打开“按楼栋拆分导出”窗口（界面一 / 界面二）
        """
        try:
            if self.ck_aldcf is None:
                self.ck_aldcf = AldcMainWindow()
            self.ck_aldcf.show()
            self.ck_aldcf.raise_()
            self.ck_aldcf.activateWindow()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开“按楼栋拆分导出”失败：\n{e}")

    # ---------- 打开学风简报中心窗口：da kai xue feng jian bao zhong xin chuang kou ----------
    def dkxfjbzxck(self):
        """
        打开“学风简报中心”窗口
        """
        try:
            if self.ck_xfjb is None:
                self.ck_xfjb = XfjbMainWindow()
            self.ck_xfjb.show()
            self.ck_xfjb.raise_()
            self.ck_xfjb.activateWindow()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开“学风简报中心”失败：\n{e}")

    # ---------- 打开院系词典中心窗口：da kai yuan xi ci dian zhong xin chuang kou ----------
    def dkyxcdzxck(self):
        """
        打开“院系词典中心”窗口（把 QWidget 包一层 QMainWindow）
        """
        try:
            if self.ck_yxcd is None:
                # 外包一层 QMainWindow 让它有自己的标题栏和尺寸
                from PySide6.QtWidgets import QMainWindow as _QMainWindow
                wrapper = _QMainWindow(self)
                widget = YxCidianWidget(wrapper)
                wrapper.setCentralWidget(widget)
                wrapper.setWindowTitle("院系词典中心")
                wrapper.resize(920, 600)
                self.ck_yxcd = wrapper
            self.ck_yxcd.show()
            self.ck_yxcd.raise_()
            self.ck_yxcd.activateWindow()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开“院系词典中心”失败：\n{e}")

    # ---------- 打开寝室区间排除窗口：da kai qin shi qu jian pai chu chuang kou ----------
    def dknsqjpcck(self):
        """
        打开“寝室区间排除（按楼栋）”窗口（把 QWidget 包一层 QMainWindow）
        """
        try:
            if self.ck_qjpc is None:
                from PySide6.QtWidgets import QMainWindow as _QMainWindow
                wrapper = _QMainWindow(self)
                widget = QjPaichuWidget(wrapper)
                wrapper.setCentralWidget(widget)
                wrapper.setWindowTitle("寝室区间排除（按楼栋）")
                wrapper.resize(920, 600)
                self.ck_qjpc = wrapper
            self.ck_qjpc.show()
            self.ck_qjpc.raise_()
            self.ck_qjpc.activateWindow()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开“寝室区间排除（按楼栋）”失败：\n{e}")


# ---------------- 入口函数 ----------------
def main():
    app = QApplication(sys.argv)
    win = XiaoYuanSuShePingTai()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
