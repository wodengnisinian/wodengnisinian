# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import tempfile
import traceback
from dataclasses import dataclass
from typing import Dict, List, Optional

from PySide6.QtCore import QThread, Signal

from .processing import (
    build_workbook_bytes,
    if2_load_and_clean, if2_build_tables, if2_save_excel,
    load_tabular_file, drop_rectified_rows,
)

@dataclass
class SettingsRun:
    input_path: str
    output_path: str
    max_sheets: int
    font_name: str
    mark_num_zero: bool
    mark_text_zero: bool
    mark_text_zero_dot: bool
    dept_dictionary: Dict[str, Dict[str, List[str]]]
    use_majority_dept: bool

@dataclass
class BriefingSettings:
    report_path: str
    rectified_path: str
    output_path: str
    match_by_room_only: bool
    require_department: bool

class Worker(QThread):
    """界面一导出"""
    progress = Signal(int)
    finished = Signal(str, dict)  # 返回路径和summary
    error = Signal(str)

    def __init__(self, s: SettingsRun):
        super().__init__()
        self.s = s
        self.summary: Dict[str, int] | None = None

    def run(self):
        try:
            self.progress.emit(10)
            data, meta = build_workbook_bytes(
                input_path=self.s.input_path,
                font_name=self.s.font_name,
                max_sheets=self.s.max_sheets,
                mark_num_zero=self.s.mark_num_zero,
                mark_text_zero=self.s.mark_text_zero,
                mark_text_zero_dot=self.s.mark_text_zero_dot,
                dept_dictionary=self.s.dept_dictionary,
                use_majority_dept=self.s.use_majority_dept,
            )
            self.summary = meta
            self.progress.emit(80)
            # 原子写入：先写入临时文件，再重命名
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".tmp", dir=os.path.dirname(self.s.output_path))
            try:
                with os.fdopen(tmp_fd, "wb") as f:
                    f.write(data)
                os.replace(tmp_path, self.s.output_path)
            except Exception:
                os.unlink(tmp_path)
                raise
            self.progress.emit(100)
            self.finished.emit(self.s.output_path, self.summary or {})
        except Exception as e:
            self.error.emit(f"Error occurred: {traceback.format_exc()}")

class WorkerIface2(QThread):
    """界面二导出"""
    progress = Signal(int)
    finished = Signal(str, dict)  # 返回路径和summary
    error = Signal(str)

    def __init__(self, input_path: str, output_path: str, exclude_params: Dict,
                 drop_zero_text: bool, drop_zero_numeric: bool,
                 use_majority_dept: bool,
                 dept_dictionary: Optional[Dict] = None,
                 fallback_df=None):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.exclude_params = exclude_params
        self.drop_zero_text = drop_zero_text
        self.drop_zero_numeric = drop_zero_numeric
        self.use_majority_dept = use_majority_dept
        self.dept_dictionary = dept_dictionary or {}
        self.fallback_df = fallback_df
        self.summary: Dict[str, int] | None = None

    def run(self):
        try:
            self.progress.emit(10)
            all_df, logs, stats = if2_load_and_clean(
                self.input_path,
                self.exclude_params,
                self.drop_zero_text,
                self.drop_zero_numeric,
                self.use_majority_dept,
                self.dept_dictionary,
                self.fallback_df,
            )
            self.summary = stats
            self.progress.emit(45)
            table1, table2 = if2_build_tables(all_df)
            self.progress.emit(70)
            if2_save_excel(table1, table2, logs, stats, self.output_path)
            self.progress.emit(100)
            self.finished.emit(self.output_path, self.summary or {})
        except Exception as e:
            self.error.emit(f"Error occurred: {traceback.format_exc()}")

class WorkerBriefing(QThread):
    progress = Signal(int)
    finished = Signal(str, dict)  # 返回路径和summary
    error = Signal(str)

    def __init__(self, settings: BriefingSettings):
        super().__init__()
        self.settings = settings
        self.summary: Dict[str, int] | None = None

    def run(self):
        try:
            self.progress.emit(10)
            report_df = load_tabular_file(self.settings.report_path)
            self.progress.emit(35)
            rect_df = load_tabular_file(self.settings.rectified_path)
            self.progress.emit(60)
            cleaned = drop_rectified_rows(
                report_df,
                rect_df,
                match_by_room_only=self.settings.match_by_room_only,
                require_department=self.settings.require_department,
            )
            self.progress.emit(85)
            cleaned.to_excel(self.settings.output_path, index=False)
            self.summary = {
                "report_rows": int(len(report_df)),
                "rect_rows": int(len(rect_df)),
                "remaining_rows": int(len(cleaned)),
            }
            self.progress.emit(100)
            self.finished.emit(self.settings.output_path, self.summary or {})
        except Exception as exc:
            self.error.emit(f"Error occurred: {traceback.format_exc()}")
