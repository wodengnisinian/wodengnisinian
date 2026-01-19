# -*- coding: utf-8 -*-
from __future__ import annotations

# ===== 版本信息 =====
"""Metadata for the building-based exporter."""

__version__ = "1.1.0.16"
__build_note__ = (
    "整合为单一数据处理脚本，保留统一清洗链；运行日志入口移至主界面左侧按钮区域，设置弹窗精简通用页；",
    "修复清洗阶段缺失列（如班级）导致统计中断的问题；",
    "界面二分组统计强化院系标量化，避免 first_dept 维度异常；",
    "运行日志增强：记录输入/输出与处理统计，界面一条件格式在空表时不再触发异常；",
    "界面二原始数据统一从“原始数据输入”弹窗获取，路径缺失时自动唤起弹窗提醒；",
    "界面一新增寝室院系多数决选项，可按人数占比统一院系判定",
    "界面一与界面二原始数据合并为同一输入文件路径，共用同一份明细数据。",
    "界面二“排除文本 0/0.0”选项调整为仅排除文本 0.0（含 0.00/0.000分 等写法），界面一总分为0明细逻辑保持不变。"
)
__history__ = """
1.0.0.0: 初始版本发布。
。。。。。。（省略部分更新效果）
1.1.0.0: 增加了可选排除与去除总分为0的功能，完善了界面一和界面二的操作流程。
1.1.0.1: UI界面排版优化
"""

import os
import re
import sys
import io
import json
import datetime as dt
from copy import deepcopy
from dataclasses import dataclass
from typing import Iterable, List, Dict, Tuple, Optional, Set, Any, Callable

import pandas as pd
import numpy as np

from PySide6.QtCore import Qt, QThread, Signal, QSettings
from PySide6.QtGui import QFontDatabase
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QLabel, QLineEdit, QPushButton, QSpinBox, QHBoxLayout, QVBoxLayout,
    QFormLayout, QTabWidget, QComboBox, QCheckBox, QDialog,
    QFrame, QStyle, QDialogButtonBox, QProgressBar, QPlainTextEdit, QGridLayout, QScrollArea, QSizePolicy, QTextEdit,
    QMenu,
)

# ---------------- 工具（界面一用） ---------------- #
HIDDEN_WS_RE = re.compile(r'[\u00A0\u2000-\u200B\u202F\u205F\u3000\uFEFF]')
SCORE_RE = re.compile(r'[-+]?\d+(?:\.\d+)?')
ZERO_TEXT_RE = re.compile(r'^0(?:\.0+)?(?:\s*分)?$', re.IGNORECASE)
BUILDING_EXTRACT_RE = re.compile(r'([^\s]*苑\s*(?:\d+|[一二三四五六七八九十]+)\s*(?:号|栋))')
DETAIL_COLUMNS = ["序号", "楼栋", "宿舍号", "院系", "班级", "学生姓名", "评分状态", "总分", "检查时间", "打分原因"]


def choose_detail_sheet(xls_path: str) -> str:
    """选择最像明细表的工作表：列数>=10 且 行数最多"""
    xls = pd.ExcelFile(xls_path)
    best_name, best_rows = None, -1
    for s in xls.sheet_names:
        try:
            t = pd.read_excel(xls_path, sheet_name=s, header=None)
            t = t.dropna(how="all").dropna(axis=1, how="all")
            if t.shape[1] >= 10 and t.shape[0] > best_rows:
                best_name, best_rows = s, t.shape[0]
        except Exception:
            pass
    return best_name or xls.sheet_names[0]


def clean_text(s):
    if pd.isna(s): return np.nan
    return HIDDEN_WS_RE.sub("", str(s)).strip()


def normalize_plain_text(value) -> str:
    """Remove exotic whitespaces and ensure a safe string output for general文本字段。"""
    text = clean_text(value)
    if pd.isna(text):
        return ""
    return str(text)


def normalize_department(value) -> str:
    """标准化院系字符串，便于判定有效/无效行。"""
    return normalize_plain_text(value)


def ensure_scalar_department(value) -> str:
    """保证院系字段为单值字符串，避免列表/集合导致分组报错。"""
    if isinstance(value, (list, tuple, set)):
        try:
            value = next(iter(value))
        except StopIteration:
            value = ""
    return normalize_department(value)


def is_valid_department(value) -> bool:
    text = normalize_department(value)
    if not text:
        return False
    return text.lower() != "nan"


def parse_score(v):
    if pd.isna(v): return np.nan
    s = str(v).strip()
    if not s:
        return np.nan
    compact = re.sub(r"\s+", "", s)
    compact = compact.replace("分", "")
    pure_match = re.match(r"^[-+]?\d+(?:\.\d+)?$", compact)
    if pure_match:
        try:
            return float(pure_match.group())
        except ValueError:
            return np.nan
    if re.search(r"\d{2,4}[-/]\d{1,2}[-/]\d{1,2}", compact):
        return np.nan
    m = SCORE_RE.search(compact)
    return float(m.group()) if m else np.nan


def pick_majority(series: pd.Series) -> str:
    """院系多数决，平局取首个院系。"""
    counts = series.value_counts()
    if counts.empty:
        return ""
    top = counts.max()
    candidates = [dept for dept, cnt in counts.items() if cnt == top]
    if len(candidates) == 1:
        return candidates[0]
    return series.iloc[0]


# ---------------- 全局通用清洗函数（约定名） ---------------- #
def qkh(df: pd.DataFrame) -> pd.DataFrame:
    """去空行：删除整行为空或全为空白的记录。"""
    if df.empty:
        return df.copy()
    blank_mask = df.apply(lambda col: col.map(lambda x: normalize_plain_text(x) == ""))
    return df.loc[~blank_mask.all(axis=1)].reset_index(drop=True)


def qbjty_text(text: object) -> str:
    """全半角统一：常用空格/逗号标准化。"""
    if pd.isna(text):
        return ""
    s = str(text)
    replacements = {
        ord("，"): ",",
        ord("。"): ".",
        ord("　"): " ",
        ord("\u3000"): " ",
    }
    s = s.translate(replacements)
    return normalize_plain_text(s)


def zfcgfh(df: pd.DataFrame, columns: Optional[List[str]] = None) -> pd.DataFrame:
    """字符串规范化：去除首尾空格与控制字符。"""
    cols = columns or df.select_dtypes(include=[object]).columns.tolist()
    for col in cols:
        df[col] = df[col].map(normalize_plain_text)
    return df


def qbjty(df: pd.DataFrame, columns: Optional[List[str]] = None) -> pd.DataFrame:
    """全半角统一：对指定列应用 qbjty_text。"""
    cols = columns or df.select_dtypes(include=[object]).columns.tolist()
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    for col in cols:
        df[col] = df[col].map(qbjty_text)
    return df


def scgjwk(df: pd.DataFrame, building_col: str = "楼栋", room_col: str = "宿舍号") -> pd.DataFrame:
    """删除关键字段为空的行（楼栋/宿舍号）。"""
    mask = (
        df[building_col].map(normalize_plain_text).eq("") |
        df[room_col].map(normalize_plain_text).eq("")
    )
    return df.loc[~mask].reset_index(drop=True)


def sshszh(series: pd.Series) -> pd.Series:
    """宿舍号数字化：解析数字，无法转换则返回 NaN。"""
    return series.map(lambda x: extract_room_num(x))


def zfshzh(series: pd.Series) -> pd.Series:
    """总分数字化：统一转为浮点数。"""
    return series.map(parse_score)


def _is_text_zero_0dot0_only(s: str) -> bool:
    """
    专供界面二使用：仅把“0.0/0.00/0.000/0.0分/0.00分...”视为文本 0 分，
    不再把纯“0/0分”当作文本 0。
    """
    if not s:
        return False
    txt = re.sub(r"\s+", "", s).replace("分", "")
    # 至少有一个小数点
    if "." not in txt:
        return False
    m = re.fullmatch(r"0\.0+", txt)
    return bool(m)


def flzf0(
    df: pd.DataFrame,
    score_col: str,
    drop_text_zero: bool,
    drop_numeric_zero: bool,
    *,
    text_mode: str = "both",
    extra_text_pred: Optional[Callable[[str], bool]] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    分离总分为 0 的记录，返回（非 0 数据，0 分数据）。

    参数含义：
    - drop_text_zero: 是否按“文本写法”去零；
    - drop_numeric_zero: 是否按“数值结果”为 0 去零；
    - text_mode:
        * "both"       : 文本 0 / 0.0 / 0分 / 0.0分 ... 都算文本 0 （界面一/旧逻辑）
        * "only_0dot0" : 仅把 “0.0 / 0.00 / 0.000 / 0.0分 / 0.00分 …” 视为文本 0（界面二）
    - extra_text_pred: 额外的文本判断函数，可选。
    """
    # 原始文本列
    score_raw = df[score_col].map(normalize_plain_text)
    # 数值列
    score_num = score_raw.map(parse_score)

    text_zero = pd.Series(False, index=df.index)
    num_zero = pd.Series(False, index=df.index)

    # ---------- 文本 0 判定 ----------
    if drop_text_zero:
        if text_mode == "only_0dot0":
            text_zero = score_raw.map(_is_text_zero_0dot0_only)
        else:
            text_zero = score_raw.map(
                lambda x: bool(ZERO_TEXT_RE.match(str(x).strip()))
            )

        if extra_text_pred is not None:
            text_zero = text_zero | score_raw.map(
                lambda x: bool(extra_text_pred(str(x)))
            )

    # ---------- 数值 0 判定 ----------
    if drop_numeric_zero:
        num_zero = score_num.fillna(np.nan) == 0

    if not drop_text_zero and not drop_numeric_zero:
        return df.copy(), df.iloc[0:0].copy()

    mask_zero = text_zero | num_zero

    zero_df = df[mask_zero].copy()
    zero_df["_zero_text"] = text_zero.loc[mask_zero].fillna(False)
    zero_df["_zero_num"] = num_zero.loc[mask_zero].fillna(False)

    keep_df = df[~mask_zero].copy()
    return keep_df, zero_df


def ajqc(df: pd.DataFrame, keys: List[str]) -> pd.DataFrame:
    """按主键去重，保留第一条。"""
    return df.drop_duplicates(subset=keys).reset_index(drop=True)


def yyqjpc(df: pd.DataFrame, ex: Dict) -> pd.DataFrame:
    """应用寝室区间排除规则。"""
    if not ex.get("enabled"):
        return df

    def _normalize_garden(cfg: Dict) -> Dict[int, Dict[str, object]]:
        out: Dict[int, Dict[str, object]] = {}
        if not isinstance(cfg, dict):
            return out
        for key, val in cfg.items():
            try:
                no = int(key)
            except (TypeError, ValueError):
                continue
            ranges = []
            for pair in (val.get("ranges") if isinstance(val, dict) else []) or []:
                try:
                    lo, hi = int(pair[0]), int(pair[1])
                    if lo > hi:
                        lo, hi = hi, lo
                    ranges.append((lo, hi))
                except Exception:
                    continue
            singles = set()
            for s in (val.get("singles") if isinstance(val, dict) else []) or []:
                try:
                    singles.add(int(s))
                except Exception:
                    continue
            if ranges or singles:
                out[no] = {"ranges": ranges, "singles": singles}
        return out

    lan_cfg = _normalize_garden(ex.get("lan", {}))
    mei_cfg = _normalize_garden(ex.get("mei", {}))

    def hit_row(row) -> bool:
        g, no = parse_garden_and_no(row.get("楼栋"))
        if g is None or no is None:
            return False
        cfg = lan_cfg.get(no) if g == "兰" else mei_cfg.get(no)
        if not cfg:
            return False
        n = extract_room_num(row.get("宿舍号"))
        if n is None:
            return False
        for lo, hi in cfg.get("ranges", []):
            if lo <= n <= hi:
                return True
        return n in cfg.get("singles", set())

    mask_ex = df.apply(hit_row, axis=1)
    return df[~mask_ex].reset_index(drop=True)


def normalize_building(s):
    """规范楼栋为 x苑x号（如：'梅苑1号'），满足“楼栋名规范化”判定。"""
    if pd.isna(s):
        return np.nan
    s = clean_text(s)
    garden, num = parse_garden_and_no(s)
    if garden and num is not None:
        return f"{garden}苑{num}号"
    m = BUILDING_EXTRACT_RE.search(s)
    if m:
        return clean_text(m.group(1))
    if "苑" in s and ("号" in s or "栋" in s):
        return s
    return s.strip()


def clean_building_text(value) -> str:
    """优先用规范化楼栋名，失败则返回基础清理后的字符串。"""
    norm = normalize_building(value)
    if pd.isna(norm):
        return normalize_plain_text(value)
    return normalize_plain_text(norm)


def room_sort_key(v):
    s = str(v).strip()
    m = re.search(r"(\d+)", s)
    return (0, int(m.group(1)), s) if m else (1, float('inf'), s)


# ---------------- 界面二：常量与解析工具（按列值排除） ---------------- #
PRESET_DEPTS_IF2: List[str] = [
    "机电工程系", "信息工程系", "艺术设计系", "经济管理系",
    "文化与旅游系", "轻工化工系", "建筑工程系", "怀卡托国际学院",
]


def default_department_dictionary() -> Dict[str, Dict[str, List[str]]]:
    return {dept: {"精确": [], "关键字": []} for dept in PRESET_DEPTS_IF2}


def normalize_dictionary(data: Optional[Dict[str, Dict[str, List[str]]]]) -> Dict[str, Dict[str, List[str]]]:
    result = default_department_dictionary()
    if not isinstance(data, dict):
        return result
    for dept, rules in data.items():
        if dept not in result:
            result[dept] = {"精确": [], "关键字": []}
        exact = rules.get("精确", []) if isinstance(rules, dict) else []
        keywords = rules.get("关键字", []) if isinstance(rules, dict) else []
        cleaned_exact = []
        for item in exact:
            norm = normalize_plain_text(item)
            if norm:
                cleaned_exact.append(norm)
        cleaned_keywords = []
        for item in keywords:
            norm = normalize_plain_text(item)
            if norm:
                cleaned_keywords.append(norm)
        result[dept]["精确"] = cleaned_exact
        result[dept]["关键字"] = cleaned_keywords
    return result


def infer_department_from_class(cls_name: str, dictionary: Dict[str, Dict[str, List[str]]]) -> str:
    if not dictionary:
        return ""
    matched = match_department_by_class_name(cls_name, dictionary)
    return matched or ""


def fill_department_with_dictionary(df: pd.DataFrame,
                                    dictionary: Dict[str, Dict[str, List[str]]],
                                    dept_col: str = "院系",
                                    class_col: str = "班级") -> pd.DataFrame:
    if dictionary is None or not isinstance(df, pd.DataFrame):
        return df
    if dept_col not in df.columns or class_col not in df.columns:
        return df
    mask = ~df[dept_col].apply(is_valid_department)
    if not mask.any():
        return df
    inferred = df.loc[mask, class_col].map(lambda cls: infer_department_from_class(cls, dictionary))
    df.loc[mask, dept_col] = inferred
    return df


OUTPUT_ORDER_IF2: List[str] = [
    "信息工程系", "建筑工程系", "怀卡托国际学院", "文化与旅游系",
    "机电工程系", "经济管理系", "艺术设计系", "轻工化工系",
]

TOTAL_ROW_NAME_IF2 = "总计"

BUILDING_NAME_RE = re.compile(r'([兰梅])苑\s*(\d+|[一二三四五六七八九十]+)\s*(?:号|栋)')

CHINESE_NUM = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
}


def parse_chinese_number(text: str) -> Optional[int]:
    if not text:
        return None
    if text.isdigit():
        return int(text)
    total = 0
    if text == "十":
        return 10
    if len(text) == 2 and text.startswith("十") and text[1] in CHINESE_NUM:
        return 10 + CHINESE_NUM[text[1]]
    if len(text) == 2 and text.endswith("十") and text[0] in CHINESE_NUM:
        return CHINESE_NUM[text[0]] * 10
    for ch in text:
        if ch in CHINESE_NUM:
            total = total * 10 + CHINESE_NUM[ch]
        else:
            return None
    return total if total > 0 else None


def if2_is_building_sheet(name: str) -> bool:
    """只遍历“兰/梅苑x号/栋”样式的楼栋表"""
    return BUILDING_NAME_RE.search(str(name).strip()) is not None


def if2_pct(x: int, base: int) -> str:
    """百分比字符串（两位小数）；分母0→0%"""
    return "0%" if base == 0 else f"{(x / base) * 100:.2f}%"


def parse_garden_and_no(bld: str) -> tuple[str | None, int | None]:
    if pd.isna(bld):
        return None, None
    s = str(bld).strip()
    m = BUILDING_NAME_RE.search(s)
    if not m:
        return None, None
    return m.group(1), parse_chinese_number(m.group(2))


def extract_room_num(x) -> int | None:
    if pd.isna(x):
        return None
    m = re.search(r'(\d+)', str(x).strip())
    return int(m.group(1)) if m else None


def split_floor_room(n: int | None) -> tuple[int | None, int | None]:
    if n is None:
        return None, None
    return n // 100, n % 100


def is_valid_interval_record(start, end) -> Tuple[bool, Optional[Tuple[int, int]]]:
    """判定“区间行”是否有效，满足要求 34。"""
    try:
        a = int(str(start))
        b = int(str(end))
    except (TypeError, ValueError):
        return False, None
    if a > b:
        return False, None
    return True, (a, b)


def is_valid_single_room(room) -> Tuple[bool, Optional[int]]:
    """判定“单间行”是否有效，满足要求 35。"""
    text = normalize_plain_text(room)
    if not text:
        return False, None
    try:
        value = int(text)
    except ValueError:
        return False, None
    return True, value


def expand_room_ranges(range_records: Iterable, existing_rooms: Optional[Iterable] = None) -> List[Dict[str, object]]:
    """将有效区间展开为单间记录，满足要求 36。"""
    seen: Set[int] = set()
    if existing_rooms:
        for room in existing_rooms:
            ok, value = is_valid_single_room(room)
            if ok and value is not None:
                seen.add(value)

    new_rooms: List[Dict[str, object]] = []
    for rec in range_records:
        if isinstance(rec, dict):
            start = rec.get("start")
            end = rec.get("end")
        else:
            try:
                start, end = rec
            except (TypeError, ValueError):
                continue
        valid, bounds = is_valid_interval_record(start, end)
        if not valid or bounds is None:
            continue
        lo, hi = bounds
        for room in range(lo, hi + 1):
            if room in seen:
                continue
            seen.add(room)
            new_rooms.append({
                "room": room,
                "remark": f"由区间 {lo}-{hi} 展开",
            })
    return new_rooms


def parse_interval_text(text: str) -> List[Tuple[int, int]]:
    ranges: List[Tuple[int, int]] = []
    if not text:
        return ranges
    for part in re.split(r'[,\uFF0C\u3001；;,，\s]+', text.strip()):
        if not part:
            continue
        if "-" not in part:
            continue
        try:
            a, b = part.split("-", 1)
            lo, hi = int(a), int(b)
            if lo > hi:
                lo, hi = hi, lo
            ranges.append((lo, hi))
        except ValueError:
            continue
    return ranges


def parse_single_text(text: str) -> List[int]:
    singles: List[int] = []
    if not text:
        return singles
    for part in re.split(r'[,\uFF0C\u3001；;,，\s]+', text.strip()):
        if not part:
            continue
        if part.isdigit():
            singles.append(int(part))
    return singles


def describe_exclusion(ex: Dict) -> str:
    """把“寝室区间排除”的配置转换成日志描述，便于追溯。"""

    def _fmt(garden_key: str, label: str) -> str:
        rows = []
        cfg = ex.get(garden_key, {}) if isinstance(ex, dict) else {}
        for num in sorted(cfg.keys()):
            ranges = cfg[num].get("ranges", []) if isinstance(cfg[num], dict) else []
            singles = cfg[num].get("singles", []) if isinstance(cfg[num], dict) else []
            rng_txt = ",".join(f"{lo}-{hi}" for lo, hi in ranges) or "无区间"
            sgl_txt = ",".join(str(s) for s in singles) or "无单间"
            rows.append(f"{num}栋：{rng_txt}；单间：{sgl_txt}")
        return f"{label}（{'; '.join(rows) if rows else '未填写'}）"

    return "；".join([_fmt("lan", "兰苑"), _fmt("mei", "梅苑")])


def match_department_by_class_name(cls_name: str, dictionary: Dict[str, Dict[str, List[str]]]) -> Optional[str]:
    """根据班级名称在“院系词典中心”中寻找院系，满足要求 37-38。"""
    cls = normalize_plain_text(cls_name)
    if not cls:
        return None

    for dept, rules in dictionary.items():
        for exact in rules.get("精确", []):
            if cls == normalize_plain_text(exact):
                return dept

    keyword_hits: List[Tuple[str, str]] = []
    for dept, rules in dictionary.items():
        for kw in rules.get("关键字", []):
            kw_norm = normalize_plain_text(kw)
            if kw_norm and kw_norm in cls:
                keyword_hits.append((dept, kw_norm))

    if not keyword_hits:
        return None

    keyword_hits.sort(key=lambda item: (-len(item[1]), item[0]))
    best_len = len(keyword_hits[0][1])
    best_candidates = [dept for dept, kw in keyword_hits if len(kw) == best_len]
    return best_candidates[0]


def drop_rectified_rows(report_df: pd.DataFrame,
                        rectified_df: pd.DataFrame,
                        match_by_room_only: bool = True,
                        require_department: bool = False) -> pd.DataFrame:
    """按照“学风简报中心”删行口径，移除已整改寝室（要求 39-40）。"""
    include_dept = (not match_by_room_only) and require_department

    def _build_key(row: pd.Series) -> Optional[Tuple[str, str, Optional[str]]]:
        building = clean_building_text(row.get("楼栋", ""))
        room = normalize_plain_text(row.get("宿舍号", ""))
        if not building or not room:
            return None
        if include_dept:
            dept = normalize_department(row.get("院系", ""))
            return building, room, dept
        return building, room, None

    rect_keys: Set[Tuple[str, str, Optional[str]]] = set()
    for _, row in rectified_df.iterrows():
        key = _build_key(row)
        if key is not None:
            rect_keys.add(key)

    if not rect_keys:
        return report_df.copy()

    mask_to_drop = report_df.apply(lambda row: (_build_key(row) in rect_keys), axis=1)
    return report_df[~mask_to_drop].reset_index(drop=True)


def load_tabular_file(path: str) -> pd.DataFrame:
    """加载 Excel/Word/PDF 为 DataFrame，用于“简报中心”导入。"""

    target_cols = {"楼栋", "宿舍号"}

    def _build_df(rows: List[List[str]]) -> pd.DataFrame:
        if not rows:
            raise ValueError("未在文件中发现可用表格数据。")
        cleaned = [[(c or "").strip() for c in row] for row in rows]
        header_idx = 0
        for idx, row in enumerate(cleaned):
            if target_cols.issubset(set(row)):
                header_idx = idx
                break
        header = cleaned[header_idx]
        data = cleaned[header_idx + 1:] if len(cleaned) > header_idx + 1 else []
        col_count = len(header)
        if any(len(r) != col_count for r in data):
            fixed = []
            for r in data:
                row = (r + [""] * col_count)[:col_count]
                fixed.append(row)
            data = fixed
        cols = [c.strip() or f"列{i + 1}" for i, c in enumerate(header)]
        return pd.DataFrame(data, columns=cols)

    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path)

    if ext in (".doc", ".docx"):
        try:
            import docx  # type: ignore
        except Exception as exc:
            raise ImportError("解析 Word 需要安装 python-docx 库。") from exc

        doc = docx.Document(path)
        fallback = None
        matched: List[pd.DataFrame] = []
        for table in doc.tables:
            rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            if not rows:
                continue
            df = _build_df(rows)
            if target_cols.issubset(set(df.columns)):
                matched.append(df)
            elif fallback is None:
                fallback = df
        if matched:
            return pd.concat(matched, ignore_index=True)
        if fallback is not None:
            return fallback
        raise ValueError("未在 Word 文件中找到表格。")

    if ext == ".pdf":
        try:
            import pdfplumber  # type: ignore
        except Exception as exc:
            raise ImportError("解析 PDF 需要安装 pdfplumber 库。") from exc

        fallback = None
        matched: List[pd.DataFrame] = []
        with pdfplumber.open(path) as pdf:  # type: ignore[attr-defined]
            for page in pdf.pages:
                tables = page.extract_tables() or []
                if not tables:
                    single = page.extract_table()
                    if single:
                        tables = [single]
                for table in tables:
                    if not table:
                        continue
                    rows = [[(cell or "").strip() for cell in row] for row in table]
                    df = _build_df(rows)
                    if target_cols.issubset(set(df.columns)):
                        matched.append(df)
                    elif fallback is None:
                        fallback = df
        if matched:
            return pd.concat(matched, ignore_index=True)
        if fallback is not None:
            return fallback
        raise ValueError("未在 PDF 文件中找到表格。")

    raise ValueError("仅支持 Excel/Word/PDF 文件。")


# ---------------- 界面二：数据装载与统计（含“可选排除 + 去零”） ---------------- #
def if2_load_and_clean(input_path: str,
                       ex: Dict,
                       drop_zero_text: bool,
                       drop_zero_numeric: bool,
                       use_majority_dept: bool,
                       dept_dictionary: Optional[Dict[str, Dict[str, List[str]]]] = None,
                       fallback_df: Optional[pd.DataFrame] = None
                       ) -> Tuple[pd.DataFrame, Dict[str, List[str]], Dict[str, int]]:
    xls = pd.ExcelFile(input_path)
    building_frames: List[pd.DataFrame] = []
    logs: Dict[str, List[str]] = {
        "ignored_non_building": [],
        "ignored_missing_columns": [],
        "used_building_sheets": [],
        "policy": [
            "仅遍历表名含“兰/梅苑×号/栋”的工作表；",
            "院系单元格如含多个，仅取第一个（全角逗号→半角）；",
            "唯一键 = (楼栋, 宿舍号, 院系first)；",
            "优秀≥90；不合格<60；其余为合格；",
            "仅统计 8 个预置系部；输出顺序按指定列表；",
        ],
    }
    if ex.get("enabled"):
        logs["policy"].append("按楼栋配置区间/单间排除宿舍。")
        logs["policy"].append("排除条件：" + describe_exclusion(ex))
    if drop_zero_text:
        logs["policy"].append("排除总分文本为 0.0（含 0.00/0.000分 等写法，仅界面二）。")
    if drop_zero_numeric:
        logs["policy"].append("排除总分数值为 0。")
    if use_majority_dept:
        logs["policy"].append("寝室院系由人数占比决定，平局时回退至单元格首个院系。")

    # ✅ 追加一条“本次开关状态”说明，方便日志查看
    logs["policy"].append(
        f"本次去零开关：文本0.0={'开启' if drop_zero_text else '关闭'}；数值0={'开启' if drop_zero_numeric else '关闭'}；"
        f"区间排查={'开启' if ex.get('enabled') else '关闭'}。"
    )

    stats = {
        "sheets_total": len(xls.sheet_names),
        "sheets_used": 0,
        "rows_raw": 0,
        "rows_after_structure": 0,
        "rows_valid_dept": 0,
        "rows_after_zero": 0,
        "rows_after_exclusion": 0,
        "rows_final": 0,
        "zero_text_removed": 0,
        "zero_numeric_removed": 0,
        "excluded_rows": 0,
    }

    def process_frame(df: pd.DataFrame, sheet_name: str = ""):
        stats["rows_raw"] += len(df)
        df.columns = df.columns.str.strip()
        must_cols = {"楼栋", "宿舍号", "院系", "总分"}
        if not must_cols.issubset(df.columns):
            if sheet_name:
                missing = sorted(must_cols - set(df.columns))
                logs["ignored_missing_columns"].append(f"{sheet_name}（缺列：{missing}）")
            return

        for optional in ("班级",):
            if optional not in df.columns:
                df[optional] = ""

        df = qkh(df)
        df = zfcgfh(df)
        df = qbjty(df, ["楼栋", "宿舍号", "院系", "班级", "总分"])
        df = scgjwk(df, "楼栋", "宿舍号")

        df["楼栋"] = df["楼栋"].map(clean_building_text)
        df["宿舍号"] = df["宿舍号"].map(normalize_plain_text)
        df["宿舍号_num"] = sshszh(df["宿舍号"])
        df = fill_department_with_dictionary(df, dept_dictionary or {}, dept_col="院系", class_col="班级")
        dept_raw = df["院系"].map(normalize_plain_text).str.replace("，", ",")
        first_dept = (
            dept_raw.str.split(",", n=1, expand=True)[0]
            .map(ensure_scalar_department)
        )
        df["first_dept"] = first_dept
        stats["rows_after_structure"] += len(df)

        # === 去零：界面二只对文本 0.0 做“文本去零”，数值 0 独立控制 ===
        df, zero_df = flzf0(
            df,
            "总分",
            drop_text_zero=drop_zero_text,
            drop_numeric_zero=drop_zero_numeric,
            text_mode="only_0dot0",
        )
        if isinstance(zero_df, pd.DataFrame):
            if "_zero_text" in zero_df.columns:
                stats["zero_text_removed"] += int(zero_df["_zero_text"].sum())
            if "_zero_num" in zero_df.columns:
                stats["zero_numeric_removed"] += int(zero_df["_zero_num"].sum())

        df = df[df["first_dept"].apply(is_valid_department)].copy()
        df["总分_num"] = zfshzh(df["总分"])
        stats["rows_valid_dept"] += len(df)
        # ✅ 记录“去零后（且院系有效）行数”
        stats["rows_after_zero"] += len(df)

        before_ex = len(df)
        df = yyqjpc(df, ex)
        stats["excluded_rows"] += int(before_ex - len(df))
        stats["rows_after_exclusion"] += len(df)

        building_frames.append(df[["楼栋", "宿舍号", "first_dept", "总分_num"]].copy())
        if sheet_name:
            logs["used_building_sheets"].append(sheet_name)

    for sheet in xls.sheet_names:
        if not if2_is_building_sheet(sheet):
            logs["ignored_non_building"].append(sheet)
            continue

        df = xls.parse(sheet, dtype=str)
        stats["sheets_used"] += 1
        process_frame(df, sheet)

    if not building_frames and fallback_df is not None:
        stats["sheets_used"] += 1
        process_frame(fallback_df.copy(), "界面一源数据")

    if not building_frames:
        return pd.DataFrame(columns=["楼栋", "宿舍号", "first_dept", "总分_num"]), logs, stats

    all_df = pd.concat(building_frames, ignore_index=True)

    keep_cols = ["楼栋", "宿舍号", "first_dept", "总分_num"]
    all_df = all_df.loc[:, [c for c in keep_cols if c in all_df.columns]]
    all_df = all_df.loc[:, ~all_df.columns.duplicated()]
    all_df["first_dept"] = all_df["first_dept"].map(ensure_scalar_department)

    if all_df.empty:
        return all_df, logs, stats

    if use_majority_dept:
        all_df["first_dept_final"] = all_df.groupby(["楼栋", "宿舍号"])['first_dept'].transform(pick_majority)
    else:
        all_df["first_dept_final"] = all_df["first_dept"]

    all_df["first_dept_final"] = all_df["first_dept_final"].map(ensure_scalar_department)
    all_df = all_df[all_df["first_dept_final"].isin(PRESET_DEPTS_IF2)]
    all_df = ajqc(all_df, ["楼栋", "宿舍号", "first_dept_final"])
    all_df = all_df.rename(columns={"first_dept_final": "first_dept"})
    all_df["first_dept"] = all_df["first_dept"].map(ensure_scalar_department)
    stats["rows_final"] = int(len(all_df))
    return all_df, logs, stats


def if2_build_tables(all_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """生成表一/表二"""
    all_df = all_df.dropna(subset=["总分_num"])
    all_df = all_df.loc[:, [c for c in ["楼栋", "宿舍号", "first_dept", "总分_num"] if c in all_df.columns]].copy()
    all_df = all_df.loc[:, ~all_df.columns.duplicated()]
    all_df["first_dept"] = all_df["first_dept"].map(ensure_scalar_department)
    excellent_mask = all_df["总分_num"] >= 90
    fail_mask = all_df["总分_num"] < 60

    checked_counts = all_df.groupby("first_dept").size()
    excellent_counts = all_df[excellent_mask].groupby("first_dept").size()
    fail_counts = all_df[fail_mask].groupby("first_dept").size()

    def pick(s: pd.Series, dept: str) -> int:
        return int(s.get(dept, 0))

    # 表一：优秀/不及格
    table1_rows = []
    for i, dept in enumerate(OUTPUT_ORDER_IF2, start=1):
        table1_rows.append({
            "序号": i,
            "系部": dept,
            "优秀寝室": pick(excellent_counts, dept),
            "不合格寝室": pick(fail_counts, dept),
        })
    table1_rows.append({
        "序号": len(table1_rows) + 1,
        "系部": TOTAL_ROW_NAME_IF2,
        "优秀寝室": sum(r["优秀寝室"] for r in table1_rows),
        "不合格寝室": sum(r["不合格寝室"] for r in table1_rows),
    })
    table1 = pd.DataFrame(table1_rows, columns=["序号", "系部", "优秀寝室", "不合格寝室"])

    # 表二：检查/优秀/合格/不合格 + 各率
    table2_rows = []
    for i, dept in enumerate(OUTPUT_ORDER_IF2, start=1):
        checked = pick(checked_counts, dept)
        excellent = pick(excellent_counts, dept)
        fail = pick(fail_counts, dept)
        qualified = checked - excellent - fail
        table2_rows.append({
            "序号": i,
            "系部": dept,
            "检查寝室/间": checked,
            "优秀寝室/间": excellent,
            "优秀率": if2_pct(excellent, checked),
            "合格寝室/间": qualified,
            "合格率": if2_pct(qualified, checked),
            "不合格寝室/间": fail,
            "不合格率": if2_pct(fail, checked),
        })

    total_checked = sum(r["检查寝室/间"] for r in table2_rows)
    total_excellent = sum(r["优秀寝室/间"] for r in table2_rows)
    total_fail = sum(r["不合格寝室/间"] for r in table2_rows)
    total_qualified = total_checked - total_excellent - total_fail

    table2_rows.append({
        "序号": len(table2_rows) + 1,
        "系部": TOTAL_ROW_NAME_IF2,
        "检查寝室/间": total_checked,
        "优秀寝室/间": total_excellent,
        "优秀率": if2_pct(total_excellent, total_checked),
        "合格寝室/间": total_qualified,
        "合格率": if2_pct(total_qualified, total_checked),
        "不合格寝室/间": total_fail,
        "不合格率": if2_pct(total_fail, total_checked),
    })
    table2 = pd.DataFrame(
        table2_rows,
        columns=[
            "序号", "系部", "检查寝室/间", "优秀寝室/间", "优秀率",
            "合格寝室/间", "合格率", "不合格寝室/间", "不合格率"
        ],
    )
    return table1, table2


def if2_save_excel(
    table1: pd.DataFrame,
    table2: pd.DataFrame,
    logs: Dict[str, List[str]],
    stats: Dict[str, int],
    output_path: str
) -> None:
    """
    界面二：按模板输出 3 个工作表：
    1）优秀与不及格表
    2）检查与各率（公式与模板一致）
    3）日志（日期 + 操作内容），日志中增加数据概况、阶段占比、去零/排除统计。
    """
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        wb = writer.book

        # 通用单元格样式
        header_fmt = wb.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        cell_fmt = wb.add_format({
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })
        pct_fmt = wb.add_format({
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "num_format": "0.00%",   # 百分比两位小数
        })

        # ---------- 先把 table1 / table2 做个类型清洗 ----------
        t1 = table1.copy()
        for col in ["优秀寝室", "不合格寝室"]:
            if col in t1.columns:
                t1[col] = pd.to_numeric(t1[col], errors="coerce").fillna(0).astype(int)

        t2 = table2.copy()
        for col in ["检查寝室/间", "优秀寝室/间", "不合格寝室/间"]:
            if col in t2.columns:
                t2[col] = pd.to_numeric(t2[col], errors="coerce").fillna(0).astype(int)

        # 小工具：安全计算百分比
        def _pct(num: int, den: int) -> str:
            if not den:
                return "0%"
            return f"{(num / den) * 100:.2f}%"

        # =====================================================
        #  Sheet1：优秀与不及格表
        # =====================================================
        sheet1_name = "优秀与不及格表"
        ws1 = wb.add_worksheet(sheet1_name)

        headers1 = ["序号", "系部", "优秀寝室", "不合格寝室"]
        for col, h in enumerate(headers1):
            ws1.write(0, col, h, header_fmt)

        ordered_rows: List[Dict[str, Any]] = []
        for dept in OUTPUT_ORDER_IF2:
            row = t1[t1["系部"] == dept]
            if row.empty:
                ordered_rows.append({"系部": dept, "优秀寝室": 0, "不合格寝室": 0})
            else:
                r0 = row.iloc[0]
                ordered_rows.append({
                    "系部": dept,
                    "优秀寝室": int(r0.get("优秀寝室", 0)),
                    "不合格寝室": int(r0.get("不合格寝室", 0)),
                })

        total_row = t1[t1["系部"] == TOTAL_ROW_NAME_IF2]
        if not total_row.empty:
            tr = total_row.iloc[0]
            total_excellent = int(tr.get("优秀寝室", 0))
            total_fail = int(tr.get("不合格寝室", 0))
        else:
            total_excellent = sum(r["优秀寝室"] for r in ordered_rows)
            total_fail = sum(r["不合格寝室"] for r in ordered_rows)

        start_row_s1 = 1
        for i, r in enumerate(ordered_rows):
            row_idx = start_row_s1 + i
            excel_row_no = row_idx + 1
            ws1.write_formula(row_idx, 0, "=ROW()-1", cell_fmt)
            ws1.write(row_idx, 1, r["系部"], cell_fmt)
            ws1.write_number(row_idx, 2, r["优秀寝室"], cell_fmt)
            ws1.write_number(row_idx, 3, r["不合格寝室"], cell_fmt)

        total_row_idx_s1 = start_row_s1 + len(ordered_rows)
        total_excel_row_s1 = total_row_idx_s1 + 1
        first_data_excel_row_s1 = start_row_s1 + 1
        last_data_excel_row_s1 = first_data_excel_row_s1 + len(ordered_rows) - 1

        ws1.write_formula(total_row_idx_s1, 0, "=ROW()-1", cell_fmt)
        ws1.write(total_row_idx_s1, 1, TOTAL_ROW_NAME_IF2, cell_fmt)
        ws1.write_formula(
            total_row_idx_s1, 2,
            f"=SUM(C{first_data_excel_row_s1}:C{last_data_excel_row_s1})",
            cell_fmt,
        )
        ws1.write_formula(
            total_row_idx_s1, 3,
            f"=SUM(D{first_data_excel_row_s1}:D{last_data_excel_row_s1})",
            cell_fmt,
        )

        ws1.set_column(0, 0, 8)
        ws1.set_column(1, 1, 16)
        ws1.set_column(2, 3, 14)

        # =====================================================
        #  Sheet2：检查与各率
        # =====================================================
        sheet2_name = "检查与各率"
        ws2 = wb.add_worksheet(sheet2_name)

        headers2 = [
            "序号", "系部", "检查寝室/间", "优秀寝室/间", "优秀率",
            "合格寝室/间", "合格率", "不合格寝室/间", "不合格率",
        ]
        for col, h in enumerate(headers2):
            ws2.write(0, col, h, header_fmt)

        dept_counts: Dict[str, Tuple[int, int, int]] = {}
        if not t2.empty:
            for _, row in t2.iterrows():
                dept = row["系部"]
                chk = int(row.get("检查寝室/间", 0))
                exc = int(row.get("优秀寝室/间", 0))
                fail = int(row.get("不合格寝室/间", 0))
                dept_counts[dept] = (chk, exc, fail)

        start_row_s2 = 1
        for i, dept in enumerate(OUTPUT_ORDER_IF2):
            row_idx = start_row_s2 + i
            excel_row_no = row_idx + 1

            chk, exc, fail = dept_counts.get(dept, (0, 0, 0))

            ws2.write_formula(row_idx, 0, "=ROW()-1", cell_fmt)
            ws2.write(row_idx, 1, dept, cell_fmt)
            ws2.write_number(row_idx, 2, chk, cell_fmt)
            ws2.write_formula(
                row_idx, 3,
                f"='{sheet1_name}'!C{excel_row_no}",
                cell_fmt,
            )
            ws2.write_formula(
                row_idx, 4,
                f"=IF(C{excel_row_no}>0,D{excel_row_no}/C{excel_row_no},0)",
                pct_fmt,
            )
            ws2.write_formula(
                row_idx, 7,
                f"='{sheet1_name}'!D{excel_row_no}",
                cell_fmt,
            )
            ws2.write_formula(
                row_idx, 5,
                f"=C{excel_row_no}-D{excel_row_no}-H{excel_row_no}",
                cell_fmt,
            )
            ws2.write_formula(
                row_idx, 6,
                f"=IF(C{excel_row_no}>0,F{excel_row_no}/C{excel_row_no},0)",
                pct_fmt,
            )
            ws2.write_formula(
                row_idx, 8,
                f"=IF(C{excel_row_no}>0,H{excel_row_no}/C{excel_row_no},0)",
                pct_fmt,
            )

        total_row_idx_s2 = start_row_s2 + len(OUTPUT_ORDER_IF2)
        total_excel_row_s2 = total_row_idx_s2 + 1
        first_data_excel_row_s2 = start_row_s2 + 1
        last_data_excel_row_s2 = first_data_excel_row_s2 + len(OUTPUT_ORDER_IF2) - 1

        ws2.write_formula(total_row_idx_s2, 0, "=ROW()-1", cell_fmt)
        ws2.write(total_row_idx_s2, 1, "合计", cell_fmt)

        ws2.write_formula(
            total_row_idx_s2, 2,
            f"=SUM(C{first_data_excel_row_s2}:C{last_data_excel_row_s2})",
            cell_fmt,
        )
        ws2.write_formula(
            total_row_idx_s2, 3,
            f"='{sheet1_name}'!C{total_excel_row_s1}",
            cell_fmt,
        )
        ws2.write_formula(
            total_row_idx_s2, 4,
            f"=IF(C{total_excel_row_s2}>0,D{total_excel_row_s2}/C{total_excel_row_s2},0)",
            pct_fmt,
        )
        ws2.write_formula(
            total_row_idx_s2, 7,
            f"='{sheet1_name}'!D{total_excel_row_s1}",
            cell_fmt,
        )
        ws2.write_formula(
            total_row_idx_s2, 5,
            f"=C{total_excel_row_s2}-D{total_excel_row_s2}-H{total_excel_row_s2}",
            cell_fmt,
        )
        ws2.write_formula(
            total_row_idx_s2, 6,
            f"=IF(C{total_excel_row_s2}>0,F{total_excel_row_s2}/C{total_excel_row_s2},0)",
            pct_fmt,
        )
        ws2.write_formula(
            total_row_idx_s2, 8,
            f"=IF(C{total_excel_row_s2}>0,H{total_excel_row_s2}/C{total_excel_row_s2},0)",
            pct_fmt,
        )

        ws2.set_column(0, 0, 8)
        ws2.set_column(1, 1, 16)
        ws2.set_column(2, 8, 14)

        # =====================================================
        #  Sheet3：日志（详细到爆炸）
        # =====================================================
        wslog = wb.add_worksheet("日志")
        wslog.write(0, 0, "日期", header_fmt)
        wslog.write(0, 1, "操作内容", header_fmt)

        ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_lines: List[str] = []

        def _s(key: str) -> int:
            if not stats:
                return 0
            try:
                return int(stats.get(key, 0))
            except Exception:
                return 0

        rows_raw = _s("rows_raw")
        rows_struct = _s("rows_after_structure")
        rows_valid_dept = _s("rows_valid_dept")
        rows_after_zero = _s("rows_after_zero")
        rows_after_ex = _s("rows_after_exclusion")
        rows_final = _s("rows_final")
        z_text = _s("zero_text_removed")
        z_num = _s("zero_numeric_removed")
        excl = _s("excluded_rows")
        sheets_total = _s("sheets_total")
        sheets_used = _s("sheets_used")
        sheets_other = max(sheets_total - sheets_used, 0)

        policies = logs.get("policy", [])
        if policies:
            log_lines.append("统计口径：" + "；".join(policies))

        if rows_raw or rows_final:
            overview = (
                f"运行概况：原始行={rows_raw}；"
                f"结构化后={rows_struct}；"
                f"有效院系行={rows_valid_dept}；"
                f"去零后={rows_after_zero}；"
                f"区间/单间排除后={rows_after_ex}；"
                f"最终统计行={rows_final}。"
            )
            log_lines.append(overview)

            phase = (
                "阶段占比："
                f"结构化/原始={_pct(rows_struct, rows_raw)}；"
                f"有效院系/结构化={_pct(rows_valid_dept, rows_struct)}；"
                f"去零后/有效院系={_pct(rows_after_zero, rows_valid_dept)}；"
                f"排除后/去零后={_pct(rows_after_ex, rows_after_zero)}；"
                f"最终/原始={_pct(rows_final, rows_raw)}。"
            )
            log_lines.append(phase)

        if z_text or z_num:
            detail_parts: List[str] = []
            if z_text:
                detail_parts.append(
                    f"文本 0.0 行数={z_text}（占结构化行 {_pct(z_text, rows_struct)}）"
                )
            if z_num:
                detail_parts.append(
                    f"数值 0 行数={z_num}（占结构化行 {_pct(z_num, rows_struct)}）"
                )
            log_lines.append("去零明细：" + "；".join(detail_parts))
        else:
            log_lines.append("去零明细：本次未检测到文本 0.0 或数值 0 需要排除的记录。")

        if excl:
            log_lines.append(
                f"区间/单间排除：共排除 {excl} 行（占去零后行 {_pct(excl, rows_after_zero)}）。"
            )
        else:
            log_lines.append("区间/单间排除：未配置或未命中需排除的寝室。")

        if sheets_total:
            log_lines.append(
                f"工作表统计：总工作表数={sheets_total}；"
                f"被识别为楼栋表={sheets_used}；"
                f"其它类型工作表={sheets_other}。"
            )

        for key, label in [
            ("used_building_sheets", "已使用楼栋表"),
            ("ignored_non_building", "忽略（非楼栋表）"),
            ("ignored_missing_columns", "忽略（缺必需列）"),
        ]:
            items = logs.get(key, [])
            for s in items:
                log_lines.append(f"{label}：{s}")

        if not log_lines:
            log_lines.append("本次运行未记录到特殊说明。")

        for i, text in enumerate(log_lines, start=1):
            wslog.write(i, 0, ts, cell_fmt)
            wslog.write(i, 1, text, cell_fmt)

        wslog.set_column(0, 0, 20)
        wslog.set_column(1, 1, 80)


# ---------------- 界面一（保留） ---------------- #
def load_detail_dataframe(input_path: str, dept_dictionary: Dict[str, Dict[str, List[str]]]) -> Tuple[
    pd.DataFrame, Dict[str, int]]:
    cols_required = DETAIL_COLUMNS
    sheet = choose_detail_sheet(input_path)

    preview = pd.read_excel(input_path, sheet_name=sheet, header=None, nrows=1)
    header_hits = 0
    if not preview.empty:
        header_hits = sum(str(v).strip() in cols_required for v in preview.iloc[0].tolist())
    skiprows = 1 if header_hits >= 3 else 0

    raw_num = pd.read_excel(input_path, sheet_name=sheet, header=None, skiprows=skiprows)
    raw_str = pd.read_excel(input_path, sheet_name=sheet, header=None, skiprows=skiprows, dtype=str)

    for raw in (raw_num, raw_str):
        if raw.shape[1] < 10:
            for _ in range(10 - raw.shape[1]):
                raw[raw.shape[1]] = np.nan
    raw_num = raw_num.iloc[:, :10]; raw_num.columns = cols_required
    raw_str = raw_str.iloc[:, :10]; raw_str.columns = cols_required

    blank_mask = raw_str.apply(lambda col: col.map(lambda x: normalize_plain_text(x) == ""))
    keep_mask = ~blank_mask.all(axis=1)
    raw_num = raw_num.loc[keep_mask].reset_index(drop=True)
    raw_str = raw_str.loc[keep_mask].reset_index(drop=True)

    raw_num["总分_rawstr"] = raw_str["总分"].astype(str)

    raw_num["院系"] = raw_num["院系"].apply(normalize_department)
    raw_num = fill_department_with_dictionary(raw_num, dept_dictionary, dept_col="院系", class_col="班级")
    valid_mask = raw_num["院系"].apply(is_valid_department)
    df = raw_num[valid_mask].copy()

    df["楼栋_norm"] = df["楼栋"].map(normalize_building)
    meta = {
        "sheet": sheet,
        "raw_rows": int(len(raw_num)),
        "valid_rows": int(len(df)),
    }
    return df, meta


def build_workbook_bytes(
        input_path: str,
        font_name: str,
        max_sheets: int,
        mark_num_zero: bool,
        mark_text_zero: bool,
        mark_text_zero_dot: bool,
        dept_dictionary: Dict[str, Dict[str, List[str]]],
        use_majority_dept: bool,
) -> Tuple[bytes, Dict[str, int]]:
    """
    界面一：
      - 每栋 6 张正常表（你原来按楼栋拆分那套）
      - 额外 2 张总表（不及格寝室明细 / 0.0分寝室明细）
      - 目录 + 版本信息 保留
      - ✅ 每张明细表首行开启筛选
      - ✅ 不及格明细按“楼栋 + 宿舍号”从小到大排序
    """
    # 先加载/清洗明细
    df1, meta = load_detail_dataframe(input_path, dept_dictionary)
    cols_required = DETAIL_COLUMNS
    sheet = meta["sheet"]

    # 多数决院系（可选）
    if use_majority_dept and not df1.empty:
        df1["院系"] = (
            df1.groupby(["楼栋_norm", "宿舍号"])["院系"]
            .transform(pick_majority)
            .fillna(df1["院系"])
        )

    # 拆出一个精简视图用于楼栋分表
    slim = df1[["序号", "楼栋_norm", "宿舍号", "院系", "总分", "总分_rawstr"]].copy()

    def sort_key_bld(x):
        text = str(x or "").strip()
        prefix = text.split("苑", 1)[0]
        has_garden = 0 if "苑" in text else 1
        m = re.search(r"(\d+)", text)
        num = int(m.group(1)) if m else float("inf")
        return (prefix, has_garden, num, text)

    candidates_series = slim["楼栋_norm"].dropna().astype(str).str.strip()
    candidates_series = candidates_series[candidates_series.astype(str).str.lower() != "nan"]
    candidates = [c for c in candidates_series.unique().tolist() if c]
    candidates = sorted(candidates, key=sort_key_bld)
    if max_sheets > 0:
        targets = candidates[:max_sheets]
    else:
        targets = candidates

    # ==== 这里重新做一次总分解析，给 2 张汇总表用 ====
    df_full = df1.copy()

    # 保障文本列干净
    df_full["总分_rawstr"] = df_full["总分_rawstr"].map(normalize_plain_text)
    df_full["总分_num"] = df_full["总分_rawstr"].map(parse_score)

    score_raw = df_full["总分_rawstr"]
    score_num = df_full["总分_num"]

    # 0.0 分表：仅保留“0.0 / 0.00 / 0.000 / 0.0分 / 0.00分 …”这类写法
    zero_mask = score_raw.map(_is_text_zero_0dot0_only)
    zero_df = df_full[zero_mask].copy()

    # 不及格表：0 ≤ 总分 < 60，排除掉已经进 0.0 分表的行
    # （只要 parse_score 有效数值即可，NaN 自动被排除）
    fail_mask = score_num.notna() & (score_num >= 0) & (score_num < 60) & (~zero_mask)
    fail_df = df_full[fail_mask].copy()

    # ✅【新增】按“楼栋 + 宿舍号”排序不及格/0.0 明细
    # 使用已有的 room_sort_key，保证“101, 102, 201...”这种自然顺序
    for _tmp in (fail_df, zero_df):
        if not _tmp.empty:
            _tmp["_room_key_"] = _tmp["宿舍号"].map(room_sort_key)
            # 楼栋内按宿舍号升序；也可以只按 _room_key_ 排，如果你想完全无视楼栋顺序的话
            _tmp.sort_values(by=["楼栋_norm", "_room_key_"], inplace=True)
            _tmp.drop(columns=["_room_key_"], inplace=True)

    import xlsxwriter

    buffer = io.BytesIO()
    summary_meta = {
        "sheet": sheet,
        "raw_rows": meta["raw_rows"],
        "valid_rows": meta["valid_rows"],
        "sheet_count": int(len(targets)),       # 楼栋表数量
        "fail_rows": int(len(fail_df)),         # 不及格寝室条数（含 0 分，但已排除 0.0 系）
        "zero_rows": int(len(zero_df)),         # 0.0 分条数
    }

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        wb = writer.book
        header_bold = wb.add_format({
            "bold": True, "font_name": font_name, "font_size": 11,
            "align": "center", "valign": "vcenter", "border": 1
        })
        header_norm = wb.add_format({
            "bold": False, "font_name": font_name, "font_size": 11,
            "align": "center", "valign": "vcenter", "border": 1
        })
        cell_fmt = wb.add_format({
            "font_name": font_name, "font_size": 11,
            "align": "center", "valign": "vcenter", "border": 1
        })
        red_fill = wb.add_format({
            "font_name": font_name, "font_size": 11,
            "bg_color": "#FFCCCC"
        })

        # ---------- 界面一：楼栋分表（每栋 6 张内部分表，逻辑保持不变） ----------
        summary_rows = []
        keep_headers = ["序号", "楼栋", "宿舍号", "院系", "总分"]
        for name in targets:
            sub = slim[slim["楼栋_norm"].astype(str).str.strip() == name].copy()
            sub["_room_key_"] = sub["宿舍号"].map(room_sort_key)
            sub = sub.sort_values("_room_key_").drop(columns=["_room_key_"])
            out = sub[["序号", "楼栋_norm", "宿舍号", "院系", "总分", "总分_rawstr"]].copy()
            out = out.rename(columns={"楼栋_norm": "楼栋", "总分_rawstr": "总分原值(隐藏)"})

            ws_name = (name if len(name) <= 31 else name[:31]) or "未命名"
            out.to_excel(writer, sheet_name=ws_name, index=False, header=False, startrow=1)
            ws = writer.sheets[ws_name]

            headers = keep_headers + ["总分原值(隐藏)"]
            for j, col in enumerate(headers):
                fmt = header_bold if col in keep_headers else header_norm
                ws.write(0, j, col, fmt)

            ws.set_column(0, len(headers) - 1, 12, cell_fmt)
            ws.set_row(0, 18)
            ws.set_column(5, 5, None, None, {"hidden": True})

            nrows = len(out)
            if nrows >= 1:
                for r in range(1, nrows + 1):
                    ws.write_formula(r, 0, "=ROW()-1", cell_fmt)

                start_row, end_row = 2, nrows + 1
                rng = f"A{start_row}:E{end_row}"
                if mark_num_zero:
                    ws.conditional_format(rng, {"type": "formula", "criteria": "=$E2=0", "format": red_fill})
                if mark_text_zero:
                    ws.conditional_format(rng, {"type": "formula", "criteria": '=$F2="0"', "format": red_fill})
                if mark_text_zero_dot:
                    ws.conditional_format(rng, {"type": "formula", "criteria": '=$F2="0.0"', "format": red_fill})

                # ✅【新增】首行开启筛选
                ws.autofilter(0, 0, nrows, len(headers) - 1)

            summary_rows.append([ws_name, nrows])

        # ---------- 2 张总表：不及格寝室明细 / 0.0分寝室明细 ----------
        def _export_special(df_src: pd.DataFrame, sheet_name: str):
            if df_src.empty:
                # 空表就不建工作表，避免一堆空 sheet
                return

            # 防止有列缺失，缺的先补空
            base_cols = [
                "序号", "楼栋", "宿舍号", "院系", "总分",
                "班级", "学生姓名", "评分状态", "检查时间", "打分原因",
            ]
            for col in base_cols:
                if col not in df_src.columns:
                    df_src[col] = ""

            sp = df_src[base_cols + ["总分_rawstr"]].copy()
            sp = sp.rename(columns={"总分_rawstr": "总分原值(隐藏)"})

            sp.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=1)
            ws_sp = writer.sheets[sheet_name]

            cols_all = base_cols + ["总分原值(隐藏)"]
            for j, col in enumerate(cols_all):
                fmt = header_bold if col in ["序号", "楼栋", "宿舍号", "院系", "总分"] else header_norm
                ws_sp.write(0, j, col, fmt)

            ws_sp.set_row(0, 18)
            ws_sp.set_column(0, len(cols_all) - 1, 12, cell_fmt)
            # 隐藏“总分原值(隐藏)”列
            ws_sp.set_column(len(cols_all) - 1, len(cols_all) - 1, None, None, {"hidden": True})

            for r in range(1, len(sp) + 1):
                ws_sp.write_formula(r, 0, "=ROW()-1", cell_fmt)

            # ✅【新增】不及格/0.0 明细也加上筛选
            ws_sp.autofilter(0, 0, len(sp), len(cols_all) - 1)

        _export_special(fail_df, "不及格寝室明细")
        _export_special(zero_df, "0.0分寝室明细")

        # 目录
        toc = pd.DataFrame(summary_rows, columns=["工作表名", "行数"])
        toc.to_excel(writer, sheet_name="目录", index=False, header=False, startrow=1)
        ws_toc = writer.sheets["目录"]
        ws_toc.write(0, 0, "工作表名", header_norm)
        ws_toc.write(0, 1, "行数", header_norm)
        ws_toc.set_column(0, 1, 18, cell_fmt)
        ws_toc.set_row(0, 18)

        # 版本信息
        meta_df = pd.DataFrame({
            "字段": ["版本号", "构建时间", "说明"],
            "值": [__version__, dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), __build_note__],
        })
        try:
            meta_df.to_excel(writer, sheet_name="版本信息", index=False)
        except Exception:
            # 某些极端情况下 sheet 名冲突就算了，不影响主流程
            pass

    buffer.seek(0)
    return buffer.getvalue(), summary_meta



# ---------------- 任务线程 ---------------- #
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
    finished = Signal(str)
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
            with open(self.s.output_path, "wb") as f:
                f.write(data)
            self.progress.emit(100)
            self.finished.emit(self.s.output_path)
        except Exception as e:
            self.error.emit(str(e))


class WorkerIface2(QThread):
    """界面二导出（按列值排除 + 去零）"""
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, input_path: str, output_path: str, exclude_params: Dict,
                 drop_zero_text: bool, drop_zero_numeric: bool,
                 use_majority_dept: bool,
                 dept_dictionary: Optional[Dict[str, Dict[str, List[str]]]] = None,
                 fallback_df: Optional[pd.DataFrame] = None):
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
            self.finished.emit(self.output_path)
        except Exception as e:
            self.error.emit(str(e))


class WorkerBriefing(QThread):
    progress = Signal(int)
    finished = Signal(str)
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
            self.finished.emit(self.settings.output_path)
        except Exception as exc:
            self.error.emit(str(exc))


# ---------------- 自定义控件 ---------------- #
class DropLineEdit(QLineEdit):
    """支持拖拽 Excel 文件的输入框"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        allowed_ext = (".xlsx", ".xls", ".doc", ".docx", ".pdf")
        if e.mimeData().hasUrls():
            urls = e.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith(allowed_ext):
                e.acceptProposedAction()
                return
        super().dragEnterEvent(e)

    def dropEvent(self, e):
        allowed_ext = (".xlsx", ".xls", ".doc", ".docx", ".pdf")
        urls = e.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path.lower().endswith(allowed_ext):
                self.setText(path)
                return
        super().dropEvent(e)


# ---------------- 运行日志工具 ---------------- #
def append_runtime_log(qs: QSettings, text: str):
    """把运行日志追加到 QSettings（通用设置中展示）"""
    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prev = qs.value("runtime/logs", "")
    newlog = f"[{ts}] {text}"
    qs.setValue("runtime/logs", (prev + "\n" + newlog).strip())


def get_bool_setting(qs: QSettings, key: str, default: bool) -> bool:
    value = qs.value(key, default)
    if isinstance(value, str):
        return value.lower() == "true"
    return bool(value)


def set_bool_setting(qs: QSettings, key: str, value: bool):
    qs.setValue(key, bool(value))


def read_dictionary_setting(qs: QSettings) -> Dict[str, Dict[str, List[str]]]:
    raw = qs.value("dictionary/json", "")
    if not raw:
        return default_department_dictionary()
    try:
        data = json.loads(raw)
    except Exception:
        return default_department_dictionary()
    return normalize_dictionary(data)


def save_dictionary_setting(qs: QSettings, data: Dict[str, Dict[str, List[str]]]):
    normalized = normalize_dictionary(data)
    qs.setValue("dictionary/json", json.dumps(normalized, ensure_ascii=False, indent=2))


# ---------------- 设置中心（只保留通用 + 版本/日志） ---------------- #
class SettingsDialog(QDialog):
    """只包含通用外观/目录 + 版本历史 + 运行日志展示"""

    def __init__(self, parent=None, qsettings: QSettings | None = None, show_logs: bool = False):
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

        # —— 通用 —— #
        pg_general = QWidget(); tabs.addTab(pg_general, "通用")
        g1 = QFormLayout(pg_general)
        self.cmb_theme = QComboBox(); self.cmb_theme.addItems(["浅色", "深色"])
        self.chk_open_after = QCheckBox("导出完成后自动打开文件夹")
        self.edt_default_dir = QLineEdit(); self.edt_default_dir.setPlaceholderText("默认输出目录（留空=跟随输入文件所在目录）")
        btn_pick_dir = QPushButton("选择…")

        def pick_dir():
            d = QFileDialog.getExistingDirectory(self, "选择默认输出目录", "")
            if d:
                self.edt_default_dir.setText(d)

        btn_pick_dir.clicked.connect(pick_dir)
        row = QHBoxLayout(); row.addWidget(self.edt_default_dir, 1); row.addWidget(btn_pick_dir)
        g1.addRow("主题：", self.cmb_theme)
        g1.addRow("", self.chk_open_after)
        g1.addRow("默认输出目录：", row)

        self.font_combo = QComboBox(); self.font_combo.setEditable(False)
        try:
            families = list(dict.fromkeys(
                ["仿宋_GB2312", "FangSong", "宋体", "SimSun", "黑体", "SimHei", "微软雅黑", "Microsoft YaHei", "等线", "DengXian"]
                + list(QFontDatabase.families())
            ))
        except TypeError:
            families = ["仿宋_GB2312", "FangSong", "宋体", "SimSun", "黑体", "SimHei", "微软雅黑", "Microsoft YaHei",
                        "等线", "DengXian"]
        self.font_combo.addItems(families)
        if "仿宋_GB2312" in families:
            self.font_combo.setCurrentText("仿宋_GB2312")
        self.font_custom = QLineEdit(); self.font_custom.setPlaceholderText("自定义字体名（可选）")
        row_font = QHBoxLayout(); row_font.setSpacing(8)
        row_font.addWidget(self.font_combo, 1); row_font.addWidget(self.font_custom, 1)
        g1.addRow("Excel 字体：", row_font)
        hint_ex = QLabel("区间示例：101-120,201-220；单间示例：101,203（界面二可直接输入空框）")
        hint_ex.setWordWrap(True); hint_ex.setStyleSheet("color:#6B7280;")
        g1.addRow("排除示例：", hint_ex)

        # —— 版本与历史 —— #
        pg_version = QWidget(); tabs.addTab(pg_version, "版本信息")
        v1 = QVBoxLayout(pg_version)
        lbl_version = QLabel(f"版本：{__version__}\n\n更新说明：\n{__build_note__}\n\n历史版本：\n{__history__}")
        lbl_version.setWordWrap(True)
        v1.addWidget(lbl_version)

        # —— 运行日志 —— #
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
            "font_sel": "仿宋_GB2312",
            "font_custom": "",
        }

    def load(self):
        d = self.defaults()
        self.cmb_theme.setCurrentText(self.qs.value("theme", d["theme"]))
        self.chk_open_after.setChecked(get_bool_setting(self.qs, "open_after", d["open_after"]))
        self.edt_default_dir.setText(self.qs.value("default_dir", d["default_dir"]))
        fsel = self.qs.value("font_sel", d["font_sel"])
        if fsel and fsel in [self.font_combo.itemText(i) for i in range(self.font_combo.count())]:
            self.font_combo.setCurrentText(fsel)
        self.font_custom.setText(self.qs.value("font_custom", d["font_custom"]))
        logs = self.qs.value("runtime/logs", "")
        self.txt_logs.setPlainText(logs)

    def restore_defaults(self):
        d = self.defaults()
        self.cmb_theme.setCurrentText(d["theme"])
        self.chk_open_after.setChecked(d["open_after"])
        self.edt_default_dir.setText(d["default_dir"])
        self.font_combo.setCurrentText(d["font_sel"])
        self.font_custom.setText(d["font_custom"])

    def accept(self):
        self.qs.setValue("theme", self.cmb_theme.currentText())
        set_bool_setting(self.qs, "open_after", self.chk_open_after.isChecked())
        self.qs.setValue("default_dir", self.edt_default_dir.text().strip())
        self.qs.setValue("font_sel", self.font_combo.currentText().strip())
        self.qs.setValue("font_custom", self.font_custom.text().strip())
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

        tip = QLabel(
            "在此维护“班级 → 院系”的映射。左侧快捷按钮切换院系，右侧编辑精确班级与关键字；每行一个词条，可用中文/英文逗号分隔。",
        )
        tip.setWordWrap(True)
        layout.addWidget(tip)

        body = QFrame(self)
        body_layout = QHBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(12)

        left_col = QVBoxLayout()
        left_col.setSpacing(8)
        left_col.addWidget(QLabel("院系快捷：点击切换，右键删除"))

        btn_scroll = QScrollArea()
        btn_scroll.setObjectName("DeptButtonScroll")
        btn_scroll.setWidgetResizable(True)
        btn_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        btn_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        btn_container = QWidget()
        btn_container.setObjectName("DeptButtonArea")
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

        right_col = QVBoxLayout()
        right_col.setSpacing(8)
        right_col.addWidget(QLabel("词典编辑"))

        tabs = QTabWidget(self)
        self.tabs = tabs
        self.edits: Dict[str, Tuple[QPlainTextEdit, QPlainTextEdit]] = {}
        for dept in sorted(self._data.keys()):
            self._add_department_tab(dept, self._data.get(dept, {}))
        self._refresh_dept_buttons()
        self.tabs.currentChanged.connect(lambda _: self._sync_dept_buttons())

        right_col.addWidget(tabs, 1)
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
            btn.customContextMenuRequested.connect(
                lambda pos, d=dept, b=btn: self._on_dept_btn_context_menu(pos, d, b)
            )
            self.dept_btn_layout.addWidget(btn)
            self.dept_btns[dept] = btn
        self.dept_btn_layout.addStretch(1)
        self._sync_dept_buttons()

    def _sync_dept_buttons(self):
        current = self.tabs.tabText(self.tabs.currentIndex()) if self.tabs.count() else None
        for dept, btn in list(self.dept_btns.items()):
            if btn.isCheckable():
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
        if not dept:
            return
        if dept in self.edits:
            return
        page = QWidget()
        form = QFormLayout(page)
        form.setLabelAlignment(Qt.AlignRight)
        txt_exact = QPlainTextEdit()
        txt_exact.setPlaceholderText("精确匹配：每行一个班级名称。")
        txt_kw = QPlainTextEdit()
        txt_kw.setPlaceholderText("关键字匹配：每行一个关键字，自动按长度优先。")
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
            ret = QMessageBox.question(
                self,
                "确认删除",
                f"确定删除院系“{dept}”及其全部班级/关键字吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if ret == QMessageBox.Yes:
                self._remove_department(dept)

    def _remove_department(self, dept: str):
        if not dept:
            return
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
        name = normalize_plain_text(self.new_dept_input.text())
        if not name:
            QMessageBox.information(self, "提示", "请输入院系名称后再新增。")
            return
        if name in self.edits:
            QMessageBox.information(self, "提示", "该院系已存在，无需重复添加。")
            for i in range(self.tabs.count()):
                if self.tabs.tabText(i) == name:
                    self.tabs.setCurrentIndex(i)
                    break
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
                norm = normalize_plain_text(piece)
                if norm:
                    parts.append(norm)
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


# ---------------- GUI（主窗口，页面内包含各自的设置项） ---------------- #
class MainWindow(QMainWindow):
    ORG = "DormHealth"
    APP = "LHExporter"

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"普查数据清洗软件  v{__version__}")
        self.setMinimumSize(1100, 760)
        self.qs = QSettings(self.ORG, self.APP)
        self.dept_dictionary = read_dictionary_setting(self.qs)
        self.settings_dialog: SettingsDialog | None = None
        self.last_progress_logged: int = 0

        self.nav_buttons = []

        header = QFrame(self); header.setObjectName("TopPanel")
        header.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(24, 10, 24, 6)
        header_layout.setSpacing(10)

        hdr_title = QLabel("普查数据清洗软件", self)
        hdr_title.setObjectName("HeaderTitle")
        hdr_hint = QLabel(
            "做这个软件的目的，就是为了摸鱼，其次，赠与恩师使用，如果有BUG，请联系作者本人，电话：15085629880，微信号：x57531，邮箱：891368650@qq.com，restoredefaults：恢复默认设置/OK：确认更改设置/cancel：取消更改。",
            self
        )
        hdr_hint.setObjectName("HeaderHint")
        hdr_hint.setWordWrap(True)
        header_layout.addWidget(hdr_title, 0, Qt.AlignLeft)
        header_layout.addWidget(hdr_hint, 0, Qt.AlignLeft)
        header_layout.addStretch(1)

        workspace_shell = QFrame(self)
        workspace_shell.setObjectName("WorkspaceShell")
        self.workspace_shell = workspace_shell
        workspace_layout_outer = QVBoxLayout(workspace_shell)
        workspace_layout_outer.setContentsMargins(20, 0, 20, 0)
        workspace_layout_outer.setSpacing(0)

        self.overlay_container = QFrame(workspace_shell)
        self.overlay_container.setObjectName("OverlayContainer")
        self.overlay_container.hide()
        overlay_layout = QVBoxLayout(self.overlay_container)
        overlay_layout.setContentsMargins(0, 0, 0, 0)
        overlay_layout.setSpacing(0)

        overlay_mask = QFrame(self.overlay_container)
        overlay_mask.setObjectName("OverlayMask")
        mask_layout = QVBoxLayout(overlay_mask)
        mask_layout.setContentsMargins(48, 40, 48, 40)
        mask_layout.setSpacing(12)

        overlay_card = QFrame(overlay_mask); overlay_card.setObjectName("OverlayCard")
        overlay_card.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.MinimumExpanding)
        self.overlay_card = overlay_card
        card_layout = QVBoxLayout(overlay_card)
        card_layout.setContentsMargins(20, 16, 20, 16)
        card_layout.setSpacing(12)

        overlay_header = QHBoxLayout(); overlay_header.setSpacing(8)
        overlay_title = QLabel("原始数据输入面板", self); overlay_title.setObjectName("HeaderTitle")
        overlay_tip = QLabel(
            "支持 分表导出/楼栋遍历与排除 共用原始Excel，以及整改/简报初稿文件，路径相互独立，可滚动查看完整输入项。",
            self
        )
        overlay_tip.setObjectName("HeaderHint"); overlay_tip.setWordWrap(True)
        overlay_header.addWidget(overlay_title)
        overlay_header.addStretch(1)
        btn_close_overlay = QPushButton("收起", self)
        btn_close_overlay.clicked.connect(self.toggle_overlay)
        overlay_header.addWidget(btn_close_overlay)

        card_layout.addLayout(overlay_header)
        card_layout.addWidget(overlay_tip)

        scroll = QScrollArea(self.overlay_container)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setObjectName("OverlayScroll")
        scroll_content = QWidget()
        scroll_layout = QFormLayout(scroll_content)
        scroll_layout.setLabelAlignment(Qt.AlignRight)
        scroll_layout.setContentsMargins(4, 4, 4, 4)
        scroll_layout.setSpacing(10)

        sp_icon = getattr(QStyle, "SP_DirOpenIcon", None) or getattr(QStyle, "SP_DirIcon", None) or getattr(
            QStyle, "SP_DialogOpenButton", None)

        self.input_edit = DropLineEdit(self)
        self.input_edit.setPlaceholderText("原始Excel（分表导出/楼栋遍历与排除 共用，.xlsx / .xls）")
        btn_in_side = QPushButton("浏览…", self)
        if sp_icon is not None:
            btn_in_side.setIcon(self.style().standardIcon(sp_icon))
        btn_in_side.clicked.connect(self.choose_input)
        row_input_side = QHBoxLayout(); row_input_side.setSpacing(8)
        row_input_side.addWidget(self.input_edit, 1)
        row_input_side.addWidget(btn_in_side)
        scroll_layout.addRow("原始数据：", row_input_side)

        self.rect_input_edit = DropLineEdit(self)
        self.rect_input_edit.setPlaceholderText("整改清单（Excel / Word / PDF）")
        btn_rect = QPushButton("浏览…", self)
        if sp_icon is not None:
            btn_rect.setIcon(self.style().standardIcon(sp_icon))
        btn_rect.clicked.connect(self.choose_rectify_input)
        row_rect = QHBoxLayout(); row_rect.setSpacing(8)
        row_rect.addWidget(self.rect_input_edit, 1)
        row_rect.addWidget(btn_rect)
        scroll_layout.addRow("整改数据：", row_rect)

        self.brief_report_edit = DropLineEdit(self)
        self.brief_report_edit.setPlaceholderText("简报初稿（Excel / Word / PDF）")
        btn_report = QPushButton("浏览…", self)
        if sp_icon is not None:
            btn_report.setIcon(self.style().standardIcon(sp_icon))
        btn_report.clicked.connect(self.choose_brief_report)
        row_report = QHBoxLayout(); row_report.setSpacing(8)
        row_report.addWidget(self.brief_report_edit, 1)
        row_report.addWidget(btn_report)
        scroll_layout.addRow("简报初稿：", row_report)

        scroll.setWidget(scroll_content)
        card_layout.addWidget(scroll, 1)

        overlay_hint = QLabel("点击“收起”即可收起浮层，下方按钮区域与工作区将重新可见。", self)
        overlay_hint.setObjectName("HeaderHint")
        overlay_hint.setWordWrap(True)
        card_layout.addWidget(overlay_hint)

        self.brief_report_display = QLineEdit(self)
        self.brief_report_display.setReadOnly(True)
        self.brief_report_display.setPlaceholderText("简报初稿路径会在此同步显示")
        self.brief_report_edit.textChanged.connect(self.brief_report_display.setText)

        mask_layout.addWidget(overlay_card, 0, Qt.AlignHCenter | Qt.AlignTop)
        mask_layout.addStretch(1)
        overlay_layout.addWidget(overlay_mask)

        workspace = QFrame(self); workspace.setObjectName("Workspace")
        workspace_layout = QHBoxLayout(workspace)
        workspace_layout.setContentsMargins(0, 14, 0, 14)
        workspace_layout.setSpacing(16)

        workspace_layout_outer.addWidget(workspace)
        self.tabs = QTabWidget(self); self.tabs.setObjectName("WorkTabs"); self.tabs.tabBar().hide()

        tab1 = QWidget(self); self.tabs.addTab(tab1, "分表导出")
        t1_layout = QVBoxLayout(tab1); t1_layout.setContentsMargins(14, 14, 14, 14); t1_layout.setSpacing(12)

        card1 = QFrame(self); card1.setObjectName("Card")
        c1 = QVBoxLayout(card1); c1.setContentsMargins(16, 16, 16, 16); c1.setSpacing(12)
        c1.addWidget(self._build_section_header("分表导出", "按楼栋拆分、序号公式、0 分标红与目录统计（与 楼栋遍历与排除 共用原始数据）"))
        form1 = QFormLayout(); form1.setLabelAlignment(Qt.AlignRight)

        self.spin_max = QSpinBox(self); self.spin_max.setRange(1, 50)
        self.chk_mark_num_zero = QCheckBox("标红：数字 0", self)
        self.chk_mark_text_zero = QCheckBox('标红：文本 "0"', self)
        self.chk_mark_text_zero_dot = QCheckBox('标红：文本 "0.0"', self)
        self.chk_use_majority_dept_ui1 = QCheckBox("寝室院系按人数占比判定（平局取首个院系）", self)
        self.output_edit1 = QLineEdit(self); self.output_edit1.setPlaceholderText("保存为 .xlsx …")
        btn_out1 = QPushButton("保存到…", self)
        sp_save = getattr(QStyle, "SP_DialogSaveButton", None) or getattr(QStyle, "SP_DriveHDIcon", None)
        if sp_save is not None:
            btn_out1.setIcon(self.style().standardIcon(sp_save))
        btn_out1.clicked.connect(self.choose_output1)
        row1 = QHBoxLayout(); row1.addWidget(self.output_edit1); row1.addWidget(btn_out1)

        mark_row = QHBoxLayout(); mark_row.setSpacing(10)
        for cb in (self.chk_mark_num_zero, self.chk_mark_text_zero, self.chk_mark_text_zero_dot):
            mark_row.addWidget(cb)
        mark_row.addStretch(1)
        form1.addRow("最多 x苑x号 工作表数：", self.spin_max)
        form1.addRow("标红选项：", mark_row)
        form1.addRow("院系判定：", self.chk_use_majority_dept_ui1)
        form1.addRow("输出文件：", row1)
        c1.addLayout(form1)

        self.btn_run1 = QPushButton("导出", self); self.btn_run1.setObjectName("PrimaryButton")
        self.btn_run1.setMinimumHeight(44)
        self.btn_run1.clicked.connect(self.start_run1)
        c1.addWidget(self.btn_run1)

        tip1 = QLabel("广告位")
        tip1.setObjectName("HelperText"); tip1.setWordWrap(True)
        c1.addWidget(tip1)
        t1_layout.addWidget(card1)

        tab2 = QWidget(self); self.tabs.addTab(tab2, "楼栋遍历与排除")
        t2_layout = QVBoxLayout(tab2); t2_layout.setContentsMargins(14, 14, 14, 14); t2_layout.setSpacing(12)

        card2 = QFrame(self); card2.setObjectName("Card")
        c2 = QVBoxLayout(card2); c2.setContentsMargins(16, 16, 16, 16); c2.setSpacing(12)
        c2.addWidget(self._build_section_header("楼栋遍历与排除", "区间/单间排除、去零开关、院系判定与八大系部统计（与 分表导出 共用原始数据）"))
        form2 = QFormLayout(); form2.setLabelAlignment(Qt.AlignRight)

        self.exclude_buildings = list(range(1, 4))
        self.chk_excl_enable_ui2 = QCheckBox("启用“寝室区间排除”")
        self.chk_excl_enable_ui2.stateChanged.connect(lambda _: self._update_exclusion_summary())
        self.exclusion_cfg: Dict[str, Dict[int, Dict[str, str]]] = {"lan": {}, "mei": {}}

        self.btn_config_exclusion = QPushButton("配置区间/单间…", self)
        self.btn_config_exclusion.clicked.connect(self.open_exclusion_dialog)
        self.lbl_exclusion_summary = QLabel("未配置")
        self.lbl_exclusion_summary.setObjectName("HeaderHint")

        ex_head_row = QHBoxLayout(); ex_head_row.setSpacing(10)
        ex_head_row.addWidget(self.chk_excl_enable_ui2)
        ex_head_row.addWidget(self.btn_config_exclusion)
        ex_head_row.addWidget(self.lbl_exclusion_summary, 1)

        self.chk_use_majority_dept = QCheckBox("寝室院系按人数占比判定（平局取首个院系）")
        self.chk_drop_zero_text_ui2 = QCheckBox("排除文本 0.0")
        self.chk_drop_zero_numeric_ui2 = QCheckBox("排除数值 0")
        self.chk_drop_zero_text_ui2.setChecked(True)
        self.chk_drop_zero_numeric_ui2.setChecked(True)

        self.output_edit2 = QLineEdit(self); self.output_edit2.setPlaceholderText("保存为 .xlsx …")
        btn_out2 = QPushButton("保存到…", self)
        if sp_save is not None:
            btn_out2.setIcon(self.style().standardIcon(sp_save))
        btn_out2.clicked.connect(self.choose_output2)
        row2 = QHBoxLayout(); row2.addWidget(self.output_edit2); row2.addWidget(btn_out2)

        self.lbl_info_i2 = QLabel(
            "楼栋遍历与排除使用的原始数据即为左侧“原始数据输入”中选择的Excel；区间/单间排除采用弹窗填写（兰/梅苑 1-3 栋），可选启用人数占比院系判定，并分开排除文本 0.0 / 数值 0 分。")
        self.lbl_info_i2.setWordWrap(True)

        c2.addLayout(ex_head_row)

        opt_row = QHBoxLayout(); opt_row.setSpacing(12)
        for cb in (self.chk_use_majority_dept, self.chk_drop_zero_text_ui2, self.chk_drop_zero_numeric_ui2):
            opt_row.addWidget(cb)
        opt_row.addStretch(1)
        form2.addRow("统计选项：", opt_row)
        form2.addRow("输出文件：", row2)

        c2.addLayout(form2)
        c2.addWidget(self.lbl_info_i2)

        self.btn_run2 = QPushButton("导出", self); self.btn_run2.setObjectName("PrimaryButton")
        self.btn_run2.setMinimumHeight(44)
        self.btn_run2.clicked.connect(self.start_run2)
        c2.addWidget(self.btn_run2)

        tip2 = QLabel("广告位")
        tip2.setObjectName("HelperText"); tip2.setWordWrap(True)
        c2.addWidget(tip2)

        t2_layout.addWidget(card2)

        tab3 = QWidget(self); self.tabs.addTab(tab3, "学风简报中心")
        t3_layout = QVBoxLayout(tab3); t3_layout.setContentsMargins(14, 14, 14, 14); t3_layout.setSpacing(12)

        card3 = QFrame(self); card3.setObjectName("Card")
        c3 = QVBoxLayout(card3); c3.setContentsMargins(16, 16, 16, 16); c3.setSpacing(12)
        c3.addWidget(self._build_section_header("简报中心 · 整改删行", "原始简报与整改清单统一从顶栏输入，支持院系严格匹配开关"))
        form3 = QFormLayout(); form3.setLabelAlignment(Qt.AlignRight)

        btn_report2 = QPushButton("浏览…"); btn_report2.clicked.connect(self.choose_brief_report)
        row_report2 = QHBoxLayout()
        row_report2.addWidget(self.brief_report_display, 1)
        row_report2.addWidget(btn_report2)

        self.chk_match_room_only = QCheckBox('仅按“楼栋+宿舍号”匹配删除')
        self.chk_require_dept = QCheckBox('需要院系同时匹配（更严格）')

        self.brief_output_edit = QLineEdit(); self.brief_output_edit.setPlaceholderText("保存输出 Excel…")
        btn_out3 = QPushButton("保存到…"); btn_out3.clicked.connect(self.choose_brief_output)
        row_out3 = QHBoxLayout(); row_out3.addWidget(self.brief_output_edit, 1); row_out3.addWidget(btn_out3)

        form3.addRow("简报初稿：", row_report2)
        brief_opts = QHBoxLayout(); brief_opts.setSpacing(12)
        brief_opts.addWidget(self.chk_match_room_only)
        brief_opts.addWidget(self.chk_require_dept)
        brief_opts.addStretch(1)
        form3.addRow("匹配口径：", brief_opts)
        form3.addRow("输出文件：", row_out3)
        c3.addLayout(form3)

        self.btn_run3 = QPushButton("执行删行", self); self.btn_run3.setObjectName("PrimaryButton")
        self.btn_run3.setMinimumHeight(44)
        self.btn_run3.clicked.connect(self.start_run3)
        c3.addWidget(self.btn_run3)

        tip3 = QLabel("广告位")
        tip3.setObjectName("HelperText"); tip3.setWordWrap(True)
        c3.addWidget(tip3)

        t3_layout.addWidget(card3)

        self.tabs.currentChanged.connect(self._sync_nav_buttons)

        nav_panel = QFrame(self); nav_panel.setObjectName("NavPanel"); nav_panel.setMinimumWidth(220)
        nav_layout = QVBoxLayout(nav_panel)
        nav_layout.setContentsMargins(16, 16, 16, 16)
        nav_layout.setSpacing(12)

        raw_card = QFrame(self); raw_card.setObjectName("NavCard")
        raw_lay = QVBoxLayout(raw_card); raw_lay.setContentsMargins(12, 12, 12, 12); raw_lay.setSpacing(6)
        self.raw_data_btn = QPushButton("原始数据输入", self)
        self.raw_data_btn.setObjectName("PrimaryButton")
        self.raw_data_btn.setCheckable(True)
        self.raw_data_btn.clicked.connect(self.toggle_overlay)
        raw_lay.addWidget(self.raw_data_btn)

        nav_layout.addWidget(raw_card)

        nav_layout.addWidget(self._create_nav_entry("分表导出", "", 0))
        nav_layout.addWidget(self._create_nav_entry("楼栋遍历与排除", "", 1))
        nav_layout.addWidget(self._create_nav_entry("简报中心", "", 2))
        nav_layout.addStretch(1)

        dict_btn = QPushButton("院系词典", self); dict_btn.clicked.connect(self.open_dictionary_dialog)
        nav_layout.addWidget(dict_btn)

        log_btn = QPushButton("运行日志", self); log_btn.clicked.connect(self.open_logs_dialog)
        nav_layout.addWidget(log_btn)

        settings_btn = QPushButton("设置", self); settings_btn.clicked.connect(self.open_settings)
        nav_layout.addWidget(settings_btn)

        workspace_layout.addWidget(nav_panel, 3)
        workspace_layout.addWidget(self.tabs, 10)
        workspace_layout.setStretch(0, 3)
        workspace_layout.setStretch(1, 10)

        bottom = QFrame(self); bottom.setObjectName("BottomBar")
        bottom_layout = QHBoxLayout(bottom)
        bottom_layout.setContentsMargins(16, 10, 16, 10)
        bottom_layout.setSpacing(8)
        self.progress = QProgressBar(self); self.progress.setValue(0); self.progress.setTextVisible(True)
        bottom_layout.addWidget(self.progress)

        st = self.statusBar()
        ver_lbl = QLabel(f"版本：{__version__}", self)
        st.addPermanentWidget(ver_lbl)

        central = QWidget(self)
        root = QVBoxLayout(central); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)
        root.addWidget(header)
        root.addWidget(workspace_shell, 1)
        root.setStretch(0, 1)
        root.setStretch(1, 5)
        root.addWidget(bottom)
        self.setCentralWidget(central)

        self.worker: QThread | None = None

        self.load_basic_settings()
        self.apply_theme()
        self._sync_nav_buttons(self.tabs.currentIndex())

    def _qss(self, dark: bool) -> str:
        if not dark:
            accent = "#007AFF"; accent_hover = "#0A84FF"; accent_press = "#0051C6"
            bg = "#F5F5F7"; card = "#FFFFFF"; border = "#E5E5EA"
            text_muted = "#6C6C70"; text = "#1C1C1E"; header_soft = "rgba(255,255,255,0.7)"
        else:
            accent = "#0A84FF"; accent_hover = "#3A7CFF"; accent_press = "#246BFF"
            bg = "#000000"; card = "#1C1C1E"; border = "#2C2C2E"
            text_muted = "#8E8E93"; text = "#F2F2F7"; header_soft = "rgba(28,28,30,0.9)"
        return f"""
        QMainWindow {{ background: {bg}; color:{text}; }}
        #TopPanel {{ background: {card}; border-bottom: 1px solid {border}; }}
        #HeaderTitle {{ font-size: 17px; font-weight: 600; margin-bottom: 4px; }}
        #HeaderCaption {{ color: {text_muted}; font-size: 13px; margin-top: 4px; }}
        #HeaderHint {{ color: {text_muted}; font-size: 12px; }}
        #HeaderDivider {{ border: none; background: {border}; height: 1px; }}

        #WorkspaceShell {{ background: transparent; position: relative; }}
        #OverlayContainer {{ background: transparent; }}
        #OverlayMask {{ background: rgba(0,0,0,0.35); }}
        #OverlayCard {{ background: {card}; border: 1px solid {border}; border-radius: 18px; box-shadow: 0 18px 48px rgba(0,0,0,0.18); }}
        #OverlayScroll {{ border: 0px; }}
        #NavPanel {{ background: {card}; border: 1px solid {border}; border-radius: 18px; }}
        #NavTitle {{ font-size: 15px; font-weight: 600; }}
        #NavHint {{ color: {text_muted}; font-size: 12px; }}
        #NavCard {{ background: transparent; border: 0px; border-radius: 0px; }}
        QPushButton#NavButton {{
            height: 40px; border-radius: 12px; font-weight: 600;
            background: transparent; border: 1px solid {border}; color:{text};
        }}
        QPushButton#NavButton:hover {{ border-color: {accent_hover}; color:{accent_hover}; }}
        QPushButton#NavButton:checked {{ background: {card}; color:{accent}; border-color: {accent}; }}

        QPushButton#DeptButton {{
            min-width: 120px; padding: 6px 12px; border-radius: 12px;
            background: transparent; border: 1px solid {border}; color:{text};
        }}
        QPushButton#DeptButton:hover {{ border-color: {accent_hover}; color:{accent_hover}; }}
        QPushButton#DeptButton:checked {{ background: transparent; color:{accent}; border-color: {accent}; box-shadow: none; }}
        QPushButton#DeptButton:pressed {{ background: transparent; border-color: {accent_press}; }}

        #DeptButtonArea {{ background: transparent; border: 0px; }}

        QTabWidget#WorkTabs::pane {{ border: 0px; }}
        QTabBar::tab {{
            padding: 0px; margin: 0px; border: none; height: 0px; width: 0px;
        }}

        #Card {{ background: {card}; border: 1px solid {border}; border-radius: 18px; }}

        QPushButton {{
            height: 36px; padding: 6px 18px; border-radius: 12px;
            background: {card}; border: 1px solid {border}; color:{text};
        }}
        QPushButton:hover {{ background: {'#2C2C2E' if dark else '#E5E5EA'}; }}
        QPushButton:pressed {{ background: {'#3A3A3C' if dark else '#D1D1D6'}; }}

        QPushButton#PrimaryButton {{
            background: {accent}; color: white; border: 1px solid {accent}; font-weight: 600;
        }}
        QPushButton#PrimaryButton:hover {{ background: {accent_hover}; border-color: {accent_hover}; }}
        QPushButton#PrimaryButton:pressed {{ background: {accent_press}; border-color: {accent_press}; }}

        #SectionHeader {{ background: transparent; }}
        #SectionTitle {{ font-size: 15px; font-weight: 600; }}
        #SectionSub {{ color: {text_muted}; font-size: 12px; }}

        QLineEdit, QSpinBox, QComboBox {{
            height: 34px; padding: 6px 10px; border-radius: 12px;
            border: 1px solid {border}; background: {card}; color:{text};
        }}
        QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{ border: 1px solid {accent}; }}
        QProgressBar {{
            border: 1px solid {border}; border-radius: 8px; text-align: center; height: 20px; color:{text};
        }}
        QProgressBar::chunk {{ background-color: {accent}; border-radius: 8px; }}
        #BottomBar {{ background: {header_soft}; border-top: 1px solid {border}; }}
        QLabel#HelperText {{ color: {text_muted}; font-size: 13px; }}
        """

    def _create_nav_entry(self, title: str, subtitle: str, tab_index: int) -> QFrame:
        card = QFrame(self); card.setObjectName("NavCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)

        btn = QPushButton(title, self)
        btn.setObjectName("NavButton")
        btn.setCheckable(True)
        btn.clicked.connect(lambda _, idx=tab_index: self.tabs.setCurrentIndex(idx))
        layout.addWidget(btn)

        if subtitle:
            lab = QLabel(subtitle, self)
            lab.setObjectName("NavHint")
            lab.setWordWrap(True)
            layout.addWidget(lab)

        self.nav_buttons.append((tab_index, btn))
        return card

    def _build_section_header(self, title: str, subtitle: str | None = None) -> QFrame:
        box = QFrame(self); box.setObjectName("SectionHeader")
        lay = QVBoxLayout(box); lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(4)

        title_lab = QLabel(title, self); title_lab.setObjectName("SectionTitle")
        lay.addWidget(title_lab)

        if subtitle:
            sub_lab = QLabel(subtitle, self)
            sub_lab.setObjectName("SectionSub")
            sub_lab.setWordWrap(True)
            lay.addWidget(sub_lab)

        return box

    def _sync_nav_buttons(self, index: int):
        for idx, btn in self.nav_buttons:
            blocked = btn.blockSignals(True)
            btn.setChecked(idx == index)
            btn.blockSignals(blocked)

    def _resize_overlay_card(self):
        if not hasattr(self, "overlay_card"):
            return
        width = int(self.workspace_shell.width() * 0.625)
        height = int(self.workspace_shell.height() * 0.65)
        width = max(760, width)
        height = max(520, height)
        self.overlay_card.setMinimumSize(width, height)
        self.overlay_card.setMaximumWidth(int(self.workspace_shell.width() * 0.9))

    def toggle_overlay(self):
        visible = not self.overlay_container.isVisible()
        self.overlay_container.setVisible(visible)
        if hasattr(self, "raw_data_btn") and self.raw_data_btn:
            self.raw_data_btn.setChecked(visible)
        if visible:
            self.overlay_container.raise_()
            self.overlay_container.setGeometry(self.workspace_shell.rect())
            self._resize_overlay_card()

    def _show_overlay(self):
        if not self.overlay_container.isVisible():
            self.toggle_overlay()
        else:
            self.overlay_container.raise_()
            self.overlay_container.setGeometry(self.workspace_shell.rect())
            self._resize_overlay_card()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "overlay_container") and self.overlay_container.isVisible():
            self.overlay_container.setGeometry(self.workspace_shell.rect())
            self._resize_overlay_card()

    def apply_theme(self):
        theme = self.qs.value("theme", "浅色")
        dark = (theme == "深色")
        self.setStyleSheet(self._qss(dark))

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
            self.lbl_exclusion_summary.setText("未配置具体区间/单间")
        else:
            self.lbl_exclusion_summary.setText(f"兰苑{lan_cnt}栋，梅苑{mei_cnt}栋")

    def open_exclusion_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("区间/单间排除设置")
        dlg.setMinimumSize(840, 560)
        lay = QVBoxLayout(dlg); lay.setContentsMargins(16, 16, 16, 16); lay.setSpacing(12)
        tip = QLabel("按楼栋填写区间与单间，留空表示不排除；兰/梅苑均提供 1-3 栋输入，并支持滚动查看。")
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
            card = QFrame(); card.setObjectName("ExGardenCard")
            card_lay = QVBoxLayout(card); card_lay.setContentsMargins(10, 10, 10, 10); card_lay.setSpacing(8)
            caption = QLabel(title); caption.setObjectName("HeaderCaption")
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
                rng.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                sgl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
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

    def load_basic_settings(self):
        self.input_edit.setText(self.qs.value("input_path", ""))
        self.rect_input_edit.setText(self.qs.value("rectify/path", ""))
        self.output_edit1.setText(self.qs.value("ui1/out", ""))
        self.output_edit2.setText(self.qs.value("ui2/out", ""))
        self.brief_report_edit.setText(self.qs.value("brief/report", ""))
        self.brief_output_edit.setText(self.qs.value("brief/out", ""))

        self.spin_max.setValue(int(self.qs.value("ui1/max", 6)))
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
        self.chk_match_room_only.setChecked(get_bool_setting(self.qs, "brief/match_room_only", True))
        self.chk_require_dept.setChecked(get_bool_setting(self.qs, "brief/require_dept", False))

    def save_basic_settings(self):
        self.qs.setValue("input_path", self.input_edit.text().strip())
        self.qs.setValue("rectify/path", self.rect_input_edit.text().strip())
        self.qs.setValue("ui1/out", self.output_edit1.text().strip())
        self.qs.setValue("ui2/out", self.output_edit2.text().strip())
        self.qs.setValue("brief/report", self.brief_report_edit.text().strip())
        self.qs.setValue("brief/out", self.brief_output_edit.text().strip())

        self.qs.setValue("ui1/max", self.spin_max.value())
        set_bool_setting(self.qs, "ui1/mark_num", self.chk_mark_num_zero.isChecked())
        set_bool_setting(self.qs, "ui1/mark_txt0", self.chk_mark_text_zero.isChecked())
        set_bool_setting(self.qs, "ui1/mark_txt0dot", self.chk_mark_text_zero_dot.isChecked())
        set_bool_setting(self.qs, "ui1/majority_dept", self.chk_use_majority_dept_ui1.isChecked())

        set_bool_setting(self.qs, "ui2/ex/enabled", self.chk_excl_enable_ui2.isChecked())
        self._save_exclusion_cfg()
        set_bool_setting(self.qs, "ui2/majority_dept", self.chk_use_majority_dept.isChecked())
        set_bool_setting(self.qs, "ui2/drop_zero_text", self.chk_drop_zero_text_ui2.isChecked())
        set_bool_setting(self.qs, "ui2/drop_zero_numeric", self.chk_drop_zero_numeric_ui2.isChecked())
        set_bool_setting(self.qs, "brief/match_room_only", self.chk_match_room_only.isChecked())
        set_bool_setting(self.qs, "brief/require_dept", self.chk_require_dept.isChecked())

    def reset_settings(self):
        self.qs.clear()
        self.dept_dictionary = read_dictionary_setting(self.qs)
        self.load_basic_settings()
        self.apply_theme()
        QMessageBox.information(self, "提示", "已重置设置（不影响已导出的文件）。")

    def open_settings(self, show_logs: bool = False):
        dlg = SettingsDialog(self, self.qs, show_logs=show_logs)
        self.settings_dialog = dlg
        dlg.finished.connect(lambda _: setattr(self, "settings_dialog", None))
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

    def _log_runtime(self, message: str):
        append_runtime_log(self.qs, message)

    def _log_progress(self, value: int):
        if value >= 100 or value - self.last_progress_logged >= 20:
            self._log_runtime(f"当前进度：{value}%")
            self.last_progress_logged = value

    def on_progress(self, value: int):
        self.progress.setValue(value)
        self._log_progress(value)

    def _default_output_dir(self, input_path: str) -> str:
        d = self.qs.value("default_dir", "").strip()
        if d and os.path.isdir(d):
            return d
        return os.path.dirname(input_path) if input_path else os.getcwd()

    def choose_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择原始 Excel（界面一 / 界面二 共用）",
            "",
            "Excel 文件 (*.xlsx *.xls)"
        )
        if path:
            self.input_edit.setText(path)
            base = os.path.splitext(os.path.basename(path))[0]
            out_dir = self._default_output_dir(path)
            if not self.output_edit1.text():
                self.output_edit1.setText(os.path.join(out_dir, f"{base}_表一.xlsx"))
            if not self.output_edit2.text():
                self.output_edit2.setText(os.path.join(out_dir, f"{base}_表二.xlsx"))
            self.save_basic_settings()

    def choose_rectify_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择整改清单",
            self.rect_input_edit.text().strip(),
            "支持 Excel/Word/PDF (*.xlsx *.xls *.doc *.docx *.pdf)"
        )
        if path:
            self.rect_input_edit.setText(path)
            self.save_basic_settings()

    def choose_output1(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存 分表导出 Excel", self.output_edit1.text().strip(),
                                              "Excel 文件 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_edit1.setText(path)
            self.save_basic_settings()

    def choose_output2(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存 楼栋遍历与排除 Excel", self.output_edit2.text().strip(),
                                              "Excel 文件 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_edit2.setText(path)
            self.save_basic_settings()

    def choose_brief_report(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择简报初稿",
            self.brief_report_edit.text().strip(),
            "支持 Excel/Word/PDF (*.xlsx *.xls *.doc *.docx *.pdf)"
        )
        if path:
            self.brief_report_edit.setText(path)
            if not self.brief_output_edit.text():
                base = os.path.splitext(os.path.basename(path))[0]
                out_dir = self._default_output_dir(path)
                self.brief_output_edit.setText(os.path.join(out_dir, f"{base}_删行后.xlsx"))
            self.save_basic_settings()

    def choose_brief_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存删行简报", self.brief_output_edit.text().strip(),
                                              "Excel 文件 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.brief_output_edit.setText(path)
            self.save_basic_settings()

    def _excel_font_name(self) -> str:
        custom = str(self.qs.value("font_custom", "")).strip()
        base = str(self.qs.value("font_sel", "仿宋_GB2312") or "仿宋_GB2312").strip()
        return custom or base or "仿宋_GB2312"

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

        return {
            "enabled": enabled,
            "lan": collect("lan"),
            "mei": collect("mei"),
        }

    def start_run1(self):
        in_path = self.input_edit.text().strip()
        out_path = self.output_edit1.text().strip()
        if not in_path or not os.path.isfile(in_path):
            QMessageBox.warning(self, "提示", "请先在“原始数据输入”中选择有效的原始 Excel 文件（.xlsx/.xls）。")
            self._show_overlay()
            self.input_edit.setFocus()
            return
        if not out_path:
            QMessageBox.warning(self, "提示", "请指定 分表导出 输出文件路径（.xlsx）。")
            return

        s = SettingsRun(
            input_path=in_path,
            output_path=out_path,
            max_sheets=int(self.spin_max.value()),
            font_name=self._excel_font_name(),
            mark_num_zero=bool(self.chk_mark_num_zero.isChecked()),
            mark_text_zero=bool(self.chk_mark_text_zero.isChecked()),
            mark_text_zero_dot=bool(self.chk_mark_text_zero_dot.isChecked()),
            dept_dictionary=deepcopy(self.dept_dictionary),
            use_majority_dept=bool(self.chk_use_majority_dept_ui1.isChecked()),
        )

        self.btn_run1.setEnabled(False)
        self.btn_run2.setEnabled(False)
        self.btn_run3.setEnabled(False)
        self.progress.setValue(5)
        self.last_progress_logged = 0
        self._log_runtime(f"开始生成分表导出数据：{os.path.basename(out_path)}")

        self.worker = Worker(s)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def start_run2(self):
        in_path = self.input_edit.text().strip()
        out_path = self.output_edit2.text().strip()
        if not in_path or not os.path.isfile(in_path):
            QMessageBox.warning(self, "提示", "请先在“原始数据输入”中选择有效的原始 Excel 文件（.xlsx/.xls）。")
            self._show_overlay()
            self.input_edit.setFocus()
            return
        if not out_path:
            QMessageBox.warning(self, "提示", "请指定 楼栋遍历与排除 输出文件路径（.xlsx）。")
            return

        self.btn_run1.setEnabled(False)
        self.btn_run2.setEnabled(False)
        self.btn_run3.setEnabled(False)
        self.progress.setValue(5)
        self.last_progress_logged = 0
        self._log_runtime(f"开始生成楼栋遍历与排除数据：{os.path.basename(out_path)}")

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
        report_path = self.brief_report_edit.text().strip()
        rect_path = self.rect_input_edit.text().strip()
        output_path = self.brief_output_edit.text().strip()
        if not report_path or not os.path.isfile(report_path):
            QMessageBox.warning(self, "提示", "请选择有效的简报初稿文件（Excel/Word/PDF）。")
            return
        if not rect_path or not os.path.isfile(rect_path):
            QMessageBox.warning(self, "提示", "请在顶部“整改数据”栏选择有效的整改清单（Excel/Word/PDF）。")
            return
        if not output_path:
            QMessageBox.warning(self, "提示", "请指定输出文件路径（.xlsx）。")
            return

        self.btn_run1.setEnabled(False)
        self.btn_run2.setEnabled(False)
        self.btn_run3.setEnabled(False)
        self.progress.setValue(5)
        self.last_progress_logged = 0
        self._log_runtime(f"开始生成简报中心：{os.path.basename(output_path)}")

        settings = BriefingSettings(
            report_path=report_path,
            rectified_path=rect_path,
            output_path=output_path,
            match_by_room_only=self.chk_match_room_only.isChecked(),
            require_department=self.chk_require_dept.isChecked(),
        )
        self.worker = WorkerBriefing(settings)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def show_about(self):
        QMessageBox.information(
            self, "关于本程序",
            f"版本：{__version__}\n\n说明：{__build_note__}"
        )

    def _open_folder(self, path: str):
        try:
            folder = os.path.dirname(path)
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception:
            pass

    def on_finished(self, out_path: str):
        self.btn_run1.setEnabled(True)
        self.btn_run2.setEnabled(True)
        self.btn_run3.setEnabled(True)
        self.progress.setValue(100)
        self.save_basic_settings()
        summary_msg = None

        if isinstance(self.worker, Worker):
            meta = getattr(self.worker, "summary", {}) or {}
            summary_msg = (
                f"分表导出 完成：{os.path.basename(out_path)}；"
                f"明细表={meta.get('sheet', '')}，有效行={meta.get('valid_rows', 0)}，"
                f"楼栋表={meta.get('sheet_count', 0)}，"
                f"不及格条数={meta.get('fail_rows', 0)}，"
                f"总分为0.0条数={meta.get('zero_rows', 0)}。"
            )
        elif isinstance(self.worker, WorkerIface2):
            stat = getattr(self.worker, "summary", {}) or {}
            summary_msg = (
                f"楼栋遍历与排除 完成：{os.path.basename(out_path)}；"
                f"原始行={stat.get('rows_raw', 0)}，结构化={stat.get('rows_after_structure', 0)}，"
                f"去零后={stat.get('rows_after_zero', 0)}，排除后={stat.get('rows_after_exclusion', 0)}，"
                f"最终计数={stat.get('rows_final', 0)}。"
            )
        elif isinstance(self.worker, WorkerBriefing):
            stat = getattr(self.worker, "summary", {}) or {}
            summary_msg = (
                f"简报中心 完成：{os.path.basename(out_path)}；"
                f"初稿行={stat.get('report_rows', 0)}，整改行={stat.get('rect_rows', 0)}，"
                f"剩余行={stat.get('remaining_rows', 0)}。"
            )

        if summary_msg:
            self._log_runtime(summary_msg)
        else:
            self._log_runtime(f"已生成：{out_path}")
        if get_bool_setting(self.qs, "open_after", True):
            self._open_folder(out_path)
        QMessageBox.information(self, "完成", f"已生成：\n{out_path}\n版本：{__version__}")

    def on_error(self, msg: str):
        self.btn_run1.setEnabled(True)
        self.btn_run2.setEnabled(True)
        self.btn_run3.setEnabled(True)
        self.progress.setValue(0)
        self._log_runtime(f"处理失败：{msg}")
        QMessageBox.critical(self, "错误", f"处理失败：\n{msg}")


def main():
    os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "1")
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
