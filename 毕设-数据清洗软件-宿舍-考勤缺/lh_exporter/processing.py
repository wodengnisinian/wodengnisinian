# -*- coding: utf-8 -*-
from __future__ import annotations

"""
Core data processing logic (ported from your first monolithic script).
UI code is separated into other modules.
"""

# ===== 版本信息 =====
__version__ = "2.0.0.01"
__build_note__ = (
    "全新界面设计，采用现代化卡片式布局，提升用户体验；",
    "整合为单一数据处理脚本，保留统一清洗链；运行日志入口移至主界面左侧按钮区域，设置弹窗精简通用页；",
    "修复清洗阶段缺失列（如班级）导致统计中断的问题；",
    "界面二分组统计强化院系标量化，避免 first_dept 维度异常；",
    "运行日志增强：记录输入/输出与处理统计，界面一条件格式在空表时不再触发异常；",
    "界面二原始数据统一从“原始数据输入”弹窗获取，路径缺失时自动唤起弹窗提醒；",
    "界面一新增寝室院系多数决选项，可按人数占比统一院系判定",
    "界面一与界面二原始数据合并为同一输入文件路径，共用同一份明细数据。",
    "界面二“排除文本 0/0.0”选项调整为仅排除文本 0.0（含 0.00/0.000分 等写法），界面一总分为0明细逻辑保持不变。",
    "优化考勤检查标签页，添加卡片式框架和功能状态显示；",
    "修复进度条样式，添加边框使其更清晰；",
    "优化每个工作表最多楼栋数的输入方式，添加输入框和加减按钮；",
    "修复不存在的btn_run3属性导致的错误。"
)
__history__ = """
1.0.0.0: 初始版本发布。
。。。。。。（省略部分更新效果）
1.1.0.0: 增加了可选排除与去除总分为0的功能，完善了界面一和界面二的操作流程。
1.1.0.1: UI界面排版优化
2.0.0.01: 全新界面设计，采用现代化卡片式布局，提升用户体验；整合数据处理脚本；修复各种问题；优化界面交互。
"""

import os
import re
import io
import datetime as dt
from copy import deepcopy
from typing import Iterable, List, Dict, Tuple, Optional, Set, Any, Callable

import pandas as pd
import numpy as np

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
            # 优化：只读取前50行来检测，减少内存消耗
            t = pd.read_excel(xls_path, sheet_name=s, header=None, nrows=50)
            t = t.dropna(how="all").dropna(axis=1, how="all")
            if t.shape[1] >= 10 and t.shape[0] > best_rows:
                best_name, best_rows = s, t.shape[0]
        except Exception:
            pass
    return best_name or xls.sheet_names[0]

def clean_text(s):
    if pd.isna(s): 
        return np.nan
    return HIDDEN_WS_RE.sub("", str(s)).strip()

def normalize_plain_text(value) -> str:
    text = clean_text(value)
    if pd.isna(text):
        return ""
    return str(text)

def normalize_department(value) -> str:
    return normalize_plain_text(value)

def ensure_scalar_department(value) -> str:
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
    if pd.isna(v): 
        return np.nan
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
    """
    score_raw = df[score_col].map(normalize_plain_text)
    score_num = score_raw.map(parse_score)

    text_zero = pd.Series(False, index=df.index)
    num_zero = pd.Series(False, index=df.index)

    if drop_text_zero:
        if text_mode == "only_0dot0":
            text_zero = score_raw.map(_is_text_zero_0dot0_only)
        else:
            text_zero = score_raw.map(lambda x: bool(ZERO_TEXT_RE.match(str(x).strip())))
        if extra_text_pred is not None:
            text_zero = text_zero | score_raw.map(lambda x: bool(extra_text_pred(str(x))))

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

def match_department_by_class_name(cls_name: str, dictionary: Dict[str, Dict[str, List[str]]]) -> Optional[str]:
    """根据班级名称在“院系词典中心”中寻找院系。"""
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
CHINESE_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}

def parse_chinese_number(text: str) -> Optional[int]:
    if not text:
        return None
    if text.isdigit():
        return int(text)
    if text == "十":
        return 10
    if len(text) == 2 and text.startswith("十") and text[1] in CHINESE_NUM:
        return 10 + CHINESE_NUM[text[1]]
    if len(text) == 2 and text.endswith("十") and text[0] in CHINESE_NUM:
        return CHINESE_NUM[text[0]] * 10
    total = 0
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

def parse_interval_text(text: str) -> List[Tuple[int, int]]:
    ranges: List[Tuple[int, int]] = []
    if not text:
        return ranges
    for part in re.split(r'[,\uFF0C\u3001；;,，\s]+', text.strip()):
        if not part or "-" not in part:
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
        if part and part.isdigit():
            singles.append(int(part))
    return singles

def describe_exclusion(ex: Dict) -> str:
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
    """规范楼栋为 x苑x号（如：'梅苑1号'）。"""
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
    norm = normalize_building(value)
    if pd.isna(norm):
        return normalize_plain_text(value)
    return normalize_plain_text(norm)

def room_sort_key(v):
    s = str(v).strip()
    m = re.search(r"(\d+)", s)
    return (0, int(m.group(1)), s) if m else (1, float('inf'), s)

# ---------------- 简报中心：导入 Excel/Word/PDF ---------------- #
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

    if ext == ".csv":
        return pd.read_csv(path)

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

    raise ValueError("仅支持 Excel/Word/PDF/CSV 文件。")

def drop_rectified_rows(report_df: pd.DataFrame,
                        rectified_df: pd.DataFrame,
                        match_by_room_only: bool = True,
                        require_department: bool = False) -> pd.DataFrame:
    """按照删行口径，移除已整改寝室。使用向量化操作提高性能。"""
    include_dept = (not match_by_room_only) and require_department

    # 向量化处理：构建关键列
    report_df = report_df.copy()
    rectified_df = rectified_df.copy()

    # 清理楼栋和宿舍号
    report_df['_building'] = report_df['楼栋'].astype(str).apply(clean_building_text)
    report_df['_room'] = report_df['宿舍号'].astype(str).apply(normalize_plain_text)
    rectified_df['_building'] = rectified_df['楼栋'].astype(str).apply(clean_building_text)
    rectified_df['_room'] = rectified_df['宿舍号'].astype(str).apply(normalize_plain_text)

    if include_dept:
        report_df['_dept'] = report_df['院系'].astype(str).apply(normalize_department)
        rectified_df['_dept'] = rectified_df['院系'].astype(str).apply(normalize_department)
        # 构建复合键
        report_df['_key'] = report_df['_building'] + '|' + report_df['_room'] + '|' + report_df['_dept']
        rectified_df['_key'] = rectified_df['_building'] + '|' + rectified_df['_room'] + '|' + rectified_df['_dept']
    else:
        # 仅使用楼栋和宿舍号
        report_df['_key'] = report_df['_building'] + '|' + report_df['_room']
        rectified_df['_key'] = rectified_df['_building'] + '|' + rectified_df['_room']

    # 使用isin进行向量化过滤
    rect_keys = set(rectified_df['_key'].dropna())
    if not rect_keys:
        return report_df.drop(columns=['_building', '_room', '_key'])

    mask_to_keep = ~report_df['_key'].isin(rect_keys)
    result = report_df[mask_to_keep].drop(columns=['_building', '_room', '_key', '_dept'] if include_dept else ['_building', '_room', '_key'])

    return result.reset_index(drop=True)

# ---------------- 界面二：数据装载与统计 ---------------- #
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
        first_dept = (dept_raw.str.split(",", n=1, expand=True)[0].map(ensure_scalar_department))
        df["first_dept"] = first_dept
        stats["rows_after_structure"] += len(df)

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

def if2_save_excel(table1: pd.DataFrame, table2: pd.DataFrame, logs: Dict[str, List[str]], stats: Dict[str, int], output_path: str) -> None:
    """界面二：按模板输出 3 个工作表：优秀与不及格表 / 检查与各率 / 日志"""
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 1})
        cell_fmt = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})
        pct_fmt = wb.add_format({"align": "center", "valign": "vcenter", "border": 1, "num_format": "0.00%"})

        t1 = table1.copy()
        for col in ["优秀寝室", "不合格寝室"]:
            if col in t1.columns:
                t1[col] = pd.to_numeric(t1[col], errors="coerce").fillna(0).astype(int)

        t2 = table2.copy()
        for col in ["检查寝室/间", "优秀寝室/间", "不合格寝室/间"]:
            if col in t2.columns:
                t2[col] = pd.to_numeric(t2[col], errors="coerce").fillna(0).astype(int)

        def _pct(num: int, den: int) -> str:
            if not den:
                return "0%"
            return f"{(num / den) * 100:.2f}%"

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
                ordered_rows.append({"系部": dept, "优秀寝室": int(r0.get("优秀寝室", 0)), "不合格寝室": int(r0.get("不合格寝室", 0))})

        start_row_s1 = 1
        for i, r in enumerate(ordered_rows):
            row_idx = start_row_s1 + i
            ws1.write_formula(row_idx, 0, "=ROW()-1", cell_fmt)
            ws1.write(row_idx, 1, r["系部"], cell_fmt)
            ws1.write_number(row_idx, 2, r["优秀寝室"], cell_fmt)
            ws1.write_number(row_idx, 3, r["不合格寝室"], cell_fmt)

        total_row_idx_s1 = start_row_s1 + len(ordered_rows)
        first_data_excel_row_s1 = start_row_s1 + 1
        last_data_excel_row_s1 = first_data_excel_row_s1 + len(ordered_rows) - 1

        ws1.write_formula(total_row_idx_s1, 0, "=ROW()-1", cell_fmt)
        ws1.write(total_row_idx_s1, 1, TOTAL_ROW_NAME_IF2, cell_fmt)
        ws1.write_formula(total_row_idx_s1, 2, f"=SUM(C{first_data_excel_row_s1}:C{last_data_excel_row_s1})", cell_fmt)
        ws1.write_formula(total_row_idx_s1, 3, f"=SUM(D{first_data_excel_row_s1}:D{last_data_excel_row_s1})", cell_fmt)

        ws1.set_column(0, 0, 8)
        ws1.set_column(1, 1, 16)
        ws1.set_column(2, 3, 14)

        sheet2_name = "检查与各率"
        ws2 = wb.add_worksheet(sheet2_name)
        headers2 = ["序号", "系部", "检查寝室/间", "优秀寝室/间", "优秀率", "合格寝室/间", "合格率", "不合格寝室/间", "不合格率"]
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
            ws2.write_formula(row_idx, 3, f"='{sheet1_name}'!C{excel_row_no}", cell_fmt)
            ws2.write_formula(row_idx, 4, f"=IF(C{excel_row_no}>0,D{excel_row_no}/C{excel_row_no},0)", pct_fmt)
            ws2.write_formula(row_idx, 7, f"='{sheet1_name}'!D{excel_row_no}", cell_fmt)
            ws2.write_formula(row_idx, 5, f"=C{excel_row_no}-D{excel_row_no}-H{excel_row_no}", cell_fmt)
            ws2.write_formula(row_idx, 6, f"=IF(C{excel_row_no}>0,F{excel_row_no}/C{excel_row_no},0)", pct_fmt)
            ws2.write_formula(row_idx, 8, f"=IF(C{excel_row_no}>0,H{excel_row_no}/C{excel_row_no},0)", pct_fmt)

        total_row_idx_s2 = start_row_s2 + len(OUTPUT_ORDER_IF2)
        total_excel_row_s2 = total_row_idx_s2 + 1
        total_excel_row_s1 = total_row_idx_s1 + 1
        first_data_excel_row_s2 = start_row_s2 + 1
        last_data_excel_row_s2 = first_data_excel_row_s2 + len(OUTPUT_ORDER_IF2) - 1

        ws2.write_formula(total_row_idx_s2, 0, "=ROW()-1", cell_fmt)
        ws2.write(total_row_idx_s2, 1, "合计", cell_fmt)
        ws2.write_formula(total_row_idx_s2, 2, f"=SUM(C{first_data_excel_row_s2}:C{last_data_excel_row_s2})", cell_fmt)
        ws2.write_formula(total_row_idx_s2, 3, f"='{sheet1_name}'!C{total_excel_row_s1}", cell_fmt)
        ws2.write_formula(total_row_idx_s2, 4, f"=IF(C{total_excel_row_s2}>0,D{total_excel_row_s2}/C{total_excel_row_s2},0)", pct_fmt)
        ws2.write_formula(total_row_idx_s2, 7, f"='{sheet1_name}'!D{total_excel_row_s1}", cell_fmt)
        ws2.write_formula(total_row_idx_s2, 5, f"=C{total_excel_row_s2}-D{total_excel_row_s2}-H{total_excel_row_s2}", cell_fmt)
        ws2.write_formula(total_row_idx_s2, 6, f"=IF(C{total_excel_row_s2}>0,F{total_excel_row_s2}/C{total_excel_row_s2},0)", pct_fmt)
        ws2.write_formula(total_row_idx_s2, 8, f"=IF(C{total_excel_row_s2}>0,H{total_excel_row_s2}/C{total_excel_row_s2},0)", pct_fmt)

        ws2.set_column(0, 0, 8)
        ws2.set_column(1, 1, 16)
        ws2.set_column(2, 8, 14)

        wslog = wb.add_worksheet("日志")
        wslog.write(0, 0, "日期", header_fmt)
        wslog.write(0, 1, "操作内容", header_fmt)

        ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_lines: List[str] = []

        def _s(key: str) -> int:
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
            log_lines.append(
                f"运行概况：原始行={rows_raw}；结构化后={rows_struct}；有效院系行={rows_valid_dept}；去零后={rows_after_zero}；"
                f"区间/单间排除后={rows_after_ex}；最终统计行={rows_final}。"
            )
            log_lines.append(
                "阶段占比："
                f"结构化/原始={_pct(rows_struct, rows_raw)}；"
                f"有效院系/结构化={_pct(rows_valid_dept, rows_struct)}；"
                f"去零后/有效院系={_pct(rows_after_zero, rows_valid_dept)}；"
                f"排除后/去零后={_pct(rows_after_ex, rows_after_zero)}；"
                f"最终/原始={_pct(rows_final, rows_raw)}。"
            )

        if z_text or z_num:
            parts = []
            if z_text:
                parts.append(f"文本0.0行数={z_text}（占结构化行 {_pct(z_text, rows_struct)}）")
            if z_num:
                parts.append(f"数值0行数={z_num}（占结构化行 {_pct(z_num, rows_struct)}）")
            log_lines.append("去零明细：" + "；".join(parts))
        else:
            log_lines.append("去零明细：本次未检测到文本 0.0 或数值 0 需要排除的记录。")

        if excl:
            log_lines.append(f"区间/单间排除：共排除 {excl} 行（占去零后行 {_pct(excl, rows_after_zero)}）。")
        else:
            log_lines.append("区间/单间排除：未配置或未命中需排除的寝室。")

        if sheets_total:
            log_lines.append(f"工作表统计：总工作表数={sheets_total}；被识别为楼栋表={sheets_used}；其它类型工作表={sheets_other}。")

        for key, label in [
            ("used_building_sheets", "已使用楼栋表"),
            ("ignored_non_building", "忽略（非楼栋表）"),
            ("ignored_missing_columns", "忽略（缺必需列）"),
        ]:
            for s in logs.get(key, []):
                log_lines.append(f"{label}：{s}")

        if not log_lines:
            log_lines.append("本次运行未记录到特殊说明。")

        for i, text in enumerate(log_lines, start=1):
            wslog.write(i, 0, ts, cell_fmt)
            wslog.write(i, 1, text, cell_fmt)

        wslog.set_column(0, 0, 20)
        wslog.set_column(1, 1, 80)

# ---------------- 界面一 ---------------- #
def load_detail_dataframe(input_path: str, dept_dictionary: Dict[str, Dict[str, List[str]]]) -> Tuple[pd.DataFrame, Dict[str, int]]:
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
    meta = {"sheet": sheet, "raw_rows": int(len(raw_num)), "valid_rows": int(len(df))}
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
    分表导出（界面一）
    """
    df1, meta = load_detail_dataframe(input_path, dept_dictionary)
    sheet = meta["sheet"]

    if use_majority_dept and not df1.empty:
        df1["院系"] = (
            df1.groupby(["楼栋_norm", "宿舍号"])["院系"]
            .transform(pick_majority)
            .fillna(df1["院系"])
        )

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
    targets = candidates[:max_sheets] if max_sheets > 0 else candidates

    df_full = df1.copy()
    df_full["总分_rawstr"] = df_full["总分_rawstr"].map(normalize_plain_text)
    df_full["总分_num"] = df_full["总分_rawstr"].map(parse_score)

    score_raw = df_full["总分_rawstr"]
    score_num = df_full["总分_num"]

    zero_mask = score_raw.map(_is_text_zero_0dot0_only)
    zero_df = df_full[zero_mask].copy()
    fail_mask = score_num.notna() & (score_num >= 0) & (score_num < 60) & (~zero_mask)
    fail_df = df_full[fail_mask].copy()

    for _tmp in (fail_df, zero_df):
        if not _tmp.empty:
            _tmp["_room_key_"] = _tmp["宿舍号"].map(room_sort_key)
            _tmp.sort_values(by=["楼栋_norm", "_room_key_"], inplace=True)
            _tmp.drop(columns=["_room_key_"], inplace=True)

    buffer = io.BytesIO()
    summary_meta = {
        "sheet": sheet,
        "raw_rows": meta["raw_rows"],
        "valid_rows": meta["valid_rows"],
        "sheet_count": int(len(targets)),
        "fail_rows": int(len(fail_df)),
        "zero_rows": int(len(zero_df)),
    }

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        wb = writer.book
        header_bold = wb.add_format({"bold": True, "font_name": font_name, "font_size": 11, "align": "center", "valign": "vcenter", "border": 1})
        header_norm = wb.add_format({"bold": False, "font_name": font_name, "font_size": 11, "align": "center", "valign": "vcenter", "border": 1})
        cell_fmt = wb.add_format({"font_name": font_name, "font_size": 11, "align": "center", "valign": "vcenter", "border": 1})
        red_fill = wb.add_format({"font_name": font_name, "font_size": 11, "bg_color": "#FFCCCC"})

        summary_rows = []
        keep_headers = ["序号", "楼栋", "宿舍号", "院系", "总分"]
        used_sheet_names: set = set()  # 跟踪已使用的工作表名称
        for name in targets:
            sub = slim[slim["楼栋_norm"].astype(str).str.strip() == name].copy()
            sub["_room_key_"] = sub["宿舍号"].map(room_sort_key)
            sub = sub.sort_values("_room_key_").drop(columns=["_room_key_"])
            out = sub[["序号", "楼栋_norm", "宿舍号", "院系", "总分", "总分_rawstr"]].copy()
            out = out.rename(columns={"楼栋_norm": "楼栋", "总分_rawstr": "总分原值(隐藏)"})

            # 处理工作表名称，防止重名
            base_name = (name if len(name) <= 28 else name[:28]) or "未命名"
            ws_name = base_name
            suffix = 1
            while ws_name in used_sheet_names:
                suffix_str = f"_{suffix}"
                ws_name = base_name[:31 - len(suffix_str)] + suffix_str
                suffix += 1
            used_sheet_names.add(ws_name)

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

                ws.autofilter(0, 0, nrows, len(headers) - 1)

            summary_rows.append([ws_name, nrows])

        def _export_special(df_src: pd.DataFrame, sheet_name: str):
            if df_src.empty:
                return
            base_cols = ["序号", "楼栋", "宿舍号", "院系", "总分", "班级", "学生姓名", "评分状态", "检查时间", "打分原因"]
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
            ws_sp.set_column(len(cols_all) - 1, len(cols_all) - 1, None, None, {"hidden": True})

            for r in range(1, len(sp) + 1):
                ws_sp.write_formula(r, 0, "=ROW()-1", cell_fmt)

            ws_sp.autofilter(0, 0, len(sp), len(cols_all) - 1)

        _export_special(fail_df, "不及格寝室明细")
        _export_special(zero_df, "0.0分寝室明细")

        toc = pd.DataFrame(summary_rows, columns=["工作表名", "行数"])
        toc.to_excel(writer, sheet_name="目录", index=False, header=False, startrow=1)
        ws_toc = writer.sheets["目录"]
        ws_toc.write(0, 0, "工作表名", header_norm)
        ws_toc.write(0, 1, "行数", header_norm)
        ws_toc.set_column(0, 1, 18, cell_fmt)
        ws_toc.set_row(0, 18)

        meta_df = pd.DataFrame({
            "字段": ["版本号", "构建时间", "说明"],
            "值": [__version__, dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), __build_note__],
        })
        try:
            meta_df.to_excel(writer, sheet_name="版本信息", index=False)
        except Exception:
            pass

    buffer.seek(0)
    return buffer.getvalue(), summary_meta

