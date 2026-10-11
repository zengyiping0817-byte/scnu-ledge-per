# -*- coding: utf-8 -*-
"""
学社衔接台账生成核心脚本
================================
输入：智慧团建导出的《[XX团委]毕业生团组织转接统计信息列表YYYYMMDDHHMMSS.xlsx》
输出：
  1. 【YYYYMMDD】毕业生组织关系转接发起率统计.xlsx
  2. 【YYYYMMDD】毕业生组织关系转接完成率统计.xlsx
  3. 【YYYYMMDD】学社衔接情况.docx   （以模板为底本填充，保留全部样式）

用法：
  python build_ledger.py --input <导出文件.xlsx> [--date YYYY-MM-DD] [--outdir <目录>]

对比逻辑说明：
  · 数据日期 = 本次运行日期（或 --date 指定）
  · 上一次日期：读取 outdir 下的 history.json（若存在）中记录的最近一次日期；
    不存在则本次运行「仅记录当前日期」，对比列留空待下次填充。
  · 每次运行都会把本次快照写入 history.json，供下次对比使用。
"""

import argparse
import json
import os
import re
import sys
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt

# 同目录模块
sys.path.insert(0, str(Path(__file__).parent))
from config import (  # noqa: E402
    NAME_REPLACE, SCOPE, SCHOOL_ROOT,
    COL_ORG, COL_NOT_TRANSFERRED, COL_INITIATED, COL_NOT_INITIATED, COL_COMPLETED,
    OUT_INITIATE_XLSX, OUT_COMPLETE_XLSX, OUT_WORD, WORD_TITLE, LEDGER_ROOT_LABEL,
    OUT_PDF, PREV_INITIATE_KEYS, PREV_COMPLETE_KEYS, PREV_WORD_KEYS,
)

HISTORY_FILE = "history.json"
ASSETS_DIR = Path(__file__).parent.parent / "assets"
TEMPLATE_DOCX = ASSETS_DIR / "template.docx"


# ===========================================================================
# 工具
# ===========================================================================
def to_int(v):
    if v is None:
        return 0
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "—", "None"):
        return 0
    try:
        return int(float(s))
    except ValueError:
        return 0


def strip_prefix(name: str) -> str:
    """剥离团组织名称前缀与后缀，得到学院核心名。"""
    n = str(name or "").strip()
    n = re.sub(r"^中国共产主义青年团", "", n)
    n = re.sub(r"^共青团", "", n)
    n = re.sub(r"^华南师范大学", "", n)
    return n


def norm_core(core: str) -> str:
    """去掉结尾的「委员会」等后缀，得到可比对的核心名。"""
    n = core.strip()
    n = re.sub(r"(委员会|团委)$", "", n)
    return n


def map_college(core: str) -> str:
    """
    核心名 -> 台账学院名。
    返回台账学院名；若不属于台账统计范围内的「学院级」组织，返回 None。
    """
    base = norm_core(core)
    # 1) 名称映射（合并/改名）
    for frag, target in NAME_REPLACE:
        if frag in base:
            return target
    # 2) 精确命中白名单
    if base in SCOPE:
        return base
    return None


# ===========================================================================
# 1. 解析导出文件
# ===========================================================================
def parse_export(xlsx_path: Path):
    """
    返回:
      root: dict(initiated, not_initiated, completed, not_transferred)
      colleges: {台账学院名: {initiated, not_initiated, completed, not_transferred}}
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    idx = {name: i for i, name in enumerate(header)}

    required = [COL_ORG, COL_NOT_TRANSFERRED, COL_INITIATED, COL_NOT_INITIATED, COL_COMPLETED]
    for col in required:
        if col not in idx:
            # 容错：列名可能带空格或顺序不同
            hit = [h for h in header if col.replace(" ", "") == h.replace(" ", "")]
            if hit:
                idx[col] = header.index(hit[0])
            else:
                raise ValueError(f"导出文件缺少列：{col}\n实际列：{header}")

    root = None
    raw = {}   # 台账学院名 -> 累计数值

    for r in rows[1:]:
        if r is None or r[idx[COL_ORG]] is None:
            continue
        raw_name = str(r[idx[COL_ORG]]).strip()
        core = strip_prefix(raw_name)

        vals = {
            "not_transferred": to_int(r[idx[COL_NOT_TRANSFERRED]]),
            "initiated": to_int(r[idx[COL_INITIATED]]),
            "not_initiated": to_int(r[idx[COL_NOT_INITIATED]]),
            "completed": to_int(r[idx[COL_COMPLETED]]),
        }

        # 学校整体行：原始名恰为「华南师范大学委员会」
        if core in ("委员会", "") or norm_core(core) == "华南师范大学":
            root = vals
            continue

        # 只统计「白名单 + 映射后」命中的学院级组织
        mapped = map_college(core)
        # 排除团总支 / 年级 / 专业支部等下属组织
        is_sub = any(k in core for k in
                     ("团总支", "总支部", "支部", "年级", "专业", "班级"))
        if mapped and not is_sub:
            if mapped not in raw:
                raw[mapped] = {"initiated": 0, "not_initiated": 0,
                               "completed": 0, "not_transferred": 0}
            for k in raw[mapped]:
                raw[mapped][k] += vals[k]

    if root is None:
        # 兜底：用各学院合计
        root = {"initiated": 0, "not_initiated": 0, "completed": 0, "not_transferred": 0}
        for v in raw.values():
            for k in root:
                root[k] += v[k]

    return root, raw


def rate_initiate(d):
    den = d["initiated"] + d["not_initiated"]
    return (d["initiated"] / den) if den else None


def rate_complete(d):
    den = d["completed"] + d["not_transferred"]
    return (d["completed"] / den) if den else None


# ===========================================================================
# 1b. 解析「上一期台账」文件（两个 Excel + Word/PDF）
# ===========================================================================
def _read_rate_from_cell(v):
    """把率单元格读成 float（0~1）：支持 0.8024 浮点、'80.24%' 字符串、公式结果。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if s in ("", "-", "—", "——", "None"):
        return None
    if s.endswith("%"):
        try:
            return float(s[:-1]) / 100.0
        except ValueError:
            return None
    try:
        f = float(s)
    except ValueError:
        return None
    # 若明显是百分数写法（>1），换算为小数
    return f / 100.0 if f > 1.5 else f


def parse_prev_rate_xlsx(xlsx_path: Path, metric: str):
    """
    从「上一期台账」的率统计 Excel 中，读取该期的「本期」数据作为我们的「上期」。

    台账 Excel 有两种列布局（均为 8 或 9 列）：
      列 A=组织名称 | B,C=前一期数 | D=前一期率 | E,F=本期数 | G=本期率 | [H排序] | [I变化值]
    我们需要的是这份台账「本期」的数据（E/F/G），因为它才是我们要对比的「上一期」。

    返回：{台账学院名: {"rate": float|None, "raw": {k1:..,k2:..}|None}}
      · 率优先取 G 列（本期率）；若 G 空则退化为按 E/F 计算。
      · 兼容不同列数（8 列旧格式 / 9 列新格式）。
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))

    if metric == "initiate":
        key1, key2 = "initiated", "not_initiated"
    else:
        key1, key2 = "not_transferred", "completed"

    def cell(r, i):
        return r[i] if (r is not None and len(r) > i) else None

    result = {}
    for r in rows:
        if not r or r[0] is None:
            continue
        name = str(r[0]).strip()
        if name in ("组织名称", ""):
            continue
        core = strip_prefix(name)
        mapped = map_college(core)
        target = mapped or core
        if target == LEDGER_ROOT_LABEL or core in ("委员会", "") or norm_core(core) == LEDGER_ROOT_LABEL:
            target = LEDGER_ROOT_LABEL
        # 本期：E(4)=key1, F(5)=key2, G(6)=率
        raw = {key1: to_int(cell(r, 4)), key2: to_int(cell(r, 5))}
        rate = _read_rate_from_cell(cell(r, 6))
        if rate is None:
            den = raw[key1] + raw[key2]
            if den:
                num = raw[key1] if metric == "initiate" else raw[key2]
                rate = num / den
        result[target] = {"rate": rate, "raw": raw}

    return result


def parse_prev_word(docx_path: Path):
    """
    从「上一期」的 Word 台账中读取上期/上上期的率（表 0，7 列）。

    表结构：
      行0 大类表头；行1 期次表头（[名称, 上期, 本期, 变化值, 上期, 本期, 变化值]）
      行2起 数据：[名称, 发起率上期, 发起率本期, 变化值, 完成率上期, 完成率本期, 变化值]

    返回：
      {
        "prev_date_label": "8月31日",
        "cur_date_label": "9月7日",
        "rows": {台账学院名: {
            "initiate_prev": float|None, "initiate_cur": float|None,
            "complete_prev": float|None, "complete_cur": float|None,
        }},
        "school": {同上，针对「华南师范大学」行},
      }
    """
    doc = Document(str(docx_path))
    out = {"prev_date_label": None, "cur_date_label": None,
           "rows": {}, "school": None}
    if not doc.tables:
        return out
    table = doc.tables[0]

    # 期次表头
    try:
        hdr = table.rows[1]
        out["prev_date_label"] = hdr.cells[1].text.strip()
        out["cur_date_label"] = hdr.cells[2].text.strip()
    except Exception:
        pass

    for ri in range(2, len(table.rows)):
        cells = table.rows[ri].cells
        name = cells[0].text.strip()
        if not name:
            continue
        rec = {
            "initiate_prev": _read_rate_from_cell(cells[1].text),
            "initiate_cur": _read_rate_from_cell(cells[2].text),
            "complete_prev": _read_rate_from_cell(cells[4].text),
            "complete_cur": _read_rate_from_cell(cells[5].text),
        }
        if name == LEDGER_ROOT_LABEL:
            out["school"] = rec
            continue
        core = strip_prefix(name)
        mapped = map_college(core)
        out["rows"][mapped or core] = rec

    return out


def parse_prev_pdf(pdf_path: Path):
    """
    从「上一期」的 PDF 台账（由本脚本的 Word 导出）中读取率。
    PDF 文本流是扁平的，按「名称 + 6 个百分数」的模式逐行抓取。

    返回结构与 parse_prev_word 相同（rows/school/cur_date_label/prev_date_label）。
    """
    out = {"prev_date_label": None, "cur_date_label": None,
           "rows": {}, "school": None}
    try:
        import pymupdf
    except Exception:
        try:
            import fitz as pymupdf  # noqa
        except Exception:
            return out

    try:
        doc = pymupdf.open(str(pdf_path))
    except Exception:
        return out

    text = "\n".join(page.get_text() for page in doc)

    # 期次标签：正文「（数据更新至9月7日）」= 本期；表头里（副标题之后）首个日期 = 上期
    m = re.search(r"数据更新至\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", text)
    if m:
        out["cur_date_label"] = f"{int(m.group(1))}月{int(m.group(2))}日"
    head_zone = text[m.end():] if m else text
    heads = re.findall(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日", head_zone)
    if heads:
        out["prev_date_label"] = f"{int(heads[0][0])}月{int(heads[0][1])}日"

    # 名称集合：优先用 config 的 SCOPE + 学校行
    names = [LEDGER_ROOT_LABEL] + list(SCOPE)

    # 定位每个名称在文本中的位置（表头里也可能出现名称，取最后一次出现的更稳）
    name_pos = []
    for nm in names:
        idx = text.rfind(nm)
        if idx >= 0:
            name_pos.append((idx, nm))
    name_pos.sort()

    # token 正则：百分数、变化值（带符号百分数）、占位符「——」
    token_re = re.compile(r"([+-]?\d+(?:\.\d+)?)\s*%|(——|—)")

    for i, (idx, nm) in enumerate(name_pos):
        end = name_pos[i + 1][0] if i + 1 < len(name_pos) else len(text)
        seg = text[idx + len(nm):end]
        tokens = []
        for mt in token_re.finditer(seg):
            if mt.group(1) is not None:
                tokens.append(float(mt.group(1)) / 100.0)
            else:
                tokens.append(None)  # 「——」占位
        # 每家机构固定 6 个 token：发起上期/发起本期/变化 + 完成上期/完成本期/变化
        if len(tokens) < 6:
            continue
        t = tokens[:6]
        rec = {
            "initiate_prev": t[0],
            "initiate_cur": t[1],
            "complete_prev": t[3],
            "complete_cur": t[4],
        }
        if nm == LEDGER_ROOT_LABEL:
            out["school"] = rec
        else:
            out["rows"][nm] = rec

    return out


def classify_prev_files(paths):
    """把用户给的「上期台账」文件按类型分到 initiate/complete/word。"""
    buckets = {"initiate": [], "complete": [], "word": []}
    for p in paths:
        name = Path(p).name
        ext = Path(p).suffix.lower()
        if ext in (".docx", ".doc", ".pdf"):
            if any(k in name for k in PREV_WORD_KEYS):
                buckets["word"].append(Path(p))
            continue
        if ext in (".xlsx", ".xls", ".csv"):
            if any(k in name for k in PREV_INITIATE_KEYS):
                buckets["initiate"].append(Path(p))
            elif any(k in name for k in PREV_COMPLETE_KEYS):
                buckets["complete"].append(Path(p))
    return buckets


def parse_prev_ledger(prev_files):
    """
    综合解析「上一期台账」（两个 Excel + 可选 Word），返回统一结构：
      {
        "date_label": "9月7日",           # 上期数据日期标签（用于表头）
        "prev_date_label": "8月31日",      # 上上期（Word 里带出，用于对照，可选）
        "initiate": {名称: rate},          # 上期 发起率
        "complete": {名称: rate},          # 上期 完成率
        "school_initiate": rate|None,
        "school_complete": rate|None,
        "raw_initiate": {名称: {initiated,not_initiated}},   # 可选
        "raw_complete": {名称: {not_transferred,completed}}, # 可选
      }
    仅用 Excel 的率统计表即可；若能找到 Word，则用 Word 里的「本期率」做交叉校验并
    以 Word 为准（Word 是最终口径）。
    """
    b = classify_prev_files(prev_files)
    result = {
        "date_label": None, "prev_date_label": None,
        "initiate": {}, "complete": {},
        "school_initiate": None, "school_complete": None,
        "raw_initiate": {}, "raw_complete": {},
    }

    if b["initiate"]:
        d = parse_prev_rate_xlsx(b["initiate"][0], "initiate")
        if LEDGER_ROOT_LABEL in d:
            result["school_initiate"] = d[LEDGER_ROOT_LABEL]["rate"]
            d.pop(LEDGER_ROOT_LABEL, None)
        for k, v in d.items():
            result["initiate"][k] = v["rate"]
            if v["raw"]:
                result["raw_initiate"][k] = v["raw"]

    if b["complete"]:
        d = parse_prev_rate_xlsx(b["complete"][0], "complete")
        if LEDGER_ROOT_LABEL in d:
            result["school_complete"] = d[LEDGER_ROOT_LABEL]["rate"]
            d.pop(LEDGER_ROOT_LABEL, None)
        for k, v in d.items():
            result["complete"][k] = v["rate"]
            if v["raw"]:
                result["raw_complete"][k] = v["raw"]

    if b["word"]:
        # PDF 不解析（内容同 Word），仅当有 docx 时解析
        docx_candidates = [p for p in b["word"] if p.suffix.lower() in (".docx", ".doc")]
        pdf_candidates = [p for p in b["word"] if p.suffix.lower() == ".pdf"]
        w = None
        if docx_candidates:
            w = parse_prev_word(docx_candidates[0])
        elif pdf_candidates:
            # 无 Word 时，退化为从 PDF 文本抓取
            w = parse_prev_pdf(pdf_candidates[0])
        if w:
            result["date_label"] = w.get("cur_date_label")
            result["prev_date_label"] = w.get("prev_date_label")
            # 台账文档为准：覆盖 Excel 读到的率
            for k, rec in w["rows"].items():
                result["initiate"][k] = rec["initiate_cur"]
                result["complete"][k] = rec["complete_cur"]
            if w.get("school"):
                result["school_initiate"] = w["school"]["initiate_cur"]
                result["school_complete"] = w["school"]["complete_cur"]

    return result, b


# ===========================================================================
# 2. history.json —— 保存上一次快照
# ===========================================================================
def load_history(outdir: Path):
    p = outdir / HISTORY_FILE
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_history(outdir: Path, hist: dict):
    (outdir / HISTORY_FILE).write_text(
        json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8")


def find_prev_snapshot(hist: dict, current_date: str):
    """返回最近一次（早于 current_date）的快照及其日期。"""
    dates = sorted([d for d in hist.keys() if d < current_date])
    if not dates:
        return None, None
    last = dates[-1]
    return last, hist[last]


def fmt_rate(v):
    return "——" if v is None else f"{v*100:.2f}%"


def fmt_delta(cur, prev):
    if cur is None or prev is None:
        return "——"
    d = (cur - prev) * 100
    if abs(d) < 0.005:
        return "——"
    return f"{d:+.2f}%"


def cn_date(iso: str):
    """2026-09-07 -> 9月7日"""
    y, m, d = iso.split("-")
    return f"{int(m)}月{int(d)}日"


# ===========================================================================
# 3. 生成 Excel（发起率 / 完成率）
# ===========================================================================
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="E2EFDA")
RED = "FF0000"


def build_rate_xlsx(out_path: Path, metric: str, cur_date, prev_date,
                    root_cur, root_prev, rows_cur, rows_prev, ref_totals=None,
                    show_dates=False):
    """
    metric: 'initiate' | 'complete'

    归一化数据模型（兼容两种输入来源）：
      rows_cur / rows_prev : { 名称: R }  R = {"rate": float|None, "raw": dict|None}
      root_cur / root_prev : 同上结构的 R（学校整体行）
      · raw 存在时（来自系统导出）：数值列写原始数，率列写公式（可追溯）
      · raw 为空时（来自上期台账 Excel/Word）：数值列留空，率列写死数值
    行结构：组织名称 | 上期数1 | 上期数2 | 上期率 | 本期数1 | 本期数2 | 本期率 | 排序 | 变化值
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    if metric == "initiate":
        c1, c2 = "已发起过转接的毕业生团员数", "未发起过转接的毕业生团员数"
        col1, col2 = "initiated", "not_initiated"
    else:
        c1, c2 = "当年未转出毕业团员数", "已完成组织关系转接的毕业生团员数"
        col1, col2 = "not_transferred", "completed"

    # 第1行（列结构：A组织名称 | B,C 上期数 | D 上期率 | E,F 本期数 | G 本期率 | H 排序 | I 变化值）
    ws["A1"] = "组织名称"
    ws["B1"] = c1
    ws["C1"] = c2
    ws["D1"] = "转接发起率" if metric == "initiate" else "转接完成率"
    ws["E1"] = c1
    ws["F1"] = c2
    ws["G1"] = "转接发起率" if metric == "initiate" else "转接完成率"
    ws["H1"] = "排序"
    ws["I1"] = "变化值"

    # 第2行：A1:A2 合并；H1:H2、I1:I2 分别合并（与用户台账一致）
    ws.merge_cells("A1:A2")
    ws.merge_cells("H1:H2")
    ws.merge_cells("I1:I2")
    # 第2行各期上方显示期次日期（上期/本期），一眼分清
    if show_dates:
        for col in ("B", "C", "D"):
            ws[f"{col}2"] = cn_date(prev_date) if prev_date else "—"
        for col in ("E", "F", "G"):
            ws[f"{col}2"] = cn_date(cur_date)
    if ref_totals:
        prev_tot, cur_tot = ref_totals
        for col in ("B", "C", "D"):
            ws[f"{col}2"] = prev_tot
        for col in ("E", "F", "G"):
            ws[f"{col}2"] = cur_tot

    # 数据行
    r = 3

    # 率公式：发起率 = 已发起/(已发起+未发起) → B/(B+C)；完成率 = 已完成/(已完成+未转出) → C/(C+B)
    # 两者本质都是「分子列 / 两列之和」，列 B、C 的语义调换后除数顺序不同
    if metric == "initiate":
        def rate_formula(row):
            return f"=B{row}/(B{row}+C{row})", f"=E{row}/(E{row}+F{row})"
    else:
        def rate_formula(row):
            return f"=C{row}/(C{row}+B{row})", f"=F{row}/(F{row}+E{row})"

    def put_row(name, R_cur, R_prev, order=None):
        nonlocal r
        R_cur = R_cur or {}
        R_prev = R_prev or {}
        rate_cur = R_cur.get("rate")
        rate_prev = R_prev.get("rate")
        raw_cur = R_cur.get("raw")
        raw_prev = R_prev.get("raw")
        f_prev, f_cur = rate_formula(r)
        ws.cell(row=r, column=1, value=name)
        # 上期两数（有原始数才写）
        ws.cell(row=r, column=2, value=(raw_prev or {}).get(col1))
        ws.cell(row=r, column=3, value=(raw_prev or {}).get(col2))
        # 上期率：有原始数 → 公式；仅有率 → 直接写数值
        if raw_prev:
            ws.cell(row=r, column=4, value=f_prev)
        elif rate_prev is not None:
            ws.cell(row=r, column=4, value=round(rate_prev, 6))
        # 本期两数
        ws.cell(row=r, column=5, value=(raw_cur or {}).get(col1))
        ws.cell(row=r, column=6, value=(raw_cur or {}).get(col2))
        if raw_cur:
            ws.cell(row=r, column=7, value=f_cur)
        elif rate_cur is not None:
            ws.cell(row=r, column=7, value=round(rate_cur, 6))
        ws.cell(row=r, column=8, value=order)          # 排序号（学校行留空）
        # 变化值：两边都有率才写
        if rate_prev is not None and rate_cur is not None:
            if raw_prev and raw_cur:
                ws.cell(row=r, column=9, value=f"=G{r}-D{r}")
            else:
                ws.cell(row=r, column=9, value=round(rate_cur - rate_prev, 6))
        r += 1
        return rate_cur

    # 学校整体行（第 3 行）
    put_row(LEDGER_ROOT_LABEL, root_cur, root_prev)
    school_row = 3

    # 各学院：按本次率降序，并编号 1..N
    ordered = sorted(rows_cur.items(),
                     key=lambda kv: (kv[1].get("rate") if kv[1].get("rate") is not None else -1),
                     reverse=True)
    for idx, (name, R_cur) in enumerate(ordered, start=1):
        put_row(name, R_cur, rows_prev.get(name), order=idx)

    # 样式
    for row in ws.iter_rows(min_row=1, max_row=r - 1, max_col=9):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.column in (4, 7, 9):
                cell.number_format = "0.00%"
    # 表头（第1行 + 第2行日期/参考数格子）统一浅绿底 + 方正仿宋加粗
    for col in "ABCDEFGHI":
        c = ws[f"{col}1"]
        c.font = Font(name="方正仿宋_GBK", size=11, bold=True)
        c.fill = HEADER_FILL
    for col in "BCDEFG":
        c = ws[f"{col}2"]
        c.font = Font(name="Times New Roman", size=11, bold=True)
        c.fill = HEADER_FILL
    # 数据行：普通行用默认黑字（组织名称列用深灰 333333）；仅「华南师范大学」整行红色
    for row in ws.iter_rows(min_row=3, max_row=r - 1, max_col=9):
        is_school = (row[0].row == school_row)
        for cell in row:
            if is_school:
                cell.font = Font(name="Times New Roman", size=11, color=RED)
            elif cell.column == 1:
                cell.font = Font(name="Times New Roman", size=11, color="333333")
            else:
                cell.font = Font(name="Times New Roman", size=11)

    # 列宽（与用户台账一致；H 用默认，I 为 9.13）
    widths = {"A": 33.86, "B": 31.40, "C": 29.86, "D": 14.93,
              "E": 31.40, "F": 29.86, "G": 14.93, "I": 9.13}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

    wb.save(out_path)
    return out_path


# ===========================================================================
# 4. 生成 Word（以模板为底本）
# ===========================================================================
def set_cell_text(cell, text, font="Times New Roman", size=11,
                  bold=None, color=None):
    """改写单元格文本，尽量保留原 run 样式。"""
    # 保留第一个 run 的样式作为模板
    para = cell.paragraphs[0]
    if para.runs:
        run = para.runs[0]
        run.text = text
        for extra in para.runs[1:]:
            extra.text = ""
    else:
        run = para.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    # 设置东亚字体
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    if rFonts is None:
        from docx.oxml.ns import qn
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.append(rFonts)
    from docx.oxml.ns import qn
    if font == "Times New Roman":
        rFonts.set(qn('w:eastAsia'), '等线')
    else:
        rFonts.set(qn('w:eastAsia'), font)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def set_cell_shading(cell, fill_hex):
    """设置单元格底色（solid）。fill_hex e.g. 'C5E0B3' / 'F7CAAC' / None 表示清除。"""
    from docx.oxml.ns import qn
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = tcPr.makeelement(qn('w:shd'), {})
        tcPr.append(shd)
    if fill_hex:
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_hex)
    else:
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'auto')


def build_word(out_path: Path, base_docx: Path, cur_date, prev_date,
               root_cur, root_prev, rows_cur, rows_prev):
    """
    以模板为底本生成 Word 台账。

    归一化数据模型（与 Excel 复用）：
      rows_cur / rows_prev : { 名称: {"initiate": float|None, "complete": float|None} }
      root_cur / root_prev : 同结构（学校整体行）

    高亮规则（按「本期」的率排名，各取前 3）：
      · 绿色 C5E0B3 —— 本期转接发起率排名前 3：覆盖「组织名称 + 发起率三列(C1/C2/C3)」
      · 橙色 F7CAAC —— 本期转接完成率排名前 3：覆盖「组织名称 + 完成率三列(C4/C5/C6)」
      · 重叠时：组织名称取绿；发起率列只跟绿走，完成率列只跟橙走
        （即某组织只进完成率前3、未进发起率前3时，只有名称与完成率列标橙，发起率列不标色）
    变化值：本期率 − 上期率；为 0（含无上期数据）显示「——」，为正加「+」号，为负自带「-」。
    """
    from docx.shared import RGBColor
    doc = Document(str(base_docx))

    def g(d, k):
        """取率：d 为 {initiate, complete} 结构。"""
        if not d:
            return None
        return d.get(k)

    # 标题与副标题
    for r in doc.paragraphs[0].runs[1:]:
        r.text = ""
    doc.paragraphs[0].runs[0].text = f"{cur_date[:4]}年{WORD_TITLE}"

    sub = doc.paragraphs[1]
    for r in sub.runs[1:]:
        r.text = ""
    sub.runs[0].text = f"（数据更新至{cn_date(cur_date)}）"

    table = doc.tables[0]

    # 表头期次（第1行表头为深红底白字，模板中已存在，无需改色）
    prev_label = cn_date(prev_date) if prev_date else "—"
    set_cell_text(table.rows[1].cells[1], prev_label, font="Times New Roman", size=11, bold=True)
    set_cell_text(table.rows[1].cells[2], cn_date(cur_date), font="Times New Roman", size=11, bold=True)
    set_cell_text(table.rows[1].cells[4], prev_label, font="Times New Roman", size=11, bold=True)
    set_cell_text(table.rows[1].cells[5], cn_date(cur_date), font="Times New Roman", size=11, bold=True)

    # 数据行集合（学校行 + 各学院，按本期发起率降序）
    data_rows = [(LEDGER_ROOT_LABEL, root_cur, root_prev, None)]
    ordered = sorted(rows_cur.items(),
                     key=lambda kv: (g(kv[1], "initiate") if g(kv[1], "initiate") is not None else -1),
                     reverse=True)
    for name, d_cur in ordered:
        data_rows.append((name, d_cur, rows_prev.get(name), name))

    # 增删行：模板数据区行数 = 总行数 - 2
    need = len(data_rows)
    while len(table.rows) - 2 < need:
        new_tr = deepcopy(table.rows[-1]._tr)
        table._tbl.append(new_tr)
    while len(table.rows) - 2 > need:
        table._tbl.remove(table.rows[-1]._tr)

    # 计算本期发起率/完成率前 3（学院级，不含学校行）
    college_rows = [(n, dc) for (n, dc, dp, key) in data_rows if key is not None]

    def top3(metric):
        scored = [(n, g(dc, metric)) for (n, dc) in college_rows if g(dc, metric) is not None]
        scored.sort(key=lambda x: x[1], reverse=True)
        return [n for n, _ in scored[:3]]

    top_init = set(top3("initiate"))
    top_comp = set(top3("complete"))

    GREY = RGBColor(0x33, 0x33, 0x33)
    BLACK = RGBColor(0, 0, 0)
    RED = RGBColor(0xFF, 0, 0)

    for i, (name, d_cur, d_prev, key) in enumerate(data_rows):
        row = table.rows[i + 2]
        ri = g(d_cur, "initiate")
        ci = g(d_cur, "complete")
        rp = g(d_prev, "initiate")
        cp = g(d_prev, "complete")

        is_school = key is None
        name_font = "方正仿宋_GBK"
        if is_school:
            # 学校行：整体红色字体，无底色
            set_cell_text(row.cells[0], name, font=name_font, size=11, bold=False, color=RED)
            vals = [fmt_rate(rp), fmt_rate(ri), fmt_delta(ri, rp),
                    fmt_rate(cp), fmt_rate(ci), fmt_delta(ci, cp)]
            for j, v in enumerate(vals):
                set_cell_text(row.cells[j + 1], v, font="Times New Roman", size=11, color=RED)
            continue

        # 学院行
        green_name = key in top_init
        orange_name = key in top_comp
        name_color = GREY
        set_cell_text(row.cells[0], name, font=name_font, size=11, bold=False, color=name_color)

        vals = [fmt_rate(rp), fmt_rate(ri), fmt_delta(ri, rp),
                fmt_rate(cp), fmt_rate(ci), fmt_delta(ci, cp)]
        for j, v in enumerate(vals):
            set_cell_text(row.cells[j + 1], v, font="Times New Roman", size=11, color=BLACK)

        # 清空整行底色
        for c in row.cells:
            set_cell_shading(c, None)
        # 名称列：绿优先
        if green_name:
            set_cell_shading(row.cells[0], "C5E0B3")
        elif orange_name:
            set_cell_shading(row.cells[0], "F7CAAC")
        # 发起率三列（C1/C2/C3）：只跟绿
        if green_name:
            for j in (1, 2, 3):
                set_cell_shading(row.cells[j], "C5E0B3")
        # 完成率三列（C4/C5/C6）：只跟橙
        if orange_name:
            for j in (4, 5, 6):
                set_cell_shading(row.cells[j], "F7CAAC")

    doc.save(str(out_path))
    return out_path


def _row_label(name):
    return LEDGER_ROOT_LABEL if name == "华南师范大学" else name


# ===========================================================================
# 4b. Word -> PDF（优先 Word COM，保证排版与 .docx 完全一致）
# ===========================================================================
def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> bool:
    """
    把 docx 转成 PDF。优先使用本机 Microsoft Word（win32com），排版最保真；
    失败则尝试 LibreOffice（soffice）；再失败则返回 False（不中断主流程）。
    """
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()

    # 方案 1：Microsoft Word COM
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0
        except Exception:
            pass
        try:
            doc = word.Documents.Open(str(docx_path), ReadOnly=True)
            doc.SaveAs(str(pdf_path), FileFormat=17)  # wdFormatPDF
            doc.Close(False)
            if pdf_path.exists():
                return True
        finally:
            try:
                word.Quit()
            except Exception:
                pass
    except Exception as e:
        print(f"[提示] 用 Word 转 PDF 失败：{e}")

    # 方案 2：LibreOffice
    try:
        import subprocess
        for exe in ("soffice", "libreoffice",
                    r"C:\Program Files\LibreOffice\program\soffice.exe"):
            try:
                subprocess.run([exe, "--headless", "--convert-to", "pdf",
                                "--outdir", str(pdf_path.parent), str(docx_path)],
                               check=True, capture_output=True, timeout=180)
                if pdf_path.exists():
                    return True
            except (FileNotFoundError, subprocess.CalledProcessError):
                continue
    except Exception as e:
        print(f"[提示] 用 LibreOffice 转 PDF 失败：{e}")

    return False


# ===========================================================================
# 5. 主流程
# ===========================================================================
def _norm_export(root, rows, metric):
    """把系统导出的原始数值模型，归一化为 {名称: {"rate":..,"raw":..}}。"""
    rate_fn = rate_initiate if metric == "initiate" else rate_complete
    out = {}
    for name, d in rows.items():
        out[name] = {"rate": rate_fn(d), "raw": d}
    root_r = {"rate": rate_fn(root), "raw": root} if root else None
    return root_r, out


def _norm_prev(prev_ledger, metric):
    """把「上期台账」解析结果，归一化为 {名称: {"rate":..,"raw":..}}。"""
    rates = prev_ledger["initiate"] if metric == "initiate" else prev_ledger["complete"]
    raws = prev_ledger["raw_initiate"] if metric == "initiate" else prev_ledger["raw_complete"]
    school_rate = (prev_ledger["school_initiate"] if metric == "initiate"
                   else prev_ledger["school_complete"])
    out = {}
    for name, rate in rates.items():
        out[name] = {"rate": rate, "raw": raws.get(name)}
    root_r = {"rate": school_rate, "raw": None}
    return root_r, out


def _merge_word_model(init_root, init_rows, comp_root, comp_rows):
    """合并发起率、完成率两套数据，形成 Word 的行模型。"""
    out = {}
    for name in set(init_rows) | set(comp_rows):
        out[name] = {
            "initiate": (init_rows.get(name) or {}).get("rate"),
            "complete": (comp_rows.get(name) or {}).get("rate"),
        }
    root = {
        "initiate": (init_root or {}).get("rate"),
        "complete": (comp_root or {}).get("rate"),
    }
    return root, out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="本期数据文件：智慧团建导出 xlsx")
    ap.add_argument("--prev-files", nargs="*", default=None,
                    help="上一期台账文件（可多个）：发起率/完成率统计 xlsx + 学社衔接情况 docx/pdf。"
                         "提供后将直接从这些文件读取上期数据，无需 history.json。")
    ap.add_argument("--date", default=None, help="本次数据日期 YYYY-MM-DD，缺省取今天")
    ap.add_argument("--prev-date", default=None,
                    help="上一期数据日期 YYYY-MM-DD（配合 --prev-files 使用，"
                         "不填则尝试从文件名或 Word 中读取）")
    ap.add_argument("--outdir", default=None, help="输出目录，缺省为输入文件同目录")
    ap.add_argument("--no-history", action="store_true", help="不读写 history.json")
    ap.add_argument("--no-pdf", action="store_true", help="不生成 PDF（默认生成）")
    ap.add_argument("--ref-prev", type=int, default=None,
                    help="上一期上级参考总数（如省级毕业生团员总数），用于第2行")
    ap.add_argument("--ref-cur", type=int, default=None,
                    help="本期上级参考总数，用于第2行")
    ap.add_argument("--ref-total", type=int, default=None,
                    help="仅当无历史对比时，写入本期上级参考总数")
    ap.add_argument("--no-dates", action="store_true",
                    help="不在 Excel 第2行显示期次日期（默认显示）")
    args = ap.parse_args()

    inp = Path(args.input).resolve()
    if not inp.exists():
        print(f"[错误] 找不到输入文件：{inp}")
        sys.exit(1)

    cur_date = args.date or date.today().isoformat()
    outdir = Path(args.outdir).resolve() if args.outdir else inp.parent
    outdir.mkdir(parents=True, exist_ok=True)

    # ---- 本期数据 ----
    root_cur_raw, rows_cur = parse_export(inp)
    init_root_cur, init_rows_cur = _norm_export(root_cur_raw, rows_cur, "initiate")
    comp_root_cur, comp_rows_cur = _norm_export(root_cur_raw, rows_cur, "complete")

    # ---- 上期数据：两种来源 ----
    prev_date = None
    init_root_prev = comp_root_prev = None
    init_rows_prev, comp_rows_prev = {}, {}

    if args.prev_files:
        # 来源 A：用户直接给的上一期台账文件（两个 Excel + Word）
        prev_files = [Path(p).resolve() for p in args.prev_files]
        for p in prev_files:
            if not p.exists():
                print(f"[错误] 找不到上期文件：{p}")
                sys.exit(1)
        led, buckets = parse_prev_ledger(prev_files)
        print(f"[信息] 上期文件识别：发起率表 {len(buckets['initiate'])} 个、"
              f"完成率表 {len(buckets['complete'])} 个、Word/PDF {len(buckets['word'])} 个")
        if not buckets["initiate"] and not buckets["complete"] and not buckets["word"]:
            print("[警告] 未能从 --prev-files 识别出可用的上期文件，"
                  "请确认文件名包含「发起率 / 完成率 / 学社衔接情况」。")
        # 上期日期
        if args.prev_date:
            prev_date = args.prev_date
        elif led.get("date_label"):
            prev_date = _label_to_iso(led["date_label"], cur_date)
        else:
            prev_date = _infer_date_from_files(prev_files, cur_date)
        init_root_prev, init_rows_prev = _norm_prev(led, "initiate")
        comp_root_prev, comp_rows_prev = _norm_prev(led, "complete")
        # 校验：若实际上没读到任何率数据，则不声称有对比基准
        got_any = bool(led["initiate"]) or bool(led["complete"])
        if not got_any:
            print("[警告] 未能从 --prev-files 读出上期率数据，本次将不生成对比列。")
            prev_date = None
        print(f"[信息] 上期数据日期：{prev_date or '未识别'}")

    else:
        # 来源 B：history.json 自动对比
        if not args.no_history:
            hist = load_history(outdir)
            prev_date, prev_snap = find_prev_snapshot(hist, cur_date)
            if prev_snap:
                p_root = prev_snap.get("root")
                p_cols = prev_snap.get("colleges", {})
                init_root_prev, init_rows_prev = _norm_export(p_root, p_cols, "initiate")
                comp_root_prev, comp_rows_prev = _norm_export(p_root, p_cols, "complete")
            hist[cur_date] = {"root": root_cur_raw, "colleges": rows_cur}
            save_history(outdir, hist)

    # 写历史（用上期文件或本期数据都记录一份本期快照）
    if args.prev_files and not args.no_history:
        hist = load_history(outdir)
        hist[cur_date] = {"root": root_cur_raw, "colleges": rows_cur}
        save_history(outdir, hist)

    tag = cur_date.replace("-", "")
    f1 = outdir / f"【{tag}】{OUT_INITIATE_XLSX}.xlsx"
    f2 = outdir / f"【{tag}】{OUT_COMPLETE_XLSX}.xlsx"
    f3 = outdir / f"【{tag}】{OUT_WORD}.docx"
    f4 = outdir / f"【{tag}】{OUT_PDF}.pdf"

    ref = None
    if args.ref_prev is not None or args.ref_cur is not None:
        ref = (args.ref_prev, args.ref_cur)

    build_rate_xlsx(f1, "initiate", cur_date, prev_date, init_root_cur, init_root_prev,
                    init_rows_cur, init_rows_prev, ref, show_dates=not args.no_dates)
    build_rate_xlsx(f2, "complete", cur_date, prev_date, comp_root_cur, comp_root_prev,
                    comp_rows_cur, comp_rows_prev, ref, show_dates=not args.no_dates)

    # Word 用「发起率+完成率」合并模型
    word_root_cur, word_rows_cur = _merge_word_model(
        init_root_cur, init_rows_cur, comp_root_cur, comp_rows_cur)
    word_root_prev, word_rows_prev = _merge_word_model(
        init_root_prev, init_rows_prev, comp_root_prev, comp_rows_prev)

    if TEMPLATE_DOCX.exists():
        build_word(f3, TEMPLATE_DOCX, cur_date, prev_date,
                   word_root_cur, word_root_prev, word_rows_cur, word_rows_prev)
    else:
        print(f"[警告] 未找到模板 {TEMPLATE_DOCX}，跳过 Word 生成")

    # PDF（第 4 个文件）
    pdf_ok = False
    if not args.no_pdf and f3.exists():
        pdf_ok = convert_docx_to_pdf(f3, f4)

    print("生成完成：")
    print("  ", f1)
    print("  ", f2)
    print("  ", f3)
    if pdf_ok:
        print("  ", f4)
    elif not args.no_pdf:
        print(f"   [注意] PDF 未生成（缺 Word/LibreOffice？），可先用 --no-pdf 跳过")
    if prev_date:
        print(f"对比基准日期：{prev_date}")
    else:
        print("本次无上期对比数据（对比列留空）")


def _label_to_iso(label, ref_iso):
    """'9月7日' -> '2026-09-07'（年份取参考日期所在年，跨年则取上一年）。"""
    m = re.match(r"\s*(\d{1,2})月(\d{1,2})日", str(label or ""))
    if not m:
        return None
    mo, da = int(m.group(1)), int(m.group(2))
    ref_y = int(str(ref_iso)[:4])
    try:
        cand = date(ref_y, mo, da)
    except ValueError:
        return None
    # 若推算出的日期晚于参考日期，说明是上一年
    if cand.isoformat() >= str(ref_iso):
        try:
            cand = date(ref_y - 1, mo, da)
        except ValueError:
            pass
    return cand.isoformat()


def _infer_date_from_files(prev_files, cur_iso):
    """从文件名里的 8 位日期（如 20260907 / 【20260907】）推断上期日期。"""
    for p in prev_files:
        m = re.search(r"(20\d{2})(\d{2})(\d{2})", Path(p).name)
        if m:
            iso = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            if iso < str(cur_iso):
                return iso
    return None


if __name__ == "__main__":
    main()
