"""분기 사업분석 탭 — 엑셀 레이아웃·수식은 그대로 두고, 라벨별 대시보드 값만 반영한다."""
from __future__ import annotations

import os
import re
import shutil
from datetime import date
from typing import Any

import pandas as pd
import streamlit as st

try:
    from openpyxl import load_workbook
except Exception:  # pragma: no cover
    load_workbook = None

QA_DIR = os.path.join("uploaded_cache", "quarterly_analysis")
QA_TEMPLATE = os.path.join(QA_DIR, "template.xlsx")
QA_WORK = os.path.join(QA_DIR, "work.xlsx")
QA_SHEET = "초안_v2"

# 엑셀에 쓰는 입력 칸만. 수식 칸(K6/K7/K9, D26…)·서술 칸은 건드리지 않는다.
# month = 분기 안 슬롯(1·2·3). 실제 달력 월은 선택한 분기로 매핑한다.
# kind: goal | sales_mil | prior_sales_mil | rate_vs_goal | rate_vs_prior | qty_ton | qty_bt | ...
_GOAL_CELLS = ("E6", "G6", "I6")
_QA_FIELDS: list[dict[str, Any]] = [
    # ■ 목표 / 달성 — 목표는 사용자가 직접 입력
    {"cell": "E6", "label": "목표 · {m1}", "kind": "goal", "month": 1},
    {"cell": "G6", "label": "목표 · {m2}", "kind": "goal", "month": 2},
    {"cell": "I6", "label": "목표 · {m3}", "kind": "goal", "month": 3},
    {"cell": "E7", "label": "실적 · {m1}", "kind": "sales_mil", "month": 1},
    {"cell": "G7", "label": "실적 · {m2}", "kind": "sales_mil", "month": 2},
    {"cell": "I7", "label": "실적 · {m3}", "kind": "sales_mil", "month": 3},
    {"cell": "E8", "label": "달성률 · {m1}", "kind": "rate_vs_goal", "month": 1, "goal_cell": "E6"},
    {"cell": "G8", "label": "달성률 · {m2}", "kind": "rate_vs_goal", "month": 2, "goal_cell": "G6"},
    {"cell": "I8", "label": "달성률 · {m3}", "kind": "rate_vs_goal", "month": 3, "goal_cell": "I6"},
    {"cell": "K8", "label": "달성률 · 합계", "kind": "rate_vs_goal_sum"},
    {"cell": "E9", "label": "전년 매출 · {m1}", "kind": "prior_sales_mil", "month": 1},
    {"cell": "G9", "label": "전년 매출 · {m2}", "kind": "prior_sales_mil", "month": 2},
    {"cell": "I9", "label": "전년 매출 · {m3}", "kind": "prior_sales_mil", "month": 3},
    {"cell": "E10", "label": "전년 비교 · {m1}", "kind": "rate_vs_prior", "month": 1},
    {"cell": "G10", "label": "전년 비교 · {m2}", "kind": "rate_vs_prior", "month": 2},
    {"cell": "I10", "label": "전년 비교 · {m3}", "kind": "rate_vs_prior", "month": 3},
    {"cell": "K10", "label": "전년 비교 · 합계", "kind": "rate_vs_prior_sum"},
    # ■ 실적 / 분석 — 당해·전년 수량 (TON / BT). Bulk 합계·전년대비 일부는 수식/텍스트 유지
    {"cell": "D21", "label": "L-O2 · 당해 {m1}", "kind": "qty_ton", "month": 1, "item": "L-O2", "year": "curr"},
    {"cell": "E21", "label": "L-O2 · 전년 {m1}", "kind": "qty_ton", "month": 1, "item": "L-O2", "year": "prior"},
    {"cell": "F21", "label": "L-O2 · 전년대비 {m1}", "kind": "qty_rate", "month": 1, "item": "L-O2"},
    {"cell": "G21", "label": "L-O2 · 당해 {m2}", "kind": "qty_ton", "month": 2, "item": "L-O2", "year": "curr"},
    {"cell": "H21", "label": "L-O2 · 전년 {m2}", "kind": "qty_ton", "month": 2, "item": "L-O2", "year": "prior"},
    {"cell": "I21", "label": "L-O2 · 전년대비 {m2}", "kind": "qty_rate", "month": 2, "item": "L-O2"},
    {"cell": "J21", "label": "L-O2 · 당해 {m3}", "kind": "qty_ton", "month": 3, "item": "L-O2", "year": "curr"},
    {"cell": "K21", "label": "L-O2 · 전년 {m3}", "kind": "qty_ton", "month": 3, "item": "L-O2", "year": "prior"},
    {"cell": "L21", "label": "L-O2 · 전년대비 {m3}", "kind": "qty_rate", "month": 3, "item": "L-O2"},
    {"cell": "D22", "label": "L-N2 · 당해 {m1}", "kind": "qty_ton", "month": 1, "item": "L-N2", "year": "curr"},
    {"cell": "E22", "label": "L-N2 · 전년 {m1}", "kind": "qty_ton", "month": 1, "item": "L-N2", "year": "prior"},
    {"cell": "F22", "label": "L-N2 · 전년대비 {m1}", "kind": "qty_rate", "month": 1, "item": "L-N2"},
    {"cell": "G22", "label": "L-N2 · 당해 {m2}", "kind": "qty_ton", "month": 2, "item": "L-N2", "year": "curr"},
    {"cell": "H22", "label": "L-N2 · 전년 {m2}", "kind": "qty_ton", "month": 2, "item": "L-N2", "year": "prior"},
    {"cell": "I22", "label": "L-N2 · 전년대비 {m2}", "kind": "qty_rate", "month": 2, "item": "L-N2"},
    {"cell": "J22", "label": "L-N2 · 당해 {m3}", "kind": "qty_ton", "month": 3, "item": "L-N2", "year": "curr"},
    {"cell": "K22", "label": "L-N2 · 전년 {m3}", "kind": "qty_ton", "month": 3, "item": "L-N2", "year": "prior"},
    {"cell": "L22", "label": "L-N2 · 전년대비 {m3}", "kind": "qty_rate", "month": 3, "item": "L-N2"},
    {"cell": "D23", "label": "L-AR · 당해 {m1}", "kind": "qty_ton", "month": 1, "item": "L-AR", "year": "curr"},
    {"cell": "E23", "label": "L-AR · 전년 {m1}", "kind": "qty_ton", "month": 1, "item": "L-AR", "year": "prior"},
    {"cell": "F23", "label": "L-AR · 전년대비 {m1}", "kind": "qty_rate", "month": 1, "item": "L-AR"},
    {"cell": "G23", "label": "L-AR · 당해 {m2}", "kind": "qty_ton", "month": 2, "item": "L-AR", "year": "curr"},
    {"cell": "H23", "label": "L-AR · 전년 {m2}", "kind": "qty_ton", "month": 2, "item": "L-AR", "year": "prior"},
    {"cell": "I23", "label": "L-AR · 전년대비 {m2}", "kind": "qty_rate", "month": 2, "item": "L-AR"},
    {"cell": "J23", "label": "L-AR · 당해 {m3}", "kind": "qty_ton", "month": 3, "item": "L-AR", "year": "curr"},
    {"cell": "K23", "label": "L-AR · 전년 {m3}", "kind": "qty_ton", "month": 3, "item": "L-AR", "year": "prior"},
    {"cell": "L23", "label": "L-AR · 전년대비 {m3}", "kind": "qty_rate", "month": 3, "item": "L-AR"},
    {"cell": "D24", "label": "L-CO2 · 당해 {m1}", "kind": "qty_ton", "month": 1, "item": "L-CO2", "year": "curr"},
    {"cell": "E24", "label": "L-CO2 · 전년 {m1}", "kind": "qty_ton", "month": 1, "item": "L-CO2", "year": "prior"},
    {"cell": "F24", "label": "L-CO2 · 전년대비 {m1}", "kind": "qty_rate", "month": 1, "item": "L-CO2"},
    {"cell": "G24", "label": "L-CO2 · 당해 {m2}", "kind": "qty_ton", "month": 2, "item": "L-CO2", "year": "curr"},
    {"cell": "H24", "label": "L-CO2 · 전년 {m2}", "kind": "qty_ton", "month": 2, "item": "L-CO2", "year": "prior"},
    {"cell": "I24", "label": "L-CO2 · 전년대비 {m2}", "kind": "qty_rate", "month": 2, "item": "L-CO2"},
    {"cell": "J24", "label": "L-CO2 · 당해 {m3}", "kind": "qty_ton", "month": 3, "item": "L-CO2", "year": "curr"},
    {"cell": "K24", "label": "L-CO2 · 전년 {m3}", "kind": "qty_ton", "month": 3, "item": "L-CO2", "year": "prior"},
    {"cell": "L24", "label": "L-CO2 · 전년대비 {m3}", "kind": "qty_rate", "month": 3, "item": "L-CO2"},
    {"cell": "D25", "label": "LPG · 당해 {m1}", "kind": "qty_ton", "month": 1, "item": "LPG", "year": "curr"},
    {"cell": "E25", "label": "LPG · 전년 {m1}", "kind": "qty_ton", "month": 1, "item": "LPG", "year": "prior"},
    {"cell": "F25", "label": "LPG · 전년대비 {m1}", "kind": "qty_rate", "month": 1, "item": "LPG"},
    {"cell": "G25", "label": "LPG · 당해 {m2}", "kind": "qty_ton", "month": 2, "item": "LPG", "year": "curr"},
    {"cell": "H25", "label": "LPG · 전년 {m2}", "kind": "qty_ton", "month": 2, "item": "LPG", "year": "prior"},
    {"cell": "I25", "label": "LPG · 전년대비 {m2}", "kind": "qty_rate", "month": 2, "item": "LPG"},
    {"cell": "J25", "label": "LPG · 당해 {m3}", "kind": "qty_ton", "month": 3, "item": "LPG", "year": "curr"},
    {"cell": "K25", "label": "LPG · 전년 {m3}", "kind": "qty_ton", "month": 3, "item": "LPG", "year": "prior"},
    {"cell": "L25", "label": "LPG · 전년대비 {m3}", "kind": "qty_rate", "month": 3, "item": "LPG"},
    {"cell": "F26", "label": "Bulk 합계 · 전년대비 {m1}", "kind": "bulk_rate", "month": 1},
    {"cell": "I26", "label": "Bulk 합계 · 전년대비 {m2}", "kind": "bulk_rate", "month": 2},
    {"cell": "L26", "label": "Bulk 합계 · 전년대비 {m3}", "kind": "bulk_rate", "month": 3},
    {"cell": "D27", "label": "Gas Cylinder · 당해 {m1}", "kind": "qty_bt", "month": 1, "year": "curr"},
    {"cell": "E27", "label": "Gas Cylinder · 전년 {m1}", "kind": "qty_bt", "month": 1, "year": "prior"},
    {"cell": "F27", "label": "Gas Cylinder · 전년대비 {m1}", "kind": "cyl_rate", "month": 1},
    {"cell": "G27", "label": "Gas Cylinder · 당해 {m2}", "kind": "qty_bt", "month": 2, "year": "curr"},
    {"cell": "H27", "label": "Gas Cylinder · 전년 {m2}", "kind": "qty_bt", "month": 2, "year": "prior"},
    {"cell": "I27", "label": "Gas Cylinder · 전년대비 {m2}", "kind": "cyl_rate", "month": 2},
    {"cell": "J27", "label": "Gas Cylinder · 당해 {m3}", "kind": "qty_bt", "month": 3, "year": "curr"},
    {"cell": "K27", "label": "Gas Cylinder · 전년 {m3}", "kind": "qty_bt", "month": 3, "year": "prior"},
    {"cell": "L27", "label": "Gas Cylinder · 전년대비 {m3}", "kind": "cyl_rate", "month": 3},
]

_BULK_ITEMS = ("L-O2", "L-N2", "L-AR", "L-CO2", "LPG")
_QA_QUARTER_MONTHS = {
    1: (1, 2, 3),
    2: (4, 5, 6),
    3: (7, 8, 9),
    4: (10, 11, 12),
}


def _qa_quarter_options(*, start_year: int = 2020, end_year: int | None = None) -> list[str]:
    """연도·분기 목록 — selectbox에서 스크롤로 고른다."""
    today = date.today()
    end = int(end_year or (today.year + 1))
    opts: list[str] = []
    for y in range(int(start_year), end + 1):
        for q in (1, 2, 3, 4):
            opts.append(f"{y}년 {q}분기")
    return opts


def _qa_parse_quarter_label(label: str) -> tuple[int, int]:
    raw = str(label or "").strip()
    m = re.match(r"(\d{4})\s*년\s*([1-4])\s*분기", raw)
    if not m:
        today = date.today()
        q = (today.month - 1) // 3 + 1
        return today.year, q
    return int(m.group(1)), int(m.group(2))


def _qa_default_quarter_label() -> str:
    today = date.today()
    q = (today.month - 1) // 3 + 1
    # 분기 초·중이면 직전 분기 기본
    if today.day < 10 and today.month in (1, 4, 7, 10):
        if q == 1:
            return f"{today.year - 1}년 4분기"
        return f"{today.year}년 {q - 1}분기"
    return f"{today.year}년 {q}분기"


def _qa_slot_month(months: tuple[int, int, int], slot: int) -> int:
    idx = max(1, min(3, int(slot))) - 1
    return int(months[idx])


def _qa_field_label(tmpl: str, months: tuple[int, int, int]) -> str:
    return (
        str(tmpl)
        .replace("{m1}", f"{months[0]}월")
        .replace("{m2}", f"{months[1]}월")
        .replace("{m3}", f"{months[2]}월")
    )


def _qa_ensure_workfile() -> str:
    os.makedirs(QA_DIR, exist_ok=True)
    if not os.path.isfile(QA_WORK):
        src = QA_TEMPLATE if os.path.isfile(QA_TEMPLATE) else None
        if not src:
            raise FileNotFoundError(
                f"분기 사업분석 템플릿이 없습니다: {QA_TEMPLATE}"
            )
        shutil.copy2(src, QA_WORK)
    return QA_WORK


def _qa_read_cells(path: str, cells: list[str]) -> dict[str, Any]:
    if load_workbook is None:
        raise RuntimeError("openpyxl 이 필요합니다.")
    wb = load_workbook(path, data_only=False)
    if QA_SHEET not in wb.sheetnames:
        wb.close()
        raise KeyError(f"시트 '{QA_SHEET}' 없음: {wb.sheetnames}")
    ws = wb[QA_SHEET]
    out = {addr: ws[addr].value for addr in cells}
    wb.close()
    return out


def _qa_write_cells(path: str, values: dict[str, Any]) -> int:
    """입력 칸만 덮어쓴다. 수식·병합·서식은 유지."""
    if load_workbook is None:
        raise RuntimeError("openpyxl 이 필요합니다.")
    wb = load_workbook(path, data_only=False)
    ws = wb[QA_SHEET]
    n = 0
    for addr, val in values.items():
        cell = ws[addr]
        cur = cell.value
        if isinstance(cur, str) and cur.startswith("="):
            continue
        if cur == val:
            continue
        cell.value = val
        n += 1
    wb.save(path)
    wb.close()
    return n


def _qa_arrow_pct(curr: float, base: float) -> str:
    if base == 0:
        return "-" if curr == 0 else "▲"
    pct = round((curr - base) / abs(base) * 100)
    if pct > 0:
        return f"▲{pct}%"
    if pct < 0:
        return f"▼{abs(pct)}%"
    return "0%"


def _qa_prepare_df(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    work = df.copy()
    if "매출일_dt" not in work.columns and "매출일" in work.columns:
        work["매출일_dt"] = pd.to_datetime(work["매출일"], errors="coerce")
    if "매출일_dt" not in work.columns:
        return pd.DataFrame()
    work["매출일_dt"] = pd.to_datetime(work["매출일_dt"], errors="coerce")
    if "매출액" in work.columns:
        work["매출액"] = pd.to_numeric(work["매출액"], errors="coerce").fillna(0.0)
    else:
        work["매출액"] = 0.0
    if "출고량" in work.columns:
        work["출고량"] = pd.to_numeric(work["출고량"], errors="coerce").fillna(0.0)
    else:
        work["출고량"] = 0.0
    if "품목명" not in work.columns:
        work["품목명"] = ""
    work = work.dropna(subset=["매출일_dt"])
    work["연도"] = work["매출일_dt"].dt.year
    work["월"] = work["매출일_dt"].dt.month
    return work


def _qa_item_mask(series: pd.Series, item: str) -> pd.Series:
    s = series.astype(str)
    u = s.str.upper().str.replace(" ", "", regex=False)
    is_bulk = u.str.contains("BULK", na=False) | s.str.contains("벌크", na=False)
    if item == "L-O2":
        return is_bulk & ~u.str.contains("CO2", na=False) & (
            u.str.contains("O2", na=False) | s.str.contains("산소", na=False)
        )
    if item == "L-N2":
        return is_bulk & (u.str.contains("N2", na=False) | s.str.contains("질소", na=False))
    if item == "L-AR":
        return is_bulk & (
            u.str.contains("AR", na=False)
            | s.str.contains("아르곤|아르", na=False)
        )
    if item == "L-CO2":
        return is_bulk & (u.str.contains("CO2", na=False) | s.str.contains("탄산", na=False))
    if item == "LPG":
        return is_bulk & (u.str.contains("LPG", na=False) | s.str.contains("엘피지|프로판", na=False))
    return pd.Series(False, index=series.index)


def _qa_is_listed_bulk(series: pd.Series) -> pd.Series:
    """실적/분석 표의 L-O2·L-N2·L-AR·L-CO2·LPG 해당 여부."""
    m = pd.Series(False, index=series.index)
    for item in _BULK_ITEMS:
        m = m | _qa_item_mask(series, item)
    return m


def _qa_is_cylinder(series: pd.Series) -> pd.Series:
    """실적/분석에서 위 5종(L-O2~LPG) 외는 전부 Gas Cylinder(실린더 매출)로 집계."""
    s = series.astype(str).str.strip()
    valid = s.str.len().gt(0) & ~s.isin({"", "nan", "None", "NaN"})
    return valid & ~_qa_is_listed_bulk(series)


def _qa_month_sales_mil(work: pd.DataFrame, year: int, month: int) -> float:
    if work.empty:
        return 0.0
    hit = work[(work["연도"] == year) & (work["월"] == month)]
    # 부가세 별도 백만원 — 매출액(공급가) / 1,000,000
    return float(round(hit["매출액"].sum() / 1_000_000.0))


def _qa_month_qty_ton(work: pd.DataFrame, year: int, month: int, item: str) -> float:
    if work.empty:
        return 0.0
    m = (work["연도"] == year) & (work["월"] == month) & _qa_item_mask(work["품목명"], item)
    # kg → TON
    return float(round(work.loc[m, "출고량"].sum() / 1000.0))


def _qa_month_qty_bt(work: pd.DataFrame, year: int, month: int) -> float:
    if work.empty:
        return 0.0
    m = (work["연도"] == year) & (work["월"] == month) & _qa_is_cylinder(work["품목명"])
    return float(round(work.loc[m, "출고량"].sum()))


def compute_dashboard_values(
    df: pd.DataFrame | None,
    *,
    year: int,
    months: tuple[int, int, int],
    goals: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """라벨(셀)별 대시보드 계산값. 목표는 사용자가 넣은 goals를 쓴다."""
    work = _qa_prepare_df(df)
    prior = year - 1
    goal_map = goals or {}
    m1, m2, m3 = (int(months[0]), int(months[1]), int(months[2]))
    cal_months = (m1, m2, m3)
    sales = {m: _qa_month_sales_mil(work, year, m) for m in cal_months}
    prior_sales = {m: _qa_month_sales_mil(work, prior, m) for m in cal_months}
    ton_curr = {
        (item, m): _qa_month_qty_ton(work, year, m, item)
        for item in _BULK_ITEMS
        for m in cal_months
    }
    ton_prior = {
        (item, m): _qa_month_qty_ton(work, prior, m, item)
        for item in _BULK_ITEMS
        for m in cal_months
    }
    bt_curr = {m: _qa_month_qty_bt(work, year, m) for m in cal_months}
    bt_prior = {m: _qa_month_qty_bt(work, prior, m) for m in cal_months}

    def _goal(cell: str) -> float:
        raw = goal_map.get(cell)
        try:
            return float(raw)
        except (TypeError, ValueError):
            return 0.0

    out: dict[str, Any] = {}
    for f in _QA_FIELDS:
        addr = f["cell"]
        kind = f["kind"]
        slot = int(f.get("month") or 0)
        month = _qa_slot_month(cal_months, slot) if slot else 0
        if kind == "goal":
            out[addr] = int(round(_goal(addr)))
            continue
        if kind == "sales_mil":
            out[addr] = int(sales[month])
            continue
        if kind == "prior_sales_mil":
            out[addr] = int(prior_sales[month])
            continue
        if kind == "rate_vs_goal":
            g = _goal(str(f.get("goal_cell") or ""))
            out[addr] = _qa_arrow_pct(sales[month], g)
            continue
        if kind == "rate_vs_goal_sum":
            gsum = _goal("E6") + _goal("G6") + _goal("I6")
            out[addr] = _qa_arrow_pct(sum(sales.values()), gsum)
            continue
        if kind == "rate_vs_prior":
            out[addr] = _qa_arrow_pct(sales[month], prior_sales[month])
            continue
        if kind == "rate_vs_prior_sum":
            out[addr] = _qa_arrow_pct(sum(sales.values()), sum(prior_sales.values()))
            continue
        item = str(f.get("item") or "")
        if kind == "qty_ton":
            src = ton_curr if f.get("year") == "curr" else ton_prior
            out[addr] = int(src[(item, month)])
            continue
        if kind == "qty_rate":
            out[addr] = _qa_arrow_pct(ton_curr[(item, month)], ton_prior[(item, month)])
            continue
        if kind == "bulk_rate":
            c = sum(ton_curr[(it, month)] for it in _BULK_ITEMS)
            p = sum(ton_prior[(it, month)] for it in _BULK_ITEMS)
            out[addr] = _qa_arrow_pct(c, p)
            continue
        if kind == "qty_bt":
            src = bt_curr if f.get("year") == "curr" else bt_prior
            out[addr] = int(src[month])
            continue
        if kind == "cyl_rate":
            out[addr] = _qa_arrow_pct(bt_curr[month], bt_prior[month])
            continue
    return out


def _qa_preview_table(
    excel_vals: dict[str, Any],
    dash_vals: dict[str, Any],
    months: tuple[int, int, int],
) -> pd.DataFrame:
    rows = []
    for f in _QA_FIELDS:
        addr = f["cell"]
        rows.append(
            {
                "셀": addr,
                "라벨": _qa_field_label(str(f["label"]), months),
                "엑셀 현재": excel_vals.get(addr),
                "반영 값": dash_vals.get(addr),
                "구분": "목표(직접입력)" if f["kind"] == "goal" else "대시보드",
            }
        )
    return pd.DataFrame(rows)


def _qa_num_or(raw: Any, default: float = 0.0) -> float:
    try:
        if raw is None or raw == "":
            return float(default)
        return float(raw)
    except (TypeError, ValueError):
        return float(default)


def render_quarterly_analysis_tab(
    df: pd.DataFrame | None = None, latest_update_str: str = ""
) -> None:
    """분기 사업분석 — 분기 스크롤 선택·목표 직접입력 후 엑셀 반영."""
    st.markdown(
        "<div class='sub-header dashboard-tab-panel-head'>📊 분기 사업분석</div>",
        unsafe_allow_html=True,
    )
    st.caption(
        "임시 탭 · 엑셀 레이아웃·수식은 그대로 두고, 라벨 값만 「반영」으로 씁니다. "
        f"시트 `{QA_SHEET}` · 작업파일 `{QA_WORK}`"
    )
    if load_workbook is None:
        st.error("openpyxl 이 없어 엑셀을 열 수 없습니다.")
        return

    try:
        path = _qa_ensure_workfile()
    except FileNotFoundError as err:
        st.error(str(err))
        st.info("템플릿 xlsx를 uploaded_cache/quarterly_analysis/template.xlsx 에 넣어 주세요.")
        return

    q_opts = _qa_quarter_options()
    default_q = _qa_default_quarter_label()
    if "qa_quarter" not in st.session_state:
        st.session_state["qa_quarter"] = default_q if default_q in q_opts else q_opts[-5]
    elif st.session_state.get("qa_quarter") not in q_opts:
        st.session_state["qa_quarter"] = default_q if default_q in q_opts else q_opts[-5]

    q_lab = st.selectbox(
        "분기 선택 (스크롤)",
        options=q_opts,
        key="qa_quarter",
        help="목록을 스크롤해 연도·분기를 고릅니다.",
    )
    year, quarter = _qa_parse_quarter_label(str(q_lab))
    months = _QA_QUARTER_MONTHS.get(quarter, (1, 2, 3))
    st.caption(f"선택: {year}년 {quarter}분기 · {months[0]}–{months[2]}월")

    cells = [f["cell"] for f in _QA_FIELDS]
    try:
        excel_vals = _qa_read_cells(path, cells)
    except Exception as err:
        st.error(f"엑셀 읽기 오류: {err}")
        return

    st.markdown("##### 목표 (백만원, 직접 입력)")
    g1, g2, g3 = st.columns(3)
    # 분기 바뀔 때 엑셀 목표로 시드. 같은 분기에서는 사용자 입력을 유지.
    seed_key = f"qa_goal_seed_{year}_{quarter}"
    if st.session_state.get("qa_goal_seed_for") != seed_key:
        st.session_state["qa_goal_seed_for"] = seed_key
        st.session_state["qa_goal_m1"] = int(_qa_num_or(excel_vals.get("E6"), 0))
        st.session_state["qa_goal_m2"] = int(_qa_num_or(excel_vals.get("G6"), 0))
        st.session_state["qa_goal_m3"] = int(_qa_num_or(excel_vals.get("I6"), 0))
    with g1:
        goal_m1 = st.number_input(
            f"목표 · {months[0]}월",
            min_value=0,
            max_value=1_000_000,
            step=1,
            key="qa_goal_m1",
        )
    with g2:
        goal_m2 = st.number_input(
            f"목표 · {months[1]}월",
            min_value=0,
            max_value=1_000_000,
            step=1,
            key="qa_goal_m2",
        )
    with g3:
        goal_m3 = st.number_input(
            f"목표 · {months[2]}월",
            min_value=0,
            max_value=1_000_000,
            step=1,
            key="qa_goal_m3",
        )
    user_goals = {"E6": int(goal_m1), "G6": int(goal_m2), "I6": int(goal_m3)}

    dash_vals = compute_dashboard_values(
        df, year=int(year), months=months, goals=user_goals
    )
    preview = _qa_preview_table(excel_vals, dash_vals, months)

    c1, c2, c3 = st.columns([1.2, 1.2, 2])
    with c1:
        apply = st.button("엑셀에 반영", type="primary", key="qa_apply", width="stretch")
    with c2:
        reset = st.button("템플릿으로 초기화", key="qa_reset", width="stretch")
    with c3:
        if latest_update_str:
            st.caption(f"대시보드 기준 시각: {latest_update_str}")
        if os.path.isfile(path):
            with open(path, "rb") as fh:
                st.download_button(
                    "작업 엑셀 다운로드",
                    data=fh.read(),
                    file_name=f"{year}년{quarter}분기사업분석.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="qa_dl",
                )

    if reset:
        if os.path.isfile(QA_TEMPLATE):
            shutil.copy2(QA_TEMPLATE, QA_WORK)
            st.session_state.pop("qa_goal_seed_for", None)
            st.success("템플릿으로 작업 파일을 다시 만들었습니다.")
            st.rerun()
        else:
            st.error("템플릿 파일이 없습니다.")

    if apply:
        to_write = {
            addr: dash_vals[addr]
            for addr in cells
            if addr in dash_vals and dash_vals[addr] is not None
        }
        try:
            n = _qa_write_cells(path, to_write)
            st.success(f"엑셀에 {n}개 칸을 반영했습니다. (수식·레이아웃 유지 · 목표 포함)")
            st.rerun()
        except Exception as err:
            st.error(f"반영 실패: {err}")

    st.markdown("##### 라벨별 값")
    st.dataframe(preview, width="stretch", hide_index=True)
    st.caption(
        "목표는 위 입력값을 엑셀 E6/G6/I6에 반영합니다. "
        "합계(K6/K7/K9)·Bulk 합계(SUM) 수식 칸은 쓰지 않습니다."
    )
