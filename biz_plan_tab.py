"""분기 사업분석 — 로컬 전용.

목표/달성·실적/분석 두 표를 매출에서 취합하고, 목표만 직접 고친다.
저장은 양식 파일의 셀 값만 바꾸고 서식 XML은 그대로 둔다.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import unicodedata
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd
import streamlit as st

QUARTERS = {1: (1, 2, 3), 2: (4, 5, 6), 3: (7, 8, 9), 4: (10, 11, 12)}
MONEY_COLS = (5, 7, 9)  # E G I
QTY_NOW_COLS = (4, 7, 10)  # D G J
QTY_PREV_COLS = (5, 8, 11)  # E H K
QTY_PCT_COLS = (6, 9, 12)  # F I L
_PRODUCTS = ("LO2", "LN2", "LAR", "LCO2", "LPG", "CYL")
_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
_NOISE = ("운송", "운반", "검사", "입금", "수수료", "임대", "차액", "잔액", "이월")
_LN2_LITER_KG = 0.808


def _nfc(text: str) -> str:
    return unicodedata.normalize("NFC", str(text or ""))


def find_plan_root() -> Path | None:
    """스크린샷 경로: Drive › 다른 컴퓨터 › 내 컴퓨터 (1) › Desktop › 업무 › 사업계획."""
    try:
        cached = st.session_state.get("_bp_plan_root")
    except Exception:
        cached = None
    if isinstance(cached, str) and cached:
        path = Path(cached)
        if path.is_dir():
            return path
    found = _scan_plan_root()
    if found is not None:
        try:
            st.session_state["_bp_plan_root"] = str(found)
        except Exception:
            pass
    return found


def _scan_plan_root() -> Path | None:
    base = Path.home() / "Library" / "CloudStorage"
    if not base.is_dir():
        return None
    for drive in base.iterdir():
        if not drive.name.startswith("GoogleDrive-"):
            continue
        other = drive / "다른 컴퓨터"
        if not other.is_dir():
            continue
        for comp in other.iterdir():
            name = _nfc(comp.name)
            if "내 컴퓨터 (1)" not in name:
                continue
            return comp / "Desktop" / "업무" / "사업계획"
    return None


def _forms_under(root: Path) -> Path | None:
    """사업계획 폴더 안의 양식, 없으면 이미 있는 분기 파일을 고른다."""
    if not root.is_dir():
        return None
    drafts: list[Path] = []
    quarters: list[Path] = []
    for path in root.rglob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        name = _nfc(path.name)
        if "26Q3" in name and "draft" in name.lower():
            drafts.append(path)
        elif name.endswith("분기 사업계획.xlsx"):
            quarters.append(path)
    if drafts:
        return sorted(drafts, key=lambda p: _nfc(str(p)))[0]
    if quarters:
        return sorted(quarters, key=lambda p: _nfc(str(p)))[0]
    return None


def find_template() -> Path | None:
    root = find_plan_root()
    cands = [Path.home() / "Desktop" / "26Q3 사업분석_draft 261016.xlsx"]
    if root is not None:
        cands.append(root / "2026" / "26Q3 사업분석_draft 261016.xlsx")
    for path in cands:
        if path.is_file():
            return path
    if root is not None:
        found = _forms_under(root)
        if found is not None:
            return found
    desktop_quarter = Path.home() / "Desktop" / "4분기 사업계획.xlsx"
    if desktop_quarter.is_file():
        return desktop_quarter
    return None


def _create_quarter_book(path: Path, year: int, quarter: int) -> None:
    """양식이 없을 때 같은 칸 구조의 통합문서를 그 경로에 만든다."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{str(int(year))[2:]}Q{int(quarter)}"
    months = QUARTERS[int(quarter)]
    yy = str(int(year))[2:]
    py = str(int(year) - 1)[2:]
    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    head_fill = PatternFill("solid", fgColor="1F4E79")
    head_font = Font(color="FFFFFF", bold=True, name="맑은 고딕")
    label_font = Font(name="맑은 고딕")
    ws["C2"] = f"’{yy}년 {int(quarter)}분기 사업분석"
    ws["C2"].font = Font(bold=True, size=16, name="맑은 고딕")
    ws["C5"] = "구분"
    ws["C5"].font = head_font
    ws["C5"].fill = head_fill
    for col, month in zip(MONEY_COLS, months):
        cell = ws.cell(5, col, f"{month}월")
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin
    for offset, label in enumerate(("목표", "실적", "달성율", "전년매출", "전년비교"), start=6):
        ws.cell(offset, 4, label).font = label_font
        for col in MONEY_COLS:
            cell = ws.cell(offset, col, 0)
            cell.border = thin
            cell.alignment = Alignment(horizontal="right")
    for row in range(12, 20):
        ws.cell(row, 3, "")
    ws["C20"] = "■ 실적 / 분석"
    qty = 22
    ws.cell(qty, 3, "구분").font = head_font
    ws.cell(qty, 3).fill = head_fill
    for col, month in zip(QTY_NOW_COLS, months):
        ws.cell(qty, col, f"’{yy}년 {month}월")
    for col, month in zip(QTY_PREV_COLS, months):
        ws.cell(qty, col, f"’{py}년 {month}월")
    for i, name in enumerate(("LO2", "LN2", "LAR", "LCO2", "LPG", "Gas Cylinder", "벌크 합계")):
        row = qty + 1 + i
        ws.cell(row, 3, name).font = label_font
        for col in (*QTY_NOW_COLS, *QTY_PREV_COLS):
            cell = ws.cell(row, col, 0)
            cell.border = thin
    ws["C32"] = " ※ 원인 / 분석"
    for row in range(33, 44):
        ws.cell(row, 3, "")
    ws["C44"] = " ※ 전망"
    for row in range(45, 61):
        ws.cell(row, 3, "")
    wb.save(path)
    wb.close()


def quarter_path(year: int, quarter: int) -> Path | None:
    root = find_plan_root()
    if root is None:
        return None
    return root / str(int(year)) / f"{int(quarter)}분기 사업계획.xlsx"


def arrow_pct(current: float, base: float) -> str:
    if not base:
        return ""
    pct = int(round((float(current) - float(base)) / float(base) * 100))
    if pct > 0:
        return f"▲{pct}%"
    if pct < 0:
        return f"▼{abs(pct)}%"
    return "0%"


def product_bucket(name: str) -> str | None:
    """벌크는 품목 톤, 벌크가 아닌 실린더는 병(CYL). 운송·검사 같은 줄은 뺀다."""
    text = str(name or "").strip()
    if not text or text.lower() == "nan":
        return None
    if any(word in text for word in _NOISE):
        return None
    upper = text.upper().replace(" ", "")
    bulk = "BULK" in upper or "벌크" in text
    if bulk and ("CO2" in upper or "탄산" in text):
        return "LCO2"
    if bulk and ("LPG" in upper or "프로판" in text):
        return "LPG"
    if bulk and ("AR" in upper or "아르" in text):
        return "LAR"
    if bulk and ("N2" in upper or "질소" in text):
        return "LN2"
    if bulk and ("O2" in upper or "산소" in text):
        return "LO2"
    if bulk:
        return None
    return "CYL"


def _tons(name: str, qty: float, bucket: str) -> float:
    amount = float(qty or 0)
    if bucket == "CYL":
        return amount
    lowered = str(name).lower()
    if "liter" in lowered or "리터" in str(name):
        factor = _LN2_LITER_KG if bucket == "LN2" else 1.0
        return amount * factor / 1000.0
    return amount / 1000.0


def _month_frame(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    if "매출일_dt" in df.columns:
        dt = pd.to_datetime(df["매출일_dt"], errors="coerce")
        mask = (dt.dt.year == int(year)) & (dt.dt.month == int(month))
        return df.loc[mask]
    if "연도" not in df.columns or "월" not in df.columns:
        return pd.DataFrame()
    years = df["연도"].astype(str).str.slice(0, 4)
    months = df["월"].astype(str)
    want = {f"{int(month)}월", f"{int(month):02d}월"}
    return df.loc[(years == str(int(year))) & months.isin(want)]


def _sum_money(frame: pd.DataFrame) -> int:
    """부가세 별도 매출액을 백만원으로 반올림."""
    if frame is None or frame.empty or "매출액" not in frame.columns:
        return 0
    won = pd.to_numeric(frame["매출액"], errors="coerce").fillna(0).sum()
    return int(round(float(won) / 1_000_000.0))


def _sum_qty(frame: pd.DataFrame) -> dict[str, int]:
    out = {key: 0.0 for key in _PRODUCTS}
    if frame is None or frame.empty or "품목명" not in frame.columns or "출고량" not in frame.columns:
        return {key: 0 for key in _PRODUCTS}
    qty = pd.to_numeric(frame["출고량"], errors="coerce").fillna(0)
    for name, amount in zip(frame["품목명"].astype(str), qty):
        bucket = product_bucket(name)
        if bucket is None:
            continue
        out[bucket] += _tons(name, float(amount), bucket)
    return {key: int(round(val)) for key, val in out.items()}


def aggregate_quarter(df: pd.DataFrame, year: int, quarter: int) -> dict:
    months = QUARTERS[int(quarter)]
    now_money, prev_money = [], []
    now_qty = {key: [] for key in _PRODUCTS}
    prev_qty = {key: [] for key in _PRODUCTS}
    dated = None
    if isinstance(df, pd.DataFrame) and not df.empty and "매출일_dt" in df.columns:
        dated = pd.to_datetime(df["매출일_dt"], errors="coerce")
    for month in months:
        if dated is not None:
            cur = df.loc[(dated.dt.year == int(year)) & (dated.dt.month == int(month))]
            old = df.loc[(dated.dt.year == int(year) - 1) & (dated.dt.month == int(month))]
        else:
            cur = _month_frame(df, int(year), month)
            old = _month_frame(df, int(year) - 1, month)
        now_money.append(_sum_money(cur))
        prev_money.append(_sum_money(old))
        cur_q = _sum_qty(cur)
        old_q = _sum_qty(old)
        for key in _PRODUCTS:
            now_qty[key].append(cur_q[key])
            prev_qty[key].append(old_q[key])
    return {
        "months": months,
        "money_now": now_money,
        "money_prev": prev_money,
        "qty_now": now_qty,
        "qty_prev": prev_qty,
    }


def _fmt_mil(value: int) -> str:
    return f"{int(value):,}"


def quarter_money(df: pd.DataFrame, year: int, quarter: int) -> int:
    total = 0
    for month in QUARTERS[int(quarter)]:
        total += _sum_money(_month_frame(df, int(year), month))
    return int(total)


def sales_goal_lines(year: int, quarter: int, annual: int, df: pd.DataFrame) -> list[str]:
    """목표/달성 아래 1번 문장. 매출이 있으면 분기 실적과 달성률을 채운다."""
    annual = int(annual or 0)
    quarter = int(quarter)
    yy = str(int(year))[2:]
    actuals = [quarter_money(df, year, q) for q in (1, 2, 3, 4)]
    prev = quarter_money(df, int(year) - 1, quarter)
    q_now = actuals[quarter - 1]
    ytd = sum(actuals[:quarter])

    def _ratio(num: float, den: float) -> int | None:
        if not den:
            return None
        return int(round(float(num) / float(den) * 100))

    q_pct = _ratio(q_now, annual)
    yoy = _ratio(q_now - prev, prev)
    ytd_pct = _ratio(ytd, annual)
    goal_bit = f"목표대비 {q_pct}% 달성" if q_pct is not None else "목표대비 달성"
    if yoy is None:
        yoy_bit = "전년 실적 없음"
    elif yoy > 0:
        yoy_bit = f"전년 분기 대비 {yoy}% 상승"
    elif yoy < 0:
        yoy_bit = f"전년 분기 대비 {abs(yoy)}% 하락"
    else:
        yoy_bit = "전년 분기 대비 보합"
    ytd_bit = f"목표대비 {ytd_pct}% 달성" if ytd_pct is not None else "목표대비 달성"
    indent = " " * 31
    parts = [f"{i}분기 실적 {_fmt_mil(actuals[i - 1])}백만" for i in range(1, quarter + 1)]
    return [
        f"1. {yy}년 매출목표      : {_fmt_mil(annual)}백만 / {quarter}분기 실적 {_fmt_mil(q_now)}백만 / {goal_bit}",
        f"{indent}전년 {quarter}분기 실적 {_fmt_mil(prev)}백만  {yoy_bit}",
        f"{indent}{' / '.join(parts)}",
        f"{indent}{ytd_bit}",
    ]


def growth_top(df: pd.DataFrame, year: int, quarter: int, limit: int = 30) -> list[tuple[str, int, int, int]]:
    """전년 같은 분기보다 매출이 늘어난 거래처. (이름, 올해 원, 전년 원, 증가 원)."""
    if not isinstance(df, pd.DataFrame) or df.empty or "거래처" not in df.columns or "매출액" not in df.columns:
        return []
    months = set(QUARTERS[int(quarter)])

    def _slice(which_year: int) -> pd.DataFrame:
        if "매출일_dt" in df.columns:
            dated = pd.to_datetime(df["매출일_dt"], errors="coerce")
            return df.loc[(dated.dt.year == int(which_year)) & (dated.dt.month.isin(months))]
        parts = [_month_frame(df, which_year, month) for month in months]
        parts = [part for part in parts if isinstance(part, pd.DataFrame) and not part.empty]
        if not parts:
            return pd.DataFrame()
        return pd.concat(parts, ignore_index=True)

    def _totals(frame: pd.DataFrame) -> dict[str, float]:
        if frame is None or frame.empty:
            return {}
        money = pd.to_numeric(frame["매출액"], errors="coerce").fillna(0)
        names = frame["거래처"].astype(str).str.strip()
        return money.groupby(names).sum().to_dict()

    now = _totals(_slice(int(year)))
    old = _totals(_slice(int(year) - 1))
    rows = []
    for name in set(now) | set(old):
        if not name or name.lower() == "nan" or name in {"미지정", "(거래처명 없음)"}:
            continue
        current = float(now.get(name) or 0)
        previous = float(old.get(name) or 0)
        if current <= previous:
            continue
        rows.append((name, int(round(current)), int(round(previous)), int(round(current - previous))))
    rows.sort(key=lambda item: item[3], reverse=True)
    return rows[: int(limit)]


def cause_lines(picked: list[str], ranked: list[tuple[str, int, int, int]], year: int, quarter: int) -> list[str]:
    yy = str(int(year) - 1)[2:]
    amounts = {name: delta for name, _now, _old, delta in ranked}
    lines = [f"2. {yy}년 {int(quarter)}분기 대비 매출성장업체"]
    for name in picked:
        man = int(round(float(amounts.get(name) or 0) / 10000.0))
        lines.append(f"      • {name} : {man:,}만")
    return lines


def _marker_rows(ws) -> dict[str, int | None]:
    perf = cause = outlook = nxt = None
    for row in ws.iter_rows(min_row=1, max_row=120, max_col=3):
        cell = row[2]
        raw = str(cell.value or "")
        text = _norm_label(raw)
        if not text:
            continue
        if perf is None and "실적" in text and "분석" in text:
            perf = cell.row
        elif cause is None and "원인" in text:
            cause = cell.row
        elif outlook is None and "전망" in text:
            outlook = cell.row
        elif nxt is None and raw.strip().startswith("■") and perf and cell.row > (outlook or perf):
            nxt = cell.row
    return {"perf": perf, "cause": cause, "outlook": outlook, "next": nxt}


def _label_bucket(label: str) -> str | None:
    text = str(label or "").replace(" ", "").replace("-", "").upper()
    if "합계" in str(label):
        return "BULK"
    if "CYL" in text or "GAS" in text or "실린더" in str(label) or "병" in str(label):
        return "CYL"
    if "CO2" in text or "LCO2" in text:
        return "LCO2"
    if "LPG" in text:
        return "LPG"
    if "LAR" in text or text.endswith("AR"):
        return "LAR"
    if "LN2" in text or text.endswith("N2"):
        return "LN2"
    if "LO2" in text or text.endswith("O2"):
        return "LO2"
    return None


def _norm_label(value) -> str:
    return str(value or "").replace(" ", "").replace("\n", "")


def locate_sheet(path: Path) -> dict:
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb.active
    money_header = None
    qty_header = None
    for row in ws.iter_rows(min_row=1, max_row=80, max_col=12):
        label = _norm_label(row[2].value)
        if label == "구분":
            if money_header is None:
                money_header = row[2].row
            elif qty_header is None:
                qty_header = row[2].row
    roles: dict[str, int] = {}
    if money_header:
        for row in ws.iter_rows(min_row=money_header + 1, max_row=money_header + 8, max_col=4):
            kind = _norm_label(row[3].value)
            if "목표" in kind:
                roles["target"] = row[3].row
            elif "실적" in kind:
                roles["actual"] = row[3].row
            elif "달성" in kind:
                roles["rate"] = row[3].row
            elif "전년매출" in kind:
                roles["prior"] = row[3].row
            elif "전년비교" in kind:
                roles["yoy"] = row[3].row
    products = []
    if qty_header:
        for row in ws.iter_rows(min_row=qty_header + 1, max_row=qty_header + 20, max_col=3):
            label = str(row[2].value or "").strip()
            if not label:
                continue
            if label.startswith("※") or "원인" in label:
                break
            bucket = _label_bucket(label)
            if bucket:
                products.append((row[2].row, bucket, label))
    title = None
    for row in ws.iter_rows(min_row=1, max_row=3, max_col=3):
        if row[2].value and "사업분석" in str(row[2].value):
            title = row[2].coordinate
            break
    markers = _marker_rows(ws)
    last_money = max(roles.values()) if roles else (money_header or 0)
    note_rows = list(range(int(last_money) + 2, int(markers["perf"]))) if markers["perf"] and last_money else []
    cause_rows = (
        list(range(int(markers["cause"]) + 1, int(markers["outlook"])))
        if markers["cause"] and markers["outlook"]
        else []
    )
    outlook_end = int(markers["next"] or ((markers["outlook"] or 0) + 17))
    outlook_rows = list(range(int(markers["outlook"]) + 1, outlook_end)) if markers["outlook"] else []
    return {
        "sheet": ws.title,
        "money_header": money_header,
        "qty_header": qty_header,
        "roles": roles,
        "products": products,
        "title": title,
        "note_rows": note_rows,
        "cause_rows": cause_rows,
        "outlook_rows": outlook_rows,
        "ws": ws,
        "wb": wb,
    }


def _close_located(info: dict | None) -> None:
    wb = info.get("wb") if isinstance(info, dict) else None
    if wb is None:
        return
    try:
        wb.close()
    except Exception:
        pass


def _col_letter(index: int) -> str:
    letters = ""
    n = int(index)
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _cell(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"


def _is_formula(cell) -> bool:
    return isinstance(cell.value, str) and str(cell.value).startswith("=")


def _put_lines(writes: dict[str, object], rows: list[int], lines: list[str] | None) -> None:
    if lines is None:
        return
    for i, row in enumerate(rows):
        writes[_cell(int(row), 3)] = lines[i] if i < len(lines) else ""


def read_text_blocks(path: Path | None) -> dict:
    empty = {"notes": [], "causes": [], "outlooks": [], "note_rows": [], "cause_rows": [], "outlook_rows": [], "annual": 0}
    if path is None or not path.is_file():
        return empty
    info = locate_sheet(path)
    try:
        ws = info["ws"]

        def grab(rows: list[int]) -> list[str]:
            lines = []
            for row in rows:
                value = ws.cell(int(row), 3).value
                lines.append("" if value is None else str(value))
            while lines and not str(lines[-1]).strip():
                lines.pop()
            return lines

        notes = grab(info["note_rows"])
        annual = 0
        if notes and "매출목표" in notes[0]:
            found = re.search(r"([\d,]{3,})\s*백만", notes[0])
            if found:
                annual = int(found.group(1).replace(",", ""))
        return {
            "notes": notes,
            "causes": grab(info["cause_rows"]),
            "outlooks": grab(info["outlook_rows"]),
            "note_rows": list(info["note_rows"]),
            "cause_rows": list(info["cause_rows"]),
            "outlook_rows": list(info["outlook_rows"]),
            "annual": annual,
        }
    finally:
        _close_located(info)


def build_writes(
    path: Path,
    year: int,
    quarter: int,
    targets: list[float],
    agg: dict,
    notes: list[str] | None = None,
    causes: list[str] | None = None,
    outlooks: list[str] | None = None,
) -> dict[str, object]:
    info = locate_sheet(path)
    ws = info["ws"]
    writes: dict[str, object] = {}
    yy = str(int(year))[2:]
    py = str(int(year) - 1)[2:]
    months = agg["months"]
    if info["title"]:
        writes[info["title"]] = f"’{yy}년 {quarter}분기 사업분석"
    header = info["money_header"]
    roles = info["roles"]
    if header:
        for col, month in zip(MONEY_COLS, months):
            writes[_cell(header, col)] = f"{month}월"
    if roles.get("target"):
        writes[_cell(roles["target"], 3)] = f"{int(year)}년"
        for col, value in zip(MONEY_COLS, targets):
            ref = _cell(roles["target"], col)
            if not _is_formula(ws[ref]):
                writes[ref] = int(round(float(value or 0)))
    now = [int(v) for v in agg["money_now"]]
    prev = [int(v) for v in agg["money_prev"]]
    tgt = [int(round(float(v or 0))) for v in targets]
    if roles.get("actual"):
        for col, value in zip(MONEY_COLS, now):
            ref = _cell(roles["actual"], col)
            if not _is_formula(ws[ref]):
                writes[ref] = value
    if roles.get("prior"):
        for col, value in zip(MONEY_COLS, prev):
            ref = _cell(roles["prior"], col)
            if not _is_formula(ws[ref]):
                writes[ref] = value
    if roles.get("rate"):
        for col, actual, goal in zip(MONEY_COLS, now, tgt):
            writes[_cell(roles["rate"], col)] = arrow_pct(actual, goal)
        writes[_cell(roles["rate"], 11)] = arrow_pct(sum(now), sum(tgt))
    if roles.get("yoy"):
        for col, actual, old in zip(MONEY_COLS, now, prev):
            writes[_cell(roles["yoy"], col)] = arrow_pct(actual, old)
        writes[_cell(roles["yoy"], 11)] = arrow_pct(sum(now), sum(prev))
    for col in (5, 7, 9, 11):
        for role in ("target", "actual", "prior"):
            row = roles.get(role)
            if not row:
                continue
            ref = _cell(row, col)
            if _is_formula(ws[ref]):
                if role == "target":
                    writes[ref] = sum(tgt)
                elif role == "actual":
                    writes[ref] = sum(now)
                else:
                    writes[ref] = sum(prev)
    qty_header = info["qty_header"]
    if qty_header:
        for col, month in zip(QTY_NOW_COLS, months):
            writes[_cell(qty_header, col)] = f"’{yy}년 {month}월"
        for col, month in zip(QTY_PREV_COLS, months):
            writes[_cell(qty_header, col)] = f"’{py}년 {month}월"
    bulk_now = [0, 0, 0]
    bulk_prev = [0, 0, 0]
    for row, bucket, _label in info["products"]:
        if bucket == "BULK":
            continue
        cur = agg["qty_now"].get(bucket) or [0, 0, 0]
        old = agg["qty_prev"].get(bucket) or [0, 0, 0]
        if bucket != "CYL":
            for i in range(3):
                bulk_now[i] += int(cur[i])
                bulk_prev[i] += int(old[i])
        for i, col in enumerate(QTY_NOW_COLS):
            ref = _cell(row, col)
            if not _is_formula(ws[ref]):
                writes[ref] = int(cur[i])
        for i, col in enumerate(QTY_PREV_COLS):
            ref = _cell(row, col)
            if not _is_formula(ws[ref]):
                writes[ref] = int(old[i])
        for i, col in enumerate(QTY_PCT_COLS):
            writes[_cell(row, col)] = arrow_pct(cur[i], old[i])
    for row, bucket, _label in info["products"]:
        if bucket != "BULK":
            continue
        for i, col in enumerate(QTY_NOW_COLS):
            ref = _cell(row, col)
            if _is_formula(ws[ref]):
                writes[ref] = bulk_now[i]
            else:
                writes[ref] = bulk_now[i]
        for i, col in enumerate(QTY_PREV_COLS):
            ref = _cell(row, col)
            writes[ref] = bulk_prev[i]
        for i, col in enumerate(QTY_PCT_COLS):
            writes[_cell(row, col)] = arrow_pct(bulk_now[i], bulk_prev[i])
    _put_lines(writes, info.get("note_rows") or [], notes)
    _put_lines(writes, info.get("cause_rows") or [], causes)
    _put_lines(writes, info.get("outlook_rows") or [], outlooks)
    _close_located(info)
    return writes


def _shared_strings(xml: bytes) -> list[str]:
    root = ET.fromstring(xml)
    out = []
    for si in root.findall("m:si", _NS):
        parts = []
        for node in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            parts.append(node.text or "")
        out.append("".join(parts))
    return out


def _xml_text(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _append_shared(sst: str, text: str, strings: list[str]) -> tuple[str, int]:
    if text in strings:
        return sst, strings.index(text)
    idx = len(strings)
    strings.append(text)
    space = ' xml:space="preserve"' if text != text.strip() else ""
    piece = f"<si><t{space}>{_xml_text(text)}</t></si>"
    if "</sst>" not in sst:
        raise ValueError("sharedStrings에 끝 태그가 없습니다.")
    sst = sst.replace("</sst>", piece + "</sst>", 1)
    sst = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{idx + 1}"', sst, count=1)
    match = re.search(r'\bcount="(\d+)"', sst)
    if match:
        sst = sst.replace(match.group(0), f'count="{int(match.group(1)) + 1}"', 1)
    return sst, idx


def _patch_sheet(xml: str, writes: dict[str, object], strings: list[str], sst: str) -> tuple[str, str]:
    for ref, value in writes.items():
        found = re.search(rf'<c r="{re.escape(ref)}"[^>]*/>', xml)
        self_closed = found is not None
        if not found:
            found = re.search(rf'<c r="{re.escape(ref)}"[^>]*>.*?</c>', xml, re.S)
        if not found:
            row_no = re.search(r"(\d+)$", ref)
            if row_no and re.search(rf'<row r="{row_no.group(1)}"[^>]*>', xml):
                xml = re.sub(
                    rf'(<row r="{row_no.group(1)}"[^>]*>)',
                    rf'\1<c r="{ref}"/>',
                    xml,
                    count=1,
                )
                found = re.search(rf'<c r="{re.escape(ref)}"[^>]*/>', xml)
                self_closed = found is not None
            if not found:
                continue
        tag = found.group(0)
        if self_closed:
            open_tag = tag[:-2] + ">"
            inner = ""
            close = "</c>"
        else:
            split = tag.find(">")
            open_tag, inner, close = tag[: split + 1], tag[split + 1 : -4], "</c>"
        formula = "<f" in inner
        if isinstance(value, str) and not formula:
            sst, idx = _append_shared(sst, value, strings)
            open_tag = re.sub(r'\s+t="[^"]*"', "", open_tag)
            open_tag = open_tag[:-1] + ' t="s">'
            inner = f"<v>{idx}</v>"
        else:
            number = int(round(float(value or 0)))
            if formula:
                if re.search(r"<v>.*?</v>", inner):
                    inner = re.sub(r"<v>.*?</v>", f"<v>{number}</v>", inner, count=1)
                else:
                    inner = inner + f"<v>{number}</v>"
            else:
                open_tag = re.sub(r'\s+t="s"', "", open_tag)
                inner = f"<v>{number}</v>"
        xml = xml[: found.start()] + open_tag + inner + close + xml[found.end() :]
    return xml, sst


def _empty_sst() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'count="0" uniqueCount="0"></sst>'
    )


def _ensure_sst_content_type(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    if "sharedStrings.xml" in text:
        return xml
    piece = (
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    )
    return text.replace("</Types>", piece + "</Types>", 1).encode("utf-8")


def _ensure_sst_rel(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    if "sharedStrings" in text:
        return xml
    piece = (
        '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml" Id="rIdSST"/>'
    )
    return text.replace("</Relationships>", piece + "</Relationships>", 1).encode("utf-8")


def write_values(path: Path, writes: dict[str, object]) -> None:
    """셀 값만 교체한다. styles.xml 바이트는 유지한다."""
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin:
        sheet_name = "xl/worksheets/sheet1.xml"
        sst_name = "xl/sharedStrings.xml"
        names = set(zin.namelist())
        had_sst = sst_name in names
        sheet = zin.read(sheet_name).decode("utf-8")
        sst = zin.read(sst_name).decode("utf-8") if had_sst else _empty_sst()
        strings = _shared_strings(sst.encode("utf-8"))
        sheet, sst = _patch_sheet(sheet, writes, strings, sst)
        with zipfile.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == sheet_name:
                    data = sheet.encode("utf-8")
                elif item.filename == sst_name:
                    data = sst.encode("utf-8")
                elif not had_sst and item.filename == "[Content_Types].xml":
                    data = _ensure_sst_content_type(data)
                elif not had_sst and item.filename == "xl/_rels/workbook.xml.rels":
                    data = _ensure_sst_rel(data)
                zout.writestr(item, data)
            if not had_sst:
                zout.writestr(sst_name, sst.encode("utf-8"))
    os.replace(tmp, path)


def save_quarter(
    year: int,
    quarter: int,
    targets: list[float],
    sales_df: pd.DataFrame,
    notes: list[str] | None = None,
    causes: list[str] | None = None,
    outlooks: list[str] | None = None,
) -> Path:
    root = find_plan_root()
    if root is None:
        raise FileNotFoundError("사업계획 폴더를 찾지 못했습니다.")
    year_dir = root / str(int(year))
    year_dir.mkdir(parents=True, exist_ok=True)
    dest = year_dir / f"{int(quarter)}분기 사업계획.xlsx"
    if not dest.is_file():
        template = find_template()
        if template is not None and template.resolve() != dest.resolve():
            shutil.copy(template, dest)
        else:
            _create_quarter_book(dest, year, quarter)
    agg = aggregate_quarter(sales_df, year, quarter)
    writes = build_writes(dest, year, quarter, targets, agg, notes=notes, causes=causes, outlooks=outlooks)
    write_values(dest, writes)
    return dest


def read_targets(path: Path, year: int, quarter: int) -> list[int] | None:
    if not path.is_file():
        return None
    info = locate_sheet(path)
    try:
        row = info["roles"].get("target")
        header = info["money_header"]
        if not row or not header:
            return None
        ws = info["ws"]
        months = []
        for col in MONEY_COLS:
            text = str(ws.cell(header, col).value or "")
            found = re.search(r"(\d+)", text)
            months.append(int(found.group(1)) if found else 0)
        if tuple(months) != QUARTERS[int(quarter)]:
            return None
        values = []
        for col in MONEY_COLS:
            cell = ws.cell(row, col)
            if _is_formula(cell):
                values.append(0)
                continue
            try:
                values.append(int(round(float(cell.value or 0))))
            except (TypeError, ValueError):
                values.append(0)
        return values
    finally:
        _close_located(info)


def _money_table(year: int, months: tuple[int, ...], targets: list[int], agg: dict) -> pd.DataFrame:
    now = agg["money_now"]
    prev = agg["money_prev"]
    cols = [f"{m}월" for m in months] + ["합계"]
    rows = {
        "목표": targets + [sum(targets)],
        "실적": now + [sum(now)],
        "달성률": [arrow_pct(a, t) for a, t in zip(now, targets)] + [arrow_pct(sum(now), sum(targets))],
        "전년 매출": prev + [sum(prev)],
        "전년 비교": [arrow_pct(a, b) for a, b in zip(now, prev)] + [arrow_pct(sum(now), sum(prev))],
    }
    frame = pd.DataFrame(rows).T
    frame.columns = cols
    frame.index.name = f"{year}년"
    return frame.reset_index().rename(columns={"index": "구분"})


def _qty_table(year: int, months: tuple[int, ...], agg: dict) -> pd.DataFrame:
    yy = str(year)[2:]
    py = str(year - 1)[2:]
    labels = [
        ("LO2", "LO2"),
        ("LN2", "LN2"),
        ("LAR", "LAR"),
        ("LCO2", "LCO2"),
        ("LPG", "LPG"),
        ("BULK", "Bulk 합계"),
        ("CYL", "Gas Cylinder"),
    ]
    bulk_now = [0, 0, 0]
    bulk_prev = [0, 0, 0]
    for key in ("LO2", "LN2", "LAR", "LCO2", "LPG"):
        for i in range(3):
            bulk_now[i] += agg["qty_now"][key][i]
            bulk_prev[i] += agg["qty_prev"][key][i]
    records = []
    for key, label in labels:
        if key == "BULK":
            cur, old = bulk_now, bulk_prev
        else:
            cur, old = agg["qty_now"][key], agg["qty_prev"][key]
        row = {"구분": label}
        for i, month in enumerate(months):
            row[f"’{yy}년 {month}월"] = cur[i]
            row[f"’{py}년 {month}월"] = old[i]
            row[f"전년 대비 {month}월"] = arrow_pct(cur[i], old[i])
        records.append(row)
    return pd.DataFrame(records)


def _preview_source(year: int, quarter: int) -> Path | None:
    dest = quarter_path(year, quarter)
    if dest is not None and dest.is_file():
        return dest
    return find_template()


def _preview_token(
    src: Path,
    year: int,
    quarter: int,
    targets: list[int],
    agg: dict,
    notes: list[str] | None = None,
    causes: list[str] | None = None,
    outlooks: list[str] | None = None,
) -> tuple:
    return (
        str(src),
        int(src.stat().st_mtime_ns),
        int(year),
        int(quarter),
        tuple(int(v) for v in targets),
        tuple(int(v) for v in (agg.get("money_now") or [])),
        tuple(int(v) for v in (agg.get("money_prev") or [])),
        tuple(notes or ()),
        tuple(causes or ()),
        tuple(outlooks or ()),
    )


def _ql_png(xlsx: Path) -> bytes:
    """맥 미리보기로 엑셀 시트를 원본 모습 그대로 이미지로 만든다."""
    with tempfile.TemporaryDirectory(prefix="bp_ql_") as td:
        proc = subprocess.run(
            ["qlmanage", "-t", "-s", "1800", "-o", td, str(xlsx)],
            check=False,
            timeout=40,
            capture_output=True,
        )
        pngs = sorted(Path(td).glob("*.png"))
        if proc.returncode != 0 or not pngs:
            detail = (proc.stderr or proc.stdout or b"").decode("utf-8", "replace").strip()
            raise RuntimeError(detail or "엑셀 원본 이미지를 만들지 못했습니다.")
        return pngs[0].read_bytes()


def sheet_preview_png(
    year: int,
    quarter: int,
    targets: list[int],
    sales_df: pd.DataFrame,
    notes: list[str] | None = None,
    causes: list[str] | None = None,
    outlooks: list[str] | None = None,
) -> bytes | None:
    """저장과 같은 숫자만 임시 파일에 넣고, 원본 서식 이미지를 돌려준다. 저장 파일은 그대로 둔다."""
    src = _preview_source(year, quarter)
    if src is None:
        return None
    frame = sales_df if isinstance(sales_df, pd.DataFrame) else pd.DataFrame()
    agg = _cached_aggregate(frame, year, quarter)
    token = _preview_token(src, year, quarter, targets, agg, notes, causes, outlooks)
    try:
        cached = st.session_state.get("_bp_sheet_preview")
    except Exception:
        cached = None
    if isinstance(cached, tuple) and len(cached) == 2 and cached[0] == token and isinstance(cached[1], (bytes, bytearray)):
        return bytes(cached[1])
    with tempfile.TemporaryDirectory(prefix="bp_fill_") as td:
        copy = Path(td) / "preview.xlsx"
        shutil.copy(src, copy)
        write_values(copy, build_writes(copy, year, quarter, targets, agg, notes=notes, causes=causes, outlooks=outlooks))
        png = _ql_png(copy)
    try:
        st.session_state["_bp_sheet_preview"] = (token, png)
    except Exception:
        pass
    return png


def _cached_targets(path: Path | None, year: int, quarter: int) -> list[int] | None:
    if path is None or not path.is_file():
        return None
    try:
        sig = (str(path), path.stat().st_mtime_ns, int(year), int(quarter))
        bag = st.session_state.get("_bp_target_cache")
        if not isinstance(bag, dict):
            bag = {}
        if sig in bag:
            return bag[sig]
    except Exception:
        return read_targets(path, year, quarter)
    values = read_targets(path, year, quarter)
    try:
        bag[sig] = values
        st.session_state["_bp_target_cache"] = bag
    except Exception:
        pass
    return values


def _cached_aggregate(df: pd.DataFrame, year: int, quarter: int) -> dict:
    frame = df if isinstance(df, pd.DataFrame) else pd.DataFrame()
    sig = (int(year), int(quarter), id(frame), int(getattr(frame, "shape", (0,))[0]))
    try:
        bag = st.session_state.get("_bp_agg_cache")
        if isinstance(bag, dict) and sig in bag:
            return bag[sig]
    except Exception:
        bag = None
    agg = aggregate_quarter(frame, year, quarter)
    if isinstance(bag, dict):
        if len(bag) > 8:
            bag.clear()
        bag[sig] = agg
        try:
            st.session_state["_bp_agg_cache"] = bag
        except Exception:
            pass
    else:
        try:
            st.session_state["_bp_agg_cache"] = {sig: agg}
        except Exception:
            pass
    return agg


def _sync_lines(key: str, lines: list[str]) -> None:
    st.session_state[key] = list(lines)
    prefix = f"{key}__"
    for name in list(st.session_state.keys()):
        if str(name).startswith(prefix):
            del st.session_state[name]
    for i, text in enumerate(lines):
        st.session_state[f"{prefix}{i}"] = text


def _line_editor(key: str, max_lines: int, add_label: str) -> list[str]:
    lines = list(st.session_state.get(key) or [])
    remove_at = None
    for i in range(len(lines)):
        left, right = st.columns([10, 1.3])
        with left:
            lines[i] = st.text_input("줄", key=f"{key}__{i}", label_visibility="collapsed")
        with right:
            if st.button("삭제", key=f"{key}_del_{i}"):
                remove_at = i
    if remove_at is not None:
        lines.pop(remove_at)
        _sync_lines(key, lines)
        st.rerun(scope="fragment")
    st.session_state[key] = lines
    if len(lines) < int(max_lines):
        if st.button(add_label, key=f"{key}_add"):
            lines.append("")
            _sync_lines(key, lines)
            st.rerun(scope="fragment")
    elif int(max_lines) > 0:
        st.caption("엑셀에 더 쓸 줄이 없습니다.")
    return lines


def _rest_after_goal(notes: list[str]) -> list[str]:
    if notes and "매출목표" in str(notes[0]):
        return list(notes[4:])
    return list(notes)


@st.fragment
def render_biz_plan_tab(sales_df: pd.DataFrame) -> None:
    """로컬 분기 사업분석. 연도·분기와 목표를 고르고 두 표를 보여 준다."""
    st.markdown("##### 분기 사업분석")
    root = find_plan_root()
    if root is None:
        st.error("사업계획 폴더를 찾지 못했습니다. 이 맥의 Google Drive 사업계획 폴더가 있어야 합니다.")
        return
    today = date.today()
    if "bp_year" not in st.session_state:
        st.session_state["bp_year"] = int(today.year)
    if "bp_quarter" not in st.session_state:
        st.session_state["bp_quarter"] = (today.month - 1) // 3 + 1
    frame = sales_df if isinstance(sales_df, pd.DataFrame) else pd.DataFrame()
    edit_col, view_col = st.columns([1, 1.15])
    with edit_col:
        pick_year, pick_quarter = st.columns(2)
        with pick_year:
            year = int(st.number_input("연도", min_value=2000, max_value=2100, step=1, key="bp_year"))
        with pick_quarter:
            quarter = int(
                st.selectbox(
                    "분기",
                    [1, 2, 3, 4],
                    format_func=lambda q: f"{q}분기 ({QUARTERS[q][0]}–{QUARTERS[q][2]}월)",
                    key="bp_quarter",
                )
            )
        months = QUARTERS[quarter]
        dest = quarter_path(year, quarter)
        stored = _cached_targets(dest, year, quarter)
        if stored is None:
            template = find_template()
            stored = _cached_targets(template, year, quarter) if template else None
        if stored is None:
            stored = [0, 0, 0]
        boxes = st.columns(3)
        targets: list[int] = []
        for i, month in enumerate(months):
            key = f"bp_tgt_{year}_{quarter}_{month}"
            if key not in st.session_state:
                st.session_state[key] = int(stored[i])
            with boxes[i]:
                targets.append(
                    int(
                        st.number_input(
                            f"{month}월 목표 (백만원, 부가세 별도)",
                            min_value=0,
                            step=1,
                            key=key,
                        )
                    )
                )
        agg = _cached_aggregate(frame, year, quarter)
        source = dest if dest is not None and dest.is_file() else find_template()
        stamp = int(source.stat().st_mtime_ns) if source is not None and source.is_file() else 0
        text_sig = (int(year), int(quarter), str(source or ""), stamp)
        if st.session_state.get("_bp_text_sig") != text_sig:
            loaded = read_text_blocks(source)
            st.session_state["_bp_text_loaded"] = loaded
            st.session_state["_bp_text_sig"] = text_sig
            st.session_state[f"_bp_goal_auto_{year}_{quarter}"] = None
            st.session_state[f"_bp_cause_sel_{year}_{quarter}"] = None
            annual_key = f"bp_annual_{year}"
            if annual_key not in st.session_state and loaded.get("annual"):
                st.session_state[annual_key] = int(loaded["annual"])
        loaded = st.session_state.get("_bp_text_loaded") or {}
        note_cap = len(loaded.get("note_rows") or []) or 8
        cause_cap = len(loaded.get("cause_rows") or []) or 11
        outlook_cap = len(loaded.get("outlook_rows") or []) or 16
        annual_key = f"bp_annual_{year}"
        if annual_key not in st.session_state:
            st.session_state[annual_key] = int(loaded.get("annual") or 0)
        annual = int(
            st.number_input(
                "연간 매출목표 (백만원, 부가세 별도)",
                min_value=0,
                step=1,
                key=annual_key,
            )
        )
        auto = sales_goal_lines(year, quarter, annual, frame)
        notes_key = f"bp_notes_{year}_{quarter}"
        if notes_key not in st.session_state:
            _sync_lines(notes_key, auto + _rest_after_goal(list(loaded.get("notes") or [])))
        prev_auto = st.session_state.get(f"_bp_goal_auto_{year}_{quarter}")
        notes_now = list(st.session_state.get(notes_key) or [])
        refresh_goal = bool(st.session_state.pop(f"_bp_goal_refresh_{year}_{quarter}", False))
        if prev_auto is None or refresh_goal or notes_now[:4] == list(prev_auto):
            if prev_auto is None:
                rest = _rest_after_goal(list(loaded.get("notes") or []))
            else:
                rest = notes_now[4:]
            merged = list(auto) + rest
            if merged != notes_now:
                _sync_lines(notes_key, merged)
            st.session_state[f"_bp_goal_auto_{year}_{quarter}"] = list(auto)
        st.markdown("**목표 / 달성**")
        st.caption("단위: 백만원, 부가세 별도. 목표는 위 칸에서 바꿉니다.")
        st.dataframe(_money_table(year, months, targets, agg), use_container_width=True, hide_index=True)
        st.caption("1번은 매출이 있으면 다시 채웁니다. 줄은 추가·삭제할 수 있고, 오른쪽 이미지와 저장 파일에 들어갑니다.")
        notes = _line_editor(notes_key, note_cap, "목표 문구 줄 추가")
        if st.button("1번을 실적으로 다시 채우기", key=f"bp_goal_btn_{year}_{quarter}"):
            st.session_state[f"_bp_goal_refresh_{year}_{quarter}"] = True
            st.rerun(scope="fragment")
        st.markdown("**실적 / 분석**")
        st.caption("벌크는 톤, 벌크가 아닌 실린더는 병입니다.")
        st.dataframe(_qty_table(year, months, agg), use_container_width=True, hide_index=True)
        st.markdown("**원인 / 분석**")
        st.caption(
            f"전년 같은 분기보다 매출이 늘어난 거래처 상위 30입니다. 고르면 원인/분석에 들어가고, 엑셀에는 {cause_cap}줄까지 씁니다."
        )
        ranked = growth_top(frame, year, quarter, 30)
        cause_key = f"bp_causes_{year}_{quarter}"
        if cause_key not in st.session_state:
            _sync_lines(cause_key, list(loaded.get("causes") or []))
        if ranked:
            show = pd.DataFrame(
                {
                    "순위": list(range(1, len(ranked) + 1)),
                    "거래처": [item[0] for item in ranked],
                    "올해(만)": [int(round(item[1] / 10000.0)) for item in ranked],
                    "전년(만)": [int(round(item[2] / 10000.0)) for item in ranked],
                    "증가(만)": [int(round(item[3] / 10000.0)) for item in ranked],
                }
            )
            st.dataframe(show, use_container_width=True, hide_index=True)
            options = [item[0] for item in ranked]
            picked = st.multiselect("원인/분석에 넣을 거래처", options, key=f"bp_grow_{year}_{quarter}")
            sel_key = f"_bp_cause_sel_{year}_{quarter}"
            prev_sel = st.session_state.get(sel_key)
            if prev_sel is None:
                st.session_state[sel_key] = tuple(picked)
            elif tuple(picked) != tuple(prev_sel):
                written = cause_lines(list(picked), ranked, year, quarter)[:cause_cap]
                _sync_lines(cause_key, written)
                st.session_state[sel_key] = tuple(picked)
                st.rerun(scope="fragment")
        else:
            st.caption("비교할 매출 자료가 없습니다.")
        causes = _line_editor(cause_key, cause_cap, "원인 줄 추가")
        st.markdown("**전망**")
        outlook_key = f"bp_outlook_{year}_{quarter}"
        if outlook_key not in st.session_state:
            _sync_lines(outlook_key, list(loaded.get("outlooks") or []))
        outlooks = _line_editor(outlook_key, outlook_cap, "전망 줄 추가")
        year_dir = root / str(year)
        st.caption(f"저장 위치 {year_dir / f'{quarter}분기 사업계획.xlsx'}")
        if st.button("저장", type="primary", key="bp_save"):
            try:
                saved = save_quarter(year, quarter, targets, frame, notes=notes, causes=causes, outlooks=outlooks)
                st.success(f"저장했습니다. {saved}")
            except Exception as exc:
                st.error(f"저장하지 못했습니다. {exc}")
    with view_col:
        png = None
        try:
            with st.spinner("엑셀 원본을 맞추고 있습니다."):
                png = sheet_preview_png(
                    year, quarter, targets, frame, notes=notes, causes=causes, outlooks=outlooks
                )
        except Exception as exc:
            st.caption(f"엑셀 원본을 보여 주지 못했습니다. {exc}")
        if png:
            st.image(png, use_container_width=True)
            st.caption("엑셀 원본입니다. 숫자와 문구는 지금 입력과 같고, 저장을 누르면 파일에 남습니다.")
        elif _preview_source(year, quarter) is None:
            st.caption("보여 줄 엑셀 원본이 없습니다.")
