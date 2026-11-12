"""분기 사업분석 — 로컬 전용.

목표/달성·실적/분석 두 표를 매출에서 취합하고, 목표만 직접 고친다.
저장은 양식 파일의 셀 값만 바꾸고 서식 XML은 그대로 둔다.
"""
from __future__ import annotations

import os
import re
import shutil
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


def find_template() -> Path | None:
    root = find_plan_root()
    cands = [Path.home() / "Desktop" / "26Q3 사업분석_draft 261016.xlsx"]
    if root is not None:
        cands.append(root / "2026" / "26Q3 사업분석_draft 261016.xlsx")
    for path in cands:
        if path.is_file():
            return path
    return None


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
    return {
        "sheet": ws.title,
        "money_header": money_header,
        "qty_header": qty_header,
        "roles": roles,
        "products": products,
        "title": title,
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


def build_writes(path: Path, year: int, quarter: int, targets: list[float], agg: dict) -> dict[str, object]:
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
    piece = f"<si><t>{_xml_text(text)}</t></si>"
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
        found = re.search(rf'<c r="{re.escape(ref)}"[^>]*>.*?</c>', xml, re.S)
        self_closed = False
        if not found:
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
            if 't="s"' not in open_tag:
                open_tag = open_tag[:-1] + ' t="s">'
            if re.search(r"<v>.*?</v>", inner):
                inner = re.sub(r"<v>.*?</v>", f"<v>{idx}</v>", inner, count=1)
            else:
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


def write_values(path: Path, writes: dict[str, object]) -> None:
    """셀 값만 교체한다. styles.xml 바이트는 유지한다."""
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin:
        sheet_name = "xl/worksheets/sheet1.xml"
        sst_name = "xl/sharedStrings.xml"
        sheet = zin.read(sheet_name).decode("utf-8")
        sst = zin.read(sst_name).decode("utf-8")
        strings = _shared_strings(sst.encode("utf-8"))
        sheet, sst = _patch_sheet(sheet, writes, strings, sst)
        with zipfile.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == sheet_name:
                    data = sheet.encode("utf-8")
                elif item.filename == sst_name:
                    data = sst.encode("utf-8")
                zout.writestr(item, data)
    os.replace(tmp, path)


def save_quarter(year: int, quarter: int, targets: list[float], sales_df: pd.DataFrame) -> Path:
    root = find_plan_root()
    if root is None:
        raise FileNotFoundError("사업계획 폴더를 찾지 못했습니다.")
    year_dir = root / str(int(year))
    year_dir.mkdir(parents=True, exist_ok=True)
    dest = year_dir / f"{int(quarter)}분기 사업계획.xlsx"
    if not dest.is_file():
        template = find_template()
        if template is None:
            raise FileNotFoundError("사업분석 양식 파일을 찾지 못했습니다.")
        shutil.copy(template, dest)
    agg = aggregate_quarter(sales_df, year, quarter)
    writes = build_writes(dest, year, quarter, targets, agg)
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
    agg = _cached_aggregate(sales_df if isinstance(sales_df, pd.DataFrame) else pd.DataFrame(), year, quarter)
    st.markdown("**목표 / 달성**")
    st.caption("단위: 백만원, 부가세 별도. 목표는 위 칸에서 바꿉니다.")
    st.dataframe(_money_table(year, months, targets, agg), use_container_width=True, hide_index=True)
    st.markdown("**실적 / 분석**")
    st.caption("벌크는 톤, 벌크가 아닌 실린더는 병입니다.")
    st.dataframe(_qty_table(year, months, agg), use_container_width=True, hide_index=True)
    year_dir = root / str(year)
    st.caption(f"저장 위치 {year_dir / f'{quarter}분기 사업계획.xlsx'}")
    if st.button("저장", type="primary", key="bp_save"):
        try:
            saved = save_quarter(year, quarter, targets, sales_df if isinstance(sales_df, pd.DataFrame) else pd.DataFrame())
            st.success(f"저장했습니다. {saved}")
        except Exception as exc:
            st.error(f"저장하지 못했습니다. {exc}")
