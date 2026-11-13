"""사업계획서 — 로컬 전용.

PPT 16페이지를 방마다 나누고, 저장은 그 페이지의 숫자와 그래프만 바꾼다.
가스코아산은 아산, 그 외는 화성. 금액은 부가세 포함.
"""
from __future__ import annotations

import io
import json
import re
import shutil
import subprocess
import unicodedata
import zipfile
from functools import lru_cache
from pathlib import Path
from xml.sax.saxutils import escape, unescape

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from pptx import Presentation

SALES_DIR = Path(__file__).resolve().parent / "uploaded_cache" / "sales"
PAGES = tuple(range(1, 17))
_AT = re.compile(r"(<a:t(?:\s[^>]*)?>)([^<]*)(</a:t>)")
_YEAR_TOKEN = {
    "2023": "\ue000",
    "2024": "\ue001",
    "2025": "\ue002",
    "2026": "\ue003",
    "23년": "\ue004",
    "24년": "\ue005",
    "25년": "\ue006",
    "26년": "\ue007",
}
PLAN_ASAN = (457, 429, 470, 490, 475, 481, 494, 450, 437, 509, 498, 505)  # 1월→12월
PLAN_HWA = (1081, 942, 1054, 1205, 1333, 1508, 1764, 1638, 1535, 1365, 1162, 1178)
PLAN_CHART = {"아산": 5695, "화성": 15764, "전체": 21461}
PRIOR_PLAN_TOTAL = {2024: 21085, 2025: 21567, 2026: 21461}
PLAN_TONS = {"산소": 6502, "질소": 11381, "알곤": 7491, "탄산": 50688, "LPG": 319}
POLICY_TITLE = "영업방침 및 영업목표"
REVIEW_MARK = "사업계획 대비"
STRATEGY_TITLE = "년도 세부 영업전략"
POLICY_DEFAULTS = ("영업방침 및 영업목표",)
REVIEW_DEFAULTS = (
    "총 판매현황",
    "세부실적 / 품목별 판매현황",
    "세부실적 / 품목별 BULK 현황",
    "25년도 담당별 실적",
    "분석 / 향후 추진",
)
STRATEGY_DEFAULTS = (
    "신규영업 및 투자계획",
    "지역별 산업단지 영업계획",
    "악성 채권현황 및 업체 관리방안",
)
DIRECTIVE_LABEL = "영업 방침"
FOCUS_LABEL = "중점운영 목표"
TASK_LABEL = "추진 과제"
_GOAL_MARK = {"directive": "영업방침", "focus": "중점운영목표", "task": "추진과제"}
DIRECTIVE_DEFAULTS = (
    "신규영업 강화 / 수익성 위주 영업",
    "기술영업 / 관리강화",
    "고객만족 영업",
    "거래처 관리",
)
FOCUS_DEFAULTS = (
    "직 거래처  개발",
    "단가인상 / 채권관리",
    "안전 / 전문교육 실시",
    "기존 거래처 관리(주기적 방문)",
)
TASK_DEFAULTS = (
    "신규 산업단지 집중 영업",
    "신규 충전소 / 판매점 선점",
    "단가인상 적기 반영",
    "부실채권 관리 / 여신점검",
    "고객응대 / 서비스 강화",
    "주기적인 안전점검(분기 1회)",
)
_STRATEGY_BULLET = (
    '<a:p><a:pPr marL="0" lvl="0" indent="0"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc>'
    '<a:spcBef><a:spcPts val="864"/></a:spcBef><a:buNone/></a:pPr>'
    '<a:r><a:rPr lang="en-US" altLang="ko-KR" sz="2700" dirty="0"><a:solidFill><a:srgbClr val="242529"/></a:solidFill>'
    '<a:latin typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/>'
    '<a:ea typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/></a:rPr>'
    "<a:t>   - </a:t></a:r>"
    '<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="2700" dirty="0"><a:solidFill><a:srgbClr val="242529"/></a:solidFill>'
    '<a:latin typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/>'
    '<a:ea typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/></a:rPr>'
    "<a:t>신규영업 및 투자계획</a:t></a:r>"
    '<a:endParaRPr lang="en-US" altLang="ko-KR" sz="2700" dirty="0"><a:solidFill><a:srgbClr val="242529"/></a:solidFill>'
    '<a:latin typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/>'
    '<a:ea typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/></a:endParaRPr></a:p>'
)
BULK_ROWS = ("O2", "N2", "AR", "CO2", "LPG(T/K)")
CYL_ROWS = (
    "O2(L.G.C)",
    "N2(L.G.C)",
    "Ar(L.G.C)",
    "O2",
    "N2",
    "Ar",
    "CO2",
    "고순도",
    "Mix",
    "He",
    "LPG 외 기타",
)
_TON_OF = {"산소": "O2", "질소": "N2", "알곤": "AR", "탄산": "CO2", "LPG": "LPG(T/K)"}
_CLIENT_GAS = {"LO2": "O2", "LN2": "N2", "LAR": "AR", "LCO2": "CO2", "LPG": "LPG"}
PAGE8_PRODUCTS = ("LO2", "LN2", "LAR", "LCO2", "LPG")
_DELTA_RED = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'
_DELTA_DARK = '<a:solidFill><a:schemeClr val="bg2"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>'


def _nfc(text: str) -> str:
    return unicodedata.normalize("NFC", str(text or ""))


def find_plan_root() -> Path | None:
    """Drive › 다른 컴퓨터 › 내 컴퓨터 (1) › Desktop › 업무 › 사업계획."""
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
            if "내 컴퓨터 (1)" in _nfc(comp.name):
                return comp / "Desktop" / "업무" / "사업계획"
    return None


def find_template() -> Path | None:
    desktop = Path.home() / "Desktop"
    if not desktop.is_dir():
        return None
    for path in desktop.iterdir():
        if path.suffix.lower() == ".pptx" and "20260119" in path.name and not path.name.startswith("~$"):
            return path
    return None


def proposal_name(year: int) -> str:
    """2026 → 26년 사업계획서.pptx."""
    return f"{int(year) % 100:02d}년 사업계획서.pptx"


def resolve_proposal_path(root: Path, year: int) -> Path:
    """같은 이름의 파일이 있으면 그 파일을 이어서 고친다."""
    name = proposal_name(year)
    folder = root / str(int(year))
    for candidate in (folder / name, root / name):
        if candidate.is_file():
            return candidate
    return folder / name


def proposal_path(year: int) -> Path | None:
    root = find_plan_root()
    if root is None:
        return None
    return resolve_proposal_path(root, year)


def _num(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.replace(r"[^\d.-]", "", regex=True)
    return pd.to_numeric(cleaned, errors="coerce").fillna(0).round()


def _month_of(value: object) -> int | None:
    text = str(value or "").strip()
    matched = re.match(r"(\d{4})-(\d{1,2})-\d{1,2}", text)
    if matched:
        month = int(matched.group(2))
        return month if 1 <= month <= 12 else None
    matched = re.match(r"(\d{1,2})/", text)
    if not matched:
        return None
    month = int(matched.group(1))
    return month if 1 <= month <= 12 else None


def is_asan(name: object) -> bool:
    text = re.sub(r"^[zZ]", "", str(name or "").strip()).strip()
    return text.startswith("가스코아산")


def _year_files(year: int) -> list[Path]:
    annual = SALES_DIR / f"{int(year)}.csv"
    if annual.is_file():
        return [annual]
    return sorted(SALES_DIR.glob(f"{int(year)}[0-1][0-9].csv"))


@lru_cache(maxsize=8)
def load_year(year: int) -> pd.DataFrame:
    """날짜가 있는 매출만. 합계액이 있으면 부가세 포함, 없으면 매출액×1.1."""
    frames: list[pd.DataFrame] = []
    for path in _year_files(year):
        frame = pd.read_csv(path, encoding="utf-8-sig")
        frame.columns = frame.columns.astype(str).str.strip()
        if "거래처" not in frame.columns or "품목명" not in frame.columns:
            continue
        if "합계액" in frame.columns:
            money = _num(frame["합계액"])
        elif "매출액" in frame.columns:
            money = _num(frame["매출액"]) * 1.1
        else:
            continue
        qty = _num(frame["출고량"]) if "출고량" in frame.columns else 0.0
        months = frame["매출일"].map(_month_of) if "매출일" in frame.columns else None
        if months is None:
            continue
        part = pd.DataFrame(
            {
                "거래처": frame["거래처"].fillna("").astype(str).str.strip(),
                "품목명": frame["품목명"].fillna("").astype(str).str.strip(),
                "월": months,
                "금액": money,
                "수량": qty,
            }
        )
        part = part[part["월"].between(1, 12)]
        part = part[~part["품목명"].str.contains("이월", na=False)]
        frames.append(part)
    if not frames:
        return pd.DataFrame(columns=["거래처", "품목명", "월", "금액", "수량"])
    return pd.concat(frames, ignore_index=True)


def _through(frame: pd.DataFrame, month: int) -> pd.DataFrame:
    if frame.empty:
        return frame
    return frame[frame["월"] <= int(month)]


def plant_won(year: int, month: int) -> dict[str, dict[int, float]]:
    frame = _through(load_year(year), month)
    out = {"아산": {m: 0.0 for m in range(1, 13)}, "화성": {m: 0.0 for m in range(1, 13)}}
    if frame.empty:
        out["전체"] = {m: 0.0 for m in range(1, 13)}
        return out
    frame = frame.copy()
    frame["사업장"] = frame["거래처"].map(lambda name: "아산" if is_asan(name) else "화성")
    grouped = frame.groupby(["사업장", "월"])["금액"].sum()
    for (plant, mon), value in grouped.items():
        out[plant][int(mon)] = float(value)
    out["전체"] = {m: out["아산"][m] + out["화성"][m] for m in range(1, 13)}
    return out


def millions(month_map: dict[int, float], month: int) -> list[int]:
    """1월→12월. 고른 달 이후는 0. 각 달은 백만원 반올림."""
    return [int(round(month_map.get(m, 0.0) / 1_000_000)) if m <= int(month) else 0 for m in range(1, 13)]


def annual_million(month_map: dict[int, float], month: int) -> int:
    """연간 합계는 달별 반올림을 더하지 않고, 원 합계를 한 번에 백만원으로 맞춘다."""
    total = sum(month_map.get(m, 0.0) for m in range(1, int(month) + 1))
    return int(round(total / 1_000_000))


def _item_key(name: str) -> str:
    text = str(name or "")
    upper = text.upper().replace(" ", "")
    bulk = "BULK" in upper or "벌크" in text
    mix = upper.startswith("MIX")
    lgc = "LGC" in upper

    def has(*keys: str) -> bool:
        return any(key in upper or key in text for key in keys)

    if bulk and not mix:
        if has("CO2", "탄산"):
            return "B:CO2"
        if has("LPG"):
            return "B:LPG"
        if has("N2", "질소"):
            return "B:N2"
        if has("AR", "아르곤", "아르"):
            return "B:AR"
        if has("O2", "산소"):
            return "B:O2"
        return "C:기타"
    if mix:
        return "C:Mix"
    if lgc:
        if has("CO2", "탄산"):
            return "C:CO2"
        if has("N2", "질소"):
            return "C:N2L"
        if has("AR", "아르"):
            return "C:ARL"
        if has("O2", "산소"):
            return "C:O2L"
    he = bool(re.search(r"(^|[^A-Z])HE([^A-Z]|$)", upper))
    if ("UHP" in upper or "고순도" in text) and not he and "AIR" not in upper and not lgc:
        return "C:UHP"
    if he:
        return "C:He"
    if has("CO2", "탄산"):
        return "C:CO2"
    if has("LPG"):
        return "C:기타"
    if has("N2", "질소"):
        return "C:N2"
    if has("AR", "아르") and not has("CO2"):
        return "C:Ar"
    if has("O2", "산소") and not has("CO2"):
        return "C:O2"
    return "C:기타"


def _qty_kg(name: str, qty: float) -> float:
    upper = str(name or "").upper()
    if "BULK" in upper and ("LITER" in upper or "리터" in str(name)):
        return float(qty) * 0.808
    return float(qty)


def item_totals(year: int, month: int) -> dict[str, dict[str, float]]:
    frame = _through(load_year(year), month)
    rows = {key: {"금액": 0.0, "수량": 0.0} for key in ("B:O2", "B:N2", "B:AR", "B:CO2", "B:LPG", "C:O2L", "C:N2L", "C:ARL", "C:O2", "C:N2", "C:Ar", "C:CO2", "C:UHP", "C:Mix", "C:He", "C:기타")}
    if frame.empty:
        return rows
    for name, money, qty in zip(frame["품목명"], frame["금액"], frame["수량"], strict=False):
        key = _item_key(str(name))
        rows[key]["금액"] += float(money)
        rows[key]["수량"] += _qty_kg(str(name), float(qty))
    return rows


def _fmt_int(value: float) -> str:
    return f"{int(round(value)):,}"


def _fmt_pct(value: float) -> str:
    tenths = int(value * 10)
    whole, frac = divmod(abs(tenths), 10)
    sign = "-" if tenths < 0 else ""
    if frac == 0:
        return f"{sign}{whole}%"
    return f"{sign}{whole}.{frac}%"


def _line(cur: dict[str, float], prev: dict[str, float], total: float, months: int) -> dict[str, str]:
    divisor = max(int(months), 1)
    avg = cur["수량"] / divisor
    prev_avg = prev["수량"] / divisor
    delta = (avg - prev_avg) / prev_avg * 100 if prev_avg else 0.0
    share = cur["금액"] / total * 100 if total else 0.0
    return {
        "수량": _fmt_int(cur["수량"]),
        "월평균": _fmt_int(avg),
        "월평균 증감": _fmt_pct(delta),
        "금액": _fmt_int(cur["금액"]),
        "전년대비": _fmt_int(cur["금액"] - prev["금액"]),
        "비중": _fmt_pct(share),
    }


def item_sheet(year: int, month: int) -> list[dict[str, str]]:
    cur = item_totals(year, month)
    prev = item_totals(year - 1, month)
    keys = {
        "O2": "B:O2",
        "N2": "B:N2",
        "AR": "B:AR",
        "CO2": "B:CO2",
        "LPG(T/K)": "B:LPG",
        "O2(L.G.C)": "C:O2L",
        "N2(L.G.C)": "C:N2L",
        "Ar(L.G.C)": "C:ARL",
        "O2c": "C:O2",
        "N2c": "C:N2",
        "Arc": "C:Ar",
        "CO2c": "C:CO2",
        "고순도": "C:UHP",
        "Mix": "C:Mix",
        "He": "C:He",
        "LPG 외 기타": "C:기타",
    }
    total = sum(bucket["금액"] for bucket in cur.values())
    order = [
        ("Bulk", "O2", "O2"),
        ("", "N2", "N2"),
        ("", "AR", "AR"),
        ("", "CO2", "CO2"),
        ("", "LPG(T/K)", "LPG(T/K)"),
        ("Cylinder", "O2(L.G.C)", "O2(L.G.C)"),
        ("", "N2(L.G.C)", "N2(L.G.C)"),
        ("", "Ar(L.G.C)", "Ar(L.G.C)"),
        ("", "O2", "O2c"),
        ("", "N2", "N2c"),
        ("", "Ar", "Arc"),
        ("", "CO2", "CO2c"),
        ("", "고순도", "고순도"),
        ("", "Mix", "Mix"),
        ("", "He", "He"),
        ("", "LPG 외 기타", "LPG 외 기타"),
    ]
    body = []
    for group, label, key in order:
        body.append({"구분": group, "품목": label, **_line(cur[keys[key]], prev[keys[key]], total, month)})
    body.append({"구분": "총합계", "품목": "", "수량": _fmt_int(total), "월평균": "", "월평균 증감": "", "금액": "", "전년대비": "", "비중": ""})
    return body


def tons_of(year: int, month: int) -> dict[str, int]:
    totals = item_totals(year, month)
    named = {"산소": "B:O2", "질소": "B:N2", "알곤": "B:AR", "탄산": "B:CO2", "LPG": "B:LPG"}
    out = {label: int(totals[key]["수량"] / 1000) for label, key in named.items()}
    out["합계"] = sum(out.values())
    return out


@lru_cache(maxsize=8)
def _client_bulk_index(year: int, month: int) -> dict[str, dict[str, float]]:
    """접힌 거래처 이름 → 벌크 품목별 kg. 한 해에 한 번만 만든다."""
    frame = _through(load_year(year), month)
    index: dict[str, dict[str, float]] = {}
    if frame.empty:
        return index
    for client, item, qty in zip(frame["거래처"], frame["품목명"], frame["수량"], strict=False):
        kind = _item_key(str(item))
        if not kind.startswith("B:"):
            continue
        gas = kind[2:]
        folded = re.sub(r"\s+", "", str(client))
        bucket = index.setdefault(folded, {})
        bucket[gas] = bucket.get(gas, 0.0) + _qty_kg(str(item), float(qty))
    return index


def client_kg(year: int, month: int, client: str, product: str) -> float | None:
    if not client or "신규" in client or client.strip() in {"-", "—"}:
        return None
    key = _CLIENT_GAS.get(re.sub(r"[^A-Za-z0-9]", "", product))
    if key is None:
        return None
    needle = re.sub(r"\s+", "", client.strip())
    if not needle:
        return 0.0
    total = 0.0
    found = False
    for name, gases in _client_bulk_index(int(year), int(month)).items():
        if needle in name:
            found = True
            total += gases.get(key, 0.0)
    return total if found else 0.0


ADDRESS_PATH = Path(__file__).resolve().parent / "uploaded_cache" / "address.csv"


def _fold_name(value: object) -> str:
    text = re.sub(r"\s+", "", str(value or ""))
    return re.sub(r"^[zZ]", "", text)


@lru_cache(maxsize=1)
def _staff_pairs() -> tuple[tuple[str, str], ...]:
    """접힌 업체명 → 영업담당자. 긴 이름부터 맞춘다."""
    if not ADDRESS_PATH.is_file():
        return ()
    frame = pd.read_csv(ADDRESS_PATH, encoding="utf-8-sig")
    frame.columns = frame.columns.astype(str).str.strip()
    if "영업담당자" not in frame.columns:
        return ()
    pairs: list[tuple[str, str]] = []
    for _, row in frame.iterrows():
        staff = str(row.get("영업담당자") or "").strip()
        if staff in {"", "담당자없음", "nan", "None", "미지정"}:
            continue
        for column in ("사업체명", "상호"):
            if column not in frame.columns:
                continue
            name = _fold_name(row.get(column))
            if len(name) >= 2:
                pairs.append((name, staff))
    pairs.sort(key=lambda item: len(item[0]), reverse=True)
    return tuple(pairs)


def staff_of(client: object) -> str:
    folded = _fold_name(client)
    if len(folded) < 2:
        return ""
    exact = ""
    loose = ""
    for name, staff in _staff_pairs():
        if name == folded:
            exact = staff
            break
        if not loose and (folded.startswith(name) or name.startswith(folded)):
            loose = staff
    return exact or loose


def staff_key_in_label(label: str) -> str:
    folded = _fold_name(label)
    keys = sorted({staff for _name, staff in _staff_pairs()}, key=len, reverse=True)
    for key in keys:
        if _fold_name(key) and _fold_name(key) in folded:
            return key
    return ""


@lru_cache(maxsize=8)
def _raw_won_index(year: int, month: int) -> dict[str, float]:
    frame = _through(load_year(year), month)
    totals: dict[str, float] = {}
    if frame.empty:
        return totals
    for client, money in zip(frame["거래처"], frame["금액"], strict=False):
        key = _fold_name(client)
        if not key:
            continue
        totals[key] = totals.get(key, 0.0) + float(money)
    return totals


def _million_of(year: int, month: int, client: str) -> int:
    needle = _fold_name(client)
    if not needle:
        return 0
    totals = _raw_won_index(int(year), int(month))
    if needle in totals:
        return int(round(totals[needle] / 1_000_000))
    picked = [won for name, won in totals.items() if name.startswith(needle) or needle.startswith(name)]
    if not picked:
        return 0
    return int(round(sum(picked) / 1_000_000))


def volume_change_text(plan_year: int, month: int, client: str) -> str:
    """물량증감. 전년 대비 분석년의 백만원 증감률이다."""
    current = _million_of(int(plan_year) - 1, month, client)
    previous = _million_of(int(plan_year) - 2, month, client)
    if previous <= 0 and current <= 0:
        return ""
    if previous <= 0:
        return "신규"
    pct = int(round((current - previous) / previous * 100))
    if pct > 0:
        return f"▲{pct}% 증대"
    if pct < 0:
        return f"▼{abs(pct)}%"
    return "0%"


def staff_top_clients(plan_year: int, month: int, staff: str) -> list[dict]:
    """담당자 실적 1~10위. 금액은 부가세 포함 백만원이다."""
    key = str(staff or "").strip()
    if not key:
        return []
    frame = _through(load_year(int(plan_year) - 1), month)
    raw: dict[str, float] = {}
    label: dict[str, str] = {}
    if not frame.empty:
        for client, money in zip(frame["거래처"], frame["금액"], strict=False):
            if staff_of(client) != key:
                continue
            folded = _fold_name(client)
            if not folded:
                continue
            raw[folded] = raw.get(folded, 0.0) + float(money)
            current = label.get(folded, "")
            name = str(client).strip()
            if len(name) > len(current):
                label[folded] = name
    ranked = sorted(raw.items(), key=lambda item: item[1], reverse=True)[:10]
    rows: list[dict] = []
    for index, (folded, _won) in enumerate(ranked, start=1):
        client = label.get(folded) or folded
        rows.append(
            {
                "rank": index,
                "client": client,
                "prev": _million_of(int(plan_year) - 2, month, client),
                "curr": _million_of(int(plan_year) - 1, month, client),
                "change": volume_change_text(plan_year, month, client),
            }
        )
    return rows


_GAS_LABEL = {
    "B:O2": "LO2(BULK)",
    "B:N2": "LN2(BULK)",
    "B:AR": "LAr(BULK)",
    "B:CO2": "LCO2(BULK)",
    "B:LPG": "LPG(BULK)",
    "C:O2L": "LO2(LGC)",
    "C:N2L": "LN2(LGC)",
    "C:ARL": "LAr(LGC)",
    "C:O2": "O2(Cyl)",
    "C:N2": "N2(Cyl)",
    "C:Ar": "Ar(Cyl)",
    "C:CO2": "CO2(Cyl)",
    "C:Mix": "Mix Gas(Cyl)",
    "C:He": "He(Cyl)",
    "C:UHP": "UHP(Cyl)",
    "C:기타": "기타",
}


@lru_cache(maxsize=8)
def _client_item_won(year: int, month: int) -> dict[str, tuple[tuple[str, float], ...]]:
    frame = _through(load_year(year), month)
    buckets: dict[str, dict[str, float]] = {}
    if frame.empty:
        return {}
    for client, item, money in zip(frame["거래처"], frame["품목명"], frame["금액"], strict=False):
        key = _fold_name(client)
        if not key:
            continue
        kind = _item_key(str(item))
        bucket = buckets.setdefault(key, {})
        bucket[kind] = bucket.get(kind, 0.0) + float(money)
    return {key: tuple(sorted(kinds.items(), key=lambda item: item[1], reverse=True)) for key, kinds in buckets.items()}


def client_gas_text(plan_year: int, month: int, client: str) -> str:
    """분석년 매출에서 가스 종류를 적는다. 두 종류면 나란히, 더 많으면 대표 하나에 외를 붙인다."""
    needle = _fold_name(client)
    if not needle:
        return ""
    kinds = _client_item_won(int(plan_year) - 1, month)
    picked = kinds.get(needle)
    if picked is None:
        merged: dict[str, float] = {}
        for name, rows in kinds.items():
            if name.startswith(needle) or needle.startswith(name):
                for key, won in rows:
                    merged[key] = merged.get(key, 0.0) + won
        picked = tuple(sorted(merged.items(), key=lambda item: item[1], reverse=True))
    ranked = [(key, won) for key, won in picked if won > 0 and _GAS_LABEL.get(key, "") not in {"", "기타"}]
    if not ranked:
        ranked = [(key, won) for key, won in picked if won > 0 and _GAS_LABEL.get(key)]
    labels: list[str] = []
    for key, _won in ranked:
        label = _GAS_LABEL[key]
        if label not in labels:
            labels.append(label)
    if not labels:
        return ""
    if len(labels) == 1:
        return labels[0]
    if len(labels) == 2:
        return f"{labels[0]}, {labels[1]}"
    return f"{labels[0]} 외"


def staff_new_clients(plan_year: int, month: int, staff: str) -> list[dict]:
    """전년도에 없고 분석년에 있는 담당 거래처 전부."""
    key = str(staff or "").strip()
    if not key:
        return []
    previous = set(_raw_won_index(int(plan_year) - 2, month))
    frame = _through(load_year(int(plan_year) - 1), month)
    raw: dict[str, float] = {}
    label: dict[str, str] = {}
    if not frame.empty:
        for client, money in zip(frame["거래처"], frame["금액"], strict=False):
            if staff_of(client) != key:
                continue
            folded = _fold_name(client)
            if not folded or folded in previous:
                continue
            raw[folded] = raw.get(folded, 0.0) + float(money)
            name = str(client).strip()
            if len(name) > len(label.get(folded, "")):
                label[folded] = name
    rows: list[dict] = []
    for folded, won in sorted(raw.items(), key=lambda item: item[1], reverse=True):
        if won <= 0:
            continue
        client = label.get(folded) or folded
        rows.append(
            {
                "client": client,
                "gas": client_gas_text(plan_year, month, client),
                "won": int(round(won / 1_000_000)),
            }
        )
    return rows


def _shift_year_text(text: str, year: int) -> str:
    if int(year) == 2026 and "2023" not in text and "23년" not in text:
        return text
    shifted = text
    for src, token in _YEAR_TOKEN.items():
        shifted = shifted.replace(src, token)
    values = {
        "\ue000": str(year - 3),
        "\ue001": str(year - 2),
        "\ue002": str(year - 1),
        "\ue003": str(year),
        "\ue004": f"{(year - 3) % 100:02d}년",
        "\ue005": f"{(year - 2) % 100:02d}년",
        "\ue006": f"{(year - 1) % 100:02d}년",
        "\ue007": f"{year % 100:02d}년",
    }
    for token, value in values.items():
        shifted = shifted.replace(token, value)
    return shifted


def _map_text(xml: str, fn) -> str:
    return _AT.sub(lambda match: match.group(1) + fn(match.group(2)) + match.group(3), xml)


def _para_texts(para: str) -> list[str]:
    return [unescape(match.group(2)) for match in _AT.finditer(para)]


def _set_nth_text(block: str, index: int, text: str) -> str:
    seen = {"n": 0}
    safe = escape(text)

    def repl(match: re.Match[str]) -> str:
        seen["n"] += 1
        if seen["n"] != index:
            return match.group(0)
        return match.group(1) + safe + match.group(3)

    return _AT.sub(repl, block)


def _bullet_line(texts: list[str]) -> str | None:
    if not texts or "-" not in texts[0]:
        return None
    head = texts[0].strip().lstrip("-").strip()
    rest = "".join(texts[1:])
    line = f"{head}{rest}".strip() if head else rest.strip()
    return line or None


def _shape_lines(xml: str, kind: str) -> list[str] | None:
    for shape in re.finditer(r"<p:sp\b.*?</p:sp>", xml, flags=re.S):
        block = shape.group(0)
        if kind == "strategy" and STRATEGY_TITLE not in block:
            continue
        if kind == "review" and REVIEW_MARK not in block:
            continue
        if kind == "policy" and (POLICY_TITLE not in block or REVIEW_MARK in block or STRATEGY_TITLE in block):
            continue
        paras = re.findall(r"<a:p\b.*?</a:p>", block, flags=re.S)
        lines = []
        for para in paras[1:]:
            line = _bullet_line(_para_texts(para))
            if line:
                lines.append(line)
        return lines
    return None


def strategy_lines_in(xml: str) -> list[str] | None:
    """2페이지 3번 세부전략. 그 칸이 없으면 None, 항목이 없으면 빈 목록."""
    return _shape_lines(xml, "strategy")


def _read_slide2(year: int) -> str | None:
    path = proposal_path(year)
    if path is None or not path.is_file():
        return None
    try:
        return zipfile.ZipFile(path).read("ppt/slides/slide2.xml").decode("utf-8")
    except (OSError, KeyError, zipfile.BadZipFile):
        return None


def read_toc_lines(year: int) -> tuple[list[str], list[str], list[str]]:
    xml = _read_slide2(year)
    defaults = (list(POLICY_DEFAULTS), list(REVIEW_DEFAULTS), list(STRATEGY_DEFAULTS))
    if xml is None:
        return defaults
    found = [_shape_lines(xml, kind) for kind in ("policy", "review", "strategy")]
    return tuple(default if lines is None else lines for default, lines in zip(defaults, found, strict=False))


def read_strategy_lines(year: int) -> list[str]:
    return read_toc_lines(year)[2]


def _fill_bullets(block: str, lines: list[str], year: int, kind: str) -> str:
    paras = list(re.finditer(r"<a:p\b.*?</a:p>", block, flags=re.S))
    if not paras:
        return block
    title = paras[0].group(0)
    texts = _para_texts(title)
    if kind == "strategy" and texts and texts[0].isdigit():
        title = _set_nth_text(title, 1, f"{int(year) % 100:02d}")
    if kind == "review" and len(texts) >= 4 and texts[0].isdigit() and texts[2].isdigit():
        title = _set_nth_text(title, 1, f"{int(year) % 100:02d}")
        title = _set_nth_text(title, 3, f"{(int(year) - 1) % 100:02d}")
    sample = next((para.group(0) for para in paras[1:] if _bullet_line(_para_texts(para.group(0)))), _STRATEGY_BULLET)
    tail = ""
    if len(paras) > 1 and _bullet_line(_para_texts(paras[-1].group(0))) is None:
        tail = paras[-1].group(0)
    built = [title, *[_set_nth_text(sample, 2, line) for line in lines]]
    if tail:
        built.append(tail)
    return block[: paras[0].start()] + "".join(built) + block[paras[-1].end() :]


def _apply_toc_lines(
    xml: str,
    year: int,
    policy_lines: list[str] | None,
    review_lines: list[str] | None,
    strategy_lines: list[str] | None,
) -> str:
    """2페이지 1·2·3번 항목만 바꾼다. 넘기지 않은 칸은 그대로 둔다."""

    def repl(match: re.Match[str]) -> str:
        block = match.group(0)
        if STRATEGY_TITLE in block and strategy_lines is not None:
            return _fill_bullets(block, strategy_lines, year, "strategy")
        if REVIEW_MARK in block and review_lines is not None:
            return _fill_bullets(block, review_lines, year, "review")
        if POLICY_TITLE in block and policy_lines is not None:
            return _fill_bullets(block, policy_lines, year, "policy")
        return block

    return re.sub(r"<p:sp\b.*?</p:sp>", repl, xml, flags=re.S)


_NUMBERED_ITEM = (
    '<a:p><a:pPr marL="514350" indent="-514350"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc>'
    '<a:spcBef><a:spcPts val="864"/></a:spcBef><a:buAutoNum type="arabicPeriod"/></a:pPr>'
    '<a:r><a:rPr lang="ko-KR" altLang="en-US" dirty="0"><a:solidFill><a:srgbClr val="242529"/></a:solidFill>'
    '<a:latin typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/>'
    '<a:ea typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/></a:rPr>'
    "<a:t>거래처 관리</a:t></a:r></a:p>"
)


def _compact_title(text: str) -> str:
    return re.sub(r"\s+", "", text)


def _goal_lines(xml: str, kind: str) -> list[str] | None:
    mark = _GOAL_MARK[kind]
    for shape in re.finditer(r"<p:sp\b.*?</p:sp>", xml, flags=re.S):
        paras = re.findall(r"<a:p\b.*?</a:p>", shape.group(0), flags=re.S)
        if not paras:
            continue
        if _compact_title("".join(_para_texts(paras[0]))) != mark:
            continue
        lines = []
        for para in paras[1:]:
            line = "".join(_para_texts(para)).strip()
            if line:
                lines.append(line)
        return lines
    return None


def _read_slide3(year: int) -> str | None:
    path = proposal_path(year)
    if path is None or not path.is_file():
        return None
    try:
        return zipfile.ZipFile(path).read("ppt/slides/slide3.xml").decode("utf-8")
    except (OSError, KeyError, zipfile.BadZipFile):
        return None


def read_goal_lines(year: int) -> tuple[list[str], list[str], list[str]]:
    xml = _read_slide3(year)
    defaults = (list(DIRECTIVE_DEFAULTS), list(FOCUS_DEFAULTS), list(TASK_DEFAULTS))
    if xml is None:
        return defaults
    found = [_goal_lines(xml, kind) for kind in ("directive", "focus", "task")]
    return tuple(default if lines is None else lines for default, lines in zip(defaults, found, strict=False))


def _fill_numbered(block: str, lines: list[str]) -> str:
    paras = list(re.finditer(r"<a:p\b.*?</a:p>", block, flags=re.S))
    if not paras:
        return block
    title = paras[0].group(0)
    sample = next((para.group(0) for para in paras[1:] if "".join(_para_texts(para.group(0))).strip()), _NUMBERED_ITEM)
    built = [title, *[_set_cell_text(sample, line) for line in lines]]
    return block[: paras[0].start()] + "".join(built) + block[paras[-1].end() :]


def _apply_goal_lines(
    xml: str,
    directive_lines: list[str] | None,
    focus_lines: list[str] | None,
    task_lines: list[str] | None,
) -> str:
    """3페이지 영업 방침, 중점운영 목표, 추진 과제만 바꾼다."""
    chosen = {
        "영업방침": directive_lines,
        "중점운영목표": focus_lines,
        "추진과제": task_lines,
    }

    def repl(match: re.Match[str]) -> str:
        block = match.group(0)
        paras = re.findall(r"<a:p\b.*?</a:p>", block, flags=re.S)
        if not paras:
            return block
        mark = _compact_title("".join(_para_texts(paras[0])))
        lines = chosen.get(mark)
        if lines is None:
            return block
        return _fill_numbered(block, lines)

    return re.sub(r"<p:sp\b.*?</p:sp>", repl, xml, flags=re.S)


def _set_cell_text(cell_xml: str, text: str) -> str:
    seen = {"n": 0}
    safe = escape(text)

    def repl(match: re.Match[str]) -> str:
        seen["n"] += 1
        return match.group(1) + (safe if seen["n"] == 1 else "") + match.group(3)

    return _AT.sub(repl, cell_xml)


def _update_table(table_xml: str, cells: dict[tuple[int, int], str]) -> str:
    rows = list(re.finditer(r"<a:tr\b.*?</a:tr>", table_xml, flags=re.S))
    if not rows:
        return table_xml
    pieces = [table_xml[: rows[0].start()]]
    for row_index, row in enumerate(rows):
        row_xml = row.group(0)
        found = list(re.finditer(r"<a:tc\b.*?</a:tc>", row_xml, flags=re.S))
        if not found:
            pieces.append(row_xml)
        else:
            built = [row_xml[: found[0].start()]]
            for col_index, cell in enumerate(found):
                block = cell.group(0)
                if (row_index, col_index) in cells:
                    block = _set_cell_text(block, cells[(row_index, col_index)])
                built.append(block)
                nxt = found[col_index + 1].start() if col_index + 1 < len(found) else len(row_xml)
                built.append(row_xml[cell.end() : nxt])
            pieces.append("".join(built))
        nxt_row = rows[row_index + 1].start() if row_index + 1 < len(rows) else len(table_xml)
        pieces.append(table_xml[row.end() : nxt_row])
    return "".join(pieces)


def _update_tables(xml: str, tables: list[dict[tuple[int, int], str]]) -> str:
    found = list(re.finditer(r"<a:tbl\b.*?</a:tbl>", xml, flags=re.S))
    if not found:
        return xml
    pieces = [xml[: found[0].start()]]
    for index, table in enumerate(found):
        block = table.group(0)
        if index < len(tables) and tables[index]:
            block = _update_table(block, tables[index])
        pieces.append(block)
        nxt = found[index + 1].start() if index + 1 < len(found) else len(xml)
        pieces.append(xml[table.end() : nxt])
    return "".join(pieces)


def _fill_cache(block: str, values: list[object]) -> str:
    cursor = {"i": 0}

    def repl(match: re.Match[str]) -> str:
        if cursor["i"] >= len(values) or values[cursor["i"]] is None:
            cursor["i"] += 1
            return match.group(0)
        value = values[cursor["i"]]
        cursor["i"] += 1
        return f"<c:v>{value}</c:v>"

    return re.sub(r"<c:v>[^<]*</c:v>", repl, block)


def _patch_chart_xml(xml: str, categories: list[str], series: dict[str, list[object]]) -> str:
    def repl_ser(match: re.Match[str]) -> str:
        block = match.group(0)
        name_match = re.search(r"<c:tx>.*?<c:v>([^<]*)</c:v>", block, flags=re.S)
        name = name_match.group(1) if name_match else ""
        cat = re.search(r"<c:cat>.*?</c:cat>", block, flags=re.S)
        if cat:
            filled = _fill_cache(cat.group(0), categories)
            block = block[: cat.start()] + filled + block[cat.end() :]
        if name in series:
            num = re.search(r"<c:numCache>.*?</c:numCache>", block, flags=re.S)
            if num:
                filled = _fill_cache(num.group(0), series[name])
                block = block[: num.start()] + filled + block[num.end() :]
        return block

    return re.sub(r"<c:ser>.*?</c:ser>", repl_ser, xml, flags=re.S)


def _patch_workbook(blob: bytes, cells: dict[str, object]) -> bytes:
    book = load_workbook(io.BytesIO(blob))
    sheet = book.active
    for coord, value in cells.items():
        sheet[coord].value = value
    out = io.BytesIO()
    book.save(out)
    return out.getvalue()


def _embedding(parts: dict[str, bytes], chart_file: str) -> str | None:
    rels = parts.get(f"ppt/charts/_rels/{Path(chart_file).name}.rels", b"").decode("utf-8")
    matched = re.search(r'Target="(?:\.\./)?embeddings/([^"]+)"', rels)
    if not matched:
        return None
    return "ppt/embeddings/" + matched.group(1)


def _prior_total(year: int) -> int:
    return int(PRIOR_PLAN_TOTAL.get(int(year), 0))


def _slide4_numbers(year: int, month: int, plan: dict[str, int]) -> dict[str, str]:
    older = plant_won(year - 2, month)
    mid = plant_won(year - 1, month)
    return {
        "5,695": _fmt_int(plan["아산"]),
        "15,764": _fmt_int(plan["화성"]),
        "21,461": _fmt_int(plan["전체"]),
        "4,417": _fmt_int(annual_million(mid["아산"], month)),
        "18,121": _fmt_int(annual_million(mid["전체"], month)),
        "4,213": _fmt_int(annual_million(older["아산"], month)),
        "13,293": _fmt_int(annual_million(older["화성"], month)),
        "17,506": _fmt_int(annual_million(older["전체"], month)),
        "21,567": _fmt_int(_prior_total(year - 1) or annual_million(mid["전체"], month)),
    }


def _chart1(year: int, month: int, plan: dict[str, int]) -> tuple[list[str], dict[str, list[int]]]:
    cats = [f"{year}(P)", str(year - 1), str(year - 2)]
    y1 = plant_won(year - 1, month)
    y2 = plant_won(year - 2, month)
    asan = [plan["아산"], annual_million(y1["아산"], month), annual_million(y2["아산"], month)]
    hwa = [plan["화성"], annual_million(y1["화성"], month), annual_million(y2["화성"], month)]
    total = [plan["전체"], annual_million(y1["전체"], month), annual_million(y2["전체"], month)]
    series = {"아산": asan, "화성": hwa, "전체": total, "P": [_prior_total(year - 1), _prior_total(year - 2)]}
    return cats, series


def _chart3(year: int, month: int, tons_plan: dict[str, int]) -> tuple[list[str], dict[str, list[int]]]:
    cats = [str(year - 2), str(year - 1), f"{year}년(P)"]
    older = tons_of(year - 2, month)
    mid = tons_of(year - 1, month)
    series = {}
    for label, key in (("탄산", "탄산"), ("알곤", "알곤"), ("질소", "질소"), ("산소", "산소"), ("LPG(T/K)", "LPG")):
        series[label if label != "LPG(T/K)" else "LPG(T/K)"] = [older[key], mid[key], int(tons_plan[key])]
    # 차트 계열 이름은 LPG(T/K)
    series["LPG(T/K)"] = [older["LPG"], mid["LPG"], int(tons_plan["LPG"])]
    return cats, series


def _month_cells(values: list[int], total: int | None = None) -> list[str]:
    """PPT 열 순서: 합계, 12월 … 1월."""
    head = sum(values) if total is None else int(total)
    return [_fmt_int(head), *[_fmt_int(values[m]) for m in range(11, -1, -1)]]


def _slide5_cells(year: int, month: int, asan_months: list[int], hwa_months: list[int]) -> dict[tuple[int, int], str]:
    years = (year - 2, year - 1, year)
    cells: dict[tuple[int, int], str] = {}
    for block, y in enumerate(years):
        if y == year:
            rows = [
                (hwa_months, sum(hwa_months)),
                (asan_months, sum(asan_months)),
                (
                    [a + b for a, b in zip(hwa_months, asan_months, strict=False)],
                    sum(hwa_months) + sum(asan_months),
                ),
            ]
        else:
            won = plant_won(y, month)
            rows = [
                (millions(won[label], month), annual_million(won[label], month))
                for label in ("화성", "아산", "전체")
            ]
        for offset, (row_values, total) in enumerate(rows):
            row_index = 1 + block * 3 + offset
            for col, text in enumerate(_month_cells(row_values, total)):
                cells[(row_index, col)] = text
            cells[(row_index, 13)] = ("화 성", "아 산", "전 체")[offset]
    return cells


def _put_item_row(cells: dict[tuple[int, int], str], excel_row: int, row: dict[str, str]) -> None:
    cells[(excel_row, 2)] = row["수량"]
    cells[(excel_row, 3)] = row["월평균"]
    cells[(excel_row, 4)] = row["월평균 증감"]
    cells[(excel_row, 5)] = row["금액"]
    cells[(excel_row, 6)] = row["전년대비"]
    cells[(excel_row, 7)] = row["비중"]


def _slide6_cells(year: int, month: int) -> list[dict[tuple[int, int], str]]:
    """표 행: 2–6 벌크, 7 소계, 8–18 실린더, 19 소계, 20 총합계."""
    tables = []
    for y in (year - 2, year - 1):
        sheet = item_sheet(y, month)
        body = sheet[:-1]
        cells: dict[tuple[int, int], str] = {(0, 2): f"{y}년"}
        for offset, row in enumerate(body[:5]):
            _put_item_row(cells, 2 + offset, row)
        for offset, row in enumerate(body[5:]):
            _put_item_row(cells, 8 + offset, row)
        cells[(20, 2)] = sheet[-1]["수량"]
        tables.append(cells)
    return tables


def _slide7_cells(year: int, month: int, tons_plan: dict[str, int]) -> dict[tuple[int, int], str]:
    older = tons_of(year - 2, month)
    mid = tons_of(year - 1, month)
    labels = ("산소", "질소", "알곤", "탄산", "LPG")
    cells: dict[tuple[int, int], str] = {
        (0, 1): f"{year - 2}년",
        (0, 2): f"{year - 1}년",
        (0, 3): f"{year}년(P)",
    }
    for index, label in enumerate(labels, start=1):
        cells[(index, 1)] = _fmt_int(older[label])
        cells[(index, 2)] = _fmt_int(mid[label])
        cells[(index, 3)] = _fmt_int(tons_plan[label])
    cells[(6, 1)] = _fmt_int(older["합계"])
    cells[(6, 2)] = _fmt_int(mid["합계"])
    cells[(6, 3)] = _fmt_int(sum(int(tons_plan[label]) for label in labels))
    return cells


BULK_DEFAULTS = (
    "- 26년 벌크 전년대비 목표 17% 증대",
    "- 26년 LPG(T/K) 전년대비 62% 증대 (라쿨 설비 증설)",
    "- 26년 탄산 전년대비 25% 증대",
    "(1~3월, 10~12월 : 월 4,400톤, 4~6월 :5,900톤, 7~9월 : 7,400톤)",
)


def bulk_growth_percent(year: int, month: int, tons_plan: dict[str, int] | None = None) -> int:
    """전년 합계 대비 올해 계획 합계의 증가율. 슬라이드 삼각형 숫자다."""
    labels = ("산소", "질소", "알곤", "탄산", "LPG")
    chosen = tons_plan or PLAN_TONS
    prev = int(tons_of(int(year) - 1, month)["합계"])
    plan = sum(int(chosen[label]) for label in labels)
    if prev <= 0:
        return 0
    return int(round((plan - prev) / prev * 100))


def _tune_bulk_line(line: str, year: int, pct: int) -> str:
    text = re.sub(r"(?<!\d)\d{2}(?=년)", f"{int(year) % 100:02d}", str(line).strip())
    if "벌크" in text and "목표" in text:
        text = re.sub(r"\d+%", f"{int(pct)}%", text, count=1)
    return text


def _bulk_lines_in(xml: str) -> list[str] | None:
    for shape in re.finditer(r"<p:sp\b.*?</p:sp>", xml, flags=re.S):
        block = shape.group(0)
        if "outerShdw" not in block or 'sz="2400"' not in block or "벌크운영 계획" in block:
            continue
        lines = []
        for para in re.findall(r"<a:p\b.*?</a:p>", block, flags=re.S):
            line = "".join(_para_texts(para)).strip()
            if line:
                lines.append(line)
        return lines
    return None


def read_bulk_lines(year: int) -> list[str]:
    path = proposal_path(year)
    if path is None or not path.is_file():
        return list(BULK_DEFAULTS)
    try:
        xml = zipfile.ZipFile(path).read("ppt/slides/slide7.xml").decode("utf-8")
    except (OSError, KeyError, zipfile.BadZipFile):
        return list(BULK_DEFAULTS)
    found = _bulk_lines_in(xml)
    return list(BULK_DEFAULTS) if found is None else found


def _fill_plain_lines(block: str, lines: list[str]) -> str:
    paras = list(re.finditer(r"<a:p\b.*?</a:p>", block, flags=re.S))
    if not paras:
        return block
    sample = paras[0].group(0)
    built = [_set_cell_text(sample, line) for line in lines] or [_set_cell_text(sample, "")]
    return block[: paras[0].start()] + "".join(built) + block[paras[-1].end() :]


def _patch_goal_percent(block: str, pct: int) -> str:
    def para_repl(match: re.Match[str]) -> str:
        para = match.group(0)
        joined = "".join(_para_texts(para))
        if "벌크" not in joined or "목표" not in joined:
            return para

        def text_repl(text_match: re.Match[str]) -> str:
            if not re.fullmatch(r"\d+%\s*", text_match.group(2)):
                return text_match.group(0)
            space = " " if text_match.group(2).endswith(" ") else ""
            return text_match.group(1) + f"{int(pct)}%{space}" + text_match.group(3)

        return _AT.sub(text_repl, para)

    return re.sub(r"<a:p\b.*?</a:p>", para_repl, block, flags=re.S)


def _apply_bulk_plan(
    xml: str,
    year: int,
    month: int,
    tons_plan: dict[str, int],
    lines: list[str] | None,
) -> str:
    pct = bulk_growth_percent(year, month, tons_plan)
    tuned = None if lines is None else [_tune_bulk_line(line, year, pct) for line in lines if str(line).strip()]

    def repl(match: re.Match[str]) -> str:
        block = match.group(0)
        text = "".join(piece for para in re.findall(r"<a:p\b.*?</a:p>", block, flags=re.S) for piece in _para_texts(para)).strip()
        if re.fullmatch(r"\d+%", text):
            return _set_cell_text(block, f"{pct}%")
        if "outerShdw" in block and 'sz="2400"' in block and "벌크운영 계획" not in block:
            if tuned is None:
                return _patch_goal_percent(block, pct)
            return _fill_plain_lines(block, tuned)
        return block

    return re.sub(r"<p:sp\b.*?</p:sp>", repl, xml, flags=re.S)


@lru_cache(maxsize=1)
def _page8_template() -> tuple[tuple[str, str, str, str], ...]:
    template = find_template()
    if template is None:
        return ()
    table = next(shape.table for shape in Presentation(str(template)).slides[7].shapes if shape.has_table)
    product = ""
    rows = []
    for row_index in range(2, len(table.rows)):
        raw = table.cell(row_index, 0).text.strip()
        if raw:
            product = raw
        rows.append(
            (
                product,
                table.cell(row_index, 1).text.strip(),
                table.cell(row_index, 3).text.strip(),
                table.cell(row_index, 5).text.strip(),
            )
        )
    return tuple(rows)


def _parse_int(text: str) -> int | None:
    digits = re.sub(r"[^\d-]", "", str(text or ""))
    if digits in {"", "-"}:
        return None
    try:
        return int(digits)
    except ValueError:
        return None


def page8_figures(year: int, month: int, product: str, client: str, budget_text: str) -> tuple[str, str, str]:
    """실적 집계, 예상 표시, 증감. 예상이 비어 있고 실적이 있으면 증감은 실적의 마이너스다."""
    name = str(client or "").strip()
    raw_budget = str(budget_text or "").strip()
    budget = None if raw_budget in {"", "-"} else _parse_int(raw_budget)
    actual = client_kg(year - 1, month, name, product) if name else None
    actual_txt = "-" if actual is None else _fmt_int(actual)
    actual_n = 0 if actual is None else int(round(actual))
    if budget is None:
        budget_txt = "-"
        delta_n = None if actual is None else -actual_n
    else:
        budget_txt = _fmt_int(budget)
        delta_n = int(budget) - actual_n
    delta_txt = "-" if delta_n is None else _fmt_int(delta_n)
    return actual_txt, budget_txt, delta_txt


def _row_cell_texts(row_xml: str) -> list[str]:
    cells = re.findall(r"<a:tc\b.*?</a:tc>", row_xml, flags=re.S)
    return ["".join(_para_texts(cell)).strip() for cell in cells]


def _page8_rows_in(xml: str) -> list[dict[str, str]]:
    table = re.search(r"<a:tbl\b.*?</a:tbl>", xml, flags=re.S)
    if table is None:
        return []
    rows = re.findall(r"<a:tr\b.*?</a:tr>", table.group(0), flags=re.S)
    current = ""
    found: list[dict[str, str]] = []
    for row in rows[2:]:
        cells = _row_cell_texts(row)
        if len(cells) < 6:
            continue
        product = cells[0].replace("\u3000", "").strip()
        if product:
            current = product
        if current not in PAGE8_PRODUCTS:
            continue
        budget = cells[3].strip()
        note = cells[5].replace("\u3000", "").strip()
        found.append(
            {
                "product": current,
                "client": cells[1].strip(),
                "budget": "" if budget in {"", "-"} else budget,
                "note": "" if note in {"", "-"} else note,
            }
        )
    return found


def read_page8_rows(year: int) -> list[dict[str, str]]:
    path = proposal_path(year)
    if path is None or not path.is_file():
        return []
    try:
        xml = zipfile.ZipFile(path).read("ppt/slides/slide8.xml").decode("utf-8")
    except (OSError, KeyError, zipfile.BadZipFile):
        return []
    return _page8_rows_in(xml)


def _set_row_cells(row_xml: str, values: list[str]) -> str:
    found = list(re.finditer(r"<a:tc\b.*?</a:tc>", row_xml, flags=re.S))
    if not found:
        return row_xml
    pieces = [row_xml[: found[0].start()]]
    for index, cell in enumerate(found):
        block = _set_cell_text(cell.group(0), values[index]) if index < len(values) else cell.group(0)
        pieces.append(block)
        nxt = found[index + 1].start() if index + 1 < len(found) else len(row_xml)
        pieces.append(row_xml[cell.end() : nxt])
    return "".join(pieces)


def _paint_delta(row_xml: str, negative: bool) -> str:
    found = list(re.finditer(r"<a:tc\b.*?</a:tc>", row_xml, flags=re.S))
    if len(found) < 5:
        return row_xml
    cell = found[4].group(0)
    fill = _DELTA_RED if negative else _DELTA_DARK

    def repl(match: re.Match[str]) -> str:
        body = match.group(0)
        if "FF0000" in body or 'val="bg2"' in body:
            return fill
        return body

    painted = re.sub(r"<a:solidFill>.*?</a:solidFill>", repl, cell, flags=re.S)
    return row_xml[: found[4].start()] + painted + row_xml[found[4].end() :]


def _span_row(row_xml: str, count: int) -> str:
    row_xml = re.sub(r"\srowSpan=\"\d+\"", "", row_xml, count=1)
    row_xml = re.sub(r"\svMerge=\"\d+\"", "", row_xml, count=1)
    if count > 1:
        row_xml = re.sub(r"<a:tc\b", f'<a:tc rowSpan="{count}"', row_xml, count=1)
    return row_xml


def _apply_page8(xml: str, year: int, month: int, rows: list[dict]) -> str:
    """8페이지 데이터 행만 다시 쓴다. 글꼴과 테두리는 원래 행을 복사한다."""
    table_match = re.search(r"<a:tbl\b.*?</a:tbl>", xml, flags=re.S)
    if table_match is None:
        return xml
    table = table_match.group(0)
    found = list(re.finditer(r"<a:tr\b.*?</a:tr>", table, flags=re.S))
    if len(found) < 4:
        return xml
    head = found[2].group(0)
    cont = found[3].group(0)
    grouped: dict[str, list[dict]] = {product: [] for product in PAGE8_PRODUCTS}
    for row in rows:
        product = str(row.get("product") or "").strip()
        client = str(row.get("client") or "").strip()
        if product not in grouped or not client:
            continue
        grouped[product].append(row)
    built = [found[0].group(0), found[1].group(0)]
    for product, items in grouped.items():
        if not items:
            continue
        for index, item in enumerate(items):
            sample = head if index == 0 else cont
            row_xml = _span_row(sample, len(items)) if index == 0 else sample
            actual_txt, budget_txt, delta_txt = page8_figures(
                year, month, product, str(item.get("client") or ""), str(item.get("budget") or "")
            )
            delta_n = _parse_int(delta_txt)
            row_xml = _set_row_cells(
                row_xml,
                [
                    product if index == 0 else "",
                    str(item.get("client") or "").strip(),
                    actual_txt,
                    budget_txt,
                    delta_txt,
                    str(item.get("note") or "").strip(),
                ],
            )
            built.append(_paint_delta(row_xml, delta_n is not None and delta_n < 0))
    new_table = table[: found[0].start()] + "".join(built) + table[found[-1].end() :]
    return xml[: table_match.start()] + new_table + xml[table_match.end() :]


def _blank_cell(text: str) -> str:
    cleaned = str(text or "").replace("\u3000", "").strip()
    return cleaned or "\u3000"


def _page10_groups_in(xml: str) -> list[dict]:
    table = re.search(r"<a:tbl\b.*?</a:tbl>", xml, flags=re.S)
    if table is None:
        return []
    rows = re.findall(r"<a:tr\b.*?</a:tr>", table.group(0), flags=re.S)
    groups: list[dict] = []
    current: dict | None = None
    for row in rows[2:]:
        cells = _row_cell_texts(row)
        if len(cells) < 6:
            continue
        staff = cells[0].replace("\u3000", "").strip()
        if staff:
            current = {"staff": staff, "news": [], "growth": []}
            groups.append(current)
        if current is None:
            continue
        client = cells[1].replace("\u3000", "").strip()
        gas = cells[2].replace("\u3000", "").strip()
        growth = cells[3].replace("\u3000", "").strip()
        if client or gas:
            current["news"].append({"client": client, "gas": gas})
        if growth:
            current["growth"].append({"client": growth})
    return groups


def read_page10_groups(year: int, page: int = 10) -> list[dict]:
    path = proposal_path(year)
    if path is None or not path.is_file():
        return []
    try:
        xml = zipfile.ZipFile(path).read(f"ppt/slides/slide{int(page)}.xml").decode("utf-8")
    except (OSError, KeyError, zipfile.BadZipFile):
        return []
    return _page10_groups_in(xml)


def _apply_page10_heading(xml: str, year: int) -> str:
    """제목의 실적 연도만 맞춘다. 입사년도는 표 문안 그대로 둔다."""
    yy = f"{(int(year) - 1) % 100:02d}"

    def shape(match: re.Match[str]) -> str:
        block = match.group(0)
        if "담당별" not in block:
            return block

        def run(text_match: re.Match[str]) -> str:
            body = text_match.group(2)
            found = re.fullmatch(r"(\s*)\d{2}", body)
            if found is None:
                return text_match.group(0)
            return text_match.group(1) + found.group(1) + yy + text_match.group(3)

        return _AT.sub(run, block)

    return re.sub(r"<p:sp\b.*?</p:sp>", shape, xml, flags=re.S)


def _apply_page10(xml: str, year: int, month: int, groups: list[dict]) -> str:
    """10페이지 데이터 행만 다시 쓴다. 글꼴과 테두리는 원래 행을 복사한다."""
    table_match = re.search(r"<a:tbl\b.*?</a:tbl>", xml, flags=re.S)
    if table_match is None:
        return xml
    table = table_match.group(0)
    found = list(re.finditer(r"<a:tr\b.*?</a:tr>", table, flags=re.S))
    if len(found) < 4:
        return xml
    head = found[2].group(0)
    cont = found[3].group(0)
    built = [found[0].group(0), found[1].group(0)]
    for group in groups:
        staff = str(group.get("staff") or "").strip()
        news = [
            item
            for item in group.get("news") or []
            if str(item.get("client") or "").strip() or str(item.get("gas") or "").strip()
        ]
        growth = [item for item in group.get("growth") or [] if str(item.get("client") or "").strip()]
        count = max(len(news), len(growth))
        if not staff or count == 0:
            continue
        for index in range(count):
            sample = _span_row(head, count) if index == 0 else cont
            new_item = news[index] if index < len(news) else {}
            grow_item = growth[index] if index < len(growth) else {}
            client = str(grow_item.get("client") or "").strip()
            sample = _set_row_cells(
                sample,
                [
                    staff if index == 0 else "\u3000",
                    _blank_cell(str(new_item.get("client") or "")),
                    _blank_cell(str(new_item.get("gas") or "")),
                    _blank_cell(client),
                    _blank_cell(volume_change_text(year, month, client) if client else ""),
                    "\u3000",
                ],
            )
            built.append(sample)
    new_table = table[: found[0].start()] + "".join(built) + table[found[-1].end() :]
    xml = xml[: table_match.start()] + new_table + xml[table_match.end() :]
    return _apply_page10_heading(xml, year)


def _slide8_cells(year: int, month: int) -> dict[tuple[int, int], str]:
    cells: dict[tuple[int, int], str] = {}
    for index, (product, client, budget_text, _note) in enumerate(_page8_template()):
        actual = client_kg(year - 1, month, client, product)
        if actual is None:
            continue
        row = 2 + index
        cells[(row, 2)] = _fmt_int(actual)
        budget = _parse_int(budget_text)
        if budget is not None:
            cells[(row, 4)] = _fmt_int(budget - actual)
    return cells


def compose_staff_map(legend: bytes, dots: bytes, size: tuple[int, int] = (1600, 1274)) -> bytes:
    """담당자 색 범례를 위에, 거래처 점 지도를 아래에 붙인다."""
    from PIL import Image

    width, height = size
    canvas = Image.new("RGB", size, "white")
    legend_im = Image.open(io.BytesIO(legend)).convert("RGBA")
    dots_im = Image.open(io.BytesIO(dots)).convert("RGBA")
    pad = 28
    legend_w = width - pad * 2
    legend_h = max(1, round(legend_im.height * legend_w / legend_im.width))
    legend_im = legend_im.resize((legend_w, legend_h), Image.Resampling.LANCZOS)
    canvas.paste(legend_im, (pad, pad), legend_im)
    top = pad + legend_h + 16
    room_w = width - pad * 2
    room_h = height - top - pad
    scale = min(room_w / dots_im.width, room_h / dots_im.height)
    dots_w = max(1, round(dots_im.width * scale))
    dots_h = max(1, round(dots_im.height * scale))
    dots_im = dots_im.resize((dots_w, dots_h), Image.Resampling.LANCZOS)
    left = pad + (room_w - dots_w) // 2
    canvas.paste(dots_im, (left, top), dots_im)
    out = io.BytesIO()
    canvas.save(out, format="PNG")
    return out.getvalue()


def replace_named_part(path: Path, name: str, data: bytes) -> None:
    """압축 안의 파일 하나만 바꾼다. 슬라이드 XML은 그대로 둔다."""
    with zipfile.ZipFile(path, "r") as zin:
        parts = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    if name not in parts:
        raise KeyError(name)
    parts[name] = data
    _write_package(path, path, parts)


def _cell_plain(cell_xml: str) -> str:
    lines: list[str] = []
    for para in re.findall(r"<a:p\b.*?</a:p>", cell_xml, flags=re.S):
        text = "".join(_para_texts(para)).replace("\u3000", "").strip()
        if text:
            lines.append(text)
    return "\n".join(lines)


_PAGE14_RUN = (
    '<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="2300" b="0" dirty="0">'
    '<a:latin typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/>'
    '<a:ea typeface="HY견고딕" panose="02030600000101010101" pitchFamily="18" charset="-127"/>'
    "</a:rPr><a:t>\u3000</a:t></a:r>"
)


def _set_cell_lines(cell_xml: str, text: str) -> str:
    lines = [line.strip() for line in str(text or "").splitlines() if line.strip()] or ["\u3000"]
    paras = list(re.finditer(r"<a:p\b.*?</a:p>", cell_xml, flags=re.S))
    if not paras:
        return _set_cell_text(cell_xml, lines[0] if lines != ["\u3000"] else "")
    template = paras[0].group(0)
    if "<a:t" not in template:
        template = template.replace("</a:p>", _PAGE14_RUN + "</a:p>", 1)
    built = [_set_cell_text(template, line) for line in lines]
    return cell_xml[: paras[0].start()] + "".join(built) + cell_xml[paras[-1].end() :]


def _set_row_lines(row_xml: str, values: list[str]) -> str:
    found = list(re.finditer(r"<a:tc\b.*?</a:tc>", row_xml, flags=re.S))
    if not found:
        return row_xml
    pieces = [row_xml[: found[0].start()]]
    for index, cell in enumerate(found):
        block = _set_cell_lines(cell.group(0), values[index]) if index < len(values) else cell.group(0)
        pieces.append(block)
        nxt = found[index + 1].start() if index + 1 < len(found) else len(row_xml)
        pieces.append(row_xml[cell.end() : nxt])
    return "".join(pieces)


def _clear_row_spans(row_xml: str) -> str:
    row_xml = re.sub(r'\srowSpan="\d+"', "", row_xml)
    return re.sub(r'\svMerge="\d+"', "", row_xml)


def _page14_groups_in(xml: str) -> list[dict]:
    table = re.search(r"<a:tbl\b.*?</a:tbl>", xml, flags=re.S)
    if table is None:
        return []
    rows = re.findall(r"<a:tr\b.*?</a:tr>", table.group(0), flags=re.S)
    groups: list[dict] = []
    current: dict | None = None
    for row in rows[1:]:
        cells = re.findall(r"<a:tc\b.*?</a:tc>", row, flags=re.S)
        if len(cells) < 4:
            continue
        staff = _cell_plain(cells[0])
        if staff:
            current = {"staff": staff, "rows": []}
            groups.append(current)
        if current is None:
            continue
        region = _cell_plain(cells[1])
        zone = _cell_plain(cells[2])
        note = _cell_plain(cells[3])
        if region or zone or note:
            current["rows"].append({"region": region, "zone": zone, "note": note})
    return groups


def read_page14_groups(year: int) -> list[dict]:
    path = proposal_path(year)
    if path is None or not path.is_file():
        return []
    try:
        xml = zipfile.ZipFile(path).read("ppt/slides/slide14.xml").decode("utf-8")
    except (OSError, KeyError, zipfile.BadZipFile):
        return []
    return _page14_groups_in(xml)


def _apply_page14(xml: str, groups: list[dict]) -> str:
    """14페이지 지역 표만 다시 쓴다. 글꼴과 테두리는 원래 행을 복사한다."""
    table_match = re.search(r"<a:tbl\b.*?</a:tbl>", xml, flags=re.S)
    if table_match is None:
        return xml
    table = table_match.group(0)
    found = list(re.finditer(r"<a:tr\b.*?</a:tr>", table, flags=re.S))
    if len(found) < 4:
        return xml
    head = _clear_row_spans(found[2].group(0))
    cont = _clear_row_spans(found[3].group(0))
    note_cell = re.findall(r"<a:tc\b.*?</a:tc>", head, flags=re.S)[3]
    built = [found[0].group(0)]
    for group in groups:
        staff = str(group.get("staff") or "").strip()
        rows = [
            item
            for item in group.get("rows") or []
            if str(item.get("region") or "").strip() or str(item.get("zone") or "").strip() or str(item.get("note") or "").strip()
        ]
        if not staff or not rows:
            continue
        for index, item in enumerate(rows):
            sample = head if index == 0 else cont
            if index == 0 and len(rows) > 1:
                sample = re.sub(r"<a:tc\b", f'<a:tc rowSpan="{len(rows)}"', sample, count=1)
            elif index > 0:
                sample = re.sub(r"<a:tc\b", '<a:tc vMerge="1"', sample, count=1)
            note_at = list(re.finditer(r"<a:tc\b.*?</a:tc>", sample, flags=re.S))
            if len(note_at) >= 4:
                sample = sample[: note_at[3].start()] + note_cell + sample[note_at[3].end() :]
            built.append(
                _set_row_lines(
                    sample,
                    [
                        staff if index == 0 else "\u3000",
                        str(item.get("region") or ""),
                        str(item.get("zone") or ""),
                        str(item.get("note") or ""),
                    ],
                )
            )
    new_table = table[: found[0].start()] + "".join(built) + table[found[-1].end() :]
    return xml[: table_match.start()] + new_table + xml[table_match.end() :]


def _edit_page14(year: int, groups: list[dict]) -> list[dict]:
    """14페이지 구분·담당지역·산업단지·비고를 넣고 뺀다."""
    store = f"bppt_p14_{year}"
    seq_key = f"bppt_p14_seq_{year}"
    if store not in st.session_state:
        seeded: list[dict] = []
        nxt = 1
        for group in groups:
            rows = []
            for row in group.get("rows") or []:
                rows.append(
                    {
                        "id": nxt,
                        "region": row.get("region") or "",
                        "zone": row.get("zone") or "",
                        "note": row.get("note") or "",
                    }
                )
                nxt += 1
            if not rows:
                rows.append({"id": nxt, "region": "", "zone": "", "note": ""})
                nxt += 1
            seeded.append({"id": nxt, "staff": group.get("staff") or "", "rows": rows})
            nxt += 1
        st.session_state[store] = seeded
        st.session_state[seq_key] = nxt
    people = st.session_state[store]
    head_staff, head_region, head_zone, head_note, _head_delete = st.columns([1.1, 1.6, 2.4, 1.2, 0.7])
    head_staff.markdown("**구분**")
    head_region.markdown("**담당지역**")
    head_zone.markdown("**주요 산업단지**")
    head_note.markdown("**비고**")
    saved: list[dict] = []
    for person in list(people):
        pid = person["id"]
        staff_key = f"bppt_p14s_{year}_{pid}"
        if staff_key not in st.session_state:
            st.session_state[staff_key] = person.get("staff") or ""
        kept_rows: list[dict] = []
        for index, row in enumerate(list(person["rows"])):
            rid = row["id"]
            region_key = f"bppt_p14r_{year}_{pid}_{rid}"
            zone_key = f"bppt_p14z_{year}_{pid}_{rid}"
            note_key = f"bppt_p14n_{year}_{pid}_{rid}"
            if region_key not in st.session_state:
                st.session_state[region_key] = row.get("region") or ""
            if zone_key not in st.session_state:
                st.session_state[zone_key] = row.get("zone") or ""
            if note_key not in st.session_state:
                st.session_state[note_key] = row.get("note") or ""
            staff_col, region_col, zone_col, note_col, delete_col = st.columns([1.1, 1.6, 2.4, 1.2, 0.7])
            if index == 0:
                staff_col.text_input("구분", key=staff_key, label_visibility="collapsed")
            region_col.text_area("담당지역", key=region_key, height=70, label_visibility="collapsed")
            zone_col.text_area("주요 산업단지", key=zone_key, height=70, label_visibility="collapsed")
            note_col.text_input("비고", key=note_key, label_visibility="collapsed")
            region = str(st.session_state.get(region_key) or "")
            zone = str(st.session_state.get(zone_key) or "")
            note = str(st.session_state.get(note_key) or "")
            row["region"], row["zone"], row["note"] = region, zone, note
            if delete_col.button("삭제", key=f"bppt_p14del_{year}_{pid}_{rid}"):
                if len(person["rows"]) == 1:
                    st.session_state[store] = [item for item in people if item["id"] != pid]
                else:
                    person["rows"] = [item for item in person["rows"] if item["id"] != rid]
                st.rerun(scope="fragment")
            if region.strip() or zone.strip() or note.strip():
                kept_rows.append({"region": region, "zone": zone, "note": note})
        add_col, _rest = st.columns([1, 4])
        if add_col.button("지역 추가", key=f"bppt_p14add_{year}_{pid}"):
            seq = int(st.session_state.get(seq_key) or 1)
            st.session_state[seq_key] = seq + 1
            person["rows"].append({"id": seq, "region": "", "zone": "", "note": ""})
            st.rerun(scope="fragment")
        staff = str(st.session_state.get(staff_key) or "").strip()
        person["staff"] = staff
        if staff and kept_rows:
            saved.append({"staff": staff, "rows": kept_rows})
    if st.button("구분 추가", key=f"bppt_p14new_{year}"):
        seq = int(st.session_state.get(seq_key) or 1)
        st.session_state[seq_key] = seq + 2
        people.append({"id": seq, "staff": "", "rows": [{"id": seq + 1, "region": "", "zone": "", "note": ""}]})
        st.rerun(scope="fragment")
    return saved


def page14_map_png(year: int = 2026) -> bytes | None:
    path = proposal_path(year)
    if path is None or not path.is_file():
        return None
    try:
        with zipfile.ZipFile(path) as book:
            return book.read("ppt/media/image3.png")
    except (OSError, KeyError, zipfile.BadZipFile):
        return None


def _write_package(path: Path, source: Path, parts: dict[str, bytes]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, parts.get(item.filename, zin.read(item.filename)))
    tmp.replace(path)


def _keep_page14_title(xml: str) -> str:
    """'선제적 접근'이 줄 끝에서 끊기지 않게 제목 칸을 넓히고 한 줄로 잇는다."""
    idx = xml.find(">선제적<")
    if idx < 0:
        return xml
    start = xml.rfind("<p:sp>", 0, idx)
    end = xml.find("</p:sp>", idx)
    if start < 0 or end < 0:
        return xml
    shape = xml[start:end]
    if 'wrap="none"' not in shape:
        shape = shape.replace("<a:bodyPr>", '<a:bodyPr wrap="none">', 1)

    def widen(match: re.Match[str]) -> str:
        width = int(match.group(1))
        height = match.group(2)
        if width >= 22000000:
            return match.group(0)
        return f'<a:ext cx="22000000" cy="{height}"/>'

    shape = re.sub(r'<a:ext cx="(\d+)" cy="(\d+)"/>', widen, shape, count=1)
    return xml[:start] + shape + xml[end:]


def save_page(
    page: int,
    year: int,
    month: int,
    plan: dict[str, int] | None = None,
    tons_plan: dict[str, int] | None = None,
    asan_months: list[int] | None = None,
    hwa_months: list[int] | None = None,
    policy_lines: list[str] | None = None,
    review_lines: list[str] | None = None,
    strategy_lines: list[str] | None = None,
    directive_lines: list[str] | None = None,
    focus_lines: list[str] | None = None,
    task_lines: list[str] | None = None,
    bulk_lines: list[str] | None = None,
    page8_rows: list[dict] | None = None,
    page10_groups: list[dict] | None = None,
    page11_groups: list[dict] | None = None,
    page14_groups: list[dict] | None = None,
    path: Path | None = None,
    template: Path | None = None,
) -> Path:
    """같은 파일이 있으면 그 파일의 해당 페이지만 갱신한다. 없을 때만 양식을 복사한다."""
    dest = path or proposal_path(year)
    if dest is None:
        raise FileNotFoundError("사업계획 폴더를 찾지 못했습니다.")
    dest.parent.mkdir(parents=True, exist_ok=True)
    if not dest.is_file():
        template = template or find_template()
        if template is None or not template.is_file():
            raise FileNotFoundError("사업계획서 양식을 찾지 못했습니다.")
        shutil.copy(template, dest)
    source = dest
    targets = {**PLAN_CHART, **(plan or {})}
    ton_targets = {**PLAN_TONS, **(tons_plan or {})}
    asan_row = list(asan_months) if asan_months is not None else list(PLAN_ASAN)
    hwa_row = list(hwa_months) if hwa_months is not None else list(PLAN_HWA)
    with zipfile.ZipFile(source, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}
    slide_name = f"ppt/slides/slide{int(page)}.xml"
    xml = parts[slide_name].decode("utf-8")
    xml = _map_text(xml, lambda text: _shift_year_text(text, year))
    if page == 2 and any(lines is not None for lines in (policy_lines, review_lines, strategy_lines)):
        cleaned = [
            None if lines is None else [str(line).strip() for line in lines if str(line).strip()]
            for lines in (policy_lines, review_lines, strategy_lines)
        ]
        xml = _apply_toc_lines(xml, year, cleaned[0], cleaned[1], cleaned[2])
    if page == 3 and any(lines is not None for lines in (directive_lines, focus_lines, task_lines)):
        cleaned = [
            None if lines is None else [str(line).strip() for line in lines if str(line).strip()]
            for lines in (directive_lines, focus_lines, task_lines)
        ]
        xml = _apply_goal_lines(xml, cleaned[0], cleaned[1], cleaned[2])
    if page == 4:
        numbers = _slide4_numbers(year, month, targets)
        xml = _map_text(xml, lambda text: numbers.get(text, text))
    elif page == 5:
        xml = _update_tables(xml, [_slide5_cells(year, month, asan_row, hwa_row)])
    elif page == 6:
        xml = _update_tables(xml, _slide6_cells(year, month))
    elif page == 7:
        xml = _update_tables(xml, [_slide7_cells(year, month, ton_targets)])
        xml = _apply_bulk_plan(xml, year, month, ton_targets, bulk_lines)
    elif page == 8:
        if page8_rows is not None:
            xml = _apply_page8(xml, year, month, page8_rows)
        else:
            xml = _update_tables(xml, [_slide8_cells(year, month)])
    elif page == 10 and page10_groups is not None:
        xml = _apply_page10(xml, year, month, page10_groups)
    elif page == 11 and page11_groups is not None:
        xml = _apply_page10(xml, year, month, page11_groups)
    elif page == 14:
        if page14_groups is not None:
            xml = _apply_page14(xml, page14_groups)
        xml = _keep_page14_title(xml)
    parts[slide_name] = xml.encode("utf-8")
    if page == 4:
        cats, series = _chart1(year, month, targets)
        for chart_name in ("ppt/charts/chart1.xml", "ppt/charts/chart2.xml"):
            if chart_name not in parts:
                continue
            chart_xml = parts[chart_name].decode("utf-8")
            if chart_name.endswith("chart1.xml"):
                parts[chart_name] = _patch_chart_xml(chart_xml, cats, series).encode("utf-8")
                embed = _embedding(parts, chart_name)
                if embed and embed in parts:
                    y1a, y2a = series["아산"][1], series["아산"][2]
                    y1h, y2h = series["화성"][1], series["화성"][2]
                    parts[embed] = _patch_workbook(
                        parts[embed],
                        {
                            "A2": f"{year}(P)",
                            "B2": targets["아산"],
                            "C2": targets["화성"],
                            "D2": targets["전체"],
                            "A3": year - 1,
                            "B3": y1a,
                            "C3": y1h,
                            "D3": series["전체"][1],
                            "E3": _prior_total(year - 1),
                            "A4": year - 2,
                            "B4": y2a,
                            "C4": y2h,
                            "D4": series["전체"][2],
                            "E4": _prior_total(year - 2),
                        },
                    )
            else:
                y3 = plant_won(year - 3, month)
                a3 = sum(millions(y3["아산"], month))
                h3 = sum(millions(y3["화성"], month))
                parts[chart_name] = _patch_chart_xml(
                    chart_xml,
                    [f"{year - 1}(P)", str(year - 2), str(year - 3)],
                    {"아산": [0, series["아산"][2], a3], "화성": [0, series["화성"][2], h3], "전체": [_prior_total(year - 1), series["전체"][2], a3 + h3]},
                ).encode("utf-8")
                embed = _embedding(parts, chart_name)
                if embed and embed in parts:
                    parts[embed] = _patch_workbook(
                        parts[embed],
                        {
                            "A2": f"{year - 1}(P)",
                            "D2": _prior_total(year - 1),
                            "A3": year - 2,
                            "B3": series["아산"][2],
                            "C3": series["화성"][2],
                            "D3": series["전체"][2],
                            "A4": year - 3,
                            "B4": a3,
                            "C4": h3,
                            "D4": a3 + h3,
                        },
                    )
    if page == 7:
        cats, series = _chart3(year, month, ton_targets)
        chart_name = "ppt/charts/chart3.xml"
        if chart_name in parts:
            parts[chart_name] = _patch_chart_xml(parts[chart_name].decode("utf-8"), cats, series).encode("utf-8")
            embed = _embedding(parts, chart_name)
            if embed and embed in parts:
                older = tons_of(year - 2, month)
                mid = tons_of(year - 1, month)
                parts[embed] = _patch_workbook(
                    parts[embed],
                    {
                        "A2": year - 2,
                        "B2": older["탄산"],
                        "C2": older["알곤"],
                        "D2": older["질소"],
                        "E2": older["산소"],
                        "F2": older["LPG"],
                        "A3": year - 1,
                        "B3": mid["탄산"],
                        "C3": mid["알곤"],
                        "D3": mid["질소"],
                        "E3": mid["산소"],
                        "F3": mid["LPG"],
                        "A4": f"{year}년(P)",
                        "B4": int(ton_targets["탄산"]),
                        "C4": int(ton_targets["알곤"]),
                        "D4": int(ton_targets["질소"]),
                        "E4": int(ton_targets["산소"]),
                        "F4": int(ton_targets["LPG"]),
                    },
                )
    _write_package(dest, source, parts)
    return dest


def _plant_frame(year: int, month: int) -> pd.DataFrame:
    won = plant_won(year, month)
    rows = []
    for label in ("아산", "화성", "전체"):
        vals = millions(won[label], month)
        rows.append({"구분": label, **{f"{m}월": vals[m - 1] for m in range(1, 13)}, "합계": annual_million(won[label], month)})
    return pd.DataFrame(rows)


def _edit_page8(year: int, month: int, product: str, rows: list[dict[str, str]]) -> list[dict[str, str]]:
    store = f"bppt_p8_{product}_{year}"
    if store not in st.session_state:
        st.session_state[store] = [
            {"id": index + 1, "client": row["client"], "budget": row["budget"], "note": row["note"]}
            for index, row in enumerate(rows)
        ]
    items = st.session_state[store]
    kept: list[dict[str, str]] = []
    for item in items:
        client_key = f"bppt_p8c_{product}_{year}_{item['id']}"
        budget_key = f"bppt_p8b_{product}_{year}_{item['id']}"
        note_key = f"bppt_p8n_{product}_{year}_{item['id']}"
        if client_key not in st.session_state:
            st.session_state[client_key] = item["client"]
        if budget_key not in st.session_state:
            st.session_state[budget_key] = item["budget"]
        if note_key not in st.session_state:
            st.session_state[note_key] = item["note"]
        client_col, actual_col, budget_col, delta_col, note_col, delete_col = st.columns([2.1, 1.5, 1.5, 1.4, 2.2, 0.7])
        client_col.text_input("거래처", key=client_key, label_visibility="collapsed")
        client = str(st.session_state.get(client_key) or "")
        budget = str(st.session_state.get(budget_key) or "")
        note = str(st.session_state.get(note_key) or "")
        actual_txt, _budget_txt, delta_txt = page8_figures(year, month, product, client, budget)
        actual_col.markdown(actual_txt)
        budget_col.text_input("예상", key=budget_key, label_visibility="collapsed")
        delta_col.markdown(delta_txt)
        note_col.text_input("비고", key=note_key, label_visibility="collapsed")
        if delete_col.button("삭제", key=f"bppt_p8del_{product}_{year}_{item['id']}"):
            st.session_state[store] = [row for row in items if row["id"] != item["id"]]
            st.rerun(scope="fragment")
        item["client"] = client
        item["budget"] = budget
        item["note"] = note
        if client.strip():
            kept.append({"product": product, "client": client.strip(), "budget": budget.strip(), "note": note.strip()})
    if st.button("거래처 추가", key=f"bppt_p8add_{product}_{year}"):
        nxt = max((item["id"] for item in items), default=0) + 1
        items.append({"id": nxt, "client": "", "budget": "", "note": ""})
        st.rerun(scope="fragment")
    return kept


def _edit_page10(year: int, month: int, groups: list[dict], page: int = 10) -> list[dict]:
    prefix = f"bppt_p{int(page)}"
    store = f"{prefix}_{year}"
    if store not in st.session_state:
        seeded = []
        for index, group in enumerate(groups, start=1):
            seeded.append(
                {
                    "id": index,
                    "staff": group.get("staff") or "",
                    "news": [
                        {"id": nindex, "client": item.get("client") or "", "gas": item.get("gas") or ""}
                        for nindex, item in enumerate(group.get("news") or [], start=1)
                    ],
                    "growth": [
                        {"id": gindex, "client": item.get("client") or ""}
                        for gindex, item in enumerate(group.get("growth") or [], start=1)
                    ],
                }
            )
        st.session_state[store] = seeded
    people = st.session_state[store]
    saved: list[dict] = []
    for person in people:
        pid = person["id"]
        staff_key = f"{prefix}s_{year}_{pid}"
        if staff_key not in st.session_state:
            st.session_state[staff_key] = person.get("staff") or ""
        st.markdown("---")
        name_col, delete_col = st.columns([6, 1])
        name_col.text_input("담당", key=staff_key)
        staff = str(st.session_state.get(staff_key) or "").strip()
        person["staff"] = staff
        staff_key_name = staff_key_in_label(staff)
        if delete_col.button("담당 삭제", key=f"{prefix}del_{year}_{pid}"):
            st.session_state[store] = [row for row in people if row["id"] != pid]
            st.rerun(scope="fragment")
        st.markdown("**신규**")
        newcomers = staff_new_clients(year, month, staff_key_name) if staff_key_name else []
        if newcomers:
            st.caption(f"{year - 1}년에 새로 거래한 업체 {len(newcomers)}곳입니다. 전체 적용을 누르면 업체명과 가스종류가 들어갑니다.")
            st.dataframe(
                pd.DataFrame(
                    {
                        "거래처": [row["client"] for row in newcomers],
                        "가스종류": [row["gas"] for row in newcomers],
                        f"{year - 1}년(백만)": [row["won"] for row in newcomers],
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )
            apply_col, pick_col, one_col = st.columns([1.2, 3, 1])
            if apply_col.button("전체 적용", key=f"{prefix}nall_{year}_{pid}"):
                already = {_fold_name(row.get("client")) for row in person["news"]}
                nxt = max((item["id"] for item in person["news"]), default=0)
                for row in newcomers:
                    folded = _fold_name(row["client"])
                    if not folded or folded in already:
                        continue
                    nxt += 1
                    already.add(folded)
                    person["news"].append({"id": nxt, "client": row["client"], "gas": row["gas"]})
                    st.session_state[f"{prefix}n_{year}_{pid}_{nxt}"] = row["client"]
                    st.session_state[f"{prefix}g_{year}_{pid}_{nxt}"] = row["gas"]
                st.rerun(scope="fragment")
            new_options = [row["client"] for row in newcomers]
            new_labels = {row["client"]: f"{row['client']}  {row['gas']}" for row in newcomers}
            pick_col.selectbox(
                "신규 업체",
                new_options,
                format_func=lambda name: new_labels.get(name, name),
                key=f"{prefix}npick_{year}_{pid}",
            )
            if one_col.button("선택 추가", key=f"{prefix}none_{year}_{pid}"):
                chosen = str(st.session_state.get(f"{prefix}npick_{year}_{pid}") or "").strip()
                gas = next((row["gas"] for row in newcomers if row["client"] == chosen), "")
                already = {_fold_name(row.get("client")) for row in person["news"]}
                if chosen and _fold_name(chosen) not in already:
                    nxt = max((item["id"] for item in person["news"]), default=0) + 1
                    person["news"].append({"id": nxt, "client": chosen, "gas": gas})
                    st.session_state[f"{prefix}n_{year}_{pid}_{nxt}"] = chosen
                    st.session_state[f"{prefix}g_{year}_{pid}_{nxt}"] = gas
                    st.rerun(scope="fragment")
        elif staff_key_name:
            st.caption(f"{year - 1}년에 전년 대비 신규 업체가 없습니다.")
        else:
            st.caption("담당 이름에 김명현처럼 적으면 신규 업체가 나옵니다.")
        news_kept = []
        for item in person["news"]:
            client_key = f"{prefix}n_{year}_{pid}_{item['id']}"
            gas_key = f"{prefix}g_{year}_{pid}_{item['id']}"
            if client_key not in st.session_state:
                st.session_state[client_key] = item.get("client") or ""
            known_gas = client_gas_text(year, month, str(st.session_state.get(client_key) or item.get("client") or ""))
            if gas_key not in st.session_state:
                st.session_state[gas_key] = item.get("gas") or known_gas
            elif not str(st.session_state.get(gas_key) or "").strip() and known_gas:
                st.session_state[gas_key] = known_gas
            client_col, gas_col, row_delete = st.columns([3, 2, 0.8])
            client_col.text_input("업체명", key=client_key, label_visibility="collapsed")
            gas_col.text_input("가스종류", key=gas_key, label_visibility="collapsed")
            if row_delete.button("삭제", key=f"{prefix}ndel_{year}_{pid}_{item['id']}"):
                person["news"] = [row for row in person["news"] if row["id"] != item["id"]]
                st.rerun(scope="fragment")
            client = str(st.session_state.get(client_key) or "").strip()
            gas = str(st.session_state.get(gas_key) or "").strip()
            item["client"] = client
            item["gas"] = gas
            if client or gas:
                news_kept.append({"client": client, "gas": gas})
        if st.button("신규 추가", key=f"{prefix}nadd_{year}_{pid}"):
            nxt = max((item["id"] for item in person["news"]), default=0) + 1
            person["news"].append({"id": nxt, "client": "", "gas": ""})
            st.rerun(scope="fragment")
        st.markdown("**물량증대**")
        ranks = staff_top_clients(year, month, staff_key_name) if staff_key_name else []
        if ranks:
            st.dataframe(
                pd.DataFrame(
                    {
                        "순위": [row["rank"] for row in ranks],
                        "거래처": [row["client"] for row in ranks],
                        f"{year - 2}년(백만)": [row["prev"] for row in ranks],
                        f"{year - 1}년(백만)": [row["curr"] for row in ranks],
                        "물량증감": [row["change"] for row in ranks],
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )
            pick_col, add_col = st.columns([4, 1])
            pick_key = f"{prefix}pick_{year}_{pid}"
            options = [row["client"] for row in ranks]
            labels = {row["client"]: f"{row['rank']}위 {row['client']}  {row['change']}" for row in ranks}
            pick_col.selectbox("1~10위", options, format_func=lambda name: labels.get(name, name), key=pick_key)
            if add_col.button("선택 추가", key=f"{prefix}gadd_{year}_{pid}"):
                chosen = str(st.session_state.get(pick_key) or "").strip()
                already = {str(row.get("client") or "").strip() for row in person["growth"]}
                if chosen and chosen not in already:
                    nxt = max((item["id"] for item in person["growth"]), default=0) + 1
                    person["growth"].append({"id": nxt, "client": chosen})
                    st.rerun(scope="fragment")
        else:
            st.caption("담당 이름에 김명현처럼 적으면 그 담당 1~10위가 나옵니다.")
        growth_kept = []
        for item in person["growth"]:
            grow_key = f"{prefix}c_{year}_{pid}_{item['id']}"
            if grow_key not in st.session_state:
                st.session_state[grow_key] = item.get("client") or ""
            client_col, delta_col, row_delete = st.columns([3, 1.6, 0.8])
            client_col.text_input("물량증대 업체", key=grow_key, label_visibility="collapsed")
            client = str(st.session_state.get(grow_key) or "").strip()
            item["client"] = client
            delta_col.markdown(volume_change_text(year, month, client) if client else "\u00a0")
            if row_delete.button("삭제", key=f"{prefix}gdel_{year}_{pid}_{item['id']}"):
                person["growth"] = [row for row in person["growth"] if row["id"] != item["id"]]
                st.rerun(scope="fragment")
            if client:
                growth_kept.append({"client": client})
        if staff and (news_kept or growth_kept):
            saved.append({"staff": staff, "news": news_kept, "growth": growth_kept})
    if st.button("담당 추가", key=f"{prefix}add_{year}"):
        nxt = max((person["id"] for person in people), default=0) + 1
        people.append({"id": nxt, "staff": "", "news": [], "growth": []})
        st.rerun(scope="fragment")
    return saved


def _edit_lines(year: int, slot: str, lines: list[str], page: int = 2) -> list[str]:
    store = f"bppt_s{page}_{slot}_{year}"
    if store not in st.session_state:
        st.session_state[store] = [{"id": index + 1, "text": line} for index, line in enumerate(lines)]
    items = st.session_state[store]
    for item in items:
        key = f"bppt_s{page}_{slot}_{year}_{item['id']}"
        if key in st.session_state:
            item["text"] = st.session_state[key]
    kept = [str(item["text"]).strip() for item in items if str(item["text"]).strip()]
    for item in items:
        key = f"bppt_s{page}_{slot}_{year}_{item['id']}"
        if key not in st.session_state:
            st.session_state[key] = item["text"]
        left, right = st.columns([8, 1])
        left.text_input(slot, key=key, label_visibility="collapsed")
        if right.button("삭제", key=f"bppt_s{page}_del_{slot}_{year}_{item['id']}"):
            st.session_state[store] = [row for row in items if row["id"] != item["id"]]
            st.session_state.pop(key, None)
            st.rerun(scope="fragment")
    if st.button("항목 추가", key=f"bppt_s{page}_add_{slot}_{year}"):
        nxt = max((item["id"] for item in items), default=0) + 1
        items.append({"id": nxt, "text": ""})
        st.rerun(scope="fragment")
    return kept


_PREVIEW_DIR = Path(__file__).resolve().parent / ".bppt_preview"
_RENDER_SWIFT = r'''import Foundation
import PDFKit
import AppKit

let args = CommandLine.arguments
guard args.count == 4, let pageNo = Int(args[2]), let doc = PDFDocument(url: URL(fileURLWithPath: args[1])), let page = doc.page(at: pageNo - 1) else {
    fputs("usage: bppt_render pdf page png\n", stderr)
    exit(1)
}
let bounds = page.bounds(for: .mediaBox)
let scale: CGFloat = 1.6
let size = NSSize(width: bounds.width * scale, height: bounds.height * scale)
let image = NSImage(size: size)
image.lockFocus()
NSColor.white.set()
NSRect(origin: .zero, size: size).fill()
if let ctx = NSGraphicsContext.current?.cgContext {
    ctx.saveGState()
    ctx.scaleBy(x: scale, y: scale)
    page.draw(with: .mediaBox, to: ctx)
    ctx.restoreGState()
}
image.unlockFocus()
guard let tiff = image.tiffRepresentation, let rep = NSBitmapImageRep(data: tiff), let png = rep.representation(using: .png, properties: [:]) else {
    fputs("render failed\n", stderr)
    exit(2)
}
try png.write(to: URL(fileURLWithPath: args[3]))
'''


def _preview_base() -> Path:
    found = find_template()
    if found is not None and found.is_file():
        return found
    for year in (2026, 2027, 2025):
        path = proposal_path(year)
        if path is not None and path.is_file():
            return path
    raise FileNotFoundError("미리보기에 쓸 사업계획서 양식을 찾지 못했습니다.")


def _ensure_renderer() -> Path:
    _PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    binary = _PREVIEW_DIR / "bppt_render"
    baked = Path("/tmp/bppt_render")
    if baked.is_file() and not binary.is_file():
        shutil.copy(baked, binary)
        binary.chmod(0o755)
    if binary.is_file():
        return binary
    source = _PREVIEW_DIR / "render.swift"
    source.write_text(_RENDER_SWIFT, encoding="utf-8")
    subprocess.run(["swiftc", "-O", str(source), "-o", str(binary)], check=True, timeout=180)
    return binary


def _export_pdf(pptx: Path) -> Path:
    name = "bppt_preview.pdf"
    candidates = [
        Path("/tmp") / name,
        Path.home() / "Library/Containers/com.microsoft.Powerpoint/Data/tmp" / name,
    ]
    for path in candidates:
        try:
            if path.is_file():
                path.unlink()
        except OSError:
            pass
    script = (
        'tell application "Microsoft PowerPoint"\n'
        f'open POSIX file "{pptx}"\n'
        f'set pres to presentation "{pptx.name}"\n'
        f'save pres in "/tmp/{name}" as save as PDF\n'
        "close pres saving no\n"
        "end tell\n"
    )
    subprocess.run(["osascript"], input=script, text=True, check=True, timeout=90)
    found = [path for path in candidates if path.is_file()]
    if not found:
        raise FileNotFoundError("미리보기 PDF를 만들지 못했습니다.")
    return max(found, key=lambda path: path.stat().st_mtime)


def _build_preview_deck(specs: list[dict], dest: Path) -> None:
    base = _preview_base()
    if dest.exists():
        dest.unlink()
    for spec in specs:
        save_page(
            int(spec["page"]),
            int(spec["year"]),
            int(spec["month"]),
            plan=spec["plan"],
            tons_plan=spec["tons"],
            asan_months=spec["asan"],
            hwa_months=spec["hwa"],
            policy_lines=spec.get("policy"),
            review_lines=spec.get("review"),
            strategy_lines=spec.get("strategy"),
            directive_lines=spec.get("directive"),
            focus_lines=spec.get("focus"),
            task_lines=spec.get("task"),
            bulk_lines=spec.get("bulk"),
            page8_rows=spec.get("page8"),
            page10_groups=spec.get("page10"),
            page11_groups=spec.get("page11"),
            page14_groups=spec.get("map") if int(spec.get("page") or 0) == 14 else None,
            path=dest,
            template=base,
        )


def _render_pngs(pdf: Path) -> dict[int, bytes]:
    renderer = _ensure_renderer()
    folder = _PREVIEW_DIR / "pages"
    folder.mkdir(parents=True, exist_ok=True)
    images: dict[int, bytes] = {}
    for page in PAGES:
        png = folder / f"{page}.png"
        subprocess.run([str(renderer), str(pdf), str(page), str(png)], check=True, timeout=30)
        images[page] = png.read_bytes()
    return images


def _preview_token(specs: list[dict]) -> str:
    return json.dumps(specs, ensure_ascii=False, sort_keys=True)


def _changed_preview_pages(old: object, new: list[dict]) -> list[int] | None:
    if not isinstance(old, list) or len(old) != len(new):
        return None
    changed: list[int] = []
    for before, after in zip(old, new, strict=False):
        if before != after:
            changed.append(int(after.get("page", 0)))
    return changed


def _saved_preview_specs() -> list | None:
    raw = st.session_state.get("_bppt_preview_token")
    if not isinstance(raw, str) or not raw:
        return None
    try:
        specs = json.loads(raw)
    except json.JSONDecodeError:
        return None
    return specs if isinstance(specs, list) else None


def _render_one_png(pdf: Path, page: int) -> bytes:
    renderer = _ensure_renderer()
    folder = _PREVIEW_DIR / "pages"
    folder.mkdir(parents=True, exist_ok=True)
    png = folder / f"{page}.png"
    subprocess.run([str(renderer), str(pdf), str(page), str(png)], check=True, timeout=30)
    return png.read_bytes()


def _refresh_one_page(spec: dict) -> None:
    """이미 만든 미리보기에서 그 페이지만 고친다."""
    deck = _PREVIEW_DIR / "preview.pptx"
    page = int(spec["page"])
    save_page(
        page,
        int(spec["year"]),
        int(spec["month"]),
        plan=spec.get("plan"),
        tons_plan=spec.get("tons"),
        asan_months=spec.get("asan"),
        hwa_months=spec.get("hwa"),
        policy_lines=spec.get("policy"),
        review_lines=spec.get("review"),
        strategy_lines=spec.get("strategy"),
        directive_lines=spec.get("directive"),
        focus_lines=spec.get("focus"),
        task_lines=spec.get("task"),
        bulk_lines=spec.get("bulk"),
        page8_rows=spec.get("page8"),
        page10_groups=spec.get("page10"),
        page11_groups=spec.get("page11"),
        page14_groups=spec.get("map") if int(spec.get("page") or 0) == 14 else None,
        path=deck,
        template=deck,
    )
    pngs = dict(st.session_state.get("_bppt_preview_pngs") or {})
    pngs[page] = _render_one_png(_export_pdf(deck), page)
    st.session_state["_bppt_preview_pngs"] = pngs


def _refresh_previews(specs: list[dict]) -> None:
    token = _preview_token(specs)
    if st.session_state.get("_bppt_preview_token") == token:
        return
    changed = _changed_preview_pages(_saved_preview_specs(), specs)
    deck = _PREVIEW_DIR / "preview.pptx"
    try:
        _PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
        if changed and len(changed) == 1 and deck.is_file():
            spec = next(item for item in specs if int(item.get("page", 0)) == changed[0])
            with st.spinner(f"{changed[0]}페이지 미리보기를 맞추고 있습니다."):
                _refresh_one_page(spec)
        else:
            with st.spinner("슬라이드 미리보기를 만들고 있습니다."):
                _build_preview_deck(specs, deck)
                st.session_state["_bppt_preview_pngs"] = _render_pngs(_export_pdf(deck))
        st.session_state["_bppt_preview_error"] = ""
    except Exception as exc:
        st.session_state["_bppt_preview_pngs"] = {}
        st.session_state["_bppt_preview_error"] = f"미리보기를 만들지 못했습니다. {exc}"
    st.session_state["_bppt_preview_token"] = token
    from streamlit.runtime.scriptrunner_utils.script_run_context import get_script_run_ctx

    ctx = get_script_run_ctx()
    if ctx is not None and ctx.fragment_ids_this_run:
        st.rerun(scope="fragment")
    st.rerun()


def _show_page_preview(page: int) -> None:
    pngs = st.session_state.get("_bppt_preview_pngs") or {}
    image = pngs.get(page)
    if isinstance(image, (bytes, bytearray)):
        st.image(image, use_container_width=True)
        st.caption("수정이 반영된 이 페이지입니다. 파일에 남기려면 저장을 누릅니다.")
        return
    error = st.session_state.get("_bppt_preview_error") or ""
    st.caption(error or "이 페이지 미리보기를 준비하고 있습니다.")


@st.fragment
def render_biz_proposal_tab(_sales_df: pd.DataFrame | None = None) -> None:
    """로컬 사업계획서. 16페이지가 각각 방이고, 저장은 그 페이지만 반영한다."""
    st.markdown("##### 사업계획서")
    st.caption("가스코아산은 아산, 그 외는 화성입니다. 금액은 부가세 포함이고, 월은 그 달까지의 실적입니다. 4페이지 2026년은 목표라서 고칠 수 있습니다. 저장 파일은 26년 사업계획서.pptx이고, 같은 파일이 있으면 그 파일에 이어서 넣습니다.")
    rooms = st.tabs([f"{page}페이지" for page in PAGES])
    preview_specs: list[dict] = []
    for page, room in zip(PAGES, rooms, strict=False):
        with room:
            edit_col, view_col = st.columns([1, 1.15])
            with edit_col:
                left, mid, _right = st.columns([1, 1, 2])
                year = int(left.number_input("년도", min_value=2021, max_value=2035, value=2026, step=1, key=f"bppt_y_{page}"))
                month = int(mid.selectbox("월", list(range(1, 13)), index=11, format_func=lambda value: f"{value}월", key=f"bppt_m_{page}"))
                plan = dict(PLAN_CHART)
                tons_plan = dict(PLAN_TONS)
                asan_months = list(PLAN_ASAN)
                hwa_months = list(PLAN_HWA)
                policy_lines = None
                review_lines = None
                strategy_lines = None
                directive_lines = None
                focus_lines = None
                task_lines = None
                bulk_lines = None
                page8_rows = None
                page10_groups = None
                page11_groups = None
                page14_groups = None
                if page == 4:
                    st.dataframe(_plant_frame(year - 2, month).assign(연도=year - 2), hide_index=True, use_container_width=True)
                    st.dataframe(_plant_frame(year - 1, month).assign(연도=year - 1), hide_index=True, use_container_width=True)
                    c1, c2, c3 = st.columns(3)
                    plan["아산"] = int(c1.number_input("아산 목표(백만원)", min_value=0, value=PLAN_CHART["아산"], step=1, key=f"bppt_asan_{year}"))
                    plan["화성"] = int(c2.number_input("화성 목표(백만원)", min_value=0, value=PLAN_CHART["화성"], step=1, key=f"bppt_hwa_{year}"))
                    plan["전체"] = int(c3.number_input("전체 목표(백만원)", min_value=0, value=PLAN_CHART["전체"], step=1, key=f"bppt_all_{year}"))
                    st.bar_chart(
                        pd.DataFrame(
                            {"아산": [plan["아산"], annual_million(plant_won(year - 1, month)["아산"], month), annual_million(plant_won(year - 2, month)["아산"], month)],
                             "화성": [plan["화성"], annual_million(plant_won(year - 1, month)["화성"], month), annual_million(plant_won(year - 2, month)["화성"], month)]},
                            index=[f"{year}(P)", str(year - 1), str(year - 2)],
                        )
                    )
                elif page == 5:
                    st.dataframe(_plant_frame(year - 2, month), hide_index=True, use_container_width=True)
                    st.dataframe(_plant_frame(year - 1, month), hide_index=True, use_container_width=True)
                    grid = pd.DataFrame({"월": list(range(1, 13)), "화성": hwa_months, "아산": asan_months})
                    edited = st.data_editor(grid, hide_index=True, num_rows="fixed", key=f"bppt_grid_{year}", use_container_width=True)
                    if isinstance(edited, pd.DataFrame) and {"화성", "아산"} <= set(edited.columns):
                        hwa_months = [int(v) for v in edited["화성"].tolist()]
                        asan_months = [int(v) for v in edited["아산"].tolist()]
                    st.caption(f"{year}년 목표 합계 화성 {sum(hwa_months):,} · 아산 {sum(asan_months):,}")
                elif page == 6:
                    st.dataframe(pd.DataFrame(item_sheet(year - 1, month)), hide_index=True, use_container_width=True)
                elif page == 7:
                    older = tons_of(year - 2, month)
                    mid = tons_of(year - 1, month)
                    st.dataframe(
                        pd.DataFrame(
                            [
                                {"품목": label, str(year - 2): older[label], str(year - 1): mid[label], f"{year}(P)": PLAN_TONS[label]}
                                for label in ("산소", "질소", "알곤", "탄산", "LPG", "합계")
                                if label != "합계"
                            ]
                        ),
                        hide_index=True,
                        use_container_width=True,
                    )
                    cols = st.columns(5)
                    for col, label in zip(cols, ("산소", "질소", "알곤", "탄산", "LPG"), strict=False):
                        tons_plan[label] = int(col.number_input(f"{label} 목표(톤)", min_value=0, value=PLAN_TONS[label], step=1, key=f"bppt_ton_{year}_{label}"))
                    pct = bulk_growth_percent(year, month, tons_plan)
                    st.caption(f"벌크운영 계획 문안을 고치고 줄을 추가하거나 삭제할 수 있습니다. 삼각형 {pct}%는 {year - 1}년 대비 {year}년 계획 합계의 증가율입니다.")
                    bulk_store = f"bppt_s7_bulk_{year}"
                    if bulk_store in st.session_state:
                        for item in st.session_state[bulk_store]:
                            item["text"] = _tune_bulk_line(item["text"], year, pct)
                            item_key = f"bppt_s7_bulk_{year}_{item['id']}"
                            if item_key in st.session_state:
                                st.session_state[item_key] = _tune_bulk_line(st.session_state[item_key], year, pct)
                    bulk_now = [_tune_bulk_line(line, year, pct) for line in read_bulk_lines(year)]
                    st.markdown(f"**{year}년 벌크운영 계획**")
                    bulk_lines = _edit_lines(year, "bulk", bulk_now, page=7)
                elif page == 8:
                    st.caption("거래처를 넣으면 그 해 실적이 집계됩니다. 예상과 비고는 비워 둘 수 있고, 증감량은 예상이 있으면 예상에서 실적을 뺀 값입니다.")
                    loaded = read_page8_rows(year)
                    page8_rows = []
                    head_client, head_actual, head_budget, head_delta, head_note, _head_delete = st.columns([2.1, 1.5, 1.5, 1.4, 2.2, 0.7])
                    head_client.markdown("**거래처**")
                    head_actual.markdown(f"**{year - 1} 실적**")
                    head_budget.markdown("**예상**")
                    head_delta.markdown("**증감**")
                    head_note.markdown("**비고**")
                    for product in PAGE8_PRODUCTS:
                        st.markdown(f"**{product}**")
                        page8_rows.extend(
                            _edit_page8(year, month, product, [row for row in loaded if row["product"] == product])
                        )
                elif page == 14:
                    st.caption("구분, 담당지역, 산업단지를 고치고 줄을 넣거나 뺄 수 있습니다. 오른쪽 슬라이드는 바로 바뀌고, 저장을 누르면 파일에 들어갑니다.")
                    page14_groups = _edit_page14(year, read_page14_groups(year))
                    st.markdown("**담당 지도**")
                    map_png = page14_map_png(year)
                    if map_png:
                        st.image(map_png, use_container_width=True)
                elif page in (10, 11):
                    st.caption("담당별로 신규 업체와 물량증대를 넣고 뺄 수 있습니다. 신규는 전년 대비 새 업체와 가스종류를 불러오고, 물량증대는 1~10위를 고르면 증감이 계산됩니다. 비고는 비워 둡니다.")
                    loaded_groups = read_page10_groups(year, page)
                    if page == 10:
                        page10_groups = _edit_page10(year, month, loaded_groups, page=10)
                    else:
                        page11_groups = _edit_page10(year, month, loaded_groups, page=11)
                elif page == 2:
                    yy = int(year) % 100
                    prev = (int(year) - 1) % 100
                    st.caption("1번, 2번, 3번 항목을 고치고 줄을 추가하거나 삭제할 수 있습니다.")
                    policy_now, review_now, strategy_now = read_toc_lines(year)
                    st.markdown(f"**1. {POLICY_TITLE}**")
                    policy_lines = _edit_lines(year, "policy", policy_now)
                    st.markdown(f"**2. {yy}년 사업계획 대비 {prev}년 실적**")
                    review_lines = _edit_lines(year, "review", review_now)
                    st.markdown(f"**3. {yy:02d}{STRATEGY_TITLE}**")
                    strategy_lines = _edit_lines(year, "strategy", strategy_now)
                elif page == 3:
                    st.caption("1번, 2번, 3번 항목을 고치고 줄을 추가하거나 삭제할 수 있습니다.")
                    directive_now, focus_now, task_now = read_goal_lines(year)
                    st.markdown(f"**1. {DIRECTIVE_LABEL}**")
                    directive_lines = _edit_lines(year, "directive", directive_now, page=3)
                    st.markdown(f"**2. {FOCUS_LABEL}**")
                    focus_lines = _edit_lines(year, "focus", focus_now, page=3)
                    st.markdown(f"**3. {TASK_LABEL}**")
                    task_lines = _edit_lines(year, "task", task_now, page=3)
                else:
                    st.caption(f"이 페이지는 문안의 년도를 {year}년 기준으로 바꿉니다. 표 서식은 그대로입니다.")
                if st.button(f"{page}페이지 저장", key=f"bppt_save_{page}"):
                    try:
                        saved = save_page(
                            page,
                            year,
                            month,
                            plan=plan,
                            tons_plan=tons_plan,
                            asan_months=asan_months,
                            hwa_months=hwa_months,
                            policy_lines=policy_lines,
                            review_lines=review_lines,
                            strategy_lines=strategy_lines,
                            directive_lines=directive_lines,
                            focus_lines=focus_lines,
                            task_lines=task_lines,
                            bulk_lines=bulk_lines,
                            page8_rows=page8_rows,
                            page10_groups=page10_groups,
                            page11_groups=page11_groups,
                            page14_groups=page14_groups,
                        )
                        st.success(f"{page}페이지를 저장했습니다. {saved}")
                    except FileNotFoundError as exc:
                        st.error(str(exc))
                    except Exception as exc:
                        st.error(f"{page}페이지 저장 오류: {exc}")
            with view_col:
                _show_page_preview(page)
            preview_specs.append(
                {
                    "page": page,
                    "year": int(year),
                    "month": int(month),
                    "plan": dict(plan) if page == 4 else None,
                    "tons": dict(tons_plan) if page == 7 else None,
                    "asan": list(asan_months) if page == 5 else None,
                    "hwa": list(hwa_months) if page == 5 else None,
                    "policy": list(policy_lines) if page == 2 and policy_lines is not None else None,
                    "review": list(review_lines) if page == 2 and review_lines is not None else None,
                    "strategy": list(strategy_lines) if page == 2 and strategy_lines is not None else None,
                    "directive": list(directive_lines) if page == 3 and directive_lines is not None else None,
                    "focus": list(focus_lines) if page == 3 and focus_lines is not None else None,
                    "task": list(task_lines) if page == 3 and task_lines is not None else None,
                    "bulk": list(bulk_lines) if page == 7 and bulk_lines is not None else None,
                    "page8": list(page8_rows) if page == 8 and page8_rows is not None else None,
                    "page10": list(page10_groups) if page == 10 and page10_groups is not None else None,
                    "page11": list(page11_groups) if page == 11 and page11_groups is not None else None,
                    "map": list(page14_groups) if page == 14 and page14_groups is not None else None,
                }
            )
    if preview_specs:
        _refresh_previews(preview_specs)
