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


def client_kg(year: int, month: int, client: str, product: str) -> float | None:
    if not client or "신규" in client or client.strip() in {"-", "—"}:
        return None
    key = _CLIENT_GAS.get(re.sub(r"[^A-Za-z0-9]", "", product))
    if key is None:
        return None
    frame = _through(load_year(year), month)
    if frame.empty:
        return 0.0
    mask = frame["거래처"].str.contains(re.escape(client.strip()), na=False)
    picked = frame[mask]
    if picked.empty:
        return 0.0
    total = 0.0
    for name, qty in zip(picked["품목명"], picked["수량"], strict=False):
        if _item_key(str(name)) == f"B:{key}":
            total += _qty_kg(str(name), float(qty))
    return total


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


def _write_package(path: Path, source: Path, parts: dict[str, bytes]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, parts.get(item.filename, zin.read(item.filename)))
    tmp.replace(path)


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
    elif page == 8:
        xml = _update_tables(xml, [_slide8_cells(year, month)])
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


def _refresh_previews(specs: list[dict]) -> None:
    token = _preview_token(specs)
    if st.session_state.get("_bppt_preview_token") == token:
        return
    try:
        _PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
        deck = _PREVIEW_DIR / "preview.pptx"
        with st.spinner("슬라이드 미리보기를 만들고 있습니다."):
            _build_preview_deck(specs, deck)
            pngs = _render_pngs(_export_pdf(deck))
        st.session_state["_bppt_preview_pngs"] = pngs
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
                elif page == 8:
                    preview = []
                    for product, client, budget, note in _page8_template():
                        actual = client_kg(year - 1, month, client, product)
                        preview.append(
                            {
                                "품목": product,
                                "거래처": client,
                                f"{year - 1} 실적(kg)": None if actual is None else int(round(actual)),
                                f"{year} 예산(kg)": budget,
                                "비고": note,
                            }
                        )
                    st.dataframe(pd.DataFrame(preview), hide_index=True, use_container_width=True)
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
                }
            )
    if preview_specs:
        _refresh_previews(preview_specs)
