"""시장조사 탭 — Drive「Desktop/업무/시장조사」자료를 지역·공급사·소스로 정리해 조회."""
from __future__ import annotations

import os
import re
import shutil
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    from openpyxl import load_workbook
except Exception:  # pragma: no cover
    load_workbook = None

MR_CACHE_DIR = os.path.join("uploaded_cache", "market_research")
MR_DRIVE_CANDIDATES = [
    os.path.expanduser(
        "~/Library/CloudStorage/GoogleDrive-3023526@gmail.com/"
        "다른 컴퓨터/내 컴퓨터 (1)/Desktop/업무/시장조사"
    ),
    os.path.expanduser(
        "~/Library/CloudStorage/GoogleDrive-3023526@gmail.com/"
        "다른 컴퓨터/내 컴퓨터/업무/시장조사"
    ),
    os.path.join("Desktop", "업무", "시장조사"),
]

_SKIP_SHEET = re.compile(
    r"스케줄|원장|용기재고|재고현황|Sheet1\s*\(|^\d+$",
    re.I,
)

_REGION_RULES: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"화성|동탄|향남|팔탄|마도|봉담|우정|장안|비봉|양감|정남|서신"), "화성"),
    (re.compile(r"평택|송탄|팽성|청북|포승|안중"), "평택"),
    (re.compile(r"안성"), "안성"),
    (re.compile(r"오산"), "오산"),
    (re.compile(r"용인|처인|기흥|수지도"), "용인"),
    (re.compile(r"수원"), "수원"),
    (re.compile(r"안산|단원|상록"), "안산"),
    (re.compile(r"시흥|정왕|배곧"), "시흥"),
    (re.compile(r"광명"), "광명"),
    (re.compile(r"군포"), "군포"),
    (re.compile(r"의왕"), "의왕"),
    (re.compile(r"안양|동안|만안"), "안양"),
    (re.compile(r"과천"), "과천"),
    (re.compile(r"성남|분당|판교|수정구|중원구"), "성남"),
    (re.compile(r"광주(?:시|도)?(?!\s*광역시)|경기\s*광주"), "경기광주"),
    (re.compile(r"이천"), "이천"),
    (re.compile(r"여주"), "여주"),
    (re.compile(r"김포"), "김포"),
    (re.compile(r"부천"), "부천"),
    (re.compile(r"인천|남동|연수|부평|계양|서구\s*금곡|미추홀"), "인천"),
    (re.compile(r"서울|금천|가산|구로|영등포|강서|송파"), "서울"),
    (re.compile(r"당진"), "당진"),
    (re.compile(r"서산|태안|홍성|예산|아산|천안|공주|보령|논산|부여"), "충남"),
    (re.compile(r"음성|진천|증평|충주|청주|제천|옥천|영동|괴산"), "충북"),
    (re.compile(r"대전"), "대전"),
    (re.compile(r"세종"), "세종"),
    (re.compile(r"부산|강서구\s*미음|사하|사상"), "부산"),
    (re.compile(r"창원|진해|마산|김해|양산|함안|거제|통영|진주"), "경남"),
    (re.compile(r"울산"), "울산"),
    (re.compile(r"대구"), "대구"),
    (re.compile(r"광주광|광주시\s*서구|남구\s*서문대"), "광주"),
    (re.compile(r"여수|광양|순천|영암|무안|나주|목포"), "전남"),
    (re.compile(r"전주|군산|익산|완주"), "전북"),
    (re.compile(r"강원|원주|춘천|강릉"), "강원"),
    (re.compile(r"충남|충청남"), "충남"),
    (re.compile(r"충북|충청북"), "충북"),
    (re.compile(r"경기"), "경기기타"),
]


def _s(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    t = str(v).strip()
    if t.lower() in {"nan", "none", "null"}:
        return ""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", t))


def _company_key(name: str) -> str:
    """중복 병합용 업체 키 — ㈜/주식회사/공백·괄호 제거."""
    t = _s(name)
    t = re.sub(
        r"(주식회사|유한회사|유한책임회사|\(주\)|㈜|㈔|㈜)",
        "",
        t,
        flags=re.I,
    )
    t = re.sub(r"[\s\(\)（）\[\]【】·\.\-_/,，、xX×]+", "", t)
    return t.casefold()


def _merge_unique_text(values, *, sep: str = " · ", max_parts: int = 12) -> str:
    seen: set[str] = set()
    parts: list[str] = []
    for v in values:
        s = _s(v)
        if not s:
            continue
        for piece in re.split(r"\s*[|/·;；]\s*", s):
            piece = piece.strip()
            if not piece:
                continue
            key = piece.casefold()
            if key in seen:
                continue
            seen.add(key)
            parts.append(piece)
            if len(parts) >= max_parts:
                break
        if len(parts) >= max_parts:
            break
    return sep.join(parts)


def _best_text(values) -> str:
    """가장 길고 알찬 값 하나."""
    best = ""
    for v in values:
        s = _s(v)
        if len(s) > len(best):
            best = s
    return best


def _best_region(values) -> str:
    ranked = []
    for v in values:
        s = _s(v) or "미분류"
        ranked.append(s)
    for s in ranked:
        if s and s != "미분류":
            return s
    return ranked[0] if ranked else "미분류"


def merge_duplicate_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """같은 업체키 행을 1건으로 병합. (병합 DF, 제거된 중복 건수)."""
    if df.empty:
        return df, 0
    work = df.copy()
    work["업체키"] = work["업체명"].map(_company_key)
    work = work[work["업체키"].astype(str).str.len() >= 2]
    before = len(work)
    groups = []
    for key, g in work.groupby("업체키", sort=False):
        if len(g) == 1:
            row = g.iloc[0].to_dict()
            row["병합건수"] = 1
            groups.append(row)
            continue
        row = {
            "업체키": key,
            "업체명": _best_text(g["업체명"]),
            "지역": _best_region(g["지역"]),
            "주소": _best_text(g["주소"]),
            "업종": _merge_unique_text(g["업종"], max_parts=6),
            "사용가스": _merge_unique_text(g["사용가스"], max_parts=8),
            "공급사": _merge_unique_text(g["공급사"], max_parts=8),
            "담당자": _merge_unique_text(g["담당자"], max_parts=6),
            "연락처": _merge_unique_text(g["연락처"], max_parts=6),
            "비고": _merge_unique_text(g["비고"], max_parts=10),
            "출처": _merge_unique_text(g["출처"], max_parts=8),
            "파일": _merge_unique_text(g["파일"], max_parts=6),
            "시트": _merge_unique_text(g["시트"], max_parts=8),
            "병합건수": int(len(g)),
        }
        groups.append(row)
    out = pd.DataFrame(groups)
    removed = before - len(out)
    return out.reset_index(drop=True), removed


def _norm_header(v) -> str:
    t = _s(v)
    t = t.replace(" ", "").replace("\u3000", "")
    return t


def infer_region(*texts: str, sheet_hint: str = "") -> str:
    blob = " ".join([sheet_hint, *[t for t in texts if t]])
    if not blob.strip():
        return "미분류"
    for pat, name in _REGION_RULES:
        if pat.search(blob):
            return name
    return "미분류"


def _find_header_row(rows: list[tuple], keywords: list[str], scan: int = 12) -> int | None:
    keys = [k.replace(" ", "") for k in keywords]
    for i, row in enumerate(rows[:scan]):
        cells = [_norm_header(c) for c in row]
        hit = sum(1 for k in keys if any(k in c for c in cells if c))
        if hit >= max(2, min(3, len(keys) // 2 + 1)):
            return i
    return None


def _row_dict(header: list[str], row: tuple) -> dict[str, str]:
    out: dict[str, str] = {}
    for i, h in enumerate(header):
        if not h:
            continue
        val = _s(row[i]) if i < len(row) else ""
        if h in out and out[h] and val:
            out[h] = f"{out[h]} / {val}"
        elif h not in out or not out[h]:
            out[h] = val
    return out


def _pick(d: dict[str, str], *cands: str) -> str:
    for c in cands:
        c2 = c.replace(" ", "")
        for k, v in d.items():
            kn = k.replace(" ", "")
            if c2 == kn or c2 in kn or kn in c2:
                if v:
                    return v
    return ""


def ensure_market_research_cache() -> str:
    """Drive 원본이 있으면 캐시로 동기화. 없으면 배포 시드 사용."""
    os.makedirs(MR_CACHE_DIR, exist_ok=True)
    for cand in MR_DRIVE_CANDIDATES:
        if cand and os.path.isdir(cand):
            try:
                for root, _dirs, files in os.walk(cand):
                    rel = os.path.relpath(root, cand)
                    dest_root = (
                        MR_CACHE_DIR if rel == "." else os.path.join(MR_CACHE_DIR, rel)
                    )
                    os.makedirs(dest_root, exist_ok=True)
                    for name in files:
                        if name.startswith("~$") or name.startswith("."):
                            continue
                        if not name.lower().endswith((".xlsx", ".xls", ".csv")):
                            continue
                        src = os.path.join(root, name)
                        dst = os.path.join(dest_root, name)
                        try:
                            if (not os.path.exists(dst)) or (
                                os.path.getmtime(src) > os.path.getmtime(dst) + 1
                            ):
                                shutil.copy2(src, dst)
                        except Exception:
                            pass
                return MR_CACHE_DIR
            except Exception:
                continue
    return MR_CACHE_DIR


def _list_xlsx(root: str) -> list[Path]:
    out: list[Path] = []
    p = Path(root)
    if not p.is_dir():
        return out
    for f in sorted(p.rglob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        out.append(f)
    return out


def _read_sheet_rows(path: Path, sheet: str, max_rows: int = 20000) -> list[tuple]:
    if load_workbook is None:
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            return []
        ws = wb[sheet]
        rows: list[tuple] = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            rows.append(tuple(row))
            if i >= max_rows:
                break
        return rows
    finally:
        wb.close()


def _parse_lco2(path: Path) -> list[dict]:
    if load_workbook is None or not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    records: list[dict] = []
    try:
        for sn in wb.sheetnames:
            rows = []
            ws = wb[sn]
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                rows.append(tuple(row))
                if i >= 5000:
                    break
            hi = _find_header_row(
                rows, ["상호명", "소재지", "공급사", "위치", "제품명"]
            )
            if hi is None:
                continue
            header = [_norm_header(c) for c in rows[hi]]
            last_region = sn.strip() or "미분류"
            for row in rows[hi + 1 :]:
                d = _row_dict(header, row)
                name = _pick(d, "상호명", "상호", "업체명")
                if not name or len(name) < 2:
                    continue
                loc = _pick(d, "위치", "지역")
                addr = _pick(d, "소재지", "주소")
                if _pick(d, "지역") and len(_pick(d, "지역")) <= 6:
                    last_region = _pick(d, "지역")
                region = infer_region(loc, addr, name, sheet_hint=last_region)
                records.append(
                    {
                        "출처": "LCO2경쟁사",
                        "파일": path.name,
                        "시트": sn.strip(),
                        "지역": region,
                        "업체명": name,
                        "주소": addr or loc,
                        "업종": _pick(d, "고압가스종류", "업종"),
                        "사용가스": _pick(d, "제품명", "사용가스"),
                        "공급사": _pick(d, "공급사", "현공급처"),
                        "담당자": _pick(d, "담당", "담당자"),
                        "연락처": _pick(d, "연락처", "전화"),
                        "비고": _pick(d, "비고", "대납업체", "월사용량"),
                    }
                )
    finally:
        wb.close()
    return records


def _parse_region_survey(path: Path) -> list[dict]:
    """시장조사 (67).xlsx — 지역 시트별 표준 양식."""
    if load_workbook is None or not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    records: list[dict] = []
    try:
        for sn in wb.sheetnames:
            rows = []
            ws = wb[sn]
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                rows.append(tuple(row))
                if i >= 5000:
                    break
            hi = _find_header_row(
                rows, ["상호", "위치", "현공급처", "용기보유", "업종"]
            )
            if hi is None:
                # 헤더 없이 데이터만 있는 음성 시트 등
                hi = _find_header_row(rows, ["상호", "위치", "연락처"])
            header = (
                [_norm_header(c) for c in rows[hi]]
                if hi is not None
                else ["NO", "상호", "위치", "연락처", "담당", "업종", "용기보유현황", "현공급처", "비고"]
            )
            start = (hi + 1) if hi is not None else 4
            # 헤더가 약하면 고정 매핑
            if hi is None or not any("상호" in h for h in header):
                for row in rows[start:]:
                    vals = [_s(c) for c in row[:9]]
                    if len(vals) < 3 or not vals[1]:
                        continue
                    if vals[1] in {"상호", "상 호", "NO"}:
                        continue
                    addr = vals[2]
                    region = infer_region(addr, sn, sheet_hint=sn)
                    records.append(
                        {
                            "출처": "지역시장조사",
                            "파일": path.name,
                            "시트": sn,
                            "지역": region,
                            "업체명": vals[1],
                            "주소": addr,
                            "업종": vals[5] if len(vals) > 5 else "",
                            "사용가스": vals[6] if len(vals) > 6 else "",
                            "공급사": vals[7] if len(vals) > 7 else "",
                            "담당자": vals[4] if len(vals) > 4 else "",
                            "연락처": vals[3] if len(vals) > 3 else "",
                            "비고": vals[8] if len(vals) > 8 else "",
                        }
                    )
                continue
            for row in rows[start:]:
                d = _row_dict(header, row)
                name = _pick(d, "상호", "상호명", "업체명")
                if not name or name in {"상호", "상 호"}:
                    continue
                addr = _pick(d, "위치", "주소", "소재지", "지역")
                region = infer_region(addr, sn, sheet_hint=sn)
                records.append(
                    {
                        "출처": "지역시장조사",
                        "파일": path.name,
                        "시트": sn,
                        "지역": region,
                        "업체명": name,
                        "주소": addr,
                        "업종": _pick(d, "업종", "생산품목"),
                        "사용가스": _pick(d, "용기보유현황", "사용가스", "용기"),
                        "공급사": _pick(d, "현공급처", "공급처", "공급사"),
                        "담당자": _pick(d, "담당", "담당자"),
                        "연락처": _pick(d, "연락처", "전화", "전화번호"),
                        "비고": _pick(d, "비고", "비 고"),
                    }
                )
    finally:
        wb.close()
    return records


def _parse_visit_notes(path: Path, source_label: str) -> list[dict]:
    """김진혁/mail 등 — 시트명≈지역, 업체명·공급처·사용가스."""
    if load_workbook is None or not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    records: list[dict] = []
    try:
        for sn in wb.sheetnames:
            if _SKIP_SHEET.search(sn):
                continue
            if "통합" in sn and "모든내용" in sn:
                pass  # keep
            rows = []
            ws = wb[sn]
            # 넓은 시트(화성)는 앞 열만
            for i, row in enumerate(ws.iter_rows(values_only=True, max_col=14)):
                rows.append(tuple(row))
                if i >= 8000:
                    break
            hi = _find_header_row(
                rows, ["업체명", "지역", "공급처", "사용가스", "담당자"]
            )
            if hi is None:
                continue
            header = [_norm_header(c) for c in rows[hi]]
            for row in rows[hi + 1 :]:
                d = _row_dict(header, row)
                name = _pick(d, "업체명", "상호", "입체명")
                if not name or len(name) < 2:
                    continue
                if name.endswith("x") and len(name) <= 3:
                    continue
                addr = _pick(d, "지역", "주소", "위치", "회사위치")
                region = infer_region(addr, sn, sheet_hint=sn)
                records.append(
                    {
                        "출처": source_label,
                        "파일": path.name,
                        "시트": sn,
                        "지역": region,
                        "업체명": name.rstrip("x").strip() or name,
                        "주소": addr,
                        "업종": _pick(d, "생산품목", "업종", "종목"),
                        "사용가스": _pick(d, "사용가스", "가스"),
                        "공급사": _pick(d, "공급처", "현공급처", "공급사"),
                        "담당자": _pick(d, "담당자", "담당"),
                        "연락처": "",
                        "비고": _pick(d, "비고", "특이사항", "세부사항"),
                    }
                )
    finally:
        wb.close()
    return records


def _parse_factory_registry(path: Path) -> list[dict]:
    """화성공장등록검색 — 대규모 공장 DB (샘플/필터용)."""
    if not path.exists():
        return []
    try:
        df = pd.read_excel(path, sheet_name=0, header=0, dtype=str)
    except Exception:
        return []
    cols = {str(c).strip(): c for c in df.columns}
    def col(*names):
        for n in names:
            for k, c in cols.items():
                if n in k.replace(" ", ""):
                    return c
        return None

    c_name = col("회사명")
    c_addr = col("공장주소", "주소")
    c_prod = col("생산품")
    c_ind = col("업종명", "대표업종")
    c_tel = col("전화번호")
    c_park = col("산업단지")
    if c_name is None:
        return []
    records: list[dict] = []
    for _, r in df.iterrows():
        name = _s(r.get(c_name))
        if not name:
            continue
        addr = _s(r.get(c_addr)) if c_addr is not None else ""
        region = infer_region(addr, sheet_hint="화성")
        records.append(
            {
                "출처": "화성공장등록",
                "파일": path.name,
                "시트": "등록공장",
                "지역": region if region != "미분류" else "화성",
                "업체명": name,
                "주소": addr,
                "업종": _s(r.get(c_ind)) if c_ind is not None else "",
                "사용가스": "",
                "공급사": "",
                "담당자": "",
                "연락처": _s(r.get(c_tel)) if c_tel is not None else "",
                "비고": " / ".join(
                    x
                    for x in [
                        _s(r.get(c_prod)) if c_prod is not None else "",
                        _s(r.get(c_park)) if c_park is not None else "",
                    ]
                    if x
                ),
            }
        )
    return records


def _parse_seojin(path: Path) -> list[dict]:
    if load_workbook is None or not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    records: list[dict] = []
    try:
        for sn in wb.sheetnames:
            if _SKIP_SHEET.search(sn) and "화성" not in sn:
                continue
            rows = []
            ws = wb[sn]
            for i, row in enumerate(ws.iter_rows(values_only=True, max_col=12)):
                rows.append(tuple(row))
                if i >= 3000:
                    break
            hi = _find_header_row(rows, ["입체명", "주소", "담당자", "업체명"])
            if hi is None:
                hi = _find_header_row(rows, ["업체명", "지역", "공급처"])
            if hi is None:
                continue
            header = [_norm_header(c) for c in rows[hi]]
            for row in rows[hi + 1 :]:
                d = _row_dict(header, row)
                name = _pick(d, "입체명", "업체명", "상호")
                if not name:
                    continue
                addr = _pick(d, "주소", "지역", "위치")
                records.append(
                    {
                        "출처": "서진산업가스",
                        "파일": path.name,
                        "시트": sn,
                        "지역": infer_region(addr, sn, sheet_hint=sn),
                        "업체명": name,
                        "주소": addr,
                        "업종": _pick(d, "종목", "업종", "생산품목"),
                        "사용가스": "",
                        "공급사": "서진산업가스",
                        "담당자": _pick(d, "담당자"),
                        "연락처": _pick(d, "전화번호", "연락처"),
                        "비고": _pick(d, "특이사항", "비고"),
                    }
                )
    finally:
        wb.close()
    return records


@st.cache_data(show_spinner=False, ttl=600)
def load_market_research_frame(_cache_sig: str) -> tuple[pd.DataFrame, int, int]:
    """모든 시장조사 엑셀을 합치고 중복 업체를 병합.

    Returns
    -------
    (merged_df, raw_count, removed_dup_count)
    """
    root = ensure_market_research_cache()
    files = {p.name: p for p in _list_xlsx(root)}
    records: list[dict] = []

    for name, p in files.items():
        nfc = unicodedata.normalize("NFC", name)
        n = nfc.lower()
        if "lco2" in n or "경쟁사" in nfc:
            records.extend(_parse_lco2(p))
        elif "시장조사 (67)" in nfc or nfc.startswith("시장조사 (67)"):
            records.extend(_parse_region_survey(p))
        elif "mail" in n or "시장조사ㅡ" in nfc or "시장조사-" in nfc:
            records.extend(_parse_visit_notes(p, "방문조사(mail)"))
        elif "화성" in nfc and "공장" in nfc:
            records.extend(_parse_factory_registry(p))
        elif "김진혁" in nfc:
            records.extend(_parse_visit_notes(p, "방문조사(김진혁)"))
        elif "서진" in nfc:
            records.extend(_parse_seojin(p))
        elif "스페셜" in nfc or "스폐셜" in nfc:
            continue

    empty_cols = [
        "출처",
        "파일",
        "시트",
        "지역",
        "업체명",
        "주소",
        "업종",
        "사용가스",
        "공급사",
        "담당자",
        "연락처",
        "비고",
        "업체키",
        "병합건수",
    ]
    if not records:
        return pd.DataFrame(columns=empty_cols), 0, 0

    df = pd.DataFrame(records)
    df = df[df["업체명"].map(lambda x: len(_s(x)) >= 2)].copy()
    raw_n = len(df)
    merged, removed = merge_duplicate_rows(df)
    return merged, raw_n, removed


def _cache_signature() -> str:
    root = ensure_market_research_cache()
    parts = []
    for p in _list_xlsx(root):
        try:
            parts.append(f"{p}:{os.path.getmtime(p):.0f}:{os.path.getsize(p)}")
        except OSError:
            parts.append(str(p))
    return "|".join(parts) or "empty"


def _metric_box(label: str, value: str) -> str:
    return (
        f"<div class='metric-box'><div class='metric-label'>{label}</div>"
        f"<div class='metric-value'>{value}</div></div>"
    )


def render_market_research_tab(latest_update_str: str = "") -> None:
    """시장조사 탭 UI."""
    st.markdown(
        "<div class='sub-header dashboard-tab-panel-head'>🔎 시장조사</div>",
        unsafe_allow_html=True,
    )
    st.caption(
        "경로: Google Drive › 다른 컴퓨터 › Desktop › 업무 › 시장조사  "
        "· 지역·공급사·출처별로 정리한 통합 조회"
    )

    sig = _cache_signature()
    with st.spinner("시장조사 자료 정리·중복 병합 중…"):
        df, raw_n, removed_n = load_market_research_frame(sig)

    if df.empty:
        st.warning(
            f"`{MR_CACHE_DIR}` 에 엑셀이 없습니다. "
            "맥에서 Drive「업무/시장조사」동기화 후 새로고침하세요."
        )
        return

    # —— 상단 KPI
    n_all = len(df)
    n_merged_rows = int((df["병합건수"] > 1).sum()) if "병합건수" in df.columns else 0
    n_region = df.loc[df["지역"] != "미분류", "지역"].nunique()
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(_metric_box("병합 후 업체", f"{n_all:,}"), unsafe_allow_html=True)
    c2.markdown(
        _metric_box("중복 병합", f"{removed_n:,}건↓"),
        unsafe_allow_html=True,
    )
    c3.markdown(_metric_box("지역 수", f"{n_region:,}"), unsafe_allow_html=True)
    c4.markdown(
        _metric_box("원본→병합", f"{raw_n:,}→{n_all:,}"),
        unsafe_allow_html=True,
    )
    if removed_n:
        st.caption(
            f"같은 업체명(㈜/주식회사·공백 무시)은 1건으로 합쳤습니다. "
            f"병합된 업체 {n_merged_rows:,}곳 · 공급사·비고·출처는 · 로 합쳐 표시합니다."
        )

    # —— 필터
    regions = sorted(
        [r for r in df["지역"].dropna().unique().tolist() if r],
        key=lambda x: (x == "미분류", x),
    )
    _src_set: set[str] = set()
    for v in df["출처"].dropna().astype(str):
        for p in re.split(r"\s*[·|/]\s*", v):
            p = p.strip()
            if p:
                _src_set.add(p)
    sources = sorted(_src_set)
    suppliers = sorted(
        [s for s in df["공급사"].dropna().unique().tolist() if str(s).strip()],
        key=lambda x: str(x),
    )[:400]

    f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.4, 1.6])
    with f1:
        sel_region = st.multiselect(
            "지역",
            options=regions,
            default=[],
            key="mr_region",
            placeholder="전체 지역",
        )
    with f2:
        sel_src = st.multiselect(
            "출처",
            options=sources,
            default=[],
            key="mr_src",
            placeholder="전체 출처",
        )
    with f3:
        sel_sup = st.multiselect(
            "공급사",
            options=suppliers,
            default=[],
            key="mr_sup",
            placeholder="전체 공급사",
        )
    with f4:
        q = st.text_input(
            "검색 (업체·주소·가스·비고)",
            key="mr_q",
            placeholder="예: 화성, LCO2, 린데…",
        )

    view = df
    if sel_region:
        view = view[view["지역"].isin(sel_region)]
    if sel_src:
        src_mask = False
        for s in sel_src:
            src_mask = src_mask | view["출처"].fillna("").astype(str).str.contains(
                re.escape(s), case=False, regex=True, na=False
            )
        view = view[src_mask]
    if sel_sup:
        sup_mask = False
        for s in sel_sup:
            sup_mask = sup_mask | view["공급사"].fillna("").astype(str).str.contains(
                re.escape(s), case=False, regex=True, na=False
            )
        view = view[sup_mask]
    if q.strip():
        qq = q.strip()
        mask = False
        for col in ("업체명", "주소", "업종", "사용가스", "공급사", "비고", "담당자"):
            mask = mask | view[col].fillna("").astype(str).str.contains(
                qq, case=False, regex=False, na=False
            )
        view = view[mask]

    tab_overview, tab_region, tab_supplier, tab_factory, tab_files = st.tabs(
        ["📋 통합 목록", "🗺️ 지역별", "🏭 공급사별", "🏗️ 화성공장 DB", "📁 원본 파일"]
    )

    show_cols = [
        "지역",
        "업체명",
        "주소",
        "업종",
        "사용가스",
        "공급사",
        "담당자",
        "연락처",
        "비고",
        "출처",
        "시트",
        "병합건수",
    ]

    with tab_overview:
        st.caption(f"필터 결과 **{len(view):,}**건 / 고유 업체 **{view['업체키'].nunique():,}**")
        st.dataframe(
            view[show_cols],
            width="stretch",
            hide_index=True,
            height=520,
        )
        csv = view[show_cols].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "필터 결과 CSV 다운로드",
            data=csv,
            file_name="시장조사_필터결과.csv",
            mime="text/csv",
            key="mr_dl_csv",
        )

    with tab_region:
        left, right = st.columns([1, 1.4])
        region_counts = (
            view.groupby("지역", dropna=False)
            .agg(레코드=("업체명", "count"), 고유업체=("업체키", "nunique"))
            .sort_values("레코드", ascending=False)
            .reset_index()
        )
        with left:
            st.markdown("**지역별 건수**")
            st.dataframe(region_counts, width="stretch", hide_index=True, height=420)
        with right:
            pick = st.selectbox(
                "지역 상세",
                options=region_counts["지역"].tolist() or ["미분류"],
                key="mr_region_pick",
            )
            sub = view[view["지역"] == pick][show_cols]
            st.caption(f"{pick} · {len(sub):,}건")
            st.dataframe(sub, width="stretch", hide_index=True, height=420)

    with tab_supplier:
        # 공급사 비어있으면 제외하고 상위
        sup = view[view["공급사"].fillna("").astype(str).str.strip() != ""].copy()
        if sup.empty:
            st.info("공급사 정보가 있는 행이 없습니다. 다른 출처를 선택해 보세요.")
        else:
            sc = (
                sup.groupby("공급사")
                .agg(레코드=("업체명", "count"), 고유업체=("업체키", "nunique"))
                .sort_values("레코드", ascending=False)
                .head(50)
                .reset_index()
            )
            a, b = st.columns([1, 1.4])
            with a:
                st.markdown("**공급사 TOP**")
                st.dataframe(sc, width="stretch", hide_index=True, height=420)
            with b:
                pick_s = st.selectbox(
                    "공급사 상세",
                    options=sc["공급사"].tolist(),
                    key="mr_sup_pick",
                )
                sub = view[view["공급사"] == pick_s][show_cols]
                st.caption(f"{pick_s} · {len(sub):,}건")
                st.dataframe(sub, width="stretch", hide_index=True, height=420)

    with tab_factory:
        fac = view[view["출처"].fillna("").astype(str).str.contains("화성공장등록", regex=False, na=False)]
        st.caption(
            "화성공장 등록 검색 DB — 가스 공급 여부와 무관한 **등록 공장 전체**. "
            f"현재 필터 기준 {len(fac):,}건"
        )
        q2 = st.text_input("공장 DB 추가 검색", key="mr_fac_q", placeholder="업종·주소·회사명")
        fac2 = fac
        if q2.strip():
            qq = q2.strip()
            m = False
            for col in ("업체명", "주소", "업종", "비고"):
                m = m | fac2[col].fillna("").astype(str).str.contains(
                    qq, case=False, regex=False, na=False
                )
            fac2 = fac2[m]
        st.dataframe(
            fac2[show_cols],
            width="stretch",
            hide_index=True,
            height=480,
        )

    with tab_files:
        root = ensure_market_research_cache()
        files = _list_xlsx(root)
        st.markdown(f"**캐시 폴더:** `{root}`")
        rows = []
        for p in files:
            try:
                sz = os.path.getsize(p)
            except OSError:
                sz = 0
            rows.append(
                {
                    "파일": str(p.relative_to(root)) if root in str(p) else p.name,
                    "크기(KB)": round(sz / 1024, 1),
                }
            )
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        by_src = (
            df.groupby("출처")
            .agg(레코드=("업체명", "count"), 고유업체=("업체키", "nunique"))
            .sort_values("레코드", ascending=False)
            .reset_index()
        )
        st.markdown("**출처별 적재량**")
        st.dataframe(by_src, width="stretch", hide_index=True)
        if latest_update_str:
            st.caption(f"대시보드 기준 시각: {latest_update_str}")
