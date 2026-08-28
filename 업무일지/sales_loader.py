"""공문 탭용 매출 DataFrame 로더 — uploaded_cache/sales/*.csv.

메인 app.py 의 파서와 동일한 규칙(종속처 분리·담당자 추론)을 공문 탭에 맞게 축약.
"""
from __future__ import annotations

import io
import os
import re
from typing import Iterable

import numpy as np
import pandas as pd
import streamlit as st

CACHE_DIR = "uploaded_cache"
SALES_DIR = os.path.join(CACHE_DIR, "sales")

EMPTY_CLIENT_LABEL = "(거래처명 없음)"

_REMARK_SUBCLIENT_SKIP_RE = re.compile(
    r"이관|변경|출고|회수|반납|정리|차액|으로|부터|까지|취소|오류|수정|"
    r"미검|공병|대납|직접|고객사|용기|고압|일반으로|합계|이월|"
    r"거래분|전자어음|입금|반품|상계|가수|물품대|임대료|실린더|벌크|본사|"
    r"여수|운반|운임|수수료|보증|회수"
)
_REMARK_QTY_RE = re.compile(r"\d[\d,]*\s*(?:kg|KG|톤|t|T|L|리터)\b")
_REMARK_GAS_ONLY_RE = re.compile(
    r"^(?:O2|N2|AR|CO2|AIR|He|HE|LO2|LN2|LPG)(?:\s*[,/·+]?\s*(?:O2|N2|AR|CO2|AIR|He|HE|LO2|LN2|LPG|\d[\d,]*\s*(?:kg|L)?))*\s*$",
    re.I,
)


def _is_valid_client_name(name) -> bool:
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return False
    s = str(name).strip()
    if not s:
        return False
    if s == EMPTY_CLIENT_LABEL:
        return True
    return s.lower() not in {"nan", "none", "nat", "null"}


def _normalize_manual_client_key(name) -> str | None:
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return EMPTY_CLIENT_LABEL
    s = str(name).strip()
    if not s or s.lower() in {"nan", "none", "nat"}:
        return EMPTY_CLIENT_LABEL
    return s


def _parse_sales_filename_year_month(file_name: str):
    base = os.path.basename(str(file_name or ""))
    m = re.match(r"^(20\d{2})(\d{2})?(?:\.csv)?$", base, flags=re.IGNORECASE)
    if m:
        return m.group(1), m.group(2)
    m2 = re.search(r"(20\d{2})(\d{2})?", base)
    if m2:
        return m2.group(1), m2.group(2)
    return None, None


def _dedupe_sales_file_meta(file_meta: list) -> list:
    if not file_meta:
        return file_meta
    rows = []
    for item in file_meta:
        try:
            name, path = item[0], item[1]
        except Exception:
            continue
        y, mm = _parse_sales_filename_year_month(name)
        try:
            sz = int(item[3]) if len(item) > 3 else int(os.path.getsize(path))
        except Exception:
            sz = 0
        try:
            mt = int(item[2]) if len(item) > 2 else 0
        except Exception:
            mt = 0
        rows.append({"name": name, "item": item, "y": y, "mm": mm, "sz": sz, "mt": mt})

    drop: set[str] = set()
    winners = {}
    for r in rows:
        if not r["y"] or r["sz"] <= 0:
            continue
        key = (r["y"], r["sz"])
        prev = winners.get(key)
        if prev is None:
            winners[key] = r
            continue
        score = (1 if r["mm"] else 0, r["mt"], len(r["name"]))
        pscore = (1 if prev["mm"] else 0, prev["mt"], len(prev["name"]))
        if score >= pscore:
            winners[key] = r
    for r in rows:
        if not r["y"] or r["sz"] <= 0:
            continue
        w = winners.get((r["y"], r["sz"]))
        if w is not None and r["name"] != w["name"]:
            drop.add(r["name"])

    by_y: dict = {}
    for r in rows:
        if r["name"] in drop or not r["y"]:
            continue
        by_y.setdefault(r["y"], []).append(r)
    for grp in by_y.values():
        annuals = [r for r in grp if not r["mm"]]
        monthlies = [r for r in grp if r["mm"]]
        if annuals and monthlies:
            for a in annuals:
                drop.add(a["name"])

    return [r["item"] for r in rows if r["name"] not in drop]


def parse_date_series_robust(series, default_year="2026"):
    if series.empty:
        return pd.Series(pd.NaT, index=series.index)
    s_str = series.astype(str).str.strip()
    parsed = pd.Series(pd.NaT, index=series.index)
    digits = s_str.str.replace(r"\D", "", regex=True)
    cond_8 = (digits.str.len() == 8) & parsed.isna()
    if cond_8.any():
        parsed[cond_8] = pd.to_datetime(digits[cond_8], format="%Y%m%d", errors="coerce")
    cond_6 = (digits.str.len() == 6) & parsed.isna()
    if cond_6.any():
        parsed[cond_6] = pd.to_datetime("20" + digits[cond_6], format="%Y%m%d", errors="coerce")
    remaining_mask = parsed.isna() & (s_str != "") & (s_str != "nan") & (s_str != "None")
    if remaining_mask.any():
        rem_series = s_str[remaining_mask]
        parts_df = rem_series.str.split(r"[-/.\s]+", expand=True)
        parts_df = parts_df.apply(lambda col: col.str.strip()).replace("", None)
        num_parts = parts_df.notna().sum(axis=1)
        cond_3 = num_parts >= 3
        if cond_3.any():
            sub = parts_df[cond_3]
            y = sub[0].astype(str).str.strip()
            y = np.where(y.str.len() == 2, "20" + y, y)
            m = sub[1].astype(str).str.strip().str.zfill(2)
            d = sub[2].astype(str).str.strip().str.zfill(2)
            parsed.loc[sub.index] = pd.to_datetime(y + "-" + m + "-" + d, format="%Y-%m-%d", errors="coerce")
        cond_2 = num_parts == 2
        if cond_2.any():
            sub = parts_df[cond_2]
            m = sub[0].astype(str).str.strip().str.zfill(2)
            d = sub[1].astype(str).str.strip().str.zfill(2)
            parsed.loc[sub.index] = pd.to_datetime(
                str(default_year) + "-" + m + "-" + d, format="%Y-%m-%d", errors="coerce"
            )
    valid_range = (parsed >= pd.Timestamp("2000-01-01")) & (parsed <= pd.Timestamp("2099-12-31"))
    parsed[~valid_range] = pd.NaT
    return parsed


def _drop_sales_noise_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    noise = pd.Series(False, index=df.index)
    if "거래처" in df.columns:
        noise |= df["거래처"].astype(str).str.match(r"^\d{4}-\d{2}-\d{2}\s", na=False)
    if "품목명" in df.columns:
        noise |= df["품목명"].astype(str).str.contains(r"이월\s*미수|\[이월", na=False)
    return df.loc[~noise].copy()


def _is_remark_subclient_note(note) -> bool:
    s = str(note or "").strip()
    if not s or s.lower() in ("nan", "none", "-", "없음"):
        return False
    if _REMARK_SUBCLIENT_SKIP_RE.search(s):
        return False
    if _REMARK_QTY_RE.search(s):
        return False
    if _REMARK_GAS_ONLY_RE.match(s):
        return False
    if len(s) < 2:
        return False
    if re.fullmatch(r"[\d\s./\-~,]+", s):
        return False
    if re.match(r"^\d{1,2}/\d{1,2}\b", s):
        return False
    return True


def _remark_parent_base_name(client_name):
    s = str(client_name or "").strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    if re.search(r"\([^)]+\)\s*$", s):
        return None
    base = s.rstrip(".").strip()
    return base or None


def expand_remark_subclients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or "거래처" not in df.columns:
        return df if df is not None else pd.DataFrame()
    out = df.copy()
    out["거래처"] = out["거래처"].fillna("").astype(str).str.strip()
    if "거래처_원본" not in out.columns:
        out["거래처_원본"] = out["거래처"]
    else:
        blank_orig = out["거래처_원본"].isna() | (out["거래처_원본"].astype(str).str.strip() == "")
        out.loc[blank_orig, "거래처_원본"] = out.loc[blank_orig, "거래처"]
    if "비고" not in out.columns:
        return out
    client = out["거래처"]
    note = out["비고"].fillna("").astype(str).str.strip()
    base = client.map(_remark_parent_base_name)
    sub_mask = base.notna() & note.map(_is_remark_subclient_note)
    if sub_mask.any():
        out.loc[sub_mask, "거래처"] = base.loc[sub_mask] + "(" + note.loc[sub_mask] + ")"
    return out


def _apply_manual_staff_mapping(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or "거래처" not in df.columns:
        return df if df is not None else pd.DataFrame()
    out = df.copy()
    blank = ~out["거래처"].fillna("").astype(str).str.strip().map(_is_valid_client_name)
    if blank.any():
        out.loc[blank, "거래처"] = EMPTY_CLIENT_LABEL
    manual_map_path = os.path.join(CACHE_DIR, "manual_staff_mapping.csv")
    if not os.path.exists(manual_map_path):
        return out
    if "담당자" not in out.columns:
        out["담당자"] = "미지정"
    try:
        manual_df = pd.read_csv(manual_map_path)
        if "거래처" not in manual_df.columns or "담당자" not in manual_df.columns:
            return out
        manual_dict = {}
        for _, mrow in manual_df.iterrows():
            ck = _normalize_manual_client_key(mrow["거래처"])
            if ck is None:
                continue
            staff = str(mrow["담당자"]).strip() if pd.notna(mrow["담당자"]) else ""
            if staff and staff not in ("nan", "None", "NaN", "미지정"):
                manual_dict[ck] = staff
        if manual_dict:
            mask_manual = out["거래처"].isin(manual_dict.keys())
            out.loc[mask_manual, "담당자"] = out.loc[mask_manual, "거래처"].map(manual_dict)
    except Exception:
        pass
    return out


def _parse_sales_uploaded_tuples(file_tuples: Iterable[tuple]) -> pd.DataFrame:
    df_list = []
    for file_name, content in file_tuples:
        try:
            decoded_text = None
            for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
                try:
                    decoded_text = content.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if decoded_text is None:
                decoded_text = content.decode("utf-8", errors="replace")
            lines = [line for line in decoded_text.splitlines() if line.strip()]
            if not lines:
                continue
            header_idx = 0
            max_matches = 0
            for i, line in enumerate(lines[:30]):
                cells = re.split(r",|\t", line)
                matches = sum(
                    1
                    for cell in cells
                    if any(kw in cell for kw in ["거래처", "품목", "매출", "단가", "수량", "담당", "일자"])
                )
                if matches > max_matches:
                    max_matches = matches
                    header_idx = i
                if max_matches >= 3:
                    break
            df = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])), on_bad_lines="skip", engine="python")
            df.columns = df.columns.astype(str).str.strip()
            cols = list(df.columns)

            def find_col(priority_keywords, exclude_keywords=None):
                exclude_keywords = exclude_keywords or []
                clean_cols = [c.replace(" ", "") for c in cols]
                for kw in priority_keywords:
                    kw_clean = kw.replace(" ", "")
                    for orig_c, clean_c in zip(cols, clean_cols):
                        if any(ex in clean_c for ex in exclude_keywords):
                            continue
                        if kw_clean == clean_c:
                            return orig_c
                for kw in priority_keywords:
                    kw_clean = kw.replace(" ", "")
                    for orig_c, clean_c in zip(cols, clean_cols):
                        if any(ex in clean_c for ex in exclude_keywords):
                            continue
                        if kw_clean in clean_c:
                            return orig_c
                return None

            rename_dict = {}
            c_staff = find_col(["담당자", "담당자명", "영업담당", "영업사원", "담당"], ["코드", "ID", "번호"])
            c_client = find_col(["거래처", "거래처명", "상호명", "고객명", "회사명", "상호", "고객"], ["코드", "ID", "번호", "담당", "영업"])
            c_item = find_col(["품목명", "제품명", "상품명", "품목", "제품"], ["코드", "ID", "번호", "규격"])
            c_sales = find_col(["매출액", "금액", "매출"], ["일", "자", "수량", "량", "단가"])
            c_qty = find_col(["출고량", "수량", "출고"], ["액", "금액", "단가"])
            c_price = find_col(["단가", "단 가", "판매단가", "공급단가"], ["액", "금액", "수량", "량"])
            c_date = find_col(["매출일자", "매출일", "일자", "날짜", "출고일"])
            c_note = find_col(["비고", "적요", "메모", "특이사항"], ["코드", "ID"])
            if c_client:
                rename_dict[c_client] = "거래처"
            if c_item:
                rename_dict[c_item] = "품목명"
            if c_staff:
                rename_dict[c_staff] = "담당자"
            if c_sales:
                rename_dict[c_sales] = "매출액"
            if c_qty:
                rename_dict[c_qty] = "출고량"
            if c_price:
                rename_dict[c_price] = "단가"
            if c_date:
                rename_dict[c_date] = "매출일자_raw"
            if c_note:
                rename_dict[c_note] = "비고"
            df = df.rename(columns=rename_dict)
            df = _drop_sales_noise_rows(df)
            for req in ["거래처", "품목명", "담당자"]:
                if req not in df.columns:
                    df[req] = "미지정"
            file_year = _parse_sales_filename_year_month(file_name)[0] or "2026"
            date_col = "매출일자_raw" if "매출일자_raw" in df.columns else df.columns[0]
            df["매출일_dt"] = parse_date_series_robust(df[date_col], default_year=file_year)
            if "매출액" in df.columns:
                df["매출액"] = pd.to_numeric(
                    df["매출액"].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce"
                ).fillna(0)
            else:
                df["매출액"] = 0
            if "출고량" in df.columns:
                df["출고량"] = pd.to_numeric(
                    df["출고량"].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce"
                ).fillna(0)
            else:
                df["출고량"] = 0
            if "단가" in df.columns:
                df["단가"] = pd.to_numeric(
                    df["단가"].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce"
                ).fillna(0)
            else:
                df["단가"] = 0
            df["거래처"] = df["거래처"].fillna("").astype(str).str.strip()
            df["담당자"] = df["담당자"].fillna("미지정").astype(str).str.strip()
            blank_client = ~df["거래처"].map(_is_valid_client_name)
            if blank_client.any():
                df.loc[blank_client, "거래처"] = EMPTY_CLIENT_LABEL
            df = expand_remark_subclients(df)
            df = df.dropna(subset=["매출일_dt"])
            if not df.empty:
                df_list.append(df)
        except Exception:
            continue
    result_df = pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()
    if not result_df.empty and "거래처" in result_df.columns and "담당자" in result_df.columns:
        invalid_staff_markers = ["", "nan", "None", "NAT", "NaN", "담당자없음", "지정안함", "없음"]
        result_df["담당자"] = result_df["담당자"].fillna("미지정").astype(str).str.strip()
        result_df["담당자"] = result_df["담당자"].replace(invalid_staff_markers, "미지정")
        temp_df = result_df.dropna(subset=["매출일_dt"]).sort_values("매출일_dt")
        valid_staff_map = temp_df[~temp_df["담당자"].isin(["미지정"])].groupby("거래처")["담당자"].last().to_dict()
        mask_unassigned = result_df["담당자"] == "미지정"
        result_df.loc[mask_unassigned, "담당자"] = (
            result_df.loc[mask_unassigned, "거래처"].map(valid_staff_map).fillna("미지정")
        )
        result_df.loc[result_df["담당자"].isin(invalid_staff_markers), "담당자"] = "미지정"
    result_df = _apply_manual_staff_mapping(result_df)
    if not result_df.empty and "거래처" in result_df.columns and "담당자" in result_df.columns:
        mask_closed = result_df["거래처"].astype(str).str.strip().str.match(r"^[zZ]", na=False)
        if mask_closed.any():
            result_df.loc[mask_closed, "담당자"] = "거래종료"
    is_deposit = result_df["품목명"].astype(str).str.contains("입금", na=False) if "품목명" in result_df.columns else False
    if isinstance(is_deposit, pd.Series) and is_deposit.any():
        result_df = result_df[~is_deposit].copy()
    return result_df


def _collect_sales_file_meta() -> list:
    if not os.path.isdir(SALES_DIR):
        return []
    meta = []
    for fn in sorted(os.listdir(SALES_DIR)):
        if not fn.lower().endswith(".csv"):
            continue
        path = os.path.join(SALES_DIR, fn)
        if not os.path.isfile(path):
            continue
        try:
            st_info = os.stat(path)
            meta.append((fn, path, int(st_info.st_mtime_ns), int(st_info.st_size)))
        except OSError:
            meta.append((fn, path, 0, 0))
    return _dedupe_sales_file_meta(meta)


def _manual_staff_map_cache_token() -> tuple:
    path = os.path.join(CACHE_DIR, "manual_staff_mapping.csv")
    try:
        st_info = os.stat(path)
        return (int(st_info.st_mtime_ns), int(st_info.st_size))
    except OSError:
        return (0, 0)


@st.cache_data(show_spinner="매출 데이터 불러오는 중…")
def _load_sales_cached(_meta_sig: tuple, _manual_token: tuple) -> pd.DataFrame:
    meta = _meta_sig[1] if isinstance(_meta_sig, tuple) and len(_meta_sig) == 2 else []
    tuples = []
    for item in meta:
        try:
            f_name, f_path = item[0], item[1]
            with open(f_path, "rb") as sf:
                tuples.append((f_name, sf.read()))
        except Exception:
            continue
    return _parse_sales_uploaded_tuples(tuples)


def load_sales_for_letter_tab() -> tuple[pd.DataFrame, str]:
    """(sales_df, latest_update_str) — 공문 탭용."""
    meta = _collect_sales_file_meta()
    sig = (tuple((m[0], m[2], m[3]) for m in meta), meta)
    manual_token = _manual_staff_map_cache_token()
    df = _load_sales_cached(sig, manual_token)
    latest = "데이터 없음"
    if not df.empty and "매출일_dt" in df.columns:
        latest_dt = df["매출일_dt"].max()
        if pd.notnull(latest_dt):
            latest = latest_dt.strftime("%Y-%m-%d")
    return df, latest
