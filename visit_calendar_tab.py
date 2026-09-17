"""방문·할일 캘린더 — 가로 일자 줄, 왼쪽 지정 납품 목록, 오른쪽 할일."""
from __future__ import annotations

import calendar
import html
import json
import os
import re
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from typing import Any

import pandas as pd
import streamlit as st

VC_DIR = os.path.join("uploaded_cache", "visit_calendar")
VC_STORE = os.path.join(VC_DIR, "store.json")
_WEEKDAYS = ("월", "화", "수", "목", "금", "토", "일")
_CYLINDER = "실린더"
_VC_SAT_BG, _VC_SAT_FG = "#f4f8ff", "#3b6fd8"
_VC_SUN_BG, _VC_SUN_FG = "#fff6f5", "#d23b3b"


def _s(v) -> str:
    if v is None:
        return ""
    t = str(v).strip()
    if t.lower() in {"nan", "none", "null"}:
        return ""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", t))


def _company_key(name: str) -> str:
    t = _s(name)
    t = re.sub(r"(주식회사|유한회사|유한책임회사|\(주\)|㈜|㈔)", "", t, flags=re.I)
    t = re.sub(r"[\s\(\)（）\[\]【】·\.\-_/,，、]+", "", t)
    return t.casefold()


def _iso(d: date | str | None) -> str:
    if d is None or d == "":
        return ""
    try:
        if pd.isna(d):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(d, datetime):
        return d.date().isoformat()
    if isinstance(d, date):
        return d.isoformat()
    s = _s(d)
    m = re.match(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if not m:
        return ""
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
    except ValueError:
        return ""


def _is_bulk_item(name: str) -> bool:
    s = str(name or "")
    return ("BULK" in s.upper()) or ("벌크" in s)


def _bulk_short(name: str) -> str:
    raw = _s(name)
    u = raw.upper()
    if "CO2" in u or "이산화" in raw:
        return "CO2"
    if "N2" in u or "질소" in raw:
        return "N2"
    if "O2" in u or "산소" in raw:
        return "O2"
    if "AR" in u or "아르곤" in raw:
        return "AR"
    return raw.split()[0][:8] or "벌크"


def _fmt_qty(v) -> str:
    try:
        n = float(v)
    except (TypeError, ValueError):
        return ""
    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n)):,}"
    return f"{n:,.1f}"


def _empty_store() -> dict:
    return {"todos": [], "visits": []}


def load_store() -> dict:
    if not os.path.exists(VC_STORE):
        return _empty_store()
    try:
        with open(VC_STORE, encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return _empty_store()
        todos = [x for x in (data.get("todos") or []) if isinstance(x, dict)]
        visits = [x for x in (data.get("visits") or []) if isinstance(x, dict)]
        return {"todos": todos, "visits": visits}
    except Exception:
        return _empty_store()


def save_store(store: dict) -> None:
    os.makedirs(VC_DIR, exist_ok=True)
    payload = {
        "todos": [x for x in (store.get("todos") or []) if isinstance(x, dict)],
        "visits": [x for x in (store.get("visits") or []) if isinstance(x, dict)],
    }
    with open(VC_STORE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def add_todo(fields: dict) -> dict:
    title = _s(fields.get("title") or fields.get("업체명") or fields.get("제목") or fields.get("client"))
    if len(title) < 1:
        raise ValueError("업체명을 입력하세요.")
    item = {
        "id": uuid.uuid4().hex[:12],
        "title": title,
        "done": False,
        "starred": bool(fields.get("starred")),
        "due": _iso(fields.get("due")) or date.today().isoformat(),
        "staff": _s(fields.get("staff")),
        "client": _s(fields.get("client") or title),
        "note": _s(fields.get("note") or fields.get("세부사항")),
        "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    store = load_store()
    store["todos"].insert(0, item)
    save_store(store)
    return item


def toggle_todo(todo_id: str, *, done: bool | None = None) -> bool:
    store = load_store()
    for t in store["todos"]:
        if str(t.get("id")) == str(todo_id):
            t["done"] = (not bool(t.get("done"))) if done is None else bool(done)
            save_store(store)
            return True
    return False


def delete_todo(todo_id: str) -> bool:
    store = load_store()
    n0 = len(store["todos"])
    store["todos"] = [t for t in store["todos"] if str(t.get("id")) != str(todo_id)]
    if len(store["todos"]) == n0:
        return False
    save_store(store)
    return True


def toggle_todo_star(todo_id: str, *, starred: bool | None = None) -> bool:
    store = load_store()
    for t in store["todos"]:
        if str(t.get("id")) == str(todo_id):
            t["starred"] = (not bool(t.get("starred"))) if starred is None else bool(starred)
            save_store(store)
            return True
    return False


def _todo_due_date(raw) -> date | None:
    iso = _iso(raw)
    if not iso:
        return None
    try:
        y, m, d = iso.split("-")
        return date(int(y), int(m), int(d))
    except ValueError:
        return None


def update_todo(todo_id: str, fields: dict) -> dict:
    title = _s(fields.get("title") or fields.get("업체명") or fields.get("제목") or fields.get("client"))
    if len(title) < 1:
        raise ValueError("업체명을 입력하세요.")
    store = load_store()
    for t in store["todos"]:
        if str(t.get("id")) != str(todo_id):
            continue
        t["title"] = title
        t["client"] = _s(fields.get("client") or title)
        t["note"] = _s(fields.get("note") if "note" in fields else fields.get("세부사항", t.get("note")))
        if "due" in fields:
            t["due"] = _iso(fields.get("due")) or t.get("due") or date.today().isoformat()
        save_store(store)
        return t
    raise ValueError("할일을 찾을 수 없습니다.")


def month_visit_rows(store: dict, month: date, staff: str = "") -> list[dict]:
    """이번 달 기방문·방문예정. 담당자가 있으면 그 사람 일정만."""
    prefix = f"{month.year:04d}-{month.month:02d}-"
    staff_s = _s(staff)
    rows: list[dict] = []
    for v in store.get("visits") or []:
        if not isinstance(v, dict):
            continue
        iso = _iso(v.get("date"))
        if not iso.startswith(prefix):
            continue
        vs = _s(v.get("staff"))
        if staff_s and vs and vs != staff_s:
            continue
        stt = _s(v.get("status")) or "done"
        if stt not in {"done", "planned"}:
            stt = "done"
        rows.append(
            {
                "id": str(v.get("id") or ""),
                "iso": iso,
                "status": stt,
                "client": _s(v.get("client")),
                "note": _s(v.get("note")),
                "staff": vs,
            }
        )
    rows.sort(key=lambda r: (r["iso"], 0 if r["status"] == "planned" else 1, r["client"]))
    return rows


def _direct_visit_on(store: dict, day: date | str, client: str) -> dict | None:
    iso = _iso(day)
    key = _company_key(client)
    if not iso or not key:
        return None
    for v in store.get("visits") or []:
        if _iso(v.get("date")) == iso and _company_key(v.get("client") or "") == key:
            return v
    return None


def add_visit(fields: dict) -> dict:
    client = _s(fields.get("client"))
    if len(client) < 2:
        raise ValueError("거래처를 2글자 이상 선택하거나 입력하세요.")
    status = _s(fields.get("status")) or "done"
    if status not in {"done", "planned"}:
        status = "done"
    item = {
        "id": uuid.uuid4().hex[:12],
        "date": _iso(fields.get("date")) or date.today().isoformat(),
        "staff": _s(fields.get("staff")) or "김혁수",
        "client": client,
        "note": _s(fields.get("note")),
        "source": "직접입력",
        "status": status,
    }
    store = load_store()
    store["visits"].insert(0, item)
    save_store(store)
    return item


def delete_visit(visit_id: str) -> bool:
    store = load_store()
    n0 = len(store["visits"])
    store["visits"] = [v for v in store["visits"] if str(v.get("id")) != str(visit_id)]
    if len(store["visits"]) == n0:
        return False
    save_store(store)
    return True


def clear_day_visit(day: date | str, client: str) -> bool:
    """해당 날·거래처의 방문·방문예정을 지운다. 달력·월간 일정·내역에서 사라진다."""
    hit = _direct_visit_on(load_store(), day, client)
    if not hit:
        return False
    return delete_visit(str(hit.get("id") or ""))


_SKIP_STAFF = {"", "미지정", "전체", "전체 담당자"}
_SKIP_CLIENT = {"", "미지정", "전체", "전체 거래처"}


def _staff_from_store(store: dict | None = None) -> list[str]:
    """기본 담당자 + 방문·할일 저장분. 상단에서 이미 만든 담당자 목록이 있으면 합친다."""
    names = ["김혁수"]
    src = store if isinstance(store, dict) else load_store()
    for row in list(src.get("visits") or []) + list(src.get("todos") or []):
        if not isinstance(row, dict):
            continue
        s = _s(row.get("staff"))
        if s and s not in names and s not in _SKIP_STAFF:
            names.append(s)
    try:
        for s in st.session_state.get("_dash_staff_opts_list") or []:
            s = _s(s)
            if s and s not in names and s not in _SKIP_STAFF:
                names.append(s)
        filt = st.session_state.get("dash_filter_staff_sb_v33") or st.session_state.get(
            "dash_filter_staff_sb_v32"
        )
        if isinstance(filt, str):
            s = _s(filt)
            if s and s not in names and s not in _SKIP_STAFF:
                names.append(s)
    except Exception:
        pass
    return names


def _clients_from_store(staff: str, store: dict | None = None) -> list[str]:
    names: list[str] = []
    src = store if isinstance(store, dict) else load_store()
    want = _s(staff)
    for row in list(src.get("visits") or []) + list(src.get("todos") or []):
        if not isinstance(row, dict):
            continue
        if want and _s(row.get("staff")) not in {"", want}:
            continue
        c = _s(row.get("client") or row.get("title"))
        if c and c not in names and c not in _SKIP_CLIENT:
            names.append(c)
    names.sort()
    return names


def _sales_client_names(df: pd.DataFrame | None, staff: str) -> list[str]:
    """선택 담당자의 매출 거래처. 담당자·행수가 같으면 다시 훑지 않는다."""
    if df is None or df.empty or not staff:
        return []
    if "담당자" not in df.columns or "거래처" not in df.columns:
        return []
    n = int(len(df))
    box = _vc_state_dict("_vc_staff_clients")
    sig = (staff, n, "c2")
    hit = box.get("names")
    if box.get("sig") == sig and isinstance(hit, list):
        return hit
    work = df[df["담당자"].astype(str).str.strip() == staff]
    names = {
        _s(x)
        for x in pd.unique(work["거래처"].astype(str))
        if _s(x) and _s(x) not in _SKIP_CLIENT
    }
    if "거래처_원본" in work.columns:
        names |= {
            _s(x)
            for x in pd.unique(work["거래처_원본"].astype(str))
            if _s(x) and _s(x) not in _SKIP_CLIENT
        }
    out = sorted(names)
    box.clear()
    box["sig"] = sig
    box["names"] = out
    return out


def _staff_clients(df: pd.DataFrame | None, staff: str, store: dict | None = None) -> list[str]:
    """담당자 매출 거래처 전부 + 방문·할일에 적힌 이름."""
    names = set(_clients_from_store(staff, store))
    names.update(_sales_client_names(df, staff))
    try:
        cur = _s(st.session_state.get("vc_client"))
        if cur and cur not in _SKIP_CLIENT:
            names.add(cur)
    except Exception:
        pass
    return sorted(names)


def _staff_names(df: pd.DataFrame | None) -> list[str]:
    """담당자 목록. 매출 담당자 열 전체 스캔은 하지 않는다."""
    box = _vc_state_dict("_vc_staff_names")
    names = _staff_from_store()
    if box.get("names") != names:
        box.clear()
        box["names"] = names
    return names


def _sales_by_client(df: pd.DataFrame | None, staff: str) -> dict[str, list[dict]]:
    """담당자 매출을 거래처 키로 한 번만 묶어 두고, 거래처 변경 때 재계산하지 않는다."""
    if df is None or df.empty:
        return {}
    n = int(len(df))
    sig = (staff, n, tuple(df.columns[:12]), "idx1")
    box = _vc_state_dict("_vc_sales_by_client")
    hit = box.get("map")
    if box.get("sig") == sig and isinstance(hit, dict):
        return hit
    if "거래처" not in df.columns or "매출일_dt" not in df.columns:
        box.clear()
        box["sig"] = sig
        box["map"] = {}
        return {}
    work = df
    if staff and "담당자" in work.columns:
        work = work[work["담당자"].astype(str).str.strip() == staff]
    if work.empty:
        box.clear()
        box["sig"] = sig
        box["map"] = {}
        return {}
    names = work["거래처"].astype(str)
    keymap = {x: _company_key(x) for x in pd.unique(names)}
    keys = names.map(keymap)
    isos = pd.to_datetime(work["매출일_dt"], errors="coerce").dt.strftime("%Y-%m-%d")
    if "품목명" in work.columns:
        items = work["품목명"].astype(str)
        skip = items.str.contains(r"이월\s*미수|\[이월", na=False, regex=True)
    else:
        items = pd.Series("", index=work.index, dtype=str)
        skip = pd.Series(False, index=work.index)
    qtys = (
        pd.to_numeric(work["출고량"], errors="coerce")
        if "출고량" in work.columns
        else pd.Series(0.0, index=work.index)
    )
    amts = (
        pd.to_numeric(work["매출액"], errors="coerce")
        if "매출액" in work.columns
        else pd.Series(0.0, index=work.index)
    )
    bulk = items.str.upper().str.contains("BULK", na=False) | items.str.contains("벌크", na=False)
    out: dict[str, list[dict]] = {}
    for key, iso, item, qty, amt, is_bulk, cname, dropped in zip(
        keys.tolist(),
        isos.tolist(),
        items.tolist(),
        qtys.tolist(),
        amts.tolist(),
        bulk.tolist(),
        names.tolist(),
        skip.tolist(),
    ):
        if dropped or not key or not iso or iso == "NaT":
            continue
        name = _s(item) or "납품"
        if name.lower() in {"nan", "none"}:
            name = "납품"
        out.setdefault(str(key), []).append(
            {
                "date": iso,
                "item": name,
                "qty": 0.0 if pd.isna(qty) else float(qty),
                "amount": 0.0 if pd.isna(amt) else float(amt),
                "staff": staff,
                "client": _s(cname),
                "source": "납품",
                "bulk": bool(is_bulk),
            }
        )
    for rows in out.values():
        rows.sort(key=lambda x: (str(x.get("date") or ""), str(x.get("item") or "")), reverse=True)
    box.clear()
    box["sig"] = sig
    box["map"] = out
    return out


def _sales_delivery_rows(df: pd.DataFrame | None, staff: str, client: str) -> list[dict]:
    """선택 거래처의 납품 상세(날짜·품목·출고량·매출액)."""
    if df is None or df.empty or not client:
        return []
    key = _company_key(client)
    if not key:
        return []
    return list(_sales_by_client(df, staff).get(key) or [])


def _t2d_delivery_rows(
    df: pd.DataFrame | None, staff: str, client: str, month: date | None = None
) -> list[dict]:
    """이미 좁힌 거래처 매출만 납품 행으로. 방문탭 전체 맵·full_df 스캔을 하지 않는다."""
    if df is None or df.empty or len(_s(client)) < 2:
        return []
    if "매출일_dt" not in df.columns:
        return []
    work = df
    if staff and "담당자" in work.columns:
        col = work["담당자"]
        if col.dtype == object:
            if not bool(col.eq(staff).all()):
                work = work.loc[col.astype(str).str.strip() == staff]
        else:
            work = work.loc[col.astype(str).str.strip() == staff]
        if work.empty:
            return []
    if month is not None:
        dts = pd.to_datetime(work["매출일_dt"], errors="coerce")
        work = work.loc[(dts.dt.year == month.year) & (dts.dt.month == month.month)]
        if work.empty:
            return []
    names = work["거래처"].astype(str) if "거래처" in work.columns else pd.Series("", index=work.index, dtype=str)
    isos = pd.to_datetime(work["매출일_dt"], errors="coerce").dt.strftime("%Y-%m-%d")
    if "품목명" in work.columns:
        items = work["품목명"].astype(str)
        skip = items.str.contains(r"이월\s*미수|\[이월", na=False, regex=True)
    else:
        items = pd.Series("", index=work.index, dtype=str)
        skip = pd.Series(False, index=work.index)
    qtys = (
        pd.to_numeric(work["출고량"], errors="coerce")
        if "출고량" in work.columns
        else pd.Series(0.0, index=work.index)
    )
    amts = (
        pd.to_numeric(work["매출액"], errors="coerce")
        if "매출액" in work.columns
        else pd.Series(0.0, index=work.index)
    )
    bulk = items.str.upper().str.contains("BULK", na=False) | items.str.contains("벌크", na=False)
    out: list[dict] = []
    for iso, item, qty, amt, is_bulk, cname, dropped in zip(
        isos.tolist(),
        items.tolist(),
        qtys.tolist(),
        amts.tolist(),
        bulk.tolist(),
        names.tolist(),
        skip.tolist(),
    ):
        if dropped or not iso or iso == "NaT":
            continue
        name = _s(item) or "납품"
        if name.lower() in {"nan", "none"}:
            name = "납품"
        out.append(
            {
                "date": iso,
                "item": name,
                "qty": 0.0 if pd.isna(qty) else float(qty),
                "amount": 0.0 if pd.isna(amt) else float(amt),
                "staff": staff,
                "client": _s(cname),
                "source": "납품",
                "bulk": bool(is_bulk),
            }
        )
    out.sort(key=lambda x: (str(x.get("date") or ""), str(x.get("item") or "")), reverse=True)
    return out



def delivery_rows_on(rows: list[dict], d: date | str | None) -> list[dict]:
    iso = _iso(d)
    if not iso:
        return []
    return [r for r in (rows or []) if r.get("date") == iso]


def _sales_visit_dates(df: pd.DataFrame | None, staff: str, client: str) -> list[dict]:
    rows = _sales_delivery_rows(df, staff, client)
    days = sorted({r.get("date") or "" for r in rows if r.get("date")}, reverse=True)
    return [{"date": d, "staff": staff, "client": client, "source": "납품", "note": ""} for d in days]


def _vc_state_dict(name: str) -> dict:
    try:
        cur = st.session_state.get(name)
        if not isinstance(cur, dict):
            cur = {}
            st.session_state[name] = cur
        return cur
    except Exception:
        return {}


def _worklog_visit_dates(client: str) -> list[dict]:
    if len(_s(client)) < 2:
        return []
    key = _company_key(client)
    cached = _ensure_worklog_index()
    hits = cached.get(key) or []
    out = []
    for iso in hits:
        out.append({"date": iso, "staff": "", "client": client, "source": "업무일지", "note": ""})
    return out


def _ensure_worklog_index() -> dict[str, list[str]]:
    try:
        cached = st.session_state.get("_vc_wl_idx")
        if isinstance(cached, dict) and st.session_state.get("_vc_wl_idx_ready"):
            return cached
    except Exception:
        cached = None
    built = _build_worklog_client_index()
    try:
        st.session_state["_vc_wl_idx"] = built
        st.session_state["_vc_wl_idx_ready"] = True
    except Exception:
        pass
    return built


def _build_worklog_client_index() -> dict[str, list[str]]:
    idx: dict[str, set[str]] = {}
    try:
        import worklog_tab as wt
    except Exception:
        return {}
    try:
        days = sorted(wt.list_saved_worklog_dates())
    except Exception:
        return {}
    for iso in days:
        try:
            d = date.fromisoformat(iso)
        except ValueError:
            continue
        try:
            cells = wt.read_worklog_cells(d)
        except Exception:
            continue
        names = set()
        for r in getattr(wt, "WL_CLIENT_ROWS", range(8, 40)):
            name = _s(cells.get(f"C{r}"))
            if name:
                names.add(_company_key(name))
        try:
            extras = wt.read_worklog_extra_page_cells(d)
        except Exception:
            extras = []
        for extra in extras or []:
            if not isinstance(extra, dict):
                continue
            for r in getattr(wt, "WL_CLIENT_ROWS", range(8, 40)):
                name = _s(extra.get(f"C{r}"))
                if name:
                    names.add(_company_key(name))
        for k in names:
            if not k:
                continue
            idx.setdefault(k, set()).add(iso)
    return {k: sorted(v, reverse=True) for k, v in idx.items()}


def merge_visit_history(df: pd.DataFrame | None, staff: str, client: str) -> list[dict]:
    """직접입력 + 업무일지 + 납품일을 날짜 내림차순으로 합친다."""
    if len(_s(client)) < 2:
        return []
    key = _company_key(client)
    rows: list[dict] = []
    seen: set[tuple[str, str]] = set()
    store = load_store()
    for v in store.get("visits") or []:
        if _company_key(v.get("client") or "") != key:
            continue
        if staff and _s(v.get("staff")) and _s(v.get("staff")) != staff:
            continue
        iso = _iso(v.get("date"))
        if not iso:
            continue
        sig = (iso, "직접입력")
        if sig in seen:
            continue
        seen.add(sig)
        rows.append({**v, "date": iso, "source": "직접입력"})
    for v in _worklog_visit_dates(client):
        sig = (v["date"], "업무일지")
        if sig in seen:
            continue
        seen.add(sig)
        rows.append(v)
    for v in _sales_visit_dates(df, staff, client):
        sig = (v["date"], "납품")
        if sig in seen:
            continue
        seen.add(sig)
        rows.append(v)
    rows.sort(key=lambda x: str(x.get("date") or ""), reverse=True)
    return rows


def _cached_history(df: pd.DataFrame | None, staff: str, client: str) -> list[dict]:
    """담당자·거래처가 같으면 납품·업무일지 이력을 다시 긁지 않는다."""
    if len(_s(client)) < 2:
        return []
    store_mtime = os.path.getmtime(VC_STORE) if os.path.exists(VC_STORE) else 0
    n = 0 if df is None or df.empty else int(len(df))
    sig = (staff, _company_key(client), n, store_mtime)
    box = _vc_state_dict("_vc_hist_cache")
    if box.get("sig") == sig and isinstance(box.get("rows"), list):
        return box["rows"]
    rows = merge_visit_history(df, staff, client)
    box.clear()
    box["sig"] = sig
    box["rows"] = rows
    return rows


def marked_dates(store: dict, history: list[dict], month: date) -> dict[str, set[str]]:
    """날짜 → {todo, visit, worklog, sales} 표시용."""
    chips = calendar_chips(store, history, month)
    return {iso: {c.get("kind") or "visit" for c in items} for iso, items in chips.items()}


def calendar_chips(
    store: dict,
    history: list[dict],
    month: date,
    deliveries: list[dict] | None = None,
) -> dict[str, list[dict]]:
    """구글 캘린더 칸에 붙일 이벤트 칩. 벌크는 품목+충전량, 그외는 실린더."""
    visible = {
        d.isoformat()
        for week in calendar.Calendar(firstweekday=0).monthdatescalendar(month.year, month.month)
        for d in week
    }
    out: dict[str, list[dict]] = {}
    seen: set[tuple[str, str, str]] = set()

    def _add(iso: str, label: str, kind: str, color: str) -> None:
        if not iso or not label or iso not in visible:
            return
        sig = (iso, kind, label)
        if sig in seen:
            return
        seen.add(sig)
        out.setdefault(iso, []).append({"label": label[:18], "kind": kind, "color": color})

    for t in store.get("todos") or []:
        if t.get("done"):
            continue
        _add(_iso(t.get("due")), _s(t.get("title")) or "할일", "todo", "#f6bf26")
    for v in store.get("visits") or []:
        stt = _s(v.get("status")) or "done"
        kind = "planned" if stt == "planned" else "visit"
        color = "#e37400" if kind == "planned" else "#137333"
        _add(_iso(v.get("date")), _s(v.get("client")) or ("방문예정" if kind == "planned" else "방문"), kind, color)
    for v in history or []:
        src = str(v.get("source") or "")
        if src == "업무일지":
            _add(_iso(v.get("date")), "업무일지", "worklog", "#0b8043")
        elif src == "직접입력":
            _add(_iso(v.get("date")), _s(v.get("client")) or "방문", "visit", "#1a73e8")
    by_day: dict[str, list[dict]] = {}
    for r in deliveries or []:
        iso = _iso(r.get("date"))
        if iso:
            by_day.setdefault(iso, []).append(r)
    for iso, items in by_day.items():
        totals: dict[str, float] = {}
        has_other = False
        for x in items:
            if x.get("bulk") or _is_bulk_item(x.get("item") or ""):
                key = _bulk_short(str(x.get("item") or ""))
                totals[key] = totals.get(key, 0.0) + float(x.get("qty") or 0)
            else:
                has_other = True
        for name, qty in totals.items():
            q = _fmt_qty(qty)
            _add(iso, f"{name} {q}" if q else name, "bulk", "#1a73e8")
        if has_other:
            _add(iso, _CYLINDER, "other", "#8d6e63")
    return out


def delivery_chips(month: date, deliveries: list[dict] | None = None) -> dict[str, list[dict]]:
    """납품(벌크·실린더)만 달력 칩으로 남긴다."""
    return calendar_chips({}, [], month, deliveries)


def _vc_css() -> str:
    return """
    <style>
    .vc-toolbar {
      display: flex; align-items: center; gap: 10px;
      padding: 2px 0 4px; color: #3c4043; min-height: 32px;
    }
    .vc-toolbar .vc-title { font-size: 20px; font-weight: 500; letter-spacing: -0.3px; line-height: 32px; }
    div[class*="st-key-vc_jump_today"],
    div[class*="st-key-vc_prev_month"],
    div[class*="st-key-vc_next_month"] {
      display: flex !important; justify-content: flex-end !important;
    }
    div[class*="st-key-vc_jump_today"] button {
      min-height: 28px !important; height: 28px !important;
      padding: 0 12px !important; border-radius: 8px !important;
      background: #f1f3f4 !important; color: #3c4043 !important;
      border: none !important; box-shadow: none !important;
      font-size: 12px !important; font-weight: 600 !important;
      letter-spacing: -0.2px !important;
    }
    div[class*="st-key-vc_prev_month"] button,
    div[class*="st-key-vc_next_month"] button {
      min-height: 28px !important; height: 28px !important;
      min-width: 28px !important; padding: 0 !important;
      border-radius: 8px !important; background: #f1f3f4 !important;
      color: #3c4043 !important; border: none !important;
      box-shadow: none !important; font-size: 15px !important;
      font-weight: 600 !important; line-height: 28px !important;
    }
    div[class*="st-key-vc_jump_today"] button:hover,
    div[class*="st-key-vc_prev_month"] button:hover,
    div[class*="st-key-vc_next_month"] button:hover {
      background: #e8eaed !important;
    }
    div[class*="st-key-vc_strip_"] { margin: 0 !important; }
    div[class*="st-key-vc_strip_"] button {
      min-height: 2.7rem !important; height: 2.7rem !important;
      padding: 2px 0 !important; border-radius: 10px !important;
      font-size: 10px !important; font-weight: 600 !important;
      background: #fff !important; color: #3c4043 !important;
      border: 1px solid #eceff3 !important; box-shadow: none !important;
      white-space: pre-line !important; line-height: 1.15 !important;
    }
    div[data-testid="stHorizontalBlock"]:has(div[class*="st-key-vc_strip_"]) {
      gap: 3px !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.vc-wd) {
      gap: 3px !important;
      margin-bottom: -0.55rem !important;
    }
    .vc-wd {
      text-align: center; font-size: 10px; font-weight: 700;
      line-height: 1.1; padding: 0; letter-spacing: -0.2px;
    }
    div[data-testid="stHorizontalBlock"]:has(.vc-visit-nm) {
      gap: 3px !important;
      margin-top: -0.4rem !important;
    }
    .vc-visit-nm {
      text-align: center; font-size: 9px; font-weight: 700;
      color: #137333; line-height: 1.15; padding-top: 1px;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    .vc-visit-nm.vc-planned {
      color: #c47d00; font-style: italic; font-weight: 700;
      border-bottom: 1px dashed #e37400;
    }
    .vc-month-col {
      border: 1px solid #e8eaed; border-radius: 12px;
      padding: 8px 10px 10px; max-height: 240px; overflow-y: auto;
    }
    .vc-month-col.done { background: #f7fbf8; border-color: #cfe8d4; }
    .vc-month-col.plan { background: #fffbf5; border-color: #f5d9a8; }
    .vc-month-h { font-size: 15px; font-weight: 700; padding: 0 2px 6px; }
    .vc-month-h.done { color: #137333; }
    .vc-month-h.plan { color: #c47d00; }
    .vc-month-h span { font-weight: 500; font-size: 13px; color: #5f6368; }
    .vc-month-row {
      display: flex; gap: 10px; align-items: baseline;
      padding: 5px 4px; border-bottom: 1px solid #ececec; font-size: 14px;
    }
    .vc-month-row:last-child { border-bottom: none; }
    .vc-month-row .d { font-weight: 700; min-width: 72px; color: #3c4043; flex: 0 0 auto; }
    .vc-month-row .n { font-weight: 500; color: #202124; line-height: 1.35; }
    .vc-month-row .note { color: #80868b; font-size: 12px; padding-left: 6px; }
    .vc-month-row.sel { background: #e8f0fe; border-radius: 6px; }
    .vc-month-row.mine .n { color: #1a73e8; }
    .vc-month-empty { color: #9aa0a6; font-size: 13px; padding: 8px 4px; }
    div[class*="st-key-vc_todo_q"] input,
    div[class*="st-key-vc_todo_name"] input,
    div[class*="st-key-vc_todo_detail"] input {
      border-radius: 10px !important; background: #f1f3f4 !important;
    }
    div[class*="st-key-vc_todo_pick_"] button {
      min-height: auto !important; height: auto !important;
      padding: 0 !important; justify-content: flex-start !important;
      text-align: left !important; background: transparent !important;
      border: none !important; box-shadow: none !important;
      color: #202124 !important; font-size: 15px !important;
      font-weight: 500 !important; line-height: 1.25 !important;
    }
    .vc-todo-edit div[class*="st-key-vc_todo_pick_"] button { color: #1a73e8 !important; }
    div[class*="st-key-vc_todo_chk_"] button,
    div[class*="st-key-vc_todo_star_"] button,
    div[class*="st-key-vc_todo_x_"] button {
      min-height: 2.05rem !important; height: 2.05rem !important;
      border-radius: 999px !important; background: #fff !important;
      border: 1px solid #dadce0 !important; box-shadow: none !important;
      color: #5f6368 !important; font-size: 16px !important;
    }
    div[class*="st-key-vc_todo_star_"] button {
      color: #c5c7ca !important;
    }
    div[class*="st-key-vc_visit_chk_"] button,
    div[class*="st-key-vc_plan_chk_"] button,
    div[class*="st-key-vc_visit_del_"] button {
      min-height: 26px !important; height: 26px !important;
      padding: 0 10px !important; border-radius: 6px !important;
      font-size: 12px !important; font-weight: 600 !important;
      letter-spacing: -0.2px !important; box-shadow: none !important;
      border: none !important; background: #f1f3f4 !important;
      color: #3c4043 !important;
    }
    div[class*="st-key-vc_visit_chk_"] button:hover,
    div[class*="st-key-vc_plan_chk_"] button:hover {
      background: #e8eaed !important;
    }
    div[class*="st-key-vc_visit_del_"] button {
      background: transparent !important; color: #9aa0a6 !important;
      padding: 0 8px !important;
    }
    div[class*="st-key-vc_visit_del_"] button:hover {
      background: #fce8e6 !important; color: #d23b3b !important;
    }
    div[class*="st-key-vc_ms_x_"] button,
    div[class*="st-key-vc_log_x_"] button {
      min-height: 22px !important; height: 22px !important;
      min-width: 22px !important; padding: 0 !important;
      border-radius: 6px !important; font-size: 13px !important;
      font-weight: 600 !important; box-shadow: none !important;
      border: none !important; background: transparent !important;
      color: #9aa0a6 !important;
    }
    div[class*="st-key-vc_ms_x_"] button:hover,
    div[class*="st-key-vc_log_x_"] button:hover {
      background: #fce8e6 !important; color: #d23b3b !important;
    }
    .vc-todo-name { font-size: 15px; font-weight: 500; color: #202124; line-height: 1.25; }
    .vc-todo-detail { font-size: 13px; color: #80868b; line-height: 1.3; padding-top: 2px; }
    .vc-todo-row { border-bottom: 1px solid #e8eaed; padding: 2px 0 6px; }
    .vc-todo-row.vc-todo-edit {
      background: #e8f0fe; border-radius: 10px; margin: 0 -6px; padding: 4px 6px 8px;
    }
    .vc-todo-done .vc-todo-name { color: #9aa0a6; text-decoration: line-through; }
    .vc-todo-done .vc-todo-detail { color: #bdc1c6; }
    .vc-todo-done div[class*="st-key-vc_todo_pick_"] button { color: #9aa0a6 !important; text-decoration: line-through !important; }
    </style>
    """


def _shift_month(d: date, delta: int) -> date:
    y, m = d.year, d.month + delta
    while m < 1:
        m += 12
        y -= 1
    while m > 12:
        m -= 12
        y += 1
    return date(y, m, 1)


_VC_BUTTON_PREFIXES = (
    "vc_day_",
    "vc_strip_",
    "vc_prev_month",
    "vc_next_month",
    "vc_jump_today",
    "vc_visit_save_",
    "vc_visit_chk_",
    "vc_plan_chk_",
    "vc_visit_del_",
    "vc_ms_x_",
    "vc_log_x_",
    "vc_todo_del",
    "vc_todo_done",
    "vc_todo_chk_",
    "vc_todo_star_",
    "vc_todo_x_",
    "vc_todo_pick_",
    "vc_todo_show_hidden",
    "vc_todo_cancel_edit",
)


def _purge_vc_button_keys() -> None:
    """버튼 값은 session_state로 넣을 수 없다. 예전 백업·클릭 잔여를 지운다."""
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and k.startswith(_VC_BUTTON_PREFIXES):
            st.session_state.pop(k, None)
    bak = st.session_state.get("_dash_bak_visit")
    if isinstance(bak, dict):
        for k in list(bak):
            if isinstance(k, str) and k.startswith(_VC_BUTTON_PREFIXES):
                bak.pop(k, None)


def _on_pick_day(d: date) -> None:
    st.session_state["_vc_selected"] = d
    st.session_state["_vc_month"] = date(d.year, d.month, 1)
    st.session_state["_vc_open_delivery"] = True


def _on_shift_month(delta: int) -> None:
    cur = st.session_state.get("_vc_month")
    if not isinstance(cur, date):
        cur = date.today().replace(day=1)
    new = _shift_month(cur, delta)
    st.session_state["_vc_month"] = new
    sel = st.session_state.get("_vc_selected")
    if isinstance(sel, date) and (sel.year, sel.month) != (new.year, new.month):
        last = calendar.monthrange(new.year, new.month)[1]
        st.session_state["_vc_selected"] = date(new.year, new.month, min(sel.day, last))


def _on_jump_today() -> None:
    today = date.today()
    st.session_state["_vc_month"] = date(today.year, today.month, 1)
    st.session_state["_vc_selected"] = today


def _on_todo_done(todo_id: str) -> None:
    toggle_todo(str(todo_id))


def _on_todo_star(todo_id: str) -> None:
    toggle_todo_star(str(todo_id))


def _on_todo_delete(todo_id: str) -> None:
    delete_todo(str(todo_id))
    if str(st.session_state.get("_vc_edit_todo") or "") == str(todo_id):
        _clear_todo_edit()


def _clear_todo_edit() -> None:
    st.session_state.pop("_vc_edit_todo", None)
    st.session_state.pop("vc_todo_name", None)
    st.session_state.pop("vc_todo_detail", None)
    st.session_state.pop("vc_todo_due", None)


def _on_pick_todo(todo_id: str) -> None:
    tid = str(todo_id)
    if str(st.session_state.get("_vc_edit_todo") or "") == tid:
        _clear_todo_edit()
        return
    store = load_store()
    hit = next((t for t in store.get("todos") or [] if str(t.get("id")) == tid), None)
    if not hit:
        return
    st.session_state["_vc_edit_todo"] = tid
    st.session_state["vc_todo_name"] = _s(hit.get("title") or hit.get("client"))
    st.session_state["vc_todo_detail"] = _s(hit.get("note"))
    due = _todo_due_date(hit.get("due"))
    if due:
        st.session_state["vc_todo_due"] = due
        st.session_state["_vc_selected"] = due
        st.session_state["_vc_month"] = date(due.year, due.month, 1)


def _on_cancel_todo_edit() -> None:
    _clear_todo_edit()


def _on_toggle_show_done() -> None:
    st.session_state["vc_hide_done"] = not bool(st.session_state.get("vc_hide_done", True))


def _on_toggle_visit(d: date, staff: str, client: str, status: str = "done") -> None:
    if len(_s(client)) < 2:
        return
    if status not in {"done", "planned"}:
        status = "done"
    hit = _direct_visit_on(load_store(), d, client)
    if hit:
        cur = _s(hit.get("status")) or "done"
        if cur == status:
            delete_visit(str(hit.get("id") or ""))
        else:
            store = load_store()
            for v in store.get("visits") or []:
                if str(v.get("id")) == str(hit.get("id")):
                    v["status"] = status
                    save_store(store)
                    break
    else:
        add_visit({"date": d, "staff": staff, "client": client, "status": status})
    st.session_state.pop("_vc_hist_cache", None)


def _on_delete_day_visit(d: date, client: str) -> None:
    clear_day_visit(d, client)
    st.session_state.pop("_vc_hist_cache", None)


def _on_delete_visit_id(visit_id: str) -> None:
    delete_visit(str(visit_id))
    st.session_state.pop("_vc_hist_cache", None)


def _vc_rerun() -> None:
    """달력·할일 조작은 fragment만 다시 그린다."""
    try:
        st.rerun(scope="fragment")
    except Exception:
        st.rerun()


def render_visit_calendar_tab(df: pd.DataFrame | None = None, latest_update_str: str = "") -> None:
    """방문 미팅 캘린더 + 할일 목록."""
    st.markdown(
        "<div class='sub-header dashboard-tab-panel-head'>📅 방문·할일</div>",
        unsafe_allow_html=True,
    )
    _purge_vc_button_keys()

    @st.fragment
    def _visit_body() -> None:
        _render_visit_body(df, latest_update_str)

    _visit_body()


def _render_visit_body(df: pd.DataFrame | None, latest_update_str: str) -> None:
    today = date.today()
    if "_vc_month" not in st.session_state:
        st.session_state["_vc_month"] = date(today.year, today.month, 1)
    if "_vc_selected" not in st.session_state:
        st.session_state["_vc_selected"] = today

    store0 = load_store()
    staffs = _staff_from_store(store0)
    default_staff = "김혁수" if "김혁수" in staffs else (staffs[0] if staffs else "")
    if "vc_staff" not in st.session_state:
        st.session_state["vc_staff"] = st.session_state.get("_vc_staff") or default_staff
    if "vc_client" not in st.session_state:
        prev = _s(st.session_state.get("_vc_client"))
        st.session_state["vc_client"] = prev or None
    elif st.session_state.get("vc_client") == "":
        st.session_state["vc_client"] = None

    month = st.session_state.get("_vc_month") or date(today.year, today.month, 1)
    st.markdown(_vc_css(), unsafe_allow_html=True)

    head_l, head_r = st.columns([2.35, 1.05])
    with head_l:
        st.markdown(
            f"<div class='vc-toolbar'><span class='vc-title'>{month.year}년 {month.month}월</span></div>",
            unsafe_allow_html=True,
        )
    with head_r:
        n1, n2, n3 = st.columns([1.15, 0.42, 0.42], gap="small")
        with n1:
            st.button("오늘", key="vc_jump_today", width="content", on_click=_on_jump_today)
        with n2:
            st.button("‹", key="vc_prev_month", width="content", on_click=_on_shift_month, args=(-1,))
        with n3:
            st.button("›", key="vc_next_month", width="content", on_click=_on_shift_month, args=(1,))
    f1, f2 = st.columns([1, 1])
    with f1:
        staff = st.selectbox("담당자", options=staffs or [""], key="vc_staff")
    clients = _staff_clients(df, staff, store0)
    with f2:
        client = (
            st.selectbox(
                "거래처",
                options=clients or [""],
                index=None,
                placeholder="거래처명 입력",
                key="vc_client",
                accept_new_options=True,
                help="이 담당자의 매출 거래처가 모두 나옵니다. 고르면 달력·방문·할일에 연동됩니다.",
            )
            or ""
        )

    st.session_state["_vc_staff"] = staff
    st.session_state["_vc_client"] = client

    @st.fragment
    def _visit_day_block() -> None:
        sel: date = st.session_state.get("_vc_selected") or today
        mon = st.session_state.get("_vc_month") or month
        stf = str(st.session_state.get("vc_staff") or staff or "")
        cli = str(st.session_state.get("vc_client") or client or "")
        store = load_store()
        deliveries = _sales_delivery_rows(df, stf, cli) if cli else []
        prefix = f"{mon.year:04d}-{mon.month:02d}-"
        month_deliveries = [r for r in deliveries if str(r.get("date") or "").startswith(prefix)]
        history = _worklog_visit_dates(cli) if cli else []
        chips = calendar_chips(store, history, mon, month_deliveries)
        day_deliveries = delivery_rows_on(deliveries, sel)
        _render_day_strip(mon, sel, chips, today)
        st.caption("달력 아래 짧은 이름은 표시용입니다. 이번 달 일정은 아래 기방문·방문예정 목록에서 보세요.")
        _render_month_schedule(mon, store, stf, sel, cli)

        left, right = st.columns([1, 1], gap="medium")
        with left:
            if cli and day_deliveries:
                _render_delivery_list(cli, sel, day_deliveries, deliveries)
            elif cli:
                st.markdown(
                    f"<div style='font-size:16px;font-weight:500;color:#3c4043;padding:8px 0 4px'>"
                    f"지정 납품 목록 · {cli}</div>",
                    unsafe_allow_html=True,
                )
                st.caption("이 날 지정 납품이 없습니다. 위 줄에서 벌크·실린더가 있는 날을 누르세요.")
            else:
                st.caption("담당자와 거래처를 고르면 지정 납품 목록이 여기에 표시됩니다.")
            _render_day_agenda(sel, store, stf, cli)
            _render_visit_log(store, stf, cli)

        with right:
            _render_todo_panel(sel, stf, cli, store)

        if latest_update_str:
            st.caption(f"대시보드 기준 시각: {latest_update_str}")

    _visit_day_block()


def _client_short(name: str) -> str:
    t = _s(name)
    t = re.sub(r"\(.*$", "", t)
    t = re.sub(r"(주식회사|유한회사|유한책임회사|\(주\)|㈜|㈔)", "", t)
    t = re.sub(r"\s+", "", t)
    return (t or _s(name))[:8]


def _strip_visit_mark(items: list[dict]) -> tuple[str, str]:
    visit_lab = ""
    plan_lab = ""
    for c in items or []:
        lab = _s(c.get("label"))
        if c.get("kind") == "visit" and lab and lab != "방문":
            visit_lab = _client_short(lab)
        elif c.get("kind") == "planned" and lab:
            plan_lab = _client_short(lab)
    if visit_lab:
        return visit_lab, "visit"
    if plan_lab:
        return f"예·{plan_lab[:6]}", "planned"
    return "", ""


def _strip_visit_name(items: list[dict]) -> str:
    return _strip_visit_mark(items)[0]


def _strip_tag(items: list[dict]) -> str:
    kinds = {c.get("kind") for c in (items or [])}
    has_bulk = "bulk" in kinds
    has_cyl = "other" in kinds
    if has_bulk and has_cyl:
        return "벌·실"
    if has_bulk:
        return "벌크"
    if has_cyl:
        return _CYLINDER
    if "todo" in kinds:
        return "할일"
    if "worklog" in kinds:
        return "일지"
    return ""


def _weekday_name(d: date) -> str:
    return _WEEKDAYS[d.weekday()]


def _weekday_color(d: date) -> str:
    wd = d.weekday()
    if wd == 5:
        return _VC_SAT_FG
    if wd == 6:
        return _VC_SUN_FG
    return "#6b7280"


def _strip_cell_kind(items: list[dict], wd: int) -> str:
    kinds = {c.get("kind") for c in (items or [])}
    if "bulk" in kinds and "other" in kinds:
        return "mix"
    if "bulk" in kinds:
        return "bulk"
    if "other" in kinds:
        return "other"
    if "visit" in kinds:
        return "visit"
    if "planned" in kinds:
        return "planned"
    if "todo" in kinds:
        return "todo"
    if "worklog" in kinds:
        return "worklog"
    if wd == 5:
        return "sat"
    if wd == 6:
        return "sun"
    return ""


def _strip_days_payload(
    month: date, chips: dict[str, list[dict]], selected: date, today: date
) -> list[dict]:
    last = calendar.monthrange(month.year, month.month)[1]
    out: list[dict] = []
    for day in range(1, last + 1):
        d = date(month.year, month.month, day)
        iso = d.isoformat()
        items = chips.get(iso) or []
        nm, mark = _strip_visit_mark(items)
        out.append(
            {
                "iso": iso,
                "day": day,
                "wd": _weekday_name(d),
                "wdc": _weekday_color(d),
                "tag": _strip_tag(items),
                "name": nm,
                "mark": mark,
                "kind": _strip_cell_kind(items, d.weekday()),
                "sel": iso == selected.isoformat(),
                "today": iso == today.isoformat(),
            }
        )
    return out


_VC_STRIP_HTML = """<div class="vc-strip-root"></div>"""
_VC_STRIP_CSS = f"""
.vc-strip-root {{ width:100%; }}
.vc-strip-wd, .vc-strip-days, .vc-strip-nm {{
  display:flex; gap:3px; width:100%;
}}
.vc-strip-wd {{ margin-bottom:2px; }}
.vc-strip-nm {{ margin-top:2px; }}
.vc-strip-wd span {{
  flex:1; min-width:0; text-align:center; font-size:10px; font-weight:700;
  line-height:1.1; letter-spacing:-0.2px;
}}
.vc-strip-days button {{
  flex:1; min-width:0; min-height:2.7rem; height:2.7rem; padding:2px 0;
  border-radius:10px; font-size:10px; font-weight:600; background:#fff;
  color:#3c4043; border:1px solid #eceff3; box-shadow:none; cursor:pointer;
  white-space:pre-line; line-height:1.15;
}}
.vc-strip-days button.mix {{ background:#ece6f8; color:#5b4b8a; }}
.vc-strip-days button.bulk {{ background:#e8f0fe; color:#1a73e8; }}
.vc-strip-days button.other {{ background:#f6eee4; color:#8d6e63; }}
.vc-strip-days button.visit {{ background:#e6f4ea; color:#137333; }}
.vc-strip-days button.planned {{ background:#fff4e5; color:#c47d00; }}
.vc-strip-days button.todo {{ background:#fef7e0; color:#b06000; }}
.vc-strip-days button.worklog {{ background:#e6f4ea; color:#0b8043; }}
.vc-strip-days button.sat {{ background:{_VC_SAT_BG}; color:{_VC_SAT_FG}; }}
.vc-strip-days button.sun {{ background:{_VC_SUN_BG}; color:{_VC_SUN_FG}; }}
.vc-strip-days button.today {{ box-shadow:inset 0 0 0 1.5px #1a73e8; }}
.vc-strip-days button.sel {{
  background:#1a73e8 !important; color:#fff !important; border-color:#1a73e8 !important;
  box-shadow:none;
}}
.vc-strip-nm span {{
  flex:1; min-width:0; text-align:center; font-size:9px; font-weight:700;
  color:#137333; line-height:1.15; padding-top:1px;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
}}
.vc-strip-nm span.planned {{
  color:#c47d00; font-style:italic; border-bottom:1px dashed #e37400;
}}
"""
_VC_STRIP_JS = r"""
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => (
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]
  ));
}

export default function (component) {
  const { data, parentElement, setStateValue } = component;
  const root = parentElement.querySelector(".vc-strip-root");
  if (!root) return;
  const days = Array.isArray(data && data.days) ? data.days : [];
  const selected = String((data && data.selected) || "");
  let html = '<div class="vc-strip-wd">';
  for (const d of days) {
    html += '<span style="color:' + esc(d.wdc || "#6b7280") + '">' + esc(d.wd) + "</span>";
  }
  html += '</div><div class="vc-strip-days">';
  for (const d of days) {
    const iso = String(d.iso || "");
    const cls = [d.kind || "", iso === selected ? "sel" : "", d.today ? "today" : ""]
      .filter(Boolean)
      .join(" ");
    const lab = d.tag ? String(d.day) + "\n" + String(d.tag) : String(d.day);
    html += '<button type="button" data-iso="' + esc(iso) + '" class="' + cls + '">' + esc(lab) + "</button>";
  }
  html += "</div>";
  if (!(data && data.hideNames)) {
    html += '<div class="vc-strip-nm">';
    for (const d of days) {
      const nm = String(d.name || "");
      const mcls = d.mark === "planned" ? "planned" : "";
      html += "<span class=\"" + mcls + "\">" + (nm ? esc(nm) : "&nbsp;") + "</span>";
    }
    html += "</div>";
  }
  if (root.innerHTML !== html) root.innerHTML = html;
  if (!root.dataset.vcBound) {
    root.dataset.vcBound = "1";
    root.addEventListener("click", (ev) => {
      const btn = ev.target && ev.target.closest ? ev.target.closest("[data-iso]") : null;
      if (!btn || !root.contains(btn)) return;
      const iso = btn.getAttribute("data-iso") || "";
      if (!iso) return;
      root.querySelectorAll("[data-iso]").forEach((el) => el.classList.toggle("sel", el === btn));
      setStateValue("iso", iso);
    });
  }
}
"""
try:
    _VC_STRIP = st.components.v2.component(
        "visit_day_strip_v1",
        html=_VC_STRIP_HTML,
        css=_VC_STRIP_CSS,
        js=_VC_STRIP_JS,
    )
except Exception:  # pragma: no cover
    _VC_STRIP = None


def _on_strip_iso_change() -> None:
    raw = st.session_state.get("vc_strip_host")
    iso = ""
    if raw is not None:
        iso = str(getattr(raw, "iso", "") or "")
        if not iso and isinstance(raw, dict):
            iso = str(raw.get("iso") or "")
    try:
        d = date.fromisoformat(iso[:10])
    except ValueError:
        return
    if st.session_state.get("_vc_selected") == d:
        return
    _on_pick_day(d)


def _render_day_strip(month: date, selected: date, chips: dict[str, list[dict]], today: date) -> None:
    """이번 달 1일~말일을 가로로 한 줄에 요일·벌크·실린더와 함께 보여 고른다."""
    days = _strip_days_payload(month, chips, selected, today)
    if _VC_STRIP is not None:
        _VC_STRIP(
            key="vc_strip_host",
            data={"days": days, "selected": selected.isoformat(), "today": today.isoformat()},
            default={"iso": selected.isoformat()},
            on_iso_change=_on_strip_iso_change,
        )
        return
    last = calendar.monthrange(month.year, month.month)[1]
    cols = st.columns(last, gap="small")
    for i, cell in enumerate(days):
        d = date.fromisoformat(cell["iso"])
        with cols[i]:
            st.button(
                f"{cell['day']}\n{cell['tag']}" if cell["tag"] else str(cell["day"]),
                key=f"vc_strip_{cell['iso']}",
                type="secondary",
                width="stretch",
                on_click=_on_pick_day,
                args=(d,),
            )


def _month_row_html(row: dict, selected: date, client: str) -> str:
    iso = row.get("iso") or ""
    try:
        d = date.fromisoformat(iso)
        dlab = f"{d.day}일 ({_WEEKDAYS[d.weekday()]})"
    except ValueError:
        dlab = iso
    cls = "vc-month-row"
    sel_iso = selected.isoformat() if isinstance(selected, date) else ""
    if iso == sel_iso:
        cls += " sel"
    sel_key = _company_key(client)
    if sel_key and _company_key(row.get("client") or "") == sel_key:
        cls += " mine"
    note = html.escape(_s(row.get("note")))
    note_h = f"<span class='note'>{note}</span>" if note else ""
    return (
        f"<div class='{cls}'><span class='d'>{html.escape(dlab)}</span>"
        f"<span class='n'>{html.escape(_s(row.get('client')) or '-')}</span>{note_h}</div>"
    )


def _month_schedule_items_html(
    rows: list[dict], selected: date, client: str, kind: str
) -> str:
    title = "기방문" if kind == "visit" else "방문예정"
    hcls = "done" if kind == "visit" else "plan"
    parts = [
        f"<div class='vc-month-col {hcls}'>",
        f"<div class='vc-month-h {hcls}'>{title} <span>{len(rows)}건</span></div>",
    ]
    if not rows:
        parts.append(f"<div class='vc-month-empty'>이 달 {title}이 없습니다.</div>")
    for r in rows:
        parts.append(_month_row_html(r, selected, client))
    parts.append("</div>")
    return "".join(parts)


def _render_month_col(
    title: str, rows: list[dict], selected: date, client: str, kind: str, key_pfx: str
) -> None:
    hcls = "done" if kind == "visit" else "plan"
    st.markdown(
        f"<div class='vc-month-col {hcls}'>"
        f"<div class='vc-month-h {hcls}'>{title} <span>{len(rows)}건</span></div>",
        unsafe_allow_html=True,
    )
    if not rows:
        st.markdown(
            f"<div class='vc-month-empty'>이 달 {title}이 없습니다.</div></div>",
            unsafe_allow_html=True,
        )
        return
    for r in rows:
        lab, btn = st.columns([1.55, 0.22])
        with lab:
            st.markdown(_month_row_html(r, selected, client), unsafe_allow_html=True)
        with btn:
            vid = str(r.get("id") or "")
            if vid:
                st.button(
                    "×",
                    key=f"{key_pfx}{vid}",
                    help="삭제하면 달력·월간 일정·내역에서 사라집니다.",
                    on_click=_on_delete_visit_id,
                    args=(vid,),
                )
    st.markdown("</div>", unsafe_allow_html=True)


def _render_month_schedule(
    month: date, store: dict, staff: str, selected: date, client: str
) -> None:
    rows = month_visit_rows(store, month, staff)
    done = [r for r in rows if r.get("status") != "planned"]
    planned = [r for r in rows if r.get("status") == "planned"]
    st.markdown(
        f"<div style='font-size:16px;font-weight:500;color:#3c4043;padding:10px 0 6px'>"
        f"{month.year}년 {month.month}월 기방문 · 방문예정"
        f"<span style='font-size:13px;font-weight:400;color:#5f6368'> · 기방문 {len(done)} · 방문예정 {len(planned)}</span>"
        f"</div>",
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2, gap="medium")
    with c1:
        _render_month_col("기방문", done, selected, client, "visit", "vc_ms_x_")
    with c2:
        _render_month_col("방문예정", planned, selected, client, "planned", "vc_ms_x_")


def _render_day_agenda(selected: date, store: dict, staff: str, client: str) -> None:
    iso = selected.isoformat()
    hit = _direct_visit_on(store, selected, client)
    cur = _s((hit or {}).get("status")) or "done"
    done = bool(hit) and cur != "planned"
    planned = bool(hit) and cur == "planned"
    off = len(_s(client)) < 2
    cluster, _rest = st.columns([1.15, 1.85])
    with cluster:
        b1, b2, b3 = st.columns([0.85, 1.15, 0.65], gap="small")
        with b1:
            st.button(
                "방문",
                key=f"vc_visit_chk_{iso}",
                type="secondary",
                width="content",
                disabled=off,
                on_click=_on_toggle_visit,
                args=(selected, staff, client, "done"),
                help="실제 방문. 달력 아래에 초록 업체명.",
            )
        with b2:
            st.button(
                "방문예정",
                key=f"vc_plan_chk_{iso}",
                type="secondary",
                width="content",
                disabled=off,
                on_click=_on_toggle_visit,
                args=(selected, staff, client, "planned"),
                help="방문 예정. 달력 아래에 주황 예·업체명.",
            )
        with b3:
            st.button(
                "삭제",
                key=f"vc_visit_del_{iso}",
                type="secondary",
                width="content",
                disabled=off or not hit,
                on_click=_on_delete_day_visit,
                args=(selected, client),
                help="이 날 방문·방문예정을 지웁니다. 달력과 목록에서 사라집니다.",
            )
    on_css = []
    if done:
        on_css.append(
            f'div[class*="st-key-vc_visit_chk_{iso}"] button{{background:#e6f4ea!important;color:#137333!important;}}'
        )
    if planned:
        on_css.append(
            f'div[class*="st-key-vc_plan_chk_{iso}"] button{{background:#fff4e5!important;color:#c47d00!important;}}'
        )
    if on_css:
        st.markdown(f"<style>{''.join(on_css)}</style>", unsafe_allow_html=True)


def _render_visit_log(store: dict, staff: str, client: str) -> None:
    st.markdown(
        "<div style='font-size:16px;font-weight:500;color:#3c4043;padding:16px 0 6px'>방문 내역</div>",
        unsafe_allow_html=True,
    )
    rows = [v for v in (store.get("visits") or []) if isinstance(v, dict)]
    if client:
        key = _company_key(client)
        rows = [v for v in rows if _company_key(v.get("client") or "") == key]
    elif staff:
        rows = [v for v in rows if not _s(v.get("staff")) or _s(v.get("staff")) == staff]
    rows = sorted(rows, key=lambda v: _iso(v.get("date")) or "", reverse=True)
    extra = max(0, len(rows) - 40)
    rows = rows[:40]
    if not rows:
        st.caption("체크한 방문 내역이 없습니다.")
        return
    for v in rows:
        d = _iso(v.get("date"))
        wd = ""
        try:
            if d:
                wd = f" ({_WEEKDAYS[date.fromisoformat(d).weekday()]})"
        except ValueError:
            wd = ""
        note = f" · {_s(v.get('note'))}" if _s(v.get("note")) else ""
        kind = "방문예정" if (_s(v.get("status")) or "done") == "planned" else "방문"
        lab, btn = st.columns([1.55, 0.18])
        with lab:
            st.markdown(f"- **{d}{wd}** · {kind} · {v.get('client') or '-'}{note}")
        with btn:
            vid = str(v.get("id") or "")
            if vid:
                st.button(
                    "×",
                    key=f"vc_log_x_{vid}",
                    help="삭제하면 달력·월간 일정·내역에서 사라집니다.",
                    on_click=_on_delete_visit_id,
                    args=(vid,),
                )
    if extra:
        st.caption(f"최근 40건만 표시 · 나머지 {extra}건")


def _render_delivery_list(
    client: str,
    selected: date,
    day_rows: list[dict],
    all_rows: list[dict],
    *,
    heading: str | None = None,
) -> None:
    wd = _WEEKDAYS[selected.weekday()]
    title = heading or "지정 납품 목록"
    st.markdown(
        f"<div style='font-size:16px;font-weight:500;color:#3c4043;padding:4px 0 6px'>"
        f"{title} · {client} · {selected.month}월 {selected.day}일 ({wd})</div>",
        unsafe_allow_html=True,
    )
    ordered = sorted(
        day_rows,
        key=lambda r: (0 if (r.get("bulk") or _is_bulk_item(r.get("item") or "")) else 1, str(r.get("item") or "")),
    )
    n_bulk = sum(1 for r in ordered if r.get("bulk") or _is_bulk_item(r.get("item") or ""))
    n_other = len(ordered) - n_bulk
    st.caption(f"벌크 {n_bulk}건 · {_CYLINDER} {n_other}건")
    table = pd.DataFrame(
        [
            {
                "구분": "벌크" if (r.get("bulk") or _is_bulk_item(r.get("item") or "")) else _CYLINDER,
                "날짜": r.get("date"),
                "품목": r.get("item"),
                "충전량": _fmt_qty(r.get("qty")) if (r.get("bulk") or _is_bulk_item(r.get("item") or "")) else "",
                "출고량": "" if (r.get("bulk") or _is_bulk_item(r.get("item") or "")) else _fmt_qty(r.get("qty")),
                "매출액": _fmt_qty(r.get("amount")),
            }
            for r in ordered
        ]
    )
    st.dataframe(table, width="stretch", hide_index=True, height=min(380, 90 + 36 * max(len(ordered), 3)))


def _render_todo_panel(selected: date, staff: str, client: str, store: dict) -> None:
    st.markdown(
        "<div style='font-size:16px;font-weight:500;color:#3c4043;padding:2px 0 8px'>내 할일 목록</div>",
        unsafe_allow_html=True,
    )
    q = st.text_input("검색", key="vc_todo_q", placeholder="업체명 · 세부사항 검색", label_visibility="collapsed")
    if "vc_hide_done" not in st.session_state:
        st.session_state["vc_hide_done"] = True
    hide_done = bool(st.session_state.get("vc_hide_done", True))
    all_todos = list(store.get("todos") or [])
    hidden_n = sum(1 for t in all_todos if t.get("done"))

    todos = all_todos
    qq = _s(q).casefold()
    if qq:
        todos = [
            t
            for t in todos
            if qq in _s(t.get("title")).casefold()
            or qq in _s(t.get("client")).casefold()
            or qq in _s(t.get("note")).casefold()
        ]
    if hide_done:
        todos = [t for t in todos if not t.get("done")]
    todos.sort(key=lambda t: (0 if t.get("starred") else 1, 1 if t.get("done") else 0, _iso(t.get("due")) or "9999", _s(t.get("title"))))
    edit_id = str(st.session_state.get("_vc_edit_todo") or "")
    edit_item = next((t for t in all_todos if str(t.get("id")) == edit_id), None)
    if edit_id and not edit_item:
        st.session_state.pop("_vc_edit_todo", None)
        edit_id = ""

    if not todos:
        st.caption("할일이 없습니다. 아래에서 업체명과 세부사항을 넣어 추가하세요.")
    else:
        for t in todos:
            tid = str(t.get("id") or "")
            plain_name = _s(t.get("title") or t.get("client")) or "할일"
            detail = html.escape(_s(t.get("note")))
            due_s = _iso(t.get("due"))
            if due_s:
                detail = f"{detail} · {due_s[5:]}" if detail else due_s[5:]
            row_cls = "vc-todo-row"
            if t.get("done"):
                row_cls += " vc-todo-done"
            if tid == edit_id:
                row_cls += " vc-todo-edit"
            st.markdown(f"<div class='{row_cls}'>", unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns([0.16, 1.42, 0.2, 0.18])
            with c1:
                st.button(
                    "●" if t.get("done") else "○",
                    key=f"vc_todo_chk_{tid}",
                    width="stretch",
                    help="완료 표시",
                    on_click=_on_todo_done,
                    args=(tid,),
                )
            with c2:
                st.button(
                    plain_name,
                    key=f"vc_todo_pick_{tid}",
                    width="stretch",
                    help="다시 눌러 편집",
                    on_click=_on_pick_todo,
                    args=(tid,),
                )
                st.markdown(
                    f"<div class='vc-todo-detail'>{detail or '&nbsp;'}</div>",
                    unsafe_allow_html=True,
                )
            with c3:
                st.button(
                    "★" if t.get("starred") else "☆",
                    key=f"vc_todo_star_{tid}",
                    width="stretch",
                    help="중요 표시",
                    on_click=_on_todo_star,
                    args=(tid,),
                )
            with c4:
                st.button(
                    "×",
                    key=f"vc_todo_x_{tid}",
                    width="stretch",
                    help="삭제",
                    on_click=_on_todo_delete,
                    args=(tid,),
                )
            st.markdown("</div>", unsafe_allow_html=True)

    if edit_item:
        cap, cancel = st.columns([1.35, 0.65])
        with cap:
            st.markdown(
                "<div style='font-size:14px;font-weight:500;color:#1a73e8;padding:8px 0 4px'>할 일 수정</div>",
                unsafe_allow_html=True,
            )
        with cancel:
            st.button(
                "편집 취소",
                key="vc_todo_cancel_edit",
                width="stretch",
                on_click=_on_cancel_todo_edit,
            )
    form_title = "" if edit_item else (
        "<div style='font-size:14px;font-weight:500;color:#1a73e8;padding:8px 0 4px'>할 일 추가</div>"
    )
    with st.form("vc_todo_add_form", clear_on_submit=True):
        if form_title:
            st.markdown(form_title, unsafe_allow_html=True)
        name_in = st.text_input("업체명", placeholder=client or "업체명", key="vc_todo_name")
        detail_in = st.text_input("세부사항", placeholder="세부사항", key="vc_todo_detail")
        due_default = _todo_due_date(edit_item.get("due")) if edit_item else selected
        if "vc_todo_due" not in st.session_state:
            st.session_state["vc_todo_due"] = due_default or selected
        due = st.date_input("날짜", format="YYYY/MM/DD", key="vc_todo_due")
        save_label = "저장" if edit_item else "＋ 할 일 추가"
        if st.form_submit_button(save_label, type="primary"):
            try:
                payload = {
                    "title": name_in or client,
                    "note": detail_in,
                    "due": due,
                    "staff": staff,
                    "client": name_in or client,
                }
                if edit_item:
                    update_todo(edit_id, payload)
                    st.session_state.pop("_vc_edit_todo", None)
                else:
                    add_todo(payload)
                if isinstance(due, date):
                    st.session_state["_vc_selected"] = due
                    st.session_state["_vc_month"] = date(due.year, due.month, 1)
                _vc_rerun()
            except ValueError as e:
                st.error(str(e))

    _rest, btm = st.columns([1.1, 0.9])
    with btm:
        st.button(
            "숨긴 할일 보기" if hide_done else "완료 숨기기",
            key="vc_todo_show_hidden",
            width="stretch",
            disabled=hide_done and hidden_n == 0,
            on_click=_on_toggle_show_done,
            help="완료로 숨긴 할일을 다시 보여 줍니다.",
        )
        if hide_done and hidden_n:
            st.caption(f"완료 {hidden_n}건 숨김")
        elif not hide_done:
            st.caption("완료 할일을 함께 표시 중")


_T2D_BUTTON_PREFIXES = (
    "t2d_jump_today",
    "t2d_prev_month",
    "t2d_next_month",
    "t2d_strip_",
)


def _t2d_css() -> str:
    return """
    <style>
    div[class*="st-key-tab2_delivery_status"] {
      background: #fff;
      border: 1px solid #e7edf4;
      border-radius: 16px;
      padding: 12px 16px 14px;
      margin: 0 0 14px;
      box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }
    .t2d-label {
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.14em;
      color: #64748b;
      line-height: 1;
      margin: 2px 0 10px;
    }
    .t2d-meta {
      font-size: 13px;
      font-weight: 500;
      color: #475569;
      margin: -2px 0 8px;
    }
    .t2d-empty {
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 76px;
      background: #f8fafc;
      border-radius: 12px;
      color: #94a3b8;
      font-size: 14px;
      font-weight: 500;
      letter-spacing: -0.2px;
      text-align: center;
      padding: 18px 16px;
    }
    .t2d-toolbar {
      display: flex; align-items: center; gap: 10px;
      padding: 0 0 2px; color: #3c4043; min-height: 32px;
    }
    .t2d-toolbar .t2d-title {
      font-size: 20px; font-weight: 500; letter-spacing: -0.3px; line-height: 32px;
    }
    div[class*="st-key-t2d_jump_today"],
    div[class*="st-key-t2d_prev_month"],
    div[class*="st-key-t2d_next_month"] {
      display: flex !important; justify-content: flex-end !important;
    }
    div[class*="st-key-t2d_jump_today"] button {
      min-height: 28px !important; height: 28px !important;
      padding: 0 12px !important; border-radius: 8px !important;
      background: #f1f3f4 !important; color: #3c4043 !important;
      border: none !important; box-shadow: none !important;
      font-size: 12px !important; font-weight: 600 !important;
      letter-spacing: -0.2px !important;
    }
    div[class*="st-key-t2d_prev_month"] button,
    div[class*="st-key-t2d_next_month"] button {
      min-height: 28px !important; height: 28px !important;
      min-width: 28px !important; padding: 0 !important;
      border-radius: 8px !important; background: #f1f3f4 !important;
      color: #3c4043 !important; border: none !important;
      box-shadow: none !important; font-size: 15px !important;
      font-weight: 600 !important; line-height: 28px !important;
    }
    div[class*="st-key-t2d_jump_today"] button:hover,
    div[class*="st-key-t2d_prev_month"] button:hover,
    div[class*="st-key-t2d_next_month"] button:hover {
      background: #e8eaed !important;
    }
    div[class*="st-key-t2d_strip_"] { margin: 0 !important; }
    div[class*="st-key-t2d_strip_"] button {
      min-height: 2.7rem !important; height: 2.7rem !important;
      padding: 2px 0 !important; border-radius: 10px !important;
      font-size: 10px !important; font-weight: 600 !important;
      background: #fff !important; color: #3c4043 !important;
      border: 1px solid #eceff3 !important; box-shadow: none !important;
      white-space: pre-line !important; line-height: 1.15 !important;
    }
    </style>
    """


def _purge_t2d_button_keys() -> None:
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and k.startswith(_T2D_BUTTON_PREFIXES):
            st.session_state.pop(k, None)


def _on_t2d_pick_day(d: date) -> None:
    st.session_state["_t2d_selected"] = d
    st.session_state["_t2d_month"] = date(d.year, d.month, 1)


def _on_t2d_shift_month(delta: int) -> None:
    cur = st.session_state.get("_t2d_month")
    if not isinstance(cur, date):
        cur = date.today().replace(day=1)
    new = _shift_month(cur, delta)
    st.session_state["_t2d_month"] = new
    sel = st.session_state.get("_t2d_selected")
    if isinstance(sel, date) and (sel.year, sel.month) != (new.year, new.month):
        last = calendar.monthrange(new.year, new.month)[1]
        st.session_state["_t2d_selected"] = date(new.year, new.month, min(sel.day, last))


def _on_t2d_jump_today() -> None:
    today = date.today()
    st.session_state["_t2d_month"] = date(today.year, today.month, 1)
    st.session_state["_t2d_selected"] = today


def _on_t2d_strip_iso_change() -> None:
    raw = st.session_state.get("t2d_strip_host")
    iso = ""
    if raw is not None:
        iso = str(getattr(raw, "iso", "") or "")
        if not iso and isinstance(raw, dict):
            iso = str(raw.get("iso") or "")
    try:
        d = date.fromisoformat(iso[:10])
    except ValueError:
        return
    if st.session_state.get("_t2d_selected") == d:
        return
    _on_t2d_pick_day(d)


def _render_t2d_day_strip(
    month: date, selected: date, chips: dict[str, list[dict]], today: date
) -> None:
    days = _strip_days_payload(month, chips, selected, today)
    if _VC_STRIP is not None:
        _VC_STRIP(
            key="t2d_strip_host",
            data={
                "days": days,
                "selected": selected.isoformat(),
                "today": today.isoformat(),
                "hideNames": True,
            },
            default={"iso": selected.isoformat()},
            on_iso_change=_on_t2d_strip_iso_change,
        )
        return
    last = calendar.monthrange(month.year, month.month)[1]
    cols = st.columns(last, gap="small")
    for i, cell in enumerate(days):
        d = date.fromisoformat(cell["iso"])
        with cols[i]:
            st.button(
                f"{cell['day']}\n{cell['tag']}" if cell["tag"] else str(cell["day"]),
                key=f"t2d_strip_{cell['iso']}",
                type="secondary",
                width="stretch",
                on_click=_on_t2d_pick_day,
                args=(d,),
            )


def render_tab2_delivery_status(
    df: pd.DataFrame | None, staff: str, client: str
) -> None:
    """거래처 분석 상단 납품현황. 담당자·거래처는 메인 고정바 값을 쓴다."""
    _purge_t2d_button_keys()
    st.markdown(_t2d_css(), unsafe_allow_html=True)
    with st.container(key="tab2_delivery_status"):
        st.markdown("<div class='t2d-label'>납품현황</div>", unsafe_allow_html=True)
        staff = _s(staff)
        client = _s(client)

        @st.fragment
        def _t2d_body() -> None:
            today = date.today()
            if "_t2d_month" not in st.session_state:
                st.session_state["_t2d_month"] = date(today.year, today.month, 1)
            if "_t2d_selected" not in st.session_state:
                st.session_state["_t2d_selected"] = today
            if len(client) < 2:
                st.markdown(
                    "<div class='t2d-empty'>담당자와 거래처를 고르면 지정 납품 목록이 여기에 표시됩니다.</div>",
                    unsafe_allow_html=True,
                )
                return
            month = st.session_state.get("_t2d_month") or date(today.year, today.month, 1)
            selected = st.session_state.get("_t2d_selected") or today
            if not isinstance(month, date):
                month = date(today.year, today.month, 1)
            if not isinstance(selected, date):
                selected = today
            meta_bits = [html.escape(x) for x in (staff, client) if x]
            if meta_bits:
                st.markdown(
                    f"<div class='t2d-meta'>{' · '.join(meta_bits)}</div>",
                    unsafe_allow_html=True,
                )
            head_l, head_r = st.columns([2.35, 1.05])
            with head_l:
                st.markdown(
                    f"<div class='t2d-toolbar'><span class='t2d-title'>{month.year}년 {month.month}월</span></div>",
                    unsafe_allow_html=True,
                )
            with head_r:
                n1, n2, n3 = st.columns([1.15, 0.42, 0.42], gap="small")
                with n1:
                    st.button(
                        "오늘",
                        key="t2d_jump_today",
                        width="content",
                        on_click=_on_t2d_jump_today,
                    )
                with n2:
                    st.button(
                        "‹",
                        key="t2d_prev_month",
                        width="content",
                        on_click=_on_t2d_shift_month,
                        args=(-1,),
                    )
                with n3:
                    st.button(
                        "›",
                        key="t2d_next_month",
                        width="content",
                        on_click=_on_t2d_shift_month,
                        args=(1,),
                    )
            deliveries = _t2d_delivery_rows(df, staff, client, month=month)
            month_deliveries = deliveries
            chips = delivery_chips(month, month_deliveries)
            _render_t2d_day_strip(month, selected, chips, today)
            day_rows = delivery_rows_on(deliveries, selected)
            if day_rows:
                _render_delivery_list(client, selected, day_rows, deliveries, heading="납품 내역")
            else:
                st.markdown(
                    "<div class='t2d-empty' style='min-height:56px;margin-top:10px'>"
                    "이 날 납품 내역이 없습니다. 벌크·실린더 표시가 있는 날을 선택하세요.</div>",
                    unsafe_allow_html=True,
                )

        _t2d_body()

