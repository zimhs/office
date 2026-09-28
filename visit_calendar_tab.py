"""방문·할일 캘린더 — 가로 일자 줄, 왼쪽 지정 납품 목록, 오른쪽 월간 일정."""
from __future__ import annotations

import calendar
import html
import json
import os
import re
import sys
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from typing import Any

import pandas as pd
import streamlit as st

VC_DIR = os.path.join("uploaded_cache", "visit_calendar")
VC_STORE = os.path.join(VC_DIR, "store.json")
_WEEKDAYS = ("월", "화", "수", "목", "금", "토", "일")
_CAL_HEADERS = ("일", "월", "화", "수", "목", "금", "토")
_CAL_FIRST = calendar.SUNDAY
_CYLINDER = "실린더"
_VC_SAT_BG, _VC_SAT_FG = "#eef4fc", "#3b6fd8"
_VC_SUN_BG, _VC_SUN_FG = "#fff3f1", "#d23b3b"
_MCAL_SLOTS = 5


def _vc_is_touch_ui() -> bool:
    """iPad/터치. 맥·데스크톱 브라우저는 False — 레이아웃을 바꾸지 않는다."""
    try:
        if st.session_state.get("force_touch_ui") is True:
            return True
        v = st.query_params.get("touch_ui", "")
        if isinstance(v, list):
            v = v[0] if v else ""
        if str(v) == "1":
            st.session_state["force_touch_ui"] = True
            return True
    except Exception:
        pass
    try:
        cookies = getattr(st.context, "cookies", None)
        if cookies is not None and str(cookies.get("dashboard_touch", "")) == "1":
            st.session_state["force_touch_ui"] = True
            return True
    except Exception:
        pass
    try:
        headers = getattr(st.context, "headers", None)
        ua = ""
        if headers is not None:
            ua = str(headers.get("User-Agent") or headers.get("user-agent") or "")
        if ua and (
            re.search(r"iPad|iPhone|iPod", ua)
            or ("Macintosh" in ua and "Mobile" in ua)
        ):
            st.session_state["force_touch_ui"] = True
            return True
        plat = ""
        mobile = ""
        if headers is not None:
            plat = str(
                headers.get("Sec-CH-UA-Platform")
                or headers.get("sec-ch-ua-platform")
                or ""
            )
            mobile = str(
                headers.get("Sec-CH-UA-Mobile")
                or headers.get("sec-ch-ua-mobile")
                or ""
            )
        if "ipad" in plat.lower() or "ios" in plat.lower() or mobile.strip() in {"?1", "1", "true"}:
            st.session_state["force_touch_ui"] = True
            return True
    except Exception:
        pass
    return bool(st.session_state.get("force_touch_ui"))


def _vc_is_mac_local() -> bool:
    """맥 로컬과 Cloud 방문 달력은 같다. 가벼운 경로는 쓰지 않는다."""
    if _vc_is_streamlit_cloud():
        return True
    return not _vc_is_touch_ui()


def _vc_is_streamlit_cloud() -> bool:
    """Cloud만 True. 로컬 맥은 False. app.py _is_streamlit_cloud 와 같은 신호."""
    try:
        if (os.environ.get("STREAMLIT_RUNTIME_ENVIRONMENT") or "").strip().lower() == "cloud":
            return True
    except Exception:
        pass
    for _k in ("STREAMLIT_CLOUD", "IS_STREAMLIT_CLOUD"):
        try:
            _v = (os.environ.get(_k) or "").strip().lower()
            if _v in ("1", "true", "yes"):
                return True
        except Exception:
            pass
    try:
        if os.path.isdir("/mount/src"):
            return True
        cwd = os.path.abspath(os.getcwd())
        if cwd.startswith("/mount/src"):
            return True
    except Exception:
        pass
    try:
        if (os.environ.get("HOME") or "").rstrip("/") == "/home/adminuser":
            return True
    except Exception:
        pass
    try:
        headers = getattr(getattr(st, "context", None), "headers", None)
        host = ""
        if headers is not None:
            host = str(headers.get("host") or headers.get("Host") or "")
        if "streamlit.app" in host.lower():
            return True
    except Exception:
        pass
    return False


def _vc_is_darwin_local() -> bool:
    """납품줄 v2는 맥 Desktop 프로세스만. Cloud에 올리면 방문 탭이 끝나지 않는다."""
    try:
        if sys.platform != "darwin":
            return False
    except Exception:
        return False
    if _vc_is_streamlit_cloud():
        return False
    return True


def _vc_use_mcal_date_input() -> bool:
    """맥 로컬·Cloud 데스크톱은 date_input. 아이패드 date_input은 끝나지 않는다."""
    if _vc_is_darwin_local():
        return True
    return _vc_is_streamlit_cloud() and not _vc_is_touch_ui()


def _vc_use_ipad_layout() -> bool:
    """아이패드 전폭 레이아웃. Cloud에는 쓰지 않는다 — 로컬과 같은 2열."""
    if _vc_is_streamlit_cloud():
        return False
    return _vc_is_touch_ui()


def _vc_drop_touch_media(css: str) -> str:
    """Cloud·맥 데스크톱에는 아이패드 @media를 넣지 않는다."""
    if _vc_use_ipad_layout():
        return css
    out: list[str] = []
    i = 0
    needle = "@media (hover: none) and (pointer: coarse)"
    while True:
        j = css.find(needle, i)
        if j < 0:
            out.append(css[i:])
            break
        out.append(css[i:j])
        brace = css.find("{", j)
        if brace < 0:
            break
        depth = 0
        k = brace
        while k < len(css):
            ch = css[k]
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    k += 1
                    break
            k += 1
        i = k
    return "".join(out)


_VC_STRIP = None
_VC_STRIP_READY = False


def _vc_use_day_strip_component() -> bool:
    """맥 로컬만 납품줄 v2. Cloud에는 iframe을 만들지 않는다."""
    if not _vc_is_darwin_local():
        return False
    return _vc_ensure_strip() is not None


def _vc_clear_strip_hosts_if_unused() -> None:
    """v2를 안 쓰는 세션에 남은 iframe 호스트 값이 있으면 Safari가 계속 기다린다."""
    if _vc_use_day_strip_component():
        return
    st.session_state.pop("vc_strip_host", None)
    st.session_state.pop("t2d_strip_host", None)


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
    try:
        st.session_state.pop("_vc_day_data", None)
    except Exception:
        pass


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


def month_cal_weeks(month: date) -> list[list[date]]:
    """월 달력 격자(일~토). 앞뒤 달 날짜도 칸을 채운다."""
    return calendar.Calendar(firstweekday=_CAL_FIRST).monthdatescalendar(month.year, month.month)


def schedule_cells(
    store: dict,
    days: list[date],
    staff: str = "",
    history: list[dict] | None = None,
    client: str = "",
) -> dict[str, list[dict]]:
    """칸별 표시. visit=방문, planned=방문예정, prior=기방문(업무일지)."""
    want = {d.isoformat() for d in days}
    staff_s = _s(staff)
    client_key = _company_key(client)
    out: dict[str, list[dict]] = {iso: [] for iso in want}
    seen: set[tuple[str, str, str]] = set()

    def _add(iso: str, kind: str, label: str, *, mine: bool = False) -> None:
        if iso not in want:
            return
        lab = _s(label) or kind
        sig = (iso, kind, lab)
        if sig in seen:
            return
        seen.add(sig)
        out[iso].append({"kind": kind, "label": lab, "mine": mine})

    for v in store.get("visits") or []:
        if not isinstance(v, dict):
            continue
        iso = _iso(v.get("date"))
        vs = _s(v.get("staff"))
        if staff_s and vs and vs != staff_s:
            continue
        stt = _s(v.get("status")) or "done"
        kind = "planned" if stt == "planned" else "visit"
        name = _s(v.get("client")) or ("방문예정" if kind == "planned" else "방문")
        _add(iso, kind, name, mine=bool(client_key and _company_key(name) == client_key))
    for h in history or []:
        if not isinstance(h, dict):
            continue
        if str(h.get("source") or "") != "업무일지":
            continue
        _add(_iso(h.get("date")), "prior", "기방문")
    return out


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
    n = int(len(df))
    my = (month.year, month.month) if isinstance(month, date) else None
    amt = 0.0
    if "매출액" in df.columns:
        try:
            amt = float(pd.to_numeric(df["매출액"], errors="coerce").sum())
        except Exception:
            amt = 0.0
    sig = (staff, client, my, n, round(amt, 2), "t2d3")
    box = _vc_state_dict("_t2d_row_cache")
    hit = box.get("rows")
    if box.get("sig") == sig and isinstance(hit, list):
        return list(hit)
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
    prices = (
        pd.to_numeric(work["단가"], errors="coerce")
        if "단가" in work.columns
        else pd.Series(float("nan"), index=work.index)
    )
    bulk = items.str.upper().str.contains("BULK", na=False) | items.str.contains("벌크", na=False)
    out: list[dict] = []
    for iso, item, qty, amt, price, is_bulk, cname, dropped in zip(
        isos.tolist(),
        items.tolist(),
        qtys.tolist(),
        amts.tolist(),
        prices.tolist(),
        bulk.tolist(),
        names.tolist(),
        skip.tolist(),
    ):
        if dropped or not iso or iso == "NaT":
            continue
        name = _s(item) or "납품"
        if name.lower() in {"nan", "none"}:
            name = "납품"
        q = 0.0 if pd.isna(qty) else float(qty)
        a = 0.0 if pd.isna(amt) else float(amt)
        if price is not None and not pd.isna(price) and float(price) != 0:
            unit = float(price)
        elif q:
            unit = a / q
        else:
            unit = None
        out.append(
            {
                "date": iso,
                "item": name,
                "qty": q,
                "amount": a,
                "unit_price": unit,
                "staff": staff,
                "client": _s(cname),
                "source": "납품",
                "bulk": bool(is_bulk),
            }
        )
    out.sort(key=lambda x: (str(x.get("date") or ""), str(x.get("item") or "")), reverse=True)
    box.clear()
    box["sig"] = sig
    box["rows"] = out
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


def _vc_day_payload(
    df: pd.DataFrame | None, staff: str, client: str, month: date
) -> tuple[dict, list[dict], list[dict]]:
    """담당·거래처가 같으면 월 이동에도 납품·일지 맵을 다시 만들지 않는다."""
    ck = (staff, client)
    box = _vc_state_dict("_vc_day_data")
    if box.get("ck") == ck and isinstance(box.get("store"), dict):
        return box["store"], list(box.get("deliveries") or []), list(box.get("history") or [])
    store = load_store()
    deliveries = _sales_delivery_rows(df, staff, client) if client else []
    history = _worklog_visit_dates(client) if client else []
    box.clear()
    box["ck"] = ck
    box["store"] = store
    box["deliveries"] = deliveries
    box["history"] = history
    return store, deliveries, history


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
        for week in month_cal_weeks(month)
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
    return _vc_drop_touch_media("""
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
    div[class*="st-key-vc_mcal_box"] {
      border: 1.5px solid #9aa8bc !important;
      border-radius: 12px !important;
      overflow: hidden !important;
      background: #fff !important;
      margin: 0 0 12px !important;
      padding: 0 !important;
    }
    /* vc-grid-stColumn-v3 */
    .vc-mcal {
      padding: 10px 10px 8px;
    }
    .vc-mcal-title {
      font-size: 16px; font-weight: 600; color: #3c4043;
      padding: 0 2px 6px; letter-spacing: -0.2px;
    }
    .vc-mcal-leg {
      display: flex; gap: 10px; flex-wrap: wrap; font-size: 11px;
      color: #5f6368; padding: 0 2px 8px;
    }
    .vc-mcal-leg i {
      display: inline-block; width: 8px; height: 8px; border-radius: 50%;
      margin-right: 4px; vertical-align: 1px;
    }
    .vc-mcal-wd {
      text-align: center; font-size: 10px; font-weight: 700;
      letter-spacing: -0.2px; padding: 2px 0 4px; color: #6b7280;
      background: transparent;
    }
    .vc-mcal-wd.sun { color: #d23b3b; }
    .vc-mcal-wd.sat { color: #3b6fd8; }
    div[class*="st-key-vc_mcal_box"] .stHorizontalBlock {
      gap: 6px !important; margin: 0 6px 6px !important;
    }
    div[class*="st-key-vc_mcal_box"] .stColumn {
      border: none !important;
      padding: 0 !important; min-width: 0 !important;
      background: transparent !important;
    }
    div[class*="st-key-vc_mcal_box"] .stHorizontalBlock:has(div[class*="st-key-vc_mcal_2"]) .stColumn {
      background: #fff !important;
      border: 1px solid #eceff3 !important;
      border-radius: 10px !important;
    }
    div[class*="st-key-vc_mcal_box"] .stHorizontalBlock:has(div[class*="st-key-vc_mcal_2"]) > .stColumn:nth-child(1) {
      background: #fff3f1 !important;
    }
    div[class*="st-key-vc_mcal_box"] .stHorizontalBlock:has(div[class*="st-key-vc_mcal_2"]) > .stColumn:nth-child(7) {
      background: #eef4fc !important;
    }
    .vc-mcal-slots, .vc-mcal-chips {
      display: flex; flex-direction: column; align-items: center; gap: 0;
      padding: 0 2px 4px; text-align: center;
    }
    .vc-mcal-slot, .vc-mcal-chip {
      display: block; width: 100%; font-size: 10px; font-weight: 700; line-height: 1.2;
      min-height: 1.22rem; padding: 1px 2px;
      border: none; border-radius: 0; box-shadow: none;
      text-align: center;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    .vc-mcal-slot.empty { background: transparent; color: transparent; }
    .vc-mcal-slot.visit, .vc-mcal-chip.visit { background: #e6f4ea; color: #137333; }
    .vc-mcal-slot.planned, .vc-mcal-chip.planned { background: #fff4e5; color: #c47d00; }
    .vc-mcal-slot.prior, .vc-mcal-chip.prior { background: #e8f0fe; color: #1a73e8; }
    .vc-mcal-slot.mine, .vc-mcal-chip.mine { box-shadow: none; }
    .vc-mcal-more { font-size: 10px; color: #80868b; text-align: center; }
    div[class*="st-key-vc_mcal_2"] { margin: 0 !important; }
    div[class*="st-key-vc_mcal_2"] button {
      min-height: 8.2rem !important; height: auto !important;
      padding: 6px 2px 8px !important; border-radius: 10px !important;
      font-size: 12px !important; font-weight: 600 !important;
      background: #fff !important; color: #3c4043 !important;
      border: 1px solid #eceff3 !important; box-shadow: none !important;
      justify-content: center !important; text-align: center !important;
      white-space: pre-line !important; line-height: 1.25 !important;
    }
    div[class*="st-key-vc_mcal_2"] button:disabled {
      background: #fafbfc !important; color: #c4c7cc !important; border: none !important;
    }
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
    /* 아이패드·터치만. 맥 hover 레이아웃은 그대로 */
    @media (hover: none) and (pointer: coarse) {
      div[class*="st-key-vc_jump_today"] button,
      div[class*="st-key-vc_prev_month"] button,
      div[class*="st-key-vc_next_month"] button {
        min-height: 2.6rem !important; height: 2.6rem !important;
        min-width: 2.6rem !important; font-size: 15px !important;
      }
      div[class*="st-key-vc_strip_"] button {
        min-height: 3.1rem !important; height: 3.1rem !important;
        font-size: 12px !important;
      }
      div[class*="st-key-vc_mcal_2"] button {
        min-height: 8.2rem !important; height: auto !important;
        font-size: 14px !important;
        touch-action: manipulation;
      }
      div[class*="st-key-vc_visit_chk_"] button,
      div[class*="st-key-vc_plan_chk_"] button,
      div[class*="st-key-vc_visit_del_"] button {
        min-height: 2.7rem !important; height: 2.7rem !important;
        font-size: 15px !important; padding: 0 14px !important;
        touch-action: manipulation;
      }
      div[class*="st-key-vc_todo_chk_"] button,
      div[class*="st-key-vc_todo_star_"] button,
      div[class*="st-key-vc_todo_x_"] button {
        min-height: 2.6rem !important; height: 2.6rem !important;
      }
      .vc-mcal-slot, .vc-mcal-chip { font-size: 11px; }
      .vc-month-col { -webkit-overflow-scrolling: touch; }
    }
    /* iPad Mini 7 세로(CSS 744×1133) — 칸은 전폭, 납품줄은 가로 스크롤 */
    @media (hover: none) and (pointer: coarse) and (max-width: 850px) and (orientation: portrait) {
      .vc-toolbar .vc-title { font-size: 17px; line-height: 28px; }
      div[class*="st-key-vc_mcal_2"] button {
        min-height: 7.2rem !important;
        font-size: 13px !important;
      }
      .vc-mcal-slot, .vc-mcal-chip { font-size: 10px; }
      div[class*="st-key-vc_mcal_box"] .stHorizontalBlock { gap: 4px !important; margin: 0 4px 4px !important; }
    }
    /* iPad Mini 7 가로(CSS 1133×744) — 높이만 줄이고 7칸 전폭 유지 */
    @media (hover: none) and (pointer: coarse) and (min-width: 851px) and (max-width: 1180px) and (orientation: landscape) {
      .vc-toolbar { padding: 0 0 2px; min-height: 26px; }
      .vc-toolbar .vc-title { font-size: 18px; line-height: 26px; }
      div[class*="st-key-vc_jump_today"] button,
      div[class*="st-key-vc_prev_month"] button,
      div[class*="st-key-vc_next_month"] button {
        min-height: 2.3rem !important; height: 2.3rem !important;
      }
      div[class*="st-key-vc_mcal_2"] button {
        min-height: 6.6rem !important;
        padding: 4px 2px 6px !important;
      }
      div[class*="st-key-vc_mcal_box"] .stHorizontalBlock { gap: 4px !important; margin: 0 4px 4px !important; }
    }
    </style>
    """)


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
    "vc_mcal_",
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


def _apply_pending_mcal_pick() -> None:
    """스케줄 달력 클릭을 purge 전에 반영한다. 지우면 일자 이동이 사라진다."""
    for k in list(st.session_state.keys()):
        if not (isinstance(k, str) and k.startswith("vc_mcal_") and len(k) > 8 and k[8].isdigit()):
            continue
        if st.session_state.get(k) is not True:
            continue
        try:
            _on_pick_day(date.fromisoformat(k[8:18]))
        except ValueError:
            pass
        st.session_state.pop(k, None)


def _apply_pending_month_nav() -> None:
    """월 이동 클릭을 purge 전에 반영한다. 전체 재실행이면 로딩이 길어진다."""
    if st.session_state.get("vc_prev_month") is True:
        _on_shift_month(-1)
        st.session_state.pop("vc_prev_month", None)
    elif st.session_state.get("vc_next_month") is True:
        _on_shift_month(1)
        st.session_state.pop("vc_next_month", None)
    elif st.session_state.get("vc_jump_today") is True:
        _on_jump_today()
        st.session_state.pop("vc_jump_today", None)


_VC_KEEP_KEYS = frozenset({"vc_strip_host", "vc_mcal_date", "vc_mcal_box", "vc_mcal_iso"})


def _is_vc_purge_key(k: str) -> bool:
    """일자 버튼 잔여만 지운다. 스케줄 날짜 위젯·컨테이너는 지우면 재실행이 안 끝난다."""
    if k in _VC_KEEP_KEYS:
        return False
    if k.startswith("vc_mcal_"):
        return len(k) > 8 and k[8].isdigit()
    return k.startswith(_VC_BUTTON_PREFIXES)


def _purge_vc_button_keys() -> None:
    """버튼 값은 session_state로 넣을 수 없다. 예전 백업·클릭 잔여를 지운다."""
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and _is_vc_purge_key(k):
            st.session_state.pop(k, None)
    bak = st.session_state.get("_dash_bak_visit")
    if isinstance(bak, dict):
        for k in list(bak):
            if isinstance(k, str) and _is_vc_purge_key(k):
                bak.pop(k, None)


def _as_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _vc_date_field(label: str, *, key: str, default: date):
    """맥·Cloud 데스크톱은 date_input. 아이패드는 네이티브 달력이 핸드셰이크를 안 끝내서 selectbox."""
    if key not in st.session_state:
        st.session_state[key] = default
    raw = st.session_state.get(key)
    coerced = _as_date(raw)
    if coerced is not None and raw != coerced:
        st.session_state[key] = coerced
    if _vc_use_mcal_date_input():
        return st.date_input(label, format="YYYY/MM/DD", key=key)
    base = coerced or default
    opts = [base + timedelta(days=i) for i in range(-14, 61)]
    if base not in opts:
        opts = [base] + opts
    return st.selectbox(label, options=opts, key=key)


def _on_pick_day(d: date) -> None:
    """일자만 고른다. 월은 ‹ › · 오늘에서만 바꾼다."""
    iso = d.isoformat()
    st.session_state["_vc_selected"] = d
    st.session_state["_vc_open_delivery"] = True
    if st.session_state.get("vc_mcal_date") != d:
        st.session_state["vc_mcal_date"] = d
    if st.session_state.get("vc_mcal_iso") != iso:
        st.session_state["vc_mcal_iso"] = iso


def _on_mcal_date_change() -> None:
    d = _as_date(st.session_state.get("vc_mcal_date"))
    if d is None:
        return
    st.session_state["_vc_selected"] = d
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
        clamped = date(new.year, new.month, min(sel.day, last))
        st.session_state["_vc_selected"] = clamped
        st.session_state["vc_mcal_date"] = clamped
        st.session_state["vc_mcal_iso"] = clamped.isoformat()


def _on_jump_today() -> None:
    today = date.today()
    st.session_state["_vc_month"] = date(today.year, today.month, 1)
    st.session_state["_vc_selected"] = today
    st.session_state["vc_mcal_date"] = today
    st.session_state["vc_mcal_iso"] = today.isoformat()


def _sync_selected_from_mcal_widget() -> None:
    """스케줄 날짜 위젯 → 선택일. 위젯 키는 여기서 쓰지 않는다(재실행 루프 방지)."""
    if _vc_use_mcal_date_input():
        d = _as_date(st.session_state.get("vc_mcal_date"))
    else:
        raw = st.session_state.get("vc_mcal_iso")
        if not isinstance(raw, str) or len(raw) < 10:
            return
        try:
            d = date.fromisoformat(raw[:10])
        except ValueError:
            return
        mon = st.session_state.get("_vc_month")
        if isinstance(mon, date) and (d.year, d.month) != (mon.year, mon.month):
            return
    if d is None:
        return
    if st.session_state.get("_vc_selected") != d:
        st.session_state["_vc_selected"] = d
        st.session_state["_vc_open_delivery"] = True


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
        if _vc_is_touch_ui():
            return
        st.rerun()


def _apply_query_day_pick() -> None:
    try:
        raw = st.query_params.get("vc_pick", "")
        if isinstance(raw, list):
            raw = raw[0] if raw else ""
        raw = str(raw or "").strip()
        if not raw:
            return
        d = date.fromisoformat(raw[:10])
    except ValueError:
        return
    _on_pick_day(d)
    try:
        del st.query_params["vc_pick"]
    except Exception:
        pass


def _vc_pick_href(iso: str) -> str:
    bits = [f"vc_pick={iso}"]
    try:
        tu = st.query_params.get("touch_ui", "")
        if isinstance(tu, list):
            tu = tu[0] if tu else ""
        if str(tu) == "1":
            bits.append("touch_ui=1")
    except Exception:
        pass
    return "?" + "&".join(bits)


def render_visit_calendar_tab(df: pd.DataFrame | None = None, latest_update_str: str = "") -> None:
    """방문 미팅 캘린더 + 할일 목록."""
    st.markdown(
        "<div class='sub-header dashboard-tab-panel-head'>📅 방문·할일</div>",
        unsafe_allow_html=True,
    )
    _apply_query_day_pick()
    _vc_clear_strip_hosts_if_unused()
    _apply_pending_month_nav()
    _apply_pending_mcal_pick()
    _purge_vc_button_keys()
    _render_visit_body(df, latest_update_str)


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
        if prev:
            st.session_state["vc_client"] = prev

    month = st.session_state.get("_vc_month") or date(today.year, today.month, 1)
    st.markdown(_vc_css(), unsafe_allow_html=True)

    def _visit_day_block() -> None:
        _apply_pending_month_nav()
        _apply_pending_mcal_pick()
        _sync_selected_from_mcal_widget()
        sel: date = st.session_state.get("_vc_selected") or today
        mon = st.session_state.get("_vc_month") or month
        f1, f2 = st.columns([1, 1])
        with f1:
            staff = st.selectbox("담당자", options=staffs or [""], key="vc_staff")
        clients = _staff_clients(df, staff, store0)
        with f2:
            if _vc_is_mac_local():
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
            else:
                q = st.text_input(
                    "거래처 검색",
                    key="vc_client_q",
                    placeholder="거래처명 입력",
                )
                qq = _s(q).casefold()
                cur = _s(st.session_state.get("vc_client"))
                opts = [
                    c
                    for c in (clients or [])
                    if qq and qq in c.casefold()
                ][:40]
                if not qq:
                    opts = [cur] if cur else [""]
                if cur and cur not in opts:
                    opts = [cur] + [x for x in opts if x != cur]
                client = (
                    st.selectbox(
                        "거래처",
                        options=opts or [""],
                        key="vc_client",
                        help="검색어를 넣으면 담당자 매출 거래처가 나옵니다.",
                    )
                    or ""
                )
        st.session_state["_vc_staff"] = staff
        st.session_state["_vc_client"] = client
        stf = str(staff or "")
        cli = str(client or "")
        head_l, head_r = st.columns([2.35, 1.05])
        with head_l:
            st.markdown(
                f"<div class='vc-toolbar'><span class='vc-title'>{mon.year}년 {mon.month}월</span></div>",
                unsafe_allow_html=True,
            )
        with head_r:
            n1, n2, n3 = st.columns([1.15, 0.42, 0.42], gap="small")
            with n1:
                st.button("오늘", key="vc_jump_today", width="content")
            with n2:
                st.button("‹", key="vc_prev_month", width="content")
            with n3:
                st.button("›", key="vc_next_month", width="content")
        store, deliveries, history = _vc_day_payload(df, stf, cli, mon)
        prefix = f"{mon.year:04d}-{mon.month:02d}-"
        month_deliveries = [r for r in deliveries if str(r.get("date") or "").startswith(prefix)]
        chips = delivery_chips(mon, month_deliveries)
        day_deliveries = delivery_rows_on(deliveries, sel)
        _render_day_strip(mon, sel, chips, today)
        st.caption("위 줄은 납품(벌크·실린더)만 표시합니다. 방문은 오른쪽 월간 스케줄에서 보세요.")

        def _vc_delivery_col() -> None:
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
            _render_visit_log(store, stf, cli)

        if not _vc_use_ipad_layout():
            left, right = st.columns([1, 1], gap="medium")
            with left:
                _vc_delivery_col()
                _render_todo_panel(sel, stf, cli, store)
            with right:
                _render_month_cal(mon, sel, store, stf, cli, history, today)
                _render_day_agenda(sel, store, stf, cli)
                _render_month_schedule(mon, store, stf, sel, cli)
        else:
            # 아이패드: 월간 달력을 전폭으로 — 반쪽 7칸은 펜슬·손가락으로 누를 수 없다.
            _render_month_cal(mon, sel, store, stf, cli, history, today)
            _render_day_agenda(sel, store, stf, cli)
            _vc_delivery_col()
            _render_month_schedule(mon, store, stf, sel, cli)
            _render_todo_panel(sel, stf, cli, store)

        if latest_update_str:
            st.caption(f"대시보드 기준 시각: {latest_update_str}")

    if _vc_is_mac_local():
        st.fragment(_visit_day_block)()
    else:
        # 아이패드: 날짜 조작이 전 탭을 다시 돌리지 않게 fragment만 다시 그린다.
        st.fragment(_visit_day_block)()


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
@media (hover: none) and (pointer: coarse) and (orientation: portrait) {{
  .vc-strip-root {{ overflow-x:auto; -webkit-overflow-scrolling:touch; }}
  .vc-strip-wd, .vc-strip-days, .vc-strip-nm {{ width:max-content; min-width:100%; }}
  .vc-strip-days button {{ flex:0 0 2.2rem; min-width:2.2rem; }}
  .vc-strip-wd span, .vc-strip-nm span {{ flex:0 0 2.2rem; min-width:2.2rem; }}
}}
@media (hover: none) and (pointer: coarse) and (orientation: landscape) {{
  .vc-strip-days button {{ min-height:2.4rem; height:2.4rem; }}
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
def _vc_ensure_strip():
    """맥 로컬 첫 사용 때만 v2를 등록한다. Cloud에는 iframe을 만들지 않는다."""
    global _VC_STRIP, _VC_STRIP_READY
    if _VC_STRIP_READY:
        return _VC_STRIP
    _VC_STRIP_READY = True
    if not _vc_is_darwin_local():
        _VC_STRIP = None
        return None
    try:
        _VC_STRIP = st.components.v2.component(
            "visit_day_strip_v1",
            html=_VC_STRIP_HTML,
            css=_VC_STRIP_CSS,
            js=_VC_STRIP_JS,
        )
    except Exception:  # pragma: no cover
        _VC_STRIP = None
    return _VC_STRIP


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
    if _vc_use_day_strip_component():
        _VC_STRIP(
            key="vc_strip_host",
            data={"days": days, "selected": selected.isoformat(), "today": today.isoformat()},
            default={"iso": selected.isoformat()},
            on_iso_change=_on_strip_iso_change,
        )
        return
    if _vc_is_darwin_local():
        _render_strip_buttons(days, "vc_strip_", _on_pick_day)
        return
    st.markdown(_strip_fallback_html(days, selected.isoformat()), unsafe_allow_html=True)


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


def _mcal_grid_css() -> str:
    """달력 fragment 안에 넣는다. 탭 바깥 CSS는 새로고침 전엔 안 바뀐다."""
    return _vc_drop_touch_media("""
    <style>
    /* vc-grid-stColumn-v3 */
    div[class*="st-key-vc_mcal_box"] .stHorizontalBlock { gap: 6px !important; margin: 0 6px 6px !important; }
    div[class*="st-key-vc_mcal_box"] .stColumn {
      border: none !important; padding: 0 !important; min-width: 0 !important;
      background: transparent !important;
    }
    div[class*="st-key-vc_mcal_box"] .stHorizontalBlock:has(div[class*="st-key-vc_mcal_2"]) .stColumn {
      background: #fff !important;
      border: 1px solid #eceff3 !important;
      border-radius: 10px !important;
    }
    div[class*="st-key-vc_mcal_box"] .stHorizontalBlock:has(div[class*="st-key-vc_mcal_2"]) > .stColumn:nth-child(1) {
      background: #fff3f1 !important;
    }
    div[class*="st-key-vc_mcal_box"] .stHorizontalBlock:has(div[class*="st-key-vc_mcal_2"]) > .stColumn:nth-child(7) {
      background: #eef4fc !important;
    }
    .vc-mcal-wd {
      text-align: center; font-size: 10px; font-weight: 700;
      letter-spacing: -0.2px; padding: 2px 0 4px; color: #6b7280;
      background: transparent;
    }
    .vc-mcal-wd.sun { color: #d23b3b; }
    .vc-mcal-wd.sat { color: #3b6fd8; }
    div[class*="st-key-vc_mcal_2"] { margin: 0 !important; }
    div[class*="st-key-vc_mcal_2"] button {
      min-height: 8.2rem !important; height: auto !important;
      padding: 6px 2px 8px !important; border-radius: 10px !important;
      white-space: pre-line !important; line-height: 1.25 !important;
      justify-content: center !important; text-align: center !important;
      cursor: pointer !important;
    }
    .vc-mcal-slots, .vc-mcal-chips { text-align: center; align-items: center; }
    .vc-mcal-slot, .vc-mcal-chip {
      border: none !important; border-radius: 0 !important; box-shadow: none !important;
      text-align: center !important;
    }
    @media (hover: none) and (pointer: coarse) and (max-width: 850px) and (orientation: portrait) {
      div[class*="st-key-vc_mcal_2"] button { min-height: 7.2rem !important; font-size: 13px !important; }
      div[class*="st-key-vc_mcal_box"] .stHorizontalBlock { gap: 4px !important; margin: 0 4px 4px !important; }
    }
    @media (hover: none) and (pointer: coarse) and (min-width: 851px) and (max-width: 1180px) and (orientation: landscape) {
      div[class*="st-key-vc_mcal_2"] button { min-height: 6.6rem !important; padding: 4px 2px 6px !important; }
      div[class*="st-key-vc_mcal_box"] .stHorizontalBlock { gap: 4px !important; margin: 0 4px 4px !important; }
      .vc-mcal-table td { min-height: 6.6rem; padding: 4px 2px 6px; }
    }
    .vc-mcal-head {
      display: grid; grid-template-columns: repeat(7, 1fr); gap: 6px;
      margin: 0 6px 4px;
    }
    .vc-mcal-th {
      text-align: center; font-size: 10px; font-weight: 700;
      letter-spacing: -0.2px; padding: 2px 0 4px; color: #6b7280;
    }
    .vc-mcal-th.sun { color: #d23b3b; }
    .vc-mcal-th.sat { color: #3b6fd8; }
    .vc-mcal-table {
      width: 100%; border-collapse: separate; border-spacing: 6px;
      table-layout: fixed; margin: 0 0 8px;
    }
    .vc-mcal-table td {
      background: #fff; border: 1px solid #eceff3; border-radius: 10px;
      vertical-align: top; padding: 6px 2px 8px; min-height: 8.2rem;
      text-align: center;
    }
    .vc-mcal-table td.sun { background: #fff3f1; }
    .vc-mcal-table td.sat { background: #eef4fc; }
    .vc-mcal-table td.sel { background: #e8eaed; }
    .vc-mcal-table td.out { color: #80868b; }
    .vc-mcal-table td.today { box-shadow: inset 0 0 0 1.5px #9aa8bc; }
    .vc-mcal-table td .vc-mcal-hit,
    .vc-mcal-table td a.vc-mcal-hit {
      color: inherit; text-decoration: none; display: block; min-height: 8.2rem;
      cursor: pointer; -webkit-tap-highlight-color: rgba(26,115,232,0.25);
    }
    .vc-mcal-num { font-size: 12px; font-weight: 600; color: #3c4043; padding-bottom: 4px; }
    .vc-strip-html { width: 100%; margin: 0 0 8px; }
    .vc-strip-html .vc-strip-wd,
    .vc-strip-html .vc-strip-days,
    .vc-strip-html .vc-strip-nm {
      display: flex; gap: 3px; width: 100%;
    }
    .vc-strip-html .vc-strip-wd { margin-bottom: 2px; }
    .vc-strip-html .vc-strip-nm { margin-top: 2px; }
    .vc-strip-html .vc-strip-wd span {
      flex: 1; min-width: 0; text-align: center; font-size: 10px; font-weight: 700;
      line-height: 1.1;
    }
    .vc-strip-html .vc-strip-days span,
    .vc-strip-html .vc-strip-days a {
      flex: 1; min-width: 0; min-height: 2.7rem; padding: 2px 0;
      border-radius: 10px; font-size: 10px; font-weight: 600; background: #fff;
      color: #3c4043; border: 1px solid #eceff3; text-align: center;
      white-space: pre-line; line-height: 1.15; display: flex;
      align-items: center; justify-content: center;
      text-decoration: none; cursor: pointer;
      -webkit-tap-highlight-color: rgba(26,115,232,0.25);
    }
    .vc-strip-html .vc-strip-days span.mix,
    .vc-strip-html .vc-strip-days a.mix { background: #ece6f8; color: #5b4b8a; }
    .vc-strip-html .vc-strip-days span.bulk,
    .vc-strip-html .vc-strip-days a.bulk { background: #e8f0fe; color: #1a73e8; }
    .vc-strip-html .vc-strip-days span.other,
    .vc-strip-html .vc-strip-days a.other { background: #f6eee4; color: #8d6e63; }
    .vc-strip-html .vc-strip-days span.sat,
    .vc-strip-html .vc-strip-days a.sat { background: #eef4fc; color: #3b6fd8; }
    .vc-strip-html .vc-strip-days span.sun,
    .vc-strip-html .vc-strip-days a.sun { background: #fff3f1; color: #d23b3b; }
    .vc-strip-html .vc-strip-days span.today,
    .vc-strip-html .vc-strip-days a.today { box-shadow: inset 0 0 0 1.5px #1a73e8; }
    .vc-strip-html .vc-strip-days span.sel,
    .vc-strip-html .vc-strip-days a.sel {
      background: #1a73e8 !important; color: #fff !important; border-color: #1a73e8 !important;
    }
    .vc-strip-html .vc-strip-nm span {
      flex: 1; min-width: 0; text-align: center; font-size: 9px; font-weight: 700;
      color: #137333; line-height: 1.15; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    }
    @media (hover: none) and (pointer: coarse) and (orientation: portrait) {
      .vc-strip-html { overflow-x: auto; -webkit-overflow-scrolling: touch; }
      .vc-strip-html .vc-strip-wd,
      .vc-strip-html .vc-strip-days,
      .vc-strip-html .vc-strip-nm { width: max-content; min-width: 100%; }
      .vc-strip-html .vc-strip-days span,
      .vc-strip-html .vc-strip-days a { flex: 0 0 2.2rem; min-width: 2.2rem; }
      .vc-strip-html .vc-strip-wd span, .vc-strip-html .vc-strip-nm span { flex: 0 0 2.2rem; min-width: 2.2rem; }
    }
    </style>
    """)


def _mcal_chips_html(marks: list[dict]) -> str:
    """일자 칸에 값 5줄을 항상 그린다. 비어 있으면 빈 칸."""
    tags = {"visit": "방문", "planned": "예정", "prior": "기방문"}
    parts: list[str] = []
    extra = max(0, len(marks) - _MCAL_SLOTS)
    for i in range(_MCAL_SLOTS):
        if i < len(marks):
            m = marks[i]
            kind = str(m.get("kind") or "visit")
            raw = _s(m.get("label")) or tags.get(kind, "")
            lab = html.escape(_client_short(raw) if kind != "prior" else raw)
            mine = " mine" if m.get("mine") else ""
            more = f" +{extra}" if i == _MCAL_SLOTS - 1 and extra else ""
            parts.append(
                f"<span class='vc-mcal-slot {html.escape(kind)}{mine}'>{lab}{more}</span>"
            )
        else:
            parts.append("<span class='vc-mcal-slot empty'></span>")
    return "<div class='vc-mcal-slots'>" + "".join(parts) + "</div>"


def _mcal_button_label(d: date, marks: list[dict]) -> str:
    """일자+값 5줄을 한 버튼에 넣어 칸 전체를 누를 수 있게 한다."""
    tags = {"visit": "방문", "planned": "예정", "prior": "기방문"}
    lines = [str(d.day)]
    extra = max(0, len(marks) - _MCAL_SLOTS)
    for i in range(_MCAL_SLOTS):
        if i < len(marks):
            m = marks[i]
            kind = str(m.get("kind") or "visit")
            raw = _s(m.get("label")) or tags.get(kind, "")
            lab = _client_short(raw) if kind != "prior" else raw
            more = f" +{extra}" if i == _MCAL_SLOTS - 1 and extra else ""
            lines.append(f"{lab}{more}")
        else:
            lines.append("\u00a0")
    return "\n".join(lines)


def _mcal_head_html() -> str:
    ths: list[str] = []
    for i, wd in enumerate(_CAL_HEADERS):
        cls = " sun" if i == 0 else (" sat" if i == 6 else "")
        ths.append(f"<div class='vc-mcal-th{cls}'>{wd}</div>")
    return "<div class='vc-mcal-head'>" + "".join(ths) + "</div>"


def _mcal_month_html(
    month: date,
    selected: date,
    today: date,
    weeks: list[list[date]],
    cells: dict[str, list[dict]],
    *,
    pick_href: bool = False,
) -> str:
    sel_iso = selected.isoformat()
    today_iso = today.isoformat()
    rows: list[str] = []
    for week in weeks:
        tds: list[str] = []
        for i, d in enumerate(week):
            iso = d.isoformat()
            out = d.month != month.month
            cls = ["sun" if i == 0 else "sat" if i == 6 else ""]
            if out:
                cls.append("out")
            elif iso == sel_iso:
                cls.append("sel")
            if not out and iso == today_iso:
                cls.append("today")
            klass = " ".join(x for x in cls if x)
            body = (
                f"<div class='vc-mcal-num'>{d.day}</div>"
                f"{_mcal_chips_html(cells.get(iso) or [])}"
            )
            if pick_href:
                href = html.escape(_vc_pick_href(iso), quote=True)
                hit = f"<a class='vc-mcal-hit' href='{href}'>{body}</a>"
            else:
                hit = f"<div class='vc-mcal-hit'>{body}</div>"
            tds.append(f"<td class='{klass}'>{hit}</td>")
        rows.append("<tr>" + "".join(tds) + "</tr>")
    return "<table class='vc-mcal-table'><tbody>" + "".join(rows) + "</tbody></table>"


def _strip_fallback_html(days: list[dict], selected_iso: str) -> str:
    wd = "".join(
        f'<span style="color:{html.escape(str(d.get("wdc") or "#6b7280"))}">{html.escape(str(d.get("wd") or ""))}</span>'
        for d in days
    )
    cells = []
    names = []
    for d in days:
        iso = str(d.get("iso") or "")
        cls = " ".join(
            x
            for x in (
                str(d.get("kind") or ""),
                "sel" if iso == selected_iso else "",
                "today" if d.get("today") else "",
            )
            if x
        )
        tag = str(d.get("tag") or "")
        lab = f"{d.get('day')}<br>{html.escape(tag)}" if tag else str(d.get("day") or "")
        cells.append(f'<span class="{html.escape(cls)}">{lab}</span>')
        nm = str(d.get("name") or "")
        mcls = "planned" if d.get("mark") == "planned" else ""
        names.append(f'<span class="{mcls}">{html.escape(nm) if nm else "&nbsp;"}</span>')
    return (
        '<div class="vc-strip-html">'
        f'<div class="vc-strip-wd">{wd}</div>'
        f'<div class="vc-strip-days">{"".join(cells)}</div>'
        f'<div class="vc-strip-nm">{"".join(names)}</div>'
        "</div>"
    )


def _render_month_cal(
    month: date,
    selected: date,
    store: dict,
    staff: str,
    client: str,
    history: list[dict],
    today: date,
) -> None:
    weeks = month_cal_weeks(month)
    days = [d for w in weeks for d in w]
    cells = schedule_cells(store, days, staff, history, client)
    n_visit = sum(1 for items in cells.values() for x in items if x.get("kind") == "visit")
    n_plan = sum(1 for items in cells.values() for x in items if x.get("kind") == "planned")
    n_prior = sum(1 for items in cells.values() for x in items if x.get("kind") == "prior")
    st.markdown(_mcal_grid_css(), unsafe_allow_html=True)
    with st.container(key="vc_mcal_box"):
        st.markdown(
            f"<div class='vc-mcal'>"
            f"<div class='vc-mcal-title'>{month.year}년 {month.month}월 스케줄</div>"
            f"<div class='vc-mcal-leg'>"
            f"<span><i style='background:#137333'></i>방문 {n_visit}</span>"
            f"<span><i style='background:#c47d00'></i>방문예정 {n_plan}</span>"
            f"<span><i style='background:#1a73e8'></i>기방문 {n_prior}</span>"
            f"</div></div>",
            unsafe_allow_html=True,
        )
        raw = st.session_state.get("vc_mcal_date")
        coerced = _as_date(raw)
        if coerced is not None and raw != coerced:
            st.session_state["vc_mcal_date"] = coerced
        elif "vc_mcal_date" not in st.session_state:
            st.session_state["vc_mcal_date"] = selected
        if _vc_use_mcal_date_input():
            st.markdown(
                _mcal_head_html()
                + _mcal_month_html(
                    month, selected, today, weeks, cells,
                    pick_href=_vc_is_darwin_local(),
                ),
                unsafe_allow_html=True,
            )
            first, last = weeks[0][0], weeks[-1][-1]
            st.date_input(
                "스케줄 날짜",
                min_value=first,
                max_value=last,
                format="YYYY/MM/DD",
                key="vc_mcal_date",
            )
        else:
            st.markdown(
                _mcal_head_html()
                + _mcal_month_html(
                    month, selected, today, weeks, cells, pick_href=False
                ),
                unsafe_allow_html=True,
            )


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
    if not _vc_is_mac_local():
        with c1:
            st.markdown(
                _month_schedule_items_html(done, selected, client, "visit"),
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                _month_schedule_items_html(planned, selected, client, "planned"),
                unsafe_allow_html=True,
            )
        return
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
    wd = _WEEKDAYS[selected.weekday()]
    st.markdown(
        f"<div style='font-size:13px;font-weight:600;color:#5f6368;padding:2px 0 6px'>"
        f"선택한 날 · {selected.month}월 {selected.day}일 ({wd})</div>",
        unsafe_allow_html=True,
    )
    b1, b2, b3 = st.columns([0.85, 1.15, 0.65], gap="small")
    with b1:
        st.button(
            "방문",
            key=f"vc_visit_chk_{iso}",
            type="secondary",
            width="stretch",
            disabled=off,
            on_click=_on_toggle_visit,
            args=(selected, staff, client, "done"),
            help="실제 방문. 달력에 초록 업체명.",
        )
    with b2:
        st.button(
            "방문예정",
            key=f"vc_plan_chk_{iso}",
            type="secondary",
            width="stretch",
            disabled=off,
            on_click=_on_toggle_visit,
            args=(selected, staff, client, "planned"),
            help="방문 예정. 달력에 주황 예·업체명.",
        )
    with b3:
        st.button(
            "삭제",
            key=f"vc_visit_del_{iso}",
            type="secondary",
            width="stretch",
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
    if not _vc_is_mac_local():
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
            st.markdown(f"- **{d}{wd}** · {kind} · {v.get('client') or '-'}{note}")
        ids = [str(v.get("id") or "") for v in rows if str(v.get("id") or "")]
        labels = {
            str(v.get("id") or ""): f"{_iso(v.get('date'))} · {_s(v.get('client')) or '-'}"
            for v in rows
            if str(v.get("id") or "")
        }
        if ids:
            pick = st.selectbox(
                "삭제할 방문",
                options=[""] + ids,
                format_func=lambda i: labels.get(i, "선택") if i else "선택",
                key="vc_light_log_sel",
            )
            st.button(
                "선택한 방문 삭제",
                key="vc_light_log_del",
                disabled=not pick,
                on_click=_on_delete_visit_id,
                args=(str(pick),),
            )
        if extra:
            st.caption(f"최근 40건만 표시 · 나머지 {extra}건")
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
    show_unit_price: bool = False,
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
    if not _vc_is_mac_local():
        lines = [
            "|구분|품목|충전량|출고량|매출액|",
            "|---|---|---|---|---|",
        ]
        for r in ordered:
            bulk = bool(r.get("bulk") or _is_bulk_item(r.get("item") or ""))
            kind = "벌크" if bulk else _CYLINDER
            qty = _fmt_qty(r.get("qty"))
            fill = qty if bulk else ""
            out = "" if bulk else qty
            lines.append(
                f"|{kind}|{html.escape(_s(r.get('item')))}|{fill}|{out}|{_fmt_qty(r.get('amount'))}|"
            )
        st.markdown("\n".join(lines))
        return
    table = pd.DataFrame(
        [
            {
                "구분": "벌크" if (r.get("bulk") or _is_bulk_item(r.get("item") or "")) else _CYLINDER,
                "날짜": r.get("date"),
                "품목": r.get("item"),
                **(
                    {"단가": _fmt_qty(r.get("unit_price"))}
                    if show_unit_price
                    else {}
                ),
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
    elif not _vc_is_mac_local():
        labels: dict[str, str] = {}
        for t in todos:
            tid = str(t.get("id") or "")
            if not tid:
                continue
            plain_name = _s(t.get("title") or t.get("client")) or "할일"
            due_s = _iso(t.get("due"))
            mark = "★ " if t.get("starred") else ""
            done_m = "완료 · " if t.get("done") else ""
            extra = f" · {due_s[5:]}" if due_s else ""
            note = _s(t.get("note"))
            note_h = f" · {note}" if note else ""
            st.markdown(f"- {mark}{done_m}**{plain_name}**{extra}{note_h}")
            labels[tid] = f"{mark}{plain_name}{extra}"
        ids = list(labels)
        pick = st.selectbox(
            "할일 선택",
            options=[""] + ids,
            format_func=lambda i: labels.get(i, "선택") if i else "선택",
            key="vc_light_todo_sel",
        )
        a1, a2, a3, a4 = st.columns(4)
        with a1:
            st.button(
                "완료",
                key="vc_light_todo_done",
                disabled=not pick,
                on_click=_on_todo_done,
                args=(str(pick),),
            )
        with a2:
            st.button(
                "별표",
                key="vc_light_todo_star",
                disabled=not pick,
                on_click=_on_todo_star,
                args=(str(pick),),
            )
        with a3:
            st.button(
                "편집",
                key="vc_light_todo_edit",
                disabled=not pick,
                on_click=_on_pick_todo,
                args=(str(pick),),
            )
        with a4:
            st.button(
                "삭제",
                key="vc_light_todo_del",
                disabled=not pick,
                on_click=_on_todo_delete,
                args=(str(pick),),
            )
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
        due = _vc_date_field("날짜", key="vc_todo_due", default=due_default or selected)
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


def _apply_pending_t2d_strip_pick() -> None:
    """납품달력 클릭을 purge 전에 반영한다. 호스트 키는 지우지 않는다."""
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


def _apply_pending_t2d_month_nav() -> None:
    if st.session_state.get("t2d_prev_month") is True:
        _on_t2d_shift_month(-1)
        st.session_state.pop("t2d_prev_month", None)
    elif st.session_state.get("t2d_next_month") is True:
        _on_t2d_shift_month(1)
        st.session_state.pop("t2d_next_month", None)
    elif st.session_state.get("t2d_jump_today") is True:
        _on_t2d_jump_today()
        st.session_state.pop("t2d_jump_today", None)


def _purge_t2d_button_keys() -> None:
    for k in list(st.session_state.keys()):
        if not (isinstance(k, str) and k.startswith(_T2D_BUTTON_PREFIXES)):
            continue
        if k == "t2d_strip_host":
            continue
        st.session_state.pop(k, None)


def _on_t2d_pick_day(d: date) -> None:
    """일자만 고른다. 월은 ‹ › · 오늘에서만 바꾼다."""
    st.session_state["_t2d_selected"] = d


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


def _render_strip_buttons(days: list[dict], key_pfx: str, on_pick) -> None:
    cols = st.columns(len(days) or 1, gap="small")
    for i, cell in enumerate(days):
        d = date.fromisoformat(cell["iso"])
        with cols[i]:
            st.button(
                f"{cell['day']}\n{cell['tag']}" if cell["tag"] else str(cell["day"]),
                key=f"{key_pfx}{cell['iso']}",
                type="secondary",
                width="stretch",
                on_click=on_pick,
                args=(d,),
            )


def _render_t2d_day_strip(
    month: date, selected: date, chips: dict[str, list[dict]], today: date
) -> None:
    days = _strip_days_payload(month, chips, selected, today)
    if _vc_use_day_strip_component():
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
    _render_strip_buttons(days, "t2d_strip_", _on_t2d_pick_day)


def render_tab2_delivery_status(
    df: pd.DataFrame | None, staff: str, client: str
) -> None:
    """거래처 분석 상단 납품현황. 담당자·거래처는 메인 고정바 값을 쓴다."""
    _vc_clear_strip_hosts_if_unused()
    _apply_pending_t2d_month_nav()
    _apply_pending_t2d_strip_pick()
    _purge_t2d_button_keys()
    st.markdown(_t2d_css(), unsafe_allow_html=True)
    with st.container(key="tab2_delivery_status"):
        st.markdown("<div class='t2d-label'>납품현황</div>", unsafe_allow_html=True)
        staff = _s(staff)
        client = _s(client)

        def _t2d_body() -> None:
            _apply_pending_t2d_month_nav()
            _apply_pending_t2d_strip_pick()
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
                    st.button("오늘", key="t2d_jump_today", width="content")
                with n2:
                    st.button("‹", key="t2d_prev_month", width="content")
                with n3:
                    st.button("›", key="t2d_next_month", width="content")
            deliveries = _t2d_delivery_rows(df, staff, client, month=month)
            month_deliveries = deliveries
            chips = delivery_chips(month, month_deliveries)
            _render_t2d_day_strip(month, selected, chips, today)
            day_rows = delivery_rows_on(deliveries, selected)
            if day_rows:
                _render_delivery_list(
                    client,
                    selected,
                    day_rows,
                    deliveries,
                    heading="납품 내역",
                    show_unit_price=True,
                )
            else:
                st.markdown(
                    "<div class='t2d-empty' style='min-height:56px;margin-top:10px'>"
                    "이 날 납품 내역이 없습니다. 벌크·실린더 표시가 있는 날을 선택하세요.</div>",
                    unsafe_allow_html=True,
                )

        if _vc_is_mac_local():
            st.fragment(_t2d_body)()
        else:
            _t2d_body()

