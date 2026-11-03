"""구글 캘린더·할 일 — API로 월간 달력과 할 일 목록을 한 화면에 붙여 작업한다."""
from __future__ import annotations

import calendar
import html
import json
import os
import secrets
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, time, timedelta, timezone
from typing import Any
from zoneinfo import ZoneInfo

import requests
import streamlit as st

_ROOT = os.path.dirname(os.path.abspath(__file__))
_HOME_GC = os.path.join(os.path.expanduser("~"), ".office-dashboard")
_ST_GC = os.path.join(_ROOT, ".streamlit")
GC_DIR = os.path.join(_ROOT, "uploaded_cache", "google_calendar")
GC_TOKEN = os.path.join(GC_DIR, "token.json")
GC_OAUTH = os.path.join(GC_DIR, "oauth_client.json")
GC_TOKEN_HOME = os.path.join(_HOME_GC, "google_calendar_token.json")
GC_OAUTH_HOME = os.path.join(_HOME_GC, "google_oauth_client.json")
GC_TOKEN_ST = os.path.join(_ST_GC, "gc_token.json")
GC_OAUTH_ST = os.path.join(_ST_GC, "gc_oauth_client.json")
GC_TASKS = os.path.join(GC_DIR, "tasks_local.json")
GC_TASKS_HOME = os.path.join(_HOME_GC, "google_tasks.json")
GC_TASKS_ST = os.path.join(_ST_GC, "gc_tasks.json")
GC_KR_HOLIDAY = "ko.south_korea#holiday@group.v.calendar.google.com"
_TZ = ZoneInfo("Asia/Seoul")
_SCOPES = (
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/tasks",
)
_AUTH = "https://accounts.google.com/o/oauth2/v2/auth"
_TOKEN = "https://oauth2.googleapis.com/token"
_CAL_API = "https://www.googleapis.com/calendar/v3"
_TASK_API = "https://tasks.googleapis.com/tasks/v1"
_CLOUD_REDIRECT = "https://office-g8ryabkapprkpjmfwa5aypw.streamlit.app/"
_EV_MEM: dict[str, dict[str, list[dict]]] = {}


def calendar_embed_url(srcs: list[str] | None = None, *, height: int = 780) -> str:
    """하위 호환 — 공개 내장 URL. 본 탭은 API 화면을 쓴다."""
    items = [s.strip() for s in (srcs or [GC_KR_HOLIDAY]) if s and str(s).strip()]
    q = {
        "height": str(int(height)),
        "wkst": "1",
        "ctz": "Asia/Seoul",
        "hl": "ko",
        "mode": "MONTH",
        "showTitle": "0",
        "showPrint": "0",
        "showTabs": "1",
        "showCalendars": "1",
        "showTz": "0",
        "src": items[0],
    }
    extra = "".join(
        "&src=" + urllib.parse.quote(s, safe="") for s in items[1:]
    )
    return "https://calendar.google.com/calendar/embed?" + urllib.parse.urlencode(q) + extra


def month_weeks(year: int, month: int) -> list[list[date]]:
    """일요일 시작 주. Google 월간과 같다."""
    return calendar.Calendar(firstweekday=6).monthdatescalendar(year, month)


def iso_week_label(week: list[date]) -> int:
    return int(week[3].isocalendar()[1])


def _gc_is_holiday_day(items: list[dict] | None) -> bool:
    return any(str((ev or {}).get("cal_id") or "") == GC_KR_HOLIDAY for ev in (items or []))


def event_day_key(item: dict) -> str:
    start = (item or {}).get("start") or {}
    raw = str(start.get("date") or start.get("dateTime") or "")
    return raw[:10]


def _mkdir() -> None:
    for d in (GC_DIR, _HOME_GC, _ST_GC):
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass


def _token_paths() -> list[str]:
    return [GC_TOKEN_HOME, GC_TOKEN_ST, GC_TOKEN]


def _oauth_paths() -> list[str]:
    return [GC_OAUTH_HOME, GC_OAUTH_ST, GC_OAUTH]


def _read_json(path: str) -> dict:
    try:
        if not path or not os.path.exists(path):
            return {}
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _write_json(path: str, data: dict) -> None:
    try:
        parent = os.path.dirname(path)
        if parent:
            os.makedirs(parent, exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except Exception:
        pass


def _usable_token(data: dict) -> bool:
    return bool(data.get("refresh_token") or data.get("access_token"))


def _secret(name: str) -> str:
    try:
        secrets_obj = getattr(st, "secrets", None)
        if secrets_obj is not None:
            for key in (name, name.upper()):
                try:
                    v = secrets_obj.get(key) if hasattr(secrets_obj, "get") else secrets_obj[key]
                except Exception:
                    v = None
                if v is not None and str(v).strip():
                    return str(v).strip()
            for section in ("google_oauth", "google_calendar"):
                try:
                    nested = secrets_obj.get(section) if hasattr(secrets_obj, "get") else None
                except Exception:
                    nested = None
                if nested is not None:
                    try:
                        v = nested.get(name) if hasattr(nested, "get") else nested[name]
                    except Exception:
                        v = None
                    if v is not None and str(v).strip():
                        return str(v).strip()
    except Exception:
        pass
    return (os.environ.get(name) or os.environ.get(name.upper()) or "").strip()


def load_oauth_client() -> tuple[str, str]:
    cid = _secret("google_oauth_client_id") or _secret("client_id")
    csec = _secret("google_oauth_client_secret") or _secret("client_secret")
    data: dict = {}
    for path in _oauth_paths():
        data = _read_json(path)
        if data:
            break
    if data:
        for path in _oauth_paths():
            if not os.path.exists(path):
                _write_json(path, data)
    installed = data.get("web") or data.get("installed") or data
    if isinstance(installed, dict):
        cid = cid or str(installed.get("client_id") or "").strip()
        csec = csec or str(installed.get("client_secret") or "").strip()
    return cid, csec


def save_oauth_client(client_id: str, client_secret: str) -> None:
    payload = {
        "web": {
            "client_id": client_id.strip(),
            "client_secret": client_secret.strip(),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": _TOKEN,
        }
    }
    for path in _oauth_paths():
        _write_json(path, payload)


def load_token() -> dict:
    tok = st.session_state.get("_gc_token")
    if isinstance(tok, dict) and _usable_token(tok):
        return tok
    best: dict = {}
    for path in _token_paths():
        data = _read_json(path)
        if _usable_token(data):
            best = data
            break
    if best:
        st.session_state["_gc_token"] = best
        for path in _token_paths():
            if not os.path.exists(path):
                _write_json(path, best)
    return best


def save_token(tok: dict) -> None:
    if not isinstance(tok, dict):
        return
    tok = dict(tok)
    if not tok.get("refresh_token"):
        prev = st.session_state.get("_gc_token")
        if not (isinstance(prev, dict) and prev.get("refresh_token")):
            prev = {}
            for path in _token_paths():
                prev = _read_json(path)
                if prev.get("refresh_token"):
                    break
        if isinstance(prev, dict) and prev.get("refresh_token"):
            tok["refresh_token"] = prev["refresh_token"]
    for path in _token_paths():
        _write_json(path, tok)
    st.session_state["_gc_token"] = tok


def clear_token() -> None:
    st.session_state.pop("_gc_token", None)
    _EV_MEM.clear()
    for path in _token_paths():
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            pass


def redirect_uri() -> str:
    try:
        headers = getattr(st.context, "headers", None) or {}
        host = str(headers.get("Host") or headers.get("host") or "")
        if "streamlit.app" in host:
            return _CLOUD_REDIRECT
        if host:
            h = host.split(",")[0].strip()
            if h.startswith("localhost") or h.startswith("127.0.0.1"):
                return f"http://{h}/"
            return f"https://{h}/"
    except Exception:
        pass
    return "http://localhost:8501/"


def _redirect_candidates(preferred: str = "") -> list[str]:
    seen: list[str] = []
    for raw in (
        preferred,
        redirect_uri(),
        "http://127.0.0.1:8501/",
        "http://localhost:8501/",
        "http://127.0.0.1:8501",
        "http://localhost:8501",
        _CLOUD_REDIRECT,
    ):
        u = str(raw or "").strip()
        if u and u not in seen:
            seen.append(u)
        if u.endswith("/"):
            alt = u[:-1]
        elif u:
            alt = u + "/"
        else:
            alt = ""
        if alt and alt not in seen:
            seen.append(alt)
    return seen


def auth_url(client_id: str, state: str) -> str:
    ru = redirect_uri()
    try:
        st.session_state["_gc_oauth_redirect"] = ru
    except Exception:
        pass
    q = {
        "client_id": client_id,
        "redirect_uri": ru,
        "response_type": "code",
        "scope": " ".join(_SCOPES),
        "access_type": "offline",
        "include_granted_scopes": "true",
        "prompt": "consent",
        "state": state,
        "hl": "ko",
    }
    return _AUTH + "?" + urllib.parse.urlencode(q)


def _token_error_text(r: requests.Response) -> str:
    err, desc = "", ""
    try:
        data = r.json()
        err = str((data or {}).get("error") or "")
        desc = str((data or {}).get("error_description") or "")
    except Exception:
        pass
    if err == "redirect_uri_mismatch":
        return "리다이렉트 주소가 Google Cloud에 등록된 URI와 다릅니다."
    if err == "invalid_grant":
        return "인증이 만료됐거나 이미 사용되었습니다. 다시 연결해 주세요."
    if err == "invalid_client":
        return "OAuth 클라이언트 ID 또는 비밀번호가 올바르지 않습니다."
    if desc:
        return desc
    if err:
        return err
    return f"토큰 요청 실패 ({r.status_code})"


def exchange_code(client_id: str, client_secret: str, code: str) -> dict:
    preferred = ""
    try:
        preferred = str(st.session_state.get("_gc_oauth_redirect") or "")
    except Exception:
        preferred = ""
    last: requests.Response | None = None
    last_mismatch = False
    for ru in _redirect_candidates(preferred):
        r = requests.post(
            _TOKEN,
            data={
                "code": code,
                "client_id": client_id,
                "client_secret": client_secret,
                "redirect_uri": ru,
                "grant_type": "authorization_code",
            },
            timeout=30,
        )
        if r.ok:
            data = r.json()
            data["_obtained"] = datetime.now(timezone.utc).isoformat()
            try:
                st.session_state["_gc_oauth_redirect"] = ru
            except Exception:
                pass
            return data
        last = r
        err = ""
        try:
            err = str((r.json() or {}).get("error") or "")
        except Exception:
            err = ""
        if err == "redirect_uri_mismatch":
            last_mismatch = True
            continue
        break
    if last is not None:
        raise RuntimeError(_token_error_text(last))
    if last_mismatch:
        raise RuntimeError("리다이렉트 주소가 Google Cloud에 등록된 URI와 다릅니다.")
    raise RuntimeError("구글 토큰을 받지 못했습니다.")


def refresh_token(client_id: str, client_secret: str, tok: dict) -> dict:
    rt = str(tok.get("refresh_token") or "")
    if not rt:
        raise RuntimeError("재로그인 필요")
    r = requests.post(
        _TOKEN,
        data={
            "refresh_token": rt,
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "refresh_token",
        },
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    data["refresh_token"] = rt
    data["_obtained"] = datetime.now(timezone.utc).isoformat()
    return data


def _expired(tok: dict) -> bool:
    exp = tok.get("expires_in")
    obtained = tok.get("_obtained")
    if not exp or not obtained:
        return False
    try:
        t0 = datetime.fromisoformat(str(obtained).replace("Z", "+00:00"))
        return datetime.now(timezone.utc) > t0 + timedelta(seconds=int(exp) - 60)
    except Exception:
        return False


def access_token() -> str:
    cid, csec = load_oauth_client()
    tok = load_token()
    if not tok:
        return ""
    at = str(tok.get("access_token") or "")
    need = (not at) or _expired(tok)
    if need and cid and csec and tok.get("refresh_token"):
        try:
            tok = refresh_token(cid, csec, tok)
            save_token(tok)
            at = str(tok.get("access_token") or "")
        except Exception:
            pass
    return at


def api_get(url: str, params: dict | None = None, token: str = "") -> dict:
    tok = token or access_token()
    if not tok:
        raise RuntimeError("구글 연결이 필요합니다.")
    r = requests.get(
        url,
        params=params or {},
        headers={"Authorization": f"Bearer {tok}"},
        timeout=30,
    )
    if r.status_code == 401:
        cid, csec = load_oauth_client()
        tok_d = load_token()
        if cid and csec and tok_d.get("refresh_token"):
            tok_d = refresh_token(cid, csec, tok_d)
            save_token(tok_d)
            r = requests.get(
                url,
                params=params or {},
                headers={"Authorization": f"Bearer {tok_d.get('access_token')}"},
                timeout=30,
            )
    r.raise_for_status()
    return r.json() if r.content else {}


def api_json(method: str, url: str, body: dict | None = None) -> dict:
    tok = access_token()
    if not tok:
        raise RuntimeError("구글 연결이 필요합니다.")
    r = requests.request(
        method,
        url,
        json=body or {},
        headers={"Authorization": f"Bearer {tok}", "Content-Type": "application/json"},
        timeout=30,
    )
    r.raise_for_status()
    return r.json() if r.content else {}


def list_calendars() -> list[dict]:
    data = api_get(f"{_CAL_API}/users/me/calendarList")
    items = data.get("items") or []
    out = []
    for it in items:
        if str(it.get("selected", True)).lower() in ("false", "0"):
            continue
        role = str(it.get("accessRole") or "")
        out.append(
            {
                "id": it.get("id") or "",
                "summary": it.get("summary") or it.get("id") or "",
                "color": it.get("backgroundColor") or "#7b8794",
                "primary": bool(it.get("primary")),
                "writable": role in ("owner", "writer"),
            }
        )
    return out


def list_events(cal_id: str, start: date, end: date, token: str = "") -> list[dict]:
    tmin = datetime.combine(start, time.min, tzinfo=_TZ).isoformat()
    tmax = datetime.combine(end, time.max, tzinfo=_TZ).isoformat()
    data = api_get(
        f"{_CAL_API}/calendars/{urllib.parse.quote(cal_id, safe='')}/events",
        {
            "timeMin": tmin,
            "timeMax": tmax,
            "singleEvents": "true",
            "orderBy": "startTime",
            "maxResults": "250",
            "fields": "items(id,summary,description,start,end)",
        },
        token=token,
    )
    return list(data.get("items") or [])


def _writable_calendars(cals: list[dict]) -> list[dict]:
    out = []
    for c in cals or []:
        cid = str(c.get("id") or "")
        if not cid or cid == GC_KR_HOLIDAY or not c.get("writable"):
            continue
        out.append(c)
    out.sort(key=lambda c: (not bool(c.get("primary")), str(c.get("summary") or "")))
    return out


def insert_event(
    cal_id: str, title: str, day: date, all_day: bool, notes: str = ""
) -> dict:
    cid = str(cal_id or "").strip() or "primary"
    if cid == GC_KR_HOLIDAY:
        raise RuntimeError("공휴일 캘린더에는 일정을 넣을 수 없습니다. 업무 캘린더를 고르세요.")
    body: dict[str, Any] = {"summary": title}
    if str(notes or "").strip():
        body["description"] = str(notes).strip()
    if all_day:
        body["start"] = {"date": day.isoformat()}
        body["end"] = {"date": (day + timedelta(days=1)).isoformat()}
    else:
        s = datetime.combine(day, time(9, 0), tzinfo=_TZ)
        e = s + timedelta(hours=1)
        body["start"] = {"dateTime": s.isoformat(), "timeZone": "Asia/Seoul"}
        body["end"] = {"dateTime": e.isoformat(), "timeZone": "Asia/Seoul"}
    return api_json("POST", f"{_CAL_API}/calendars/{urllib.parse.quote(cid, safe='')}/events", body)


def api_delete(url: str) -> None:
    tok = access_token()
    if not tok:
        raise RuntimeError("구글 연결이 필요합니다.")
    r = requests.delete(url, headers={"Authorization": f"Bearer {tok}"}, timeout=30)
    if r.status_code == 401:
        cid, csec = load_oauth_client()
        tok_d = load_token()
        if cid and csec and tok_d.get("refresh_token"):
            tok_d = refresh_token(cid, csec, tok_d)
            save_token(tok_d)
            r = requests.delete(
                url,
                headers={"Authorization": f"Bearer {tok_d.get('access_token')}"},
                timeout=30,
            )
    if r.status_code in (404, 410):
        return
    r.raise_for_status()


def delete_event(cal_id: str, event_id: str) -> None:
    cid = str(cal_id or "").strip()
    eid = str(event_id or "").strip()
    if not cid or not eid or cid == GC_KR_HOLIDAY:
        return
    api_delete(
        f"{_CAL_API}/calendars/{urllib.parse.quote(cid, safe='')}/events/{urllib.parse.quote(eid, safe='')}"
    )


def patch_event(cal_id: str, event_id: str, fields: dict) -> dict:
    cid = str(cal_id or "").strip()
    eid = str(event_id or "").strip()
    if not cid or not eid or cid == GC_KR_HOLIDAY:
        return {}
    return api_json(
        "PATCH",
        f"{_CAL_API}/calendars/{urllib.parse.quote(cid, safe='')}/events/"
        f"{urllib.parse.quote(eid, safe='')}",
        fields,
    )


def list_tasklists() -> list[dict]:
    data = api_get(f"{_TASK_API}/users/@me/lists")
    return list(data.get("items") or [])


def list_tasks(list_id: str) -> list[dict]:
    data = api_get(
        f"{_TASK_API}/lists/{urllib.parse.quote(list_id, safe='')}/tasks",
        {"showCompleted": "false", "showHidden": "false", "maxResults": "100"},
    )
    items = [t for t in (data.get("items") or []) if t.get("status") != "completed"]
    items.sort(key=lambda t: (t.get("position") or "", t.get("title") or ""))
    return items


def add_task(list_id: str, title: str, notes: str = "") -> dict:
    body: dict[str, Any] = {"title": title}
    if str(notes or "").strip():
        body["notes"] = str(notes).strip()
    return api_json(
        "POST",
        f"{_TASK_API}/lists/{urllib.parse.quote(list_id, safe='')}/tasks",
        body,
    )


def patch_task(list_id: str, task_id: str, fields: dict) -> dict:
    return api_json(
        "PATCH",
        f"{_TASK_API}/lists/{urllib.parse.quote(list_id, safe='')}/tasks/{urllib.parse.quote(task_id, safe='')}",
        fields,
    )


def complete_task(list_id: str, task_id: str) -> dict:
    return api_json(
        "PATCH",
        f"{_TASK_API}/lists/{urllib.parse.quote(list_id, safe='')}/tasks/{urllib.parse.quote(task_id, safe='')}",
        {"status": "completed"},
    )


def _gc_api_box() -> dict:
    return st.session_state.setdefault("_gc_api", {})


def _gc_cached_cals() -> list[dict]:
    box = _gc_api_box()
    if box.get("cals") is None:
        box["cals"] = list_calendars()
    return list(box["cals"] or [])


def _gc_cached_tlists() -> list[dict]:
    box = _gc_api_box()
    if box.get("tlists") is None:
        box["tlists"] = list_tasklists()
    return list(box["tlists"] or [])


def _ev_key(year: int, month: int) -> str:
    return f"ev:{year}-{month:02d}"


def _merge_event_days(
    cached: dict[str, list[dict]], fresh: dict[str, list[dict]]
) -> dict[str, list[dict]]:
    index: dict[tuple[str, str], tuple[str, dict]] = {}
    extras: dict[str, list[dict]] = {}

    def _ingest(src: dict[str, list[dict]], overwrite: bool) -> None:
        for day_k, items in (src or {}).items():
            for ev in items or []:
                eid = str(ev.get("id") or "")
                cid = str(ev.get("cal_id") or "")
                if not eid:
                    extras.setdefault(day_k, []).append(ev)
                    continue
                key = (cid, eid)
                if overwrite or key not in index:
                    index[key] = (day_k, ev)

    _ingest(cached, True)
    _ingest(fresh, True)
    out = dict(extras)
    for day_k, ev in index.values():
        out.setdefault(day_k, []).append(ev)
    return out


def _gc_cached_events(
    cals: list[dict], year: int, month: int, *, force: bool = False
) -> dict[str, list[dict]]:
    box = _gc_api_box()
    key = _ev_key(year, month)
    now = datetime.now(timezone.utc).timestamp()
    cached = box.get(key)
    if cached is None:
        cached = _EV_MEM.get(key)
    if cached is not None and not force:
        box[key] = cached
        if key + ":at" not in box:
            box[key + ":at"] = now
        return cached
    fresh = _events_by_day(cals, year, month) if cals else {}
    data = _merge_event_days(cached or {}, fresh) if cached else fresh
    box[key] = data
    box[key + ":at"] = now
    _EV_MEM[key] = data
    return data


def _gc_sync_google(year: int, month: int) -> tuple[int, int]:
    """대시보드↔구글 캘린더·할 일을 한 번에 맞춘다. (일정 수, 할 일 수)"""
    box = _gc_api_box()
    try:
        box["cals"] = list_calendars()
    except Exception:
        pass
    cals = list(box.get("cals") or [])
    key = _ev_key(year, month)
    cached = dict(box.get(key) or _EV_MEM.get(key) or {})
    pushed: dict[str, list[dict]] = {}
    for day_k, items in cached.items():
        row: list[dict] = []
        for ev in items or []:
            eid = str(ev.get("id") or "")
            writable = bool(ev.get("writable"))
            local = eid.startswith("local-") or bool(ev.get("local_only"))
            if writable and local:
                try:
                    day = date.fromisoformat(str(ev.get("day") or day_k)[:10])
                    created = insert_event(
                        str(ev.get("cal_id") or "primary"),
                        str(ev.get("title") or ""),
                        day,
                        True,
                        str(ev.get("description") or ""),
                    )
                    meta = {
                        "id": ev.get("cal_id"),
                        "color": ev.get("color"),
                        "summary": ev.get("cal"),
                        "writable": True,
                    }
                    for c in cals:
                        if str(c.get("id")) == str(ev.get("cal_id")):
                            meta = c
                            break
                    row.append(_event_chip(created, meta) or ev)
                except Exception:
                    row.append(ev)
            else:
                row.append(ev)
        pushed[day_k] = row
    fresh = _events_by_day(cals, year, month) if cals else {}
    merged = _merge_event_days(pushed, fresh)
    now = datetime.now(timezone.utc).timestamp()
    box[key] = merged
    box[key + ":at"] = now
    _EV_MEM[key] = merged
    try:
        box["tlists"] = list_tasklists()
    except Exception:
        pass
    tlists = list(box.get("tlists") or [])
    lid = str(st.session_state.get("_gc_lid") or "")
    if not lid and tlists:
        lid = str(tlists[0].get("id") or "")
        st.session_state["_gc_lid"] = lid
    n_tasks = 0
    if lid:
        n_tasks = len(merge_tasks_with_google(lid))
    n_ev = sum(len(v or []) for v in merged.values())
    return n_ev, n_tasks


def _gc_cache_add_event(day: date, chip: dict) -> None:
    key = _ev_key(day.year, day.month)
    box = _gc_api_box()
    by = dict(box.get(key) or _EV_MEM.get(key) or {})
    day_k = day.isoformat()
    by[day_k] = list(by.get(day_k) or []) + [chip]
    box[key] = by
    box[key + ":at"] = datetime.now(timezone.utc).timestamp()
    _EV_MEM[key] = by


def _gc_cache_patch_event(cal_id: str, event_id: str, fields: dict) -> None:
    cid, eid = str(cal_id), str(event_id)

    def _up(by: dict) -> dict:
        out: dict[str, list[dict]] = {}
        for day_k, items in (by or {}).items():
            nxt = []
            for e in items or []:
                if str(e.get("id")) == eid and str(e.get("cal_id")) == cid:
                    e = dict(e)
                    e.update(fields)
                nxt.append(e)
            if nxt:
                out[day_k] = nxt
        return out

    box = _gc_api_box()
    for k in list(box.keys()):
        if str(k).startswith("ev:") and isinstance(box.get(k), dict):
            box[k] = _up(box[k])
    for k in list(_EV_MEM.keys()):
        _EV_MEM[k] = _up(_EV_MEM[k])


def _gc_find_event(by_day: dict[str, list[dict]], event_id: str) -> dict | None:
    eid = str(event_id or "")
    if not eid:
        return None
    for items in (by_day or {}).values():
        for ev in items or []:
            if str(ev.get("id") or "") == eid:
                return ev
    return None


def _gc_cache_remove_event(cal_id: str, event_id: str) -> None:
    cid, eid = str(cal_id), str(event_id)

    def _strip(by: dict) -> dict:
        out: dict[str, list[dict]] = {}
        for day_k, items in (by or {}).items():
            kept = [
                e
                for e in (items or [])
                if not (str(e.get("id")) == eid and str(e.get("cal_id")) == cid)
            ]
            if kept:
                out[day_k] = kept
        return out

    box = _gc_api_box()
    for k in list(box.keys()):
        if str(k).startswith("ev:") and isinstance(box.get(k), dict):
            box[k] = _strip(box[k])
    for k in list(_EV_MEM.keys()):
        _EV_MEM[k] = _strip(_EV_MEM[k])
    now = datetime.now(timezone.utc).timestamp()
    for k in list(box.keys()):
        if str(k).startswith("ev:") and str(k).endswith(":at"):
            box[k] = now


def _task_store_paths() -> list[str]:
    return [GC_TASKS_HOME, GC_TASKS_ST, GC_TASKS]


def _slim_task(t: dict) -> dict:
    return {
        "id": str((t or {}).get("id") or ""),
        "title": str((t or {}).get("title") or ""),
        "notes": str((t or {}).get("notes") or ""),
        "status": str((t or {}).get("status") or "needsAction"),
        "due": str((t or {}).get("due") or ""),
        "position": str((t or {}).get("position") or ""),
        "local_only": bool((t or {}).get("local_only")),
    }


def load_local_task_lists() -> dict:
    for path in _task_store_paths():
        data = _read_json(path)
        if isinstance(data, dict) and isinstance(data.get("lists"), dict):
            return data
    return {"lists": {}}


def save_local_task_lists(data: dict) -> None:
    payload = {"lists": dict((data or {}).get("lists") or {})}
    for path in _task_store_paths():
        _write_json(path, payload)


def load_local_tasks(list_id: str) -> list[dict]:
    lists = load_local_task_lists().get("lists") or {}
    items = lists.get(list_id) or []
    return [_slim_task(t) for t in items if isinstance(t, dict)]


def save_local_tasks(list_id: str, tasks: list[dict]) -> None:
    data = load_local_task_lists()
    data.setdefault("lists", {})[list_id] = [_slim_task(t) for t in tasks]
    save_local_task_lists(data)
    _gc_set_tasks(list_id, data["lists"][list_id])
    _gc_api_box()[f"tk_at:{list_id}"] = datetime.now(timezone.utc).timestamp()


def merge_tasks_with_google(list_id: str) -> list[dict]:
    """대시보드 할 일을 구글에 보내고, 구글에만 있는 할 일은 여기에도 붙인다."""
    pushed: list[dict] = []
    for t in load_local_tasks(list_id):
        title = str(t.get("title") or "").strip()
        notes = str(t.get("notes") or "")
        tid = str(t.get("id") or "")
        if not title:
            continue
        if t.get("local_only") or not tid or tid.startswith("local-"):
            try:
                pushed.append(_slim_task(add_task(list_id, title, notes)))
            except Exception:
                pushed.append(_slim_task(t))
            continue
        pushed.append(_slim_task(t))
    seen = {str(t.get("id") or "") for t in pushed}
    try:
        remote = [_slim_task(x) for x in (list_tasks(list_id) if list_id else [])]
    except Exception:
        save_local_tasks(list_id, pushed)
        return pushed
    remote_ids = {str(x.get("id") or "") for x in remote}
    out: list[dict] = []
    for t in pushed:
        tid = str(t.get("id") or "")
        if t.get("local_only") or tid.startswith("local-"):
            out.append(t)
            continue
        if tid in remote_ids:
            out.append(t)
    for r in remote:
        rid = str(r.get("id") or "")
        if rid and rid not in seen:
            out.append(r)
            seen.add(rid)
    save_local_tasks(list_id, out)
    return out


def import_tasks_from_google(list_id: str) -> list[dict]:
    return merge_tasks_with_google(list_id)


def _gc_refresh_tasks(list_id: str) -> list[dict]:
    return merge_tasks_with_google(list_id)


def _gc_cached_tasks(list_id: str, *, force: bool = False) -> list[dict]:
    if not list_id:
        return []
    box = _gc_api_box()
    now = datetime.now(timezone.utc).timestamp()
    age = now - float(box.get(f"tk_at:{list_id}") or 0)
    lists = load_local_task_lists().get("lists") or {}
    if not force and list_id in lists and age < 20:
        local = [_slim_task(t) for t in (lists.get(list_id) or []) if isinstance(t, dict)]
        _gc_set_tasks(list_id, local)
        return local
    return merge_tasks_with_google(list_id)


def _gc_set_tasks(list_id: str, tasks: list[dict]) -> None:
    box = _gc_api_box()
    box[f"tk:{list_id}"] = list(tasks)


def _event_chip(it: dict, cal: dict) -> dict | None:
    key = event_day_key(it)
    if not key:
        return None
    stt = (it.get("start") or {}).get("dateTime") or ""
    hhmm = ""
    if "T" in stt:
        try:
            dt = datetime.fromisoformat(stt.replace("Z", "+00:00")).astimezone(_TZ)
            hhmm = dt.strftime("%H:%M")
        except Exception:
            hhmm = ""
    return {
        "title": it.get("summary") or "(제목 없음)",
        "color": cal.get("color") or "#7b8794",
        "cal": cal.get("summary") or "",
        "hhmm": hhmm,
        "day": key,
        "description": str(it.get("description") or ""),
        "id": str(it.get("id") or ""),
        "cal_id": str(cal.get("id") or ""),
        "writable": bool(cal.get("writable"))
        and str(cal.get("id") or "") != GC_KR_HOLIDAY,
    }


def _events_by_day(cals: list[dict], year: int, month: int) -> dict[str, list[dict]]:
    weeks = month_weeks(year, month)
    start, end = weeks[0][0], weeks[-1][-1]
    tok = access_token()

    def _one(cal: dict) -> list[tuple[str, dict]]:
        try:
            items = list_events(cal["id"], start, end, token=tok) if tok else []
        except Exception:
            return []
        out: list[tuple[str, dict]] = []
        for it in items:
            chip = _event_chip(it, cal)
            if chip:
                out.append((str(chip.get("day")), chip))
        return out

    by: dict[str, list[dict]] = {}
    workers = min(8, max(1, len(cals)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futs = [pool.submit(_one, cal) for cal in cals]
        for fut in as_completed(futs):
            try:
                rows = fut.result()
            except Exception:
                continue
            for key, chip in rows:
                by.setdefault(key, []).append(chip)
    return by


def _month_html(
    year: int,
    month: int,
    by_day: dict[str, list[dict]],
    today: date,
    selected: date | None = None,
) -> str:
    weeks = month_weeks(year, month)
    heads = "".join(
        f'<th class="gc-dow{" gc-sun" if i == 0 else " gc-sat" if i == 6 else ""}">'
        f'{d}</th>'
        for i, d in enumerate(("일", "월", "화", "수", "목", "금", "토"))
    )
    rows = []
    for week in weeks:
        wn = iso_week_label(week)
        cells = [f'<td class="gc-wn">{wn}</td>']
        for d in week:
            key = d.isoformat()
            outside = d.month != month
            cls = ["gc-cell"]
            if outside:
                cls.append("gc-out")
            if d == today:
                cls.append("gc-today")
            if selected is not None and d == selected:
                cls.append("gc-sel")
            if d.weekday() == 6:
                cls.append("gc-sun")
            if d.weekday() == 5:
                cls.append("gc-sat")
            if _gc_is_holiday_day(by_day.get(key)):
                cls.append("gc-hol")
            chips = []
            for ev in (by_day.get(key) or [])[:4]:
                label = html.escape((ev.get("hhmm") + " " if ev.get("hhmm") else "") + ev["title"])
                col = html.escape(str(ev.get("color") or "#c5c9ce"))
                xbtn = ""
                if ev.get("writable") and ev.get("id") and ev.get("cal_id"):
                    xbtn = '<span class="gc-x">×</span>'
                chips.append(
                    f'<div class="gc-chip" style="background:{col}22;border-left:3px solid {col}">'
                    f'<span class="gc-chip-t">{label}</span>{xbtn}</div>'
                )
            more = len(by_day.get(key) or []) - 4
            if more > 0:
                chips.append(f'<div class="gc-more">+{more}</div>')
            cells.append(
                f'<td class="{" ".join(cls)}"><div class="gc-num">{d.day}</div>'
                f'{"".join(chips)}</td>'
            )
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return f"""
<style>
.gc-wrap {{ font-family: "Google Sans", "Noto Sans KR", system-ui, sans-serif; color:#3c4043; }}
.gc-table {{ width:100%; border-collapse:separate; border-spacing:6px; table-layout:fixed; background:transparent; }}
.gc-table th {{ font-weight:500; font-size:12px; padding:8px 6px; text-align:center; color:#70757a; }}
.gc-wn {{ width:36px; color:#80868b; font-size:11px; text-align:center; background:transparent; border:none; }}
.gc-cell {{ height:72px; vertical-align:top; border:1px solid #eceff3; border-radius:10px; padding:8px 4px; background:#fff; }}
.gc-out {{ color:#bdc1c6; }}
.gc-today {{ background:#e8f4ff; }}
.gc-today .gc-num {{ background:#1a73e8; color:#fff; border-radius:50%; width:22px; height:22px;
  display:inline-flex; align-items:center; justify-content:center; }}
.gc-sel {{ background:#e8f0fe; }}
.gc-sun, .gc-sat, .gc-hol {{ background:#fdecea; }}
.gc-today.gc-sun, .gc-today.gc-sat, .gc-today.gc-hol {{ background:#e8f4ff; }}
.gc-sel.gc-sun, .gc-sel.gc-sat, .gc-sel.gc-hol {{ background:#e8f0fe; }}
.gc-num {{ font-size:12px; font-weight:500; margin-bottom:2px; }}
.gc-sun, .gc-sun .gc-num, .gc-hol, .gc-hol .gc-num {{ color:#d93025; }}
.gc-sat, .gc-sat .gc-num {{ color:#c62828; }}
.gc-chip {{ font-size:11px; line-height:1.25; padding:1px 2px 1px 4px; margin:1px 0; border-radius:2px;
  display:flex; align-items:center; gap:2px; overflow:hidden; }}
.gc-chip-t {{ flex:1; min-width:0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.gc-x {{ flex:0 0 auto; color:#80868b; font-size:13px; padding:0 3px; line-height:1; cursor:pointer; }}
.gc-more {{ font-size:11px; color:#5f6368; }}
</style>
<div class="gc-wrap"><table class="gc-table"><thead><tr><th class="gc-wn"></th>{heads}</tr></thead>
<tbody>{"".join(rows)}</tbody></table></div>
"""


def _gc_set_del_event(cal_id: str, event_id: str) -> None:
    st.session_state["_gc_pending_del_event"] = {"cal": cal_id, "id": event_id}


def _gc_set_sel_day(day: date) -> None:
    st.session_state["_gc_sel_day"] = day


def _gc_set_sel_event(day: date, event_id: str) -> None:
    st.session_state["_gc_sel_day"] = day
    st.session_state["_gc_sel_event"] = str(event_id or "")
    st.session_state["_gc_memo_open"] = True


def _gc_close_memo_pop() -> None:
    st.session_state["_gc_memo_open"] = False


def _gc_open_add_pop() -> None:
    day = st.session_state.get("_gc_sel_day")
    if not isinstance(day, date):
        day = datetime.now(_TZ).date()
    st.session_state["_gc_add_pop"] = True
    st.session_state["gc_ev_day"] = day
    st.session_state.pop("gc_ev_title", None)
    st.session_state.pop("gc_ev_note", None)


def _gc_close_add_pop() -> None:
    st.session_state["_gc_add_pop"] = False


def _render_month_grid(
    year: int,
    month: int,
    by_day: dict[str, list[dict]],
    today: date,
    selected: date | None = None,
) -> None:
    """표 달력. 제목은 클릭하지 않고, 쓸 수 있는 일정만 ×로 지운다."""
    st.markdown(
        """
<style>
div[class*="st-key-gc_calhead"] [data-testid="stHorizontalBlock"],
div[class*="st-key-gc_calhead"] .stHorizontalBlock {
  display: flex !important;
  flex-wrap: nowrap !important;
  gap: 6px !important;
  margin-bottom: 4px !important;
  overflow: visible !important;
}
div[class*="st-key-gc_calgrid"] [data-testid="stHorizontalBlock"],
div[class*="st-key-gc_calgrid"] .stHorizontalBlock {
  display: flex !important;
  flex-wrap: nowrap !important;
  gap: 6px !important;
  margin-bottom: 6px !important;
  overflow: hidden !important;
}
div[class*="st-key-gc_calhead"] [data-testid="stColumn"],
div[class*="st-key-gc_calhead"] .stColumn {
  padding: 2px 2px 6px !important;
  background: transparent !important;
  border: none !important;
  min-width: 0 !important;
  min-height: 1.6rem !important;
  flex: 1 1 0 !important;
  overflow: visible !important;
}
div[class*="st-key-gc_calhead"] [data-testid="stMarkdownContainer"],
div[class*="st-key-gc_calhead"] [data-testid="stMarkdownContainer"] p,
div[class*="st-key-gc_calhead"] [data-testid="stMarkdownContainer"] div,
div[class*="st-key-gc_calhead"] .gc-dow-lab {
  font-size: 13px !important;
  font-weight: 700 !important;
  line-height: 1.4 !important;
  text-align: center !important;
  overflow: visible !important;
  margin: 0 !important;
}
div[class*="st-key-gc_calgrid"] [data-testid="stColumn"],
div[class*="st-key-gc_calgrid"] .stColumn {
  border: 1px solid #eceff3;
  border-radius: 10px;
  min-height: 88px;
  min-width: 0 !important;
  flex: 1 1 0 !important;
  padding: 4px 2px !important;
  background: #fff;
}
div[class*="st-key-gc_calhead"] [data-testid="stHorizontalBlock"] > div:first-child,
div[class*="st-key-gc_calgrid"] [data-testid="stHorizontalBlock"] > div:first-child,
div[class*="st-key-gc_calhead"] .stHorizontalBlock > div:first-child,
div[class*="st-key-gc_calgrid"] .stHorizontalBlock > div:first-child {
  flex: 0 0 28px !important;
  width: 28px !important;
  max-width: 28px !important;
  background: transparent !important;
  border: none !important;
  border-radius: 0 !important;
  box-shadow: none !important;
  min-height: 0 !important;
}
div[class*="st-key-gc_calgrid"] [data-testid="stColumn"]:has([class*="st-key-gc_cellsel"]) {
  background: #e8f0fe !important;
  box-shadow: none !important;
}
div[class*="st-key-gc_calgrid"] [data-testid="stColumn"]:has([class*="st-key-gc_celltoday"]) {
  background: #e8f4ff !important;
}
div[class*="st-key-gc_calgrid"] [data-testid="stColumn"]:has([class*="st-key-gc_cellpink"]) {
  background: #fdecea !important;
}
div[class*="st-key-gc_calgrid"] [data-testid="stVerticalBlock"] {
  gap: 2px !important;
}
div[class*="st-key-gc_calhead"] [data-testid="stMarkdownContainer"] p,
div[class*="st-key-gc_calgrid"] [data-testid="stMarkdownContainer"] p {
  margin: 0 !important;
}
div[class*="st-key-gc_chiprow_"],
div[class*="st-key-gc_evsel_"] {
  position: relative !important;
}
div[class*="st-key-gc_xd_"] {
  position: absolute !important;
  top: 0 !important;
  right: 1px !important;
  z-index: 3 !important;
  width: 1.2rem !important;
}
div[class*="st-key-gc_xd_"] button {
  min-height: 1.3rem !important;
  height: 1.3rem !important;
  min-width: 1.15rem !important;
  width: 1.15rem !important;
  padding: 0 !important;
  font-size: 13px !important;
  line-height: 1 !important;
  color: #5f6368 !important;
  border: none !important;
  background: transparent !important;
}
div[class*="st-key-gc_chiprow_"] [data-testid="stHorizontalBlock"],
div[class*="st-key-gc_evsel_"] [data-testid="stHorizontalBlock"] {
  display: flex !important;
  flex-wrap: nowrap !important;
  gap: 0 !important;
  align-items: center !important;
}
div[class*="st-key-gc_evlist_"] {
  max-height: 96px !important;
  overflow-y: auto !important;
  overflow-x: hidden !important;
  padding: 0 2px 4px !important;
}
.gc-ev-lab {
  font-size: 11px !important;
  font-weight: 400 !important;
  line-height: 1.25 !important;
  color: #174ea6 !important;
  white-space: nowrap !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  padding: 1px 16px 1px 4px !important;
}
div[class*="st-key-gc_evpick_"] {
  position: absolute !important;
  inset: 0 !important;
  opacity: 0 !important;
  z-index: 2 !important;
}
div[class*="st-key-gc_evpick_"] button {
  min-height: 1.2rem !important;
  height: 100% !important;
  width: 100% !important;
  padding: 0 4px !important;
  border: none !important;
  outline: none !important;
  background: transparent !important;
  box-shadow: none !important;
}
div[class*="st-key-gc_evsel_"] div[class*="st-key-gc_evpick_"] button {
  background: #d2e3fc !important;
}
div[class*="st-key-gc_memo_pop"],
div[class*="st-key-gc_add_pop"] {
  position: fixed !important;
  top: 18% !important;
  left: 50% !important;
  transform: translateX(-50%) !important;
  z-index: 10050 !important;
  width: min(520px, 92vw) !important;
  background: #fff !important;
  border: 1px solid #e0e3e8 !important;
  border-radius: 12px !important;
  box-shadow: 0 12px 40px rgba(32, 33, 36, 0.22) !important;
  padding: 14px 16px 12px !important;
}
div[class*="st-key-gc_day_"] button {
  min-height: 1.35rem !important;
  height: 1.35rem !important;
  width: 100% !important;
  padding: 0 2px !important;
  font-size: 12px !important;
  font-weight: 600 !important;
  border: none !important;
  outline: none !important;
  border-radius: 8px !important;
  background: transparent !important;
  color: inherit !important;
  box-shadow: none !important;
  justify-content: center !important;
  align-items: center !important;
  line-height: 1.2 !important;
  text-align: center !important;
}
div[class*="st-key-gc_day_"] button:focus,
div[class*="st-key-gc_day_"] button:focus-visible,
div[class*="st-key-gc_day_"] button:hover {
  border: none !important;
  outline: none !important;
  box-shadow: none !important;
}
div[class*="st-key-gc_calgrid"] [data-testid="stColumn"]:has([class*="st-key-gc_cellsel"]),
div[class*="st-key-gc_cellsel"] {
  background: #e8f0fe !important;
  box-shadow: none !important;
}
div[class*="st-key-gc_celltoday"] {
  background: #e8f4ff !important;
}
div[class*="st-key-gc_cellpink"] {
  background: #fdecea !important;
}
</style>
""",
        unsafe_allow_html=True,
    )
    with st.container(key="gc_calhead"):
        head = st.columns([0.38] + [1] * 7, gap=None)
        with head[0]:
            st.markdown('<div class="gc-wn"></div>', unsafe_allow_html=True)
        for i, name in enumerate(("일", "월", "화", "수", "목", "금", "토")):
            color = "#d93025" if i == 0 else "#1a73e8" if i == 6 else "#70757a"
            with head[i + 1]:
                st.markdown(
                    f'<div class="gc-dow-lab" style="color:{color}">{name}</div>',
                    unsafe_allow_html=True,
                )
    weeks = month_weeks(year, month)
    day_rules: list[str] = []
    planned: list[list[tuple]] = []
    for week in weeks:
        row = []
        for d in week:
            iso = d.isoformat()
            stamp = d.strftime("%Y%m%d")
            is_sel = selected is not None and d == selected
            items = by_day.get(iso) or []
            is_hol = _gc_is_holiday_day(items)
            is_pink = (not is_sel) and (d != today) and (d.weekday() >= 5 or is_hol)
            if d == today:
                bg = "#e8f4ff"
            elif d.weekday() >= 5 or is_hol:
                bg = "#fdecea"
            else:
                bg = "#fff"
            if is_sel and d != today:
                bg = "#e8f0fe"
            if d.month != month:
                color = "#bdc1c6"
            elif d.weekday() == 6 or is_hol:
                color = "#d93025"
            elif d.weekday() == 5:
                color = "#c62828"
            else:
                color = "#3c4043"
            btn_extra = (
                "border:none!important;outline:none!important;box-shadow:none!important;"
                "width:100%!important;max-width:none!important;min-width:0!important;"
                "cursor:pointer!important;pointer-events:auto!important;"
            )
            if d == today and d.month == month:
                color = "#1a73e8"
                btn_extra += "font-weight:700!important;"
            elif is_sel:
                color = "#174ea6"
                btn_extra += (
                    "background:#d2e3fc!important;font-weight:700!important;"
                )
            day_rules.append(
                f"div[class*='st-key-gc_calgrid'] [data-testid='stColumn']"
                f":has([class*='st-key-gc_cell_{stamp}']),"
                f"div[class*='st-key-gc_cell_{stamp}']{{background:{bg}!important;}}"
                f"div[class*='st-key-gc_day_{stamp}'] button{{color:{color}!important;{btn_extra}}}"
            )
            row.append((d, stamp, items, iso))
        planned.append(row)
    if day_rules:
        st.markdown("<style>" + "".join(day_rules) + "</style>", unsafe_allow_html=True)
    with st.container(key="gc_calgrid"):
        for week, row in zip(weeks, planned):
            cols = st.columns([0.38] + [1] * 7, gap=None, vertical_alignment="top")
            with cols[0]:
                st.markdown(
                    f'<div style="color:#80868b;font-size:11px;text-align:center;'
                    f'padding-top:4px">{iso_week_label(week)}</div>',
                    unsafe_allow_html=True,
                )
            for i, (d, stamp, items, iso) in enumerate(row):
                with cols[i + 1]:
                    with st.container(key=f"gc_cell_{stamp}"):
                        st.button(
                            str(d.day),
                            key=f"gc_day_{stamp}",
                            type="tertiary",
                            width="stretch",
                            on_click=_gc_set_sel_day,
                            args=(d,),
                        )
                        if items:
                            with st.container(key=f"gc_evlist_{stamp}"):
                                for idx, ev in enumerate(items):
                                    eid = str(ev.get("id") or "")
                                    cid = str(ev.get("cal_id") or "")
                                    lab = (
                                        (str(ev.get("hhmm") or "") + " ")
                                        if ev.get("hhmm")
                                        else ""
                                    ) + str(ev.get("title") or "")
                                    if len(lab) > 14:
                                        lab = lab[:13] + "…"
                                    lab = lab or "(일정)"
                                    with st.container(key=f"gc_chiprow_{stamp}_{idx}"):
                                        st.markdown(
                                            '<div class="gc-ev-lab" style="font-size:11px;'
                                            "font-weight:400;line-height:1.25;color:#174ea6;"
                                            "white-space:nowrap;overflow:hidden;"
                                            "text-overflow:ellipsis;padding:1px 16px 1px 4px\">"
                                            f"{html.escape(lab)}</div>",
                                            unsafe_allow_html=True,
                                        )
                                        st.button(
                                            lab,
                                            key=f"gc_evpick_{stamp}_{idx}",
                                            type="tertiary",
                                            width="stretch",
                                            on_click=_gc_set_sel_event,
                                            args=(d, eid or iso),
                                        )
                                        if ev.get("writable") and eid and cid:
                                            st.button(
                                                "×",
                                                key=f"gc_xd_{stamp}_{idx}",
                                                type="tertiary",
                                                width="content",
                                                on_click=_gc_set_del_event,
                                                args=(cid, eid),
                                            )


def _render_event_memo_pop(ev: dict) -> None:
    """일정 클릭 메모. 대화상자 대신 조각 안 팝업이라 대시보드를 다시 켜지 않는다."""
    title = str(ev.get("title") or "(제목 없음)")
    eid = str(ev.get("id") or "")
    cid = str(ev.get("cal_id") or "")
    note = str(ev.get("description") or "")
    writable = bool(ev.get("writable"))
    nk = "gc_pop_note_" + (eid[-24:] if eid else "x")
    if nk not in st.session_state:
        st.session_state[nk] = note
    with st.container(key="gc_memo_pop"):
        st.markdown(f"**{html.escape(title)}**")
        meta = " · ".join(
            p for p in (str(ev.get("day") or ""), str(ev.get("cal") or "")) if p
        )
        if meta:
            st.caption(meta)
        if writable:
            st.text_area("메모", key=nk, height=90, placeholder="일정 메모")
            b1, b2 = st.columns(2)
            with b1:
                if st.button("메모 저장", type="primary", key="gc_pop_save"):
                    st.session_state["_gc_pending_ev_note"] = {
                        "cal": cid,
                        "id": eid,
                        "notes": st.session_state.get(nk) or "",
                    }
                    st.session_state["_gc_memo_open"] = False
                    st.rerun(scope="fragment")
            with b2:
                st.button(
                    "닫기",
                    key="gc_pop_close",
                    type="tertiary",
                    on_click=_gc_close_memo_pop,
                )
        else:
            st.markdown(
                "<div style='font-size:12px;color:#31333F;white-space:pre-wrap;"
                f"line-height:1.4'>{html.escape(note) if note else '메모 없음'}</div>",
                unsafe_allow_html=True,
            )
            st.button(
                "닫기",
                key="gc_pop_close_ro",
                type="tertiary",
                on_click=_gc_close_memo_pop,
            )


def _render_add_event_pop(cals: list[dict]) -> None:
    """일자 선택 후 일정 추가. 대화상자 대신 조각 안 팝업이라 날짜 버튼을 건드리지 않는다."""
    with st.container(key="gc_add_pop"):
        h1, h2 = st.columns([6, 1])
        with h1:
            st.markdown("**일정 추가**")
        with h2:
            st.button(
                "닫기",
                key="gc_add_close",
                type="tertiary",
                on_click=_gc_close_add_pop,
            )
        et = st.text_input("제목", key="gc_ev_title")
        ed = st.date_input("날짜", key="gc_ev_day")
        en = st.text_area("메모", key="gc_ev_note", height=70, placeholder="일정 메모")
        writable = _writable_calendars(cals)
        cal_opts = {c["summary"]: c["id"] for c in writable} or {"기본": "primary"}
        if st.session_state.get("gc_ev_cal") not in cal_opts:
            st.session_state["gc_ev_cal"] = next(iter(cal_opts))
        pick = st.selectbox("캘린더", options=list(cal_opts.keys()), key="gc_ev_cal")
        if st.button("일정 저장", type="primary", key="gc_ev_save"):
            title = (et or "").strip()
            if len(title) < 1:
                st.error("제목을 입력하세요.")
            else:
                st.session_state["_gc_pending_event"] = {
                    "cal": cal_opts[pick],
                    "title": title,
                    "day": ed,
                    "notes": en or "",
                }
                st.session_state["_gc_add_pop"] = False
                st.rerun(scope="fragment")


def _qp_pop(*names: str) -> None:
    try:
        qp = st.query_params
        for name in names:
            try:
                del qp[name]
            except Exception:
                pass
    except Exception:
        pass


def _oauth_query_one(name: str) -> str:
    try:
        v = st.query_params.get(name)
        if isinstance(v, list):
            v = v[0] if v else ""
        return str(v or "").strip()
    except Exception:
        return ""


def _consume_oauth_redirect() -> None:
    err = _oauth_query_one("error")
    code = _oauth_query_one("code")
    state = _oauth_query_one("state")
    if err:
        st.error("구글 연결이 취소되었습니다.")
        _qp_pop("error", "code", "state", "scope", "prompt", "authuser")
        return
    if not code:
        return
    if st.session_state.get("_gc_oauth_code_used") == code:
        _qp_pop("code", "state", "scope", "prompt", "authuser", "hd", "session_state")
        return
    if _usable_token(load_token()):
        st.session_state["_gc_oauth_code_used"] = code
        _qp_pop("code", "state", "scope", "prompt", "authuser", "hd", "session_state")
        return
    expect = str(st.session_state.get("_gc_oauth_state") or "")
    if expect and state and state != expect:
        st.error("구글 연결 상태가 맞지 않습니다. 다시 연결해 주세요.")
        return
    cid, csec = load_oauth_client()
    if not cid or not csec:
        st.error("OAuth 클라이언트 ID/비밀번호가 없습니다.")
        return
    st.session_state["_gc_oauth_code_used"] = code
    try:
        tok = exchange_code(cid, csec, code)
        save_token(tok)
        st.session_state.pop("_gc_oauth_state", None)
        _qp_pop("code", "state", "scope", "prompt", "authuser", "hd", "session_state")
        st.success("구글 캘린더·할 일이 연결되었습니다.")
        st.rerun()
    except Exception as e:
        st.session_state.pop("_gc_oauth_code_used", None)
        st.error(f"구글 연결 실패: {e}")


def _setup_expander() -> None:
    cid, csec = load_oauth_client()
    with st.expander("구글 연결 설정 (최초 1회)", expanded=not bool(cid and csec)):
        st.markdown(
            "Google Cloud에서 **캘린더 API · Tasks API**를 켜고, OAuth 클라이언트(웹)를 만든 뒤 "
            "아래 리다이렉트 URI를 등록하세요."
        )
        st.code(
            "http://localhost:8501/\n"
            "http://127.0.0.1:8501/\n"
            f"{_CLOUD_REDIRECT}",
            language=None,
        )
        i1, i2 = st.columns(2)
        with i1:
            nid = st.text_input("클라이언트 ID", value=cid, type="default", key="gc_cid")
        with i2:
            nsec = st.text_input("클라이언트 비밀번호", value=csec, type="password", key="gc_csec")
        if st.button("설정 저장", key="gc_save_oauth"):
            if len((nid or "").strip()) < 8 or len((nsec or "").strip()) < 8:
                st.error("클라이언트 ID와 비밀번호를 넣어 주세요.")
            else:
                save_oauth_client(nid, nsec)
                st.success("저장했습니다. 아래에서 구글 계정을 연결하세요.")
                st.rerun()


@st.fragment
def _gc_cal_panel() -> None:
    """달력 월 이동·일정 저장만 이 조각을 다시 그린다."""
    today = datetime.now(_TZ).date()
    pending_m = st.session_state.pop("_gc_pending_month", None)
    if isinstance(pending_m, date):
        st.session_state["gc_year"] = pending_m.year
        st.session_state["gc_month"] = pending_m.month
    elif not st.session_state.get("gc_year"):
        vc_mon = st.session_state.get("_vc_month")
        if isinstance(vc_mon, date):
            st.session_state["gc_year"] = vc_mon.year
            st.session_state["gc_month"] = vc_mon.month
    y = int(st.session_state.get("gc_year") or today.year)
    m = int(st.session_state.get("gc_month") or today.month)

    pending_ev = st.session_state.pop("_gc_pending_event", None)
    if isinstance(pending_ev, dict):
        day = pending_ev["day"]
        cal_id = str(pending_ev.get("cal") or "primary")
        if cal_id == GC_KR_HOLIDAY:
            wlist = _writable_calendars(_gc_api_box().get("cals") or [])
            cal_id = str((wlist[0].get("id") if wlist else "") or "primary")
        cal_meta = {"id": cal_id, "summary": "", "color": "#1a73e8", "writable": True}
        for c in (_gc_api_box().get("cals") or []):
            if str(c.get("id")) == cal_id:
                cal_meta = c
                break
        notes = str(pending_ev.get("notes") or "")
        chip = {
            "title": str(pending_ev.get("title") or ""),
            "color": cal_meta.get("color") or "#1a73e8",
            "cal": cal_meta.get("summary") or "",
            "hhmm": "",
            "day": day.isoformat(),
            "description": notes,
            "id": "local-" + secrets.token_hex(4),
            "cal_id": cal_id,
            "writable": True,
        }
        _gc_cache_add_event(day, chip)
        st.session_state.pop("gc_ev_title", None)
        st.session_state.pop("gc_ev_note", None)
        try:
            created = insert_event(
                cal_id, str(pending_ev.get("title") or ""), day, True, notes
            )
            real = _event_chip(created, cal_meta)
            if real:
                _gc_cache_remove_event(str(chip.get("cal_id")), str(chip.get("id")))
                _gc_cache_add_event(day, real)
        except Exception as e:
            st.error(f"일정 저장 실패: {e}")
    pending_note = st.session_state.pop("_gc_pending_ev_note", None)
    if isinstance(pending_note, dict) and pending_note.get("id"):
        cid = str(pending_note.get("cal") or "")
        eid = str(pending_note.get("id") or "")
        notes = str(pending_note.get("notes") or "")
        _gc_cache_patch_event(cid, eid, {"description": notes})
        if eid and not eid.startswith("local-") and cid:
            try:
                patch_event(cid, eid, {"description": notes})
            except Exception as e:
                st.error(f"메모 저장 실패: {e}")
    pending_del = st.session_state.pop("_gc_pending_del_event", None)
    if isinstance(pending_del, dict) and pending_del.get("id"):
        cid = str(pending_del.get("cal") or "primary")
        eid = str(pending_del.get("id") or "")
        _gc_cache_remove_event(cid, eid)
        if eid and not eid.startswith("local-"):
            try:
                delete_event(cid, eid)
            except Exception as e:
                st.error(f"일정 삭제 실패: {e}")

    try:
        cals = _gc_cached_cals()
    except Exception as e:
        st.error(f"캘린더를 불러오지 못했습니다: {e}")
        cals = []

    top = st.columns([1.1, 1.1, 4.2, 1.6])
    with top[0]:
        if st.button("오늘", key="gc_today"):
            st.session_state["_gc_pending_month"] = date(today.year, today.month, 1)
            st.rerun(scope="fragment")
    with top[1]:
        n1, n2 = st.columns(2)
        with n1:
            if st.button("‹", key="gc_prev"):
                d = date(y, m, 1) - timedelta(days=1)
                st.session_state["_gc_pending_month"] = date(d.year, d.month, 1)
                st.rerun(scope="fragment")
        with n2:
            if st.button("›", key="gc_next"):
                d = date(y, m, 28) + timedelta(days=8)
                st.session_state["_gc_pending_month"] = date(d.year, d.month, 1)
                st.rerun(scope="fragment")
    with top[2]:
        st.markdown(f"**{y}년 {m}월**")
    with top[3]:
        st.button(
            "일정 추가",
            type="primary",
            key="gc_add_open",
            on_click=_gc_open_add_pop,
        )

    y = int(st.session_state.get("gc_year") or today.year)
    m = int(st.session_state.get("gc_month") or today.month)
    by_day = _gc_cached_events(cals, y, m)
    sel_day = st.session_state.get("_gc_sel_day")
    _render_month_grid(
        y, m, by_day, today, selected=sel_day if isinstance(sel_day, date) else None
    )
    if st.session_state.get("_gc_memo_open"):
        pop_ev = _gc_find_event(
            by_day, str(st.session_state.get("_gc_sel_event") or "")
        )
        if pop_ev:
            _render_event_memo_pop(pop_ev)
        else:
            st.session_state["_gc_memo_open"] = False
    if st.session_state.get("_gc_add_pop"):
        _render_add_event_pop(cals)


@st.fragment
def _gc_task_panel() -> None:
    """할 일만 이 조각을 다시 그린다."""
    lid = str(st.session_state.get("_gc_lid") or "")
    pending_new = st.session_state.pop("_gc_pending_new", None)
    if pending_new:
        st.session_state["gc_new_task"] = ""
        st.session_state["gc_new_note"] = ""
    if pending_new and lid:
        if isinstance(pending_new, dict):
            title_in = str(pending_new.get("title") or "")
            notes_in = str(pending_new.get("notes") or "")
        else:
            title_in, notes_in = str(pending_new), ""
        created: dict[str, Any]
        try:
            created = _slim_task(add_task(lid, title_in, notes_in))
        except Exception as e:
            created = {
                "id": "local-" + secrets.token_hex(8),
                "title": title_in,
                "notes": notes_in,
                "status": "needsAction",
                "due": "",
                "position": "",
                "local_only": True,
            }
            st.warning(f"구글 전송은 실패했고 대시보드에만 저장했습니다: {e}")
        cur = load_local_tasks(lid)
        cur.append(created)
        save_local_tasks(lid, cur)
    pending_note = st.session_state.pop("_gc_pending_note", None)
    if isinstance(pending_note, dict) and lid and pending_note.get("id"):
        nid = str(pending_note.get("id") or "")
        ntxt = str(pending_note.get("notes") or "")
        cur = []
        for t in load_local_tasks(lid):
            if str(t.get("id")) == nid:
                t = dict(t)
                t["notes"] = ntxt
                if not t.get("local_only") and nid and not nid.startswith("local-"):
                    try:
                        patch_task(lid, nid, {"notes": ntxt})
                    except Exception as e:
                        t["local_only"] = True
                        st.warning(f"구글 메모 반영 실패, 대시보드에만 저장: {e}")
            cur.append(t)
        save_local_tasks(lid, cur)
    pending_done = st.session_state.pop("_gc_pending_done", None)
    if pending_done and lid:
        nid = str(pending_done)
        left = []
        for t in load_local_tasks(lid):
            if str(t.get("id")) == nid:
                if not t.get("local_only") and nid and not nid.startswith("local-"):
                    try:
                        complete_task(lid, nid)
                    except Exception as e:
                        st.warning(f"구글 완료 반영 실패: {e}")
                continue
            left.append(t)
        save_local_tasks(lid, left)
    try:
        tlists = _gc_cached_tlists()
    except Exception as e:
        st.error(f"할 일 목록을 불러오지 못했습니다: {e}")
        tlists = []
    st.markdown(
        """
<style>
div[class*="st-key-gc_ne_"] button,
div[class*="st-key-gc_toggle_new_note"] button,
div[class*="st-key-gc_ns_"] button {
  min-height: 1.55rem !important;
  height: 1.55rem !important;
  min-width: 0 !important;
  padding: 0 0.4rem !important;
  font-size: 11px !important;
  line-height: 1.2 !important;
  white-space: nowrap !important;
  border-radius: 6px !important;
}
div[class*="st-key-gc_td_"] p,
div[class*="st-key-gc_td_"] label,
div[class*="st-key-gc_td_"] [data-testid="stWidgetLabel"],
div[class*="st-key-gc_td_"] [data-testid="stWidgetLabel"] p {
  font-size: 11px !important;
  line-height: 1.2 !important;
}
</style>
""",
        unsafe_allow_html=True,
    )
    st.markdown("**할 일 목록**")
    if not tlists:
        st.caption("할 일 목록이 없습니다.")
        return
    names = [t.get("title") or "할 일" for t in tlists]
    li = st.selectbox(
        "목록",
        options=list(range(len(names))),
        format_func=lambda i: names[i],
        key="gc_tlist",
    )
    lid = str(tlists[int(li)].get("id") or "")
    st.session_state["_gc_lid"] = lid
    try:
        tasks = _gc_cached_tasks(lid) if lid else []
    except Exception as e:
        st.error(str(e))
        tasks = []
    edit_id = str(st.session_state.get("_gc_note_edit") or "")
    for t in tasks:
        tid = str(t.get("id") or "")
        title = t.get("title") or "(제목 없음)"
        due = str((t.get("due") or ""))[:10]
        label = title if not due else f"{title}  · {due}"
        note = str(t.get("notes") or "").strip()
        row = st.columns([8.2, 1.15], gap="small", vertical_alignment="top")
        open_this = edit_id == tid
        with row[0]:
            if st.checkbox(label, value=False, key=f"gc_td_{tid}"):
                st.session_state["_gc_pending_done"] = tid
                st.rerun(scope="fragment")
            if note and not open_this:
                st.markdown(
                    "<div style='font-size:12px;color:#5f6368;line-height:1.35;"
                    "margin:-6px 0 6px 1.7rem;white-space:pre-wrap'>"
                    f"{html.escape(note)}</div>",
                    unsafe_allow_html=True,
                )
            if open_this:
                st.text_input(
                    "메모 입력",
                    value=note,
                    key=f"gc_note_{tid}",
                    label_visibility="collapsed",
                    placeholder="할 일 아래 메모",
                )
        with row[1]:
            if st.button(
                "닫기" if open_this else "메모",
                key=f"gc_ne_{tid}",
                type="tertiary",
                width="content",
            ):
                if open_this:
                    st.session_state["_gc_pending_note"] = {
                        "id": tid,
                        "notes": st.session_state.get(f"gc_note_{tid}") or "",
                    }
                    st.session_state["_gc_note_edit"] = ""
                else:
                    st.session_state["_gc_note_edit"] = tid
                st.rerun(scope="fragment")
    add_row = st.columns([6.4, 1.1, 1.15], gap="small", vertical_alignment="bottom")
    with add_row[0]:
        nt = st.text_input(
            "새 할 일",
            key="gc_new_task",
            placeholder="할 일 제목",
            label_visibility="collapsed",
        )
    with add_row[1]:
        if st.button(
            "메모",
            key="gc_toggle_new_note",
            type="tertiary",
            width="content",
        ):
            st.session_state["_gc_show_new_note"] = not bool(
                st.session_state.get("_gc_show_new_note")
            )
            st.rerun(scope="fragment")
    with add_row[2]:
        add_clicked = st.button("추가", key="gc_add_task", width="content")
    nn = ""
    if st.session_state.get("_gc_show_new_note"):
        nn = st.text_input(
            "새 메모",
            key="gc_new_note",
            placeholder="할 일 아래 메모",
            label_visibility="collapsed",
        )
    if add_clicked:
        title = (nt or "").strip()
        if len(title) < 1:
            st.error("제목을 입력하세요.")
        else:
            st.session_state["_gc_pending_new"] = {
                "title": title,
                "notes": nn or "",
            }
            st.session_state["_gc_show_new_note"] = False
            st.rerun(scope="fragment")


def _gc_work_panel() -> None:
    today = datetime.now(_TZ).date()
    y = int(st.session_state.get("gc_year") or today.year)
    m = int(st.session_state.get("gc_month") or today.month)
    bar = st.columns([1.35, 6.2])
    with bar[0]:
        if st.button("연동실행", type="primary", key="gc_sync_run"):
            try:
                n_ev, n_tk = _gc_sync_google(y, m)
                st.session_state["_gc_sync_msg"] = (
                    f"연동했습니다. 일정 {n_ev}건 · 할 일 {n_tk}건"
                )
            except Exception as e:
                st.session_state["_gc_sync_msg"] = f"연동 실패: {e}"
            st.rerun()
    with bar[1]:
        msg = st.session_state.pop("_gc_sync_msg", None)
        if msg:
            st.caption(msg)
    left, right = st.columns([3.15, 1.12], gap="small")
    with left:
        _gc_cal_panel()
    with right:
        _gc_task_panel()


def render_google_calendar_tab(
    df=None, latest_update_str: str = ""
) -> None:
    """상단 담당자·거래처·납품줄 + 구글 월간 달력·할 일."""
    try:
        import visit_calendar_tab as _vc

        _vc.render_visit_delivery_header(df, latest_update_str)
    except Exception as e:
        st.error(f"납품 달력을 표시하지 못했습니다: {e}")
    _consume_oauth_redirect()
    cid, csec = load_oauth_client()
    if not (cid and csec):
        _setup_expander()
    connected = bool(access_token())
    if not connected:
        if cid and csec:
            st.session_state.setdefault("_gc_oauth_state", secrets.token_urlsafe(16))
            st.link_button(
                "구글 계정 연결",
                auth_url(cid, st.session_state["_gc_oauth_state"]),
                type="primary",
                width="stretch",
            )
        return
    _gc_work_panel()
