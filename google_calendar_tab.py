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
    "https://www.googleapis.com/auth/gmail.modify",
)
_AUTH = "https://accounts.google.com/o/oauth2/v2/auth"
_TOKEN = "https://oauth2.googleapis.com/token"
_CAL_API = "https://www.googleapis.com/calendar/v3"
_TASK_API = "https://tasks.googleapis.com/tasks/v1"
_GMAIL_API = "https://gmail.googleapis.com/gmail/v1"
GC_GMAIL_LABEL = "신일가스"
GC_GMAIL_PERSONAL = "개인업무"
GC_GMAIL_INBOX = "INBOX"
GC_MAIL_CHIP_TITLE = "메일"
GC_GMAIL_SYNC = os.path.join(GC_DIR, "gmail_shinil_sync.json")
GC_GMAIL_SYNC_HOME = os.path.join(_HOME_GC, "gmail_shinil_sync.json")
GC_EV_CACHE = os.path.join(GC_DIR, "events_cache.json")
GC_EV_CACHE_HOME = os.path.join(_HOME_GC, "gc_events_cache.json")
GC_CALS_CACHE = os.path.join(GC_DIR, "cals_cache.json")
GC_CALS_CACHE_HOME = os.path.join(_HOME_GC, "gc_cals_cache.json")
GC_MAIL_SYNC_BATCH = 80
_CLOUD_REDIRECT = "https://office-g8ryabkapprkpjmfwa5aypw.streamlit.app/"
_EV_MEM: dict[str, dict[str, list[dict]]] = {}
_EV_BG = ThreadPoolExecutor(max_workers=2, thread_name_prefix="gc-ev")
_EV_BG_FUTS: dict[str, Any] = {}


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


def mini_month_weeks(year: int, month: int) -> list[list[date]]:
    """월요일 시작 주. 구글 왼쪽 작은 달력과 같다."""
    return calendar.Calendar(firstweekday=0).monthdatescalendar(year, month)


def iso_week_label(week: list[date]) -> int:
    return int(week[3].isocalendar()[1])


def _gc_is_holiday_day(items: list[dict] | None) -> bool:
    return any(str((ev or {}).get("cal_id") or "") == GC_KR_HOLIDAY for ev in (items or []))


def _gc_cached_holiday_keys(year: int, month: int) -> set[str]:
    """이미 불러온 일정만 본다. 작은 달력 월 이동에서 API를 다시 치지 않는다."""
    weeks = mini_month_weeks(year, month)
    months = {(d.year, d.month) for week in weeks for d in week}
    box = _gc_api_box()
    out: set[str] = set()
    for yy, mm in months:
        cached = box.get(_ev_key(yy, mm)) or _EV_MEM.get(_ev_key(yy, mm)) or {}
        for key, items in cached.items():
            if _gc_is_holiday_day(items):
                out.add(str(key))
    return out


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
    if err == "unauthorized_client":
        return "이 OAuth 클라이언트로 토큰을 갱신할 수 없습니다. 다시 연결해 주세요."
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
    rt = str(tok.get("refresh_token") or "").strip()
    if not rt:
        raise RuntimeError("재로그인 필요")
    r = requests.post(
        _TOKEN,
        data={
            "refresh_token": rt,
            "client_id": str(client_id or "").strip(),
            "client_secret": str(client_secret or "").strip(),
            "grant_type": "refresh_token",
        },
        timeout=30,
    )
    if not r.ok:
        raise RuntimeError(_token_error_text(r))
    data = r.json() if r.content else {}
    if not isinstance(data, dict):
        data = {}
    new_rt = str(data.get("refresh_token") or "").strip()
    data["refresh_token"] = new_rt or rt
    data["_obtained"] = datetime.now(timezone.utc).isoformat()
    if not data.get("expires_in"):
        data["expires_in"] = 3500
    return data


def _expired(tok: dict) -> bool:
    if not str((tok or {}).get("access_token") or "").strip():
        return True
    exp = tok.get("expires_in")
    obtained = tok.get("_obtained")
    if not exp or not obtained:
        return True
    try:
        t0 = datetime.fromisoformat(str(obtained).replace("Z", "+00:00"))
        return datetime.now(timezone.utc) > t0 + timedelta(seconds=int(exp) - 60)
    except Exception:
        return True


def _is_reauth_error(exc: BaseException) -> bool:
    resp = getattr(exc, "response", None)
    err = ""
    try:
        if resp is not None:
            err = str((resp.json() or {}).get("error") or "")
    except Exception:
        err = ""
    blob = (err + " " + str(exc)).lower()
    return err in ("invalid_grant", "invalid_client", "unauthorized_client") or any(
        s in blob
        for s in (
            "재로그인",
            "만료",
            "invalid_grant",
            "invalid_client",
            "unauthorized_client",
        )
    )


def _auth_error_message(exc: BaseException) -> str:
    resp = getattr(exc, "response", None)
    if resp is not None:
        try:
            return _token_error_text(resp)
        except Exception:
            pass
    msg = str(exc or "").strip()
    if "oauth2.googleapis.com" in msg or "bad request" in msg.lower():
        return "구글 인증이 만료되었습니다. 다시 연결해 주세요."
    return msg or "구글 인증이 만료되었습니다. 다시 연결해 주세요."


def _handle_refresh_failure(exc: BaseException) -> None:
    if _is_reauth_error(exc):
        clear_token()
        try:
            st.session_state["_gc_need_reauth"] = True
        except Exception:
            pass


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
        except Exception as e:
            _handle_refresh_failure(e)
            if _is_reauth_error(e):
                return ""
    return at


def _retry_after_refresh(request_fn):
    cid, csec = load_oauth_client()
    tok_d = load_token()
    if not (cid and csec and tok_d.get("refresh_token")):
        raise RuntimeError("구글 연결이 필요합니다.")
    try:
        tok_d = refresh_token(cid, csec, tok_d)
        save_token(tok_d)
    except Exception as e:
        _handle_refresh_failure(e)
        raise RuntimeError(_auth_error_message(e)) from e
    return request_fn(str(tok_d.get("access_token") or ""))


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

        def _once(at: str):
            return requests.get(
                url,
                params=params or {},
                headers={"Authorization": f"Bearer {at}"},
                timeout=30,
            )

        r = _retry_after_refresh(_once)
    if r.status_code == 401:
        raise RuntimeError("구글 인증이 만료되었습니다. 다시 연결해 주세요.")
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
    if r.status_code == 401:

        def _once(at: str):
            return requests.request(
                method,
                url,
                json=body or {},
                headers={"Authorization": f"Bearer {at}", "Content-Type": "application/json"},
                timeout=30,
            )

        r = _retry_after_refresh(_once)
    if r.status_code == 401:
        raise RuntimeError("구글 인증이 만료되었습니다. 다시 연결해 주세요.")
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
            "fields": "items(id,summary,description,start,end,extendedProperties)",
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


class GmailScopeError(RuntimeError):
    """토큰에 Gmail 읽기 권한이 없을 때."""


def gmail_thread_url(thread_id: str) -> str:
    tid = str(thread_id or "").strip()
    if not tid:
        return "https://mail.google.com/mail/u/0/#inbox"
    return "https://mail.google.com/mail/u/0/#all/" + urllib.parse.quote(tid, safe="")


def gmail_received_date(internal_ms: Any) -> date | None:
    try:
        ms = int(internal_ms)
    except (TypeError, ValueError):
        return None
    if ms <= 0:
        return None
    return datetime.fromtimestamp(ms / 1000.0, tz=_TZ).date()


def _gmail_header(headers: list | None, name: str) -> str:
    want = str(name or "").lower()
    for h in headers or []:
        if str((h or {}).get("name") or "").lower() == want:
            return str((h or {}).get("value") or "").strip()
    return ""


def mail_event_title(_subject: str = "") -> str:
    return GC_MAIL_CHIP_TITLE


def is_mail_calendar_event(ev: dict | None) -> bool:
    title = str((ev or {}).get("title") or (ev or {}).get("summary") or "").strip()
    desc = str((ev or {}).get("description") or "")
    if title == GC_MAIL_CHIP_TITLE or title.startswith("[메일]"):
        return True
    return "메일 열기:" in desc or "mail.google.com/mail" in desc


def mail_subject_from_event(ev: dict | None) -> str:
    title = str((ev or {}).get("title") or (ev or {}).get("summary") or "").strip()
    if title.startswith("[메일]"):
        return title[len("[메일]") :].strip()
    if title and title != GC_MAIL_CHIP_TITLE:
        return title
    note = str((ev or {}).get("description") or "")
    first = (note.splitlines() or [""])[0].strip()
    if first.startswith("제목:"):
        return first[3:].strip()
    if first and not first.startswith("보낸사람:") and not first.startswith("받은날:") and not first.startswith("메일 열기:"):
        return first
    return ""


def calendar_chip_label(ev: dict | None) -> str:
    """칸에는 짧게. 메일은 '메일'만."""
    if is_mail_calendar_event(ev):
        return GC_MAIL_CHIP_TITLE
    title = str((ev or {}).get("title") or "").strip()
    hhmm = str((ev or {}).get("hhmm") or "").strip()
    lab = ((hhmm + " ") if hhmm else "") + title
    if len(lab) > 14:
        lab = lab[:13] + "…"
    return lab or "(일정)"


def mail_event_notes(
    *,
    subject: str,
    sender: str,
    received: date,
    snippet: str,
    thread_id: str,
) -> str:
    url = gmail_thread_url(thread_id)
    sub = str(subject or "").strip() or "(제목 없음)"
    snip = " ".join(str(snippet or "").split())
    if len(snip) > 280:
        snip = snip[:277] + "…"
    lines = [
        f"제목: {sub}",
        f"보낸사람: {str(sender or '').strip() or '-'}",
        f"받은날: {received.isoformat()}",
        "",
    ]
    if snip:
        lines.extend([snip, ""])
    lines.append(f"메일 열기: {url}")
    return "\n".join(lines)


def gmail_inbox_query_since(since: date) -> str:
    prev = since - timedelta(days=1)
    return f"in:inbox after:{prev.year}/{prev.month:02d}/{prev.day:02d}"


def match_gmail_label(labels: list[dict], want: str = GC_GMAIL_LABEL) -> dict:
    want = str(want or GC_GMAIL_LABEL).strip()
    items = [x for x in (labels or []) if isinstance(x, dict) and str(x.get("id") or "")]
    exact = [x for x in items if str(x.get("name") or "").strip() == want]
    if exact:
        return exact[0]
    leaf = [
        x
        for x in items
        if str(x.get("name") or "").strip().split("/")[-1] == want
    ]
    if leaf:
        return leaf[0]
    part = [x for x in items if want in str(x.get("name") or "")]
    return part[0] if part else {}


def token_has_gmail_modify() -> bool:
    tok = load_token()
    return "gmail.modify" in str((tok or {}).get("scope") or "")


def _gmail_error_blob(exc: BaseException) -> str:
    resp = getattr(exc, "response", None)
    text = ""
    try:
        text = str(getattr(resp, "text", "") or "")
    except Exception:
        text = ""
    return (text + " " + str(exc)).lower()


def _gmail_error_is_api_off(exc: BaseException) -> bool:
    blob = _gmail_error_blob(exc)
    return "has not been used" in blob or "it is disabled" in blob or "access not configured" in blob


def _gmail_error_is_scope(exc: BaseException) -> bool:
    resp = getattr(exc, "response", None)
    code = 0
    try:
        code = int(getattr(resp, "status_code", 0) or 0)
    except Exception:
        code = 0
    blob = _gmail_error_blob(exc)
    if _gmail_error_is_api_off(exc):
        return False
    return code in (401, 403) and (
        "insufficient" in blob
        or "access_token_scope" in blob
        or "insufficientpermissions" in blob
        or "forbidden" in blob
    )


def _gmail_raise(exc: BaseException, *, write: bool = False) -> None:
    if _gmail_error_is_api_off(exc):
        raise RuntimeError(
            "Google Cloud에서 Gmail API를 켠 뒤 잠시 기다려 주세요."
        ) from exc
    if _gmail_error_is_scope(exc):
        raise GmailScopeError(
            "Gmail 폴더 이동 권한이 없습니다. 아래 버튼으로 구글 계정을 다시 연결해 주세요."
            if write
            else "Gmail 권한이 없습니다. 아래 버튼으로 구글 계정을 다시 연결해 주세요."
        ) from exc
    raise exc


def _gmail_get(path: str, params: dict | None = None) -> dict:
    try:
        return api_get(f"{_GMAIL_API}/users/me/{path.lstrip('/')}", params)
    except Exception as e:
        _gmail_raise(e, write=False)
        raise


def _gmail_json(method: str, path: str, body: dict | None = None) -> dict:
    try:
        return api_json(method, f"{_GMAIL_API}/users/me/{path.lstrip('/')}", body)
    except Exception as e:
        _gmail_raise(e, write=True)
        raise


def _gc_render_gmail_reconnect(*, key: str) -> None:
    cid, _csec = load_oauth_client()
    if not cid:
        return
    st.session_state.setdefault("_gc_oauth_state", secrets.token_urlsafe(16))
    st.caption(
        "Google에서 Gmail 수정 권한을 허용한 뒤, 반드시 이 화면으로 돌아와야 저장됩니다."
    )
    st.link_button(
        "Gmail 권한 다시 연결",
        auth_url(cid, st.session_state["_gc_oauth_state"]),
        type="primary",
        width="stretch",
        key=key,
    )
    if st.button("연결 해제 후 처음부터", key=f"{key}_wipe"):
        clear_token()
        st.session_state.pop("_gc_need_gmail", None)
        st.session_state.pop("_gc_oauth_code_used", None)
        st.rerun()


def gmail_ids_from_event(ev: dict | None) -> dict[str, str]:
    ev = ev or {}
    mid = str(ev.get("gmail_msg") or "").strip()
    tid = str(ev.get("gmail_thread") or "").strip()
    desc = str(ev.get("description") or "")
    if not tid:
        for part in desc.replace("\n", " ").split():
            if "#all/" in part:
                tid = part.split("#all/", 1)[-1].strip().rstrip(").,]")
                break
    return {"msg": mid, "thread": tid}


def gmail_move_body(label_id: str) -> dict[str, list[str]]:
    return {
        "addLabelIds": [str(label_id)],
        "removeLabelIds": [GC_GMAIL_INBOX],
    }


def ensure_gmail_label(name: str = GC_GMAIL_LABEL) -> dict:
    want = str(name or GC_GMAIL_LABEL).strip() or GC_GMAIL_LABEL
    lab = match_gmail_label(list_gmail_labels(), want)
    if lab:
        return lab
    created = _gmail_json(
        "POST",
        "labels",
        {
            "name": want,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        },
    )
    return created if isinstance(created, dict) else {}


def gmail_move_inbox_to_label(
    *, thread_id: str = "", msg_id: str = "", label_name: str = GC_GMAIL_LABEL
) -> dict:
    if not token_has_gmail_modify():
        raise GmailScopeError(
            "Gmail 폴더 이동 권한이 없습니다. 아래 버튼으로 구글 계정을 다시 연결해 주세요."
        )
    lab = ensure_gmail_label(label_name)
    lid = str(lab.get("id") or "").strip()
    if not lid:
        raise RuntimeError(f"Gmail에 '{label_name}' 폴더를 만들지 못했습니다.")
    body = gmail_move_body(lid)
    tid = str(thread_id or "").strip()
    mid = str(msg_id or "").strip()
    if tid:
        return _gmail_json("POST", f"threads/{urllib.parse.quote(tid, safe='')}/modify", body)
    if mid:
        return _gmail_json("POST", f"messages/{urllib.parse.quote(mid, safe='')}/modify", body)
    raise RuntimeError("옮길 메일을 찾지 못했습니다.")


def gmail_trash_path(*, thread_id: str = "", msg_id: str = "") -> str:
    tid = str(thread_id or "").strip()
    mid = str(msg_id or "").strip()
    if tid:
        return f"threads/{urllib.parse.quote(tid, safe='')}/trash"
    if mid:
        return f"messages/{urllib.parse.quote(mid, safe='')}/trash"
    raise RuntimeError("지울 메일을 찾지 못했습니다.")


def gmail_trash_mail(*, thread_id: str = "", msg_id: str = "") -> dict:
    if not token_has_gmail_modify():
        raise GmailScopeError(
            "Gmail 삭제 권한이 없습니다. 아래 버튼으로 구글 계정을 다시 연결해 주세요."
        )
    return _gmail_json("POST", gmail_trash_path(thread_id=thread_id, msg_id=msg_id), {})


def _drop_gmail_sync_id(msg_id: str) -> None:
    mid = str(msg_id or "").strip()
    if not mid:
        return
    state = load_gmail_sync_state()
    ids = dict(state.get("ids") or {})
    if mid in ids:
        ids.pop(mid, None)
        state["ids"] = ids
        save_gmail_sync_state(state)


def list_gmail_labels() -> list[dict]:
    data = _gmail_get("labels")
    return list(data.get("labels") or [])


def list_gmail_message_ids(
    label_id: str = GC_GMAIL_INBOX, *, limit: int = 500, query: str = ""
) -> list[dict]:
    lid = str(label_id or GC_GMAIL_INBOX).strip() or GC_GMAIL_INBOX
    out: list[dict] = []
    token = ""
    cap = max(1, min(int(limit or 500), 2000))
    while len(out) < cap:
        params: dict[str, Any] = {
            "labelIds": lid,
            "maxResults": str(min(100, cap - len(out))),
        }
        if str(query or "").strip():
            params["q"] = str(query).strip()
        if token:
            params["pageToken"] = token
        data = _gmail_get("messages", params)
        for it in data.get("messages") or []:
            mid = str((it or {}).get("id") or "")
            if mid:
                out.append(
                    {
                        "id": mid,
                        "threadId": str((it or {}).get("threadId") or mid),
                    }
                )
            if len(out) >= cap:
                break
        token = str(data.get("nextPageToken") or "")
        if not token:
            break
    return out


def gmail_message_meta(msg_id: str) -> dict:
    mid = str(msg_id or "").strip()
    if not mid:
        return {}
    data = _gmail_get(
        f"messages/{urllib.parse.quote(mid, safe='')}",
        {
            "format": "metadata",
            "metadataHeaders": ["Subject", "From", "Date"],
            "fields": "id,threadId,internalDate,snippet,payload/headers",
        },
    )
    headers = list(((data.get("payload") or {}).get("headers")) or [])
    received = gmail_received_date(data.get("internalDate"))
    return {
        "id": str(data.get("id") or mid),
        "threadId": str(data.get("threadId") or mid),
        "subject": _gmail_header(headers, "Subject"),
        "sender": _gmail_header(headers, "From"),
        "snippet": str(data.get("snippet") or ""),
        "day": received,
    }


def _gmail_sync_paths() -> list[str]:
    return [GC_GMAIL_SYNC_HOME, GC_GMAIL_SYNC]


def load_gmail_sync_state() -> dict:
    for path in _gmail_sync_paths():
        data = _read_json(path)
        if data:
            return data
    return {"ids": {}}


def save_gmail_sync_state(data: dict) -> None:
    payload = dict(data or {})
    if not isinstance(payload.get("ids"), dict):
        payload["ids"] = {}
    for path in _gmail_sync_paths():
        _write_json(path, payload)


def insert_mail_day_event(
    cal_id: str,
    *,
    title: str,
    day: date,
    notes: str,
    msg_id: str,
    thread_id: str,
) -> dict:
    cid = str(cal_id or "").strip() or "primary"
    if cid == GC_KR_HOLIDAY:
        raise RuntimeError("공휴일 캘린더에는 일정을 넣을 수 없습니다.")
    url = gmail_thread_url(thread_id)
    body: dict[str, Any] = {
        "summary": mail_event_title(title),
        "description": str(notes or "").strip(),
        "start": {"date": day.isoformat()},
        "end": {"date": (day + timedelta(days=1)).isoformat()},
        "source": {"title": "Gmail", "url": url},
        "extendedProperties": {
            "private": {
                "shinilMail": "1",
                "gmailMsgId": str(msg_id or "").strip(),
                "gmailThread": str(thread_id or "").strip(),
            }
        },
    }
    return api_json(
        "POST",
        f"{_CAL_API}/calendars/{urllib.parse.quote(cid, safe='')}/events",
        body,
    )


def _list_synced_mail_ids_from_calendar(cal_id: str) -> set[str]:
    cid = str(cal_id or "").strip() or "primary"
    found: set[str] = set()
    token = ""
    for _ in range(20):
        params: dict[str, Any] = {
            "privateExtendedProperty": "shinilMail=1",
            "maxResults": "250",
            "singleEvents": "true",
            "fields": "items(extendedProperties),nextPageToken",
        }
        if token:
            params["pageToken"] = token
        try:
            data = api_get(
                f"{_CAL_API}/calendars/{urllib.parse.quote(cid, safe='')}/events",
                params,
            )
        except Exception:
            break
        for it in data.get("items") or []:
            priv = ((it or {}).get("extendedProperties") or {}).get("private") or {}
            mid = str(priv.get("gmailMsgId") or "").strip()
            if mid:
                found.add(mid)
        token = str(data.get("nextPageToken") or "")
        if not token:
            break
    return found


def _primary_writable_cal_id(cals: list[dict] | None = None) -> str:
    items = _writable_calendars(cals or [])
    for c in items:
        if c.get("primary"):
            return str(c.get("id") or "primary")
    if items:
        return str(items[0].get("id") or "primary")
    return "primary"


def _retitle_mail_calendar_events(cal_id: str) -> int:
    """이미 넣은 메일 일정의 칸 제목을 '메일'로 줄인다. 본문은 메모에 둔다."""
    cid = str(cal_id or "").strip() or "primary"
    n = 0
    token = ""
    for _ in range(20):
        params: dict[str, Any] = {
            "privateExtendedProperty": "shinilMail=1",
            "maxResults": "250",
            "singleEvents": "true",
            "fields": "items(id,summary,description),nextPageToken",
        }
        if token:
            params["pageToken"] = token
        try:
            data = api_get(
                f"{_CAL_API}/calendars/{urllib.parse.quote(cid, safe='')}/events",
                params,
            )
        except Exception:
            break
        for it in data.get("items") or []:
            eid = str((it or {}).get("id") or "")
            summary = str((it or {}).get("summary") or "").strip()
            if not eid or summary == GC_MAIL_CHIP_TITLE:
                continue
            fields: dict[str, Any] = {"summary": GC_MAIL_CHIP_TITLE}
            desc = str((it or {}).get("description") or "")
            sub = mail_subject_from_event(
                {"title": summary, "description": desc}
            )
            if sub and f"제목: {sub}" not in desc:
                fields["description"] = f"제목: {sub}\n" + (desc.lstrip("\n") if desc else "")
            try:
                patch_event(cid, eid, fields)
                n += 1
            except Exception:
                continue
        token = str(data.get("nextPageToken") or "")
        if not token:
            break
    return n


def sync_shinil_gmail_to_calendar(
    *,
    label_name: str = GC_GMAIL_LABEL,
    cal_id: str = "",
    limit: int = GC_MAIL_SYNC_BATCH,
    since: date | None = None,
) -> dict[str, Any]:
    """신일가스 라벨 메일을 받은 날 종일 일정으로 넣고, 칸에는 '메일'만 보여 준다."""
    labels = list_gmail_labels()
    lab = match_gmail_label(labels, label_name)
    if not lab:
        return {
            "added": 0,
            "skipped": 0,
            "remain": 0,
            "label": label_name,
            "error": f"Gmail에 '{label_name}' 폴더(라벨)가 없습니다.",
        }
    lid = str(lab.get("id") or "")
    listed = list_gmail_message_ids(lid, limit=500)
    state = load_gmail_sync_state()
    done = {
        str(k): str(v)
        for k, v in dict(state.get("ids") or {}).items()
        if str(k).strip()
    }
    cid = str(cal_id or "").strip() or _primary_writable_cal_id()
    cal_done = _list_synced_mail_ids_from_calendar(cid)
    for mid in cal_done:
        done.setdefault(mid, "cal")
    pending = [it for it in listed if str(it.get("id") or "") not in done]
    batch = pending[: max(0, int(limit or GC_MAIL_SYNC_BATCH))]
    added = 0
    skipped = len(listed) - len(pending)
    months: set[tuple[int, int]] = set()
    for it in batch:
        mid = str(it.get("id") or "")
        try:
            meta = gmail_message_meta(mid)
        except GmailScopeError:
            raise
        except Exception:
            continue
        day = meta.get("day")
        if not isinstance(day, date):
            continue
        if since is not None and day < since:
            continue
        notes = mail_event_notes(
            subject=str(meta.get("subject") or ""),
            sender=str(meta.get("sender") or ""),
            received=day,
            snippet=str(meta.get("snippet") or ""),
            thread_id=str(meta.get("threadId") or it.get("threadId") or mid),
        )
        try:
            created = insert_mail_day_event(
                cid,
                title=str(meta.get("subject") or ""),
                day=day,
                notes=notes,
                msg_id=mid,
                thread_id=str(meta.get("threadId") or mid),
            )
        except Exception:
            continue
        done[mid] = day.isoformat()
        added += 1
        months.add((day.year, day.month))
        chip = _event_chip(created, {"id": cid, "summary": "", "color": "#1a73e8", "writable": True})
        if chip:
            chip["title"] = GC_MAIL_CHIP_TITLE
            _gc_cache_add_event(day, chip)
    try:
        _retitle_mail_calendar_events(cid)
    except Exception:
        pass
    state["ids"] = done
    state["label"] = str(lab.get("name") or label_name)
    if since is not None:
        state["since"] = since.isoformat()
    state["cal_id"] = cid
    state["at"] = datetime.now(_TZ).isoformat()
    save_gmail_sync_state(state)
    box = _gc_api_box()
    for yy, mm in months:
        key = _ev_key(yy, mm)
        box.pop(key + ":at", None)
        box.pop(key, None)
        _EV_MEM.pop(key, None)
    return {
        "added": added,
        "skipped": skipped,
        "remain": max(0, len(pending) - added),
        "label": str(lab.get("name") or label_name),
        "error": "",
    }


def api_delete(url: str) -> None:
    tok = access_token()
    if not tok:
        raise RuntimeError("구글 연결이 필요합니다.")
    r = requests.delete(url, headers={"Authorization": f"Bearer {tok}"}, timeout=30)
    if r.status_code == 401:

        def _once(at: str):
            return requests.delete(url, headers={"Authorization": f"Bearer {at}"}, timeout=30)

        try:
            r = _retry_after_refresh(_once)
        except Exception:
            return
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


def list_tasks(list_id: str, *, include_completed: bool = False) -> list[dict]:
    data = api_get(
        f"{_TASK_API}/lists/{urllib.parse.quote(list_id, safe='')}/tasks",
        {
            "showCompleted": "true" if include_completed else "false",
            "showHidden": "true" if include_completed else "false",
            "maxResults": "100",
        },
    )
    items = list(data.get("items") or [])
    if not include_completed:
        items = [t for t in items if t.get("status") != "completed"]
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


def uncomplete_task(list_id: str, task_id: str) -> dict:
    return api_json(
        "PATCH",
        f"{_TASK_API}/lists/{urllib.parse.quote(list_id, safe='')}/tasks/{urllib.parse.quote(task_id, safe='')}",
        {"status": "needsAction"},
    )


def _gc_api_box() -> dict:
    return st.session_state.setdefault("_gc_api", {})


def _gc_load_disk_events() -> None:
    if _EV_MEM:
        return
    for path in (GC_EV_CACHE_HOME, GC_EV_CACHE):
        data = _read_json(path)
        if data:
            for key, val in data.items():
                if isinstance(val, dict):
                    _EV_MEM[key] = val
            return


def _gc_save_disk_events() -> None:
    payload = {k: v for k, v in _EV_MEM.items() if isinstance(v, dict)}
    if not payload:
        return
    _write_json(GC_EV_CACHE_HOME, payload)
    _write_json(GC_EV_CACHE, payload)


def _gc_save_disk_cals(cals: list[dict]) -> None:
    blob = {"cals": list(cals or [])}
    _write_json(GC_CALS_CACHE_HOME, blob)
    _write_json(GC_CALS_CACHE, blob)


def _gc_cached_cals() -> list[dict]:
    box = _gc_api_box()
    if box.get("cals") is None:
        for path in (GC_CALS_CACHE_HOME, GC_CALS_CACHE):
            items = _read_json(path).get("cals")
            if isinstance(items, list) and items:
                box["cals"] = items
                break
        if box.get("cals") is None:
            box["cals"] = list_calendars()
            _gc_save_disk_cals(box["cals"])
    return list(box["cals"] or [])


def _gc_cached_tlists() -> list[dict]:
    box = _gc_api_box()
    if box.get("tlists") is None:
        box["tlists"] = list_tasklists()
    return list(box["tlists"] or [])


def _ev_key(year: int, month: int) -> str:
    return f"ev:{year}-{month:02d}"


def _gc_shift_month(year: int, month: int, delta: int) -> tuple[int, int]:
    m0 = year * 12 + (month - 1) + int(delta)
    return m0 // 12, m0 % 12 + 1


def _gc_prefetch_one(cals: list[dict], year: int, month: int) -> Any:
    """옆달 일정을 백그라운드에서 받아, ‹ › 클릭 때 API 대기를 없앤다."""
    key = _ev_key(year, month)
    if key in _EV_MEM:
        return None
    fut = _EV_BG_FUTS.get(key)
    if fut is not None and not fut.done():
        return fut
    snap = list(cals or [])

    def _job() -> dict[str, list[dict]]:
        data = _events_by_day(snap, year, month) if snap else {}
        _EV_MEM[key] = data
        _gc_save_disk_events()
        return data

    fut = _EV_BG.submit(_job)
    _EV_BG_FUTS[key] = fut
    return fut


def _gc_schedule_prefetch(cals: list[dict], year: int, month: int) -> None:
    if not cals:
        return
    for delta in (-1, 1):
        yy, mm = _gc_shift_month(year, month, delta)
        _gc_prefetch_one(cals, yy, mm)


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
    _gc_load_disk_events()
    cached = box.get(key)
    if cached is None:
        cached = _EV_MEM.get(key)
    if cached is not None and not force:
        box[key] = cached
        if key + ":at" not in box:
            box[key + ":at"] = now
        return cached
    fresh: dict[str, list[dict]] = {}
    fut = _EV_BG_FUTS.get(key)
    if fut is not None:
        try:
            fresh = fut.result(timeout=25) or {}
        except Exception:
            fresh = _events_by_day(cals, year, month) if cals else {}
    else:
        fresh = _events_by_day(cals, year, month) if cals else {}
    data = _merge_event_days(cached or {}, fresh) if cached else fresh
    box[key] = data
    box[key + ":at"] = now
    _EV_MEM[key] = data
    _gc_save_disk_events()
    return data


def _gc_sync_google(year: int, month: int) -> tuple[int, int]:
    """대시보드↔구글 캘린더·할 일을 한 번에 맞춘다. (일정 수, 할 일 수)"""
    box = _gc_api_box()
    try:
        box["cals"] = list_calendars()
        _gc_save_disk_cals(box["cals"])
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
    _gc_save_disk_events()
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
    _gc_save_disk_events()


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
    _gc_save_disk_events()


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
    _gc_save_disk_events()


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
        remote = [
            _slim_task(x)
            for x in (list_tasks(list_id, include_completed=True) if list_id else [])
        ]
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
    priv = ((it.get("extendedProperties") or {}).get("private") or {})
    return {
        "title": it.get("summary") or "(제목 없음)",
        "color": cal.get("color") or "#7b8794",
        "cal": cal.get("summary") or "",
        "hhmm": hhmm,
        "day": key,
        "description": str(it.get("description") or ""),
        "id": str(it.get("id") or ""),
        "cal_id": str(cal.get("id") or ""),
        "gmail_msg": str(priv.get("gmailMsgId") or ""),
        "gmail_thread": str(priv.get("gmailThread") or ""),
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
                label = html.escape(calendar_chip_label(ev))
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
.gc-chip {{ font-size:10px; line-height:1.2; padding:1px 2px 1px 4px; margin:1px 0; border-radius:2px;
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


def _gc_mini_set_month(day: date) -> None:
    st.session_state["gc_mini_year"] = day.year
    st.session_state["gc_mini_month"] = day.month


def _gc_set_view_month(day: date) -> None:
    first = date(day.year, day.month, 1)
    st.session_state["gc_year"] = first.year
    st.session_state["gc_month"] = first.month
    st.session_state["_gc_pending_month"] = first


def _gc_set_sel_event(day: date, event_id: str, chip: str = "") -> None:
    st.session_state["_gc_sel_day"] = day
    st.session_state["_gc_sel_event"] = str(event_id or "")
    st.session_state["_gc_sel_chip"] = str(chip or "")
    st.session_state["_gc_memo_open"] = True


def _gc_close_memo_pop() -> None:
    st.session_state["_gc_memo_open"] = False
    st.session_state.pop("_gc_sel_chip", None)


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


def _render_mini_month(
    year: int,
    month: int,
    today: date,
    selected: date | None = None,
    holidays: set[str] | None = None,
) -> None:
    """구글 왼쪽 작은 달력. 날짜 칸 키(gc_day_/gc_cell_)는 바꾸지 않는다."""
    weeks = mini_month_weeks(year, month)
    hols = holidays or set()
    st.markdown(
        """
<style>
div[class*="st-key-gc_mini"] { max-width: 248px !important; }
div[class*="st-key-gc_mini"] [data-testid="stVerticalBlock"] { gap: 2px !important; }
div[class*="st-key-gc_mini"] [data-testid="stHorizontalBlock"] {
  gap: 0 !important; margin: 0 !important; flex-wrap: nowrap !important;
}
div[class*="st-key-gc_mini"] [data-testid="stColumn"] {
  padding: 0 !important; min-width: 0 !important;
}
div[class*="st-key-gc_mini"] .gc-mini-ym {
  font-size: 13px !important; font-weight: 600 !important;
  color: #3c4043 !important; margin: 0 !important; padding: 4px 2px 6px;
}
div[class*="st-key-gc_mini"] .gc-mini-dow,
div[class*="st-key-gc_mini"] .gc-mini-wn {
  font-size: 10px !important; color: #70757a !important;
  text-align: center !important; margin: 0 !important; line-height: 1.6 !important;
}
div[class*="st-key-gc_mini"] .gc-mini-dow-pink { color: #d93025 !important; }
div[class*="st-key-gc_mini"] .gc-mini-wn { color: #9aa0a6 !important; }
div[class*="st-key-gc_mini"] div[class*="st-key-gc_mp"],
div[class*="st-key-gc_mini"] div[class*="st-key-gc_mn"] { width: 1.4rem !important; }
div[class*="st-key-gc_mini"] button {
  min-height: 22px !important; height: 22px !important;
  width: 22px !important; min-width: 22px !important;
  padding: 0 !important; font-size: 11px !important; font-weight: 400 !important;
  border: none !important; border-radius: 50% !important;
  background: transparent !important; box-shadow: none !important;
  color: #3c4043 !important; line-height: 22px !important;
}
div[class*="st-key-gc_mdt_"] button {
  background: #1a73e8 !important; color: #fff !important; font-weight: 500 !important;
}
div[class*="st-key-gc_mdp_"] button {
  background: #fdecea !important; color: #d93025 !important;
}
div[class*="st-key-gc_mdq_"] button {
  background: #fdecea !important; color: #e8a0a0 !important;
}
div[class*="st-key-gc_mdo_"] button { color: #bdc1c6 !important; }
div[class*="st-key-gc_mds_"] button { background: #e8f0fe !important; color: #174ea6 !important; }
div[class*="st-key-gc_mp"] button, div[class*="st-key-gc_mn"] button {
  min-height: 1.4rem !important; height: 1.4rem !important;
  width: 1.4rem !important; color: #5f6368 !important; font-size: 14px !important;
}
</style>
""",
        unsafe_allow_html=True,
    )
    with st.container(key="gc_mini"):
        hd = st.columns([3.4, 0.7, 0.7], vertical_alignment="center")
        with hd[0]:
            st.markdown(
                f'<div class="gc-mini-ym">{html.escape(f"{year}년 {month}월")}</div>',
                unsafe_allow_html=True,
            )
        with hd[1]:
            st.button(
                "‹",
                key="gc_mp",
                type="tertiary",
                on_click=_gc_mini_set_month,
                args=(date(year, month, 1) - timedelta(days=1),),
            )
        with hd[2]:
            st.button(
                "›",
                key="gc_mn",
                type="tertiary",
                on_click=_gc_mini_set_month,
                args=(date(year, month, 28) + timedelta(days=8),),
            )
        labs = st.columns([0.72] + [1] * 7, gap=None)
        with labs[0]:
            st.markdown('<div class="gc-mini-wn"></div>', unsafe_allow_html=True)
        for i, lab in enumerate(("월", "화", "수", "목", "금", "토", "일")):
            with labs[i + 1]:
                cls = "gc-mini-dow gc-mini-dow-pink" if i >= 5 else "gc-mini-dow"
                st.markdown(
                    f'<div class="{cls}">{lab}</div>', unsafe_allow_html=True
                )
        for week in weeks:
            row = st.columns([0.72] + [1] * 7, gap=None)
            with row[0]:
                st.markdown(
                    f'<div class="gc-mini-wn">{int(week[3].isocalendar()[1])}</div>',
                    unsafe_allow_html=True,
                )
            for i, d in enumerate(week):
                stamp = d.isoformat()
                is_pink = d.weekday() >= 5 or stamp in hols
                if d == today:
                    wrap = f"gc_mdt_{stamp}"
                elif selected is not None and d == selected:
                    wrap = f"gc_mds_{stamp}"
                elif is_pink and d.month != month:
                    wrap = f"gc_mdq_{stamp}"
                elif is_pink:
                    wrap = f"gc_mdp_{stamp}"
                elif d.month != month:
                    wrap = f"gc_mdo_{stamp}"
                else:
                    wrap = f"gc_mdi_{stamp}"
                with row[i + 1]:
                    with st.container(key=wrap):
                        st.button(
                            str(d.day),
                            key=f"gc_md_{stamp}",
                            type="tertiary",
                            on_click=_gc_set_sel_day,
                            args=(d,),
                        )


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
div[class*="st-key-gc_calgrid"] [data-testid="stColumn"]:has(.gc-js-sel),
div[class*="st-key-gc_cell_"].gc-js-sel,
div.gc-js-sel {
  background: #e8f0fe !important;
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
  top: 50% !important;
  right: 1px !important;
  transform: translateY(-50%) !important;
  z-index: 3 !important;
  width: 1.2rem !important;
  height: 1.2rem !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
}
div[class*="st-key-gc_xd_"] button {
  min-height: 1.2rem !important;
  height: 1.2rem !important;
  min-width: 1.2rem !important;
  width: 1.2rem !important;
  padding: 0 !important;
  font-size: 13px !important;
  line-height: 1.2rem !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
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
div[class*="st-key-gc_chiprow_"] div[class*="st-key-gc_evpick_"] button,
div[class*="st-key-gc_evsel_"] div[class*="st-key-gc_evpick_"] button {
  padding-right: 16px !important;
}
div[class*="st-key-gc_evpick_"] button {
  min-height: 1.2rem !important;
  height: 1.2rem !important;
  width: 100% !important;
  padding: 0 4px !important;
  font-size: 10px !important;
  font-weight: 400 !important;
  border: none !important;
  outline: none !important;
  border-radius: 6px !important;
  background: #e8f0fe !important;
  color: #174ea6 !important;
  box-shadow: none !important;
  justify-content: flex-start !important;
  text-align: left !important;
  white-space: nowrap !important;
  overflow: hidden !important;
}
div[class*="st-key-gc_evpick_"] *,
div[class*="st-key-gc_evpick_"] button p,
div[class*="st-key-gc_evpick_"] button span,
div[class*="st-key-gc_evpick_"] [data-testid="stMarkdownContainer"],
div[class*="st-key-gc_evpick_"] [data-testid="stMarkdownContainer"] p,
div[class*="st-key-gc_evpick_"] [data-testid^="stBaseButton"] p {
  font-size: 10px !important;
  line-height: 1.2 !important;
  font-weight: 400 !important;
}
div[class*="st-key-gc_evsel_"] div[class*="st-key-gc_evpick_"] button {
  background: #d2e3fc !important;
}
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
div[class*="st-key-gc_memo_pop"] {
  position: fixed !important;
  top: 18% !important;
  left: 50% !important;
  transform: none !important;
  z-index: 10050 !important;
  width: min(340px, 92vw) !important;
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
    static_rules: list[str] = []
    sel_stamp = (
        selected.strftime("%Y%m%d") if isinstance(selected, date) else ""
    )
    today_stamp = today.strftime("%Y%m%d")
    planned: list[list[tuple]] = []
    for week in weeks:
        row = []
        for d in week:
            iso = d.isoformat()
            stamp = d.strftime("%Y%m%d")
            items = by_day.get(iso) or []
            is_hol = _gc_is_holiday_day(items)
            if d.weekday() >= 5 or is_hol:
                static_rules.append(
                    f"div[class*='st-key-gc_calgrid'] [data-testid='stColumn']"
                    f":has([class*='st-key-gc_cell_{stamp}']),"
                    f"div[class*='st-key-gc_cell_{stamp}']{{background:#fdecea!important;}}"
                )
            if d.month != month:
                color = "#bdc1c6"
            elif d.weekday() == 6 or is_hol:
                color = "#d93025"
            elif d.weekday() == 5:
                color = "#c62828"
            else:
                color = "#3c4043"
            extra = ""
            if d == today and d.month == month:
                color = "#1a73e8"
                extra = "font-weight:700!important;"
                static_rules.append(
                    f"div[class*='st-key-gc_calgrid'] [data-testid='stColumn']"
                    f":has([class*='st-key-gc_cell_{stamp}']),"
                    f"div[class*='st-key-gc_cell_{stamp}']{{background:#e8f4ff!important;}}"
                )
            static_rules.append(
                f"div[class*='st-key-gc_day_{stamp}'] button{{color:{color}!important;{extra}}}"
            )
            row.append((d, stamp, items, iso))
        planned.append(row)
    if static_rules:
        st.markdown(
            "<style>" + "".join(static_rules) + "</style>",
            unsafe_allow_html=True,
        )
    if sel_stamp and sel_stamp != today_stamp:
        st.markdown(
            "<style>"
            f"div[class*='st-key-gc_calgrid'] [data-testid='stColumn']"
            f":has([class*='st-key-gc_cell_{sel_stamp}']),"
            f"div[class*='st-key-gc_cell_{sel_stamp}']{{background:#e8f0fe!important;}}"
            f"div[class*='st-key-gc_day_{sel_stamp}'] button{{"
            "background:#d2e3fc!important;color:#174ea6!important;font-weight:700!important;}}"
            "</style>",
            unsafe_allow_html=True,
        )
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
                                    lab = calendar_chip_label(ev)
                                    with st.container(key=f"gc_chiprow_{stamp}_{idx}"):
                                        st.button(
                                            lab,
                                            key=f"gc_evpick_{stamp}_{idx}",
                                            type="tertiary",
                                            width="stretch",
                                            on_click=_gc_set_sel_event,
                                            args=(d, eid or iso, f"{stamp}_{idx}"),
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
    _gc_shrink_event_titles()


def _gc_shrink_event_titles() -> None:
    """일정 제목만 10px. 월 이동마다 스크립트를 넣지 않는다."""
    st.markdown(
        "<style>html body div[class*='st-key-gc_evpick_'] button,"
        "html body div[class*='st-key-gc_evpick_'] button p,"
        "html body div[class*='st-key-gc_evpick_'] button span,"
        "html body div[class*='st-key-gc_evpick_'] button div,"
        "html body div[class*='st-key-gc_evpick_'] [data-testid='stMarkdownContainer'] p"
        "{font-size:10px!important;line-height:1.2!important;font-weight:400!important}</style>",
        unsafe_allow_html=True,
    )


def _render_event_memo_pop(ev: dict) -> None:
    """일정 클릭 메모. 대화상자 대신 조각 안 팝업이라 대시보드를 다시 켜지 않는다."""
    title = str(ev.get("title") or "(제목 없음)")
    eid = str(ev.get("id") or "")
    cid = str(ev.get("cal_id") or "")
    note = str(ev.get("description") or "")
    if is_mail_calendar_event(ev):
        title = GC_MAIL_CHIP_TITLE
        sub = mail_subject_from_event(ev)
        if sub and f"제목: {sub}" not in note:
            note = f"제목: {sub}\n" + note.lstrip()
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
        move_msg = st.session_state.get("_gc_mail_move_msg")
        if is_mail_calendar_event(ev) and move_msg:
            st.caption(move_msg)
        need_mail_auth = bool(
            is_mail_calendar_event(ev)
            and (
                st.session_state.get("_gc_need_gmail")
                or not token_has_gmail_modify()
            )
        )
        if need_mail_auth:
            if not move_msg:
                st.caption("폴더 이동·삭제는 Gmail 권한을 다시 연결해야 합니다.")
            _gc_render_gmail_reconnect(key="gc_gmail_reauth_memo")
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
            if is_mail_calendar_event(ev):
                ids = gmail_ids_from_event({**ev, "description": st.session_state.get(nk) or note})
                st.caption("메일 이동")
                m1, m2, m3 = st.columns(3)
                with m1:
                    if st.button("개인업무", key="gc_mail_move_personal"):
                        st.session_state["_gc_pending_mail_move"] = {
                            "cal": cid,
                            "id": eid,
                            "msg": ids.get("msg") or "",
                            "thread": ids.get("thread") or "",
                            "label": GC_GMAIL_PERSONAL,
                        }
                        st.rerun(scope="fragment")
                with m2:
                    if st.button("신일가스", key="gc_mail_move_shinil"):
                        st.session_state["_gc_pending_mail_move"] = {
                            "cal": cid,
                            "id": eid,
                            "msg": ids.get("msg") or "",
                            "thread": ids.get("thread") or "",
                            "label": GC_GMAIL_LABEL,
                        }
                        st.rerun(scope="fragment")
                with m3:
                    if st.button("삭제", key="gc_mail_del"):
                        st.session_state["_gc_pending_mail_del"] = {
                            "cal": cid,
                            "id": eid,
                            "msg": ids.get("msg") or "",
                            "thread": ids.get("thread") or "",
                        }
                        st.session_state["_gc_memo_open"] = False
                        st.rerun(scope="fragment")
            elif eid:
                if st.button("삭제", key="gc_pop_del"):
                    st.session_state["_gc_pending_del_event"] = {"cal": cid, "id": eid}
                    st.session_state["_gc_memo_open"] = False
                    st.rerun(scope="fragment")
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


def _gc_place_memo_pop(chip: str) -> None:
    """메모 팝업을 클릭한 일정 제목 바로 아래로 붙인다. 날짜 버튼 키는 바꾸지 않는다."""
    key = "gc_evpick_" + str(chip or "").replace("\\", "").replace('"', "")
    st.html(
        f"""
<script>
(function () {{
  var key = {json.dumps(key)};
  function place() {{
    var pop = window.parent.document.querySelector('[class*="st-key-gc_memo_pop"]')
      || document.querySelector('[class*="st-key-gc_memo_pop"]');
    var doc = pop ? pop.ownerDocument : document;
    var win = doc.defaultView || window;
    var anchor = doc.querySelector('[class*="st-key-' + key + '"] button')
      || doc.querySelector('[class*="st-key-' + key + '"]');
    if (!pop || !anchor) return;
    var r = anchor.getBoundingClientRect();
    var pw = Math.min(340, win.innerWidth - 16);
    var ph = pop.getBoundingClientRect().height || 240;
    var left = Math.min(Math.max(8, r.left), win.innerWidth - pw - 8);
    var top = r.bottom + 6;
    if (top + ph > win.innerHeight - 8) {{
      top = Math.max(8, r.top - ph - 6);
    }}
    pop.style.setProperty("position", "fixed", "important");
    pop.style.setProperty("top", top + "px", "important");
    pop.style.setProperty("left", left + "px", "important");
    pop.style.setProperty("transform", "none", "important");
    pop.style.setProperty("width", pw + "px", "important");
  }}
  place();
  setTimeout(place, 50);
  setTimeout(place, 200);
}})();
</script>
""",
        unsafe_allow_javascript=True,
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
        st.session_state.pop("_gc_need_gmail", None)
        _qp_pop("code", "state", "scope", "prompt", "authuser", "hd", "session_state")
        if token_has_gmail_modify():
            st.success("구글 캘린더·할 일·Gmail이 연결되었습니다.")
        else:
            st.warning(
                "연결은 됐지만 Gmail 수정 권한이 없습니다. "
                "연결 해제 후 처음부터 다시 연결해 주세요."
            )
        st.rerun()
    except Exception as e:
        st.session_state.pop("_gc_oauth_code_used", None)
        st.error(f"구글 연결 실패: {e}")


def _setup_expander() -> None:
    cid, csec = load_oauth_client()
    with st.expander("구글 연결 설정 (최초 1회)", expanded=not bool(cid and csec)):
        st.markdown(
            "Google Cloud에서 **캘린더 API · Tasks API · Gmail API**를 켜고, OAuth 클라이언트(웹)를 만든 뒤 "
            "아래 리다이렉트 URI를 등록하세요. 메일 연동·신일가스 폴더 이동은 구글 계정을 다시 연결해야 합니다."
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
    pending_move = st.session_state.pop("_gc_pending_mail_move", None)
    if isinstance(pending_move, dict) and (pending_move.get("thread") or pending_move.get("msg")):
        try:
            folder = str(pending_move.get("label") or GC_GMAIL_LABEL).strip() or GC_GMAIL_LABEL
            gmail_move_inbox_to_label(
                thread_id=str(pending_move.get("thread") or ""),
                msg_id=str(pending_move.get("msg") or ""),
                label_name=folder,
            )
            st.session_state["_gc_mail_move_msg"] = f"받은편지함에서 {folder} 폴더로 옮겼습니다."
        except GmailScopeError as e:
            st.session_state["_gc_need_gmail"] = True
            st.session_state["_gc_mail_move_msg"] = str(e)
        except Exception as e:
            st.session_state["_gc_mail_move_msg"] = f"메일 이동 실패: {e}"
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
    pending_mail_del = st.session_state.pop("_gc_pending_mail_del", None)
    if isinstance(pending_mail_del, dict) and pending_mail_del.get("id"):
        cid = str(pending_mail_del.get("cal") or "primary")
        eid = str(pending_mail_del.get("id") or "")
        try:
            gmail_trash_mail(
                thread_id=str(pending_mail_del.get("thread") or ""),
                msg_id=str(pending_mail_del.get("msg") or ""),
            )
            _drop_gmail_sync_id(str(pending_mail_del.get("msg") or ""))
            st.session_state["_gc_mail_move_msg"] = "메일과 캘린더에서 삭제했습니다."
        except GmailScopeError as e:
            st.session_state["_gc_need_gmail"] = True
            st.session_state["_gc_mail_move_msg"] = str(e)
        except Exception as e:
            st.session_state["_gc_mail_move_msg"] = f"메일 삭제 실패: {e}"
        _gc_cache_remove_event(cid, eid)
        if eid and not eid.startswith("local-"):
            try:
                delete_event(cid, eid)
            except Exception as e:
                st.error(f"일정 삭제 실패: {e}")
        st.session_state["_gc_memo_open"] = False
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
        st.session_state["_gc_memo_open"] = False

    try:
        cals = _gc_cached_cals()
    except Exception as e:
        st.error(f"캘린더를 불러오지 못했습니다: {e}")
        cals = []

    st.markdown(
        """
<style>
div[class*="st-key-gc_cal_top"] {
  position: relative !important;
}
div[class*="st-key-gc_cal_top"] > div[data-testid="stVerticalBlock"] {
  gap: 0 !important;
  position: relative !important;
}
div[class*="st-key-gc_ym"] {
  height: 0 !important;
  min-height: 0 !important;
  margin: 0 !important;
  padding: 0 !important;
  overflow: visible !important;
  pointer-events: none !important;
}
div[class*="st-key-gc_ym"] p,
div[class*="st-key-gc_ym"] .gc-ym {
  position: absolute !important;
  left: 50% !important;
  top: 50% !important;
  transform: translate(-50%, -50%) !important;
  font-size: 18px !important;
  font-weight: 700 !important;
  line-height: 1.2 !important;
  margin: 0 !important;
  text-align: center !important;
  white-space: nowrap !important;
  z-index: 2 !important;
  pointer-events: none !important;
}
div[class*="st-key-gc_head_acts"] [data-testid="stHorizontalBlock"] {
  gap: 4px !important;
  flex-wrap: nowrap !important;
  justify-content: flex-end !important;
}
div[class*="st-key-gc_head_acts"] button {
  min-height: 2rem !important;
  padding-left: 0.65rem !important;
  padding-right: 0.65rem !important;
  white-space: nowrap !important;
}
</style>
""",
        unsafe_allow_html=True,
    )
    with st.container(key="gc_cal_top"):
        with st.container(key="gc_ym"):
            st.markdown(
                f'<div class="gc-ym">{html.escape(f"{y}년 {m}월")}</div>',
                unsafe_allow_html=True,
            )
        top = st.columns([1.05, 1.05, 3.15, 2.55], vertical_alignment="center")
        with top[0]:
            st.button(
                "오늘",
                key="gc_today",
                on_click=_gc_set_view_month,
                args=(date(today.year, today.month, 1),),
            )
        with top[1]:
            n1, n2 = st.columns(2)
            with n1:
                st.button(
                    "‹",
                    key="gc_prev",
                    on_click=_gc_set_view_month,
                    args=(date(y, m, 1) - timedelta(days=1),),
                )
            with n2:
                st.button(
                    "›",
                    key="gc_next",
                    on_click=_gc_set_view_month,
                    args=(date(y, m, 28) + timedelta(days=8),),
                )
        with top[2]:
            pass
        with top[3]:
            with st.container(key="gc_head_acts"):
                a1, a2 = st.columns(2, gap="small")
                with a1:
                    st.button(
                        "일정 추가",
                        type="primary",
                        key="gc_add_open",
                        on_click=_gc_open_add_pop,
                    )
                with a2:
                    if st.button("연동실행", type="primary", key="gc_sync_run"):
                        try:
                            n_ev, n_tk = _gc_sync_google(y, m)
                            mail = sync_shinil_gmail_to_calendar()
                            extra = ""
                            err = str(mail.get("error") or "")
                            if err:
                                extra = f" · 신일가스 메일: {err}"
                            else:
                                extra = (
                                    f" · 신일가스 메일 {int(mail.get('added') or 0)}건 추가"
                                )
                                remain = int(mail.get("remain") or 0)
                                if remain:
                                    extra += f" · 남은 {remain}건은 다시 연동"
                            st.session_state["_gc_sync_msg"] = (
                                f"연동했습니다. 일정 {n_ev}건 · 할 일 {n_tk}건{extra}"
                            )
                            st.session_state.pop("_gc_need_gmail", None)
                        except GmailScopeError as e:
                            st.session_state["_gc_need_gmail"] = True
                            st.session_state["_gc_sync_msg"] = str(e)
                        except Exception as e:
                            st.session_state["_gc_sync_msg"] = f"연동 실패: {e}"
                        st.rerun()
    msg = st.session_state.pop("_gc_sync_msg", None)
    if msg:
        st.caption(msg)
    move_msg = st.session_state.get("_gc_mail_move_msg")
    if move_msg:
        st.caption(move_msg)
    if st.session_state.get("_gc_need_gmail") or (
        not token_has_gmail_modify() and bool(load_token())
    ):
        _gc_render_gmail_reconnect(key="gc_gmail_reauth_head")

    y = int(st.session_state.get("gc_year") or today.year)
    m = int(st.session_state.get("gc_month") or today.month)
    by_day = _gc_cached_events(cals, y, m)
    if st.session_state.get("_gc_pf_ym") != (y, m):
        st.session_state["_gc_pf_ym"] = (y, m)
        _gc_schedule_prefetch(cals, y, m)
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
            _gc_place_memo_pop(str(st.session_state.get("_gc_sel_chip") or ""))
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
                t = dict(t)
                t["status"] = "completed"
                if not t.get("local_only") and nid and not nid.startswith("local-"):
                    try:
                        complete_task(lid, nid)
                    except Exception as e:
                        st.warning(f"구글 완료 반영 실패: {e}")
            left.append(t)
        save_local_tasks(lid, left)
    pending_restore = st.session_state.pop("_gc_pending_restore", None)
    if pending_restore and lid:
        nid = str(pending_restore)
        left = []
        for t in load_local_tasks(lid):
            if str(t.get("id")) == nid:
                t = dict(t)
                t["status"] = "needsAction"
                if not t.get("local_only") and nid and not nid.startswith("local-"):
                    try:
                        uncomplete_task(lid, nid)
                    except Exception as e:
                        st.warning(f"구글 복원 반영 실패: {e}")
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
div[class*="st-key-gc_ns_"] button,
div[class*="st-key-gc_tr_"] button {
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
    open_tasks = [t for t in tasks if str(t.get("status") or "") != "completed"]
    done_tasks = [t for t in tasks if str(t.get("status") or "") == "completed"]
    for t in open_tasks:
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
    with st.expander(f"완료된 할 일 ({len(done_tasks)})", expanded=False):
        if not done_tasks:
            st.caption("완료된 할 일이 없습니다.")
        for t in done_tasks:
            tid = str(t.get("id") or "")
            title = t.get("title") or "(제목 없음)"
            due = str((t.get("due") or ""))[:10]
            label = title if not due else f"{title}  · {due}"
            note = str(t.get("notes") or "").strip()
            row = st.columns([8.2, 1.15], gap="small", vertical_alignment="top")
            with row[0]:
                st.markdown(
                    "<div style='font-size:13px;color:#80868b;text-decoration:line-through;"
                    f"line-height:1.35'>{html.escape(label)}</div>",
                    unsafe_allow_html=True,
                )
                if note:
                    st.markdown(
                        "<div style='font-size:12px;color:#9aa0a6;line-height:1.35;"
                        "margin:0 0 6px 0;white-space:pre-wrap'>"
                        f"{html.escape(note)}</div>",
                        unsafe_allow_html=True,
                    )
            with row[1]:
                if st.button(
                    "복원",
                    key=f"gc_tr_{tid}",
                    type="tertiary",
                    width="content",
                ):
                    st.session_state["_gc_pending_restore"] = tid
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


@st.fragment
def _gc_mini_panel() -> None:
    """작은 달력만 월을 바꾸고, 화면 전체를 다시 켜지 않는다."""
    today = datetime.now(_TZ).date()
    if not st.session_state.get("gc_mini_year"):
        st.session_state["gc_mini_year"] = today.year
        st.session_state["gc_mini_month"] = today.month
    y = int(st.session_state.get("gc_mini_year") or today.year)
    m = int(st.session_state.get("gc_mini_month") or today.month)
    sel = st.session_state.get("_gc_sel_day")
    _render_mini_month(
        y,
        m,
        today,
        sel if isinstance(sel, date) else None,
        holidays=_gc_cached_holiday_keys(y, m),
    )


def _gc_work_panel() -> None:
    """상단 가로 달력 아래: 왼쪽 작은 달력 · 가운데 월간 · 오른쪽 할 일."""
    mini, cal, tasks = st.columns([0.82, 2.70, 1.05], gap="small")
    with mini:
        _gc_mini_panel()
    with cal:
        _gc_cal_panel()
    with tasks:
        _gc_task_panel()


def render_google_calendar_tab(
    df=None, latest_update_str: str = ""
) -> None:
    """상단 담당자·거래처·납품줄 + 구글 월간 달력·할 일."""
    st.html(
        """
<script>
(function () {
  try {
    var w = window.parent || window;
    if (w.__gcEvpickObs) { w.__gcEvpickObs.disconnect(); w.__gcEvpickObs = null; }
  } catch (e) {}
  var id = "gc-evpick-fs";
  var css = 'html body div[class*="st-key-gc_evpick_"] button,' +
    'html body div[class*="st-key-gc_evpick_"] button p,' +
    'html body div[class*="st-key-gc_evpick_"] button span,' +
    'html body div[class*="st-key-gc_evpick_"] button div,' +
    'html body div[class*="st-key-gc_evpick_"] [data-testid="stMarkdownContainer"] p' +
    '{font-size:10px!important;line-height:1.2!important;font-weight:400!important;}';
  function put(doc) {
    if (!doc) return;
    var s = doc.getElementById(id);
    if (!s) {
      s = doc.createElement("style");
      s.id = id;
      (doc.head || doc.documentElement).appendChild(s);
    }
    s.textContent = css;
  }
  put(document);
  try { put(window.parent.document); } catch (e) {}
  function bindDayPick(doc) {
    if (!doc || doc.__gcDayPick) return;
    doc.__gcDayPick = true;
    doc.addEventListener("click", function (ev) {
      var t = ev.target;
      if (!t || !t.closest) return;
      var host = t.closest('[class*="st-key-gc_day_"]');
      if (!host) return;
      var m = String(host.className || "").match(/st-key-gc_day_([0-9]{8})/);
      if (!m) return;
      var stamp = m[1];
      doc.querySelectorAll(".gc-js-sel").forEach(function (n) {
        n.classList.remove("gc-js-sel");
      });
      var cell = doc.querySelector('[class*="st-key-gc_cell_' + stamp + '"]');
      if (cell) {
        cell.classList.add("gc-js-sel");
        var col = cell.closest('[data-testid="stColumn"]');
        if (col) col.classList.add("gc-js-sel");
      }
    }, true);
  }
  bindDayPick(document);
  try { bindDayPick(window.parent.document); } catch (e2) {}
})();
</script>
""",
        unsafe_allow_javascript=True,
    )
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
        if st.session_state.get("_gc_need_reauth"):
            st.warning("구글 인증이 만료되었습니다. 다시 연결해 주세요.")
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
