"""구글 캘린더·할 일 — API로 월간 달력과 할 일 목록을 한 화면에 붙여 작업한다."""
from __future__ import annotations

import calendar
import html
import json
import os
import secrets
import urllib.parse
from datetime import date, datetime, time, timedelta, timezone
from typing import Any
from zoneinfo import ZoneInfo

import requests
import streamlit as st

GC_DIR = os.path.join("uploaded_cache", "google_calendar")
GC_TOKEN = os.path.join(GC_DIR, "token.json")
GC_OAUTH = os.path.join(GC_DIR, "oauth_client.json")
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


def event_day_key(item: dict) -> str:
    start = (item or {}).get("start") or {}
    raw = str(start.get("date") or start.get("dateTime") or "")
    return raw[:10]


def _mkdir() -> None:
    os.makedirs(GC_DIR, exist_ok=True)


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
    if os.path.exists(GC_OAUTH):
        try:
            with open(GC_OAUTH, "r", encoding="utf-8") as f:
                data = json.load(f)
            installed = data.get("web") or data.get("installed") or data
            cid = cid or str(installed.get("client_id") or "").strip()
            csec = csec or str(installed.get("client_secret") or "").strip()
        except Exception:
            pass
    return cid, csec


def save_oauth_client(client_id: str, client_secret: str) -> None:
    _mkdir()
    with open(GC_OAUTH, "w", encoding="utf-8") as f:
        json.dump(
            {
                "web": {
                    "client_id": client_id.strip(),
                    "client_secret": client_secret.strip(),
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": _TOKEN,
                }
            },
            f,
            ensure_ascii=False,
            indent=2,
        )


def load_token() -> dict:
    if os.path.exists(GC_TOKEN):
        try:
            with open(GC_TOKEN, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and data.get("access_token"):
                return data
        except Exception:
            return {}
    tok = st.session_state.get("_gc_token")
    return tok if isinstance(tok, dict) else {}


def save_token(tok: dict) -> None:
    _mkdir()
    try:
        with open(GC_TOKEN, "w", encoding="utf-8") as f:
            json.dump(tok, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    st.session_state["_gc_token"] = tok


def clear_token() -> None:
    st.session_state.pop("_gc_token", None)
    try:
        if os.path.exists(GC_TOKEN):
            os.remove(GC_TOKEN)
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


def auth_url(client_id: str, state: str) -> str:
    q = {
        "client_id": client_id,
        "redirect_uri": redirect_uri(),
        "response_type": "code",
        "scope": " ".join(_SCOPES),
        "access_type": "offline",
        "include_granted_scopes": "true",
        "prompt": "consent",
        "state": state,
        "hl": "ko",
    }
    return _AUTH + "?" + urllib.parse.urlencode(q)


def exchange_code(client_id: str, client_secret: str, code: str) -> dict:
    r = requests.post(
        _TOKEN,
        data={
            "code": code,
            "client_id": client_id,
            "client_secret": client_secret,
            "redirect_uri": redirect_uri(),
            "grant_type": "authorization_code",
        },
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    data["_obtained"] = datetime.now(timezone.utc).isoformat()
    return data


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
    if _expired(tok) and cid and csec:
        try:
            tok = refresh_token(cid, csec, tok)
            save_token(tok)
        except Exception:
            return str(tok.get("access_token") or "")
    return str(tok.get("access_token") or "")


def api_get(url: str, params: dict | None = None) -> dict:
    tok = access_token()
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
        out.append(
            {
                "id": it.get("id") or "",
                "summary": it.get("summary") or it.get("id") or "",
                "color": it.get("backgroundColor") or "#7b8794",
                "primary": bool(it.get("primary")),
            }
        )
    return out


def list_events(cal_id: str, start: date, end: date) -> list[dict]:
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
        },
    )
    return list(data.get("items") or [])


def insert_event(cal_id: str, title: str, day: date, all_day: bool) -> dict:
    body: dict[str, Any] = {"summary": title}
    if all_day:
        body["start"] = {"date": day.isoformat()}
        body["end"] = {"date": (day + timedelta(days=1)).isoformat()}
    else:
        s = datetime.combine(day, time(9, 0), tzinfo=_TZ)
        e = s + timedelta(hours=1)
        body["start"] = {"dateTime": s.isoformat(), "timeZone": "Asia/Seoul"}
        body["end"] = {"dateTime": e.isoformat(), "timeZone": "Asia/Seoul"}
    return api_json("POST", f"{_CAL_API}/calendars/{urllib.parse.quote(cal_id, safe='')}/events", body)


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


def add_task(list_id: str, title: str) -> dict:
    return api_json(
        "POST",
        f"{_TASK_API}/lists/{urllib.parse.quote(list_id, safe='')}/tasks",
        {"title": title},
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


def _gc_cached_events(cals: list[dict], year: int, month: int) -> dict[str, list[dict]]:
    box = _gc_api_box()
    key = f"ev:{year}-{month:02d}"
    if box.get(key) is None:
        box[key] = _events_by_day(cals, year, month) if cals else {}
    return box[key]


def _gc_cached_tasks(list_id: str) -> list[dict]:
    box = _gc_api_box()
    key = f"tk:{list_id}"
    if list_id and key not in box:
        box[key] = list_tasks(list_id)
    return list(box.get(key) or [])


def _gc_set_tasks(list_id: str, tasks: list[dict]) -> None:
    _gc_api_box()[f"tk:{list_id}"] = list(tasks)


def _events_by_day(cals: list[dict], year: int, month: int) -> dict[str, list[dict]]:
    weeks = month_weeks(year, month)
    start, end = weeks[0][0], weeks[-1][-1]
    by: dict[str, list[dict]] = {}
    for cal in cals:
        try:
            items = list_events(cal["id"], start, end)
        except Exception:
            continue
        color = cal.get("color") or "#7b8794"
        name = cal.get("summary") or ""
        for it in items:
            key = event_day_key(it)
            if not key:
                continue
            stt = (it.get("start") or {}).get("dateTime") or ""
            hhmm = ""
            if "T" in stt:
                try:
                    dt = datetime.fromisoformat(stt.replace("Z", "+00:00")).astimezone(_TZ)
                    hhmm = dt.strftime("%H:%M")
                except Exception:
                    hhmm = ""
            by.setdefault(key, []).append(
                {
                    "title": it.get("summary") or "(제목 없음)",
                    "color": color,
                    "cal": name,
                    "hhmm": hhmm,
                }
            )
    return by


def _month_html(year: int, month: int, by_day: dict[str, list[dict]], today: date) -> str:
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
            if d.weekday() == 6:
                cls.append("gc-sun")
            if d.weekday() == 5:
                cls.append("gc-sat")
            chips = []
            for ev in (by_day.get(key) or [])[:4]:
                label = html.escape((ev.get("hhmm") + " " if ev.get("hhmm") else "") + ev["title"])
                col = html.escape(str(ev.get("color") or "#c5c9ce"))
                chips.append(
                    f'<div class="gc-chip" style="background:{col}22;border-left:3px solid {col}">{label}</div>'
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
.gc-table {{ width:100%; border-collapse:collapse; table-layout:fixed; background:#fff; }}
.gc-table th {{ font-weight:500; font-size:12px; padding:8px 6px; text-align:center; color:#70757a; }}
.gc-wn {{ width:36px; color:#80868b; font-size:11px; text-align:center; background:#fafafa; }}
.gc-cell {{ height:92px; vertical-align:top; border:1px solid #e8eaed; padding:4px; }}
.gc-out {{ color:#bdc1c6; background:#fcfcfd; }}
.gc-today .gc-num {{ background:#1a73e8; color:#fff; border-radius:50%; width:22px; height:22px;
  display:inline-flex; align-items:center; justify-content:center; }}
.gc-num {{ font-size:12px; font-weight:500; margin-bottom:2px; }}
.gc-sun, .gc-sun .gc-num {{ color:#d93025; }}
.gc-sat, .gc-sat .gc-num {{ color:#1a73e8; }}
.gc-chip {{ font-size:11px; line-height:1.25; padding:1px 4px; margin:1px 0; border-radius:2px;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.gc-more {{ font-size:11px; color:#5f6368; }}
</style>
<div class="gc-wrap"><table class="gc-table"><thead><tr><th class="gc-wn"></th>{heads}</tr></thead>
<tbody>{"".join(rows)}</tbody></table></div>
"""


def _consume_oauth_redirect() -> None:
    try:
        qp = st.query_params
        err = str(qp.get("error") or "")
        code = str(qp.get("code") or "")
        state = str(qp.get("state") or "")
    except Exception:
        return
    if err:
        st.error("구글 연결이 취소되었습니다.")
        try:
            st.query_params.clear()
        except Exception:
            pass
        return
    if not code:
        return
    expect = str(st.session_state.get("_gc_oauth_state") or "")
    if expect and state and state != expect:
        st.error("구글 연결 상태가 맞지 않습니다. 다시 연결해 주세요.")
        return
    cid, csec = load_oauth_client()
    if not cid or not csec:
        st.error("OAuth 클라이언트 ID/비밀번호가 없습니다.")
        return
    try:
        tok = exchange_code(cid, csec, code)
        save_token(tok)
        st.session_state.pop("_gc_oauth_state", None)
        try:
            st.query_params.clear()
        except Exception:
            pass
        st.success("구글 캘린더·할 일이 연결되었습니다.")
        st.rerun()
    except Exception as e:
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
def _gc_work_panel() -> None:
    """할 일 추가·완료는 이 조각만 다시 그린다. 엑셀·납품줄·다른 탭은 다시 안 읽는다."""
    today = datetime.now(_TZ).date()
    vc_mon = st.session_state.get("_vc_month")
    if isinstance(vc_mon, date):
        st.session_state["gc_year"] = vc_mon.year
        st.session_state["gc_month"] = vc_mon.month
    y = int(st.session_state.get("gc_year") or today.year)
    m = int(st.session_state.get("gc_month") or today.month)

    lid = str(st.session_state.get("_gc_lid") or "")
    box = _gc_api_box()
    pending_new = st.session_state.pop("_gc_pending_new", None)
    if pending_new and lid:
        try:
            created = add_task(lid, str(pending_new))
            cur = list(box.get(f"tk:{lid}") or [])
            cur.append(created)
            _gc_set_tasks(lid, cur)
        except Exception as e:
            st.error(f"추가 실패: {e}")
    pending_done = st.session_state.pop("_gc_pending_done", None)
    if pending_done and lid:
        try:
            complete_task(lid, str(pending_done))
            _gc_set_tasks(
                lid,
                [
                    t
                    for t in (box.get(f"tk:{lid}") or [])
                    if str(t.get("id")) != str(pending_done)
                ],
            )
        except Exception as e:
            st.error(f"완료 처리 실패: {e}")
    pending_ev = st.session_state.pop("_gc_pending_event", None)
    if isinstance(pending_ev, dict):
        try:
            day = pending_ev["day"]
            insert_event(
                str(pending_ev.get("cal") or "primary"),
                str(pending_ev.get("title") or ""),
                day,
                True,
            )
            box.pop(f"ev:{day.year}-{day.month:02d}", None)
        except Exception as e:
            st.error(f"일정 저장 실패: {e}")

    try:
        cals = _gc_cached_cals()
    except Exception as e:
        st.error(f"캘린더를 불러오지 못했습니다: {e}")
        cals = []
    try:
        tlists = _gc_cached_tlists()
    except Exception as e:
        st.error(f"할 일 목록을 불러오지 못했습니다: {e}")
        tlists = []

    top = st.columns([1.1, 1.1, 4, 1.6])
    with top[0]:
        if st.button("오늘", key="gc_today"):
            st.session_state["gc_year"] = today.year
            st.session_state["gc_month"] = today.month
            st.rerun(scope="fragment")
    with top[1]:
        n1, n2 = st.columns(2)
        with n1:
            if st.button("‹", key="gc_prev"):
                d = date(y, m, 1) - timedelta(days=1)
                st.session_state["gc_year"] = d.year
                st.session_state["gc_month"] = d.month
                st.rerun(scope="fragment")
        with n2:
            if st.button("›", key="gc_next"):
                d = date(y, m, 28) + timedelta(days=8)
                st.session_state["gc_year"] = d.year
                st.session_state["gc_month"] = d.month
                st.rerun(scope="fragment")
    with top[2]:
        st.markdown(f"**{y}년 {m}월**")
    with top[3]:
        if st.button("연결 해제", key="gc_logout"):
            clear_token()
            st.session_state.pop("_gc_api", None)
            st.rerun()

    y = int(st.session_state.get("gc_year") or today.year)
    m = int(st.session_state.get("gc_month") or today.month)
    left, right = st.columns([3.15, 1.12], gap="small")
    with left:
        by_day = _gc_cached_events(cals, y, m)
        st.markdown(_month_html(y, m, by_day, today), unsafe_allow_html=True)
        with st.expander("일정 추가", expanded=False):
            et = st.text_input("제목", key="gc_ev_title")
            ed = st.date_input("날짜", value=today, key="gc_ev_day")
            cal_opts = {c["summary"]: c["id"] for c in cals} or {"기본": "primary"}
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
                    }
                    st.rerun(scope="fragment")
    with right:
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
        for t in tasks:
            tid = str(t.get("id") or "")
            title = t.get("title") or "(제목 없음)"
            due = str((t.get("due") or ""))[:10]
            label = title if not due else f"{title}  · {due}"
            if st.checkbox(label, value=False, key=f"gc_td_{tid}"):
                st.session_state["_gc_pending_done"] = tid
                st.rerun(scope="fragment")
        nt = st.text_input("새 할 일", key="gc_new_task", placeholder="할 일 추가")
        if st.button("추가", key="gc_add_task"):
            title = (nt or "").strip()
            if len(title) < 1:
                st.error("내용을 입력하세요.")
            else:
                st.session_state["_gc_pending_new"] = title
                st.session_state["gc_new_task"] = ""
                st.rerun(scope="fragment")


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
