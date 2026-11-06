"""메일 탭 하단 — 신일가스 Gmail 폴더 제목·본문 보기."""
from __future__ import annotations

import base64
import html
import re
import urllib.parse
from datetime import date

import streamlit as st

import google_calendar_tab as gc

_GF_ROWS = "_gf_shinil_rows"
_GF_BODY = "_gf_shinil_body"
_GF_ERR = "_gf_shinil_err"
_MAX_LIST = 40


def _b64url(data: str) -> str:
    raw = str(data or "").strip()
    if not raw:
        return ""
    pad = "=" * ((4 - len(raw) % 4) % 4)
    try:
        return base64.urlsafe_b64decode(raw + pad).decode("utf-8", "replace")
    except Exception:
        return ""


def _strip_html(raw: str) -> str:
    text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", str(raw or ""))
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    return re.sub(r"[ \t]+\n", "\n", re.sub(r"\n{3,}", "\n\n", text)).strip()


def gmail_plain_body(payload: dict | None) -> str:
    """Gmail payload에서 읽기용 텍스트만 고른다. HTML은 태그를 뗀다."""
    node = payload if isinstance(payload, dict) else {}
    mime = str(node.get("mimeType") or "").lower()
    data = str((node.get("body") or {}).get("data") or "")
    parts = list(node.get("parts") or [])
    if data and mime.startswith("text/plain"):
        return _b64url(data).strip()
    plain = ""
    rich = ""
    for part in parts:
        got = gmail_plain_body(part)
        child_mime = str((part or {}).get("mimeType") or "").lower()
        if got and child_mime.startswith("text/plain") and not plain:
            plain = got
        elif got and not rich:
            rich = got
    if plain:
        return plain
    if rich:
        return rich
    if data:
        decoded = _b64url(data)
        if "html" in mime:
            return _strip_html(decoded)
        return decoded.strip()
    return ""


def _row_label(row: dict) -> str:
    day = row.get("day")
    day_s = day.isoformat() if isinstance(day, date) else ""
    sub = str(row.get("subject") or "(제목 없음)").strip()
    if len(sub) > 72:
        sub = sub[:70] + "…"
    return f"{day_s}  {sub}".strip()


def _load_folder_rows() -> list[dict]:
    labels = gc.list_gmail_labels()
    lab = gc.match_gmail_label(labels, gc.GC_GMAIL_LABEL)
    if not lab:
        raise RuntimeError("Gmail에 신일가스 폴더가 없습니다.")
    listed = gc.list_gmail_message_ids(str(lab.get("id") or ""), limit=_MAX_LIST)
    rows: list[dict] = []
    for it in listed:
        mid = str((it or {}).get("id") or "")
        if not mid:
            continue
        try:
            meta = gc.gmail_message_meta(mid)
        except Exception:
            continue
        if not meta.get("id"):
            continue
        rows.append(meta)
    return rows


def _load_body(msg_id: str) -> str:
    mid = str(msg_id or "").strip()
    if not mid:
        return ""
    data = gc._gmail_get(
        f"messages/{urllib.parse.quote(mid, safe='')}",
        {"format": "full", "fields": "id,snippet,payload"},
    )
    text = gmail_plain_body(data.get("payload") if isinstance(data, dict) else {})
    if text:
        return text
    return str((data or {}).get("snippet") or "").strip()


@st.fragment
def render_shinil_gmail_folder() -> None:
    """메일 탭 하단. 신일가스 라벨 제목·본문만 본다."""
    st.divider()
    st.markdown("##### 신일가스 메일 폴더")
    if not gc.access_token():
        st.caption("구글캘린더 탭에서 구글 계정을 연결하면 이 폴더가 보입니다.")
        return
    top = st.columns([1.2, 4])
    with top[0]:
        reload = st.button("폴더 새로고침", key="gf_shinil_reload")
    if reload or not isinstance(st.session_state.get(_GF_ROWS), list):
        try:
            st.session_state[_GF_ROWS] = _load_folder_rows()
            st.session_state.pop(_GF_BODY, None)
            st.session_state.pop(_GF_ERR, None)
        except gc.GmailScopeError as e:
            st.session_state[_GF_ERR] = str(e)
            st.session_state[_GF_ROWS] = []
        except Exception as e:
            st.session_state[_GF_ERR] = str(e)
            st.session_state[_GF_ROWS] = []
    err = str(st.session_state.get(_GF_ERR) or "")
    if err:
        st.caption(err)
        return
    rows = list(st.session_state.get(_GF_ROWS) or [])
    if not rows:
        st.caption("신일가스 폴더에 메일이 없습니다.")
        return
    ids = [str(r.get("id") or "") for r in rows if r.get("id")]
    labels = {str(r.get("id")): _row_label(r) for r in rows}
    picked = st.selectbox(
        "제목",
        options=ids,
        format_func=lambda i: labels.get(i, i),
        key="gf_shinil_sel",
    )
    row = next((r for r in rows if str(r.get("id")) == str(picked or "")), None)
    if not row:
        return
    sender = str(row.get("sender") or "-")
    st.caption(f"보낸사람: {sender}")
    url = gc.gmail_thread_url(str(row.get("threadId") or row.get("id") or ""))
    st.markdown(f"[지메일에서 열기]({url})")
    body_box = st.session_state.get(_GF_BODY)
    if not isinstance(body_box, dict) or str(body_box.get("id")) != str(picked):
        try:
            text = _load_body(str(picked))
        except gc.GmailScopeError as e:
            st.caption(str(e))
            return
        except Exception as e:
            st.caption(f"본문을 읽지 못했습니다: {e}")
            return
        body_box = {"id": str(picked), "text": text}
        st.session_state[_GF_BODY] = body_box
    body_txt = str(body_box.get("text") or "")
    st.markdown(
        "<div style='font-size:13px;line-height:1.45;white-space:pre-wrap;"
        "border:1px solid #e2e8f0;border-radius:8px;padding:10px 12px;"
        "max-height:280px;overflow:auto;background:#f8fafc'>"
        f"{html.escape(body_txt) if body_txt else '본문이 없습니다.'}</div>",
        unsafe_allow_html=True,
    )
