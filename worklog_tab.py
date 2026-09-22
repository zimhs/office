"""일일업무일지 탭 — 엑셀 양식 그대로 표시/편집/날짜별 저장/출력."""
from __future__ import annotations

import calendar
import html
import io
import json
import os
import platform
import re
import shutil
import subprocess
import threading
import time
from datetime import date
from functools import lru_cache
from typing import Any

import streamlit as st
import streamlit.components.v1 as components
from streamlit.errors import StreamlitAPIException
try:
    from streamlit.errors import StreamlitDuplicateElementKey as _WLDupKeyError
except Exception:  # 구버전 호환
    _WLDupKeyError = Exception

from dev_mode import dev_caption, is_dev_mode

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    load_workbook = None
    get_column_letter = None


def _wl_rerun(*, full: bool = False) -> None:
    """업무일지 fragment 안이면 fragment만 다시 실행 (전체 앱 로딩 방지).

    full=True 는 저장·날짜변경처럼 왼쪽 요약도 같이 갱신해야 할 때 사용.
    위젯 on_change 콜백 안에서는 호출하지 말 것 — Streamlit이 이미 rerun 중이라
    중복 rerun이 Cached ForwardMsg MISS·특수기호 실패를 유발한다.
    """
    if full:
        st.rerun()
        return
    now = time.time()
    last = float(st.session_state.get("_wl_frag_rerun_ts") or 0)
    if now - last < 0.12:
        return
    st.session_state["_wl_frag_rerun_ts"] = now
    try:
        st.rerun(scope="fragment")
    except (StreamlitAPIException, RuntimeError):
        st.rerun()


def _pin_worklog_scroll() -> None:
    """날짜 변경·저장 후 Streamlit이 본문을 위로 끌어올리지 않게 한다."""
    st.session_state["wl_scroll_pin"] = int(st.session_state.get("wl_scroll_pin") or 0) + 1
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and (k.startswith("wl_focus_ln_") or k.startswith("wl_focus_caret_")):
            st.session_state.pop(k, None)
    st.session_state.pop("wl_active_cell_key", None)


def _render_worklog_scroll_lock(iso: str) -> None:
    _WL_SCROLL_LOCK(
        key="wl_scroll_lock",
        data={"iso": iso, "pin": int(st.session_state.get("wl_scroll_pin") or 0)},
        height=1,
    )


def _wl_quiet_ui() -> bool:
    try:
        if st.session_state.get("force_touch_ui") is True: return True
        v = st.query_params.get("touch_ui", "")
        if isinstance(v, (list, tuple)): v = v[0] if v else ""
        if str(v).strip() in ("1", "true", "True"): return True
    except Exception: pass
    try: return platform.system() != "Darwin"
    except Exception: return True


def _wl_is_streamlit_cloud() -> bool:
    try:
        env = (os.environ.get("STREAMLIT_RUNTIME_ENVIRONMENT") or "").strip().lower()
        if env == "cloud":
            return True
    except Exception:
        pass
    try:
        if os.path.abspath(os.getcwd()).startswith("/mount/src"):
            return True
    except Exception:
        pass
    return False


def _wl_is_ipad_ui() -> bool:
    """iPad/touch UI — app.is_touch_ui()와 동일 기준. 맥·데스크톱 브라우저는 False."""
    try:
        if st.session_state.get("force_touch_ui") is True:
            return True
        v = st.query_params.get("touch_ui", "")
        if isinstance(v, (list, tuple)):
            v = v[0] if v else ""
        if str(v).strip() in ("1", "true", "True"):
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
    except Exception:
        pass
    return bool(st.session_state.get("force_touch_ui"))


def _invalidate_saved_dates_cache() -> None:
    st.session_state.pop("wl_saved_dates_cache", None)
    st.session_state.pop("wl_stored_cells_memo", None)


def _remember_calendar_saved_date(d: date) -> None:
    """저장한 날짜를 달력 • 캐시에 바로 넣는다. 빈 캐시 때문에 점이 빠지지 않게 한다."""
    iso = d.isoformat()
    cached = st.session_state.get("wl_saved_dates_cache")
    if not isinstance(cached, set):
        cached = set(list_saved_worklog_dates())
    cached.add(iso)
    st.session_state["wl_saved_dates_cache"] = cached


def _drop_saved_date_from_cache(iso: str) -> None:
    """달력 • 캐시에서 그 날짜만 뺀다. 월별 xlsx 전체 재스캔을 피한다."""
    cached = st.session_state.get("wl_saved_dates_cache")
    if isinstance(cached, set):
        cached.discard(str(iso))
    else:
        st.session_state.pop("wl_saved_dates_cache", None)
    _invalidate_stored_cells_memo(str(iso))


def _invalidate_stored_cells_memo(iso: str | None = None) -> None:
    if iso is None:
        st.session_state.pop("wl_stored_cells_memo", None)
        return
    memo = st.session_state.get("wl_stored_cells_memo")
    if not isinstance(memo, dict):
        return
    pfx = f"{iso}|"
    for k in list(memo):
        if str(k).startswith(pfx):
            memo.pop(k, None)

WORKLOG_DIR = os.path.join("uploaded_cache", "worklog")
WORKLOG_TEMPLATE = os.path.join(WORKLOG_DIR, "template.xlsx")
WORKLOG_TEMPLATE_SRC = os.path.expanduser("~/Desktop/업무일지.xlsx")
WORKLOG_ARCHIVE_REL = os.path.join("Desktop", "업무", "일지")

# =====================================================================
# 💡 원본.xlsx 양식 기준 행/열 매핑 (인쇄·미리보기 일치)
# 제목 E2 / 날짜 C5 / 내용헤더 G7 / 비고 Y7:AB / 본문 8~39 / 익일 40~43 / 특이 44~47
# =====================================================================
WL_MIN_ROW, WL_MAX_ROW = 1, 47  # 본문·익일·특이까지(원본 로고는 47행 아래 여백)
WL_MIN_COL, WL_MAX_COL = 3, 28  # C ~ AB

WL_DATE_CELL = "C5"
WL_CLIENT_ROWS = list(range(8, 40))   # 8행 ~ 39행 (총 32줄)
WL_CONTENT_ROWS = list(range(8, 40))  # 8행 ~ 39행 (총 32줄)
WL_SHEET_N = len(WL_CONTENT_ROWS)     # 본문 32칸 (익일·특이 제외)
WL_MAX_PAGES = 12                     # 업무입력 1페이지 + 추가 페이지
WL_NEXT_ROWS = list(range(40, 44))    # 40행 ~ 43행 (총 4줄)
WL_NOTE_ROWS = list(range(44, 48))    # 44행 ~ 47행 (총 4줄)

WL_CONTENT_COL_START = 7  # G
WL_CONTENT_COL_END = 24   # X
WL_CLIENT_COL_START = 3   # C
WL_CLIENT_COL_END = 6     # F
WL_REMARK_COL_START = 25  # Y — 내용 우측 입력칸(비고)
WL_REMARK_COL_END = 28    # AB

# 화면 엑셀 미리보기 배율 (인쇄 print_mode=True 와 무관 — 인쇄는 Excel pageSetup 그대로)
_WL_PREVIEW_SCALE = 0.65
# 업무입력 표시만 한 단계 크게. 줄바꿈·미리보기·인쇄는 원본 14pt.
_WL_INPUT_LOOK_SCALE = 0.75
# 웹/Mac: Batang(윈도우) 미설치 → 한글 글리프가 고딕으로 떨어짐.
# Nanum Myeongjo(CDN)를 최우선으로 두어 바탕체와 같은 명조 계열을 강제.
_WL_FONT_STACK = "'Nanum Myeongjo','Apple Myungjo','Batang','BatangChe','바탕체','바탕','바탕글',serif"
_WL_FONT_FACE_CSS = "@import url('https://fonts.googleapis.com/css2?family=Nanum+Myeongjo:wght@400;700&display=swap');"
# 로컬 반영 확인용 (탭 상단에 표시)
_WL_UI_BUILD = "2026-09-22 · 입력칸 여백활용 · 내용줄=미리보기"
_WL_MOVE_BLOCK_MSG = "이미 저장된 데이터가 있으면 자료를 옮길 수 없습니다."


class WorklogSaveBlockedError(Exception):
    """이미 저장된 날짜에 후입력 저장 시도."""


# =====================================================================
# 💡 [컴포넌트 복구] 기존 환경에 등록된 안전한 v2 컴포넌트 이름 유지
# =====================================================================

_WL_LINES_HTML = """
<div class="wl-lines"></div>
"""

# 줄바꿈은 원본 14pt 칸(8·37·8자). 입력 표시만 미리보기 배율로 줄여 작성하기 쉽게 한다.
_WL_LINES_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Nanum+Myeongjo:wght@400;700&display=swap');
.wl-lines { display: flex; flex-direction: column; width: 100%; border: 1px solid #94A3B8; border-radius: 4px; overflow: hidden; background: #fff; box-sizing: border-box; }
.wl-row { display: flex; align-items: center; width: 100%; border-bottom: 1px solid #E2E8F0; box-sizing: border-box; }
.wl-row:last-child { border-bottom: none; }
.wl-row input {
  flex: 1 1 auto;
  min-width: 0;
  width: 100%;
  height: var(--wl-row-h, 30px);
  padding: 0 1px;
  border: none;
  background: transparent;
  color: #0F172A;
  font-family: 'Nanum Myeongjo','Apple Myungjo','Batang','BatangChe','바탕체','바탕',serif !important;
  font-size: var(--wl-look-pt, var(--wl-show-pt, 14pt));
  line-height: var(--wl-row-h, 30px);
  outline: none;
  box-sizing: border-box;
}
.wl-lines.client .wl-row input { background: #F8FAFC; text-align: center; }
.wl-lines.remark .wl-row input { background: #FFFBEB; text-align: left; }
.wl-row input:focus { background: #E0F2FE; }
.wl-row.wl-sel input {
  background: #93C5FD !important;
  color: #0F172A;
}
.wl-row.wl-sel-part input {
  background: linear-gradient(to right,
    transparent var(--wl-sel-l, 0%),
    #93C5FD var(--wl-sel-l, 0%),
    #93C5FD var(--wl-sel-r, 100%),
    transparent var(--wl-sel-r, 100%)) !important;
}
"""

_WL_LINES_JS = r"""
const __wlLinesInst = new WeakMap();

export default function (component) {
  const { data, parentElement, setStateValue } = component;
  const root = parentElement.querySelector(".wl-lines");
  if (!root) return;

  const maxU = Number((data && data.max_u) || 64);
  const cellW = Number((data && data.cell_w) || 666);
  const fontPt = Number((data && data.font_pt) || 14);
  const lookScale = Number((data && data.look_scale) || 0.65);
  const origW = cellW > 8 ? Math.max(8, cellW - 2) : 0;
  const _wlOrigCanvas = document.createElement("canvas");
  const variant = String((data && data.variant) || "content");
  const side8 = variant === "client" || variant === "remark";
  const iso = String((data && data.iso) || "");
  const slot = String((data && data.slot) != null ? data.slot : "0");
  const rev = Number((data && data.rev) || 0);
  const focusReq = Number((data && data.focus));
  const fixedRows = Math.max(0, Number((data && data.fixed_rows) || 0));
  const incoming = Array.isArray(data && data.lines)
    ? data.lines.map((x) => String(x ?? ""))
    : [""];
  const replace = Number((data && data.replace) || 0) === 1;
  const memKey = variant + "|" + slot + "|" + String(fixedRows) + "|" + iso;
  const mem = (window.__wlLinesMem = window.__wlLinesMem || {});

  try {
    root.className = "wl-lines" + (variant === "client" ? " client" : variant === "remark" ? " remark" : "");
    root.style.width = "100%";
    root.style.maxWidth = "100%";
    root.style.minWidth = "";
    root.style.setProperty("--wl-show-pt", fontPt + "pt");
    root.style.setProperty("--wl-row-h", "30px");
  } catch (e0) {}

  let inst = __wlLinesInst.get(root);
  const isoChanged = !inst || String(inst.iso || "") !== iso;
  if (isoChanged) {
    try { if (inst && inst.ro) inst.ro.disconnect(); } catch (eD) {}
    try { if (inst && inst.dragOff) inst.dragOff(); } catch (eDrag) {}
    inst = { lines: null, rev: null, rebuilding: false, lastEmitted: null, iso: iso, ro: null, dragOff: null, drag: null, dragSel: null };
    __wlLinesInst.set(root, inst);
  }
  mem[memKey] = inst;

  function charUnits(ch) {
    if (ch === " " || ch === "\t" || ch === "\u00a0") return 0.58;
    const o = ch.charCodeAt(0);
    if (
      (o >= 0xac00 && o <= 0xd7a3) ||
      (o >= 0x1100 && o <= 0x11ff) ||
      (o >= 0x3130 && o <= 0x318f) ||
      (o >= 0x2e80 && o <= 0x9fff) ||
      (o >= 0xff00 && o <= 0xffef)
    )
      return 2;
    return 1.1;
  }
  function lstripWs(s) {
    return String(s || "").replace(/^[\s\u00a0\u3000]+/, "");
  }
  function displayUnits(s) {
    let w = 0;
    s = s || "";
    for (let i = 0; i < s.length; i++) w += charUnits(s.charAt(i));
    return w;
  }
  function measureOrigPx(s) {
    try {
      const ctx = _wlOrigCanvas.getContext("2d");
      ctx.font = "400 " + fontPt + "pt 'Nanum Myeongjo','Apple Myungjo','Batang','BatangChe','바탕체',serif";
      return ctx.measureText(s || "").width;
    } catch (eM) {
      return 0;
    }
  }
  function inputInnerW() {
    const el = root.querySelector("input");
    if (!el) return Math.max(0, root.clientWidth || 0);
    try {
      const cs = window.getComputedStyle(el);
      const pad = (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0);
      return Math.max(0, (el.clientWidth || 0) - pad);
    } catch (eW) {
      return Math.max(0, el.clientWidth || 0);
    }
  }
  function applyFillScale() {
    const scale = lookScale > 0.2 && lookScale < 1 ? lookScale : 0.65;
    const look = Math.max(7, fontPt * scale);
    root.style.setProperty("--wl-show-pt", look + "pt");
    root.style.setProperty("--wl-look-pt", look + "pt");
    root.style.width = "100%";
    root.style.maxWidth = "100%";
  }
  function measureShowPx(s) {
    try {
      const ctx = _wlOrigCanvas.getContext("2d");
      const show = root.style.getPropertyValue("--wl-show-pt") || (fontPt + "pt");
      ctx.font = "400 " + show + " 'Nanum Myeongjo','Apple Myungjo','Batang','BatangChe','바탕체',serif";
      return ctx.measureText(s || "").width;
    } catch (eS) {
      return 0;
    }
  }
  function lineOver(s) {
    // 거래처·비고는 원본 한글 8자(maxU)가 한 줄에 들어가야 한다.
    if (side8 && displayUnits(s) <= maxU) return false;
    if (origW > 8) {
      const px = measureOrigPx(s);
      if (px > 0 && px > origW) return true;
    }
    // 내용칸: 입력칸이 넓어도 인쇄미리보기(원본 14pt)와 같은 글자수
    if (!side8) return displayUnits(s) > maxU;
    const w = inputInnerW();
    if (w > 8) {
      const px = measureShowPx(s);
      if (px > 0 && px > w) return true;
    }
    return displayUnits(s) > maxU;
  }
  function fitByOrigPx(s, maxPx) {
    if (!s) return { head: "", tail: "" };
    if (measureOrigPx(s) <= maxPx) return { head: s, tail: "" };
    let lo = 0;
    let hi = s.length;
    while (lo < hi) {
      const mid = Math.ceil((lo + hi) / 2);
      if (measureOrigPx(s.slice(0, mid)) <= maxPx) lo = mid;
      else hi = mid - 1;
    }
    if (lo <= 0) return { head: s.slice(0, 1), tail: lstripWs(s.slice(1)) };
    return { head: s.slice(0, lo), tail: lstripWs(s.slice(lo)) };
  }
  function fitByUnits(s, max) {
    if (!s) return { head: "", tail: "" };
    if (displayUnits(s) <= max) return { head: s, tail: "" };
    let acc = 0;
    for (let i = 0; i < s.length; i++) {
      const cu = charUnits(s.charAt(i));
      if (acc + cu > max) {
        if (i === 0) return { head: s.slice(0, 1), tail: lstripWs(s.slice(1)) };
        return { head: s.slice(0, i), tail: lstripWs(s.slice(i)) };
      }
      acc += cu;
    }
    return { head: s, tail: "" };
  }
  function fitByShowPx(s, maxPx) {
    if (!s) return { head: "", tail: "" };
    if (measureShowPx(s) <= maxPx) return { head: s, tail: "" };
    let lo = 0;
    let hi = s.length;
    while (lo < hi) {
      const mid = Math.ceil((lo + hi) / 2);
      if (measureShowPx(s.slice(0, mid)) <= maxPx) lo = mid;
      else hi = mid - 1;
    }
    if (lo <= 0) return { head: s.slice(0, 1), tail: lstripWs(s.slice(1)) };
    return { head: s.slice(0, lo), tail: lstripWs(s.slice(lo)) };
  }
  function fitLine(s) {
    if (side8 && displayUnits(s) <= maxU) return { head: s, tail: "" };
    if (origW > 8 && measureOrigPx(s) > origW) return fitByOrigPx(s, origW);
    if (!side8) return fitByUnits(s, maxU);
    const w = inputInnerW();
    if (w > 8 && measureShowPx(s) > w) return fitByShowPx(s, w);
    return fitByUnits(s, maxU);
  }
  function normalize(arr) {
    let out = (arr || []).map((x) => String(x ?? ""));
    if (fixedRows > 0) {
      while (out.length < fixedRows) out.push("");
      if (out.length > fixedRows) out = out.slice(0, fixedRows);
      return out;
    }
    if (!out.length) out.push("");
    if (out[out.length - 1] !== "") out.push("");
    return out;
  }
  function fillEmptyFrom(src, dest) {
    if (!src || !src.length) return dest;
    const out = dest.slice();
    for (let i = 0; i < out.length && i < src.length; i++) {
      if (!(out[i] || "") && (src[i] || "")) out[i] = src[i];
    }
    return out;
  }
  function mergeIncoming(next) {
    let out = normalize(next);
    const inputs = root.querySelectorAll("input[data-idx]");
    if (inputs.length) {
      const dom = [];
      inputs.forEach((inp) => dom.push(inp.value || ""));
      out = fillEmptyFrom(normalize(dom), out);
    }
    out = fillEmptyFrom(inst.lines, out);
    return out;
  }
  function readDomLines() {
    const inputs = root.querySelectorAll("input[data-idx]");
    if (!inputs.length) return normalize(inst.lines || incoming);
    const out = [];
    inputs.forEach((inp) => out.push(inp.value || ""));
    return normalize(out);
  }
  function emit(next, focusIdx) {
    const out = normalize(next);
    inst.lines = out;
    const sig = out.join("\n");
    // 값이 같으면 Python rerun을 만들지 않는다.
    // (blur/클릭마다 setStateValue → 저장·추가 버튼 클릭이 삼켜지거나 Cached ForwardMsg MISS)
    if (sig === inst.lastEmitted && typeof focusIdx !== "number") {
      return out;
    }
    inst.lastEmitted = sig;
    setStateValue("lines", out);
    if (typeof focusIdx === "number") setStateValue("focus", focusIdx);
    return out;
  }
  function localOnly(next) {
    inst.lines = normalize(next);
    return inst.lines;
  }
  function focusAt(idx) {
    requestAnimationFrame(() => {
      const el = root.querySelector('input[data-idx="' + idx + '"]');
      if (!el) return;
      try {
        el.focus({ preventScroll: true });
        const n = (el.value || "").length;
        el.setSelectionRange(n, n);
      } catch (e) {
        try {
          el.focus();
        } catch (e2) {}
      }
      if (typeof window.wlScrollCellIntoTabView === "function") {
        window.wlScrollCellIntoTabView(el);
      } else {
        try {
          const tabs = document.querySelector('[data-testid="stTabs"] [role="tablist"]');
          const topLimit = Math.max(tabs ? tabs.getBoundingClientRect().bottom : 0, 56) + 10;
          const rect = el.getBoundingClientRect();
          if (rect.top < topLimit) window.scrollBy(0, rect.top - topLimit);
          else if (rect.bottom > window.innerHeight - 16) window.scrollBy(0, rect.bottom - (window.innerHeight - 16));
        } catch (e3) {}
      }
    });
  }
  function shouldApplyDataFocus() {
    if (!Number.isFinite(focusReq) || focusReq < 0) return false;
    const ae = document.activeElement;
    if (ae && root.contains(ae) && String(ae.tagName || "").toUpperCase() === "INPUT") {
      const curIdx = Number(ae.dataset && ae.dataset.idx);
      if (Number.isFinite(curIdx) && curIdx !== Number(focusReq)) return false;
    }
    return true;
  }

  function offsetFromX(inp, clientX) {
    if (!inp) return 0;
    const s = inp.value || "";
    try {
      const rect = inp.getBoundingClientRect();
      const cs = window.getComputedStyle(inp);
      const padL = parseFloat(cs.paddingLeft) || 0;
      const padR = parseFloat(cs.paddingRight) || 0;
      let x = clientX - rect.left - padL;
      const inner = Math.max(1, (inp.clientWidth || 0) - padL - padR);
      if ((cs.textAlign || "") === "center") {
        const tw = measureShowPx(s);
        x = x - Math.max(0, (inner - tw) / 2);
      }
      if (x <= 0) return 0;
      let lo = 0;
      let hi = s.length;
      while (lo < hi) {
        const mid = Math.ceil((lo + hi) / 2);
        if (measureShowPx(s.slice(0, mid)) <= x) lo = mid;
        else hi = mid - 1;
      }
      return lo;
    } catch (eOff) {
      try { return Number(inp.selectionStart || 0); } catch (e2) { return s.length; }
    }
  }
  function inputAtPoint(x, y) {
    let el = null;
    try { el = document.elementFromPoint(x, y); } catch (eP) { return null; }
    while (el) {
      if (el.tagName === "INPUT" && root.contains(el)) return el;
      el = el.parentElement;
    }
    const rows = root.querySelectorAll(".wl-row input");
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i].getBoundingClientRect();
      if (y >= r.top && y <= r.bottom) return rows[i];
    }
    return null;
  }
  function normSel(sel) {
    if (!sel) return null;
    let aIdx = Number(sel.aIdx), aOff = Number(sel.aOff), bIdx = Number(sel.bIdx), bOff = Number(sel.bOff);
    if (aIdx > bIdx || (aIdx === bIdx && aOff > bOff)) {
      const t = aIdx; aIdx = bIdx; bIdx = t;
      const o = aOff; aOff = bOff; bOff = o;
    }
    return { aIdx, aOff, bIdx, bOff };
  }
  function clearPaintSel() {
    root.querySelectorAll(".wl-row").forEach((row) => {
      row.classList.remove("wl-sel", "wl-sel-part");
      row.style.removeProperty("--wl-sel-l");
      row.style.removeProperty("--wl-sel-r");
    });
  }
  function paintSel() {
    clearPaintSel();
    const sel = normSel(inst.dragSel);
    if (!sel) return;
    const rows = root.querySelectorAll(".wl-row");
    for (let i = sel.aIdx; i <= sel.bIdx; i++) {
      const row = rows[i];
      if (!row) continue;
      const inp = row.querySelector("input");
      const text = (inp && inp.value) || "";
      const from = i === sel.aIdx ? sel.aOff : 0;
      const to = i === sel.bIdx ? sel.bOff : text.length;
      if (from <= 0 && to >= text.length) {
        row.classList.add("wl-sel");
      } else {
        row.classList.add("wl-sel-part");
        const inner = inputInnerW() || 1;
        const tw = measureShowPx(text);
        let base = 0;
        try {
          if (inp && window.getComputedStyle(inp).textAlign === "center") {
            base = Math.max(0, (inner - tw) / 2);
          }
        } catch (eC) {}
        const l = Math.max(0, ((base + measureShowPx(text.slice(0, from))) / inner) * 100);
        const r = Math.min(100, ((base + measureShowPx(text.slice(0, to))) / inner) * 100);
        row.style.setProperty("--wl-sel-l", l + "%");
        row.style.setProperty("--wl-sel-r", r + "%");
      }
    }
  }
  function selectedText() {
    const sel = normSel(inst.dragSel);
    if (!sel) return "";
    const lines = readDomLines();
    if (sel.aIdx === sel.bIdx) return (lines[sel.aIdx] || "").slice(sel.aOff, sel.bOff);
    const out = [(lines[sel.aIdx] || "").slice(sel.aOff)];
    for (let i = sel.aIdx + 1; i < sel.bIdx; i++) out.push(lines[i] || "");
    out.push((lines[sel.bIdx] || "").slice(0, sel.bOff));
    return out.join("\n");
  }
  function hasDragSel() {
    const sel = normSel(inst.dragSel);
    return !!(sel && (sel.aIdx !== sel.bIdx || sel.aOff !== sel.bOff));
  }
  function applyLinesToDom(lines, focusIdx) {
    const out = normalize(lines);
    inst.lines = out;
    const inputs = root.querySelectorAll("input[data-idx]");
    if (inputs.length === out.length && inputs.length > 0) {
      for (let k = 0; k < inputs.length; k++) {
        if (inputs[k].value !== out[k]) inputs[k].value = out[k];
      }
    } else {
      rebuild(typeof focusIdx === "number" ? focusIdx : -1);
    }
    emit(out, typeof focusIdx === "number" ? focusIdx : null);
    if (typeof focusIdx === "number") focusAt(focusIdx);
  }
  function deleteSelected() {
    const sel = normSel(inst.dragSel);
    if (!sel) return false;
    const lines = readDomLines();
    if (sel.aIdx === sel.bIdx) {
      const s = lines[sel.aIdx] || "";
      lines[sel.aIdx] = s.slice(0, sel.aOff) + s.slice(sel.bOff);
    } else {
      lines[sel.aIdx] = (lines[sel.aIdx] || "").slice(0, sel.aOff) + (lines[sel.bIdx] || "").slice(sel.bOff);
      for (let i = sel.aIdx + 1; i <= sel.bIdx; i++) lines[i] = "";
    }
    inst.dragSel = null;
    clearPaintSel();
    applyLinesToDom(lines, sel.aIdx);
    return true;
  }
  function bindDragSelect() {
    if (inst.dragOff) return;
    const onDown = (e) => {
      if (e.button !== 0) return;
      const inp = e.target && e.target.tagName === "INPUT" ? e.target : null;
      if (!inp || !root.contains(inp)) return;
      inst.drag = { aIdx: Number(inp.dataset.idx || 0), aOff: offsetFromX(inp, e.clientX), moved: false };
    };
    const onMove = (e) => {
      if (!inst.drag || (e.buttons & 1) === 0) return;
      const inp = inputAtPoint(e.clientX, e.clientY);
      if (!inp) return;
      const bIdx = Number(inp.dataset.idx || 0);
      const bOff = offsetFromX(inp, e.clientX);
      if (bIdx !== inst.drag.aIdx || Math.abs(bOff - inst.drag.aOff) > 1) inst.drag.moved = true;
      if (!inst.drag.moved) return;
      try { e.preventDefault(); } catch (ePrev) {}
      inst.dragSel = { aIdx: inst.drag.aIdx, aOff: inst.drag.aOff, bIdx, bOff };
      paintSel();
    };
    const onUp = () => {
      if (inst.drag && !inst.drag.moved) {
        inst.dragSel = null;
        clearPaintSel();
      }
      inst.drag = null;
    };
    const onCopy = (e) => {
      if (!hasDragSel()) return;
      const t = selectedText();
      if (!t) return;
      try {
        e.preventDefault();
        e.clipboardData.setData("text/plain", t);
      } catch (eCpy) {}
    };
    const onCut = (e) => {
      if (!hasDragSel()) return;
      const t = selectedText();
      try {
        e.preventDefault();
        if (t) e.clipboardData.setData("text/plain", t);
      } catch (eCut) {}
      deleteSelected();
    };
    const onKey = (e) => {
      const inEditor = root.contains(document.activeElement) || !!inst.dragSel;
      if (!inEditor) return;
      if ((e.metaKey || e.ctrlKey) && String(e.key || "").toLowerCase() === "a") {
        const lines = readDomLines();
        let last = lines.length - 1;
        while (last > 0 && !(lines[last] || "")) last -= 1;
        inst.dragSel = { aIdx: 0, aOff: 0, bIdx: last, bOff: (lines[last] || "").length };
        paintSel();
        try { e.preventDefault(); } catch (eA) {}
        return;
      }
      if (hasDragSel() && (e.key === "Backspace" || e.key === "Delete")) {
        try { e.preventDefault(); } catch (eDel) {}
        deleteSelected();
      }
    };
    root.addEventListener("mousedown", onDown);
    document.addEventListener("mousemove", onMove, true);
    document.addEventListener("mouseup", onUp, true);
    document.addEventListener("copy", onCopy, true);
    document.addEventListener("cut", onCut, true);
    document.addEventListener("keydown", onKey, true);
    inst.dragOff = function () {
      try { root.removeEventListener("mousedown", onDown); } catch (e1) {}
      try { document.removeEventListener("mousemove", onMove, true); } catch (e2) {}
      try { document.removeEventListener("mouseup", onUp, true); } catch (e3) {}
      try { document.removeEventListener("copy", onCopy, true); } catch (e4) {}
      try { document.removeEventListener("cut", onCut, true); } catch (e5) {}
      try { document.removeEventListener("keydown", onKey, true); } catch (e6) {}
    };
  }

  function insertLineAfter(j, inputEl) {
    const raw = [];
    root.querySelectorAll("input[data-idx]").forEach((inp) => {
      raw.push(inp.value || "");
    });
    if (!raw.length) raw.push("");
    while (raw.length <= j) raw.push("");
    raw[j] = (inputEl && inputEl.value) || raw[j] || "";
    if (fixedRows > 0) {
      const next = Math.min(j + 1, Math.max(fixedRows - 1, 0));
      emit(raw, next);
      focusAt(next);
      return;
    }
    raw.splice(j + 1, 0, "");
    if (raw[raw.length - 1] !== "") raw.push("");
    emit(raw, j + 1);
    rebuild(j + 1);
  }

  function rebuild(focusIdx) {
    inst.rebuilding = true;
    root.innerHTML = "";
    const lines = normalize(inst.lines || incoming);
    inst.lines = lines;
    lines.forEach((line, j) => {
      const row = document.createElement("div");
      row.className = "wl-row";
      const input = document.createElement("input");
      input.type = "text";
      input.value = line;
      input.placeholder = "";
      input.dataset.idx = String(j);
      const commitValue = (mode) => {
        if (inst.rebuilding) return;
        const cur = readDomLines();
        let j0 = j;
        let v = cur[j0] || "";
        if (!lineOver(v)) {
          // 입력 중에는 로컬만. 한글 음절마다 setStateValue 하면
          // Cached ForwardMsg MISS → ERROR → 버튼 전부 먹통.
          if (mode === "blur" || mode === "force") emit(cur, null);
          else localOnly(cur);
          return;
        }
        while (j0 < cur.length && lineOver(cur[j0] || "")) {
          const ft = fitLine(cur[j0] || "");
          if (!ft.tail || ft.head === (cur[j0] || "")) break;
          cur[j0] = ft.head;
          const nextTail = lstripWs(ft.tail);
          if (j0 + 1 < cur.length) cur[j0 + 1] = nextTail + (cur[j0 + 1] || "");
          else if (fixedRows > 0) break;
          else cur.splice(j0 + 1, 0, nextTail);
          j0 += 1;
        }
        const focusIdx = Math.min(j0, Math.max(cur.length - 1, 0));
        emit(cur, focusIdx);
        rebuild(focusIdx);
      };
      input.addEventListener("input", (e) => {
        if (e.isComposing) return;
        inst.dragSel = null;
        clearPaintSel();
        commitValue("type");
      });
      input.addEventListener("compositionend", () => {
        commitValue("type");
      });
      input.addEventListener("blur", () => {
        commitValue("blur");
      });
      input.addEventListener("keydown", (e) => {
        const nav = (typeof window !== "undefined" && window.wlNavWorklogArrow)
          || (typeof window !== "undefined" && window.parent && window.parent !== window && window.parent.wlNavWorklogArrow);
        if (nav && (e.key === "ArrowUp" || e.key === "ArrowDown" || e.key === "ArrowLeft" || e.key === "ArrowRight")) {
          try {
            if (nav(input, e)) return;
          } catch (eNav) {}
        }
        if (e.key === "ArrowUp" || e.key === "ArrowDown") {
          const dest = e.key === "ArrowUp" ? j - 1 : j + 1;
          const n = (inst.lines || lines || []).length;
          if (dest >= 0 && dest < n) {
            e.preventDefault();
            e.stopPropagation();
            focusAt(dest);
          }
          return;
        }
        if (e.key !== "Enter") return;
        const typing = (input.value || "") !== "";
        if (e.isComposing && typing) return;
        e.preventDefault();
        e.stopPropagation();
        insertLineAfter(j, input);
      });
      row.appendChild(input);
      root.appendChild(row);
    });
    inst.rebuilding = false;
    applyFillScale();
    if (typeof focusIdx === "number" && focusIdx >= 0) focusAt(focusIdx);
  }

  inst.iso = iso;
  if (inst.rev !== rev || replace || isoChanged) {
    inst.rev = rev;
    // 날짜 전환·프로그램 시드는 전날 DOM/메모를 빈칸에 되살리지 않는다.
    const next = (replace || isoChanged) ? normalize(incoming) : mergeIncoming(incoming);
    const inputs = root.querySelectorAll("input[data-idx]");
    // 행 수가 같으면 innerHTML 전체 재생성 없이 값만 제자리로 갱신한다.
    // (특수기호 삽입 등으로 rev만 오를 때 전체 rebuild 하면 입력칸이 잠깐 높이 0으로
    //  무너졌다 복구돼 "밑으로 껌벅 내려갔다 올라오는" 현상이 생긴다.)
    if (inputs.length === next.length && inputs.length > 0) {
      inst.lines = next;
      inst.lastEmitted = next.join("\n");
      for (let k = 0; k < inputs.length; k++) {
        if (inputs[k].value !== next[k]) inputs[k].value = next[k];
      }
      if (shouldApplyDataFocus()) focusAt(focusReq);
      applyFillScale();
    } else {
      inst.lines = next;
      inst.lastEmitted = next.join("\n");
      rebuild(shouldApplyDataFocus() ? focusReq : -1);
    }
  } else if (!root.childElementCount) {
    inst.lines = inst.lines ? mergeIncoming(incoming) : normalize(incoming);
    rebuild(shouldApplyDataFocus() ? focusReq : -1);
  } else {
    inst.lines = mergeIncoming(incoming);
    applyFillScale();
  }
  if (!inst.ro) {
    try {
      inst.ro = new ResizeObserver(function () { applyFillScale(); });
      inst.ro.observe(root);
    } catch (eRo) {}
  }
  bindDragSelect();
}
"""

_WL_LINES_EDITOR = st.components.v2.component(
    "worklog_entry_lines_v53",
    html=_WL_LINES_HTML,
    css=_WL_LINES_CSS,
    js=_WL_LINES_JS,
)

_WL_ENTER_HOOK_JS = r"""
export default function (component) {
  const { data, setTriggerValue } = component;
  const iso = (data && data.iso) || "";
  const focusKey = (data && data.focus_key) || "";
  const focusCaret = data && data.focus_caret != null && data.focus_caret !== "" ? Number(data.focus_caret) : null;
  let lastSent = ""; let lastSig = ""; let lastAt = 0;
  // remount 때마다 document 리스너가 쌓이면 click/focus마다 setTriggerValue가
  // 폭주하고 Cached ForwardMsg MISS → 저장·추가 등 버튼이 전부 먹통이 된다.
  if (typeof window !== "undefined" && window.__wlEnterHookOff) {
    try { window.__wlEnterHookOff(); } catch (eOff) {}
    window.__wlEnterHookOff = null;
  }
  if (typeof window !== "undefined" && typeof window.__wlApplyScrollPin === "function") {
    try { window.__wlApplyScrollPin(); } catch (ePin) {}
  }

  function walkHosts(el) {
    const out = [];
    let n = el;
    while (n && n !== document && n !== window) {
      out.push(n);
      const root = n.getRootNode && n.getRootNode();
      if (root && root !== document && root.host) {
        n = root.host;
        continue;
      }
      n = n.parentElement || null;
    }
    return out;
  }

  function eventTargetInput(e) {
    const path = (e.composedPath && e.composedPath()) || [e.target];
    for (let i = 0; i < path.length; i++) {
      const n = path[i];
      if (!n || !n.tagName) continue;
      const tag = String(n.tagName).toUpperCase();
      if (tag === "INPUT" || tag === "TEXTAREA") return n;
    }
    return e.target;
  }

  function queryDeep(root, selector) {
    const out = [];
    const seen = new Set();
    function walk(node) {
      if (!node || seen.has(node)) return;
      seen.add(node);
      if (node.querySelectorAll) {
        try {
          node.querySelectorAll(selector).forEach((el) => out.push(el));
        } catch (e1) {}
      }
      const kids = node.querySelectorAll ? node.querySelectorAll("*") : [];
      for (let i = 0; i < kids.length; i++) {
        if (kids[i].shadowRoot) walk(kids[i].shadowRoot);
      }
    }
    walk(root);
    return out;
  }

  function slotOk(part) {
    const s = String(part || "");
    return s === String(iso || "") || s === "ui";
  }

  function resolveKey(t) {
    if (!t) return null;
    const tag = String(t.tagName || "").toUpperCase();
    if (tag === "TEXTAREA") {
      const wrap = t.closest ? t.closest('[class*="st-key-wl_next_area_"],[class*="st-key-wl_notes_area_"]') : null;
      if (!wrap) return null;
      const cls = Array.prototype.find.call(wrap.classList || [], (c) => {
        const s = String(c);
        return s.indexOf("st-key-wl_next_area_") !== -1 || s.indexOf("st-key-wl_notes_area_") !== -1;
      });
      if (!cls) return null;
      const key = String(cls).replace(/^st-key-/, "");
      const m = /^(wl_(?:next|notes)_area)_(\d{4}-\d{2}-\d{2})$/.exec(key);
      if (!m || m[2] !== iso) return null;
      return { key: key, kind: m[1], ei: -1, lj: -1 };
    }
    if (tag !== "INPUT") return null;
    const chain = walkHosts(t);
    for (let i = 0; i < chain.length; i++) {
      const wrap = chain[i];
      if (!wrap.classList) continue;
      const cls = Array.prototype.find.call(wrap.classList || [], (c) => {
        const s = String(c);
        return s.indexOf("st-key-wl_ent_ln_") !== -1 || s.indexOf("st-key-wl_ent_cl_") !== -1 || s.indexOf("st-key-wl_ent_rm_") !== -1;
      });
      if (cls) {
        const key = String(cls).replace(/^st-key-/, "");
        const m = /^(wl_ent_ln|wl_ent_cl|wl_ent_rm)_(\d{4}-\d{2}-\d{2}|ui)_(\d+)_(\d+)(?:_g\d+)?$/.exec(key);
        if (m && slotOk(m[2])) return { key: key, kind: m[1], ei: Number(m[3]), lj: Number(m[4]) };
      }
    }
    let wrap = null;
    let cls2 = null;
    for (let i = 0; i < chain.length; i++) {
      const n = chain[i];
      if (!n.classList) continue;
      cls2 = Array.prototype.find.call(n.classList || [], (c) => {
        const s = String(c);
        return s.indexOf("st-key-wl_lines_comp_") !== -1 || s.indexOf("st-key-wl_clients_comp_") !== -1 || s.indexOf("st-key-wl_remarks_comp_") !== -1;
      });
      if (cls2) { wrap = n; break; }
    }
    if (!wrap || !cls2) return null;
    const compKey = String(cls2).replace(/^st-key-/, "");
    const m2 = /^(wl_(?:lines|clients|remarks)_comp)_(\d{4}-\d{2}-\d{2}|ui)_(\d+)/.exec(compKey);
    if (!m2 || !slotOk(m2[2])) return null;
    const kind = m2[1] === "wl_clients_comp" ? "wl_ent_cl" : m2[1] === "wl_remarks_comp" ? "wl_ent_rm" : "wl_ent_ln";
    const ei = Number(m2[3]);
    const lj = Number(t.dataset && t.dataset.idx != null ? t.dataset.idx : 0);
    const logicalKey = kind + "_" + iso + "_" + ei + "_" + lj;
    return { key: logicalKey, kind: kind, ei: ei, lj: lj };
  }

  const COLS = ["wl_ent_cl", "wl_ent_ln", "wl_ent_rm"];

  function listKindInputs(kind) {
    const needle = kind === "wl_ent_cl" ? "st-key-wl_ent_cl_" : kind === "wl_ent_rm" ? "st-key-wl_ent_rm_" : "st-key-wl_ent_ln_";
    const compNeedle = kind === "wl_ent_cl" ? "st-key-wl_clients_comp_" : kind === "wl_ent_rm" ? "st-key-wl_remarks_comp_" : "st-key-wl_lines_comp_";
    const wraps = queryDeep(document, 'div[class*="' + needle + '"],div[class*="' + compNeedle + '"]');
    const out = [];
    const seen = new Set();
    for (let w = 0; w < wraps.length; w++) {
      const nodes = queryDeep(wraps[w], "input[data-idx], input");
      for (let i = 0; i < nodes.length; i++) {
        const el = nodes[i];
        if (seen.has(el)) continue;
        const info = resolveKey(el);
        if (!info || info.kind !== kind) continue;
        seen.add(el);
        out.push(el);
      }
    }
    out.sort((a, b) => {
      const ia = resolveKey(a);
      const ib = resolveKey(b);
      return (ia ? ia.lj : 0) - (ib ? ib.lj : 0);
    });
    return out;
  }

  function findCell(kind, lj, ei) {
    const list = listKindInputs(kind);
    let fallback = null;
    for (let i = 0; i < list.length; i++) {
      const p = resolveKey(list[i]);
      if (!p || Number(p.lj) !== Number(lj)) continue;
      if (ei == null || !Number.isFinite(Number(ei)) || Number(p.ei) === Number(ei)) return list[i];
      if (!fallback) fallback = list[i];
    }
    return fallback;
  }

  function findInputForFocusKey(focusKey) {
    if (!focusKey) return null;
    let m = /^(wl_ent_ln|wl_ent_cl|wl_ent_rm)_(\d{4}-\d{2}-\d{2}|ui)_(\d+)_(\d+)(?:_g\d+)?$/.exec(focusKey);
    if (m) {
      const kind = m[1];
      const ei = m[3];
      const lj = m[4];
      const base = kind === "wl_ent_cl" ? "st-key-wl_clients_comp_" : kind === "wl_ent_rm" ? "st-key-wl_remarks_comp_" : "st-key-wl_lines_comp_";
      const slots = [m[2], "ui", iso];
      for (let s = 0; s < slots.length; s++) {
        const slot = String(slots[s] || "");
        if (!slot) continue;
        const wraps = queryDeep(document, 'div[class*="' + base + slot + "_" + ei + '"]');
        for (let i = 0; i < wraps.length; i++) {
          const found = queryDeep(wraps[i], '.wl-lines input[data-idx="' + lj + '"]');
          if (found.length) return found[0];
        }
      }
      const legacy = document.querySelector('div[class*="st-key-' + focusKey + '"] input');
      if (legacy) return legacy;
    }
    return document.querySelector(
      'div[class*="st-key-' + focusKey + '"] input, div[class*="st-key-' + focusKey + '"] textarea'
    );
  }

  function findScroller(start) {
    let n = start;
    const chain = [];
    while (n && n !== document) {
      chain.push(n);
      const root = n.getRootNode && n.getRootNode();
      if (root && root !== document && root.host) {
        n = root.host;
        continue;
      }
      n = n.parentElement;
    }
    for (let i = 0; i < chain.length; i++) {
      const node = chain[i];
      if (!node || node.nodeType !== 1) continue;
      try {
        const st = window.getComputedStyle(node);
        const oy = st.overflowY || "";
        if ((oy === "auto" || oy === "scroll" || oy === "overlay") && node.scrollHeight > node.clientHeight + 8) {
          return node;
        }
      } catch (e0) {}
    }
    return document.querySelector("section.main") || document.scrollingElement || document.documentElement;
  }

  function scrollCellIntoTabView(el) {
    if (!el || !el.getBoundingClientRect) return;
    const tabs = document.querySelector('[data-testid="stTabs"] [role="tablist"]');
    const tabBottom = tabs ? tabs.getBoundingClientRect().bottom : 0;
    const topLimit = Math.max(tabBottom, 56) + 10;
    const bottomLimit = (window.innerHeight || 800) - 16;
    const rect = el.getBoundingClientRect();
    let delta = 0;
    if (rect.top < topLimit) delta = rect.top - topLimit;
    else if (rect.bottom > bottomLimit) delta = rect.bottom - bottomLimit;
    if (!delta) return;
    const scroller = findScroller(el);
    if (scroller && scroller !== document.documentElement && scroller !== document.body && scroller !== document.scrollingElement) {
      scroller.scrollTop += delta;
    } else {
      window.scrollBy(0, delta);
    }
  }
  if (typeof window !== "undefined") window.wlScrollCellIntoTabView = scrollCellIntoTabView;

  function focusInput(el, caret) {
    if (!el) return false;
    try {
      el.focus({ preventScroll: true });
      const n = (el.value || "").length;
      let pos = n;
      if (caret === "start") pos = 0;
      else if (caret === "end") pos = n;
      else if (typeof caret === "number" && Number.isFinite(caret)) {
        pos = Math.max(0, Math.min(n, Math.floor(caret)));
      }
      el.setSelectionRange(pos, pos);
    } catch (e1) {
      try { el.focus(); } catch (e2) {}
    }
    try {
      scrollCellIntoTabView(el);
    } catch (e3) {}
    return true;
  }

  function findPeer(info, targetKind) {
    const list = listKindInputs(targetKind);
    let best = null; let bestScore = 1e9;
    for (let i = 0; i < list.length; i++) {
      const p = resolveKey(list[i]);
      if (!p) continue;
      const score = Math.abs(p.ei - info.ei) * 1000 + Math.abs(p.lj - info.lj);
      if (score < bestScore) {
        bestScore = score;
        best = list[i];
      }
    }
    return best;
  }

  function emit(key, value) {
    const now = Date.now();
    const v = String(value || "");
    const sig = key + "\\0" + v;
    if (sig === lastSig && now - lastAt < 320) return;
    lastSig = sig; lastAt = now;
    const payload = JSON.stringify({ key: key, t: now, v: v });
    if (payload === lastSent) return;
    lastSent = payload;
    setTriggerValue("enter", payload);
  }

  const onKey = (e) => {
    if (e.isComposing || e.keyCode === 229) return;
    const t = eventTargetInput(e);
    const info = resolveKey(t);
    if (!info) return;
    const start = typeof t.selectionStart === "number" ? t.selectionStart : 0;
    const end = typeof t.selectionEnd === "number" ? t.selectionEnd : start;
    const len = String(t.value || "").length;
    const caretAll = start === end;
    const tag = String(t.tagName || "").toUpperCase();

    if (e.key === "Enter") {
      // 익일업무·특이사항 textarea: Enter = 다음 줄, ⌘/Ctrl+Enter = 반영
      if (tag === "TEXTAREA") {
        if (e.metaKey || e.ctrlKey) {
          e.preventDefault();
          e.stopPropagation();
          emit(info.key, t.value || "");
        }
        return;
      }
      // 거래처·내용·비고는 CCv2가 다음 줄로 옮긴다. 여기서 가로채 다시 심으면 이전 칸 값이 빠진다.
      if (info.kind === "wl_ent_cl" || info.kind === "wl_ent_ln" || info.kind === "wl_ent_rm") return;
      e.preventDefault();
      e.stopPropagation();
      emit(info.key, t.value || "");
      return;
    }

    if (tag === "TEXTAREA") return;
    if (tryArrowNav(t, info, e)) return;
  };

  function tryArrowNav(t, info, e) {
    if (!t || !info || !e) return false;
    const start = typeof t.selectionStart === "number" ? t.selectionStart : 0;
    const end = typeof t.selectionEnd === "number" ? t.selectionEnd : start;
    const len = String(t.value || "").length;
    const caretAll = start === end;
    if (e.key === "ArrowUp") {
      const el = findCell(info.kind, info.lj - 1, info.ei);
      if (!el) return false;
      e.preventDefault();
      e.stopPropagation();
      focusInput(el, "end");
      return true;
    }
    if (e.key === "ArrowDown") {
      const el = findCell(info.kind, info.lj + 1, info.ei);
      if (!el) return false;
      e.preventDefault();
      e.stopPropagation();
      focusInput(el, "end");
      return true;
    }
    if (e.key === "ArrowLeft" && caretAll && start === 0) {
      const ci = COLS.indexOf(info.kind);
      let kind = info.kind;
      let lj = info.lj;
      if (ci > 0) kind = COLS[ci - 1];
      else if (info.lj > 0) { kind = COLS[2]; lj = info.lj - 1; }
      else return false;
      const el = findCell(kind, lj, info.ei);
      if (!el) return false;
      e.preventDefault();
      e.stopPropagation();
      focusInput(el, "end");
      return true;
    }
    if (e.key === "ArrowRight" && caretAll && start === len) {
      const ci = COLS.indexOf(info.kind);
      let kind = info.kind;
      let lj = info.lj;
      if (ci >= 0 && ci < COLS.length - 1) kind = COLS[ci + 1];
      else { kind = COLS[0]; lj = info.lj + 1; }
      const el = findCell(kind, lj, info.ei);
      if (!el) return false;
      e.preventDefault();
      e.stopPropagation();
      focusInput(el, "start");
      return true;
    }
    return false;
  }

  if (typeof window !== "undefined") {
    window.wlNavWorklogArrow = function (el, ev) {
      const info = resolveKey(el);
      if (!info) return false;
      return tryArrowNav(el, info, ev);
    };
  }

  document.addEventListener("keydown", onKey, true);

  let focusTimers = [];
  const clearFocusTimers = () => {
    for (let i = 0; i < focusTimers.length; i++) {
      try { clearTimeout(focusTimers[i]); } catch (eT) {}
    }
    focusTimers = [];
  };
  const onUserTakeover = (e) => {
    const t = eventTargetInput(e);
    if (resolveKey(t) || (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA"))) {
      clearFocusTimers();
    }
  };
  document.addEventListener("pointerdown", onUserTakeover, true);
  document.addEventListener("keydown", onUserTakeover, true);

  if (focusKey) {
    const go = () => {
      const active = document.activeElement;
      const tag = String((active && active.tagName) || "").toUpperCase();
      if (tag === "INPUT" || tag === "TEXTAREA") {
        const cur = resolveKey(active);
        if (cur && String(cur.key) === String(focusKey)) return true;
        return false;
      }
      const el = findInputForFocusKey(focusKey);
      if (!el) return false;
      const caret = focusCaret != null && Number.isFinite(focusCaret) ? focusCaret : "end";
      focusInput(el, caret);
      return true;
    };
    go();
    focusTimers.push(setTimeout(go, 50));
  }

  const off = () => {
    clearFocusTimers();
    document.removeEventListener("keydown", onKey, true);
    document.removeEventListener("pointerdown", onUserTakeover, true);
    document.removeEventListener("keydown", onUserTakeover, true);
  };
  if (typeof window !== "undefined") window.__wlEnterHookOff = off;
  return off;
}
"""

_WL_ENTER_HOOK = st.components.v2.component(
    "worklog_cell_nav_hook_v32",
    js=_WL_ENTER_HOOK_JS,
)

_WL_SCROLL_LOCK_JS = r"""
export default function (component) {
  const { data } = component;
  const iso = String((data && data.iso) || "");
  const pin = Number((data && data.pin) || 0);

  function scrollerList() {
    const out = [];
    const seen = new Set();
    const add = (el) => {
      if (!el || seen.has(el)) return;
      seen.add(el);
      out.push(el);
    };
    add(document.querySelector("section.main"));
    add(document.querySelector('[data-testid="stMain"]'));
    add(document.querySelector('[data-testid="stAppViewContainer"]'));
    add(document.querySelector(".stApp"));
    add(document.scrollingElement);
    add(document.documentElement);
    add(document.body);
    return out;
  }
  function readSnap() {
    const main = document.querySelector("section.main") || document.querySelector('[data-testid="stMain"]');
    return {
      win: window.scrollY || document.documentElement.scrollTop || document.body.scrollTop || 0,
      main: main ? (main.scrollTop || 0) : 0,
    };
  }
  function writeSnap(snap) {
    if (!snap) return;
    const y = Number(snap.win || 0);
    const m = Number(snap.main || 0);
    try { window.scrollTo(0, y); } catch (e0) {}
    try { document.documentElement.scrollTop = y; } catch (e1) {}
    try { document.body.scrollTop = y; } catch (e2) {}
    scrollerList().forEach((el) => {
      try { el.scrollTop = (el === document.documentElement || el === document.body) ? y : m; } catch (e3) {}
    });
    const main = document.querySelector("section.main") || document.querySelector('[data-testid="stMain"]');
    if (main) {
      try { main.scrollTop = m || y; } catch (e4) {}
    }
  }
  function freeze(ms) {
    window.__wlScrollSnap = readSnap();
    window.__wlMainScrollPinUntil = Date.now() + (ms || 1200);
  }
  function applyPin() {
    const until = window.__wlMainScrollPinUntil || 0;
    if (Date.now() > until) return;
    writeSnap(window.__wlScrollSnap);
  }
  function isDateChrome(el) {
    let n = el;
    for (let i = 0; i < 12 && n && n !== document; i++) {
      const cls = String((n.className && n.className.baseVal) || n.className || "");
      if (
        cls.indexOf("st-key-wl_date") >= 0 ||
        cls.indexOf("st-key-wl_cal") >= 0 ||
        cls.indexOf("st-key-wl_day_") >= 0 ||
        cls.indexOf("st-key-wl_today") >= 0 ||
        cls.indexOf("st-key-wl_prev_month") >= 0 ||
        cls.indexOf("st-key-wl_next_month") >= 0 ||
        cls.indexOf("st-key-wl_save_btn_") >= 0 ||
        cls.indexOf("st-key-wl_del_") >= 0
      ) return true;
      const root = n.getRootNode && n.getRootNode();
      n = (root && root !== document && root.host) ? root.host : n.parentElement;
    }
    return false;
  }

  if (typeof window !== "undefined") {
    window.__wlApplyScrollPin = applyPin;
    if (!window.__wlScrollIntoViewPatched) {
      window.__wlScrollIntoViewPatched = true;
      const orig = Element.prototype.scrollIntoView;
      Element.prototype.scrollIntoView = function () {
        if (Date.now() <= (window.__wlMainScrollPinUntil || 0)) return;
        return orig.apply(this, arguments);
      };
    }
    if (!window.__wlScrollLockListen) {
      window.__wlScrollLockListen = true;
      document.addEventListener("pointerdown", (e) => {
        if (isDateChrome(e.target)) freeze(500);
      }, true);
      document.addEventListener("scroll", () => {
        if (Date.now() <= (window.__wlMainScrollPinUntil || 0)) return;
        window.__wlScrollSnap = readSnap();
      }, true);
    }
    const prevIso = window.__wlScrollLockIso;
    const prevPin = Number(window.__wlScrollLockPin || 0);
    window.__wlScrollLockIso = iso;
    window.__wlScrollLockPin = pin;
    if ((prevIso && prevIso !== iso) || (pin && pin !== prevPin)) {
      window.__wlMainScrollPinUntil = Date.now() + 500;
    }
    if (window.__wlScrollSnap && Date.now() <= (window.__wlMainScrollPinUntil || 0)) {
      applyPin();
      requestAnimationFrame(applyPin);
      setTimeout(applyPin, 80);
    }
  }
}
"""

_WL_SCROLL_LOCK = st.components.v2.component(
    "worklog_scroll_lock_v3",
    js=_WL_SCROLL_LOCK_JS,
)

# 자주 쓰는 순 (앞쪽 = 우선 표시)

# 왼쪽 미리보기 — iframe(components.html) 재부착 없이 글자만 갱신
_WL_PREVIEW_HOST_HTML = """<div class="wl-live-preview"></div>"""
_WL_PREVIEW_HOST_CSS = """
.wl-live-preview {
  width: 100%;
  max-height: 820px;
  overflow: auto;
  background: #fff;
  box-sizing: border-box;
}
.wl-live-preview.excel {
  max-height: 1100px;
  border: 1px solid #94A3B8;
}
.wl-preview-page { display: block; }
"""
_WL_PREVIEW_HOST_JS = r"""
const __wlPrev = new WeakMap();

function applyPatches(root, patches) {
  if (!root || !patches || typeof patches !== "object") return;
  for (const key of Object.keys(patches)) {
    const el = root.querySelector('[data-wl="' + key + '"]');
    if (el) el.innerHTML = String(patches[key] ?? "");
  }
}

function scrollToPreviewPage(host, page) {
  if (!host || page < 1) return;
  const marked = host.querySelector('.wl-preview-page[data-wl-page="' + page + '"]');
  const tables = host.querySelectorAll(".wl-sheet");
  const el = marked || tables[page - 1] || (page > 1 ? host.querySelector('[data-wl^="p' + page + '-"]') : null);
  if (!el) return;
  const top = el.getBoundingClientRect().top - host.getBoundingClientRect().top + host.scrollTop;
  try {
    host.scrollTo({ top: Math.max(0, top - 6), behavior: "smooth" });
  } catch (e0) {
    host.scrollTop = Math.max(0, top - 6);
  }
}

export default function (component) {
  const { data, parentElement } = component;
  const host = parentElement.querySelector(".wl-live-preview");
  if (!host) return;

  const mode = String((data && data.mode) || "summary");
  const html = String((data && data.html) || "");
  const rev = String((data && data.rev) || "");
  const height = Number((data && data.height) || 0);
  const patches = (data && data.patches) || null;
  const page = Math.max(1, Number((data && data.page) || 1));

  host.className = "wl-live-preview" + (mode === "excel" ? " excel" : "");
  if (height > 0) host.style.maxHeight = height + "px";

  let inst = __wlPrev.get(host);
  if (!inst) {
    inst = { rev: "", mode: "", html: "", page: 0 };
    __wlPrev.set(host, inst);
  }
  if (html) inst.html = html;

  if (inst.rev !== rev || inst.mode !== mode || !host.childElementCount) {
    if (inst.html) {
      inst.rev = rev;
      inst.mode = mode;
      host.innerHTML = inst.html;
      inst.page = 0;
    }
  }
  if (patches) applyPatches(host, patches);
  if (mode === "excel" && inst.page !== page) {
    inst.page = page;
    requestAnimationFrame(() => scrollToPreviewPage(host, page));
  }
}
"""

_WL_PREVIEW_HOST = st.components.v2.component(
    "worklog_live_preview_v2",
    html=_WL_PREVIEW_HOST_HTML,
    css=_WL_PREVIEW_HOST_CSS,
    js=_WL_PREVIEW_HOST_JS,
)

_WL_PRINT_LAUNCH_JS = r"""
export default function (component) {
  const { data, parentElement } = component;
  const html = (data && data.html) || "";
  const n = String((data && data.n) || "");
  if (!html || !n) return;
  try {
    if (window.__wlPrintN === n) return;
    window.__wlPrintN = n;
  } catch (e0) {}
  const hostDoc = (parentElement && parentElement.ownerDocument)
    || document;
  let frame = hostDoc.getElementById("wl-print-frame");
  if (!frame) {
    frame = hostDoc.createElement("iframe");
    frame.id = "wl-print-frame";
    frame.setAttribute("aria-hidden", "true");
    frame.setAttribute("title", "일일업무일지 인쇄");
    frame.style.cssText = "position:fixed;left:-10000px;top:0;width:794px;height:1123px;border:0;visibility:hidden;pointer-events:none;";
    hostDoc.body.appendChild(frame);
  }
  try {
    frame.srcdoc = html;
  } catch (e2) {
    try {
      const doc = frame.contentDocument || (frame.contentWindow && frame.contentWindow.document);
      if (!doc) return;
      doc.open();
      doc.write(html);
      doc.close();
    } catch (e3) {}
  }
}
"""

_WL_PRINT_LAUNCH = st.components.v2.component(
    "worklog_print_launch_v2",
    js=_WL_PRINT_LAUNCH_JS,
)


_WL_UI_SLOT = "ui"


def _ui_iso(iso: str | None = None) -> str:
    """입력칸 CCv2 키 — 날짜가 바뀌어도 같은 위젯을 유지한다. 칸 값은 날짜 전환 때 그날 저장본으로 교체한다."""
    return _WL_UI_SLOT


def _entry_lines_inst_key(iso: str, entry_i: int) -> str: return f"wl_lines_inst_{_ui_iso(iso)}_{entry_i}"
def _entry_clients_inst_key(iso: str, entry_i: int) -> str: return f"wl_clients_inst_{_ui_iso(iso)}_{entry_i}"
def _entry_remarks_inst_key(iso: str, entry_i: int) -> str: return f"wl_remarks_inst_{_ui_iso(iso)}_{entry_i}"

def _entry_lines_comp_key(iso: str, entry_i: int) -> str:
    # 인스턴스 번호로 CCv2 위젯을 재마운트 — 동일 키에 남은 구 result.lines 부활 차단
    g = int(st.session_state.get(_entry_lines_inst_key(iso, entry_i), 0) or 0)
    return f"wl_lines_comp_{_ui_iso(iso)}_{entry_i}_i{g}"

def _entry_lines_rev_key(iso: str, entry_i: int) -> str: return f"wl_ent_rev_{_ui_iso(iso)}_{entry_i}"
def _entry_lines_live_key(iso: str, entry_i: int) -> str: return f"wl_lines_live_{_ui_iso(iso)}_{entry_i}"

def _entry_clients_comp_key(iso: str, entry_i: int) -> str:
    g = int(st.session_state.get(_entry_clients_inst_key(iso, entry_i), 0) or 0)
    return f"wl_clients_comp_{_ui_iso(iso)}_{entry_i}_i{g}"

def _entry_clients_rev_key(iso: str, entry_i: int) -> str: return f"wl_clients_rev_{_ui_iso(iso)}_{entry_i}"
def _entry_clients_live_key(iso: str, entry_i: int) -> str: return f"wl_clients_live_{_ui_iso(iso)}_{entry_i}"

def _entry_remarks_comp_key(iso: str, entry_i: int) -> str:
    g = int(st.session_state.get(_entry_remarks_inst_key(iso, entry_i), 0) or 0)
    return f"wl_remarks_comp_{_ui_iso(iso)}_{entry_i}_i{g}"

def _entry_remarks_rev_key(iso: str, entry_i: int) -> str: return f"wl_remarks_rev_{_ui_iso(iso)}_{entry_i}"
def _entry_remarks_live_key(iso: str, entry_i: int) -> str: return f"wl_remarks_live_{_ui_iso(iso)}_{entry_i}"

def _bump_entry_lines_comp_inst(iso: str, entry_i: int) -> None:
    st.session_state.pop(_entry_lines_comp_key(iso, entry_i), None)
    st.session_state.pop(f"wl_lines_comp_{iso}_{entry_i}", None)  # legacy key
    st.session_state.pop(f"wl_lines_user_edit_{iso}_{entry_i}", None)
    k = _entry_lines_inst_key(iso, entry_i)
    st.session_state[k] = int(st.session_state.get(k, 0) or 0) + 1

def _bump_entry_clients_comp_inst(iso: str, entry_i: int) -> None:
    st.session_state.pop(_entry_clients_comp_key(iso, entry_i), None)
    st.session_state.pop(f"wl_clients_comp_{iso}_{entry_i}", None)  # legacy key
    st.session_state.pop(f"wl_clients_user_edit_{iso}_{entry_i}", None)
    k = _entry_clients_inst_key(iso, entry_i)
    st.session_state[k] = int(st.session_state.get(k, 0) or 0) + 1

def _bump_entry_remarks_comp_inst(iso: str, entry_i: int) -> None:
    st.session_state.pop(_entry_remarks_comp_key(iso, entry_i), None)
    st.session_state.pop(f"wl_remarks_comp_{iso}_{entry_i}", None)
    st.session_state.pop(f"wl_remarks_user_edit_{iso}_{entry_i}", None)
    k = _entry_remarks_inst_key(iso, entry_i)
    st.session_state[k] = int(st.session_state.get(k, 0) or 0) + 1

def _scrub_dummy_label(val: str) -> str:
    s = (val or "").strip()
    if re.fullmatch(r"거래처\d+", s) or re.fullmatch(r"내용\d+", s) or re.fullmatch(r"비고\d+", s): return ""
    return val or ""

# 원본 14pt 명조 칸 폭. 한글=1em, 띄어쓰기≈0.42em, ASCII≈0.55em.
# 단위는 한글 1자=2. 칸 폭은 원본.xlsx G~X·C~F·Y~AB 그대로.
_WL_SPACE_UNITS = 0.58
_WL_ASCII_UNITS = 1.1
_WL_CONTENT_HANGUL = 37
_WL_CLIENT_HANGUL = 8
_WL_REMARK_HANGUL = 8
_WL_ORIG_CONTENT_PX = 666
_WL_ORIG_SIDE_PX = 148
# 입력칸 테두리·패딩 보정 + 거래처·비고를 입력에서 더 넓게.
_WL_INPUT_SIDE_CHROME_PX = 16
_WL_INPUT_SIDE_WIDEN_PX = 44

def _char_units(ch: str) -> float:
    if not ch: return 0.0
    if ch in " \t\u00a0": return _WL_SPACE_UNITS
    o = ord(ch)
    if (0xAC00 <= o <= 0xD7A3 or 0x1100 <= o <= 0x11FF or 0x3130 <= o <= 0x318F or 0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF): return 2.0
    return _WL_ASCII_UNITS

def _lstrip_line_ws(s: str) -> str:
    return (s or "").lstrip(" \t\u00a0\u3000")


def _display_units(s: str) -> float:
    """미리보기·원본 칸과 같이 앞 공백도 폭에 넣는다. 칸이 차면 바로 다음 줄."""
    return float(sum(_char_units(ch) for ch in (s or "")))

_WL_BODY_FONT_NAME = "바탕체"
_WL_BODY_FONT_PT = 14.0  # 원본.xlsx 본문 글자 크기와 동일
_WL_LOOK_PT_DELTA = 0.0  # 입력 표시는 _WL_INPUT_LOOK_SCALE. 원본·인쇄는 14pt.

def _input_look_scale() -> float:
    return float(_WL_INPUT_LOOK_SCALE)

def _look_font_pt() -> float:
    return max(7.0, float(_WL_BODY_FONT_PT) * float(_input_look_scale()) + float(_WL_LOOK_PT_DELTA))

def _input_cell_px(kind: str) -> int:
    """입력칸 표시 폭. 글자 배율에 맞추고 거래처·비고는 더 넓힌다."""
    orig = int(_orig_cell_px(kind))
    scaled = max(40, int(round(orig * (_look_font_pt() / float(_WL_BODY_FONT_PT)))))
    if kind in ("client", "remark"):
        return scaled + int(_WL_INPUT_SIDE_CHROME_PX) + int(_WL_INPUT_SIDE_WIDEN_PX)
    return scaled

def _set_body_font(cell) -> None:
    try: cell.font = cell.font.copy(name=_WL_BODY_FONT_NAME, size=float(_WL_BODY_FONT_PT))
    except Exception: pass

# =====================================================================
# 💡 [핵심] 글자 넘침 현상 원천 차단 (엄격한 max_units 설정)
# =====================================================================
# 14pt 바탕체 기준 내용칸 한 줄 한도 (한글 1자=2단위). 자동 다음칸 이동 임계값.
@lru_cache(maxsize=1)
def _content_line_units() -> int: return int(_WL_CONTENT_HANGUL) * 2

@lru_cache(maxsize=1)
def _client_line_units() -> int: return int(_WL_CLIENT_HANGUL) * 2

# 비고 Y~AB — 거래처와 같이 한글 8자.
@lru_cache(maxsize=1)
def _remark_line_units() -> int: return int(_WL_REMARK_HANGUL) * 2

def _hangul_line_limit(max_u: int) -> int:
    return max(1, int(max_u) // 2)


@lru_cache(maxsize=8)
def _orig_cell_px(kind: str) -> int:
    """엑셀 미리보기 원본 병합칸 폭(px). 입력 줄바꿈과 미리보기를 같게 맞춘다."""
    if kind == "content":
        c0, c1, fallback = WL_CONTENT_COL_START, WL_CONTENT_COL_END, _WL_ORIG_CONTENT_PX
    elif kind == "remark":
        c0, c1, fallback = WL_REMARK_COL_START, WL_REMARK_COL_END, _WL_ORIG_SIDE_PX
    else:
        c0, c1, fallback = WL_CLIENT_COL_START, WL_CLIENT_COL_END, _WL_ORIG_SIDE_PX
    try:
        path = WORKLOG_TEMPLATE
        if load_workbook is not None and path and os.path.exists(path):
            wb = load_workbook(path, data_only=False)
            ws = wb.active
            px = sum(_excel_width_to_px(_excel_col_width(ws, c)) for c in range(c0, c1 + 1))
            wb.close()
            if px >= 40:
                return int(px)
    except Exception:
        pass
    return int(fallback)


def _fit_preview_col_widths(col_widths: list[int], layout_scale: float = 1.0) -> list[int]:
    """원본.xlsx 열 폭을 그대로 둔다."""
    return list(col_widths)


def _pad_sheet_lines(lines: list[str] | None, n: int = WL_SHEET_N) -> list[str]:
    """본문 32칸. 익일업무·특이사항은 여기 넣지 않는다."""
    out = [str(x or "") for x in (lines or [])][:n]
    while len(out) < n:
        out.append("")
    return out


def _sheet_lines_from_cells(cells: dict) -> tuple[list[str], list[str], list[str]]:
    clients, contents, remarks = [], [], []
    for r in WL_CONTENT_ROWS:
        g = str(cells.get(f"G{r}", "") or "")
        if g in (_WL_SOFT_BLANK, "\u00a0"):
            g = ""
        clients.append(str(cells.get(f"C{r}", "") or ""))
        contents.append(g)
        remarks.append(str(cells.get(f"Y{r}", "") or ""))
    return _pad_sheet_lines(clients), _pad_sheet_lines(contents), _pad_sheet_lines(remarks)


def _pack_sheet_to_cells(
    d: date,
    clients: list[str],
    contents: list[str],
    remarks: list[str],
    next_day: list[str] | None = None,
    notes: list[str] | None = None,
) -> dict:
    """본문 32칸만 행 대 행으로 넣고, 익일·특이는 기존 D열 칸에 둔다."""
    cells = _empty_cells(d)
    clients, contents, remarks = (
        _pad_sheet_lines(clients),
        _pad_sheet_lines(contents),
        _pad_sheet_lines(remarks),
    )
    for i, r in enumerate(WL_CONTENT_ROWS):
        cells[f"C{r}"] = clients[i]
        cells[f"G{r}"] = contents[i]
        cells[f"Y{r}"] = remarks[i]
    max_u = _content_line_units()
    for t_list, t_rows in ((next_day, WL_NEXT_ROWS), (notes, WL_NOTE_ROWS)):
        chunks = _panel_lines_to_cells(t_list, max_u)
        for i, r in enumerate(t_rows):
            cells[f"D{r}"] = chunks[i] if i < len(chunks) else ""
    return cells


def _sheet_entry_from_lines(
    clients: list[str] | None,
    contents: list[str] | None,
    remarks: list[str] | None,
) -> dict:
    clients = _pad_sheet_lines(clients)
    contents = _pad_sheet_lines(contents)
    remarks = _pad_sheet_lines(remarks)
    return {
        "client": "\n".join(x for x in clients if str(x).strip()),
        "client_lines": clients,
        "content": "\n".join(x for x in contents if str(x).strip()),
        "lines": contents,
        "remarks": "\n".join(x for x in remarks if str(x).strip()),
        "remark_lines": remarks,
        "blank_after": 0,
    }


def _sheet_entry_from_cells(cells: dict) -> dict:
    return _sheet_entry_from_lines(*_sheet_lines_from_cells(cells or {}))


def _empty_sheet_entry() -> dict:
    return _sheet_entry_from_lines([], [], [])


def _extra_page_sheet_name(page_n: int) -> str:
    """일자 파일 추가 시트. 2페이지 → p2. 숫자만이면 달력이 일로 착각한다."""
    return f"p{max(2, int(page_n))}"


def _archive_extra_sheet_name(d: date, page_n: int) -> str:
    """월별 파일 추가 시트. 16일 2페이지 → 16p2."""
    return f"{d.day}p{max(2, int(page_n))}"


def _extra_page_n_from_sheet_name(name: str) -> int | None:
    s = str(name or "").strip()
    m = re.fullmatch(r"p(\d+)", s)
    if m:
        n = int(m.group(1))
        return n if n >= 2 else None
    m = re.fullmatch(r"(\d{1,2})p(\d+)", s)
    if m:
        n = int(m.group(2))
        return n if n >= 2 else None
    return None


def _is_day_extra_sheet_name(name: str) -> bool:
    return bool(re.fullmatch(r"p\d+", str(name or "").strip()))


def _is_archive_extra_sheet_name(name: str, d: date | None = None) -> bool:
    s = str(name or "").strip()
    if d is not None:
        return bool(re.fullmatch(rf"{d.day}p\d+", s))
    return bool(re.fullmatch(r"\d{1,2}p\d+", s))


def _page_count_key(iso: str) -> str:
    return f"wl_page_count_{iso}"


def _page_idx_key(iso: str) -> str:
    return f"wl_page_idx_{iso}"


def _page_count_for(iso: str) -> int:
    n = int(st.session_state.get(_page_count_key(iso), 1) or 1)
    return max(1, min(WL_MAX_PAGES, n))


def _page_idx_for(iso: str) -> int:
    n = _page_count_for(iso)
    i = int(st.session_state.get(_page_idx_key(iso), 0) or 0)
    return max(0, min(n - 1, i))


def _page_snap_key(iso: str, entry_i: int) -> str:
    return f"wl_page_snap_{iso}_{int(entry_i)}"


def _page_lines_have_text(parts: tuple[list[str], list[str], list[str]] | None) -> bool:
    if not parts:
        return False
    return any(str(x).strip() for col in parts for x in (col or []))


def _live_sheet_lines_at(iso: str, entry_i: int) -> tuple[list[str], list[str], list[str]]:
    return (
        _pad_sheet_lines(_clients_from_widgets(iso, entry_i, keep_trailing_empty=True)),
        _pad_sheet_lines(_lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=True)),
        _pad_sheet_lines(_remarks_from_widgets(iso, entry_i, keep_trailing_empty=True)),
    )


def _snapshot_worklog_page(iso: str, entry_i: int | None = None) -> None:
    """보이는 페이지를 스냅샷. 페이지 전환 때 CCv2 언마운트가 빈 값으로 덮는 것을 막는다."""
    i = _page_idx_for(iso) if entry_i is None else int(entry_i)
    live = _live_sheet_lines_at(iso, i)
    prev = st.session_state.get(_page_snap_key(iso, i))
    if (not _page_lines_have_text(live)) and isinstance(prev, (list, tuple)) and len(prev) == 3 and _page_lines_have_text(tuple(prev)):
        return
    st.session_state[_page_snap_key(iso, i)] = live


def _sheet_lines_from_widgets_at(iso: str, entry_i: int) -> tuple[list[str], list[str], list[str]]:
    live = _live_sheet_lines_at(iso, entry_i)
    snap = st.session_state.get(_page_snap_key(iso, entry_i))
    snap_p = None
    if isinstance(snap, (list, tuple)) and len(snap) == 3:
        snap_p = (
            _pad_sheet_lines(snap[0]),
            _pad_sheet_lines(snap[1]),
            _pad_sheet_lines(snap[2]),
        )
    if entry_i != _page_idx_for(iso) and snap_p is not None:
        return snap_p
    if (not _page_lines_have_text(live)) and snap_p is not None and _page_lines_have_text(snap_p):
        return snap_p
    return live


def _detach_extra_pages(cells: dict | None) -> tuple[dict, list[dict]]:
    src = dict(cells or {})
    extras = src.pop("_extra_pages", None)
    if not isinstance(extras, list):
        extras = []
    clean: list[dict] = []
    for extra in extras:
        if isinstance(extra, dict):
            item = {k: v for k, v in extra.items() if k != "_extra_pages"}
            clean.append(item)
    return src, clean


def _attach_extra_pages(cells: dict, extras: list[dict] | None) -> dict:
    out = dict(cells or {})
    if extras:
        out["_extra_pages"] = [dict(x) for x in extras if isinstance(x, dict)]
    else:
        out.pop("_extra_pages", None)
    return out


def _sheet_page_has_body_text(cells: dict | None) -> bool:
    src, _extras = _detach_extra_pages(cells)
    return any(
        str(src.get(f"G{r}", "") or "").strip()
        or str(src.get(f"C{r}", "") or "").strip()
        or str(src.get(f"Y{r}", "") or "").strip()
        for r in WL_CONTENT_ROWS
    )


def _collapse_empty_worklog_pages(cells: dict) -> dict:
    """앞쪽 빈 페이지를 접고, 내용 없는 추가 페이지는 저장하지 않는다."""
    body, extras = _detach_extra_pages(cells)
    extras = [dict(e) for e in extras if isinstance(e, dict)]
    for extra in extras:
        extra.pop("_extra_pages", None)
    while extras and not _sheet_page_has_body_text(body):
        nxt = extras.pop(0)
        for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
            k = f"D{r}"
            if not str(nxt.get(k, "") or "").strip() and str(body.get(k, "") or "").strip():
                nxt[k] = body[k]
        if body.get("date"):
            nxt["date"] = body.get("date")
        body = nxt
    extras = [e for e in extras if _sheet_page_has_body_text(e)]
    return _attach_extra_pages(body, extras)


def _html_page_break() -> str:
    return (
        '<div class="wl-page-break" style="page-break-before:always;break-before:page;'
        'height:18px;min-height:18px;"></div>'
    )


def _sheet_row_usage(
    clients: list[str] | None,
    contents: list[str] | None,
    remarks: list[str] | None,
) -> dict:
    """본문 32칸만 센다. 맨 끝 빈 칸은 사용으로 치지 않는다. 익일·특이는 제외."""
    clients, contents, remarks = (
        _pad_sheet_lines(clients),
        _pad_sheet_lines(contents),
        _pad_sheet_lines(remarks),
    )
    used = 0
    for i in range(WL_SHEET_N):
        if any(str(x).strip() for x in (clients[i], contents[i], remarks[i])):
            used = i + 1
    total = WL_SHEET_N
    return {
        "total": total,
        "used": used,
        "remaining": max(0, total - used),
        "per_entry": [used],
        "last_row": WL_CONTENT_ROWS[-1] if WL_CONTENT_ROWS else 39,
        "next_row": WL_CONTENT_ROWS[used] if used < total else None,
        "overflow": False,
    }

def _wl_col_limit_label(title: str, max_u: int) -> str:
    n = _hangul_line_limit(max_u)
    return (
        f"<div style='font-size:11px;font-weight:700;color:#334155;margin:0 0 4px;line-height:1.35;white-space:nowrap;overflow:hidden;min-height:1.4em;'>"
        f"{title} <span style='font-weight:500;color:#94A3B8;'>원본 한글 {n}자</span></div>"
    )

def _fit_by_units(s: str, max_units: int | None = None) -> tuple[str, str]:
    if max_units is None: max_units = _content_line_units()
    if not s: return "", ""
    if _display_units(s) <= max_units: return s, ""
    acc = 0.0
    for i, ch in enumerate(s):
        cu = _char_units(ch)
        if acc + cu > max_units:
            head = s[:i] if i else s[:1]
            tail = s[i:] if i else s[1:]
            return head, _lstrip_line_ws(tail)
        acc += cu
    return s, ""

def _chunk_text(text: str, max_units: int | None = None) -> list[str]:
    if max_units is None: max_units = _content_line_units()
    s = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not s: return []
    out: list[str] = []
    for para in s.split("\n"):
        rest = para
        if rest == "" and not out: continue
        if rest == "": continue
        while rest:
            head, rest = _fit_by_units(rest, max_units)
            if not head and rest: head, rest = rest[:1], rest[1:]
            out.append(head)
            if not rest: break
    return out

def _spill_column(cells: dict, rows: list[int], col: str, max_u: int | None = None) -> dict:
    if max_u is None: max_u = _content_line_units()
    vals = [str(cells.get(f"{col}{r}", "") or "") for r in rows]
    for i in range(len(vals)):
        while _display_units(vals[i]) > max_u and i + 1 < len(vals):
            head, tail = _fit_by_units(vals[i], max_u)
            vals[i] = head
            vals[i + 1] = _lstrip_line_ws(tail) + vals[i + 1]
        if _display_units(vals[i]) > max_u:
            head, _tail = _fit_by_units(vals[i], max_u)
            vals[i] = head
    out = dict(cells)
    for r, v in zip(rows, vals): out[f"{col}{r}"] = v
    return out

def _spill_all_content(cells: dict) -> dict:
    body, extras = _detach_extra_pages(cells)
    body = _spill_column(body, WL_CONTENT_ROWS, "G")
    body = _spill_column(body, WL_CONTENT_ROWS, "Y", _remark_line_units())
    body = _spill_column(body, WL_NEXT_ROWS, "D")
    body = _spill_column(body, WL_NOTE_ROWS, "D")
    spilled_extras = [_spill_all_content(extra) for extra in extras]
    return _attach_extra_pages(body, spilled_extras)

# 💡 템플릿 준비: git의 uploaded_cache/worklog/template.xlsx 를 우선 사용
# (예전엔 ~/Desktop/업무일지.xlsx mtime이 더 新し면 덮어써서 로고·양식 반영이 깨짐)
def _ensure_dirs() -> None:
    os.makedirs(WORKLOG_DIR, exist_ok=True)
    # Desktop 파일은 "캐시 템플릿이 없을 때만" 보충. 있으면 절대 덮어쓰지 않음.
    if (not os.path.exists(WORKLOG_TEMPLATE)) and os.path.exists(WORKLOG_TEMPLATE_SRC):
        try:
            shutil.copy2(WORKLOG_TEMPLATE_SRC, WORKLOG_TEMPLATE)
        except Exception:
            pass

def _iter_google_drive_roots() -> list[str]:
    cloud = os.path.join(os.path.expanduser("~"), "Library", "CloudStorage")
    if not os.path.isdir(cloud): return []
    roots: list[str] = []
    try:
        for name in sorted(os.listdir(cloud)):
            if name.startswith("GoogleDrive"): roots.append(os.path.join(cloud, name))
    except OSError: return []
    return roots

def resolve_worklog_archive_root() -> str | None:
    """…/Desktop/업무/일지 또는 Drive 경로. 세션 캐시로 상단 필터 rerun 시 Drive 스캔 생략."""
    cached = st.session_state.get("_wl_archive_root_cache")
    if isinstance(cached, dict) and "path" in cached:
        return cached.get("path")
    candidates: list[str] = []
    home = os.path.expanduser("~")
    candidates.append(os.path.join(home, "Desktop", "업무", "일지"))
    for groot in _iter_google_drive_roots():
        for other_name in ("다른 컴퓨터", "Computers"):
            other = os.path.join(groot, other_name)
            if not os.path.isdir(other): continue
            try: pcs = sorted(os.listdir(other))
            except OSError: continue
            pcs.sort(key=lambda n: (0 if "(1)" in n else 1, n))
            for pc in pcs: candidates.append(os.path.join(other, pc, WORKLOG_ARCHIVE_REL))
    found: str | None = None
    existing = [p for p in candidates if os.path.isdir(p)]
    if existing:
        found = existing[0]
    else:
        for p in candidates:
            parent = os.path.dirname(p)
            grand = os.path.dirname(parent)
            if os.path.isdir(grand):
                try:
                    os.makedirs(p, exist_ok=True)
                    found = p
                    break
                except OSError:
                    continue
    st.session_state["_wl_archive_root_cache"] = {"path": found}
    return found


def worklog_archive_path(d: date) -> str | None:
    """호환용: 월별 통합 파일 경로 (…/일지/2026/8월.xlsx)."""
    return worklog_archive_month_path(d)


def worklog_archive_year_dir(d: date, *, create: bool = True) -> str | None:
    """연도 폴더: …/일지/2026. create=True 이면 없을 때 자동 생성."""
    root = resolve_worklog_archive_root()
    if not root:
        return None
    year_dir = os.path.join(root, str(d.year))
    if create:
        try:
            os.makedirs(year_dir, exist_ok=True)
        except OSError:
            return None
    elif not os.path.isdir(year_dir):
        return None
    return year_dir


def worklog_archive_month_dir(d: date) -> str | None:
    """구버전 월 하위폴더(…/2026/8월). 레거시 정리용."""
    year_dir = worklog_archive_year_dir(d, create=False)
    if not year_dir:
        return None
    return os.path.join(year_dir, f"{d.month}월")


def worklog_archive_month_path(d: date, *, create_year: bool = True) -> str | None:
    """달이 바뀌면 9월.xlsx 신규. 경로: …/일지/2026/9월.xlsx (연도 폴더 직하)."""
    year_dir = worklog_archive_year_dir(d, create=create_year)
    if not year_dir:
        return None
    return os.path.join(year_dir, f"{d.month}월.xlsx")


def _cleanup_legacy_day_archive_files(d: date, month_path: str | None) -> None:
    """예전 일자별 xlsx·구 월폴더 안 일자파일을 정리."""
    year_dir = worklog_archive_year_dir(d, create=False)
    candidates: list[str] = []
    if year_dir:
        candidates.append(os.path.join(year_dir, f"{d.isoformat()}.xlsx"))
        # 구경로: …/2026/8월/YYYY-MM-DD.xlsx , …/2026/8월/8월.xlsx
        old_month_dir = os.path.join(year_dir, f"{d.month}월")
        candidates.append(os.path.join(old_month_dir, f"{d.isoformat()}.xlsx"))
        candidates.append(os.path.join(old_month_dir, f"{d.month}월.xlsx"))
    month_abs = os.path.abspath(month_path) if month_path else ""
    for path in candidates:
        try:
            if month_abs and os.path.abspath(path) == month_abs:
                continue
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            pass
    # 구 월 폴더가 비면 제거
    if year_dir:
        old_month_dir = os.path.join(year_dir, f"{d.month}월")
        try:
            if os.path.isdir(old_month_dir) and not os.listdir(old_month_dir):
                os.rmdir(old_month_dir)
        except OSError:
            pass


def worklog_archive_sheet_title(d: date) -> str:
    """월별 파일 안 워크시트명 = 일 (예: 27)."""
    return str(d.day)


def _worklog_archive_legacy_sheet_title(d: date) -> str:
    """구버전 워크시트명 (YYYY-MM-DD). 존재·삭제 조회용."""
    return d.isoformat()


def _worklog_archive_sheet_titles_for_lookup(d: date) -> tuple[str, ...]:
    """현재(일) + 레거시(YYYY-MM-DD) 워크시트명."""
    cur = worklog_archive_sheet_title(d)
    leg = _worklog_archive_legacy_sheet_title(d)
    return (cur,) if cur == leg else (cur, leg)


def _resolve_archive_sheet_name(sheetnames: list[str] | tuple[str, ...], d: date) -> str | None:
    for title in _worklog_archive_sheet_titles_for_lookup(d):
        if title in sheetnames:
            return title
    return None


def _archive_sheet_sort_key(name: str) -> tuple:
    if name.isdigit():
        return (0, int(name))
    try:
        return (1, date.fromisoformat(name))
    except ValueError:
        return (2, name)


def _invalidate_worklog_presence_cache(d: date | None = None) -> None:
    """날짜 존재 캐시 무효화 (저장·삭제 후)."""
    if d is not None:
        iso = d.isoformat()
        for k in (
            f"wl_arch_exists_{iso}",
            f"wl_presence_{iso}",
            f"wl_presence_{iso}_fast",
            f"wl_presence_{iso}_all",
        ):
            st.session_state.pop(k, None)
        return
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and (k.startswith("wl_arch_exists_") or k.startswith("wl_presence_")):
            st.session_state.pop(k, None)


def worklog_date_exists_in_archive(d: date) -> bool:
    """월별 xlsx에 해당 날짜 시트가 있는지 (구 일자파일 포함)."""
    iso = d.isoformat()
    ck = f"wl_arch_exists_{iso}"
    cached = st.session_state.get(ck)
    if isinstance(cached, bool):
        return cached
    titles = _worklog_archive_sheet_titles_for_lookup(d)
    found = False
    month_path = worklog_archive_month_path(d, create_year=False)
    if month_path and os.path.exists(month_path) and load_workbook is not None:
        try:
            wb = load_workbook(month_path, read_only=True)
            try:
                if _resolve_archive_sheet_name(wb.sheetnames, d):
                    found = True
            finally:
                wb.close()
        except Exception:
            pass
    if not found:
        year_dir = worklog_archive_year_dir(d, create=False)
        if year_dir:
            for p in (
                os.path.join(year_dir, f"{d.isoformat()}.xlsx"),
                os.path.join(year_dir, f"{d.month}월", f"{d.isoformat()}.xlsx"),
                os.path.join(year_dir, f"{d.month}월", f"{d.month}월.xlsx"),
            ):
                if not os.path.exists(p):
                    continue
                if p.endswith(f"{d.month}월.xlsx") and load_workbook is not None:
                    try:
                        wb = load_workbook(p, read_only=True)
                        try:
                            if _resolve_archive_sheet_name(wb.sheetnames, d):
                                found = True
                                break
                        finally:
                            wb.close()
                    except Exception:
                        pass
                elif os.path.basename(p) == f"{d.isoformat()}.xlsx":
                    found = True
                    break
    st.session_state[ck] = found
    return found


def worklog_date_exists_on_drive(d: date) -> bool:
    try:
        from drive_autoload import resolve_drive_worklog_archive_dir, resolve_drive_worklog_dir

        drv = resolve_drive_worklog_dir()
        if drv and os.path.isfile(os.path.join(drv, f"{d.isoformat()}.xlsx")):
            return True
        # Drive에 올린 월별 통합 파일 시트도 존재로 간주
        arch_dir = resolve_drive_worklog_archive_dir(d.year)
        if not arch_dir or load_workbook is None:
            return False
        month_p = os.path.join(arch_dir, f"{d.month}월.xlsx")
        if not os.path.isfile(month_p):
            return False
        wb = load_workbook(month_p, read_only=True)
        try:
            return _resolve_archive_sheet_name(wb.sheetnames, d) is not None
        finally:
            wb.close()
    except Exception:
        return False


def worklog_date_exists_on_cloud(d: date) -> bool:
    try:
        from worklog_remote_sync import worklog_date_exists_on_cloud as _cloud_exists

        return bool(_cloud_exists(d, WORKLOG_DIR))
    except Exception:
        return False


def detect_worklog_date_presence(d: date, *, include_remote: bool = True) -> dict:
    """로컬 캐시·월별보관·(선택) Drive·Cloud Gist 에 해당 날짜 일지 존재 여부."""
    iso = d.isoformat()
    cache_k = f"wl_presence_{iso}_{'all' if include_remote else 'fast'}"
    cached = st.session_state.get(cache_k)
    if isinstance(cached, dict):
        return cached
    local = os.path.isfile(worklog_path(d))
    archive = False
    drive = False
    cloud = False
    if not local:
        archive = worklog_date_exists_in_archive(d)
        if include_remote:
            if not archive:
                drive = worklog_date_exists_on_drive(d)
            if not archive and not drive:
                cloud = worklog_date_exists_on_cloud(d)
    locs: list[str] = []
    if local:
        locs.append("로컬 캐시")
    if archive:
        root = resolve_worklog_archive_root()
        if root:
            locs.append(f"일지/{d.year}/{d.month}월.xlsx")
        else:
            locs.append(f"월별파일({d.month}월.xlsx)")
    if drive:
        locs.append("Drive worklog")
    if cloud:
        locs.append("Cloud Gist")
    out = {
        "local": local,
        "archive": archive,
        "drive": drive,
        "cloud": cloud,
        "any": bool(local or archive or drive or cloud),
        "locations": locs,
    }
    st.session_state[cache_k] = out
    return out


def check_worklog_save_allowed(d: date, *, had_local_at_open: bool) -> tuple[bool, str]:
    """날짜 중복 시 후입력 저장 차단.

    달력에 • 가 없는 날은 무조건 허용한다 (빈 시트·남은 일자파일·캐시 불일치).
    • 가 있어도 칸이 비어 있으면 허용. 실제 내용이 있는 • 날만 차단.
    had_local_at_open=True(덮어쓰기/이미 연 날)면 로컬 파일이 없어도 허용.
    Cloud Gist·Drive에만 있으면 맥 로컬 저장은 허용 (저장 시 로컬+아카이브 반영).
    """
    if had_local_at_open:
        return True, ""
    iso = d.isoformat()
    if iso not in _saved_dates_for_calendar():
        return True, ""
    local_has = False
    if os.path.isfile(worklog_path(d)):
        try:
            local_has = _worklog_cells_have_draft(read_worklog_cells(d))
        except Exception:
            local_has = True
    archive_has = worklog_archive_has_saved_content(d) if not local_has else False
    if not local_has and not archive_has:
        return True, ""
    locs: list[str] = []
    if local_has:
        locs.append("로컬 캐시")
    if archive_has:
        root = resolve_worklog_archive_root()
        if root:
            locs.append(f"일지/{d.year}/{d.month}월.xlsx")
        else:
            locs.append(f"월별파일({d.month}월.xlsx)")
    detail = ", ".join(locs) if locs else "저장소"
    return (
        False,
        f"{d.isoformat()} 일지가 이미 있습니다 ({detail}). "
        "수정하려면 달력에서 해당 날짜(•)를 선택하거나 삭제 후 다시 저장하세요.",
    )


def describe_worklog_archive_target(d: date) -> str:
    """UI용 저장 경로 설명 (Desktop/업무/일지/YYYY/N월.xlsx#일)."""
    root = resolve_worklog_archive_root()
    month_path = worklog_archive_month_path(d)
    if root and month_path:
        return f"{root}/{d.year}/{d.month}월.xlsx#{worklog_archive_sheet_title(d)}"
    if month_path:
        return f"{month_path}#{worklog_archive_sheet_title(d)}"
    home = os.path.expanduser("~")
    return f"{home}/Desktop/업무/일지/{d.year}/{d.month}월.xlsx#{worklog_archive_sheet_title(d)}"


def _copy_worksheet_cross_workbook(src_ws, dst_ws) -> None:
    """다른 통합문서 간 시트 복사 — openpyxl WorksheetCopy + 인쇄·화면 설정.

    merged_cells 객체 shallow copy / _style 일괄 복사는 Excel에서 병합·테두리가
    깨지므로, 셀 서식은 속성별 copy 후 merge_cells()로 병합을 다시 적용한다.
    """
    from copy import copy as _cpy

    for (row, col), source_cell in src_ws._cells.items():
        target_cell = dst_ws.cell(column=col, row=row)
        target_cell._value = source_cell._value
        target_cell.data_type = source_cell.data_type
        if source_cell.has_style:
            try:
                target_cell.font = _cpy(source_cell.font)
                target_cell.border = _cpy(source_cell.border)
                target_cell.fill = _cpy(source_cell.fill)
                target_cell.number_format = source_cell.number_format
                target_cell.protection = _cpy(source_cell.protection)
                target_cell.alignment = _cpy(source_cell.alignment)
            except Exception:
                pass
        if source_cell.hyperlink:
            target_cell._hyperlink = _cpy(source_cell.hyperlink)
        if source_cell.comment:
            target_cell.comment = _cpy(source_cell.comment)

    for attr in ("row_dimensions", "column_dimensions"):
        src_dims = getattr(src_ws, attr)
        dst_dims = getattr(dst_ws, attr)
        for key, dim in src_dims.items():
            dst_dims[key] = _cpy(dim)
            dst_dims[key].worksheet = dst_ws

    dst_ws.sheet_format = _cpy(src_ws.sheet_format)
    dst_ws.sheet_properties = _cpy(src_ws.sheet_properties)
    dst_ws.page_margins = _cpy(src_ws.page_margins)
    dst_ws.page_setup = _cpy(src_ws.page_setup)
    dst_ws.print_options = _cpy(src_ws.print_options)

    try:
        dst_ws.print_area = src_ws.print_area
    except Exception:
        pass
    try:
        dst_ws.print_title_rows = src_ws.print_title_rows
        dst_ws.print_title_cols = src_ws.print_title_cols
    except Exception:
        pass
    try:
        dst_ws.sheet_view = _cpy(src_ws.sheet_view)
    except Exception:
        pass
    try:
        if getattr(src_ws, "views", None):
            dst_ws.views = _cpy(src_ws.views)
    except Exception:
        pass

    # 병합은 셀·서식 복사 후 마지막에 적용 (merged_cells shallow copy 금지)
    try:
        for mr in list(src_ws.merged_cells.ranges):
            dst_ws.merge_cells(str(mr))
    except Exception:
        pass
    _ensure_worklog_body_merges(dst_ws)

    try:
        from openpyxl.drawing.image import Image as XLImage

        for img in list(getattr(src_ws, "_images", []) or []):
            try:
                data = img._data()
                bio = io.BytesIO(data)
                new_img = XLImage(bio)
                if getattr(img, "width", None):
                    new_img.width = img.width
                if getattr(img, "height", None):
                    new_img.height = img.height
                if getattr(img, "anchor", None) is not None:
                    new_img.anchor = img.anchor
                dst_ws.add_image(new_img)
            except Exception:
                continue
    except Exception:
        pass


def _clone_worksheet_to_workbook(src_ws, dst_wb, title: str):
    """날짜일지 시트를 월별 통합 파일로 복사 (원본 인쇄·열 너비·화면 배율 유지)."""
    if title in dst_wb.sheetnames:
        del dst_wb[title]
    dst = dst_wb.create_sheet(title)
    _copy_worksheet_cross_workbook(src_ws, dst)
    return dst


def _migrate_legacy_month_workbook(d: date, month_path: str) -> None:
    """구경로 …/2026/8월/8월.xlsx → …/2026/8월.xlsx 로 1회 이동."""
    if not month_path:
        return
    if os.path.exists(month_path):
        return
    year_dir = worklog_archive_year_dir(d)
    if not year_dir:
        return
    legacy_month = os.path.join(year_dir, f"{d.month}월", f"{d.month}월.xlsx")
    if not os.path.exists(legacy_month):
        return
    try:
        os.makedirs(os.path.dirname(month_path), exist_ok=True)
        shutil.move(legacy_month, month_path)
    except OSError:
        try:
            shutil.copy2(legacy_month, month_path)
            os.remove(legacy_month)
        except OSError:
            pass


def _rename_day_extra_sheets_to_archive(wb, d: date) -> None:
    """일자 파일 p2 → 월별 16p2."""
    for name in list(wb.sheetnames):
        n = _extra_page_n_from_sheet_name(name)
        if not n or not _is_day_extra_sheet_name(name):
            continue
        new_title = _archive_extra_sheet_name(d, n)
        if name == new_title:
            continue
        if new_title in wb.sheetnames:
            del wb[new_title]
        wb[name].title = new_title


def _sync_archive_extra_sheets(day_wb, month_wb, d: date) -> None:
    """월별 파일에 해당 날짜 추가 페이지 시트를 맞춘다."""
    wanted: set[str] = set()
    active_title = day_wb.active.title if day_wb.worksheets else ""
    for name in day_wb.sheetnames:
        if name == active_title:
            continue
        n = _extra_page_n_from_sheet_name(name)
        if not n:
            continue
        arch_name = _archive_extra_sheet_name(d, n)
        wanted.add(arch_name)
        _clone_worksheet_to_workbook(day_wb[name], month_wb, arch_name)
    for name in list(month_wb.sheetnames):
        if _is_archive_extra_sheet_name(name, d) and name not in wanted:
            del month_wb[name]


def upsert_worklog_archive_sheet(d: date, day_xlsx_path: str, *, allow_overwrite: bool = True) -> str | None:
    """일자 파일을 월별 xlsx의 날짜 시트로 반영. 달이 바뀌면 N월.xlsx 신규 생성."""
    if load_workbook is None:
        return None
    month_path = worklog_archive_month_path(d)
    if not month_path or not day_xlsx_path or not os.path.exists(day_xlsx_path):
        return None
    sheet_title = worklog_archive_sheet_title(d)
    legacy_title = _worklog_archive_legacy_sheet_title(d)
    if not allow_overwrite and worklog_date_exists_in_archive(d):
        return month_path if os.path.exists(month_path) else None
    _migrate_legacy_month_workbook(d, month_path)

    # 첫 월 파일: 일지 xlsx를 그대로 복사 → 인쇄 미리보기·열 너비 100% 유지
    if not os.path.exists(month_path):
        shutil.copy2(day_xlsx_path, month_path)
        month_wb = load_workbook(month_path)
        try:
            ws = month_wb.active
            if ws.title != sheet_title:
                ws.title = sheet_title
            _rename_day_extra_sheets_to_archive(month_wb, d)
            month_wb.save(month_path)
        finally:
            month_wb.close()
        _cleanup_legacy_day_archive_files(d, month_path)
        return month_path

    day_wb = load_workbook(day_xlsx_path)
    try:
        src_ws = day_wb.active
        month_wb = load_workbook(month_path)
        try:
            # 레거시 YYYY-MM-DD 시트 → 일(27) 시트로 통일
            if legacy_title != sheet_title and legacy_title in month_wb.sheetnames:
                del month_wb[legacy_title]
            _clone_worksheet_to_workbook(src_ws, month_wb, sheet_title)
            # 시트 이름 일자순 정렬 (일-only·레거시 ISO 모두)
            names = [
                n
                for n in month_wb.sheetnames
                if n != sheet_title and not n.startswith("_")
            ]
            names.append(sheet_title)
            names.sort(key=_archive_sheet_sort_key)
            for i, name in enumerate(names):
                month_wb.move_sheet(name, offset=i - month_wb.sheetnames.index(name))
            _sync_archive_extra_sheets(day_wb, month_wb, d)
            month_wb.save(month_path)
        finally:
            month_wb.close()
    finally:
        day_wb.close()

    # 예전 일자별 파일·구 월폴더 경로 정리
    _cleanup_legacy_day_archive_files(d, month_path)
    return month_path


def delete_worklog_archive_sheet_at(month_path: str, d: date) -> str | None:
    """지정한 월별 xlsx에서 해당 날짜 시트 삭제. 남은 시트가 없으면 파일 삭제."""
    if load_workbook is None or not month_path or not os.path.exists(month_path):
        return None
    removed = None
    try:
        wb = load_workbook(month_path)
        try:
            for sheet_title in list(wb.sheetnames):
                if sheet_title in _worklog_archive_sheet_titles_for_lookup(d) or _is_archive_extra_sheet_name(sheet_title, d):
                    del wb[sheet_title]
                    removed = month_path
            remaining = [n for n in wb.sheetnames if n and not n.startswith("_")]
            if not remaining:
                wb.close()
                try:
                    os.remove(month_path)
                except OSError:
                    pass
                return month_path
            if removed:
                wb.save(month_path)
        finally:
            try:
                wb.close()
            except Exception:
                pass
    except Exception:
        return None
    return removed


def delete_worklog_archive_sheet(d: date) -> str | None:
    """월별 파일에서 해당 날짜 시트 삭제. 시트가 없으면 파일 삭제."""
    month_path = worklog_archive_month_path(d, create_year=False)
    if not month_path:
        return None
    _cleanup_legacy_day_archive_files(d, month_path)
    return delete_worklog_archive_sheet_at(month_path, d)


def worklog_path(d: date) -> str: return os.path.join(WORKLOG_DIR, f"{d.isoformat()}.xlsx")

def _list_archive_saved_dates() -> set[str]:
    """맥 경로 …/일지/YYYY/N월.xlsx 시트(일) → 달력 • 표시용."""
    out: set[str] = set()
    if load_workbook is None:
        return out
    root = resolve_worklog_archive_root()
    if not root or not os.path.isdir(root):
        return out
    try:
        year_names = os.listdir(root)
    except OSError:
        return out
    for year_name in year_names:
        if not (year_name.isdigit() and len(year_name) == 4):
            continue
        year = int(year_name)
        year_dir = os.path.join(root, year_name)
        if not os.path.isdir(year_dir):
            continue
        try:
            files = os.listdir(year_dir)
        except OSError:
            continue
        for fname in files:
            m = re.match(r"^(\d{1,2})월\.xlsx$", fname)
            if not m:
                continue
            month = int(m.group(1))
            if month < 1 or month > 12:
                continue
            path = os.path.join(year_dir, fname)
            try:
                wb = load_workbook(path, read_only=True)
                try:
                    for name in wb.sheetnames:
                        if name.isdigit():
                            try:
                                out.add(date(year, month, int(name)).isoformat())
                            except ValueError:
                                pass
                        elif _is_archive_extra_sheet_name(name) or _is_day_extra_sheet_name(name):
                            continue
                        else:
                            try:
                                out.add(date.fromisoformat(name).isoformat())
                            except ValueError:
                                pass
                finally:
                    wb.close()
            except Exception:
                continue
    return out


def list_saved_worklog_dates() -> set[str]:
    cached = st.session_state.get("wl_saved_dates_cache")
    if isinstance(cached, set): return cached
    _ensure_dirs()
    out: set[str] = set()
    try:
        for name in os.listdir(WORKLOG_DIR):
            if name.endswith(".xlsx") and len(name) >= 15 and name[0:4].isdigit() and name not in {"template.xlsx"} and not name.startswith("_preview_") and "_인쇄" not in name:
                out.add(name.replace(".xlsx", ""))
    except OSError:
        pass
    out.update(_list_archive_saved_dates())
    st.session_state["wl_saved_dates_cache"] = out
    return out


def _saved_dates_for_calendar() -> set[str]:
    """달력 • — 실제로 저장된 날짜만. 날짜만 바꾼 상태는 예전 날에 •가 남는다."""
    return set(list_saved_worklog_dates())

def format_worklog_date(d: date) -> str:
    weeks = "월화수목금토일"
    return f"{d.strftime('%Y-%m-%d')} ({weeks[d.weekday()]})"

def _clear_content_cells(ws) -> None:
    for r in WL_CLIENT_ROWS:
        try: ws.cell(r, 3).value = None
        except AttributeError: pass
    for r in WL_CONTENT_ROWS:
        try: ws.cell(r, 7).value = None
        except AttributeError: pass
        try: ws.cell(r, WL_REMARK_COL_START).value = None
        except AttributeError: pass
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        try: ws.cell(r, 4).value = None
        except AttributeError: pass

def _empty_cells(d: date) -> dict:
    cells = {"date": format_worklog_date(d)}
    for r in WL_CLIENT_ROWS: cells[f"C{r}"] = ""
    for r in WL_CONTENT_ROWS:
        cells[f"G{r}"] = ""
        cells[f"Y{r}"] = ""
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS: cells[f"D{r}"] = ""
    return cells

def _cells_from_worksheet(ws, d: date) -> dict:
    cells = {"date": format_worklog_date(d)}
    for r in WL_CLIENT_ROWS:
        v = ws.cell(r, 3).value
        cells[f"C{r}"] = "" if v is None else str(v)
    for r in WL_CONTENT_ROWS:
        v = ws.cell(r, 7).value
        cells[f"G{r}"] = "" if v is None else str(v)
        y = ws.cell(r, WL_REMARK_COL_START).value
        cells[f"Y{r}"] = "" if y is None else str(y)
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        v = ws.cell(r, 4).value
        cells[f"D{r}"] = "" if v is None else str(v)
    try:
        c_date = ws[WL_DATE_CELL].value
        if c_date is not None and not str(c_date).startswith("="):
            cells["date"] = str(c_date)
    except Exception:
        pass
    return cells


def read_worklog_cells(d: date) -> dict:
    path = worklog_path(d)
    if not os.path.exists(path) or load_workbook is None:
        arch = read_worklog_cells_from_archive(d)
        return arch if arch is not None else _empty_cells(d)
    wb = load_workbook(path, data_only=False)
    try:
        return _cells_from_worksheet(wb.active, d)
    finally:
        wb.close()


def read_worklog_cells_from_archive(d: date) -> dict | None:
    if load_workbook is None:
        return None
    month_path = worklog_archive_month_path(d, create_year=False)
    if not month_path or not os.path.exists(month_path):
        return None
    try:
        wb = load_workbook(month_path, data_only=False)
        try:
            name = _resolve_archive_sheet_name(wb.sheetnames, d)
            if not name:
                return None
            return _cells_from_worksheet(wb[name], d)
        finally:
            wb.close()
    except Exception:
        return None

# 💡 강제 템플릿 덮어쓰기 로직 적용
def _apply_cells_to_worksheet(ws, d: date, cells: dict, *, include_panels: bool = True) -> None:
    """활성/추가 시트에 본문 칸을 쓴다. 추가 페이지는 익일·특이를 비운다."""
    body, _extras = _detach_extra_pages(cells)
    try:
        date_cell = ws[WL_DATE_CELL]
        date_cell.value = body.get("date") or format_worklog_date(d)
        _set_body_font(date_cell)
    except AttributeError:
        pass

    for r in WL_CLIENT_ROWS:
        try:
            cell = ws.cell(r, 3)
            cell.value = (body.get(f"C{r}", "") or None)
            _set_body_font(cell)
            try: cell.alignment = cell.alignment.copy(horizontal="center", vertical="center", wrapText=False)
            except Exception: pass
        except AttributeError: pass
    for r in WL_CONTENT_ROWS:
        try:
            cell = ws.cell(r, 7)
            cell.value = (body.get(f"G{r}", "") or None)
            _set_body_font(cell)
            try: cell.alignment = cell.alignment.copy(wrapText=True, shrinkToFit=False, vertical="top")
            except Exception: pass
        except AttributeError: pass
    for r in WL_CONTENT_ROWS:
        try:
            cell = ws.cell(r, WL_REMARK_COL_START)
            cell.value = (body.get(f"Y{r}", "") or None)
            _set_body_font(cell)
            try: cell.alignment = cell.alignment.copy(wrapText=True, shrinkToFit=False, vertical="top")
            except Exception: pass
        except AttributeError: pass
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        try:
            cell = ws.cell(r, 4)
            if include_panels:
                cell.value = (body.get(f"D{r}", "") or None)
            else:
                cell.value = None
            _set_body_font(cell)
            try: cell.alignment = cell.alignment.copy(wrapText=True, shrinkToFit=False, vertical="top")
            except Exception: pass
        except AttributeError: pass


def _sync_day_extra_sheets(path: str, d: date, extra_pages: list[dict]) -> None:
    if load_workbook is None or not path or not os.path.exists(path):
        return
    wb = load_workbook(path)
    try:
        src = wb.active
        wanted: list[str] = []
        for i, extra in enumerate(extra_pages or []):
            page_n = i + 2
            title = _extra_page_sheet_name(page_n)
            wanted.append(title)
            extra_body, _ = _detach_extra_pages(extra if isinstance(extra, dict) else {})
            extra_body["date"] = extra_body.get("date") or format_worklog_date(d)
            if title in wb.sheetnames:
                ws = wb[title]
            else:
                try:
                    ws = wb.copy_worksheet(src)
                    ws.title = title
                except Exception:
                    ws = wb.create_sheet(title)
                    _copy_worksheet_cross_workbook(src, ws)
                _clear_content_cells(ws)
            _apply_cells_to_worksheet(ws, d, extra_body, include_panels=False)
            _ensure_worklog_body_merges(ws)
        for name in list(wb.sheetnames):
            if name != src.title and _is_day_extra_sheet_name(name) and name not in wanted:
                del wb[name]
        wb.save(path)
    finally:
        try:
            wb.close()
        except Exception:
            pass


def _extra_page_cells_from_workbook(wb, d: date) -> list[dict]:
    found: dict[int, dict] = {}
    active_title = wb.active.title if wb.worksheets else ""
    for name in wb.sheetnames:
        if name == active_title:
            continue
        n = _extra_page_n_from_sheet_name(name)
        if not n:
            continue
        if _is_archive_extra_sheet_name(name) and not _is_archive_extra_sheet_name(name, d):
            continue
        found[n] = _cells_from_worksheet(wb[name], d)
    return [found[n] for n in sorted(found)]


def read_worklog_extra_page_cells(d: date) -> list[dict]:
    """2페이지부터의 본문 칸. 없으면 빈 목록."""
    path = worklog_path(d)
    if os.path.exists(path) and load_workbook is not None:
        try:
            wb = load_workbook(path, data_only=False)
            try:
                extras = _extra_page_cells_from_workbook(wb, d)
                if extras:
                    return extras
            finally:
                wb.close()
        except Exception:
            pass
    return _read_extra_pages_from_archive(d)


def _read_extra_pages_from_archive(d: date) -> list[dict]:
    if load_workbook is None:
        return []
    month_path = worklog_archive_month_path(d, create_year=False)
    if not month_path or not os.path.exists(month_path):
        return []
    try:
        wb = load_workbook(month_path, data_only=False)
        try:
            return _extra_page_cells_from_workbook(wb, d)
        finally:
            wb.close()
    except Exception:
        return []


def write_cells_to_path(path: str, d: date, cells: dict, *, force_template: bool = False) -> None:
    if load_workbook is None: raise RuntimeError("openpyxl 이 필요합니다.")
    _ensure_dirs()
    body, extras = _detach_extra_pages(cells)
    if force_template or not os.path.exists(path):
        if not os.path.exists(WORKLOG_TEMPLATE): raise FileNotFoundError("업무일지 템플릿이 없습니다.")
        shutil.copy2(WORKLOG_TEMPLATE, path)
    wb = load_workbook(path)
    ws = wb.active
    if force_template: _clear_content_cells(ws)
    _apply_cells_to_worksheet(ws, d, body, include_panels=True)
    _ensure_worklog_body_merges(ws)
    wb.save(path)
    wb.close()
    if extras:
        try:
            _sync_day_extra_sheets(path, d, extras)
        except Exception:
            pass

def save_worklog_cells(d: date, cells: dict, *, force: bool = False, allow_overwrite: bool = False) -> str:
    """저장: 로컬 캐시 + 월별 일지 + Drive + (가능하면) Cloud Gist.

    allow_overwrite=False(기본) — 해당 날짜가 이미 있으면 후입력 저장 차단.
    force=True — Drive/Cloud push 시 원격 덮어쓰기(기존 일지 수정 시).
    """
    ok, block_msg = check_worklog_save_allowed(d, had_local_at_open=allow_overwrite)
    if not ok:
        raise WorklogSaveBlockedError(block_msg)
    path = worklog_path(d)
    cells = _spill_all_content(cells)
    write_cells_to_path(path, d, cells, force_template=True)
    _invalidate_saved_dates_cache()
    _remember_calendar_saved_date(d)
    _invalidate_worklog_presence_cache(d)
    try:
        from worklog_remote_sync import clear_worklog_day_deleted, invalidate_gist_days_cache
        clear_worklog_day_deleted(d.isoformat(), WORKLOG_DIR)
        invalidate_gist_days_cache()
    except Exception:
        pass
    st.session_state.pop("wl_last_archive_path", None)
    st.session_state.pop("wl_last_archive_err", None)
    st.session_state.pop("wl_last_drive_path", None)
    st.session_state.pop("wl_last_drive_month_path", None)
    st.session_state.pop("wl_last_drive_conflict", None)
    st.session_state.pop("wl_last_cloud_gist", None)
    st.session_state.pop("wl_last_cloud_err", None)
    st.session_state.pop("wl_last_archive_sheet", None)
    st.session_state.pop("wl_last_archive_target", None)
    try:
        archive = upsert_worklog_archive_sheet(d, path, allow_overwrite=allow_overwrite or force)
        if archive:
            st.session_state["wl_last_archive_path"] = archive
            st.session_state["wl_last_archive_sheet"] = worklog_archive_sheet_title(d)
            st.session_state["wl_last_archive_target"] = describe_worklog_archive_target(d)
    except Exception as e:
        st.session_state["wl_last_archive_err"] = str(e)
    try:
        from drive_autoload import push_worklog_day_to_drive, push_worklog_month_archive_to_drive

        drv = push_worklog_day_to_drive(path, WORKLOG_DIR, force=force)
        if drv:
            st.session_state["wl_last_drive_path"] = drv
        elif not force:
            pres = detect_worklog_date_presence(d)
            if pres.get("drive"):
                st.session_state["wl_last_drive_conflict"] = d.isoformat()
        arch_path = st.session_state.get("wl_last_archive_path")
        if arch_path and os.path.isfile(arch_path):
            mdrv = push_worklog_month_archive_to_drive(arch_path, year=d.year, force=force)
            if mdrv:
                st.session_state["wl_last_drive_month_path"] = mdrv
    except Exception:
        pass
    return path


def create_worklog_day_local(d: date) -> dict:
    """선택한 날짜의 엑셀 업무일지를 로컬에만 만든다 (캐시 + 월별 시트).

    Drive/Cloud는 작성 후 「저장」할 때 반영. 이미 있으면 덮어쓰지 않는다.
    """
    _ensure_dirs()
    path = worklog_path(d)
    archive_exists = worklog_date_exists_in_archive(d)
    local_exists = os.path.isfile(path)
    archive = None
    if not local_exists:
        write_cells_to_path(path, d, _empty_cells(d), force_template=True)
    if not archive_exists:
        archive = upsert_worklog_archive_sheet(d, path, allow_overwrite=False)
    else:
        archive = worklog_archive_month_path(d, create_year=False)
    _invalidate_saved_dates_cache()
    _invalidate_worklog_presence_cache(d)
    created = not (local_exists or archive_exists)
    iso = d.isoformat()
    st.session_state[f"wl_saved_ok_{iso}"] = True
    ctx = dict(st.session_state.get(f"wl_open_ctx_{iso}") or {})
    ctx["had_local"] = True
    st.session_state[f"wl_open_ctx_{iso}"] = ctx
    cells = read_worklog_cells(d) if os.path.isfile(path) else _empty_cells(d)
    try:
        _publish_view_cells(d, cells)
    except Exception:
        pass
    st.session_state.pop(f"wl_sum_sig_v25_{iso}", None)
    st.session_state.pop(f"wl_sum_html_v25_{iso}", None)
    st.session_state[f"wl_left_excel_on_{iso}"] = False
    saved_set = st.session_state.get("wl_saved_dates_cache")
    if isinstance(saved_set, set):
        saved_set.add(iso)
    target = describe_worklog_archive_target(d)
    if created:
        msg = f"로컬 엑셀을 만들었습니다 · {target}"
    else:
        msg = f"이미 있는 날짜입니다 · {target}"
    return {
        "created": created,
        "path": path,
        "archive": archive,
        "target": target,
        "msg": msg,
        "cells": cells,
    }


def _purge_worklog_day_preview_cache(d: date) -> None:
    """삭제 후 왼쪽 요약·엑셀 HTML 캐시 제거 (불필요한 재렌더·로딩 방지)."""
    iso = d.isoformat()
    for k in (
        f"wl_sum_sig_{iso}",
        f"wl_sum_html_{iso}",
        f"wl_print_cells_sig_{iso}",
        f"wl_print_cells_path_{iso}",
        f"wl_left_excel_on_{iso}",
        f"wl_left_excel_path_{iso}",
        f"wl_left_excel_sig_v24_{iso}",
        f"wl_left_excel_html_v24_{iso}",
        f"wl_left_excel_h_v24_{iso}",
        f"wl_left_excel_html_v27_{iso}",
        f"wl_left_excel_h_v27_{iso}",
        f"wl_left_excel_skel_v27_{iso}",
        f"wl_form_sig_v14_{iso}",
        f"wl_remote_pull_tried_{iso}",
    ):
        st.session_state.pop(k, None)
    for k in list(st.session_state.keys()):
        if not isinstance(k, str):
            continue
        if iso in k and (
            k.startswith("wl_print_html_cache_")
            or k.startswith("wl_print_html_meta_")
            or k.startswith("wl_left_excel_html_")
        ):
            st.session_state.pop(k, None)


def _delete_worklog_day_remote_sync(d: date) -> tuple[list[str], str]:
    """Drive 원격 삭제 (느림 — UI 블로킹 방지용 분리)."""
    removed: list[str] = []
    try:
        from drive_autoload import delete_worklog_day_from_drive

        for dn in delete_worklog_day_from_drive(d, WORKLOG_DIR):
            removed.append(f"Drive:{dn}")
    except Exception:
        pass
    return removed, ""


def _worklog_remote_delete_job(d: date) -> None:
    try:
        _delete_worklog_day_remote_sync(d)
        try:
            month_path = worklog_archive_month_path(d, create_year=False)
            if month_path and os.path.isfile(month_path):
                from drive_autoload import push_worklog_month_archive_to_drive
                push_worklog_month_archive_to_drive(month_path, year=d.year, force=True)
        except Exception:
            pass
        try:
            from worklog_remote_sync import invalidate_gist_days_cache
            invalidate_gist_days_cache()
        except Exception:
            pass
    except Exception:
        pass


def _schedule_worklog_remote_delete(d: date) -> None:
    iso = d.isoformat()
    lock_k = f"_wl_remote_del_started_{iso}"
    if st.session_state.get(lock_k):
        return
    st.session_state[lock_k] = True
    threading.Thread(target=_worklog_remote_delete_job, args=(d,), daemon=True).start()


def _take_day_action_flags(iso: str) -> dict:
    return {
        "save": bool(st.session_state.pop(f"wl_do_save_{iso}", None)),
        "add": bool(st.session_state.pop(f"wl_do_add_{iso}", None)),
        "del": st.session_state.pop(f"wl_do_del_{iso}", None),
    }


def _put_day_action_flags(iso: str, flags: dict | None) -> None:
    if not flags:
        return
    if flags.get("save"):
        st.session_state[f"wl_do_save_{iso}"] = True
    if flags.get("add"):
        st.session_state[f"wl_do_add_{iso}"] = True
    if flags.get("del") is not None:
        st.session_state[f"wl_do_del_{iso}"] = flags["del"]


def _merge_day_action_flags(*groups: dict | None) -> dict:
    out = {"save": False, "add": False, "del": None}
    for g in groups:
        if not g:
            continue
        out["save"] = out["save"] or bool(g.get("save"))
        out["add"] = out["add"] or bool(g.get("add"))
        if g.get("del") is not None:
            out["del"] = g.get("del")
    return out


def _queue_worklog_save(iso: str) -> None:
    """저장 버튼 on_click — 지금 열린 편집 날짜에 저장한다."""
    try:
        _snapshot_worklog_page(iso)
    except Exception:
        pass
    st.session_state[f"wl_do_save_{iso}"] = True


def _queue_worklog_add(iso: str) -> None:
    """항목/페이지 추가 on_click — fragment를 한 번 더 rerun 하지 않는다."""
    try:
        _snapshot_worklog_page(iso)
    except Exception:
        pass
    st.session_state[f"wl_do_add_{iso}"] = True


def _queue_worklog_page(iso: str, idx: int) -> None:
    try:
        _snapshot_worklog_page(iso)
    except Exception:
        pass
    st.session_state[_page_idx_key(iso)] = max(0, int(idx))
    st.session_state.pop(f"wl_focus_ln_{iso}", None)
    st.session_state.pop(f"wl_focus_caret_{iso}", None)
    st.session_state.pop("wl_active_cell_key", None)
    st.session_state.pop("wl_active_cell_sel", None)
    if int(idx) > 0:
        html_k = f"wl_left_excel_html_v27_{iso}"
        cached = str(st.session_state.get(html_k) or "")
        page_n = int(idx) + 1
        if cached and f'data-wl-page="{page_n}"' not in cached and f'data-wl="p{page_n}-' not in cached:
            st.session_state[f"wl_left_excel_rebuild_{iso}"] = True


def _add_worklog_input_page(iso: str) -> None:
    n = _page_count_for(iso)
    if n >= WL_MAX_PAGES:
        return
    st.session_state[_page_count_key(iso)] = n + 1
    st.session_state[_page_idx_key(iso)] = n
    _seed_entry_clients(iso, n, [""])
    _apply_entry_lines(iso, n, [""], remount_comp=True)
    _seed_entry_remarks(iso, n, [""])
    _snapshot_worklog_page(iso, n)
    st.session_state.pop(f"wl_left_excel_html_v27_{iso}", None)
    st.session_state[f"wl_left_excel_rebuild_{iso}"] = True


def _queue_worklog_del_entry(iso: str, idx: int) -> None:
    """항목 삭제 on_click."""
    st.session_state[f"wl_do_del_{iso}"] = int(idx)


def _queue_close_delete_popover() -> None:
    """삭제 확정 후 팝업을 닫는다.

    콜백에서 popover 키를 False로 두면 Streamlit이 프론트 상태(열림)로 다시 덮는다.
    위젯을 그리기 전에 인스턴스를 올려 닫힌 팝업으로 다시 단다.
    """
    st.session_state["wl_del_day_force_close"] = True
    st.session_state["wl_skip_sync_once"] = True


def _worklog_delete_popover_key() -> str:
    inst = int(st.session_state.get("wl_del_day_inst") or 0)
    return f"wl_del_day_open_{inst}"


def _flush_worklog_delete_popover() -> None:
    """삭제 popover를 만들기 전에 호출. force_close면 닫힌 새 인스턴스로 교체."""
    if not st.session_state.pop("wl_del_day_force_close", None):
        return
    inst = int(st.session_state.get("wl_del_day_inst") or 0)
    st.session_state.pop("wl_del_day_open", None)
    st.session_state.pop(f"wl_del_day_open_{inst}", None)
    st.session_state[f"wl_del_day_open_{inst}"] = False
    st.session_state["wl_del_day_inst"] = inst + 1


def _on_cancel_delete_day() -> None:
    """삭제 인라인 확인 취소 — 확인 UI만 닫는다."""
    _pin_worklog_scroll()
    st.session_state["wl_del_confirm_open"] = False


def _on_confirm_delete_day() -> None:
    """확정 on_click — 삭제를 예약하고 인라인 확인 UI를 닫는다."""
    _pin_worklog_scroll()
    d = st.session_state.get("worklog_selected")
    if isinstance(d, date):
        st.session_state["wl_do_delete_day"] = d.isoformat()
    elif isinstance(d, str) and d:
        st.session_state["wl_do_delete_day"] = d
    st.session_state["wl_del_confirm_open"] = False
    _queue_close_delete_popover()


def _run_pending_worklog_day_delete() -> bool:
    """wl_do_delete_day 플래그가 있으면 로컬 삭제 실행. 처리했으면 True."""
    del_iso = st.session_state.pop("wl_do_delete_day", None)
    if not del_iso:
        return False
    try:
        d_del = date.fromisoformat(str(del_iso))
        delete_worklog_day(d_del, remote=False)
        _schedule_worklog_remote_delete(d_del)
        prev = [str(x) for x in (st.session_state.get("wl_purge_dates") or []) if str(x) != str(del_iso)]
        st.session_state["wl_purge_dates"] = prev
        if st.session_state.get("wl_date_retarget_from") == str(del_iso):
            st.session_state.pop("wl_date_retarget_from", None)
        _queue_close_delete_popover()
        _pin_worklog_scroll()
        try:
            _prepare_worklog_day_state(d_del, skip_remote_pull=True)
        except Exception:
            pass
        return True
    except Exception:
        _queue_close_delete_popover()
        return False


def purge_worklog_day_files(d: date, *, remote: bool = False) -> list[str]:
    """그 날짜의 저장 파일·월별 시트만 지운다. 다른 날짜 위젯은 건드리지 않는다."""
    _ensure_dirs()
    iso = d.isoformat()
    removed: list[str] = []
    targets = [
        worklog_path(d),
        os.path.join(WORKLOG_DIR, f"_preview_{iso}.xlsx"),
        os.path.join(WORKLOG_DIR, f"일일업무일지_{iso}_인쇄.xlsx"),
    ]
    try:
        for name in os.listdir(WORKLOG_DIR):
            if name == "template.xlsx" or not name.endswith(".xlsx"):
                continue
            if name == f"{iso}.xlsx" or name.startswith(f"{iso}_") or name.startswith(f"_preview_{iso}") or name.startswith(f"일일업무일지_{iso}"):
                targets.append(os.path.join(WORKLOG_DIR, name))
    except OSError:
        pass
    seen: set[str] = set()
    for path in targets:
        if path in seen:
            continue
        seen.add(path)
        if os.path.exists(path):
            try:
                os.remove(path)
                removed.append(os.path.basename(path))
            except OSError:
                pass
    try:
        arch = delete_worklog_archive_sheet(d)
        if arch:
            removed.append(f"{os.path.basename(arch)}#{worklog_archive_sheet_title(d)}")
    except Exception:
        pass
    try:
        from worklog_remote_sync import mark_worklog_day_deleted
        mark_worklog_day_deleted(iso, WORKLOG_DIR)
    except Exception:
        pass
    if remote:
        extra, _note = _delete_worklog_day_remote_sync(d)
        removed.extend(extra)
        try:
            from worklog_remote_sync import invalidate_gist_days_cache
            invalidate_gist_days_cache()
        except Exception:
            pass
    _drop_saved_date_from_cache(iso)
    _invalidate_worklog_presence_cache(d)
    _purge_worklog_day_preview_cache(d)
    return removed


def delete_worklog_day(d: date, *, remote: bool = True) -> list[str]:
    """선택한 그 날짜만 삭제. 옮긴 다른 날짜 파일·입력은 유지한다."""
    iso = d.isoformat()
    removed = purge_worklog_day_files(d, remote=remote)
    _clear_date_widget_state(d)
    _gone = {
        "local": False,
        "archive": False,
        "drive": False,
        "cloud": False,
        "any": False,
        "locations": [],
    }
    st.session_state[f"wl_open_ctx_{iso}"] = {"had_local": False, "presence": _gone}
    st.session_state[f"wl_presence_{iso}_fast"] = _gone
    st.session_state[f"wl_presence_{iso}_all"] = _gone
    st.session_state[f"wl_arch_exists_{iso}"] = False
    st.session_state.pop(f"wl_saved_ok_{iso}", None)
    empty = [_empty_sheet_entry()]
    st.session_state[_boot_key(d)] = True
    st.session_state[_entries_key(d)] = empty
    st.session_state[_next_key(d)] = ""
    st.session_state[_notes_key(d)] = ""
    st.session_state[f"wl_entry_count_{iso}"] = 1
    st.session_state[_page_count_key(iso)] = 1
    st.session_state[_page_idx_key(iso)] = 0
    msg = f"삭제 완료" + (f": {', '.join(removed)}" if removed else " (저장본 없음, 입력만 초기화)")
    if not remote:
        msg += " · Cloud/Drive 정리 중"
    st.session_state[f"wl_pending_sync_{iso}"] = {"entries": empty, "next": "", "notes": "", "extra_pages": [], "msg": msg}
    try:
        _publish_view_cells(d, _empty_cells(d))
    except Exception:
        st.session_state.pop(_view_cells_key(d), None)
    st.session_state["_wl_drive_sync_ts"] = time.time()
    return removed

def _worklog_cells_have_draft(cells: dict | None) -> bool:
    src, extras = _detach_extra_pages(cells)
    if any(
        str(src.get(f"G{r}", "") or "").strip()
        or str(src.get(f"C{r}", "") or "").strip()
        or str(src.get(f"Y{r}", "") or "").strip()
        for r in WL_CONTENT_ROWS
    ):
        return True
    if any(str(src.get(f"D{r}", "") or "").strip() for r in WL_NEXT_ROWS + WL_NOTE_ROWS):
        return True
    return any(_worklog_cells_have_draft(extra) for extra in extras)


def _read_leftover_archive_day_cells(d: date) -> dict | None:
    """구 경로 일자파일(…/2026/2026-09-03.xlsx)에 내용이 있으면 읽는다."""
    if load_workbook is None:
        return None
    year_dir = worklog_archive_year_dir(d, create=False)
    if not year_dir:
        return None
    for p in (
        os.path.join(year_dir, f"{d.isoformat()}.xlsx"),
        os.path.join(year_dir, f"{d.month}월", f"{d.isoformat()}.xlsx"),
    ):
        if not os.path.isfile(p):
            continue
        try:
            wb = load_workbook(p, data_only=False)
            try:
                return _cells_from_worksheet(wb.active, d)
            finally:
                wb.close()
        except Exception:
            continue
    return None


def worklog_archive_has_saved_content(d: date) -> bool:
    """월별 시트·구 일자파일에 실제 입력(거래처/내용/예정/비고)이 있는지.

    시트만 있고 칸이 비어 있으면 False — 달력 • 없는 날과 같게 취급한다.
    """
    cells = read_worklog_cells_from_archive(d)
    if cells is not None and _worklog_cells_have_draft(cells):
        return True
    leftover = _read_leftover_archive_day_cells(d)
    return leftover is not None and _worklog_cells_have_draft(leftover)


def _worklog_day_is_persisted(d: date) -> bool:
    """로컬 파일·월별 시트·방금 저장 표시가 있으면 저장된 날."""
    iso = d.isoformat()
    if os.path.exists(worklog_path(d)):
        return True
    if st.session_state.get(f"wl_saved_ok_{iso}"):
        return True
    ctx = st.session_state.get(f"wl_open_ctx_{iso}") or {}
    if ctx.get("had_local"):
        return True
    try:
        if iso in list_saved_worklog_dates():
            return True
    except Exception:
        pass
    try:
        return bool(worklog_date_exists_in_archive(d))
    except Exception:
        return False


def _worklog_day_has_saved_or_draft(d: date) -> bool:
    if _worklog_day_is_persisted(d):
        return True
    try:
        cells = _cells_from_widgets(d)
    except Exception:
        cells = read_worklog_cells(d)
    return _worklog_cells_have_draft(cells)


def _load_cells_for_reassign(old: date) -> dict:
    try:
        cells = _cells_from_widgets(old)
        if _worklog_cells_have_draft(cells):
            return cells
    except Exception:
        pass
    cells = read_worklog_cells(old)
    cells = _attach_extra_pages(cells, read_worklog_extra_page_cells(old))
    if _worklog_cells_have_draft(cells):
        return cells
    arch = read_worklog_cells_from_archive(old)
    if arch is not None:
        return _attach_extra_pages(arch, _read_extra_pages_from_archive(old))
    return cells


def _patch_saved_dates_after_move(old: date, new: date) -> None:
    """달력 • 를 즉시 맞춘다. 7일 빼고 3일 넣음."""
    cached = st.session_state.get("wl_saved_dates_cache")
    if isinstance(cached, set):
        cached.discard(old.isoformat())
        cached.add(new.isoformat())
        st.session_state["wl_saved_dates_cache"] = cached
    else:
        _invalidate_saved_dates_cache()
    _invalidate_worklog_presence_cache(old)
    _invalidate_worklog_presence_cache(new)


def _mark_worklog_day_writable(d: date, *, had_local: bool) -> None:
    """이동·저장 후 이 날짜에 다시 저장할 수 있게 연다."""
    iso = d.isoformat()
    ctx = dict(st.session_state.get(f"wl_open_ctx_{iso}") or {})
    ctx["had_local"] = bool(had_local)
    st.session_state[f"wl_open_ctx_{iso}"] = ctx
    if had_local:
        st.session_state[f"wl_saved_ok_{iso}"] = True


def _flush_queued_date_pick() -> None:
    """날짜칸 위젯을 그리기 전에, 이전 런에서 미뤄 둔 값을 넣는다."""
    st.session_state["_wl_date_pick_live"] = False
    nxt = st.session_state.pop("_wl_date_pick_next", None)
    if isinstance(nxt, date):
        st.session_state["wl_date_pick"] = nxt
        st.session_state["wl_date_sync"] = nxt.isoformat()


def _set_wl_date_pick(d: date) -> None:
    """위젯이 이미 뜨면 키를 쓰지 않고 다음 런으로만 넘긴다."""
    st.session_state["wl_date_sync"] = d.isoformat()
    if st.session_state.get("_wl_date_pick_live"):
        st.session_state["_wl_date_pick_next"] = d
        return
    try:
        st.session_state["wl_date_pick"] = d
    except Exception:
        st.session_state["_wl_date_pick_next"] = d


def _switch_worklog_selected_date(new: date) -> None:
    st.session_state["worklog_selected"] = new
    st.session_state["worklog_month"] = date(new.year, new.month, 1)
    _set_wl_date_pick(new)


def apply_worklog_date_change(old: date, new: date) -> str:
    """저장된 7일도 날짜칸을 3일로 바꾸면 그 일지가 3일이 된다.

    대상 날짜에 이미 저장된 데이터가 있으면 옮기지 않는다.
    빈 날은 날짜만 전환한다. 성공 시 빈 문자열.
    """
    if old == new:
        return ""
    if _worklog_dest_has_saved_data(new):
        return _WL_MOVE_BLOCK_MSG
    if _worklog_day_has_saved_or_draft(old):
        try:
            reassign_worklog_date(old, new, overwrite_dest=True)
        except (FileExistsError, WorklogSaveBlockedError) as e:
            return str(e)
        return ""
    _clear_date_widget_state(old)
    _switch_worklog_selected_date(new)
    return ""


def _remember_purge_date(d: date) -> None:
    """저장 때 지울 예전 날짜. 날짜만 바꾼 뒤에도 잃지 않게 목록으로 쌓는다."""
    iso = d.isoformat()
    prev = [str(x) for x in (st.session_state.get("wl_purge_dates") or []) if x]
    if iso not in prev:
        prev.append(iso)
    st.session_state["wl_purge_dates"] = prev
    if not st.session_state.get("wl_date_retarget_from"):
        st.session_state["wl_date_retarget_from"] = iso


def _take_purge_dates(target: date) -> list[date]:
    raw = [str(x) for x in (st.session_state.get("wl_purge_dates") or []) if x]
    rf = st.session_state.get("wl_date_retarget_from")
    if isinstance(rf, str) and rf:
        raw.append(rf)
    out: list[date] = []
    seen: set[str] = set()
    for iso in raw:
        try:
            d = date.fromisoformat(iso)
        except ValueError:
            continue
        if d == target or d.isoformat() in seen:
            continue
        seen.add(d.isoformat())
        out.append(d)
    return out


def _worklog_editor_occupied(d: date) -> bool:
    """화면 세션이나 저장본이 있으면 그 날짜를 연다 (내용을 덮어 복사하지 않음)."""
    if st.session_state.get(_boot_key(d)) and _entries_key(d) in st.session_state:
        entries = st.session_state.get(_entries_key(d)) or []
        if any(str((e or {}).get("content") or "").strip() or str((e or {}).get("client") or "").strip() for e in entries if isinstance(e, dict)):
            return True
        if str(st.session_state.get(_next_key(d), "") or "").strip() or str(st.session_state.get(_notes_key(d), "") or "").strip():
            return True
        iso = d.isoformat()
        for i in range(1, _page_count_for(iso)):
            cl, co, rm = _sheet_lines_from_widgets_at(iso, i)
            if any(str(x).strip() for x in cl + co + rm):
                return True
    if os.path.isfile(worklog_path(d)):
        try:
            return _worklog_cells_have_draft(read_worklog_cells(d))
        except Exception:
            return True
    try:
        if d.isoformat() in _saved_dates_for_calendar():
            return True
    except Exception:
        pass
    try:
        return bool(worklog_archive_has_saved_content(d))
    except Exception:
        return False


def _worklog_dest_has_saved_data(d: date) -> bool:
    """이미 저장된 일지가 있는지. 빈 시트·잔여 빈 파일은 제외."""
    path = worklog_path(d)
    if os.path.isfile(path):
        try:
            if _worklog_cells_have_draft(read_worklog_cells(d)):
                return True
        except Exception:
            return True
    try:
        return bool(worklog_archive_has_saved_content(d))
    except Exception:
        return False


def _block_move_to_saved_date(stay: date) -> None:
    st.session_state["wl_date_err"] = _WL_MOVE_BLOCK_MSG
    _set_wl_date_pick(stay)
    st.session_state["wl_skip_sync_once"] = True


def _worklog_day_already_saved(d: date) -> bool:
    """이 날짜에 이미 저장한 일지가 있으면 True. 날짜 변경 시 지우지 않기 위함."""
    if st.session_state.get(f"wl_saved_ok_{d.isoformat()}"):
        return True
    return _worklog_dest_has_saved_data(d)


def try_retarget_worklog_editor_date(old: date, new: date) -> bool:
    """업무일지 날짜만 바꾼다. 입력 중인 내용은 그대로 새 날짜로 옮긴다."""
    if old == new:
        return True
    try:
        _snapshot_worklog_page(old.isoformat())
    except Exception:
        pass
    retarget_worklog_editor_date(old, new)
    _pin_worklog_scroll()
    st.session_state["wl_skip_sync_once"] = True
    return True


def _date_move_mode_on() -> bool:
    return bool(st.session_state.get("wl_date_move_mode"))


def _on_toggle_date_move_mode() -> None:
    st.session_state["wl_date_move_mode"] = not _date_move_mode_on()
    st.session_state.pop("wl_date_err", None)


def _move_worklog_editor_to_date(new: date) -> bool:
    """편집중인 글의 날짜만 바꾼다. 대상에 이미 저장본이 있으면 옮기지 않는다."""
    old = st.session_state.get("worklog_selected")
    if not isinstance(old, date):
        st.session_state["wl_date_move_mode"] = False
        _open_worklog_saved_date(new)
        return True
    if old == new:
        st.session_state["wl_date_move_mode"] = False
        return True
    if _worklog_dest_has_saved_data(new):
        _block_move_to_saved_date(old)
        return False
    st.session_state.pop(_parked_cells_key(old), None)
    st.session_state.pop(_parked_cells_key(new), None)
    try_retarget_worklog_editor_date(old, new)
    st.session_state["wl_date_move_mode"] = False
    return True


def _parked_cells_key(d: date) -> str:
    return f"wl_parked_cells_{d.isoformat()}"


def _park_shared_editor_to_date(d: date) -> None:
    """지금 입력칸을 그 날짜 초안으로 남겨 둔다. 달력으로 다시 오면 그대로 연다."""
    iso = d.isoformat()
    try:
        cells = _cells_from_widgets(d)
    except Exception:
        return
    if not isinstance(cells, dict):
        return
    st.session_state[_parked_cells_key(d)] = cells
    st.session_state[_entries_key(d)] = [_sheet_entry_from_cells(cells)]
    st.session_state[_view_cells_key(d)] = cells
    st.session_state[_boot_key(d)] = True
    try:
        _snapshot_worklog_page(iso)
    except Exception:
        pass


def _bind_editor_from_cells(d: date, cells: dict, *, remount_comp: bool = False) -> dict:
    """공유 입력칸에 그날 값을 넣는다. remount 없이 rev만 올려 달력 이동 로딩을 줄인다."""
    packed = dict(cells or _empty_cells(d))
    body, extras_cells = _detach_extra_pages(packed)
    extras = [_sheet_entry_from_cells(x) for x in (extras_cells or [])]
    _, nd, nt = _entries_from_cells(body)
    _seed_day_entry_widgets(
        d,
        [_sheet_entry_from_cells(body)],
        "\n".join(nd),
        "\n".join(nt),
        extra_pages=extras,
        remount_comp=remount_comp,
    )
    st.session_state[_boot_key(d)] = True
    return packed


def _reload_worklog_date_from_storage(d: date) -> dict:
    """그 날짜 저장본으로 입력칸·요약을 다시 심는다. 다른 날 초안은 복사하지 않는다."""
    iso = d.isoformat()
    parked = st.session_state.get(_parked_cells_key(d))
    if isinstance(parked, dict):
        return _bind_editor_from_cells(d, parked, remount_comp=False)
    st.session_state.pop(_boot_key(d), None)
    st.session_state.pop(_entries_key(d), None)
    st.session_state.pop(_view_cells_key(d), None)
    st.session_state.pop(f"wl_pending_sync_{iso}", None)
    try:
        cells = _stored_cells_for_date(d)
    except Exception:
        cells = _empty_cells(d)
    return _bind_editor_from_cells(d, cells or _empty_cells(d), remount_comp=False)


def _open_worklog_saved_date(new: date) -> None:
    """그 날짜를 연다. 저장본을 업무내용·요약에 불러온다."""
    old = st.session_state.get("worklog_selected")
    if isinstance(old, date) and old != new:
        try:
            _park_shared_editor_to_date(old)
        except Exception:
            pass
    st.session_state["worklog_month"] = date(new.year, new.month, 1)
    _switch_worklog_selected_date(new)
    st.session_state.pop("wl_date_err", None)
    st.session_state.pop("wl_pending_date_change", None)
    st.session_state["wl_skip_sync_once"] = True
    try:
        cells = _reload_worklog_date_from_storage(new)
        _publish_view_cells(new, cells)
    except Exception:
        _prepare_worklog_day_state(new, skip_remote_pull=True)
    _pin_worklog_scroll()


def retarget_worklog_editor_date(old: date, new: date) -> None:
    """저장 전 날짜만 바꾼다. 파일·월별 시트는 건드리지 않는다.

    화면 내용만 새 날짜로 옮긴다. 예전 날짜 데이터는 저장 버튼을 누를 때 삭제한다.
    """
    if old == new:
        return
    try:
        cells = _cells_from_widgets(old)
    except Exception:
        cells = read_worklog_cells(old)
    cells = dict(cells or _empty_cells(new))
    cells["date"] = format_worklog_date(new)
    entries = [_sheet_entry_from_cells(cells)]
    extra_entries = [_sheet_entry_from_cells(x) for x in (_detach_extra_pages(cells)[1] or _extra_page_cells_from_widgets(old))]
    _, nd, nt = _entries_from_cells(cells)
    next_txt = "\n".join(nd)
    notes_txt = "\n".join(nt)
    _remember_purge_date(old)
    _switch_worklog_selected_date(new)
    st.session_state[_boot_key(new)] = True
    st.session_state[_entries_key(new)] = entries
    st.session_state[_next_key(new)] = next_txt
    st.session_state[_notes_key(new)] = notes_txt
    st.session_state[f"wl_entry_count_{new.isoformat()}"] = 1
    st.session_state[f"wl_pending_sync_{new.isoformat()}"] = {
        "entries": entries, "next": next_txt, "notes": notes_txt, "extra_pages": extra_entries, "msg": "",
        "keep_editor": True,
    }
    try:
        _publish_view_cells(new, cells)
    except Exception:
        pass
    st.session_state.pop(f"wl_saved_ok_{new.isoformat()}", None)
    st.session_state.pop("wl_date_err", None)
    st.session_state.pop("wl_pending_date_change", None)


def commit_worklog_date_save(source: date, target: date, cells: dict) -> str:
    """고른 날짜에 저장한다. 날짜만 바꾼 경우 예전 날짜 파일은 저장 때 지운다."""
    cells = dict(cells or {})
    cells["date"] = format_worklog_date(target)
    path = save_worklog_cells(target, cells, force=True, allow_overwrite=True)
    purge = _take_purge_dates(target)
    if source != target and source not in purge:
        purge.append(source)
    for old in purge:
        purge_worklog_day_files(old, remote=False)
        try:
            _schedule_worklog_remote_delete(old)
        except Exception:
            pass
        try:
            _clear_date_widget_state(old)
        except Exception:
            pass
        _invalidate_worklog_presence_cache(old)
    _switch_worklog_selected_date(target)
    _mark_worklog_day_writable(target, had_local=True)
    st.session_state.pop("wl_date_err", None)
    st.session_state.pop("wl_date_retarget_from", None)
    st.session_state.pop("wl_purge_dates", None)
    st.session_state.pop("wl_pending_date_change", None)
    _invalidate_saved_dates_cache()
    _remember_calendar_saved_date(target)
    _invalidate_worklog_presence_cache(target)
    return path


def consume_left_date_pick_move(selected: date) -> tuple[date, bool]:
    """왼쪽 날짜칸이 selected와 다르면 연다. 날짜변경 모드면 편집중 내용을 옮긴다."""
    picked = st.session_state.get("wl_date_pick")
    if not isinstance(picked, date) or picked == selected:
        return selected, False
    if _date_move_mode_on():
        ok = _move_worklog_editor_to_date(picked)
        return (picked if ok else selected), ok
    _open_worklog_saved_date(picked)
    return picked, True


def _on_wl_date_pick_change() -> None:
    """날짜칸: 기본은 그날 저장본을 연다. 날짜변경이 켜져 있으면 편집중 내용을 옮긴다."""
    st.session_state["_wl_date_pick_live"] = False
    picked = st.session_state.get("wl_date_pick")
    selected = st.session_state.get("worklog_selected")
    if not isinstance(picked, date):
        return
    if isinstance(selected, date) and picked == selected:
        return
    if _date_move_mode_on():
        _move_worklog_editor_to_date(picked)
        return
    _open_worklog_saved_date(picked)


def _on_wl_cal_day(iso: str) -> None:
    """달력: 기본은 그날을 연다. 날짜변경이 켜져 있으면 편집중 내용을 옮긴다."""
    try:
        new = date.fromisoformat(iso)
    except ValueError:
        return
    st.session_state["worklog_month"] = date(new.year, new.month, 1)
    old = st.session_state.get("worklog_selected")
    if isinstance(old, date) and old == new:
        return
    st.session_state["_wl_date_pick_live"] = False
    st.session_state.pop(f"wl_do_save_{iso}", None)
    if isinstance(old, date):
        st.session_state.pop(f"wl_do_save_{old.isoformat()}", None)
    if _date_move_mode_on():
        _move_worklog_editor_to_date(new)
        return
    _open_worklog_saved_date(new)


def _run_pending_worklog_date_change() -> bool:
    """대기 중인 날짜 변경은 대상 날짜를 연다. 파일은 저장 버튼에서 기록한다."""
    pending = st.session_state.pop("wl_pending_date_change", None)
    if not pending:
        return False
    try:
        old = date.fromisoformat(str(pending[0]))
        new = date.fromisoformat(str(pending[1]))
    except Exception:
        return False
    old_flags = _take_day_action_flags(old.isoformat())
    new_flags = _take_day_action_flags(new.isoformat())
    _open_worklog_saved_date(new)
    merged = _merge_day_action_flags(old_flags, new_flags)
    merged["save"] = False
    _put_day_action_flags(new.isoformat(), merged)
    return True


def reassign_worklog_date(old: date, new: date, *, overwrite_dest: bool = True) -> str:
    if old == new: return "same"
    if not overwrite_dest:
        ok, block_msg = check_worklog_save_allowed(new, had_local_at_open=False)
        if not ok:
            raise FileExistsError(block_msg)
    cells = _load_cells_for_reassign(old)
    cells["date"] = format_worklog_date(new)
    old_local = os.path.exists(worklog_path(old))
    old_persisted = _worklog_day_is_persisted(old)
    should_write = old_persisted or old_local or _worklog_cells_have_draft(cells)
    if should_write:
        save_worklog_cells(new, cells, force=True, allow_overwrite=overwrite_dest)
    if old_persisted or old_local:
        for path in (worklog_path(old), _preview_path(old), _print_xlsx_path(old)):
            if os.path.exists(path):
                try: os.remove(path)
                except OSError: pass
        try:
            delete_worklog_archive_sheet(old)
        except Exception:
            pass
        try:
            from worklog_remote_sync import mark_worklog_day_deleted
            mark_worklog_day_deleted(old.isoformat(), WORKLOG_DIR)
        except Exception:
            pass
        try:
            _schedule_worklog_remote_delete(old)
        except Exception:
            pass
    entries = [_sheet_entry_from_cells(cells)]
    extra_entries = [_sheet_entry_from_cells(x) for x in _detach_extra_pages(cells)[1]]
    _, nd, nt = _entries_from_cells(cells)
    _clear_date_widget_state(old)
    _switch_worklog_selected_date(new)
    st.session_state[_boot_key(new)] = True
    st.session_state[_entries_key(new)] = entries
    st.session_state[_next_key(new)] = "\n".join(nd)
    st.session_state[_notes_key(new)] = "\n".join(nt)
    st.session_state[f"wl_entry_count_{new.isoformat()}"] = 1
    st.session_state[f"wl_pending_sync_{new.isoformat()}"] = {
        "entries": entries,
        "extra_pages": extra_entries,
        "next": "\n".join(nd),
        "notes": "\n".join(nt),
        "msg": "",
    }
    _mark_worklog_day_writable(new, had_local=True)
    _invalidate_saved_dates_cache()
    _patch_saved_dates_after_move(old, new)
    return "moved" if should_write else "retargeted"

def _cell_fill_color(cell) -> str | None:
    try:
        fill = cell.fill
        if not fill or fill.fill_type is None: return None
        fg = fill.fgColor
        if fg is None: return None
        if getattr(fg, "type", None) == "rgb" and fg.rgb and fg.rgb != "00000000":
            rgb = str(fg.rgb)
            if len(rgb) == 8: rgb = rgb[2:]
            return f"#{rgb}"
    except Exception: return None
    return None

def _border_css(cell) -> str:
    parts = []
    try:
        b = cell.border
        for side_name, css_side in (("left", "border-left"), ("right", "border-right"), ("top", "border-top"), ("bottom", "border-bottom")):
            side = getattr(b, side_name, None)
            if side and side.style:
                color = "#111"
                try:
                    c = side.color
                    if c is not None and getattr(c, "type", None) == "rgb" and c.rgb:
                        rgb = str(c.rgb)
                        if len(rgb) == 8: rgb = rgb[2:]
                        if rgb and rgb != "00000000": color = f"#{rgb}"
                except Exception: pass
                parts.append(f"{css_side}:1px solid {color};")
    except Exception: pass
    return "".join(parts)

def _excel_col_width(ws, col_idx: int) -> float:
    best = None
    for dim in ws.column_dimensions.values():
        if dim.min is None or dim.max is None or dim.width is None: continue
        if dim.min <= col_idx <= dim.max:
            span = dim.max - dim.min
            if best is None or span < best[0]: best = (span, float(dim.width))
    if best: return best[1]
    return float(ws.sheet_format.defaultColWidth or 8.43)

def _excel_row_height_px(ws, row: int) -> int:
    """엑셀 행 높이 → px (인위적 늘리기 없음)."""
    h = ws.row_dimensions[row].height
    if h:
        return max(1, int(round(float(h) * 96 / 72)))
    return 20  # Excel 기본 ~15pt

def _wl_cell_font(cell, *, is_content: bool, is_client: bool, is_body_d: bool, is_date: bool) -> tuple[str, float]:
    """셀 font → (font-stack, pt). 본문 계열은 바탕체·엑셀 pt 그대로."""
    font = cell.font
    force_batang = is_content or is_client or is_body_d or is_date
    fname = (font.name or "").strip()
    if force_batang or fname in ("바탕", "바탕체", "바탕글", "Batang", "BatangChe"):
        stack = _WL_FONT_STACK
    elif fname in ("맑은 고딕", "Malgun Gothic"):
        stack = "'Malgun Gothic','Apple SD Gothic Neo',sans-serif"
    elif fname:
        stack = f"'{html.escape(fname)}',{_WL_FONT_STACK}"
    else:
        stack = _WL_FONT_STACK
    if font.size:
        fsize_pt = float(font.size)
    elif force_batang:
        fsize_pt = float(_WL_BODY_FONT_PT)
    else:
        fsize_pt = 11.0
    return stack, fsize_pt

def _excel_width_to_px(width: float) -> int:
    """엑셀 열 너비 → px (화면 100% 기준, 원본과 동일 체감)."""
    try: w = float(width)
    except (TypeError, ValueError): w = 8.43
    return max(10, int(w * 7 + 5))

def _worklog_sheet_pixel_size(path: str) -> tuple[int, int]:
    if load_workbook is None or not path or not os.path.exists(path): return 900, 1312
    try:
        wb = load_workbook(path, data_only=False)
        ws = wb.active
        total_w = sum(_excel_width_to_px(_excel_col_width(ws, c)) for c in range(WL_MIN_COL, WL_MAX_COL + 1))
        total_h = 0
        for r in range(WL_MIN_ROW, WL_MAX_ROW + 1):
            total_h += _excel_row_height_px(ws, r) + 1
        # 원본 로고(특이사항 하단, Excel 표시 높이 61 + 간격)
        total_h += 61 + 10
        n_pages = 1
        active_title = ws.title
        for name in wb.sheetnames:
            if name == active_title:
                continue
            if _extra_page_n_from_sheet_name(name):
                n_pages += 1
        if n_pages > 1:
            total_h = total_h * n_pages + 18 * (n_pages - 1)
        wb.close()
        return max(1, total_w), max(1, total_h)
    except Exception: return 900, 1312


def _excel_page_scale(path: str) -> float | None:
    """원본.xlsx pageSetup.scale(예: 69) → 0.69. 없으면 None."""
    if load_workbook is None or not path or not os.path.exists(path):
        return None
    try:
        wb = load_workbook(path, data_only=False)
        ws = wb.active
        sc = ws.page_setup.scale
        wb.close()
        if sc is None:
            return None
        v = float(sc)
        if v > 1.5:  # Excel stores percent (69)
            v = v / 100.0
        if 0.3 <= v <= 1.5:
            return v
    except Exception:
        return None
    return None


def _excel_page_margins_css(path: str) -> str:
    """원본 pageMargins → @page margin CSS (인치). 기본 0.55in 0.51in."""
    top = right = bottom = left = None
    if load_workbook is not None and path and os.path.exists(path):
        try:
            wb = load_workbook(path, data_only=False)
            ws = wb.active
            pm = ws.page_margins
            top, right, bottom, left = pm.top, pm.right, pm.bottom, pm.left
            wb.close()
        except Exception:
            pass
    # 원본.xlsx 기본값(인치)
    t = float(top if top is not None else 0.55)
    r = float(right if right is not None else 0.51)
    b = float(bottom if bottom is not None else 0.55)
    l = float(left if left is not None else 0.51)
    return f"{t:.2f}in {r:.2f}in {b:.2f}in {l:.2f}in"


def _excel_print_scale(path: str | None = None) -> float:
    """인쇄 배율 = 원본 Excel pageSetup.scale 그대로(추가 축소 없음)."""
    s = _excel_page_scale(path) if path else None
    if s is not None:
        return float(s)
    return 0.69  # 원본.xlsx 기본


def _scaled_view_frame_size(path: str, scale: float) -> tuple[int, int]:
    w, h = _worklog_sheet_pixel_size(path)
    s = float(scale) if scale and scale > 0 else 1.0
    return int(w * s) + 28, int(h * s) + 64

def _worklog_logo_bytes(xlsx_path: str | None = None) -> tuple[bytes, str, int, int] | None:
    """원본 템플릿 내장 로고(320×61) 우선. (bytes, mime, w, h)"""
    _here = os.path.dirname(os.path.abspath(__file__))
    if xlsx_path and os.path.exists(xlsx_path):
        try:
            import zipfile
            with zipfile.ZipFile(xlsx_path) as zf:
                for name in ("xl/media/image1.jpeg", "xl/media/image1.jpg", "xl/media/image1.png"):
                    if name in zf.namelist():
                        data = zf.read(name)
                        mime = "image/jpeg" if name.endswith((".jpeg", ".jpg")) else "image/png"
                        return data, mime, 320, 61  # 원본 Excel 표시 크기
        except Exception:
            pass
    for p, mime in (
        (os.path.join(WORKLOG_DIR, "shinil_logo.jpeg"), "image/jpeg"),
        (os.path.join(_here, "logo.png"), "image/png"),
        ("logo.png", "image/png"),
    ):
        if os.path.exists(p):
            try:
                with open(p, "rb") as f:
                    return f.read(), mime, 320, 61
            except Exception:
                continue
    return None


def _worklog_body_merge_bands() -> tuple[tuple[int, int, list[int]], ...]:
    """본문 한 행 = 거래처 C~F / 내용 G~X / 비고 Y~AB."""
    return (
        (WL_CLIENT_COL_START, WL_CLIENT_COL_END, WL_CLIENT_ROWS),
        (WL_CONTENT_COL_START, WL_CONTENT_COL_END, WL_CONTENT_ROWS),
        (WL_REMARK_COL_START, WL_REMARK_COL_END, WL_CONTENT_ROWS),
    )


def _ensure_worklog_body_merges(ws) -> None:
    """템플릿에 빠진 본문 병합(특히 비고 Y~AB)을 행마다 다시 붙인다."""
    if ws is None:
        return
    for c0, c1, rows in _worklog_body_merge_bands():
        for r in rows:
            ok = False
            overlap: list[str] = []
            for mr in list(ws.merged_cells.ranges):
                if mr.max_row < r or mr.min_row > r:
                    continue
                if mr.max_col < c0 or mr.min_col > c1:
                    continue
                if mr.min_row == r and mr.max_row == r and mr.min_col == c0 and mr.max_col == c1:
                    ok = True
                    break
                overlap.append(str(mr))
            if ok:
                continue
            for ref in overlap:
                try:
                    ws.unmerge_cells(ref)
                except Exception:
                    pass
            try:
                ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
            except Exception:
                pass


def _inject_body_col_merges(
    merge_map: dict[tuple[int, int], tuple[int, int]],
    skip: set[tuple[int, int]],
) -> None:
    """미리보기 HTML — 템플릿 병합이 없어도 본문 칸 폭을 원본과 같게 둔다."""
    for c0, c1, rows in _worklog_body_merge_bands():
        cs_need = c1 - c0 + 1
        for r in rows:
            if (r, c0) in skip:
                continue
            rs, cs = merge_map.get((r, c0), (1, 1))
            if cs < cs_need:
                merge_map[(r, c0)] = (max(1, int(rs)), cs_need)
            for c in range(c0 + 1, c1 + 1):
                skip.add((r, c))
                merge_map.pop((r, c), None)


def _worksheet_to_table_html(
    ws,
    path: str,
    *,
    include_logo: bool = True,
    layout_scale: float = 1.0,
    data_prefix: str = "",
) -> str:
    pfx = str(data_prefix or "")

    merge_map: dict[tuple[int, int], tuple[int, int]] = {}
    skip: set[tuple[int, int]] = set()
    for mr in ws.merged_cells.ranges:
        if mr.max_row < WL_MIN_ROW or mr.min_row > WL_MAX_ROW: continue
        if mr.max_col < WL_MIN_COL or mr.min_col > WL_MAX_COL: continue
        max_r, max_c = min(mr.max_row, WL_MAX_ROW), min(mr.max_col, WL_MAX_COL)
        min_r, min_c = max(mr.min_row, WL_MIN_ROW), max(mr.min_col, WL_MIN_COL)
        if min_r > max_r or min_c > max_c: continue
        rs, cs = max_r - min_r + 1, max_c - min_c + 1
        tl_r, tl_c = mr.min_row, mr.min_col
        if tl_r < WL_MIN_ROW or tl_c < WL_MIN_COL: tl_r, tl_c = min_r, min_c
        merge_map[(tl_r, tl_c)] = (rs, cs)
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                if (r, c) != (tl_r, tl_c): skip.add((r, c))
    _inject_body_col_merges(merge_map, skip)

    col_widths = []
    total_w = 0.0
    ls = float(layout_scale) if layout_scale and layout_scale > 0 else 1.0
    for c in range(WL_MIN_COL, WL_MAX_COL + 1):
        w = _excel_col_width(ws, c)
        px = max(1, int(round(_excel_width_to_px(w) * ls)))
        col_widths.append(px)
        total_w += px

    rows_html = []
    for r in range(WL_MIN_ROW, WL_MAX_ROW + 1):
        height_px = max(1, int(round(_excel_row_height_px(ws, r) * ls)))
        tds = []
        for c in range(WL_MIN_COL, WL_MAX_COL + 1):
            if (r, c) in skip: continue
            cell = ws.cell(r, c)
            rs, cs = merge_map.get((r, c), (1, 1))
            val = cell.value
            text = "" if isinstance(val, str) and val.startswith("=") else ("" if val is None else str(val))

            if c == 3 and r == WL_NEXT_ROWS[0] and not text.strip(): text = "익일업무"
            if c == 3 and r == WL_NOTE_ROWS[0] and not text.strip(): text = "특이사항"

            is_content = c == WL_CONTENT_COL_START and WL_CONTENT_ROWS[0] <= r <= WL_CONTENT_ROWS[-1]
            is_client = c == WL_CLIENT_COL_START and WL_CLIENT_ROWS[0] <= r <= WL_CLIENT_ROWS[-1]
            is_remark = c == WL_REMARK_COL_START and WL_CONTENT_ROWS[0] <= r <= WL_CONTENT_ROWS[-1]
            is_body_d = c == 4 and (r in WL_NEXT_ROWS or r in WL_NOTE_ROWS)
            is_date = (c == 3 and r == 5) or (str(cell.coordinate) == WL_DATE_CELL)
            font_stack, fsize_pt = _wl_cell_font(cell, is_content=is_content or is_remark, is_client=is_client, is_body_d=is_body_d, is_date=is_date)
            # 인쇄 시 Excel pageSetup.scale을 글자 pt·셀 크기에 미리 반영(브라우저 추가 축소 방지)
            if ls != 1.0:
                fsize_pt = round(float(fsize_pt) * ls, 2)
            font = cell.font
            bold = "bold" if font.bold else "normal"
            align = cell.alignment
            ha = align.horizontal or "left"
            va = align.vertical or "middle"
            if ha == "general": ha = "left"
            if va == "center": va = "middle"

            _text_clean = str(text).replace(" ", "").replace("\u3000", "").replace("\n", "") if text else ""
            is_side_label = c == 3 and (_text_clean in ("익일업무", "특이사항") or (r == WL_NEXT_ROWS[0] and rs >= 3) or (r == WL_NOTE_ROWS[0] and rs >= 3))
            is_vertical = is_side_label or getattr(align, "textRotation", 0) in (255, 90) or (_text_clean == "결재" and rs >= 2 and cs == 1)

            if is_vertical: ha, va = "center", "middle"
            elif is_client: ha, va = "center", "middle"  # 거래처 칸 항상 가운데
            elif is_body_d: ha, va = "left", "top"  # 익일/특이 — 빈 줄·다음 줄 위치 유지
            elif is_content or is_remark: ha = "left"
                
            fill = _cell_fill_color(cell) or "#FFFFFF"
            border = _border_css(cell)
            span = ""
            if rs > 1: span += f' rowspan="{rs}"'
            if cs > 1: span += f' colspan="{cs}"'
                
            if text in (_WL_SOFT_BLANK, "\u00a0"): text = ""
            elif is_content and text.strip() == "" and text != "": text = ""
                
            if is_vertical and text.strip():
                chars = [ch for ch in _text_clean if ch]
                esc = "<br>".join(html.escape(ch) for ch in chars)
            elif is_body_d and not str(text).strip():
                esc = "&nbsp;"  # 익일/특이 빈 행 — 인쇄 미리보기에서 줄 위치 유지
            else:
                esc = html.escape(text).replace(" ", "&nbsp;").replace("\n", "<br>")
                
            if is_content:
                # 원본.xlsx wrapText — 바탕체 14pt 원본 칸 폭에서 줄바꿈
                white, overflow, text_overflow, break_css = "pre-wrap", "hidden", "clip", "break-all"
                wrap_css = "overflow-wrap:anywhere;"
            elif is_client or is_remark:
                # 입력 1줄 = 미리보기 1줄. 비고는 칸 안에서 다시 쪼개지 않는다.
                white, overflow, text_overflow, break_css = "nowrap", "hidden", "clip", "keep-all"
                wrap_css = "overflow-wrap:normal;"
            elif is_vertical:
                white, overflow, text_overflow, break_css = "normal", "hidden", "clip", "keep-all"
                wrap_css = "overflow-wrap:normal;"
            else:
                white, overflow, text_overflow, break_css = "pre-wrap", "hidden", "clip", "break-word"
                wrap_css = "overflow-wrap:anywhere;"
                
            c0 = c - WL_MIN_COL
            span_w = sum(col_widths[c0 : c0 + max(cs, 1)]) if c0 >= 0 else 0
            width_css = f"width:{span_w}px;min-width:{span_w}px;max-width:{span_w}px;" if span_w else ""

            if is_vertical:
                pad_css = "padding:4px 1px;"
            else:
                pad_css = "padding:0;"
            
            if r < WL_CLIENT_ROWS[0] and not is_vertical:
                style = (
                    f"box-sizing:border-box;{width_css}"
                    f"font-family:{font_stack};font-size:{fsize_pt}pt;font-weight:{bold};"
                    f"text-align:center; vertical-align:middle;"
                    f"background:{fill}; {border} {pad_css}"
                )
            else:
                line_css = "line-height:1.35;" if is_vertical else "line-height:1.25;"
                
                style = (
                    f"box-sizing:border-box;{width_css}"
                    f"font-family:{font_stack};font-size:{fsize_pt}pt;font-weight:{bold};"
                    f"text-align:{ha};vertical-align:{va};"
                    f"background:{fill};{border}"
                    f"{pad_css}white-space:{white};overflow:{overflow};"
                    f"text-overflow:{text_overflow};word-break:{break_css};"
                    f"{wrap_css}"
                    f"{line_css}"
                )
            wl_key = ""
            if is_date:
                wl_key = "date"
            elif is_content:
                wl_key = f"G{r}"
            elif is_client:
                wl_key = f"C{r}"
            elif is_remark:
                wl_key = f"Y{r}"
            elif is_body_d:
                wl_key = f"D{r}"
            data_attr = f' data-wl="{pfx}{wl_key}"' if wl_key else ""
            tds.append(f'<td{span}{data_attr} style="{style}">{esc}</td>')
        rows_html.append(f'<tr style="height:{height_px}px;box-sizing:border-box;">{"".join(tds)}</tr>')

    colgroup = "".join(f'<col style="width:{w}px">' for w in col_widths)

    import base64
    logo_tr = ""
    # 원본: 특이사항(47행) 바로 아래 중앙, Excel 표시 크기 320×61
    if include_logo:
        logo = _worklog_logo_bytes(path)
        if logo:
            try:
                data, mime, nat_w, nat_h = logo
                b64_img = base64.b64encode(data).decode("utf-8")
                ls = float(layout_scale) if layout_scale and layout_scale > 0 else 1.0
                logo_w = max(1, int(round(nat_w * ls)))
                logo_h = max(1, int(round(nat_h * ls)))
                pad_top = max(4, int(round(8 * ls)))  # 원본과 비슷한 위쪽 간격
                logo_tr = (
                    f'<tr style="border:none;background:#fff;height:{logo_h + pad_top + 2}px;">'
                    f'<td colspan="26" style="text-align:center;padding:{pad_top}px 0 0 0;border:none;vertical-align:top;">'
                    f'<img src="data:{mime};base64,{b64_img}" alt="신일가스 로고" '
                    f'width="{logo_w}" height="{logo_h}" '
                    f'style="width:{logo_w}px;height:{logo_h}px;max-width:none;display:block;margin:0 auto;">'
                    "</td></tr>"
                )
            except Exception:
                pass

    page_n = 1
    m_p = re.match(r"p(\d+)-$", pfx)
    if m_p:
        page_n = max(2, int(m_p.group(1)))
    return (
        f'<div class="wl-preview-page" data-wl-page="{page_n}">'
        f'<table class="wl-sheet" style="border-collapse:collapse;table-layout:fixed;width:{int(total_w)}px;background:#fff;box-sizing:border-box;">'
        f"<colgroup>{colgroup}</colgroup>"
        f"<tbody>{''.join(rows_html)}{logo_tr}</tbody>"
        f"</table></div>"
    )


def _worklog_html_sheet_pages(wb) -> list[tuple[object, str]]:
    pages: list[tuple[object, str]] = [(wb.active, "")]
    active_title = wb.active.title if wb.worksheets else ""
    extras: list[tuple[int, str]] = []
    for name in wb.sheetnames:
        if name == active_title:
            continue
        n = _extra_page_n_from_sheet_name(name)
        if n:
            extras.append((n, name))
    extras.sort()
    for n, name in extras:
        pages.append((wb[name], f"p{n}-"))
    return pages


def workbook_to_html(path: str, *, include_logo: bool = True, layout_scale: float = 1.0) -> str:
    if load_workbook is None:
        return "<p>openpyxl 필요</p>"
    wb = load_workbook(path, data_only=False)
    try:
        parts: list[str] = []
        for i, (ws, prefix) in enumerate(_worklog_html_sheet_pages(wb)):
            if i:
                parts.append(_html_page_break())
            parts.append(
                _worksheet_to_table_html(
                    ws, path, include_logo=include_logo, layout_scale=layout_scale, data_prefix=prefix,
                )
            )
        return "".join(parts)
    finally:
        wb.close()

def _a4_print_fit(raw_w: int, raw_h: int, *, path: str | None = None) -> float:
    """인쇄 배율 = 원본 Excel pageSetup.scale 그대로(추가 맞춤 축소 없음)."""
    return _excel_print_scale(path)

def render_worklog_view_html(path: str, *, print_mode: bool = False, scale: float | None = None, auto_print: bool = False, wrap_height: str | None = None) -> str:
    # 인쇄: Excel 배율(69%)을 HTML 글씨·셀에 직접 반영 → CSS transform으로 또 줄이지 않음
    excel_print_s = _excel_print_scale(path)
    layout_s = float(excel_print_s) if print_mode else 1.0
    sheet = workbook_to_html(path, include_logo=True, layout_scale=layout_s)
    raw_w, raw_h = _worklog_sheet_pixel_size(path)
    print_fit = excel_print_s
    # print_mode에서는 이미 layout_scale 적용된 크기
    if print_mode:
        scaled_w = max(1, int(round(raw_w * layout_s)))
        scaled_h = max(1, int(round(raw_h * layout_s)))
        page_margins = _excel_page_margins_css(path)
        pct = int(round(excel_print_s * 100))
    else:
        scaled_w, scaled_h = max(1, int(round(raw_w * print_fit))), max(1, int(round(raw_h * print_fit)))
        page_margins = "8mm"
        pct = int(round(excel_print_s * 100))

    toolbar = ""
    if print_mode:
        view_scale = 1.0
        toolbar = f"""<div class="toolbar no-print"><button type="button" id="wl-print-btn">인쇄하기</button><span class="hint">원본 Excel 인쇄 설정 그대로입니다 (배율 {pct}%, 여백 동일). 브라우저 인쇄창에서 「용지에 맞춤」을 끄세요.</span></div>"""
    else:
        view_scale = float(scale if scale is not None else _WL_PREVIEW_SCALE)
    frame_w, frame_h = _scaled_view_frame_size(path, view_scale)
    if print_mode:
        # 이미 69%로 렌더됨 → 추가 zoom/transform 없음
        scale_css = f"zoom:1;width:{scaled_w}px;max-width:none;"
        scale_css_fallback = f"transform:none;width:{scaled_w}px;"
        wrap_h, wrap_w, wrap_overflow, body_overflow, body_h = f"{scaled_h}px", f"{scaled_w}px", "hidden", "auto", "auto"
    elif view_scale >= 1:
        scale_css = "zoom:1;width:fit-content;"
        scale_css_fallback = "transform:scale(1);transform-origin:top left;width:fit-content;"
        wrap_h, wrap_w, wrap_overflow, body_overflow, body_h = "auto", "100%", "visible", "visible", "auto"
    else:
        s = float(view_scale)
        scale_css = f"zoom:{s};width:fit-content;"
        scale_css_fallback = f"transform:scale({s});transform-origin:top left;margin-bottom:{(s - 1) * raw_h:.1f}px;width:fit-content;"
        wrap_h, wrap_w, wrap_overflow, body_overflow, body_h = "auto", f"{frame_w}px", "visible", "hidden", f"{frame_h}px"
    if wrap_height is not None: wrap_h = wrap_height

    fit_print_js = f"""
        var wlZoom = 1;
        var wlOrigFit = 1;
        function wlApplyZoom(s) {{
          var sheet = document.querySelector('.sheet-scale');
          var wrap = document.querySelector('.wrap');
          var table = document.querySelector('.wl-sheet');
          if (!sheet || !wrap || !table) return;
          wlZoom = 1;
          sheet.style.transform = 'none';
          sheet.style.zoom = '1';
          wrap.style.maxWidth = 'none';
          wrap.style.width = '{scaled_w}px';
          wrap.style.height = '{scaled_h}px';
          wrap.style.overflow = 'hidden';
          wrap.style.margin = '0 auto';
        }}
        function wlFitToA4() {{ wlApplyZoom(1); }}
        try {{ window.wlFitToA4 = wlFitToA4; }} catch (e0) {{}}
    """
    go_print_js = """
        function goPrint() {
          try { wlFitToA4(); } catch (e0) {}
          var fire = function() {
            setTimeout(function() {
              try { window.focus(); window.print(); } catch (e) {}
            }, 150);
          };
          if (document.fonts && document.fonts.ready) {
            document.fonts.ready.then(fire).catch(fire);
          } else {
            setTimeout(fire, 400);
          }
        }
    """
    auto_script = f"""<script>(function() {{ {fit_print_js} {go_print_js} var btn = document.getElementById('wl-print-btn'); if (btn) btn.addEventListener('click', function(ev) {{ ev.preventDefault(); goPrint(); }}); window.addEventListener('beforeprint', function() {{ try {{ wlFitToA4(); }} catch (e2) {{}} }}); function boot() {{ try {{ wlFitToA4(); }} catch (e3) {{}} {"setTimeout(goPrint, 500);" if auto_print else ""} }} if (document.readyState === 'complete') setTimeout(boot, 250); else window.addEventListener('load', function() {{ setTimeout(boot, 250); }}); }})();</script>""" if print_mode else ""
    fallback_block = f"@supports not (zoom: 1) {{ .sheet-scale {{ {scale_css_fallback} }} }}" if scale_css_fallback else ""
    if print_mode:
        print_media = f"""@media print {{ html, body {{ overflow:visible !important; height:auto !important; width:auto !important; margin:0 !important; padding:0 !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }} .no-print, .toolbar {{ display:none !important; }} .wrap {{ overflow:visible !important; max-width:none !important; width:{scaled_w}px !important; height:auto !important; border:none !important; margin:0 auto !important; padding:0 !important; }} .sheet-scale {{ zoom:1 !important; transform:none !important; width:{scaled_w}px !important; margin:0 !important; }} .wl-sheet, .wl-sheet td, .wl-sheet tr {{ font-family:{_WL_FONT_STACK} !important; -webkit-print-color-adjust:exact; print-color-adjust:exact; }} .wl-page-break {{ page-break-before:always !important; break-before:page !important; height:0 !important; min-height:0 !important; }} }}"""
    else:
        print_media = "@media print { html, body { overflow:visible !important; } }"
    
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>일일업무일지</title><meta name="viewport" content="width=device-width, initial-scale=1"><link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Nanum+Myeongjo:wght@400;700&display=swap" rel="stylesheet"><style>{_WL_FONT_FACE_CSS} @page {{ size: A4 portrait; margin: {page_margins}; }} html, body {{ margin:0; padding:0; background:#fff; overflow:{body_overflow} !important; height:{body_h}; }} body {{ padding:{"6px" if not print_mode else "0"}; box-sizing:border-box; font-family:{_WL_FONT_STACK} !important; }} .toolbar {{ margin-bottom:10px; display:flex; gap:10px; align-items:center; flex-wrap:wrap; }} .toolbar button {{ padding:8px 14px; font-size:14px; border:1px solid #334155; border-radius:6px; background:#1E293B; color:#fff; cursor:pointer; }} .toolbar button.secondary {{ background:#F8FAFC; color:#334155; border-color:#CBD5E1; cursor:default; }} .toolbar .hint {{ font:12px/1.45 sans-serif; color:#64748B; max-width:42rem; }} .wrap {{ overflow:{wrap_overflow} !important; height:{wrap_h}; width:{wrap_w}; max-width:{"none" if print_mode else "100%"}; border:{"none" if print_mode else "1px solid #94A3B8"}; background:#fff; box-sizing:border-box; padding:0; }} .sheet-scale {{ {scale_css} }} .wl-sheet {{ border-collapse:collapse; table-layout:fixed; font-family:{_WL_FONT_STACK} !important; }} .wl-sheet, .wl-sheet td, .wl-sheet tr {{ box-sizing:border-box; font-family:{_WL_FONT_STACK} !important; }} .wl-sheet td[data-wl^="G"] {{ white-space:pre-wrap !important; overflow:hidden !important; word-break:break-all !important; overflow-wrap:anywhere !important; }} .wl-sheet td[data-wl^="Y"], .wl-sheet td[data-wl^="C"] {{ white-space:nowrap !important; overflow:hidden !important; word-break:keep-all !important; overflow-wrap:normal !important; }} .wl-page-break {{ page-break-before:always; break-before:page; }} {fallback_block} {print_media}</style></head><body>{toolbar}<div class="wrap"><div class="sheet-scale">{sheet}</div></div>{auto_script}</body></html>"""

def _entry_blank_after(ent: dict | None, default: int = 1) -> int:
    try: n = int((ent or {}).get("blank_after", default))
    except (TypeError, ValueError): n = default
    return max(0, min(10, n))

_WL_SOFT_BLANK = " "

def _grouped_entries_from_cells(cells: dict) -> list[dict]:
    entries: list[dict] = []
    blank_run = 0
    for r in WL_CLIENT_ROWS:
        raw_c, raw_g = str(cells.get(f"C{r}", "") or ""), str(cells.get(f"G{r}", "") or "")
        raw_y = str(cells.get(f"Y{r}", "") or "")
        client_raw = _scrub_dummy_label(raw_c)
        client_stripped = client_raw.strip()
        remark_raw = _scrub_dummy_label(raw_y)
        remark_stripped = remark_raw.strip()
        soft_blank = raw_g == _WL_SOFT_BLANK or raw_g == "\u00a0"
        g_whitespace_only = (not client_stripped and not remark_stripped and not soft_blank and raw_g != "" and raw_g.strip() == "")
        fully_empty = (not client_stripped and not remark_stripped and not soft_blank and not g_whitespace_only and raw_c.strip() == "" and raw_g.strip() == "" and raw_y.strip() == "")
        content = "" if (soft_blank or g_whitespace_only) else _scrub_dummy_label(raw_g)

        def _finish_row(ent: dict, c_val: str, g_val: str, y_val: str) -> None:
            ent.setdefault("client_lines", []).append(c_val)
            ent.setdefault("lines", []).append(g_val)
            ent.setdefault("remark_lines", []).append(y_val)
            ent["client"] = "\n".join(ent.get("client_lines") or [])
            ent["content"] = "\n".join(ent.get("lines") or [])
            ent["remarks"] = "\n".join(ent.get("remark_lines") or [])

        def _new_entry(c_val: str, g_val: str, y_val: str) -> dict:
            return {
                "client": c_val,
                "client_lines": [c_val],
                "content": g_val,
                "lines": [g_val],
                "remarks": y_val,
                "remark_lines": [y_val],
                "blank_after": 1,
            }
        
        if fully_empty: 
            blank_run += 1
            continue

        start_new = (not entries) or blank_run > 0
        if client_stripped and entries and blank_run == 0:
            prev_c = str((entries[-1].get("client_lines") or [""])[-1] or "")
            if _display_units(prev_c) < _client_line_units():
                start_new = True
        if start_new and not (
            g_whitespace_only or (soft_blank and not client_stripped and not remark_stripped)
        ):
            if entries:
                entries[-1]["blank_after"] = max(0, min(10, blank_run))
            blank_run = 0
            entries.append(_new_entry(client_raw, content or "", remark_raw))
            continue

        if g_whitespace_only or (soft_blank and not client_stripped and not remark_stripped):
            if not entries or blank_run > 0:
                if entries: entries[-1]["blank_after"] = max(0, min(10, blank_run))
                blank_run = 0
                entries.append(_new_entry(client_raw, "", remark_raw))
            else:
                blank_run = 0
                _finish_row(entries[-1], client_raw, "", remark_raw)
            continue
            
        if not entries or blank_run > 0:
            if entries: entries[-1]["blank_after"] = max(0, min(10, blank_run))
            blank_run = 0
            entries.append(_new_entry(client_raw, content or "", remark_raw))
        else:
            _finish_row(entries[-1], client_raw, content or "", remark_raw)
    return entries

def _entry_client_lines(ent: dict | None) -> list[str]:
    if not ent: return []
    raw = ent.get("client_lines")
    if isinstance(raw, list) and raw: src = [str(x or "") for x in raw]
    else: 
        raw_c = str(ent.get("client") or "")
        src = raw_c.split('\n') if raw_c else []
    max_u, out = _client_line_units(), []
    for line in src:
        if not str(line or "").strip(): 
            out.append(str(line or ""))
            continue
        out.extend(_chunk_text(str(line or ""), max_u) or [str(line or "")])
    return out

def _entry_pack_lines(ent: dict) -> list[str]:
    max_u = _content_line_units()
    raw = ent.get("lines")
    src = [str(x or "") for x in raw] if isinstance(raw, list) else _chunk_text(str(ent.get("content") or ""), max_u) or []
    out: list[str] = []
    for line in src:
        if not str(line or "").strip(): 
            out.append(str(line or ""))
            continue
        out.extend(_chunk_text(line, max_u) or [line])
    if not out and (_entry_client_lines(ent) or _entry_remark_lines(ent)): out = [""]
    return out

def _entry_remark_lines(ent: dict | None) -> list[str]:
    if not ent: return []
    raw = ent.get("remark_lines")
    if isinstance(raw, list) and raw: src = [str(x or "") for x in raw]
    else:
        raw_r = str(ent.get("remarks") or "")
        src = raw_r.split("\n") if raw_r else []
    max_u, out = _remark_line_units(), []
    for line in src:
        if not str(line or "").strip():
            out.append(str(line or ""))
            continue
        out.extend(_chunk_text(str(line or ""), max_u) or [str(line or "")])
    return out

def _content_row_usage(entries: list[dict] | None) -> dict:
    ents = list(entries or [])
    if len(ents) == 1:
        e = ents[0]
        cl = e.get("client_lines") if isinstance(e.get("client_lines"), list) else []
        ln = e.get("lines") if isinstance(e.get("lines"), list) else []
        rm = e.get("remark_lines") if isinstance(e.get("remark_lines"), list) else []
        if len(cl) >= WL_SHEET_N or len(ln) >= WL_SHEET_N or len(rm) >= WL_SHEET_N:
            return _sheet_row_usage(cl, ln, rm)
    total, used, wrote_any, prev_gap, per_entry = len(WL_CONTENT_ROWS), 0, False, 0, []
    for ent in ents:
        clients, pack_lines, remarks = _entry_client_lines(ent), _entry_pack_lines(ent), _entry_remark_lines(ent)
        if not any(str(x).strip() for x in clients) and not any((x or "").strip() for x in pack_lines) and not any(str(x).strip() for x in remarks): per_entry.append(0); continue
        gap = prev_gap if wrote_any else 0
        lines = max(len(clients), len(pack_lines), len(remarks), 1)
        need = gap + lines
        per_entry.append(need)
        used += need; wrote_any = True; prev_gap = _entry_blank_after(ent, 1)
    return {"total": total, "used": used, "remaining": max(0, total - used), "per_entry": per_entry, "last_row": WL_CONTENT_ROWS[-1] if WL_CONTENT_ROWS else 41, "next_row": WL_CONTENT_ROWS[used] if used < total else None, "overflow": used > total}

def _render_row_remain_gauge(usage: dict, *, height_px: int = 980) -> None:
    total, used, rem, overflow = max(1, int(usage.get("total") or 1)), max(0, int(usage.get("used") or 0)), max(0, int(usage.get("remaining") or 0)), bool(usage.get("overflow"))
    if overflow: accent, used_color, rem_color, label, big, rem_show, used_show = "#DC2626", "#FECACA", "#FEE2E2", "초과", f"+{used - total}", 0, total
    elif rem <= 3: accent, used_color, rem_color, label, big, rem_show, used_show = "#D97706", "#E2E8F0", "#FBBF24", "남음", str(rem), rem, min(used, total)
    else: accent, used_color, rem_color, label, big, rem_show, used_show = "#0F766E", "#E2E8F0", "#14B8A6", "남음", str(rem), rem, min(used, total)
    segs = [f'<div style="flex:1 1 0;min-height:3px;border-radius:3px;background:{used_color if i < used_show else rem_color};margin:0 0 {max(1, int(2 if total <= 20 else 1))}px 0;opacity:{0.55 if i < used_show else 1};"></div>' for i in range(total)]
    if segs: segs[-1] = segs[-1].replace(f"margin:0 0 {max(1, int(2 if total <= 20 else 1))}px 0;", "margin:0;")
    next_txt = f"다음 G{usage.get('next_row')}" if usage.get("next_row") else "칸 끝"
    st.markdown(f"""<div style="display:flex;flex-direction:column;align-items:center;justify-content:flex-start;gap:6px;padding:2px 2px 0;min-height:{height_px}px;height:{height_px}px;font-family:'Pretendard','Apple SD Gothic Neo',sans-serif;"><div style="text-align:center;line-height:1.1;flex:0 0 auto;"><div style="font-size:10px;font-weight:700;letter-spacing:.06em;color:#64748B;">{label}</div><div style="font-size:22px;font-weight:800;color:{accent};margin-top:1px;">{big}</div><div style="font-size:10px;color:#94A3B8;">칸</div></div><div style="display:flex;flex-direction:column;justify-content:flex-start;flex:1 1 auto;width:26px;min-height:{max(520, height_px - 120)}px;height:{max(520, height_px - 120)}px;padding:4px 3px;border-radius:10px;background:#F8FAFC;box-shadow:inset 0 0 0 1px #E2E8F0;" title="전체 {total}칸 · 사용 {used}칸 · 남음 {rem}칸">{"".join(segs)}</div><div style="text-align:center;line-height:1.25;flex:0 0 auto;"><div style="font-size:11px;font-weight:800;color:#0F172A;"><span style="color:{accent};">{rem_show}</span><span style="color:#94A3B8;font-weight:600;"> / {total}</span></div><div style="font-size:10px;color:#64748B;">남은칸 / 전체칸</div><div style="font-size:10px;color:#94A3B8;margin-top:2px;">사용 {used_show}칸 · {next_txt}</div></div></div>""", unsafe_allow_html=True)

def _textarea_lines(raw: str) -> list[str]:
    """익일업무·특이사항 textarea — 선행·중간 빈 줄 위치 유지."""
    if raw is None:
        return []
    return str(raw).splitlines()


def _panel_lines_from_cells(cells: dict, rows: list[int], *, col: str = "D") -> list[str]:
    """엑셀 행 → textarea 줄. 맨 끝 빈 행만 제거. soft blank는 빈 줄로."""
    lines: list[str] = []
    for r in rows:
        raw = str(cells.get(f"{col}{r}") or "")
        if raw in (_WL_SOFT_BLANK, "\u00a0"):
            lines.append("")
        else:
            lines.append(raw)
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def _panel_lines_to_cells(lines: list[str] | None, max_u: int) -> list[str]:
    """textarea 줄 → 엑셀 D열 행 목록. 빈 줄은 NBSP로 유지(빈 문자열이 None으로 저장되는 것 방지)."""
    chunks: list[str] = []
    for t in lines or []:
        t = str(t or "")
        if t.strip() == "":
            chunks.append("\u00a0")
        else:
            chunks.extend(_chunk_text(t, max_u) or [t])
    return chunks


def _pack_entries_to_cells(d: date, entries: list[dict], next_day: list[str] | None = None, notes: list[str] | None = None) -> dict:
    cells, max_u, row_i, rows, wrote_any, prev_gap = _empty_cells(d), _content_line_units(), 0, WL_CONTENT_ROWS, False, 0
    for ent in entries or []:
        clients, chunks, remarks = _entry_client_lines(ent), _entry_pack_lines(ent), _entry_remark_lines(ent)
        if not any(str(x).strip() for x in clients) and not any((x or "").strip() for x in chunks) and not any(str(x).strip() for x in remarks): continue
        if wrote_any: row_i += max(0, int(prev_gap))
        for j in range(max(len(clients), len(chunks), len(remarks), 1)):
            if row_i >= len(rows): break
            r = rows[row_i]
            c_val = clients[j] if j < len(clients) else ""
            g_val = chunks[j] if j < len(chunks) else ""
            y_val = remarks[j] if j < len(remarks) else ""
            
            cells[f"C{r}"] = c_val
            if (j < len(chunks) and not str(chunks[j]).strip()) or (j < len(clients) and not str(clients[j]).strip() and j >= len(chunks) and not str(y_val).strip()):
                cells[f"G{r}"] = _WL_SOFT_BLANK
            else:
                cells[f"G{r}"] = g_val
            cells[f"Y{r}"] = y_val
            row_i += 1
        wrote_any, prev_gap = True, _entry_blank_after(ent, 1)
    for t_list, t_rows in ((next_day, WL_NEXT_ROWS), (notes, WL_NOTE_ROWS)):
        chunks = _panel_lines_to_cells(t_list, max_u)
        for i, r in enumerate(t_rows): cells[f"D{r}"] = chunks[i] if i < len(chunks) else ""
    return cells


def _flatten_to_sheet_entry(d: date, entries_list: list[dict] | None) -> dict:
    """여러 항목을 본문 32칸 한 장으로 편다. 익일·특이는 넣지 않는다."""
    ents = list(entries_list or [])
    if len(ents) == 1:
        e = ents[0]
        cl = e.get("client_lines") if isinstance(e.get("client_lines"), list) else []
        ln = e.get("lines") if isinstance(e.get("lines"), list) else []
        rm = e.get("remark_lines") if isinstance(e.get("remark_lines"), list) else []
        if len(cl) >= WL_SHEET_N or len(ln) >= WL_SHEET_N or len(rm) >= WL_SHEET_N:
            return _sheet_entry_from_lines(cl, ln, rm)
    return _sheet_entry_from_cells(_pack_entries_to_cells(d, ents))

def _entries_from_cells(cells: dict) -> tuple[list[tuple[str, str]], list[str], list[str]]:
    rows = [_summary_row_from_entry(e) for e in _grouped_entries_from_cells(cells)]
    next_day = _panel_lines_from_cells(cells, WL_NEXT_ROWS)
    notes = _panel_lines_from_cells(cells, WL_NOTE_ROWS)
    return rows, next_day, notes


def _summary_row_from_entry(ent: dict) -> tuple[str, str]:
    """요약/미리보기용 — 거래처 줄은 이어 붙이고, 내용은 빈줄만 뺀다."""
    client = " ".join(ln.strip() for ln in _entry_client_lines(ent) if str(ln).strip())
    content_lines = [ln for ln in _entry_pack_lines(ent) if str(ln).strip()]
    return client, "\n".join(content_lines)

_WL_SUMMARY_PREVIEW_CSS = (
    "@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');"
    " .wl-sum-preview { font-family:'Pretendard','Apple SD Gothic Neo',sans-serif; color:#0F172A; }"
    " .wl-sum-preview .card { background:linear-gradient(180deg,#F8FAFC 0%,#FFFFFF 48px); border:1px solid #E2E8F0;"
    " border-radius:14px; box-shadow:0 1px 2px rgba(15,23,42,.04); overflow:hidden; }"
    " .wl-sum-preview .head { padding:16px 18px 14px; border-bottom:1px solid #E2E8F0;"
    " background:linear-gradient(135deg,#0F766E 0%,#0E7490 55%,#0369A1 100%); color:#fff; }"
    " .wl-sum-preview .head .title { font-size:19px; font-weight:750; letter-spacing:-.02em; }"
    " .wl-sum-preview .head .sub { margin-top:4px; font-size:14px; opacity:.94; }"
    " .wl-sum-preview .sec { padding:14px 16px 8px; }"
    " .wl-sum-preview .sec h3 { margin:0 0 10px; font-size:12px; font-weight:700; letter-spacing:.06em;"
    " color:#64748B; text-transform:uppercase; }"
    " .wl-sum-preview .item { display:flex; gap:12px; align-items:flex-start; padding:12px; margin-bottom:8px;"
    " background:#fff; border:1px solid #E2E8F0; border-radius:10px; }"
    " .wl-sum-preview .item .body { flex:1; min-width:0; display:flex; flex-direction:column; gap:0; }"
    " .wl-sum-preview .idx { flex:0 0 28px; height:28px; border-radius:8px; background:#CCFBF1; color:#0F766E;"
    " font-weight:700; font-size:13px; display:flex; align-items:center; justify-content:center; }"
    " .wl-sum-preview .client { font-size:16px; font-weight:700; color:#134E4A; margin:0 0 2px 0; padding:0; white-space:pre-wrap; line-height:1.4; }"
    " .wl-sum-preview .content { font-size:15px; line-height:1.55; color:#1E293B; white-space:pre-wrap; word-break:break-word; margin:0; padding:0; }"
    " .wl-sum-preview .muted { color:#94A3B8; font-weight:500; }"
    " .wl-sum-preview .empty { padding:18px; text-align:center; color:#94A3B8; font-size:13px;"
    " border:1px dashed #CBD5E1; border-radius:10px; background:#F8FAFC; }"
    " .wl-sum-preview .panel { margin:0 16px 14px; padding:12px 14px; border-radius:12px; border:1px solid #E2E8F0; background:#fff; }"
    " .wl-sum-preview .panel.next { border-left:4px solid #2563EB; }"
    " .wl-sum-preview .panel.note { border-left:4px solid #D97706; }"
    " .wl-sum-preview .panel h3 { margin:0 0 8px; font-size:13px; font-weight:700; color:#1E293B; }"
    " .wl-sum-preview .line { display:flex; gap:8px; align-items:flex-start; padding:7px 0; font-size:15px;"
    " line-height:1.55; color:#1E293B; border-bottom:1px solid #F1F5F9; }"
    " .wl-sum-preview .line:last-child { border-bottom:none; }"
    " .wl-sum-preview .dot { width:7px; height:7px; margin-top:7px; border-radius:50%; background:#94A3B8; flex:0 0 auto; }"
    " .wl-sum-preview .panel.next .dot { background:#2563EB; }"
    " .wl-sum-preview .panel.note .dot { background:#D97706; }"
    " .wl-sum-preview .foot { padding:10px 16px 14px; font-size:11px; color:#94A3B8; }"
)

def render_readable_preview_html(d: date, cells: dict) -> str:
    body, extras = _detach_extra_pages(cells)
    rows, next_day, notes = _entries_from_cells(body)
    for extra in extras:
        extra_rows, _, _ = _entries_from_cells(extra)
        rows.extend(extra_rows)
    date_label = html.escape(body.get("date") or format_worklog_date(d))
    if rows:
        work_items = []
        for i, (client, content) in enumerate(rows, 1):
            client_show = str(client or "").strip()
            content_show = str(content or "").strip()
            c = html.escape(client_show) if client_show else "<span class='muted'>(거래처 없음)</span>"
            t = html.escape(content_show).replace("\n", "<br>") if content_show else "<span class='muted'>—</span>"
            work_items.append(f"""<div class="item"><div class="idx">{i}</div><div class="body"><div class="client">{c}</div><div class="content">{t}</div></div></div>""")
        work_html = "".join(work_items)
    else: work_html = "<div class='empty'>등록된 업무 내용이 없습니다.</div>"
    def _lines(items: list[str], empty_msg: str) -> str:
        if not items:
            return f"<div class='empty'>{empty_msg}</div>"
        out = []
        for x in items:
            if not str(x).strip():
                out.append("<div class='line'><span class='dot' style='visibility:hidden'></span><span>&nbsp;</span></div>")
            else:
                out.append(f"<div class='line'><span class='dot'></span><span>{html.escape(x)}</span></div>")
        return "".join(out)
    _iframe_css = _WL_SUMMARY_PREVIEW_CSS.replace(".wl-sum-preview ", "") + " html, body { margin:0; padding:0; background:transparent; } body { padding:4px; }"
    return (
        f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>{_iframe_css}</style></head><body>'
        f'<div class="card"><div class="head"><div class="title">일일업무일지</div><div class="sub">{date_label}</div></div>'
        f'<div class="sec"><h3>거래처 · 내용</h3>{work_html}</div>'
        f'<div class="panel next"><h3>익일업무</h3>{_lines(next_day, "익일업무 없음")}</div>'
        f'<div class="panel note"><h3>특 이 사 항</h3>{_lines(notes, "특이사항 없음")}</div>'
        f'<div class="foot">인쇄는 상단 「인쇄창열기」를 사용하세요.</div></div></body></html>'
    )

def _entries_key(d: date) -> str: return f"wl_entries_{d.isoformat()}"
def _next_key(d: date) -> str: return f"wl_next_{d.isoformat()}"
def _notes_key(d: date) -> str: return f"wl_notes_{d.isoformat()}"
def _boot_key(d: date) -> str: return f"worklog_booted_{d.isoformat()}"

def _stored_cells_memo_key(d: date) -> str:
    path = worklog_path(d)
    try:
        arch = worklog_archive_month_path(d, create_year=False) or ""
    except Exception:
        arch = ""
    try:
        pm = os.path.getmtime(path) if os.path.isfile(path) else 0.0
    except OSError:
        pm = 0.0
    try:
        am = os.path.getmtime(arch) if arch and os.path.isfile(arch) else 0.0
    except OSError:
        am = 0.0
    return f"{d.isoformat()}|{pm}|{am}"


def _stored_cells_for_date(d: date) -> dict:
    """로컬 파일이 비어 있으면 월별 저장본을 쓴다. 빈 잔여 파일이 내용을 가리지 않게."""
    memo = st.session_state.setdefault("wl_stored_cells_memo", {})
    mk = _stored_cells_memo_key(d)
    hit = memo.get(mk) if isinstance(memo, dict) else None
    if isinstance(hit, dict):
        return dict(hit)
    try:
        packed = _attach_extra_pages(read_worklog_cells(d), read_worklog_extra_page_cells(d))
        if _worklog_cells_have_draft(packed):
            if isinstance(memo, dict):
                memo[mk] = packed
            return packed
    except Exception:
        packed = None
    try:
        arch = read_worklog_cells_from_archive(d)
        if arch is not None:
            arch_p = _attach_extra_pages(arch, _read_extra_pages_from_archive(d))
            if _worklog_cells_have_draft(arch_p):
                if isinstance(memo, dict):
                    memo[mk] = arch_p
                return arch_p
    except Exception:
        pass
    out = packed or _empty_cells(d)
    if isinstance(memo, dict):
        memo[mk] = out
    return out


def _init_widget_state(d: date) -> dict:
    bk, ek = _boot_key(d), _entries_key(d)
    if st.session_state.get(bk) and ek in st.session_state: return {}
    packed = _stored_cells_for_date(d)
    cells, extras = _detach_extra_pages(packed)
    entries = [_sheet_entry_from_cells(cells)]
    st.session_state[ek] = entries
    iso = d.isoformat()
    st.session_state[f"wl_entry_count_{iso}"] = 1
    st.session_state[_page_count_key(iso)] = max(1, min(WL_MAX_PAGES, 1 + len(extras)))
    st.session_state[_page_idx_key(iso)] = 0
    _, next_day, notes = _entries_from_cells(cells)
    st.session_state[_next_key(d)] = "\n".join(next_day)
    st.session_state[_notes_key(d)] = "\n".join(notes)
    sheet = entries[0] if entries else _empty_sheet_entry()
    _seed_entry_clients(iso, 0, sheet.get("client_lines") or [""])
    _apply_entry_lines(iso, 0, [str(x or "") for x in (sheet.get("lines") or [])], remount_comp=True)
    _seed_entry_remarks(iso, 0, sheet.get("remark_lines") or [""])
    _snapshot_worklog_page(iso, 0)
    for i, extra in enumerate(extras, start=1):
        ent = _sheet_entry_from_cells(extra)
        _seed_entry_clients(iso, i, ent.get("client_lines") or [""])
        _apply_entry_lines(iso, i, [str(x or "") for x in (ent.get("lines") or [])], remount_comp=True)
        _seed_entry_remarks(iso, i, ent.get("remark_lines") or [""])
        _snapshot_worklog_page(iso, i)
    st.session_state[bk] = True
    return _attach_extra_pages(cells, extras)

def _entry_line_count_key(iso: str, entry_i: int) -> str: return f"wl_ent_lc_{iso}_{entry_i}"
def _entry_line_gen_key(iso: str, entry_i: int) -> str: return f"wl_ent_gen_{iso}_{entry_i}"
def _entry_line_key(iso: str, entry_i: int, line_j: int) -> str: return f"wl_ent_ln_{iso}_{entry_i}_{line_j}_g{int(st.session_state.get(_entry_line_gen_key(iso, entry_i), 0) or 0)}"

def _bump_entry_line_gen(iso: str, entry_i: int) -> None:
    k = _entry_line_gen_key(iso, entry_i)
    old_g, old_n = int(st.session_state.get(k, 0) or 0), int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old_n, 0) + 8):
        st.session_state.pop(f"wl_ent_ln_{iso}_{entry_i}_{j}_g{old_g}", None); st.session_state.pop(f"wl_ent_ln_{iso}_{entry_i}_{j}", None)
    st.session_state[k] = old_g + 1

def _widget_line_parts(iso: str, entry_i: int) -> list[str] | None:
    lc = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
    if lc > 0:
        return [str(st.session_state.get(_entry_line_key(iso, entry_i, j), "") or "") for j in range(lc)]
    raw = str(st.session_state.get(f"wl_ent_t_{iso}_{entry_i}", "") or "")
    return raw.splitlines() if raw else None


def _lines_from_entry_widgets(iso: str, entry_i: int, *, keep_trailing_empty: bool = True) -> list[str]:
    live, cs = st.session_state.get(_entry_lines_live_key(iso, entry_i)), st.session_state.get(_entry_lines_comp_key(iso, entry_i))
    parts = _coalesce_editor_lines(
        [str(x or "") for x in live] if isinstance(live, list) else None,
        _comp_state_lines(cs),
        _widget_line_parts(iso, entry_i),
    )
    if not keep_trailing_empty:
        while parts and parts[-1] == "": parts.pop()
    elif parts and parts[-1] != "": parts = list(parts) + [""]
    elif not parts and keep_trailing_empty: parts = [""]
    return parts

def _content_from_entry_lines(iso: str, entry_i: int) -> str: return "\n".join(_lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=False))

def _comp_focus_seen_key(iso: str, entry_i: int, kind: str) -> str:
    return f"wl_comp_focus_seen_{kind}_{iso}_{entry_i}"


def _comp_send_focus(iso: str, entry_i: int, kind: str) -> int:
    """CCv2 data.focus is one-shot. Sticky widget focus must not steal the clicked cell."""
    fk = st.session_state.get(f"wl_focus_ln_{iso}")
    if not isinstance(fk, str) or not fk:
        return -1
    prefix = {"ln": "wl_ent_ln", "cl": "wl_ent_cl", "rm": "wl_ent_rm"}.get(kind) or ""
    m = re.match(rf"^{re.escape(prefix)}_{re.escape(iso)}_{int(entry_i)}_(\d+)(?:_g\d+)?$", fk)
    if not m:
        return -1
    try:
        return int(m.group(1))
    except (TypeError, ValueError):
        return -1


def _mark_comp_focus_seen(iso: str, entry_i: int, kind: str, fj: int) -> None:
    st.session_state[_comp_focus_seen_key(iso, entry_i, kind)] = int(fj)


def _maybe_remember_comp_focus(iso: str, entry_i: int, kind: str, cur: dict, key_fn) -> None:
    """Remember wrap/enter moves only. Blur keeps the previous widget focus and must not restore it."""
    if not isinstance(cur, dict):
        return
    try:
        fj = int(cur.get("focus", -1))
    except (TypeError, ValueError):
        fj = -1
    seen_k = _comp_focus_seen_key(iso, entry_i, kind)
    if seen_k not in st.session_state:
        st.session_state[seen_k] = fj if fj >= 0 else -1
        return
    seen = st.session_state.get(seen_k)
    if fj < 0 or fj == seen:
        return
    st.session_state[seen_k] = fj
    synced = cur.get("lines") if isinstance(cur.get("lines"), list) else []
    caret = cur.get("caret") if isinstance(cur.get("caret"), dict) else {}
    try:
        pos = int(caret.get("s", len(str((synced[fj] if 0 <= fj < len(synced) else "") or ""))))
    except (TypeError, ValueError, IndexError):
        pos = 0
    _remember_active_cell(iso, key_fn(iso, entry_i, fj), pos)


def _seed_comp_focus_seen(iso: str, entry_i: int, kind: str, focus_n: int) -> None:
    seen_k = _comp_focus_seen_key(iso, entry_i, kind)
    if seen_k not in st.session_state:
        st.session_state[seen_k] = focus_n if isinstance(focus_n, int) and focus_n >= 0 else -1


def _comp_state_lines(cs) -> list[str] | None:
    if isinstance(cs, dict) and isinstance(cs.get("lines"), list):
        return [str(x or "") for x in cs.get("lines") or []]
    lines = getattr(cs, "lines", None)
    if isinstance(lines, list):
        return [str(x or "") for x in lines]
    return None


def _coalesce_editor_lines(*groups: list | None) -> list[str]:
    """Keep the longest known editor text. Empty CCv2 props must not wipe a typed cell."""
    best: list[str] | None = None
    best_score = -1
    for g in groups:
        if not isinstance(g, list):
            continue
        score = sum(len(str(x or "")) for x in g)
        if score > best_score:
            best_score = score
            best = [str(x or "") for x in g]
    return best if best is not None else [""]


def _set_comp_lines_state(iso: str, entry_i: int, chunks: list[str], *, focus_j: int | None = None) -> None:
    ck = _entry_lines_comp_key(iso, entry_i)
    prev = st.session_state.get(ck) if isinstance(st.session_state.get(ck), dict) else {}
    fj = prev.get("focus", -1) if focus_j is None else max(0, min(int(focus_j), max(len(chunks) - 1, 0)))
    new_state = {"lines": list(chunks), "focus": fj}
    if isinstance(prev.get("caret"), dict):
        new_state["caret"] = dict(prev.get("caret") or {})
    # 위젯 키는 중첩 mutate 금지 — 통째로 교체. 이전 값 잔존 시 pop 후 재설정.
    # 이미 instantiate된 CCv2 키를 건드리면 StreamlitAPIException → 탭 ERROR → 버튼 전부 먹통.
    try:
        st.session_state.pop(ck, None)
        st.session_state[ck] = new_state
    except StreamlitAPIException:
        _bump_entry_lines_comp_inst(iso, entry_i)
        ck = _entry_lines_comp_key(iso, entry_i)
        st.session_state[ck] = new_state
    if focus_j is not None:
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_line_key(iso, entry_i, int(fj))
        _mark_comp_focus_seen(iso, entry_i, "ln", int(fj))
    st.session_state[_entry_lines_live_key(iso, entry_i)] = list(chunks)


def _norm_editor_lines(lines: list | None) -> list[str]:
    return _pad_sheet_lines(lines)


def _result_lines_list(result: Any) -> list[str] | None:
    if result is None:
        return None
    if hasattr(result, "lines") and isinstance(result.lines, list):
        return _norm_editor_lines(result.lines)
    if isinstance(result, dict) and isinstance(result.get("lines"), list):
        return _norm_editor_lines(result.get("lines"))
    return None


def _bound_editor_key(iso: str, entry_i: int, kind: str) -> str:
    return f"wl_bound_editor_{kind}_{iso}_{entry_i}"


def _forced_editor_key(iso: str, entry_i: int, kind: str) -> str:
    return {
        "ln": f"wl_force_comp_lines_{iso}_{entry_i}",
        "cl": f"wl_force_comp_clients_{iso}_{entry_i}",
        "rm": f"wl_force_comp_remarks_{iso}_{entry_i}",
    }[kind]


def _mark_editor_bound(iso: str, entry_i: int, kind: str) -> None:
    st.session_state[_bound_editor_key(iso, entry_i, kind)] = True


def _clear_editor_bound(iso: str, entry_i: int, kind: str) -> None:
    st.session_state.pop(_bound_editor_key(iso, entry_i, kind), None)


def _program_editor_lines(iso: str, entry_i: int, kind: str, live: Any) -> tuple[list[str] | None, bool]:
    """날짜 시드(강제)일 때만 칸을 통째로 교체한다. 칸 이동·입력 중에는 화면 값을 유지한다."""
    forced = st.session_state.get(_forced_editor_key(iso, entry_i, kind))
    if isinstance(forced, list):
        return _pad_sheet_lines(forced), True
    return None, False


def _pick_editor_out_lines(
    *,
    session_lines: list[str],
    result: Any,
    forced: Any,
    changed_key: str,
) -> list[str]:
    """CCv2 result.lines가 저장/시드 이후 구값으로 session을 덮지 않게 선택.

    - forced(프로그램 시드) 우선
    - 사용자 편집 콜백이 난 런만 result 채택
    - 그 외에는 session(live) 유지 ← 저장 후 구값 부활 차단
    """
    session_n = _norm_editor_lines(session_lines)
    if isinstance(forced, list):
        st.session_state.pop(changed_key, None)
        return _norm_editor_lines(forced)
    user_changed = bool(st.session_state.pop(changed_key, None))
    result_n = _result_lines_list(result)
    if user_changed and result_n is not None:
        return _coalesce_editor_lines(result_n, session_n)
    return session_n


def _apply_entry_lines(iso: str, entry_i: int, lines: list[str], *, focus_j: int | None = None, bump_gen: bool = False, remount_comp: bool = False) -> None:
    if bump_gen: _bump_entry_line_gen(iso, entry_i)
    if remount_comp: _bump_entry_lines_comp_inst(iso, entry_i)
    chunks = _pad_sheet_lines(lines)
    old = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(chunks)) + 3): st.session_state.pop(_entry_line_key(iso, entry_i, j), None)
    st.session_state[_entry_line_count_key(iso, entry_i)] = len(chunks)
    for j, line in enumerate(chunks): st.session_state[_entry_line_key(iso, entry_i, j)] = line
    st.session_state[f"wl_ent_t_{iso}_{entry_i}"] = "\n".join(chunks)
    _set_comp_lines_state(iso, entry_i, chunks, focus_j=focus_j)
    st.session_state[_entry_lines_rev_key(iso, entry_i)] = int(st.session_state.get(_entry_lines_rev_key(iso, entry_i), 0) or 0) + 1
    # 컴포넌트가 이전 result.lines로 덮어쓰지 않도록 가드
    st.session_state[f"wl_force_comp_lines_{iso}_{entry_i}"] = list(chunks)
    _mark_editor_bound(iso, entry_i, "ln")
    if focus_j is not None:
        fj = max(0, min(int(focus_j), len(chunks) - 1))
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_line_key(iso, entry_i, fj)
        st.session_state[f"wl_focus_caret_{iso}"] = len(str(chunks[fj] or ""))

def _insert_line_after(iso: str, entry_i: int, line_j: int) -> None:
    cur = _pad_sheet_lines(_lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=True))
    key = _entry_line_key(iso, entry_i, line_j)
    if key in st.session_state: cur[line_j] = str(st.session_state.get(key) or "")
    _apply_entry_lines(iso, entry_i, cur, focus_j=min(line_j + 1, WL_SHEET_N - 1))

def _entry_client_count_key(iso: str, entry_i: int) -> str: return f"wl_ent_clc_{iso}_{entry_i}"
def _entry_client_key(iso: str, entry_i: int, line_j: int) -> str: return f"wl_ent_cl_{iso}_{entry_i}_{line_j}"

def _clients_from_widgets(iso: str, entry_i: int, *, keep_trailing_empty: bool = False) -> list[str]:
    live, cs = st.session_state.get(_entry_clients_live_key(iso, entry_i)), st.session_state.get(_entry_clients_comp_key(iso, entry_i))
    lc = int(st.session_state.get(_entry_client_count_key(iso, entry_i), 0) or 0)
    if lc > 0:
        widget_parts = [str(st.session_state.get(_entry_client_key(iso, entry_i, j), "") or "") for j in range(lc)]
    else:
        raw = str(st.session_state.get(f"wl_ent_c_{iso}_{entry_i}", "") or "")
        widget_parts = raw.splitlines() if raw else None
    parts = _coalesce_editor_lines(
        [str(x or "") for x in live] if isinstance(live, list) else None,
        _comp_state_lines(cs),
        widget_parts,
    )
    if not keep_trailing_empty:
        while parts and not str(parts[-1]).strip(): parts.pop()
    elif parts and str(parts[-1]).strip() != "": parts = list(parts) + [""]
    elif not parts and keep_trailing_empty: parts = [""]
    return parts

def _set_comp_clients_state(iso: str, entry_i: int, chunks: list[str], *, focus_j: int | None = None) -> None:
    ck = _entry_clients_comp_key(iso, entry_i)
    prev = st.session_state.get(ck) if isinstance(st.session_state.get(ck), dict) else {}
    fj = prev.get("focus", -1) if focus_j is None else max(0, min(int(focus_j), max(len(chunks) - 1, 0)))
    new_state = {"lines": list(chunks), "focus": fj}
    if isinstance(prev.get("caret"), dict):
        new_state["caret"] = dict(prev.get("caret") or {})
    try:
        st.session_state.pop(ck, None)
        st.session_state[ck] = new_state
    except StreamlitAPIException:
        _bump_entry_clients_comp_inst(iso, entry_i)
        ck = _entry_clients_comp_key(iso, entry_i)
        st.session_state[ck] = new_state
    if focus_j is not None:
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_client_key(iso, entry_i, int(fj))
        _mark_comp_focus_seen(iso, entry_i, "cl", int(fj))
    st.session_state[_entry_clients_live_key(iso, entry_i)] = list(chunks)

def _apply_entry_clients(iso: str, entry_i: int, lines: list[str], *, focus_j: int | None = None, remount_comp: bool = False) -> None:
    if remount_comp: _bump_entry_clients_comp_inst(iso, entry_i)
    chunks = _pad_sheet_lines(lines)
    old = int(st.session_state.get(_entry_client_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(chunks)) + 3): st.session_state.pop(_entry_client_key(iso, entry_i, j), None)
    st.session_state[_entry_client_count_key(iso, entry_i)] = len(chunks)
    for j, line in enumerate(chunks): st.session_state[_entry_client_key(iso, entry_i, j)] = line
    filled = list(chunks)
    while filled and filled[-1] == "": filled.pop()
    st.session_state[f"wl_ent_c_{iso}_{entry_i}"] = "\n".join(filled)
    _set_comp_clients_state(iso, entry_i, chunks, focus_j=focus_j)
    st.session_state[_entry_clients_rev_key(iso, entry_i)] = int(st.session_state.get(_entry_clients_rev_key(iso, entry_i), 0) or 0) + 1
    st.session_state[f"wl_force_comp_clients_{iso}_{entry_i}"] = list(chunks)
    _mark_editor_bound(iso, entry_i, "cl")
    if focus_j is not None:
        fj = max(0, min(int(focus_j), len(chunks) - 1))
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_client_key(iso, entry_i, fj)
        st.session_state[f"wl_focus_caret_{iso}"] = len(str(chunks[fj] or ""))

def _seed_entry_clients(iso: str, entry_i: int, client: str | list[str], *, remount_comp: bool = True) -> None:
    max_u = _client_line_units()
    src = [str(x or "") for x in client] if isinstance(client, list) else (str(client or "").split('\n') if str(client or "") else [""])
    chunks: list[str] = []
    for line in src:
        s = str(line or "")
        if not s.strip():
            chunks.append(s)
            continue
        chunks.extend(_chunk_text(s, max_u) or [s])
    if not chunks: chunks = [""]
    fj = max(0, len(chunks) - 1)
    for j, line in enumerate(chunks):
        if _display_units(line) >= max_u: fj = min(j + 1, len(chunks))
    _apply_entry_clients(iso, entry_i, chunks, focus_j=fj, remount_comp=remount_comp)

def _insert_client_after(iso: str, entry_i: int, line_j: int) -> None:
    cur = _pad_sheet_lines(_clients_from_widgets(iso, entry_i, keep_trailing_empty=True))
    key = _entry_client_key(iso, entry_i, line_j)
    if key in st.session_state: cur[line_j] = str(st.session_state.get(key) or "")
    _apply_entry_clients(iso, entry_i, cur, focus_j=min(line_j + 1, WL_SHEET_N - 1))

def _entry_remark_count_key(iso: str, entry_i: int) -> str: return f"wl_ent_rmc_{iso}_{entry_i}"
def _entry_remark_key(iso: str, entry_i: int, line_j: int) -> str: return f"wl_ent_rm_{iso}_{entry_i}_{line_j}"

def _remarks_from_widgets(iso: str, entry_i: int, *, keep_trailing_empty: bool = False) -> list[str]:
    live, cs = st.session_state.get(_entry_remarks_live_key(iso, entry_i)), st.session_state.get(_entry_remarks_comp_key(iso, entry_i))
    lc = int(st.session_state.get(_entry_remark_count_key(iso, entry_i), 0) or 0)
    if lc > 0:
        widget_parts = [str(st.session_state.get(_entry_remark_key(iso, entry_i, j), "") or "") for j in range(lc)]
    else:
        raw = str(st.session_state.get(f"wl_ent_r_{iso}_{entry_i}", "") or "")
        widget_parts = raw.splitlines() if raw else None
    parts = _coalesce_editor_lines(
        [str(x or "") for x in live] if isinstance(live, list) else None,
        _comp_state_lines(cs),
        widget_parts,
    )
    if not keep_trailing_empty:
        while parts and not str(parts[-1]).strip(): parts.pop()
    elif parts and str(parts[-1]).strip() != "": parts = list(parts) + [""]
    elif not parts and keep_trailing_empty: parts = [""]
    return parts

def _set_comp_remarks_state(iso: str, entry_i: int, chunks: list[str], *, focus_j: int | None = None) -> None:
    ck = _entry_remarks_comp_key(iso, entry_i)
    prev = st.session_state.get(ck) if isinstance(st.session_state.get(ck), dict) else {}
    fj = prev.get("focus", -1) if focus_j is None else max(0, min(int(focus_j), max(len(chunks) - 1, 0)))
    new_state = {"lines": list(chunks), "focus": fj}
    if isinstance(prev.get("caret"), dict):
        new_state["caret"] = dict(prev.get("caret") or {})
    try:
        st.session_state.pop(ck, None)
        st.session_state[ck] = new_state
    except StreamlitAPIException:
        _bump_entry_remarks_comp_inst(iso, entry_i)
        ck = _entry_remarks_comp_key(iso, entry_i)
        st.session_state[ck] = new_state
    if focus_j is not None:
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_remark_key(iso, entry_i, int(fj))
        _mark_comp_focus_seen(iso, entry_i, "rm", int(fj))
    st.session_state[_entry_remarks_live_key(iso, entry_i)] = list(chunks)

def _apply_entry_remarks(iso: str, entry_i: int, lines: list[str], *, focus_j: int | None = None, remount_comp: bool = False) -> None:
    if remount_comp: _bump_entry_remarks_comp_inst(iso, entry_i)
    chunks = _pad_sheet_lines(lines)
    old = int(st.session_state.get(_entry_remark_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(chunks)) + 3): st.session_state.pop(_entry_remark_key(iso, entry_i, j), None)
    st.session_state[_entry_remark_count_key(iso, entry_i)] = len(chunks)
    for j, line in enumerate(chunks): st.session_state[_entry_remark_key(iso, entry_i, j)] = line
    filled = list(chunks)
    while filled and filled[-1] == "": filled.pop()
    st.session_state[f"wl_ent_r_{iso}_{entry_i}"] = "\n".join(filled)
    _set_comp_remarks_state(iso, entry_i, chunks, focus_j=focus_j)
    st.session_state[_entry_remarks_rev_key(iso, entry_i)] = int(st.session_state.get(_entry_remarks_rev_key(iso, entry_i), 0) or 0) + 1
    st.session_state[f"wl_force_comp_remarks_{iso}_{entry_i}"] = list(chunks)
    _mark_editor_bound(iso, entry_i, "rm")
    if focus_j is not None:
        fj = max(0, min(int(focus_j), len(chunks) - 1))
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_remark_key(iso, entry_i, fj)
        st.session_state[f"wl_focus_caret_{iso}"] = len(str(chunks[fj] or ""))

def _seed_entry_remarks(iso: str, entry_i: int, remarks: str | list[str], *, remount_comp: bool = True) -> None:
    max_u = _remark_line_units()
    src = [str(x or "") for x in remarks] if isinstance(remarks, list) else (str(remarks or "").split("\n") if str(remarks or "") else [""])
    chunks: list[str] = []
    for line in src:
        s = str(line or "")
        if not s.strip():
            chunks.append(s)
            continue
        chunks.extend(_chunk_text(s, max_u) or [s])
    if not chunks: chunks = [""]
    fj = max(0, len(chunks) - 1)
    for j, line in enumerate(chunks):
        if _display_units(line) >= max_u: fj = min(j + 1, len(chunks))
    _apply_entry_remarks(iso, entry_i, chunks, focus_j=fj, remount_comp=remount_comp)

def _split_overflow_parts(parts: list[str], max_u: int) -> list[str]:
    out, i, n = len(parts), 0, len(parts)
    out = []
    while i < n:
        s = str(parts[i] or "")
        if _display_units(s) > max_u:
            pieces = _chunk_text(s, max_u) or [s]
            out.append(pieces[0])
            j = i + 1
            for ov in pieces[1:]:
                if j < n and str(parts[j] or "") == ov: out.append(str(parts[j] or "")); j += 1
                else: out.append(ov)
            i = j
        else: out.append(s); i += 1
    return out if out else [""]

def _dedupe_overflow_tail(pieces: list[str], tail: list[str]) -> list[str]:
    if len(pieces) <= 1: return list(tail)
    rest = list(tail)
    for ov in pieces[1:]:
        if rest and str(rest[0] or "") == ov: rest.pop(0)
        else: break
    return rest

def _commit_enter_on_cell(kind: str, iso: str, entry_i: int, line_j: int, value: str) -> None:
    value = str(value or "")
    is_client = kind == "wl_ent_cl"
    is_remark = kind == "wl_ent_rm"
    max_u = _client_line_units() if is_client else (_remark_line_units() if is_remark else _content_line_units())
    if is_client:
        cur = _clients_from_widgets(iso, entry_i, keep_trailing_empty=True)
    elif is_remark:
        cur = _remarks_from_widgets(iso, entry_i, keep_trailing_empty=True)
    else:
        cur = _lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=True)
    new = _pad_sheet_lines(cur)
    line_j = max(0, min(int(line_j), WL_SHEET_N - 1))
    pieces = _chunk_text(value, max_u) or [value] if _display_units(value) > max_u else [value]
    for k, p in enumerate(pieces):
        dest = line_j + k
        if dest >= WL_SHEET_N:
            break
        if k == 0:
            new[dest] = p
        else:
            new[dest] = p + str(new[dest] or "")
    focus = min(line_j + 1, WL_SHEET_N - 1)
    if is_client: _apply_entry_clients(iso, entry_i, new, focus_j=focus)
    elif is_remark: _apply_entry_remarks(iso, entry_i, new, focus_j=focus)
    else: _apply_entry_lines(iso, entry_i, new, focus_j=focus, bump_gen=True)
    st.session_state.pop(f"wl_enter_done_{iso}", None)

def _mount_entry_client_editor(iso: str, entry_i: int, max_u: int) -> list[str]:
    ck, live_key, cs = _entry_clients_comp_key(iso, entry_i), _entry_clients_live_key(iso, entry_i), st.session_state.get(_entry_clients_comp_key(iso, entry_i))
    live = st.session_state.get(live_key)
    prog, replace = _program_editor_lines(iso, entry_i, "cl", live)
    if prog is not None:
        lines, focus = prog, -1
    elif isinstance(cs, dict) and isinstance(cs.get("lines"), list): lines, focus = [str(x or "") for x in cs.get("lines") or []], cs.get("focus", -1)
    elif isinstance(live, list): lines, focus = [str(x or "") for x in live or []], -1
    elif int(st.session_state.get(_entry_client_count_key(iso, entry_i), 0) or 0) > 0: lines, focus = [str(st.session_state.get(_entry_client_key(iso, entry_i, j), "") or "") for j in range(int(st.session_state.get(_entry_client_count_key(iso, entry_i), 1) or 1))], -1
    else:
        raw = str(st.session_state.get(f"wl_ent_c_{iso}_{entry_i}", "") or "")
        if raw:
            _seed_entry_clients(iso, entry_i, raw)
            cs = st.session_state.get(ck)
            if isinstance(cs, dict) and isinstance(cs.get("lines"), list): lines, focus = [str(x or "") for x in cs.get("lines") or []], cs.get("focus", -1)
            else: lines, focus = [""], -1
        else: lines, focus = [""], -1
        replace = False
        live = st.session_state.get(live_key)
    if prog is None:
        lines = _coalesce_editor_lines(lines, live if isinstance(live, list) else None, _comp_state_lines(cs))
    if any(_display_units(p) > max_u for p in lines):
        fixed, focus = _split_overflow_parts(lines, max_u), 0
        for j, line in enumerate(fixed):
            if _display_units(line) >= max_u: focus = min(j + 1, len(fixed))
        if fixed != lines:
            _apply_entry_clients(iso, entry_i, fixed, focus_j=focus)
            cs = st.session_state.get(ck)
            if isinstance(cs, dict) and isinstance(cs.get("lines"), list): lines, focus = [str(x or "") for x in cs.get("lines") or []], cs.get("focus", focus)
            else: lines = fixed
    lines = _pad_sheet_lines(lines)
    try: focus_n = int(focus)
    except (TypeError, ValueError): focus_n = -1

    def _on_clients_change() -> None:
        cur = st.session_state.get(ck)
        synced = _comp_state_lines(cur)
        if synced is not None:
            st.session_state[f"wl_clients_user_edit_{iso}_{entry_i}"] = True
            _clear_editor_bound(iso, entry_i, "cl")
            st.session_state[live_key] = synced
            st.session_state[_entry_client_count_key(iso, entry_i)] = len(synced)
            for j, line in enumerate(synced): st.session_state[_entry_client_key(iso, entry_i, j)] = line
            filled = list(synced)
            while filled and filled[-1] == "": filled.pop()
            st.session_state[f"wl_ent_c_{iso}_{entry_i}"] = "\n".join(filled)
            focus_cur = cur if isinstance(cur, dict) else {"lines": synced, "focus": getattr(cur, "focus", -1)}
            _maybe_remember_comp_focus(iso, entry_i, "cl", focus_cur, _entry_client_key)

    _seed_comp_focus_seen(iso, entry_i, "cl", focus_n)
    result = _WL_LINES_EDITOR(
        key=ck,
        data={"iso": iso, "slot": str(entry_i), "lines": lines, "focus": _comp_send_focus(iso, entry_i, "cl"), "max_u": int(max_u), "cell_w": int(_orig_cell_px("client")), "font_pt": float(_WL_BODY_FONT_PT), "look_scale": float(_input_look_scale()), "variant": "client", "fixed_rows": WL_SHEET_N, "rev": int(st.session_state.get(_entry_clients_rev_key(iso, entry_i), 0) or 0), "replace": 1 if replace else 0},
        default={"lines": lines},
        on_lines_change=_on_clients_change,
    )
    
    forced = st.session_state.pop(f"wl_force_comp_clients_{iso}_{entry_i}", None)
    out = _pick_editor_out_lines(
        session_lines=lines,
        result=result,
        forced=forced,
        changed_key=f"wl_clients_user_edit_{iso}_{entry_i}",
    )
    st.session_state[live_key] = out
    old = int(st.session_state.get(_entry_client_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(out)) + 3): st.session_state.pop(_entry_client_key(iso, entry_i, j), None)
    st.session_state[_entry_client_count_key(iso, entry_i)] = len(out)
    for j, line in enumerate(out): st.session_state[_entry_client_key(iso, entry_i, j)] = line
    filled = list(out)
    while filled and filled[-1] == "": filled.pop()
    st.session_state[f"wl_ent_c_{iso}_{entry_i}"] = "\n".join(filled)
    return out

def _mount_entry_lines_editor(iso: str, entry_i: int, max_u: int) -> list[str]:
    ck, live_key, cs = _entry_lines_comp_key(iso, entry_i), _entry_lines_live_key(iso, entry_i), st.session_state.get(_entry_lines_comp_key(iso, entry_i))
    live = st.session_state.get(live_key)
    prog, replace = _program_editor_lines(iso, entry_i, "ln", live)
    if prog is not None:
        lines, focus = prog, -1
    elif isinstance(cs, dict) and isinstance(cs.get("lines"), list): lines, focus = [str(x or "") for x in cs.get("lines") or []], cs.get("focus", -1)
    elif isinstance(live, list): lines, focus = [str(x or "") for x in live or []], -1
    else: lines, focus = [""], -1
    if prog is None:
        lines = _coalesce_editor_lines(lines, live if isinstance(live, list) else None, _comp_state_lines(cs))
    if any(_display_units(p) > max_u for p in lines):
        fixed, focus = _split_overflow_parts(lines, max_u), 0
        for j, line in enumerate(fixed):
            if _display_units(line) >= max_u: focus = min(j + 1, len(fixed))
        if fixed != lines:
            _apply_entry_lines(iso, entry_i, fixed, focus_j=focus)
            cs = st.session_state.get(ck)
            if isinstance(cs, dict) and isinstance(cs.get("lines"), list): lines, focus = [str(x or "") for x in cs.get("lines") or []], cs.get("focus", focus)
            else: lines = fixed
    lines = _pad_sheet_lines(lines)
    try: focus_n = int(focus)
    except (TypeError, ValueError): focus_n = -1

    def _on_lines_change() -> None:
        cur = st.session_state.get(ck)
        synced = _comp_state_lines(cur)
        if synced is not None:
            st.session_state[f"wl_lines_user_edit_{iso}_{entry_i}"] = True
            _clear_editor_bound(iso, entry_i, "ln")
            st.session_state[live_key] = synced
            st.session_state[_entry_line_count_key(iso, entry_i)] = len(synced)
            for j, line in enumerate(synced): st.session_state[_entry_line_key(iso, entry_i, j)] = line
            filled = list(synced)
            while filled and filled[-1] == "": filled.pop()
            st.session_state[f"wl_ent_t_{iso}_{entry_i}"] = "\n".join(filled)
            focus_cur = cur if isinstance(cur, dict) else {"lines": synced, "focus": getattr(cur, "focus", -1)}
            _maybe_remember_comp_focus(iso, entry_i, "ln", focus_cur, _entry_line_key)

    _seed_comp_focus_seen(iso, entry_i, "ln", focus_n)
    result = _WL_LINES_EDITOR(
        key=ck,
        data={"iso": iso, "slot": str(entry_i), "lines": lines, "focus": _comp_send_focus(iso, entry_i, "ln"), "max_u": int(max_u), "cell_w": int(_orig_cell_px("content")), "font_pt": float(_WL_BODY_FONT_PT), "look_scale": float(_input_look_scale()), "variant": "content", "fixed_rows": WL_SHEET_N, "rev": int(st.session_state.get(_entry_lines_rev_key(iso, entry_i), 0) or 0), "replace": 1 if replace else 0},
        default={"lines": lines},
        on_lines_change=_on_lines_change,
    )
    
    forced = st.session_state.pop(f"wl_force_comp_lines_{iso}_{entry_i}", None)
    out = _pick_editor_out_lines(
        session_lines=lines,
        result=result,
        forced=forced,
        changed_key=f"wl_lines_user_edit_{iso}_{entry_i}",
    )
    st.session_state[live_key] = out
    old = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(out)) + 3): st.session_state.pop(_entry_line_key(iso, entry_i, j), None)
    st.session_state[_entry_line_count_key(iso, entry_i)] = len(out)
    for j, line in enumerate(out): st.session_state[_entry_line_key(iso, entry_i, j)] = line
    filled = list(out)
    while filled and filled[-1] == "": filled.pop()
    st.session_state[f"wl_ent_t_{iso}_{entry_i}"] = "\n".join(filled)
    return out

def _mount_entry_remark_editor(iso: str, entry_i: int, max_u: int) -> list[str]:
    ck, live_key, cs = _entry_remarks_comp_key(iso, entry_i), _entry_remarks_live_key(iso, entry_i), st.session_state.get(_entry_remarks_comp_key(iso, entry_i))
    live = st.session_state.get(live_key)
    prog, replace = _program_editor_lines(iso, entry_i, "rm", live)
    if prog is not None:
        lines, focus = prog, -1
    elif isinstance(cs, dict) and isinstance(cs.get("lines"), list): lines, focus = [str(x or "") for x in cs.get("lines") or []], cs.get("focus", -1)
    elif isinstance(live, list): lines, focus = [str(x or "") for x in live or []], -1
    elif int(st.session_state.get(_entry_remark_count_key(iso, entry_i), 0) or 0) > 0: lines, focus = [str(st.session_state.get(_entry_remark_key(iso, entry_i, j), "") or "") for j in range(int(st.session_state.get(_entry_remark_count_key(iso, entry_i), 1) or 1))], -1
    else:
        raw = str(st.session_state.get(f"wl_ent_r_{iso}_{entry_i}", "") or "")
        if raw:
            _seed_entry_remarks(iso, entry_i, raw)
            cs = st.session_state.get(ck)
            if isinstance(cs, dict) and isinstance(cs.get("lines"), list): lines, focus = [str(x or "") for x in cs.get("lines") or []], cs.get("focus", -1)
            else: lines, focus = [""], -1
        else: lines, focus = [""], -1
        live = st.session_state.get(live_key)
    if prog is None:
        lines = _coalesce_editor_lines(lines, live if isinstance(live, list) else None, _comp_state_lines(cs))
    if any(_display_units(p) > max_u for p in lines):
        fixed, focus = _split_overflow_parts(lines, max_u), 0
        for j, line in enumerate(fixed):
            if _display_units(line) >= max_u: focus = min(j + 1, len(fixed))
        if fixed != lines:
            _apply_entry_remarks(iso, entry_i, fixed, focus_j=focus)
            cs = st.session_state.get(ck)
            if isinstance(cs, dict) and isinstance(cs.get("lines"), list): lines, focus = [str(x or "") for x in cs.get("lines") or []], cs.get("focus", focus)
            else: lines = fixed
    lines = _pad_sheet_lines(lines)
    try: focus_n = int(focus)
    except (TypeError, ValueError): focus_n = -1

    def _on_remarks_change() -> None:
        cur = st.session_state.get(ck)
        synced = _comp_state_lines(cur)
        if synced is not None:
            st.session_state[f"wl_remarks_user_edit_{iso}_{entry_i}"] = True
            _clear_editor_bound(iso, entry_i, "rm")
            st.session_state[live_key] = synced
            st.session_state[_entry_remark_count_key(iso, entry_i)] = len(synced)
            for j, line in enumerate(synced): st.session_state[_entry_remark_key(iso, entry_i, j)] = line
            filled = list(synced)
            while filled and filled[-1] == "": filled.pop()
            st.session_state[f"wl_ent_r_{iso}_{entry_i}"] = "\n".join(filled)
            focus_cur = cur if isinstance(cur, dict) else {"lines": synced, "focus": getattr(cur, "focus", -1)}
            _maybe_remember_comp_focus(iso, entry_i, "rm", focus_cur, _entry_remark_key)

    _seed_comp_focus_seen(iso, entry_i, "rm", focus_n)
    result = _WL_LINES_EDITOR(
        key=ck,
        data={"iso": iso, "slot": str(entry_i), "lines": lines, "focus": _comp_send_focus(iso, entry_i, "rm"), "max_u": int(max_u), "cell_w": int(_orig_cell_px("remark")), "font_pt": float(_WL_BODY_FONT_PT), "look_scale": float(_input_look_scale()), "variant": "remark", "fixed_rows": WL_SHEET_N, "rev": int(st.session_state.get(_entry_remarks_rev_key(iso, entry_i), 0) or 0), "replace": 1 if replace else 0},
        default={"lines": lines},
        on_lines_change=_on_remarks_change,
    )

    forced = st.session_state.pop(f"wl_force_comp_remarks_{iso}_{entry_i}", None)
    out = _pick_editor_out_lines(
        session_lines=lines,
        result=result,
        forced=forced,
        changed_key=f"wl_remarks_user_edit_{iso}_{entry_i}",
    )
    st.session_state[live_key] = out
    old = int(st.session_state.get(_entry_remark_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(out)) + 3): st.session_state.pop(_entry_remark_key(iso, entry_i, j), None)
    st.session_state[_entry_remark_count_key(iso, entry_i)] = len(out)
    for j, line in enumerate(out): st.session_state[_entry_remark_key(iso, entry_i, j)] = line
    filled = list(out)
    while filled and filled[-1] == "": filled.pop()
    st.session_state[f"wl_ent_r_{iso}_{entry_i}"] = "\n".join(filled)
    return out


def _remember_active_cell(iso: str, fk: str, pos: int) -> None:
    st.session_state["wl_active_cell_key"] = fk
    st.session_state[f"wl_focus_ln_{iso}"] = fk
    st.session_state["wl_active_cell_sel"] = (pos, pos)
    st.session_state[f"wl_focus_caret_{iso}"] = pos


def _sync_editor_focus_from_comp(iso: str, entry_i: int, comp_key: str, key_fn) -> None:
    cur = st.session_state.get(comp_key)
    if not isinstance(cur, dict):
        return
    try:
        fj = int(cur.get("focus", -1))
    except (TypeError, ValueError):
        fj = -1
    if fj < 0:
        return
    caret = cur.get("caret") if isinstance(cur.get("caret"), dict) else {}
    try:
        pos = int(caret.get("s", 0))
    except (TypeError, ValueError):
        pos = 0
    fk = key_fn(iso, entry_i, fj)
    _remember_active_cell(iso, fk, pos)


def _last_used_line_index(lines: list[str]) -> int:
    for j in range(len(lines) - 1, -1, -1):
        if str(lines[j] or ""):
            return j
    return 0


def _seed_entry_lines(iso: str, entry_i: int, content: str, *, focus_j: int | None = None, focus_last: bool = False) -> None:
    max_u = _content_line_units()
    src = str(content or "").split('\n') if str(content or "") else [""]
    chunks = []
    for line in src:
        if not line.strip(): 
            chunks.append(line)
            continue
        chunks.extend(_chunk_text(line, max_u) or [line])
    if not chunks: chunks = [""]
    fj = focus_j
    if fj is None and focus_last:
        fj = max(0, len(chunks) - 1)
        for j, line in enumerate(chunks):
            if _display_units(line) >= max_u: fj = min(j + 1, len(chunks))
    _apply_entry_lines(iso, entry_i, chunks, focus_j=fj, remount_comp=True)

def _read_editor_entries(d: date) -> list[dict]:
    iso = d.isoformat()
    n = int(st.session_state.get(f"wl_entry_count_{iso}", 0) or 0)
    stored = st.session_state.get(_entries_key(d)) or [{"client": "", "content": "", "blank_after": 1}]
    if n <= 0: n = len(stored)
    out: list[dict] = []
    for i in range(n):
        ck, gk, lc = f"wl_ent_c_{iso}_{i}", f"wl_ent_gap_{iso}_{i}", int(st.session_state.get(_entry_line_count_key(iso, i), 0) or 0)
        if ck in st.session_state or lc > 0 or f"wl_ent_t_{iso}_{i}" in st.session_state or f"wl_ent_r_{iso}_{i}" in st.session_state or int(st.session_state.get(_entry_client_count_key(iso, i), 0) or 0) > 0 or int(st.session_state.get(_entry_remark_count_key(iso, i), 0) or 0) > 0:
            if int(st.session_state.get(_entry_client_count_key(iso, i), 0) or 0) > 0:
                client_lines = _clients_from_widgets(iso, i, keep_trailing_empty=False)
                client = "\n".join(client_lines)
            else:
                client = str(st.session_state.get(ck, "") or "")
                client_lines = _entry_client_lines({"client": client})
            lines = _lines_from_entry_widgets(iso, i, keep_trailing_empty=False)
            content = "\n".join(lines)
            if int(st.session_state.get(_entry_remark_count_key(iso, i), 0) or 0) > 0 or f"wl_ent_r_{iso}_{i}" in st.session_state or isinstance(st.session_state.get(_entry_remarks_live_key(iso, i)), list):
                remark_lines = _remarks_from_widgets(iso, i, keep_trailing_empty=False)
                remarks = "\n".join(remark_lines)
            else:
                remarks, remark_lines = "", []
            blank_after = _entry_blank_after({"blank_after": st.session_state.get(gk)}, 1) if gk in st.session_state else (_entry_blank_after(stored[i], 1) if i < len(stored) else 1)
        elif i < len(stored):
            client = str(stored[i].get("client") or "")
            if not client and isinstance(stored[i].get("client_lines"), list): client = "\n".join(str(x or "") for x in stored[i].get("client_lines") or [])
            client_lines = _entry_client_lines(stored[i])
            lines = stored[i].get("lines")
            if not isinstance(lines, list): lines = _chunk_text(str(stored[i].get("content") or ""), _content_line_units()) or []
            content = ("\n".join(str(x or "") for x in lines) if lines else str(stored[i].get("content") or ""))
            remark_lines = stored[i].get("remark_lines")
            if not isinstance(remark_lines, list): remark_lines = _entry_remark_lines(stored[i])
            remarks = ("\n".join(str(x or "") for x in remark_lines) if remark_lines else str(stored[i].get("remarks") or ""))
            blank_after = _entry_blank_after(stored[i], 1)
        else:
            client, content, remarks, blank_after, lines, client_lines, remark_lines = "", "", "", 1, [], [], []
        # 주의: 위젯이 비어 있다고 해서 stored(이전 저장값)로 되살리지 않음.
        # 사용자가 지운 뒤 저장하면 구값이 되살아나던 원인이었음.
        out.append({"client": client, "client_lines": client_lines, "content": content, "lines": lines, "remarks": remarks, "remark_lines": remark_lines, "blank_after": blank_after})
    return out or [{"client": "", "client_lines": [], "content": "", "lines": [], "remarks": "", "remark_lines": [], "blank_after": 1}]

def _sheet_lines_from_widgets(iso: str) -> tuple[list[str], list[str], list[str]]:
    return _sheet_lines_from_widgets_at(iso, 0)


def _extra_page_cells_from_widgets(d: date) -> list[dict]:
    iso = d.isoformat()
    n = _page_count_for(iso)
    out: list[dict] = []
    for i in range(1, n):
        clients, contents, remarks = _sheet_lines_from_widgets_at(iso, i)
        out.append(_pack_sheet_to_cells(d, clients, contents, remarks, [], []))
    return out


def _cells_from_widgets(d: date) -> dict:
    iso = d.isoformat()
    clients, contents, remarks = _sheet_lines_from_widgets(iso)
    nk, ok = f"wl_next_area_{iso}", f"wl_notes_area_{iso}"
    next_raw = str(st.session_state.get(nk) or st.session_state.get(_next_key(d), "") or "")
    notes_raw = str(st.session_state.get(ok) or st.session_state.get(_notes_key(d), "") or "")
    cells = _pack_sheet_to_cells(d, clients, contents, remarks, _textarea_lines(next_raw), _textarea_lines(notes_raw))
    return _attach_extra_pages(cells, _extra_page_cells_from_widgets(d))

def _seed_day_entry_widgets(
    d: date,
    entries_list: list[dict],
    next_txt: str,
    notes_txt: str,
    extra_pages: list[dict] | None = None,
    remount_comp: bool = True,
) -> None:
    """저장/추가/삭제 후 입력 위젯을 entries 기준으로 다시 심는다. CCv2 인스턴스 키도 갱신."""
    iso = d.isoformat()
    ek = _entries_key(d)
    if extra_pages is None:
        extra_pages = []
        for i in range(1, _page_count_for(iso)):
            extra_pages.append(_sheet_entry_from_lines(*_sheet_lines_from_widgets_at(iso, i)))
    extra_pages = [e for e in extra_pages if isinstance(e, dict)]
    old_n = max(
        int(st.session_state.get(f"wl_entry_count_{iso}", 0) or 0),
        _page_count_for(iso),
        1 + len(extra_pages),
        len(entries_list),
    )
    for i in range(old_n + 2):
        st.session_state.pop(f"wl_ent_c_{iso}_{i}", None)
        st.session_state.pop(f"wl_ent_t_{iso}_{i}", None)
        st.session_state.pop(f"wl_ent_gap_{iso}_{i}", None)
        st.session_state.pop(f"wl_exp_{iso}_{i}", None)
        st.session_state.pop(_entry_lines_comp_key(iso, i), None)
        st.session_state.pop(f"wl_lines_comp_{iso}_{i}", None)
        st.session_state.pop(_entry_lines_rev_key(iso, i), None)
        st.session_state.pop(_entry_lines_live_key(iso, i), None)
        st.session_state.pop(f"wl_force_comp_lines_{iso}_{i}", None)
        st.session_state.pop(f"wl_lines_user_edit_{iso}_{i}", None)
        st.session_state.pop(_entry_clients_comp_key(iso, i), None)
        st.session_state.pop(f"wl_clients_comp_{iso}_{i}", None)
        st.session_state.pop(_entry_clients_rev_key(iso, i), None)
        st.session_state.pop(_entry_clients_live_key(iso, i), None)
        st.session_state.pop(f"wl_force_comp_clients_{iso}_{i}", None)
        st.session_state.pop(f"wl_clients_user_edit_{iso}_{i}", None)
        st.session_state.pop(_entry_remarks_comp_key(iso, i), None)
        st.session_state.pop(f"wl_remarks_comp_{iso}_{i}", None)
        st.session_state.pop(_entry_remarks_rev_key(iso, i), None)
        st.session_state.pop(_entry_remarks_live_key(iso, i), None)
        st.session_state.pop(f"wl_force_comp_remarks_{iso}_{i}", None)
        st.session_state.pop(f"wl_remarks_user_edit_{iso}_{i}", None)
        st.session_state.pop(f"wl_ent_r_{iso}_{i}", None)
        old_lc = int(st.session_state.get(_entry_line_count_key(iso, i), 0) or 0)
        for j in range(old_lc + 3): st.session_state.pop(_entry_line_key(iso, i, j), None)
        st.session_state.pop(_entry_line_count_key(iso, i), None)
        old_cc = int(st.session_state.get(_entry_client_count_key(iso, i), 0) or 0)
        for j in range(old_cc + 3): st.session_state.pop(_entry_client_key(iso, i, j), None)
        st.session_state.pop(_entry_client_count_key(iso, i), None)
        old_rc = int(st.session_state.get(_entry_remark_count_key(iso, i), 0) or 0)
        for j in range(old_rc + 3): st.session_state.pop(_entry_remark_key(iso, i, j), None)
        st.session_state.pop(_entry_remark_count_key(iso, i), None)
    st.session_state.pop(f"wl_next_area_{iso}", None)
    st.session_state.pop(f"wl_notes_area_{iso}", None)
    sheet = _flatten_to_sheet_entry(d, entries_list)
    entries_list = [sheet]
    st.session_state[ek] = entries_list
    st.session_state[f"wl_entry_count_{iso}"] = 1
    _seed_entry_clients(iso, 0, sheet.get("client_lines") or [""], remount_comp=remount_comp)
    _apply_entry_lines(iso, 0, [str(x or "") for x in (sheet.get("lines") or [])], remount_comp=remount_comp)
    _seed_entry_remarks(iso, 0, sheet.get("remark_lines") or [""], remount_comp=remount_comp)
    _snapshot_worklog_page(iso, 0)
    for i, extra in enumerate(extra_pages, start=1):
        if i >= WL_MAX_PAGES:
            break
        cl = extra.get("client_lines") if isinstance(extra.get("client_lines"), list) else []
        ln = extra.get("lines") if isinstance(extra.get("lines"), list) else []
        rm = extra.get("remark_lines") if isinstance(extra.get("remark_lines"), list) else []
        if not cl and not ln and not rm:
            extra_ent = _sheet_entry_from_cells(extra) if any(str(extra.get(k, "") or "") for k in extra) else _empty_sheet_entry()
            cl, ln, rm = extra_ent.get("client_lines") or [""], extra_ent.get("lines") or [""], extra_ent.get("remark_lines") or [""]
        _seed_entry_clients(iso, i, cl or [""], remount_comp=remount_comp)
        _apply_entry_lines(iso, i, [str(x or "") for x in (ln or [])], remount_comp=remount_comp)
        _seed_entry_remarks(iso, i, rm or [""], remount_comp=remount_comp)
        _snapshot_worklog_page(iso, i)
    st.session_state[_page_count_key(iso)] = max(1, min(WL_MAX_PAGES, 1 + len(extra_pages)))
    st.session_state[_page_idx_key(iso)] = min(_page_idx_for(iso), st.session_state[_page_count_key(iso)] - 1)
    st.session_state[f"wl_next_area_{iso}"] = next_txt
    st.session_state[f"wl_notes_area_{iso}"] = notes_txt
    st.session_state[_next_key(d)] = next_txt
    st.session_state[_notes_key(d)] = notes_txt


def _view_cells_key(d: date) -> str: return f"wl_view_cells_{d.isoformat()}"

def _publish_view_cells(d: date, cells: dict) -> None:
    iso = d.isoformat()
    cells = dict(cells or {})
    prev = st.session_state.get(_view_cells_key(d))
    if isinstance(prev, dict):
        try:
            if json.dumps(prev, ensure_ascii=False, sort_keys=True) == json.dumps(cells, ensure_ascii=False, sort_keys=True):
                return
        except Exception:
            pass
    st.session_state[_view_cells_key(d)] = cells
    for k in (
        f"wl_sum_sig_{iso}", f"wl_sum_html_{iso}",
        f"wl_sum_sig_v25_{iso}", f"wl_sum_html_v25_{iso}",
        f"wl_left_excel_sig_v24_{iso}", f"wl_left_excel_html_v24_{iso}", f"wl_left_excel_h_v24_{iso}",
        f"wl_left_excel_sig_v25_{iso}", f"wl_left_excel_html_v25_{iso}", f"wl_left_excel_h_v25_{iso}",
        f"wl_sum_sig_v26_{iso}", f"wl_sum_html_v26_{iso}",
    ):
        st.session_state.pop(k, None)

def _view_cells_for_preview(d: date) -> dict:
    key = _view_cells_key(d)
    cached = st.session_state.get(key)
    if isinstance(cached, dict) and cached:
        return cached
    try:
        if os.path.exists(worklog_path(d)):
            cells = _attach_extra_pages(read_worklog_cells(d), read_worklog_extra_page_cells(d))
            st.session_state[key] = cells
            return cells
    except Exception:
        pass
    cells = _empty_cells(d)
    st.session_state[key] = cells
    return cells


def _draft_cells_for_left_preview(d: date) -> dict:
    """왼쪽 요약/엑셀 — 현재 입력 위젯 값 (실시간). 위젯이 비면 그날 저장본."""
    try:
        live = _cells_from_widgets(d)
        if _worklog_cells_have_draft(live):
            return live
    except Exception:
        live = None
    published = st.session_state.get(_view_cells_key(d))
    if isinstance(published, dict) and published and _worklog_cells_have_draft(published):
        return published
    try:
        return _view_cells_for_preview(d)
    except Exception:
        return live or _empty_cells(d)


def _clear_date_widget_state(d: date) -> None:
    iso = d.isoformat()
    prefixes = (f"wl_ent_c_{iso}_", f"wl_ent_t_{iso}_", f"wl_ent_r_{iso}_", f"wl_ent_gap_{iso}_", f"wl_ent_ln_{iso}_", f"wl_ent_lc_{iso}_", f"wl_ent_gen_{iso}_", f"wl_ent_cl_{iso}_", f"wl_ent_clc_{iso}_", f"wl_ent_rm_{iso}_", f"wl_ent_rmc_{iso}_", f"wl_ent_rev_{iso}_", f"wl_lines_comp_{iso}_", f"wl_lines_live_{iso}_", f"wl_lines_inst_{iso}_", f"wl_clients_comp_{iso}_", f"wl_clients_live_{iso}_", f"wl_clients_inst_{iso}_", f"wl_clients_rev_{iso}_", f"wl_remarks_comp_{iso}_", f"wl_remarks_live_{iso}_", f"wl_remarks_inst_{iso}_", f"wl_remarks_rev_{iso}_", f"wl_force_comp_lines_{iso}_", f"wl_force_comp_clients_{iso}_", f"wl_force_comp_remarks_{iso}_", f"wl_lines_user_edit_{iso}_", f"wl_clients_user_edit_{iso}_", f"wl_remarks_user_edit_{iso}_", f"wl_exp_{iso}_", f"wl_entries_{iso}", f"wl_next_{iso}", f"wl_notes_{iso}", f"wl_next_area_{iso}", f"wl_notes_area_{iso}", f"wl_entry_count_{iso}", f"worklog_booted_{iso}", f"wl_save_btn_{iso}", f"wl_focus_ln_{iso}", f"wl_do_save_{iso}", f"wl_flash_save_{iso}", f"wl_view_cells_{iso}", f"wl_parked_cells_{iso}")
    for k in list(st.session_state.keys()):
        if not isinstance(k, str): continue
        if k in prefixes or any(k.startswith(p) for p in prefixes if p.endswith("_")): del st.session_state[k]
        elif k.startswith(f"wl_page_snap_{iso}_"): del st.session_state[k]
        elif k in {f"wl_entries_{iso}", f"wl_next_{iso}", f"wl_notes_{iso}", f"wl_entry_count_{iso}", f"worklog_booted_{iso}", f"wl_next_area_{iso}", f"wl_notes_area_{iso}", f"wl_pending_sync_{iso}", f"wl_do_add_{iso}", f"wl_do_del_{iso}", f"wl_focus_ln_{iso}", f"wl_do_save_{iso}", f"wl_flash_save_{iso}", f"wl_view_cells_{iso}", f"wl_open_ctx_{iso}", f"wl_saved_ok_{iso}", f"wl_page_count_{iso}", f"wl_page_idx_{iso}"}: del st.session_state[k]

def _preview_path(d: date) -> str: return os.path.join(WORKLOG_DIR, f"_preview_{d.isoformat()}.xlsx")

def _build_preview_file(d: date, cells: dict) -> str:
    _ensure_dirs()
    if not os.path.exists(WORKLOG_TEMPLATE): raise FileNotFoundError("업무일지 템플릿이 없습니다.")
    dst = _preview_path(d)
    write_cells_to_path(dst, d, cells, force_template=True)
    return dst

def _excel_app_path() -> str | None:
    for p in ("/Applications/Microsoft Excel.app", os.path.expanduser("~/Applications/Microsoft Excel.app")):
        if os.path.isdir(p): return p
    return None

def _print_xlsx_path(d: date) -> str: return os.path.join(WORKLOG_DIR, f"일일업무일지_{d.isoformat()}_인쇄.xlsx")

def prepare_print_xlsx(d: date, cells: dict) -> str:
    _ensure_dirs()
    dst = _print_xlsx_path(d)
    write_cells_to_path(dst, d, cells, force_template=True)
    return os.path.abspath(dst)

def open_excel_print_preview(xlsx_path: str, *, prefer_print_dialog: bool = True) -> tuple[bool, str]:
    abs_path = os.path.abspath(xlsx_path)
    if not os.path.exists(abs_path): return False, "미리보기용 엑셀 파일이 없습니다."
    if platform.system() != "Darwin": return False, "Excel 인쇄 화면은 맥에서만 자동 연결됩니다."
    if not _excel_app_path():
        try: subprocess.Popen(["open", abs_path], start_new_session=True); return True, "파일을 열었습니다. Excel이 없다면 설치 후 다시 시도해 주세요."
        except Exception as e: return False, f"파일 열기 실패: {e}"
    ap = abs_path.replace("\\", "\\\\").replace('"', '\\"')
    if prefer_print_dialog:
        script = f'''set targetFile to POSIX file "{ap}"\ntell application "Microsoft Excel"\nactivate\nopen targetFile\ndelay 1.2\nend tell\ntell application "System Events"\nif exists process "Microsoft Excel" then\ntell process "Microsoft Excel"\nset frontmost to true\ndelay 0.4\nkeystroke "p" using {{command down}}\nend tell\nend if\nend tell\nreturn true'''
        ok_msg = "Excel에서 열어 인쇄(미리보기) 화면까지 연결했습니다."
    else:
        script = f'''set targetFile to POSIX file "{ap}"\nset previewDone to false\ntell application "Microsoft Excel"\nactivate\nopen targetFile\ndelay 1.3\ntry\nprint preview active sheet\nset previewDone to true\nend try\nif previewDone is false then\ntry\nprint preview\nset previewDone to true\nend try\nend if\nend tell\nif previewDone is false then\ntell application "System Events"\nif exists process "Microsoft Excel" then\ntell process "Microsoft Excel"\nset frontmost to true\ndelay 0.35\ntry\nclick menu item "인쇄 미리 보기" of menu "파일" of menu bar 1\nset previewDone to true\nend try\nif previewDone is false then\nkeystroke "p" using {{command down}}\nset previewDone to true\nend if\nend tell\nend if\nend tell\nend if\nreturn previewDone'''
        ok_msg = "Excel에서 열어 인쇄 미리보기까지 연결했습니다."
    try:
        r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=50)
        if r.returncode == 0: return True, ok_msg
        subprocess.Popen(["open", "-a", "Microsoft Excel", abs_path], start_new_session=True)
        err = (r.stderr or r.stdout or "").strip()
        hint = " (손쉬운 사용 권한을 허용하면 자동 연결됩니다)" if "Not authorized" in err or "assistive" in err.lower() or "1002" in err else ""
        return True, "Excel에서 파일을 열었습니다. ⌘P로 인쇄화면을 여세요." + hint
    except Exception as e:
        try: subprocess.Popen(["open", "-a", "Microsoft Excel", abs_path], start_new_session=True); return True, f"Excel에서 열었습니다. (자동화 실패: {e})"
        except Exception as e2: return False, f"실행 실패: {e2}"

def _launch_browser_print_dialog(xlsx_path: str) -> None:
    """숨은 iframe에 인쇄 HTML을 넣고 인쇄 대화상자만 연다. 새 탭은 열지 않는다."""
    st.session_state["wl_print_panel"] = False
    abs_path = os.path.abspath(xlsx_path)
    if not os.path.exists(abs_path): st.error("인쇄용 파일이 없습니다."); return
    cache_k, meta_k = f"wl_print_html_cache_{abs_path}", f"wl_print_html_meta_{abs_path}"
    try: mtime = os.path.getmtime(abs_path)
    except OSError: mtime = 0.0
    cached, meta = st.session_state.get(cache_k), st.session_state.get(meta_k) or {}
    cache_ver = "v27"
    if isinstance(cached, str) and cached and meta.get("mtime") == mtime and meta.get("path") == abs_path and meta.get("ver") == cache_ver:
        stamped = cached
        nonce = int(st.session_state.get("wl_print_n", 0)) + 1
        st.session_state["wl_print_n"] = nonce
    else:
        try:
            doc_html = render_worklog_view_html(abs_path, print_mode=True, auto_print=True, scale=1.0)
        except Exception as e: st.error(f"인쇄 문서 준비 실패: {e}"); return
        nonce = int(st.session_state.get("wl_print_n", 0)) + 1
        st.session_state["wl_print_n"] = nonce
        stamped = doc_html.replace("<body>", f'<body data-wl-print="{nonce}">', 1)
        st.session_state[cache_k] = stamped
        st.session_state[meta_k] = {"mtime": mtime, "path": abs_path, "ver": cache_ver}
    _WL_PRINT_LAUNCH(
        key=f"wl_print_launch_{nonce}",
        data={"html": stamped, "n": str(nonce)},
        height=1,
    )
    st.caption("인쇄 창이 안 뜨면 브라우저 인쇄 권한을 허용해 주세요.")

def _open_worklog_print_panel(xlsx_path: str, *, auto: bool = False) -> None:
    st.session_state["wl_print_panel"] = True
    st.session_state["wl_dialog_preview_path"] = os.path.abspath(xlsx_path)
    st.session_state["wl_print_auto_once"] = False

def _render_worklog_print_panel() -> bool:
    if not st.session_state.get("wl_print_panel"): return False
    path = st.session_state.get("wl_dialog_preview_path")
    st.markdown("""<style>.dashboard-filter-sticky, #dashboard-top-shield, #dashboard-sticky-spacer { display: none !important; } section.main .block-container { padding-top: 0.4rem !important; }</style>""", unsafe_allow_html=True)
    top1, top2 = st.columns([1.1, 3])
    with top1:
        if st.button("← 본화면으로", type="primary", width="stretch", key="wl_print_back_home"):
            st.session_state["wl_print_panel"] = False; st.session_state["wl_print_auto_once"] = False; _wl_rerun()
    with top2:
        st.markdown("##### 인쇄 미리보기")
        st.caption("자동 팝업 없음 · 「인쇄하기」만 누르면 인쇄 창이 열립니다.")
    if not path or not os.path.exists(str(path)):
        st.error("인쇄용 파일이 없습니다. 본화면으로 돌아가 다시 시도해 주세요.")
        return True
    path = str(path)
    try:
        print_html = render_worklog_view_html(path, print_mode=True, auto_print=False, scale=1.0)
        _, raw_h = _worklog_sheet_pixel_size(path)
        components.html(print_html, height=min(920, max(480, int(raw_h * _a4_print_fit(*_worklog_sheet_pixel_size(path), path=path)) + 72)), scrolling=True)
    except Exception as e: st.error(f"인쇄 미리보기 표시 실패: {e}")
    b1, b2 = st.columns(2)
    with b1:
        try:
            with open(path, "rb") as f: xbytes = f.read()
            st.download_button("엑셀 다운로드", data=xbytes, file_name=os.path.basename(path) or "일일업무일지.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch", key="wl_print_panel_dl")
        except Exception: pass
    with b2:
        if platform.system() == "Darwin":
            if st.button("Excel로 인쇄", width="stretch", key="wl_print_panel_excel"):
                ok, msg = open_excel_print_preview(path, prefer_print_dialog=True)
                (st.success if ok else st.warning)(msg)
    return True

@st.dialog("원본 엑셀 양식 미리보기", width="large")
def _worklog_form_preview_dialog() -> None:
    path = st.session_state.get("wl_dialog_preview_path")
    if not path or not os.path.exists(str(path)):
        st.error("미리보기 파일을 만들 수 없습니다. 템플릿·입력을 확인해 주세요.")
        return
    path = str(path)
    try:
        scale = _WL_PREVIEW_SCALE
        print_html = render_worklog_view_html(path, print_mode=False, auto_print=False, scale=scale)
        _, frame_h = _scaled_view_frame_size(path, scale)
        components.html(print_html, height=min(900, max(480, int(frame_h))), scrolling=True)
    except Exception as e:
        st.error(f"미리보기 표시 실패: {e}")
        return
    b1, b2, b3 = st.columns(3)
    with b1:
        try:
            with open(path, "rb") as f: xbytes = f.read()
            st.download_button("엑셀 다운로드", data=xbytes, file_name=os.path.basename(path) or "일일업무일지_미리보기.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch", key="wl_dialog_dl_xlsx")
        except Exception: st.caption("엑셀 다운로드를 준비하지 못했습니다.")
    with b2:
        if st.button("인쇄창열기", width="stretch", key="wl_dialog_browser_print"): _launch_browser_print_dialog(path)
    with b3:
        if platform.system() == "Darwin":
            if st.button("Excel 인쇄 화면", width="stretch", key="wl_dialog_open_excel"):
                ok, msg = open_excel_print_preview(path, prefer_print_dialog=True)
                if ok: st.success(msg)
                else: st.warning(msg)


def _prepare_excel_preview(d: date, cells: dict) -> str:
    try: return prepare_print_xlsx(d, cells)
    except Exception: return _build_preview_file(d, cells)

def _clear_wl_cal_nav_keys() -> None:
    """달력 ◀▶오늘 버튼 키는 session_state에 넣으면 안 된다."""
    for k in ("wl_prev_month", "wl_next_month", "wl_today"):
        st.session_state.pop(k, None)


def _render_month_calendar(selected: date, saved: set[str]) -> date | None:
    _clear_wl_cal_nav_keys()
    if "worklog_month" not in st.session_state: st.session_state["worklog_month"] = date(selected.year, selected.month, 1)
    month_anchor: date = st.session_state["worklog_month"]

    st.markdown("""<style>div[data-testid="stPopoverBody"] { max-width: 340px !important; width: 340px !important; padding: 0.55rem 0.6rem 0.65rem !important; } div[data-testid="stPopoverBody"] div[class*="st-key-wl_day_"] button, div[data-testid="stPopoverBody"] div[class*="st-key-wl_prev_month"] button, div[data-testid="stPopoverBody"] div[class*="st-key-wl_next_month"] button, div[data-testid="stPopoverBody"] div[class*="st-key-wl_today"] button { min-height: 2.15rem !important; height: 2.15rem !important; padding: 0 0.2rem !important; font-size: 0.95rem !important; font-weight: 600 !important; line-height: 1.1 !important; border-radius: 7px !important; } div[data-testid="stPopoverBody"] [data-testid="stCaptionContainer"] { font-size: 0.74rem !important; margin-bottom: 0.3rem !important; } div[data-testid="stPopoverBody"] [data-testid="stHorizontalBlock"] { gap: 0.3rem !important; } div[data-testid="stPopoverBody"] [data-testid="column"] { padding: 0 !important; }</style>""", unsafe_allow_html=True)

    nav = st.columns([0.85, 0.85, 2.6, 1.2], gap="small")
    with nav[0]:
        if st.button("◀", key="wl_prev_month", width="stretch"):
            y, m = month_anchor.year, month_anchor.month - 1
            if m < 1: y, m = y - 1, 12
            st.session_state["worklog_month"] = date(y, m, 1)
            _wl_rerun()
    with nav[1]:
        if st.button("▶", key="wl_next_month", width="stretch"):
            y, m = month_anchor.year, month_anchor.month + 1
            if m > 12: y, m = y + 1, 1
            st.session_state["worklog_month"] = date(y, m, 1)
            _wl_rerun()
    with nav[2]:
        st.markdown(f"<div style='text-align:center;font-weight:800;font-size:16px;padding:3px 0;line-height:1.2;color:#0F172A;'>{month_anchor.year}년 {month_anchor.month}월</div>", unsafe_allow_html=True)
    with nav[3]:
        st.button(
            "오늘",
            key="wl_today",
            width="stretch",
            on_click=_on_wl_cal_day,
            args=(date.today().isoformat(),),
        )

    st.caption("• = 저장됨 · 보기 전용 · 날짜를 누르면 그 날을 엽니다")
    weeks = ["월", "화", "수", "목", "금", "토", "일"]
    head = st.columns(7, gap="small")
    for i, w in enumerate(weeks):
        color = "#DC2626" if i == 6 else ("#2563EB" if i == 5 else "#64748B")
        head[i].markdown(f"<div style='text-align:center;font-size:13px;font-weight:700;color:{color};line-height:1;padding:2px 0 4px 0;'>{w}</div>", unsafe_allow_html=True)

    cal = calendar.Calendar(firstweekday=0)
    clicked = None
    for week in cal.monthdayscalendar(month_anchor.year, month_anchor.month):
        cols = st.columns(7, gap="small")
        for i, day in enumerate(week):
            with cols[i]:
                if day == 0:
                    st.markdown("<div style='height:2.15rem'></div>", unsafe_allow_html=True)
                    continue
                d = date(month_anchor.year, month_anchor.month, day)
                is_sel = d == selected
                has = d.isoformat() in saved
                label = f"{day}•" if has else f"{day}"
                st.button(
                    label,
                    key=f"wl_day_{d.isoformat()}_{'s' if has else 'n'}",
                    width="stretch",
                    type="primary" if is_sel else "secondary",
                    on_click=_on_wl_cal_day,
                    args=(d.isoformat(),),
                )
    return clicked



def _worklog_summary_html_body(full_html: str) -> str:
    if "<body>" in full_html:
        return full_html.split("<body>", 1)[1].rsplit("</body>", 1)[0].strip()
    return full_html


def _wl_preview_cell_html(key: str, raw: Any) -> str:
    """엑셀 미리보기 칸 패치용 — workbook_to_html 이스케이프와 동일."""
    text = "" if raw is None else str(raw)
    if text in (_WL_SOFT_BLANK, "\u00a0"):
        text = ""
    if key.startswith("G") and text.strip() == "" and text != "":
        text = ""
    if key.startswith("D") and not text.strip():
        return "&nbsp;"
    return html.escape(text).replace(" ", "&nbsp;").replace("\n", "<br>")


def _wl_preview_patches(cells: dict | None) -> dict[str, str]:
    src, extras = _detach_extra_pages(cells)
    out = {"date": html.escape(str(src.get("date") or ""))}
    for r in WL_CLIENT_ROWS:
        out[f"C{r}"] = _wl_preview_cell_html(f"C{r}", src.get(f"C{r}", ""))
    for r in WL_CONTENT_ROWS:
        out[f"G{r}"] = _wl_preview_cell_html(f"G{r}", src.get(f"G{r}", ""))
        out[f"Y{r}"] = _wl_preview_cell_html(f"Y{r}", src.get(f"Y{r}", ""))
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        out[f"D{r}"] = _wl_preview_cell_html(f"D{r}", src.get(f"D{r}", ""))
    for i, extra in enumerate(extras, start=2):
        pfx = f"p{i}-"
        out[f"{pfx}date"] = html.escape(str(extra.get("date") or src.get("date") or ""))
        for r in WL_CLIENT_ROWS:
            out[f"{pfx}C{r}"] = _wl_preview_cell_html(f"C{r}", extra.get(f"C{r}", ""))
        for r in WL_CONTENT_ROWS:
            out[f"{pfx}G{r}"] = _wl_preview_cell_html(f"G{r}", extra.get(f"G{r}", ""))
            out[f"{pfx}Y{r}"] = _wl_preview_cell_html(f"Y{r}", extra.get(f"Y{r}", ""))
        for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
            out[f"{pfx}D{r}"] = "&nbsp;"
    return out


def _excel_preview_host_html(path: str, *, scale: float | None = None) -> str:
    """엑셀 양식을 유지형 미리보기 호스트에 넣을 HTML (전체 문서/iframe 아님)."""
    s = float(scale if scale is not None else _WL_PREVIEW_SCALE)
    sheet = workbook_to_html(path, include_logo=True, layout_scale=1.0)
    scale_css = f"zoom:{s};" if s != 1 else "zoom:1;"
    fallback = (
        f"@supports not (zoom: 1) {{ .sheet-scale {{ transform:scale({s});"
        f"transform-origin:top left; }} }}"
        if s != 1
        else ""
    )
    return (
        f"<style>{_WL_FONT_FACE_CSS}"
        f" .wrap {{ overflow:visible; width:100%; background:#fff; box-sizing:border-box; }}"
        f" .sheet-scale {{ {scale_css} width:fit-content; }}"
        f" .wl-sheet {{ border-collapse:collapse; table-layout:fixed; font-family:{_WL_FONT_STACK} !important; }}"
        f" .wl-sheet, .wl-sheet td, .wl-sheet tr {{ box-sizing:border-box; font-family:{_WL_FONT_STACK} !important; }}"
        f" .wl-sheet td[data-wl^='G'], .wl-sheet td[data-wl*='-G'] {{"
        f" white-space:pre-wrap !important; overflow:hidden !important;"
        f" word-break:break-all !important; overflow-wrap:anywhere !important; }}"
        f" .wl-sheet td[data-wl^='Y'], .wl-sheet td[data-wl*='-Y'],"
        f" .wl-sheet td[data-wl^='C'], .wl-sheet td[data-wl*='-C'] {{"
        f" white-space:nowrap !important; overflow:hidden !important;"
        f" word-break:keep-all !important; overflow-wrap:normal !important; }}"
        f" .wl-page-break {{ page-break-before:always; break-before:page; }}"
        f" {fallback}</style>"
        f'<div class="wrap"><div class="sheet-scale">{sheet}</div></div>'
    )


def _mount_worklog_live_preview(
    *,
    key: str,
    mode: str,
    html: str,
    rev: str,
    height: int,
    patches: dict[str, str] | None = None,
    page: int = 1,
) -> None:
    _WL_PREVIEW_HOST(
        key=key,
        data={
            "mode": mode,
            "html": html or "",
            "rev": str(rev or ""),
            "height": int(height or 0),
            "patches": patches or None,
            "page": max(1, int(page or 1)),
        },
        default={},
    )


def _render_worklog_summary_block(selected: date, cells: dict) -> None:
    """요약 카드 — 유지형 호스트에 HTML만 갱신 (iframe 재부착 없음)."""
    draft_sig = json.dumps({"cells": cells, "build": _WL_UI_BUILD}, ensure_ascii=False, sort_keys=True)
    sum_sig_k = f"wl_sum_sig_v26_{selected.isoformat()}"
    sum_html_k = f"wl_sum_html_v26_{selected.isoformat()}"
    if st.session_state.get(sum_sig_k) != draft_sig:
        view_html = render_readable_preview_html(selected, cells)
        st.session_state[sum_sig_k] = draft_sig
        st.session_state[sum_html_k] = view_html
    else:
        view_html = st.session_state.get(sum_html_k)
        if not view_html:
            view_html = render_readable_preview_html(selected, cells)
            st.session_state[sum_html_k] = view_html
    body = _worklog_summary_html_body(view_html)
    host_html = (
        f'<style>{_WL_SUMMARY_PREVIEW_CSS}</style>'
        f'<div class="wl-sum-preview">{body}</div>'
    )
    _mount_worklog_live_preview(
        key="wl_sum_host",
        mode="summary",
        html=host_html,
        rev=draft_sig,
        height=820,
    )


def _try_pull_remote_worklog_day(d: date) -> bool:
    """로컬에 없을 때 Drive 동기화 후 파일 존재 여부 확인 (Gist 미사용)."""
    iso = d.isoformat()
    if os.path.isfile(worklog_path(d)):
        return False
    tried_k = f"wl_remote_pull_tried_{iso}"
    if st.session_state.get(tried_k):
        return False
    st.session_state[tried_k] = True
    if _wl_is_streamlit_cloud():
        try:
            from drive_autoload import sync_dashboard_copy_on_boot

            sync_dashboard_copy_on_boot(
                os.path.dirname(WORKLOG_DIR),
                force_refresh=False,
                include_worklog=True,
            )
        except Exception:
            pass
    else:
        try:
            from drive_autoload import sync_worklog_bidirectional

            sync_worklog_bidirectional(WORKLOG_DIR, force=True)
        except Exception:
            pass
    if os.path.isfile(worklog_path(d)):
        _invalidate_saved_dates_cache()
        _invalidate_worklog_presence_cache(d)
        st.session_state.pop(_boot_key(d), None)
        st.session_state.pop(f"wl_open_ctx_{iso}", None)
        st.session_state.pop(f"wl_remote_pull_tried_{iso}", None)
        return True
    return False


def _dashboard_filters_changed_this_run() -> bool:
    """상단 필터가 직전 기록과 다르면 True (시그니처는 아직 갱신하지 않음)."""
    prev = st.session_state.get("_wl_dash_filter_sig")
    if prev is None:
        return False
    return prev != _dashboard_top_filter_sig()


def _prepare_worklog_day_state(selected: date, *, skip_remote_pull: bool = False) -> None:
    """날짜별 위젯 초기화 + 저장 직후 pending 시드 (페이지 rerun 시 1회)."""
    iso = selected.isoformat()
    pending_move = st.session_state.get(f"wl_pending_sync_{iso}")
    moving = bool(st.session_state.get("wl_date_retarget_from") or st.session_state.get("wl_purge_dates"))
    if not skip_remote_pull and not os.path.isfile(worklog_path(selected)) and not isinstance(pending_move, dict) and not moving:
        if _try_pull_remote_worklog_day(selected):
            iso = selected.isoformat()
    open_k = f"wl_open_ctx_{iso}"
    if open_k not in st.session_state:
        pres = detect_worklog_date_presence(selected, include_remote=False)
        st.session_state[open_k] = {"had_local": bool(pres.get("local")), "presence": pres}
    _init_widget_state(selected)
    pending = st.session_state.pop(f"wl_pending_sync_{iso}", None)
    if isinstance(pending, dict):
        if not pending.get("keep_editor"):
            _seed_day_entry_widgets(
                selected,
                pending.get("entries") or [{"client": "", "content": ""}],
                pending.get("next") or "",
                pending.get("notes") or "",
                extra_pages=pending.get("extra_pages"),
            )
        if pending.get("msg"):
            st.session_state[f"wl_flash_save_{iso}"] = {
                "msg": pending["msg"],
                "cloud_err": pending.get("cloud_err"),
            }


def _render_worklog_left_preview(selected: date) -> None:
    """왼쪽 요약/엑셀 — 입력 위젯 실시간 반영 (요약은 soft blank 제거)."""
    draft = _draft_cells_for_left_preview(selected)
    st.markdown("##### 업무일지 보기")
    p1, p2 = st.columns(2)
    with p1:
        do_print = st.button(
            "엑셀 미리보기", width="stretch", key="wl_print_btn",
            help="현재 입력을 왼쪽에 원본 엑셀 양식으로 반영합니다.",
        )
    with p2:
        do_open_print = st.button(
            "인쇄창열기", width="stretch", key="wl_open_print_btn",
            help="브라우저 인쇄 창을 엽니다.", type="primary",
        )

    def _resolve_print_xlsx() -> str:
        cells_dl = _cells_from_widgets(selected)
        sig = json.dumps(cells_dl, ensure_ascii=False, sort_keys=True)
        sig_k = f"wl_print_cells_sig_{selected.isoformat()}"
        path_k = f"wl_print_cells_path_{selected.isoformat()}"
        prev_sig = st.session_state.get(sig_k)
        prev_path = str(st.session_state.get(path_k) or "")
        if prev_sig == sig and prev_path and os.path.isfile(prev_path):
            st.session_state[f"wl_left_excel_path_{selected.isoformat()}"] = prev_path
            return prev_path
        out = os.path.abspath(_prepare_excel_preview(selected, cells_dl))
        if prev_path and prev_path != out:
            st.session_state.pop(f"wl_print_html_cache_{prev_path}", None)
            st.session_state.pop(f"wl_print_html_meta_{prev_path}", None)
        st.session_state.pop(f"wl_print_html_cache_{out}", None)
        st.session_state.pop(f"wl_print_html_meta_{out}", None)
        st.session_state[sig_k] = sig
        st.session_state[path_k] = out
        st.session_state[f"wl_left_excel_path_{selected.isoformat()}"] = out
        return out

    if do_open_print:
        try:
            xlsx_abs = _resolve_print_xlsx()
            _launch_browser_print_dialog(xlsx_abs)
        except Exception as e:
            st.error(f"인쇄 창을 열지 못했습니다: {e}")

    _left_excel_key = f"wl_left_excel_on_{selected.isoformat()}"
    _left_path_key = f"wl_left_excel_path_{selected.isoformat()}"
    if do_print:
        cells_now = _cells_from_widgets(selected)
        try:
            _publish_view_cells(selected, cells_now)
            xlsx_abs = _prepare_excel_preview(selected, cells_now)
            st.session_state[_left_excel_key] = True
            st.session_state[_left_path_key] = xlsx_abs
            st.session_state["wl_dialog_preview_path"] = xlsx_abs
            form_sig = json.dumps(cells_now, ensure_ascii=False)
            st.session_state[f"wl_form_sig_v14_{selected.isoformat()}"] = form_sig
            st.session_state["_wl_force_form_sig"] = form_sig
            st.session_state[f"wl_left_excel_rebuild_{selected.isoformat()}"] = True
            st.session_state.pop(f"wl_left_excel_html_v27_{selected.isoformat()}", None)
            draft = dict(cells_now)
        except Exception as e:
            st.error(f"미리보기 생성 중 오류가 발생했습니다: {e}")
            st.session_state[_left_excel_key] = False

    _show_excel_left = bool(st.session_state.get(_left_excel_key))
    if _show_excel_left:
        sw1, sw2 = st.columns([1, 1])
        with sw1:
            st.caption("원본 엑셀 양식 적용 중")
        with sw2:
            if st.button("요약 보기로", width="stretch", key=f"wl_left_to_summary_{selected.isoformat()}"):
                try:
                    _publish_view_cells(selected, _cells_from_widgets(selected))
                except Exception:
                    pass
                st.session_state[_left_excel_key] = False
        xlsx_left = st.session_state.get(_left_path_key) or ""
        if xlsx_left and os.path.exists(str(xlsx_left)):
            try:
                cells_view = _draft_cells_for_left_preview(selected)
                html_k = f"wl_left_excel_html_v27_{selected.isoformat()}"
                h_k = f"wl_left_excel_h_v27_{selected.isoformat()}"
                skel_k = f"wl_left_excel_skel_v27_{selected.isoformat()}"
                rebuild_k = f"wl_left_excel_rebuild_{selected.isoformat()}"
                scale_l = _WL_PREVIEW_SCALE
                page_n = _page_idx_for(selected.isoformat()) + 1
                need_skel = bool(st.session_state.pop(rebuild_k, None)) or not st.session_state.get(html_k)
                cached_html = str(st.session_state.get(html_k) or "")
                if (
                    not need_skel
                    and page_n > 1
                    and cached_html
                    and f'data-wl-page="{page_n}"' not in cached_html
                    and f'data-wl="p{page_n}-' not in cached_html
                ):
                    need_skel = True
                if need_skel:
                    xlsx_left = _prepare_excel_preview(selected, cells_view)
                    st.session_state[_left_path_key] = xlsx_left
                    inline = _excel_preview_host_html(str(xlsx_left), scale=scale_l)
                    _, fh = _scaled_view_frame_size(str(xlsx_left), scale_l)
                    st.session_state[html_k] = inline
                    st.session_state[h_k] = fh
                    st.session_state[skel_k] = str(int(st.session_state.get(skel_k) or 0) + 1)
                else:
                    inline = st.session_state.get(html_k) or ""
                    fh = st.session_state.get(h_k)
                    if not inline:
                        inline = _excel_preview_host_html(str(xlsx_left), scale=scale_l)
                        _, fh = _scaled_view_frame_size(str(xlsx_left), scale_l)
                        st.session_state[html_k] = inline
                        st.session_state[h_k] = fh
                _mount_worklog_live_preview(
                    key="wl_excel_host",
                    mode="excel",
                    html=inline,
                    rev=str(st.session_state.get(skel_k) or "1"),
                    height=min(1100, max(560, int(fh or 600))),
                    patches=_wl_preview_patches(cells_view),
                    page=page_n,
                )
                if st.button("크게 보기", width="stretch", key=f"wl_left_excel_big_{selected.isoformat()}"):
                    st.session_state["wl_dialog_preview_path"] = str(xlsx_left)
                    _worklog_form_preview_dialog()
            except Exception as e:
                st.warning(f"엑셀 양식 표시 실패: {e}")
                st.session_state[_left_excel_key] = False
        else:
            st.info("엑셀 미리보기 파일이 없습니다. 다시 「엑셀 미리보기」를 눌러 주세요.")
    else:
        try:
            _render_worklog_summary_block(selected, draft)
        except Exception as e:
            if _wl_quiet_ui():
                st.info("업무일지 요약을 표시하지 못했습니다. 입력 후 다시 확인해 주세요.")
            else:
                st.error(f"요약 보기 오류: {e}")


def _sync_left_preview_snapshot(d: date, *, focus_sig: str | None = None) -> None:
    """왼쪽 요약 스냅샷만 갱신. fragment rerun은 호출하지 않음 (위젯 콜백 rerun에 맡김)."""
    iso = d.isoformat()
    if focus_sig is not None:
        prev = st.session_state.get(f"wl_left_focus_sig_{iso}")
        if prev == focus_sig:
            return
        st.session_state[f"wl_left_focus_sig_{iso}"] = focus_sig
    try:
        _publish_view_cells(d, _cells_from_widgets(d))
    except Exception:
        pass


def _request_left_preview_refresh(d: date, *, focus_sig: str | None = None) -> None:
    """저장·엑셀 미리보기 등 — 왼쪽 요약 스냅샷 갱신 + fragment rerun 예약."""
    _sync_left_preview_snapshot(d, focus_sig=focus_sig)
    iso = d.isoformat()
    # 엑셀 미리보기 모드면 iframe 갱신 생략 (요약 보기로 돌아올 때 publish 반영)
    if st.session_state.get(f"wl_left_excel_on_{iso}"):
        return
    st.session_state["wl_need_left_refresh"] = True


def _wl_finish_edit_fragment() -> None:
    """편집 fragment 마무리 — 왼쪽 요약 갱신 시 fragment만 rerun (전체 앱·sync 생략)."""
    if st.session_state.pop("wl_need_left_refresh", None):
        _wl_rerun()



def _render_worklog_date_toolbar(selected: date) -> None:
    """날짜칸·달력·삭제. 본문(거래처/내용/비고)과 같은 폭·왼쪽선에 맞춘다."""
    saved = _saved_dates_for_calendar()
    with st.container(key="wl_date_bar", horizontal=True, gap="small", vertical_alignment="bottom"):
        if "wl_date_pick" not in st.session_state:
            _set_wl_date_pick(selected)
        st.date_input(
            "업무일지 날짜",
            format="YYYY/MM/DD",
            key="wl_date_pick",
            on_change=_on_wl_date_pick_change,
            help="기본은 그 날 저장본을 엽니다. 「날짜변경」을 켠 뒤 날짜를 고르면 지금 입력 중인 내용이 그 날로 옮겨집니다.",
            width="stretch",
        )
        st.session_state["_wl_date_pick_live"] = True
        with st.popover("📅 달력", width="content", key="wl_cal_pop"):
            _render_month_calendar(selected, saved)
        _move_on = _date_move_mode_on()
        st.button(
            "날짜변경",
            width="content",
            key="wl_date_move_btn",
            type="primary" if _move_on else "secondary",
            on_click=_on_toggle_date_move_mode,
            help="켜 둔 뒤 날짜를 고르면 지금 편집중인 글의 날짜만 바뀝니다. 저장해야 확정됩니다.",
        )
        if st.button("삭제", width="content", key="wl_del_open_btn"):
            _pin_worklog_scroll()
            st.session_state["wl_del_confirm_open"] = True
    if _date_move_mode_on():
        st.caption("날짜변경 켜짐 · 날짜를 고르면 지금 입력 중인 내용이 그 날로 옮겨집니다. 저장해야 확정됩니다.")
    # 삭제 확인 — 팝오버 대신 세션 상태 기반 인라인 UI. 확정/취소 후 확실히 사라진다.
    # (st.popover 를 코드로 닫으면 프론트가 다시 열어버려 '깜박→부활'하는 문제를 회피)
    if st.session_state.get("wl_del_confirm_open"):
        with st.container(border=True):
            st.caption(f"🗑️ {selected.isoformat()} 일지만 삭제 · 다른 날짜는 그대로 둡니다")
            _dc1, _dc2 = st.columns([1, 1], gap="small")
            _dc1.button("확정 삭제", type="primary", width="stretch", key="wl_del_day_yes", on_click=_on_confirm_delete_day)
            _dc2.button("취소", width="stretch", key="wl_del_day_no", on_click=_on_cancel_delete_day)
    _date_err = st.session_state.pop("wl_date_err", None)
    if _date_err:
        st.error(_date_err)


def _render_worklog_input_panel(selected: date) -> None:
    """오른쪽 게이지+입력. 칸 이동 시 published 스냅샷 갱신(동일 fragment rerun)."""
    _run_pending_worklog_date_change()
    st.session_state.pop("wl_need_app_rerun", None)
    selected = st.session_state.get("worklog_selected") or selected

    try:
        iso_g = selected.isoformat()
        _c, _g, _y = _sheet_lines_from_widgets_at(iso_g, _page_idx_for(iso_g))
        _gauge_usage = _sheet_row_usage(_c, _g, _y)
    except Exception:
        _gauge_usage = _sheet_row_usage([], [], [])
    col_gauge, col_input = st.columns([0.14, 1], gap="small")
    with col_gauge:
            st.markdown("<div style='height:0.35rem'></div>", unsafe_allow_html=True)
            _render_row_remain_gauge(_gauge_usage, height_px=980)

    with col_input:
            _render_worklog_date_toolbar(selected)
            iso = selected.isoformat()
            ek = _entries_key(selected)
            if ek not in st.session_state or not st.session_state[ek]: st.session_state[ek] = [_empty_sheet_entry()]

            def _seed_entry_widgets(entries_list: list[dict], next_txt: str, notes_txt: str) -> None:
                _seed_day_entry_widgets(selected, entries_list, next_txt, notes_txt)

            flash = st.session_state.pop(f"wl_flash_save_{iso}", None)
            if isinstance(flash, dict) and flash.get("msg"):
                if flash.get("cloud_err"):
                    st.warning(flash["msg"])
                else:
                    st.success(flash["msg"])

            st.session_state.pop(f"wl_do_del_{iso}", None)
            if st.session_state.pop(f"wl_do_add_{iso}", None):
                _add_worklog_input_page(iso)
            entries = list(st.session_state[ek] or [_empty_sheet_entry()])
            sheet = entries[0] if entries else _empty_sheet_entry()
            ck = f"wl_ent_c_{iso}_0"
            if ck not in st.session_state: st.session_state[ck] = sheet.get("client") or ""
            if int(st.session_state.get(_entry_client_count_key(iso, 0), 0) or 0) <= 0:
                cl0 = sheet.get("client_lines")
                if isinstance(cl0, list) and cl0: _seed_entry_clients(iso, 0, cl0)
                else: _seed_entry_clients(iso, 0, sheet.get("client") or "")
            if int(st.session_state.get(_entry_line_count_key(iso, 0), 0) or 0) <= 0:
                lines0 = sheet.get("lines")
                if isinstance(lines0, list): _apply_entry_lines(iso, 0, [str(x or "") for x in lines0], remount_comp=True)
                else: _seed_entry_lines(iso, 0, sheet.get("content") or "")
            if int(st.session_state.get(_entry_remark_count_key(iso, 0), 0) or 0) <= 0:
                rl0 = sheet.get("remark_lines")
                if isinstance(rl0, list) and rl0: _seed_entry_remarks(iso, 0, rl0)
                else: _seed_entry_remarks(iso, 0, sheet.get("remarks") or "")
            st.session_state[f"wl_entry_count_{iso}"] = 1
            nk = f"wl_next_area_{iso}"
            ok = f"wl_notes_area_{iso}"
            if nk not in st.session_state: st.session_state[nk] = st.session_state.get(_next_key(selected), "")
            if ok not in st.session_state: st.session_state[ok] = st.session_state.get(_notes_key(selected), "")

            def _wl_entry_editor():
                d = st.session_state.get("worklog_selected") or selected
                iso2 = d.isoformat()
                max_u = _content_line_units()
                # 저장은 2단계: 버튼 → flush 후 다음 런에서 실제 저장 (직전 입력 누락/구값 복원 방지)
                do_save = bool(st.session_state.pop(f"wl_do_save_{iso2}", None))
                if st.session_state.pop(f"wl_do_add_{iso2}", None):
                    _add_worklog_input_page(iso2)
                pi = _page_idx_for(iso2)

                _force_open = st.session_state.pop(f"wl_force_expand_{iso2}", None)
                if isinstance(_force_open, int): st.session_state[f"wl_exp_{iso2}_{_force_open}"] = True

                del_ln = st.session_state.pop(f"wl_do_del_ln_{iso2}", None)
                if isinstance(del_ln, (list, tuple)) and len(del_ln) == 2:
                    dei, dlj = int(del_ln[0]), int(del_ln[1])
                    cur = _lines_from_entry_widgets(iso2, dei, keep_trailing_empty=False)
                    if 0 <= dlj < len(cur): cur.pop(dlj)
                    if not cur: cur = [""]
                    _apply_entry_lines(iso2, dei, cur, focus_j=min(dlj, max(len(cur) - 1, 0)))

                ins_ln = st.session_state.pop(f"wl_do_insert_ln_{iso2}", None)
                if isinstance(ins_ln, (list, tuple)) and len(ins_ln) == 2:
                    _insert_line_after(iso2, int(ins_ln[0]), int(ins_ln[1]))

                del_cl = st.session_state.pop(f"wl_do_del_cl_{iso2}", None)
                if isinstance(del_cl, (list, tuple)) and len(del_cl) == 2:
                    dei, dlj = int(del_cl[0]), int(del_cl[1])
                    cur = _clients_from_widgets(iso2, dei, keep_trailing_empty=False)
                    if 0 <= dlj < len(cur): cur.pop(dlj)
                    if not cur: cur = [""]
                    _apply_entry_clients(iso2, dei, cur, focus_j=min(dlj, max(len(cur) - 1, 0)))

                ins_cl = st.session_state.pop(f"wl_do_insert_cl_{iso2}", None)
                if isinstance(ins_cl, (list, tuple)) and len(ins_cl) == 2:
                    _insert_client_after(iso2, int(ins_cl[0]), int(ins_cl[1]))

                ent_req = st.session_state.pop(f"wl_do_enter_cell_{iso2}", None)
                if isinstance(ent_req, dict):
                    _commit_enter_on_cell(str(ent_req.get("kind") or ""), iso2, int(ent_req.get("ei") or 0), int(ent_req.get("lj") or 0), str(ent_req.get("v") or ""))

                def _on_enter_trigger():
                    hook = st.session_state.get("wl_enter_hook_nav") or st.session_state.get(f"wl_enter_hook_{iso2}") or {}
                    payload = str(hook.get("enter") or "") if isinstance(hook, dict) else ""
                    if not payload: return
                    done_k = f"wl_enter_done_{iso2}"
                    try:
                        obj = json.loads(payload)
                        key, val = str(obj.get("key") or ""), str(obj.get("v") or "")
                    except Exception:
                        key, val = payload.split(":", 1)[0], ""
                    sig = f"{key}\0{val}"
                    if st.session_state.get(done_k) == sig: return
                    st.session_state[done_k] = sig
                    m = re.match(r"^(wl_ent_ln|wl_ent_cl|wl_ent_rm)_(\d{4}-\d{2}-\d{2})_(\d+)_(\d+)(?:_g\d+)?$", key)
                    if not m or m.group(2) != iso2: return
                    if m.group(1) in ("wl_ent_ln", "wl_ent_cl", "wl_ent_rm"):
                        return
                    st.session_state[f"wl_do_enter_cell_{iso2}"] = {"kind": m.group(1), "ei": int(m.group(3)), "lj": int(m.group(4)), "v": val}

                def _on_focus_trigger():
                    hook = st.session_state.get("wl_enter_hook_nav") or st.session_state.get(f"wl_enter_hook_{iso2}") or {}
                    if not isinstance(hook, dict):
                        return
                    fk = str(hook.get("focus") or "")
                    if fk.startswith("wl_next_area_") or fk.startswith("wl_notes_area_"):
                        st.session_state["wl_active_cell_key"] = fk
                        return
                    if fk.startswith("wl_ent_ln_") or fk.startswith("wl_ent_cl_") or fk.startswith("wl_ent_rm_"):
                        st.session_state["wl_active_cell_key"] = fk
                        st.session_state[f"wl_focus_ln_{iso2}"] = fk

                def _on_caret_trigger():
                    hook = st.session_state.get("wl_enter_hook_nav") or st.session_state.get(f"wl_enter_hook_{iso2}") or {}
                    if not isinstance(hook, dict):
                        return
                    raw = hook.get("caret")
                    if not raw:
                        return
                    try:
                        obj = json.loads(str(raw))
                        fk, s, e = str(obj.get("key") or ""), int(obj.get("s") or 0), int(obj.get("e") or int(obj.get("s") or 0))
                    except Exception:
                        return
                    if fk.startswith("wl_next_area_") or fk.startswith("wl_notes_area_"):
                        st.session_state["wl_active_cell_key"] = fk
                        st.session_state["wl_active_cell_sel"] = (s, e)
                        return
                    if fk.startswith("wl_ent_ln_") or fk.startswith("wl_ent_cl_") or fk.startswith("wl_ent_rm_"):
                        st.session_state["wl_active_cell_key"] = fk
                        st.session_state[f"wl_focus_ln_{iso2}"] = fk
                        st.session_state["wl_active_cell_sel"] = (s, e)
                        st.session_state[f"wl_focus_caret_{iso2}"] = s

                _cu = _client_line_units()
                _ru = _remark_line_units()
                if int(st.session_state.get(_entry_client_count_key(iso2, pi), 0) or 0) <= 0:
                    if pi <= 0:
                        stored_e = st.session_state.get(_entries_key(d)) or []
                        if stored_e and isinstance(stored_e[0].get("client_lines"), list):
                            _seed_entry_clients(iso2, 0, stored_e[0].get("client_lines") or [""])
                        else:
                            _seed_entry_clients(iso2, 0, str(st.session_state.get(f"wl_ent_c_{iso2}_0", "") or ""))
                    else:
                        _seed_entry_clients(iso2, pi, str(st.session_state.get(f"wl_ent_c_{iso2}_{pi}", "") or ""))
                if int(st.session_state.get(_entry_line_count_key(iso2, pi), 0) or 0) <= 0:
                    if pi <= 0:
                        lines0 = None
                        stored_e = st.session_state.get(_entries_key(d)) or []
                        if stored_e and isinstance(stored_e[0].get("lines"), list): lines0 = stored_e[0].get("lines")
                        if isinstance(lines0, list): _apply_entry_lines(iso2, 0, [str(x or "") for x in lines0], remount_comp=True)
                        else: _seed_entry_lines(iso2, 0, str(st.session_state.get(f"wl_ent_t_{iso2}_0", "") or ""))
                    else:
                        _seed_entry_lines(iso2, pi, str(st.session_state.get(f"wl_ent_t_{iso2}_{pi}", "") or ""))
                if int(st.session_state.get(_entry_remark_count_key(iso2, pi), 0) or 0) <= 0:
                    if pi <= 0:
                        stored_e = st.session_state.get(_entries_key(d)) or []
                        if stored_e and isinstance(stored_e[0].get("remark_lines"), list):
                            _seed_entry_remarks(iso2, 0, stored_e[0].get("remark_lines") or [""])
                        else:
                            _seed_entry_remarks(iso2, 0, str(st.session_state.get(f"wl_ent_r_{iso2}_0", "") or (stored_e[0].get("remarks") if stored_e else "") or ""))
                    else:
                        _seed_entry_remarks(iso2, pi, str(st.session_state.get(f"wl_ent_r_{iso2}_{pi}", "") or ""))

                _cw, _gw, _rw = _input_cell_px("client"), _input_cell_px("content"), _input_cell_px("remark")
                _side = max(int(_cw), int(_rw))
                st.markdown(
                    f"""<style>
                    div[class*="st-key-wl_entry_sheet"],
                    div[class*="st-key-wl_entry_sheet"] > div {{
                      width: 100% !important; max-width: 100% !important;
                    }}
                    div[class*="st-key-wl_entry_sheet"] [data-testid="stHorizontalBlock"] {{
                      display: flex !important;
                      flex-wrap: nowrap !important;
                      justify-content: flex-start !important;
                      align-items: stretch !important;
                      gap: 0 !important;
                      width: 100% !important;
                      max-width: 100% !important;
                    }}
                    div[class*="st-key-wl_entry_sheet"] [data-testid="stHorizontalBlock"] > div:nth-child(1),
                    div[class*="st-key-wl_entry_sheet"] [data-testid="stColumn"]:nth-child(1),
                    div[class*="st-key-wl_entry_sheet"] [data-testid="column"]:nth-child(1),
                    div[class*="st-key-wl_entry_sheet"] [data-testid="stHorizontalBlock"] > div:nth-child(3),
                    div[class*="st-key-wl_entry_sheet"] [data-testid="stColumn"]:nth-child(3),
                    div[class*="st-key-wl_entry_sheet"] [data-testid="column"]:nth-child(3) {{
                      width: {_side}px !important; min-width: {_side}px !important; max-width: {_side}px !important;
                      flex: 0 0 {_side}px !important; padding: 0 !important;
                    }}
                    div[class*="st-key-wl_entry_sheet"] [data-testid="stHorizontalBlock"] > div:nth-child(2),
                    div[class*="st-key-wl_entry_sheet"] [data-testid="stColumn"]:nth-child(2),
                    div[class*="st-key-wl_entry_sheet"] [data-testid="column"]:nth-child(2) {{
                      flex: 1 1 auto !important; width: auto !important;
                      min-width: {_gw}px !important; max-width: none !important; padding: 0 !important;
                    }}
                    div[class*="st-key-wl_clients_comp_"], div[class*="st-key-wl_lines_comp_"], div[class*="st-key-wl_remarks_comp_"] {{ width: 100% !important; max-width: 100% !important; }}
                    div[class*="st-key-wl_clients_comp_"] .wl-lines, div[class*="st-key-wl_lines_comp_"] .wl-lines, div[class*="st-key-wl_remarks_comp_"] .wl-lines {{ margin: 0; width: 100% !important; max-width: 100% !important; border-radius: 0; }}
                    div[class*="st-key-wl_lines_comp_"] > div, div[class*="st-key-wl_remarks_comp_"] > div, div[class*="st-key-wl_clients_comp_"] > div {{ width: 100% !important; }}
                    div[class*="st-key-wl_page_bar_"] {{ margin: 0.15rem 0 0.45rem 0 !important; }}
                    div[class*="st-key-wl_page_btn_"] button, div[class*="st-key-wl_page_plus_"] button {{
                      min-height: 2.4rem !important; height: 2.4rem !important;
                      padding: 0 0.85rem !important; font-size: 0.875rem !important; font-weight: 600 !important;
                      border-radius: 0.5rem !important; line-height: 1 !important;
                    }}
                    div[class*="st-key-wl_page_plus_"] button {{ min-width: 2.4rem !important; padding: 0 0.7rem !important; }}
                    </style>""",
                    unsafe_allow_html=True,
                )

                n_pages = _page_count_for(iso2)
                with st.container(key=f"wl_page_bar_{iso2}", horizontal=True, gap="small"):
                    for _pi in range(n_pages):
                        st.button(
                            f"{_pi + 1}페이지",
                            key=f"wl_page_btn_{iso2}_{_pi}",
                            type="primary" if _pi == pi else "secondary",
                            width="content",
                            on_click=_queue_worklog_page,
                            args=(iso2, _pi),
                        )
                    st.button(
                        "＋",
                        key=f"wl_page_plus_{iso2}",
                        width="content",
                        disabled=n_pages >= WL_MAX_PAGES,
                        on_click=_queue_worklog_add,
                        args=(iso2,),
                        help="같은 형식의 입력 페이지를 추가합니다.",
                    )

                # 입력만 가로 여백을 채운다. 줄바꿈·인쇄미리보기는 원본 14pt 칸 폭.
                with st.container(key="wl_entry_sheet"):
                    col_client, col_content, col_remark = st.columns([_side, _gw, _side], gap="small")

                    with col_client:
                        st.markdown(_wl_col_limit_label("거래처", _cu), unsafe_allow_html=True)
                        _mount_entry_client_editor(iso2, pi, _cu)
                    with col_content:
                        st.markdown(_wl_col_limit_label("내용", max_u), unsafe_allow_html=True)
                        _mount_entry_lines_editor(iso2, pi, max_u)
                    with col_remark:
                        st.markdown(_wl_col_limit_label("비고", _ru), unsafe_allow_html=True)
                        _mount_entry_remark_editor(iso2, pi, _ru)

                if pi == 0:
                    st.markdown("<div style='font-size:12px;font-weight:700;color:#334155;margin:12px 0 4px;'>익일업무 <span style='font-weight:500;color:#94A3B8;'>(줄바꿈 = 항목 구분 · Enter=다음 줄)</span></div>", unsafe_allow_html=True)
                    st.markdown(
                        f"""<style>
                        div[class*="st-key-wl_next_area_"] textarea,
                        div[class*="st-key-wl_notes_area_"] textarea {{
                          font-family: {_WL_FONT_STACK} !important;
                          font-size: 11pt !important;
                          line-height: 1.45 !important;
                        }}
                        </style>""",
                        unsafe_allow_html=True,
                    )
                    st.text_area("익일업무", key=f"wl_next_area_{iso2}", label_visibility="collapsed", height=110)
                    st.markdown("<div style='font-size:12px;font-weight:700;color:#334155;margin:12px 0 4px;'>특 이 사 항 <span style='font-weight:500;color:#94A3B8;'>(줄바꿈 = 항목 구분 · Enter=다음 줄)</span></div>", unsafe_allow_html=True)
                    st.text_area("특이사항", key=f"wl_notes_area_{iso2}", label_visibility="collapsed", height=100)

                st.button(
                    "저장",
                    type="primary",
                    width="stretch",
                    key=f"wl_save_btn_{iso2}",
                    on_click=_queue_worklog_save,
                    args=(iso2,),
                )
                if do_save:
                    try:
                        source_d = d
                        picked = st.session_state.get("wl_date_pick")
                        retarget_from = st.session_state.get("wl_date_retarget_from")
                        if isinstance(retarget_from, str) and retarget_from:
                            try:
                                source_d = date.fromisoformat(retarget_from)
                            except ValueError:
                                source_d = d
                        if isinstance(picked, date) and picked != d:
                            try_retarget_worklog_editor_date(d, picked)
                            d = st.session_state.get("worklog_selected") or picked
                        target_d = d
                        pack_d = d
                        pack_iso = pack_d.isoformat()
                        clients_now, contents_now, remarks_now = _sheet_lines_from_widgets(pack_iso)
                        next_txt = str(st.session_state.get(f"wl_next_area_{pack_iso}", "") or st.session_state.get(_next_key(pack_d), "") or "")
                        notes_txt = str(st.session_state.get(f"wl_notes_area_{pack_iso}", "") or st.session_state.get(_notes_key(pack_d), "") or "")
                        cells = _pack_sheet_to_cells(
                            target_d, clients_now, contents_now, remarks_now,
                            _textarea_lines(next_txt),
                            _textarea_lines(notes_txt),
                        )
                        extra_cells = _extra_page_cells_from_widgets(pack_d)
                        extra_date = format_worklog_date(target_d)
                        for extra in extra_cells:
                            extra["date"] = extra_date
                        cells = _collapse_empty_worklog_pages(_attach_extra_pages(cells, extra_cells))
                        extra_cells = _detach_extra_pages(cells)[1]
                        path = commit_worklog_date_save(source_d, target_d, cells)
                        d = target_d
                        iso2 = d.isoformat()
                        st.session_state[f"wl_saved_ok_{iso2}"] = True
                        ctx = dict(st.session_state.get(f"wl_open_ctx_{iso2}") or {})
                        ctx["had_local"] = True
                        st.session_state[f"wl_open_ctx_{iso2}"] = ctx
                        _publish_view_cells(d, cells)
                        seed_entries = [_sheet_entry_from_lines(clients_now, contents_now, remarks_now)]
                        seed_extras = [_sheet_entry_from_cells(x) for x in extra_cells]
                        st.session_state[_entries_key(d)] = seed_entries
                        st.session_state[_next_key(d)] = next_txt
                        st.session_state[_notes_key(d)] = notes_txt
                        arch = (st.session_state.get("wl_last_archive_path") or "")
                        arch_target = st.session_state.get("wl_last_archive_target") or describe_worklog_archive_target(d)
                        drv = st.session_state.get("wl_last_drive_path") or ""
                        mdrv = st.session_state.get("wl_last_drive_month_path") or ""
                        gist = st.session_state.get("wl_last_cloud_gist") or ""
                        cerr = st.session_state.get("wl_last_cloud_err") or ""
                        drv_cf = st.session_state.get("wl_last_drive_conflict") or ""
                        msg = f"저장 완료: {os.path.basename(path)}"
                        if arch_target and not _wl_quiet_ui():
                            msg += f" · {arch_target}"
                        elif arch and not _wl_quiet_ui():
                            _sh = st.session_state.get("wl_last_archive_sheet") or worklog_archive_sheet_title(d)
                            msg += f" · 일지/{d.year}/{os.path.basename(arch)}#{_sh}"
                        if drv:
                            msg += " · Drive"
                        if mdrv and not _wl_quiet_ui():
                            msg += f" · Drive일지/{d.year}"
                        if gist:
                            msg += f" · Cloud OK (gist `{gist}`)"
                        elif cerr == "duplicate_date":
                            msg += " · Cloud: 선입력본 유지(덮어쓰기 안 함)"
                        elif cerr:
                            msg += f" · Cloud 실패: {cerr}"
                        elif not _wl_quiet_ui():
                            msg += " · Cloud미연동: secrets에 github_token"
                        if drv_cf and not drv:
                            msg += f" · Drive: 선입력본 유지({drv_cf})"
                        st.session_state[f"wl_pending_sync_{iso2}"] = {
                            "entries": seed_entries,
                            "extra_pages": seed_extras,
                            "next": next_txt,
                            "notes": notes_txt,
                            "msg": msg,
                            "cloud_err": cerr,
                        }
                        # 저장 직후 강제 pull은 구 원격본이 로컬을 덮을 수 있음 — push는 save_worklog_cells에서 이미 함
                        st.session_state["_wl_drive_sync_ts"] = time.time()
                        _pin_worklog_scroll()
                        _wl_rerun()
                    except WorklogSaveBlockedError as e:
                        st.error(str(e))
                    except Exception as e:
                        if _wl_quiet_ui():
                            st.error("저장에 실패했습니다. 입력 내용을 확인한 뒤 다시 시도해 주세요.")
                        else:
                            st.error(f"저장 실패: {e}")

                focus_key = st.session_state.pop(f"wl_focus_ln_{iso2}", None)
                focus_caret = st.session_state.pop(f"wl_focus_caret_{iso2}", None)
                if isinstance(focus_key, str) and (focus_key.startswith("wl_ent_ln_") or focus_key.startswith("wl_ent_cl_") or focus_key.startswith("wl_ent_rm_")):
                    try:
                        _m = re.match(r"^wl_ent_(?:ln|cl|rm)_\d{4}-\d{2}-\d{2}_(\d+)_", focus_key)
                        if _m: st.session_state[f"wl_exp_{iso2}_{int(_m.group(1))}"] = True
                    except Exception: pass
                else:
                    # 익일업무/특이사항 등은 포커스 강제 복원 금지
                    focus_key = None
                    focus_caret = None

                # Enter 줄바꿈만 서버로. focus/caret trigger는 클릭마다 rerun·ERROR를 만들어 버튼을 죽인다.
                _WL_ENTER_HOOK(
                    key="wl_enter_hook_nav",
                    data={"iso": iso2, "focus_key": focus_key if isinstance(focus_key, str) else "", "focus_caret": (int(focus_caret) if isinstance(focus_caret, (int, float)) else ""), "client_max_u": _client_line_units(), "content_max_u": _content_line_units(), "remark_max_u": _remark_line_units()},
                    on_enter_change=_on_enter_trigger,
                    height=1,
                )
            _wl_entry_editor()


def _dashboard_top_filter_sig() -> tuple:
    """상단 고정바 담당자·거래처·품목·기간 시그니처 (app.py 키와 동일)."""
    return (
        st.session_state.get("dash_filter_staff_sb_v33")
        or st.session_state.get("dash_filter_staff_sb_v32"),
        st.session_state.get("dash_filter_client_sb_v33")
        or st.session_state.get("dash_filter_client_sb_v32"),
        st.session_state.get("dash_filter_item_sb_v33")
        or st.session_state.get("dash_filter_item_sb_v32"),
        st.session_state.get("dash_filter_start"),
        st.session_state.get("dash_filter_end"),
    )


def _maybe_sync_worklog_remote() -> None:
    """페이지 전체 rerun 시에만 Drive/Gist 동기화 (fragment 입력 rerun 제외).

    상단 담당자·거래처·품목 변경으로 인한 전체 rerun에서는 네트워크 sync를 건너뛰어
    필터 조작 시 로딩 스파이크를 막는다. (강제 sync 버튼은 제외)
    """
    if st.session_state.pop("wl_skip_sync_once", None):
        return
    _force = bool(st.session_state.pop("_wl_drive_sync_force", None))
    _filt = _dashboard_top_filter_sig()
    _prev_filt = st.session_state.get("_wl_dash_filter_sig")
    if not _force and _prev_filt is not None and _prev_filt != _filt:
        st.session_state["_wl_dash_filter_sig"] = _filt
        return
    st.session_state["_wl_dash_filter_sig"] = _filt
    try:
        from drive_autoload import sync_worklog_bidirectional

        _now = time.time()
        _prev = float(st.session_state.get("_wl_drive_sync_ts") or 0)
        _on_cloud = _wl_is_streamlit_cloud()
        # 로컬 Mac: 필터 rerun이 잦아 간격을 넉넉히 (강제 버튼은 즉시)
        _sync_iv = 20 if _on_cloud else 300
        if not (_force or (_now - _prev >= _sync_iv)):
            return
        st.session_state["_wl_drive_sync_ts"] = _now
        _wl_sync: dict = {"ok": True, "skipped": True, "copied": [], "conflicts": []}
        if not _on_cloud:
            _wl_sync = sync_worklog_bidirectional(WORKLOG_DIR, force=_force)
        elif _force:
            try:
                from drive_autoload import sync_dashboard_copy_on_boot

                _wl_sync = sync_dashboard_copy_on_boot(
                    os.path.dirname(WORKLOG_DIR),
                    force_refresh=True,
                    include_worklog=True,
                )
            except Exception as _dre:
                _wl_sync = {"ok": False, "error": str(_dre), "copied": [], "conflicts": []}
        _remote_sync: dict = {"ok": True, "skipped": True, "copied": [], "conflicts": []}
        st.session_state["_wl_last_wl_sync"] = _wl_sync
        st.session_state["_wl_last_remote_sync"] = _remote_sync
        _conflicts: list[str] = []
        for _src in (_wl_sync, _remote_sync):
            if isinstance(_src, dict):
                for _c in (_src.get("conflicts") or []):
                    if _c not in _conflicts:
                        _conflicts.append(_c)
        if _conflicts:
            st.session_state["_wl_sync_conflicts"] = _conflicts
        _copied_n = len((_wl_sync or {}).get("copied") or []) + len((_remote_sync or {}).get("copied") or [])
        if _copied_n:
            _invalidate_saved_dates_cache()
            try:
                from worklog_remote_sync import invalidate_gist_days_cache
                invalidate_gist_days_cache()
            except Exception:
                pass
    except Exception:
        pass


def _render_worklog_sync_ui() -> None:
    """동기화 결과·충돌 안내 (fragment 밖). dev 모드(?dev=1)에서만 표시."""
    if not is_dev_mode():
        return
    try:
        from drive_remote_fetch import drive_remote_configured

        _on_cloud = _wl_is_streamlit_cloud()
        _wl_sync = st.session_state.get("_wl_last_wl_sync") or {}
        _copied_n = len((_wl_sync or {}).get("copied") or [])
        if _copied_n:
            st.caption(f"일지 동기화 · {_copied_n}개" + (" (Drive)" if _on_cloud else " (Drive/Cloud)"))
        if _on_cloud:
            if st.button("↻ Drive에서 일지 가져오기", key="wl_drive_pull_btn", width="stretch"):
                st.session_state["_wl_drive_sync_force"] = True
                st.rerun()
        if st.session_state.get("_wl_sync_conflicts"):
            _cf = list(st.session_state.get("_wl_sync_conflicts") or [])
            st.warning(
                "로컬·클라우드 일지가 다릅니다(자동 덮어쓰기 안 함): "
                + ", ".join(_cf[:8])
                + ("…" if len(_cf) > 8 else "")
            )
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("이 기기 → Drive", key="wl_cf_push_local", width="stretch", help="이 기기 내용으로 Drive를 맞춥니다."):
                    try:
                        from drive_autoload import resolve_drive_conflict

                        for name in list(_cf):
                            try:
                                resolve_drive_conflict(name, WORKLOG_DIR, prefer="local")
                            except Exception:
                                pass
                        st.session_state.pop("_wl_sync_conflicts", None)
                        st.session_state["_wl_drive_sync_force"] = True
                        _invalidate_saved_dates_cache()
                        st.rerun()
                    except Exception as e:
                        st.error(str(e) if not _wl_quiet_ui() else "동기화에 실패했습니다.")
            with c2:
                if st.button("Drive → 이 기기", key="wl_cf_pull_drive", width="stretch", help="Drive 내용으로 이 기기를 맞춥니다."):
                    try:
                        from drive_autoload import resolve_drive_conflict

                        for name in list(_cf):
                            try:
                                resolve_drive_conflict(name, WORKLOG_DIR, prefer="drive")
                            except Exception:
                                pass
                        st.session_state.pop("_wl_sync_conflicts", None)
                        st.session_state["_wl_drive_sync_force"] = True
                        _invalidate_saved_dates_cache()
                        st.rerun()
                    except Exception as e:
                        st.error(str(e) if not _wl_quiet_ui() else "동기화에 실패했습니다.")
            with c3:
                if st.button("나중에", key="wl_cf_dismiss", width="stretch"):
                    st.session_state.pop("_wl_sync_conflicts", None)
                    st.rerun()
        elif _on_cloud and not drive_remote_configured():
            if not st.session_state.get("_wl_remote_setup_hint"):
                st.session_state["_wl_remote_setup_hint"] = True
                st.caption(
                    "Cloud: secrets에 drive_uproad_folder_id + google_service_account(또는 API key) 설정"
                )
        elif not _on_cloud:
            st.caption("맥: Drive「dashboard 복사본/worklog」↔ 로컬 양방향")
    except Exception:
        pass


def render_worklog_tab(latest_update_str: str = "") -> None:
    if load_workbook is None:
        st.error("openpyxl 이 필요합니다. `pip install openpyxl` 후 다시 실행하세요.")
        return
    try:
        _ensure_dirs()
        if not os.path.exists(WORKLOG_TEMPLATE):
            if _wl_quiet_ui(): st.error("업무일지 템플릿이 없습니다. 사이드바에서 자료를 올린 뒤 `uploaded_cache/worklog/template.xlsx` 를 준비하세요.")
            else: st.error("업무일지 템플릿을 찾을 수 없습니다. `Desktop/업무일지.xlsx` 또는 `uploaded_cache/worklog/template.xlsx` 를 준비하세요.")
            return
    except Exception as e:
        if _wl_quiet_ui(): st.error("업무일지 템플릿을 준비하지 못했습니다. 파일을 다시 확인해 주세요.")
        else: st.error(f"템플릿 준비 실패: {e}")
        return

    # 반영 확인용 + 예전 미리보기 HTML 캐시 무효화(한 번)
    # 제목과 같은 블록에 폰트를 넣어 빈 markdown 여백이 생기지 않게 함
    st.markdown(
        f"<link rel='stylesheet' href='https://fonts.googleapis.com/css2?family=Nanum+Myeongjo:wght@400;700&display=swap'>"
        f"<style>section.main {{ font-family:{_WL_FONT_STACK}; }}</style>"
        "<div class='sub-header dashboard-tab-panel-head'>📝 일일업무일지</div>",
        unsafe_allow_html=True,
    )
    dev_caption(f"업무일지 빌드 {_WL_UI_BUILD}")
    _filt_changed = _dashboard_filters_changed_this_run()
    _arch_root = resolve_worklog_archive_root()
    if is_dev_mode():
        if _arch_root:
            st.caption(f"월별 저장 경로: `{_arch_root}/{{연도}}/{{N}}월.xlsx` (날짜=시트명)")
        else:
            st.caption("월별 저장 경로: `Desktop/업무/일지/{연도}/{N}월.xlsx` (Google Drive 동기화 시 「다른 컴퓨터/내 컴퓨터/Desktop/업무/일지」)")
        try:
            from drive_remote_fetch import drive_remote_configured

            if _wl_is_streamlit_cloud():
                if drive_remote_configured():
                    st.caption("☁ Drive uproad **연동됨** — 재시작·↻ 버튼으로 최신 일지 로드")
                else:
                    st.caption(
                        "☁ Drive **미연동** — secrets에 drive_uproad_folder_id + google_service_account 필요"
                    )
            else:
                st.caption("☁ 맥: Drive「dashboard 복사본/worklog」↔ 로컬")
        except Exception:
            pass
    if not st.session_state.get("_wl_cache_bust_v24"):
        st.session_state["_wl_cache_bust_v24"] = True
        for k in list(st.session_state.keys()):
            if not isinstance(k, str):
                continue
            if k.startswith("wl_left_excel_html_") or k.startswith("wl_left_excel_sig_") or k.startswith("wl_print_html_cache_"):
                st.session_state.pop(k, None)


    if "worklog_selected" not in st.session_state:
        st.session_state["worklog_selected"] = date.today()
    selected: date = st.session_state["worklog_selected"]

    _flush_queued_date_pick()
    _run_pending_worklog_day_delete()
    if _run_pending_worklog_date_change():
        selected = st.session_state.get("worklog_selected") or selected
        st.session_state.pop("wl_need_app_rerun", None)
    selected, _tab_moved = consume_left_date_pick_move(selected)
    if _tab_moved:
        selected = st.session_state.get("worklog_selected") or selected
        st.session_state.pop("wl_need_app_rerun", None)

    _boot = not st.session_state.get("_wl_boot_sync_done")
    _prepare_worklog_day_state(selected, skip_remote_pull=True)
    if _boot:
        st.session_state["_wl_boot_sync_done"] = True
    else:
        _maybe_sync_worklog_remote()

    if st.session_state.get("wl_print_panel"):
        _render_worklog_print_panel()
        return

    if "wl_date_pick" not in st.session_state:
        _set_wl_date_pick(selected)

    st.markdown(
        """<style>div[class*="st-key-wl_del_day_open"] button, div[class*="st-key-wl_del_day_yes"] button { font-size: 0.72rem !important; padding: 0.12rem 0.4rem !important; min-height: 1.55rem !important; }</style>""",
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <style>
        div[data-testid="column"]:nth-of-type(2) {
            position: sticky;
            top: 4rem;
            align-self: flex-start;
            z-index: 99;
        }
        /* iPad Mini 7 세로: 입력칸 sticky가 필터를 가리지 않게. 가로는 기존 유지 */
        @media (max-width: 850px) {
            div[data-testid="column"]:nth-of-type(2) {
                position: static !important;
                top: auto !important;
            }
        }
        @media (hover: none) and (pointer: coarse) {
            div[class*="st-key-wl_date_pick"] [data-baseweb="input"],
            div[class*="st-key-wl_cal_pop"] button,
            div[class*="st-key-wl_date_move_btn"] button,
            div[class*="st-key-wl_save_btn_"] button,
            div[class*="st-key-wl_add_btn_"] button,
            div[class*="st-key-wl_del_btn_"] button,
            div[class*="st-key-wl_print_btn"] button,
            div[class*="st-key-wl_open_print_btn"] button {
                min-height: 2.7rem !important;
                touch-action: manipulation;
            }
            div[class*="st-key-wl_lines_comp_"] textarea,
            div[class*="st-key-wl_clients_comp_"] input,
            div[class*="st-key-wl_remarks_comp_"] input {
                touch-action: manipulation;
                font-size: 16px !important;
            }
        }
        @media (min-width: 851px) and (max-width: 1180px) and (orientation: landscape) {
            div[data-testid="column"]:nth-of-type(2) {
                top: 3.2rem;
            }
        }
        div[class*="st-key-wl_sum_host_"],
        div[class*="st-key-wl_excel_host_"] {
            min-height: 240px;
        }
        /* JS-only Enter hook이 빈 박스로 저장 버튼을 덮지 않게 */
        div[class*="st-key-wl_enter_hook_"],
        div[class*="st-key-wl_scroll_lock"] {
            height: 0 !important;
            min-height: 0 !important;
            max-height: 0 !important;
            overflow: hidden !important;
            margin: 0 !important;
            padding: 0 !important;
            pointer-events: none !important;
        }
        /* 저장 후에도 날짜칸이 잠긴 것처럼 보이지 않게 */
        div[class*="st-key-wl_date_pick"] [data-baseweb="input"],
        div[class*="st-key-wl_date_pick"] input {
            background-color: #ffffff !important;
            color: #111827 !important;
            cursor: pointer !important;
            opacity: 1 !important;
        }
        div[class*="st-key-wl_date_bar"] { margin: 0 0 0.15rem 0 !important; }
        div[class*="st-key-wl_date_pick"] { margin-bottom: 0 !important; }
        div[class*="st-key-wl_date_pick"] [data-baseweb="input"] {
            min-height: 2.4rem !important;
            border-radius: 0.5rem !important;
        }
        div[class*="st-key-wl_cal_pop"] button,
        div[class*="st-key-wl_date_move_btn"] button,
        div[class*="st-key-wl_del_open_btn"] button {
            min-height: 2.4rem !important;
            height: 2.4rem !important;
            border-radius: 0.5rem !important;
            padding: 0 0.85rem !important;
            font-size: 0.875rem !important;
            font-weight: 600 !important;
        }
        div[class*="st-key-wl_save_btn_"] button,
        div[class*="st-key-wl_add_btn_"] button,
        div[class*="st-key-wl_del_btn_"] button,
        div[class*="st-key-wl_print_btn"] button,
        div[class*="st-key-wl_open_print_btn"] button {
            position: relative !important;
            z-index: 5 !important;
            pointer-events: auto !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    _render_worklog_sync_ui()

    # 업무일지 전체(미리보기+달력툴바+편집칸)를 하나의 fragment로 묶는다.
    # 달력 날짜 클릭·월 이동 등 업무일지 내부 조작이 대시보드 12탭 전체를 다시 그리지 않고
    # 업무일지 영역만 fragment 스코프로 갱신 → 달력 클릭/날짜 이동 로딩이 크게 짧아진다.
    @st.fragment
    def _worklog_body() -> None:
        _run_pending_worklog_day_delete()
        sel: date = st.session_state.get("worklog_selected") or selected
        _prepare_worklog_day_state(sel, skip_remote_pull=True)
        _render_worklog_scroll_lock(sel.isoformat())
        col_preview, col_edit = st.columns([1, 1.14], gap="small")
        with col_preview:
            _render_worklog_left_preview(sel)
        with col_edit:
            st.markdown("##### 업무 입력")
            _render_worklog_input_panel(sel)
            _wl_finish_edit_fragment()

    _worklog_body()
