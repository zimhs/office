"""일일업무일지 탭 — 엑셀 양식 그대로 표시/편집/날짜별 저장/출력."""
from __future__ import annotations

import calendar
import html
import json
import os
import platform
import re
import shutil
import subprocess
import time
import unicodedata
from datetime import date
from functools import lru_cache

import streamlit as st
import streamlit.components.v1 as components
from streamlit.errors import StreamlitAPIException

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    load_workbook = None
    get_column_letter = None


def _wl_rerun() -> None:
    """업무일지 fragment 안이면 fragment만 다시 실행 (전체 앱 로딩 방지)."""
    try:
        st.rerun(scope="fragment")
    except StreamlitAPIException:
        st.rerun()


def _wl_quiet_ui() -> bool:
    """아이패드·클라우드·비 macOS — 맥 전용 경로/Excel 경고를 숨긴다."""
    try:
        if st.session_state.get("force_touch_ui") is True:
            return True
        v = st.query_params.get("touch_ui", "")
        if isinstance(v, (list, tuple)):
            v = v[0] if v else ""
        if str(v).strip() in ("1", "true", "True"):
            return True
    except Exception:
        pass
    try:
        return platform.system() != "Darwin"
    except Exception:
        return True


def _invalidate_saved_dates_cache() -> None:
    st.session_state.pop("wl_saved_dates_cache", None)

WORKLOG_DIR = os.path.join("uploaded_cache", "worklog")
WORKLOG_TEMPLATE = os.path.join(WORKLOG_DIR, "template.xlsx")
WORKLOG_TEMPLATE_SRC = os.path.expanduser("~/Desktop/업무일지.xlsx")
# Google Drive「다른 컴퓨터/…/Desktop/업무/일지」또는 로컬 Desktop/업무/일지
WORKLOG_ARCHIVE_REL = os.path.join("Desktop", "업무", "일지")

WL_MIN_ROW, WL_MAX_ROW = 1, 47
WL_MIN_COL, WL_MAX_COL = 3, 28  # C ~ AB

# 작성 칸 (라벨 C40/C44 등은 건드리지 않음)
WL_CLIENT_ROWS = list(range(9, 40))  # C
WL_CONTENT_ROWS = list(range(9, 40))  # G
WL_NEXT_ROWS = list(range(40, 44))  # D (익일업무 본문)
WL_NOTE_ROWS = list(range(44, 48))  # D (특이사 항 본문)

# 내용 병합 열 G~X (원본 양식). 한 줄 폭은 템플릿에서 계산.
WL_CONTENT_COL_START = 7  # G
WL_CONTENT_COL_END = 24  # X
# 거래처 병합 열 C~F
WL_CLIENT_COL_START = 3  # C
WL_CLIENT_COL_END = 6  # F

# 내용 칸 편집기 (CCv2) — Enter로 다음 칸 생성/이동을 JS에서 직접 처리
_WL_LINES_HTML = """
<div class="wl-lines"></div>
"""

_WL_LINES_CSS = """
.wl-lines { display: flex; flex-direction: column; gap: 4px; width: 100%; max-width: 420px; }
.wl-row { display: flex; gap: 4px; align-items: center; width: 100%; }
.wl-row input {
  flex: 1 1 auto;
  min-width: 0;
  width: 100%;
  height: 1.55rem;
  padding: 0.1rem 0.35rem;
  border: 1px solid #94A3B8;
  border-radius: 4px;
  background: #fff;
  color: #0F172A;
  font-size: 0.75rem;
  line-height: 1.2;
  outline: none;
}
.wl-row input:focus {
  border-color: #0F766E;
  box-shadow: 0 0 0 1px #0F766E;
}
.wl-row button {
  flex: 0 0 40px;
  height: 1.55rem;
  border: 1px solid #CBD5E1;
  border-radius: 4px;
  background: #F8FAFC;
  color: #334155;
  font-size: 0.68rem;
  padding: 0;
  cursor: pointer;
}
.wl-row button:hover { background: #F1F5F9; }
.wl-row button.hidden { visibility: hidden; pointer-events: none; }
"""

_WL_LINES_JS = r"""
const __wlLinesInst = new WeakMap();

export default function (component) {
  const { data, parentElement, setStateValue } = component;
  const root = parentElement.querySelector(".wl-lines");
  if (!root) return;

  const maxU = Number((data && data.max_u) || 70);
  const rev = Number((data && data.rev) || 0);
  const focusReq = Number((data && data.focus));
  const incoming = Array.isArray(data && data.lines)
    ? data.lines.map((x) => String(x ?? ""))
    : [""];

  let inst = __wlLinesInst.get(root);
  if (!inst) {
    inst = { lines: null, rev: null, rebuilding: false };
    __wlLinesInst.set(root, inst);
  }

  function charUnits(ch) {
    const o = ch.charCodeAt(0);
    if (
      (o >= 0xac00 && o <= 0xd7a3) ||
      (o >= 0x1100 && o <= 0x11ff) ||
      (o >= 0x3130 && o <= 0x318f) ||
      (o >= 0x2e80 && o <= 0x9fff) ||
      (o >= 0xff00 && o <= 0xffef)
    )
      return 2;
    return 1;
  }
  function displayUnits(s) {
    let w = 0;
    s = s || "";
    for (let i = 0; i < s.length; i++) w += charUnits(s.charAt(i));
    return w;
  }
  function fitByUnits(s, max) {
    if (!s) return { head: "", tail: "" };
    if (displayUnits(s) <= max) return { head: s, tail: "" };
    let acc = 0;
    for (let i = 0; i < s.length; i++) {
      const cu = charUnits(s.charAt(i));
      if (acc + cu > max) {
        if (i === 0) return { head: s.slice(0, 1), tail: s.slice(1) };
        return { head: s.slice(0, i), tail: s.slice(i) };
      }
      acc += cu;
    }
    return { head: s, tail: "" };
  }
  function normalize(arr) {
    const out = (arr || []).map((x) => String(x ?? ""));
    if (!out.length) out.push("");
    if (out[out.length - 1] !== "") out.push("");
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
    setStateValue("lines", out);
    if (typeof focusIdx === "number") setStateValue("focus", focusIdx);
    return out;
  }
  function focusAt(idx) {
    requestAnimationFrame(() => {
      const el = root.querySelector('input[data-idx="' + idx + '"]');
      if (!el) return;
      try {
        el.focus({ preventScroll: false });
        const n = (el.value || "").length;
        el.setSelectionRange(n, n);
      } catch (e) {
        try {
          el.focus();
        } catch (e2) {}
      }
    });
  }

  function insertLineAfter(j, inputEl) {
    // DOM 값을 그대로 읽고, 현재 칸 바로 아래 빈 칸을 무조건 추가
    const raw = [];
    root.querySelectorAll("input[data-idx]").forEach((inp) => {
      raw.push(inp.value || "");
    });
    if (!raw.length) raw.push("");
    while (raw.length <= j) raw.push("");
    raw[j] = (inputEl && inputEl.value) || raw[j] || "";
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
      input.placeholder = j + 1 + "칸 (빈 칸도 저장 · Enter=다음 칸)";
      input.dataset.idx = String(j);
      const commitValue = () => {
        if (inst.rebuilding) return;
        const cur = readDomLines();
        let v = cur[j] || "";
        if (displayUnits(v) > maxU) {
          const ft = fitByUnits(v, maxU);
          cur[j] = ft.head;
          if (j + 1 < cur.length) cur[j + 1] = ft.tail + (cur[j + 1] || "");
          else cur.splice(j + 1, 0, ft.tail);
          emit(cur, j + 1);
          rebuild(j + 1);
          return;
        }
        emit(cur, null);
      };
      input.addEventListener("input", (e) => {
        if (e.isComposing) return;
        commitValue();
      });
      input.addEventListener("compositionend", () => {
        commitValue();
      });
      input.addEventListener("blur", () => {
        commitValue();
      });
      input.addEventListener("keydown", (e) => {
        // 빈 칸 Enter도 동작: IME keyCode 229는 값이 있을 때만 무시
        if (e.key !== "Enter") return;
        const typing = (input.value || "") !== "";
        if (e.isComposing && typing) return;
        e.preventDefault();
        e.stopPropagation();
        insertLineAfter(j, input);
      });
      const del = document.createElement("button");
      del.type = "button";
      del.textContent = "삭제";
      const isLastEmpty = j === lines.length - 1 && (line || "") === "";
      if (isLastEmpty) del.classList.add("hidden");
      del.addEventListener("click", () => {
        const cur = readDomLines();
        cur.splice(j, 1);
        if (!cur.length) cur.push("");
        const fj = Math.min(j, Math.max(cur.length - 1, 0));
        emit(cur, fj);
        rebuild(fj);
      });
      row.appendChild(input);
      row.appendChild(del);
      root.appendChild(row);
    });
    inst.rebuilding = false;
    if (typeof focusIdx === "number" && focusIdx >= 0) focusAt(focusIdx);
  }

  // rev가 바뀔 때만 Python 값으로 강제 동기화. 그 외에는 DOM/로컬 유지(입력 유실 방지).
  if (inst.rev !== rev) {
    inst.rev = rev;
    inst.lines = normalize(incoming);
    rebuild(Number.isFinite(focusReq) ? focusReq : -1);
  } else if (!root.childElementCount) {
    if (!inst.lines) inst.lines = normalize(incoming);
    rebuild(Number.isFinite(focusReq) ? focusReq : -1);
  } else {
    // DOM이 Python 상태보다 앞선 경우에만 반영 (저장 유실 방지, 무한 rerun 방지)
    const dom = readDomLines();
    const base = normalize(incoming);
    let same = dom.length === base.length;
    if (same) {
      for (let i = 0; i < dom.length; i++) {
        if (dom[i] !== base[i]) {
          same = false;
          break;
        }
      }
    }
    if (!same) emit(dom, null);
    else inst.lines = dom;
  }
}
"""

_WL_LINES_EDITOR = st.components.v2.component(
    "worklog_entry_lines",
    html=_WL_LINES_HTML,
    css=_WL_LINES_CSS,
    js=_WL_LINES_JS,
)


def _entry_lines_comp_key(iso: str, entry_i: int) -> str:
    return f"wl_lines_comp_{iso}_{entry_i}"


def _entry_lines_rev_key(iso: str, entry_i: int) -> str:
    return f"wl_ent_rev_{iso}_{entry_i}"


def _entry_lines_live_key(iso: str, entry_i: int) -> str:
    return f"wl_lines_live_{iso}_{entry_i}"


def _scrub_dummy_label(val: str) -> str:
    """예전 placeholder가 값으로 남은 경우(거래처1, 내용3 등) 제거."""
    s = (val or "").strip()
    if re.fullmatch(r"거래처\d+", s) or re.fullmatch(r"내용\d+", s):
        return ""
    return val or ""


def _char_units(ch: str) -> int:
    """엑셀 동아시아 폭: 전각·한글·호환자모=2, 그 외=1."""
    ea = unicodedata.east_asian_width(ch)
    if ea in ("F", "W", "A"):
        return 2
    o = ord(ch)
    if (
        0xAC00 <= o <= 0xD7A3
        or 0x1100 <= o <= 0x11FF
        or 0x3130 <= o <= 0x318F
        or 0x2E80 <= o <= 0x9FFF
        or 0xF900 <= o <= 0xFAFF
        or 0xFF00 <= o <= 0xFFEF
    ):
        return 2
    return 1


def _display_units(s: str) -> int:
    return sum(_char_units(ch) for ch in (s or ""))


def _excel_col_width(ws, col_idx: int) -> float:
    """openpyxl 단일문자 키가 잘못된 폭을 줄 수 있어 범위(min/max)로 조회."""
    best = None  # (span, width)
    for dim in ws.column_dimensions.values():
        if dim.min is None or dim.max is None or dim.width is None:
            continue
        if dim.min <= col_idx <= dim.max:
            span = dim.max - dim.min
            if best is None or span < best[0]:
                best = (span, float(dim.width))
    if best:
        return best[1]
    try:
        return float(ws.sheet_format.defaultColWidth or 8.43)
    except Exception:
        return 8.43


def _excel_width_to_px(width: float) -> int:
    """엑셀 열 폭 → CSS px.

    원본 스크린샷 대비 내용칸 채움(~92%)에 맞춘 계수.
    """
    try:
        w = float(width)
    except (TypeError, ValueError):
        w = 8.43
    return max(10, int(w * 7 + 5))


# 아이패드/브라우저에서 엑셀 14pt보다 좁게 렌더되는 보정 (원본 채움 실측)
_WL_FONT_PX_SCALE = 1.08
# 원본 양식 본문 글꼴 (엑셀·하단 미리보기)
_WL_BODY_FONT_NAME = "바탕체"
_WL_BODY_FONT_PT = 14.0


def _set_body_font(cell) -> None:
    """거래처·내용 등 본문 칸에 바탕체 적용."""
    try:
        cell.font = cell.font.copy(
            name=_WL_BODY_FONT_NAME, size=float(_WL_BODY_FONT_PT)
        )
    except Exception:
        pass


@lru_cache(maxsize=1)
def _content_line_units() -> int:
    """원본 G:X 병합 폭 기준 — 칸을 거의 채우되 우측이 넘치지 않게 여유를 둔다.

    반각=1, 한글=2. 장문이 원본 테두리를 살짝 넘기던 것을 막아
    약 70단위(~한글 35자)에서 다음 칸으로 넘긴다.
    """
    fallback = 70
    if load_workbook is None or not os.path.exists(WORKLOG_TEMPLATE):
        return fallback
    try:
        wb = load_workbook(WORKLOG_TEMPLATE, data_only=False)
        ws = wb.active
        total = sum(
            _excel_col_width(ws, c)
            for c in range(WL_CONTENT_COL_START, WL_CONTENT_COL_END + 1)
        )
        wb.close()
        units = int(total * (11 / 14) * 1.06)
        return max(66, min(units, 71))
    except Exception:
        return fallback


@lru_cache(maxsize=1)
def _client_line_units() -> int:
    """원본 C:F 거래처 병합 폭 기준 — 칸을 넘기면 다음 행으로(반각=1)."""
    fallback = 18
    if load_workbook is None or not os.path.exists(WORKLOG_TEMPLATE):
        return fallback
    try:
        wb = load_workbook(WORKLOG_TEMPLATE, data_only=False)
        ws = wb.active
        total = sum(
            _excel_col_width(ws, c)
            for c in range(WL_CLIENT_COL_START, WL_CLIENT_COL_END + 1)
        )
        wb.close()
        units = int(total * (11 / 14) * 1.05)
        return max(12, min(units, 28))
    except Exception:
        return fallback


def _fit_by_units(s: str, max_units: int | None = None) -> tuple[str, str]:
    """max_units 안에 들어가는 접두 / 나머지."""
    if max_units is None:
        max_units = _content_line_units()
    if not s:
        return "", ""
    if _display_units(s) <= max_units:
        return s, ""
    acc = 0
    for i, ch in enumerate(s):
        cu = _char_units(ch)
        if acc + cu > max_units:
            return s[:i] if i else s[:1], s[i:] if i else s[1:]
        acc += cu
    return s, ""


def _chunk_text(text: str, max_units: int | None = None) -> list[str]:
    """개행 우선, 긴 줄은 내용칸 폭으로 잘라 행 목록 생성."""
    if max_units is None:
        max_units = _content_line_units()
    s = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not s.strip():
        return []
    out: list[str] = []
    for para in s.split("\n"):
        rest = para
        if rest == "" and not out:
            continue
        if rest == "":
            continue
        while rest:
            head, rest = _fit_by_units(rest, max_units)
            if not head and rest:
                head, rest = rest[:1], rest[1:]
            out.append(head)
            if not rest:
                break
    return out


def _spill_column(cells: dict, rows: list[int], col: str) -> dict:
    """긴 값을 아래 칸으로 넘김 (저장/표시 공통)."""
    max_u = _content_line_units()
    vals = [str(cells.get(f"{col}{r}", "") or "") for r in rows]
    for i in range(len(vals)):
        while _display_units(vals[i]) > max_u and i + 1 < len(vals):
            head, tail = _fit_by_units(vals[i], max_u)
            vals[i] = head
            vals[i + 1] = tail + vals[i + 1]
        if _display_units(vals[i]) > max_u:
            head, _tail = _fit_by_units(vals[i], max_u)
            vals[i] = head
    out = dict(cells)
    for r, v in zip(rows, vals):
        out[f"{col}{r}"] = v
    return out


def _spill_all_content(cells: dict) -> dict:
    cells = _spill_column(cells, WL_CONTENT_ROWS, "G")
    cells = _spill_column(cells, WL_NEXT_ROWS, "D")
    cells = _spill_column(cells, WL_NOTE_ROWS, "D")
    return cells


def _ensure_dirs() -> None:
    os.makedirs(WORKLOG_DIR, exist_ok=True)
    if not os.path.exists(WORKLOG_TEMPLATE) and os.path.exists(WORKLOG_TEMPLATE_SRC):
        shutil.copy2(WORKLOG_TEMPLATE_SRC, WORKLOG_TEMPLATE)


def _iter_google_drive_roots() -> list[str]:
    cloud = os.path.join(os.path.expanduser("~"), "Library", "CloudStorage")
    if not os.path.isdir(cloud):
        return []
    roots: list[str] = []
    try:
        for name in sorted(os.listdir(cloud)):
            if name.startswith("GoogleDrive"):
                roots.append(os.path.join(cloud, name))
    except OSError:
        return []
    return roots


def resolve_worklog_archive_root() -> str | None:
    """업무/일지 보관 루트 (년도 폴더의 부모).

    우선순위:
    1) Google Drive「다른 컴퓨터/*/Desktop/업무/일지」(이미 있는 경로)
    2) ~/Desktop/업무/일지
    3) Google Drive에서 Desktop/업무 까지 보이면 일지 폴더 생성
    """
    candidates: list[str] = []
    home = os.path.expanduser("~")
    candidates.append(os.path.join(home, "Desktop", "업무", "일지"))

    for groot in _iter_google_drive_roots():
        for other_name in ("다른 컴퓨터", "Computers"):
            other = os.path.join(groot, other_name)
            if not os.path.isdir(other):
                continue
            try:
                pcs = sorted(os.listdir(other))
            except OSError:
                continue
            # 「내 컴퓨터 (1)」을 우선
            pcs.sort(key=lambda n: (0 if "(1)" in n else 1, n))
            for pc in pcs:
                candidates.append(os.path.join(other, pc, WORKLOG_ARCHIVE_REL))

    existing = [p for p in candidates if os.path.isdir(p)]
    if existing:
        return existing[0]

    # 없으면 Desktop/업무 아래에 일지 생성 시도
    for p in candidates:
        parent = os.path.dirname(p)  # …/업무
        grand = os.path.dirname(parent)  # …/Desktop
        if os.path.isdir(grand):
            try:
                os.makedirs(p, exist_ok=True)
                return p
            except OSError:
                continue
    return None


def worklog_archive_path(d: date) -> str | None:
    """달력 외 보관 경로: …/일지/{년도}/{YYYY-MM-DD}.xlsx (년도 폴더 없으면 생성)."""
    root = resolve_worklog_archive_root()
    if not root:
        return None
    year_dir = os.path.join(root, str(d.year))
    try:
        os.makedirs(year_dir, exist_ok=True)
    except OSError:
        return None
    return os.path.join(year_dir, f"{d.isoformat()}.xlsx")


def worklog_path(d: date) -> str:
    return os.path.join(WORKLOG_DIR, f"{d.isoformat()}.xlsx")


def list_saved_worklog_dates() -> set[str]:
    cached = st.session_state.get("wl_saved_dates_cache")
    if isinstance(cached, set):
        return cached
    _ensure_dirs()
    out: set[str] = set()
    for name in os.listdir(WORKLOG_DIR):
        if (
            name.endswith(".xlsx")
            and len(name) >= 15
            and name[0:4].isdigit()
            and name not in {"template.xlsx"}
            and not name.startswith("_preview_")
            and "_인쇄" not in name
        ):
            out.add(name.replace(".xlsx", ""))
    st.session_state["wl_saved_dates_cache"] = out
    return out


def format_worklog_date(d: date) -> str:
    weeks = "월화수목금토일"
    return f"{d.strftime('%Y-%m-%d')} ({weeks[d.weekday()]})"


def _clear_content_cells(ws) -> None:
    for r in WL_CLIENT_ROWS:
        ws.cell(r, 3).value = None
    for r in WL_CONTENT_ROWS:
        ws.cell(r, 7).value = None
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        ws.cell(r, 4).value = None


def _empty_cells(d: date) -> dict:
    cells = {"date": format_worklog_date(d)}
    for r in WL_CLIENT_ROWS:
        cells[f"C{r}"] = ""
    for r in WL_CONTENT_ROWS:
        cells[f"G{r}"] = ""
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        cells[f"D{r}"] = ""
    return cells


def read_worklog_cells(d: date) -> dict:
    path = worklog_path(d)
    if not os.path.exists(path) or load_workbook is None:
        return _empty_cells(d)
    wb = load_workbook(path, data_only=False)
    ws = wb.active
    cells = {"date": format_worklog_date(d)}
    for r in WL_CLIENT_ROWS:
        v = ws.cell(r, 3).value
        cells[f"C{r}"] = "" if v is None else str(v)
    for r in WL_CONTENT_ROWS:
        v = ws.cell(r, 7).value
        cells[f"G{r}"] = "" if v is None else str(v)
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        v = ws.cell(r, 4).value
        cells[f"D{r}"] = "" if v is None else str(v)
    c5 = ws["C5"].value
    if c5 is not None and not str(c5).startswith("="):
        cells["date"] = str(c5)
    wb.close()
    return cells


def write_cells_to_path(path: str, d: date, cells: dict, *, blank_base: bool = False) -> None:
    """양식·병합·스타일을 유지한 채 path에 셀 값 기록."""
    if load_workbook is None:
        raise RuntimeError("openpyxl 이 필요합니다.")
    _ensure_dirs()
    if not os.path.exists(path):
        if not os.path.exists(WORKLOG_TEMPLATE):
            raise FileNotFoundError("업무일지 템플릿이 없습니다.")
        shutil.copy2(WORKLOG_TEMPLATE, path)
        blank_base = True
    wb = load_workbook(path)
    ws = wb.active
    if blank_base:
        _clear_content_cells(ws)
    ws["C5"] = cells.get("date") or format_worklog_date(d)
    for r in WL_CLIENT_ROWS:
        cell = ws.cell(r, 3)
        cell.value = (cells.get(f"C{r}", "") or None)
        _set_body_font(cell)
    for r in WL_CONTENT_ROWS:
        cell = ws.cell(r, 7)
        cell.value = (cells.get(f"G{r}", "") or None)
        _set_body_font(cell)
        # 칸 안에서 줄바꿈·넘침으로 다음 행을 가리지 않도록
        try:
            cell.alignment = cell.alignment.copy(wrapText=False, shrinkToFit=False)
        except Exception:
            pass
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        cell = ws.cell(r, 4)
        cell.value = (cells.get(f"D{r}", "") or None)
        _set_body_font(cell)
        try:
            cell.alignment = cell.alignment.copy(wrapText=False, shrinkToFit=False)
        except Exception:
            pass
    wb.save(path)
    wb.close()


def save_worklog_cells(d: date, cells: dict) -> str:
    """달력용 캐시 저장 + Desktop/업무/일지 복사(맥) + Drive「dashboard 복사본/worklog」공유.

    Drive 폴더가 있으면 맥·아이패드(Drive) 공통 저장. Streamlit Cloud만 쓰면
    서버 디스크라 Drive와 즉시 공유되지 않음(맥에서 Drive 동기화 시 반영).
    """
    path = worklog_path(d)
    is_new = not os.path.exists(path)
    cells = _spill_all_content(cells)
    write_cells_to_path(path, d, cells, blank_base=is_new)
    _invalidate_saved_dates_cache()
    st.session_state.pop("wl_last_archive_path", None)
    st.session_state.pop("wl_last_archive_err", None)
    st.session_state.pop("wl_last_drive_path", None)
    try:
        archive = worklog_archive_path(d)
        if archive:
            shutil.copy2(path, archive)
            st.session_state["wl_last_archive_path"] = archive
    except Exception as e:
        st.session_state["wl_last_archive_err"] = str(e)
    try:
        from drive_autoload import push_worklog_day_to_drive

        drv = push_worklog_day_to_drive(path, WORKLOG_DIR)
        if drv:
            st.session_state["wl_last_drive_path"] = drv
    except Exception:
        pass
    return path


def delete_worklog_day(d: date) -> list[str]:
    """해당일 저장 파일·미리보기·인쇄본 삭제. 세션 입력 상태도 비움."""
    _ensure_dirs()
    iso = d.isoformat()
    removed: list[str] = []
    # 날짜 관련 파일 전부 제거 (저장본·미리보기·인쇄본)
    targets = [
        worklog_path(d),
        os.path.join(WORKLOG_DIR, f"_preview_{iso}.xlsx"),
        os.path.join(WORKLOG_DIR, f"일일업무일지_{iso}_인쇄.xlsx"),
    ]
    try:
        from drive_autoload import resolve_drive_worklog_dir

        _ddr = resolve_drive_worklog_dir()
        if _ddr:
            targets.append(os.path.join(_ddr, f"{iso}.xlsx"))
    except Exception:
        pass
    for name in os.listdir(WORKLOG_DIR):
        if iso in name and name.endswith(".xlsx") and name != "template.xlsx":
            targets.append(os.path.join(WORKLOG_DIR, name))
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

    _invalidate_saved_dates_cache()
    _clear_date_widget_state(d)
    # 빈 상태로 강제 리셋 (다음 _init 이 파일을 다시 안 읽도록 boot=True + 빈 entries)
    empty = [{"client": "", "content": "", "lines": [], "blank_after": 1}]
    st.session_state[_boot_key(d)] = True
    st.session_state[_entries_key(d)] = empty
    st.session_state[_next_key(d)] = ""
    st.session_state[_notes_key(d)] = ""
    st.session_state[f"wl_entry_count_{iso}"] = 1
    st.session_state[f"wl_pending_sync_{iso}"] = {
        "entries": empty,
        "next": "",
        "notes": "",
        "msg": (
            f"삭제 완료"
            + (f": {', '.join(removed)}" if removed else " (저장본 없음, 입력만 초기화)")
        ),
    }
    return removed


def reassign_worklog_date(old: date, new: date) -> str:
    """입력/저장본을 새 날짜로 옮긴다. new 에 기존 저장본이 있으면 오류."""
    if old == new:
        return "same"
    if os.path.exists(worklog_path(new)):
        raise FileExistsError(f"{new.isoformat()} 에 이미 저장된 일지가 있습니다.")

    # 현재 입력 우선, 없으면 파일
    try:
        cells = _cells_from_widgets(old)
    except Exception:
        cells = read_worklog_cells(old)
    cells["date"] = format_worklog_date(new)

    old_saved = os.path.exists(worklog_path(old))
    if old_saved or any(
        str(cells.get(f"G{r}", "") or "").strip()
        or str(cells.get(f"C{r}", "") or "").strip()
        for r in WL_CONTENT_ROWS
    ) or any(str(cells.get(f"D{r}", "") or "").strip() for r in WL_NEXT_ROWS + WL_NOTE_ROWS):
        save_worklog_cells(new, cells)

    if old_saved:
        for path in (worklog_path(old), _preview_path(old), _print_xlsx_path(old)):
            if os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass

    entries = _grouped_entries_from_cells(cells) or [
        {"client": "", "content": "", "lines": [], "blank_after": 1}
    ]
    _, nd, nt = _entries_from_cells(cells)
    _clear_date_widget_state(old)
    st.session_state["worklog_selected"] = new
    st.session_state["worklog_month"] = date(new.year, new.month, 1)
    st.session_state[_boot_key(new)] = True
    st.session_state[_entries_key(new)] = entries
    st.session_state[_next_key(new)] = "\n".join(nd)
    st.session_state[_notes_key(new)] = "\n".join(nt)
    st.session_state[f"wl_entry_count_{new.isoformat()}"] = len(entries)
    st.session_state[f"wl_pending_sync_{new.isoformat()}"] = {
        "entries": entries,
        "next": "\n".join(nd),
        "notes": "\n".join(nt),
        "msg": "",
    }
    _invalidate_saved_dates_cache()
    return "moved" if old_saved else "retargeted"


def _cell_fill_color(cell) -> str | None:
    try:
        fill = cell.fill
        if not fill or fill.fill_type is None:
            return None
        fg = fill.fgColor
        if fg is None:
            return None
        if getattr(fg, "type", None) == "rgb" and fg.rgb and fg.rgb != "00000000":
            rgb = str(fg.rgb)
            if len(rgb) == 8:
                rgb = rgb[2:]
            return f"#{rgb}"
    except Exception:
        return None
    return None


def _border_css(cell) -> str:
    """엑셀에 실제로 있는 테두리만 반영. 없으면 빈 칸 격자(가짜 박스)를 그리지 않음."""
    parts = []
    try:
        b = cell.border
        for side_name, css_side in (
            ("left", "border-left"),
            ("right", "border-right"),
            ("top", "border-top"),
            ("bottom", "border-bottom"),
        ):
            side = getattr(b, side_name, None)
            if side and side.style:
                color = "#111"
                try:
                    c = side.color
                    if c is not None and getattr(c, "type", None) == "rgb" and c.rgb:
                        rgb = str(c.rgb)
                        if len(rgb) == 8:
                            rgb = rgb[2:]
                        if rgb and rgb != "00000000":
                            color = f"#{rgb}"
                except Exception:
                    pass
                parts.append(f"{css_side}:1px solid {color};")
    except Exception:
        pass
    return "".join(parts)


def _worklog_sheet_pixel_size(path: str) -> tuple[int, int]:
    """표시 범위(C~AB, 1~47)의 비스케일 픽셀 폭·높이."""
    if load_workbook is None or not path or not os.path.exists(path):
        return 900, 1312
    try:
        wb = load_workbook(path, data_only=False)
        ws = wb.active
        total_w = 0
        for c in range(WL_MIN_COL, WL_MAX_COL + 1):
            total_w += _excel_width_to_px(_excel_col_width(ws, c))
        total_h = 0
        for r in range(WL_MIN_ROW, WL_MAX_ROW + 1):
            h = ws.row_dimensions[r].height
            # Excel pt → CSS px (96dpi). 행 테두리 1px 여유.
            total_h += (int(round(float(h) * 96 / 72)) if h else 28) + 1
        wb.close()
        return max(1, total_w), max(1, total_h)
    except Exception:
        return 900, 1312


def _scaled_view_frame_size(path: str, scale: float) -> tuple[int, int]:
    """scale 적용 후 iframe에 넣을 폭·높이 (하단 잘림 방지 여유 포함)."""
    w, h = _worklog_sheet_pixel_size(path)
    s = float(scale) if scale and scale > 0 else 1.0
    # body padding·테두리·서브픽셀·Streamlit iframe 여유
    return int(w * s) + 28, int(h * s) + 64


def workbook_to_html(path: str) -> str:
    """원본 엑셀 양식 범위를 HTML 테이블로 변환 (병합·테두리·폰트 유지)."""
    if load_workbook is None:
        return "<p>openpyxl 필요</p>"
    wb = load_workbook(path, data_only=False)
    ws = wb.active

    merge_map: dict[tuple[int, int], tuple[int, int]] = {}
    skip: set[tuple[int, int]] = set()
    for mr in ws.merged_cells.ranges:
        if mr.max_row < WL_MIN_ROW or mr.min_row > WL_MAX_ROW:
            continue
        if mr.max_col < WL_MIN_COL or mr.min_col > WL_MAX_COL:
            continue
        # 양식 밖까지 뻗는 병합은 표시 범위로 자름
        max_r = min(mr.max_row, WL_MAX_ROW)
        max_c = min(mr.max_col, WL_MAX_COL)
        min_r = max(mr.min_row, WL_MIN_ROW)
        min_c = max(mr.min_col, WL_MIN_COL)
        if min_r > max_r or min_c > max_c:
            continue
        rs = max_r - min_r + 1
        cs = max_c - min_c + 1
        # 원본 top-left가 범위 밖이면 범위 내 첫 셀을 top-left로
        tl_r, tl_c = mr.min_row, mr.min_col
        if tl_r < WL_MIN_ROW or tl_c < WL_MIN_COL:
            tl_r, tl_c = min_r, min_c
        merge_map[(tl_r, tl_c)] = (rs, cs)
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                if (r, c) != (tl_r, tl_c):
                    skip.add((r, c))

    col_widths = []
    total_w = 0.0
    for c in range(WL_MIN_COL, WL_MAX_COL + 1):
        w = _excel_col_width(ws, c)
        px = _excel_width_to_px(w)
        col_widths.append(px)
        total_w += px

    rows_html = []
    for r in range(WL_MIN_ROW, WL_MAX_ROW + 1):
        h = ws.row_dimensions[r].height
        # 행 높이도 pt→px 동일 공식 (표시·측정 일치)
        height_px = int(round(float(h) * 96 / 72)) if h else 28
        tds = []
        for c in range(WL_MIN_COL, WL_MAX_COL + 1):
            if (r, c) in skip:
                continue
            cell = ws.cell(r, c)
            rs, cs = merge_map.get((r, c), (1, 1))
            val = cell.value
            if isinstance(val, str) and val.startswith("="):
                text = ""
            else:
                text = "" if val is None else str(val)
            font = cell.font
            fname = font.name or _WL_BODY_FONT_NAME
            # 엑셀 한글 폰트명 정규화 — 본문은 바탕(바탕글) 계열
            if fname in ("맑은 고딕", "Malgun Gothic"):
                fname_css = "Malgun Gothic"
            elif fname in (
                "바탕",
                "바탕체",
                "바탕글",
                "Batang",
                "BatangChe",
            ):
                fname_css = "Batang"
            else:
                fname_css = fname
            # 엑셀 셀 글꼴 그대로 → CSS px + 브라우저 보정
            is_content = c == 7 and 9 <= r <= 39
            is_client = c == 3 and 9 <= r <= 39
            is_body_d = c == 4 and (
                r in WL_NEXT_ROWS or r in WL_NOTE_ROWS
            )
            fsize_pt = float(font.size or _WL_BODY_FONT_PT)
            # 본문 칸은 항상 바탕체로 표시 (원본 적용)
            if is_content or is_client or is_body_d:
                fname_css = "Batang"
                fsize_pt = float(font.size or _WL_BODY_FONT_PT)
            fsize_px = fsize_pt * 96.0 / 72.0 * _WL_FONT_PX_SCALE
            line_h_px = max(fsize_px * 1.12, fsize_px + 1.0)
            bold = "bold" if font.bold else "normal"
            align = cell.alignment
            ha = align.horizontal or "left"
            va = align.vertical or "middle"
            if ha == "general":
                ha = "left"
            # 익일업무·특이사항 라벨(C열 병합): 원본처럼 세로 글자
            is_side_label = c == 3 and (
                (r == 40 and rs >= 3) or (r == 44 and rs >= 3)
            )
            if is_side_label:
                ha = "center"
                va = "middle"
            fill = _cell_fill_color(cell) or "#FFFFFF"
            border = _border_css(cell)
            span = ""
            if rs > 1:
                span += f' rowspan="{rs}"'
            if cs > 1:
                span += f' colspan="{cs}"'
            # 소프트 빈 줄(항목 안 빈 칸)은 화면에서도 비어 보이게
            if text in (_WL_SOFT_BLANK, "\u00a0"):
                text = ""
            elif is_content and text.strip() == "" and text != "":
                text = ""
            # 연속 공백·선행 공백이 HTML에서 사라지지 않게
            if is_side_label and text.strip():
                # 공백 제거 후 글자마다 세로 배치 (익일업무 / 특이사항)
                chars = [ch for ch in text.replace(" ", "").replace("\u3000", "") if ch]
                esc = "<br>".join(html.escape(ch) for ch in chars)
            else:
                esc = (
                    html.escape(text)
                    .replace(" ", "&nbsp;")
                    .replace("\n", "<br>")
                )
            # 내용칸: 엑셀처럼 한 줄 + 옆 빈 칸으로 넘침 표시 (clip 금지)
            if is_content or is_client:
                white = "nowrap"
                overflow = "visible"
                text_overflow = "clip"
                zidx = "position:relative;z-index:1;"
            elif is_side_label:
                white = "normal"
                overflow = "hidden"
                text_overflow = "clip"
                zidx = ""
            else:
                white = "pre-wrap"
                overflow = "visible"
                text_overflow = "clip"
                zidx = ""
            # 병합 셀 폭 = 포함 열 합 (고정 레이아웃에서 원본과 같은 채움감)
            c0 = c - WL_MIN_COL
            span_w = sum(col_widths[c0 : c0 + max(cs, 1)]) if c0 >= 0 else 0
            width_css = f"width:{span_w}px;min-width:{span_w}px;max-width:{span_w}px;" if span_w else ""
            # 바탕글(바탕체) 우선 — Nanum/고딕으로 대체되지 않게
            if is_content or is_client or is_body_d or is_side_label:
                font_stack = (
                    "'Batang','BatangChe','바탕','바탕체','바탕글',"
                    "'Apple Myungjo','AppleMyungjo','Nanum Myeongjo',serif"
                )
            else:
                font_stack = (
                    f"'{html.escape(fname_css)}','Batang','BatangChe',"
                    f"'Apple Myungjo','Malgun Gothic',serif"
                )
            pad_css = "padding:4px 1px;" if is_side_label else "padding:0 2px;"
            line_css = (
                f"line-height:{max(fsize_px * 1.35, fsize_px + 2):.2f}px;"
                if is_side_label
                else f"line-height:{line_h_px:.2f}px;"
            )
            style = (
                f"box-sizing:border-box;{width_css}{zidx}"
                f"font-family:{font_stack};"
                f"font-size:{fsize_px:.4f}px;font-weight:{bold};"
                f"text-align:{ha};vertical-align:{va};"
                f"background:{fill};{border}"
                f"{pad_css}white-space:{white};overflow:{overflow};"
                f"text-overflow:{text_overflow};word-break:keep-all;"
                f"height:{height_px}px;min-height:{height_px}px;"
                f"{line_css}"
            )
            tds.append(f'<td{span} style="{style}">{esc}</td>')
        rows_html.append(
            f'<tr style="height:{height_px}px;box-sizing:border-box;">'
            f'{"".join(tds)}</tr>'
        )

    colgroup = "".join(f'<col style="width:{w}px">' for w in col_widths)
    wb.close()
    return f"""
    <table class="wl-sheet" style="border-collapse:collapse;table-layout:fixed;width:{int(total_w)}px;background:#fff;box-sizing:border-box;">
      <colgroup>{colgroup}</colgroup>
      <tbody>{"".join(rows_html)}</tbody>
    </table>
    """


def _a4_print_fit(raw_w: int, raw_h: int) -> float:
    """A4 인쇄 가능 영역에 양식 전체가 들어가도록 축소 비율."""
    # A4 210×297mm, 여백 5mm → 약 200×287mm ≈ 756×1085px @96dpi
    if raw_w <= 0 or raw_h <= 0:
        return 1.0
    fit = min(1.0, 756.0 / float(raw_w), 1085.0 / float(raw_h))
    # 테두리·서브픽셀 잘림 방지
    return max(0.25, min(1.0, fit * 0.97))


def render_worklog_view_html(
    path: str,
    *,
    print_mode: bool = False,
    scale: float = 0.55,
    auto_print: bool = False,
    wrap_height: str | None = None,
) -> str:
    """원본 엑셀 양식 HTML (인쇄/보고용)."""
    sheet = workbook_to_html(path)
    raw_w, raw_h = _worklog_sheet_pixel_size(path)
    print_fit = _a4_print_fit(raw_w, raw_h)
    scaled_w = max(1, int(round(raw_w * print_fit)))
    scaled_h = max(1, int(round(raw_h * print_fit)))

    toolbar = ""
    if print_mode:
        scale = 1.0
        if auto_print:
            # 자동 print가 막힐 때 사용자가 직접 누를 수 있는 버튼
            toolbar = """
        <div class="toolbar no-print">
          <button type="button" id="wl-print-btn">인쇄 창 열기</button>
          <span class="hint">인쇄 대화상자가 안 뜨면 이 버튼을 누르세요.</span>
        </div>
        """
        else:
            toolbar = """
        <div class="toolbar no-print">
          <button type="button" id="wl-print-btn">인쇄하기</button>
          <span class="hint">「인쇄하기」를 누르면 인쇄 창이 열립니다. 대상에서 <b>프린터</b>를 선택하세요.</span>
        </div>
        """
    frame_w, frame_h = _scaled_view_frame_size(path, scale)
    # print_mode: transform+고정 wrap으로 전체가 한 장에 보이게 (zoom만 쓰면 잘리는 경우 있음)
    if print_mode:
        s_view = print_fit
        scale_css = (
            f"transform:scale({s_view});transform-origin:top left;"
            f"width:{raw_w}px;"
        )
        scale_css_fallback = ""
        wrap_h = f"{scaled_h}px"
        wrap_w = f"{scaled_w}px"
        wrap_overflow = "hidden"
        body_overflow = "auto"
        body_h = "auto"
    elif scale >= 1:
        scale_css = "zoom:1;width:fit-content;"
        scale_css_fallback = (
            "transform:scale(1);transform-origin:top left;width:fit-content;"
        )
        wrap_h = "auto"
        wrap_w = "100%"
        wrap_overflow = "visible"
        body_overflow = "visible"
        body_h = "auto"
    else:
        s = float(scale)
        scale_css = f"zoom:{s};width:fit-content;"
        scale_css_fallback = (
            f"transform:scale({s});transform-origin:top left;"
            f"margin-bottom:{(s - 1) * raw_h:.1f}px;width:fit-content;"
        )
        wrap_h = "auto"
        wrap_w = f"{frame_w}px"
        wrap_overflow = "visible"
        body_overflow = "hidden"
        body_h = f"{frame_h}px"
    if wrap_height is not None:
        wrap_h = wrap_height
    # auto_print: 자동 1회 시도 + 버튼으로 재시도 (먹통 방지)
    auto_script = ""
    if auto_print:
        auto_script = """
        <script>
          (function() {
            function goPrint() {
              try { window.focus(); window.print(); } catch (e) {}
            }
            var btn = document.getElementById('wl-print-btn');
            if (btn) btn.addEventListener('click', function(ev) {
              ev.preventDefault();
              goPrint();
            });
            if (document.readyState === 'complete') setTimeout(goPrint, 250);
            else window.addEventListener('load', function() { setTimeout(goPrint, 250); });
          })();
        </script>
        """
    elif print_mode:
        auto_script = """
        <script>
          (function() {
            var btn = document.getElementById('wl-print-btn');
            if (btn) btn.addEventListener('click', function(ev) {
              ev.preventDefault();
              try { window.focus(); window.print(); } catch (e) {}
            });
          })();
        </script>
        """
    fallback_block = ""
    if scale_css_fallback:
        fallback_block = f"""
  @supports not (zoom: 1) {{
    .sheet-scale {{ {scale_css_fallback} }}
  }}
"""
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>일일업무일지</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  @page {{ size: A4 portrait; margin: 5mm; }}
  html, body {{
    margin:0; padding:0; background:#fff;
    overflow:{body_overflow} !important;
    height:{body_h};
  }}
  body {{ padding:{"0" if print_mode else "6px"}; box-sizing:border-box; }}
  .toolbar {{ margin-bottom:10px; display:flex; gap:10px; align-items:center; flex-wrap:wrap; }}
  .toolbar button {{
    padding:8px 14px; font-size:14px; border:1px solid #334155; border-radius:6px;
    background:#1E293B; color:#fff; cursor:pointer;
  }}
  .toolbar button.secondary {{
    background:#F8FAFC; color:#334155; border-color:#CBD5E1; cursor:default;
  }}
  .toolbar .hint {{ font:12px/1.45 sans-serif; color:#64748B; max-width:42rem; }}
  .wrap {{
    overflow:{wrap_overflow} !important; height:{wrap_h};
    width:{wrap_w}; max-width:100%;
    border:{"none" if print_mode else "1px solid #94A3B8"}; background:#fff;
    box-sizing:border-box;
  }}
  .sheet-scale {{ {scale_css} }}
  .wl-sheet {{ border-collapse:collapse; table-layout:fixed; }}
  .wl-sheet, .wl-sheet td, .wl-sheet tr {{ box-sizing:border-box; }}
  {fallback_block}
  @media print {{
    html, body {{
      overflow:hidden !important; height:auto !important; width:auto !important;
      margin:0 !important; padding:0 !important;
      -webkit-print-color-adjust:exact; print-color-adjust:exact;
    }}
    .no-print, .toolbar {{ display:none !important; }}
    .wrap {{
      overflow:hidden !important;
      width:{scaled_w}px !important;
      height:{scaled_h}px !important;
      max-width:none !important;
      border:none !important;
      margin:0 !important;
    }}
    .sheet-scale {{
      transform:scale({print_fit:.4f}) !important;
      transform-origin:top left !important;
      zoom:normal !important;
      margin:0 !important;
      width:{raw_w}px !important;
    }}
  }}
</style></head>
<body>
  {toolbar}
  <div class="wrap"><div class="sheet-scale">{sheet}</div></div>
  {auto_script}
</body></html>"""

def _entry_blank_after(ent: dict | None, default: int = 1) -> int:
    """항목 뒤에 띄울 빈 칸 수 (0~10)."""
    try:
        n = int((ent or {}).get("blank_after", default))
    except (TypeError, ValueError):
        n = default
    return max(0, min(10, n))


# 항목 안 빈 내용줄 표식 (완전 빈 칸=항목 사이 간격과 구분)
_WL_SOFT_BLANK = " "


def _grouped_entries_from_cells(cells: dict) -> list[dict]:
    """엑셀 행 → 논리 항목.

    - C·G 모두 완전 빈 행 뒤에 오면 새 항목 (앞 빈 행 수 = 이전 blank_after)
    - G만 공백(소프트 빈 줄·여러 칸 공백)이면 같은 항목의 빈 줄
    - 빈 행 없이 이어지면 같은 항목: 거래처·내용 칸을 각각 이어 붙임
    """
    entries: list[dict] = []
    blank_run = 0
    for r in WL_CLIENT_ROWS:
        raw_c = str(cells.get(f"C{r}", "") or "")
        raw_g = str(cells.get(f"G{r}", "") or "")
        client = _scrub_dummy_label(raw_c).strip()
        soft_blank = raw_g == _WL_SOFT_BLANK or raw_g == "\u00a0"
        # G만 공백(1칸 이상) = 항목 안 빈 줄. C·G 모두 비어 있으면 항목 구분.
        g_whitespace_only = (
            not client
            and not soft_blank
            and raw_g != ""
            and raw_g.strip() == ""
        )
        fully_empty = (
            not client
            and not soft_blank
            and not g_whitespace_only
            and raw_c.strip() == ""
            and raw_g.strip() == ""
        )
        content = "" if (soft_blank or g_whitespace_only) else _scrub_dummy_label(raw_g)

        if fully_empty:
            blank_run += 1
            continue

        if g_whitespace_only or soft_blank:
            blank_run = 0
            if entries:
                entries[-1].setdefault("lines", []).append("")
                entries[-1]["content"] = "\n".join(entries[-1].get("lines") or [])
            continue

        if not entries or blank_run > 0:
            if entries:
                entries[-1]["blank_after"] = max(0, min(10, blank_run))
            blank_run = 0
            client_lines = [client] if client else []
            lines = [content or ""]
            entries.append(
                {
                    "client": "\n".join(client_lines),
                    "client_lines": client_lines,
                    "content": "\n".join(lines),
                    "lines": lines,
                    "blank_after": 1,
                }
            )
        else:
            if blank_run > 0:
                for _ in range(blank_run):
                    entries[-1].setdefault("lines", []).append("")
                    entries[-1].setdefault("client_lines", []).append("")
                blank_run = 0
            ent = entries[-1]
            cl = ent.setdefault("client_lines", [])
            if not cl and str(ent.get("client") or "").strip():
                cl[:] = [str(ent.get("client") or "").strip()]
            if client:
                cl.append(client)
            ent["client"] = "\n".join(x for x in cl if (x or "").strip())
            ent.setdefault("lines", []).append(content or "")
            ent["content"] = "\n".join(ent.get("lines") or [])
    return entries


def _entry_client_lines(ent: dict | None) -> list[str]:
    """거래처 칸 목록 (줄바꿈·칸폭 초과 = 다음 C칸)."""
    if not ent:
        return []
    raw = ent.get("client_lines")
    if isinstance(raw, list) and raw:
        src = [str(x or "") for x in raw]
    else:
        src = str(ent.get("client") or "").splitlines() or (
            [str(ent.get("client") or "")] if str(ent.get("client") or "").strip() else []
        )
    max_u = _client_line_units()
    out: list[str] = []
    for line in src:
        s = str(line or "")
        if not s.strip():
            # 중간 빈 줄은 유지하지 않음 (거래처는 연속 칸)
            continue
        parts = _chunk_text(s, max_u) or [s]
        out.extend(parts)
    return out


def _entry_pack_lines(ent: dict) -> list[str]:
    """저장용 칸 목록(빈 칸 포함). 긴 줄만 폭에 맞게 분할."""
    max_u = _content_line_units()
    raw = ent.get("lines")
    if isinstance(raw, list):
        src = [str(x or "") for x in raw]
    else:
        src = _chunk_text(str(ent.get("content") or ""), max_u) or []
    out: list[str] = []
    for line in src:
        if line == "":
            out.append("")
            continue
        parts = _chunk_text(line, max_u) or [line]
        out.extend(parts)
    if not out and _entry_client_lines(ent):
        out = [""]
    return out


def _content_row_usage(entries: list[dict] | None) -> dict:
    """원본 내용칸(G9~G39) 사용량. 항목별 blank_after·빈 칸 포함."""
    total = len(WL_CONTENT_ROWS)
    used = 0
    wrote_any = False
    prev_gap = 0
    per_entry: list[int] = []
    for ent in entries or []:
        clients = _entry_client_lines(ent)
        pack_lines = _entry_pack_lines(ent)
        if not any(str(x).strip() for x in clients) and not any(
            (x or "").strip() for x in pack_lines
        ):
            per_entry.append(0)
            continue
        gap = prev_gap if wrote_any else 0
        lines = max(len(clients), len(pack_lines), 1)
        need = gap + lines
        per_entry.append(need)
        used += need
        wrote_any = True
        prev_gap = _entry_blank_after(ent, 1)
    remaining = max(0, total - used)
    last_row = WL_CONTENT_ROWS[-1] if WL_CONTENT_ROWS else 39
    next_row = (
        WL_CONTENT_ROWS[used]
        if used < total
        else None
    )
    return {
        "total": total,
        "used": used,
        "remaining": remaining,
        "per_entry": per_entry,
        "last_row": last_row,
        "next_row": next_row,
        "overflow": used > total,
    }


def _render_row_remain_gauge(usage: dict, *, height_px: int = 980) -> None:
    """보기·입력 사이 세로 게이지: 전체 내용칸 대비 남은 칸수.
    저장 버튼 근처까지 세로로 길게 (폭은 좁게 유지).
    """
    total = max(1, int(usage.get("total") or 1))
    used = max(0, int(usage.get("used") or 0))
    rem = max(0, int(usage.get("remaining") or 0))
    overflow = bool(usage.get("overflow"))

    if overflow:
        accent = "#DC2626"
        used_color = "#FECACA"
        rem_color = "#FEE2E2"
        label = "초과"
        big = f"+{used - total}"
        rem_show = 0
        used_show = total
    elif rem <= 3:
        accent = "#D97706"
        used_color = "#E2E8F0"
        rem_color = "#FBBF24"
        label = "남음"
        big = str(rem)
        rem_show = rem
        used_show = min(used, total)
    else:
        accent = "#0F766E"
        used_color = "#E2E8F0"
        rem_color = "#14B8A6"
        label = "남음"
        big = str(rem)
        rem_show = rem
        used_show = min(used, total)

    # 위에서부터 사용칸(회색) → 아래 남은칸(색) 순서의 세그먼트
    segs: list[str] = []
    gap = max(1, int(2 if total <= 20 else 1))
    for i in range(total):
        # i=0 이 맨 위
        is_used = i < used_show
        bg = used_color if is_used else rem_color
        segs.append(
            f'<div style="flex:1 1 0;min-height:3px;border-radius:3px;background:{bg};'
            f'margin:0 0 {gap}px 0;opacity:{0.55 if is_used else 1};"></div>'
        )
    # 마지막 margin 제거용
    if segs:
        segs[-1] = segs[-1].replace(f"margin:0 0 {gap}px 0;", "margin:0;")

    next_row = usage.get("next_row")
    next_txt = f"다음 G{next_row}" if next_row else "칸 끝"
    # 상단 숫자·하단 라벨 영역 제외한 막대 본체 높이
    bar_h = max(520, height_px - 120)

    st.markdown(
        f"""
<div style="
  display:flex;flex-direction:column;align-items:center;justify-content:flex-start;
  gap:6px;padding:2px 2px 0;min-height:{height_px}px;height:{height_px}px;
  font-family:'Pretendard','Apple SD Gothic Neo',sans-serif;
">
  <div style="text-align:center;line-height:1.1;flex:0 0 auto;">
    <div style="font-size:10px;font-weight:700;letter-spacing:.06em;color:#64748B;">{label}</div>
    <div style="font-size:22px;font-weight:800;color:{accent};margin-top:1px;">{big}</div>
    <div style="font-size:10px;color:#94A3B8;">칸</div>
  </div>

  <div style="
    display:flex;flex-direction:column;justify-content:flex-start;flex:1 1 auto;
    width:26px;min-height:{bar_h}px;height:{bar_h}px;padding:4px 3px;
    border-radius:10px;background:#F8FAFC;
    box-shadow:inset 0 0 0 1px #E2E8F0;
  " title="전체 {total}칸 · 사용 {used}칸 · 남음 {rem}칸">
    {"".join(segs)}
  </div>

  <div style="text-align:center;line-height:1.25;flex:0 0 auto;">
    <div style="font-size:11px;font-weight:800;color:#0F172A;">
      <span style="color:{accent};">{rem_show}</span>
      <span style="color:#94A3B8;font-weight:600;"> / {total}</span>
    </div>
    <div style="font-size:10px;color:#64748B;">남은칸 / 전체칸</div>
    <div style="font-size:10px;color:#94A3B8;margin-top:2px;">사용 {used_show}칸 · {next_txt}</div>
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )


def _pack_entries_to_cells(
    d: date,
    entries: list[dict],
    next_day: list[str] | None = None,
    notes: list[str] | None = None,
) -> dict:
    """논리 항목 → 원본 엑셀 C/G/D 칸.

    긴 내용은 내용칸(G:X) 폭에 맞게 아래 행으로 분배하고,
    항목 사이 빈 칸은 각 항목의 blank_after 값을 따른다.
    """
    cells = _empty_cells(d)
    max_u = _content_line_units()
    row_i = 0
    rows = WL_CONTENT_ROWS
    wrote_any = False
    prev_gap = 0
    for ent in entries or []:
        clients = _entry_client_lines(ent)
        chunks = _entry_pack_lines(ent)
        if not any(str(x).strip() for x in clients) and not any(
            (x or "").strip() for x in chunks
        ):
            continue
        # 이전 항목이 지정한 빈 칸만큼 띄움
        if wrote_any:
            row_i += max(0, int(prev_gap))
        n = max(len(clients), len(chunks), 1)
        for j in range(n):
            if row_i >= len(rows):
                break
            r = rows[row_i]
            cells[f"C{r}"] = clients[j] if j < len(clients) else ""
            # 항목 안 빈 내용줄은 공백 한 칸으로 남겨, 항목 사이 완전 빈 행과 구분
            if j < len(chunks):
                g = chunks[j]
                cells[f"G{r}"] = _WL_SOFT_BLANK if g == "" else g
            else:
                cells[f"G{r}"] = ""
            row_i += 1
        wrote_any = True
        prev_gap = _entry_blank_after(ent, 1)

    def _pack_d(texts: list[str], target_rows: list[int]) -> None:
        chunks: list[str] = []
        for t in texts or []:
            chunks.extend(_chunk_text(t, max_u))
        for i, r in enumerate(target_rows):
            cells[f"D{r}"] = chunks[i] if i < len(chunks) else ""

    _pack_d(next_day or [], WL_NEXT_ROWS)
    _pack_d(notes or [], WL_NOTE_ROWS)
    return cells


def _entries_from_cells(cells: dict) -> tuple[list[tuple[str, str]], list[str], list[str]]:
    """미리보기용: 그룹된 (거래처, 내용) / 익일업무 / 특이사 항."""
    rows = [
        (e.get("client") or "", e.get("content") or "")
        for e in _grouped_entries_from_cells(cells)
    ]
    next_day = [(cells.get(f"D{r}") or "").strip() for r in WL_NEXT_ROWS]
    next_day = [x for x in next_day if x]
    notes = [(cells.get(f"D{r}") or "").strip() for r in WL_NOTE_ROWS]
    notes = [x for x in notes if x]
    return rows, next_day, notes


def render_readable_preview_html(d: date, cells: dict) -> str:
    """화면용: 양식 무시, 텍스트만 읽기 쉽게."""
    rows, next_day, notes = _entries_from_cells(cells)
    date_label = html.escape(cells.get("date") or format_worklog_date(d))

    if rows:
        work_items = []
        for i, (client, content) in enumerate(rows, 1):
            c = (
                html.escape(client).replace("\n", "<br>")
                if client
                else "<span class='muted'>(거래처 없음)</span>"
            )
            t = html.escape(content) if content else "<span class='muted'>—</span>"
            work_items.append(
                f"""<div class="item">
                  <div class="idx">{i}</div>
                  <div class="body">
                    <div class="client">{c}</div>
                    <div class="content">{t}</div>
                  </div>
                </div>"""
            )
        work_html = "".join(work_items)
    else:
        work_html = "<div class='empty'>등록된 업무 내용이 없습니다.</div>"

    def _lines(items: list[str], empty_msg: str) -> str:
        if not items:
            return f"<div class='empty'>{empty_msg}</div>"
        return "".join(f"<div class='line'><span class='dot'></span><span>{html.escape(x)}</span></div>" for x in items)

    next_html = _lines(next_day, "익일업무 없음")
    note_html = _lines(notes, "특이사 항 없음")

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');
  html, body {{ margin:0; padding:0; background:transparent; }}
  body {{
    font-family:'Pretendard', 'Apple SD Gothic Neo', sans-serif;
    color:#0F172A; padding:4px;
  }}
  .card {{
    background:linear-gradient(180deg, #F8FAFC 0%, #FFFFFF 48px);
    border:1px solid #E2E8F0; border-radius:14px;
    box-shadow:0 1px 2px rgba(15,23,42,.04);
    overflow:hidden;
  }}
  .head {{
    padding:16px 18px 14px;
    border-bottom:1px solid #E2E8F0;
    background:linear-gradient(135deg, #0F766E 0%, #0E7490 55%, #0369A1 100%);
    color:#fff;
  }}
  .head .title {{ font-size:18px; font-weight:750; letter-spacing:-.02em; }}
  .head .sub {{ margin-top:4px; font-size:13px; opacity:.92; }}
  .sec {{ padding:14px 16px 8px; }}
  .sec h3 {{
    margin:0 0 10px; font-size:12px; font-weight:700; letter-spacing:.06em;
    color:#64748B; text-transform:uppercase;
  }}
  .item {{
    display:flex; gap:12px; align-items:flex-start;
    padding:12px 12px; margin-bottom:8px;
    background:#fff; border:1px solid #E2E8F0; border-radius:10px;
  }}
  .item:hover {{ border-color:#99F6E4; background:#F0FDFA; }}
  .idx {{
    flex:0 0 28px; height:28px; border-radius:8px;
    background:#CCFBF1; color:#0F766E; font-weight:700; font-size:13px;
    display:flex; align-items:center; justify-content:center;
  }}
  .client {{ font-size:15px; font-weight:700; color:#134E4A; margin-bottom:4px; white-space:pre-wrap; }}
  .content {{ font-size:14px; line-height:1.55; color:#334155; white-space:pre-wrap; word-break:break-word; }}
  .muted {{ color:#94A3B8; font-weight:500; }}
  .empty {{
    padding:18px; text-align:center; color:#94A3B8; font-size:13px;
    border:1px dashed #CBD5E1; border-radius:10px; background:#F8FAFC;
  }}
  .panel {{
    margin:0 16px 14px; padding:12px 14px;
    border-radius:12px; border:1px solid #E2E8F0; background:#fff;
  }}
  .panel.next {{ border-left:4px solid #2563EB; }}
  .panel.note {{ border-left:4px solid #D97706; }}
  .panel h3 {{ margin:0 0 8px; font-size:13px; font-weight:700; color:#1E293B; }}
  .line {{
    display:flex; gap:8px; align-items:flex-start;
    padding:6px 0; font-size:14px; line-height:1.5; color:#334155;
    border-bottom:1px solid #F1F5F9;
  }}
  .line:last-child {{ border-bottom:none; }}
  .dot {{
    width:7px; height:7px; margin-top:7px; border-radius:50%;
    background:#94A3B8; flex:0 0 auto;
  }}
  .panel.next .dot {{ background:#2563EB; }}
  .panel.note .dot {{ background:#D97706; }}
  .foot {{
    padding:10px 16px 14px; font-size:11px; color:#94A3B8;
  }}
</style></head>
<body>
  <div class="card">
    <div class="head">
      <div class="title">일일업무일지</div>
      <div class="sub">{date_label}</div>
    </div>
    <div class="sec">
      <h3>거래처 · 내용</h3>
      {work_html}
    </div>
    <div class="panel next">
      <h3>익일업무</h3>
      {next_html}
    </div>
    <div class="panel note">
      <h3>특 이 사 항</h3>
      {note_html}
    </div>
    <div class="foot">인쇄는 상단 「프린터 화면」을 사용하세요.</div>
  </div>
</body></html>"""


def _entries_key(d: date) -> str:
    return f"wl_entries_{d.isoformat()}"


def _next_key(d: date) -> str:
    return f"wl_next_{d.isoformat()}"


def _notes_key(d: date) -> str:
    return f"wl_notes_{d.isoformat()}"


def _boot_key(d: date) -> str:
    return f"worklog_booted_{d.isoformat()}"


def _init_widget_state(d: date) -> dict:
    """엑셀 → 항목(거래처/내용) 세션 상태로 로드. 이미 boot면 디스크 재읽기 생략."""
    bk = _boot_key(d)
    ek = _entries_key(d)
    if st.session_state.get(bk) and ek in st.session_state:
        return {}
    cells = read_worklog_cells(d)
    # 예전 칸별입력 boot 만 있는 경우도 항목 모드로 마이그레이션
    entries = _grouped_entries_from_cells(cells)
    if not entries:
        entries = [{"client": "", "content": ""}]
    st.session_state[ek] = entries
    _, next_day, notes = _entries_from_cells(cells)
    st.session_state[_next_key(d)] = "\n".join(next_day)
    st.session_state[_notes_key(d)] = "\n".join(notes)
    st.session_state[bk] = True
    return cells


def _entry_line_count_key(iso: str, entry_i: int) -> str:
    return f"wl_ent_lc_{iso}_{entry_i}"


def _entry_line_gen_key(iso: str, entry_i: int) -> str:
    return f"wl_ent_gen_{iso}_{entry_i}"


def _entry_line_key(iso: str, entry_i: int, line_j: int) -> str:
    """세대(gN)를 넣어 칸 분할 후 Streamlit이 이전 긴 값을 복원하지 않게 함."""
    g = int(st.session_state.get(_entry_line_gen_key(iso, entry_i), 0) or 0)
    return f"wl_ent_ln_{iso}_{entry_i}_{line_j}_g{g}"


def _bump_entry_line_gen(iso: str, entry_i: int) -> None:
    k = _entry_line_gen_key(iso, entry_i)
    old_g = int(st.session_state.get(k, 0) or 0)
    old_n = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old_n, 0) + 8):
        st.session_state.pop(f"wl_ent_ln_{iso}_{entry_i}_{j}_g{old_g}", None)
        st.session_state.pop(f"wl_ent_ln_{iso}_{entry_i}_{j}", None)
    st.session_state[k] = old_g + 1


def _lines_from_entry_widgets(iso: str, entry_i: int, *, keep_trailing_empty: bool = True) -> list[str]:
    """칸 위젯 → 줄 목록. 중간 빈 칸은 유지. 끝의 편집용 빈 칸 1개는 선택적 유지."""
    lc = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
    if lc > 0:
        parts = [
            str(st.session_state.get(_entry_line_key(iso, entry_i, j), "") or "")
            for j in range(lc)
        ]
    else:
        live = st.session_state.get(_entry_lines_live_key(iso, entry_i))
        if isinstance(live, list):
            parts = [str(x or "") for x in live]
        else:
            raw = str(st.session_state.get(f"wl_ent_t_{iso}_{entry_i}", "") or "")
            parts = [raw] if raw else ([""] if keep_trailing_empty else [])
    if not keep_trailing_empty:
        while parts and parts[-1] == "":
            parts.pop()
    elif parts and parts[-1] != "":
        parts = list(parts) + [""]
    elif not parts and keep_trailing_empty:
        parts = [""]
    return parts


def _content_from_entry_lines(iso: str, entry_i: int) -> str:
    """칸 입력들을 이어붙여 항목 내용 문자열로(미리보기용)."""
    return "\n".join(_lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=False))


def _set_comp_lines_state(
    iso: str, entry_i: int, chunks: list[str], *, focus_j: int | None = None
) -> None:
    """CCv2 상태 키를 깨지 않고 lines/focus만 갱신."""
    ck = _entry_lines_comp_key(iso, entry_i)
    if ck not in st.session_state or not isinstance(st.session_state.get(ck), dict):
        st.session_state[ck] = {"lines": chunks, "focus": -1}
    else:
        st.session_state[ck]["lines"] = chunks
    if focus_j is not None:
        fj = max(0, min(int(focus_j), max(len(chunks) - 1, 0)))
        st.session_state[ck]["focus"] = fj
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_line_key(iso, entry_i, fj)
    st.session_state[_entry_lines_live_key(iso, entry_i)] = list(chunks)


def _apply_entry_lines(
    iso: str,
    entry_i: int,
    lines: list[str],
    *,
    focus_j: int | None = None,
    bump_gen: bool = False,
) -> None:
    """줄 목록을 text_input 키에 반영. 끝에 편집용 빈 칸 1개 유지."""
    if bump_gen:
        _bump_entry_line_gen(iso, entry_i)
    chunks = [str(x or "") for x in (lines or [])]
    if not chunks or chunks[-1] != "":
        chunks.append("")
    old = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(chunks)) + 3):
        st.session_state.pop(_entry_line_key(iso, entry_i, j), None)
    st.session_state[_entry_line_count_key(iso, entry_i)] = len(chunks)
    for j, line in enumerate(chunks):
        st.session_state[_entry_line_key(iso, entry_i, j)] = line
    st.session_state[f"wl_ent_t_{iso}_{entry_i}"] = "\n".join(chunks)
    st.session_state[_entry_lines_live_key(iso, entry_i)] = list(chunks)
    if focus_j is not None:
        fj = max(0, min(int(focus_j), len(chunks) - 1))
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_line_key(iso, entry_i, fj)


def _insert_line_after(iso: str, entry_i: int, line_j: int) -> None:
    """현재 칸 바로 아래 빈 칸을 만들고 포커스 이동."""
    cur = _lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=True)
    if not cur:
        cur = [""]
    while len(cur) <= line_j:
        cur.append("")
    key = _entry_line_key(iso, entry_i, line_j)
    if key in st.session_state:
        cur[line_j] = str(st.session_state.get(key) or "")
    cur.insert(line_j + 1, "")
    _apply_entry_lines(iso, entry_i, cur, focus_j=line_j + 1)


def _entry_client_count_key(iso: str, entry_i: int) -> str:
    return f"wl_ent_clc_{iso}_{entry_i}"


def _entry_client_key(iso: str, entry_i: int, line_j: int) -> str:
    return f"wl_ent_cl_{iso}_{entry_i}_{line_j}"


def _clients_from_widgets(iso: str, entry_i: int, *, keep_trailing_empty: bool = False) -> list[str]:
    """거래처 칸 위젯 → 목록."""
    lc = int(st.session_state.get(_entry_client_count_key(iso, entry_i), 0) or 0)
    if lc > 0:
        parts = [
            str(st.session_state.get(_entry_client_key(iso, entry_i, j), "") or "")
            for j in range(lc)
        ]
    else:
        raw = str(st.session_state.get(f"wl_ent_c_{iso}_{entry_i}", "") or "")
        parts = raw.splitlines() if raw else ([""] if keep_trailing_empty else [])
    if not keep_trailing_empty:
        while parts and not str(parts[-1]).strip():
            parts.pop()
    elif parts and str(parts[-1]).strip() != "":
        parts = list(parts) + [""]
    elif not parts and keep_trailing_empty:
        parts = [""]
    return parts


def _apply_entry_clients(
    iso: str, entry_i: int, lines: list[str], *, focus_j: int | None = None
) -> None:
    """거래처 칸 text_input 키에 반영. 끝에 편집용 빈 칸 1개 유지."""
    chunks = [str(x or "") for x in (lines or [])]
    if not chunks or chunks[-1] != "":
        chunks.append("")
    old = int(st.session_state.get(_entry_client_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(chunks)) + 3):
        st.session_state.pop(_entry_client_key(iso, entry_i, j), None)
    st.session_state[_entry_client_count_key(iso, entry_i)] = len(chunks)
    for j, line in enumerate(chunks):
        st.session_state[_entry_client_key(iso, entry_i, j)] = line
    filled = list(chunks)
    while filled and filled[-1] == "":
        filled.pop()
    st.session_state[f"wl_ent_c_{iso}_{entry_i}"] = "\n".join(filled)
    if focus_j is not None:
        fj = max(0, min(int(focus_j), len(chunks) - 1))
        st.session_state[f"wl_focus_ln_{iso}"] = _entry_client_key(iso, entry_i, fj)


def _seed_entry_clients(iso: str, entry_i: int, client: str | list[str]) -> None:
    max_u = _client_line_units()
    if isinstance(client, list):
        src = [str(x or "") for x in client]
    else:
        raw = str(client or "")
        src = raw.splitlines() if raw else [""]
    chunks: list[str] = []
    for line in src:
        s = str(line or "")
        if not s.strip():
            if not chunks:
                chunks = [""]
            continue
        chunks.extend(_chunk_text(s, max_u) or [s])
    if not chunks:
        chunks = [""]
    fj = max(0, len(chunks) - 1)
    for j, line in enumerate(chunks):
        if _display_units(line) >= max_u:
            fj = min(j + 1, len(chunks))
    _apply_entry_clients(iso, entry_i, chunks, focus_j=fj)


def _insert_client_after(iso: str, entry_i: int, line_j: int) -> None:
    cur = _clients_from_widgets(iso, entry_i, keep_trailing_empty=True)
    if not cur:
        cur = [""]
    while len(cur) <= line_j:
        cur.append("")
    key = _entry_client_key(iso, entry_i, line_j)
    if key in st.session_state:
        cur[line_j] = str(st.session_state.get(key) or "")
    cur.insert(line_j + 1, "")
    _apply_entry_clients(iso, entry_i, cur, focus_j=line_j + 1)


def _split_overflow_parts(parts: list[str], max_u: int) -> list[str]:
    """칸 경계를 유지한 채, 폭 초과인 칸만 잘라 다음 칸으로 넘긴다.

    이미 다음에 같은 넘침 조각이 있으면 중복 삽입하지 않는다.
    """
    out: list[str] = []
    i = 0
    n = len(parts)
    while i < n:
        s = str(parts[i] or "")
        if _display_units(s) > max_u:
            pieces = _chunk_text(s, max_u) or [s]
            out.append(pieces[0])
            j = i + 1
            for ov in pieces[1:]:
                if j < n and str(parts[j] or "") == ov:
                    out.append(str(parts[j] or ""))
                    j += 1
                else:
                    out.append(ov)
            i = j
        else:
            out.append(s)
            i += 1
    if not out:
        out = [""]
    return out


def _dedupe_overflow_tail(pieces: list[str], tail: list[str]) -> list[str]:
    """분할 결과의 넘침 조각이 tail 앞에 이미 있으면 건너뛴다."""
    if len(pieces) <= 1:
        return list(tail)
    rest = list(tail)
    for ov in pieces[1:]:
        if rest and str(rest[0] or "") == ov:
            rest.pop(0)
        else:
            break
    return rest


def _commit_enter_on_cell(
    kind: str, iso: str, entry_i: int, line_j: int, value: str
) -> None:
    """입력값 반영 → 칸폭 초과 시 분할 → 다음 칸(넘친 내용) 포커스."""
    value = str(value or "")
    if kind == "wl_ent_cl":
        max_u = _client_line_units()
        cur = _clients_from_widgets(iso, entry_i, keep_trailing_empty=True)
        while len(cur) <= line_j:
            cur.append("")
        if cur and cur[-1] == "":
            cur = cur[:-1]
        head, tail = cur[:line_j], cur[line_j + 1 :]
        if _display_units(value) > max_u:
            pieces = _chunk_text(value, max_u) or [value]
        else:
            pieces = [value]
        if len(pieces) == 1:
            new = head + pieces + [""] + list(tail)
            focus = line_j + 1
        else:
            new = head + pieces + _dedupe_overflow_tail(pieces, list(tail))
            focus = line_j + 1
            if focus >= len(new):
                new.append("")
        _apply_entry_clients(iso, entry_i, new, focus_j=min(focus, len(new) - 1))
        return

    max_u = _content_line_units()
    cur = _lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=True)
    while len(cur) <= line_j:
        cur.append("")
    if cur and cur[-1] == "":
        cur = cur[:-1]
    # 현재 칸은 emit 값만 사용 (위젯에 남은 긴 값과 이중 분할 방지)
    head, tail = cur[:line_j], cur[line_j + 1 :]
    if _display_units(value) > max_u:
        pieces = _chunk_text(value, max_u) or [value]
    else:
        pieces = [value]
    if len(pieces) == 1:
        new = head + pieces + [""] + list(tail)
        focus = line_j + 1
        _apply_entry_lines(
            iso, entry_i, new, focus_j=min(focus, len(new) - 1), bump_gen=False
        )
        return
    new = head + pieces + _dedupe_overflow_tail(pieces, list(tail))
    focus = line_j + 1
    if focus >= len(new):
        new.append("")
    _apply_entry_lines(
        iso, entry_i, new, focus_j=min(focus, len(new) - 1), bump_gen=True
    )


def _mount_entry_client_editor(iso: str, entry_i: int, max_u: int) -> list[str]:
    """거래처 칸: text_input + Enter/＋ 다음 칸 · 폭 초과 시 자동 분할."""
    # 거래처 칸만 옅은 민트 (내용 칸과 구분). 첫 칸 마운트 시에만 스타일 1회 출력.
    if entry_i == 0:
        st.markdown(
            """
<style>
/* 일일업무일지 — 거래처 text_input 배경 */
div[class*="st-key-wl_ent_cl_"] [data-baseweb="base-input"] > div,
div[class*="st-key-wl_ent_cl_"] [data-baseweb="input"] > div,
div[class*="st-key-wl_ent_cl_"] input {
  background-color: #E7F5F2 !important;
  border-color: #B7DDD4 !important;
}
div[class*="st-key-wl_ent_cl_"] [data-baseweb="base-input"] > div:focus-within,
div[class*="st-key-wl_ent_cl_"] input:focus {
  background-color: #DFF3EE !important;
  border-color: #7CBCAD !important;
}
/* 업무입력칸: 세로 가운데 · 글자는 왼쪽부터 */
div[class*="st-key-wl_ent_cl_"] [data-baseweb="base-input"],
div[class*="st-key-wl_ent_cl_"] [data-baseweb="input"],
div[class*="st-key-wl_ent_ln_"] [data-baseweb="base-input"],
div[class*="st-key-wl_ent_ln_"] [data-baseweb="input"] {
  min-height: 2.35rem !important;
}
div[class*="st-key-wl_ent_cl_"] [data-baseweb="base-input"] > div,
div[class*="st-key-wl_ent_cl_"] [data-baseweb="input"] > div,
div[class*="st-key-wl_ent_ln_"] [data-baseweb="base-input"] > div,
div[class*="st-key-wl_ent_ln_"] [data-baseweb="input"] > div {
  display: flex !important;
  align-items: center !important;
  min-height: 2.35rem !important;
  padding-top: 0 !important;
  padding-bottom: 0 !important;
}
div[class*="st-key-wl_ent_cl_"] input,
div[class*="st-key-wl_ent_ln_"] input {
  text-align: left !important;
  line-height: 1.25 !important;
  padding-top: 0 !important;
  padding-bottom: 0 !important;
  padding-left: 0.5rem !important;
  height: 2.1rem !important;
  min-height: 2.1rem !important;
}
/* 삭제 / ＋ 버튼 글자 가운데 정렬 */
div[class*="st-key-wl_cl_del_"] button,
div[class*="st-key-wl_cl_add_"] button,
div[class*="st-key-wl_ln_del_"] button,
div[class*="st-key-wl_ln_add_"] button {
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  text-align: center !important;
  padding: 0 !important;
  line-height: 1.15 !important;
  min-height: 2.4rem !important;
  height: 2.4rem !important;
}
div[class*="st-key-wl_cl_del_"] button p,
div[class*="st-key-wl_cl_add_"] button p,
div[class*="st-key-wl_ln_del_"] button p,
div[class*="st-key-wl_ln_add_"] button p,
div[class*="st-key-wl_cl_del_"] button span,
div[class*="st-key-wl_cl_add_"] button span,
div[class*="st-key-wl_ln_del_"] button span,
div[class*="st-key-wl_ln_add_"] button span {
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  margin: 0 !important;
  padding: 0 !important;
  line-height: 1.15 !important;
  width: 100%;
  text-align: center !important;
}
</style>
            """,
            unsafe_allow_html=True,
        )

    if int(st.session_state.get(_entry_client_count_key(iso, entry_i), 0) or 0) <= 0:
        raw = str(st.session_state.get(f"wl_ent_c_{iso}_{entry_i}", "") or "")
        _seed_entry_clients(iso, entry_i, raw)

    lc = int(st.session_state.get(_entry_client_count_key(iso, entry_i), 1) or 1)
    parts = [
        str(st.session_state.get(_entry_client_key(iso, entry_i, j), "") or "")
        for j in range(lc)
    ]
    if any(_display_units(p) > max_u for p in parts):
        fixed = _split_overflow_parts(parts, max_u)
        focus = 0
        for j, line in enumerate(fixed):
            if _display_units(line) >= max_u:
                focus = min(j + 1, len(fixed))
        _apply_entry_clients(iso, entry_i, fixed, focus_j=focus)
        lc = int(st.session_state.get(_entry_client_count_key(iso, entry_i), 1) or 1)

    for j in range(lc):
        row_l, row_r = st.columns([8, 1], gap="small")
        with row_l:
            st.text_input(
                f"거래처 {entry_i + 1}-{j + 1}",
                key=_entry_client_key(iso, entry_i, j),
                label_visibility="collapsed",
                placeholder="",
                autocomplete="off",
            )
        with row_r:
            is_last_empty = (
                j == lc - 1
                and str(
                    st.session_state.get(_entry_client_key(iso, entry_i, j), "") or ""
                )
                == ""
            )
            if is_last_empty:
                if st.form_submit_button(
                    "＋",
                    key=f"wl_cl_add_{iso}_{entry_i}_{j}",
                    width="stretch",
                    help="아래 거래처 칸 추가",
                ):
                    st.session_state[f"wl_do_insert_cl_{iso}"] = (entry_i, j)
                    _wl_rerun()
            elif st.form_submit_button(
                "삭제",
                key=f"wl_cl_del_{iso}_{entry_i}_{j}",
                width="stretch",
                help=f"거래처 {j + 1}칸 삭제",
            ):
                st.session_state[f"wl_do_del_cl_{iso}"] = (entry_i, j)
                _wl_rerun()

    out = _clients_from_widgets(iso, entry_i, keep_trailing_empty=True)
    st.session_state[f"wl_ent_c_{iso}_{entry_i}"] = "\n".join(
        _clients_from_widgets(iso, entry_i, keep_trailing_empty=False)
    )
    return out


def _mount_entry_lines_editor(iso: str, entry_i: int, max_u: int) -> list[str]:
    """내용 칸: Streamlit text_input (입력 안정) + 칸 추가/삭제."""
    if int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0) <= 0:
        live = st.session_state.get(_entry_lines_live_key(iso, entry_i))
        if isinstance(live, list) and live:
            _apply_entry_lines(iso, entry_i, [str(x or "") for x in live])
        else:
            _apply_entry_lines(iso, entry_i, [""])

    lc = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 1) or 1)
    parts = [
        str(st.session_state.get(_entry_line_key(iso, entry_i, j), "") or "")
        for j in range(lc)
    ]
    if any(_display_units(p) > max_u for p in parts):
        fixed = _split_overflow_parts(parts, max_u)
        focus = 0
        for j, line in enumerate(fixed):
            if _display_units(line) >= max_u:
                focus = min(j + 1, len(fixed))
        _apply_entry_lines(
            iso, entry_i, fixed, focus_j=focus, bump_gen=True
        )
        lc = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 1) or 1)

    for j in range(lc):
        row_l, row_r = st.columns([8, 1], gap="small")
        with row_l:
            st.text_input(
                f"내용 {entry_i + 1}-{j + 1}",
                key=_entry_line_key(iso, entry_i, j),
                label_visibility="collapsed",
                placeholder="",
                autocomplete="off",
            )
        with row_r:
            is_last_empty = (
                j == lc - 1
                and str(st.session_state.get(_entry_line_key(iso, entry_i, j), "") or "")
                == ""
            )
            if is_last_empty:
                if st.form_submit_button(
                    "＋",
                    key=f"wl_ln_add_{iso}_{entry_i}_{j}",
                    width="stretch",
                    help="아래 칸 추가",
                ):
                    st.session_state[f"wl_do_insert_ln_{iso}"] = (entry_i, j)
                    _wl_rerun()
            elif st.form_submit_button(
                "삭제",
                key=f"wl_ln_del_{iso}_{entry_i}_{j}",
                width="stretch",
                help=f"{j + 1}칸 삭제",
            ):
                st.session_state[f"wl_do_del_ln_{iso}"] = (entry_i, j)
                _wl_rerun()

    out = _lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=True)
    st.session_state[_entry_lines_live_key(iso, entry_i)] = out
    st.session_state[f"wl_ent_t_{iso}_{entry_i}"] = "\n".join(
        _lines_from_entry_widgets(iso, entry_i, keep_trailing_empty=False)
    )
    return out


# Enter / 칸초과 / 방향키 이동 (document 키 훅)
_WL_ENTER_HOOK_JS = r"""
export default function (component) {
  const { data, setTriggerValue } = component;
  const iso = (data && data.iso) || "";
  const focusKey = (data && data.focus_key) || "";
  const focusCaret =
    data && data.focus_caret != null && data.focus_caret !== ""
      ? Number(data.focus_caret)
      : null;
  const clientMax = Number((data && data.client_max_u) || 15);
  const contentMax = Number((data && data.content_max_u) || 70);
  let lastSent = "";
  let lastSig = "";
  let lastAt = 0;
  let lastFocusKey = "";

  function charUnits(ch) {
    const o = ch.charCodeAt(0);
    if (
      (o >= 0xac00 && o <= 0xd7a3) ||
      (o >= 0x1100 && o <= 0x11ff) ||
      (o >= 0x3130 && o <= 0x318f) ||
      (o >= 0x2e80 && o <= 0x9fff) ||
      (o >= 0xff00 && o <= 0xffef)
    )
      return 2;
    return 1;
  }
  function displayUnits(s) {
    let w = 0;
    s = String(s || "");
    for (let i = 0; i < s.length; i++) w += charUnits(s.charAt(i));
    return w;
  }
  function fitByUnits(s, max) {
    s = String(s || "");
    if (displayUnits(s) <= max) return { head: s, tail: "" };
    let acc = 0;
    for (let i = 0; i < s.length; i++) {
      const cu = charUnits(s.charAt(i));
      if (acc + cu > max) {
        if (i === 0) return { head: s.slice(0, 1), tail: s.slice(1) };
        return { head: s.slice(0, i), tail: s.slice(i) };
      }
      acc += cu;
    }
    return { head: s, tail: "" };
  }
  function resolveKey(t) {
    if (!t || String(t.tagName || "").toUpperCase() !== "INPUT") return null;
    const wrap = t.closest
      ? t.closest('[class*="st-key-wl_ent_ln_"],[class*="st-key-wl_ent_cl_"]')
      : null;
    if (!wrap) return null;
    const cls = Array.prototype.find.call(wrap.classList || [], (c) => {
      const s = String(c);
      return (
        s.indexOf("st-key-wl_ent_ln_") !== -1 ||
        s.indexOf("st-key-wl_ent_cl_") !== -1
      );
    });
    if (!cls) return null;
    const key = String(cls).replace(/^st-key-/, "");
    const m = /^(wl_ent_ln|wl_ent_cl)_(\d{4}-\d{2}-\d{2})_(\d+)_(\d+)(?:_g\d+)?$/.exec(key);
    if (!m || m[2] !== iso) return null;
    return {
      key: key,
      kind: m[1],
      ei: Number(m[3]),
      lj: Number(m[4]),
      maxU: m[1] === "wl_ent_cl" ? clientMax : contentMax,
    };
  }
  function listKindInputs(kind) {
    const needle =
      kind === "wl_ent_cl" ? "st-key-wl_ent_cl_" : "st-key-wl_ent_ln_";
    const nodes = document.querySelectorAll(
      'div[class*="' + needle + '"] input'
    );
    const out = [];
    for (let i = 0; i < nodes.length; i++) {
      const info = resolveKey(nodes[i]);
      if (info) out.push(nodes[i]);
    }
    return out;
  }
  function focusInput(el, caret) {
    if (!el) return false;
    try {
      el.focus({ preventScroll: false });
      const n = (el.value || "").length;
      let pos = n;
      if (caret === "start") pos = 0;
      else if (caret === "end") pos = n;
      else if (typeof caret === "number" && Number.isFinite(caret)) {
        pos = Math.max(0, Math.min(n, Math.floor(caret)));
      }
      el.setSelectionRange(pos, pos);
    } catch (e1) {
      try {
        el.focus();
      } catch (e2) {}
    }
    return true;
  }
  function isComposingTarget(t) {
    return !!(t && (t.isComposing || t.composing));
  }
  function findPeer(info, targetKind) {
    const list = listKindInputs(targetKind);
    let best = null;
    let bestScore = 1e9;
    for (let i = 0; i < list.length; i++) {
      const p = resolveKey(list[i]);
      if (!p) continue;
      // 같은 항목 우선, 줄 번호 가까운 칸
      const score =
        Math.abs(p.ei - info.ei) * 1000 + Math.abs(p.lj - info.lj);
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
    // input + compositionend 이중 발행 / 같은 넘침 반복 차단
    if (sig === lastSig && now - lastAt < 700) return;
    lastSig = sig;
    lastAt = now;
    const payload = JSON.stringify({ key: key, t: now, v: v });
    if (payload === lastSent) return;
    lastSent = payload;
    setTriggerValue("enter", payload);
  }

  const onKey = (e) => {
    if (e.isComposing || e.keyCode === 229) return;
    const info = resolveKey(e.target);
    if (!info) return;
    const t = e.target;
    const start = typeof t.selectionStart === "number" ? t.selectionStart : 0;
    const end = typeof t.selectionEnd === "number" ? t.selectionEnd : start;
    const len = String(t.value || "").length;
    const caretAll = start === end;

    if (e.key === "Enter") {
      e.preventDefault();
      e.stopPropagation();
      // 서버 rerun 없이 다음 칸으로만 이동 (form 입력 중 로딩·값 유실 방지)
      // 칸 추가는 「＋」/저장 시 서버에서 처리
      const list = listKindInputs(info.kind);
      const idx = list.indexOf(t);
      if (idx >= 0 && idx < list.length - 1) {
        focusInput(list[idx + 1], "end");
      }
      return;
    }

    // 방향키: 한 줄 입력이므로 위/아래는 항상 칸 이동
    // 좌/우는 커서 끝·앞에서 거래처↔내용 이동
    if (e.key === "ArrowUp") {
      const list = listKindInputs(info.kind);
      const idx = list.indexOf(t);
      if (idx > 0) {
        e.preventDefault();
        e.stopPropagation();
        focusInput(list[idx - 1], "end");
      } else {
        // 맨 위면 반대편 열의 같은 위치(또는 직전)로
        const peerKind =
          info.kind === "wl_ent_cl" ? "wl_ent_ln" : "wl_ent_cl";
        const peer = findPeer(info, peerKind);
        if (peer && peer !== t) {
          e.preventDefault();
          e.stopPropagation();
          focusInput(peer, "end");
        }
      }
      return;
    }
    if (e.key === "ArrowDown") {
      const list = listKindInputs(info.kind);
      const idx = list.indexOf(t);
      if (idx >= 0 && idx < list.length - 1) {
        e.preventDefault();
        e.stopPropagation();
        focusInput(list[idx + 1], "end");
      } else {
        const peerKind =
          info.kind === "wl_ent_cl" ? "wl_ent_ln" : "wl_ent_cl";
        const peer = findPeer(
          { ...info, lj: info.lj + 1 },
          peerKind
        );
        if (peer && peer !== t) {
          e.preventDefault();
          e.stopPropagation();
          focusInput(peer, "start");
        }
      }
      return;
    }
    if (e.key === "ArrowLeft" && caretAll && start === 0) {
      const peer =
        info.kind === "wl_ent_ln"
          ? findPeer(info, "wl_ent_cl")
          : (() => {
              const list = listKindInputs("wl_ent_cl");
              const idx = list.indexOf(t);
              // 거래처 칸에서 왼쪽: 이전 거래처 칸
              return idx > 0 ? list[idx - 1] : null;
            })();
      if (peer) {
        e.preventDefault();
        e.stopPropagation();
        focusInput(peer, "end");
      }
      return;
    }
    if (e.key === "ArrowRight" && caretAll && start === len) {
      const peer =
        info.kind === "wl_ent_cl"
          ? findPeer(info, "wl_ent_ln")
          : (() => {
              const list = listKindInputs("wl_ent_ln");
              const idx = list.indexOf(t);
              return idx >= 0 && idx < list.length - 1 ? list[idx + 1] : null;
            })();
      if (peer) {
        e.preventDefault();
        e.stopPropagation();
        focusInput(peer, "start");
      }
      return;
    }
  };

  // 칸 폭 초과 시 DOM만 자르고 서버로는 보내지 않음 (입력 중 로딩 방지)
  // Enter 시에만 emit → 서버에서 다음 칸 분할
  const onInput = (e) => {
    if (e.isComposing) return;
    const t = e.target;
    const info = resolveKey(t);
    if (!info) return;
    const v = String(t.value || "");
    if (displayUnits(v) <= info.maxU) return;
    const { head } = fitByUnits(v, info.maxU);
    try {
      t.value = head;
    } catch (err) {}
  };

  const onCompEnd = (e) => {
    const t = e.target;
    const info = resolveKey(t);
    if (!info) return;
    const v = String(t.value || "");
    if (displayUnits(v) <= info.maxU) return;
    const { head } = fitByUnits(v, info.maxU);
    try {
      t.value = head;
    } catch (err) {}
  };

  document.addEventListener("keydown", onKey, true);
  document.addEventListener("input", onInput, true);
  document.addEventListener("compositionend", onCompEnd, true);
  // focus → setTriggerValue 제거: 칸 이동만으로 fragment 로딩이 나던 원인

  if (focusKey) {
    const go = () => {
      const el = document.querySelector(
        'div[class*="st-key-' + focusKey + '"] input'
      );
      if (!el) return false;
      const caret =
        focusCaret != null && Number.isFinite(focusCaret) ? focusCaret : "end";
      focusInput(el, caret);
      return true;
    };
    go();
    setTimeout(go, 50);
    setTimeout(go, 150);
    setTimeout(go, 350);
    setTimeout(go, 600);
  }

  return () => {
    document.removeEventListener("keydown", onKey, true);
    document.removeEventListener("input", onInput, true);
    document.removeEventListener("compositionend", onCompEnd, true);
  };
}
"""

_WL_ENTER_HOOK = st.components.v2.component(
    "worklog_cell_nav_hook_v11",
    js=_WL_ENTER_HOOK_JS,
)



_WL_SPECIAL_CHARS = [
    "※",
    "★",
    "☆",
    "○",
    "●",
    "◎",
    "▲",
    "△",
    "■",
    "□",
    "→",
    "←",
    "·",
    "～",
    "【",
    "】",
    "「",
    "」",
    "①",
    "②",
    "③",
    "✓",
    "◇",
    "◆",
]


_WL_SPECIAL_BAR_HTML = """
<div class="wl-sp"></div>
"""

_WL_SPECIAL_BAR_CSS = """
.wl-sp {
  display: flex;
  flex-wrap: nowrap;
  gap: 4px;
  width: 100%;
  overflow-x: auto;
  padding: 2px 0;
  box-sizing: border-box;
}
.wl-sp button {
  flex: 1 0 auto;
  min-width: 1.35rem;
  height: 1.55rem;
  margin: 0;
  padding: 0 2px;
  border: 1px solid #cbd5e1;
  border-radius: 0.35rem;
  background: #f8fafc;
  color: #0f172a;
  font-size: 0.85rem;
  line-height: 1.55rem;
  cursor: pointer;
}
.wl-sp button:hover {
  background: #e2e8f0;
}
"""

_WL_SPECIAL_BAR_JS = r"""
export default function (component) {
  const { data, parentElement, setTriggerValue } = component;
  const iso = (data && data.iso) || "";
  const chars = (data && data.chars) || [];
  let last = null;

  function resolveKey(t) {
    if (!t || String(t.tagName || "").toUpperCase() !== "INPUT") return null;
    const wrap = t.closest
      ? t.closest('[class*="st-key-wl_ent_ln_"],[class*="st-key-wl_ent_cl_"]')
      : null;
    if (!wrap) return null;
    const cls = Array.prototype.find.call(wrap.classList || [], (c) => {
      const s = String(c);
      return (
        s.indexOf("st-key-wl_ent_ln_") !== -1 ||
        s.indexOf("st-key-wl_ent_cl_") !== -1
      );
    });
    if (!cls) return null;
    const key = String(cls).replace(/^st-key-/, "");
    const m = /^(wl_ent_ln|wl_ent_cl)_(\d{4}-\d{2}-\d{2})_(\d+)_(\d+)(?:_g\d+)?$/.exec(key);
    if (!m || m[2] !== iso) return null;
    return { key: key, el: t };
  }

  function remember(t) {
    const info = resolveKey(t);
    if (!info) return;
    let s = 0;
    let e = 0;
    try {
      s = typeof info.el.selectionStart === "number" ? info.el.selectionStart : 0;
      e = typeof info.el.selectionEnd === "number" ? info.el.selectionEnd : s;
    } catch (err) {}
    last = { key: info.key, el: info.el, s: s, e: e };
  }

  function insertChar(ch) {
    const t = document.activeElement;
    const live = resolveKey(t);
    const src = live || last;
    if (!src || !src.key) {
      try {
        setTriggerValue("insert", JSON.stringify({ miss: true, ch: ch, t: Date.now() }));
      } catch (err) {}
      return;
    }
    const el = live ? live.el : src.el;
    let s = 0;
    let e = 0;
    try {
      if (el && typeof el.selectionStart === "number") {
        s = el.selectionStart;
        e = typeof el.selectionEnd === "number" ? el.selectionEnd : s;
      } else if (last && last.key === src.key) {
        s = last.s;
        e = last.e;
      } else {
        s = e = String((el && el.value) || "").length;
      }
    } catch (err2) {
      s = e = String((el && el.value) || "").length;
    }
    const cur = String((el && el.value) || "");
    const next = cur.slice(0, s) + ch + cur.slice(e);
    const pos = s + ch.length;
    try {
      if (el) {
        el.value = next;
        el.setSelectionRange(pos, pos);
      }
    } catch (err3) {}
    try {
      setTriggerValue(
        "insert",
        JSON.stringify({ key: src.key, v: next, s: pos, ch: ch, t: Date.now() })
      );
    } catch (err4) {}
  }

  let root = parentElement.querySelector(".wl-sp");
  if (!root) {
    root = document.createElement("div");
    root.className = "wl-sp";
    parentElement.appendChild(root);
  }
  if (root.childElementCount !== chars.length) {
    root.innerHTML = "";
    for (let i = 0; i < chars.length; i++) {
      const ch = String(chars[i] || "");
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = ch;
      btn.title = ch + " 삽입";
      btn.addEventListener("pointerdown", function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        insertChar(ch);
      });
      root.appendChild(btn);
    }
  }

  const onFocusIn = (e) => remember(e.target);
  const onSel = (e) => {
    if (e && (e.isComposing || e.keyCode === 229)) return;
    remember(e.target || document.activeElement);
  };
  document.addEventListener("focusin", onFocusIn, true);
  document.addEventListener("keyup", onSel, true);
  document.addEventListener("mouseup", onSel, true);

  return () => {
    document.removeEventListener("focusin", onFocusIn, true);
    document.removeEventListener("keyup", onSel, true);
    document.removeEventListener("mouseup", onSel, true);
  };
}
"""

_WL_SPECIAL_BAR = st.components.v2.component(
    "worklog_special_bar_v1",
    html=_WL_SPECIAL_BAR_HTML,
    css=_WL_SPECIAL_BAR_CSS,
    js=_WL_SPECIAL_BAR_JS,
)


def _render_worklog_special_chars(iso: str) -> None:
    """자주 쓰는 특수문자 — HTML 버튼, 커서 위치에 삽입."""

    def _on_insert():
        stt = st.session_state.get(f"wl_sp_bar_{iso}") or {}
        if isinstance(stt, dict):
            raw = stt.get("insert")
        else:
            raw = getattr(stt, "insert", None)
        if not raw:
            return
        try:
            obj = json.loads(str(raw))
            fk = str(obj.get("key") or "")
            pos = int(obj.get("s") or 0)
        except Exception:
            return
        if obj.get("miss"):
            st.session_state["wl_special_msg"] = "칸을 먼저 클릭한 뒤 특수문자를 누르세요"
            return
        if not (fk.startswith("wl_ent_ln_") or fk.startswith("wl_ent_cl_")):
            return
        sig = f"{fk}\0{obj.get('v')}\0{pos}\0{obj.get('t')}"
        done_k = f"wl_special_done_{iso}"
        if st.session_state.get(done_k) == sig:
            return
        st.session_state[done_k] = sig
        st.session_state[f"wl_do_special_{iso}"] = {
            "key": fk,
            "v": str(obj.get("v") if obj.get("v") is not None else ""),
            "s": pos,
            "ch": str(obj.get("ch") or ""),
        }

    _WL_SPECIAL_BAR(
        key=f"wl_sp_bar_{iso}",
        data={"iso": iso, "chars": list(_WL_SPECIAL_CHARS)},
        on_insert_change=_on_insert,
        width="stretch",
        height=36,
    )


def _apply_special_insert(iso: str, fk: str, val: str, pos: int, ch: str) -> None:
    """위젯 생성 전: 특수문자가 반영된 칸 값을 심고 포커스를 되돌린다."""
    val = str(val or "")
    pos = max(0, min(int(pos or 0), len(val)))
    m = re.match(
        r"^(wl_ent_ln|wl_ent_cl)_(\d{4}-\d{2}-\d{2})_(\d+)_(\d+)",
        fk,
    )
    if not m or m.group(2) != iso:
        return
    kind, ei, lj = m.group(1), int(m.group(3)), int(m.group(4))
    if kind == "wl_ent_ln":
        cur = _lines_from_entry_widgets(iso, ei, keep_trailing_empty=True)
        while len(cur) <= lj:
            cur.append("")
        cur[lj] = val
        _apply_entry_lines(iso, ei, cur, focus_j=lj, bump_gen=True)
    else:
        cur = _clients_from_widgets(iso, ei, keep_trailing_empty=True)
        while len(cur) <= lj:
            cur.append("")
        cur[lj] = val
        _apply_entry_clients(iso, ei, cur, focus_j=lj)
    st.session_state["wl_active_cell_key"] = st.session_state.get(
        f"wl_focus_ln_{iso}"
    ) or fk
    st.session_state["wl_active_cell_sel"] = (pos, pos)
    st.session_state[f"wl_focus_caret_{iso}"] = pos
    st.session_state["wl_special_msg"] = f"「{ch}」삽입" if ch else "특수문자 삽입"


def _seed_entry_lines(
    iso: str,
    entry_i: int,
    content: str,
    *,
    focus_j: int | None = None,
    focus_last: bool = False,
) -> None:
    """내용 문자열 → 칸 목록으로 시드."""
    max_u = _content_line_units()
    chunks = _chunk_text(content or "", max_u) or [""]
    fj = focus_j
    if fj is None and focus_last:
        fj = max(0, len(chunks) - 1)
        for j, line in enumerate(chunks):
            if _display_units(line) >= max_u:
                fj = min(j + 1, len(chunks))
    _apply_entry_lines(iso, entry_i, chunks, focus_j=fj)


def _read_editor_entries(d: date) -> list[dict]:
    """위젯 키에서 현재 항목 목록 수집."""
    iso = d.isoformat()
    n = int(st.session_state.get(f"wl_entry_count_{iso}", 0) or 0)
    stored = st.session_state.get(_entries_key(d)) or [
        {"client": "", "content": "", "blank_after": 1}
    ]
    if n <= 0:
        n = len(stored)
    out: list[dict] = []
    for i in range(n):
        ck = f"wl_ent_c_{iso}_{i}"
        gk = f"wl_ent_gap_{iso}_{i}"
        lc = int(st.session_state.get(_entry_line_count_key(iso, i), 0) or 0)
        if ck in st.session_state or lc > 0 or f"wl_ent_t_{iso}_{i}" in st.session_state or int(
            st.session_state.get(_entry_client_count_key(iso, i), 0) or 0
        ) > 0:
            if int(st.session_state.get(_entry_client_count_key(iso, i), 0) or 0) > 0:
                client_lines = _clients_from_widgets(
                    iso, i, keep_trailing_empty=False
                )
                client = "\n".join(client_lines)
            else:
                client = str(st.session_state.get(ck, "") or "")
                client_lines = _entry_client_lines({"client": client})
            lines = _lines_from_entry_widgets(iso, i, keep_trailing_empty=False)
            content = "\n".join(lines)
            if gk in st.session_state:
                blank_after = _entry_blank_after({"blank_after": st.session_state.get(gk)}, 1)
            elif i < len(stored):
                blank_after = _entry_blank_after(stored[i], 1)
            else:
                blank_after = 1
        elif i < len(stored):
            client = str(stored[i].get("client") or "")
            if not client and isinstance(stored[i].get("client_lines"), list):
                client = "\n".join(str(x or "") for x in stored[i].get("client_lines") or [])
            client_lines = _entry_client_lines(stored[i])
            lines = stored[i].get("lines")
            if not isinstance(lines, list):
                lines = _chunk_text(str(stored[i].get("content") or ""), _content_line_units()) or []
            content = (
                "\n".join(str(x or "") for x in lines)
                if lines
                else str(stored[i].get("content") or "")
            )
            blank_after = _entry_blank_after(stored[i], 1)
        else:
            client, content, blank_after, lines, client_lines = "", "", 1, [], []
        out.append(
            {
                "client": client,
                "client_lines": client_lines,
                "content": content,
                "lines": lines,
                "blank_after": blank_after,
            }
        )
    return out or [{"client": "", "client_lines": [], "content": "", "lines": [], "blank_after": 1}]


def _cells_from_widgets(d: date) -> dict:
    """항목 입력 → 원본 엑셀 칸 구조."""
    entries = _read_editor_entries(d)
    next_raw = str(st.session_state.get(_next_key(d), "") or "")
    notes_raw = str(st.session_state.get(_notes_key(d), "") or "")
    # text_area 키가 있으면 우선
    nk = f"wl_next_area_{d.isoformat()}"
    ok = f"wl_notes_area_{d.isoformat()}"
    if nk in st.session_state:
        next_raw = str(st.session_state.get(nk) or "")
    if ok in st.session_state:
        notes_raw = str(st.session_state.get(ok) or "")
    next_day = [x.strip() for x in next_raw.splitlines() if x.strip()]
    notes = [x.strip() for x in notes_raw.splitlines() if x.strip()]
    return _pack_entries_to_cells(d, entries, next_day, notes)


def _clear_date_widget_state(d: date) -> None:
    iso = d.isoformat()
    prefixes = (
        f"wl_ent_c_{iso}_",
        f"wl_ent_t_{iso}_",
        f"wl_ent_gap_{iso}_",
        f"wl_ent_ln_{iso}_",
        f"wl_ent_lc_{iso}_",
        f"wl_ent_gen_{iso}_",
        f"wl_ent_cl_{iso}_",
        f"wl_ent_clc_{iso}_",
        f"wl_ent_rev_{iso}_",
        f"wl_lines_comp_{iso}_",
        f"wl_lines_live_{iso}_",
        f"wl_exp_{iso}_",
        f"wl_entries_{iso}",
        f"wl_next_{iso}",
        f"wl_notes_{iso}",
        f"wl_next_area_{iso}",
        f"wl_notes_area_{iso}",
        f"wl_entry_count_{iso}",
        f"worklog_booted_{iso}",
        f"wl_save_btn_{iso}",
        f"wl_focus_ln_{iso}",
    )
    for k in list(st.session_state.keys()):
        if not isinstance(k, str):
            continue
        if k in prefixes or any(k.startswith(p) for p in prefixes if p.endswith("_")):
            del st.session_state[k]
        elif k in {
            f"wl_entries_{iso}",
            f"wl_next_{iso}",
            f"wl_notes_{iso}",
            f"wl_entry_count_{iso}",
            f"worklog_booted_{iso}",
            f"wl_next_area_{iso}",
            f"wl_notes_area_{iso}",
            f"wl_pending_sync_{iso}",
            f"wl_do_add_{iso}",
            f"wl_do_del_{iso}",
            f"wl_focus_ln_{iso}",
        }:
            del st.session_state[k]


def _preview_path(d: date) -> str:
    return os.path.join(WORKLOG_DIR, f"_preview_{d.isoformat()}.xlsx")


def _build_preview_file(d: date, cells: dict) -> str:
    """날짜 파일은 건드리지 않고 미리보기용 임시 xlsx 생성.

    항상 템플릿에서 새로 만들어, 이전 저장본 잔여 값이 섞이지 않게 한다.
    """
    _ensure_dirs()
    if not os.path.exists(WORKLOG_TEMPLATE):
        raise FileNotFoundError("업무일지 템플릿이 없습니다.")
    dst = _preview_path(d)
    shutil.copy2(WORKLOG_TEMPLATE, dst)
    write_cells_to_path(dst, d, cells, blank_base=True)
    return dst


def _excel_app_path() -> str | None:
    candidates = [
        "/Applications/Microsoft Excel.app",
        os.path.expanduser("~/Applications/Microsoft Excel.app"),
    ]
    for p in candidates:
        if os.path.isdir(p):
            return p
    return None


def _print_xlsx_path(d: date) -> str:
    """Excel에서 열 인쇄용 파일 (한글 파일명)."""
    return os.path.join(WORKLOG_DIR, f"일일업무일지_{d.isoformat()}_인쇄.xlsx")


def prepare_print_xlsx(d: date, cells: dict) -> str:
    """현재 입력값을 원본 양식 xlsx로 만들어 절대경로 반환."""
    _ensure_dirs()
    dst = _print_xlsx_path(d)
    src = worklog_path(d) if os.path.exists(worklog_path(d)) else WORKLOG_TEMPLATE
    shutil.copy2(src, dst)
    write_cells_to_path(dst, d, cells, blank_base=not os.path.exists(worklog_path(d)))
    return os.path.abspath(dst)


def open_excel_print_preview(
    xlsx_path: str, *, prefer_print_dialog: bool = True
) -> tuple[bool, str]:
    """
    macOS + Microsoft Excel:
    1) 원본 xlsx 열기
    2) prefer_print_dialog=True → ⌘P 인쇄(미리보기) 대화상자 (요청 화면)
       prefer_print_dialog=False → Excel 인쇄 미리보기 메뉴 우선
    """
    abs_path = os.path.abspath(xlsx_path)
    if not os.path.exists(abs_path):
        return False, "미리보기용 엑셀 파일이 없습니다."
    if platform.system() != "Darwin":
        return False, (
            "Excel 인쇄 화면은 맥에서 로컬 대시보드 실행 시에만 자동 연결됩니다. "
            "Cloud에서는 「엑셀 저장본」다운로드 후 Excel에서 ⌘P 로 여세요."
        )

    if not _excel_app_path():
        try:
            subprocess.Popen(["open", abs_path], start_new_session=True)
            return True, "파일을 열었습니다. Excel이 없다면 설치 후 다시 시도해 주세요."
        except Exception as e:
            return False, f"파일 열기 실패: {e}"

    ap = abs_path.replace("\\", "\\\\").replace('"', '\\"')
    if prefer_print_dialog:
        # 사용자가 원하는 macOS「프린트」대화상자(미리보기+Excel 옵션)
        script = f'''
set targetFile to POSIX file "{ap}"
tell application "Microsoft Excel"
    activate
    open targetFile
    delay 1.2
end tell
tell application "System Events"
    if exists process "Microsoft Excel" then
        tell process "Microsoft Excel"
            set frontmost to true
            delay 0.4
            keystroke "p" using {{command down}}
        end tell
    end if
end tell
return true
'''
        ok_msg = "Excel에서 열어 인쇄(미리보기) 화면까지 연결했습니다."
    else:
        script = f'''
set targetFile to POSIX file "{ap}"
set previewDone to false
tell application "Microsoft Excel"
    activate
    open targetFile
    delay 1.3
    try
        print preview active sheet
        set previewDone to true
    end try
    if previewDone is false then
        try
            print preview
            set previewDone to true
        end try
    end if
end tell
if previewDone is false then
    tell application "System Events"
        if exists process "Microsoft Excel" then
            tell process "Microsoft Excel"
                set frontmost to true
                delay 0.35
                try
                    click menu item "인쇄 미리 보기" of menu "파일" of menu bar 1
                    set previewDone to true
                end try
                if previewDone is false then
                    keystroke "p" using {{command down}}
                    set previewDone to true
                end if
            end tell
        end if
    end tell
end if
return previewDone
'''
        ok_msg = "Excel에서 열어 인쇄 미리보기까지 연결했습니다."

    try:
        r = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=50,
        )
        if r.returncode == 0:
            return True, ok_msg
        subprocess.Popen(
            ["open", "-a", "Microsoft Excel", abs_path],
            start_new_session=True,
        )
        err = (r.stderr or r.stdout or "").strip()
        hint = ""
        if "Not authorized" in err or "assistive" in err.lower() or "1002" in err:
            hint = (
                " (시스템 설정 → 개인정보 보호 → 손쉬운 사용에서 "
                "Terminal/Cursor 제어를 허용하면 ⌘P 화면까지 자동으로 열립니다)"
            )
        return True, (
            "Excel에서 파일을 열었습니다. 인쇄 화면은 Excel에서 ⌘P 로 열어 주세요."
            + hint
        )
    except subprocess.TimeoutExpired:
        return False, "Excel 응답 시간 초과. Excel에서 파일을 직접 열어 주세요."
    except Exception as e:
        try:
            subprocess.Popen(
                ["open", "-a", "Microsoft Excel", abs_path],
                start_new_session=True,
            )
            return True, f"Excel에서 파일을 열었습니다. (인쇄 화면 자동화 실패: {e})"
        except Exception as e2:
            return False, f"Excel 실행 실패: {e2}"


def _launch_browser_print_dialog(xlsx_path: str) -> None:
    """본화면을 유지한 채, 브라우저 인쇄 대화상자를 연다.

    Streamlit 버튼 클릭 → 재실행 후에는 window.open이 막히므로,
    components.html에 문서를 직접 넣고 window.print()를 호출한다.
    (JSON을 <script>에 넣으면 </script> 때문에 스크립트가 깨져 버튼이 먹통처럼 보였음)
    """
    st.session_state["wl_print_panel"] = False
    abs_path = os.path.abspath(xlsx_path)
    if not os.path.exists(abs_path):
        st.error("인쇄용 파일이 없습니다.")
        return
    try:
        doc_html = render_worklog_view_html(
            abs_path, print_mode=True, auto_print=True, scale=1.0
        )
    except Exception as e:
        st.error(f"인쇄 문서 준비 실패: {e}")
        return
    # 클릭마다 iframe을 다시 마운트해야 print가 다시 실행됨
    nonce = int(st.session_state.get("wl_print_n", 0)) + 1
    st.session_state["wl_print_n"] = nonce
    stamped = doc_html.replace("<title>", f"<title><!--wl-print-{nonce}-->", 1)
    # 상단 「인쇄 창 열기」버튼이 보이도록 높이 확보 (본문은 print CSS로 출력)
    components.html(stamped, height=56, scrolling=False)
    st.caption("인쇄 창이 열립니다. 안 뜨면 위 「인쇄 창 열기」를 눌러 주세요.")


def _open_worklog_print_panel(xlsx_path: str, *, auto: bool = False) -> None:
    """인쇄 미리보기 패널(폴백). 가능하면 _launch_browser_print_dialog 사용."""
    st.session_state["wl_print_panel"] = True
    st.session_state["wl_dialog_preview_path"] = os.path.abspath(xlsx_path)
    st.session_state["wl_print_auto_once"] = False
    _ = auto


def _render_worklog_print_panel() -> bool:
    """인쇄 미리보기만 표시 (자동 window.print 팝업 없음)."""
    if not st.session_state.get("wl_print_panel"):
        return False
    path = st.session_state.get("wl_dialog_preview_path")
    st.markdown(
        """
        <style>
        /* 인쇄 패널일 때 상단 고정 필터/탭 숨김 — 미리보기만 */
        .dashboard-filter-sticky,
        #dashboard-top-shield,
        #dashboard-sticky-spacer { display: none !important; }
        section.main .block-container { padding-top: 0.4rem !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    top1, top2 = st.columns([1.1, 3])
    with top1:
        if st.button(
            "← 본화면으로",
            type="primary",
            width="stretch",
            key="wl_print_back_home",
        ):
            st.session_state["wl_print_panel"] = False
            st.session_state["wl_print_auto_once"] = False
            _wl_rerun()
    with top2:
        st.markdown("##### 인쇄 미리보기")
        st.caption("자동 팝업 없음 · 「인쇄하기」만 누르면 인쇄 창이 열립니다.")
    if not path or not os.path.exists(str(path)):
        st.error("인쇄용 파일이 없습니다. 본화면으로 돌아가 다시 시도해 주세요.")
        return True
    path = str(path)
    try:
        print_html = render_worklog_view_html(
            path, print_mode=True, auto_print=False, scale=1.0
        )
        components.html(print_html, height=900, scrolling=True)
    except Exception as e:
        st.error(f"인쇄 미리보기 표시 실패: {e}")
    b1, b2 = st.columns(2)
    with b1:
        try:
            with open(path, "rb") as f:
                xbytes = f.read()
            st.download_button(
                "엑셀 다운로드",
                data=xbytes,
                file_name=os.path.basename(path) or "일일업무일지.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                key="wl_print_panel_dl",
            )
        except Exception:
            pass
    with b2:
        if platform.system() == "Darwin":
            if st.button("Excel로 인쇄", width="stretch", key="wl_print_panel_excel"):
                ok, msg = open_excel_print_preview(path, prefer_print_dialog=True)
                (st.success if ok else st.warning)(msg)
    return True


@st.dialog("원본 엑셀 양식 미리보기", width="large")
def _worklog_form_preview_dialog() -> None:
    """큰 화면용 엑셀 양식 미리보기 (보조)."""
    path = st.session_state.get("wl_dialog_preview_path")
    if not path or not os.path.exists(str(path)):
        st.error("미리보기 파일을 만들 수 없습니다. 템플릿·입력을 확인해 주세요.")
        return
    path = str(path)
    try:
        scale = 0.62
        print_html = render_worklog_view_html(
            path, print_mode=False, auto_print=False, scale=scale
        )
        _, frame_h = _scaled_view_frame_size(path, scale)
        components.html(
            print_html,
            height=min(760, max(420, int(frame_h))),
            scrolling=True,
        )
    except Exception as e:
        st.error(f"미리보기 표시 실패: {e}")
        return
    b1, b2, b3 = st.columns(3)
    with b1:
        try:
            with open(path, "rb") as f:
                xbytes = f.read()
            st.download_button(
                "엑셀 다운로드",
                data=xbytes,
                file_name=os.path.basename(path) or "일일업무일지_미리보기.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                key="wl_dialog_dl_xlsx",
            )
        except Exception:
            st.caption("엑셀 다운로드를 준비하지 못했습니다.")
    with b2:
        if st.button("프린터 화면", width="stretch", key="wl_dialog_browser_print"):
            _launch_browser_print_dialog(path)
    with b3:
        if platform.system() == "Darwin":
            if st.button(
                "Excel 인쇄 화면",
                width="stretch",
                key="wl_dialog_open_excel",
            ):
                ok, msg = open_excel_print_preview(path, prefer_print_dialog=True)
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)
        else:
            st.caption("Excel 앱 인쇄는 맥 로컬에서 가능합니다.")


def _prepare_excel_preview(d: date, cells: dict) -> str:
    """현재 입력 → 엑셀 미리보기 파일 경로."""
    try:
        return prepare_print_xlsx(d, cells)
    except Exception:
        return _build_preview_file(d, cells)


def _render_month_calendar(selected: date, saved: set[str]) -> date | None:
    """팝업 내부용 월 달력 (컴팩트)."""
    if "worklog_month" not in st.session_state:
        st.session_state["worklog_month"] = date(selected.year, selected.month, 1)
    month_anchor: date = st.session_state["worklog_month"]

    st.markdown(
        """
        <style>
        /* 달력 팝업 폭·날짜 칸 축소 */
        div[data-testid="stPopoverBody"] {
          max-width: 268px !important;
          width: 268px !important;
          padding: 0.35rem 0.45rem 0.5rem !important;
        }
        div[data-testid="stPopoverBody"] div[class*="st-key-wl_day_"] button,
        div[data-testid="stPopoverBody"] div[class*="st-key-wl_prev_month"] button,
        div[data-testid="stPopoverBody"] div[class*="st-key-wl_next_month"] button,
        div[data-testid="stPopoverBody"] div[class*="st-key-wl_today"] button {
          min-height: 1.55rem !important;
          height: 1.55rem !important;
          padding: 0 0.15rem !important;
          font-size: 0.7rem !important;
          line-height: 1.1 !important;
          border-radius: 5px !important;
        }
        div[data-testid="stPopoverBody"] [data-testid="stCaptionContainer"] {
          font-size: 0.65rem !important;
          margin-bottom: 0.15rem !important;
        }
        div[data-testid="stPopoverBody"] [data-testid="stHorizontalBlock"] {
          gap: 0.2rem !important;
        }
        div[data-testid="stPopoverBody"] [data-testid="column"] {
          padding: 0 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    nav = st.columns([0.85, 0.85, 2.6, 1.2], gap="small")
    with nav[0]:
        if st.button("◀", key="wl_prev_month", width="stretch"):
            y, m = month_anchor.year, month_anchor.month - 1
            if m < 1:
                y, m = y - 1, 12
            st.session_state["worklog_month"] = date(y, m, 1)
            _wl_rerun()
    with nav[1]:
        if st.button("▶", key="wl_next_month", width="stretch"):
            y, m = month_anchor.year, month_anchor.month + 1
            if m > 12:
                y, m = y + 1, 1
            st.session_state["worklog_month"] = date(y, m, 1)
            _wl_rerun()
    with nav[2]:
        st.markdown(
            f"<div style='text-align:center;font-weight:700;font-size:12px;padding:2px 0;line-height:1.2;'>"
            f"{month_anchor.year}년 {month_anchor.month}월</div>",
            unsafe_allow_html=True,
        )
    with nav[3]:
        if st.button("오늘", key="wl_today", width="stretch"):
            today = date.today()
            st.session_state["worklog_selected"] = today
            st.session_state["worklog_month"] = date(today.year, today.month, 1)
            # date_input 동기화는 다음 런 위젯 생성 전에 수행
            st.session_state["wl_date_sync"] = ""
            _wl_rerun()

    st.caption("• = 저장됨 · 날짜 탭 → 해당일")
    weeks = ["월", "화", "수", "목", "금", "토", "일"]
    head = st.columns(7, gap="small")
    for i, w in enumerate(weeks):
        color = "#DC2626" if i == 6 else ("#2563EB" if i == 5 else "#64748B")
        head[i].markdown(
            f"<div style='text-align:center;font-size:10px;font-weight:600;color:{color};"
            f"line-height:1;padding:0 0 1px 0;'>{w}</div>",
            unsafe_allow_html=True,
        )

    cal = calendar.Calendar(firstweekday=0)
    clicked = None
    for week in cal.monthdayscalendar(month_anchor.year, month_anchor.month):
        cols = st.columns(7, gap="small")
        for i, day in enumerate(week):
            with cols[i]:
                if day == 0:
                    st.markdown("<div style='height:1.55rem'></div>", unsafe_allow_html=True)
                    continue
                d = date(month_anchor.year, month_anchor.month, day)
                is_sel = d == selected
                has = d.isoformat() in saved
                label = f"{day}•" if has else f"{day}"
                btn_type = "primary" if is_sel else "secondary"
                if st.button(
                    label,
                    key=f"wl_day_{d.isoformat()}",
                    width="stretch",
                    type=btn_type,
                ):
                    clicked = d
    return clicked


def render_worklog_tab(latest_update_str: str = "") -> None:
    """일일업무일지 탭 본체. (latest_update_str은 호환용 — 표시하지 않음)"""

    if load_workbook is None:
        st.error("openpyxl 이 필요합니다. `pip install openpyxl` 후 다시 실행하세요.")
        return
    try:
        _ensure_dirs()
        if not os.path.exists(WORKLOG_TEMPLATE):
            if _wl_quiet_ui():
                st.error(
                    "업무일지 템플릿이 없습니다. "
                    "사이드바에서 자료를 올린 뒤 "
                    "`uploaded_cache/worklog/template.xlsx` 를 준비하세요."
                )
            else:
                st.error(
                    "업무일지 템플릿을 찾을 수 없습니다. "
                    "`Desktop/업무일지.xlsx` 또는 "
                    "`uploaded_cache/worklog/template.xlsx` 를 준비하세요."
                )
            return
    except Exception as e:
        if _wl_quiet_ui():
            st.error("업무일지 템플릿을 준비하지 못했습니다. 파일을 다시 확인해 주세요.")
        else:
            st.error(f"템플릿 준비 실패: {e}")
        return

    @st.fragment
    def _worklog_main() -> None:
        if "worklog_selected" not in st.session_state:
            st.session_state["worklog_selected"] = date.today()
        selected: date = st.session_state["worklog_selected"]

        # Drive 동기화는 잦은 버튼/입력 rerun마다 돌리면 로딩이 심해짐 → 90초마다
        try:
            from drive_autoload import sync_worklog_bidirectional

            _now = time.time()
            _prev = float(st.session_state.get("_wl_drive_sync_ts") or 0)
            _force = bool(st.session_state.pop("_wl_drive_sync_force", None))
            if _force or (_now - _prev >= 90):
                st.session_state["_wl_drive_sync_ts"] = _now
                _wl_sync = sync_worklog_bidirectional(WORKLOG_DIR)
            else:
                _wl_sync = {"ok": True, "skipped": True, "copied": []}
            if (
                isinstance(_wl_sync, dict)
                and _wl_sync.get("ok")
                and not _wl_sync.get("skipped")
                and _wl_sync.get("copied")
            ):
                _invalidate_saved_dates_cache()
                if not _wl_quiet_ui():
                    st.caption(
                        f"Drive 업무일지 동기화 · {len(_wl_sync.get('copied') or [])}개"
                    )
            elif isinstance(_wl_sync, dict) and _wl_sync.get("skipped") and _wl_quiet_ui():
                if not st.session_state.get("_wl_cloud_sync_hint"):
                    st.session_state["_wl_cloud_sync_hint"] = True
                    st.caption(
                        "Cloud 공용 저장소에 저장됩니다. "
                        "맥에서도 같은 streamlit.app 주소로 열면 일지가 맞습니다."
                    )
            elif (
                isinstance(_wl_sync, dict)
                and _wl_sync.get("skipped")
                and not _wl_quiet_ui()
            ):
                if not st.session_state.get("_wl_local_cloud_hint"):
                    st.session_state["_wl_local_cloud_hint"] = True
                    _u = (
                        os.environ.get("DASHBOARD_CLOUD_URL")
                        or "https://office-g8ryabkapprkpjmfwa5aypw.streamlit.app"
                    )
                    st.info(
                        "로컬 Streamlit입니다. 아이패드와 일지를 맞추려면 사이드바 "
                        f"**Cloud에서 열기**로 접속하세요.\n\n{_u}"
                    )
        except Exception:
            pass

        # 삭제 요청은 위젯 생성 전에 처리
        if st.session_state.pop("wl_do_delete_day", None) == selected.isoformat():
            delete_worklog_day(selected)

        saved = list_saved_worklog_dates()
        _init_widget_state(selected)
        draft = _cells_from_widgets(selected)

        # 인쇄 패널이 열려 있으면 본화면 대신 인쇄 UI만 (팝업 중첩·취소 반복 방지)
        if st.session_state.get("wl_print_panel"):
            _render_worklog_print_panel()
            return

        # 날짜 선택(저장 후에도 변경 가능) · 달력 · 작은 삭제
        if st.session_state.get("wl_date_sync") != selected.isoformat():
            st.session_state["wl_date_pick"] = selected
            st.session_state["wl_date_sync"] = selected.isoformat()

        st.markdown(
            """
            <style>
            div[class*="st-key-wl_del_day_open"] button,
            div[class*="st-key-wl_del_day_yes"] button {
              font-size: 0.72rem !important;
              padding: 0.12rem 0.4rem !important;
              min-height: 1.55rem !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        # 왼쪽 요약 보기 · 가운데 남은행 게이지 · 오른쪽 입력
        try:
            _gauge_usage = _content_row_usage(_read_editor_entries(selected))
        except Exception:
            _gauge_usage = _content_row_usage([])

        col_preview, col_gauge, col_input = st.columns([1, 0.14, 1], gap="small")

        with col_preview:
            st.markdown("##### 업무일지 보기")
            # 버튼은 큰 iframe 위에 둠 (클릭 가로채기 방지)
            p1, p2, p3 = st.columns(3)
            with p1:
                do_print = st.button(
                    "엑셀 미리보기",
                    width="stretch",
                    key="wl_print_btn",
                    help="왼쪽 칸에 원본 엑셀 양식을 적용합니다.",
                )
            with p2:
                # 맥 로컬: Excel ⌘P / Cloud·공통: 브라우저 프린터 화면
                do_saved_excel = st.button(
                    "엑셀 저장본",
                    width="stretch",
                    key="wl_dl_btn",
                    help="맥 로컬은 Excel 인쇄 화면, Cloud는 브라우저 프린터 화면을 엽니다.",
                )
            with p3:
                do_browser_print = st.button(
                    "프린터 화면",
                    width="stretch",
                    key="wl_browser_print_btn",
                    help="브라우저 인쇄(프린터) 창을 바로 엽니다. Cloud에서도 됩니다.",
                    type="primary",
                )

            def _resolve_print_xlsx() -> str:
                cells_dl = _cells_from_widgets(selected)
                path_saved = worklog_path(selected)
                if os.path.exists(path_saved):
                    write_cells_to_path(
                        path_saved, selected, cells_dl, blank_base=False
                    )
                    return os.path.abspath(path_saved)
                return _prepare_excel_preview(selected, cells_dl)

            if do_browser_print or (
                do_saved_excel and platform.system() != "Darwin"
            ):
                try:
                    with st.spinner("인쇄 창 준비 중…"):
                        xlsx_abs = _resolve_print_xlsx()
                        _launch_browser_print_dialog(xlsx_abs)
                except Exception as e:
                    st.error(f"프린터 화면을 열지 못했습니다: {e}")

            if do_saved_excel and platform.system() == "Darwin":
                try:
                    xlsx_abs = _resolve_print_xlsx()
                    ok, msg = open_excel_print_preview(
                        xlsx_abs, prefer_print_dialog=True
                    )
                    if ok:
                        st.success(msg)
                    else:
                        st.warning(msg)
                        _launch_browser_print_dialog(xlsx_abs)
                except Exception as e:
                    st.error(f"엑셀 저장본 연결 실패: {e}")
                    try:
                        _launch_browser_print_dialog(_resolve_print_xlsx())
                    except Exception:
                        pass

            _left_excel_key = f"wl_left_excel_on_{selected.isoformat()}"
            _left_path_key = f"wl_left_excel_path_{selected.isoformat()}"
            if do_print:
                cells_now = _cells_from_widgets(selected)
                try:
                    xlsx_abs = _prepare_excel_preview(selected, cells_now)
                    st.session_state[_left_excel_key] = True
                    st.session_state[_left_path_key] = xlsx_abs
                    st.session_state["wl_dialog_preview_path"] = xlsx_abs
                    # 하단 원본 양식 캐시도 즉시 갱신
                    form_sig = json.dumps(cells_now, ensure_ascii=False, sort_keys=True)
                    st.session_state[f"wl_form_sig_v14_{selected.isoformat()}"] = None
                    st.session_state["_wl_force_form_sig"] = form_sig
                    # 맥 로컬: Excel 실행 + 인쇄 미리보기까지. Cloud는 화면 양식만.
                    if platform.system() == "Darwin":
                        ok, msg = open_excel_print_preview(
                            xlsx_abs, prefer_print_dialog=True
                        )
                        if ok:
                            st.success(msg)
                        else:
                            st.warning(msg)
                    elif _wl_quiet_ui():
                        st.caption(
                            "Cloud에서는 왼쪽 미리보기와 「프린터 화면」으로 확인하세요. "
                            "「엑셀 저장본」다운로드 후 맥 Excel에서 ⌘P 로 인쇄할 수도 있습니다."
                        )
                except Exception as e:
                    st.session_state[_left_excel_key] = False
                    if _wl_quiet_ui():
                        st.error(
                            "엑셀 미리보기를 적용하지 못했습니다. "
                            "「엑셀 저장본」다운로드를 이용해 주세요."
                        )
                    else:
                        st.error(f"엑셀 미리보기 오류: {e}")

            _show_excel_left = bool(st.session_state.get(_left_excel_key))
            if _show_excel_left:
                sw1, sw2 = st.columns([1, 1])
                with sw1:
                    st.caption("원본 엑셀 양식 적용 중")
                with sw2:
                    if st.button(
                        "요약 보기로",
                        width="stretch",
                        key=f"wl_left_to_summary_{selected.isoformat()}",
                    ):
                        st.session_state[_left_excel_key] = False
                        _wl_rerun()
                xlsx_left = st.session_state.get(_left_path_key) or ""
                if xlsx_left and os.path.exists(str(xlsx_left)):
                    try:
                        # 입력 바뀌면 미리보기 파일 다시 생성
                        cells_live = _cells_from_widgets(selected)
                        live_sig = json.dumps(
                            cells_live, ensure_ascii=False, sort_keys=True
                        )
                        sig_k = f"wl_left_excel_sig_{selected.isoformat()}"
                        if st.session_state.get(sig_k) != live_sig:
                            xlsx_left = _prepare_excel_preview(selected, cells_live)
                            st.session_state[_left_path_key] = xlsx_left
                            st.session_state[sig_k] = live_sig
                        scale_l = 0.48
                        excel_html = render_worklog_view_html(
                            str(xlsx_left),
                            print_mode=False,
                            auto_print=False,
                            scale=scale_l,
                        )
                        _, fh = _scaled_view_frame_size(str(xlsx_left), scale_l)
                        components.html(
                            excel_html,
                            height=min(820, max(480, int(fh))),
                            scrolling=True,
                        )
                        if st.button(
                            "크게 보기",
                            width="stretch",
                            key=f"wl_left_excel_big_{selected.isoformat()}",
                        ):
                            st.session_state["wl_dialog_preview_path"] = str(xlsx_left)
                            _worklog_form_preview_dialog()
                    except Exception as e:
                        st.warning(f"엑셀 양식 표시 실패: {e}")
                        st.session_state[_left_excel_key] = False
                else:
                    st.info("엑셀 미리보기 파일이 없습니다. 다시 「엑셀 미리보기」를 눌러 주세요.")
            else:
                try:
                    view_html = render_readable_preview_html(selected, draft)
                    components.html(view_html, height=820, scrolling=True)
                except Exception as e:
                    if _wl_quiet_ui():
                        st.info(
                            "업무일지 요약을 표시하지 못했습니다. 입력 후 다시 확인해 주세요."
                        )
                    else:
                        st.error(f"요약 보기 오류: {e}")

        with col_gauge:
            # 오른쪽 「저장」버튼 근처까지 세로로 맞춤 (폭 26px 유지)
            st.markdown(
                "<div style='height:0.35rem'></div>",
                unsafe_allow_html=True,
            )
            _render_row_remain_gauge(_gauge_usage, height_px=1020)

        with col_input:
            st.markdown("##### 업무 입력")

            bar_date, bar_cal, bar_del = st.columns([2.4, 1.1, 0.7], gap="small")
            with bar_date:
                picked = st.date_input(
                    "업무일지 날짜",
                    key="wl_date_pick",
                    help="저장 후에도 날짜를 바꿀 수 있습니다. 빈 날짜로 바꾸면 이 일지가 그 날짜로 옮겨집니다.",
                )
                _render_worklog_special_chars(selected.isoformat())
                msg = st.session_state.pop("wl_special_msg", None)
                if msg:
                    st.caption(msg)
            with bar_cal:
                st.markdown(
                    "<div style='height:1.55rem'></div>", unsafe_allow_html=True
                )
                with st.popover("📅 달력", width="content"):
                    clicked = _render_month_calendar(selected, saved)
                    if clicked is not None and clicked != selected:
                        _clear_date_widget_state(selected)
                        st.session_state["worklog_selected"] = clicked
                        st.session_state["worklog_month"] = date(
                            clicked.year, clicked.month, 1
                        )
                        st.session_state["wl_date_sync"] = ""
                        _wl_rerun()
            with bar_del:
                st.markdown(
                    "<div style='height:1.55rem'></div>", unsafe_allow_html=True
                )
                with st.popover("삭제", width="content", key="wl_del_day_open"):
                    st.caption("이 날짜 일지 전체 삭제")
                    if st.button(
                        "확정", type="primary", width="content", key="wl_del_day_yes"
                    ):
                        st.session_state["wl_do_delete_day"] = selected.isoformat()
                        _wl_rerun()

            if isinstance(picked, date) and picked != selected:
                if os.path.exists(worklog_path(picked)):
                    _clear_date_widget_state(selected)
                    st.session_state["worklog_selected"] = picked
                    st.session_state["worklog_month"] = date(
                        picked.year, picked.month, 1
                    )
                    st.session_state["wl_date_sync"] = ""
                    _wl_rerun()
                else:
                    try:
                        reassign_worklog_date(selected, picked)
                        st.session_state["wl_date_sync"] = ""
                        _wl_rerun()
                    except FileExistsError as e:
                        st.error(str(e))
                        st.session_state["wl_date_sync"] = ""
                        _wl_rerun()

            iso = selected.isoformat()
            ek = _entries_key(selected)
            if ek not in st.session_state or not st.session_state[ek]:
                st.session_state[ek] = [{"client": "", "content": ""}]

            def _seed_entry_widgets(entries_list: list[dict], next_txt: str, notes_txt: str) -> None:
                """위젯 생성 전에만 호출 — 키를 지우고 다시 심는다."""
                old_n = int(st.session_state.get(f"wl_entry_count_{iso}", 0) or 0)
                for i in range(max(old_n, len(entries_list)) + 2):
                    st.session_state.pop(f"wl_ent_c_{iso}_{i}", None)
                    st.session_state.pop(f"wl_ent_t_{iso}_{i}", None)
                    st.session_state.pop(f"wl_ent_gap_{iso}_{i}", None)
                    st.session_state.pop(f"wl_exp_{iso}_{i}", None)
                    st.session_state.pop(_entry_lines_comp_key(iso, i), None)
                    st.session_state.pop(_entry_lines_rev_key(iso, i), None)
                    st.session_state.pop(_entry_lines_live_key(iso, i), None)
                    old_lc = int(st.session_state.get(_entry_line_count_key(iso, i), 0) or 0)
                    for j in range(old_lc + 3):
                        st.session_state.pop(_entry_line_key(iso, i, j), None)
                    st.session_state.pop(_entry_line_count_key(iso, i), None)
                    old_cc = int(
                        st.session_state.get(_entry_client_count_key(iso, i), 0) or 0
                    )
                    for j in range(old_cc + 3):
                        st.session_state.pop(_entry_client_key(iso, i, j), None)
                    st.session_state.pop(_entry_client_count_key(iso, i), None)
                st.session_state.pop(f"wl_next_area_{iso}", None)
                st.session_state.pop(f"wl_notes_area_{iso}", None)
                st.session_state[ek] = entries_list
                st.session_state[f"wl_entry_count_{iso}"] = len(entries_list)
                for i, ent in enumerate(entries_list):
                    clines = ent.get("client_lines")
                    if isinstance(clines, list) and clines:
                        _seed_entry_clients(iso, i, clines)
                    else:
                        _seed_entry_clients(iso, i, ent.get("client") or "")
                    st.session_state[f"wl_ent_gap_{iso}_{i}"] = _entry_blank_after(ent, 1)
                    lines = ent.get("lines")
                    if isinstance(lines, list):
                        _apply_entry_lines(iso, i, [str(x or "") for x in lines])
                    else:
                        _seed_entry_lines(iso, i, ent.get("content") or "")
                st.session_state[f"wl_next_area_{iso}"] = next_txt
                st.session_state[f"wl_notes_area_{iso}"] = notes_txt
                st.session_state[_next_key(selected)] = next_txt
                st.session_state[_notes_key(selected)] = notes_txt

            # 저장 후 pending → 위젯 생성 전에 적용
            pending = st.session_state.pop(f"wl_pending_sync_{iso}", None)
            if isinstance(pending, dict):
                ents = pending.get("entries") or [{"client": "", "content": ""}]
                _seed_entry_widgets(
                    ents,
                    pending.get("next") or "",
                    pending.get("notes") or "",
                )
                if pending.get("msg"):
                    st.success(pending["msg"])

            # 추가/삭제 전: 위젯에 있던 최신 값 반영
            add_clicked = st.session_state.pop(f"wl_do_add_{iso}", False)
            del_idx = st.session_state.pop(f"wl_do_del_{iso}", None)
            if add_clicked or isinstance(del_idx, int):
                entries = _read_editor_entries(selected)
                if isinstance(del_idx, int) and 0 <= del_idx < len(entries):
                    entries.pop(del_idx)
                if add_clicked:
                    entries.append(
                        {"client": "", "content": "", "lines": [""], "blank_after": 1}
                    )
                if not entries:
                    entries = [{"client": "", "content": "", "blank_after": 1}]
                next_txt = str(
                    st.session_state.get(f"wl_next_area_{iso}")
                    or st.session_state.get(_next_key(selected), "")
                    or ""
                )
                notes_txt = str(
                    st.session_state.get(f"wl_notes_area_{iso}")
                    or st.session_state.get(_notes_key(selected), "")
                    or ""
                )
                _seed_entry_widgets(entries, next_txt, notes_txt)
                if add_clicked:
                    # 새로 추가한 항목만 펼침
                    for _i in range(len(entries)):
                        st.session_state[f"wl_exp_{iso}_{_i}"] = _i == len(entries) - 1
                    st.session_state[f"wl_force_expand_{iso}"] = len(entries) - 1
            else:
                entries = list(st.session_state[ek])
                for i, ent in enumerate(entries):
                    ck = f"wl_ent_c_{iso}_{i}"
                    gk = f"wl_ent_gap_{iso}_{i}"
                    if ck not in st.session_state:
                        st.session_state[ck] = ent.get("client") or ""
                    if gk not in st.session_state:
                        st.session_state[gk] = _entry_blank_after(ent, 1)
                    if int(
                        st.session_state.get(_entry_client_count_key(iso, i), 0) or 0
                    ) <= 0:
                        cl0 = ent.get("client_lines")
                        if isinstance(cl0, list) and cl0:
                            _seed_entry_clients(iso, i, cl0)
                        else:
                            _seed_entry_clients(iso, i, ent.get("client") or "")
                    if int(st.session_state.get(_entry_line_count_key(iso, i), 0) or 0) <= 0:
                        lines0 = ent.get("lines")
                        if isinstance(lines0, list):
                            _apply_entry_lines(iso, i, [str(x or "") for x in lines0])
                        else:
                            _seed_entry_lines(iso, i, ent.get("content") or "")
                st.session_state[f"wl_entry_count_{iso}"] = len(entries)
                nk = f"wl_next_area_{iso}"
                ok = f"wl_notes_area_{iso}"
                if nk not in st.session_state:
                    st.session_state[nk] = st.session_state.get(_next_key(selected), "")
                if ok not in st.session_state:
                    st.session_state[ok] = st.session_state.get(_notes_key(selected), "")

            def _wl_entry_editor():
                d = st.session_state.get("worklog_selected") or selected
                iso2 = d.isoformat()
                n = int(st.session_state.get(f"wl_entry_count_{iso2}", 1) or 1)
                max_u = _content_line_units()

                _force_open = st.session_state.pop(f"wl_force_expand_{iso2}", None)
                if isinstance(_force_open, int):
                    st.session_state[f"wl_exp_{iso2}_{_force_open}"] = True

                # 칸 삭제/추가 (위젯 생성 전)
                del_ln = st.session_state.pop(f"wl_do_del_ln_{iso2}", None)
                if isinstance(del_ln, (list, tuple)) and len(del_ln) == 2:
                    dei, dlj = int(del_ln[0]), int(del_ln[1])
                    cur = _lines_from_entry_widgets(iso2, dei, keep_trailing_empty=False)
                    if 0 <= dlj < len(cur):
                        cur.pop(dlj)
                    if not cur:
                        cur = [""]
                    _apply_entry_lines(
                        iso2, dei, cur, focus_j=min(dlj, max(len(cur) - 1, 0))
                    )

                ins_ln = st.session_state.pop(f"wl_do_insert_ln_{iso2}", None)
                if isinstance(ins_ln, (list, tuple)) and len(ins_ln) == 2:
                    _insert_line_after(iso2, int(ins_ln[0]), int(ins_ln[1]))

                del_cl = st.session_state.pop(f"wl_do_del_cl_{iso2}", None)
                if isinstance(del_cl, (list, tuple)) and len(del_cl) == 2:
                    dei, dlj = int(del_cl[0]), int(del_cl[1])
                    cur = _clients_from_widgets(iso2, dei, keep_trailing_empty=False)
                    if 0 <= dlj < len(cur):
                        cur.pop(dlj)
                    if not cur:
                        cur = [""]
                    _apply_entry_clients(
                        iso2, dei, cur, focus_j=min(dlj, max(len(cur) - 1, 0))
                    )

                ins_cl = st.session_state.pop(f"wl_do_insert_cl_{iso2}", None)
                if isinstance(ins_cl, (list, tuple)) and len(ins_cl) == 2:
                    _insert_client_after(iso2, int(ins_cl[0]), int(ins_cl[1]))

                # Enter 커밋(입력값 반영 + 칸 분할) — 반드시 위젯 생성 전
                ent_req = st.session_state.pop(f"wl_do_enter_cell_{iso2}", None)
                if isinstance(ent_req, dict):
                    _commit_enter_on_cell(
                        str(ent_req.get("kind") or ""),
                        iso2,
                        int(ent_req.get("ei") or 0),
                        int(ent_req.get("lj") or 0),
                        str(ent_req.get("v") or ""),
                    )

                sp_req = st.session_state.pop(f"wl_do_special_{iso2}", None)
                if isinstance(sp_req, dict):
                    try:
                        _sp_pos = int(sp_req.get("s") or 0)
                    except (TypeError, ValueError):
                        _sp_pos = 0
                    _apply_special_insert(
                        iso2,
                        str(sp_req.get("key") or ""),
                        str(sp_req.get("v") if sp_req.get("v") is not None else ""),
                        _sp_pos,
                        str(sp_req.get("ch") or ""),
                    )

                def _on_enter_trigger():
                    hook = st.session_state.get(f"wl_enter_hook_{iso2}") or {}
                    payload = ""
                    if isinstance(hook, dict):
                        payload = str(hook.get("enter") or "")
                    if not payload:
                        return
                    done_k = f"wl_enter_done_{iso2}"
                    key = ""
                    val = ""
                    try:
                        obj = json.loads(payload)
                        key = str(obj.get("key") or "")
                        val = str(obj.get("v") or "")
                    except Exception:
                        key = payload.split(":", 1)[0]
                        val = ""
                    # timestamp 달라도 같은 칸·같은 값이면 1회만 처리
                    sig = f"{key}\0{val}"
                    if st.session_state.get(done_k) == sig:
                        return
                    st.session_state[done_k] = sig
                    m = re.match(
                        r"^(wl_ent_ln|wl_ent_cl)_(\d{4}-\d{2}-\d{2})_(\d+)_(\d+)(?:_g\d+)?$",
                        key,
                    )
                    if not m or m.group(2) != iso2:
                        return
                    st.session_state[f"wl_do_enter_cell_{iso2}"] = {
                        "kind": m.group(1),
                        "ei": int(m.group(3)),
                        "lj": int(m.group(4)),
                        "v": val,
                    }

                def _on_focus_trigger():
                    hook = st.session_state.get(f"wl_enter_hook_{iso2}") or {}
                    if not isinstance(hook, dict):
                        return
                    fk = str(hook.get("focus") or "")
                    if fk.startswith("wl_ent_ln_") or fk.startswith("wl_ent_cl_"):
                        st.session_state["wl_active_cell_key"] = fk

                def _on_caret_trigger():
                    hook = st.session_state.get(f"wl_enter_hook_{iso2}") or {}
                    if not isinstance(hook, dict):
                        return
                    raw = hook.get("caret")
                    if not raw:
                        return
                    try:
                        obj = json.loads(str(raw))
                        fk = str(obj.get("key") or "")
                        s = int(obj.get("s") or 0)
                        e = int(obj.get("e") or s)
                    except Exception:
                        return
                    if fk.startswith("wl_ent_ln_") or fk.startswith("wl_ent_cl_"):
                        st.session_state["wl_active_cell_key"] = fk
                        st.session_state["wl_active_cell_sel"] = (s, e)

                focus_key = st.session_state.pop(f"wl_focus_ln_{iso2}", None)
                focus_caret = st.session_state.pop(f"wl_focus_caret_{iso2}", None)
                if isinstance(focus_key, str) and (
                    focus_key.startswith("wl_ent_ln_")
                    or focus_key.startswith("wl_ent_cl_")
                ):
                    try:
                        _m = re.match(
                            r"^wl_ent_(?:ln|cl)_\d{4}-\d{2}-\d{2}_(\d+)_",
                            focus_key,
                        )
                        if _m:
                            st.session_state[f"wl_exp_{iso2}_{int(_m.group(1))}"] = True
                    except Exception:
                        pass

                # 전체 내용칸 잔여 — form 제출 전 값은 직전 커밋 기준
                st.caption(
                    "칸에 입력하는 동안 화면이 다시 로딩되지 않습니다. "
                    "「저장」을 눌러야 미리보기·파일에 반영됩니다."
                )
                with st.form(
                    f"wl_entry_form_{iso2}",
                    clear_on_submit=False,
                    enter_to_submit=False,
                    border=False,
                ):
                    _live_entries = _read_editor_entries(d)
                    _usage = _content_row_usage(_live_entries)
                    _rem = _usage["remaining"]

                    for i in range(n):
                        if int(
                            st.session_state.get(_entry_client_count_key(iso2, i), 0) or 0
                        ) > 0:
                            _cl0 = _clients_from_widgets(
                                iso2, i, keep_trailing_empty=False
                            )
                            client_now = (_cl0[0] if _cl0 else "").strip()
                        else:
                            client_now = str(
                                st.session_state.get(f"wl_ent_c_{iso2}_{i}", "") or ""
                            ).strip()
                            if client_now:
                                client_now = client_now.splitlines()[0].strip()
                        body_now = _content_from_entry_lines(iso2, i).strip().replace(
                            "\n", " "
                        )
                        if len(body_now) > 24:
                            body_now = body_now[:24] + "…"
                        label = f"항목 {i + 1}"
                        if client_now:
                            label += f" · {client_now}"
                        elif body_now:
                            label += f" · {body_now}"
                        else:
                            label += " · (비어 있음)"

                        exp_key = f"wl_exp_{iso2}_{i}"
                        # 기본: 마지막 항목만 펼침 (키로 이후 사용자 토글 유지)
                        default_open = i == n - 1
                        if exp_key not in st.session_state:
                            st.session_state[exp_key] = default_open

                        with st.expander(
                            label,
                            expanded=bool(st.session_state.get(exp_key)),
                            key=exp_key,
                        ):
                            if st.form_submit_button(
                                "이 항목 삭제",
                                key=f"wl_del_btn_{iso2}_{i}",
                                width="stretch",
                            ):
                                st.session_state[f"wl_do_del_{iso2}"] = i
                                _wl_rerun()
                            _cu = _client_line_units()
                            if int(
                                st.session_state.get(
                                    _entry_client_count_key(iso2, i), 0
                                )
                                or 0
                            ) <= 0:
                                stored_e = st.session_state.get(_entries_key(d)) or []
                                if i < len(stored_e) and isinstance(
                                    stored_e[i].get("client_lines"), list
                                ):
                                    _seed_entry_clients(
                                        iso2,
                                        i,
                                        stored_e[i].get("client_lines") or [""],
                                    )
                                else:
                                    _seed_entry_clients(
                                        iso2,
                                        i,
                                        str(
                                            st.session_state.get(
                                                f"wl_ent_c_{iso2}_{i}", ""
                                            )
                                            or ""
                                        ),
                                    )
                            if (
                                int(
                                    st.session_state.get(
                                        _entry_line_count_key(iso2, i), 0
                                    )
                                    or 0
                                )
                                <= 0
                            ):
                                lines0 = None
                                stored_e = st.session_state.get(_entries_key(d)) or []
                                if i < len(stored_e) and isinstance(
                                    stored_e[i].get("lines"), list
                                ):
                                    lines0 = stored_e[i].get("lines")
                                if isinstance(lines0, list):
                                    _apply_entry_lines(
                                        iso2, i, [str(x or "") for x in lines0]
                                    )
                                else:
                                    _seed_entry_lines(
                                        iso2,
                                        i,
                                        str(
                                            st.session_state.get(
                                                f"wl_ent_t_{iso2}_{i}", ""
                                            )
                                            or ""
                                        ),
                                    )
                            col_client, col_content = st.columns(
                                [1, 3.2], gap="small"
                            )
                            with col_client:
                                _mount_entry_client_editor(iso2, i, _cu)
                            with col_content:
                                _mount_entry_lines_editor(iso2, i, max_u)
                            _filled = len(
                                _lines_from_entry_widgets(
                                    iso2, i, keep_trailing_empty=False
                                )
                            )
                            gap_key = f"wl_ent_gap_{iso2}_{i}"
                            if gap_key not in st.session_state:
                                st.session_state[gap_key] = _entry_blank_after(
                                    (
                                        _live_entries[i]
                                        if i < len(_live_entries)
                                        else None
                                    ),
                                    1,
                                )
                            st.number_input(
                                "다음 항목 전 빈 칸 수",
                                min_value=0,
                                max_value=10,
                                step=1,
                                key=gap_key,
                                help="이 항목 다음에 원본 엑셀에서 비워 둘 행 수. "
                                "다른 거래처 항목과 구분하려면 1 이상 권장 "
                                "(0이면 거래처 줄바꿈과 구분이 어려울 수 있음)",
                            )
                            _ent_rows = (
                                _usage["per_entry"][i]
                                if i < len(_usage["per_entry"])
                                else max(_filled, 1)
                            )
                            st.caption(
                                f"이 항목 약 {_ent_rows}행 사용 · "
                                f"전체 남은 {_rem}행 (마지막 칸 G{_usage['last_row']})"
                            )

                    if st.form_submit_button(
                        "＋ 항목 추가", key=f"wl_add_btn_{iso2}", width="stretch"
                    ):
                        st.session_state[f"wl_do_add_{iso2}"] = True
                        _wl_rerun()

                    st.markdown(
                        "<div style='font-size:12px;font-weight:700;color:#334155;margin:12px 0 4px;'>"
                        "익일업무 <span style='font-weight:500;color:#94A3B8;'>(줄바꿈 = 항목 구분)</span></div>",
                        unsafe_allow_html=True,
                    )
                    st.text_area(
                        "익일업무",
                        key=f"wl_next_area_{iso2}",
                        label_visibility="collapsed",
                        height=90,
                    )
                    st.markdown(
                        "<div style='font-size:12px;font-weight:700;color:#334155;margin:12px 0 4px;'>"
                        "특 이 사 항 <span style='font-weight:500;color:#94A3B8;'>(줄바꿈 = 항목 구분)</span></div>",
                        unsafe_allow_html=True,
                    )
                    st.text_area(
                        "특이사항",
                        key=f"wl_notes_area_{iso2}",
                        label_visibility="collapsed",
                        height=90,
                    )

                    if st.form_submit_button(
                        "저장",
                        type="primary",
                        width="stretch",
                        key=f"wl_save_btn_{iso2}",
                    ):
                        try:
                            entries_now = _read_editor_entries(d)
                            usage_now = _content_row_usage(entries_now)
                            if usage_now.get("overflow"):
                                st.error(
                                    f"내용칸 용량 초과: {usage_now['used']}/{usage_now['total']}행. "
                                    "칸을 줄이거나 항목 사이 빈 칸 수를 낮춘 뒤 다시 저장하세요."
                                )
                            else:
                                cells = _pack_entries_to_cells(
                                    d,
                                    entries_now,
                                    [
                                        x.strip()
                                        for x in str(
                                            st.session_state.get(
                                                f"wl_next_area_{iso2}", ""
                                            )
                                            or ""
                                        ).splitlines()
                                        if x.strip()
                                    ],
                                    [
                                        x.strip()
                                        for x in str(
                                            st.session_state.get(
                                                f"wl_notes_area_{iso2}", ""
                                            )
                                            or ""
                                        ).splitlines()
                                        if x.strip()
                                    ],
                                )
                                path = save_worklog_cells(d, cells)
                                packed_entries = _grouped_entries_from_cells(cells)
                                if not packed_entries:
                                    packed_entries = [
                                        {"client": "", "content": "", "lines": []}
                                    ]
                                _, nd, nt = _entries_from_cells(cells)
                                st.session_state[_entries_key(d)] = packed_entries
                                st.session_state[_next_key(d)] = "\n".join(nd)
                                st.session_state[_notes_key(d)] = "\n".join(nt)
                                arch = (
                                    st.session_state.get("wl_last_archive_path") or ""
                                )
                                drv = st.session_state.get("wl_last_drive_path") or ""
                                msg = f"저장 완료: {os.path.basename(path)}"
                                if arch and not _wl_quiet_ui():
                                    msg += (
                                        f" · 일지/{d.year}/{os.path.basename(arch)}"
                                    )
                                if drv:
                                    msg += " · Drive 복사본/worklog"
                                st.session_state[f"wl_pending_sync_{iso2}"] = {
                                    "entries": packed_entries,
                                    "next": "\n".join(nd),
                                    "notes": "\n".join(nt),
                                    "msg": msg,
                                }
                                _wl_rerun()
                        except Exception as e:
                            if _wl_quiet_ui():
                                st.error(
                                    "저장에 실패했습니다. 입력 내용을 확인한 뒤 다시 시도해 주세요."
                                )
                            else:
                                st.error(f"저장 실패: {e}")

                # Enter 훅: 방향키·칸 이동만 (입력 중 서버 호출 없음)
                _WL_ENTER_HOOK(
                    key=f"wl_enter_hook_{iso2}",
                    data={
                        "iso": iso2,
                        "focus_key": focus_key if isinstance(focus_key, str) else "",
                        "focus_caret": (
                            int(focus_caret)
                            if isinstance(focus_caret, (int, float))
                            else ""
                        ),
                        "client_max_u": _client_line_units(),
                        "content_max_u": _content_line_units(),
                    },
                    width="stretch",
                    height=1,
                )
                ins_after = st.session_state.pop(f"wl_do_insert_ln_{iso2}", None)
                if isinstance(ins_after, (list, tuple)) and len(ins_after) == 2:
                    _insert_line_after(iso2, int(ins_after[0]), int(ins_after[1]))
                    _wl_rerun()
                ins_cl_after = st.session_state.pop(f"wl_do_insert_cl_{iso2}", None)
                if isinstance(ins_cl_after, (list, tuple)) and len(ins_cl_after) == 2:
                    _insert_client_after(
                        iso2, int(ins_cl_after[0]), int(ins_cl_after[1])
                    )
                    _wl_rerun()

            _wl_entry_editor()
    _worklog_main()
