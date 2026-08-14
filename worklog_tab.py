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

  const maxU = Number((data && data.max_u) || 69);
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

    공식 mdw=7 은 내용칸이 너무 좁아 14pt 한글이 잘림.
    원본 스크린샷 채움(~92%)에 맞춘 계수.
    """
    try:
        w = float(width)
    except (TypeError, ValueError):
        w = 8.43
    return max(10, int(w * 7 + 5))


@lru_cache(maxsize=1)
def _content_line_units() -> int:
    """원본 G:X 병합 폭 기준 — 엑셀 바탕체 14pt 한 줄에 가깝게 채운 뒤 다음 행.

    반각=1, 한글 음절/호환자모=2. 원본 스크린샷 기준 긴 줄이 칸의 ~92%를 채움.
    """
    fallback = 70
    if load_workbook is None or not os.path.exists(WORKLOG_TEMPLATE):
        return fallback
    try:
        # read_only 시트는 column_dimensions 가 없어 폭 계산 불가
        wb = load_workbook(WORKLOG_TEMPLATE, data_only=False)
        ws = wb.active
        total = sum(
            _excel_col_width(ws, c)
            for c in range(WL_CONTENT_COL_START, WL_CONTENT_COL_END + 1)
        )
        wb.close()
        # 열폭(기본글꼴 문자수) → 바탕 14pt 실측 채움에 맞춤
        units = int(total * (11 / 14) * 1.05)
        return max(60, min(units, 72))
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
        ws.cell(r, 3).value = (cells.get(f"C{r}", "") or None)
    for r in WL_CONTENT_ROWS:
        cell = ws.cell(r, 7)
        cell.value = (cells.get(f"G{r}", "") or None)
        # 칸 안에서 줄바꿈·넘침으로 다음 행을 가리지 않도록
        try:
            cell.alignment = cell.alignment.copy(wrapText=False, shrinkToFit=False)
        except Exception:
            pass
    for r in WL_NEXT_ROWS + WL_NOTE_ROWS:
        cell = ws.cell(r, 4)
        cell.value = (cells.get(f"D{r}", "") or None)
        try:
            cell.alignment = cell.alignment.copy(wrapText=False, shrinkToFit=False)
        except Exception:
            pass
    wb.save(path)
    wb.close()


def save_worklog_cells(d: date, cells: dict) -> str:
    """달력용 캐시에 저장하고, Desktop/업무/일지/{년도}에도 복사(맥 로컬).

    iPad/Cloud에는 Google Drive「다른 컴퓨터」경로가 없으므로 달력 저장만 하고
    보조 복사는 조용히 건너뛴다.
    """
    path = worklog_path(d)
    is_new = not os.path.exists(path)
    cells = _spill_all_content(cells)
    write_cells_to_path(path, d, cells, blank_base=is_new)
    _invalidate_saved_dates_cache()
    st.session_state.pop("wl_last_archive_path", None)
    st.session_state.pop("wl_last_archive_err", None)
    try:
        archive = worklog_archive_path(d)
        if archive:
            shutil.copy2(path, archive)
            st.session_state["wl_last_archive_path"] = archive
    except Exception as e:
        # 달력 저장은 이미 성공 — 보조 경로 실패는 맥에서만 안내
        st.session_state["wl_last_archive_err"] = str(e)
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
            fname = font.name or "Batang"
            # 엑셀 한글 폰트명 정규화
            if fname in ("맑은 고딕", "Malgun Gothic"):
                fname_css = "Malgun Gothic"
            elif fname in ("바탕", "바탕체", "Batang", "BatangChe"):
                fname_css = "Batang"
            else:
                fname_css = fname
            # 엑셀 pt → CSS px (96dpi). pt 단위가 iframe에서 작게 먹히는 경우 대비.
            fsize_pt = float(font.size or 12)
            fsize_px = fsize_pt * 96.0 / 72.0
            bold = "bold" if font.bold else "normal"
            align = cell.alignment
            ha = align.horizontal or "left"
            va = align.vertical or "middle"
            if ha == "general":
                ha = "left"
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
            esc = html.escape(text).replace("\n", "<br>")
            # 내용칸은 한 행=한 줄 (원본 엑셀과 동일하게 칸 밖으로 줄바꿈하지 않음)
            is_content = c == 7 and 9 <= r <= 39
            is_client = c == 3 and 9 <= r <= 39
            white = "nowrap" if (is_content or is_client) else "pre-wrap"
            overflow = "hidden" if (is_content or is_client) else "visible"
            # 병합 셀 폭 = 포함 열 합 (고정 레이아웃에서 원본과 같은 채움감)
            c0 = c - WL_MIN_COL
            span_w = sum(col_widths[c0 : c0 + max(cs, 1)]) if c0 >= 0 else 0
            width_css = f"width:{span_w}px;min-width:{span_w}px;max-width:{span_w}px;" if span_w else ""
            # macOS에 Batang 없을 때 명조 계열로 폭이 비슷하게
            style = (
                f"box-sizing:border-box;{width_css}"
                f"font-family:'{html.escape(fname_css)}','Batang','BatangChe',"
                f"'Apple Myungjo','AppleMyungjo','Nanum Myeongjo','Malgun Gothic',serif;"
                f"font-size:{fsize_px:.4f}px;font-weight:{bold};"
                f"text-align:{ha};vertical-align:{va};"
                f"background:{fill};{border}"
                f"padding:0 2px;white-space:{white};overflow:{overflow};"
                f"text-overflow:clip;word-break:keep-all;"
                f"height:{height_px}px;min-height:{height_px}px;max-height:{height_px}px;"
                f"line-height:{height_px}px;"
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
    toolbar = ""
    if print_mode:
        toolbar = """
        <div class="toolbar">
          <button onclick="window.print()">다시 인쇄 / PDF 저장</button>
          <span class="hint">인쇄창이 안 뜨면 이 버튼을 누르세요.</span>
        </div>
        """
        scale = 1.0
    frame_w, frame_h = _scaled_view_frame_size(path, scale)
    # zoom은 레이아웃까지 축소. wrap은 auto로 두어 하단이 clip 되지 않게 함.
    if print_mode or scale >= 1:
        scale_css = ""
        scale_css_fallback = ""
        wrap_h = "auto"
        wrap_w = "auto"
        wrap_overflow = "visible"
        body_overflow = "visible"
        body_h = "auto"
    else:
        s = float(scale)
        raw_w, raw_h = _worklog_sheet_pixel_size(path)
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
    auto_script = ""
    if auto_print:
        auto_script = """
        <script>
          (function() {
            var tries = 0;
            function go() {
              tries += 1;
              try { window.focus(); window.print(); } catch (e) {}
              if (tries < 3) setTimeout(go, 400 * tries);
            }
            if (document.readyState === 'complete') setTimeout(go, 200);
            else window.addEventListener('load', function() { setTimeout(go, 200); });
          })();
        </script>
        """
    fallback_block = ""
    if not print_mode and scale < 1:
        fallback_block = f"""
  @supports not (zoom: 1) {{
    .sheet-scale {{ {scale_css_fallback} }}
  }}
"""
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>일일업무일지</title>
<style>
  @page {{ size: A4 portrait; margin: 10mm; }}
  html, body {{
    margin:0; padding:0; background:#fff;
    overflow:{body_overflow} !important;
    height:{body_h};
  }}
  body {{ padding:6px; box-sizing:border-box; }}
  .toolbar {{ margin-bottom:10px; display:flex; gap:10px; align-items:center; flex-wrap:wrap; }}
  .toolbar button {{
    padding:8px 14px; font-size:14px; border:1px solid #334155; border-radius:6px;
    background:#1E293B; color:#fff; cursor:pointer;
  }}
  .toolbar .hint {{ font:12px/1.4 sans-serif; color:#64748B; }}
  .wrap {{
    overflow:{wrap_overflow} !important; height:{wrap_h};
    width:{wrap_w}; max-width:100%;
    border:1px solid #94A3B8; background:#fff;
    box-sizing:border-box;
  }}
  .sheet-scale {{ {scale_css} }}
  .wl-sheet, .wl-sheet td, .wl-sheet tr {{ box-sizing:border-box; }}
  {fallback_block}
  @media print {{
    html, body {{ overflow:visible !important; height:auto; }}
    .toolbar {{ display:none !important; }}
    .wrap {{ overflow:visible !important; height:auto; width:auto; border:none; }}
    .sheet-scale {{
      zoom:1 !important; transform:none !important; margin-bottom:0 !important;
    }}
    body {{ padding:0; }}
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

    - 완전 빈 행(C·G 둘 다 비움) 뒤에 오면 새 항목 (앞 빈 행 수 = 이전 blank_after)
    - 내용칸만 공백 한 칸(SOFT_BLANK)이면 같은 항목의 빈 줄
    - 빈 행 없이 이어지면 같은 항목: 거래처·내용 칸을 각각 이어 붙임
    """
    entries: list[dict] = []
    blank_run = 0
    for r in WL_CLIENT_ROWS:
        raw_c = str(cells.get(f"C{r}", "") or "")
        raw_g = str(cells.get(f"G{r}", "") or "")
        client = _scrub_dummy_label(raw_c).strip()
        # 소프트 빈 줄은 strip 하지 않은 채 판별
        soft_blank = raw_g == _WL_SOFT_BLANK or raw_g == "\u00a0"
        content = "" if soft_blank else _scrub_dummy_label(raw_g)
        empty = not client and not soft_blank and not (content or "").strip()
        if empty:
            blank_run += 1
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
            # 같은 항목 연속 칸 (거래처 줄바꿈 / 내용 추가 / 항목 안 빈 줄)
            if blank_run > 0:
                for _ in range(blank_run):
                    entries[-1].setdefault("lines", []).append("")
                    entries[-1].setdefault("client_lines", []).append("")
                blank_run = 0
            ent = entries[-1]
            cl = ent.setdefault("client_lines", [])
            if not cl and str(ent.get("client") or "").strip():
                cl[:] = [str(ent.get("client") or "").strip()]
            # 내용만 이어지는 행·소프트 빈 줄은 거래처 칸을 늘리지 않음
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
    <div class="foot">보고용 원본 양식이 필요하면 아래 「원본양식 인쇄」를 사용하세요.</div>
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


def _entry_line_key(iso: str, entry_i: int, line_j: int) -> str:
    return f"wl_ent_ln_{iso}_{entry_i}_{line_j}"


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


def _apply_entry_lines(iso: str, entry_i: int, lines: list[str], *, focus_j: int | None = None) -> None:
    """줄 목록을 text_input 키에 반영. 끝에 편집용 빈 칸 1개 유지."""
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
        focus = line_j + 1
        if len(pieces) == 1:
            new = head + pieces + [""] + list(tail)
        else:
            new = head + pieces + list(tail)
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
    head, tail = cur[:line_j], cur[line_j + 1 :]
    if _display_units(value) > max_u:
        pieces = _chunk_text(value, max_u) or [value]
    else:
        pieces = [value]
    focus = line_j + 1
    if len(pieces) == 1:
        new = head + pieces + [""] + list(tail)
    else:
        new = head + pieces + list(tail)
        if focus >= len(new):
            new.append("")
    _apply_entry_lines(iso, entry_i, new, focus_j=min(focus, len(new) - 1))


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
        full = "".join(parts)
        _seed_entry_clients(iso, entry_i, full)
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
                if st.button(
                    "＋",
                    key=f"wl_cl_add_{iso}_{entry_i}_{j}",
                    width="stretch",
                    help="아래 거래처 칸 추가",
                ):
                    st.session_state[f"wl_do_insert_cl_{iso}"] = (entry_i, j)
                    _wl_rerun()
            elif st.button(
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
        full = "".join(parts)
        _seed_entry_lines(iso, entry_i, full)
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
                if st.button(
                    "＋",
                    key=f"wl_ln_add_{iso}_{entry_i}_{j}",
                    width="stretch",
                    help="아래 칸 추가",
                ):
                    st.session_state[f"wl_do_insert_ln_{iso}"] = (entry_i, j)
                    _wl_rerun()
            elif st.button(
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


# Enter / 칸초과 → 다음 칸 (document 키 훅)
_WL_ENTER_HOOK_JS = r"""
export default function (component) {
  const { data, setTriggerValue } = component;
  const iso = (data && data.iso) || "";
  const focusKey = (data && data.focus_key) || "";
  const clientMax = Number((data && data.client_max_u) || 15);
  const contentMax = Number((data && data.content_max_u) || 72);
  let lastSent = "";

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
    const m = /^(wl_ent_ln|wl_ent_cl)_(\d{4}-\d{2}-\d{2})_(\d+)_(\d+)$/.exec(key);
    if (!m || m[2] !== iso) return null;
    return { key: key, kind: m[1], maxU: m[1] === "wl_ent_cl" ? clientMax : contentMax };
  }
  function emit(key, value) {
    const payload = JSON.stringify({
      key: key,
      t: Date.now(),
      v: String(value || ""),
    });
    if (payload === lastSent) return;
    lastSent = payload;
    setTriggerValue("enter", payload);
  }

  const onKey = (e) => {
    if (e.key !== "Enter") return;
    if (e.isComposing) return;
    const info = resolveKey(e.target);
    if (!info) return;
    e.preventDefault();
    e.stopPropagation();
    emit(info.key, e.target.value || "");
  };

  // 원본 칸 폭을 넘는 순간 → 다음 칸으로 (입력 중 자동)
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
    emit(info.key, v);
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
    emit(info.key, v);
  };

  document.addEventListener("keydown", onKey, true);
  document.addEventListener("input", onInput, true);
  document.addEventListener("compositionend", onCompEnd, true);
  const onFocusIn = (e) => {
    const info = resolveKey(e.target);
    if (!info) return;
    try {
      setTriggerValue("focus", info.key);
    } catch (err) {}
  };
  document.addEventListener("focusin", onFocusIn, true);

  if (focusKey) {
    const go = () => {
      const el = document.querySelector(
        'div[class*="st-key-' + focusKey + '"] input'
      );
      if (!el) return false;
      try {
        el.focus({ preventScroll: false });
        const n = (el.value || "").length;
        el.setSelectionRange(n, n);
      } catch (e1) {
        try {
          el.focus();
        } catch (e2) {}
      }
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
    document.removeEventListener("focusin", onFocusIn, true);
  };
}
"""

_WL_ENTER_HOOK = st.components.v2.component(
    "worklog_cell_nav_hook_v4",
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


def _render_worklog_special_chars() -> None:
    """자주 쓰는 특수문자 — 한 줄·작은 버튼, 클릭 시 활성 칸에 삽입."""
    st.markdown(
        """
<style>
div[class*="st-key-wl_sp_"] button {
  font-size: 0.78rem !important;
  padding: 0 !important;
  min-height: 1.35rem !important;
  height: 1.35rem !important;
  line-height: 1.35rem !important;
  border-radius: 0.35rem !important;
}
div[class*="st-key-wl_sp_"] {
  min-width: 0 !important;
}
</style>
        """,
        unsafe_allow_html=True,
    )
    cols = st.columns(len(_WL_SPECIAL_CHARS), gap="small")
    for i, ch in enumerate(_WL_SPECIAL_CHARS):
        with cols[i]:
            if st.button(
                ch,
                key=f"wl_sp_{i}",
                width="stretch",
                help=f"「{ch}」삽입 (칸을 먼저 클릭)",
            ):
                st.session_state["wl_pending_special"] = ch


def _apply_pending_special_char() -> None:
    """대기 중인 특수문자를 활성 칸에 붙이거나 클립보드로 복사."""
    ch = st.session_state.pop("wl_pending_special", None)
    if not ch:
        return
    ak = st.session_state.get("wl_active_cell_key")
    if isinstance(ak, str) and (
        ak.startswith("wl_ent_ln_") or ak.startswith("wl_ent_cl_")
    ):
        st.session_state[ak] = str(st.session_state.get(ak) or "") + str(ch)
        st.session_state["wl_special_msg"] = f"「{ch}」삽입"
        return
    # 포커스된 칸이 없으면 클립보드 복사
    try:
        components.html(
            f"""<!DOCTYPE html><html><body><script>
try {{ navigator.clipboard.writeText({json.dumps(ch)}); }} catch (e) {{}}
</script></body></html>""",
            height=0,
            scrolling=False,
        )
    except Exception:
        pass
    st.session_state["wl_special_msg"] = f"「{ch}」복사됨 · 칸 클릭 후 ⌘V"


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


def open_excel_print_preview(xlsx_path: str) -> tuple[bool, str]:
    """
    macOS + Microsoft Excel 에서 원본 파일을 **열기만** 한다.
    인쇄·인쇄 미리보기·⌘P 는 실행하지 않음 (사용자가 직접 확인 후 인쇄).
    """
    abs_path = os.path.abspath(xlsx_path)
    if not os.path.exists(abs_path):
        return False, "미리보기용 엑셀 파일이 없습니다."
    if platform.system() != "Darwin":
        return False, "Excel 미리보기는 macOS 로컬 실행에서만 지원됩니다."

    excel = _excel_app_path()
    try:
        if excel:
            subprocess.Popen(
                ["open", "-a", "Microsoft Excel", abs_path],
                start_new_session=True,
            )
        else:
            subprocess.Popen(["open", abs_path], start_new_session=True)
        return True, "Excel에서 원본 양식을 열었습니다. (인쇄는 실행하지 않음)"
    except Exception as e:
        return False, f"Excel 실행 실패: {e}"


def _render_month_calendar(selected: date, saved: set[str]) -> date | None:
    """팝업 내부용 월 달력."""
    if "worklog_month" not in st.session_state:
        st.session_state["worklog_month"] = date(selected.year, selected.month, 1)
    month_anchor: date = st.session_state["worklog_month"]

    nav = st.columns([1, 1, 3, 1.4], gap="small")
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
            f"<div style='text-align:center;font-weight:700;font-size:14px;padding:4px 0;'>"
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

    st.caption("• = 저장됨 · 날짜를 누르면 해당일 일지로 전환")
    weeks = ["월", "화", "수", "목", "금", "토", "일"]
    head = st.columns(7, gap="small")
    for i, w in enumerate(weeks):
        color = "#DC2626" if i == 6 else ("#2563EB" if i == 5 else "#64748B")
        head[i].markdown(
            f"<div style='text-align:center;font-size:11px;font-weight:600;color:{color};"
            f"line-height:1.1;padding:0 0 2px 0;'>{w}</div>",
            unsafe_allow_html=True,
        )

    cal = calendar.Calendar(firstweekday=0)
    clicked = None
    for week in cal.monthdayscalendar(month_anchor.year, month_anchor.month):
        cols = st.columns(7, gap="small")
        for i, day in enumerate(week):
            with cols[i]:
                if day == 0:
                    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
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
            st.error(
                "업무일지 템플릿을 찾을 수 없습니다. "
                "`Desktop/업무일지.xlsx` 또는 `uploaded_cache/worklog/template.xlsx` 를 준비하세요."
            )
            return
    except Exception as e:
        st.error(f"템플릿 준비 실패: {e}")
        return

    @st.fragment
    def _worklog_main() -> None:
        if "worklog_selected" not in st.session_state:
            st.session_state["worklog_selected"] = date.today()
        selected: date = st.session_state["worklog_selected"]

        # 삭제 요청은 위젯 생성 전에 처리
        if st.session_state.pop("wl_do_delete_day", None) == selected.isoformat():
            delete_worklog_day(selected)

        saved = list_saved_worklog_dates()
        _init_widget_state(selected)
        draft = _cells_from_widgets(selected)

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
            try:
                view_html = render_readable_preview_html(selected, draft)
                components.html(view_html, height=900, scrolling=True)
            except Exception as e:
                st.error(f"미리보기 오류: {e}")

            p1, p2 = st.columns(2)
            with p1:
                do_print = st.button("미리보기", width="stretch", key="wl_print_btn")
            with p2:
                path_saved = worklog_path(selected)
                xbytes = b""
                src = path_saved if os.path.exists(path_saved) else None
                if src and os.path.exists(src):
                    with open(src, "rb") as f:
                        xbytes = f.read()
                else:
                    try:
                        tmp = _build_preview_file(selected, draft)
                        with open(tmp, "rb") as f:
                            xbytes = f.read()
                    except Exception:
                        xbytes = b""
                st.download_button(
                    "엑셀 저장본",
                    data=xbytes,
                    file_name=f"일일업무일지_{selected.isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key="wl_dl_btn",
                    disabled=not xbytes,
                )

            if do_print:
                cells_now = _cells_from_widgets(selected)
                try:
                    xlsx_abs = prepare_print_xlsx(selected, cells_now)
                    ok, msg = open_excel_print_preview(xlsx_abs)
                    if ok:
                        st.success(msg)
                        st.caption(f"파일: `{xlsx_abs}`")
                    else:
                        st.warning(msg + " → 화면 원본 양식으로 대체합니다.")
                        print_html = render_worklog_view_html(
                            xlsx_abs, print_mode=False, auto_print=False, scale=0.7
                        )
                        components.html(print_html, height=700, scrolling=True)
                except Exception as e:
                    st.error(f"원본 미리보기 오류: {e}")
                    try:
                        preview = _build_preview_file(selected, cells_now)
                        print_html = render_worklog_view_html(
                            preview, print_mode=False, auto_print=False, scale=0.7
                        )
                        components.html(print_html, height=700, scrolling=True)
                    except Exception as e2:
                        st.error(f"대체 미리보기도 실패: {e2}")

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
                _render_worklog_special_chars()
                msg = st.session_state.pop("wl_special_msg", None)
                if msg:
                    st.caption(msg)
            with bar_cal:
                st.markdown(
                    "<div style='height:1.55rem'></div>", unsafe_allow_html=True
                )
                with st.popover("📅 달력", width="stretch"):
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

            _apply_pending_special_char()

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

                def _on_enter_trigger():
                    hook = st.session_state.get(f"wl_enter_hook_{iso2}") or {}
                    payload = ""
                    if isinstance(hook, dict):
                        payload = str(hook.get("enter") or "")
                    if not payload:
                        return
                    done_k = f"wl_enter_done_{iso2}"
                    if st.session_state.get(done_k) == payload:
                        return
                    st.session_state[done_k] = payload
                    key = ""
                    val = ""
                    try:
                        obj = json.loads(payload)
                        key = str(obj.get("key") or "")
                        val = str(obj.get("v") or "")
                    except Exception:
                        key = payload.split(":", 1)[0]
                        val = ""
                    m = re.match(
                        r"^(wl_ent_ln|wl_ent_cl)_(\d{4}-\d{2}-\d{2})_(\d+)_(\d+)$",
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

                focus_key = st.session_state.pop(f"wl_focus_ln_{iso2}", None)
                if isinstance(focus_key, str) and (
                    focus_key.startswith("wl_ent_ln_")
                    or focus_key.startswith("wl_ent_cl_")
                ):
                    try:
                        _ei = int(focus_key.split("_")[-2])
                        st.session_state[f"wl_exp_{iso2}_{_ei}"] = True
                    except Exception:
                        pass

                # 전체 내용칸 잔여 (입력 반영) — 가운데 게이지로 표시, 배너는 숨김
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
                        label, expanded=bool(st.session_state.get(exp_key)), key=exp_key
                    ):
                        if st.button(
                            "이 항목 삭제",
                            key=f"wl_del_btn_{iso2}_{i}",
                            width="stretch",
                        ):
                            st.session_state[f"wl_do_del_{iso2}"] = i
                            _wl_rerun()
                        _cu = _client_line_units()
                        if int(
                            st.session_state.get(_entry_client_count_key(iso2, i), 0)
                            or 0
                        ) <= 0:
                            stored_e = st.session_state.get(_entries_key(d)) or []
                            if i < len(stored_e) and isinstance(
                                stored_e[i].get("client_lines"), list
                            ):
                                _seed_entry_clients(
                                    iso2, i, stored_e[i].get("client_lines") or [""]
                                )
                            else:
                                _seed_entry_clients(
                                    iso2,
                                    i,
                                    str(
                                        st.session_state.get(f"wl_ent_c_{iso2}_{i}", "")
                                        or ""
                                    ),
                                )
                        _mount_entry_client_editor(iso2, i, _cu)
                        if int(st.session_state.get(_entry_line_count_key(iso2, i), 0) or 0) <= 0:
                            lines0 = None
                            # stored entries may have lines
                            stored_e = st.session_state.get(_entries_key(d)) or []
                            if i < len(stored_e) and isinstance(stored_e[i].get("lines"), list):
                                lines0 = stored_e[i].get("lines")
                            if isinstance(lines0, list):
                                _apply_entry_lines(iso2, i, [str(x or "") for x in lines0])
                            else:
                                _seed_entry_lines(
                                    iso2,
                                    i,
                                    str(st.session_state.get(f"wl_ent_t_{iso2}_{i}", "") or ""),
                                )
                        _mount_entry_lines_editor(iso2, i, max_u)
                        _filled = len(
                            _lines_from_entry_widgets(iso2, i, keep_trailing_empty=False)
                        )
                        gap_key = f"wl_ent_gap_{iso2}_{i}"
                        if gap_key not in st.session_state:
                            st.session_state[gap_key] = _entry_blank_after(
                                (_live_entries[i] if i < len(_live_entries) else None),
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

                # Enter 훅 + 다음 칸 포커스 (입력칸 생성 후)
                _WL_ENTER_HOOK(
                    key=f"wl_enter_hook_{iso2}",
                    data={
                        "iso": iso2,
                        "focus_key": focus_key if isinstance(focus_key, str) else "",
                        "client_max_u": _client_line_units(),
                        "content_max_u": _content_line_units(),
                    },
                    on_enter_change=_on_enter_trigger,
                    on_focus_change=_on_focus_trigger,
                    width="stretch",
                    height=1,
                )
                # Enter 콜백은 훅 마운트 시점에 옴 → 다음 런 위젯 생성 전에 반영
                if st.session_state.get(f"wl_do_enter_cell_{iso2}"):
                    _wl_rerun()
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

                if st.button("＋ 항목 추가", key=f"wl_add_btn_{iso2}", width="stretch"):
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

                if st.button(
                    "저장",
                    type="primary",
                    width="stretch",
                    key=f"wl_save_btn_{iso2}",
                ):
                    try:
                        # 방금 마운트된 칸 값을 우선 반영
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
                                        st.session_state.get(f"wl_next_area_{iso2}", "") or ""
                                    ).splitlines()
                                    if x.strip()
                                ],
                                [
                                    x.strip()
                                    for x in str(
                                        st.session_state.get(f"wl_notes_area_{iso2}", "") or ""
                                    ).splitlines()
                                    if x.strip()
                                ],
                            )
                            path = save_worklog_cells(d, cells)
                            packed_entries = _grouped_entries_from_cells(cells)
                            if not packed_entries:
                                packed_entries = [{"client": "", "content": "", "lines": []}]
                            _, nd, nt = _entries_from_cells(cells)
                            st.session_state[_entries_key(d)] = packed_entries
                            st.session_state[_next_key(d)] = "\n".join(nd)
                            st.session_state[_notes_key(d)] = "\n".join(nt)
                            arch = st.session_state.get("wl_last_archive_path") or ""
                            msg = f"저장 완료: {os.path.basename(path)}"
                            if arch:
                                msg += f" · 일지/{d.year}/{os.path.basename(arch)}"
                            # iPad/Cloud는 일지 폴더가 없어도 달력 저장만으로 성공 처리
                            st.session_state[f"wl_pending_sync_{iso2}"] = {
                                "entries": packed_entries,
                                "next": "\n".join(nd),
                                "notes": "\n".join(nt),
                                "msg": msg,
                            }
                            _wl_rerun()
                    except Exception as e:
                        st.error(f"저장 실패: {e}")

            _wl_entry_editor()

        # 하단: 원본 엑셀 양식 — 배율 유지, iframe 스크롤 없이 전체 표시
        st.markdown("---")
        st.markdown("##### 원본 엑셀 양식")
        try:
            form_cells = _cells_from_widgets(selected)
            form_sig = json.dumps(form_cells, ensure_ascii=False, sort_keys=True)
            form_scale = 0.42
            sig_key = f"wl_form_sig_v8_{selected.isoformat()}"
            html_key = f"wl_form_html_v8_{selected.isoformat()}"
            h_key = f"wl_form_h_v8_{selected.isoformat()}"
            if st.session_state.get(sig_key) != form_sig:
                form_path = _build_preview_file(selected, form_cells)
                _, frame_h = _scaled_view_frame_size(form_path, form_scale)
                st.session_state[html_key] = render_worklog_view_html(
                    form_path,
                    print_mode=False,
                    scale=form_scale,
                )
                st.session_state[h_key] = int(frame_h)
                st.session_state[sig_key] = form_sig
            # 이전 세션 캐시 제거
            for k in list(st.session_state.keys()):
                if isinstance(k, str) and (
                    k.startswith("wl_form_html_v")
                    or k.startswith("wl_form_sig_v")
                    or k.startswith("wl_form_h_v")
                ):
                    if not (
                        k.startswith("wl_form_html_v8_")
                        or k.startswith("wl_form_sig_v8_")
                        or k.startswith("wl_form_h_v8_")
                    ):
                        st.session_state.pop(k, None)
            components.html(
                st.session_state.get(html_key) or "",
                height=max(240, int(st.session_state.get(h_key) or 640)),
                scrolling=False,
            )
        except Exception as e:
            st.caption(f"원본 양식 표시 실패: {e}")
    _worklog_main()
