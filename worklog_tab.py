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
import sys
import unicodedata
from datetime import date
from functools import lru_cache

import streamlit as st
import streamlit.components.v1 as components

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    load_workbook = None
    get_column_letter = None

WORKLOG_DIR = os.path.join("uploaded_cache", "worklog")
WORKLOG_TEMPLATE = os.path.join(WORKLOG_DIR, "template.xlsx")
WORKLOG_TEMPLATE_SRC = os.path.expanduser("~/Desktop/업무일지.xlsx")
WORKLOG_GDRIVE_URL_FILE = os.path.join(WORKLOG_DIR, "gdrive_template_url.txt")
WORKLOG_GDRIVE_CACHE_DIR = os.path.join("uploaded_cache", "gdrive")
# 아이패드 Files / 맥 Desktop 등에 둔 8월(또는 업무일지) 양식 후보
_WORKLOG_TEMPLATE_NAME_HINTS = (
    "업무일지",
    "일일업무",
    "8월",
    "08월",
    "2026-08",
    "2026_08",
    "202608",
)


def _path_looks_like_worklog_xlsx(path: str) -> bool:
    name = os.path.basename(path)
    if not name.lower().endswith(".xlsx"):
        return False
    if name.startswith("~$") or name.startswith("_preview_"):
        return False
    if name == "template.xlsx":
        return False
    base = name[:-5]
    # 날짜 저장본(YYYY-MM-DD.xlsx)도 양식 소스로 허용(8월분 우선)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", base):
        return base.startswith("2026-08") or base[5:7] == "08"
    return any(h.lower() in name.lower() or h in name for h in _WORKLOG_TEMPLATE_NAME_HINTS)


def _iter_xlsx_under(root: str, *, max_depth: int = 3):
    if not root or not os.path.isdir(root):
        return
    root = os.path.abspath(root)
    try:
        for dirpath, dirnames, filenames in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            depth = 0 if rel == "." else rel.count(os.sep) + 1
            if depth > max_depth:
                dirnames[:] = []
                continue
            dirnames[:] = [
                d
                for d in dirnames
                if d not in {".git", "node_modules", "__pycache__", "Library"}
                and not d.startswith(".")
            ]
            for fn in filenames:
                if fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                    yield os.path.join(dirpath, fn)
    except Exception:
        return


def _template_search_roots() -> list[str]:
    home = os.path.expanduser("~")
    roots = [
        WORKLOG_DIR,
        WORKLOG_GDRIVE_CACHE_DIR,
        os.path.join("uploaded_cache", "ipad"),
        os.path.join("uploaded_cache", "아이패드"),
        os.path.join(home, "Desktop"),
        os.path.join(home, "Desktop", "아이패드"),
        os.path.join(home, "Desktop", "iPad"),
        os.path.join(home, "Desktop", "Files"),
        os.path.join(home, "Documents"),
        os.path.join(home, "Documents", "아이패드"),
        os.path.join(home, "Documents", "Files"),
        # iCloud Drive (맥·아이패드 Files 동기화)
        os.path.join(home, "Library", "Mobile Documents", "com~apple~CloudDocs"),
        os.path.join(
            home, "Library", "Mobile Documents", "com~apple~CloudDocs", "Downloads"
        ),
    ]
    out: list[str] = []
    seen: set[str] = set()
    for r in roots:
        try:
            ap = os.path.abspath(r)
        except Exception:
            continue
        if ap in seen:
            continue
        seen.add(ap)
        out.append(r)
    return out


def _score_template_candidate(path: str) -> tuple[int, float]:
    """높을수록 우선. 8월·업무일지 이름 가산, 최신 mtime 가산."""
    name = os.path.basename(path)
    score = 0
    if "업무일지" in name or "일일업무" in name:
        score += 50
    if any(h in name for h in ("8월", "08월", "2026-08", "2026_08", "202608")):
        score += 80
    if re.fullmatch(r"2026-08-\d{2}\.xlsx", name):
        score += 40
    if os.path.basename(os.path.dirname(path)) in {
        "아이패드",
        "ipad",
        "iPad",
        "Files",
        "worklog",
    }:
        score += 20
    try:
        mtime = os.path.getmtime(path)
    except Exception:
        mtime = 0.0
    return score, mtime


def find_worklog_template_source() -> str | None:
    """Desktop/아이패드 Files/캐시에서 8월·업무일지 엑셀을 찾는다."""
    if os.path.exists(WORKLOG_TEMPLATE_SRC) and os.path.isfile(WORKLOG_TEMPLATE_SRC):
        return WORKLOG_TEMPLATE_SRC
    cands: list[str] = []
    for root in _template_search_roots():
        for path in _iter_xlsx_under(root, max_depth=3):
            if _path_looks_like_worklog_xlsx(path):
                cands.append(path)
    if not cands:
        return None
    cands.sort(key=_score_template_candidate, reverse=True)
    return cands[0]


def _extract_gdrive_id(text: str) -> str | None:
    s = (text or "").strip()
    if not s:
        return None
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", s)
    if m:
        return m.group(1)
    m = re.search(r"/file/d/([a-zA-Z0-9_-]+)", s)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", s)
    if m:
        return m.group(1)
    if re.fullmatch(r"[a-zA-Z0-9_-]{20,}", s):
        return s
    return None


def _is_gdrive_folder_url(text: str) -> bool:
    return "/folders/" in (text or "")


def _load_saved_gdrive_url() -> str:
    try:
        if os.path.exists(WORKLOG_GDRIVE_URL_FILE):
            with open(WORKLOG_GDRIVE_URL_FILE, "r", encoding="utf-8") as f:
                return (f.read() or "").strip()
    except Exception:
        pass
    return ""


def _save_gdrive_url(url: str) -> None:
    os.makedirs(WORKLOG_DIR, exist_ok=True)
    with open(WORKLOG_GDRIVE_URL_FILE, "w", encoding="utf-8") as f:
        f.write((url or "").strip())


def _gdrive_url_from_secrets() -> str:
    try:
        secrets = getattr(st, "secrets", None)
        if secrets is None:
            return ""
        for key in ("WORKLOG_GDRIVE_URL", "worklog_gdrive_url", "GDRIVE_WORKLOG_URL"):
            try:
                val = secrets.get(key) if hasattr(secrets, "get") else secrets[key]
            except Exception:
                val = None
            if val:
                return str(val).strip()
    except Exception:
        pass
    # 환경변수(Streamlit Cloud secrets / Cursor env)
    for key in ("WORKLOG_GDRIVE_URL", "worklog_gdrive_url", "GDRIVE_WORKLOG_URL"):
        val = os.environ.get(key, "")
        if val:
            return str(val).strip()
    return ""


def _pick_august_xlsx(paths: list[str]) -> str | None:
    xlsx = [p for p in paths if p.lower().endswith(".xlsx") and not os.path.basename(p).startswith("~$")]
    if not xlsx:
        return None
    scored = [(p, _score_template_candidate(p)) for p in xlsx if _path_looks_like_worklog_xlsx(p)]
    if not scored:
        scored = [(p, _score_template_candidate(p)) for p in xlsx]
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[0][0]


def download_worklog_template_from_gdrive(url: str) -> tuple[bool, str]:
    """구글드라이브 공유 링크(파일/폴더)에서 8월·업무일지 xlsx를 template으로 저장."""
    url = (url or "").strip()
    if not url:
        return False, "구글드라이브 링크가 비어 있습니다."
    try:
        import gdown
    except Exception:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "gdown", "-q"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            import gdown  # type: ignore
        except Exception as e:
            return False, f"gdown 설치 실패: {e}"

    os.makedirs(WORKLOG_GDRIVE_CACHE_DIR, exist_ok=True)
    os.makedirs(WORKLOG_DIR, exist_ok=True)
    file_id = _extract_gdrive_id(url)
    try:
        if _is_gdrive_folder_url(url) or (
            file_id and "folders" in url
        ):
            # 폴더: 내려받은 뒤 8월/업무일지 xlsx 선택
            out_dir = os.path.join(WORKLOG_GDRIVE_CACHE_DIR, "folder")
            if os.path.isdir(out_dir):
                shutil.rmtree(out_dir, ignore_errors=True)
            os.makedirs(out_dir, exist_ok=True)
            folder_url = url if "http" in url else f"https://drive.google.com/drive/folders/{file_id}"
            gdown.download_folder(folder_url, output=out_dir, quiet=True, use_cookies=False)
            found: list[str] = []
            for root, _, files in os.walk(out_dir):
                for fn in files:
                    if fn.lower().endswith(".xlsx"):
                        found.append(os.path.join(root, fn))
            picked = _pick_august_xlsx(found)
            if not picked:
                return False, "폴더에서 8월/업무일지 xlsx를 찾지 못했습니다."
            shutil.copy2(picked, WORKLOG_TEMPLATE)
            ipad_copy = os.path.join(
                "uploaded_cache", "ipad", os.path.basename(picked) or "8월_업무일지.xlsx"
            )
            os.makedirs(os.path.dirname(ipad_copy), exist_ok=True)
            shutil.copy2(picked, ipad_copy)
        else:
            tmp_path = os.path.join(WORKLOG_GDRIVE_CACHE_DIR, "drive_download.xlsx")
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
            # 파일 공유 링크 / ID
            if file_id and "http" not in url:
                result = gdown.download(id=file_id, output=tmp_path, quiet=True)
            else:
                result = gdown.download(url=url, output=tmp_path, quiet=True, fuzzy=True)
            if not result or not os.path.exists(tmp_path) or os.path.getsize(tmp_path) < 64:
                return False, "다운로드 실패(공유를 '링크가 있는 모든 사용자'로 열어 주세요)."
            # HTML 에러 페이지가 저장된 경우 거부
            with open(tmp_path, "rb") as f:
                head = f.read(32)
            if head.lstrip().startswith(b"<") or b"<!DOCTYPE" in head[:64].upper():
                return False, "다운로드가 HTML입니다. Drive 공유 권한을 확인해 주세요."
            shutil.copy2(tmp_path, WORKLOG_TEMPLATE)
            ipad_copy = os.path.join("uploaded_cache", "ipad", "8월_업무일지_gdrive.xlsx")
            os.makedirs(os.path.dirname(ipad_copy), exist_ok=True)
            shutil.copy2(tmp_path, ipad_copy)
        try:
            _content_line_units.cache_clear()
        except Exception:
            pass
        _save_gdrive_url(url)
        return True, f"구글드라이브에서 양식 연결: {os.path.basename(WORKLOG_TEMPLATE)}"
    except Exception as e:
        return False, f"구글드라이브 불러오기 실패: {e}"


def try_fetch_template_from_gdrive() -> bool:
    """secrets/저장 URL로 템플릿 자동 동기화. 성공 시 True."""
    if os.path.exists(WORKLOG_TEMPLATE):
        return True
    url = _gdrive_url_from_secrets() or _load_saved_gdrive_url()
    if not url:
        return False
    ok, _msg = download_worklog_template_from_gdrive(url)
    return ok

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

# 내용 칸 편집기 (CCv2) — Enter로 다음 칸 생성/이동을 JS에서 직접 처리
_WL_LINES_HTML = """
<div class="wl-lines"></div>
"""

_WL_LINES_CSS = """
/* 맥·iPad 공통: 16px로 Safari 자동줌 방지, 터치 타깃 확보 */
.wl-lines { display: flex; flex-direction: column; gap: 6px; width: 100%; max-width: 100%; box-sizing: border-box; }
.wl-row { display: flex; gap: 6px; align-items: center; width: 100%; }
.wl-row input {
  flex: 1 1 auto;
  min-width: 0;
  width: 100%;
  min-height: 44px;
  height: 44px;
  padding: 0.35rem 0.5rem;
  border: 1px solid #94A3B8;
  border-radius: 6px;
  background: #fff;
  color: #0F172A;
  font-size: 16px;
  line-height: 1.25;
  outline: none;
  -webkit-appearance: none;
  appearance: none;
}
.wl-row input:focus {
  border-color: #0F766E;
  box-shadow: 0 0 0 1px #0F766E;
}
.wl-row button {
  flex: 0 0 52px;
  min-height: 44px;
  height: 44px;
  border: 1px solid #CBD5E1;
  border-radius: 6px;
  background: #F8FAFC;
  color: #334155;
  font-size: 14px;
  padding: 0;
  cursor: pointer;
  -webkit-tap-highlight-color: transparent;
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


def _is_touch_ui() -> bool:
    """iPad/터치 UI 분기. app.py is_touch_ui와 동일 기준(맥 무손실)."""
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
    except Exception:
        pass
    return bool(st.session_state.get("force_touch_ui"))


def _inject_worklog_touch_css() -> None:
    """일일업무일지 탭만 — iPad 레이아웃·버튼·iframe 안정화."""
    st.markdown(
        """
        <style>
        .worklog-touch-scope div[data-testid="stButton"] > button,
        .worklog-touch-scope div[data-testid="stDownloadButton"] > button {
            min-height: 48px !important;
            font-size: 15px !important;
            font-weight: 600 !important;
        }
        .worklog-touch-scope textarea,
        .worklog-touch-scope input,
        .worklog-touch-scope [data-baseweb="input"] input,
        .worklog-touch-scope [data-baseweb="textarea"] textarea,
        .worklog-touch-scope [data-baseweb="select"] {
            font-size: 16px !important;
            min-height: 44px !important;
        }
        .worklog-touch-scope iframe {
            max-width: 100% !important;
            -webkit-overflow-scrolling: touch !important;
        }
        .worklog-touch-scope [data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _offer_template_upload() -> bool:
    """Cloud/iPad: 구글드라이브·Files에서 8월 양식 준비."""
    # 1) 구글드라이브 자동(시크릿/저장 URL)
    if try_fetch_template_from_gdrive() and os.path.exists(WORKLOG_TEMPLATE):
        st.success("구글드라이브 8월 양식을 연결했습니다.")
        st.rerun()

    # 2) 로컬/캐시 자동
    found = find_worklog_template_source()
    if found and os.path.isfile(found):
        try:
            _ensure_dirs()
            if not os.path.exists(WORKLOG_TEMPLATE):
                shutil.copy2(found, WORKLOG_TEMPLATE)
            st.success(f"8월/업무일지 양식을 연결했습니다: `{os.path.basename(found)}`")
            st.rerun()
        except Exception as e:
            st.error(f"양식 연결 실패: {e}")

    st.warning(
        "업무일지 템플릿이 없습니다. "
        "**구글드라이브 8월 엑셀** 공유 링크를 넣거나, 아이패드 Files에서 파일을 선택하세요."
    )

    saved_url = _gdrive_url_from_secrets() or _load_saved_gdrive_url()
    gurl = st.text_input(
        "구글드라이브 링크 (8월 업무일지 파일 또는 폴더)",
        value=saved_url,
        key="wl_gdrive_url",
        placeholder="https://drive.google.com/file/d/... 또는 /folders/...",
        help="공유를 ‘링크가 있는 모든 사용자’로 연 뒤 URL을 붙여 넣으세요.",
    )
    if st.button("📥 구글드라이브에서 불러오기", key="wl_gdrive_fetch", width="stretch"):
        with st.spinner("구글드라이브에서 8월 엑셀 다운로드 중..."):
            ok, msg = download_worklog_template_from_gdrive(gurl)
        if ok:
            st.success(msg)
            st.rerun()
        st.error(msg)

    st.caption("또는 아이패드 Files에서 직접 업로드")
    up = st.file_uploader(
        "아이패드 Files → 8월 업무일지.xlsx",
        type=["xlsx"],
        key="wl_template_uploader",
        help="Files 앱 / 구글드라이브 앱에서 8월 업무일지 엑셀을 올리면 양식으로 저장됩니다.",
    )
    if up is None:
        return False
    try:
        _ensure_dirs()
        raw = up.getvalue()
        safe_name = os.path.basename(up.name or "8월_업무일지.xlsx")
        if not safe_name.lower().endswith(".xlsx"):
            safe_name += ".xlsx"
        ipad_dir = os.path.join("uploaded_cache", "ipad")
        os.makedirs(ipad_dir, exist_ok=True)
        ipad_path = os.path.join(ipad_dir, safe_name)
        with open(ipad_path, "wb") as f:
            f.write(raw)
        with open(WORKLOG_TEMPLATE, "wb") as f:
            f.write(raw)
        try:
            _content_line_units.cache_clear()
        except Exception:
            pass
        st.success(f"템플릿 저장: {safe_name}")
        st.rerun()
    except Exception as e:
        st.error(f"템플릿 저장 실패: {e}")
    return False


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


@lru_cache(maxsize=1)
def _content_line_units() -> int:
    """원본 G:X 병합 폭 기준 — 한 칸을 거의 채운 뒤 다음 행으로 넘김(반각=1)."""
    fallback = 56
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
        # 열폭=기본글꼴(맑은고딕 11) 기준, 셀 글꼴 바탕체 14 → 칸을 거의 끝까지 채움
        units = int(total * (11 / 14) * 1.05)
        return max(56, min(units, 74))
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
    os.makedirs(os.path.join("uploaded_cache", "ipad"), exist_ok=True)
    os.makedirs(WORKLOG_GDRIVE_CACHE_DIR, exist_ok=True)
    if os.path.exists(WORKLOG_TEMPLATE):
        return
    # 1순위: 구글드라이브(시크릿/저장 링크)
    try:
        if try_fetch_template_from_gdrive() and os.path.exists(WORKLOG_TEMPLATE):
            return
    except Exception:
        pass
    # 2순위: Desktop / Files / 캐시
    src = find_worklog_template_source()
    if src and os.path.isfile(src):
        try:
            shutil.copy2(src, WORKLOG_TEMPLATE)
            try:
                _content_line_units.cache_clear()
            except Exception:
                pass
        except Exception:
            pass


def worklog_path(d: date) -> str:
    return os.path.join(WORKLOG_DIR, f"{d.isoformat()}.xlsx")


def list_saved_worklog_dates() -> set[str]:
    _ensure_dirs()
    out: set[str] = set()
    for name in os.listdir(WORKLOG_DIR):
        if (
            name.endswith(".xlsx")
            and len(name) >= 15
            and name[0:4].isdigit()
            and name not in {"template.xlsx"}
            and not name.startswith("_preview_")
        ):
            out.add(name.replace(".xlsx", ""))
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
    path = worklog_path(d)
    is_new = not os.path.exists(path)
    cells = _spill_all_content(cells)
    write_cells_to_path(path, d, cells, blank_base=is_new)
    return path


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
                parts.append(f"{css_side}:1px solid #111;")
    except Exception:
        pass
    return "".join(parts) if parts else "border:1px solid #333;"


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
        px = max(12, int(float(w) * 8))
        col_widths.append(px)
        total_w += px

    rows_html = []
    for r in range(WL_MIN_ROW, WL_MAX_ROW + 1):
        h = ws.row_dimensions[r].height
        height_px = int(float(h) * 1.333) if h else 28
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
            fsize = int(font.size or 12)
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
            esc = html.escape(text).replace("\n", "<br>")
            # 내용칸은 한 행=한 줄 (원본 엑셀과 동일하게 칸 밖으로 줄바꿈하지 않음)
            is_content = c == 7 and 9 <= r <= 39
            white = "nowrap" if is_content else "pre-wrap"
            overflow = "hidden" if is_content else "visible"
            style = (
                f"font-family:'{html.escape(fname)}','Apple SD Gothic Neo','Batang',serif;"
                f"font-size:{fsize}px;font-weight:{bold};"
                f"text-align:{ha};vertical-align:{va};"
                f"background:{fill};{border}"
                f"padding:2px 4px;white-space:{white};overflow:{overflow};"
                f"word-break:keep-all;max-width:100%;"
                f"height:{height_px}px;min-height:{height_px}px;max-height:{height_px}px;"
            )
            tds.append(f'<td{span} style="{style}">{esc}</td>')
        rows_html.append(f'<tr style="height:{height_px}px">{"".join(tds)}</tr>')

    colgroup = "".join(f'<col style="width:{w}px">' for w in col_widths)
    wb.close()
    return f"""
    <table class="wl-sheet" style="border-collapse:collapse;table-layout:fixed;width:{int(total_w)}px;background:#fff;">
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
    scale_css = f"transform:scale({scale}); transform-origin:top left;" if scale < 1 else ""
    wrap_height = "auto" if print_mode else "860px"
    auto_script = ""
    if auto_print:
        auto_script = """
        <script>
          (function() {
            function go() {
              try { window.focus(); window.print(); } catch (e) {}
            }
            // iframe 로드 직후 인쇄창
            if (document.readyState === 'complete') setTimeout(go, 250);
            else window.addEventListener('load', function() { setTimeout(go, 250); });
          })();
        </script>
        """
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>일일업무일지</title>
<style>
  @page {{ size: A4 portrait; margin: 10mm; }}
  html, body {{ margin:0; padding:0; background:#fff; }}
  body {{ padding:6px; }}
  .toolbar {{ margin-bottom:10px; display:flex; gap:10px; align-items:center; flex-wrap:wrap; }}
  .toolbar button {{
    padding:8px 14px; font-size:14px; border:1px solid #334155; border-radius:6px;
    background:#1E293B; color:#fff; cursor:pointer;
  }}
  .toolbar .hint {{ font:12px/1.4 sans-serif; color:#64748B; }}
  .wrap {{
    overflow:auto; height:{wrap_height};
    border:1px solid #94A3B8; background:#fff;
  }}
  .sheet-scale {{ {scale_css} width:fit-content; }}
  @media print {{
    .toolbar {{ display:none !important; }}
    .wrap {{ overflow:visible; height:auto; border:none; }}
    .sheet-scale {{ transform:none !important; }}
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


def _grouped_entries_from_cells(cells: dict) -> list[dict]:
    """엑셀 행 → 논리 항목.

    - 거래처가 있으면 새 항목 (앞 빈 행 = 이전 blank_after)
    - 거래처 없이 내용만 있으면 이전 항목에 칸 추가 (빈 칸도 유지)
    """
    entries: list[dict] = []
    blank_run = 0
    for r in WL_CLIENT_ROWS:
        client = _scrub_dummy_label(str(cells.get(f"C{r}", "") or "")).strip()
        content = _scrub_dummy_label(str(cells.get(f"G{r}", "") or ""))
        empty = not client and not (content or "").strip()
        if empty:
            blank_run += 1
            continue
        if client or not entries:
            if entries:
                entries[-1]["blank_after"] = max(0, min(10, blank_run))
            blank_run = 0
            lines = [content or ""]
            entries.append(
                {
                    "client": client,
                    "content": "\n".join(lines),
                    "lines": lines,
                    "blank_after": 1,
                }
            )
        else:
            # 항목 내부 빈 칸 보존 후 내용 칸 추가
            if blank_run > 0:
                for _ in range(blank_run):
                    entries[-1].setdefault("lines", []).append("")
                blank_run = 0
            entries[-1].setdefault("lines", []).append(content or "")
            entries[-1]["content"] = "\n".join(entries[-1].get("lines") or [])
    return entries


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
    if not out and str(ent.get("client") or "").strip():
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
        client = str(ent.get("client") or "").strip()
        pack_lines = _entry_pack_lines(ent)
        if not client and not any((x or "").strip() for x in pack_lines):
            per_entry.append(0)
            continue
        gap = prev_gap if wrote_any else 0
        lines = max(1, len(pack_lines))
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
        client = str(ent.get("client") or "").strip()
        chunks = _entry_pack_lines(ent)
        if not client and not any((x or "").strip() for x in chunks):
            continue
        # 이전 항목이 지정한 빈 칸만큼 띄움
        if wrote_any:
            row_i += max(0, int(prev_gap))
        if not chunks:
            chunks = [""]
        for j, head in enumerate(chunks):
            if row_i >= len(rows):
                break
            r = rows[row_i]
            cells[f"C{r}"] = client if j == 0 else ""
            cells[f"G{r}"] = head
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
            c = html.escape(client) if client else "<span class='muted'>(거래처 없음)</span>"
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
  .client {{ font-size:15px; font-weight:700; color:#134E4A; margin-bottom:4px; }}
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
      <div class="sub">{date_label} · 화면용 요약 (저장은 원본 엑셀 양식 유지)</div>
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
    """엑셀 → 항목(거래처/내용) 세션 상태로 로드."""
    cells = read_worklog_cells(d)
    bk = _boot_key(d)
    ek = _entries_key(d)
    # 예전 칸별입력 boot 만 있는 경우도 항목 모드로 마이그레이션
    if not st.session_state.get(bk) or ek not in st.session_state:
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
    """칸 위젯/컴포넌트 → 줄 목록. 중간 빈 칸은 유지. 끝의 편집용 빈 칸 1개는 선택적 유지."""
    live = st.session_state.get(_entry_lines_live_key(iso, entry_i))
    ck = _entry_lines_comp_key(iso, entry_i)
    cs = st.session_state.get(ck)
    if isinstance(live, list):
        parts = [str(x or "") for x in live]
    elif isinstance(cs, dict) and isinstance(cs.get("lines"), list):
        parts = [str(x or "") for x in cs.get("lines") or []]
    else:
        lc = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
        if lc <= 0:
            raw = str(st.session_state.get(f"wl_ent_t_{iso}_{entry_i}", "") or "")
            parts = [raw] if raw else ([""] if keep_trailing_empty else [])
        else:
            parts = [
                str(st.session_state.get(_entry_line_key(iso, entry_i, j), "") or "")
                for j in range(lc)
            ]
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
    """줄 목록을 컴포넌트/레거시 키에 반영. 끝에 편집용 빈 칸 1개 유지."""
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
    _set_comp_lines_state(iso, entry_i, chunks, focus_j=focus_j)
    rk = _entry_lines_rev_key(iso, entry_i)
    st.session_state[rk] = int(st.session_state.get(rk, 0) or 0) + 1


def _mount_entry_lines_editor(iso: str, entry_i: int, max_u: int) -> list[str]:
    """내용 칸 CCv2 편집터. Enter=다음 칸 생성/이동."""
    ck = _entry_lines_comp_key(iso, entry_i)
    live_key = _entry_lines_live_key(iso, entry_i)
    cs = st.session_state.get(ck)
    if isinstance(cs, dict) and isinstance(cs.get("lines"), list):
        lines = [str(x or "") for x in cs.get("lines") or []]
        focus = cs.get("focus", -1)
    elif isinstance(st.session_state.get(live_key), list):
        lines = [str(x or "") for x in st.session_state.get(live_key) or []]
        focus = -1
    else:
        lines = [""]
        focus = -1
    if not lines or lines[-1] != "":
        lines = list(lines) + [""]
    rev = int(st.session_state.get(_entry_lines_rev_key(iso, entry_i), 0) or 0)
    try:
        focus_n = int(focus)
    except (TypeError, ValueError):
        focus_n = -1

    def _on_lines_change() -> None:
        cur = st.session_state.get(ck)
        if isinstance(cur, dict) and isinstance(cur.get("lines"), list):
            synced = [str(x or "") for x in cur.get("lines") or []]
            st.session_state[live_key] = synced
            st.session_state[_entry_line_count_key(iso, entry_i)] = len(synced)
            for j, line in enumerate(synced):
                st.session_state[_entry_line_key(iso, entry_i, j)] = line
            st.session_state[f"wl_ent_t_{iso}_{entry_i}"] = "\n".join(synced)

    result = _WL_LINES_EDITOR(
        key=ck,
        data={"lines": lines, "focus": focus_n, "max_u": int(max_u), "rev": rev},
        default={"lines": lines, "focus": focus_n},
        on_lines_change=_on_lines_change,
        on_focus_change=lambda: None,
        width="stretch",
        height="content",
    )
    out = result.lines if isinstance(getattr(result, "lines", None), list) else lines
    out = [str(x or "") for x in out]
    if not out or out[-1] != "":
        out = list(out) + [""]
    st.session_state[live_key] = out
    # 레거시 키 동기화(저장/사용량용). rev는 올리지 않음.
    old = int(st.session_state.get(_entry_line_count_key(iso, entry_i), 0) or 0)
    for j in range(max(old, len(out)) + 3):
        st.session_state.pop(_entry_line_key(iso, entry_i, j), None)
    st.session_state[_entry_line_count_key(iso, entry_i)] = len(out)
    for j, line in enumerate(out):
        st.session_state[_entry_line_key(iso, entry_i, j)] = line
    st.session_state[f"wl_ent_t_{iso}_{entry_i}"] = "\n".join(out)
    return out


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
        if ck in st.session_state or lc > 0 or f"wl_ent_t_{iso}_{i}" in st.session_state:
            client = str(st.session_state.get(ck, "") or "")
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
            client, content, blank_after, lines = "", "", 1, []
        out.append(
            {
                "client": client,
                "content": content,
                "lines": lines,
                "blank_after": blank_after,
            }
        )
    return out or [{"client": "", "content": "", "lines": [], "blank_after": 1}]


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
    """날짜 파일은 건드리지 않고 미리보기용 임시 xlsx 생성."""
    _ensure_dirs()
    src = worklog_path(d) if os.path.exists(worklog_path(d)) else WORKLOG_TEMPLATE
    dst = _preview_path(d)
    shutil.copy2(src, dst)
    write_cells_to_path(dst, d, cells, blank_base=not os.path.exists(worklog_path(d)))
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
    macOS + Microsoft Excel 에서 원본 파일을 열고 인쇄 대화상자(미리보기)를 띄움.
    원격/서버 환경이거나 Excel이 없으면 실패 → 호출측에서 HTML 폴백.
    """
    abs_path = os.path.abspath(xlsx_path)
    if not os.path.exists(abs_path):
        return False, "인쇄용 엑셀 파일이 없습니다."
    if platform.system() != "Darwin":
        return False, "Excel 인쇄 미리보기는 macOS 로컬 실행에서만 지원됩니다."

    excel = _excel_app_path()
    if not excel:
        # Excel 없으면 기본 앱으로라도 연다
        try:
            subprocess.Popen(["open", abs_path], start_new_session=True)
            return True, "기본 앱으로 엑셀 파일을 열었습니다. 앱에서 ⌘P 로 인쇄하세요."
        except Exception as e:
            return False, f"파일 열기 실패: {e}"

    # 1) Excel로 원본 파일 오픈
    # 2) ⌘P → Excel 인쇄 창(미리보기 포함)
    # System Events 접근 권한이 없으면 파일만 열린 상태로 안내
    escaped = abs_path.replace("\\", "\\\\").replace('"', '\\"')
    script = f'''
set theFile to POSIX file "{escaped}"
tell application "Microsoft Excel"
    activate
    open theFile
end tell
delay 1.0
try
    tell application "System Events"
        tell process "Microsoft Excel"
            set frontmost to true
            keystroke "p" using {{command down}}
        end tell
    end tell
    return "print"
on error
    return "opened"
end try
'''
    try:
        r = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=20,
        )
        if r.returncode != 0:
            # AppleScript 실패 시 open -a 로라도 오픈
            subprocess.Popen(
                ["open", "-a", "Microsoft Excel", abs_path],
                start_new_session=True,
            )
            return True, "Excel에서 원본 양식을 열었습니다. ⌘P 로 인쇄 미리보기를 여세요."
        out = (r.stdout or "").strip()
        if out == "print":
            return True, "Excel 원본 양식 + 인쇄 미리보기를 열었습니다."
        return True, "Excel에서 원본 양식을 열었습니다. ⌘P 로 인쇄 미리보기를 여세요."
    except subprocess.TimeoutExpired:
        return True, "Excel 실행 중입니다. 창이 뜨면 ⌘P 로 인쇄하세요."
    except Exception as e:
        try:
            subprocess.Popen(
                ["open", "-a", "Microsoft Excel", abs_path],
                start_new_session=True,
            )
            return True, f"Excel로 파일을 열었습니다(⌘P). 참고: {e}"
        except Exception as e2:
            return False, f"Excel 실행 실패: {e2}"


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
            st.rerun()
    with nav[1]:
        if st.button("▶", key="wl_next_month", width="stretch"):
            y, m = month_anchor.year, month_anchor.month + 1
            if m > 12:
                y, m = y + 1, 1
            st.session_state["worklog_month"] = date(y, m, 1)
            st.rerun()
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
            st.rerun()

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
    """일일업무일지 탭 본체."""
    touch = _is_touch_ui()
    if touch:
        _inject_worklog_touch_css()
        st.markdown("<div class='worklog-touch-scope'>", unsafe_allow_html=True)

    head_l, head_r = st.columns([4, 1])
    head_l.markdown(
        "<div class='sub-header dashboard-tab-panel-head'>일일업무일지</div>",
        unsafe_allow_html=True,
    )
    if latest_update_str:
        head_r.caption(latest_update_str)

    if load_workbook is None:
        st.error("openpyxl 이 필요합니다. `pip install openpyxl` 후 다시 실행하세요.")
        if touch:
            st.markdown("</div>", unsafe_allow_html=True)
        return
    try:
        _ensure_dirs()
        if not os.path.exists(WORKLOG_TEMPLATE):
            _offer_template_upload()
            if touch:
                st.markdown("</div>", unsafe_allow_html=True)
            return
    except Exception as e:
        st.error(f"템플릿 준비 실패: {e}")
        if touch:
            st.markdown("</div>", unsafe_allow_html=True)
        return

    if "worklog_selected" not in st.session_state:
        st.session_state["worklog_selected"] = date.today()
    selected: date = st.session_state["worklog_selected"]
    saved = list_saved_worklog_dates()
    _init_widget_state(selected)
    draft = _cells_from_widgets(selected)
    status = "저장됨" if selected.isoformat() in saved else "미저장"

    # 날짜: 팝업(popover) — 화면 공간 최소 사용
    bar_l, bar_r = st.columns([3.2, 1.2], gap="small")
    with bar_l:
        st.markdown(
            f"**선택일** {format_worklog_date(selected)} · {status}",
        )
    with bar_r:
        with st.popover("📅 날짜 선택", width="stretch"):
            clicked = _render_month_calendar(selected, saved)
            if clicked is not None and clicked != selected:
                _clear_date_widget_state(selected)
                st.session_state["worklog_selected"] = clicked
                st.rerun()

    # 맥: 좌 미리보기 / 우 입력. iPad: 세로 스택(입력 먼저)으로 터치 안정화.
    if touch:
        col_input = st.container()
        col_preview = st.container()
    else:
        col_preview, col_input = st.columns(2, gap="medium")

    with col_preview:
        st.markdown("##### 업무일지 보기")
        st.caption("요약 미리보기 · 저장하면 원본 엑셀 양식에 반영")
        try:
            view_html = render_readable_preview_html(selected, draft)
            preview_h = 620 if touch else 900
            components.html(view_html, height=preview_h, scrolling=True)
        except Exception as e:
            st.error(f"미리보기 오류: {e}")

        p1, p2, p3 = st.columns(3)
        with p1:
            do_print = st.button("원본양식 인쇄", width="stretch", key="wl_print_btn")
        with p2:
            path_saved = worklog_path(selected)
            xbytes = b""
            src = path_saved if os.path.exists(path_saved) else None
            if src and os.path.exists(src):
                with open(src, "rb") as f:
                    xbytes = f.read()
            else:
                # 미저장이면 현재 입력으로 임시 엑셀 생성 후 제공
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
        with p3:
            st.caption("인쇄 = Excel 원본" if not touch else "인쇄 = 브라우저")

    with col_input:
        st.markdown("##### 업무 입력")
        _max_u_cap = _content_line_units()
        st.caption(
            f"내용 칸이 가득 차면(반각 {_max_u_cap}자 / 한글 {_max_u_cap // 2}자) "
            "자동으로 다음 칸으로 넘어갑니다"
        )

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
            st.session_state.pop(f"wl_next_area_{iso}", None)
            st.session_state.pop(f"wl_notes_area_{iso}", None)
            st.session_state[ek] = entries_list
            st.session_state[f"wl_entry_count_{iso}"] = len(entries_list)
            for i, ent in enumerate(entries_list):
                st.session_state[f"wl_ent_c_{iso}_{i}"] = ent.get("client") or ""
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
                entries.append({"client": "", "content": "", "blank_after": 1})
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

        @st.fragment
        def _wl_entry_editor():
            d = st.session_state.get("worklog_selected") or selected
            iso2 = d.isoformat()
            n = int(st.session_state.get(f"wl_entry_count_{iso2}", 1) or 1)
            max_u = _content_line_units()

            _force_open = st.session_state.pop(f"wl_force_expand_{iso2}", None)
            if isinstance(_force_open, int):
                st.session_state[f"wl_exp_{iso2}_{_force_open}"] = True

            # 전체 내용칸 잔여 (입력 반영)
            _live_entries = _read_editor_entries(d)
            _usage = _content_row_usage(_live_entries)
            _rem = _usage["remaining"]
            _used = _usage["used"]
            _tot = _usage["total"]
            _warn = _usage["overflow"] or _rem <= 3
            _color = "#B91C1C" if _warn else "#0F766E"
            _next = _usage["next_row"]
            _next_txt = f"다음 빈 행 G{_next}" if _next else "내용칸 끝"
            st.markdown(
                f"<div style='margin:0 0 10px;padding:8px 12px;border-radius:8px;"
                f"border:1px solid {'#FECACA' if _warn else '#99F6E4'};"
                f"background:{'#FEF2F2' if _warn else '#F0FDFA'};"
                f"font-size:12px;color:{_color};font-weight:600;'>"
                f"원본 내용칸 · 사용 {_used}/{_tot}행 · "
                f"<span style='font-size:14px;'>남은 {_rem}행</span>"
                f" · 마지막 G{_usage['last_row']} · {_next_txt}"
                f"{' · 용량 초과' if _usage['overflow'] else ''}"
                f"</div>",
                unsafe_allow_html=True,
            )

            for i in range(n):
                client_now = str(st.session_state.get(f"wl_ent_c_{iso2}_{i}", "") or "").strip()
                body_now = _content_from_entry_lines(iso2, i).strip().replace("\n", " ")
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

                with st.expander(label, expanded=bool(st.session_state.get(exp_key)), key=exp_key):
                    if st.button(
                        "이 항목 삭제",
                        key=f"wl_del_btn_{iso2}_{i}",
                        width="stretch",
                    ):
                        st.session_state[f"wl_do_del_{iso2}"] = i
                        st.rerun()
                    st.text_input(
                        "거래처",
                        key=f"wl_ent_c_{iso2}_{i}",
                        placeholder="거래처명 입력",
                    )
                    st.markdown(
                        "<div style='font-size:11px;color:#64748B;margin:4px 0 2px;'>내용 "
                        f"(한 칸 ≈ 반각 {max_u}자 · <b>Enter=다음 칸 생성</b> · 가득 차면 자동 분할)</div>",
                        unsafe_allow_html=True,
                    )
                    ck = _entry_lines_comp_key(iso2, i)
                    if not (
                        isinstance(st.session_state.get(ck), dict)
                        and isinstance((st.session_state.get(ck) or {}).get("lines"), list)
                    ):
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
                        help="이 항목 다음에 원본 엑셀에서 비워 둘 행 수 (0이면 바로 붙음)",
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

            if st.button("＋ 항목 추가", key=f"wl_add_btn_{iso2}", width="stretch"):
                st.session_state[f"wl_do_add_{iso2}"] = True
                st.rerun()

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
                        st.session_state[f"wl_pending_sync_{iso2}"] = {
                            "entries": packed_entries,
                            "next": "\n".join(nd),
                            "notes": "\n".join(nt),
                            "msg": f"저장 완료: {os.path.basename(path)}",
                        }
                        st.rerun()
                except Exception as e:
                    st.error(f"저장 실패: {e}")

        _wl_entry_editor()


    if do_print:
        cells_now = _cells_from_widgets(selected)
        try:
            xlsx_abs = prepare_print_xlsx(selected, cells_now)
            # iPad/Cloud: Excel 앱 연동 불가 → 브라우저 원본 양식 인쇄만
            if touch:
                print_html = render_worklog_view_html(
                    xlsx_abs, print_mode=True, auto_print=True
                )
                components.html(print_html, height=900, scrolling=True)
            else:
                ok, msg = open_excel_print_preview(xlsx_abs)
                if ok:
                    st.success(msg)
                    st.caption(f"파일: `{xlsx_abs}`")
                else:
                    st.warning(msg + " → 브라우저 원본 양식 인쇄로 대체합니다.")
                    print_html = render_worklog_view_html(
                        xlsx_abs, print_mode=True, auto_print=True
                    )
                    components.html(print_html, height=900, scrolling=True)
        except Exception as e:
            st.error(f"원본 인쇄 오류: {e}")
            try:
                preview = _build_preview_file(selected, cells_now)
                print_html = render_worklog_view_html(
                    preview, print_mode=True, auto_print=True
                )
                components.html(print_html, height=900, scrolling=True)
            except Exception as e2:
                st.error(f"대체 인쇄도 실패: {e2}")

    if touch:
        st.markdown("</div>", unsafe_allow_html=True)
