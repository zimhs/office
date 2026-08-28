"""Streamlit Cloud → Google Drive「dashboard 복사본/uproad」원격 로드.

맥은 Drive Desktop 마운트(drive_autoload)를 쓰고, Cloud는 secrets로 Drive API 조회.
"""
from __future__ import annotations

import io
import json
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import requests

from drive_autoload import (
    DRIVE_COPY_NAME,
    _CACHE_MAP,
    _SALES_NAME_RE,
    _SKIP_ANNUAL_IF_MONTHLY,
    _atomic_copy,
)

_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.readonly"
_DRIVE_API = "https://www.googleapis.com/drive/v3"
_TIMEOUT = 90
_WORKLOG_DAY_RE = re.compile(r"^20\d{2}-\d{2}-\d{2}\.xlsx$")


def _secret_get(*keys: str) -> str:
    try:
        import streamlit as st

        secrets = getattr(st, "secrets", None)
        if secrets is None:
            return ""
        for k in keys:
            try:
                v = secrets.get(k) if hasattr(secrets, "get") else secrets[k]
            except Exception:
                v = None
            if v is not None and str(v).strip():
                return str(v).strip()
    except Exception:
        pass
    for k in keys:
        v = (os.environ.get(k) or "").strip()
        if v:
            return v
    return ""


def resolve_drive_uproad_folder_id() -> str:
    return _secret_get(
        "drive_uproad_folder_id",
        "DRIVE_UPROAD_FOLDER_ID",
        "drive_folder_id",
        "DRIVE_FOLDER_ID",
    )


def drive_remote_configured() -> bool:
    fid = resolve_drive_uproad_folder_id()
    if not fid:
        return False
    if _secret_get("google_drive_api_key", "GOOGLE_DRIVE_API_KEY"):
        return True
    return _service_account_info() is not None


def _service_account_info() -> Optional[dict]:
    try:
        import streamlit as st

        raw = st.secrets.get("google_service_account")
        if isinstance(raw, dict) and raw.get("private_key"):
            return dict(raw)
    except Exception:
        pass
    raw_env = (os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON") or "").strip()
    if raw_env:
        try:
            data = json.loads(raw_env)
            if isinstance(data, dict) and data.get("private_key"):
                return data
        except Exception:
            pass
    return None


def _auth_headers() -> Tuple[Dict[str, str], Optional[str]]:
    """(headers, api_key_query) — Bearer 또는 API key."""
    info = _service_account_info()
    if info:
        try:
            from google.oauth2 import service_account
            from google.auth.transport.requests import Request

            creds = service_account.Credentials.from_service_account_info(
                info, scopes=[_DRIVE_SCOPE]
            )
            creds.refresh(Request())
            tok = creds.token
            if tok:
                return {"Authorization": f"Bearer {tok}"}, None
        except Exception:
            pass
    api_key = _secret_get("google_drive_api_key", "GOOGLE_DRIVE_API_KEY")
    if api_key:
        return {}, api_key
    return {}, None


def _drive_get(path: str, params: Optional[dict] = None) -> Tuple[Optional[dict], Optional[str]]:
    headers, api_key = _auth_headers()
    if not headers and not api_key:
        return None, "Drive API 인증 없음 (google_service_account 또는 google_drive_api_key)"
    q = dict(params or {})
    if api_key:
        q["key"] = api_key
    try:
        r = requests.get(f"{_DRIVE_API}{path}", headers=headers, params=q, timeout=_TIMEOUT)
        if r.status_code != 200:
            return None, f"Drive API {r.status_code}: {(r.text or '')[:200]}"
        return r.json(), None
    except Exception as e:
        return None, str(e)


def _drive_download(file_id: str) -> Tuple[Optional[bytes], Optional[str]]:
    headers, api_key = _auth_headers()
    if not headers and not api_key:
        return None, "Drive API 인증 없음"
    params = {"alt": "media"}
    if api_key:
        params["key"] = api_key
    try:
        r = requests.get(
            f"{_DRIVE_API}/files/{file_id}",
            headers=headers,
            params=params,
            timeout=_TIMEOUT,
        )
        if r.status_code != 200:
            return None, f"download {r.status_code}"
        return r.content, None
    except Exception as e:
        return None, str(e)


def _list_children(folder_id: str) -> Tuple[List[dict], Optional[str]]:
    out: List[dict] = []
    page_token = None
    while True:
        params: Dict[str, Any] = {
            "q": f"'{folder_id}' in parents and trashed=false",
            "fields": "nextPageToken,files(id,name,mimeType,modifiedTime)",
            "pageSize": 1000,
        }
        if page_token:
            params["pageToken"] = page_token
        data, err = _drive_get("/files", params)
        if data is None:
            return out, err
        out.extend(data.get("files") or [])
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return out, None


def _write_bytes(dest: str, raw: bytes) -> bool:
    parent = os.path.dirname(dest)
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp = dest + ".drivepull"
    try:
        with open(tmp, "wb") as f:
            f.write(raw)
        os.replace(tmp, dest)
        return True
    except OSError:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        return False


def _dedupe_sales_names(names: List[str]) -> List[str]:
    if not _SKIP_ANNUAL_IF_MONTHLY:
        return sorted(names)
    by_year: Dict[str, List[str]] = {}
    for n in names:
        m = re.match(r"^(20\d{2})(\d{2})?\.csv$", n, re.I)
        if not m:
            continue
        by_year.setdefault(m.group(1), []).append(n)
    keep = set()
    for _y, grp in by_year.items():
        monthlies = [n for n in grp if re.match(r"^20\d{2}\d{2}\.csv$", n, re.I)]
        annuals = [n for n in grp if re.match(r"^20\d{2}\.csv$", n, re.I)]
        if monthlies:
            keep.update(monthlies)
        else:
            keep.update(annuals)
    return sorted(keep)


def sync_drive_copy_from_remote(
    cache_dir: str = "./uploaded_cache",
    *,
    force_refresh: bool = False,
) -> dict:
    """Cloud: Drive uproad 폴더 → uploaded_cache (Gist 대체)."""
    folder_id = resolve_drive_uproad_folder_id()
    if not folder_id:
        return {
            "ok": True,
            "skipped": True,
            "copied": [],
            "source": None,
            "note": f"drive_uproad_folder_id 없음 — secrets에 {DRIVE_COPY_NAME}/uproad 폴더 ID 필요",
        }
    if not drive_remote_configured():
        return {
            "ok": False,
            "skipped": True,
            "copied": [],
            "source": folder_id,
            "error": "google_service_account 또는 google_drive_api_key 필요",
        }

    files, err = _list_children(folder_id)
    if err:
        return {"ok": False, "skipped": False, "copied": [], "source": folder_id, "error": err}

    by_name = {str(f.get("name") or ""): f for f in files if f.get("name")}
    copied: List[str] = []
    os.makedirs(cache_dir, exist_ok=True)
    sales_dir = os.path.join(cache_dir, "sales")
    os.makedirs(sales_dir, exist_ok=True)

    drive_to_rel = {src_name: rel for src_name, rel, _ in _CACHE_MAP}
    for src_name, meta in by_name.items():
        if src_name not in drive_to_rel:
            continue
        rel = drive_to_rel[src_name]
        dst = os.path.join(cache_dir, rel)
        if not force_refresh and os.path.isfile(dst):
            continue
        raw, derr = _drive_download(str(meta.get("id")))
        if raw and _write_bytes(dst, raw):
            copied.append(src_name)
            name_txt = next((nt for sn, _, nt in _CACHE_MAP if sn == src_name and nt), None)
            if name_txt:
                try:
                    with open(dst + "_name.txt", "w", encoding="utf-8") as f:
                        f.write(name_txt)
                except Exception:
                    pass

    sales_names = [
        n
        for n in by_name
        if _SALES_NAME_RE.match(n) and "folder" not in str(by_name[n].get("mimeType") or "")
    ]
    wanted_sales = _dedupe_sales_names(sales_names)
    if wanted_sales:
        wanted_set = set(wanted_sales)
        if force_refresh:
            try:
                for existing in os.listdir(sales_dir):
                    if existing.endswith(".csv") and existing not in wanted_set:
                        os.remove(os.path.join(sales_dir, existing))
                        copied.append(f"-sales/{existing}")
            except OSError:
                pass
        for sn in wanted_sales:
            meta = by_name.get(sn)
            if not meta:
                continue
            dst = os.path.join(sales_dir, sn)
            if not force_refresh and os.path.isfile(dst):
                continue
            raw, _ = _drive_download(str(meta.get("id")))
            if raw and _write_bytes(dst, raw):
                copied.append(f"sales/{sn}")

    # worklog 하위 폴더
    wl_meta = by_name.get("worklog")
    if wl_meta and str(wl_meta.get("mimeType") or "").endswith("folder"):
        wl_dir = os.path.join(cache_dir, "worklog")
        os.makedirs(wl_dir, exist_ok=True)
        wl_files, wl_err = _list_children(str(wl_meta.get("id")))
        if wl_err:
            copied.append(f"worklog_err:{wl_err}")
        else:
            for wf in wl_files:
                wname = str(wf.get("name") or "")
                if not _WORKLOG_DAY_RE.match(wname) and wname != "template.xlsx":
                    continue
                wdst = os.path.join(wl_dir, wname)
                if not force_refresh and os.path.isfile(wdst):
                    continue
                raw, _ = _drive_download(str(wf.get("id")))
                if raw and _write_bytes(wdst, raw):
                    copied.append(f"worklog/{wname}")

    return {
        "ok": True,
        "skipped": False,
        "copied": copied,
        "source": f"drive:{folder_id}",
        "remote": True,
    }
