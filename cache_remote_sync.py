"""Streamlit Cloud ↔ 맥 사이드바 캐시(uploaded_cache) Gist 동기화.

Cloud 재부팅 시 git 시드(옛 데이터) 대신 Gist의 최신 CSV/캐시를 받습니다.
맥에서 사이드바 업로드·Drive 동기화 후 Gist에도 올립니다.

secrets: github_token + dashboard_cache_gist_id (없으면 첫 push 시 생성)
"""
from __future__ import annotations

import base64
import gzip
import hashlib
import json
import os
import re
import time
from typing import Any, Dict, List, Optional, Set, Tuple

import requests

from worklog_remote_sync import _GIST_API, _secret_get, _sha256_bytes, _sha256_file, resolve_github_token

_MANIFEST = "_dashboard_cache_manifest.json"
_GIST_ID_FILE = ".dashboard_cache_gist_id"
_B64_SUFFIX = ".gz.b64"
_TIMEOUT = 60
_SALES_RE = re.compile(r"^20\d{2}(\d{2})?\.csv$", re.I)

# 동기화 대상 (worklog·API키·업로드 임시 제외)
_STATIC_REL = (
    "address.csv",
    "industry.csv",
    "debt.csv",
    "tank_cache.dat",
    "tank_cache.dat_name.txt",
    "vaporizer_cache.dat",
    "vaporizer_cache.dat_name.txt",
    "integrated_cache.dat",
    "integrated_cache.dat_name.txt",
    "price_increase/mail_contacts.csv",
)


def _headers(token: str) -> Dict[str, str]:
    return {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _gist_key(rel: str) -> str:
    return rel.replace("/", "__").replace("\\", "__")


def _rel_from_gist_key(key: str) -> str:
    if key.startswith("sales__"):
        return "sales/" + key[7:]
    return key.replace("__", "/")


def resolve_dashboard_cache_gist_id(cache_dir: str = "./uploaded_cache") -> str:
    for k in ("DASHBOARD_CACHE_GIST_ID", "CACHE_GIST_ID"):
        v = (os.environ.get(k) or "").strip()
        if v:
            return v
    v = _secret_get("dashboard_cache_gist_id", "cache_gist_id", "DASHBOARD_CACHE_GIST_ID")
    if v:
        return v
    path = os.path.join(cache_dir, _GIST_ID_FILE)
    try:
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                got = (f.read() or "").strip()
                if got:
                    return got
    except OSError:
        pass
    return ""


def remember_dashboard_cache_gist_id(gist_id: str, cache_dir: str = "./uploaded_cache") -> None:
    if not gist_id:
        return
    try:
        os.makedirs(cache_dir, exist_ok=True)
        with open(os.path.join(cache_dir, _GIST_ID_FILE), "w", encoding="utf-8") as f:
            f.write(gist_id.strip() + "\n")
    except OSError:
        pass


def cache_remote_configured() -> bool:
    return bool(resolve_github_token())


def _list_local_cache_files(cache_dir: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for rel in _STATIC_REL:
        p = os.path.join(cache_dir, rel)
        if os.path.isfile(p):
            out[rel] = p
    sales_dir = os.path.join(cache_dir, "sales")
    if os.path.isdir(sales_dir):
        try:
            for name in sorted(os.listdir(sales_dir)):
                if _SALES_RE.match(name):
                    out[f"sales/{name}"] = os.path.join(sales_dir, name)
        except OSError:
            pass
    return out


def _encode_file(raw: bytes) -> str:
    return base64.b64encode(gzip.compress(raw, compresslevel=6)).decode("ascii")


def _decode_file(text: str) -> bytes:
    return gzip.decompress(base64.b64decode(text.encode("ascii"), validate=False))


def _fetch_gist(token: str, gist_id: str) -> Tuple[Optional[dict], Optional[str]]:
    try:
        r = requests.get(f"{_GIST_API}/{gist_id}", headers=_headers(token), timeout=_TIMEOUT)
        if r.status_code != 200:
            return None, f"gist GET {r.status_code}: {(r.text or '')[:200]}"
        return r.json(), None
    except Exception as e:
        return None, str(e)


def _load_manifest(files_meta: Dict[str, Any]) -> Dict[str, Any]:
    meta = files_meta.get(_MANIFEST) or {}
    content = meta.get("content") if isinstance(meta, dict) else None
    if not content:
        return {}
    try:
        return json.loads(content)
    except Exception:
        return {}


def ensure_dashboard_cache_gist(cache_dir: str = "./uploaded_cache") -> Tuple[Optional[str], Optional[str]]:
    token = resolve_github_token()
    if not token:
        return None, "github_token 없음"
    gid = resolve_dashboard_cache_gist_id(cache_dir)
    if gid:
        return gid, None
    payload = {
        "description": "office dashboard uploaded_cache sync (local ↔ streamlit cloud)",
        "public": False,
        "files": {
            _MANIFEST: {
                "content": json.dumps({"version": 1, "files": {}}, ensure_ascii=False, indent=2),
            },
            "README.md": {
                "content": "Dashboard sidebar cache sync. Do not commit gist id.\n",
            },
        },
    }
    try:
        r = requests.post(_GIST_API, headers=_headers(token), json=payload, timeout=_TIMEOUT)
        if r.status_code not in (200, 201):
            return None, f"gist 생성 실패 {r.status_code}: {(r.text or '')[:200]}"
        gid = str(r.json().get("id") or "").strip()
        if not gid:
            return None, "gist id 없음"
        remember_dashboard_cache_gist_id(gid, cache_dir)
        return gid, None
    except Exception as e:
        return None, str(e)


def push_cache_file(
    rel: str,
    local_path: str,
    cache_dir: str = "./uploaded_cache",
) -> Tuple[Optional[str], Optional[str]]:
    """단일 캐시 파일 Gist 업로드. (gist_id, error)"""
    if not rel or not local_path or not os.path.isfile(local_path):
        return None, "파일 없음"
    token = resolve_github_token()
    if not token:
        return None, "github_token 없음"
    gid, err = ensure_dashboard_cache_gist(cache_dir)
    if not gid:
        return None, err or "gist 없음"
    try:
        with open(local_path, "rb") as f:
            raw = f.read()
    except OSError as e:
        return None, str(e)
    sha = _sha256_bytes(raw)
    try:
        mtime = os.path.getmtime(local_path)
    except OSError:
        mtime = time.time()

    gist, gerr = _fetch_gist(token, gid)
    if gist is None:
        return None, gerr
    files_meta = gist.get("files") or {}
    manifest = _load_manifest(files_meta)
    files_map = dict(manifest.get("files") or {})
    files_map[rel] = {"sha256": sha, "mtime": float(mtime), "size": len(raw)}
    manifest = {"version": 1, "files": files_map}
    gkey = _gist_key(rel) + _B64_SUFFIX
    body = {
        "files": {
            gkey: {"content": _encode_file(raw)},
            _MANIFEST: {"content": json.dumps(manifest, ensure_ascii=False, indent=2)},
        }
    }
    try:
        r = requests.patch(f"{_GIST_API}/{gid}", headers=_headers(token), json=body, timeout=_TIMEOUT)
        if r.status_code not in (200, 201):
            return None, f"gist PATCH {r.status_code}: {(r.text or '')[:200]}"
        remember_dashboard_cache_gist_id(gid, cache_dir)
        return gid, None
    except Exception as e:
        return None, str(e)


def _pull_one(files_meta: dict, rel: str, dest: str) -> bool:
    gkey = _gist_key(rel) + _B64_SUFFIX
    meta = files_meta.get(gkey) or {}
    content = meta.get("content") if isinstance(meta, dict) else None
    if not content:
        return False
    try:
        raw = _decode_file(content)
    except Exception:
        return False
    parent = os.path.dirname(dest)
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp = dest + ".pulling"
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


def sync_cache_remote(
    cache_dir: str = "./uploaded_cache",
    *,
    force: bool = False,
    prefer_remote: bool = False,
    force_pull: bool = False,
) -> dict:
    """Gist ↔ uploaded_cache 동기화 (사이드바 CSV·매출).

    prefer_remote=True: Cloud 재부팅 시 Gist(최신) 우선.
    force_pull=True: prefer_remote 시 sha 일치해도 Gist에서 다시 받음(수동 새로고침).
    """
    token = resolve_github_token()
    if not token:
        return {
            "ok": True,
            "skipped": True,
            "copied": [],
            "error": None,
            "note": "github_token 없음",
        }
    gid, err = ensure_dashboard_cache_gist(cache_dir)
    if not gid:
        return {"ok": False, "skipped": False, "copied": [], "error": err or "gist 없음"}

    gist, gerr = _fetch_gist(token, gid)
    if gist is None:
        return {"ok": False, "skipped": False, "copied": [], "error": gerr, "gist_id": gid}

    os.makedirs(cache_dir, exist_ok=True)
    os.makedirs(os.path.join(cache_dir, "sales"), exist_ok=True)

    local_files = _list_local_cache_files(cache_dir)

    # 맥 force=True: 로컬 캐시 전체를 Gist에 덮어씀 (Drive 동기화·↑ Gist 버튼)
    if force and not prefer_remote:
        copied: List[str] = []
        for rel, loc in sorted(local_files.items()):
            if not os.path.isfile(loc):
                continue
            _, perr = push_cache_file(rel, loc, cache_dir)
            if not perr:
                copied.append(f"→Gist!:{rel}")
        remember_dashboard_cache_gist_id(gid, cache_dir)
        push_count = len(copied)
        return {
            "ok": True,
            "skipped": False,
            "copied": copied,
            "pull_count": 0,
            "push_count": push_count,
            "remote_count": push_count,
            "gist_id": gid,
            "source": f"gist:{gid}",
        }

    files_meta = gist.get("files") or {}
    manifest = _load_manifest(files_meta)
    remote_files: Dict[str, Any] = dict(manifest.get("files") or {})

    copied: List[str] = []
    names: Set[str] = set(local_files.keys()) | set(remote_files.keys())

    for rel in sorted(names):
        loc = local_files.get(rel) or os.path.join(cache_dir, rel)
        loc_ok = os.path.isfile(loc)
        rem = remote_files.get(rel)
        gkey = _gist_key(rel) + _B64_SUFFIX
        rem_ok = rem is not None and gkey in files_meta

        if loc_ok and not rem_ok:
            if not prefer_remote:
                _, perr = push_cache_file(rel, loc, cache_dir)
                if not perr:
                    copied.append(f"→Gist:{rel}")
            continue

        if rem_ok and not loc_ok:
            if _pull_one(files_meta, rel, loc):
                copied.append(f"←Gist:{rel}")
                try:
                    rm = float((rem or {}).get("mtime") or 0)
                    if rm > 0:
                        os.utime(loc, (rm, rm))
                except OSError:
                    pass
            continue

        if not (loc_ok and rem_ok):
            continue

        local_sha = _sha256_file(loc) or ""
        remote_sha = str((rem or {}).get("sha256") or "")
        if remote_sha and local_sha == remote_sha and not (prefer_remote and force_pull):
            continue

        try:
            lm = os.path.getmtime(loc)
        except OSError:
            lm = 0.0
        rm = float((rem or {}).get("mtime") or 0)

        if prefer_remote and rem_ok:
            if force_pull or (not loc_ok) or (local_sha != remote_sha and remote_sha):
                if _pull_one(files_meta, rel, loc):
                    copied.append(f"←Gist:{rel}")
                    try:
                        rm = float((rem or {}).get("mtime") or 0)
                        if rm > 0:
                            os.utime(loc, (rm, rm))
                    except OSError:
                        pass
            continue

        if force and loc_ok:
            _, perr = push_cache_file(rel, loc, cache_dir)
            if not perr:
                copied.append(f"→Gist!:{rel}")
            continue

        if lm > rm + 1.0:
            _, perr = push_cache_file(rel, loc, cache_dir)
            if not perr:
                copied.append(f"→Gist:{rel}")
        elif rm > lm + 1.0:
            if _pull_one(files_meta, rel, loc):
                copied.append(f"←Gist:{rel}")
                try:
                    if rm > 0:
                        os.utime(loc, (rm, rm))
                except OSError:
                    pass

    remember_dashboard_cache_gist_id(gid, cache_dir)
    remote_count = len(remote_files)
    pull_count = len([x for x in copied if str(x).startswith("←Gist:")])
    push_count = len([x for x in copied if str(x).startswith("→Gist")])
    return {
        "ok": True,
        "skipped": False,
        "copied": copied,
        "pull_count": pull_count,
        "push_count": push_count,
        "remote_count": remote_count,
        "gist_id": gid,
        "source": f"gist:{gid}",
    }
