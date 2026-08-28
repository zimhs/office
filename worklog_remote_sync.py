"""로컬 Streamlit ↔ Streamlit Cloud 업무일지 양방향 동기화 (GitHub Gist).

Drive 마운트는 맥 로컬에만 있으므로, 클라우드와 맞추려면 secrets의
github_token + worklog_gist_id 로 Secret Gist에 일자 xlsx를 올리고 받습니다.

설정 (.streamlit/secrets.toml 또는 환경변수):
  github_token / GITHUB_TOKEN / WORKLOG_GITHUB_TOKEN
  worklog_gist_id / WORKLOG_GIST_ID  (없으면 첫 저장 시 자동 생성)

주의: GitHub Secret Gist는 Discover에 안 보일 뿐, URL(id) 알면 읽을 수 있습니다.
      gist id·token 은 커밋하지 마세요.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import subprocess
import time
from typing import Any, Dict, List, Optional, Tuple

import requests

_DAY_RE = re.compile(r"^20\d{2}-\d{2}-\d{2}\.xlsx$")
_GIST_API = "https://api.github.com/gists"
_MANIFEST = "_worklog_manifest.json"
_DELETED_MANIFEST = "_worklog_deleted.json"
_B64_SUFFIX = ".b64"
_GIST_ID_FILE = ".worklog_gist_id"
_TIMEOUT = 45


def _is_day_file(name: str) -> bool:
    return bool(name) and bool(_DAY_RE.match(name))


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: str) -> Optional[str]:
    try:
        with open(path, "rb") as f:
            return _sha256_bytes(f.read())
    except OSError:
        return None


def _secret_get(*keys: str) -> str:
    try:
        import streamlit as st

        secrets = getattr(st, "secrets", None)
        if secrets is None:
            return ""
        for k in keys:
            try:
                if hasattr(secrets, "get"):
                    v = secrets.get(k)
                else:
                    v = secrets[k]
            except Exception:
                v = None
            if v is not None and str(v).strip():
                return str(v).strip()
    except Exception:
        pass
    return ""


def resolve_github_token() -> str:
    """Gist 동기화용 토큰. ghs_(GitHub App) 는 gist 권한이 없는 경우가 많아 제외."""
    candidates = []
    for k in ("WORKLOG_GITHUB_TOKEN", "GITHUB_TOKEN", "GH_TOKEN"):
        v = (os.environ.get(k) or "").strip()
        if v:
            candidates.append(v)
    v = _secret_get("worklog_github_token", "github_token", "GITHUB_TOKEN")
    if v:
        candidates.append(v)
    try:
        r = subprocess.run(
            ["gh", "auth", "token"],
            capture_output=True,
            text=True,
            timeout=8,
        )
        if r.returncode == 0 and (r.stdout or "").strip():
            candidates.append((r.stdout or "").strip())
    except Exception:
        pass
    for tok in candidates:
        # GitHub App installation tokens (ghs_) usually cannot manage gists
        if tok.startswith("ghs_"):
            continue
        return tok
    return ""


def _gist_id_path(local_dir: str) -> str:
    return os.path.join(local_dir, _GIST_ID_FILE)


def resolve_gist_id(local_dir: str = "./uploaded_cache/worklog") -> str:
    for k in ("WORKLOG_GIST_ID", "WORKLOG_GIST"):
        v = (os.environ.get(k) or "").strip()
        if v:
            return v
    v = _secret_get("worklog_gist_id", "WORKLOG_GIST_ID")
    if v:
        return v
    path = _gist_id_path(local_dir)
    try:
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                got = (f.read() or "").strip()
                if got:
                    return got
    except OSError:
        pass
    # 맥 Drive worklog 에도 id 를 두면 다른 맥 로컬이 같은 gist 를 씀
    try:
        from drive_autoload import resolve_drive_worklog_dir

        drv = resolve_drive_worklog_dir()
        if drv:
            p = os.path.join(drv, _GIST_ID_FILE)
            if os.path.isfile(p):
                with open(p, "r", encoding="utf-8") as f:
                    got = (f.read() or "").strip()
                    if got:
                        return got
    except Exception:
        pass
    return ""


def remember_gist_id(gist_id: str, local_dir: str = "./uploaded_cache/worklog") -> None:
    if not gist_id:
        return
    try:
        os.makedirs(local_dir, exist_ok=True)
        with open(_gist_id_path(local_dir), "w", encoding="utf-8") as f:
            f.write(gist_id.strip() + "\n")
    except OSError:
        pass
    try:
        from drive_autoload import resolve_drive_worklog_dir

        drv = resolve_drive_worklog_dir()
        if drv:
            with open(os.path.join(drv, _GIST_ID_FILE), "w", encoding="utf-8") as f:
                f.write(gist_id.strip() + "\n")
    except Exception:
        pass


def remote_sync_configured(local_dir: str = "./uploaded_cache/worklog") -> bool:
    return bool(resolve_github_token())


def remote_sync_ready(local_dir: str = "./uploaded_cache/worklog") -> bool:
    return bool(resolve_github_token() and resolve_gist_id(local_dir))


def _headers(token: str) -> Dict[str, str]:
    return {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _encode_file(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _decode_file(text: str) -> bytes:
    return base64.b64decode(text.encode("ascii"), validate=False)


def _load_manifest(files: Dict[str, Any]) -> Dict[str, Any]:
    meta = files.get(_MANIFEST) or {}
    content = meta.get("content") if isinstance(meta, dict) else None
    if not content:
        return {}
    try:
        return json.loads(content)
    except Exception:
        return {}


def _deleted_manifest_path(local_dir: str) -> str:
    return os.path.join(local_dir, _DELETED_MANIFEST)


def _load_deleted_days(local_dir: str) -> Dict[str, float]:
    path = _deleted_manifest_path(local_dir)
    try:
        if os.path.isfile(path):
            with open(path, encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict):
                out: Dict[str, float] = {}
                for k, v in raw.items():
                    iso = str(k).replace(".xlsx", "")
                    if _is_day_file(f"{iso}.xlsx"):
                        out[iso] = float(v)
                return out
    except Exception:
        pass
    return {}


def _save_deleted_days(local_dir: str, deleted: Dict[str, float]) -> None:
    try:
        os.makedirs(local_dir, exist_ok=True)
        with open(_deleted_manifest_path(local_dir), "w", encoding="utf-8") as f:
            json.dump(deleted, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def mark_worklog_day_deleted(iso: str, local_dir: str = "./uploaded_cache/worklog") -> None:
    """로컬 삭제 표시 — Gist/Drive 동기화가 구 파일을 되살리지 않게."""
    if not iso or not _is_day_file(f"{iso}.xlsx"):
        return
    deleted = _load_deleted_days(local_dir)
    deleted[iso] = time.time()
    _save_deleted_days(local_dir, deleted)


def clear_worklog_day_deleted(iso: str, local_dir: str = "./uploaded_cache/worklog") -> None:
    deleted = _load_deleted_days(local_dir)
    if iso in deleted:
        del deleted[iso]
        _save_deleted_days(local_dir, deleted)


def is_worklog_day_deleted(name: str, local_dir: str = "./uploaded_cache/worklog") -> bool:
    iso = name.replace(".xlsx", "") if name.endswith(".xlsx") else name
    return iso in _load_deleted_days(local_dir)


def _fetch_gist(token: str, gist_id: str) -> Tuple[Optional[dict], Optional[str]]:
    try:
        r = requests.get(
            f"{_GIST_API}/{gist_id}",
            headers=_headers(token),
            timeout=_TIMEOUT,
        )
        if r.status_code != 200:
            return None, f"gist GET {r.status_code}: {(r.text or '')[:180]}"
        return r.json(), None
    except Exception as e:
        return None, str(e)


def ensure_worklog_gist(local_dir: str = "./uploaded_cache/worklog") -> Tuple[Optional[str], Optional[str]]:
    """gist id 반환. 없으면 생성. (id, error)"""
    token = resolve_github_token()
    if not token:
        return None, "github_token 없음"
    gid = resolve_gist_id(local_dir)
    if gid:
        return gid, None
    payload = {
        "description": "office worklog sync (local ↔ streamlit cloud)",
        "public": False,
        "files": {
            _MANIFEST: {
                "content": json.dumps({"version": 1, "files": {}}, ensure_ascii=False, indent=2)
            },
            "README.md": {
                "content": (
                    "Streamlit 업무일지 양방향 동기화용 Secret Gist.\n"
                    "이 파일·gist id 를 공개 저장소에 커밋하지 마세요.\n"
                )
            },
        },
    }
    try:
        r = requests.post(
            _GIST_API,
            headers=_headers(token),
            json=payload,
            timeout=_TIMEOUT,
        )
        if r.status_code not in (200, 201):
            return None, f"gist 생성 실패 {r.status_code}: {(r.text or '')[:180]}"
        data = r.json()
        gid = str(data.get("id") or "").strip()
        if not gid:
            return None, "gist id 없음"
        remember_gist_id(gid, local_dir)
        return gid, None
    except Exception as e:
        return None, str(e)


_GIST_DAYS_CACHE: Dict[str, Any] = {"ts": 0.0, "gid": "", "names": set()}


def invalidate_gist_days_cache() -> None:
    _GIST_DAYS_CACHE["ts"] = 0.0
    _GIST_DAYS_CACHE["gid"] = ""
    _GIST_DAYS_CACHE["names"] = set()


def list_gist_day_names(
    local_dir: str = "./uploaded_cache/worklog",
    *,
    max_age: float = 120.0,
) -> set[str]:
    """Gist manifest 일자 파일명 집합 (캐시, API 호출 최소화)."""
    token = resolve_github_token()
    gid = resolve_gist_id(local_dir)
    if not token or not gid:
        return set()
    now = time.time()
    cache = _GIST_DAYS_CACHE
    if (
        cache.get("gid") == gid
        and (now - float(cache.get("ts") or 0)) < max_age
        and isinstance(cache.get("names"), set)
    ):
        return set(cache["names"])
    gist, _ = _fetch_gist(token, gid)
    if gist is None:
        return set()
    files_meta = gist.get("files") or {}
    manifest = _load_manifest(files_meta)
    names: set[str] = set(manifest.get("files") or {})
    for fname in files_meta:
        if isinstance(fname, str) and fname.endswith(_B64_SUFFIX):
            day = fname[: -len(_B64_SUFFIX)]
            if _is_day_file(day):
                names.add(day)
    cache["ts"] = now
    cache["gid"] = gid
    cache["names"] = names
    return set(names)


def worklog_date_exists_on_cloud(
    d,
    local_dir: str = "./uploaded_cache/worklog",
) -> bool:
    """Gist manifest에 해당 일자 xlsx가 있는지."""
    from datetime import date as _date

    if isinstance(d, _date):
        name = f"{d.isoformat()}.xlsx"
    else:
        name = str(d or "")
        if not name.endswith(".xlsx"):
            name = f"{name}.xlsx"
    if not _is_day_file(name):
        return False
    return name in list_gist_day_names(local_dir)


def pull_worklog_day_from_remote(
    day: date | str,
    local_dir: str = "./uploaded_cache/worklog",
) -> bool:
    """Gist에 있는 일자 xlsx를 로컬 캐시로 받음. (성공 True)"""
    from datetime import date as _date

    if isinstance(day, _date):
        name = f"{day.isoformat()}.xlsx"
    else:
        name = str(day or "")
        if not name.endswith(".xlsx"):
            name = f"{name}.xlsx"
    if not _is_day_file(name):
        return False
    if is_worklog_day_deleted(name.replace(".xlsx", ""), local_dir):
        return False
    loc = os.path.join(local_dir, name)
    if os.path.isfile(loc):
        return False
    token = resolve_github_token()
    gid = resolve_gist_id(local_dir)
    if not token or not gid:
        return False
    gist, _ = _fetch_gist(token, gid)
    if gist is None:
        return False
    if not _pull_one(gist.get("files") or {}, name, loc):
        return False
    invalidate_gist_days_cache()
    return True


def push_worklog_day_remote(
    local_path: str,
    local_dir: str = "./uploaded_cache/worklog",
    *,
    force: bool = False,
) -> Tuple[Optional[str], Optional[str]]:
    """저장 직후 일자 파일을 Gist에 올림. (gist_id, error)

    force=False 이고 Gist에 이미 해당 일자가 있으면 덮어쓰지 않음 (선입력 우선).
    """
    if not local_path or not os.path.isfile(local_path):
        return None, "일지 파일 없음"
    name = os.path.basename(local_path)
    if not _is_day_file(name):
        return None, "일자 xlsx 아님"
    token = resolve_github_token()
    if not token:
        return None, "github_token 없음"
    gid, err = ensure_worklog_gist(local_dir)
    if not gid:
        return None, err or "Gist 생성/조회 실패"
    if not force and name in list_gist_day_names(local_dir):
        return gid, "duplicate_date"
    try:
        with open(local_path, "rb") as f:
            raw = f.read()
    except OSError as e:
        return None, f"파일 읽기 실패: {e}"
    sha = _sha256_bytes(raw)
    try:
        mtime = os.path.getmtime(local_path)
    except OSError:
        mtime = time.time()

    gist, gerr = _fetch_gist(token, gid)
    if gist is None:
        return None, gerr or "Gist fetch 실패"
    files_meta = gist.get("files") or {}
    manifest = _load_manifest(files_meta)
    files_map = dict(manifest.get("files") or {})
    files_map[name] = {"sha256": sha, "mtime": float(mtime), "updated_at": time.time()}
    manifest = {"version": 1, "files": files_map}

    body = {
        "files": {
            f"{name}{_B64_SUFFIX}": {"content": _encode_file(raw)},
            _MANIFEST: {
                "content": json.dumps(manifest, ensure_ascii=False, indent=2),
            },
        }
    }
    try:
        r = requests.patch(
            f"{_GIST_API}/{gid}",
            headers=_headers(token),
            json=body,
            timeout=_TIMEOUT,
        )
        if r.status_code not in (200, 201):
            return None, f"Gist 업로드 실패 {r.status_code}: {(r.text or '')[:200]}"
        remember_gist_id(gid, local_dir)
        invalidate_gist_days_cache()
        return gid, None
    except Exception as e:
        return None, str(e)


def cloud_sync_status(local_dir: str = "./uploaded_cache/worklog") -> dict:
    """UI용 Cloud 연동 상태."""
    tok = resolve_github_token()
    gid = resolve_gist_id(local_dir)
    gid_file = _gist_id_path(local_dir)
    gid_from_file = ""
    try:
        if os.path.isfile(gid_file):
            gid_from_file = (open(gid_file, encoding="utf-8").read() or "").strip()
    except OSError:
        pass
    effective_gid = gid or gid_from_file
    return {
        "token": bool(tok),
        "gist_id": effective_gid,
        "gist_file": gid_file if os.path.isfile(gid_file) else "",
        "ready": bool(tok and effective_gid),
    }


def sync_worklog_remote(
    local_dir: str = "./uploaded_cache/worklog",
    *,
    force: bool = False,
    prefer_remote: bool = False,
    force_pull: bool = False,
) -> dict:
    """Gist ↔ 로컬 uploaded_cache/worklog 동기화.

    - 로컬만 있음 → push
    - 원격만 있음 → pull
    - 둘 다 있고 sha 다름 → conflicts (자동 덮어쓰기 안 함)
    - prefer_remote=True (Cloud): Gist 최신 우선 pull
    - force_pull=True: sha 일치해도 Gist에서 다시 받음
    """
    token = resolve_github_token()
    if not token:
        return {
            "ok": True,
            "skipped": True,
            "copied": [],
            "conflicts": [],
            "note": "github_token 없음(로컬 Drive만 가능)",
        }
    gid, err = ensure_worklog_gist(local_dir)
    if not gid:
        return {
            "ok": False,
            "skipped": False,
            "copied": [],
            "conflicts": [],
            "error": err or "gist 없음",
        }

    gist, gerr = _fetch_gist(token, gid)
    if gist is None:
        return {
            "ok": False,
            "skipped": False,
            "copied": [],
            "conflicts": [],
            "error": gerr or "gist fetch 실패",
            "gist_id": gid,
        }

    try:
        os.makedirs(local_dir, exist_ok=True)
    except OSError:
        pass

    files_meta = gist.get("files") or {}
    manifest = _load_manifest(files_meta)
    remote_files: Dict[str, Any] = dict(manifest.get("files") or {})

    # manifest 없이 b64만 있는 경우도 인식
    for fname, meta in files_meta.items():
        if not isinstance(fname, str) or not fname.endswith(_B64_SUFFIX):
            continue
        day = fname[: -len(_B64_SUFFIX)]
        if _is_day_file(day) and day not in remote_files:
            remote_files[day] = {"sha256": "", "mtime": 0.0}

    local_names = set()
    try:
        local_names.update(n for n in os.listdir(local_dir) if _is_day_file(n))
    except OSError:
        pass

    copied: List[str] = []
    conflicts: List[str] = []
    names = set(local_names) | set(remote_files.keys())

    for name in sorted(names):
        loc = os.path.join(local_dir, name)
        loc_ok = os.path.isfile(loc)
        rem = remote_files.get(name)
        rem_ok = f"{name}{_B64_SUFFIX}" in files_meta

        if loc_ok and not rem_ok:
            gid, perr = push_worklog_day_remote(loc, local_dir)
            if gid:
                copied.append(f"→Cloud:{name}")
            continue

        if rem_ok and not loc_ok:
            # 로컬에서 삭제한 날짜는 pull 금지 — Gist에서도 제거
            if is_worklog_day_deleted(name, local_dir):
                if delete_worklog_day_remote(name, local_dir)[0]:
                    copied.append(f"☁삭제:{name}")
                continue
            if _pull_one(files_meta, name, loc):
                copied.append(f"←Cloud:{name}")
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
        if remote_sha and local_sha and local_sha == remote_sha:
            if not (prefer_remote and force_pull):
                continue
        # sha 없으면 내용 비교
        if not remote_sha:
            try:
                b64 = (files_meta.get(f"{name}{_B64_SUFFIX}") or {}).get("content") or ""
                remote_sha = _sha256_bytes(_decode_file(b64)) if b64 else ""
            except Exception:
                remote_sha = ""
        if remote_sha and local_sha == remote_sha and not (prefer_remote and force_pull):
            continue

        if prefer_remote and rem_ok:
            should_pull = (
                force_pull
                or not loc_ok
                or not remote_sha
                or local_sha != remote_sha
            )
            if should_pull and not is_worklog_day_deleted(name, local_dir):
                if _pull_one(files_meta, name, loc):
                    copied.append(f"←Cloud:{name}")
                    try:
                        rm = float((rem or {}).get("mtime") or 0)
                        if rm > 0:
                            os.utime(loc, (rm, rm))
                    except OSError:
                        pass
            continue

        if force:
            # force: 로컬 우선 푸시
            gid, _ = push_worklog_day_remote(loc, local_dir)
            if gid:
                copied.append(f"→Cloud!:{name}")
            continue

        try:
            lm = os.path.getmtime(loc)
        except OSError:
            lm = 0.0
        rm = float((rem or {}).get("mtime") or 0)
        # 1초 이상 차이나면 충돌로 안내 (자동 병합 없음)
        if abs(lm - rm) < 1.0 and local_sha and remote_sha and local_sha != remote_sha:
            conflicts.append(name)
        elif local_sha != remote_sha:
            conflicts.append(name)

    remember_gist_id(gid, local_dir)
    return {
        "ok": True,
        "skipped": False,
        "copied": copied,
        "conflicts": conflicts,
        "gist_id": gid,
        "source": f"gist:{gid}",
    }


def delete_worklog_day_remote(
    day: str | date,
    local_dir: str = "./uploaded_cache/worklog",
) -> Tuple[bool, Optional[str]]:
    """Gist에서 일자 xlsx 제거. (ok, error)"""
    if isinstance(day, date):
        name = f"{day.isoformat()}.xlsx"
    else:
        name = str(day or "")
        if not name.endswith(".xlsx"):
            name = f"{name}.xlsx"
    if not _is_day_file(name):
        return False, "일자 xlsx 아님"
    token = resolve_github_token()
    if not token:
        return False, "github_token 없음"
    gid = resolve_gist_id(local_dir)
    if not gid:
        return False, "gist 없음"
    gist, gerr = _fetch_gist(token, gid)
    if gist is None:
        return False, gerr or "Gist fetch 실패"
    files_meta = gist.get("files") or {}
    b64_name = f"{name}{_B64_SUFFIX}"
    if b64_name not in files_meta and name not in (dict(_load_manifest(files_meta).get("files") or {})):
        clear_worklog_day_deleted(name.replace(".xlsx", ""), local_dir)
        return True, None
    manifest = _load_manifest(files_meta)
    files_map = dict(manifest.get("files") or {})
    files_map.pop(name, None)
    manifest = {"version": 1, "files": files_map}
    patch_files: Dict[str, Any] = {
        _MANIFEST: {"content": json.dumps(manifest, ensure_ascii=False, indent=2)},
    }
    if b64_name in files_meta:
        patch_files[b64_name] = None
    try:
        r = requests.patch(
            f"{_GIST_API}/{gid}",
            headers=_headers(token),
            json={"files": patch_files},
            timeout=_TIMEOUT,
        )
        if r.status_code not in (200, 201):
            return False, f"Gist 삭제 실패 {r.status_code}: {(r.text or '')[:200]}"
        clear_worklog_day_deleted(name.replace(".xlsx", ""), local_dir)
        return True, None
    except Exception as e:
        return False, str(e)


def _pull_one(files_meta: dict, name: str, dest: str) -> bool:
    meta = files_meta.get(f"{name}{_B64_SUFFIX}") or {}
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
    tmp = dest + ".downloading"
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


def resolve_remote_conflict(
    name: str,
    local_dir: str = "./uploaded_cache/worklog",
    *,
    prefer: str = "local",
) -> Optional[str]:
    """충돌 일자: prefer=local 이면 push, drive/cloud 이면 pull."""
    if not _is_day_file(name):
        return None
    loc = os.path.join(local_dir, name)
    if prefer == "local":
        if os.path.isfile(loc):
            gid, _ = push_worklog_day_remote(loc, local_dir)
            if gid:
                return loc
        return None
    token = resolve_github_token()
    gid = resolve_gist_id(local_dir)
    if not token or not gid:
        return None
    gist, _ = _fetch_gist(token, gid)
    if not gist:
        return None
    if _pull_one(gist.get("files") or {}, name, loc):
        return loc
    return None
