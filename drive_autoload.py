"""Google Drive「dashboard 복사본」→ uploaded_cache 자동 로드.

맥: Drive 마운트에서 최신 파일을 캐시로 복사 (프로세스당 1회).
클라우드: Drive 경로가 없으면 no-op — 배포에 포함된 uploaded_cache 시드를 사용.
API 키·일지 원본 등은 복사하지 않음.
"""
from __future__ import annotations

import os
import re
import shutil
from typing import Dict, List, Optional, Tuple

try:
    from worklog_remote_sync import _sha256_file
except Exception:
    _sha256_file = None  # type: ignore

DRIVE_COPY_NAME = "dashboard 복사본"

# Drive 루트 파일명 → (캐시 상대경로, 선택적 name.txt 내용)
_CACHE_MAP: Tuple[Tuple[str, str, Optional[str]], ...] = (
    ("주소.csv", "address.csv", None),
    ("업체대분류.csv", "industry.csv", None),
    ("채권.csv", "debt.csv", None),
    ("탱크.csv", "tank_cache.dat", "탱크.csv"),
    ("기화기.csv", "vaporizer_cache.dat", "기화기.csv"),
    ("통합탱크재고.csv", "integrated_cache.dat", "통합탱크재고.csv"),
)

_SALES_NAME_RE = re.compile(r"^20\d{2}(\d{2})?\.csv$", re.IGNORECASE)

# 연간+월별 동시 존재 시 연간(YYYY.csv) 제외용 — app.py dedupe와 맞춤
_SKIP_ANNUAL_IF_MONTHLY = True

_DONE = False


def resolve_drive_dashboard_copy() -> Optional[str]:
    """Google Drive「내 드라이브/dashboard 복사본/uproad」절대경로."""
    home = os.path.expanduser("~")
    cloud = os.path.join(home, "Library", "CloudStorage")
    if not os.path.isdir(cloud):
        return None
    try:
        names = os.listdir(cloud)
    except OSError:
        return None
    for name in names:
        if not name.startswith("GoogleDrive"):
            continue
        root = os.path.join(cloud, name)
        for drive_label in ("내 드라이브", "My Drive"):
            base = os.path.join(root, drive_label, DRIVE_COPY_NAME)
            if os.path.isdir(base):
                # 💡 [핵심 수정] dashboard 복사본 폴더 안에 uproad 폴더를 만들고 최종 목적지로 지정!
                uproad_dir = os.path.join(base, "uproad")
                try:
                    os.makedirs(uproad_dir, exist_ok=True)
                    return uproad_dir
                except OSError:
                    return base
    return None


def _atomic_copy(src: str, dst: str) -> bool:
    parent = os.path.dirname(dst)
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp = dst + ".uploading"
    try:
        shutil.copy2(src, tmp)
        os.replace(tmp, dst)
        return True
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        try:
            shutil.copy2(src, dst)
            return True
        except Exception:
            return False


def _should_replace(src: str, dst: str) -> bool:
    if not os.path.isfile(src):
        return False
    if not os.path.isfile(dst):
        return True
    try:
        ss, sd = os.stat(src), os.stat(dst)
        if ss.st_size != sd.st_size:
            return True
        # Drive가 같거나 더 최신이면 갱신
        return ss.st_mtime >= sd.st_mtime - 1.0
    except OSError:
        return True


def _cache_differs_from_drive(src: str, dst: str) -> bool:
    """캐시→Drive 푸시 시 변경 여부 (크기·시간·내용 해시)."""
    if not os.path.isfile(src):
        return False
    if not os.path.isfile(dst):
        return True
    try:
        if os.path.getsize(src) != os.path.getsize(dst):
            return True
        if abs(os.path.getmtime(src) - os.path.getmtime(dst)) >= 1.0:
            return True
    except OSError:
        return True
    if _sha256_file is not None:
        try:
            local_sha = _sha256_file(src) or ""
            drive_sha = _sha256_file(dst) or ""
            if local_sha and drive_sha and local_sha != drive_sha:
                return True
        except Exception:
            return True
    return False


def _list_drive_sales(drive_root: str) -> List[str]:
    names = []
    try:
        for n in os.listdir(drive_root):
            if _SALES_NAME_RE.match(n) and os.path.isfile(os.path.join(drive_root, n)):
                names.append(n)
    except OSError:
        return []
    if not _SKIP_ANNUAL_IF_MONTHLY:
        return sorted(names)
    by_year: Dict[str, List[str]] = {}
    for n in names:
        m = re.match(r"^(20\d{2})(\d{2})?\.csv$", n, re.I)
        if not m:
            continue
        by_year.setdefault(m.group(1), []).append(n)
    keep = set()
    for y, grp in by_year.items():
        monthlies = [n for n in grp if re.match(rf"^{y}\d{{2}}\.csv$", n, re.I)]
        annuals = [n for n in grp if re.match(rf"^{y}\.csv$", n, re.I)]
        if monthlies:
            keep.update(monthlies)
        else:
            keep.update(annuals)
    return sorted(keep)


def sync_drive_copy_into_cache(
    cache_dir: str = "./uploaded_cache",
    *,
    force_refresh: bool = False,
) -> dict:
    """Drive 복사본 → uploaded_cache. 프로세스당 1회 (force_refresh 시 항상 Drive 우선).

    Returns:
        {ok, source, copied: [..], skipped: bool, error?}
    """
    global _DONE
    if _DONE and not force_refresh:
        return {"ok": True, "skipped": True, "copied": [], "source": None}

    drive_root = resolve_drive_dashboard_copy()
    if not drive_root:
        return {
            "ok": True,
            "skipped": True,
            "copied": [],
            "source": None,
            "note": "Drive 복사본 없음(클라우드·시드 캐시 사용)",
        }

    _DONE = True
    copied: List[str] = []
    try:
        os.makedirs(cache_dir, exist_ok=True)
        sales_dir = os.path.join(cache_dir, "sales")
        os.makedirs(sales_dir, exist_ok=True)

        for src_name, rel, name_txt in _CACHE_MAP:
            src = os.path.join(drive_root, src_name)
            dst = os.path.join(cache_dir, rel)
            if not force_refresh and not _should_replace(src, dst):
                continue
            if not os.path.isfile(src):
                continue
            if _atomic_copy(src, dst):
                copied.append(src_name)
                if name_txt:
                    try:
                        with open(dst + "_name.txt", "w", encoding="utf-8") as f:
                            f.write(name_txt)
                    except Exception:
                        pass

        wanted_sales = _list_drive_sales(drive_root)
        # 매출: Drive 목록으로 캐시를 맞춘다 (옛 2026.csv 잔존 방지)
        if wanted_sales:
            wanted_set = set(wanted_sales)
            try:
                for existing in os.listdir(sales_dir):
                    if existing.endswith(".csv") and existing not in wanted_set:
                        try:
                            os.remove(os.path.join(sales_dir, existing))
                            copied.append(f"-sales/{existing}")
                        except Exception:
                            pass
            except OSError:
                pass
            for sn in wanted_sales:
                src = os.path.join(drive_root, sn)
                dst = os.path.join(sales_dir, sn)
                if force_refresh or _should_replace(src, dst):
                    if _atomic_copy(src, dst):
                        copied.append(f"sales/{sn}")

        # worklog 하위 폴더
        wl_drive = os.path.join(drive_root, "worklog")
        if os.path.isdir(wl_drive):
            wl_local = os.path.join(cache_dir, "worklog")
            os.makedirs(wl_local, exist_ok=True)
            try:
                for wname in os.listdir(wl_drive):
                    if not _is_worklog_day_file(wname) and wname != "template.xlsx":
                        continue
                    src = os.path.join(wl_drive, wname)
                    dst = os.path.join(wl_local, wname)
                    if force_refresh or _should_replace(src, dst):
                        if _atomic_copy(src, dst):
                            copied.append(f"worklog/{wname}")
            except OSError:
                pass

        return {
            "ok": True,
            "skipped": False,
            "copied": copied,
            "source": drive_root,
        }
    except Exception as e:
        return {
            "ok": False,
            "skipped": False,
            "copied": copied,
            "source": drive_root,
            "error": str(e),
        }


def sync_dashboard_copy_on_boot(
    cache_dir: str = "./uploaded_cache",
    *,
    force_refresh: bool = True,
) -> dict:
    """재시작·재부팅 시 dashboard 복사본/uproad 최신 데이터 로드.

    맥: Drive Desktop 마운트 / Cloud: Drive API (drive_remote_fetch).
    """
    local_root = resolve_drive_dashboard_copy()
    if local_root:
        return sync_drive_copy_into_cache(cache_dir, force_refresh=force_refresh)
    try:
        from drive_remote_fetch import sync_drive_copy_from_remote

        return sync_drive_copy_from_remote(cache_dir, force_refresh=force_refresh)
    except Exception as e:
        return {
            "ok": False,
            "skipped": True,
            "copied": [],
            "source": None,
            "error": str(e),
        }


def sync_cache_to_drive_copy(cache_dir: str = "./uploaded_cache", *, force: bool = False) -> dict:
    """맥 uploaded_cache → Drive「dashboard 복사본」(업로드 반영).

    Google Drive 데스크톱이 클라우드로 올리면 아이패드 Drive 폴더에도 보임.
    Streamlit Cloud 시드는 별도 git push가 필요.

    force=True: 사이드바 수동 동기화 — 내용이 같아도 캐시 기준으로 Drive에 덮어씀.
    """
    drive_root = resolve_drive_dashboard_copy()
    if not drive_root:
        return {
            "ok": False,
            "skipped": True,
            "copied": [],
            "checked": 0,
            "source": None,
            "error": "Drive「dashboard 복사본」경로 없음",
        }

    copied: List[str] = []
    checked = 0
    try:
        sales_dir = os.path.join(cache_dir, "sales")

        for drive_name, rel, _name_txt in _CACHE_MAP:
            src = os.path.join(cache_dir, rel)
            dst = os.path.join(drive_root, drive_name)
            if not os.path.isfile(src):
                continue
            checked += 1
            if not force and not _cache_differs_from_drive(src, dst):
                continue
            if _atomic_copy(src, dst):
                copied.append(drive_name)

        if os.path.isdir(sales_dir):
            cache_sales = []
            try:
                cache_sales = [
                    n
                    for n in os.listdir(sales_dir)
                    if n.endswith(".csv") and _SALES_NAME_RE.match(n)
                ]
            except OSError:
                cache_sales = []
            cache_set = set(cache_sales)
            for sn in sorted(cache_sales):
                src = os.path.join(sales_dir, sn)
                dst = os.path.join(drive_root, sn)
                if not os.path.isfile(src):
                    continue
                checked += 1
                if not force and not _cache_differs_from_drive(src, dst):
                    continue
                if _atomic_copy(src, dst):
                    copied.append(sn)
            # Drive에만 남은 옛 매출(예: 2026.csv) 제거 — 캐시에 있는 목록만 유지
            try:
                for n in os.listdir(drive_root):
                    if not _SALES_NAME_RE.match(n):
                        continue
                    if n in cache_set:
                        continue
                    try:
                        os.remove(os.path.join(drive_root, n))
                        copied.append(f"-{n}")
                    except Exception:
                        pass
            except OSError:
                pass

        return {
            "ok": True,
            "skipped": False,
            "copied": copied,
            "checked": checked,
            "source": drive_root,
        }
    except Exception as e:
        return {
            "ok": False,
            "skipped": False,
            "copied": copied,
            "checked": checked,
            "source": drive_root,
            "error": str(e),
        }


_WORKLOG_DAY_RE = re.compile(r"^20\d{2}-\d{2}-\d{2}\.xlsx$")
_WL_SYNC_DONE = False


def resolve_drive_worklog_dir() -> Optional[str]:
    """Drive「dashboard 복사본/worklog」— 맥·아이패드 공통 업무일지 폴더."""
    root = resolve_drive_dashboard_copy()
    if not root:
        return None
    path = os.path.join(root, "worklog")
    try:
        os.makedirs(path, exist_ok=True)
        return path
    except OSError:
        return None


def _is_worklog_day_file(name: str) -> bool:
    if not name or not name.endswith(".xlsx"):
        return False
    if name == "template.xlsx" or name.startswith("_preview_") or "_인쇄" in name:
        return False
    return bool(_WORKLOG_DAY_RE.match(name))


def resolve_drive_worklog_archive_dir(year: Optional[int] = None) -> Optional[str]:
    """Drive worklog/일지[/YYYY] — 월별xlsx를 맥·클라우드 Drive로 공유."""
    root = resolve_drive_worklog_dir()
    if not root:
        return None
    path = os.path.join(root, "일지")
    if year is not None:
        path = os.path.join(path, str(int(year)))
    try:
        os.makedirs(path, exist_ok=True)
        return path
    except OSError:
        return None


def sync_worklog_bidirectional(
    local_dir: str = "./uploaded_cache/worklog",
    *,
    force: bool = False,
) -> dict:
    """로컬 uploaded_cache/worklog ↔ Drive dashboard 복사본/worklog.

    일자 파일: 한쪽만 있으면 복사. 양쪽이 다르면 conflicts에 넣고 자동 덮어쓰지 않음.
    맥: Drive 마운트로 양방향. 클라우드: Drive 경로 없으면 skipped.
    """
    global _WL_SYNC_DONE
    if _WL_SYNC_DONE and not force:
        return {
            "ok": True,
            "skipped": True,
            "copied": [],
            "conflicts": [],
            "source": None,
        }
    drive_dir = resolve_drive_worklog_dir()
    if not drive_dir:
        _WL_SYNC_DONE = True
        return {
            "ok": True,
            "skipped": True,
            "copied": [],
            "conflicts": [],
            "source": None,
            "note": "Drive worklog 없음(클라우드·로컬만)",
        }
    copied: List[str] = []
    conflicts: List[str] = []
    try:
        os.makedirs(local_dir, exist_ok=True)
        names = set()
        try:
            names.update(n for n in os.listdir(local_dir) if _is_worklog_day_file(n))
        except OSError:
            pass
        try:
            names.update(n for n in os.listdir(drive_dir) if _is_worklog_day_file(n))
        except OSError:
            pass
        for name in sorted(names):
            loc = os.path.join(local_dir, name)
            drv = os.path.join(drive_dir, name)
            loc_ok = os.path.isfile(loc)
            drv_ok = os.path.isfile(drv)
            if loc_ok and not drv_ok:
                if _atomic_copy(loc, drv):
                    copied.append(f"→Drive:{name}")
                continue
            if drv_ok and not loc_ok:
                try:
                    from worklog_remote_sync import is_worklog_day_deleted

                    if is_worklog_day_deleted(name, local_dir):
                        try:
                            os.remove(drv)
                            copied.append(f"Drive삭제:{name}")
                        except OSError:
                            pass
                        continue
                except Exception:
                    pass
                if _atomic_copy(drv, loc):
                    copied.append(f"←Drive:{name}")
                continue
            if not (loc_ok and drv_ok):
                continue
            try:
                lm, dm = os.path.getmtime(loc), os.path.getmtime(drv)
                ls, ds = os.path.getsize(loc), os.path.getsize(drv)
            except OSError:
                continue
            if abs(lm - dm) < 1.0 and ls == ds:
                continue
            # 양쪽 모두 있고 내용/시간이 다르면 자동 덮어쓰지 않음 (충돌 안내)
            conflicts.append(name)
            continue
        loc_t = os.path.join(local_dir, "template.xlsx")
        drv_t = os.path.join(drive_dir, "template.xlsx")
        if os.path.isfile(loc_t) and not os.path.isfile(drv_t):
            if _atomic_copy(loc_t, drv_t):
                copied.append("→Drive:template.xlsx")
        elif os.path.isfile(drv_t) and not os.path.isfile(loc_t):
            if _atomic_copy(drv_t, loc_t):
                copied.append("←Drive:template.xlsx")
        _WL_SYNC_DONE = True
        return {
            "ok": True,
            "skipped": False,
            "copied": copied,
            "conflicts": conflicts,
            "source": drive_dir,
        }
    except Exception as e:
        _WL_SYNC_DONE = True
        return {
            "ok": False,
            "skipped": False,
            "copied": copied,
            "conflicts": conflicts,
            "source": drive_dir,
            "error": str(e),
        }


def push_worklog_day_to_drive(
    local_path: str,
    local_dir: str = "./uploaded_cache/worklog",
    *,
    force: bool = False,
) -> Optional[str]:
    """저장 직후 해당일 xlsx를 Drive worklog에 복사.

    force=False 이고 Drive 쪽이 더 최신이면 덮어쓰지 않고 None 반환
    (호출측에서 충돌 안내). force=True 이면 강제 복사.
    """
    drive_dir = resolve_drive_worklog_dir()
    if not drive_dir or not local_path or not os.path.isfile(local_path):
        return None
    name = os.path.basename(local_path)
    if not _is_worklog_day_file(name):
        return None
    dst = os.path.join(drive_dir, name)
    if (not force) and os.path.isfile(dst):
        try:
            lm, dm = os.path.getmtime(local_path), os.path.getmtime(dst)
            ls, ds = os.path.getsize(local_path), os.path.getsize(dst)
            if abs(lm - dm) >= 1.0 or ls != ds:
                # 기존 Drive 파일이 더 최신이면 충돌로 간주
                if dm > lm + 1.0:
                    return None
        except OSError:
            pass
    if _atomic_copy(local_path, dst):
        return dst
    return None


def push_worklog_month_archive_to_drive(
    local_month_path: str,
    *,
    year: Optional[int] = None,
    force: bool = False,
) -> Optional[str]:
    """월별 N월.xlsx 를 Drive worklog/일지/YYYY/ 로 복사 (로컬↔클라우드 공용)."""
    if not local_month_path or not os.path.isfile(local_month_path):
        return None
    name = os.path.basename(local_month_path)
    if not name.endswith("월.xlsx"):
        return None
    y = year
    if y is None:
        parent = os.path.basename(os.path.dirname(local_month_path))
        if parent.isdigit() and len(parent) == 4:
            y = int(parent)
    drv_dir = resolve_drive_worklog_archive_dir(y)
    if not drv_dir:
        return None
    dst = os.path.join(drv_dir, name)
    if (not force) and os.path.isfile(dst):
        try:
            lm, dm = os.path.getmtime(local_month_path), os.path.getmtime(dst)
            ls, ds = os.path.getsize(local_month_path), os.path.getsize(dst)
            if abs(lm - dm) >= 1.0 or ls != ds:
                if dm > lm + 1.0:
                    return None
        except OSError:
            pass
    if _atomic_copy(local_month_path, dst):
        return dst
    return None


def delete_worklog_day_from_drive(d: date, local_dir: str = "./uploaded_cache/worklog") -> list[str]:
    """Drive worklog 폴더에서 해당 일자 xlsx 삭제."""
    removed: List[str] = []
    drive_dir = resolve_drive_worklog_dir()
    if not drive_dir:
        return removed
    name = f"{d.isoformat()}.xlsx"
    path = os.path.join(drive_dir, name)
    if os.path.isfile(path):
        try:
            os.remove(path)
            removed.append(name)
        except OSError:
            pass
    return removed


def resolve_drive_conflict(
    name: str,
    local_dir: str = "./uploaded_cache/worklog",
    *,
    prefer: str = "local",
) -> Optional[str]:
    """충돌 일자 파일을 prefer(local|drive) 쪽으로 맞춤. 성공 시 대상 경로."""
    if not _is_worklog_day_file(name):
        return None
    drive_dir = resolve_drive_worklog_dir()
    if not drive_dir:
        return None
    loc = os.path.join(local_dir, name)
    drv = os.path.join(drive_dir, name)
    if prefer == "drive":
        if os.path.isfile(drv) and _atomic_copy(drv, loc):
            return loc
        return None
    if os.path.isfile(loc) and _atomic_copy(loc, drv):
        return drv
    return None
