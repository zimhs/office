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
    """Google Drive「내 드라이브/dashboard 복사본」절대경로. 없으면 None."""
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
            candidate = os.path.join(root, drive_label, DRIVE_COPY_NAME)
            if os.path.isdir(candidate):
                return candidate
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


def sync_drive_copy_into_cache(cache_dir: str = "./uploaded_cache") -> dict:
    """Drive 복사본 → uploaded_cache. 프로세스당 1회.

    Returns:
        {ok, source, copied: [..], skipped: bool, error?}
    """
    global _DONE
    if _DONE:
        return {"ok": True, "skipped": True, "copied": [], "source": None}
    _DONE = True

    drive_root = resolve_drive_dashboard_copy()
    if not drive_root:
        return {
            "ok": True,
            "skipped": True,
            "copied": [],
            "source": None,
            "note": "Drive 복사본 없음(클라우드·시드 캐시 사용)",
        }

    copied: List[str] = []
    try:
        os.makedirs(cache_dir, exist_ok=True)
        sales_dir = os.path.join(cache_dir, "sales")
        os.makedirs(sales_dir, exist_ok=True)

        for src_name, rel, name_txt in _CACHE_MAP:
            src = os.path.join(drive_root, src_name)
            dst = os.path.join(cache_dir, rel)
            if not _should_replace(src, dst):
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
                if _should_replace(src, dst) and _atomic_copy(src, dst):
                    copied.append(f"sales/{sn}")

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


def sync_cache_to_drive_copy(cache_dir: str = "./uploaded_cache") -> dict:
    """맥 uploaded_cache → Drive「dashboard 복사본」(업로드 반영).

    Google Drive 데스크톱이 클라우드로 올리면 아이패드 Drive 폴더에도 보임.
    Streamlit Cloud 시드는 별도 git push가 필요.
    """
    drive_root = resolve_drive_dashboard_copy()
    if not drive_root:
        return {
            "ok": False,
            "skipped": True,
            "copied": [],
            "source": None,
            "error": "Drive「dashboard 복사본」경로 없음",
        }

    copied: List[str] = []
    try:
        sales_dir = os.path.join(cache_dir, "sales")

        for drive_name, rel, _name_txt in _CACHE_MAP:
            src = os.path.join(cache_dir, rel)
            dst = os.path.join(drive_root, drive_name)
            if not os.path.isfile(src):
                continue
            if os.path.isfile(dst):
                try:
                    if (
                        os.path.getsize(src) == os.path.getsize(dst)
                        and abs(os.path.getmtime(src) - os.path.getmtime(dst)) < 1.0
                    ):
                        continue
                except OSError:
                    pass
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
                if os.path.isfile(dst):
                    try:
                        if (
                            os.path.getsize(src) == os.path.getsize(dst)
                            and abs(os.path.getmtime(src) - os.path.getmtime(dst)) < 1.0
                        ):
                            continue
                    except OSError:
                        pass
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
