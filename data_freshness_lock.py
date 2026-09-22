"""매출·채권 최신본 잠금.

한 번 더 새로운 파일이 적용되면 generation 을 올리고,
그보다 예전 파일은 캐시·Drive·uproad·git 시드에서 와도 쓰지 않는다.
최신본은 항상 적용한다.
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import os
import re
import shutil
from typing import Any, Dict, List, Optional, Tuple

LOCK_DIRNAME = ".data_lock"
LOCK_JSON = "lock.json"
DEBT_REL = "debt.csv"

_SALES_NAME_RE = re.compile(r"^(20\d{2})(\d{2})?(?:\.csv)?$", re.I)
_SALES_YMD_RE = re.compile(r"(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일")
_SALES_SCAN_HEAD = 24576
_SALES_SCAN_TAIL = 49152
_SALES_MD_RE = re.compile(r"(?<!\d)(\d{1,2})/(\d{1,2})(?!\d)")
_MONTH_COLS = tuple(f"{i}월" for i in range(1, 13))

Gen = Tuple[int, int, int]


def _lock_root(cache_dir: str) -> str:
    return os.path.join(cache_dir, LOCK_DIRNAME)


def _lock_json_path(cache_dir: str) -> str:
    return os.path.join(_lock_root(cache_dir), LOCK_JSON)


def _snapshot_path(cache_dir: str, rel: str) -> str:
    return os.path.join(_lock_root(cache_dir), rel.replace("\\", "/"))


def _file_sha256(path: str) -> str:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return ""


def _bytes_sha256(raw: bytes) -> str:
    return hashlib.sha256(raw or b"").hexdigest() if raw else ""


def _decode_text(raw: bytes, limit: int = 0) -> str:
    blob = raw[:limit] if limit and len(raw) > limit else raw
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return blob.decode(enc)
        except UnicodeDecodeError:
            continue
    return blob.decode("utf-8", errors="replace")


def parse_sales_name_period(name: str) -> int:
    base = os.path.basename(str(name or "")).strip()
    m = _SALES_NAME_RE.match(base)
    if not m:
        m2 = re.search(r"(20\d{2})(\d{2})?", base)
        if not m2:
            return 0
        y, mm = m2.group(1), m2.group(2)
    else:
        y, mm = m.group(1), m.group(2)
    return int(y) * 100 + (int(mm) if mm else 0)


def _max_sales_end_ymd(text: str, period: int) -> int:
    """헤더(YYYY년 M월 D일) 또는 본문 MM/DD 중 가장 늦은 매출일."""
    end_ymd = 0
    for yy, mo, dd in _SALES_YMD_RE.findall(text or ""):
        try:
            ymd = int(yy) * 10000 + int(mo) * 100 + int(dd)
        except ValueError:
            continue
        if ymd > end_ymd:
            end_ymd = ymd
    year = period // 100 if period >= 10000 else int(period or 0)
    if year < 2000:
        return end_ymd
    for mo, dd in _SALES_MD_RE.findall(text or ""):
        try:
            month, day = int(mo), int(dd)
        except ValueError:
            continue
        if not (1 <= month <= 12 and 1 <= day <= 31):
            continue
        ymd = year * 10000 + month * 100 + day
        if ymd > end_ymd:
            end_ymd = ymd
    return end_ymd


def sales_generation_from_bytes(name: str, raw: bytes) -> Gen:
    period = parse_sales_name_period(name)
    blob = raw or b""
    if len(blob) <= _SALES_SCAN_HEAD + _SALES_SCAN_TAIL:
        text = _decode_text(blob)
    else:
        text = (
            _decode_text(blob[:_SALES_SCAN_HEAD])
            + "\n"
            + _decode_text(blob[-_SALES_SCAN_TAIL:])
        )
    end_ymd = _max_sales_end_ymd(text, period)
    return (int(period or 0), int(end_ymd or 0), int(len(blob)))


def debt_generation_from_bytes(raw: bytes) -> Gen:
    if not raw:
        return (0, 0, 0)
    text = _decode_text(raw)
    last_month = 0
    month_sum = [0.0] * 13
    try:
        reader = csv.reader(io.StringIO(text))
        header = next(reader, None) or []
        idx_to_month: Dict[int, int] = {}
        for i, col in enumerate(header):
            key = str(col or "").strip()
            if key in _MONTH_COLS:
                idx_to_month[i] = int(key.replace("월", "") or 0)
        for row in reader:
            for i, month in idx_to_month.items():
                if i >= len(row) or month <= 0:
                    continue
                cell = str(row[i] or "").replace(",", "").strip()
                if not cell or cell in ("-", "nan", "None"):
                    continue
                try:
                    val = abs(float(cell))
                except ValueError:
                    continue
                if val > 0:
                    last_month = max(last_month, month)
                    month_sum[month] += val
    except Exception:
        pass
    return (int(last_month), int(round(month_sum[last_month])), int(len(raw)))


def _norm(gen: Optional[Gen]) -> Gen:
    if not gen:
        return (0, 0, 0)
    return (int(gen[0] or 0), int(gen[1] or 0), int(gen[2] or 0))


def generation_at_least(
    incoming: Optional[Gen],
    floor: Optional[Gen],
    *,
    kind: str = "sales",
) -> bool:
    """incoming 이 잠금/현재보다 같거나 더 최신이면 True.

    파일 크기(g2)는 채권은 비교하지 않는다. 같은 달 재업로드는 잔액·행이 줄어도 최신이다.
    매출은 헤더 종료일(g1)이 더 이전이면 예전이다. 같은 날짜인데 더 작으면 Drive 옛 파일로 본다.
    날짜를 못 읽으면 같은 달로 보고, 그때도 더 작은 파일은 예전으로 본다.
    """
    inc = _norm(incoming)
    fl = _norm(floor)
    if fl == (0, 0, 0):
        return True
    if inc[0] != fl[0]:
        return inc[0] >= fl[0]
    if kind == "debt":
        return True
    if fl[1] and inc[1] and inc[1] < fl[1]:
        return False
    if (not fl[1] or not inc[1] or inc[1] == fl[1]) and fl[2] and inc[2] and inc[2] < fl[2]:
        return False
    return True


def generation_strictly_newer(
    incoming: Optional[Gen],
    floor: Optional[Gen],
    *,
    kind: str = "sales",
) -> bool:
    """부트 Drive 동기화용. 같은 달·같은 날짜는 로컬 잠금을 유지한다."""
    inc = _norm(incoming)
    fl = _norm(floor)
    if fl == (0, 0, 0):
        return True
    if inc[0] != fl[0]:
        return inc[0] > fl[0]
    if kind == "debt":
        return False
    return bool(inc[1] and inc[1] > fl[1])


def _read_json(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, TypeError, ValueError):
        return {}


def _write_json(path: str, data: Dict[str, Any]) -> None:
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def load_lock(cache_dir: str) -> Dict[str, Any]:
    data = _read_json(_lock_json_path(cache_dir))
    if not data:
        return {"version": 1, "debt": {}, "sales": {}}
    data.setdefault("debt", {})
    data.setdefault("sales", {})
    return data


def save_lock(cache_dir: str, data: Dict[str, Any]) -> None:
    data["version"] = 1
    _write_json(_lock_json_path(cache_dir), data)


def _gen_from_lock_entry(entry: Any) -> Gen:
    if not isinstance(entry, dict):
        return (0, 0, 0)
    return (
        int(entry.get("g0") or 0),
        int(entry.get("g1") or 0),
        int(entry.get("g2") or 0),
    )


def _entry_for(gen: Gen, sha: str) -> Dict[str, Any]:
    return {"g0": gen[0], "g1": gen[1], "g2": gen[2], "sha256": sha}


def locked_debt_gen(cache_dir: str) -> Gen:
    return _gen_from_lock_entry(load_lock(cache_dir).get("debt"))


def locked_sales_gen(cache_dir: str, name: str) -> Gen:
    sales = load_lock(cache_dir).get("sales") or {}
    return _gen_from_lock_entry(sales.get(os.path.basename(name)))


def generation_from_file(path: str, *, kind: str, name: str = "") -> Gen:
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except OSError:
        return (0, 0, 0)
    if kind == "debt":
        return debt_generation_from_bytes(raw)
    return sales_generation_from_bytes(name or os.path.basename(path), raw)


def _atomic_write(path: str, raw: bytes) -> bool:
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp = path + ".freshlock"
    try:
        with open(tmp, "wb") as f:
            f.write(raw)
        os.replace(tmp, path)
        return True
    except OSError:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        try:
            with open(path, "wb") as f:
                f.write(raw)
            return True
        except OSError:
            return False


def _copy_file(src: str, dst: str) -> bool:
    parent = os.path.dirname(dst)
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp = dst + ".freshlock"
    try:
        shutil.copy2(src, tmp)
        os.replace(tmp, dst)
        return True
    except OSError:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        try:
            shutil.copy2(src, dst)
            return True
        except OSError:
            return False


def _snapshot_and_lock(cache_dir: str, rel: str, raw: bytes, gen: Gen, sha: str, *, kind: str) -> None:
    snap = _snapshot_path(cache_dir, rel)
    _atomic_write(snap, raw)
    data = load_lock(cache_dir)
    if kind == "debt":
        prev = _gen_from_lock_entry(data.get("debt"))
        if generation_at_least(gen, prev, kind="debt"):
            data["debt"] = _entry_for(gen, sha)
    else:
        key = os.path.basename(rel)
        sales = dict(data.get("sales") or {})
        prev = _gen_from_lock_entry(sales.get(key))
        if generation_at_least(gen, prev, kind="sales"):
            sales[key] = _entry_for(gen, sha)
        data["sales"] = sales
        periods = [int((sales.get(n) or {}).get("g0") or 0) for n in sales]
        data["max_sales_period"] = max(periods) if periods else 0
    save_lock(cache_dir, data)


def floor_gen(cache_dir: str, rel: str, *, kind: str, name: str, current_path: str) -> Gen:
    locked = locked_debt_gen(cache_dir) if kind == "debt" else locked_sales_gen(cache_dir, name)
    current = (0, 0, 0)
    if current_path and os.path.isfile(current_path):
        current = generation_from_file(current_path, kind=kind, name=name)
    a, b = _norm(locked), _norm(current)
    return a if a >= b else b


def write_bytes_if_newer(
    cache_dir: str,
    rel: str,
    raw: bytes,
    *,
    kind: str,
    name: str = "",
    allow_same: bool = True,
    force: bool = False,
) -> Tuple[bool, str]:
    """최신이면 캐시+스냅샷+잠금 갱신. 예전이면 False.

    allow_same: 업로드·불러오기 True. Drive 부트는 False(같은 달은 로컬 유지).
    force: 수동 동기화. 잠금을 건너뛰고 이 바이트를 최신으로 적용한다.

    Returns:
        (wrote, reason) reason: wrote / same / blocked_older / blocked_same / empty
    """
    if not raw or not cache_dir or not rel:
        return False, "empty"
    dest_name = name or os.path.basename(rel)
    dst = os.path.join(cache_dir, rel)
    if kind == "debt":
        incoming = debt_generation_from_bytes(raw)
    else:
        incoming = sales_generation_from_bytes(dest_name, raw)
    sha = _bytes_sha256(raw)
    if os.path.isfile(dst) and _file_sha256(dst) == sha:
        _snapshot_and_lock(cache_dir, rel, raw, incoming, sha, kind=kind)
        return False, "same"
    if not force:
        floor = floor_gen(cache_dir, rel, kind=kind, name=dest_name, current_path=dst)
        if not generation_at_least(incoming, floor, kind=kind):
            return False, "blocked_older"
        if not allow_same and not generation_strictly_newer(incoming, floor, kind=kind):
            return False, "blocked_same"
    if not _atomic_write(dst, raw):
        return False, "write_failed"
    _snapshot_and_lock(cache_dir, rel, raw, incoming, sha, kind=kind)
    return True, "wrote"


def copy_file_if_newer(
    src: str,
    dst: str,
    cache_dir: str,
    *,
    kind: str,
    name: str = "",
    allow_same: bool = True,
    force: bool = False,
) -> Tuple[bool, str]:
    if not src or not os.path.isfile(src):
        return False, "empty"
    try:
        with open(src, "rb") as f:
            raw = f.read()
    except OSError:
        return False, "empty"
    rel = os.path.relpath(dst, cache_dir).replace("\\", "/") if cache_dir else os.path.basename(dst)
    return write_bytes_if_newer(
        cache_dir,
        rel,
        raw,
        kind=kind,
        name=name or os.path.basename(dst),
        allow_same=allow_same,
        force=force,
    )


def seed_lock_from_cache(cache_dir: str) -> None:
    """잠금이 비어 있으면 현재 캐시를 최신 바닥으로 찍는다."""
    if not cache_dir or not os.path.isdir(cache_dir):
        return
    data = load_lock(cache_dir)
    debt_path = os.path.join(cache_dir, DEBT_REL)
    debt_entry = data.get("debt") or {}
    if os.path.isfile(debt_path) and not debt_entry.get("sha256"):
        try:
            with open(debt_path, "rb") as f:
                raw = f.read()
            gen = debt_generation_from_bytes(raw)
            _snapshot_and_lock(cache_dir, DEBT_REL, raw, gen, _bytes_sha256(raw), kind="debt")
        except OSError:
            pass
    sales_dir = os.path.join(cache_dir, "sales")
    locked_sales = (load_lock(cache_dir).get("sales") or {})
    if os.path.isdir(sales_dir):
        try:
            names = [n for n in os.listdir(sales_dir) if n.lower().endswith(".csv")]
        except OSError:
            names = []
        for name in names:
            if name in locked_sales and _norm(_gen_from_lock_entry(locked_sales.get(name))) != (0, 0, 0):
                continue
            path = os.path.join(sales_dir, name)
            try:
                with open(path, "rb") as f:
                    raw = f.read()
                gen = sales_generation_from_bytes(name, raw)
                _snapshot_and_lock(cache_dir, f"sales/{name}", raw, gen, _bytes_sha256(raw), kind="sales")
            except OSError:
                continue


def restore_locked_latest(cache_dir: str) -> List[str]:
    """git/Drive가 예전 파일로 캐시를 바꿔도 잠긴 최신본을 되돌린다.

    캐시가 스냅샷보다 더 최신이면 캐시를 스냅샷에 올린다(최신 업데이트 적용).
    """
    restored: List[str] = []
    if not cache_dir:
        return restored
    os.makedirs(cache_dir, exist_ok=True)
    seed_lock_from_cache(cache_dir)
    root = _lock_root(cache_dir)
    if not os.path.isdir(root):
        return restored

    pairs: List[Tuple[str, str]] = []
    snap_debt = _snapshot_path(cache_dir, DEBT_REL)
    if os.path.isfile(snap_debt):
        pairs.append((DEBT_REL, "debt"))
    snap_sales = os.path.join(root, "sales")
    if os.path.isdir(snap_sales):
        try:
            for name in os.listdir(snap_sales):
                if name.lower().endswith(".csv"):
                    pairs.append((f"sales/{name}", "sales"))
        except OSError:
            pass

    for rel, kind in pairs:
        snap = _snapshot_path(cache_dir, rel)
        dst = os.path.join(cache_dir, rel)
        name = os.path.basename(rel)
        if not os.path.isfile(snap):
            continue
        if os.path.isfile(dst) and _file_sha256(dst) == _file_sha256(snap):
            continue
        snap_gen = generation_from_file(snap, kind=kind, name=name)
        if os.path.isfile(dst):
            cur_gen = generation_from_file(dst, kind=kind, name=name)
            if generation_at_least(cur_gen, snap_gen, kind=kind):
                try:
                    with open(dst, "rb") as f:
                        raw = f.read()
                    _snapshot_and_lock(cache_dir, rel, raw, cur_gen, _bytes_sha256(raw), kind=kind)
                except OSError:
                    pass
                continue
        if _copy_file(snap, dst):
            restored.append(rel)
    return restored


def prune_annual_if_monthly(sales_dir: str) -> List[str]:
    """같은 해 월별 CSV가 있으면 연간 YYYY.csv 만 제거 (월별 최신은 유지)."""
    removed: List[str] = []
    if not sales_dir or not os.path.isdir(sales_dir):
        return removed
    try:
        names = [n for n in os.listdir(sales_dir) if n.lower().endswith(".csv")]
    except OSError:
        return removed
    by_year: Dict[str, List[str]] = {}
    for n in names:
        m = re.match(r"^(20\d{2})(\d{2})?\.csv$", n, re.I)
        if not m:
            continue
        by_year.setdefault(m.group(1), []).append(n)
    for year, grp in by_year.items():
        monthlies = [n for n in grp if re.match(rf"^{year}\d{{2}}\.csv$", n, re.I)]
        annuals = [n for n in grp if re.match(rf"^{year}\.csv$", n, re.I)]
        if not monthlies:
            continue
        for a in annuals:
            path = os.path.join(sales_dir, a)
            try:
                os.remove(path)
                removed.append(a)
            except OSError:
                pass
    return removed


def clear_data_lock(cache_dir: str) -> None:
    root = _lock_root(cache_dir)
    if not os.path.isdir(root):
        return
    try:
        shutil.rmtree(root)
    except OSError:
        pass


def lock_status_caption(cache_dir: str) -> str:
    data = load_lock(cache_dir)
    debt = _gen_from_lock_entry(data.get("debt"))
    max_period = int(data.get("max_sales_period") or 0)
    sales = data.get("sales") or {}
    if not max_period and isinstance(sales, dict):
        max_period = max((int((sales.get(n) or {}).get("g0") or 0) for n in sales), default=0)
    bits: List[str] = []
    if max_period:
        y, m = divmod(max_period, 100)
        bits.append(f"매출 {y}-{m:02d}" if m else f"매출 {y}")
    if debt[0]:
        bits.append(f"채권 {debt[0]}월")
    if not bits:
        return ""
    return "최신 잠금 · " + " · ".join(bits)
