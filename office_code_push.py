"""맥 Drive 동기화 버튼 → 코드만 GitHub office/main 에 올린다.

캐시·매출·비밀·방문할일 단가 파일은 넣지 않는다.
원격 main 이 앞서 있으면 fetch 후 rebase 하고 push 한다.
(fetch 없이 push 하면 rejected / fetch first 오류가 난다.)
"""
from __future__ import annotations

import os
import subprocess
from typing import Dict, List

_OFFICE_PUSH_URL = "https://github.com/zimhs/office.git"
_DENY_NAMES = {
    "visit_calendar_tab.py",
    "test_visit_calendar.py",
}


def repo_root() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def list_office_code_files(root: str | None = None) -> List[str]:
    base = root or repo_root()
    out: List[str] = []
    try:
        names = os.listdir(base)
    except OSError:
        return out
    for name in sorted(names):
        if name in _DENY_NAMES:
            continue
        if name.endswith(".py") or name == "requirements.txt":
            out.append(name)
    return out


def _run(cmd: List[str], cwd: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        cmd,
        cwd=cwd,
        text=True,
        capture_output=True,
        timeout=120,
    )


def _err_text(proc: subprocess.CompletedProcess) -> str:
    return (proc.stderr or proc.stdout or "").strip()[:240]


def push_office_code(*, message: str = "") -> Dict[str, object]:
    """코드 변경이 있으면 커밋 후 origin/main 에 push. 데이터 파일은 제외."""
    root = repo_root()
    files = [n for n in list_office_code_files(root) if os.path.isfile(os.path.join(root, n))]
    if not files:
        return {"ok": False, "skipped": True, "error": "올릴 코드 파일 없음", "files": []}

    # 원격이 앞서 있으면 push 가 rejected 된다 → 먼저 fetch
    fetch = _run(["git", "fetch", _OFFICE_PUSH_URL, "main"], root)
    if fetch.returncode != 0:
        return {
            "ok": False,
            "error": f"원격 확인 실패: {_err_text(fetch) or 'git fetch 실패'}",
            "files": files,
        }

    add = _run(["git", "add", "--"] + files, root)
    if add.returncode != 0:
        return {
            "ok": False,
            "error": _err_text(add) or "git add 실패",
            "files": files,
        }
    dirty = _run(["git", "diff", "--cached", "--quiet", "--"] + files, root)
    committed = False
    if dirty.returncode == 1:
        msg = (message or "chore: 맥 Drive 동기화에서 코드를 Cloud에 반영한다").strip()
        commit = _run(["git", "commit", "-m", msg], root)
        if commit.returncode != 0:
            return {
                "ok": False,
                "error": _err_text(commit) or "git commit 실패",
                "files": files,
            }
        committed = True

    # 로컬 커밋을 원격 main 위에 올린다. 뒤처져 있으면 fast-forward.
    rebase = _run(["git", "rebase", "FETCH_HEAD"], root)
    if rebase.returncode != 0:
        _run(["git", "rebase", "--abort"], root)
        # 로컬 커밋이 없으면 ff-only 로 맞추고 push 생략
        if not committed:
            ff = _run(["git", "merge", "--ff-only", "FETCH_HEAD"], root)
            if ff.returncode == 0:
                return {
                    "ok": True,
                    "committed": False,
                    "skipped": True,
                    "files": files,
                    "note": "원격 main 을 받아 맞춰 두었습니다. 올릴 로컬 코드 변경은 없습니다.",
                }
        return {
            "ok": False,
            "committed": committed,
            "error": (
                "원격 main 과 충돌합니다. "
                "dashboard_Update 로 최신을 받은 뒤 다시 동기화하세요. "
                f"({_err_text(rebase) or 'rebase 실패'})"
            ),
            "files": files,
        }

    ahead = _run(["git", "rev-list", "--count", "FETCH_HEAD..HEAD"], root)
    ahead_n = 0
    if ahead.returncode == 0:
        try:
            ahead_n = int((ahead.stdout or "0").strip() or "0")
        except ValueError:
            ahead_n = 0
    if ahead_n <= 0:
        return {
            "ok": True,
            "committed": committed,
            "skipped": True,
            "files": files,
            "note": "코드는 이미 Cloud(main)과 같습니다.",
        }

    push = _run(["git", "push", _OFFICE_PUSH_URL, "HEAD:main"], root)
    if push.returncode != 0:
        return {
            "ok": False,
            "committed": committed,
            "error": _err_text(push) or "git push 실패",
            "files": files,
        }
    return {
        "ok": True,
        "committed": committed,
        "skipped": not committed,
        "files": files,
    }
