"""개발·코드 수정 모드 — 빌드/동기화/경고·안내는 dev 모드에서만 표시."""
from __future__ import annotations

import streamlit as st

_UI_GATED = False


def is_dev_mode() -> bool:
    """코드 수정·디버그용 UI. URL ?dev=1 또는 secrets dev_mode=true."""
    try:
        if st.session_state.get("force_dev_mode") is True:
            return True
    except Exception:
        pass
    try:
        v = st.query_params.get("dev", "")
        if isinstance(v, (list, tuple)):
            v = v[0] if v else ""
        if str(v).strip().lower() in ("1", "true", "yes", "on"):
            try:
                st.session_state["force_dev_mode"] = True
            except Exception:
                pass
            return True
    except Exception:
        pass
    try:
        dm = st.secrets.get("dev_mode")
        if dm is True or str(dm).strip().lower() in ("1", "true", "yes", "on"):
            return True
    except Exception:
        pass
    return False


def apply_dev_ui_gate() -> None:
    """일반 사용 시 st.warning / st.info 를 숨김 (dev 모드면 그대로)."""
    global _UI_GATED
    if _UI_GATED or is_dev_mode():
        return
    _UI_GATED = True

    def _hidden(*_args, **_kwargs):
        return st.container()

    st.warning = _hidden  # type: ignore[method-assign]
    st.info = _hidden  # type: ignore[method-assign]


def dev_caption(text: str) -> None:
    if is_dev_mode():
        st.caption(text)
