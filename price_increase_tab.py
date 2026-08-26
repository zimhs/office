"""단가인상 공문 탭 — 거래처·메일·품목단가·공문작성·발송.

다른 탭과 공유 상태를 바꾸지 않음. sales_df / address는 인자로만 받음.
"""
from __future__ import annotations

import io
import json
import os
import re
import smtplib
import ssl
from copy import copy
from datetime import date, datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from typing import Any, Optional
from urllib.parse import quote

import pandas as pd
import streamlit as st

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, Side
except Exception:  # pragma: no cover
    Workbook = None  # type: ignore
    load_workbook = None  # type: ignore

PI_DIR = os.path.join("uploaded_cache", "price_increase")
PI_MAIL_CSV = os.path.join(PI_DIR, "mail_contacts.csv")
PI_TEMPLATE = os.path.join(PI_DIR, "공문양식.xlsx")
PI_DRAFTS = os.path.join(PI_DIR, "drafts")
PI_SENT_LOG = os.path.join(PI_DRAFTS, "sent_log.jsonl")
PI_UI_BUILD = "2026-08-26e · 편집가능품목표(단가➠인상단가)·일괄발송"

# 기본 공문 후보 (캐시 원본 우선)
_TEMPLATE_CANDIDATES = (
    os.path.join(PI_DIR, "탄산단가인상공문.xlsx"),
    os.path.join("uploaded_cache", "price_increase", "탄산단가인상공문.xlsx"),
    "탄산단가인상공문.xlsx",
    os.path.expanduser("~/Desktop/탄산단가인상공문.xlsx"),
    os.path.expanduser("~/Desktop/업무/탄산단가인상공문.xlsx"),
    os.path.expanduser("~/Desktop/dashboard/탄산단가인상공문.xlsx"),
)

_NOISE_ITEM = re.compile(
    r"입금|이월|단가차액|잔액정리|기화기|공사|작업비|임대|운반비|회수|보증|수수료|운임",
    re.I,
)
_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def _ensure_dirs() -> None:
    os.makedirs(PI_DIR, exist_ok=True)
    os.makedirs(PI_DRAFTS, exist_ok=True)


def _secret_get(*keys: str) -> str:
    try:
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


def _norm_name(s: Any) -> str:
    t = str(s or "").strip()
    t = re.sub(r"\s+", "", t)
    t = t.rstrip(".")
    return t.lower()


def load_mail_contacts(path: str = PI_MAIL_CSV) -> pd.DataFrame:
    if not os.path.isfile(path):
        return pd.DataFrame(columns=["거래처", "이메일", "비고"])
    try:
        df = pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        try:
            df = pd.read_csv(path, encoding="cp949")
        except Exception:
            return pd.DataFrame(columns=["거래처", "이메일", "비고"])
    return _normalize_mail_df(df)


def _normalize_mail_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["거래처", "이메일", "비고"])
    cols = {str(c).strip(): c for c in df.columns}
    name_c = None
    mail_c = None
    note_c = None
    for label, c in cols.items():
        low = label.lower().replace(" ", "")
        if name_c is None and any(k in low for k in ("거래처", "이름", "성명", "상호", "업체", "name", "회사")):
            name_c = c
        if mail_c is None and any(k in low for k in ("메일", "email", "e-mail", "이메일", "mail")):
            mail_c = c
        if note_c is None and any(k in low for k in ("비고", "메모", "소속", "그룹", "note")):
            note_c = c
    if mail_c is None:
        # 열 값에서 이메일 패턴 찾기
        for c in df.columns:
            sample = df[c].astype(str).head(30).str.cat(sep=" ")
            if _EMAIL_RE.search(sample):
                mail_c = c
                break
    if name_c is None:
        for c in df.columns:
            if c != mail_c:
                name_c = c
                break
    out = pd.DataFrame()
    out["거래처"] = df[name_c].astype(str).str.strip() if name_c is not None else ""
    raw_mail = df[mail_c].astype(str) if mail_c is not None else ""
    out["이메일"] = raw_mail.map(lambda x: (_EMAIL_RE.search(str(x)).group(0) if _EMAIL_RE.search(str(x)) else ""))
    out["비고"] = df[note_c].astype(str).str.strip() if note_c is not None else ""
    out = out[(out["거래처"] != "") & (out["거래처"] != "nan") & (out["이메일"] != "")]
    out = out.drop_duplicates(subset=["거래처", "이메일"], keep="last")
    return out.reset_index(drop=True)


def save_mail_contacts(df: pd.DataFrame, path: str = PI_MAIL_CSV) -> None:
    _ensure_dirs()
    df.to_csv(path, index=False, encoding="utf-8-sig")


def lookup_email(client: str, mail_df: pd.DataFrame) -> str:
    if not client or mail_df is None or mail_df.empty:
        return ""
    key = _norm_name(client)
    for _, row in mail_df.iterrows():
        if _norm_name(row.get("거래처")) == key:
            return str(row.get("이메일") or "").strip()
    # 부분 일치 (부모명 포함)
    for _, row in mail_df.iterrows():
        n = _norm_name(row.get("거래처"))
        if key and n and (key in n or n in key):
            return str(row.get("이메일") or "").strip()
    return ""


def list_staff_options(sales_df: pd.DataFrame) -> list[str]:
    if sales_df is None or sales_df.empty or "담당자" not in sales_df.columns:
        return []
    vals = sorted(
        {
            str(x).strip()
            for x in sales_df["담당자"].dropna().unique()
            if str(x).strip() and str(x).strip() not in ("nan", "거래종료")
        }
    )
    return vals


def list_clients_for_staff(sales_df: pd.DataFrame, staff: str) -> list[str]:
    if sales_df is None or sales_df.empty:
        return []
    df = sales_df
    if staff and staff != "전체" and "담당자" in df.columns:
        df = df[df["담당자"].astype(str).str.strip() == staff]
    names: set[str] = set()
    if "거래처" in df.columns:
        names |= set(df["거래처"].dropna().astype(str).str.strip().unique())
    if "거래처_원본" in df.columns:
        names |= set(df["거래처_원본"].dropna().astype(str).str.strip().unique())
    return sorted(n for n in names if n and n != "nan")


def classify_client_kind(name: str, sales_df: pd.DataFrame) -> str:
    """종속 | 단독(부모) | 단독."""
    if not name:
        return "단독"
    if "(" in name and name.endswith(")"):
        return "종속"
    if sales_df is not None and not sales_df.empty and "거래처_원본" in sales_df.columns:
        base = name.rstrip(".")
        kids = sales_df[
            sales_df["거래처"].astype(str).str.contains(re.escape(base) + r"\(", regex=True, na=False)
        ]
        if len(kids) > 0:
            return "단독(부모)"
    return "단독"


def latest_unit_prices(sales_df: pd.DataFrame, client: str) -> pd.DataFrame:
    """거래처 선택 시 최근 적용 단가 (품목별)."""
    empty = pd.DataFrame(columns=["품목명", "기존단가", "최근매출일", "출고량합"])
    if sales_df is None or sales_df.empty or not client:
        return empty
    df = sales_df.copy()
    # 부모 선택 시 종속 포함
    if "거래처_원본" in df.columns:
        m = (df["거래처"].astype(str).str.strip() == client) | (
            df["거래처_원본"].astype(str).str.strip() == client
        )
    else:
        m = df["거래처"].astype(str).str.strip() == client
    sub = df.loc[m]
    if sub.empty:
        return empty
    if "품목명" not in sub.columns or "단가" not in sub.columns:
        return empty
    sub = sub.copy()
    sub["품목명"] = sub["품목명"].astype(str).str.strip()
    sub = sub[~sub["품목명"].str.contains(_NOISE_ITEM, na=False)]
    sub["단가"] = pd.to_numeric(sub["단가"], errors="coerce").fillna(0)
    if "출고량" in sub.columns:
        sub["출고량"] = pd.to_numeric(sub["출고량"], errors="coerce").fillna(0)
    else:
        sub["출고량"] = 0.0
    date_col = "매출일_dt" if "매출일_dt" in sub.columns else None
    rows = []
    for item, g in sub.groupby("품목명", dropna=True):
        g2 = g[g["단가"] > 0]
        if g2.empty:
            continue
        if date_col and g2[date_col].notna().any():
            g2 = g2.sort_values(date_col)
            last = g2.iloc[-1]
            price = float(last["단가"])
            last_d = last[date_col]
            last_s = (
                last_d.strftime("%Y-%m-%d")
                if hasattr(last_d, "strftime")
                else str(last_d)[:10]
            )
        else:
            price = float(g2["단가"].mode().iloc[0]) if not g2["단가"].mode().empty else float(g2["단가"].iloc[-1])
            last_s = ""
        rows.append(
            {
                "품목명": str(item),
                "기존단가": round(price, 2),
                "최근매출일": last_s,
                "출고량합": float(g["출고량"].sum()),
            }
        )
    if not rows:
        return empty
    out = pd.DataFrame(rows).sort_values("출고량합", ascending=False).reset_index(drop=True)
    return out


def default_increase_price(old: float, pct: float) -> float:
    try:
        return round(float(old) * (1.0 + float(pct) / 100.0), 1)
    except Exception:
        return float(old or 0)


def apply_increase_by_amount(old: float, amount: float) -> float:
    try:
        return round(float(old) + float(amount), 1)
    except Exception:
        return float(old or 0)


# 탄산단가인상공문 기본 본문 (양식 셀 유지, 내용만 편집)
_DEFAULT_LETTER_PARAS: dict[str, str] = {
    "C16": "1.귀사의 무궁한 발전을 기원하며, 평소 당사에 보내주시는 신뢰와 협력에 깊은 감사를 드립니다.",
    "C18": "2.당사는 귀사와의 지속적인 파트너십 유지를 최우선 가치로 삼아, 대내외적인 원가 상승 압박 속에서도 경영 ",
    "C19": "효율화를 통해 단가 인상을 최대한 억제해 왔습니다.",
    "C21": "3.그러나 최근 지정학적 리스크 심화와 글로벌 에너지 공급망의 불안정성으로 인해 당사가 감내할 수 있는 . ",
    "C22": "임계치를 상회하는 제조 원가 상승이 발생하였습니다.",
    "C23": "이에 안정적인 품질 유지와 지속 가능한 공급 체계 확보를 위해 부득이하게 아래와 같이 단가 인상을 요청드리오니 널리 양해하여 ",
    "C24": "양해하여 주시기 바랍니다. ",
    "C26": "4. 주요 인상 요인 분석",
    "C27": "• 중동 분쟁 장기화에 따른 에너지 비용 급등: 중동 지역의 지정학적 불안정 지속으로 국제 유가 및 천연가스 가격의 변동성이 확대되었으며, 탄산 원료 가스(Raw Gas) 수급 비용 및 생산 설비 가동 에너지 비용이 급격히 상승함.",
    "C29": "• 원료 가스 확보 단가 및 제조 원가 상승: 석유화학 플랜트 가동률 변화와 원료 공급원의 제한적 수급 상황이 맞물려 원료 가스 매입 단가가 폭등하였으며, 정제 및 액화 과정의 부자재 가격 상승이 직접적인 원가 부담으로 작용함.",
    "C31": "• 물류망 불안정에 따른 운반비 가중: 유가 상승 및 요소수 등 차량 유지비 증가로 인해 내륙 운송 단가가 상향 평준화되었으며, 특히 특수 고압 탱크로리 운영비용 상승이 전체 공급가에 심대한 영향을 미침.",
    "C33": "5. 당사는 이번 조정을 바탕으로 더욱 철저한 품질 관리와 원활한 공급 체계를 구축하여 귀사의 기대에 보답할 것을 약속드립니다.",
    "C35": "                                      단가 조정 내용",
}


def ensure_default_template(path: str = PI_TEMPLATE) -> str:
    """원본 공문 양식이 없으면 기본 양식 생성."""
    _ensure_dirs()
    if os.path.isfile(path) and load_workbook is not None:
        return path
    if Workbook is None:
        return path
    wb = Workbook()
    ws = wb.active
    ws.title = "단가인상공문"
    thin = Border(
        left=Side(style="thin", color="94A3B8"),
        right=Side(style="thin", color="94A3B8"),
        top=Side(style="thin", color="94A3B8"),
        bottom=Side(style="thin", color="94A3B8"),
    )
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 28

    ws.merge_cells("B2:F2")
    ws["B2"] = "신일가스(주)"
    ws["B2"].font = Font(name="맑은 고딕", size=16, bold=True, color="0F766E")
    ws["B2"].alignment = Alignment(horizontal="center")

    ws.merge_cells("B3:F3")
    ws["B3"] = "단 가 인 상 안 내 공 문"
    ws["B3"].font = Font(name="맑은 고딕", size=18, bold=True)
    ws["B3"].alignment = Alignment(horizontal="center")

    ws["B5"] = "문서번호"
    ws["C5"] = "{{DOC_NO}}"
    ws["B6"] = "작성일"
    ws["C6"] = "{{DATE}}"
    ws["B7"] = "수신"
    ws["C7"] = "{{CLIENT}} 귀중"
    ws["B8"] = "참조"
    ws["C8"] = "{{EMAIL}}"
    ws["B9"] = "제목"
    ws.merge_cells("C9:F9")
    ws["C9"] = "{{TITLE}}"

    ws.merge_cells("B11:F11")
    ws["B11"] = (
        "항상 저희 신일가스를 이용해 주셔서 감사드립니다. "
        "원자재·물류비 상승에 따라 아래와 같이 공급 단가 조정을 안내드립니다."
    )
    ws["B11"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[11].height = 48

    headers = ["순번", "품목명", "기존단가(원)", "인상적용단가(원)", "비고"]
    for i, h in enumerate(headers, start=2):
        cell = ws.cell(14, i, h)
        cell.font = Font(name="맑은 고딕", bold=True, color="FFFFFF")
        cell.fill = __import__("openpyxl.styles", fromlist=["PatternFill"]).PatternFill(
            "solid", fgColor="0F766E"
        )
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin

    # 플레이스홀더 행 (최대 20)
    for r in range(15, 35):
        for c in range(2, 7):
            cell = ws.cell(r, c, "")
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 2, f"{{{{ITEM_{r-14}_NO}}}}")
        ws.cell(r, 3, f"{{{{ITEM_{r-14}_NAME}}}}")
        ws.cell(r, 4, f"{{{{ITEM_{r-14}_OLD}}}}")
        ws.cell(r, 5, f"{{{{ITEM_{r-14}_NEW}}}}")
        ws.cell(r, 6, f"{{{{ITEM_{r-14}_NOTE}}}}")

    ws.merge_cells("B36:F36")
    ws["B36"] = "{{BODY}}"
    ws["B36"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[36].height = 80

    ws.merge_cells("B38:F38")
    ws["B38"] = "적용예정일: {{EFFECTIVE}}"
    ws.merge_cells("B40:F40")
    ws["B40"] = "신일가스(주) 화성공장"
    ws["B40"].alignment = Alignment(horizontal="right")
    ws.merge_cells("B41:F41")
    ws["B41"] = "문의: {{CONTACT}}"
    ws["B41"].alignment = Alignment(horizontal="right")

    wb.save(path)
    wb.close()
    return path


def _is_real_letter_template(ws) -> bool:
    v9 = str(ws["C9"].value or "")
    v10 = str(ws["C10"].value or "")
    return "문서번호" in v9 or "발송일자" in v10


def _format_doc_no(client: str, doc_no: str = "") -> str:
    if doc_no:
        return doc_no
    return f"신일(1)-{date.today().strftime('%y%m')}-{abs(hash(client)) % 10000:04d}"


def _format_korean_effective(effective: str) -> str:
    t = str(effective or "").strip()
    if not t:
        return f"{date.today().strftime('%Y년 %m월 %d일')} 출고분부터"
    m = re.match(r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})", t)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{y}년 {mo:02d}월 {d:02d}일 출고분부터"
    if "년" in t and "월" in t:
        return t if "출고" in t else f"{t} 출고분부터"
    return t


def _format_c37_target_items(items: list[dict]) -> str:
    names = [str(it.get("품목명") or "").strip() for it in items if str(it.get("품목명") or "").strip()]
    joined = ", ".join(names) if names else "-"
    return f"                           • 대상 품목 : {joined} "


def _format_c38_price_lines(items: list[dict]) -> str:
    parts: list[str] = []
    for it in items:
        name = str(it.get("품목명") or "").strip()
        if not name:
            continue
        old = float(it.get("기존단가") or 0)
        new = float(it.get("인상적용단가") or 0)
        diff = new - old
        if len(items) == 1:
            parts.append(f"기존단가 {name} {old:,.0f}원 ➠ {new:,.0f}원 (+{diff:,.0f}원)")
        else:
            parts.append(f"{name} {old:,.0f}원 ➠ {new:,.0f}원 (+{diff:,.0f}원)")
    body = " / ".join(parts) if parts else "별도 협의"
    return f"                           • 단가 인상 금액 : {body} "


def _fill_real_template_cells(
    ws,
    *,
    client: str,
    title: str,
    effective: str,
    items: list[dict],
    doc_no: str,
    send_date: Optional[date] = None,
    letter_paras: Optional[dict[str, str]] = None,
    c37_override: str = "",
    c38_override: str = "",
) -> None:
    """양식(레이아웃·로고)은 유지하고 내용 셀만 채움. 적용%/인상금액 숫자는 공문에 넣지 않음."""
    send_date = send_date or date.today()
    ws["C9"].value = f"문서번호 : {doc_no}"
    ws["C10"].value = f"발송일자 : {send_date.strftime('%Y.%m.%d')}"
    ws["C11"].value = f"수    신 : {client}"
    ws["C12"].value = f"제    목 : {title}"
    paras = letter_paras if letter_paras is not None else _DEFAULT_LETTER_PARAS
    for coord, text in paras.items():
        if coord in ws:
            ws[coord].value = text
    ws["C37"].value = (
        c37_override.strip()
        if str(c37_override or "").strip()
        else _format_c37_target_items(items)
    )
    ws["C38"].value = (
        c38_override.strip()
        if str(c38_override or "").strip()
        else _format_c38_price_lines(items)
    )
    ws["C39"].value = f"                           • 시행 일자 : {_format_korean_effective(effective)}"


def fill_letter_workbook(
    template_path: str,
    *,
    client: str,
    email: str,
    title: str,
    body: str,
    effective: str,
    contact: str,
    items: list[dict],
    doc_no: str = "",
    send_date: Optional[date] = None,
    letter_paras: Optional[dict[str, str]] = None,
    c37_override: str = "",
    c38_override: str = "",
) -> bytes:
    if load_workbook is None:
        raise RuntimeError("openpyxl 필요")
    ensure_default_template(template_path)
    wb = load_workbook(template_path)
    ws = wb.active
    today = send_date or date.today()
    today_s = today.strftime("%Y-%m-%d") if hasattr(today, "strftime") else str(today)[:10]
    doc_no = _format_doc_no(client, doc_no)

    if _is_real_letter_template(ws):
        _fill_real_template_cells(
            ws,
            client=client,
            title=title,
            effective=effective,
            items=items,
            doc_no=doc_no,
            send_date=today if isinstance(today, date) else date.today(),
            letter_paras=letter_paras,
            c37_override=c37_override,
            c38_override=c38_override,
        )
    else:
        mapping = {
            "{{DOC_NO}}": doc_no,
            "{{DATE}}": today_s,
            "{{CLIENT}}": client,
            "{{EMAIL}}": email or "-",
            "{{TITLE}}": title,
            "{{BODY}}": body,
            "{{EFFECTIVE}}": effective,
            "{{CONTACT}}": contact,
        }
        for i in range(1, 21):
            if i <= len(items):
                it = items[i - 1]
                mapping[f"{{{{ITEM_{i}_NO}}}}"] = str(i)
                mapping[f"{{{{ITEM_{i}_NAME}}}}"] = str(it.get("품목명") or "")
                mapping[f"{{{{ITEM_{i}_OLD}}}}"] = f"{float(it.get('기존단가') or 0):,.1f}"
                mapping[f"{{{{ITEM_{i}_NEW}}}}"] = f"{float(it.get('인상적용단가') or 0):,.1f}"
                mapping[f"{{{{ITEM_{i}_NOTE}}}}"] = str(it.get("비고") or "")
            else:
                mapping[f"{{{{ITEM_{i}_NO}}}}"] = ""
                mapping[f"{{{{ITEM_{i}_NAME}}}}"] = ""
                mapping[f"{{{{ITEM_{i}_OLD}}}}"] = ""
                mapping[f"{{{{ITEM_{i}_NEW}}}}"] = ""
                mapping[f"{{{{ITEM_{i}_NOTE}}}}"] = ""

        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str):
                    for k, rep in mapping.items():
                        if k in v:
                            v = v.replace(k, rep)
                    cell.value = v

        for r in range(15, 35):
            if not ws.cell(r, 3).value:
                for c in range(2, 7):
                    ws.cell(r, c).value = None

    bio = io.BytesIO()
    wb.save(bio)
    wb.close()
    return bio.getvalue()


def resolve_letter_template() -> str:
    """탄산단가인상공문.xlsx 우선, 없으면 기본 생성 양식."""
    custom = st.session_state.get("pi_template_path")
    if custom and os.path.isfile(custom):
        return custom
    for p in _TEMPLATE_CANDIDATES:
        if p and os.path.isfile(p):
            return p
    return ensure_default_template(PI_TEMPLATE)


def smtp_settings() -> dict:
    """다음메일(Daum) SMTP 기본. secrets로 덮어씀."""
    provider = (_secret_get("smtp_provider", "SMTP_PROVIDER") or "daum").strip().lower()
    defaults = {
        "daum": {"host": "smtp.daum.net", "port": 465, "ssl": True},
        "kakao": {"host": "smtp.daum.net", "port": 465, "ssl": True},
        "naver": {"host": "smtp.naver.com", "port": 465, "ssl": True},
        "gmail": {"host": "smtp.gmail.com", "port": 465, "ssl": True},
    }
    base = dict(defaults.get(provider, defaults["daum"]))
    host = _secret_get("smtp_host", "SMTP_HOST") or base["host"]
    port_s = _secret_get("smtp_port", "SMTP_PORT")
    port = int(port_s) if port_s else int(base["port"])
    ssl_s = (_secret_get("smtp_ssl", "SMTP_SSL") or ("1" if base["ssl"] else "0")).strip().lower()
    use_ssl = ssl_s in ("1", "true", "yes", "ssl")
    user = _secret_get("smtp_user", "SMTP_USER", "mail_user", "daum_mail_id")
    password = _secret_get(
        "smtp_password",
        "SMTP_PASSWORD",
        "mail_password",
        "mail_app_password",
        "daum_mail_password",
    )
    from_addr = _secret_get("smtp_from", "SMTP_FROM") or user
    from_name = _secret_get("smtp_from_name", "SMTP_FROM_NAME") or "신일가스"
    return {
        "provider": provider,
        "host": host,
        "port": port,
        "ssl": use_ssl,
        "user": user,
        "password": password,
        "from_addr": from_addr,
        "from_name": from_name,
        "ready": bool(user and password),
    }


def smtp_status_label(cfg: Optional[dict] = None) -> str:
    cfg = cfg or smtp_settings()
    if not cfg.get("ready"):
        return "미연동 — secrets에 smtp_user / smtp_password 필요"
    return f"연동됨 · {cfg.get('user')} → {cfg.get('host')}:{cfg.get('port')}"


def test_smtp_connection(cfg: Optional[dict] = None) -> tuple[bool, str]:
    cfg = cfg or smtp_settings()
    if not cfg.get("ready"):
        return False, "smtp_user / smtp_password 가 secrets에 없습니다."
    host, port = cfg["host"], int(cfg["port"])
    user, password = cfg["user"], cfg["password"]
    try:
        context = ssl.create_default_context()
        if cfg.get("ssl") or port == 465:
            with smtplib.SMTP_SSL(host, port, context=context, timeout=20) as server:
                server.login(user, password)
        else:
            with smtplib.SMTP(host, port, timeout=20) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(user, password)
        return True, f"로그인 성공 ({host}:{port})"
    except smtplib.SMTPAuthenticationError:
        return (
            False,
            "로그인 실패 — 다음메일: 메일설정→POP3/IMAP 사용 ON, 아이디는 전체메일(예: id@daum.net)",
        )
    except Exception as e:
        return False, str(e)


def send_mail_smtp(
    *,
    to_addr: str,
    subject: str,
    body: str,
    attachment_bytes: Optional[bytes] = None,
    attachment_name: str = "단가인상공문.xlsx",
    cc: str = "",
) -> tuple[bool, str]:
    cfg = smtp_settings()
    if not cfg.get("ready"):
        return (
            False,
            "메일 미연동: `.streamlit/secrets.toml` 또는 Cloud Secrets에 "
            "smtp_user / smtp_password 를 넣으세요. (다음: smtp.daum.net:465)",
        )
    if not to_addr:
        return False, "수신 메일 없음"

    recipients = [a.strip() for a in re.split(r"[;,]", to_addr) if a.strip()]
    cc_list = [a.strip() for a in re.split(r"[;,]", cc or "") if a.strip()]

    msg = MIMEMultipart()
    msg["From"] = formataddr((cfg["from_name"], cfg["from_addr"]))
    msg["To"] = ", ".join(recipients)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    if attachment_bytes:
        # 한글 파일명 안전 처리
        safe_name = attachment_name or "letter.xlsx"
        part = MIMEApplication(attachment_bytes, Name=safe_name)
        part.add_header("Content-Disposition", "attachment", filename=safe_name)
        msg.attach(part)

    all_rcpt = recipients + cc_list
    try:
        context = ssl.create_default_context()
        host, port = cfg["host"], int(cfg["port"])
        if cfg.get("ssl") or port == 465:
            with smtplib.SMTP_SSL(host, port, context=context, timeout=45) as server:
                server.login(cfg["user"], cfg["password"])
                server.sendmail(cfg["from_addr"], all_rcpt, msg.as_string())
        else:
            with smtplib.SMTP(host, port, timeout=45) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                server.login(cfg["user"], cfg["password"])
                server.sendmail(cfg["from_addr"], all_rcpt, msg.as_string())
        return True, f"발송 완료 → {', '.join(all_rcpt)}"
    except smtplib.SMTPAuthenticationError:
        return False, "SMTP 인증 실패 — 아이디/비밀번호 또는 POP3/IMAP 사용 설정 확인"
    except Exception as e:
        return False, f"발송 실패: {e}"


def append_sent_log(
    *,
    client: str,
    email: str,
    subject: str,
    ok: bool,
    mode: str = "single",
    staff: str = "",
    items: int = 0,
    msg: str = "",
) -> None:
    _ensure_dirs()
    row = {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "date": date.today().isoformat(),
        "client": client,
        "email": email,
        "subject": subject,
        "ok": bool(ok),
        "mode": mode,
        "staff": staff,
        "items": int(items or 0),
        "msg": msg or "",
    }
    try:
        with open(PI_SENT_LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
    except OSError:
        pass


def load_sent_log() -> pd.DataFrame:
    cols = ["ts", "date", "client", "email", "subject", "ok", "mode", "staff", "items", "msg"]
    if not os.path.isfile(PI_SENT_LOG):
        return pd.DataFrame(columns=cols)
    rows = []
    try:
        with open(PI_SENT_LOG, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rows.append(json.loads(line))
                except Exception:
                    continue
    except OSError:
        return pd.DataFrame(columns=cols)
    if not rows:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(rows)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    if "date" not in df.columns or df["date"].isna().all():
        df["date"] = df["ts"].astype(str).str[:10]
    if "subject" not in df.columns:
        df["subject"] = ""
    return df[cols]


def last_sent_for_client(client: str, log_df: Optional[pd.DataFrame] = None) -> dict:
    """성공 발송만, 가장 최근 1건."""
    if not client:
        return {}
    df = log_df if log_df is not None else load_sent_log()
    if df is None or df.empty:
        return {}
    sub = df[df["client"].astype(str).str.strip() == str(client).strip()]
    if "ok" in sub.columns:
        sub = sub[sub["ok"].astype(str).isin(["True", "true", "1"]) | (sub["ok"] == True)]  # noqa: E712
    if sub.empty:
        return {}
    sub = sub.sort_values("ts")
    last = sub.iloc[-1].to_dict()
    return {
        "date": str(last.get("date") or str(last.get("ts") or "")[:10]),
        "subject": str(last.get("subject") or ""),
        "email": str(last.get("email") or ""),
        "ts": str(last.get("ts") or ""),
    }


def sent_summary_by_client(log_df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    df = log_df if log_df is not None else load_sent_log()
    if df is None or df.empty:
        return pd.DataFrame(columns=["거래처", "최근발송일", "최근제목", "발송횟수"])
    ok = df[df["ok"].astype(str).isin(["True", "true", "1"]) | (df["ok"] == True)].copy()  # noqa: E712
    if ok.empty:
        return pd.DataFrame(columns=["거래처", "최근발송일", "최근제목", "발송횟수"])
    ok = ok.sort_values("ts")
    rows = []
    for client, g in ok.groupby(ok["client"].astype(str).str.strip()):
        last = g.iloc[-1]
        rows.append(
            {
                "거래처": client,
                "최근발송일": str(last.get("date") or str(last.get("ts") or "")[:10]),
                "최근제목": str(last.get("subject") or ""),
                "발송횟수": int(len(g)),
            }
        )
    return pd.DataFrame(rows).sort_values("최근발송일", ascending=False).reset_index(drop=True)


def _items_key(client: str) -> str:
    return f"pi_items_{_norm_name(client) or 'none'}"


def _init_items_from_prices(client: str, price_df: pd.DataFrame, pct: float = 0.0) -> list[dict]:
    """거래처 선택 시: 거래 품목명 + 마지막 거래 단가. 인상단가는 일단 동일(계산 UI에서 적용)."""
    del pct  # 공문·초기표에 %를 자동 반영하지 않음
    items = []
    for _, row in price_df.iterrows():
        old = float(row.get("기존단가") or 0)
        items.append(
            {
                "품목명": str(row.get("품목명") or ""),
                "기존단가": old,
                "인상적용단가": old,
                "비고": "",
                "최근매출일": str(row.get("최근매출일") or ""),
            }
        )
    return items


def _items_df_for_editor(items: list[dict]) -> pd.DataFrame:
    """편집용 표: 기존거래제품명 · 단가 · 인상단가 (칸 추가/삭제·내용 수정)."""
    rows = []
    for it in items:
        old = float(it.get("기존단가") or 0)
        new = float(it.get("인상적용단가") or 0)
        rows.append(
            {
                "기존거래제품명": str(it.get("품목명") or ""),
                "단가": old,
                "단가흐름": f"{old:,.0f} ➠ {new:,.0f}",
                "인상단가": new,
                "비고": str(it.get("비고") or ""),
            }
        )
    if not rows:
        return pd.DataFrame(
            columns=["기존거래제품명", "단가", "단가흐름", "인상단가", "비고"]
        )
    return pd.DataFrame(rows)


def _items_from_editor_df(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []
    out: list[dict] = []
    for _, row in df.iterrows():
        # 구버전 컬럼명(품목명/기존단가/인상적용단가)도 허용
        name = str(
            row.get("기존거래제품명")
            if "기존거래제품명" in row.index
            else row.get("품목명")
            or ""
        ).strip()
        if not name:
            continue
        if row.get("삭제") is True:
            continue
        old = row.get("단가") if "단가" in row.index else row.get("기존단가")
        new = row.get("인상단가") if "인상단가" in row.index else row.get("인상적용단가")
        out.append(
            {
                "품목명": name,
                "기존단가": float(old or 0),
                "인상적용단가": float(new or 0),
                "비고": str(row.get("비고") or ""),
                "최근매출일": str(row.get("최근매출일") or ""),
            }
        )
    return out


def _editor_widget_key(editor_key: str) -> str:
    ver = int(st.session_state.get(f"{editor_key}_ver", 0))
    return f"{editor_key}_editor_v{ver}"


def _bump_editor_widget(editor_key: str) -> None:
    """data_editor 위젯 상태를 강제로 갱신 (일괄적용·삭제·재불러오기)."""
    vk = f"{editor_key}_ver"
    st.session_state[vk] = int(st.session_state.get(vk, 0)) + 1
    # 이전 위젯 키 잔여 상태 제거
    prev = int(st.session_state[vk]) - 1
    st.session_state.pop(f"{editor_key}_editor_v{prev}", None)
    st.session_state.pop(f"{editor_key}_editor", None)


def _apply_pct_to_items(items: list[dict], pct: float) -> list[dict]:
    out = []
    for it in items:
        old = float(it.get("기존단가") or 0)
        row = dict(it)
        row["인상적용단가"] = default_increase_price(old, pct)
        out.append(row)
    return out


def _apply_amount_to_items(items: list[dict], amount: float) -> list[dict]:
    out = []
    for it in items:
        old = float(it.get("기존단가") or 0)
        row = dict(it)
        row["인상적용단가"] = apply_increase_by_amount(old, amount)
        out.append(row)
    return out


def _default_letter_title() -> str:
    return "액체탄산 공급단가 인상 협조 요청의 건"


def _default_doc_no(client: str = "") -> str:
    return _format_doc_no(client, "")


def _default_mail_body(client: str, effective: str, items: list[dict]) -> str:
    lines = [
        f"{client} 귀중",
        "",
        "항상 저희 신일가스를 이용해 주셔서 감사드립니다.",
        "첨부와 같이 공급 단가 조정을 안내드립니다.",
        "",
        f"시행일: {_format_korean_effective(effective)}",
        "",
    ]
    for it in items:
        name = str(it.get("품목명") or "")
        old = float(it.get("기존단가") or 0)
        new = float(it.get("인상적용단가") or 0)
        lines.append(f"· {name}: {old:,.0f}원 ➠ {new:,.0f}원")
    lines.extend(["", "신일가스(주) 화성공장", "문의: 031-366-0799"])
    return "\n".join(lines)


def _build_letter_bytes(
    *,
    client: str,
    email: str,
    title: str,
    body: str,
    effective: str,
    contact: str,
    items: list[dict],
    doc_no: str = "",
    send_date: Optional[date] = None,
    letter_paras: Optional[dict[str, str]] = None,
    c37_override: str = "",
    c38_override: str = "",
) -> bytes:
    tpl = resolve_letter_template()
    return fill_letter_workbook(
        tpl,
        client=client,
        email=email,
        title=title,
        body=body,
        effective=effective,
        contact=contact,
        items=items,
        doc_no=doc_no,
        send_date=send_date,
        letter_paras=letter_paras,
        c37_override=c37_override,
        c38_override=c38_override,
    )


def _render_items_table(
    *,
    editor_key: str,
    items: list[dict],
    pct_default: float,
) -> list[dict]:
    st.markdown("##### 기존거래제품명 · 단가 ➠ 인상단가")
    st.caption(
        "거래처 선택 시 **거래 품목·마지막 거래 단가**가 자동 반영됩니다. "
        "표에서 내용 수정·행 삭제(🗑)·추가(➕) 가능. "
        "**적용 % / 인상금액은 표 계산용이며 공문에는 표시되지 않습니다.**"
    )

    mode = st.radio(
        "인상 적용 방식",
        ["퍼센테이지(%)", "인상금액(원)"],
        horizontal=True,
        key=f"{editor_key}_mode",
        help="선택한 방식으로 기존단가에 적용해 인상단가를 자동 계산합니다. 공문에는 들어가지 않습니다.",
    )
    c_val, c_apply, c_reload = st.columns([1.4, 1.1, 1.2])
    with c_val:
        if mode.startswith("퍼센트"):
            apply_val = st.number_input(
                "적용 퍼센테이지(%)",
                min_value=0.0,
                max_value=100.0,
                value=float(pct_default),
                step=0.5,
                key=f"{editor_key}_pct",
            )
        else:
            apply_val = st.number_input(
                "인상금액(원)",
                min_value=0.0,
                max_value=1_000_000.0,
                value=float(st.session_state.get(f"{editor_key}_amt", 0.0)),
                step=100.0,
                key=f"{editor_key}_amt",
            )
    with c_apply:
        st.write("")
        if st.button("인상단가에 적용", key=f"{editor_key}_apply", use_container_width=True, type="primary"):
            base = st.session_state.get(editor_key, items)
            if mode.startswith("퍼센트"):
                st.session_state[editor_key] = _apply_pct_to_items(base, apply_val)
            else:
                st.session_state[editor_key] = _apply_amount_to_items(base, apply_val)
            _bump_editor_widget(editor_key)
            st.rerun()
    with c_reload:
        st.write("")
        if st.button("매출 최종단가 다시 불러오기", key=f"{editor_key}_reload", use_container_width=True):
            st.session_state.pop(editor_key, None)
            _bump_editor_widget(editor_key)
            st.rerun()

    if editor_key not in st.session_state:
        st.session_state[editor_key] = items

    cur_items = st.session_state[editor_key]
    edit_df = _items_df_for_editor(cur_items)

    edited = st.data_editor(
        edit_df,
        column_config={
            "기존거래제품명": st.column_config.TextColumn(
                "기존거래제품명",
                help="기존 거래 제품명 (직접 수정·추가 가능)",
                required=False,
                width="medium",
            ),
            "단가": st.column_config.NumberColumn(
                "단가",
                help="마지막 거래 단가(원)",
                min_value=0.0,
                format="%.1f",
                width="small",
            ),
            "단가흐름": st.column_config.TextColumn(
                "단가 ➠ 인상단가",
                help="자동 표시",
                disabled=True,
                width="medium",
            ),
            "인상단가": st.column_config.NumberColumn(
                "인상단가",
                help="인상 적용 단가(원) — 직접 수정 가능",
                min_value=0.0,
                format="%.1f",
                width="small",
            ),
            "비고": st.column_config.TextColumn("비고", width="small"),
        },
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key=_editor_widget_key(editor_key),
    )

    if (
        edited is not None
        and not edited.empty
        and "단가" in edited.columns
        and "인상단가" in edited.columns
    ):
        edited = edited.copy()
        edited["단가흐름"] = edited.apply(
            lambda r: f"{float(r.get('단가') or 0):,.0f} ➠ {float(r.get('인상단가') or 0):,.0f}",
            axis=1,
        )

    b1, b2, b3 = st.columns(3)
    with b1:
        if st.button("빈 행 모두 제거", key=f"{editor_key}_drop_empty", use_container_width=True):
            cleaned = _items_from_editor_df(edited)
            st.session_state[editor_key] = cleaned
            _bump_editor_widget(editor_key)
            st.rerun()
    with b2:
        if st.button("표 초기화(전체 삭제)", key=f"{editor_key}_clear", use_container_width=True):
            st.session_state[editor_key] = []
            _bump_editor_widget(editor_key)
            st.rerun()
    with b3:
        if st.button("표 내용 저장", key=f"{editor_key}_save", type="primary", use_container_width=True):
            st.session_state[editor_key] = _items_from_editor_df(edited)
            st.success("품목표 저장됨 (거래처별 유지)")
            st.rerun()

    result = _items_from_editor_df(edited)
    st.session_state[editor_key] = result
    if result:
        preview = " · ".join(
            f"{it['품목명']} {float(it['기존단가']):,.0f}➠{float(it['인상적용단가']):,.0f}"
            for it in result[:5]
        )
        if len(result) > 5:
            preview += f" 외 {len(result) - 5}건"
        st.caption(f"현재 {len(result)}개 품목 — {preview}")
    return result


def _render_letter_content_editor(client: str, items: list[dict], effective_s: str) -> dict:
    """공문 양식은 유지, 내용만 수정. 기본값=단가인상공문."""
    st.markdown("##### 📄 공문 내용 (양식 유지 · 내용 편집)")
    st.caption("기본 내용은 **단가인상공문(탄산)** 입니다. 레이아웃·로고는 그대로이고 아래 텍스트만 바뀝니다.")

    d1, d2 = st.columns(2)
    with d1:
        doc_no = st.text_input(
            "문서번호",
            value=_default_doc_no(client),
            key="pi_letter_doc_no",
        )
    with d2:
        send_date = st.date_input("발송일자", value=date.today(), key="pi_letter_send_date")

    with st.expander("공문 본문 문단 편집", expanded=False):
        if "pi_letter_paras" not in st.session_state:
            st.session_state["pi_letter_paras"] = dict(_DEFAULT_LETTER_PARAS)
        paras = st.session_state["pi_letter_paras"]
        for coord in sorted(paras.keys(), key=lambda x: (int(x[1:]), x[0])):
            paras[coord] = st.text_area(
                f"본문 {coord}",
                value=paras[coord],
                height=68,
                key=f"pi_para_{coord}",
            )
        st.session_state["pi_letter_paras"] = paras
        if st.button("본문 기본값(단가인상공문)으로 되돌리기", key="pi_paras_reset"):
            st.session_state["pi_letter_paras"] = dict(_DEFAULT_LETTER_PARAS)
            for coord in _DEFAULT_LETTER_PARAS:
                st.session_state.pop(f"pi_para_{coord}", None)
            st.rerun()

    auto_c37 = _format_c37_target_items(items)
    auto_c38 = _format_c38_price_lines(items)
    with st.expander("단가 조정 내용 문구 (자동 생성 · 필요 시 수정)", expanded=True):
        c37 = st.text_area("대상 품목 줄", value=auto_c37, height=60, key="pi_c37")
        c38 = st.text_area(
            "단가 인상 금액 줄 (기존단가→인상단가만 표시, 적용%/금액 미포함)",
            value=auto_c38,
            height=60,
            key="pi_c38",
        )
        st.caption(f"시행 일자: {_format_korean_effective(effective_s)}")

    return {
        "doc_no": doc_no,
        "send_date": send_date,
        "letter_paras": dict(st.session_state.get("pi_letter_paras", _DEFAULT_LETTER_PARAS)),
        "c37_override": c37,
        "c38_override": c38,
    }


def _render_mail_settings_expander(mail_df: pd.DataFrame) -> pd.DataFrame:
    with st.expander("📇 메일 연락처 관리", expanded=False):
        st.caption(f"CSV: `{PI_MAIL_CSV}` · 거래처별 발송 이메일")
        up = st.file_uploader("연락처 CSV 업로드", type=["csv"], key="pi_mail_upload")
        if up is not None:
            try:
                raw = pd.read_csv(up, encoding="utf-8-sig")
            except Exception:
                raw = pd.read_csv(up, encoding="cp949")
            merged = _normalize_mail_df(raw)
            if not merged.empty:
                save_mail_contacts(merged)
                st.success(f"{len(merged)}건 저장")
                mail_df = load_mail_contacts()
        if not mail_df.empty:
            st.dataframe(mail_df, use_container_width=True, hide_index=True)
        else:
            st.info("연락처 CSV가 없습니다. 업로드하거나 직접 입력하세요.")
            with st.form("pi_mail_manual"):
                n = st.text_input("거래처")
                e = st.text_input("이메일")
                if st.form_submit_button("추가"):
                    if n and e:
                        add = pd.DataFrame([{"거래처": n, "이메일": e, "비고": ""}])
                        out = pd.concat([mail_df, add], ignore_index=True)
                        save_mail_contacts(out)
                        st.rerun()
    return mail_df


def _render_smtp_bar() -> dict:
    cfg = smtp_settings()
    c1, c2, c3 = st.columns([2.5, 1.2, 1.2])
    with c1:
        st.caption(f"SMTP · {smtp_status_label(cfg)}")
    with c2:
        if st.button("SMTP 연결 테스트", key="pi_smtp_test", use_container_width=True):
            ok, msg = test_smtp_connection(cfg)
            (st.success if ok else st.error)(msg)
    with c3:
        tpl = resolve_letter_template()
        st.caption(f"양식: `{os.path.basename(tpl)}`")
    return cfg


def render_price_increase_tab(sales_df: pd.DataFrame, latest_update_str: str = "") -> None:
    """단가인상 탭 — 개별·일괄 발송, 품목표, 발송 이력."""
    _ensure_dirs()
    st.markdown(
        "<div class='sub-header dashboard-tab-panel-head'>📨 단가인상 공문</div>",
        unsafe_allow_html=True,
    )
    cap = f"빌드 {PI_UI_BUILD}"
    if latest_update_str:
        cap += f" · 매출 {latest_update_str}"
    st.caption(cap)

    mail_df = load_mail_contacts()
    log_df = load_sent_log()
    smtp_cfg = _render_smtp_bar()
    mail_df = _render_mail_settings_expander(mail_df)

    tab_single, tab_bulk, tab_hist = st.tabs(["개별 발송", "일괄 발송", "발송 이력"])

    # ── 개별 발송 ──
    with tab_single:
        staff_opts = ["전체"] + list_staff_options(sales_df)
        r1c1, r1c2, r1c3 = st.columns([1.2, 2, 1.5])
        with r1c1:
            staff = st.selectbox("담당자", staff_opts, key="pi_single_staff")
        clients = list_clients_for_staff(sales_df, staff)
        with r1c2:
            client = st.selectbox("거래처", clients or [""], key="pi_single_client")
        with r1c3:
            kind = classify_client_kind(client, sales_df) if client else ""
            st.caption(f"유형: **{kind}**" if kind else "")

        if not client:
            st.info("거래처를 선택하세요.")
        else:
            last = last_sent_for_client(client, log_df)
            if last:
                st.caption(
                    f"최근 발송: **{last.get('date')}** · 제목: {last.get('subject') or '-'} "
                    f"({last.get('email') or '-'})"
                )
            else:
                st.caption("최근 발송 이력 없음")

            email_default = lookup_email(client, mail_df)
            email = st.text_input("수신 이메일", value=email_default, key="pi_single_email")

            r2c1, r2c2, r2c3 = st.columns([2, 1, 1])
            with r2c1:
                title = st.text_input(
                    "메일·공문 제목",
                    value=_default_letter_title(),
                    key="pi_single_title",
                )
            with r2c2:
                effective = st.date_input(
                    "시행일",
                    value=date.today().replace(day=1),
                    key="pi_single_eff",
                )
            with r2c3:
                contact = st.text_input("문의", value="031-366-0799", key="pi_single_contact")
            effective_s = (
                effective.strftime("%Y-%m-%d")
                if hasattr(effective, "strftime")
                else str(effective)
            )

            # 거래처·공문 입력 하단: 기존거래제품명 / 단가 ➠ 인상단가 편집 표
            price_df = latest_unit_prices(sales_df, client)
            pct_key = "pi_global_pct"
            if pct_key not in st.session_state:
                st.session_state[pct_key] = 5.0
            items_key = _items_key(client)
            # 거래처별 session_state 유지 — 재선택 시에도 편집 내용 보존
            if items_key not in st.session_state:
                st.session_state[items_key] = _init_items_from_prices(
                    client, price_df, st.session_state[pct_key]
                )
            st.session_state["pi_last_client"] = client

            items = _render_items_table(
                editor_key=items_key,
                items=st.session_state[items_key],
                pct_default=st.session_state[pct_key],
            )
            st.session_state[pct_key] = float(
                st.session_state.get(f"{items_key}_pct", st.session_state[pct_key])
            )

            if not items:
                st.warning("품목이 없습니다. 표에서 행을 추가하거나 매출 데이터를 확인하세요.")

            letter_meta = _render_letter_content_editor(client, items, effective_s)

            body = st.text_area(
                "메일 본문",
                value=_default_mail_body(client, effective_s, items),
                height=160,
                key="pi_single_body",
            )

            dl_name = f"단가인상_{client}_{date.today().strftime('%Y%m%d')}.xlsx"
            letter_kwargs = dict(
                client=client,
                email=email,
                title=title,
                body=body,
                effective=effective_s,
                contact=contact,
                items=items,
                doc_no=letter_meta.get("doc_no") or "",
                send_date=letter_meta.get("send_date"),
                letter_paras=letter_meta.get("letter_paras"),
                c37_override=letter_meta.get("c37_override") or "",
                c38_override=letter_meta.get("c38_override") or "",
            )
            if items:
                xlsx = _build_letter_bytes(**letter_kwargs)
                st.download_button(
                    "📥 공문 엑셀 미리보기",
                    data=xlsx,
                    file_name=dl_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="pi_single_dl",
                )

            if st.button("📧 공문 메일 발송", type="primary", key="pi_single_send", disabled=not items):
                if not email:
                    st.error("수신 이메일을 입력하세요.")
                elif not smtp_cfg.get("ready"):
                    st.error("SMTP secrets 미설정")
                else:
                    xlsx = _build_letter_bytes(**letter_kwargs)
                    ok, msg = send_mail_smtp(
                        to_addr=email,
                        subject=title,
                        body=body,
                        attachment_bytes=xlsx,
                        attachment_name=dl_name,
                    )
                    append_sent_log(
                        client=client,
                        email=email,
                        subject=title,
                        ok=ok,
                        mode="single",
                        staff=staff if staff != "전체" else "",
                        items=len(items),
                        msg=msg,
                    )
                    (st.success if ok else st.error)(msg)

    # ── 일괄 발송 (담당자 단위 · 거래처는 한 곳씩 개별 공문) ──
    with tab_bulk:
        st.caption(
            "담당자를 먼저 고른 뒤, 해당 담당자 거래처를 **기본 전체 선택**합니다. "
            "제외할 곳만 체크 해제하세요. 발송은 **한 거래처씩** 개별 공문으로 진행됩니다. "
            "(회사 전체 일괄 발송은 없습니다.)"
        )
        staff_list = list_staff_options(sales_df)
        if not staff_list:
            st.warning("담당자 목록이 없습니다. 매출 데이터를 확인하세요.")
            b_staff = ""
        else:
            b_staff = st.selectbox(
                "담당자 (필수)",
                staff_list,
                key="pi_bulk_staff",
                help="일괄은 담당자 단위입니다. 전체 거래처 일괄은 지원하지 않습니다.",
            )
        bulk_clients = list_clients_for_staff(sales_df, b_staff) if b_staff else []

        mode_bulk = st.radio(
            "공통 인상 적용 (표 계산용 · 공문에 %/금액 미표시)",
            ["퍼센테이지(%)", "인상금액(원)"],
            horizontal=True,
            key="pi_bulk_mode",
        )
        if mode_bulk.startswith("퍼센트"):
            pct_bulk = st.number_input(
                "공통 적용 퍼센테이지(%)",
                min_value=0.0,
                max_value=100.0,
                value=float(st.session_state.get("pi_global_pct", 5.0)),
                step=0.5,
                key="pi_bulk_pct",
            )
            amt_bulk = 0.0
            st.session_state["pi_global_pct"] = pct_bulk
        else:
            amt_bulk = st.number_input(
                "공통 인상금액(원)",
                min_value=0.0,
                max_value=1_000_000.0,
                value=0.0,
                step=100.0,
                key="pi_bulk_amt",
            )
            pct_bulk = 0.0

        title_bulk = st.text_input("공통 제목", value=_default_letter_title(), key="pi_bulk_title")
        eff_bulk = st.date_input("공통 시행일", value=date.today().replace(day=1), key="pi_bulk_eff")
        eff_bulk_s = eff_bulk.strftime("%Y-%m-%d") if hasattr(eff_bulk, "strftime") else str(eff_bulk)

        # 담당자 바뀌면 선택 위젯 리셋 → 다시 전체 선택
        prev_staff = st.session_state.get("pi_bulk_staff_prev")
        if prev_staff != b_staff:
            st.session_state["pi_bulk_staff_prev"] = b_staff
            st.session_state.pop("pi_bulk_select", None)
            ver = int(st.session_state.get("pi_bulk_sel_ver", 0)) + 1
            st.session_state["pi_bulk_sel_ver"] = ver

        rows = []
        for cl in bulk_clients:
            em = lookup_email(cl, mail_df)
            last = last_sent_for_client(cl, log_df)
            rows.append(
                {
                    "발송": True,  # 담당자 내 거래처 전체 선택 기본
                    "거래처": cl,
                    "이메일": em or "(미등록)",
                    "유형": classify_client_kind(cl, sales_df),
                    "최근발송일": last.get("date", "") if last else "",
                    "최근제목": last.get("subject", "") if last else "",
                }
            )
        bulk_df = pd.DataFrame(rows)
        if not b_staff:
            st.info("담당자를 선택하세요.")
        elif bulk_df.empty:
            st.info("해당 담당자 거래처가 없습니다.")
        else:
            sel1, sel2, sel3 = st.columns(3)
            with sel1:
                if st.button("담당자 거래처 전체 선택", key="pi_bulk_all_on", use_container_width=True):
                    st.session_state["pi_bulk_force_all"] = True
                    st.session_state["pi_bulk_force_none"] = False
                    st.session_state["pi_bulk_sel_ver"] = int(st.session_state.get("pi_bulk_sel_ver", 0)) + 1
                    st.session_state.pop("pi_bulk_select", None)
                    st.rerun()
            with sel2:
                if st.button("전체 제외", key="pi_bulk_all_off", use_container_width=True):
                    st.session_state["pi_bulk_force_all"] = False
                    st.session_state["pi_bulk_force_none"] = True
                    st.session_state["pi_bulk_sel_ver"] = int(st.session_state.get("pi_bulk_sel_ver", 0)) + 1
                    st.session_state.pop("pi_bulk_select", None)
                    st.rerun()
            with sel3:
                st.caption("체크 해제로 개별 제외")

            if st.session_state.get("pi_bulk_force_none"):
                bulk_df["발송"] = False
            elif st.session_state.get("pi_bulk_force_all", True):
                bulk_df["발송"] = True
            # force 플래그는 한 번 반영 후 유지 위해 데이터에만 적용; 위젯 키는 ver로 갱신

            edited_bulk = st.data_editor(
                bulk_df,
                column_config={
                    "발송": st.column_config.CheckboxColumn(
                        "발송",
                        help="기본 전체 선택. 제외할 거래처만 체크 해제",
                        default=True,
                    ),
                    "거래처": st.column_config.TextColumn("거래처", disabled=True),
                    "이메일": st.column_config.TextColumn("이메일", disabled=True),
                    "유형": st.column_config.TextColumn("유형", disabled=True),
                    "최근발송일": st.column_config.TextColumn("최근발송일", disabled=True),
                    "최근제목": st.column_config.TextColumn("최근제목", disabled=True),
                },
                hide_index=True,
                use_container_width=True,
                key=f"pi_bulk_select_v{int(st.session_state.get('pi_bulk_sel_ver', 0))}",
            )
            # 발송 체크된 곳 — 이메일은 실제 발송 시 검증 (한 거래처씩)
            checked = edited_bulk[edited_bulk["발송"] == True].copy()  # noqa: E712
            has_mail = checked["이메일"].astype(str).str.contains("@", na=False)
            targets = checked[has_mail]
            no_mail = checked[~has_mail]
            st.caption(
                f"담당자 **{b_staff}** · 선택 {len(checked)}/{len(bulk_df)}곳 "
                f"· 발송가능(메일) **{len(targets)}**곳 · 메일미등록 {len(no_mail)}곳"
            )
            if not no_mail.empty:
                st.warning(
                    "메일 미등록 거래처는 건너뜁니다: "
                    + ", ".join(no_mail["거래처"].astype(str).head(8).tolist())
                    + (" …" if len(no_mail) > 8 else "")
                )

            confirm = st.checkbox(
                f"선택 거래처를 **한 곳씩** 개별 공문으로 발송합니다 ({len(targets)}곳, 되돌릴 수 없음)",
                key="pi_bulk_confirm",
            )
            if st.button(
                "📧 담당자 일괄 발송 (거래처별 개별)",
                type="primary",
                key="pi_bulk_send",
                disabled=not confirm or targets.empty,
            ):
                if not smtp_cfg.get("ready"):
                    st.error("SMTP secrets 미설정")
                else:
                    prog = st.progress(0.0)
                    status = st.empty()
                    results: list[str] = []
                    n = len(targets)
                    for i, (_, row) in enumerate(targets.iterrows()):
                        cl = str(row["거래처"])
                        em = str(row["이메일"]).strip()
                        status.info(f"({i + 1}/{n}) **{cl}** 개별 공문 발송 중…")
                        pdf = latest_unit_prices(sales_df, cl)
                        items_b = _init_items_from_prices(cl, pdf, 0.0)
                        if mode_bulk.startswith("퍼센트"):
                            items_b = _apply_pct_to_items(items_b, pct_bulk)
                        else:
                            items_b = _apply_amount_to_items(items_b, amt_bulk)
                        body_b = _default_mail_body(cl, eff_bulk_s, items_b)
                        dl = f"단가인상_{cl}_{date.today().strftime('%Y%m%d')}.xlsx"
                        xlsx_b = _build_letter_bytes(
                            client=cl,
                            email=em,
                            title=title_bulk,
                            body=body_b,
                            effective=eff_bulk_s,
                            contact="031-366-0799",
                            items=items_b,
                        )
                        ok, msg = send_mail_smtp(
                            to_addr=em,
                            subject=title_bulk,
                            body=body_b,
                            attachment_bytes=xlsx_b,
                            attachment_name=dl,
                        )
                        append_sent_log(
                            client=cl,
                            email=em,
                            subject=title_bulk,
                            ok=ok,
                            mode="bulk",
                            staff=b_staff,
                            items=len(items_b),
                            msg=msg,
                        )
                        results.append(f"{'✅' if ok else '❌'} {cl}: {msg}")
                        prog.progress((i + 1) / max(n, 1))
                    status.empty()
                    for line in results:
                        st.write(line)
                    st.success(f"담당자 [{b_staff}] 거래처별 개별 발송 완료 ({n}건)")

    # ── 발송 이력 ──
    with tab_hist:
        st.caption("거래처별 최근 발송일·제목 및 전체 로그")
        summary = sent_summary_by_client(log_df)
        if summary.empty:
            st.info("발송 이력이 없습니다.")
        else:
            st.markdown("##### 거래처별 최근 발송")
            st.dataframe(summary, use_container_width=True, hide_index=True)

        if not log_df.empty:
            show = log_df.copy()
            show["ok"] = show["ok"].map(lambda x: "✅" if str(x) in ("True", "true", "1") or x is True else "❌")
            show = show.rename(
                columns={
                    "date": "발송일",
                    "client": "거래처",
                    "email": "이메일",
                    "subject": "제목",
                    "mode": "모드",
                    "staff": "담당자",
                    "items": "품목수",
                    "msg": "결과",
                    "ts": "시각",
                }
            )
            st.markdown("##### 전체 이력")
            st.dataframe(
                show[["발송일", "거래처", "제목", "이메일", "ok", "모드", "담당자", "품목수", "결과", "시각"]],
                use_container_width=True,
                hide_index=True,
            )

