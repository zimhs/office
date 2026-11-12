"""견적서 — 로컬 전용.

화면 왼쪽은 견적서, 오른쪽은 품목 단가 검색이다.
저장은 신일가스_견적서.xlsx에 시트를 하나 추가하고, 그 시트에 값만 넣는다.
"""
from __future__ import annotations

import os
import re
import shutil
import unicodedata
import zipfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

_BAD_SHEET = re.compile(r"[:\\/?*\[\]]")
_PRICE_FROM = pd.Timestamp("2020-01-01")
_LINE_ROWS = tuple(range(19, 27))
_NOTE_ROWS = (34, 35, 36, 37)
_EXCEL_EPOCH = date(1899, 12, 30)
_UNTOUCHED = (
    "xl/styles.xml",
    "xl/theme/theme1.xml",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/worksheets/_rels/sheet2.xml.rels",
    "xl/drawings/drawing1.xml",
    "xl/drawings/drawing2.xml",
    "xl/media/image1.png",
    "xl/media/image2.png",
)


def _nfc(text: str) -> str:
    return unicodedata.normalize("NFC", str(text or ""))


def _xml_text(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _num_text(value: float) -> str:
    number = float(value)
    if abs(number - round(number)) < 1e-6:
        return str(int(round(number)))
    return f"{number:.6f}".rstrip("0").rstrip(".")


def _excel_serial(day: date) -> int:
    return (day - _EXCEL_EPOCH).days


def sheet_title(client: str, on: date, taken: set[str]) -> str:
    """워크시트 제목은 상호.작성날짜. 같은 이름이 있으면 번호를 붙인다."""
    name = _BAD_SHEET.sub(" ", _nfc(client)).strip()
    name = re.sub(r"\s+", " ", name).strip(" .'")
    if not name:
        raise ValueError("상호를 입력해 주세요.")
    day = on.strftime("%Y-%m-%d")
    suffix = f".{day}"

    def fit(extra: str = "") -> str:
        room = 31 - len(suffix) - len(extra)
        head = name[: max(room, 1)]
        return (head + suffix + extra)[:31]

    base = fit()
    if base not in taken:
        return base
    number = 2
    while number < 100:
        title = fit(f" ({number})")
        if title not in taken:
            return title
        number += 1
    raise ValueError("같은 이름의 시트가 너무 많습니다.")


def find_quote_dir() -> Path | None:
    """Drive › 다른 컴퓨터 › 내 컴퓨터 (1) › Desktop › 업무 › 공사 › 견적서."""
    try:
        cached = st.session_state.get("_qt_quote_dir")
    except Exception:
        cached = None
    if isinstance(cached, str) and cached and Path(cached).is_dir():
        return Path(cached)
    found = _scan_quote_dir()
    if found is not None:
        try:
            st.session_state["_qt_quote_dir"] = str(found)
        except Exception:
            pass
    return found


def _scan_quote_dir() -> Path | None:
    base = Path.home() / "Library" / "CloudStorage"
    if not base.is_dir():
        return None
    for drive in base.iterdir():
        if not drive.name.startswith("GoogleDrive-"):
            continue
        other = drive / "다른 컴퓨터"
        if not other.is_dir():
            continue
        for comp in other.iterdir():
            if "내 컴퓨터 (1)" not in _nfc(comp.name):
                continue
            work = comp / "Desktop" / "업무"
            found = []
            for mid in ("회사", "공사"):
                folder = work / mid / "견적서"
                if folder.is_dir():
                    found.append(folder)
            for folder in found:
                if _workbook_in(folder) is not None:
                    return folder
            if found:
                return found[0]
    return None


def _workbook_in(folder: Path) -> Path | None:
    if not folder.is_dir():
        return None
    for path in folder.iterdir():
        if not path.is_file() or path.name.startswith("~$"):
            continue
        title = _nfc(path.name)
        if path.suffix.lower() == ".xlsx" and "신일" in title and "견적" in title:
            return path
    return None


def find_template() -> Path | None:
    desktop = Path.home() / "Desktop" / "신일가스_견적서.xlsx"
    if desktop.is_file():
        return desktop
    folder = _scan_quote_dir()
    if folder is not None:
        return _workbook_in(folder)
    return None


def quote_workbook() -> Path:
    folder = find_quote_dir()
    if folder is None:
        raise FileNotFoundError("견적서 폴더를 찾지 못했습니다.")
    existing = _workbook_in(folder)
    if existing is not None:
        return existing
    template = find_template()
    if template is None:
        raise FileNotFoundError("견적서 양식을 찾지 못했습니다.")
    dest = folder / "신일가스_견적서.xlsx"
    shutil.copy(template, dest)
    return dest


def _shared_strings(xml: str) -> list[str]:
    out = []
    for si in re.findall(r"<si>(.*?)</si>", xml, re.S):
        parts = re.findall(r"<t(?: [^>]*)?>(.*?)</t>", si, re.S)
        text = "".join(parts) if parts else ""
        text = (
            text.replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&amp;", "&")
        )
        out.append(text)
    return out


def _append_shared(sst: str, text: str, strings: list[str]) -> tuple[str, int]:
    count = re.search(r'\bcount="(\d+)"', sst)
    if count:
        sst = sst.replace(count.group(0), f'count="{int(count.group(1)) + 1}"', 1)
    if text in strings:
        return sst, strings.index(text)
    idx = len(strings)
    strings.append(text)
    preserve = text[:1].isspace() or text[-1:].isspace() or "\n" in text
    space = ' xml:space="preserve"' if preserve else ""
    piece = f"<si><t{space}>{_xml_text(text)}</t></si>"
    if "</sst>" not in sst:
        raise ValueError("sharedStrings에 끝 태그가 없습니다.")
    sst = sst.replace("</sst>", piece + "</sst>", 1)
    sst = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{idx + 1}"', sst, count=1)
    return sst, idx


def _cell_xml(xml: str, ref: str) -> re.Match[str] | None:
    return re.search(
        rf'<c r="{re.escape(ref)}"(?:[^>]*/>|[^>]*>.*?</c>)',
        xml,
        re.S,
    )


def _open_body(open_tag: str) -> str:
    body = open_tag[2:-1]
    return re.sub(r'\s+t="[^"]*"', "", body)


def _write_cell(xml: str, ref: str, value: object, strings: list[str], sst: str) -> tuple[str, str]:
    found = _cell_xml(xml, ref)
    if not found:
        return xml, sst
    tag = found.group(0)
    if tag.endswith("/>"):
        open_tag = tag[:-2] + ">"
        inner = ""
    else:
        split = tag.find(">")
        open_tag, inner = tag[: split + 1], tag[split + 1 : -4]
    if "<f" in inner:
        return xml, sst
    body = _open_body(open_tag)
    if value is None or value == "":
        replacement = f"<c{body}/>"
    elif isinstance(value, str):
        sst, idx = _append_shared(sst, value, strings)
        replacement = f'<c{body} t="s"><v>{idx}</v></c>'
    else:
        replacement = f"<c{body}><v>{_num_text(float(value))}</v></c>"
    xml = xml[: found.start()] + replacement + xml[found.end() :]
    return xml, sst


def _write_formula_cache(xml: str, ref: str, value: object) -> str:
    found = _cell_xml(xml, ref)
    if not found:
        return xml
    tag = found.group(0)
    if tag.endswith("/>"):
        return xml
    split = tag.find(">")
    open_tag, inner = tag[: split + 1], tag[split + 1 : -4]
    formula = re.search(r"<f\b.*?</f>|<f\b[^>]*/>", inner, re.S)
    if not formula:
        return xml
    body = _open_body(open_tag)
    if isinstance(value, str):
        cached = "<v/>" if value == "" else f"<v>{_xml_text(value)}</v>"
        replacement = f'<c{body} t="str">{formula.group(0)}{cached}</c>'
    else:
        replacement = f"<c{body}>{formula.group(0)}<v>{_num_text(float(value))}</v></c>"
    return xml[: found.start()] + replacement + xml[found.end() :]


def _line_amount(qty: float | None, price: float | None) -> float | None:
    if qty is None or price is None:
        return None
    return float(qty) * float(price)


def build_sheet(template_sheet: str, sst: str, quote: dict) -> tuple[str, str]:
    strings = _shared_strings(sst)
    xml = template_sheet
    texts = {
        "C10": quote.get("client") or "",
        "C11": quote.get("person") or "",
        "C12": quote.get("phone") or "",
        "C13": quote.get("address") or "",
        "C14": quote.get("ref") or "",
        "B17": quote.get("number") or "",
        "H17": quote.get("delivery") or "",
        "D29": quote.get("pay") or "",
        "D30": quote.get("vat") or "별도",
    }
    notes = list(quote.get("notes") or [])
    while len(notes) < 4:
        notes.append("")
    for row, note in zip(_NOTE_ROWS, notes[:4]):
        texts[f"C{row}"] = note or ""
    lines = list(quote.get("lines") or [])
    while len(lines) < 8:
        lines.append({})
    amounts: list[float | None] = []
    for row, line in zip(_LINE_ROWS, lines[:8]):
        texts[f"C{row}"] = str(line.get("품목") or "")
        texts[f"D{row}"] = str(line.get("규격") or "")
        texts[f"F{row}"] = str(line.get("단위") or "")
        texts[f"I{row}"] = str(line.get("비고") or "")
        qty = line.get("수량")
        price = line.get("단가")
        xml, sst = _write_cell(xml, f"E{row}", qty, strings, sst)
        xml, sst = _write_cell(xml, f"G{row}", price, strings, sst)
        amounts.append(_line_amount(qty if isinstance(qty, (int, float)) else None, price if isinstance(price, (int, float)) else None))
    for ref, value in texts.items():
        xml, sst = _write_cell(xml, ref, value, strings, sst)
    on = quote.get("quote_date") or date.today()
    xml, sst = _write_cell(xml, "D17", _excel_serial(on), strings, sst)
    for row, amount in zip(_LINE_ROWS, amounts):
        xml = _write_formula_cache(xml, f"H{row}", "" if amount is None else amount)
    supply = sum(amount or 0 for amount in amounts)
    vat_mode = str(quote.get("vat") or "별도")
    vat_amt = 0.0 if vat_mode == "포함" else supply * 0.1
    total = supply if vat_mode == "포함" else supply + vat_amt
    xml = _write_formula_cache(xml, "H29", supply)
    xml = _write_formula_cache(xml, "H30", vat_amt)
    xml = _write_formula_cache(xml, "H31", total)
    valid = (on + timedelta(days=7)).strftime("%Y-%m-%d") + " 까지"
    xml = _write_formula_cache(xml, "F17", valid)
    note = "※ 부가세 포함 단가" if vat_mode == "포함" else "※ 부가세 별도 · 총 청구액은 우측 합계"
    xml = _write_formula_cache(xml, "B31", note)
    xml = xml.replace('tabSelected="1"', 'tabSelected="0"', 1)
    return xml, sst


def _sheet_names(workbook_xml: str) -> list[str]:
    return re.findall(r'<sheet [^>]*name="([^"]*)"', workbook_xml)


def _next_id(text: str, pattern: str) -> int:
    found = [int(n) for n in re.findall(pattern, text)]
    return (max(found) if found else 0) + 1


def _next_part(names: list[str], stem: str) -> int:
    found = []
    for name in names:
        match = re.search(rf"{re.escape(stem)}(\d+)\.xml$", name)
        if match:
            found.append(int(match.group(1)))
    return (max(found) if found else 0) + 1


def _add_sheet_package(parts: dict[str, bytes], title: str, sheet_xml: str, sst: str) -> None:
    workbook = parts["xl/workbook.xml"].decode("utf-8")
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    types = parts["[Content_Types].xml"].decode("utf-8")
    names = list(parts)
    sheet_no = _next_part(names, "xl/worksheets/sheet")
    drawing_no = _next_part(names, "xl/drawings/drawing")
    rid = _next_id(rels, r'Id="rId(\d+)"')
    sid = _next_id(workbook, r'sheetId="(\d+)"')
    local_id = len(_sheet_names(workbook))
    quoted = "'" + title.replace("'", "''") + "'"
    sheet_xml = re.sub(
        r'xr:uid="\{[^"]+\}"',
        f'xr:uid="{{00000000-0001-0000-{sheet_no:04X}-000000000000}}"',
        sheet_xml,
        count=1,
    )
    workbook = workbook.replace(
        "</sheets>",
        f'<sheet name="{_xml_text(title)}" sheetId="{sid}" r:id="rId{rid}"/></sheets>',
        1,
    )
    defined = (
        f'<definedName name="_xlnm.Print_Area" localSheetId="{local_id}">'
        f"{quoted}!$A$1:$I$41</definedName>"
    )
    if "</definedNames>" in workbook:
        workbook = workbook.replace("</definedNames>", defined + "</definedNames>", 1)
    if "fullCalcOnLoad=" not in workbook and "<calcPr" in workbook:
        workbook = workbook.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
    rels = rels.replace(
        "</Relationships>",
        (
            f'<Relationship Id="rId{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{sheet_no}.xml"/></Relationships>'
        ),
        1,
    )
    types = types.replace(
        "</Types>",
        (
            f'<Override PartName="/xl/worksheets/sheet{sheet_no}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            f'<Override PartName="/xl/drawings/drawing{drawing_no}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            "</Types>"
        ),
        1,
    )
    sheet_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
        f'Target="../drawings/drawing{drawing_no}.xml"/></Relationships>'
    )
    parts["xl/workbook.xml"] = workbook.encode("utf-8")
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    parts["[Content_Types].xml"] = types.encode("utf-8")
    parts["xl/sharedStrings.xml"] = sst.encode("utf-8")
    parts[f"xl/worksheets/sheet{sheet_no}.xml"] = sheet_xml.encode("utf-8")
    parts[f"xl/worksheets/_rels/sheet{sheet_no}.xml.rels"] = sheet_rels.encode("utf-8")
    drawing = parts.get("xl/drawings/drawing1.xml")
    drawing_rels = parts.get("xl/drawings/_rels/drawing1.xml.rels")
    if drawing is not None:
        parts[f"xl/drawings/drawing{drawing_no}.xml"] = drawing
    if drawing_rels is not None:
        parts[f"xl/drawings/_rels/drawing{drawing_no}.xml.rels"] = drawing_rels
    app = parts.get("docProps/app.xml")
    if app is not None:
        parts["docProps/app.xml"] = _bump_app(app.decode("utf-8"), title).encode("utf-8")


def _bump_app(xml: str, title: str) -> str:
    match = re.search(
        r"<TitlesOfParts><vt:vector size=\"\d+\" baseType=\"lpstr\">(.*?)</vt:vector></TitlesOfParts>",
        xml,
        re.S,
    )
    if not match:
        return xml
    names = re.findall(r"<vt:lpstr>(.*?)</vt:lpstr>", match.group(1))
    sheets = [name for name in names if "!Print_Area" not in name]
    areas = [name for name in names if "!Print_Area" in name]
    sheets.append(_xml_text(title))
    areas.append(_xml_text(title) + "!Print_Area")
    body = "".join(f"<vt:lpstr>{name}</vt:lpstr>" for name in sheets + areas)
    xml = xml[: match.start(1)] + body + xml[match.end(1) :]
    xml = re.sub(
        r'(<TitlesOfParts><vt:vector size=")\d+(")',
        rf"\g<1>{len(sheets) + len(areas)}\2",
        xml,
        count=1,
    )
    xml = re.sub(
        r"(<vt:lpstr>워크시트</vt:lpstr></vt:variant><vt:variant><vt:i4>)\d+(</vt:i4>)",
        lambda found: found.group(1) + str(len(sheets)) + found.group(2),
        xml,
        count=1,
    )
    xml = re.sub(
        r"(<vt:lpstr>이름 지정된 범위</vt:lpstr></vt:variant><vt:variant><vt:i4>)\d+(</vt:i4>)",
        lambda found: found.group(1) + str(len(areas)) + found.group(2),
        xml,
        count=1,
    )
    return xml


def _write_package(path: Path, parts: dict[str, bytes], order: list[str]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        seen = set()
        for item in zin.infolist():
            seen.add(item.filename)
            data = parts.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
        for name in order:
            if name in seen or name not in parts:
                continue
            info = zipfile.ZipInfo(filename=name)
            info.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(info, parts[name])
    os.replace(tmp, path)


def save_quote(quote: dict, path: Path | None = None) -> tuple[Path, str]:
    """양식 시트는 그대로 두고, 상호.작성날짜 시트를 추가한 뒤 값만 넣는다."""
    book = path if path is not None else quote_workbook()
    if not book.is_file():
        template = find_template()
        if template is None:
            raise FileNotFoundError("견적서 양식을 찾지 못했습니다.")
        book.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(template, book)
    with zipfile.ZipFile(book, "r") as zin:
        parts = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        order = [item.filename for item in zin.infolist()]
    before = {name: parts[name] for name in _UNTOUCHED if name in parts}
    workbook = parts["xl/workbook.xml"].decode("utf-8")
    title = sheet_title(
        str(quote.get("client") or ""),
        quote.get("quote_date") or date.today(),
        set(_sheet_names(workbook)),
    )
    template_sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    sst = parts["xl/sharedStrings.xml"].decode("utf-8")
    sheet_xml, sst = build_sheet(template_sheet, sst, quote)
    _add_sheet_package(parts, title, sheet_xml, sst)
    for name, blob in before.items():
        if parts.get(name) != blob:
            raise RuntimeError(f"서식이 바뀌었습니다: {name}")
    new_names = [name for name in parts if name not in order]
    _write_package(book, parts, order + new_names)
    return book, title


def search_item_prices(df: pd.DataFrame | None, query: str, limit: int | None = None) -> pd.DataFrame:
    """2020년 이후 전체 거래에서, 품목 검색 결과의 거래처별 최근 단가 한 건만 돌려준다."""
    columns = ["날짜", "단가", "품목", "거래처", "출고량"]
    empty = pd.DataFrame(columns=columns)
    words = [word.casefold() for word in str(query or "").split() if word.strip()]
    if not words or df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return empty
    if "품목명" not in df.columns:
        return empty
    names = df["품목명"].astype(str).str.casefold()
    mask = pd.Series(True, index=df.index)
    for word in words:
        mask = mask & names.str.contains(word, regex=False)
    hit = df.loc[mask].copy()
    if hit.empty:
        return empty
    if "매출일_dt" in hit.columns:
        hit["_d"] = pd.to_datetime(hit["매출일_dt"], errors="coerce")
    elif "매출일" in hit.columns:
        hit["_d"] = pd.to_datetime(hit["매출일"], errors="coerce")
    else:
        return empty
    hit = hit.loc[hit["_d"] >= _PRICE_FROM]
    if hit.empty:
        return empty
    if "거래처" in hit.columns:
        hit["_client"] = hit["거래처"].map(lambda value: "" if pd.isna(value) else str(value).strip())
    else:
        hit["_client"] = ""
    hit["_i"] = range(len(hit))
    hit = hit.sort_values(
        ["_d", "_i"],
        ascending=[False, False],
        kind="mergesort",
        na_position="last",
    )
    hit = hit.drop_duplicates("_client", keep="first")
    if limit:
        hit = hit.head(int(limit))
    return pd.DataFrame(
        {
            "날짜": hit["_d"].dt.strftime("%Y-%m-%d").fillna(""),
            "단가": hit["단가"] if "단가" in hit.columns else None,
            "품목": hit["품목명"].astype(str),
            "거래처": hit["거래처"].astype(str) if "거래처" in hit.columns else "",
            "출고량": hit["출고량"] if "출고량" in hit.columns else None,
        }
    )


def _text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    return str(value).strip()


def _number(value: object) -> float | None:
    if value is None or value == "":
        return None
    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _defaults() -> dict:
    path = find_template()
    blank = {
        "client": "",
        "person": "",
        "phone": "",
        "address": "",
        "ref": "",
        "number": "",
        "delivery": "협의",
        "pay": "익월 말 현금",
        "vat": "별도",
        "notes": ["", "", "", ""],
        "lines": [
            {"품목": "", "규격": "", "수량": None, "단위": "", "단가": None, "비고": ""}
            for _ in range(8)
        ],
    }
    if path is None:
        return blank
    try:
        import openpyxl

        wb = openpyxl.load_workbook(path, data_only=False)
    except Exception:
        return blank
    try:
        ws = wb[wb.sheetnames[0]]

        def cell(ref: str) -> str:
            value = ws[ref].value
            if value is None or isinstance(value, (int, float)):
                return "" if not isinstance(value, str) else value.strip()
            return str(value).strip()

        lines = []
        for row in _LINE_ROWS:
            qty = ws[f"E{row}"].value
            price = ws[f"G{row}"].value
            lines.append(
                {
                    "품목": cell(f"C{row}"),
                    "규격": cell(f"D{row}"),
                    "수량": float(qty) if isinstance(qty, (int, float)) else None,
                    "단위": cell(f"F{row}"),
                    "단가": float(price) if isinstance(price, (int, float)) else None,
                    "비고": cell(f"I{row}"),
                }
            )
        vat = cell("D30") or "별도"
        if vat not in ("별도", "포함"):
            vat = "별도"
        return {
            "client": cell("C10"),
            "person": cell("C11"),
            "phone": cell("C12"),
            "address": cell("C13"),
            "ref": cell("C14"),
            "number": cell("B17"),
            "delivery": cell("H17") or "협의",
            "pay": cell("D29"),
            "vat": vat,
            "notes": [cell(f"C{row}") for row in _NOTE_ROWS],
            "lines": lines,
        }
    finally:
        wb.close()


def _cached_defaults() -> dict:
    try:
        cached = st.session_state.get("_qt_defaults")
    except Exception:
        cached = None
    if isinstance(cached, dict):
        return cached
    loaded = _defaults()
    try:
        st.session_state["_qt_defaults"] = loaded
    except Exception:
        pass
    return loaded


def _ensure_form() -> None:
    if st.session_state.get("_qt_form_ready"):
        return
    seed = _cached_defaults()
    st.session_state.setdefault("qt_client", seed["client"])
    st.session_state.setdefault("qt_person", seed["person"])
    st.session_state.setdefault("qt_phone", seed["phone"])
    st.session_state.setdefault("qt_address", seed["address"])
    st.session_state.setdefault("qt_ref", seed["ref"])
    st.session_state.setdefault("qt_number", seed["number"])
    st.session_state.setdefault("qt_date", date.today())
    st.session_state.setdefault("qt_delivery", seed["delivery"])
    st.session_state.setdefault("qt_pay", seed["pay"])
    st.session_state.setdefault("qt_vat", seed["vat"])
    for index, note in enumerate(seed["notes"]):
        st.session_state.setdefault(f"qt_note_{index}", note)
    st.session_state["_qt_form_ready"] = True


def _quote_from_widgets(edited: pd.DataFrame) -> dict:
    lines = []
    for _, row in edited.iterrows():
        lines.append(
            {
                "품목": _text(row.get("품목")),
                "규격": _text(row.get("규격")),
                "수량": _number(row.get("수량")),
                "단위": _text(row.get("단위")),
                "단가": _number(row.get("단가")),
                "비고": _text(row.get("비고")),
            }
        )
    on = st.session_state.get("qt_date") or date.today()
    if not isinstance(on, date):
        on = date.today()
    return {
        "client": _text(st.session_state.get("qt_client")),
        "person": _text(st.session_state.get("qt_person")),
        "phone": _text(st.session_state.get("qt_phone")),
        "address": _text(st.session_state.get("qt_address")),
        "ref": _text(st.session_state.get("qt_ref")),
        "number": _text(st.session_state.get("qt_number")),
        "quote_date": on,
        "delivery": _text(st.session_state.get("qt_delivery")),
        "pay": _text(st.session_state.get("qt_pay")),
        "vat": st.session_state.get("qt_vat") or "별도",
        "notes": [_text(st.session_state.get(f"qt_note_{i}")) for i in range(4)],
        "lines": lines,
    }


def _bind_quote_copy() -> None:
    """단가 표 복사를 살리고, 검색칸·입력칸 붙여넣기는 그 칸이 받게 한다."""
    st.components.v1.html(
        """
        <script>
        (function () {
          const doc = window.parent.document;
          const win = window.parent;
          if (doc.__qtClipBound) return;
          doc.__qtClipBound = true;
          const proto = Object.getOwnPropertyDescriptor(win.HTMLInputElement.prototype, "value");
          function writeInput(input, next, caret) {
            if (proto && proto.set) proto.set.call(input, next);
            else input.value = next;
            try { input.setSelectionRange(caret, caret); } catch (err) {}
            input.dispatchEvent(new InputEvent("input", { bubbles: true, data: next, inputType: "insertFromPaste" }));
          }
          doc.addEventListener("paste", function (e) {
            const target = e.target;
            if (!target || !target.closest) return;
            const field = target.closest(
              ".st-key-qt_item_q, .st-key-qt_client, .st-key-qt_person, .st-key-qt_phone, .st-key-qt_ref, .st-key-qt_number, .st-key-qt_address, .st-key-qt_delivery, .st-key-qt_pay, .st-key-qt_note_0, .st-key-qt_note_1, .st-key-qt_note_2, .st-key-qt_note_3"
            );
            if (!field) return;
            const input = (target.tagName === "INPUT" || target.tagName === "TEXTAREA")
              ? target
              : field.querySelector("input, textarea");
            if (!input || input.readOnly || input.disabled) return;
            const raw = e.clipboardData ? e.clipboardData.getData("text/plain") : "";
            const text = String(raw || "").replace(/\\r\\n/g, "\\n").split("\\n")[0];
            if (!text) return;
            e.preventDefault();
            e.stopPropagation();
            const start = typeof input.selectionStart === "number" ? input.selectionStart : String(input.value || "").length;
            const end = typeof input.selectionEnd === "number" ? input.selectionEnd : start;
            const value = String(input.value || "");
            writeInput(input, value.slice(0, start) + text + value.slice(end), start + text.length);
          }, true);
          doc.addEventListener("copy", function (e) {
            const target = e.target;
            if (target && target.closest && target.closest(".st-key-qt_item_q, .st-key-qt_client, .st-key-qt_person, .st-key-qt_phone, .st-key-qt_address, .st-key-qt_note_0")) return;
            const box = doc.querySelector(".st-key-qt_price_box");
            if (!box) return;
            const cells = [...box.querySelectorAll('[role="gridcell"][aria-selected="true"]')];
            if (!cells.length) return;
            const lines = new Map();
            cells.forEach(function (cell) {
              const row = cell.closest("tr");
              const rowKey = row ? row.getAttribute("aria-rowindex") || "0" : "0";
              const colKey = cell.getAttribute("aria-colindex") || "0";
              if (!lines.has(rowKey)) lines.set(rowKey, []);
              lines.get(rowKey).push([Number(colKey), (cell.innerText || "").trim()]);
            });
            const text = [...lines.values()].map(function (cols) {
              cols.sort(function (a, b) { return a[0] - b[0]; });
              return cols.map(function (pair) { return pair[1]; }).join("\\t");
            }).join("\\n");
            if (!text) return;
            e.preventDefault();
            e.stopPropagation();
            if (e.clipboardData) e.clipboardData.setData("text/plain", text);
          }, true);
        })();
        </script>
        """,
        height=0,
        width=0,
    )


def _show_prices(sales_df: pd.DataFrame) -> None:
    _bind_quote_copy()
    st.markdown("##### 품목 단가")
    st.caption("2020년 이후 전체 매출입니다. 거래처마다 최근 단가 한 건만 나옵니다.")
    query = st.text_input("품목", key="qt_item_q", placeholder="예: CO2, N2 Bulk", label_visibility="collapsed")
    words = [word for word in query.split() if word.strip()]
    if not words:
        return
    key = (3, tuple(word.casefold() for word in words), id(sales_df), 0 if sales_df is None else len(sales_df))
    bag = st.session_state.get("_qt_search_cache")
    if isinstance(bag, dict) and bag.get("key") == key:
        rows = bag["rows"]
    else:
        rows = search_item_prices(sales_df, query)
        st.session_state["_qt_search_cache"] = {"key": key, "rows": rows}
    if rows.empty:
        st.caption("이 품목의 거래가 없습니다.")
        return
    st.caption(f"{len(rows)}건 · 칸은 ⌘C, 검색칸과 단가 칸은 ⌘V.")
    with st.container(key="qt_price_box"):
        st.dataframe(
            rows,
            hide_index=True,
            use_container_width=True,
            height=640,
            column_config={
                "날짜": st.column_config.TextColumn(width="small"),
                "단가": st.column_config.NumberColumn(format="%.0f", width="small"),
                "품목": st.column_config.TextColumn(width="medium"),
                "거래처": st.column_config.TextColumn(width="medium"),
                "출고량": st.column_config.NumberColumn(format="%.4g", width="small"),
            },
        )


@st.fragment
def render_quote_tab(sales_df: pd.DataFrame) -> None:
    """로컬 견적서. 왼쪽 작성, 오른쪽 품목 단가."""
    _ensure_form()
    left, right = st.columns([1.15, 1], gap="medium")
    with left:
        st.markdown("##### 견적서")
        st.caption("신일가스(주) 공급. 빈 칸을 적어 저장하면 새 시트에 값만 들어갑니다.")
        top_l, top_r = st.columns(2)
        with top_l:
            st.text_input("상호", key="qt_client")
            st.text_input("담당", key="qt_person")
            st.text_input("연락처", key="qt_phone")
        with top_r:
            st.text_input("참조", key="qt_ref")
            st.text_input("견적번호", key="qt_number")
            st.date_input("견적일자", key="qt_date")
        st.text_input("주소", key="qt_address")
        meta_l, meta_r = st.columns(2)
        with meta_l:
            st.text_input("납기", key="qt_delivery")
            st.text_input("결제조건", key="qt_pay")
        with meta_r:
            st.selectbox("부가세", ["별도", "포함"], key="qt_vat")
            on = st.session_state.get("qt_date") or date.today()
            if isinstance(on, date):
                st.caption(f"유효기간 {(on + timedelta(days=7)):%Y-%m-%d} 까지")
        line_frame = pd.DataFrame(_cached_defaults()["lines"])
        line_frame["수량"] = pd.to_numeric(line_frame["수량"], errors="coerce")
        line_frame["단가"] = pd.to_numeric(line_frame["단가"], errors="coerce")
        edited = st.data_editor(
            line_frame,
            key="qt_line_grid",
            num_rows="fixed",
            hide_index=True,
            use_container_width=True,
            column_config={
                "품목": st.column_config.TextColumn(width="small"),
                "규격": st.column_config.TextColumn(width="small"),
                "수량": st.column_config.NumberColumn(min_value=0),
                "단위": st.column_config.TextColumn(width="small"),
                "단가": st.column_config.NumberColumn(min_value=0),
                "비고": st.column_config.TextColumn(width="medium"),
            },
        )
        quote = _quote_from_widgets(edited if isinstance(edited, pd.DataFrame) else pd.DataFrame())
        amounts = [_line_amount(line["수량"], line["단가"]) for line in quote["lines"]]
        supply = sum(amount or 0 for amount in amounts)
        vat_amt = 0.0 if quote["vat"] == "포함" else supply * 0.1
        total = supply if quote["vat"] == "포함" else supply + vat_amt
        st.caption(f"공급가액 {supply:,.0f}  ·  부가세 {vat_amt:,.0f}  ·  합계 {total:,.0f}")
        for index in range(4):
            st.text_input(f"특기사항 {index + 1}", key=f"qt_note_{index}")
        if st.button("견적서 저장", type="primary", key="qt_save"):
            try:
                saved, title = save_quote(_quote_from_widgets(edited))
                st.success(f"저장했습니다. {saved.name}  ·  {title}")
            except Exception as exc:
                st.error(str(exc))
    with right:
        _show_prices(sales_df if isinstance(sales_df, pd.DataFrame) else pd.DataFrame())
