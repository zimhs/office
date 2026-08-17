"""시장조사 종속 필터용 사전 인덱스."""
from __future__ import annotations

import re

import pandas as pd


def extract_suppliers(series, *, limit: int = 300) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for v in series.astype(str):
        for p in re.split(r"\s*[·|/]\s*", v):
            p = p.strip()
            if not p or p in seen:
                continue
            seen.add(p)
            out.append(p)
            if len(out) >= limit:
                return sorted(out)
    return sorted(out)


def _pack(sub: pd.DataFrame) -> dict:
    if sub is None or sub.empty:
        return {"cxr": {}, "cxa": [], "sr": {}, "src": {}, "sa": []}
    cxr: dict[str, list[str]] = {}
    for r, g in sub.groupby("지역", sort=False):
        vals = {str(c) for c in g["산업단지"].dropna().unique() if str(c) and str(c) != "미분류"}
        cxr[str(r)] = sorted(vals)
    cxa = sorted({str(c) for c in sub["산업단지"].dropna().unique() if str(c) and str(c) != "미분류"})
    sr: dict[str, list[str]] = {}
    for r, g in sub.groupby("지역", sort=False):
        sr[str(r)] = extract_suppliers(g["공급사"], limit=300)
    src: dict[str, list[str]] = {}
    classified = sub[sub["산업단지"].astype(str) != "미분류"]
    if not classified.empty:
        for (r, c), g in classified.groupby(["지역", "산업단지"], sort=False):
            src[str(r) + "|" + str(c)] = extract_suppliers(g["공급사"], limit=200)
    return {"cxr": cxr, "cxa": cxa, "sr": sr, "src": src, "sa": extract_suppliers(sub["공급사"], limit=300)}


def build_cascade_index(df: pd.DataFrame) -> dict:
    survey = df.loc[~df["_factory_only"]] if "_factory_only" in df.columns else df
    return {"survey": _pack(survey), "all": _pack(df)}


def complex_opts(index: dict, regions: list[str], *, include_factory: bool) -> list[str]:
    pack = index["all" if include_factory else "survey"]
    if not regions:
        return list(pack["cxa"])
    out: set[str] = set()
    for r in regions:
        out.update(pack["cxr"].get(r, []))
    return sorted(out)


def supplier_opts(index: dict, regions: list[str], complexes: list[str], *, include_factory: bool) -> list[str]:
    pack = index["all" if include_factory else "survey"]
    if complexes:
        out: set[str] = set()
        regs = regions or list(pack["cxr"].keys())
        for r in regs:
            for c in complexes:
                out.update(pack["src"].get(str(r) + "|" + str(c), []))
        return sorted(out)
    if regions:
        out2: set[str] = set()
        for r in regions:
            out2.update(pack["sr"].get(r, []))
        return sorted(out2)
    return list(pack["sa"])

