"""Config-driven validation rules.

Each rule returns a boolean mask of failing rows. The engine collects every
failure into one exceptions table (row can fail several rules) with the source
file and row number, so whoever owns the ERP data can fix it at the source.
"""
from __future__ import annotations

import numpy as np
import pandas as pd


def _not_null(df, r, ctx):
    cols = r["columns"]
    mask = df[cols].isna().any(axis=1)
    detail = df[cols].isna().apply(lambda row: "blank: " + ", ".join(c for c, v in row.items() if v), axis=1)
    return mask, detail


def _unique(df, r, ctx):
    mask = df.duplicated(subset=r["columns"], keep="first")
    return mask, "duplicate of an earlier line " + df[r["columns"]].astype(str).agg("/".join, axis=1)


def _range(df, r, ctx):
    col = df[r["column"]]
    mask = pd.Series(False, index=df.index)
    if "min" in r:
        mask |= col < r["min"]
    if "max" in r:
        mask |= col > r["max"]
    return mask & col.notna(), f"{r['column']} = " + col.astype(str)


def _allowed(df, r, ctx):
    col = df[r["column"]]
    mask = col.notna() & ~col.isin(r["values"])
    return mask, f"{r['column']} = " + col.astype(str)


def _lookup(df, r, ctx):
    valid = set(ctx[r["table"]][r["key"]].dropna())
    col = df[r["column"]]
    mask = col.notna() & ~col.isin(valid)
    return mask, f"{r['column']} {col.astype(str)} not in {r['table']} master"


def _period(df, r, ctx):
    start, end = ctx["period"]
    col = df[r["column"]]
    mask = col.isna() | (col < start) | (col > end)
    return mask, "date " + col.dt.strftime("%Y-%m-%d").fillna("unreadable") + f" outside {start:%b %Y}"


def _amount(df, r, ctx):
    expected = (df.quantity * df.unit_price).round(2)
    diff = (df.amount - expected).abs()
    mask = diff > np.maximum(r.get("tolerance", 0.01), 0.001 * expected.abs())
    return mask.fillna(False), "amount " + df.amount.astype(str) + " ≠ qty×price " + expected.astype(str)


def _outlier(df, r, ctx):
    col = df[r["column"]]
    z = (col - col.mean()) / col.std(ddof=0)
    return (z.abs() > r["z"]).fillna(False), f"{r['column']} " + col.round(0).astype(str) + f" > {r['z']}σ from mean"


RULES = {
    "not_null": _not_null, "unique": _unique, "range": _range, "allowed": _allowed,
    "lookup": _lookup, "period": _period, "amount": _amount, "outlier": _outlier,
}


def run_rules(df: pd.DataFrame, rules: list[dict], ctx: dict) -> pd.DataFrame:
    """Return one row per (data row, failed rule)."""
    out = []
    for r in rules:
        fn = RULES[r["type"]]
        mask, detail = fn(df, r, ctx)
        if isinstance(detail, str):
            detail = pd.Series(detail, index=df.index)
        hit = df[mask]
        if hit.empty:
            continue
        out.append(pd.DataFrame({
            "row_id": hit.index,
            "rule": r["name"],
            "severity": r["severity"],
            "detail": detail[mask].values,
            "source_file": hit.source_file.values,
            "source_row": hit.source_row.values,
            "order_no": hit.order_no.values,
            "line_no": hit.line_no.values,
        }))
    cols = ["row_id", "rule", "severity", "detail", "source_file", "source_row", "order_no", "line_no"]
    return pd.concat(out, ignore_index=True) if out else pd.DataFrame(columns=cols)


def split_clean(df: pd.DataFrame, exceptions: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Rows with any error are rejected; rows with only warnings are kept and flagged."""
    err_ids = set(exceptions.loc[exceptions.severity == "error", "row_id"])
    warn_ids = set(exceptions.loc[exceptions.severity == "warning", "row_id"])
    clean = df[~df.index.isin(err_ids)].copy()
    clean["has_warning"] = clean.index.isin(warn_ids)
    rejected = df[df.index.isin(err_ids)].copy()
    return clean, rejected
