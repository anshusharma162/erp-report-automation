"""Find, read and standardise ERP export files."""
from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

log = logging.getLogger("erp_report")


def read_any(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(path)
    return pd.read_csv(path)


def find(folder: Path, pattern: str) -> list[Path]:
    files = sorted(p for p in folder.glob(pattern) if p.suffix.lower() in {".csv", ".xlsx", ".xls"} and not p.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"No files matching '{pattern}' in {folder}")
    return files


def standardise_columns(df: pd.DataFrame, aliases: dict[str, list[str]]) -> pd.DataFrame:
    lookup = {a.strip().lower(): std for std, names in aliases.items() for a in names}
    renamed = {c: lookup[c.strip().lower()] for c in df.columns if c.strip().lower() in lookup}
    return df.rename(columns=renamed)


def clean(df: pd.DataFrame) -> pd.DataFrame:
    """Fix formatting issues that are not data errors (case, spaces, types)."""
    df = df.copy()
    for c in df.select_dtypes(include=["object", "string"]).columns:
        df[c] = df[c].astype("string").str.strip().replace("", pd.NA)
    if "status" in df:
        df["status"] = df["status"].str.upper()
    for c in ("quantity", "unit_price", "amount"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    if "order_date" in df and not pd.api.types.is_datetime64_any_dtype(df["order_date"]):
        df["order_date"] = parse_dates(df["order_date"])
    return df


def parse_dates(s: pd.Series) -> pd.Series:
    """ISO dates (2026-10-05) are year-first; everything else is Indian day-first (05/10/2026)."""
    s = s.astype("string")
    iso = s.str.match(r"^\d{4}-\d{2}-\d{2}", na=False)
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    out[iso] = pd.to_datetime(s[iso].str[:10], format="%Y-%m-%d", errors="coerce")
    out[~iso] = pd.to_datetime(s[~iso], dayfirst=True, format="mixed", errors="coerce")
    return out


def load_sales(folder: Path, cfg: dict) -> pd.DataFrame:
    frames = []
    for path in find(folder, cfg["inputs"]["sales"]):
        raw = read_any(path)
        df = standardise_columns(raw, cfg["column_aliases"])
        missing = [c for c in cfg["required_columns"] if c not in df.columns]
        if missing:
            raise ValueError(f"{path.name}: missing required columns {missing}. Add the header to column_aliases.")
        if "branch" not in df.columns:
            # e.g. sales_chennai_2026-10.csv -> Chennai
            df["branch"] = path.stem.split("_")[1].title() if "_" in path.stem else path.stem
        df = clean(df)
        df["source_file"] = path.name
        df["source_row"] = range(2, len(df) + 2)  # Excel row number incl. header
        log.info("read %-32s %5d rows", path.name, len(df))
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def load_master(folder: Path, pattern: str) -> pd.DataFrame:
    path = find(folder, pattern)[0]
    df = clean(read_any(path))
    log.info("read %-32s %5d rows", path.name, len(df))
    return df
