"""One command from raw ERP exports to a finished Excel report.

    python -m erp_report run --input sample_input --month 2026-10
    python -m erp_report sample --out sample_input          # create demo exports
"""
from __future__ import annotations

import argparse
import logging
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml

from .loader import load_master, load_sales
from .report import build_tables, data_quality, enrich, write_workbook
from .validate import run_rules, split_clean

log = logging.getLogger("erp_report")


def run(input_dir: str | Path, month: str, out_dir: str | Path = "output", config: str | Path = "config.yaml",
        fail_on_error_rate: float | None = None) -> dict:
    t0 = time.perf_counter()
    input_dir, out_dir = Path(input_dir), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    cfg = yaml.safe_load(Path(config).read_text(encoding="utf-8"))

    start = pd.Timestamp(f"{month}-01")
    end = start + pd.offsets.MonthEnd(0)

    # 1. read + standardise
    sales = load_sales(input_dir, cfg)
    materials = load_master(input_dir, cfg["inputs"]["materials"])
    customers = load_master(input_dir, cfg["inputs"]["customers"])
    files = sales.source_file.nunique() + 2

    # 2. validate
    ctx = {"materials": materials, "customers": customers, "period": (start, end)}
    exceptions = run_rules(sales, cfg["rules"], ctx)
    clean, rejected = split_clean(sales, exceptions)
    dq = data_quality(sales, exceptions, cfg["rules"])
    for r in dq.itertuples():
        if r.failed_rows:
            log.warning("%-24s %-8s %4d rows", r.rule, r.severity, r.failed_rows)

    error_rate = len(rejected) / len(sales)
    if fail_on_error_rate is not None and error_rate > fail_on_error_rate:
        raise SystemExit(f"Error rate {error_rate:.1%} above limit {fail_on_error_rate:.1%} – report not produced.")

    # 3. merge + aggregate
    enriched = enrich(clean, materials, customers)
    tables = build_tables(enriched, cfg)

    # 4. write
    meta = {
        "period_label": start.strftime("%B %Y"),
        "generated": datetime.now().strftime("%d-%b-%Y %H:%M"),
        "files": files,
        "rows_read": len(sales),
        "rows_rejected": len(rejected),
        "rows_warning": int(clean.has_warning.sum()),
        "runtime_s": time.perf_counter() - t0,
    }
    out_path = out_dir / f"sales_report_{month}.xlsx"
    write_workbook(out_path, tables, dq, exceptions, rejected, meta, cfg)
    exceptions.to_csv(out_dir / f"exceptions_{month}.csv", index=False)

    rep = tables["reported"]
    summary = {
        "report": str(out_path),
        "rows_read": len(sales),
        "rows_rejected": len(rejected),
        "rows_with_warnings": meta["rows_warning"],
        "revenue": round(float(rep.revenue.sum()), 2),
        "orders": int(rep.order_no.nunique()),
        "seconds": round(time.perf_counter() - t0, 2),
    }
    log.info("report written: %s", out_path)
    return summary


def main(argv=None) -> None:
    p = argparse.ArgumentParser(prog="erp_report", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run", help="build the monthly report")
    r.add_argument("--input", required=True, help="folder with the ERP exports")
    r.add_argument("--month", required=True, help="reporting month, YYYY-MM")
    r.add_argument("--out", default="output")
    r.add_argument("--config", default="config.yaml")
    r.add_argument("--max-error-rate", type=float, default=None, help="abort if more than this share of rows fail (e.g. 0.1)")
    s = sub.add_parser("sample", help="generate demo ERP exports")
    s.add_argument("--out", default="sample_input")
    s.add_argument("--month", default="2026-10")
    a = p.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S",
                        handlers=[logging.StreamHandler(sys.stdout)])
    if a.cmd == "sample":
        from .sample_data import generate

        for k, v in generate(a.out, a.month).items():
            print(f"wrote {k} ({v} rows)")
        return
    if a.out:
        Path(a.out).mkdir(parents=True, exist_ok=True)
        logging.getLogger().addHandler(logging.FileHandler(Path(a.out) / f"run_{a.month}.log", mode="w"))
    summary = run(a.input, a.month, a.out, a.config, a.max_error_rate)
    print("\n".join(f"  {k:20s} {v}" for k, v in summary.items()))


if __name__ == "__main__":
    main()
