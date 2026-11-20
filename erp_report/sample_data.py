"""Generate realistic, slightly messy ERP exports for one month.

Three branches export in three different layouts (header names, CSV vs XLSX,
date formats) and the data contains the usual problems: duplicates, blanks,
negative quantities, unknown codes, wrong-month dates and amount mismatches.
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

CATEGORIES = {"Pumps": (4000, 25000), "Valves": (800, 6000), "Motors": (6000, 40000), "Spares": (100, 1500)}
REGIONS = ["South", "West", "North", "East"]
SEGMENTS = ["OEM", "Distributor", "EPC Contractor", "Direct"]

BRANCH_LAYOUTS = {
    "Chennai": {"fmt": "csv", "date": "%d/%m/%Y", "cols": {
        "order_no": "Order No", "line_no": "Line", "order_date": "Order Date", "customer_code": "Customer",
        "material_code": "Material", "quantity": "Qty", "unit_price": "Unit Price", "amount": "Net Value",
        "status": "Status"}},
    "Pune": {"fmt": "xlsx", "date": None, "cols": {
        "order_no": "Sales Order", "line_no": "Item", "order_date": "Doc Date", "customer_code": "Sold-To",
        "material_code": "Material No.", "quantity": "Order Qty", "unit_price": "Net Price", "amount": "Amount",
        "status": "Doc Status"}},
    "Delhi": {"fmt": "csv", "date": "%Y-%m-%d", "cols": {
        "order_no": "SO Number", "line_no": "Line", "order_date": "Posting Date", "customer_code": "Cust Code",
        "material_code": "Item Code", "quantity": "Quantity", "unit_price": "Rate", "amount": "Line Amount",
        "status": "Status"}},
}


def generate(out_dir: str | Path, month: str = "2026-10", seed: int = 11) -> dict[str, int]:
    rng = np.random.default_rng(seed)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    # ---- masters
    mats = []
    for cat, (lo, hi) in CATEGORIES.items():
        for i in range(1, 16):
            price = round(float(rng.uniform(lo, hi)), -1)
            mats.append((f"{cat[:3].upper()}-{i:04d}", f"{cat[:-1]} model {i}", cat, "EA", round(price * rng.uniform(0.55, 0.8), 2), price))
    materials = pd.DataFrame(mats, columns=["material_code", "description", "category", "uom", "standard_cost", "list_price"])
    customers = pd.DataFrame({
        "customer_code": [f"C{i:04d}" for i in range(1, 61)],
        "customer_name": [f"Customer {i:02d} Pvt Ltd" for i in range(1, 61)],
        "region": rng.choice(REGIONS, 60),
        "segment": rng.choice(SEGMENTS, 60),
    })
    materials.drop(columns="list_price").to_excel(out / "material_master.xlsx", index=False)
    customers.to_csv(out / "customer_master.csv", index=False)

    start = pd.Timestamp(f"{month}-01")
    days = pd.date_range(start, start + pd.offsets.MonthEnd(0))
    price = materials.set_index("material_code").list_price
    counts = {}

    for b_idx, (branch, layout) in enumerate(BRANCH_LAYOUTS.items()):
        rows = []
        for o in range(1, int(rng.integers(140, 200))):
            order_no = f"{branch[:3].upper()}{b_idx + 1}{o:05d}"
            date = rng.choice(days)
            cust = rng.choice(customers.customer_code)
            status = rng.choice(["INVOICED", "DELIVERED", "OPEN", "CANCELLED"], p=[0.6, 0.2, 0.15, 0.05])
            for line in range(1, int(rng.integers(1, 5)) + 1):
                mat = rng.choice(materials.material_code)
                qty = int(rng.integers(1, 20))
                unit = round(float(price[mat] * rng.uniform(0.85, 1.0)), 2)  # discounts
                rows.append([order_no, line * 10, pd.Timestamp(date), cust, mat, qty, unit, round(qty * unit, 2), status])
        df = pd.DataFrame(rows, columns=list(layout["cols"].keys()))

        # ---- inject realistic data problems
        n = len(df)
        pick = lambda k: rng.choice(n, k, replace=False)
        df.loc[pick(2), "customer_code"] = None                          # blanks
        df.loc[pick(2), "quantity"] = -df.loc[pick(2), "quantity"].abs().values  # returns keyed as negatives
        df.loc[pick(2), "material_code"] = "OLD-9999"                      # material not in master
        df.loc[pick(2), "customer_code"] = "C9001"                         # new customer, not yet in master
        df.loc[pick(2), "order_date"] = start - pd.Timedelta(days=3)       # previous month
        i = pick(2)
        df.loc[i, "amount"] = (df.loc[i, "amount"] * 1.1).round(2)        # manual edit -> mismatch
        df.loc[pick(3), "status"] = df.loc[pick(3), "status"].str.lower().values  # inconsistent case (cleaned, not an error)
        df = pd.concat([df, df.sample(2, random_state=b_idx)], ignore_index=True)  # duplicate lines

        df = df.rename(columns=layout["cols"])
        date_col = layout["cols"]["order_date"]
        if layout["date"]:
            df[date_col] = pd.to_datetime(df[date_col]).dt.strftime(layout["date"])
        path = out / f"sales_{branch.lower()}_{month}.{layout['fmt']}"
        if layout["fmt"] == "xlsx":
            df.to_excel(path, index=False)
        else:
            df.to_csv(path, index=False)
        counts[path.name] = len(df)
    return counts


if __name__ == "__main__":
    for k, v in generate("sample_input").items():
        print(f"{k}: {v} rows")
