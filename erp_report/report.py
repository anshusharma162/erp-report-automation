"""Aggregations and the formatted Excel workbook."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

NAVY = "1F3864"
HEADER = PatternFill("solid", fgColor=NAVY)
ERROR_FILL = PatternFill("solid", fgColor="F8D7DA")
WARN_FILL = PatternFill("solid", fgColor="FFF3CD")
THIN = Side(style="thin", color="D0D0D0")
MONEY_COLS = {"revenue", "gross_margin", "cost", "amount", "unit_price", "avg_order_value", "standard_cost"}
PCT_COLS = {"margin_pct", "share_pct", "pass_rate_pct"}


def enrich(clean: pd.DataFrame, materials: pd.DataFrame, customers: pd.DataFrame) -> pd.DataFrame:
    df = clean.merge(materials[["material_code", "description", "category", "standard_cost"]], on="material_code", how="left")
    df = df.merge(customers[["customer_code", "customer_name", "region", "segment"]], on="customer_code", how="left")
    df["customer_name"] = df.customer_name.fillna("(not in master) " + df.customer_code.astype(str))
    df["region"] = df.region.fillna("Unknown")
    df["segment"] = df.segment.fillna("Unknown")
    df["revenue"] = df.amount
    df["cost"] = (df.quantity * df.standard_cost).round(2)
    df["gross_margin"] = df.revenue - df.cost
    return df


def _group(df: pd.DataFrame, by) -> pd.DataFrame:
    g = df.groupby(by, dropna=False).agg(
        orders=("order_no", "nunique"), lines=("order_no", "size"), quantity=("quantity", "sum"),
        revenue=("revenue", "sum"), gross_margin=("gross_margin", "sum"),
    ).reset_index()
    g["margin_pct"] = (g.gross_margin / g.revenue).round(4)
    g["share_pct"] = (g.revenue / g.revenue.sum()).round(4)
    return g.sort_values("revenue", ascending=False)


def build_tables(sales: pd.DataFrame, cfg: dict) -> dict[str, pd.DataFrame]:
    s = sales[~sales.status.isin(cfg["report"]["exclude_status"])]
    by_cust = _group(s, ["customer_code", "customer_name", "segment", "region"]).head(cfg["report"]["top_customers"])
    daily = s.groupby(s.order_date.dt.date).agg(orders=("order_no", "nunique"), revenue=("revenue", "sum")).reset_index()
    daily.columns = ["date", "orders", "revenue"]
    return {
        "by_branch": _group(s, "branch"),
        "by_category": _group(s, "category"),
        "by_region": _group(s, "region"),
        "top_customers": by_cust,
        "daily": daily,
        "status": s.groupby("status").agg(lines=("order_no", "size"), revenue=("revenue", "sum")).reset_index(),
        "reported": s,
    }


def data_quality(raw: pd.DataFrame, exceptions: pd.DataFrame, rules: list[dict]) -> pd.DataFrame:
    counts = exceptions.groupby("rule").row_id.nunique() if len(exceptions) else pd.Series(dtype=int)
    rows = [{"rule": r["name"], "severity": r["severity"], "failed_rows": int(counts.get(r["name"], 0))} for r in rules]
    dq = pd.DataFrame(rows)
    dq["pass_rate_pct"] = (1 - dq.failed_rows / len(raw)).round(4)
    return dq


# ------------------------------------------------------------------ Excel
def _write_table(ws, df: pd.DataFrame, start_row: int = 1, start_col: int = 1) -> None:
    for r_i, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
        for c_i, val in enumerate(row):
            if val is pd.NA or val is pd.NaT or (isinstance(val, float) and val != val):
                val = None
            cell = ws.cell(row=start_row + r_i, column=start_col + c_i, value=val)
            cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
            col = df.columns[c_i]
            if r_i == 0:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = HEADER
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif col in MONEY_COLS:
                cell.number_format = "#,##0"
            elif col in PCT_COLS:
                cell.number_format = "0.0%"
            elif col in {"date", "order_date"}:
                cell.number_format = "dd-mmm-yyyy"
    for c_i, col in enumerate(df.columns):
        width = max([len(str(col))] + [len(str(v)) for v in df[col].head(300)]) + 2
        ws.column_dimensions[get_column_letter(start_col + c_i)].width = max(ws.column_dimensions[get_column_letter(start_col + c_i)].width or 0, min(width, 45))


def _sheet(wb, name: str, df: pd.DataFrame, title: str | None = None):
    ws = wb.create_sheet(name)
    start = 1
    if title:
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=13, color=NAVY)
        start = 3
    _write_table(ws, df, start)
    ws.freeze_panes = ws.cell(row=start + 1, column=1)
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(df.columns))}{start + len(df)}"
    return ws


def write_workbook(path: Path, tables: dict, dq: pd.DataFrame, exceptions: pd.DataFrame,
                   rejected: pd.DataFrame, meta: dict, cfg: dict) -> None:
    cur = cfg["report"]["currency"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = f"Monthly Sales Report – {meta['period_label']}"
    ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = f"Generated {meta['generated']} from {meta['files']} ERP files · {meta['runtime_s']:.1f}s"
    ws["A2"].font = Font(italic=True, color="666666")

    rep = tables["reported"]
    kpis = [
        ("Revenue", rep.revenue.sum(), f'"{cur}"#,##0'),
        ("Gross margin", rep.gross_margin.sum(), f'"{cur}"#,##0'),
        ("Margin %", rep.gross_margin.sum() / rep.revenue.sum() if len(rep) else 0, "0.0%"),
        ("Orders", rep.order_no.nunique(), "#,##0"),
        ("Avg order value", rep.revenue.sum() / max(1, rep.order_no.nunique()), f'"{cur}"#,##0'),
        ("Rows read", meta["rows_read"], "#,##0"),
        ("Rows rejected (errors)", meta["rows_rejected"], "#,##0"),
        ("Rows with warnings", meta["rows_warning"], "#,##0"),
    ]
    for i, (label, val, fmt) in enumerate(kpis):
        r = 4 + i
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=float(val))
        c.number_format = fmt
        if "rejected" in label and val:
            c.fill = ERROR_FILL
        if "warnings" in label and val:
            c.fill = WARN_FILL
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 18

    ws["D4"] = "Revenue by branch"
    ws["D4"].font = Font(bold=True, color=NAVY)
    br = tables["by_branch"][["branch", "orders", "revenue", "margin_pct", "share_pct"]]
    _write_table(ws, br, start_row=5, start_col=4)
    chart = BarChart()
    chart.title = "Revenue by branch"
    chart.y_axis.title = cur
    chart.add_data(Reference(ws, min_col=6, min_row=5, max_row=5 + len(br)), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=4, min_row=6, max_row=5 + len(br)))
    chart.height, chart.width = 7, 14
    ws.add_chart(chart, "J4")

    ws["A14"] = "Data quality"
    ws["A14"].font = Font(bold=True, color=NAVY)
    _write_table(ws, dq, start_row=15, start_col=1)

    _sheet(wb, "By Category", tables["by_category"], "Revenue and margin by product category")
    _sheet(wb, "By Region", tables["by_region"], "Revenue by customer region")
    _sheet(wb, "Top Customers", tables["top_customers"], f"Top {cfg['report']['top_customers']} customers by revenue")
    d = _sheet(wb, "Daily Trend", tables["daily"], "Daily order intake")
    line = LineChart()
    line.title = "Daily revenue"
    line.add_data(Reference(d, min_col=3, min_row=3, max_row=3 + len(tables["daily"])), titles_from_data=True)
    line.set_categories(Reference(d, min_col=1, min_row=4, max_row=3 + len(tables["daily"])))
    line.height, line.width = 8, 18
    d.add_chart(line, "F3")

    ex = exceptions.drop(columns="row_id").sort_values(["severity", "source_file", "source_row"])
    exs = _sheet(wb, "Exceptions", ex, "Rows that failed validation – fix these in the ERP")
    sev_col = list(ex.columns).index("severity") + 1
    for row in exs.iter_rows(min_row=4, min_col=sev_col, max_col=sev_col):
        for cell in row:
            cell.fill = ERROR_FILL if cell.value == "error" else WARN_FILL

    keep = ["source_file", "source_row", "branch", "order_no", "line_no", "order_date", "customer_code", "customer_name",
            "material_code", "description", "category", "quantity", "unit_price", "amount", "status", "gross_margin", "has_warning"]
    _sheet(wb, "Clean Data", rep[[c for c in keep if c in rep.columns]])
    _sheet(wb, "Rejected Rows", rejected)
    wb.save(path)
