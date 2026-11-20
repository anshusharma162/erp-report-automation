from pathlib import Path

import pandas as pd
import pytest
import yaml
from openpyxl import load_workbook

from erp_report.cli import run
from erp_report.loader import parse_dates, standardise_columns
from erp_report.sample_data import generate
from erp_report.validate import run_rules, split_clean

ROOT = Path(__file__).resolve().parents[1]
CFG = yaml.safe_load((ROOT / "config.yaml").read_text(encoding="utf-8"))


def test_header_variants_map_to_standard_names():
    df = pd.DataFrame(columns=["SO Number", "Material No.", " qty ", "Net Price"])
    assert list(standardise_columns(df, CFG["column_aliases"]).columns) == ["order_no", "material_code", "quantity", "unit_price"]


def test_dates_iso_and_day_first():
    out = parse_dates(pd.Series(["2026-10-05", "05/10/2026", "31/10/2026"]))
    assert list(out.dt.strftime("%Y-%m-%d")) == ["2026-10-05", "2026-10-05", "2026-10-31"]


def _frame(**overrides):
    base = {
        "order_no": ["A1", "A1", "A2"], "line_no": [10, 20, 10], "order_date": pd.to_datetime(["2026-10-01"] * 3),
        "customer_code": ["C1", "C1", "C2"], "material_code": ["M1", "M1", "M1"], "quantity": [2.0, 1.0, 3.0],
        "unit_price": [100.0, 100.0, 100.0], "amount": [200.0, 100.0, 300.0], "status": ["INVOICED"] * 3,
        "source_file": ["f.csv"] * 3, "source_row": [2, 3, 4],
    }
    base.update(overrides)
    return pd.DataFrame(base)


CTX = {
    "materials": pd.DataFrame({"material_code": ["M1"]}),
    "customers": pd.DataFrame({"customer_code": ["C1", "C2"]}),
    "period": (pd.Timestamp("2026-10-01"), pd.Timestamp("2026-10-31")),
}


def test_clean_data_passes_all_rules():
    assert run_rules(_frame(), CFG["rules"], CTX).empty


@pytest.mark.parametrize("overrides,rule", [
    ({"quantity": [2.0, -1.0, 3.0], "amount": [200.0, -100.0, 300.0]}, "non_positive_quantity"),
    ({"line_no": [10, 10, 10], "order_no": ["A1", "A1", "A2"]}, "duplicate_line"),
    ({"material_code": ["M1", "X9", "M1"]}, "unknown_material"),
    ({"customer_code": ["C1", None, "C2"]}, "missing_values"),
    ({"status": ["INVOICED", "BOOKED", "INVOICED"]}, "invalid_status"),
    ({"order_date": pd.to_datetime(["2026-10-01", "2026-09-30", "2026-10-01"])}, "outside_period"),
    ({"amount": [200.0, 150.0, 300.0]}, "amount_mismatch"),
])
def test_each_rule_catches_its_problem(overrides, rule):
    ex = run_rules(_frame(**overrides), CFG["rules"], CTX)
    assert rule in set(ex.rule)
    assert set(ex[ex.rule == rule].source_row) == {3}


def test_warnings_kept_errors_rejected():
    df = _frame(customer_code=["C1", "C404", "C2"], material_code=["M1", "M1", "X9"])
    ex = run_rules(df, CFG["rules"], CTX)
    clean, rejected = split_clean(df, ex)
    assert list(rejected.source_row) == [4]                 # unknown material = error
    assert clean.set_index("source_row").has_warning[3]     # unknown customer = warning


def test_end_to_end(tmp_path):
    generate(tmp_path / "in", "2026-10")
    s = run(tmp_path / "in", "2026-10", tmp_path / "out", ROOT / "config.yaml")
    assert s["rows_rejected"] > 0 and s["revenue"] > 0
    wb = load_workbook(s["report"])
    assert {"Summary", "By Category", "Top Customers", "Exceptions", "Clean Data"} <= set(wb.sheetnames)
    ex = pd.read_csv(tmp_path / "out" / "exceptions_2026-10.csv")
    # every injected problem type is detected
    assert {"duplicate_line", "unknown_material", "unknown_customer", "outside_period",
            "non_positive_quantity", "missing_values", "amount_mismatch"} <= set(ex.rule)
