# ERP & Excel Report Automation Suite

Turns a multi-hour manual month-end task (download ERP exports from each branch, fix
headers, paste into one sheet, VLOOKUP the masters, hunt for bad rows, build pivots and
charts) into **one command**:

```bash
python -m erp_report run --input sample_input --month 2026-10
```

```
read sales_chennai_2026-10.csv    497 rows
read sales_delhi_2026-10.csv      374 rows
read sales_pune_2026-10.xlsx      468 rows
read material_master.xlsx          60 rows
read customer_master.csv           60 rows
missing_values          error       6 rows
duplicate_line          error       6 rows
non_positive_quantity   error       6 rows
unknown_material        error       6 rows
unknown_customer        warning     6 rows
outside_period          error       6 rows
amount_mismatch         warning    12 rows
unusually_large_line    warning     9 rows
report written: output/sales_report_2026-10.xlsx
  rows_read            1339
  rows_rejected        30
  rows_with_warnings   20
  seconds              1.75
```

## What it does

1. **Collects** every branch export that matches `sales_*.*` (CSV or Excel) plus the
   material and customer masters.
2. **Standardises** different ERP layouts. One branch says `Sales Order` / `Material No.`,
   another `SO Number` / `Item Code`. A mapping in `config.yaml` turns them into one schema.
   It also fixes formatting issues that aren't data errors: stray spaces, `invoiced` vs
   `INVOICED`, ISO vs day-first dates.
3. **Validates** each row against config-driven rules and records every failure with the
   **source file and Excel row number**, so it can be fixed in the ERP.
   - **error** → row excluded from the figures
   - **warning** → row kept, flagged for review
4. **Merges** with the masters (category, customer, region, standard cost) and calculates
   revenue, cost and gross margin.
5. **Writes a formatted Excel workbook** and an exceptions CSV, plus a run log.

## Validation rules (`config.yaml`)

| Rule | Type | Severity | Catches |
|---|---|---|---|
| missing_values | not_null | error | blank order / customer / material / qty |
| duplicate_line | unique | error | same order + line exported twice |
| non_positive_quantity | range | error | returns keyed as negative sales |
| negative_price | range | error | |
| invalid_status | allowed | error | status outside the allowed list |
| unknown_material | lookup | error | material not in master (obsolete codes) |
| unknown_customer | lookup | warning | new customer not yet in master |
| outside_period | period | error | documents dated in another month |
| amount_mismatch | amount | warning | net value ≠ qty × price (manual edits) |
| unusually_large_line | outlier | warning | line value more than 4σ from the mean |

Adding a rule is a config change, not a code change:

```yaml
- {name: price_too_high, type: range, column: unit_price, max: 100000, severity: warning}
```

## Output workbook

| Sheet | Contents |
|---|---|
| Summary | revenue, margin, orders, AOV, rows rejected/warned, revenue-by-branch table + chart, data-quality pass rates |
| By Category / By Region | orders, qty, revenue, gross margin, margin %, share % |
| Top Customers | top 15 by revenue |
| Daily Trend | daily orders and revenue + line chart |
| Exceptions | every failed rule with file, row, order and detail (colour-coded by severity) |
| Clean Data | the reconciled dataset behind every number |
| Rejected Rows | rows left out, for follow-up |

Cancelled orders are excluded from revenue (configurable).

## Run it

```bash
pip install -r requirements.txt
python -m erp_report sample --out sample_input --month 2026-10     # demo exports with realistic errors
python -m erp_report run --input sample_input --month 2026-10      # build the report
python -m erp_report run --input sample_input --month 2026-10 --max-error-rate 0.05   # stop if data too dirty
pytest -q
```

Can be scheduled with Windows Task Scheduler or cron, so the report is waiting in a shared
folder every month-end.

## Project structure

```
config.yaml              file patterns, header aliases, rules, report options
erp_report/
  loader.py              discover files, read CSV/XLSX, standardise headers, clean types and dates
  validate.py            rule engine (not_null, unique, range, allowed, lookup, period, amount, outlier)
  report.py              enrichment, aggregations, formatted Excel with charts
  cli.py                 `run` and `sample` commands
  sample_data.py         demo ERP exports from 3 branches in 3 different layouts
tests/                   one test per rule + end-to-end test
```

## Tech
Python · pandas · openpyxl (formatting, native Excel charts) · PyYAML · pytest
