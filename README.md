# Inventory Reconciliation Pipeline

A small, production-minded Python pipeline that turns monthly inventory workbooks into a validated reconciliation report. It separates a non-mutating `check` pass from an explicit `apply` pass, then creates a backup before updating the catalog.

> Portfolio edition: every example is synthetic. No production workbook, organization name, location code, or historical repository data is included.

## What it demonstrates

- deterministic Excel ingestion and SKU-level aggregation
- fail-closed validation for duplicate IDs, unknown SKUs, price mismatches, invalid totals, formulas, and negative stock
- dry-run reporting before mutation
- timestamped backups and atomic workbook replacement
- reproducible synthetic fixtures and integration tests

## Flow

```text
Catalog.xlsx ─┐
              ├─ validate ─ aggregate ─ reconciliation report
Transactions ─┘                         └─ optional backup + apply
```

## Run the synthetic demo

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python examples\generate_demo.py
python -m inventory_reconciliation --catalog demo-data\catalog.xlsx --transactions demo-data\transactions.xlsx --output-dir output --mode check
```

After reviewing `output\reconciliation-check.xlsx`, rerun with `--mode apply`. The catalog is backed up before any write.

## Data contract

`Catalog` columns: `sku`, `product_name`, `unit_cost`, `pack_size`, `opening_stock`, `current_stock`.

`Transactions` columns: `transaction_id`, `date`, `sku`, `direction`, `package_count`, `unit_cost`, `line_total`.

`direction` is restricted to `inbound` or `outbound`; `line_total` must equal `package_count × pack_size × unit_cost`.

## Privacy boundary

Generated workbooks, reports, and backups are ignored by Git. The public repository intentionally contains only code and a fixture generator.
