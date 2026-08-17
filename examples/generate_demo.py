from pathlib import Path

from openpyxl import Workbook


def main() -> None:
    root = Path(__file__).resolve().parents[1] / "demo-data"
    root.mkdir(exist_ok=True)

    catalog = Workbook()
    sheet = catalog.active
    sheet.title = "Catalog"
    sheet.append(["sku", "product_name", "unit_cost", "pack_size", "opening_stock", "current_stock"])
    sheet.append(["SKU-001", "Reusable bottle", 8.5, 6, 120, 120])
    sheet.append(["SKU-002", "Canvas tote", 4.2, 10, 80, 80])
    catalog.save(root / "catalog.xlsx")
    catalog.close()

    transactions = Workbook()
    sheet = transactions.active
    sheet.title = "Transactions"
    sheet.append(["transaction_id", "date", "sku", "direction", "package_count", "unit_cost", "line_total"])
    sheet.append(["TX-001", "2026-08-01", "SKU-001", "inbound", 4, 8.5, 204])
    sheet.append(["TX-002", "2026-08-02", "SKU-001", "outbound", 3, 8.5, 153])
    sheet.append(["TX-003", "2026-08-02", "SKU-002", "outbound", 2, 4.2, 84])
    transactions.save(root / "transactions.xlsx")
    transactions.close()
    print(f"Synthetic demo workbooks written to {root}")


if __name__ == "__main__":
    main()
