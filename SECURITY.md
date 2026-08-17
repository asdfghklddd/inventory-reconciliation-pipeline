# Security

This showcase contains no production workbooks or customer data. Please do not report vulnerabilities with real business files attached.

The pipeline rejects formulas in numeric input cells, validates identifiers and totals, writes through a temporary file, and creates a timestamped backup before `apply` mode mutates the catalog. Use `check` mode first and keep source workbooks under your own access controls.
