from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Sequence

from .pipeline import ReconciliationError, reconcile


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Reconcile inventory workbooks safely")
    parser.add_argument("--catalog", type=Path, required=True)
    parser.add_argument("--transactions", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    parser.add_argument("--mode", choices=("check", "apply"), default="check")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = reconcile(
            catalog_path=args.catalog,
            transactions_path=args.transactions,
            output_dir=args.output_dir,
            mode=args.mode,
        )
    except ReconciliationError as error:
        print(f"inventory-reconcile: {error}")
        return 2
    print(json.dumps(result, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
