"""
Excel anonymization via key replacement.

Replaces configured cell values with unique keys. Identical values within
a group always get the same key, preserving relationships between columns
(e.g. employee → manager links stay intact in the anonymized file).

Usage:
    python src/anonymize.py input.xlsx
    python src/anonymize.py input.xlsx --config my_config.yaml
    python src/anonymize.py input.xlsx --config examples/config_names.yaml
"""

import argparse
import json
import sys
from collections import OrderedDict
from pathlib import Path

import openpyxl
import yaml


def anonymize(excel_path: Path, config_path: Path) -> None:
    config = yaml.safe_load(config_path.read_text(encoding="utf-8"))

    suffix   = config.get("output_suffix", "_anonymized")
    map_file = config.get("save_mapping")
    out_path = excel_path.with_stem(excel_path.stem + suffix)

    wb = openpyxl.load_workbook(excel_path)
    full_mapping: dict[str, dict[str, str]] = {}

    for group in config.get("groups", []):
        group_name = group["name"]
        prefix     = group["prefix"]
        mapping: dict[str, str] = OrderedDict()
        counter  = 0
        cell_refs: list[tuple] = []

        for col_spec in group["columns"]:
            sheet_name = col_spec["sheet"]
            col_letter = col_spec["col"]
            data_from  = col_spec.get("data_from_row", 2)

            if sheet_name not in wb.sheetnames:
                print(f"  WARNING: sheet '{sheet_name}' not found, skipped.")
                continue

            ws      = wb[sheet_name]
            col_idx = openpyxl.utils.column_index_from_string(col_letter)

            for row_idx in range(data_from, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val  = str(cell.value).strip() if cell.value is not None else ""
                if not val or val in ("None", "nan"):
                    continue
                if val not in mapping:
                    counter += 1
                    mapping[val] = f"{prefix}{counter:04d}"
                cell_refs.append((ws, row_idx, col_idx, mapping[val]))

        for ws, row_idx, col_idx, new_val in cell_refs:
            ws.cell(row=row_idx, column=col_idx).value = new_val

        full_mapping[group_name] = mapping
        print(f"  Group '{group_name}': {len(mapping)} unique values replaced.")

    wb.save(out_path)
    print(f"Anonymized file : {out_path}")

    if map_file:
        map_path = excel_path.parent / map_file
        # Mapping contains original values – keep local, never commit
        map_path.write_text(
            json.dumps(full_mapping, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"Mapping saved   : {map_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Anonymize Excel files by replacing cell values with unique keys"
    )
    parser.add_argument("excel", type=Path, help="Input Excel file (.xlsx)")
    _base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent.parent
    parser.add_argument(
        "--config", type=Path,
        default=_base / "anonymize_config.yaml",
        help="Config file (default: anonymize_config.yaml in project root)",
    )
    args = parser.parse_args()

    if not args.excel.exists():
        print(f"ERROR: file not found: {args.excel}")
        raise SystemExit(1)
    if not args.config.exists():
        print(f"ERROR: config not found: {args.config}")
        raise SystemExit(1)

    print(f"Input : {args.excel}")
    print(f"Config: {args.config}")
    anonymize(args.excel, args.config)


if __name__ == "__main__":
    main()
