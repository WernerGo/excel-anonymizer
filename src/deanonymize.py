"""
Restore original values from an anonymized Excel file using a saved mapping.

Works with mapping files produced by both anonymize.py and faker_replace.py.
The mapping JSON has the structure: { group_name: { original: replacement } }.
This script inverts it and replaces all replacement values back to originals.

Usage:
    uv run src/deanonymize.py anonymized.xlsx --mapping anonymization_map.json
    uv run src/deanonymize.py file_faker.xlsx --mapping faker_map.json
"""

import argparse
import json
import sys
from pathlib import Path

import openpyxl


def deanonymize(excel_path: Path, mapping_path: Path) -> None:
    mapping_data: dict[str, dict[str, str]] = json.loads(
        mapping_path.read_text(encoding="utf-8")
    )

    # Invert all groups into one lookup: replacement → original
    reverse: dict[str, str] = {}
    for group_mapping in mapping_data.values():
        for original, replacement in group_mapping.items():
            reverse[replacement] = original

    wb = openpyxl.load_workbook(excel_path)
    replaced = 0

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = str(cell.value).strip() if cell.value is not None else ""
                if val in reverse:
                    cell.value = reverse[val]
                    replaced += 1

    out_path = excel_path.with_stem(excel_path.stem + "_restored")
    wb.save(out_path)
    print(f"Restored file : {out_path}  ({replaced} values restored)")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Restore original Excel values from an anonymization mapping"
    )
    parser.add_argument("excel", type=Path, help="Anonymized Excel file (.xlsx)")
    parser.add_argument(
        "--mapping", type=Path, required=True,
        help="JSON mapping file saved by anonymize.py or faker_replace.py",
    )
    args = parser.parse_args()

    if not args.excel.exists():
        print(f"ERROR: file not found: {args.excel}")
        raise SystemExit(1)
    if not args.mapping.exists():
        print(f"ERROR: mapping not found: {args.mapping}")
        raise SystemExit(1)

    deanonymize(args.excel, args.mapping)


if __name__ == "__main__":
    main()
