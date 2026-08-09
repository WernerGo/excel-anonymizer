"""
Excel anonymization via key replacement.

Replaces configured cell values with unique keys. Identical values within
a group always get the same key, preserving relationships between columns
(e.g. employee → manager links stay intact in the anonymized file).

Numeric columns can be scaled by a factor instead of replaced, so that
amounts become fictitious while the relationships between them survive.

Usage:
    python src/anonymize.py input.xlsx
    python src/anonymize.py input.xlsx --config my_config.yaml
    python src/anonymize.py input.xlsx --config examples/config_names.yaml
"""

import argparse
import json
import re
import sys
import zipfile
from collections import OrderedDict
from pathlib import Path

import openpyxl
import yaml

FORMULA_MODES = ("values", "keep")


def load_source(excel_path: Path, mode: str) -> openpyxl.Workbook:
    """Open the workbook, resolving formulas or keeping them.

    An .xlsx stores two things for a formula cell: the formula, and the
    result Excel last calculated for it. openpyxl loads one or the other,
    and writes back whatever it loaded. So the choice is real and it is
    lossy either way:

    ``values``
        Formulas are replaced by their last results. The output is a
        plain value workbook — what a program that reads values wants,
        and reproducible because nothing recalculates.
    ``keep``
        Formulas survive, their results do not. The file has to be
        opened and recalculated in Excel before anything can read a
        value from it.

    Args:
        excel_path: Path to the input Excel file (.xlsx).
        mode: ``values`` or ``keep``.

    Returns:
        The loaded workbook.
    """
    if mode not in FORMULA_MODES:
        raise ValueError(f"formulas: {mode!r} — expected one of {', '.join(FORMULA_MODES)}")
    if mode == "keep":
        print("  WARNING: formulas are kept, so their last results are dropped.")
        print("           Open the output in Excel and recalculate before reading values from it.")
    return openpyxl.load_workbook(excel_path, data_only=(mode == "values"))


CELL = re.compile(rb"<c\b[^>]*>(?:(?!</c>).)*</c>", re.S)


def count_uncached_formulas(excel_path: Path) -> int:
    """Count formula cells that were never calculated.

    Not the same as a formula whose result is empty. ``=IFERROR(…, "")``
    is calculated and its answer is nothing, which the file records as an
    empty value element — the cell is correct and the output should be
    empty there too. A cell that was never calculated has no value
    element at all, and in ``values`` mode it arrives empty without
    anything saying it ever held something.

    openpyxl reports both as None, so this reads the stored XML, where
    the two are plainly different.

    Args:
        excel_path: Path to the input Excel file (.xlsx).

    Returns:
        The number of formula cells that carry no result of any kind.
    """
    missing = 0
    with zipfile.ZipFile(excel_path) as archive:
        sheets = [name for name in archive.namelist() if name.startswith("xl/worksheets/sheet")]
        for sheet in sheets:
            for match in CELL.finditer(archive.read(sheet)):
                cell = match.group(0)
                if b"<f" in cell and b"<v" not in cell and b"<is" not in cell:
                    missing += 1
    return missing


def drop_sheets(wb: openpyxl.Workbook, entries: list) -> None:
    """Remove sheets that are not part of what is being passed on.

    A workbook usually carries more than the data someone needs from it:
    working views, report layouts, exports to another system. Those
    sheets hold the same values as the data sheets, so anonymizing the
    data sheets alone leaves the file full of originals. Removing them
    is both simpler and more complete than replacing them column by
    column.

    Each entry names a sheet and, where it is written that way, the
    reason it can go — the list is meant to be read by whoever has to
    agree that these sheets are dispensable.

    Args:
        wb: The open workbook.
        entries: Sheet names, or mappings with ``sheet`` and ``reason``.
    """
    names = [entry["sheet"] if isinstance(entry, dict) else entry for entry in entries]
    names = [str(name) for name in names]

    if names and not set(wb.sheetnames) - set(names):
        raise ValueError("ignore_sheets would remove every sheet")

    for name in names:
        if name not in wb.sheetnames:
            print(f"  WARNING: sheet '{name}' listed for removal but not in the file.")
            continue
        del wb[name]

    removed = [name for name in names if name not in wb.sheetnames]
    if removed:
        print(f"  Removed {len(removed)} sheets that are not part of the delivery.")
        drop_dangling_names(wb, removed)


def drop_dangling_names(wb: openpyxl.Workbook, removed: list[str]) -> None:
    """Remove defined names that point at a sheet that is no longer there.

    Deleting a sheet leaves the names that referred to it behind, and
    Excel asks about links on every open because it cannot resolve them.
    The names are part of the removed sheet's machinery, so they go with
    it.

    Args:
        wb: The open workbook.
        removed: The names of the sheets that were deleted.
    """
    targets = {f"{name}!" for name in removed} | {f"'{name}'!" for name in removed}

    dangling = [
        key
        for key, defined in wb.defined_names.items()
        if any(target in str(defined.value) for target in targets)
    ]
    for key in dangling:
        del wb.defined_names[key]

    for ws in wb.worksheets:
        local = [
            key
            for key, defined in ws.defined_names.items()
            if any(target in str(defined.value) for target in targets)
        ]
        for key in local:
            del ws.defined_names[key]
        dangling.extend(local)

    if dangling:
        print(f"  Removed {len(dangling)} defined names that pointed at them.")


def clear_calculated_columns(wb: openpyxl.Workbook) -> None:
    """Forget the column formulas of tables whose cells now hold values.

    A table column can carry a formula that Excel expects every cell of
    that column to repeat. Once the formulas have been resolved to their
    results, the cells are constants and the table still asks for the
    formula, which Excel flags on every one of them as an inconsistent
    column. The formula is no longer true of the file, so it goes.

    Args:
        wb: The open workbook.
    """
    cleared = 0
    for ws in wb.worksheets:
        for table in ws.tables.values():
            for column in table.tableColumns:
                if column.calculatedColumnFormula is not None:
                    column.calculatedColumnFormula = None
                    cleared += 1
    if cleared:
        print(f"  Cleared {cleared} table column formulas that no longer describe the cells.")


def check_every_sheet_is_accounted_for(wb: openpyxl.Workbook, config: dict, mode: str) -> None:
    """Report sheets that neither a group nor `ignore_sheets` mentions.

    A configuration is written against the file you had. The next export
    can carry a sheet nobody has seen, and it passes through untouched —
    with its values — because nothing refers to it. That is the failure
    this guards: not a wrong rule, but a missing one.

    A sheet counts as accounted for if a group names it, if
    ``ignore_sheets`` removes it, or if ``keep_sheets`` says it stays as
    it is. The third is needed as often as the other two: a sheet can be
    empty, or hold nothing but a list of currency codes.

    ``warn`` names them and continues; ``fail`` refuses. Use ``fail``
    wherever the output is supposed to be free of original values, which
    can only be claimed for a file whose every sheet was considered.

    Args:
        wb: The workbook, after the ignored sheets have been removed.
        config: The parsed configuration.
        mode: ``warn`` or ``fail``.

    Raises:
        ValueError: In ``fail`` mode, if any sheet is unaccounted for.
    """
    if mode not in ("warn", "fail"):
        raise ValueError(f"unlisted_sheets: {mode!r} — expected 'warn' or 'fail'")

    named = {
        str(col["sheet"])
        for group in config.get("groups", [])
        for col in group.get("columns", [])
    }
    named |= {
        str(entry["sheet"]) if isinstance(entry, dict) else str(entry)
        for entry in config.get("keep_sheets", [])
    }
    unlisted = [name for name in wb.sheetnames if name not in named]
    if not unlisted:
        return

    print(f"  {len(unlisted)} sheets are in the file but in no group and not in ignore_sheets:")
    for name in unlisted:
        print(f"    {name}")
    if mode == "fail":
        raise ValueError(
            f"{len(unlisted)} sheets are unaccounted for; list them in a group or in ignore_sheets"
        )


def scale_value(value: object, factor: float) -> object | None:
    """Return the value multiplied by the factor, keeping the type it had.

    An integer stays an integer: writing 1283.7 where the source held
    1000 changes the shape of the column, not just its content, and any
    check comparing the two files would report it.

    Booleans are numbers in Python but flags in a spreadsheet, so they
    are left alone.

    Args:
        value: The original cell value.
        factor: The factor to multiply by.

    Returns:
        The scaled value, or None if the cell holds nothing numeric.
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    scaled = value * factor
    return int(round(scaled)) if isinstance(value, int) else round(scaled, 6)


def iter_cells(wb: openpyxl.Workbook, columns: list[dict]) -> list[tuple]:
    """Yield every cell of the configured columns as (worksheet, row, column).

    Args:
        wb: The open workbook.
        columns: The ``columns`` entries of one group.

    Returns:
        A list of (worksheet, row index, column index) tuples.
    """
    cells: list[tuple] = []
    for col_spec in columns:
        sheet_name = str(col_spec["sheet"])
        col_letter = col_spec["col"]
        data_from = col_spec.get("data_from_row", 2)

        if sheet_name not in wb.sheetnames:
            print(f"  WARNING: sheet '{sheet_name}' not found, skipped.")
            continue

        ws = wb[sheet_name]
        col_idx = openpyxl.utils.column_index_from_string(col_letter)
        for row_idx in range(data_from, ws.max_row + 1):
            cells.append((ws, row_idx, col_idx))
    return cells


def apply_key_group(wb: openpyxl.Workbook, group: dict) -> dict[str, str]:
    """Replace the values of one group with stable keys.

    Args:
        wb: The open workbook.
        group: One entry of the ``groups`` list.

    Returns:
        The mapping from original value to key.
    """
    if "prefix" not in group:
        raise ValueError(f"group {group['name']!r}: a key group needs a prefix")

    prefix = group["prefix"]
    numbers_too = group.get("include_numbers", False)
    mapping: dict[str, str] = OrderedDict()
    replacements: list[tuple] = []

    for ws, row_idx, col_idx in iter_cells(wb, group["columns"]):
        cell = ws.cell(row=row_idx, column=col_idx)
        if isinstance(cell.value, (int, float)) and not numbers_too:
            # A key is text. Writing one over a number changes the type of
            # the column, which is a change to the file's shape rather than
            # to its content — and columns holding text among numbers are
            # common enough that keying them silently would be a trap. A
            # group that does mean to replace numbers says `include_numbers`.
            continue
        val = str(cell.value).strip() if cell.value is not None else ""
        if not val or val in ("None", "nan"):
            continue
        if val not in mapping:
            mapping[val] = f"{prefix}{len(mapping) + 1:04d}"
        replacements.append((ws, row_idx, col_idx, mapping[val]))

    for ws, row_idx, col_idx, new_val in replacements:
        ws.cell(row=row_idx, column=col_idx).value = new_val

    print(f"  Group '{group['name']}': {len(mapping)} unique values replaced.")
    return mapping


def apply_scale_group(wb: openpyxl.Workbook, group: dict) -> None:
    """Multiply the numbers of one group by the group's factor.

    One factor for a whole workbook keeps the relations between amounts
    intact — a commitment stays below its tranche total, and a sum of
    cashflows still reconciles. Ratios must therefore not be scaled: a
    set of weights adding up to 1.0 would add up to the factor. Which
    columns are amounts and which are ratios is a decision about the
    data, which is why it is made in the configuration.

    Scaling is not reversible: rounding to the source type loses the
    remainder. Scale groups are therefore not written to the mapping
    file.

    Args:
        wb: The open workbook.
        group: One entry of the ``groups`` list.
    """
    factor = group.get("factor")
    if factor is None:
        raise ValueError(f"group {group['name']!r}: a scale group needs a factor")

    replacements: list[tuple] = []
    skipped = 0

    for ws, row_idx, col_idx in iter_cells(wb, group["columns"]):
        value = ws.cell(row=row_idx, column=col_idx).value
        if value is None or value == "":
            continue
        scaled = scale_value(value, factor)
        if scaled is None:
            skipped += 1
            continue
        replacements.append((ws, row_idx, col_idx, scaled))

    for ws, row_idx, col_idx, new_val in replacements:
        ws.cell(row=row_idx, column=col_idx).value = new_val

    note = f", {skipped} non-numeric cells left alone" if skipped else ""
    print(f"  Group '{group['name']}': {len(replacements)} numbers scaled by {factor}{note}.")


def anonymize(excel_path: Path, config_path: Path) -> None:
    """Replace configured cell values with stable keys and optionally save the mapping.

    Args:
        excel_path: Path to the input Excel file (.xlsx).
        config_path: Path to the YAML config defining groups, prefixes, and columns.
    """
    config = yaml.safe_load(config_path.read_text(encoding="utf-8"))

    suffix = config.get("output_suffix", "_anonymized")
    map_file = config.get("save_mapping")
    formulas = config.get("formulas", "values")
    out_path = excel_path.with_stem(excel_path.stem + suffix)

    wb = load_source(excel_path, formulas)

    if formulas == "values":
        uncached = count_uncached_formulas(excel_path)
        if uncached:
            print(f"  WARNING: {uncached} formula cells carry no calculated result and arrive empty.")
            print("           Open the source in Excel, let it recalculate, save, and run again.")

    drop_sheets(wb, config.get("ignore_sheets", []))
    check_every_sheet_is_accounted_for(wb, config, config.get("unlisted_sheets", "warn"))

    if formulas == "values":
        clear_calculated_columns(wb)

    full_mapping: dict[str, dict[str, str]] = {}

    for group in config.get("groups", []):
        strategy = group.get("strategy", "key")
        if strategy == "key":
            full_mapping[group["name"]] = apply_key_group(wb, group)
        elif strategy == "scale":
            apply_scale_group(wb, group)
        else:
            raise ValueError(f"group {group['name']!r}: unknown strategy {strategy!r}")

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
    """Parse CLI arguments and run the anonymization."""
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
