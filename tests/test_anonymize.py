"""Tests for anonymize.py: key-based replacement."""

import json
import openpyxl
import pytest
from pathlib import Path

from anonymize import anonymize, output_path


def _load_output(excel_path: Path, suffix: str) -> openpyxl.worksheet.worksheet.Worksheet:
    out = excel_path.with_stem(excel_path.stem + suffix)
    return openpyxl.load_workbook(out)["Sheet1"]


# ---------------------------------------------------------------------------
# Basic replacement
# ---------------------------------------------------------------------------

def test_original_values_are_gone(sample_excel, anon_config):
    anonymize(sample_excel, anon_config)
    ws = _load_output(sample_excel, "_anonymized")
    data_values = [ws.cell(r, c).value for r in (2, 3) for c in (1, 2, 3, 4)]
    assert "Smith" not in data_values
    assert "Brown" not in data_values
    assert "John"  not in data_values
    assert "Jane"  not in data_values


def test_keys_have_correct_prefix(sample_excel, anon_config):
    anonymize(sample_excel, anon_config)
    ws = _load_output(sample_excel, "_anonymized")
    assert ws.cell(2, 1).value.startswith("LAST")
    assert ws.cell(2, 2).value.startswith("FIRST")


# ---------------------------------------------------------------------------
# Relationship preservation
# ---------------------------------------------------------------------------

def test_same_value_gets_same_key_across_columns(sample_excel, anon_config):
    """Smith in col A (employee) and col C (manager) must get the same key."""
    anonymize(sample_excel, anon_config)
    ws = _load_output(sample_excel, "_anonymized")
    # A2 = "Smith" key, C3 = "Smith" key → must be equal
    assert ws.cell(2, 1).value == ws.cell(3, 3).value
    # A3 = "Brown" key, C2 = "Brown" key → must be equal
    assert ws.cell(3, 1).value == ws.cell(2, 3).value


def test_different_values_get_different_keys(sample_excel, anon_config):
    anonymize(sample_excel, anon_config)
    ws = _load_output(sample_excel, "_anonymized")
    assert ws.cell(2, 1).value != ws.cell(3, 1).value  # Smith ≠ Brown


def test_groups_are_independent(sample_excel, anon_config):
    """last_names and first_names groups have separate key spaces."""
    anonymize(sample_excel, anon_config)
    ws = _load_output(sample_excel, "_anonymized")
    last_keys  = {ws.cell(r, c).value for r in (2, 3) for c in (1, 3)}
    first_keys = {ws.cell(r, c).value for r in (2, 3) for c in (2, 4)}
    assert last_keys.isdisjoint(first_keys), "Key spaces of different groups must not overlap"


# ---------------------------------------------------------------------------
# Header rows untouched
# ---------------------------------------------------------------------------

def test_header_row_is_not_anonymized(sample_excel, anon_config):
    anonymize(sample_excel, anon_config)
    ws = _load_output(sample_excel, "_anonymized")
    assert ws.cell(1, 1).value == "Last Name"
    assert ws.cell(1, 2).value == "First Name"
    assert ws.cell(1, 3).value == "Mgr Last"
    assert ws.cell(1, 4).value == "Mgr First"


# ---------------------------------------------------------------------------
# Mapping file
# ---------------------------------------------------------------------------

def test_mapping_file_is_saved(sample_excel, anon_config):
    anonymize(sample_excel, anon_config)
    map_path = sample_excel.parent / "anon_map.json"
    assert map_path.exists()


def test_mapping_contains_original_values(sample_excel, anon_config):
    anonymize(sample_excel, anon_config)
    mapping = json.loads((sample_excel.parent / "anon_map.json").read_text())
    assert "Smith" in mapping["last_names"]
    assert "Brown" in mapping["last_names"]
    assert "John"  in mapping["first_names"]
    assert "Jane"  in mapping["first_names"]


def test_mapping_values_match_output(sample_excel, anon_config):
    """Keys in the mapping must equal what's written in the output file."""
    anonymize(sample_excel, anon_config)
    mapping = json.loads((sample_excel.parent / "anon_map.json").read_text())
    ws = _load_output(sample_excel, "_anonymized")
    assert ws.cell(2, 1).value == mapping["last_names"]["Smith"]
    assert ws.cell(3, 1).value == mapping["last_names"]["Brown"]


# ---------------------------------------------------------------------------
# Roundtrip: anonymize → deanonymize
# ---------------------------------------------------------------------------

def test_roundtrip_restores_original_values(sample_excel, anon_config):
    from deanonymize import deanonymize

    anonymize(sample_excel, anon_config)
    anon_path = sample_excel.with_stem(sample_excel.stem + "_anonymized")
    map_path  = sample_excel.parent / "anon_map.json"

    deanonymize(anon_path, map_path)
    restored_path = anon_path.with_stem(anon_path.stem + "_restored")

    original = openpyxl.load_workbook(sample_excel)["Sheet1"]
    restored = openpyxl.load_workbook(restored_path)["Sheet1"]

    for row in range(1, 4):
        for col in range(1, 5):
            assert original.cell(row, col).value == restored.cell(row, col).value


# ---------------------------------------------------------------------------
# Where the copy is written
# ---------------------------------------------------------------------------

def test_an_xlsx_keeps_its_extension(tmp_path):
    assert output_path(tmp_path / "book.xlsx", "_anonymised").name == "book_anonymised.xlsx"


def test_an_xlsm_is_written_as_an_xlsx(tmp_path):
    """openpyxl does not carry the macro project over, so an .xlsm output
    would claim to be macro-enabled and hold no macros — the file Excel
    offers to repair."""
    assert output_path(tmp_path / "book.xlsm", "_anonymised").name == "book_anonymised.xlsx"


def test_a_macro_workbook_produces_a_readable_copy(sample_excel, anon_config):
    macro_file = sample_excel.with_suffix(".xlsm")
    macro_file.write_bytes(sample_excel.read_bytes())

    anonymize(macro_file, anon_config)

    out = macro_file.with_stem(macro_file.stem + "_anonymized").with_suffix(".xlsx")
    assert out.exists()
    assert not macro_file.with_stem(macro_file.stem + "_anonymized").exists()
    ws = openpyxl.load_workbook(out)["Sheet1"]
    assert ws.cell(2, 1).value.startswith("LAST"), "the data was anonymised as usual"


# ---------------------------------------------------------------------------
# A configuration that no longer fits the file
# ---------------------------------------------------------------------------

def _config_with_missing_sheets(tmp_path: Path) -> Path:
    """Three groups, four column entries, three of them naming sheets that are gone."""
    import yaml
    config = {
        "output_suffix": "_anonymized",
        "groups": [
            {
                "name": "last_names",
                "prefix": "LAST",
                "columns": [
                    {"sheet": "Sheet1", "col": "A", "data_from_row": 2},
                    {"sheet": "positions", "col": "B", "data_from_row": 2},
                ],
            },
            {
                "name": "first_names",
                "prefix": "FIRST",
                "columns": [{"sheet": "positions", "col_from": "C", "data_from_row": 2}],
            },
            {
                "name": "amounts",
                "strategy": "scale",
                "factor": 1.2,
                "columns": [{"sheet": "transfer", "col": "D", "data_from_row": 2}],
            },
        ],
    }
    path = tmp_path / "stale_config.yaml"
    path.write_text(yaml.dump(config), encoding="utf-8")
    return path


def test_every_missing_sheet_is_named_not_only_the_first(sample_excel, tmp_path, capsys):
    """A delivery that moved five things must not take five runs to find out.

    The run still stops — a column entry that reaches nothing leaves the
    values it was meant to replace in the file — but it stops after
    holding the whole configuration against the workbook.
    """
    config = _config_with_missing_sheets(tmp_path)

    with pytest.raises(ValueError) as caught:
        anonymize(sample_excel, config)

    printed = capsys.readouterr().out
    assert "'positions'" in printed
    assert "'transfer'" in printed
    assert "last_names, column B" in printed
    assert "first_names, column C onwards" in printed
    assert "amounts, column D" in printed
    assert "3 column entries match no sheet in the file (2 sheets)" in printed
    assert "match no sheet in the file" in str(caught.value)


def test_nothing_is_written_when_the_configuration_does_not_fit(sample_excel, tmp_path):
    """Half an anonymized file is worse than none."""
    config = _config_with_missing_sheets(tmp_path)

    with pytest.raises(ValueError):
        anonymize(sample_excel, config)

    assert not output_path(sample_excel, "_anonymized").exists()


def test_a_configuration_that_fits_is_not_reported(sample_excel, anon_config):
    """The check has to stay quiet when there is nothing to say."""
    import yaml
    from anonymize import unreachable_columns
    import openpyxl as _openpyxl

    wb = _openpyxl.load_workbook(sample_excel)
    config = yaml.safe_load(anon_config.read_text(encoding="utf-8"))

    assert unreachable_columns(wb, config) == {}


def test_both_kinds_of_problem_are_reported_in_one_run(sample_excel, tmp_path, capsys):
    """A new delivery is wrong in more than one way at once.

    TF6 on 9.10.2026: three entries of `ignore_sheets` naming sheets that
    are gone, column entries reaching nothing, and seven sheets nobody
    has decided about. The run stopped at the second kind, so the third
    was never reached and the list the repair needs took three runs to
    assemble.
    """
    import yaml
    config = {
        "output_suffix": "_anonymized",
        "unlisted_sheets": "fail",
        "groups": [
            {
                "name": "last_names",
                "prefix": "LAST",
                "columns": [{"sheet": "positions", "col": "A", "data_from_row": 2}],
            }
        ],
    }
    path = tmp_path / "two_kinds.yaml"
    path.write_text(yaml.dump(config), encoding="utf-8")

    with pytest.raises(ValueError) as caught:
        anonymize(sample_excel, path)

    printed = capsys.readouterr().out
    assert "match no sheet in the file" in printed      # the column entry
    assert "in no group and not in ignore_sheets" in printed  # Sheet1, unaccounted for
    assert "match no sheet in the file" in str(caught.value)
    assert "unaccounted for" in str(caught.value)


def test_the_names_of_unlisted_sheets_stay_in_the_file(sample_excel, tmp_path, capsys):
    """A sheet is often named after the transaction it holds.

    The report says how many and where the list is; it does not print
    the names. What it does print is the sheets named in the
    *configuration*, which are ours.
    """
    import yaml
    config = {
        "output_suffix": "_anonymized",
        "unlisted_sheets": "fail",
        "groups": [{"name": "n", "prefix": "N", "columns": [{"sheet": "Sheet1", "col": "A"}]}],
    }
    path = tmp_path / "all_listed.yaml"
    path.write_text(yaml.dump(config), encoding="utf-8")

    anonymize(sample_excel, path)

    assert "Sheet1" not in capsys.readouterr().out.split("Input :")[-1].split("Config:")[-1]


def test_a_column_entry_that_finds_nothing_stops_the_run(sample_excel, tmp_path, capsys):
    """A list that has moved looks exactly like this.

    `ig_statics` was rebuilt between two deliveries on 9.10.2026, the
    company lists moved a few columns and gained a heading, and the
    configuration went on naming the old letters. They were empty, so
    the group replaced nothing — and nothing complained until somebody
    read the derived file and found 1,624 LEIs in it.

    **It does not stop the run.** A column that exists and holds nothing
    is ordinary, and stopping on it would block every run for a
    condition that is usually harmless. What was missing was that
    anybody could see it.
    """
    import yaml
    config = {
        "output_suffix": "_anonymized",
        "groups": [
            {
                "name": "last_names",
                "prefix": "LAST",
                "columns": [
                    {"sheet": "Sheet1", "col": "A", "data_from_row": 2},
                    {"sheet": "Sheet1", "col": "Z", "data_from_row": 2},
                ],
            }
        ],
    }
    path = tmp_path / "moved.yaml"
    path.write_text(yaml.dump(config), encoding="utf-8")

    anonymize(sample_excel, path)   # reported, not fatal

    printed = capsys.readouterr().out
    assert "group 'last_names'" in printed
    assert "Sheet1 column Z" in printed
    assert "Sheet1 column A" not in printed


def test_an_entry_full_on_one_sheet_and_empty_on_another_is_not_reported(tmp_path):
    """A scratch range is empty on some tabs and full on the next."""
    import yaml
    from anonymize import entries_that_reach_nothing

    wb = openpyxl.Workbook()
    full = wb.active
    full.title = "1"
    full.cell(1, 1, "h"); full.cell(2, 1, "Smith")
    empty = wb.create_sheet("2")
    empty.cell(1, 1, "h")
    path = tmp_path / "tabs.xlsx"
    wb.save(path)

    config = {"groups": [{"name": "n", "prefix": "N",
                          "columns": [{"sheet_pattern": r"\d+", "col": "A", "data_from_row": 2}]}]}

    assert entries_that_reach_nothing(openpyxl.load_workbook(path), config) == {}


def test_a_column_entry_can_spare_a_heading_by_name():
    """A range says *everything right of here*, except what it names.

    TF8's cashflow tabs hold `ignore.percentagePIKmargin` inside the
    scratch area, and the migration reads it. Swept with the rest, ten of
    its cells came out as a key and the database refused six PIK steps
    for a rate that is not a number (12.10.2026). A letter will not spare
    it — the column stands at AN in 93 of 120 tabs and elsewhere in eight
    more — so the entry names the heading.
    """
    from anonymize import column_range

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.append(["a", "b", "ignore.percentagePIKmargin", "d"])
    sheet.append([1, 2, 3, 4])

    everything = list(column_range({"col_from": "B", "data_from_row": 2}, sheet))
    spared = list(
        column_range(
            {"col_from": "B", "data_from_row": 2, "except_headers": ["ignore.percentagePIKmargin"]},
            sheet,
        )
    )

    assert everything == [2, 3, 4]
    assert spared == [2, 4]


def test_a_spared_heading_is_found_wherever_it_stands():
    """The point of naming it: the tabs do not agree on where it sits."""
    from anonymize import column_range

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.append(["a", "pik", "c", "d"])
    sheet.append([1, 2, 3, 4])

    spared = list(column_range({"col_from": "A", "data_from_row": 2, "except_headers": ["PIK"]}, sheet))

    assert spared == [1, 3, 4], "matched without regard to case or position"


def test_a_formula_error_is_not_given_a_key():
    """It names nobody, and a key over it hides a defect of the file.

    Six cells of the PE tracker's `BNP ID` hold `#N/A`. Keyed like the
    rest of that column, the derived file showed six well-formed
    identifiers and the migration run against it reported nothing; the
    run against the real file reported all six (13.10.2026).
    """
    from anonymize import apply_key_group

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "t"
    sheet.append(["id"])
    sheet.append(["#N/A"])
    sheet.append(["a real one"])

    apply_key_group(book, {"name": "ids", "prefix": "ID", "columns": [{"sheet": "t", "col": "A", "data_from_row": 2}]})

    assert sheet["A2"].value == "#N/A"
    assert sheet["A3"].value == "ID0001"
