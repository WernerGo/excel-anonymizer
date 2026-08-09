"""Tests for the two settings that decide what survives a run.

``formulas`` decides whether the output holds formulas or their results,
and one of the two is always lost. ``strategy: scale`` makes amounts
fictitious without destroying the relations between them.
"""

import zipfile

import openpyxl
import pytest
import yaml
from pathlib import Path

from anonymize import anonymize, count_uncached_formulas, load_source, scale_value


@pytest.fixture
def excel_with_formulas(tmp_path: Path) -> Path:
    """A workbook whose amounts are partly typed in and partly calculated.

    openpyxl cannot write a calculated result, only the formula, so these
    cells stand for the case of a file that Excel has never recalculated.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.cell(1, 1, "name");   ws.cell(1, 2, "amount");  ws.cell(1, 3, "total")
    ws.cell(2, 1, "Smith");  ws.cell(2, 2, 1000);      ws.cell(2, 3, "=B2*2")
    ws.cell(3, 1, "Brown");  ws.cell(3, 2, 1500.5);    ws.cell(3, 3, "=B3*2")
    path = tmp_path / "amounts.xlsx"
    wb.save(path)
    return path


def write_config(tmp_path: Path, config: dict) -> Path:
    """Write a config and return its path."""
    path = tmp_path / "config.yaml"
    path.write_text(yaml.dump(config), encoding="utf-8")
    return path


# ---------------------------------------------------------------------------
# formulas: values | keep
# ---------------------------------------------------------------------------

def test_values_is_the_default(excel_with_formulas, tmp_path):
    """Most users want data out of the file, so resolving is the default."""
    config = write_config(tmp_path, {"groups": []})
    anonymize(excel_with_formulas, config)

    out = excel_with_formulas.with_stem(excel_with_formulas.stem + "_anonymized")
    ws = openpyxl.load_workbook(out)["Sheet1"]
    assert ws.cell(2, 3).value != "=B2*2", "the formula must not survive in values mode"


def test_keep_leaves_the_formulas_in_place(excel_with_formulas, tmp_path):
    """For a workbook that stays in use in Excel."""
    config = write_config(tmp_path, {"formulas": "keep", "groups": []})
    anonymize(excel_with_formulas, config)

    out = excel_with_formulas.with_stem(excel_with_formulas.stem + "_anonymized")
    ws = openpyxl.load_workbook(out)["Sheet1"]
    assert ws.cell(2, 3).value == "=B2*2"


def test_an_unknown_mode_is_refused(excel_with_formulas):
    """A typo here silently changes what the output contains."""
    with pytest.raises(ValueError, match="formulas"):
        load_source(excel_with_formulas, "cached")


def rewrite_sheets(source: Path, target: Path, change) -> Path:
    """Copy a workbook, passing every worksheet part through `change`."""
    with zipfile.ZipFile(source) as reader, zipfile.ZipFile(target, "w") as writer:
        for item in reader.infolist():
            data = reader.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                data = change(data)
            writer.writestr(item, data)
    return target


def test_a_formula_that_was_never_calculated_is_counted(excel_with_formulas, tmp_path):
    """A workbook Excel has never calculated stores the formula and no result.

    That is the case worth warning about: in `values` mode those cells
    arrive empty and nothing says they ever held anything.
    """
    never = rewrite_sheets(
        excel_with_formulas, tmp_path / "never.xlsx", lambda data: data.replace(b"<v />", b"")
    )

    assert count_uncached_formulas(never) == 2


def test_a_formula_whose_result_is_empty_is_not_counted(excel_with_formulas):
    """`=IFERROR(…, "")` is calculated and its answer is nothing.

    Excel records that as an empty value element and openpyxl reports it
    as None — the same as never calculated. Counting the two together
    made a well-formed workbook look broken: one real file reported 1061
    such cells, every one of them correct. openpyxl writes its own
    formula cells the same way, which is why this fixture stands for the
    case.
    """
    assert count_uncached_formulas(excel_with_formulas) == 0


def test_a_file_without_formulas_counts_none(tmp_path):
    """No warning where there is nothing to warn about."""
    wb = openpyxl.Workbook()
    wb.active.cell(1, 1, "plain")
    path = tmp_path / "plain.xlsx"
    wb.save(path)

    assert count_uncached_formulas(path) == 0


# ---------------------------------------------------------------------------
# strategy: scale
# ---------------------------------------------------------------------------

def test_an_integer_stays_an_integer():
    """Writing 1283.7 where the source held 1000 changes the column's shape."""
    assert scale_value(1000, 1.2837) == 1284
    assert isinstance(scale_value(1000, 1.2837), int)


def test_a_float_stays_a_float():
    assert scale_value(1500.5, 2) == 3001.0
    assert isinstance(scale_value(1500.5, 2), float)


def test_a_flag_is_not_a_number():
    """`active_flag` is 0 or 1 in these files, and scaling it is nonsense."""
    assert scale_value(True, 2) is None


def test_text_and_dates_are_left_alone():
    """A numeric column holding `#N/A` is a finding, not something to scale."""
    import datetime

    assert scale_value("#N/A", 2) is None
    assert scale_value(datetime.datetime(2026, 3, 31), 2) is None


def test_scaling_keeps_the_relation_between_two_amounts(excel_with_formulas, tmp_path):
    """The point of one factor per workbook: sums and ratios still work out."""
    config = write_config(
        tmp_path,
        {
            "groups": [
                {
                    "name": "amounts",
                    "strategy": "scale",
                    "factor": 2,
                    "columns": [{"sheet": "Sheet1", "col": "B", "data_from_row": 2}],
                }
            ]
        },
    )
    anonymize(excel_with_formulas, config)

    out = excel_with_formulas.with_stem(excel_with_formulas.stem + "_anonymized")
    ws = openpyxl.load_workbook(out)["Sheet1"]
    assert ws.cell(2, 2).value == 2000
    assert ws.cell(3, 2).value == 3001.0
    assert ws.cell(3, 2).value / ws.cell(2, 2).value == pytest.approx(1500.5 / 1000)


def test_a_scale_group_needs_a_factor(excel_with_formulas, tmp_path):
    """Without one the group would silently do nothing."""
    config = write_config(
        tmp_path,
        {"groups": [{"name": "amounts", "strategy": "scale", "columns": []}]},
    )
    with pytest.raises(ValueError, match="factor"):
        anonymize(excel_with_formulas, config)


def test_a_key_group_needs_a_prefix(excel_with_formulas, tmp_path):
    """Same reasoning from the other side."""
    config = write_config(
        tmp_path,
        {"groups": [{"name": "names", "columns": []}]},
    )
    with pytest.raises(ValueError, match="prefix"):
        anonymize(excel_with_formulas, config)


def test_an_unknown_strategy_is_refused(excel_with_formulas, tmp_path):
    """A misspelt strategy would leave the column untouched and unreported."""
    config = write_config(
        tmp_path,
        {"groups": [{"name": "amounts", "strategy": "scaled", "columns": []}]},
    )
    with pytest.raises(ValueError, match="strategy"):
        anonymize(excel_with_formulas, config)


def test_scale_groups_are_not_written_to_the_mapping(excel_with_formulas, tmp_path):
    """Rounding loses the remainder, so the mapping could not restore them anyway."""
    import json

    config = write_config(
        tmp_path,
        {
            "save_mapping": "map.json",
            "groups": [
                {
                    "name": "names",
                    "prefix": "NAME",
                    "columns": [{"sheet": "Sheet1", "col": "A", "data_from_row": 2}],
                },
                {
                    "name": "amounts",
                    "strategy": "scale",
                    "factor": 2,
                    "columns": [{"sheet": "Sheet1", "col": "B", "data_from_row": 2}],
                },
            ],
        },
    )
    anonymize(excel_with_formulas, config)

    mapping = json.loads((excel_with_formulas.parent / "map.json").read_text())
    assert set(mapping) == {"names"}


# ---------------------------------------------------------------------------
# ignore_sheets
# ---------------------------------------------------------------------------

@pytest.fixture
def excel_with_a_working_sheet(tmp_path: Path) -> Path:
    """A workbook with a data sheet and a working view repeating its values."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.cell(1, 1, "name");  ws.cell(2, 1, "Smith")

    view = wb.create_sheet("Overview")
    view.cell(1, 1, "name");  view.cell(2, 1, "Smith")

    path = tmp_path / "book.xlsx"
    wb.save(path)
    return path


def test_a_listed_sheet_is_removed(excel_with_a_working_sheet, tmp_path):
    """The whole point: a sheet that is gone cannot carry an original value."""
    config = write_config(
        tmp_path,
        {"ignore_sheets": [{"sheet": "Overview", "reason": "a view of Data"}], "groups": []},
    )
    anonymize(excel_with_a_working_sheet, config)

    out = excel_with_a_working_sheet.with_stem(excel_with_a_working_sheet.stem + "_anonymized")
    assert openpyxl.load_workbook(out).sheetnames == ["Data"]


def test_a_plain_list_of_names_works_too(excel_with_a_working_sheet, tmp_path):
    """The reason is for the reader; the tool does not insist on it."""
    config = write_config(tmp_path, {"ignore_sheets": ["Overview"], "groups": []})
    anonymize(excel_with_a_working_sheet, config)

    out = excel_with_a_working_sheet.with_stem(excel_with_a_working_sheet.stem + "_anonymized")
    assert openpyxl.load_workbook(out).sheetnames == ["Data"]


def test_removing_every_sheet_is_refused(excel_with_a_working_sheet, tmp_path):
    """An empty workbook cannot be saved, and would not be wanted anyway."""
    config = write_config(tmp_path, {"ignore_sheets": ["Data", "Overview"], "groups": []})

    with pytest.raises(ValueError, match="every sheet"):
        anonymize(excel_with_a_working_sheet, config)


def test_a_sheet_that_is_not_there_only_warns(excel_with_a_working_sheet, tmp_path, capsys):
    """One list serves several files of the same family; not all carry every sheet."""
    config = write_config(tmp_path, {"ignore_sheets": ["Gone", "Overview"], "groups": []})
    anonymize(excel_with_a_working_sheet, config)

    assert "not in the file" in capsys.readouterr().out


def test_a_removed_sheet_takes_its_values_with_it(excel_with_a_working_sheet, tmp_path):
    """Anonymizing the data sheet alone would leave the copy in the view."""
    config = write_config(
        tmp_path,
        {
            "ignore_sheets": ["Overview"],
            "groups": [
                {"name": "names", "prefix": "NAME", "columns": [{"sheet": "Data", "col": "A", "data_from_row": 2}]}
            ],
        },
    )
    anonymize(excel_with_a_working_sheet, config)

    out = excel_with_a_working_sheet.with_stem(excel_with_a_working_sheet.stem + "_anonymized")
    book = openpyxl.load_workbook(out)
    values = [cell.value for sheet in book.worksheets for row in sheet.iter_rows() for cell in row]
    assert "Smith" not in values


# ---------------------------------------------------------------------------
# keys and numbers
# ---------------------------------------------------------------------------

def test_a_key_group_leaves_numbers_alone(excel_with_formulas, tmp_path):
    """A column of amounts with one text cell in it must stay a column of amounts."""
    config = write_config(
        tmp_path,
        {
            "groups": [
                {
                    "name": "mixed",
                    "prefix": "X",
                    "columns": [{"sheet": "Sheet1", "col": "B", "data_from_row": 2}],
                }
            ]
        },
    )
    anonymize(excel_with_formulas, config)

    ws = openpyxl.load_workbook(excel_with_formulas.with_stem(excel_with_formulas.stem + "_anonymized"))["Sheet1"]
    assert ws.cell(2, 2).value == 1000
    assert ws.cell(3, 2).value == 1500.5


def test_include_numbers_says_so_explicitly(excel_with_formulas, tmp_path):
    """A tax number is numeric and still has to go — but the config has to ask."""
    config = write_config(
        tmp_path,
        {
            "groups": [
                {
                    "name": "tax",
                    "prefix": "TAXNO",
                    "include_numbers": True,
                    "columns": [{"sheet": "Sheet1", "col": "B", "data_from_row": 2}],
                }
            ]
        },
    )
    anonymize(excel_with_formulas, config)

    ws = openpyxl.load_workbook(excel_with_formulas.with_stem(excel_with_formulas.stem + "_anonymized"))["Sheet1"]
    assert ws.cell(2, 2).value.startswith("TAXNO")


# ---------------------------------------------------------------------------
# what a removed sheet and a resolved formula leave behind
# ---------------------------------------------------------------------------

def test_a_defined_name_pointing_at_a_removed_sheet_goes_with_it(tmp_path):
    """Otherwise Excel asks about links on every open and cannot resolve them."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active.cell(1, 1, "a")
    wb.create_sheet("Overview").cell(1, 1, "b")
    wb.defined_names["from_overview"] = DefinedName("from_overview", attr_text="Overview!$A$1")
    wb.defined_names["from_data"] = DefinedName("from_data", attr_text="Data!$A$1")
    path = tmp_path / "book.xlsx"
    wb.save(path)

    config = write_config(tmp_path, {"ignore_sheets": ["Overview"], "groups": []})
    anonymize(path, config)

    out = openpyxl.load_workbook(path.with_stem(path.stem + "_anonymized"))
    assert "from_overview" not in out.defined_names
    assert "from_data" in out.defined_names, "a name for a kept sheet must survive"


def test_a_quoted_sheet_name_is_recognised_too(tmp_path):
    """A sheet name with a space is quoted in a reference."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active.cell(1, 1, "a")
    wb.create_sheet("Auflistung PE").cell(1, 1, "b")
    wb.defined_names["listing"] = DefinedName("listing", attr_text="'Auflistung PE'!$A$1")
    path = tmp_path / "book.xlsx"
    wb.save(path)

    config = write_config(tmp_path, {"ignore_sheets": ["Auflistung PE"], "groups": []})
    anonymize(path, config)

    out = openpyxl.load_workbook(path.with_stem(path.stem + "_anonymized"))
    assert "listing" not in out.defined_names


def table_workbook(tmp_path: Path) -> Path:
    """A workbook with a table whose column carries a formula."""
    from openpyxl.worksheet.table import Table, TableColumn, TableFormula

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["amount", "double"])
    ws.append([1000, 2000])
    ws.append([2000, 4000])
    table = Table(
        displayName="Amounts",
        ref="A1:B3",
        tableColumns=[
            TableColumn(id=1, name="amount"),
            TableColumn(id=2, name="double", calculatedColumnFormula=TableFormula(attr_text="Amounts[[#This Row],[amount]]*2")),
        ],
    )
    ws.add_table(table)
    path = tmp_path / "table.xlsx"
    wb.save(path)
    return path


def test_a_resolved_column_leaves_no_column_formula_behind(tmp_path):
    """Excel flags every cell of the column as an inconsistent calculation."""
    path = table_workbook(tmp_path)
    config = write_config(tmp_path, {"groups": []})
    anonymize(path, config)

    out = openpyxl.load_workbook(path.with_stem(path.stem + "_anonymized"))
    formulas = [c.calculatedColumnFormula for t in out["Data"].tables.values() for c in t.tableColumns]
    assert formulas == [None, None]


def test_keeping_the_formulas_keeps_the_column_formula(tmp_path):
    """There the table and the cells still agree."""
    path = table_workbook(tmp_path)
    config = write_config(tmp_path, {"formulas": "keep", "groups": []})
    anonymize(path, config)

    out = openpyxl.load_workbook(path.with_stem(path.stem + "_anonymized"))
    formulas = [c.calculatedColumnFormula for t in out["Data"].tables.values() for c in t.tableColumns]
    assert any(formula is not None for formula in formulas)


# ---------------------------------------------------------------------------
# sheets nobody mentioned
# ---------------------------------------------------------------------------

def test_a_sheet_in_no_list_is_reported(excel_with_a_working_sheet, tmp_path, capsys):
    """The next export carries a sheet the config was never written against."""
    config = write_config(
        tmp_path,
        {"groups": [{"name": "n", "prefix": "N", "columns": [{"sheet": "Data", "col": "A"}]}]},
    )
    anonymize(excel_with_a_working_sheet, config)

    assert "Overview" in capsys.readouterr().out


def test_fail_refuses_to_write_a_file_with_an_unlisted_sheet(excel_with_a_working_sheet, tmp_path):
    """Where the output must be free of originals, unconsidered is not good enough."""
    config = write_config(
        tmp_path,
        {
            "unlisted_sheets": "fail",
            "groups": [{"name": "n", "prefix": "N", "columns": [{"sheet": "Data", "col": "A"}]}],
        },
    )
    with pytest.raises(ValueError, match="unaccounted"):
        anonymize(excel_with_a_working_sheet, config)


def test_an_ignored_sheet_counts_as_accounted_for(excel_with_a_working_sheet, tmp_path):
    """Dropping a sheet is a decision about it, so it satisfies the check."""
    config = write_config(
        tmp_path,
        {
            "unlisted_sheets": "fail",
            "ignore_sheets": ["Overview"],
            "groups": [{"name": "n", "prefix": "N", "columns": [{"sheet": "Data", "col": "A"}]}],
        },
    )
    anonymize(excel_with_a_working_sheet, config)

    out = excel_with_a_working_sheet.with_stem(excel_with_a_working_sheet.stem + "_anonymized")
    assert openpyxl.load_workbook(out).sheetnames == ["Data"]


# ---------------------------------------------------------------------------
# naming many sheets at once
# ---------------------------------------------------------------------------

@pytest.fixture
def excel_with_many_tabs(tmp_path: Path) -> Path:
    """Three numbered tabs of one shape, a data sheet, and an empty divider."""
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active.cell(1, 1, "name")
    wb.active.cell(2, 1, "Smith")
    for number in ("48", "49", "77"):
        tab = wb.create_sheet(number)
        tab.cell(1, 1, "comment")
        tab.cell(2, 1, f"note {number}")
    wb.create_sheet("AC-DC")
    path = tmp_path / "book.xlsx"
    wb.save(path)
    return path


def test_a_pattern_reaches_every_matching_sheet(excel_with_many_tabs, tmp_path):
    """TF1 has 55 cashflow tabs in production and TF8 over a hundred."""
    config = write_config(
        tmp_path,
        {
            "keep_sheets": ["Data", "AC-DC"],
            "groups": [
                {
                    "name": "comments",
                    "prefix": "TEXT",
                    "columns": [{"sheet_pattern": r"\d+", "col": "A", "data_from_row": 2}],
                }
            ],
        },
    )
    anonymize(excel_with_many_tabs, config)

    out = openpyxl.load_workbook(excel_with_many_tabs.with_stem(excel_with_many_tabs.stem + "_anonymized"))
    assert [out[tab].cell(2, 1).value.startswith("TEXT") for tab in ("48", "49", "77")] == [True] * 3
    assert out["Data"].cell(2, 1).value == "Smith", "the pattern must not reach beyond it"


def test_a_pattern_accounts_for_the_sheets_it_matches(excel_with_many_tabs, tmp_path):
    """Otherwise every tab would have to be listed twice."""
    config = write_config(
        tmp_path,
        {
            "unlisted_sheets": "fail",
            "keep_sheets": ["Data", "AC-DC"],
            "groups": [
                {
                    "name": "comments",
                    "prefix": "TEXT",
                    "columns": [{"sheet_pattern": r"\d+", "col": "A", "data_from_row": 2}],
                }
            ],
        },
    )
    anonymize(excel_with_many_tabs, config)


def test_empty_sheets_can_be_removed_as_a_class(excel_with_many_tabs, tmp_path):
    """The dividers are empty and some of them are named after a transaction."""
    config = write_config(
        tmp_path,
        {
            "ignore_sheets": [{"empty": True, "reason": "section dividers"}],
            "keep_sheets": ["Data"],
            "groups": [
                {
                    "name": "comments",
                    "prefix": "TEXT",
                    "columns": [{"sheet_pattern": r"\d+", "col": "A", "data_from_row": 2}],
                }
            ],
        },
    )
    anonymize(excel_with_many_tabs, config)

    out = openpyxl.load_workbook(excel_with_many_tabs.with_stem(excel_with_many_tabs.stem + "_anonymized"))
    assert "AC-DC" not in out.sheetnames
    assert set(out.sheetnames) == {"Data", "48", "49", "77"}


def test_a_pattern_that_matches_nothing_stops_the_run(excel_with_many_tabs, tmp_path):
    """The one mistake no other check catches.

    `'\\d+'` in single-quoted YAML is a literal backslash followed by a
    `d`, so it matches no sheet at all. The columns it named went
    unreplaced, and the configuration, the structural comparison and the
    residual check all reported the file as sound.
    """
    config = write_config(
        tmp_path,
        {
            "keep_sheets": ["Data", "AC-DC"],
            "groups": [
                {
                    "name": "comments",
                    "prefix": "TEXT",
                    "columns": [{"sheet_pattern": r"\\d+", "col": "A", "data_from_row": 2}],
                }
            ],
        },
    )
    with pytest.raises(ValueError, match="matches no sheet"):
        anonymize(excel_with_many_tabs, config)


def test_a_named_sheet_that_is_absent_from_a_group_stops_the_run(excel_with_many_tabs, tmp_path):
    """A sheet renamed between exports leaves its column unhandled."""
    config = write_config(
        tmp_path,
        {
            "keep_sheets": ["Data", "AC-DC"],
            "groups": [
                {
                    "name": "names",
                    "prefix": "NAME",
                    "columns": [{"sheet": "Daten", "col": "A", "data_from_row": 2}],
                }
            ],
        },
    )
    with pytest.raises(ValueError, match="matches no sheet"):
        anonymize(excel_with_many_tabs, config)
