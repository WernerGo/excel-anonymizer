"""Tests for the two settings that decide what survives a run.

``formulas`` decides whether the output holds formulas or their results,
and one of the two is always lost. ``strategy: scale`` makes amounts
fictitious without destroying the relations between them.
"""

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


def test_formula_cells_without_a_result_are_counted(excel_with_formulas):
    """Silent emptying is the failure mode; a number in the log is the guard."""
    assert count_uncached_formulas(excel_with_formulas) == 2


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
