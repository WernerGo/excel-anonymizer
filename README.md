# excel-anonymizer

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Anonymize Excel files while preserving relationships across columns.

**Key property:** identical values across all configured columns always get the same replacement — so relationships (e.g. employee → manager, or foreign keys across sheets) stay intact in the anonymized file.

## Motivation

Excel is still the dominant data format in many organizations — reporting, controlling, HR, and operations all live in spreadsheets. AI tools are increasingly useful for analyzing that data, but they introduce a new risk: to build or improve an analysis script with AI assistance, you often have to share the data with a developer or paste it into a prompt — exposing names, salaries, or other sensitive information that most people should never see.

The safest approach is to keep both code and data local. But even then, when you collaborate with a developer or ask an AI to help write analysis code, you need realistic data that behaves like the real thing — without being the real thing.

**excel-anonymizer** solves this: replace sensitive values in your Excel files with either abstract keys or realistic-looking fakes, share or use the result freely, and restore the originals any time via the saved mapping.

## Use case

You have an Excel file with personal data (names, departments, locations) that you want to share for testing, debugging, or review — without exposing real data. Classic tools either target databases or don't preserve cross-column relationships.

## Tools

| Script | What it does |
|--------|--------------|
| `src/generate_sample.py` | Generate a realistic fake Excel from a config (no real data needed) |
| `src/anonymize.py` | Replace values with abstract keys (`NAME0001`, `DEPT0002`, …) |
| `src/faker_replace.py` | Replace values with realistic fake names/cities/companies via Faker |
| `src/deanonymize.py` | Restore original values from a mapping file saved by `anonymize.py` or `faker_replace.py` |

## Installation

```bash
# Install uv (once, no admin rights required, no Python needed beforehand — uv manages Python itself)
# Mac/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# Install dependencies
uv sync
```

## Quick start

```bash
# 1. Generate sample data to try it out (no real Excel needed)
uv run src/generate_sample.py --config examples/config_faker.yaml

# 2a. Anonymize with abstract keys
uv run src/anonymize.py sample_data.xlsx --config examples/config_names.yaml

# 2b. Or replace with realistic fake data
uv run src/faker_replace.py sample_data.xlsx --config examples/config_faker.yaml
```

## anonymize.py — key replacement

Replaces each unique value with a stable key (`PREFIX0001`, `PREFIX0002`, …).

```bash
uv run src/anonymize.py input.xlsx
uv run src/anonymize.py input.xlsx --config examples/config_names.yaml
```

Output: `input_anonymized.xlsx` next to the original.

Config:

```yaml
output_suffix: "_anonymized"
save_mapping: "anonymization_map.json"   # optional

groups:
  - name: person_names
    prefix: "NAME"
    columns:
      - sheet: Employees
        col: B               # column letter
        data_from_row: 2     # first data row (skip headers)
      - sheet: Employees
        col: D               # same mapping → relationship preserved
        data_from_row: 2
```

### Formulas: what the output contains

An `.xlsx` stores two things for a formula cell — the formula, and the
result Excel last calculated for it. openpyxl loads one or the other and
writes back whatever it loaded, so one of the two is always lost:

```yaml
formulas: values   # default
```

| Mode | Output holds | Lost |
|------|--------------|------|
| `values` | the calculated results, as plain values | the formulas |
| `keep`   | the formulas | their results — the file must be recalculated in Excel before anything can read a value |

`values` is the default because the usual reason to anonymize a file is
to pass its data on, and a program reading the output wants values. It
also makes the output reproducible: nothing recalculates.

A workbook Excel has never calculated stores its formulas without any
result, and those cells arrive **empty** in `values` mode with nothing
saying they ever held something. The run counts them and warns:

```
WARNING: 930 formula cells carry no calculated result and arrive empty.
         Open the source in Excel, let it recalculate, save, and run again.
```

This counts only cells with no result at all. A formula whose result
*is* empty — `=IFERROR(…, "")` is the common one — is calculated, and
the file records that as an empty value element. openpyxl reports both
as None, so the check reads the stored XML, where they differ. Treating
them alike made a healthy workbook report a thousand broken cells.

### Sheets that should not travel at all

A workbook usually carries more than its data: working views, report
layouts, exports to another system. Those sheets repeat the values of
the data sheets in a different arrangement, so anonymizing the data
sheets alone leaves the file full of originals. Removing them is simpler
and more complete than replacing them column by column:

```yaml
ignore_sheets:
  - sheet: Overview
    reason: a working view assembled from the data sheets
  - sheet: Export
    reason: an export to another system, not an input
```

A plain list of names works too; the reason is for whoever has to agree
that these sheets are dispensable. Removing every sheet is refused, and
a sheet that is not in the file only warns — one list often serves
several files of the same family.

### Numbers: scaling instead of replacing

Amounts are confidential too, but replacing them with keys destroys
everything that made the file useful. A group can scale them instead:

```yaml
groups:
  - name: amounts
    strategy: scale
    factor: 1.2837
    columns:
      - sheet: Positions
        col: H
        data_from_row: 2
```

One factor for a whole workbook, so the relations between amounts
survive: a commitment stays below its tranche total, and a sum of
cashflows still reconciles against the balance it belongs to.

Two rules follow from that:

- **Do not scale ratios.** Weights that add up to 1.0 would add up to
  the factor. Rates, percentages and quotas stay in a `key` group or are
  left out entirely. Which columns are amounts and which are ratios is a
  statement about the data, and that is why it lives in the config.
- **The type is preserved.** An integer stays an integer, so a column of
  whole amounts does not turn into a column of decimals. Text, dates and
  booleans in a scaled column are left untouched and counted in the log.

Scaling is one-way — rounding loses the remainder — so scale groups are
not written to the mapping file.

### Keys and numbers

A key is text, so writing one over a number changes the type of that
column. Key groups therefore skip numeric cells: a column of amounts
with a stray text cell in it stays a column of amounts. Where a numeric
value really has to go — a tax number, an account number — the group
says so:

```yaml
  - name: tax_numbers
    prefix: "TAXNO"
    include_numbers: true
    columns: [...]
```

## faker_replace.py — realistic fake data

Replaces values with realistic fake names, cities, companies etc. via [Faker](https://faker.readthedocs.io/). Same original value always gets the same fake replacement.

```bash
uv run src/faker_replace.py input.xlsx
uv run src/faker_replace.py input.xlsx --config examples/config_faker.yaml
uv run src/faker_replace.py input.xlsx --locale de_DE
```

Output: `input_faker.xlsx` next to the original.

Config:

```yaml
output_suffix: "_faker"
save_mapping: "faker_map.json"
locale: en_US                # any Faker locale: de_DE, fr_FR, ja_JP, …

groups:
  - name: last_names
    faker_type: last_name    # see supported types below
    columns:
      - sheet: Employees
        col: B
        data_from_row: 2
      - sheet: Employees
        col: D               # manager column — same fake value as employee
        data_from_row: 2

  - name: first_names
    faker_type: first_name
    columns:
      - sheet: Employees
        col: C
        data_from_row: 2
```

**Supported `faker_type` values:**

| Type | Example output |
|------|----------------|
| `last_name` | Smith |
| `first_name` | Emily |
| `full_name` | Emily Smith |
| `company` | Acme Corp |
| `city` | Berlin |
| `department` | Engineering |
| `email` | e.smith@example.com |
| `word` | Alpha |

## deanonymize.py — restore originals

Reverses an anonymized file back to the original values using the mapping file saved during anonymization. Works with both `anonymize.py` and `faker_replace.py` mappings.

```bash
uv run src/deanonymize.py sample_data_anonymized.xlsx --mapping anonymization_map.json
uv run src/deanonymize.py sample_data_faker.xlsx --mapping faker_map.json
```

Output: `<input>_restored.xlsx` next to the input file.

> **Note:** The mapping file must not have been deleted or excluded. By default both mapping files are `.gitignore`d and kept local only — see the [Security note](#security-note) below.

## generate_sample.py — create test data

Generates a fake Excel file matching the structure of a config. Uses `faker_type` when configured, otherwise infers it from the group name.

```bash
uv run src/generate_sample.py
uv run src/generate_sample.py --config examples/config_faker.yaml
uv run src/generate_sample.py --config examples/config_full.yaml --rows 50 --output my_sample.xlsx
```

Values are drawn from a small pool with repetition — so the same name appears as both employee and manager, demonstrating relationship preservation.

## Config examples

| File | Use case |
|------|----------|
| `examples/config_names.yaml` | Key replacement for person names |
| `examples/config_full.yaml` | Key replacement for names, departments, locations |
| `examples/config_faker.yaml` | Faker replacement with separate groups for last/first names |
| `examples/config_amounts.yaml` | Key replacement for names plus scaled amounts |

## Security note

Mapping files (`*_map.json`) contain the original values and are excluded from version control via `.gitignore`. All `.xlsx` files are excluded as well. Keep both local.
