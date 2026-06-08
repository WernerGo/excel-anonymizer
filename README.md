# excel-anonymizer

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
# Install uv (once, no admin rights required)
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

## Security note

Mapping files (`*_map.json`) contain the original values and are excluded from version control via `.gitignore`. All `.xlsx` files are excluded as well. Keep both local.
