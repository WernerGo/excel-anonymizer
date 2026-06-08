# excel-anonymizer

Anonymize Excel files by replacing cell values with unique, consistent keys.

**Key property:** identical values across all configured columns always get the same key — so relationships (e.g. employee → manager, or foreign keys across sheets) are preserved in the anonymized file.

## Use case

You have an Excel file with personal data (names, departments, locations) that you want to share for testing, debugging, or review — without exposing real data. Classic tools either target databases or don't preserve cross-column relationships.

## How it works

1. Define which sheets and columns to anonymize in a YAML config file
2. Run the script — it scans all configured columns and builds a value→key mapping
3. The anonymized file is written with a suffix (e.g. `data_anonymized.xlsx`)
4. Optionally saves a JSON mapping file so you can trace keys back to originals

Example: `"Smith"` appears as employee last name (col B) and as manager last name (col D) → both get the same key `NAME0001`. The hierarchy stays intact.

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

## Usage

```bash
# Generate sample data to try it out (no real Excel needed)
uv run src/generate_sample.py
uv run src/generate_sample.py --config examples/config_full.yaml --rows 50

# Anonymize a file
uv run src/anonymize.py input.xlsx
uv run src/anonymize.py input.xlsx --config examples/config_names.yaml
```

Output: `input_anonymized.xlsx` next to the original file.

## Configuration

```yaml
output_suffix: "_anonymized"
save_mapping: "anonymization_map.json"   # optional

groups:
  - name: person_names
    prefix: "NAME"
    columns:
      - sheet: Employees     # sheet name
        col: B               # column letter
        data_from_row: 2     # first data row (skip headers)
      - sheet: Employees
        col: C
        data_from_row: 2
```

- **Groups** are independent — values are only shared within a group, not across groups
- **Multiple groups** let you use different prefixes for names, departments, locations, etc.
- **`data_from_row`** skips header rows; defaults to `2`

See `examples/` for ready-to-use configs.

## Planned extensions

- `src/faker_replace.py` — replace keys with realistic fake names via [Faker](https://faker.readthedocs.io/)
- `src/deanonymize.py` — restore original values using a saved mapping file

## Security note

The mapping file (`anonymization_map.json`) contains the original values and is excluded from version control via `.gitignore`. Keep it local.
