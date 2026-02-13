# CARES Legacy Bookmark Conversion

This project automates the conversion of legacy Microsoft Word (DOCX) bookmarks used in CARES correspondence templates into standardized JSON placeholders derived from the State CARES Data Dictionary **JSON Path** definitions.

The workflow is fully automated through `main.py` and produces auditable outputs suitable for compliance and QA review.

---

## Purpose

The State released an updated CARES Data Dictionary that defines data elements using **JSON Path** expressions rather than legacy Word bookmark tokens. This repository bridges that gap by:

- Identifying and validating legacy bookmarks in DOCX templates
- Ensuring bookmarks align with the official CARES Data Dictionary
- Replacing legacy bookmark placeholders with `{{JsonPath}}` tokens
- Producing traceable audit outputs for compliance and QA

---

## Repository Structure

```text
.
├── input/
│   └── original/                   # Original DOCX templates
├── output/
│   ├── bookmarks/                  # Templates with visible bookmark names
│   ├── converted/                 # Templates with {{JsonPath}} tokens applied
│   ├── bookmarks_tally.csv         # Per-document bookmark counts
│   ├── final_tally.csv             # Replacement audit by document
│   └── unique_bookmarks.csv        # All unique bookmarks encountered
├── DataDictionary.xlsx             # State CARES data dictionary
├── main.py                         # Runs the entire pipeline
└── README.md
```

---

## Workflow Overview

The entire process is run using `main.py`. Individual scripts should not be run directly.

### 1. Bookmark Extraction

- Original DOCX templates are scanned for valid bookmarks
- Only bookmarks listed in the Data Dictionary are processed
- Bookmarks are replaced with visible bookmark names
- Outputs are written to `output/bookmarks`
- Audit files `bookmarks_tally.csv` and `unique_bookmarks.csv` are generated

### 2. JSON Path Application

- Visible bookmark names are replaced with `{{JsonPath}}` tokens
- JSON paths are sourced from the Data Dictionary
- Token wrapping with `{{ }}` is applied automatically
- Optimized single-pass regex replacement is used for performance
- Outputs are written to `output/converted`
- `final_tally.csv` is generated for audit purposes

---

## Output Filename Prefix

During execution, the user is prompted to enter an optional output filename prefix.

```text
Enter output filename prefix (leave blank for none):
```

Examples:

- `Riverside` → `Riverside_original.docx`
- `County_CA` → `County_CA_original.docx`
- *(blank)* → `original.docx`

This allows reuse of the same templates for multiple counties or deployments without code changes.

---

## Data Dictionary Requirements

The Excel data dictionary **must** contain the following columns:

- **Bookmark Name** – legacy bookmark identifier used in DOCX templates
- **Json Path** – CARES JSON Path value (without `{{ }}`)

The system automatically wraps JSON Path values in `{{ }}` during replacement.

Example:

```text
Bookmark Name: ACDateAdopted
Json Path: CaresCorrespondence.formData.CountyLegacyDataElements.AdoptChild.DateAdopted
```

Becomes:

```text
{{CaresCorrespondence.formData.CountyLegacyDataElements.AdoptChild.DateAdopted}}
```

---

## Key Features

- Single-command execution via `main.py`
- Strict enforcement of official Bookmark Name definitions
- Optimized regex-based replacement for performance
- Table-aware DOCX parsing
- Safe handling of nested tables
- Full CSV-based audit trail
- No hard-coded county or environment values

---

## Requirements

- Python 3.10+
- pandas
- python-docx
- openpyxl

Install dependencies:

```bash
pip install pandas python-docx openpyxl
```

---

## License

MIT License

---

## Maintainer

Joshua Coda

