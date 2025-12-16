# CARES Legacy Bookmark Conversion

This project automates the conversion of legacy Microsoft Word (DOCX) bookmarks used in CARES correspondence templates into standardized token placeholders derived from the State CARES Data Dictionary **JSON Path** definitions.

It supports a full, auditable workflow for modernizing legacy correspondence templates.

---

## Purpose

The State released an updated CARES Data Dictionary that defines data elements using **JSON Path** expressions rather than legacy Word bookmark tokens. This repository bridges that gap by:

- Identifying and validating legacy bookmarks in DOCX templates
- Ensuring bookmarks align with the official CARES Data Dictionary
- Replacing legacy bookmark placeholders with `{{JsonPath}}` tokens
- Producing traceable audit outputs for compliance and QA

---

## Repository Structure

```
.
├── input/
│   └── original/               # Original DOCX templates
├── output/
│   ├── bookmarks/              # Templates with visible bookmark names
│   ├── tokens/                 # Templates with {{JsonPath}} tokens applied
│   ├── bookmarks_tally.csv     # Per-document bookmark counts
│   ├── final_tally.csv         # Replacement audit by document
│   └── unique_bookmarks.csv    # All unique bookmarks encountered
├── DataDictionary.xlsx         # State CARES data dictionary
├── extract_bookmarks_to_docx.py        # Bookmark extraction and normalization
├── apply_json_to_docx.py             # Bookmark → {{JsonPath}} replacement
└── README.md
```

---

## Data Dictionary Requirements

The Excel data dictionary **must** contain the following columns:

- **Bookmark Name** – legacy bookmark identifier used in DOCX templates
- **Json Path** – CARES JSON Path value (without `{{ }}`)

The scripts automatically wrap JSON Path values in `{{ }}` during token generation.

Example:
```
Bookmark Name: ACDateAdopted
Json Path: CaresCorrespondence.formData.CountyLegacyDataElements.AdoptChild.DateAdopted
```

Becomes:
```
{{CaresCorrespondence.formData.CountyLegacyDataElements.AdoptChild.DateAdopted}}
```

---

## Workflow

### 1. Extract and Normalize Bookmarks

Processes original DOCX files and replaces valid bookmarks with their visible bookmark names.

```bash
python apply_bookmarks_to_docx.py
```

Outputs:
- `output/bookmarks/*.docx`
- `bookmarks_tally.csv`
- `unique_bookmarks.csv`

Only bookmarks listed in the Data Dictionary are processed.

---

### 2. Apply JSON Path Tokens

Replaces visible bookmark names with `{{JsonPath}}` tokens using the Data Dictionary.

```bash
python apply_json_to_docx.py
```

Outputs:
- `output/tokens/*.docx`
- `final_tally.csv`

---

## Key Features

- Uses **Json Path** as the authoritative token source
- Strict enforcement of `Bookmark Name` from the official dictionary
- Automatic `{{ }}` token wrapping
- Table-aware DOCX parsing
- Full audit trail of replacements
- Safe handling of missing or invalid bookmarks

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

MIT License. See `LICENSE` file for details.

---

## Maintainer

Joshua Coda
