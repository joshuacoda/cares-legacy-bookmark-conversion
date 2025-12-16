import csv
import re
import shutil
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
from docx import Document

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
BOOKMARK_DOCX_DIR = Path("output/bookmarks")
TALLY_FILE = Path("output/bookmarks_tally.csv")
FINAL_TALLY_FILE = Path("output/final_tally.csv")

# UPDATED DOCX go here
TOKEN_DOCX_DIR = Path("output/json")

DATA_DICT_FILE = Path("CWS-CARES Forms Data Dictionary-CountDataElements V0.01.xlsx")


# ----------------------------------------------------------------------
# DATA DICTIONARY (bookmark → token)
# ----------------------------------------------------------------------
def load_data_dictionary(path: Path) -> Dict[str, str]:
    """
    Load mapping from Bookmark Name → token from data_dictionary.xlsx.

    Required columns:
      - "Bookmark Name"
      - "json" (will be wrapped in {{ }} if missing)
    """
    df = pd.read_excel(path)

    if "Bookmark Name" not in df.columns:
        raise ValueError("data_dictionary.xlsx must contain 'Bookmark Name' column")

    if "Json Path" not in df.columns:
        raise ValueError("data_dictionary.xlsx must contain 'Json Path' column")

    mapping: Dict[str, str] = {}

    for _, row in df.iterrows():
        bookmark = row["Bookmark Name"]
        json = row["Json Path"]

        if pd.isna(bookmark) or pd.isna(json):
            continue

        bookmark_str = str(bookmark).strip()
        json_str = str(json).strip()

        if not bookmark_str or not json_str:
            continue

        # Ensure token format {{ ... }}
        if json_str.startswith("{{") and json_str.endswith("}}"):
            token = json_str
        else:
            token = f"{{{{{json_str}}}}}"

        mapping[bookmark_str] = token

    return mapping



# ----------------------------------------------------------------------
# TALLY CSV HANDLING
# ----------------------------------------------------------------------
def detect_filename_column(fieldnames: List[str]) -> str:
    """
    Guess which column in bookmarks_tally.csv contains the DOCX filename.
    Adjust this if you know the exact column name.
    """
    candidates = ["filename", "file", "docx_file", "document", "doc_name", "document_name"]
    lower_map = {name.lower(): name for name in fieldnames}
    for cand in candidates:
        if cand in lower_map:
            return lower_map[cand]
    # Fallback: first column
    return fieldnames[0]


# ----------------------------------------------------------------------
# DOCX PARAGRAPH ITERATION (INCLUDES TABLES)
# ----------------------------------------------------------------------
def iter_paragraphs(doc: Document):
    """
    Yield all paragraphs in the document, including those inside tables (one level).
    """
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para


# ----------------------------------------------------------------------
# APPLY DATA DICTIONARY json TO DOCX
# ----------------------------------------------------------------------
def apply_data_dict_to_doc(
    doc_path: Path, bookmark_to_token: Dict[str, str]
) -> Tuple[List[str], List[str]]:
    """
    Open DOCX from output/bookmarks, replace any visible bookmark names
    with their corresponding json from data_dictionary, and save to output/json.

    Returns (bookmarks_replaced, json_used).
    """
    TOKEN_DOCX_DIR.mkdir(parents=True, exist_ok=True)

    doc = Document(doc_path)

    bookmarks_replaced_set = set()
    json_used_set = set()

    # For each run, replace any bookmark names that appear in the text.
    # Assumes bookmark names appear as plain text (e.g. 'ChildBirthDate').
    for para in iter_paragraphs(doc):
        for run in para.runs:
            text = run.text or ""
            new_text = text

            for bookmark, token in bookmark_to_token.items():
                if bookmark in new_text:
                    new_text = new_text.replace(bookmark, token)
                    bookmarks_replaced_set.add(bookmark)
                    json_used_set.add(token)

            if new_text != text:
                run.text = new_text

    # Save updated DOCX into output/json
    out_path = TOKEN_DOCX_DIR / doc_path.name
    doc.save(out_path)

    return sorted(bookmarks_replaced_set), sorted(json_used_set)


def copy_doc_without_changes(doc_path: Path) -> None:
    """
    Ensure the doc exists in output/json even if we don't replace anything.
    """
    TOKEN_DOCX_DIR.mkdir(parents=True, exist_ok=True)
    out_path = TOKEN_DOCX_DIR / doc_path.name
    shutil.copy2(doc_path, out_path)


# ----------------------------------------------------------------------
# MAIN ORCHESTRATION
# ----------------------------------------------------------------------
def main() -> None:
    if not TALLY_FILE.exists():
        raise FileNotFoundError(f"Tally file not found: {TALLY_FILE}")

    if not DATA_DICT_FILE.exists():
        raise FileNotFoundError(f"Data dictionary not found: {DATA_DICT_FILE}")

    # Load bookmark → token mapping from Excel
    bookmark_to_token = load_data_dictionary(DATA_DICT_FILE)

    FINAL_TALLY_FILE.parent.mkdir(parents=True, exist_ok=True)

    with TALLY_FILE.open("r", encoding="utf-8", newline="") as f_in:
        reader = csv.DictReader(f_in)
        fieldnames = reader.fieldnames or []
        if not fieldnames:
            raise ValueError("Tally file has no columns")

        filename_col = detect_filename_column(fieldnames)

        # New columns in final_tally:
        # - bookmarks_replaced: bookmark names actually replaced in this doc
        # - json: json actually inserted into this doc
        out_fieldnames = fieldnames + ["bookmarks_replaced", "json"]

        with FINAL_TALLY_FILE.open("w", encoding="utf-8", newline="") as f_out:
            writer = csv.DictWriter(f_out, fieldnames=out_fieldnames)
            writer.writeheader()

            for row in reader:
                docx_name = row[filename_col]
                docx_path = BOOKMARK_DOCX_DIR / docx_name

                bookmarks_replaced: List[str] = []
                json_used: List[str] = []

                if docx_path.exists():
                    bookmarks_replaced, json_used = apply_data_dict_to_doc(
                        docx_path, bookmark_to_token
                    )
                else:
                    # If the file listed in the tally doesn't exist, leave blank and continue
                    bookmarks_replaced = []
                    json_used = []

                row["bookmarks_replaced"] = ";".join(bookmarks_replaced)
                row["json"] = ";".join(json_used)
                writer.writerow(row)

    print(f"Wrote updated DOCX files to: {TOKEN_DOCX_DIR}")
    print(f"Wrote final tally CSV to: {FINAL_TALLY_FILE}")


if __name__ == "__main__":
    main()
