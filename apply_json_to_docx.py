import csv
import shutil
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
from docx import Document

# ADDED
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
BOOKMARK_DOCX_DIR = Path("output/bookmarks")
TALLY_FILE = Path("output/bookmarks_tally.csv")
FINAL_TALLY_FILE = Path("output/final_tally.csv")

TOKEN_DOCX_DIR = Path("output/json")

DATA_DICT_FILE = Path("CWS-CARES Forms Data Dictionary-CountDataElements V0.01.xlsx")

# ADDED: table sizing defaults (tune these)
DEFAULT_COL_WIDTH_INCHES = 2.0           # set to match your template
EXACT_ROW_HEIGHT_INCHES = None           # e.g. 0.35 to lock height (may clip). None = allow wrap


# ----------------------------------------------------------------------
# DATA DICTIONARY (bookmark → token)
# ----------------------------------------------------------------------
def load_data_dictionary(path: Path) -> Dict[str, str]:
    df = pd.read_excel(path)

    if "Bookmark Name" not in df.columns:
        raise ValueError("data_dictionary.xlsx must contain 'Bookmark Name' column")

    if "Json Path" not in df.columns:
        raise ValueError("data_dictionary.xlsx must contain 'Json Path' column")

    mapping: Dict[str, str] = {}

    for _, row in df.iterrows():
        bookmark = row["Bookmark Name"]
        json_path = row["Json Path"]

        if pd.isna(bookmark) or pd.isna(json_path):
            continue

        bookmark_str = str(bookmark).strip()
        json_str = str(json_path).strip()

        if not bookmark_str or not json_str:
            continue

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
    candidates = ["filename", "file", "docx_file", "document", "doc_name", "document_name"]
    lower_map = {name.lower(): name for name in fieldnames}
    for cand in candidates:
        if cand in lower_map:
            return lower_map[cand]
    return fieldnames[0]


# ----------------------------------------------------------------------
# DOCX PARAGRAPH ITERATION (INCLUDES TABLES)
# ----------------------------------------------------------------------
def iter_paragraphs(doc: Document):
    for para in doc.paragraphs:
        yield para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para


# ----------------------------------------------------------------------
# ADDED: TABLE FIXED-LAYOUT ENFORCEMENT (PREVENT WIDTH EXPANSION)
# ----------------------------------------------------------------------
def iter_tables(doc: Document):
    # includes nested tables
    for tbl in doc.tables:
        yield tbl
        for row in tbl.rows:
            for cell in row.cells:
                for nested in cell.tables:
                    yield nested


def set_table_fixed_layout(table):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    tblPr = table._tbl.tblPr
    tblLayout = tblPr.find(qn("w:tblLayout"))
    if tblLayout is None:
        tblLayout = OxmlElement("w:tblLayout")
        tblPr.append(tblLayout)
    tblLayout.set(qn("w:type"), "fixed")


def set_cell_width(cell, width):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:type"), "dxa")
    tcW.set(qn("w:w"), str(width.twips))


def enforce_fixed_table_geometry(
    doc: Document,
    default_col_width_inches: float = DEFAULT_COL_WIDTH_INCHES,
    exact_row_height_inches: float | None = EXACT_ROW_HEIGHT_INCHES,
):
    col_w = Inches(default_col_width_inches)

    for table in iter_tables(doc):
        set_table_fixed_layout(table)

        for row in table.rows:
            if exact_row_height_inches is not None:
                row.height = Inches(exact_row_height_inches)
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

            for cell in row.cells:
                set_cell_width(cell, col_w)


# ----------------------------------------------------------------------
# APPLY DATA DICTIONARY TOKENS TO DOCX
# ----------------------------------------------------------------------
def apply_data_dict_to_doc(
    doc_path: Path, bookmark_to_token: Dict[str, str]
) -> Tuple[List[str], List[str]]:

    TOKEN_DOCX_DIR.mkdir(parents=True, exist_ok=True)

    doc = Document(doc_path)

    # ADDED: do this BEFORE text replacement so Word doesn't autosize on new content
    enforce_fixed_table_geometry(doc)

    bookmarks_replaced_set = set()
    json_used_set = set()

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

    out_path = TOKEN_DOCX_DIR / doc_path.name
    doc.save(out_path)

    return sorted(bookmarks_replaced_set), sorted(json_used_set)


def copy_doc_without_changes(doc_path: Path) -> None:
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

    bookmark_to_token = load_data_dictionary(DATA_DICT_FILE)

    FINAL_TALLY_FILE.parent.mkdir(parents=True, exist_ok=True)

    with TALLY_FILE.open("r", encoding="utf-8", newline="") as f_in:
        reader = csv.DictReader(f_in)
        fieldnames = reader.fieldnames or []
        if not fieldnames:
            raise ValueError("Tally file has no columns")

        filename_col = detect_filename_column(fieldnames)
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

                row["bookmarks_replaced"] = ";".join(bookmarks_replaced)
                row["json"] = ";".join(json_used)
                writer.writerow(row)

    print(f"Wrote updated DOCX files to: {TOKEN_DOCX_DIR}")
    print(f"Wrote final tally CSV to: {FINAL_TALLY_FILE}")


if __name__ == "__main__":
    main()
