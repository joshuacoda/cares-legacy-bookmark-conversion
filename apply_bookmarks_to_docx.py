import re
import unicodedata
from pathlib import Path
import csv

import pandas as pd  # NEW: for reading Excel

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
INPUT_DIR = Path("input/original")
OUTPUT_DIR = Path("output/bookmarks")
TALLY_FILE = OUTPUT_DIR.parent / "bookmarks_tally.csv"
UNIQUE_FILE = OUTPUT_DIR.parent / "unique_bookmarks.csv"
DATA_DICT_FILE = Path("CWS-CARES Forms Data Dictionary-CountDataElements V0.01.xlsx")  # NEW


# ----------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------
def ensure_dirs():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TALLY_FILE.parent.mkdir(parents=True, exist_ok=True)


def sanitize_filename(stem: str) -> str:
    cleaned = re.sub(r"\W+", "_", stem)
    cleaned = cleaned.strip("_")
    return cleaned or "document"


def iter_bookmark_starts(doc):
    """
    Yield (bookmark_element, name) for all bookmarkStart elements in the main body.
    """
    for bkm in doc.element.body.iter(qn("w:bookmarkStart")):
        name = bkm.get(qn("w:name"))
        if name:
            yield bkm, name


def get_paragraph_element(elm):
    """
    Walk up the XML tree until we hit a w:p (paragraph) element.
    """
    while elm is not None and elm.tag != qn("w:p"):
        elm = elm.getparent()
    return elm


def replace_paragraph_with_text(p_elm, text: str):
    """
    Replace all content of a paragraph (including fields, bookmarks, runs)
    with a single text run containing `text`.
    """
    for child in list(p_elm):
        p_elm.remove(child)

    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    p_elm.append(r)


def normalize_bookmark_name(name: str) -> str:
    """
    Normalize bookmark name for CSV:
    - Strip leading/trailing whitespace
    - Normalize Unicode
    - Drop non-ASCII characters
    """
    s = name.strip()
    s = unicodedata.normalize("NFKD", s)
    s = s.encode("ascii", "ignore").decode("ascii")
    return s


def load_valid_bookmarks_from_excel(path: Path) -> set[str]:
    """
    Load valid bookmark names from DataDictionary.xlsx.
    Expects column: 'Bookmark Name'
    Only the 'Bookmark Name' values are used as a filter here.
    """
    df = pd.read_excel(path)

    if "Bookmark Name" not in df.columns:
        raise ValueError("DataDictionary.xlsx must have a 'Bookmark Name' column")

    bookmarks = (
        df["Bookmark Name"]
        .dropna()
        .astype(str)
        .str.strip()
    )

    return set(bookmarks)



def process_document(
    input_path: Path,
    unique_set: set,
    valid_bookmarks: set[str],
) -> tuple[str, int]:
    """
    Process a single .docx file:
    - For every bookmark whose name is in `valid_bookmarks`
      AND does NOT start with 'Text' or 'Check' (case-insensitive),
      remove contents of its paragraph and insert the bookmark name as plain text.
    - Bookmarks not in `valid_bookmarks` are ignored.
    """
    doc = Document(input_path)
    bookmark_count = 0

    # Gather first to avoid modifying while iterating
    bookmarks_to_process = list(iter_bookmark_starts(doc))

    for bkm, name in bookmarks_to_process:
        lower_name = name.lower()

        # Skip "Text..." and "Check..." bookmarks
        if lower_name.startswith("text") or lower_name.startswith("check"):
            continue

        # Only handle bookmarks that exist in the data_dictionary
        # (compare against the raw name; if you want case-insensitive,
        # you could store a lowercased set instead).
        if name not in valid_bookmarks:
            continue

        # Add normalized name to unique set for CSV
        cleaned = normalize_bookmark_name(name)
        if cleaned:
            unique_set.add(cleaned)

        p_elm = get_paragraph_element(bkm)
        if p_elm is None:
            continue

        # Replace paragraph contents with the bookmark name itself
        replace_paragraph_with_text(p_elm, name)
        bookmark_count += 1

    stem = sanitize_filename(input_path.stem)
    output_filename = f"{stem}.docx"
    output_path = OUTPUT_DIR / output_filename
    doc.save(output_path)

    return output_filename, bookmark_count


def write_tally(tally_rows):
    with open(TALLY_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["document_name", "number_of_bookmarks"])
        writer.writerows(tally_rows)


def write_unique_bookmarks(unique_set: set):
    """
    Write CSV of unique normalized bookmark names.
    Columns: bookmark_name, token
    'token' is left empty for now.
    """
    with open(UNIQUE_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["bookmark_name", "json"])
        for name in sorted(unique_set):
            writer.writerow([name, ""])


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    ensure_dirs()

    # Load bookmark filter from Excel
    valid_bookmarks = load_valid_bookmarks_from_excel(DATA_DICT_FILE)

    tally_rows = []
    unique_bookmarks = set()

    for path in INPUT_DIR.glob("*.docx"):
        output_name, count = process_document(path, unique_bookmarks, valid_bookmarks)
        tally_rows.append([output_name, count])

    write_tally(tally_rows)
    write_unique_bookmarks(unique_bookmarks)


if __name__ == "__main__":
    main()
