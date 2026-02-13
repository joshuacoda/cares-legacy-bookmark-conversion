import re
import unicodedata
from pathlib import Path
import csv
import copy
import traceback

import pandas as pd

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ----------------------------------------------------------------------
# CONFIG (paths are relative to this script file)
# ----------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input" / "original"
OUTPUT_DIR = BASE_DIR / "output" / "bookmarks"
TALLY_FILE = BASE_DIR / "output" / "bookmarks_tally.csv"
UNIQUE_FILE = BASE_DIR / "output" / "unique_bookmarks.csv"
DATA_DICT_FILE = BASE_DIR / "CWS-CARES Forms Data Dictionary-CountyDataElements V0.03.xlsx"


# ----------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------
def ensure_dirs():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TALLY_FILE.parent.mkdir(parents=True, exist_ok=True)


def sanitize_filename(stem: str) -> str:
    cleaned = re.sub(r"\W+", "_", stem).strip("_")
    return cleaned or "document"


def iter_bookmark_starts(doc):
    for bkm in doc.element.body.iter(qn("w:bookmarkStart")):
        name = bkm.get(qn("w:name"))
        if name:
            yield bkm, name


def normalize_bookmark_name(name: str) -> str:
    s = name.strip()
    s = unicodedata.normalize("NFKD", s)
    s = s.encode("ascii", "ignore").decode("ascii")
    return s


def load_valid_bookmarks_from_excel(path: Path) -> set[str]:
    if not path.exists():
        raise FileNotFoundError(f"Excel data dictionary not found: {path}")
    df = pd.read_excel(path)
    if "Bookmark Name" not in df.columns:
        raise ValueError("DataDictionary.xlsx must have a 'Bookmark Name' column")
    return set(df["Bookmark Name"].dropna().astype(str).str.strip())


def get_paragraph_element(elm):
    while elm is not None and elm.tag != qn("w:p"):
        elm = elm.getparent()
    return elm


def find_bookmark_end(doc, start):
    start_id = start.get(qn("w:id"))
    if start_id is None:
        return None
    for end in doc.element.body.iter(qn("w:bookmarkEnd")):
        if end.get(qn("w:id")) == start_id:
            return end
    return None


def _make_run_text(text: str):
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def set_paragraph_spacing(p_elm, before: str | None = None, after: str | None = None) -> None:
    """
    Forces paragraph spacing via w:pPr/w:spacing.
    This helps when splitting one paragraph into two causes a large visible gap.
    """
    pPr = p_elm.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elm.insert(0, pPr)

    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)

    if before is not None:
        spacing.set(qn("w:before"), str(before))
    if after is not None:
        spacing.set(qn("w:after"), str(after))


# ----------------------------------------------------------------------
# STRATEGY:
# 1) Split paragraph before bookmarkStart to preserve any prefix label text.
# 2) Replace everything inside bookmark range with bookmark name.
# ----------------------------------------------------------------------
def split_paragraph_before_bookmark(bkm_start) -> bool:
    p = get_paragraph_element(bkm_start)
    if p is None:
        return False

    p_parent = p.getparent()
    if p_parent is None:
        return False

    children = list(p)
    try:
        i_start = children.index(bkm_start)
    except ValueError:
        return False

    has_ppr = (len(children) > 0 and children[0].tag == qn("w:pPr"))
    first_movable_idx = 1 if has_ppr else 0

    # Nothing before the bookmark other than pPr
    if i_start <= first_movable_idx:
        return False

    # Create new paragraph and copy pPr if present
    new_p = OxmlElement("w:p")
    if has_ppr:
        new_p.append(copy.deepcopy(children[0]))

    # Move nodes before bookmarkStart into new paragraph
    to_move = children[first_movable_idx:i_start]
    for node in to_move:
        p.remove(node)
        new_p.append(node)

    # Insert new paragraph immediately before original paragraph
    parent_children = list(p_parent)
    try:
        p_index = parent_children.index(p)
    except ValueError:
        return False

    p_parent.insert(p_index, new_p)

    # ------------------------------------------------------------------
    # IMPORTANT SPACING FIX:
    # Word uses max(after of prev, before of next) between paragraphs.
    # When we split, we create a new boundary; if either side has spacing,
    # it can look like extra space.
    #
    # Set:
    #   - after=0 on new paragraph
    #   - before=0 on original paragraph
    # only for this split boundary.
    # ------------------------------------------------------------------
    set_paragraph_spacing(new_p, after="0")
    set_paragraph_spacing(p, before="0")

    return True


def replace_bookmark_range_with_text(doc, bkm_start, text: str) -> bool:
    bkm_end = find_bookmark_end(doc, bkm_start)
    if bkm_end is None:
        return False

    start_parent = bkm_start.getparent()
    end_parent = bkm_end.getparent()
    if start_parent is None or start_parent != end_parent:
        return False  # skip complex spanning cases

    parent = start_parent
    children = list(parent)
    try:
        i_start = children.index(bkm_start)
        i_end = children.index(bkm_end)
    except ValueError:
        return False

    for child in children[i_start + 1 : i_end]:
        parent.remove(child)

    parent.insert(i_start + 1, _make_run_text(text))
    return True


def process_document(input_path: Path, unique_set: set, valid_bookmarks: set[str]) -> tuple[str, int]:
    doc = Document(input_path)
    bookmark_count = 0

    bookmarks_to_process = list(iter_bookmark_starts(doc))

    for bkm, name in bookmarks_to_process:
        lower_name = name.lower()

        if lower_name.startswith("text") or lower_name.startswith("check"):
            continue

        if name not in valid_bookmarks:
            continue

        cleaned = normalize_bookmark_name(name)
        if cleaned:
            unique_set.add(cleaned)

        split_paragraph_before_bookmark(bkm)

        if replace_bookmark_range_with_text(doc, bkm, name):
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
    with open(UNIQUE_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["bookmark_name", "json"])
        for name in sorted(unique_set):
            writer.writerow([name, ""])


def find_input_docx_files(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        return []
    # recursive, case-insensitive .docx
    files = [p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() == ".docx"]
    return files


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    ensure_dirs()

    print(f"BASE_DIR:   {BASE_DIR}")
    print(f"INPUT_DIR:  {INPUT_DIR}")
    print(f"OUTPUT_DIR: {OUTPUT_DIR}")
    print(f"EXCEL:      {DATA_DICT_FILE}")

    valid_bookmarks = load_valid_bookmarks_from_excel(DATA_DICT_FILE)

    paths = find_input_docx_files(INPUT_DIR)
    if not paths:
        print("No .docx files found. Check INPUT_DIR and your working folder.")
        return

    print(f"Found {len(paths)} .docx file(s).")

    tally_rows = []
    unique_bookmarks = set()

    for path in paths:
        try:
            output_name, count = process_document(path, unique_bookmarks, valid_bookmarks)
            tally_rows.append([output_name, count])
            print(f"Processed: {path.name} -> {output_name} ({count} bookmark(s))")
        except Exception as e:
            print(f"ERROR processing {path}: {e}")
            traceback.print_exc()

    if tally_rows:
        write_tally(tally_rows)
        write_unique_bookmarks(unique_bookmarks)
        print(f"Wrote outputs to: {OUTPUT_DIR}")
        print(f"Tally CSV:  {TALLY_FILE}")
        print(f"Unique CSV: {UNIQUE_FILE}")


if __name__ == "__main__":
    main()
