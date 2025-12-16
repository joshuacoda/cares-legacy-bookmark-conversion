# DOCX Bookmark & Token Processing Pipeline
 A two-stage Python pipeline for transforming DOCX templates into token-ready documents for CARES.
 The system extracts Word bookmarks from DOCX files, rewrites them as readable identifiers, and then
 maps those identifiers to structured tokens defined in an Excel-based data dictionary.
 The result is a clean, fully tokenized set of DOCX templates accompanied by detailed CSV tallies for
 auditing and downstream automation.
 ## Features
 ### 1. Bookmark Processing (apply_bookmarks_to_docx.py)- Reads DOCX files from input/original/- Extracts Word bookmarks- Replaces bookmark content with the bookmark name text- Skips specific bookmark types- Outputs processed files and bookmark tally CSV
 ### 2. Token Injection (apply_tokens_to_docx.py)- Reads processed files from output/bookmarks/- Loads data_dictionary.xlsx (columns: bookmark, token)- Replaces bookmark names with tokens- Outputs tokenized DOCX and final tally CSV
 ### 3. Pipeline Orchestrator (main.py)
 Runs both steps automatically.
 ## Installation
 1. Create a virtual environment:
 python -m venv venv
 source venv/bin/activate # macOS/Linux
 venv\Scripts\activate # Windows
 2. Install dependencies:
 pip install -r requirements.txt
 3. Put original CWS CMS bookmark forms in input/original folder
 4. run main.py
 ## Requirements
 python-docx
 pandas
 openpyxl
 pyyaml
 ## Usage
 Run entire pipeline:
 python main.py
 Run individual steps:
 python apply_bookmarks_to_docx.py
 python apply_tokens_to_docx.py