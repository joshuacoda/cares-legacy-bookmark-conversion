import json
from pathlib import Path
from typing import Any, List, Tuple

# Optional: if you use YAML schemas
try:
    import yaml  # pip install pyyaml
except ImportError:
    yaml = None

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
SCHEMA_DIR = Path("input/schema")
OUTPUT_DIR = Path("output")
OUTPUT_HTML = OUTPUT_DIR / "schema_reference.html"


# ----------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------
def ensure_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def is_json(path: Path) -> bool:
    return path.suffix.lower() == ".json"


def is_yaml(path: Path) -> bool:
    return path.suffix.lower() in {".yaml", ".yml"}


def load_schema(path: Path) -> Any:
    if is_json(path):
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    if is_yaml(path) and yaml is not None:
        with path.open("r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    return None


def _recurse_paths(obj: Any, prefix: str = "") -> List[str]:
    """
    Given a JSON/YAML structure, return all dot-separated paths for leaf values.
    """
    paths: List[str] = []

    if isinstance(obj, dict):
        for key, value in obj.items():
            new_prefix = f"{prefix}.{key}" if prefix else str(key)
            paths.extend(_recurse_paths(value, new_prefix))
    elif isinstance(obj, list):
        # Treat lists as repeated structures; do not index items
        for item in obj:
            paths.extend(_recurse_paths(item, prefix))
    else:
        # Leaf node
        if prefix:
            paths.append(prefix)

    return paths


def extract_paths_from_schema(schema_data: Any) -> List[str]:
    if schema_data is None:
        return []
    return sorted(set(_recurse_paths(schema_data)))


def discover_schemas(schema_dir: Path) -> List[Path]:
    if not schema_dir.exists():
        return []
    return sorted(
        p
        for p in schema_dir.rglob("*")
        if p.is_file() and (is_json(p) or is_yaml(p))
    )


def build_rows() -> List[Tuple[str, str, str]]:
    """
    Returns a list of (schema_name, path, token) rows.

    schema_name: file stem (e.g. "CaseManagement.Schema")
    path: full dot path (e.g. "CaresCorrespondence.formData.CaseManagement.HEP.ChildDetails.ChildName")
    token: token string (e.g. "{{CaresCorrespondence.formData.CaseManagement.HEP.ChildDetails.ChildName}}")
    """
    rows: List[Tuple[str, str, str]] = []
    for schema_file in discover_schemas(SCHEMA_DIR):
        schema_name = schema_file.stem
        data = load_schema(schema_file)
        paths = extract_paths_from_schema(data)
        for path in paths:
            # Build token with double curly braces
            token = f"{{{{{path}}}}}"
            rows.append((schema_name, path, token))
    return rows


# ----------------------------------------------------------------------
# HTML GENERATION (WITH SEARCH + COPY)
# ----------------------------------------------------------------------
def generate_html(rows: List[Tuple[str, str, str]]) -> str:
    def html_escape(text: str) -> str:
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    rows_html = []
    for schema_name, path, token in rows:
        esc_schema = html_escape(schema_name)
        esc_path = html_escape(path)
        esc_token = html_escape(token)
        rows_html.append(
            f"""
            <tr>
                <td>{esc_schema}</td>
                <td><code>{esc_path}</code></td>
                <td>
                    <code class="token-text">{esc_token}</code>
                    <button class="copy-btn" type="button" data-token="{esc_token}">
                        Copy
                    </button>
                </td>
            </tr>
            """
        )

    if not rows_html:
        table_body = """
            <tr><td colspan="3">No schemas or fields found.</td></tr>
        """
    else:
        table_body = "\n".join(rows_html)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Schema Reference</title>
    <style>
        body {{
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            margin: 16px;
            background-color: #f7f7f7;
        }}
        h1 {{
            margin-bottom: 0.25rem;
        }}
        .subtitle {{
            color: #555;
            margin-bottom: 1rem;
        }}
        .search-container {{
            margin-bottom: 1rem;
        }}
        .search-input {{
            padding: 6px 10px;
            font-size: 14px;
            width: 100%;
            max-width: 400px;
            box-sizing: border-box;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            background: #ffffff;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 8px 10px;
            vertical-align: middle;
            font-size: 14px;
        }}
        th {{
            background-color: #f0f0f0;
            text-align: left;
        }}
        tr:nth-child(even) {{
            background-color: #fafafa;
        }}
        code {{
            font-family: "SF Mono", Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
            font-size: 13px;
        }}
        .copy-btn {{
            margin-left: 8px;
            padding: 3px 8px;
            font-size: 12px;
            cursor: pointer;
        }}
        .copy-btn.copied {{
            outline: none;
        }}
    </style>
</head>
<body>
    <h1>Schema → Token Reference</h1>
    <p class="subtitle">
        Search below to filter by schema name, JSON path, or token. Use the "Copy" button to copy a token.
    </p>

    <div class="search-container">
        <input
            type="text"
            id="searchBox"
            class="search-input"
            placeholder="Type to search (e.g. ChildName, PastHealthIssues, HEP)..."
        />
    </div>

    <table id="schemaTable">
        <thead>
            <tr>
                <th>Schema File</th>
                <th>JSON Path</th>
                <th>Token (for DOCX)</th>
            </tr>
        </thead>
        <tbody>
            {table_body}
        </tbody>
    </table>

    <script>
        (function() {{
            const searchBox = document.getElementById('searchBox');
            const table = document.getElementById('schemaTable');
            const tbody = table.getElementsByTagName('tbody')[0];
            const rows = Array.from(tbody.getElementsByTagName('tr'));

            function normalize(str) {{
                return (str || '').toLowerCase();
            }}

            function filterRows() {{
                const query = normalize(searchBox.value);
                rows.forEach(row => {{
                    const text = normalize(row.textContent);
                    row.style.display = text.indexOf(query) !== -1 ? '' : 'none';
                }});
            }}

            searchBox.addEventListener('input', filterRows);

            // Copy-to-clipboard logic
            const copyButtons = document.querySelectorAll('.copy-btn');

            async function copyText(text) {{
                if (navigator.clipboard && navigator.clipboard.writeText) {{
                    return navigator.clipboard.writeText(text);
                }} else {{
                    const textarea = document.createElement('textarea');
                    textarea.value = text;
                    textarea.style.position = 'fixed';
                    textarea.style.left = '-9999px';
                    document.body.appendChild(textarea);
                    textarea.focus();
                    textarea.select();
                    try {{
                        document.execCommand('copy');
                    }} finally {{
                        document.body.removeChild(textarea);
                    }}
                }}
            }}

            copyButtons.forEach(btn => {{
                btn.addEventListener('click', async () => {{
                    const token = btn.getAttribute('data-token') || '';
                    // Decode HTML entities for braces if any (simple replacement for this use-case)
                    const decoded = token
                        .replace(/&amp;/g, '&')
                        .replace(/&lt;/g, '<')
                        .replace(/&gt;/g, '>')
                        .replace(/&quot;/g, '"');
                    try {{
                        await copyText(decoded);
                        const original = btn.textContent;
                        btn.textContent = 'Copied';
                        btn.classList.add('copied');
                        setTimeout(() => {{
                            btn.textContent = original;
                            btn.classList.remove('copied');
                        }}, 1200);
                    }} catch (e) {{
                        console.error('Copy failed', e);
                    }}
                }});
            }});
        }})();
    </script>
</body>
</html>
"""
    return html


def main() -> None:
    ensure_dirs()
    rows = build_rows()
    html = generate_html(rows)
    with OUTPUT_HTML.open("w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote HTML reference to: {OUTPUT_HTML}")


if __name__ == "__main__":
    main()
