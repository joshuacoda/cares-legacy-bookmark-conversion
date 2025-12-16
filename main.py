import subprocess
import sys
from pathlib import Path


def run_script(script_path: str):
    """
    Utility to run a Python script as a subprocess.
    Stops execution if the script fails.
    """
    print(f"\n--- Running {script_path} ---")
    result = subprocess.run([sys.executable, script_path], text=True)

    if result.returncode != 0:
        raise SystemExit(f"Error: {script_path} failed with exit code {result.returncode}")


def main():
    # Ensure scripts exist
    scripts = [
        "apply_bookmarks_to_docx.py",
        "apply_json_to_docx.py",
    ]

    for script in scripts:
        if not Path(script).exists():
            raise FileNotFoundError(f"Cannot find script: {script}")

    # Run scripts in strict order
    run_script("apply_bookmarks_to_docx.py")
    run_script("apply_json_to_docx.py")

    print("\nAll processing completed successfully.")


if __name__ == "__main__":
    main()
