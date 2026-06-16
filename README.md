# fileMerger

Recursively scan a directory, extract text from supported file types, and combine everything into a single Markdown file — with optional deduplication.

Useful for feeding a codebase or document folder into an LLM context window, creating a project-wide reference document, or archiving mixed-format content.

---

## Supported formats

| Format | Extension |
|--------|-----------|
| Plain text | `.txt` |
| Markdown | `.md` |
| JSON | `.json` |
| CSV | `.csv` |
| Excel workbook | `.xlsx` |
| PDF | `.pdf` |
| Word document | `.docx` |
| PowerPoint | `.pptx` |
| Jupyter notebook | `.ipynb` |

---

## Installation

```bash
git clone https://github.com/Rajallah/fileMerger.git
cd fileMerger
pip install -r requirements.txt
```

> `tqdm` is optional — if installed, a progress bar is shown automatically.

---

## Usage

```bash
# Merge the current directory into merged.md
python fileMerger.py

# Specify a directory and output file
python fileMerger.py --directory ./docs --output combined.md

# Disable deduplication
python fileMerger.py --no-dedup

# Exclude specific folders
python fileMerger.py --exclude tests fixtures data

# Increase CSV row limit
python fileMerger.py --csv-limit 200

# Verbose mode (prints each file as it's processed)
python fileMerger.py --verbose
```

---

## Options

| Flag | Default | Description |
|------|---------|-------------|
| `--directory`, `-d` | `.` | Directory to scan |
| `--output`, `-o` | `merged.md` | Output file name |
| `--csv-limit N` | `50` | Max rows from CSV/XLSX files. `0` = unlimited |
| `--json-limit N` | `50000` | Max characters from JSON files. `0` = unlimited |
| `--exclude DIR ...` | *(none)* | Extra directory names to skip |
| `--no-dedup` | *(off)* | Disable duplicate content removal |
| `--verbose`, `-v` | *(off)* | Print each file as it is processed |
| `--quiet`, `-q` | *(off)* | Suppress all output except errors |

---

## Excluded directories (defaults)

The following are automatically skipped:

`.git`, `.hg`, `.svn`, `__pycache__`, `.mypy_cache`, `.pytest_cache`, `.ruff_cache`, `node_modules`, `.venv`, `venv`, `env`, `.env`, `dist`, `build`, `target`, `.idea`, `.vscode`

Add your own with `--exclude`.

---

## Deduplication

By default, content blocks that are identical (after lowercasing and whitespace normalisation) are included only once. This is useful when the same text appears across multiple files. Disable it with `--no-dedup`.

---

## Output structure

```
# Merged Document

## Table of Contents
- path/to/file1.md
- path/to/file2.csv
...

---

## path/to/file1.md

(content)

---

## path/to/file2.csv

(content)

...

---

## Statistics

- Files discovered: 42
- Files processed: 42
- Files included: 38
- Duplicate blocks removed: 4
```

---

## License

MIT
