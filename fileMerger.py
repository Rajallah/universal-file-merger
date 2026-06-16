import os
import sys

# Define target output filename
OUTPUT_FILENAME = "Combined_File.md"
SCRIPT_SELF_NAME = "fileMerger.py"

# Get absolute path of the output file to guarantee it is safely excluded from scanning
OUTPUT_FILEPATH_ABS = os.path.abspath(OUTPUT_FILENAME)

# Global memory set to track and eliminate exact duplicate text blocks across files
seen_content = set()

# Mapping text/code extensions to their respective Markdown block code-syntax highlights
CODE_EXTENSIONS = {
    '.py': 'python',
    '.ipynb': 'json', # Read as raw structured text if nbformat fails
    '.md': 'markdown',
    '.txt': 'text',
    '.json': 'json',
    '.js': 'javascript',
    '.html': 'html',
    '.css': 'css',
    '.xml': 'xml',
    '.yaml': 'yaml',
    '.yml': 'yaml',
    '.sh': 'bash',
    '.bat': 'batch',
    '.ini': 'ini',
    '.cfg': 'ini',
    '.log': 'text',
    '.csv': 'text', # Fallback highlight if pandas/tabulate are absent
    '.tsv': 'text'
}

RICH_EXTENSIONS = ('.docx', '.pdf', '.xlsx', '.xls', '.pptx', '.csv', '.tsv')
ALL_SUPPORTED_EXTENSIONS = tuple(list(CODE_EXTENSIONS.keys()) + list(RICH_EXTENSIONS))

def read_text_with_fallback_encodings(filepath):
    """
    Attempts to read text files using multiple encoding safe nets to avoid 
    crashing on exotic, legacy, or system-specific text characters.
    """
    encodings = ['utf-8', 'latin-1', 'cp1252', 'utf-16']
    for encoding in encodings:
        try:
            with open(filepath, 'r', encoding=encoding, errors='ignore') as f:
                return f.read().strip()
        except Exception:
            continue
    return ""

def extract_notebook_content(filepath):
    """Extracts markdown and code from an Jupyter Notebook with standard fallback protection."""
    try:
        import nbformat
        content = []
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            nb = nbformat.read(f, as_version=4)
            for cell in nb.cells:
                cell_text = cell.source.strip()
                if not cell_text: 
                    continue
                if cell_text not in seen_content:
                    seen_content.add(cell_text)
                    if cell.cell_type == 'markdown':
                        content.append(cell_text)
                    elif cell.cell_type == 'code':
                        content.append(f"```python\n{cell_text}\n```")
        return "\n\n".join(content)
    except ImportError:
        # Fallback if nbformat is missing: Read it gracefully as syntax-wrapped raw JSON text
        raw_json = read_text_with_fallback_encodings(filepath)
        return f"*Note: 'nbformat' missing. Displaying raw Notebook text layout.*\n\n```json\n{raw_json}\n```"
    except Exception as e:
        return f"*Error parsing Notebook file details: {e}*"

def extract_docx_content(filepath):
    """Extracts paragraphs from Microsoft Word documents."""
    try:
        import docx
        content = []
        doc = docx.Document(filepath)
        for paragraph in doc.paragraphs:
            para_text = paragraph.text.strip()
            if not para_text: 
                continue
            if para_text not in seen_content:
                seen_content.add(para_text)
                content.append(para_text)
        return "\n\n".join(content)
    except ImportError:
        return "*[Dependency Missing: Install 'python-docx' via pip to parse this Word document]*"
    except Exception as e:
        return f"*Error parsing Word document text stream: {e}*"

def extract_pdf_content(filepath):
    """Extracts raw text streams layout from PDF documents."""
    try:
        from pypdf import PdfReader
        content = []
        reader = PdfReader(filepath)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                text_cleaned = text.strip()
                if text_cleaned not in seen_content:
                    seen_content.add(text_cleaned)
                    content.append(text_cleaned)
        return "\n\n".join(content)
    except ImportError:
        return "*[Dependency Missing: Install 'pypdf' via pip to parse this PDF document]*"
    except Exception as e:
        return f"*Error parsing PDF raw text: {e}*"

def extract_pptx_content(filepath):
    """Extracts visual slide text segments from PowerPoint presentations."""
    try:
        from pptx import Presentation
        content = []
        prs = Presentation(filepath)
        for i, slide in enumerate(prs.slides, start=1):
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    text_line = shape.text.strip()
                    if text_line not in seen_content:
                        seen_content.add(text_line)
                        slide_text.append(text_line)
            if slide_text:
                content.append(f"### Slide {i}\n" + "\n".join(slide_text))
        return "\n\n".join(content)
    except ImportError:
        return "*[Dependency Missing: Install 'python-pptx' via pip to parse this PowerPoint presentation]*"
    except Exception as e:
        return f"*Error parsing presentation text blocks: {e}*"

def extract_tabular_content(filepath, is_csv=True):
    """Parses structural tables (CSV/TSV/Excel) converting them into elegant Markdown matrices."""
    try:
        import pandas as pd
        # Choose delimiter engine based on actual extension signatures
        if filepath.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(filepath)
        else:
            delimiter = ',' if is_csv else '\t'
            df = pd.read_csv(filepath, sep=delimiter)
        
        if len(df) > 50:
            table_md = df.head(50).to_markdown(index=False)
            return table_md + f"\n\n*Note: Table truncated dynamically. Displaying first 50 rows out of {len(df)} total rows.*\n"
        return df.to_markdown(index=False)
    except ImportError:
        # If pandas/tabulate are missing, treat CSV/TSV as code blocks so raw data is still preserved
        raw_text = read_text_with_fallback_encodings(filepath)
        return f"*Note: Table formatting missing (pip install pandas tabulate). Rendering raw rows.*\n\n```text\n{raw_text}\n```"
    except Exception as e:
        return f"*Error parsing structured data matrix: {e}*"

def main():
    print("==================================================")
    print("        Universal Code & File Merger Tool         ")
    print("==================================================")
    
    # 1. SCAN DIRECTORY AND APPLY STRATEGIC EXCLUSION RULES
    all_files = []
    try:
        for f in os.listdir('.'):
            if os.path.isdir(f):
                continue
            
            # CRITICAL SAFETY HOOKS
            if f == SCRIPT_SELF_NAME:
                continue
            if os.path.abspath(f) == OUTPUT_FILEPATH_ABS:
                continue
                
            if f.lower().endswith(ALL_SUPPORTED_EXTENSIONS):
                all_files.append(f)
    except Exception as e:
        print(f"[-] Critical failure evaluating target folder path contents: {e}")
        sys.exit(1)

    all_files.sort()
    total_count = len(all_files)
    print(f"[+] Found {total_count} matching files to process safely.\n")

    if total_count == 0:
        print("[!] No supported file structures discovered in the running directory path.")
        return

    # 2. RUN EXTRACTION PIPELINES WITH TOTAL ISOLATION PROTECTION
    processed_count = 0
    with open(OUTPUT_FILENAME, 'w', encoding='utf-8') as outfile:
        outfile.write("# Combined Comprehensive Project Manifest\n\n")
        outfile.write(f"This document houses all project elements automatically generated via `{SCRIPT_SELF_NAME}`.\n\n")
        
        for index, filename in enumerate(all_files, start=1):
            ext = os.path.splitext(filename)[1].lower()
            print(f"[{index}/{total_count}] Processing: {filename}...")
            
            text_payload = ""
            
            # Routing engines based on absolute file type extensions
            if ext == '.ipynb':
                text_payload = extract_notebook_content(filename)
            elif ext == '.docx':
                text_payload = extract_docx_content(filename)
            elif ext == '.pdf':
                text_payload = extract_pdf_content(filename)
            elif ext in ('.pptx', '.ppt'):
                text_payload = extract_pptx_content(filename)
            elif ext in ('.xlsx', '.xls'):
                text_payload = extract_tabular_content(filename, is_csv=False)
            elif ext == '.csv':
                text_payload = extract_tabular_content(filename, is_csv=True)
            elif ext == '.tsv':
                text_payload = extract_tabular_content(filename, is_csv=False)
            elif ext in CODE_EXTENSIONS:
                # Standard script/plain text file extraction
                raw_code = read_text_with_fallback_encodings(filename)
                if raw_code and raw_code not in seen_content:
                    seen_content.add(raw_code)
                    lang = CODE_EXTENSIONS[ext]
                    text_payload = f"```{lang}\n{raw_code}\n```"
            
            # Commit unique extracted elements cleanly into the master document file
            if text_payload and text_payload.strip():
                outfile.write(f"\n\n---\n## File Component: {filename}\n---\n\n")
                outfile.write(text_payload)
                processed_count += 1
            else:
                print(f"  -> Skipped: No unique text blocks extracted or contents are entirely duplicated.")

    print("\n==================================================")
    print(f"[✓] Execution successful! Compiled {processed_count} files.")
    print(f"[+] Output generated: '{OUTPUT_FILENAME}'")
    print("==================================================")

if __name__ == "__main__":
    main()
