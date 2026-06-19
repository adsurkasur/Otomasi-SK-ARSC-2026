import os
import pandas as pd
from docx import Document

# BASE_DIR is set to the parent directory (project root) because the script is in scripts/
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 1. Read Excel
excel_path = os.path.join(BASE_DIR, 'data', 'PENDATAAN ANGGOTA AKTIF ARSC PERIODE 2025_2026.xlsx')
try:
    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path)
        print("Excel Columns:")
        print(df.columns.tolist())
        print("\nFirst row of data:")
        print(df.iloc[0].to_dict())
    else:
        print(f"Excel file not found at {excel_path}")
except Exception as e:
    print(f"Error reading Excel: {e}")

# 2. Read Word from docx dir
try:
    docx_dir = os.path.join(BASE_DIR, 'docx')
    if os.path.exists(docx_dir):
        files = os.listdir(docx_dir)
        docx_files = [f for f in files if f.endswith('.docx') and not f.startswith('~$')]
        if docx_files:
            sample_file = os.path.join(docx_dir, docx_files[0])
            doc = Document(sample_file)
            print(f"\nWord Document Paragraphs for {docx_files[0]}:")
            for i, p in enumerate(doc.paragraphs):
                if p.text.strip():
                    print(f"[{i}]: {p.text}")
        else:
            print("\nNo .docx files found in docx folder.")
    else:
        print(f"\ndocx folder not found at {docx_dir}")
except Exception as e:
    print(f"Error reading Word: {e}")
