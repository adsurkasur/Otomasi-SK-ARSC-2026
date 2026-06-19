import pandas as pd
from docx import Document

# 1. Read Excel
try:
    df = pd.read_excel('PENDATAAN ANGGOTA AKTIF ARSC PERIODE 2025_2026 .xlsx')
    print("Excel Columns:")
    print(df.columns.tolist())
    print("\nFirst row of data:")
    print(df.iloc[0].to_dict())
except Exception as e:
    print(f"Error reading Excel: {e}")

# 2. Read Word
try:
    doc = Document('SK Aktif Aisyah Putri Aryadhie Zahirah.docx')
    print("\nWord Document Paragraphs:")
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            print(f"[{i}]: {p.text}")
except Exception as e:
    print(f"Error reading Word: {e}")
