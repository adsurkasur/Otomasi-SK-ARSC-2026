import os
import glob
from docx import Document
import re

# BASE_DIR is set to the parent directory (project root) because the script is in scripts/
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def main():
    docx_files = glob.glob(os.path.join(BASE_DIR, "docx", "SK Aktif*.docx"))
    
    if not docx_files:
        print("Tidak ada file 'SK Aktif*.docx' yang ditemukan di folder ini.")
        return

    number_pattern = re.compile(r"No\.\s*(\d+)/SK/ARSC")
    
    file_numbers = {}
    duplicates = {}
    
    print(f"Membaca {len(docx_files)} file SK Aktif...\n")
    for file in docx_files:
        try:
            doc = Document(file)
            found_num = None
            for p in doc.paragraphs:
                text = p.text.strip()
                if text.startswith("No.") and "/SK/ARSC/" in text:
                    match = number_pattern.search(text)
                    if match:
                        found_num = int(match.group(1))
                        break
            
            if found_num is not None:
                if found_num in file_numbers:
                    if found_num not in duplicates:
                        duplicates[found_num] = [file_numbers[found_num]]
                    duplicates[found_num].append(file)
                else:
                    file_numbers[found_num] = file
            else:
                print(f"[!] Peringatan: Nomor surat tidak ditemukan di file '{file}'")
                
        except Exception as e:
            print(f"[!] Error membaca file '{file}': {e}")

    if not file_numbers:
        print("Tidak ada nomor surat yang berhasil diekstrak.")
        return

    all_nums = sorted(file_numbers.keys())
    min_num = all_nums[0]
    max_num = all_nums[-1]
    
    expected_nums = set(range(min_num, max_num + 1))
    actual_nums = set(all_nums)
    missing_nums = sorted(list(expected_nums - actual_nums))
    
    print("=" * 60)
    print("LAPORAN PENGECEKAN NOMOR SURAT")
    print("=" * 60)
    print(f"Total surat dianalisis : {len(docx_files)}")
    print(f"Rentang nomor surat    : {min_num:03d} s/d {max_num:03d}")
    print("-" * 60)
    
    if duplicates:
        print("\n[PERINGATAN] Ditemukan Nomor Duplikat:")
        for num, files in duplicates.items():
            print(f"  - No. {num:03d} dipakai oleh {len(files)} file:")
            for f in files:
                print(f"      -> {f}")
    else:
        print("\n[OK] Tidak ada nomor yang duplikat.")
        
    if missing_nums:
        print("\n[PERINGATAN] Ditemukan Nomor yang Terlewat (Skip):")
        for num in missing_nums:
            print(f"  - No. {num:03d} tidak ditemukan di file mana pun.")
    else:
        print("\n[OK] Tidak ada urutan nomor yang terlewat.")
        
    print("\n" + "=" * 60)

if __name__ == "__main__":
    main()
