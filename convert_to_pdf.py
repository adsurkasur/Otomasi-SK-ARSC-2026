import sys
import os
import glob

try:
    from docx2pdf import convert
except ImportError:
    print("Error: Library 'docx2pdf' tidak ditemukan. Silakan install dengan 'pip install docx2pdf'")
    sys.exit(1)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def main():
    print("=== DOCX to PDF Converter ===")
    
    # Jika ada argumen yang diberikan (nama file atau folder)
    if len(sys.argv) > 1:
        target = sys.argv[1]
        if os.path.exists(target):
            if os.path.isfile(target) and target.lower().endswith(".docx"):
                pdf_target = target[:-5] + ".pdf"
                if os.path.exists(pdf_target):
                    print(f"Melewati '{target}' karena file PDF sudah ada.")
                    return
                    
            print(f"Mengonversi '{target}' ke PDF...")
            try:
                convert(target)
                print("Konversi selesai!")
            except Exception as e:
                print(f"Gagal mengonversi: {e}")
        else:
            print(f"Error: Target '{target}' tidak ditemukan.")
    else:
        # Jika tidak ada argumen, cari semua file .docx di direktori docx
        docx_files = glob.glob(os.path.join(BASE_DIR, "docx", "*.docx"))
        
        # Abaikan temporary files dari Word (biasanya diawali dengan ~$)
        docx_files = [f for f in docx_files if not f.startswith("~$")]
        
        if not docx_files:
            print("Tidak ada file .docx yang ditemukan di direktori saat ini.")
            return
        
        print(f"Ditemukan {len(docx_files)} file .docx. Memulai konversi...")
        
        sukses = 0
        gagal = 0
        dilewati = 0
        
        for file in docx_files:
            pdf_file = os.path.join(BASE_DIR, "pdf", os.path.basename(file)[:-5] + ".pdf")
            if os.path.exists(pdf_file):
                print(f"\nMelewati: {file} (sudah ada PDF)")
                dilewati += 1
                continue
                
            print(f"\nMengonversi: {file}")
            try:
                convert(file, pdf_file)
                sukses += 1
            except Exception as e:
                print(f"Gagal mengonversi '{file}': {e}")
                gagal += 1
                
        print("\n=== Ringkasan ===")
        print(f"Total file: {len(docx_files)}")
        print(f"Berhasil  : {sukses}")
        print(f"Dilewati  : {dilewati}")
        print(f"Gagal     : {gagal}")

if __name__ == "__main__":
    main()
