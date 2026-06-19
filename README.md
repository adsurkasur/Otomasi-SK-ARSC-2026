# Otomasi Surat Keterangan (SK) ARSC 2026

Proyek ini dirancang untuk mempermudah dan mengotomatiskan pembuatan **Surat Keterangan Aktif** bagi anggota pengurus ARSC (Agricultural Technology Research Student Club) Fakultas Teknologi Pertanian Universitas Brawijaya.

Aplikasi ini membaca basis data anggota aktif dari Excel, mencocokkannya dengan daftar nama yang ingin dibuatkan surat, lalu menghasilkan dokumen Word (.docx) berbasis template dengan penomoran surat yang otomatis dan terurut. Proyek ini juga dilengkapi dengan alat pemantau penomoran surat dan konversi otomatis dari format `.docx` ke `.pdf`.

---

## 📂 Struktur Direktori & File

- `scripts/`: Berisi seluruh skrip Python utama (`generate_sk.py`, `cek_nomor.py`, `convert_to_pdf.py`, `peek.py`).
- `data/` *(Diabaikan oleh Git)*: Tempat menyimpan file Excel sumber pendataan anggota aktif (misal: `PENDATAAN ANGGOTA AKTIF ARSC PERIODE 2025_2026 .xlsx`).
- `templates/`: Menyimpan file template Word (`Template_SK.docx`). Semua font pada template ini telah dikonversi secara ketat menggunakan **Times New Roman**.
- `docx/` *(Diabaikan oleh Git)*: Folder output hasil generate Surat Keterangan dalam format Word (`.docx`).
- `pdf/` *(Diabaikan oleh Git)*: Folder output hasil konversi otomatis dari format Word (`.docx`) ke format PDF (`.pdf`).
- `daftar.txt` *(Diabaikan oleh Git)*: File teks untuk memasukkan daftar nama yang ingin diproses (satu nama per baris).
- `nomor_terakhir.txt` *(Diabaikan oleh Git)*: File konfigurasi kecil untuk mencatat nomor urut surat terakhir guna melanjutkan penomoran secara berurutan.

---

## 🚀 Fitur Utama

1. **Otomatisasi SK (Generate DOCX)**
   - Mendukung pencarian nama pintar berbasis kemiripan teks (`difflib`) jika terjadi sedikit saltik (typo) pada daftar nama.
   - Penomoran otomatis menggunakan format resmi: `No. [Nomor]/SK/ARSC/[Bulan-Romawi]/[Tahun]`.
   - Mengisi otomatis seluruh placeholder pada template: `[NAMA]`, `[NIM]`, `[DEPARTEMEN]`, `[PROGRAM_STUDI]`, `[JABATAN]`, dan `[TANGGAL_SURAT]`.

2. **Validasi & Pengecekan Nomor Surat**
   - Menganalisis seluruh file di folder `docx/` untuk mendeteksi apakah terdapat nomor surat yang duplikat, terlewat, atau tidak berurutan.

3. **Konversi DOCX ke PDF Massal**
   - Mengonversi semua berkas Word (`.docx`) di folder `docx/` yang belum memiliki salinan PDF secara massal menggunakan background process Word.

4. **Sistem Virtual Environment Mandiri (Auto Setup)**
   - Menyediakan skrip `.bat` untuk otomatisasi pemasangan environment Python tanpa konfigurasi manual yang rumit di laptop/PC baru.

---

## 💻 Cara Penggunaan

### 1. Persiapan Awal
Pastikan file data anggota aktif (`PENDATAAN ANGGOTA AKTIF ARSC PERIODE 2025_2026 .xlsx`) berada di dalam folder `data/` dan template `Template_SK.docx` berada di folder `templates/`.

### 2. Membuat SK (Generate DOCX)
1. Buka file `daftar.txt`.
2. Masukkan nama-nama pengurus yang akan dibuatkan SK (satu nama per baris). Simpan perubahan.
3. Tarik (*Drag and Drop*) berkas `daftar.txt` ke atas berkas `run_generate.bat`, atau cukup klik 2x `run_generate.bat`.
4. Hasil dokumen akan tersimpan di dalam folder `docx/`.

### 3. Mengecek Nomor Surat
1. Klik 2x berkas `cek_nomor_surat.bat`.
2. Jendela prompt akan menampilkan urutan nomor surat dan mendeteksi jika ada anomali atau duplikasi penomoran.

### 4. Mengonversi ke PDF
1. Klik 2x berkas `run_convert_pdf.bat`.
2. Sistem akan mendeteksi file `.docx` baru dan mengonversinya menjadi `.pdf` di dalam folder `pdf/`.

---

## 🛠️ Pengembangan & Virtual Environment

Proyek ini menggunakan virtual environment Python (`venv`) untuk mengelola dependensi. Jika Anda memindahkan folder ini ke komputer lain, gunakan script pembantu berikut:
- **`setup_venv.bat`**: Menginstal ulang virtual environment secara otomatis dan mengunduh dependensi dari `requirements.txt`.
- **`remove_venv.bat`**: Menghapus folder `venv` yang ada apabila terjadi error atau ingin melakukan instalasi bersih.
