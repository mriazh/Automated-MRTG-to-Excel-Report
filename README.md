# 📊 Automated MRTG to Excel Report (Universal Edition)

[![Python](https://img.shields.io/badge/Python-3.8--3.12-blue?logo=python)](https://www.python.org/)
[![PaddleOCR](https://img.shields.io/badge/OCR-PaddleOCR%203.x-red)](https://github.com/PaddlePaddle/PaddleOCR)
[![Status](https://img.shields.io/badge/Status-Production--Ready-green)](#)

Bot otomatisasi untuk mengkonversi *screenshot* MRTG TelkomCare menjadi laporan Excel yang rapi. Versi ini telah dioptimalkan dengan sistem **"Total Blackout"** untuk memastikan tampilan terminal tetap bersih dan kompatibilitas tinggi di berbagai arsitektur CPU (termasuk Intel 11th Gen+).

---

## ✨ Fitur Utama
- **🧠 Intelligent OCR Mode**: Ekstrak data bandwidth (Current, Avg, Max) secara otomatis + Insert gambar ke Excel.
- **🖼️ Smart Image Only Mode**: Masukkan gambar saja tanpa proses OCR (Kecepatan 10x lipat).
- **🛡️ Silent Engine Protocol**: Semua log berisik (ccache, oneDNN, Windows patterns) disembunyikan total agar terminal tetap bersih.
- **⚡ Universal Compatibility**: Fix otomatis untuk error *PIR/oneDNN* pada CPU Intel terbaru (Tiger Lake, Alder Lake, dll).
- **📝 Audit-Ready Logging**: Log detail ekstraksi tersimpan rapi di `ocr_report.log` untuk pengecekan manual.
- **🎯 Smart Summary**: Menampilkan daftar SID yang butuh review (N/A values) secara otomatis di akhir proses.

---

## 🛠️ Cara Instalasi (Windows)

### 1. Prasyarat
- **Python 3.12 (Sangat Direkomendasikan)**: [Download 64-bit](https://www.python.org/downloads/windows/).
- **PENTING**: **JANGAN** gunakan Python 3.13 ke atas karena library OCR belum mendukung versi tersebut.

### 2. Setup Awal
Buka terminal (PowerShell atau CMD) di folder project ini:

```powershell
# 1. Buat Virtual Environment (Paksa versi 3.12 jika ada banyak versi)
py -3.12 -m venv .venv

# 2. Aktifkan Environment (PowerShell)
.\.venv\Scripts\Activate.ps1

# 3. Aktifkan Environment (Command Prompt)
.\.venv\Scripts\Activate

# 4. Install Dependensi
pip install -r requirements.txt
```

---

## 🚀 Cara Penggunaan

1. **Siapkan Data**: Letakkan folder *screenshot* MRTG di dalam folder `MRTG-Data`.
   - Struktur: `MRTG-Data/YYYYMMDD/*.png`
2. **Jalankan Bot**:
   ```powershell
   python mrtg_data_to_monthly_report.py
   ```
3. **Pilih Mode**: Masukkan angka `1` (OCR) atau `2` (Image Only).
4. **Hasil**: Laporan Excel akan tercipta secara otomatis di folder utama.

---

## 🔍 Troubleshooting (FAQ)

**Q: Kenapa terminal sangat hening saat awal dijalankan?**
> A: Script menggunakan sistem **Total Blackout** untuk menyembunyikan log inisialisasi library yang berisik. Bot sebenarnya sedang menyiapkan mesin OCR di balik layar.

**Q: Bagaimana jika ada data yang terbaca N/A?**
> A: Bot akan menandai item tersebut sebagai `⚠️ PARTIAL`. Cek daftar review di akhir proses atau buka `ocr_report.log` untuk melihat detail teks yang terdeteksi.

**Q: Muncul error saat instalasi pymupdf?**
> A: Pastikan Anda menggunakan **Python 3.12**. Versi 3.13+ akan gagal mengompilasi library ini.

---
**Build with ❤️ for Telkom-GMF Reporting Efficiency.**
