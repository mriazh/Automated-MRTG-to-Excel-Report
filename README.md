# 📊 Automated MRTG Data to Excel Report

**Script otomatis untuk membaca nilai bandwidth dari gambar grafik MRTG menggunakan OCR, lalu menyusunnya ke dalam template laporan Excel bulanan.**

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![OpenCV](https://img.shields.io/badge/OpenCV-4.x-green)](https://opencv.org/)
[![Tesseract](https://img.shields.io/badge/Tesseract-5.x-orange)](https://github.com/UB-Mannheim/tesseract)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.x-yellow)](https://openpyxl.readthedocs.io/)

---

## 📌 Fitur Utama

- ✅ **Ekstraksi Otomatis OCR** – Mengambil nilai *Inbound/Outbound (Current, Average, Maximum)* dari gambar MRTG menggunakan Tesseract OCR.
- ✅ **Pemrosesan Gambar Pintar** – Menggunakan OpenCV untuk *preprocessing* gambar (grayscale, thresholding, upscaling) guna meningkatkan akurasi teks.
- ✅ **Regex Parsing Kuat** – Menangani format nilai bandwidth yang bervariasi dari teks hasil OCR secara cerdas.
- ✅ **Otomatisasi Excel** – Memasukkan nilai hasil ekstraksi secara akurat sesuai sel yang ditentukan (*mapping* dinamis).
- ✅ **Penempatan Gambar di Excel** – Melakukan *resize* dan menempelkan (*insert*) gambar MRTG asli tepat di area sel yang telah dialokasikan.
- ✅ **Multi-Sheet Harian** – Membuat *sheet* baru secara otomatis berdasarkan tanggal gambar yang diproses.

---

## 🛠️ Prasyarat

| Software | Keterangan |
|----------|-------------|
| **Python 3.8+** | [Download](https://www.python.org/downloads/) |
| **Tesseract OCR** | [Download](https://github.com/UB-Mannheim/tesseract/wiki) – **Centang "Add to PATH"** |
| **Template Excel** | File `Report on Internet Bandwidth Utilization by Telkom (MRTG).xlsx` |

### Environment Variables (PATH)
Pastikan folder instalasi Tesseract dan Python sudah ada di dalam PATH system Anda:
```
C:\Users\<username>\AppData\Local\Python\Scripts
C:\Program Files\Tesseract-OCR
```

---

## 📦 Instalasi

1. **Clone repository** (atau pindah ke folder proyek)
   ```bash
   cd Automated-MRTG-to-Excel-Report
   ```

2. **Buat virtual environment (opsional tapi disarankan)**
   ```bash
   python -m venv venv
   venv\Scripts\activate      # Windows
   ```

3. **Install library Python yang dibutuhkan**
   ```bash
   pip install opencv-python numpy openpyxl pillow pytesseract
   ```

4. **Konfigurasi Tesseract**  
   Edit baris berikut di script `mrtg_data_to_monthly_report.py` jika path Tesseract Anda berbeda:
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
   ```

---

## 📁 Persiapan File & Struktur Folder

Sebelum menjalankan script, pastikan file dan folder berikut tersedia di direktori yang sama:

1. **`list_mrtg_data.txt`**  
   Daftar SID atau Graph Title yang akan diproses.
2. **`sid-in-out-image-position-excel.txt`**  
   File konfigurasi yang memetakan di sel Excel mana nilai OCR dan gambar MRTG harus diletakkan.
3. **`Report on Internet Bandwidth Utilization by Telkom (MRTG).xlsx`**  
   File *template* laporan Excel kosong yang akan diisi oleh script.
4. **Folder `MRTG-Data/`**  
   Folder berisi direktori tanggal (format `YYYYMMDD`) yang menyimpan gambar hasil screenshot MRTG.
   ```text
   MRTG-Data/
   ├── 20260101/
   │   ├── MRTG_4700001-0021497479.png
   │   └── MRTG_3598_20260101.png
   ├── 20260102/
   └── ...
   ```

---

## 🚀 Cara Penggunaan

1. Siapkan semua file konfigurasi dan letakkan gambar MRTG di folder `MRTG-Data/<Tanggal>/`.
2. Jalankan script di terminal:
   ```bash
   python mrtg_data_to_monthly_report.py
   ```
3. Script akan membaca folder data satu per satu, melakukan ekstraksi OCR, dan menuliskannya ke Excel.
4. **Hasil akhir** akan disimpan dengan nama `Complete_Monthly_Report.xlsx`.

---

**Selamat mencoba dan semoga pelaporan bulanan MRTG Anda menjadi lebih cepat dan efisien! 🚀**
