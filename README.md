# 📊 Automated MRTG to Excel Report

**Satu script, dua mode! Otomatis memproses gambar grafik MRTG dan menyusunnya ke dalam template laporan Excel bulanan.**

Repository ini berisi **satu script Python utama** (`mrtg_data_to_monthly_report.py`) yang mendukung dua mode eksekusi:

| Mode | Deskripsi |
|------|-----------|
| 🔍 **[1] OCR Mode** | Membaca nilai bandwidth (In/Out, Current/Avg/Max) dari gambar MRTG menggunakan Tesseract OCR, lalu memasukkan teks **dan** gambar presisi ke Excel. |
| 🖼️ **[2] Image Only** | Menempatkan gambar MRTG secara presisi ke dalam Excel **tanpa OCR** — prosesnya sangat cepat dan ringan! |

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![OpenCV](https://img.shields.io/badge/OpenCV-4.x-green)](https://opencv.org/)
[![Tesseract](https://img.shields.io/badge/Tesseract-5.x-orange)](https://github.com/UB-Mannheim/tesseract)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.x-yellow)](https://openpyxl.readthedocs.io/)

---

## 🚀 Cara Penggunaan

1. **Jalankan script utama:**
   ```bash
   python mrtg_data_to_monthly_report.py
   ```
2. **Menu interaktif akan muncul:**
   ```text
   ============================================================
     AUTOMATED MRTG TO EXCEL REPORT
   ============================================================
     Pilih mode:
     [1] OCR Mode   : Ekstrak data + insert gambar ke Excel
     [2] Image Only : Insert gambar saja ke Excel (tanpa OCR)
   ============================================================
     >> Masukkan pilihan (1/2): 
   ```
3. Pilih mode sesuai dengan template laporan dan kebutuhan Anda.
4. **Selesai!** Script akan otomatis memproses semua folder tanggal di dalam `MRTG-Data/`.

---

## 📁 Persiapan File & Struktur Folder

Semua file konfigurasi sudah berada di folder utama (*root*) repository. Pastikan Anda memiliki struktur berikut:

```text
Automated-MRTG-to-Excel-Report/
├── mrtg_data_to_monthly_report.py
├── list_mrtg_data.txt                       # Daftar data (Mode OCR)
├── list_mrtg_data_img_only.txt              # Daftar data (Mode Image Only)
├── list_mrtg_data_position.txt              # Mapping letak sel Excel (Mode OCR)
├── list_mrtg_data_position_img_only.txt     # Mapping letak sel Excel (Mode Image Only)
├── MRTG-Monthly-Report-on-Internet-Bandwidth-Utilization-by-Telkom.xlsx            # Template OCR
├── MRTG-Monthly-Report-on-Internet-Bandwidth-Utilization-by-Telkom (Img only).xlsx # Template Image Only
└── MRTG-Data/                               # Folder berisi gambar MRTG harian
    ├── 20260101/
    ├── 20260102/
    └── ...
```

---

## 🛠️ Prasyarat & Instalasi

### Library Python yang dibutuhkan:
```bash
pip install opencv-python numpy openpyxl pillow pytesseract
```

### 1. Mode OCR (Butuh Tesseract)
Jika Anda menggunakan **Mode 1 (OCR)**, pastikan [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) sudah terinstall.
Jika path instalasinya berbeda, edit baris ini di dalam `mrtg_data_to_monthly_report.py`:
```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

### 2. Mode Image Only (Tanpa Tesseract)
Jika Anda hanya menggunakan **Mode 2 (Image Only)**, instalasi Tesseract dan OpenCV **tidak diwajibkan**, karena library tersebut baru akan diload ketika Mode 1 dipilih.

---

**Cepat dan Presisi! Selamat menyelesaikan pelaporan bulanan Anda! 🚀**
