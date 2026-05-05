# 📊 Automated MRTG to Excel Report

**Kumpulan script otomatis untuk memproses gambar grafik MRTG dan menyusunnya ke dalam template laporan Excel bulanan.**

Repository ini berisi **dua versi** script dengan pendekatan berbeda:

| Versi | Folder | Deskripsi |
|-------|--------|-----------|
| 🔍 **With OCR** | `with-OCR/` | Membaca nilai bandwidth dari gambar MRTG menggunakan Tesseract OCR, lalu memasukkan data + gambar ke Excel. |
| 🖼️ **Image Only** | `image-only/` | Menempatkan gambar MRTG secara presisi ke dalam Excel **tanpa OCR** — lebih cepat dan ringan. |

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![OpenCV](https://img.shields.io/badge/OpenCV-4.x-green)](https://opencv.org/)
[![Tesseract](https://img.shields.io/badge/Tesseract-5.x-orange)](https://github.com/UB-Mannheim/tesseract)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.x-yellow)](https://openpyxl.readthedocs.io/)

---

## 🔍 Versi 1: With OCR (`with-OCR/`)

### 📌 Fitur Utama

- ✅ **Ekstraksi Otomatis OCR** – Mengambil nilai *Inbound/Outbound (Current, Average, Maximum)* dari gambar MRTG menggunakan Tesseract OCR.
- ✅ **Pemrosesan Gambar Pintar** – Menggunakan OpenCV untuk *preprocessing* gambar (grayscale, thresholding, upscaling) guna meningkatkan akurasi teks.
- ✅ **Regex Parsing Kuat** – Menangani format nilai bandwidth yang bervariasi dari teks hasil OCR secara cerdas.
- ✅ **Otomatisasi Excel** – Memasukkan nilai hasil ekstraksi secara akurat sesuai sel yang ditentukan (*mapping* dinamis).
- ✅ **Penempatan Gambar di Excel** – Melakukan *resize* dan menempelkan gambar MRTG asli tepat di area sel yang telah dialokasikan.
- ✅ **Multi-Sheet Harian** – Membuat *sheet* baru secara otomatis berdasarkan tanggal gambar yang diproses.

### 🛠️ Prasyarat

| Software | Keterangan |
|----------|-------------|
| **Python 3.8+** | [Download](https://www.python.org/downloads/) |
| **Tesseract OCR** | [Download](https://github.com/UB-Mannheim/tesseract/wiki) – **Centang "Add to PATH"** |
| **Template Excel** | File `Report on Internet Bandwidth Utilization by Telkom (MRTG).xlsx` |

### 📦 Instalasi

```bash
pip install opencv-python numpy openpyxl pillow pytesseract
```

Edit path Tesseract di script jika berbeda:
```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

### 📁 Persiapan File

Pastikan file berikut tersedia di dalam folder `with-OCR/`:

1. **`list_mrtg_data.txt`** – Daftar SID atau Graph Title yang akan diproses.
2. **`sid-in-out-image-position-excel.txt`** – Konfigurasi mapping sel Excel untuk nilai OCR dan gambar MRTG.
3. **Template Excel** – File template laporan kosong.
4. **Folder `MRTG-Data/`** – Folder berisi direktori tanggal (`YYYYMMDD`) yang menyimpan gambar hasil screenshot MRTG.

### 🚀 Cara Penggunaan

```bash
cd with-OCR
python mrtg_data_to_monthly_report.py
```

**Alur:**
- Pastikan semua file konfigurasi dan gambar MRTG sudah disiapkan di folder `MRTG-Data/<Tanggal>/`.
- Jalankan script di terminal.
- Script akan membaca folder data satu per satu, melakukan ekstraksi OCR, dan menuliskannya ke Excel.
- **Hasil akhir** akan disimpan dengan nama `Complete_Monthly_Report.xlsx`.

---

## 🖼️ Versi 2: Image Only (`image-only/`)

### 📌 Fitur Utama

- ✅ **Penempatan Gambar Super Cepat** – Memasukkan gambar ke dalam Excel tanpa overhead OCR.
- ✅ **Resize Proporsional Dinamis** – Menyesuaikan ukuran gambar secara presisi agar pas dengan area sel Excel yang telah ditentukan di konfigurasi.
- ✅ **Mapping Area Fleksibel** – Menggunakan file teks sederhana untuk memetakan ID gambar ke sel awal dan akhir (misal: `B12-L23`).
- ✅ **Multi-Sheet Harian** – Otomatis mendeteksi folder tanggal dan membuat *sheet* harian (1-31) di dalam file Excel.

### 🛠️ Prasyarat

| Software | Keterangan |
|----------|-------------|
| **Python 3.8+** | [Download](https://www.python.org/downloads/) |
| **Template Excel** | File `MENTAHAN FORMAT DAILY MRTG.xlsx` |

*(Catatan: Versi ini **tidak membutuhkan** instalasi Tesseract.)*

### 📦 Instalasi

```bash
pip install openpyxl pillow
```

### 📁 Persiapan File

Pastikan file berikut tersedia di dalam folder `image-only/`:

1. **`list_mrtg_data.txt`** – Berisi daftar urutan SID atau Graph Title yang akan di-*insert*.
2. **`sid_image-position-excel.txt`** – Konfigurasi lokasi sel (area letak gambar) di Excel untuk setiap ID.
3. **Template Excel** – File template laporan kosong.
4. **Folder `MRTG-Data/`** – Folder berisi folder berformat tanggal `YYYYMMDD` yang menyimpan gambar-gambar MRTG.

### 🚀 Cara Penggunaan

```bash
cd image-only
python script_ini.py
```

**Alur:**
- Pastikan semua gambar sudah tersimpan rapi berdasarkan tanggal di folder `MRTG-Data/`.
- Jalankan script di terminal.
- Script akan langsung memetakan semua gambar dari setiap folder tanggal ke sheet masing-masing di file Excel.
- **Hasil akhir** akan tersimpan dalam file baru bernama `Daily_Report_Complete.xlsx`.

---

## 📂 Struktur Repository

```
Automated-MRTG-to-Excel-Report/
├── with-OCR/                    # Versi dengan OCR
│   ├── mrtg_data_to_monthly_report.py
│   ├── list_mrtg_data.txt
│   ├── sid-in-out-image-position-excel.txt
│   ├── Template Excel (.xlsx)
│   └── MRTG-Data/
├── image-only/                  # Versi tanpa OCR (image saja)
│   ├── script_ini.py
│   ├── list_mrtg_data.txt
│   ├── sid_image-position-excel.txt
│   ├── Template Excel (.xlsx)
│   └── MRTG-Data/
└── README.md
```

---

**Pilih versi yang sesuai kebutuhan Anda dan selamat menyelesaikan pelaporan MRTG! 🚀**
