# 📊 Automated MRTG to Excel Report

> Bot otomatisasi untuk mengkonversi screenshot grafik MRTG TelkomCare menjadi laporan Excel yang rapi. Mendukung dua mode: ekstraksi data penuh via OCR dan mode gambar-saja untuk kecepatan maksimal.

[![Python](https://img.shields.io/badge/Python-3.12-blue?logo=python)](https://www.python.org/) [![PaddleOCR](https://img.shields.io/badge/OCR-PaddleOCR%203.x-red)](https://github.com/PaddlePaddle/PaddleOCR) [![openpyxl](https://img.shields.io/badge/Excel-openpyxl-green)](https://openpyxl.readthedocs.io/) [![License](https://img.shields.io/badge/License-MIT-yellow)](#)

---

## ✨ Fitur Utama

- **🧠 Intelligent OCR Mode** — Ekstrak data bandwidth (Current, Avg, Max) secara otomatis dari gambar grafik, lalu tulis hasilnya ke sel Excel beserta gambar aslinya.
- **🖼️ Smart Image Only Mode** — Masukkan gambar langsung ke Excel tanpa proses OCR. Kecepatan hingga 10x lebih cepat, cocok jika data angka tidak diperlukan.
- **🛡️ Silent Engine Protocol** — Semua log berisik dari library (PaddlePaddle, oneDNN) disembunyikan total agar tampilan terminal tetap bersih dan mudah dipantau.
- **⚡ Universal Compatibility** — Fix otomatis untuk error *PIR/oneDNN* pada CPU Intel generasi baru (Tiger Lake, Alder Lake, dll).
- **📝 Audit-Ready Logging** — Log detail ekstraksi tersimpan rapi di `ocr_report.log` untuk keperluan pengecekan manual.
- **🎯 Smart Summary** — Menampilkan daftar SID yang memerlukan review (nilai N/A) secara otomatis di akhir proses.

---

## 🏗️ Arsitektur

```
Automated-MRTG-to-Excel-Report/
├── mrtg_data_to_monthly_report.py   ← Script utama (entry point)
├── requirements.txt                 ← Dependencies Python
├── MRTG-Data/                       ← Folder input screenshot (isi sendiri)
│   └── YYYYMMDD/                    ← Sub-folder per tanggal
│       └── *.png                    ← File screenshot MRTG
├── list_mrtg_data.txt               ← Daftar urutan SID untuk laporan OCR
├── list_mrtg_data_img_only.txt      ← Daftar urutan SID untuk laporan Image Only
├── list_mrtg_data_position.txt      ← Konfigurasi posisi SID di Excel (OCR)
├── list_mrtg_data_position_img_only.txt ← Konfigurasi posisi SID di Excel (Image Only)
├── ocr_report.log                   ← Log hasil ekstraksi (otomatis dibuat)
├── *.xlsx                           ← Hasil laporan Excel (otomatis dibuat)
├── .gitignore
├── LICENSE
└── README.md
```

---

## 🛠️ Instalasi

### 1. Prasyarat Sistem

> **⚠️ PENTING:** Gunakan **Python 3.12**. Python 3.13 ke atas belum didukung oleh library OCR yang digunakan (`pymupdf`, `PaddleOCR`).

| Software | Versi | Keterangan |
|----------|-------|------------|
| **Python** | **3.12** (wajib) | [Download 64-bit di sini](https://www.python.org/downloads/windows/) |
| **pip** | Terbaru | Sudah termasuk dalam instalasi Python |

### 2. Clone Repository

```bash
git clone https://github.com/AdimasP/Automated-MRTG-to-Excel-Report.git
cd Automated-MRTG-to-Excel-Report
```

### 3. Buat Virtual Environment

Sangat disarankan menggunakan *virtual environment* agar tidak bentrok dengan instalasi Python lain di sistem Anda.

```powershell
# Buat virtual environment (pastikan menggunakan Python 3.12)
py -3.12 -m venv .venv

# Aktifkan environment (PowerShell)
.\.venv\Scripts\Activate.ps1

# Aktifkan environment (Command Prompt)
.\.venv\Scripts\Activate
```

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

> **Catatan:** Proses instalasi pertama akan memakan waktu cukup lama karena mengunduh library PaddlePaddle dan model OCR-nya (±500MB). Pastikan koneksi internet stabil.

---

## ⚙️ Konfigurasi

Sebelum menjalankan bot, siapkan data input Anda:

### 1. Siapkan Folder Screenshot

Letakkan hasil screenshot MRTG dari TelkomCare ke dalam folder `MRTG-Data/`, dikelompokkan per tanggal:

```
MRTG-Data/
├── 20260501/
│   ├── 4700001-0021497479.png
│   └── 4700001-0020265222.png
└── 20260502/
    ├── 4700001-0021497479.png
    └── 4700001-0020265222.png
```

### 2. Konfigurasi Daftar SID

Edit file `list_mrtg_data.txt` (untuk Mode OCR) atau `list_mrtg_data_img_only.txt` (untuk Mode Image Only) sesuai dengan SID yang dimiliki. File-file ini menentukan urutan SID yang akan masuk ke laporan Excel.

---

## 🚀 Cara Penggunaan

### Langkah 1: Pastikan Virtual Environment Aktif

```powershell
# Aktifkan terlebih dahulu jika belum
.\.venv\Scripts\Activate.ps1
```

### Langkah 2: Jalankan Script

```bash
python mrtg_data_to_monthly_report.py
```

### Langkah 3: Pilih Mode

Bot akan menampilkan menu pilihan:

```
[1] OCR Mode       — Ekstrak data angka (Current/Avg/Max) + insert gambar ke Excel
[2] Image Only     — Insert gambar saja tanpa OCR (jauh lebih cepat)
```

Masukkan angka `1` atau `2` sesuai kebutuhan, lalu tekan Enter.

### Langkah 4: Tunggu Proses Selesai

- Bot akan memproses semua gambar dalam folder `MRTG-Data/` secara otomatis.
- Progress ditampilkan di terminal secara real-time.
- Di akhir proses, bot menampilkan **Smart Summary** berisi daftar SID yang perlu dicek ulang (nilai terbaca N/A).

### Langkah 5: Ambil Hasil

File laporan Excel (`.xlsx`) akan dibuat secara otomatis di folder utama. Nama file mencerminkan rentang tanggal data yang diproses.

---

## ⚠️ Troubleshooting

**Q: Terminal sangat hening / tidak ada output saat pertama dijalankan?**
> A: Ini normal. Script menggunakan sistem **Silent Engine Protocol** untuk menyembunyikan log inisialisasi library yang berisik. Bot sedang memuat mesin OCR di balik layar. Tunggu hingga muncul output proses.

**Q: Ada data yang terbaca N/A di hasil Excel?**
> A: Bot akan menandai item tersebut sebagai `⚠️ PARTIAL`. Cek daftar review di akhir proses atau buka `ocr_report.log` untuk melihat teks mentah yang berhasil terdeteksi OCR pada gambar tersebut.

**Q: Muncul error saat instalasi `pymupdf`?**
> A: Pastikan Anda menggunakan **Python 3.12**. Versi 3.13 ke atas akan gagal mengompilasi library ini. Gunakan `py -3.12 -m venv .venv` saat membuat virtual environment.

**Q: Error `oneDNN` atau `PIR` saat pertama jalan di CPU Intel baru?**
> A: Script sudah memiliki fix otomatis untuk error ini. Pesan error tersebut akan disembunyikan dan tidak mempengaruhi hasil laporan.

---

## 📄 License

[MIT License](LICENSE)
