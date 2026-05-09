# 🚀 Automated MRTG to Excel Report

Alat bantu otomatisasi tingkat tinggi untuk mengolah hasil screenshot MRTG menjadi laporan Excel yang rapi. Bot ini bisa membaca angka bandwidth (Inbound/Outbound) secara otomatis menggunakan teknologi OCR atau sekadar melakukan *image-injection* massal ke dalam template Excel.

---

## ✨ Fitur Utama

- **🧠 Dual Engine Mode**:
  - **[1] OCR Mode**: Ekstraksi data otomatis dari grafik menggunakan **PaddleOCR** dan memasukkannya ke sel Excel yang spesifik.
  - **[2] Image Only**: Memasukkan gambar grafik MRTG ke dalam Excel tanpa ekstraksi teks (Sangat Cepat).
- **📊 Real-time Dashboard**: Tampilan terminal yang bersih dengan progres bar interaktif dan statistik (OK, Partial, Fail).
- **📏 Smart Image Scaling**: Gambar otomatis di-*resize* secara proporsional agar pas dengan area cell di template Excel lu.
- **🔇 Stealth Logging**: Log internal library (Paddle/C++) dibungkam agar terminal lu tetap bersih dan fokus pada progres.
- **🔍 Audit Trails**: Kegagalan ekstraksi atau gambar yang hilang dicatat secara detail di `ocr_report.log` untuk review manual.

---

## 🛠️ Panduan Instalasi (Penting!)

Agar bot berjalan lancar tanpa error `No module named 'paddleocr'`, ikuti langkah-langkah ini:

### 1. Prasyarat Sistem
- **Python 3.8 - 3.11** (Direkomendasikan 3.10 atau 3.11). 
- **PENTING**: Wajib menggunakan Python **64-bit**.
- **Tesseract OCR**: Pastikan terinstall di sistem jika ingin akurasi tambahan.

### 2. Setup Virtual Environment (Rekomendasi)
Masuk ke folder proyek, lalu buat dan aktifkan environment:
```powershell
# Buat environment (jika belum ada)
python -m venv .venv

# Aktifkan di Windows (PowerShell)
.\.venv\Scripts\Activate.ps1

# Aktifkan di Windows (CMD)
.venv\Scripts\activate
```

### 3. Install Library
Setelah environment aktif (ada tanda `(.venv)` di terminal), jalankan:
```bash
pip install -r requirements.txt
```

---

## 🚀 Cara Penggunaan

1. **Siapkan Data**: Masukkan folder hasil screenshot MRTG lu ke dalam folder `MRTG-Data/`.
   - Struktur: `MRTG-Data/YYYYMMDD/*.png`
2. **Siapkan Template**: Pastikan file Excel template lu sudah ada di folder utama.
3. **Konfigurasi Mapping**: Edit file `list_mrtg_data_position.txt` untuk mengatur posisi sel Excel tiap SID.
4. **Jalankan Bot**:
   ```bash
   python mrtg_data_to_monthly_report.py
   ```
5. **Pilih Mode**: Masukkan `1` untuk OCR atau `2` untuk Image Only.

---

## 📁 Struktur Proyek
- `MRTG-Data/`: Sumber gambar MRTG (organisir per folder tanggal).
- `mrtg_data_to_monthly_report.py`: Script utama (Pusat Kendali).
- `list_mrtg_data.txt`: Daftar SID/Graph-title yang ingin diproses.
- `list_mrtg_data_position.txt`: Mapping koordinat sel Excel untuk mode OCR.
- `ocr_report.log`: File audit untuk pengecekan error.

---

## ❓ Troubleshooting

| Masalah | Solusi |
| :--- | :--- |
| `No module named 'paddleocr'` | Pastikan lu sudah jalankan perintah di bagian **Langkah 2 & 3** (Aktifkan .venv dulu!). |
| `Failed to load PaddlePaddle` | Cek apakah Python lu 64-bit. Paddle tidak support Python 32-bit. |
| Tampilan progress bar berantakan | Pastikan terminal lu mendukung ANSI color (Gunakan Windows Terminal atau CMD modern). |

---
**💡 Tips**: Jika menggunakan mode OCR, pastikan gambar memiliki kualitas yang baik untuk akurasi data yang maksimal.
