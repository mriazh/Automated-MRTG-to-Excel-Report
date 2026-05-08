# Automated MRTG to Excel Report

Alat otomatisasi untuk memindahkan data grafik MRTG ke laporan bulanan Excel menggunakan **PaddleOCR (Deep Learning)**. Alat ini mengekstrak nilai *Current*, *Average*, dan *Maximum* untuk *Inbound* dan *Outbound* langsung dari gambar grafik.

## 🚀 Fitur Utama
- **OCR Canggih**: Menggunakan PaddleOCR yang lebih akurat dibanding Tesseract konvensional.
- **Fuzzy Keyword Matching**: Tetap akurat meskipun OCR salah baca sedikit (misal: `Inbound` kebaca `Inhound`).
- **Review List**: Menampilkan daftar ID yang bermasalah (N/A) di akhir proses untuk audit cepat.
- **Auto-Image Insert**: Otomatis memasukkan dan menyesuaikan ukuran gambar grafik ke dalam sel Excel.
- **Progress Tracker**: Monitoring proses dengan persentase dan ringkasan per tanggal.

## 📋 Prasyarat
- **Python 3.8 - 3.12** (Disarankan Python 3.10+).
- **Microsoft Visual C++ Build Tools** (Wajib untuk kompilasi library PaddlePaddle di Windows).

## 🛠️ Instalasi

1. **Clone Repository**:
   ```bash
   git clone https://github.com/username/Automated-MRTG-to-Excel-Report.git
   cd Automated-MRTG-to-Excel-Report
   ```

2. **Buat Virtual Environment** (Disarankan):
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. **Instal Dependencies**:
   ```bash
   pip install paddlepaddle paddleocr openpyxl Pillow
   ```

## 📖 Cara Penggunaan

1. **Siapkan Data**:
   - Taruh folder-folder tanggal (format `YYYYMMDD`) di dalam folder `MRTG-Data/`.
   - Pastikan file `list_mrtg_data_position.txt` dan `list_mrtg_data.txt` sudah sesuai dengan koordinat Excel lu.

2. **Jalankan Script**:
   ```bash
   python mrtg_data_to_monthly_report.py
   ```

3. **Pilih Mode**:
   - Pilih `[1]` untuk mode OCR lengkap.
   - Pilih `[2]` jika hanya ingin memasukkan gambar tanpa ekstraksi data.

4. **Review**:
   - Setelah selesai, cek **Review List** di terminal untuk melihat SID mana saja yang memiliki nilai `N/A`.
   - Hasil akhir akan tersimpan di `MRTG-Monthly-Report.xlsx`.

## 🛠️ Troubleshooting
- **Gagal OCR**: Pastikan gambar di folder `MRTG-Data` tidak korup dan teks di legend grafik terbaca jelas.
- **Log Terlalu Ramai**: Script sudah otomatis membungkam log internal PaddleOCR, namun jika masih muncul, pastikan environment variable `GLOG_minloglevel=3` sudah aktif.

---
**Note**: Project ini bermigrasi dari Tesseract ke PaddleOCR untuk akurasi yang lebih tinggi pada layout grafik MRTG yang kompleks.
