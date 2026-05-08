# Automated MRTG to Excel Report

Alat otomatisasi untuk memindahkan data grafik MRTG ke laporan bulanan Excel menggunakan **PaddleOCR (Deep Learning)**. Alat ini mengekstrak nilai *Current*, *Average*, dan *Maximum* untuk *Inbound* dan *Outbound* langsung dari gambar grafik.

## 🚀 Fitur Utama
- **OCR Canggih**: Menggunakan PaddleOCR yang lebih akurat dibanding Tesseract konvensional.
- **Fuzzy Keyword Matching**: Tetap akurat meskipun OCR salah baca sedikit (misal: `Inbound` kebaca `Inhound`).
- **Review List**: Menampilkan daftar ID yang bermasalah (N/A) di akhir proses untuk audit cepat.
- **Auto-Image Insert**: Otomatis memasukkan dan menyesuaikan ukuran gambar grafik ke dalam sel Excel.
- **Progress Tracker**: Monitoring proses dengan persentase dan ringkasan per tanggal.

## 📋 Prasyarat (PENTING!)
- **Python 3.10, 3.11, atau 3.12** (Sangat disarankan).
  - **Catatan**: Python 3.13 saat ini **belum didukung** oleh library `paddlepaddle`. Jika Anda menggunakan 3.13, harap instal Python 3.12.
- **Microsoft Visual C++ Build Tools** (Wajib untuk Windows).

## 🛠️ Instalasi

1. **Clone Repository**:
   ```bash
   git clone https://github.com/username/Automated-MRTG-to-Excel-Report.git
   cd Automated-MRTG-to-Excel-Report
   ```

2. **Buat Virtual Environment** (Gunakan Python 3.12):
   ```bash
   # Pastikan python yang terpanggil adalah versi 3.12
   py -3.12 -m venv .venv
   .venv\Scripts\activate
   ```

3. **Instal Dependencies**:
   ```bash
   # Update pip dulu
   python -m pip install --upgrade pip
   
   # Instal library utama
   pip install paddlepaddle paddleocr openpyxl Pillow
   ```
   *Jika `paddlepaddle` tetap tidak ditemukan, coba gunakan link resmi:*
   ```bash
   pip install paddlepaddle -i https://www.paddlepaddle.org.cn/packages/stable/cpu/
   ```

## 📖 Cara Penggunaan

1. **Siapkan Data**: Taruh folder-folder tanggal di `MRTG-Data/`.
2. **Jalankan Script**: `python mrtg_data_to_monthly_report.py`
3. **Pilih Mode**: Masukkan `1` untuk OCR.
4. **Review**: Cek **Review List** di terminal setelah selesai untuk melihat data N/A.

## 🛠️ Troubleshooting
- **Error "Could not find a version"**: Ini fiks karena versi Python lu nggak cocok (biasanya karena pake 3.13). Downgrade ke 3.12.
- **Error "Microsoft Visual C++"**: Lu harus instal [Build Tools for Visual Studio](https://visualstudio.microsoft.com/visual-cpp-build-tools/).

---
**Note**: Project ini dioptimalkan untuk akurasi tinggi pada layout grafik MRTG.
