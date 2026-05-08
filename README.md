# Automated MRTG to Excel Report

Alat otomatisasi untuk memindahkan data grafik MRTG ke laporan bulanan Excel menggunakan **PaddleOCR (Deep Learning)**. 

## 📋 Prasyarat (PENTING!)
- **Python 3.12** (Sangat disarankan).
  - **Catatan**: Python 3.13 **tidak didukung** oleh PaddlePaddle. Jika Anda terlanjur menginstal 3.13, pastikan Anda menginstal 3.12 secara berdampingan.
- **Microsoft Visual C++ Build Tools** (Wajib untuk Windows).

## 🛠️ Instalasi (Langkah Demi Langkah)

Jika Anda sebelumnya gagal instalasi, **HAPUS** dulu folder `.venv` yang lama!

1. **Hapus Virtual Environment lama (Jika ada)**:
   ```bash
   rmdir /s /q .venv
   ```

2. **Buat Virtual Environment baru (Python 3.12)**:
   ```bash
   # Pastikan menggunakan versi 3.12
   py -3.12 -m venv .venv
   ```

3. **Aktifkan Virtual Environment**:
   ```bash
   .venv\Scripts\activate
   ```

4. **Instal Dependencies**:
   ```bash
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```
   *Jika masih ada kendala pada `paddlepaddle`, gunakan index resmi:*
   ```bash
   pip install paddlepaddle -i https://www.paddlepaddle.org.cn/packages/stable/cpu/
   ```

## 📖 Cara Penggunaan
1. Taruh folder data di `MRTG-Data/`.
2. Jalankan: `python mrtg_data_to_monthly_report.py`
3. Masukkan `1` untuk OCR.
4. Cek hasil di `MRTG-Monthly-Report.xlsx`.

## 🛠️ Troubleshooting
- **Gagal Instal Paddle**: Pastikan folder `.venv` lu beneran pake Python 3.12. Cek dengan perintah `python --version` setelah aktivasi venv.
- **Visual C++ Error**: Instal [Build Tools for Visual Studio](https://visualstudio.microsoft.com/visual-cpp-build-tools/).
