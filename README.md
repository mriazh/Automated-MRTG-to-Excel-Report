# Automated MRTG to Excel Report 🚀

Alat bantu otomatis untuk mengekstrak data dari grafik MRTG (.png) dan memasukkannya ke dalam laporan Excel secara rapi. Mendukung ekstraksi data otomatis menggunakan OCR atau sekadar memasukkan gambar saja.

## ✨ Fitur Utama
- **Dual Mode**:
  - **[1] OCR Mode**: Ekstraksi data bandwidth (Inbound/Outbound) otomatis dari grafik menggunakan PaddleOCR dan memasukkannya ke sel Excel yang ditentukan.
  - **[2] Image Only**: Memasukkan gambar grafik MRTG ke dalam Excel tanpa proses ekstraksi teks (lebih cepat).
- **Sticky Progress UI**: Antarmuka terminal yang bersih dengan bar progres yang selalu nempel di bawah, memberikan statistik real-time (OK, Partial, Fail).
- **Auto-Resize**: Gambar grafik secara otomatis diubah ukurannya agar pas dengan area (cell range) di Excel.
- **Silent Logging**: Log teknis yang berisik dari library OCR disembunyikan, membuat tampilan terminal tetap fokus pada progres.
- **Detailed Error Tracking**: Item yang gagal atau perlu review dicatat dalam `ocr_report.log`.

## 🛠️ Persyaratan
- Python 3.8+
- Library yang dibutuhkan (instal lewat `requirements.txt`):
  ```bash
  pip install -r requirements.txt
  ```

## 🚀 Cara Penggunaan
1. Letakkan folder data MRTG lu di folder `MRTG-Data`.
2. Pastikan file mapping (`list_mrtg_data_position.txt`) dan daftar SID (`list_mrtg_data.txt`) sudah sesuai.
3. Jalankan script utama:
   ```bash
   python mrtg_data_to_monthly_report.py
   ```
4. Pilih mode yang diinginkan (1 atau 2) dan biarkan script bekerja.

## 📁 Struktur Folder
- `MRTG-Data/`: Folder berisi subfolder tanggal (YYYYMMDD) yang berisi file .png.
- `list_mrtg_data.txt`: Daftar SID atau judul grafik yang ingin diproses.
- `list_mrtg_data_position.txt`: Mapping posisi sel Excel untuk tiap SID (Inbound/Outbound/Image).
- `ocr_report.log`: Catatan detail jika ada kegagalan ekstraksi data.

## ⚖️ Lisensi
Project ini menggunakan lisensi MIT. Silakan gunakan dan modifikasi sesuka hati!
