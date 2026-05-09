# 🚀 Automated MRTG to Excel Report (PRO Edition)

Alat bantu otomatisasi untuk mengolah hasil screenshot MRTG menjadi laporan Excel yang rapi. Bot ini bisa membaca angka bandwidth (Inbound/Outbound) secara otomatis menggunakan teknologi OCR atau melakukan *image-injection* massal ke dalam template Excel.

---

## 📖 Daftar Isi
1. [Prasyarat Sistem](#-prasyarat-sistem)
2. [Instalasi Cepat (Untuk Pemula)](#-instalasi-cepat-untuk-pemula)
3. [Cara Penggunaan (Step-by-Step)](#-cara-penggunaan-step-by-step)
4. [Penjelasan Mode](#-penjelasan-mode)
5. [Struktur Folder](#-struktur-folder)
6. [Troubleshooting (Tanya Jawab)](#-troubleshooting-tanya-jawab)

---

## 🛠️ Prasyarat Sistem

Sebelum mulai, pastikan komponen ini sudah ada di komputer lu:

1. **Python (Versi 3.8 - 3.12)**:
   - **PENTING**: Wajib install versi **64-bit**.
   - [Download Python di sini](https://www.python.org/downloads/windows/) (Ceklis "Add Python to PATH" saat install).

---

## ⚡ Instalasi Cepat (Untuk Pemula)

Ikuti urutan perintah ini di terminal (CMD/PowerShell) tepat di dalam folder proyek ini:

```powershell
# 1. Buat lingkungan kerja terisolasi (Virtual Environment)
python -m venv .venv

# 2. Aktifkan Environment
# Jika di PowerShell:
.\.venv\Scripts\Activate.ps1
# Jika di CMD:
.venv\Scripts\activate

# 3. Install semua 'senjata' yang dibutuhkan
pip install -r requirements.txt
```

---

## 🚀 Cara Penggunaan (Step-by-Step)

Gua kasih contoh alur kerja dari awal sampe dapet laporan:

### Langkah 1: Siapkan Gambar MRTG
Masukkan folder hasil screenshot lu ke folder `MRTG-Data`. 
- Contoh: `MRTG-Data/20260101/MRTG_123456.png`

### Langkah 2: Edit Daftar Target
Buka file `list_mrtg_data.txt`, isi dengan SID atau judul grafik yang mau lu proses.
- Contoh: `1. SID : 4700001-0021497479`

### Langkah 3: Atur Posisi Excel (Mapping)
Buka file `list_mrtg_data_position.txt`. Di sini lu atur koordinat sel Excel buat naruh angka dan gambar.
- Contoh: `Inbound_Current : B10` (Artinya angka Inbound Current bakal masuk ke sel B baris 10).

### Langkah 4: Jalankan Bot
Ketik perintah ini di terminal:
```bash
python mrtg_data_to_monthly_report.py
```
- Pilih **1** jika ingin bot membaca angka (OCR).
- Pilih **2** jika hanya ingin memasukkan gambar saja (Cepat).

---

## 📊 Penjelasan Mode

| Fitur | Mode [1] OCR | Mode [2] Image Only |
| :--- | :--- | :--- |
| **Fungsi Utama** | Isi angka + Masukin Gambar | Masukin Gambar Saja |
| **Kecepatan** | ~5-7 detik per gambar | <1 detik per gambar |
| **Akurasi** | Tergantung kualitas gambar | 100% (Hanya copy-paste) |
| **Output** | Angka Bandwidth & Grafik | Grafik Saja |

---

## 📁 Struktur Folder

- `MRTG-Data/`: Sumber gambar (harus per folder tanggal YYYYMMDD).
- `mrtg_data_to_monthly_report.py`: Script utama (Pusat Kendali).
- `list_mrtg_data.txt`: Daftar target yang mau diproses.
- `list_mrtg_data_position.txt`: Peta koordinat sel Excel.
- `ocr_report.log`: "Buku tamu" yang mencatat jika ada gambar yang gagal dibaca.

---

## ❓ Troubleshooting (Tanya Jawab)

**P: Kok muncul error `No module named 'paddleocr'`?**
J: Itu karena lu belum masuk ke environment. Jalankan `.venv\Scripts\activate` dulu sampe ada tulisan `(.venv)` di depan kursor terminal lu.

**P: Gambar di Excel kok "mencong" atau nggak pas?**
J: Tenang, bot ini udah pake *Nuclear Isolation*. Pastikan lu udah ngatur `Image : B10-F25` (Range sel) di file mapping dengan benar.

**P: Kenapa progress bar-nya berantakan?**
J: Gunakan **Windows Terminal** atau CMD versi terbaru agar warna dan bar-nya muncul dengan sempurna.

**P: Python versi 3.12 bisa?**
J: **BISA!** Bot ini sudah dites lancar di Python 3.12.10 64-bit.

---
**💡 Tips**: Jika hasil OCR kurang akurat, pastikan screenshot MRTG lu bersih dari overlay atau popup saat diambil.
