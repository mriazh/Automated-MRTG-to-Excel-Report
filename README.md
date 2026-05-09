# 🚀 MRTG to Excel Report Automator (PRO VERSION)

Alat bantu otomatisasi tingkat tinggi untuk mengekstrak data bandwidth dari grafik MRTG (.png) dan menyusunnya menjadi laporan bulanan Excel yang profesional. Menggunakan teknologi **PaddleOCR** untuk akurasi pembacaan data yang maksimal.

---

## ✨ Fitur Unggulan

- **🧠 Intelligent OCR Extraction**: Menggunakan engine **PaddleOCR** untuk membaca nilai *Current*, *Average*, dan *Maximum* (Inbound/Outbound) secara otomatis dari gambar grafik.
- **🖼️ Smart Image Injection**: Secara otomatis memasukkan gambar grafik ke dalam sel Excel yang ditentukan dengan skala yang pas (Anti-Pecah).
- **🕹️ Dual Engine Mode**:
  - **OCR Mode**: Ekstrak data teks + Sisipkan gambar (Lengkap).
  - **Image Only Mode**: Hanya menyisipkan gambar massal (Super Cepat).
- **📊 Real-Time Dashboard**: Tampilan terminal modern dengan Progress Bar interaktif dan statistik instan (OK, Partial, Fail).
- **🛡️ Silent Engine Protocol**: Semua log berisik dari library C++/Paddle disembunyikan agar terminal Anda tetap bersih dan fokus pada progres.
- **📝 Automated Audit Trail**: Setiap kegagalan pembacaan atau gambar yang hilang dicatat secara rapi di `ocr_report.log` untuk pengecekan ulang.

---

## 🛠️ Persiapan & Instalasi (Panduan Awam)

Ikuti langkah-langkah ini agar bot berjalan mulus tanpa error:

### 1. Prasyarat Sistem
- **Python 3.8 - 3.12**: [Download di sini](https://www.python.org/downloads/windows/) (Wajib versi **64-bit**).
- **Checklist**: Pastikan centang "Add Python to PATH" saat instalasi.

### 2. Setup Environment (Terisolasi)
Agar tidak bentrok dengan aplikasi lain, jalankan perintah ini di terminal:

**Jika punya satu versi Python:**
```powershell
python -m venv .venv
```

**Jika punya banyak versi Python (PENTING!):**
Gunakan perintah `py` untuk memilih versi spesifik (misal mau pake 3.12):
```powershell
# Pastikan folder .venv lama dihapus dulu jika salah versi
py -3.12 -m venv .venv
```

**Aktifkan Environment:**
```powershell
# Jika di PowerShell:
.\.venv\Scripts\Activate.ps1

# Jika di CMD:
.venv\Scripts\activate
```
*(Setelah aktif, ketik `python --version` buat mastiin udah bener).*

### 3. Install "Senjata" Bot
Setelah muncul tanda `(.venv)` di terminal, jalankan:
```bash
pip install -r requirements.txt
```

---

## 🚀 Cara Menjalankan

### Langkah 1: Persiapan Data
Masukkan folder hasil screenshot MRTG Anda ke folder `MRTG-Data/`.
- Struktur folder: `MRTG-Data/YYYYMMDD/gambar_mrtg.png`

### Langkah 2: Konfigurasi Target
- **Daftar SID**: Masukkan daftar SID di file `list_mrtg_data.txt`.
- **Mapping Lokasi**: Atur koordinat sel Excel (misal: B10, C12) di file `list_mrtg_data_position.txt`.

### Langkah 3: Eksekusi
Jalankan perintah berikut:
```bash
python mrtg_data_to_monthly_report.py
```
Pilih **Mode 1** untuk laporan lengkap atau **Mode 2** untuk kecepatan tinggi.

---

## 📁 Struktur Folder Proyek
```text
.
├── MRTG-Data/                # Sumber gambar (Per tanggal YYYYMMDD)
├── mrtg_data_to_monthly_report.py  # Mesin utama pengolah data
├── list_mrtg_data.txt        # Input daftar SID / Judul Grafik
├── list_mrtg_data_position.txt # Koordinat penempatan data di Excel
├── ocr_report.log            # File catatan (Audit Trail)
└── .venv/                    # Lingkungan kerja terisolasi
```

---

## ⚠️ Troubleshooting (FAQ)

- **Q: Muncul error 'No module named paddleocr'?**
  - **A**: Pastikan Anda sudah masuk ke virtual environment (`.venv`) sebelum menjalankan script.
- **Q: Versi Python tetap versi terbaru (misal 3.14), padahal butuh 3.12?**
  - **A**: Hapus folder `.venv` Anda, lalu buat ulang menggunakan perintah: `py -3.12 -m venv .venv`.
- **Q: Gambar di Excel tidak muncul?**
  - **A**: Cek folder `MRTG-Data`, pastikan nama file gambar sesuai dengan ID yang didaftarkan.
- **Q: Hasil OCR kurang akurat?**
  - **A**: Pastikan resolusi screenshot MRTG sudah standar dan tidak ada teks yang tertutup overlay.

---
**Dibuat dengan ❤️ untuk efisiensi pelaporan tim Telkom-GMF.**
