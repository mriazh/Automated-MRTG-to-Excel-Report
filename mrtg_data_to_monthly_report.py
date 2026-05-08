import os
import re
import io
import sys
import logging
import traceback

# Suppress noisy PaddleOCR / Paddle C++ logs
os.environ['GLOG_minloglevel'] = '3'
os.environ['FLAGS_minloglevel'] = '3'
os.environ['PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK'] = 'True'
os.environ['PADDLEX_DISABLE_PRINT'] = '1'
os.environ['PADDLE_LOG_LEVEL'] = 'ERROR'

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string, get_column_letter
from PIL import Image as PILImage

# ========== KONFIGURASI ==========
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_DATA = os.path.join(BASE_DIR, "MRTG-Data")
IMAGE_SCALE = 0.98

# --- Logging (DEBUG to File, INFO to Console) ---
import sys
logger = logging.getLogger('mrtg_report')
logger.setLevel(logging.DEBUG)

# File Handler (Simpan detail OCR di sini)
fh = logging.FileHandler('ocr_report.log', encoding='utf-8')
fh.setLevel(logging.DEBUG)
fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger.addHandler(fh)

# Console Handler (Tampilkan progres saja via STDOUT)
# Kita gunakan stdout karena stderr akan kita mute global untuk membungkam Paddle
ch = logging.StreamHandler(sys.stdout)
ch.setLevel(logging.INFO)
ch.setFormatter(logging.Formatter('%(message)s'))
logger.addHandler(ch)

# JURUS PAMUNGKAS: Mute Stderr (FD 2) secara global untuk hilangkan log C++
try:
    null_fd = os.open(os.devnull, os.O_RDWR)
    os.dup2(null_fd, 2)
except: pass

def cetak_progres_bar(current, total):
    """Mencetak progress bar ala pip di bagian bawah terminal."""
    bar_length = 40
    percent = (current / total)
    filled_length = int(bar_length * percent)
    
    # Bikin bar [===>    ]
    if filled_length > 0:
        bar = '=' * (filled_length - 1) + '>' + ' ' * (bar_length - filled_length)
    else:
        bar = ' ' * bar_length
        
    p_text = f"{int(percent * 100)}%"
    # ANSI: \033[K (Clear line)
    sys.stdout.write(f"\r\033[K{{{p_text:<4} [{bar}]}}")
    sys.stdout.flush()

# --- Mode OCR ---
OCR_TEMPLATE = os.path.join(BASE_DIR, "MRTG-Monthly-Report-on-Internet-Bandwidth-Utilization-by-Telkom.xlsx")
OCR_OUTPUT   = os.path.join(BASE_DIR, "MRTG-Monthly-Report.xlsx")
OCR_MAPPING  = os.path.join(BASE_DIR, "list_mrtg_data_position.txt")
OCR_DAFTAR   = os.path.join(BASE_DIR, "list_mrtg_data.txt")

# --- Mode Image Only ---
IMG_TEMPLATE = os.path.join(BASE_DIR, "MRTG-Monthly-Report-on-Internet-Bandwidth-Utilization-by-Telkom (Img only).xlsx")
IMG_OUTPUT   = os.path.join(BASE_DIR, "MRTG-Monthly-Report-image-only.xlsx")
IMG_MAPPING  = os.path.join(BASE_DIR, "list_mrtg_data_position_img_only.txt")
IMG_DAFTAR   = os.path.join(BASE_DIR, "list_mrtg_data_img_only.txt")


# ========================================================
#  SHARED FUNCTIONS (dipakai kedua mode)
# ========================================================

def baca_daftar(filepath):
    """Baca daftar SID / Graph-title dari file teks."""
    items = []
    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split('.', 1)
            if len(parts) != 2:
                continue
            nomor = parts[0].strip()
            rest = parts[1].strip()
            if rest.startswith('SID : '):
                tipe = 'SID'
                id_val = rest.replace('SID : ', '').strip()
            elif rest.startswith('Graph-title : '):
                tipe = 'Graph-title'
                id_val = rest.replace('Graph-title : ', '').strip()
            else:
                continue
            items.append((nomor, tipe, id_val))
    return items


def get_area_size_pixels(sheet, start_row, start_col, end_row, end_col):
    """Hitung ukuran area Excel dalam pixel berdasarkan lebar kolom dan tinggi baris."""
    total_width = 0
    for col in range(start_col, end_col + 1):
        col_letter = get_column_letter(col)
        col_width = sheet.column_dimensions[col_letter].width
        if col_width is None:
            col_width = 8.43  # default Excel
        total_width += col_width * 7.4  # konversi ke pixel (estimasi)

    total_height = 0
    for row in range(start_row, end_row + 1):
        row_height = sheet.row_dimensions[row].height
        if row_height is None:
            row_height = 15  # default Excel dalam point
        total_height += row_height * 1.333  # konversi ke pixel

    return total_width, total_height


def resize_image_stretch(image_path, target_width, target_height):
    """Resize gambar ke ukuran target (stretch)."""
    with PILImage.open(image_path) as img:
        if img.mode in ('RGBA', 'LA', 'P'):
            img = img.convert('RGB')
        img_resized = img.resize((int(target_width), int(target_height)), PILImage.Resampling.LANCZOS)
        output = io.BytesIO()
        img_resized.save(output, format='PNG')
        output.seek(0)
        return output


def cari_path_gambar(folder_data, tanggal_str, tipe, id_val):
    """Cari path gambar MRTG berdasarkan tipe (SID/Graph-title) dan ID."""
    if tipe == 'SID':
        nama_file = f"MRTG_{id_val}.png"
    else:
        nama_file = f"MRTG_{id_val}_{tanggal_str}.png"

    path_gambar = os.path.join(folder_data, tanggal_str, nama_file)
    if not os.path.exists(path_gambar):
        # Fallback: cari file dengan awalan MRTG_{id_val}
        folder_tgl = os.path.join(folder_data, tanggal_str)
        if os.path.exists(folder_tgl):
            for f in os.listdir(folder_tgl):
                if f.startswith(f"MRTG_{id_val}") and f.endswith(".png"):
                    path_gambar = os.path.join(folder_tgl, f)
                    break
    return path_gambar


def tambah_gambar_di_area(sheet, image_path, start_row, start_col, end_row, end_col, scale=IMAGE_SCALE):
    """Resize gambar sesuai area dan letakkan di sheet Excel."""
    try:
        width_px, height_px = get_area_size_pixels(sheet, start_row, start_col, end_row, end_col)
        width_px = width_px * scale
        height_px = height_px * scale
        img_bytes = resize_image_stretch(image_path, width_px, height_px)
        img = XLImage(img_bytes)
        img.anchor = f"{get_column_letter(start_col)}{start_row}"
        sheet.add_image(img)
        return True
    except Exception as e:
        print(f"    Gagal tambah gambar: {e}")
        return False


def get_tanggal_list(folder_data):
    """Ambil daftar folder tanggal (format YYYYMMDD) dari folder data."""
    if not os.path.exists(folder_data):
        return []
    tanggal_list = [
        d for d in os.listdir(folder_data)
        if os.path.isdir(os.path.join(folder_data, d)) and d.isdigit() and len(d) == 8
    ]
    tanggal_list.sort()
    return tanggal_list


# ========================================================
#  MODE 1: OCR (Ekstrak data + insert gambar)
# ========================================================

def baca_mapping_ocr(filepath):
    """Baca file mapping OCR (format Service Id dengan 6 field Inbound/Outbound + Image)."""
    mapping = {}
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip()]
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith('Service Id :') or line.startswith('Service Id :'):
            id_raw = line.replace('Service Id :', '').replace('Service Id :', '').strip()
            id_clean = re.sub(r'^\(\d+\)\s*', '', id_raw)
            i += 1
            entry = {}
            for _ in range(6):
                if i >= len(lines):
                    break
                key_val = lines[i].split(':', 1)
                if len(key_val) == 2:
                    key = key_val[0].strip()
                    val = key_val[1].strip()
                    match = re.match(r'([A-Z]+)(\d+)', val)
                    if match:
                        col_letter = match.group(1)
                        row = int(match.group(2))
                        col = column_index_from_string(col_letter)
                        entry[key] = (row, col)
                i += 1
            if i < len(lines) and lines[i].startswith('Image :'):
                range_str = lines[i].replace('Image :', '').strip()
                start, end = range_str.split('-')
                start_col = column_index_from_string(re.match(r'[A-Z]+', start).group())
                start_row = int(re.search(r'\d+', start).group())
                end_col = column_index_from_string(re.match(r'[A-Z]+', end).group())
                end_row = int(re.search(r'\d+', end).group())
                entry['Image'] = ((start_row, start_col), (end_row, end_col))
                i += 1
            mapping[id_clean] = entry
        else:
            i += 1
    return mapping


def _get_ocr_engine():
    """Singleton: inisialisasi PaddleOCR sekali saja."""
    if not hasattr(_get_ocr_engine, '_engine'):
        import warnings
        import contextlib
        
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            from paddleocr import PaddleOCR
            
            # Mute stdout sementara hanya saat inisialisasi model
            with open(os.devnull, 'w') as fnull:
                with contextlib.redirect_stdout(fnull):
                    _get_ocr_engine._engine = PaddleOCR(lang='en')
            
        logger.debug("PaddleOCR engine siap.")
    return _get_ocr_engine._engine


def extract_mrtg_values(image_path):
    """Ekstrak nilai bandwidth dengan strategi Sequential Context & Fuzzy Keywords."""
    try:
        ocr = _get_ocr_engine()
        # Log ke file saja (DEBUG)
        logger.debug(f"OCR memproses: {os.path.basename(image_path)}")

        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            result_iter = ocr.predict(image_path)

            all_texts = []
            for res in result_iter:
                data = res.get('res', res) if hasattr(res, 'get') else {}
                if data:
                    texts = data.get('rec_texts', [])
                    all_texts.extend([str(t).strip() for t in texts])

        # --- DEBUG LOG: START ---
        sid_log = os.path.basename(image_path).replace('MRTG_', '').replace('.png', '')
        logger.debug(f"--- START OCR [{sid_log}] ---")
        logger.debug(f"Raw Text Detected: {all_texts}")

        # --- Helper: Cari nilai setelah keyword ---
        def find_value_after(keyword_list, texts, start_search_idx):
            """Mencari angka (+ unit) setelah salah satu keyword ditemukan."""
            for i in range(start_search_idx, len(texts)):
                t = texts[i].lower()
                # Cocokkan keyword dengan toleransi typo
                if any(kw.lower() in t for kw in keyword_list):
                    # Cari angka di i atau 3 box setelahnya
                    for j in range(i, min(i + 6, len(texts))):
                        # Pattern baru: Case-insensitive unit, dukung b/s, bps, dsb.
                        match = re.search(r'(\d+(?:[\.,]\d+)?)\s*([MkGTmkgt]?[Bb]?p?s?)', texts[j])
                        if match:
                            val = match.group(1).replace(',', '.')
                            unit_raw = match.group(2).strip()
                            
                            # Normalisasi Unit
                            unit = ""
                            u_low = unit_raw.lower()
                            if 't' in u_low: unit = "T"
                            elif 'g' in u_low: unit = "G"
                            elif 'm' in u_low: unit = "M"
                            elif 'k' in u_low: unit = "k"
                            
                            # Jika unit tidak nempel, cari di 2 box setelahnya
                            # Ini handle kasus PaddleOCR split nilai: '14.91' + 'M'
                            if not unit:
                                for k in range(j + 1, min(j + 3, len(texts))):
                                    next_t = texts[k].strip().lower()
                                    # Pastikan box berikutnya HANYA unit (pendek, tidak ada angka)
                                    if re.match(r'^[tmgkTMGK]$', next_t.strip()):
                                        if 't' in next_t: unit = "T"
                                        elif 'g' in next_t: unit = "G"
                                        elif 'm' in next_t: unit = "M"
                                        elif 'k' in next_t: unit = "k"
                                        break
                                    # Atau unit yang nempel tapi di box sendiri: 'M', 'k', dll
                                    elif next_t in ['m', 'k', 'g', 't']:
                                        if next_t == 't': unit = "T"
                                        elif next_t == 'g': unit = "G"
                                        elif next_t == 'm': unit = "M"
                                        elif next_t == 'k': unit = "k"
                                        break
                                    # Jika box berikutnya sudah berisi keyword lain, stop
                                    elif any(kw in next_t for kw in ['current', 'average', 'maximum', 'inbound', 'outbound', 'cur rent']):
                                        break
                            
                            # Default ke M jika masih kosong
                            if not unit:
                                unit = "M"
                                
                            return f"{val} {unit}".strip()
                        
                        if 'n/a' in texts[j].lower():
                            continue
            return "N/A"

        # --- Cari Section Inbound & Outbound dengan Fuzzy Matching ---
        inbound_idx = -1
        outbound_idx = -1
        
        # Keyword list yang lebih luas buat handle typo OCR
        in_kws = ['inbound', 'in-bound', 'in bound', 'inhound', '1nbound', 'nbound', 'inb']
        out_kws = ['outbound', 'out-bound', 'out bound', 'oulbound', '0utbound', 'outb']

        for i, t in enumerate(all_texts):
            t_low = t.lower()
            if any(kw in t_low for kw in in_kws):
                inbound_idx = i
            if any(kw in t_low for kw in out_kws):
                outbound_idx = i

        # Proteksi: Jika Outbound tidak ketemu di legend, cari penanda Current kedua
        if outbound_idx == -1 and inbound_idx != -1:
            for i in range(inbound_idx + 1, len(all_texts)):
                if any(kw in all_texts[i].lower() for kw in ['current', 'cur rent', 'cur ren', 'curren']):
                    # Pastikan ini Current untuk Outbound (setelah ada angka Inbound)
                    has_numbers_before = False
                    for j in range(inbound_idx + 1, i):
                        if re.search(r'\d', all_texts[j]):
                            has_numbers_before = True
                            break
                    if has_numbers_before:
                        outbound_idx = i - 1
                        break

        # Parse Inbound Area
        if inbound_idx != -1:
            search_limit = outbound_idx if outbound_idx > inbound_idx else len(all_texts)
            in_area = all_texts[inbound_idx:search_limit]
            in_vals = {
                'Current': find_value_after(['Current', 'Curren', 'Cur rent', 'Cur ren', 'Cur'], in_area, 0),
                'Average': find_value_after(['Average', 'Averaqe', 'Avera9e', 'Avera', 'Ave'], in_area, 0),
                'Maximum': find_value_after(['Maximum', 'Maximu', 'Maxlmu', 'Max'], in_area, 0)
            }
        else:
            in_vals = {'Current': 'N/A', 'Average': 'N/A', 'Maximum': 'N/A'}

        # Parse Outbound Area
        if outbound_idx != -1:
            out_area = all_texts[outbound_idx:]
            out_vals = {
                'Current': find_value_after(['Current', 'Curren', 'Cur rent', 'Cur ren', 'Cur'], out_area, 0),
                'Average': find_value_after(['Average', 'Averaqe', 'Avera9e', 'Avera', 'Ave'], out_area, 0),
                'Maximum': find_value_after(['Maximum', 'Maximu', 'Maxlmu', 'Max'], out_area, 0)
            }
        else:
            out_vals = {'Current': 'N/A', 'Average': 'N/A', 'Maximum': 'N/A'}

        result = {
            'Inbound_Current': in_vals['Current'],
            'Inbound_Average': in_vals['Average'],
            'Inbound_Maximum': in_vals['Maximum'],
            'Outbound_Current': out_vals['Current'],
            'Outbound_Average': out_vals['Average'],
            'Outbound_Maximum': out_vals['Maximum']
        }

        # --- DEBUG LOG: END ---
        logger.debug(f"Extracted Values: {result}")
        logger.debug(f"--- END OCR [{sid_log}] ---\n")

        return result
    except Exception as e:
        logger.error(f"OCR Critical Error [{os.path.basename(image_path)}]: {e}")
        return None


def tulis_nilai(sheet, entry, values):
    """Tulis nilai OCR ke sel Excel yang ditentukan."""
    for key, (row, col) in entry.items():
        if key in values and key != 'Image':
            sheet.cell(row=row, column=col, value=values[key])


def proses_tanggal_ocr(wb, tanggal_str, items, mapping, global_stats, review_list, global_context):
    """Proses satu folder tanggal untuk mode OCR (ekstrak data + insert gambar)."""
    hari = int(tanggal_str[6:8])
    sheet_name_candidates = [str(hari), f"{hari:02d}"]
    sheet = None
    for name in sheet_name_candidates:
        if name in wb.sheetnames:
            sheet = wb[name]
            break
    if sheet is None:
        sheet = wb.create_sheet(title=str(hari))
        print(f"  Sheet {str(hari)} tidak ditemukan, membuat baru.")
    else:
        print(f"  Menggunakan sheet {sheet.title} untuk tanggal {tanggal_str}")

    total = len(items)
    ok_count = 0
    na_count = 0
    fail_count = 0

    for idx, (nomor, tipe, id_val) in enumerate(items, 1):
        # Update Global Counter
        global_context['current'] += 1
        curr = global_context['current']
        total_global = global_context['total']
        
        # Tampilkan progres bar di awal sebelum proses
        cetak_progres_bar(curr-1, total_global)
        
        # Terminal format: [01/48]
        prefix = f"[{curr:02d}/{total_global:02d}]"
        
        id_clean = re.sub(r'^\(\d+\)\s*', '', id_val)
        label = f"{id_val}"

        if id_clean not in mapping:
            # Hapus baris progres bar sementara
            sys.stdout.write("\r\033[K")
            print(f"  {prefix} ⏭️  {label} (Skip: No Mapping)")
            cetak_progres_bar(curr, total_global)
            continue
        entry = mapping[id_clean]

        path_gambar = cari_path_gambar(FOLDER_DATA, tanggal_str, tipe, id_val)
        if not os.path.exists(path_gambar):
            # Hapus baris progres bar sementara
            sys.stdout.write("\r\033[K")
            print(f"  {prefix} ❌ {label} (Missing Image)")
            fail_count += 1
            global_stats['fail'] += 1
            review_list.append({'sid': id_clean, 'date': tanggal_str, 'sheet': sheet.title, 'status': 'Fail', 'na': 6})
            cetak_progres_bar(curr, total_global)
            continue

        # Hapus baris progres bar sementara (jika ada) untuk cetak status pencarian
        sys.stdout.write("\r\033[K")
        print(f"  {prefix} 🔍 {label} (Processing...)", end='', flush=True)
        
        values = extract_mrtg_values(path_gambar)
        
        # Hapus teks "Processing..." untuk diganti hasil akhir
        sys.stdout.write("\r\033[K")
        
        if not values:
            print(f"  {prefix} ❌ {label} (OCR Fail)")
            fail_count += 1
            global_stats['fail'] += 1
            review_list.append({'sid': id_clean, 'date': tanggal_str, 'sheet': sheet.title, 'status': 'Fail', 'na': 6})
        else:
            # Hitung N/A
            val_na = sum(1 for v in values.values() if v == 'N/A')
            if val_na == 0:
                print(f"  {prefix} ✅ {label} (Success)")
                ok_count += 1
                global_stats['ok'] += 1
            else:
                print(f"  {prefix} ⚠️  {label} ({6-val_na}/6 OK)")
                na_count += 1
                global_stats['partial'] += 1
                review_list.append({'sid': id_clean, 'date': tanggal_str, 'sheet': sheet.title, 'status': 'Partial', 'na': val_na})

            tulis_nilai(sheet, entry, values)
            if 'Image' in entry:
                (start_row, start_col), (end_row, end_col) = entry['Image']
                tambah_gambar_di_area(sheet, path_gambar, start_row, start_col, end_row, end_col)

        # Munculkan lagi progres bar yang terupdate
        cetak_progres_bar(curr, total_global)

    global_stats['total'] += total
    print(f"\n  📊 Ringkasan tanggal {tanggal_str}: ✅ {ok_count} OK, ⚠️ {na_count} partial, ❌ {fail_count} gagal (dari {total} item)")


# ========================================================
#  MODE 2: IMAGE ONLY (insert gambar saja, tanpa OCR)
# ========================================================

def baca_mapping_img(filepath):
    """Baca file mapping Image-only (format SID -> range)."""
    mapping = {}
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip()]
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith('SID : '):
            id_raw = line.replace('SID : ', '').strip()
            id_clean = re.sub(r'^\(\d+\)\s*', '', id_raw)
            i += 1
            if i < len(lines) and lines[i].startswith('->'):
                range_str = lines[i][2:].strip()
                start, end = range_str.split('-')
                start_col = column_index_from_string(re.match(r'[A-Z]+', start).group())
                start_row = int(re.search(r'\d+', start).group())
                end_col = column_index_from_string(re.match(r'[A-Z]+', end).group())
                end_row = int(re.search(r'\d+', end).group())
                mapping[id_clean] = ((start_row, start_col), (end_row, end_col))
                i += 1
            else:
                i += 1
        else:
            i += 1
    return mapping


def proses_tanggal_img(wb, tanggal_str, items, mapping):
    """Proses satu folder tanggal untuk mode Image-only (insert gambar saja)."""
    hari = int(tanggal_str[6:8])
    sheet_name = f"{hari:02d}"
    if sheet_name not in wb.sheetnames:
        # Fallback: coba tanpa leading zero
        sheet_name = str(hari)
        if sheet_name not in wb.sheetnames:
            print(f"  Sheet {hari:02d} tidak ditemukan")
            return
    sheet = wb[sheet_name]
    print(f"  Memproses sheet {sheet_name}...")

    for nomor, tipe, id_val in items:
        id_clean = re.sub(r'^\(\d+\)\s*', '', id_val)
        if id_clean not in mapping:
            print(f"    Peringatan: ID '{id_clean}' tidak ditemukan di mapping")
            continue
        (start_row, start_col), (end_row, end_col) = mapping[id_clean]

        path_gambar = cari_path_gambar(FOLDER_DATA, tanggal_str, tipe, id_val)
        if not os.path.exists(path_gambar):
            print(f"    Gambar tidak ditemukan: {path_gambar}")
            continue

        print(f"    Menambahkan gambar untuk {tipe} {id_val} di area {start_row},{start_col} - {end_row},{end_col}")
        tambah_gambar_di_area(sheet, path_gambar, start_row, start_col, end_row, end_col)


# ========================================================
#  MAIN
# ========================================================

def main():
    print("=" * 60)
    print("  AUTOMATED MRTG TO EXCEL REPORT")
    print("=" * 60)
    print("  Pilih mode:")
    print("  [1] OCR Mode   : Ekstrak data + insert gambar ke Excel")
    print("  [2] Image Only : Insert gambar saja ke Excel (tanpa OCR)")
    print("=" * 60)

    while True:
        pilihan = input("  >> Masukkan pilihan (1/2): ").strip()
        if pilihan in ('1', '2'):
            break
        print("  Input tidak valid. Masukkan 1 atau 2.")

    if pilihan == '1':
        # === MODE OCR ===
        print("\n>> Mode: OCR (Ekstrak data + insert gambar)\n")
        template_file = OCR_TEMPLATE
        output_file   = OCR_OUTPUT
        mapping_file  = OCR_MAPPING
        daftar_file   = OCR_DAFTAR

        if not os.path.exists(mapping_file):
            print(f"File mapping '{mapping_file}' tidak ditemukan!")
            return
        mapping = baca_mapping_ocr(mapping_file)
        print(f"Mapping berisi {len(mapping)} entri.")

        items = baca_daftar(daftar_file)
        print(f"Daftar berisi {len(items)} item.")

        tanggal_list = get_tanggal_list(FOLDER_DATA)
        if not tanggal_list:
            print(f"Folder '{FOLDER_DATA}' tidak ditemukan atau kosong!")
            return
        print(f"Ditemukan {len(tanggal_list)} folder tanggal.")

        if not os.path.exists(template_file):
            print(f"Template '{template_file}' tidak ditemukan!")
            return
        wb = load_workbook(template_file)
        print("Template berhasil dimuat.\n")

        # Global Stats
        global_stats = {'ok': 0, 'partial': 0, 'fail': 0, 'total': 0}
        review_list = []

        # Global Context for Progress tracking
        global_context = {
            'current': 0,
            'total': len(tanggal_list) * len(items)
        }

        for tgl_idx, tgl in enumerate(tanggal_list, 1):
            print(f"\nMemproses tanggal: {tgl} ({tgl_idx}/{len(tanggal_list)})")
            proses_tanggal_ocr(wb, tgl, items, mapping, global_stats, review_list, global_context)

        wb.save(output_file)
        
        # --- FINAL SUMMARY ---
        print("\n" + "="*60)
        print("  FINAL REPORT SUMMARY")
        print("="*60)
        
        total = global_stats['total']
        if total > 0:
            ok_p = (global_stats['ok'] / total * 100)
            partial_p = (global_stats['partial'] / total * 100)
            fail_p = (global_stats['fail'] / total * 100)
            
            print(f"  ✅ BERHASIL (100%) : {global_stats['ok']} items ({ok_p:.1f}%)")
            print(f"  ⚠️  PARTIAL (N/A)   : {global_stats['partial']} items ({partial_p:.1f}%)")
            print(f"  ❌ GAGAL (No Data)  : {global_stats['fail']} items ({fail_p:.1f}%)")
            print(f"  ----------------------------------------------------------")
            print(f"  TOTAL ITEM PROSES : {total} items")
        else:
            print("  Tidak ada item yang diproses.")
        print("="*60)

        if review_list:
            print("\n🔍 LIST ITEM YANG PERLU DICEK (GROUPED BY SID):")
            print("-" * 80)
            
            # Grouping by SID
            grouped_issues = {}
            for item in review_list:
                sid = item['sid']
                if sid not in grouped_issues:
                    grouped_issues[sid] = {
                        'dates': [],
                        'sheets': [],
                        'total_na': 0,
                        'partial_count': 0,
                        'fail_count': 0
                    }
                
                grouped_issues[sid]['dates'].append(item['date'])
                grouped_issues[sid]['sheets'].append(item['sheet'])
                grouped_issues[sid]['total_na'] += item['na']
                if item['status'] == 'Partial':
                    grouped_issues[sid]['partial_count'] += 1
                else:
                    grouped_issues[sid]['fail_count'] += 1

            for sid, data in grouped_issues.items():
                dates_str = ", ".join(sorted(list(set(data['dates']))))
                sheets_str = ", ".join(sorted(list(set(data['sheets'])), key=lambda x: int(x) if x.isdigit() else 0))
                
                status_parts = []
                if data['partial_count'] > 0:
                    status_parts.append(f"{data['partial_count']} partial")
                if data['fail_count'] > 0:
                    status_parts.append(f"{data['fail_count']} fail")
                status_str = " & ".join(status_parts)
                
                print(f"SID     : {sid}")
                print(f"Tanggal : {dates_str}")
                print(f"Sheet   : {sheets_str}")
                print(f"Status  : {status_str}, {data['total_na']} total nilai N/A")
                print("-" * 80)
            
            print(f"Total {len(grouped_issues)} SID unik butuh review.")
            print("💡 Tips: Cek 'ocr_report.log' untuk melihat detail teks yang terdeteksi.")
        else:
            print("\n✨ SEMPURNA! Semua data terisi 100%. Tidak ada item untuk direview.")

        print(f"\n📁 File output: {os.path.abspath(output_file)}")
        print("="*60)

    else:
        # === MODE IMAGE ONLY ===
        print("\n>> Mode: Image Only (Insert gambar saja)\n")
        template_file = IMG_TEMPLATE
        output_file   = IMG_OUTPUT
        mapping_file  = IMG_MAPPING
        daftar_file   = IMG_DAFTAR

        if not os.path.exists(mapping_file):
            print(f"File mapping '{mapping_file}' tidak ditemukan!")
            return
        mapping = baca_mapping_img(mapping_file)
        print(f"Mapping berisi {len(mapping)} entri.")

        items = baca_daftar(daftar_file)
        print(f"Daftar berisi {len(items)} item.")

        tanggal_list = get_tanggal_list(FOLDER_DATA)
        if not tanggal_list:
            print(f"Folder '{FOLDER_DATA}' tidak ditemukan atau kosong!")
            return
        print(f"Ditemukan {len(tanggal_list)} folder tanggal: {tanggal_list[:5]}...")

        if not os.path.exists(template_file):
            print(f"Template '{template_file}' tidak ditemukan!")
            return
        wb = load_workbook(template_file)
        print("Template berhasil dimuat.\n")

        for tgl in tanggal_list:
            print(f"Memproses tanggal: {tgl}")
            proses_tanggal_img(wb, tgl, items, mapping)

        wb.save(output_file)
        print("\n" + "=" * 60)
        print(f"  🎉 SELESAI! File Excel disimpan sebagai: {output_file}")
        print("=" * 60)


if __name__ == "__main__":
    main()