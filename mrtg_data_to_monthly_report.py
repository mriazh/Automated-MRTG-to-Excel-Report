import os
import re
import io
import sys
import logging
import traceback

# Suppress noisy PaddleOCR / Paddle C++ logs
os.environ['GLOG_minloglevel'] = '3'          # Suppress GLOG (C++ INFO/WARNING)
os.environ['FLAGS_minloglevel'] = '3'
os.environ['PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK'] = 'True'
os.environ['PADDLEX_DISABLE_PRINT'] = '1'
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string, get_column_letter
from PIL import Image as PILImage

# ========== KONFIGURASI ==========
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_DATA = os.path.join(BASE_DIR, "MRTG-Data")
IMAGE_SCALE = 0.98

# --- Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ocr_report.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('mrtg_report')

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
    """Singleton: inisialisasi PaddleOCR sekali saja.
    Kembali ke mode CPU paling stabil untuk menghindari error argument.
    """
    if not hasattr(_get_ocr_engine, '_engine'):
        import warnings
        import contextlib
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            from paddleocr import PaddleOCR
            
            logger.info("Menginisialisasi PaddleOCR engine (Stable Mode)...")
            
            with open(os.devnull, 'w') as fnull:
                with contextlib.redirect_stdout(fnull), contextlib.redirect_stderr(fnull):
                    # Gunakan inisialisasi paling dasar yang pasti jalan di semua versi
                    _get_ocr_engine._engine = PaddleOCR(lang='en')
            
        logger.info("PaddleOCR engine siap.")
    return _get_ocr_engine._engine


def extract_mrtg_values(image_path):
    """Ekstrak nilai bandwidth dengan strategi Sequential Context & Fuzzy Keywords."""
    try:
        ocr = _get_ocr_engine()
        logger.info(f"OCR memproses: {os.path.basename(image_path)}")

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

        if not all_texts:
            return None

        # --- Helper: Cari nilai setelah keyword ---
        def find_value_after(keyword_list, texts, start_search_idx):
            """Mencari angka (+ unit) setelah salah satu keyword ditemukan."""
            for i in range(start_search_idx, len(texts)):
                t = texts[i].lower()
                # Cocokkan keyword dengan toleransi typo
                if any(kw.lower() in t for kw in keyword_list):
                    # Cari angka di i atau 3 box setelahnya
                    for j in range(i, min(i + 4, len(texts))):
                        # Pattern baru: Mendukung integer (100) dan desimal (100.50)
                        match = re.search(r'(\d+(?:[\.,]\d+)?)\s*([MkGT]?)', texts[j])
                        if match:
                            val = match.group(1).replace(',', '.')
                            unit = match.group(2)
                            if not unit and j + 1 < len(texts):
                                next_t = texts[j+1].strip()
                                if next_t in ['M', 'k', 'G', 'T']:
                                    unit = next_t
                            return f"{val} {unit}".strip() if unit else f"{val} M"
                        
                        # Jika nemu teks N/A, jangan langsung berhenti, cek box selanjutnya siapa tahu ada angka
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
                if 'current' in all_texts[i].lower():
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
                'Current': find_value_after(['Current', 'Curren'], in_area, 0),
                'Average': find_value_after(['Average', 'Averaqe', 'Avera', 'Ave'], in_area, 0),
                'Maximum': find_value_after(['Maximum', 'Maximu', 'Max'], in_area, 0)
            }
        else:
            in_vals = {'Current': 'N/A', 'Average': 'N/A', 'Maximum': 'N/A'}

        # Parse Outbound Area
        if outbound_idx != -1:
            out_area = all_texts[outbound_idx:]
            out_vals = {
                'Current': find_value_after(['Current', 'Curren'], out_area, 0),
                'Average': find_value_after(['Average', 'Averaqe', 'Avera', 'Ave'], out_area, 0),
                'Maximum': find_value_after(['Maximum', 'Maximu', 'Max'], out_area, 0)
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

        sid = os.path.basename(image_path).replace('MRTG_', '').replace('.png', '')
        na_count = sum(1 for v in result.values() if v == 'N/A')
        status = "✅ OK" if na_count == 0 else f"⚠️ {na_count} N/A"
        logger.info(f"  [{sid}] {status} -> In: {result['Inbound_Current']}|{result['Inbound_Average']}|{result['Inbound_Maximum']}, Out: {result['Outbound_Current']}|{result['Outbound_Average']}|{result['Outbound_Maximum']}")

        return result
    except Exception as e:
        logger.error(f"OCR Error: {e}")
        return None
    except Exception as e:
        logger.error(f"OCR Error: {e}")
        return None
    except Exception as e:
        logger.error(f"OCR Error: {e}")
        return None


def tulis_nilai(sheet, entry, values):
    """Tulis nilai OCR ke sel Excel yang ditentukan."""
    for key, (row, col) in entry.items():
        if key in values and key != 'Image':
            sheet.cell(row=row, column=col, value=values[key])


def proses_tanggal_ocr(wb, tanggal_str, items, mapping, global_stats, review_list):
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
        id_clean = re.sub(r'^\(\d+\)\s*', '', id_val)
        persen = f"[{idx}/{total} — {idx*100//total}%]"
        label = f"{tipe} {id_val}"

        if id_clean not in mapping:
            print(f"    {persen} ⏭️  {label} — ID tidak ada di mapping")
            continue
        entry = mapping[id_clean]

        path_gambar = cari_path_gambar(FOLDER_DATA, tanggal_str, tipe, id_val)
        if not os.path.exists(path_gambar):
            print(f"    {persen} ❌ {label} — Gambar tidak ditemukan")
            fail_count += 1
            global_stats['fail'] += 1
            continue

        print(f"    {persen} 🔍 {label} — OCR sedang memproses...", end='', flush=True)
        values = extract_mrtg_values(path_gambar)
        if not values:
            print(f"\r    {persen} ❌ {label} — Gagal OCR (tidak ada data)")
            fail_count += 1
            global_stats['fail'] += 1
            review_list.append({'sid': id_clean, 'date': tanggal_str, 'status': 'Fail'})
            continue

        # Hitung N/A
        val_na = sum(1 for v in values.values() if v == 'N/A')
        if val_na == 0:
            print(f"\r    {persen} ✅ {label} — 6/6 nilai terdeteksi")
            ok_count += 1
            global_stats['ok'] += 1
        else:
            print(f"\r    {persen} ⚠️  {label} — {6-val_na}/6 nilai ({val_na} N/A)")
            na_count += 1
            global_stats['partial'] += 1
            review_list.append({'sid': id_clean, 'date': tanggal_str, 'status': 'Partial', 'na': val_na})

        tulis_nilai(sheet, entry, values)
        if 'Image' in entry:
            (start_row, start_col), (end_row, end_col) = entry['Image']
            tambah_gambar_di_area(sheet, path_gambar, start_row, start_col, end_row, end_col)

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

        for tgl_idx, tgl in enumerate(tanggal_list, 1):
            print(f"\nMemproses tanggal: {tgl} ({tgl_idx}/{len(tanggal_list)})")
            proses_tanggal_ocr(wb, tgl, items, mapping, global_stats, review_list)

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
            print("\n🔍 LIST ITEM YANG PERLU DICEK (REVIEW LIST):")
            print(f"{'No':<4} | {'SID':<20} | {'Sheet':<6} | {'Tanggal':<10} | {'Status':<8} | {'Detail'}")
            print("-" * 85)
            for i, item in enumerate(review_list):
                # Ambil detail N/A atau status
                detail = f"{item['na']} nilai N/A" if 'na' in item else "Gagal total"
                # Cari sheet name (biasanya sheet title udah ada di item dari proses_tanggal_ocr)
                sheet_info = item.get('sheet', '??')
                print(f"{i+1:<4} | {item['sid']:<20} | {sheet_info:<6} | {item['date']:<10} | {item['status']:<8} | {detail}")
            print("-" * 85)
            print(f"Total {len(review_list)} item butuh review.")
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
