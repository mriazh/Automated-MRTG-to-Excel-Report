import os
import re
import io
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string, get_column_letter
from PIL import Image as PILImage

# ========== KONFIGURASI ==========
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_DATA = os.path.join(BASE_DIR, "MRTG-Data")
IMAGE_SCALE = 0.98

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
    """Singleton: inisialisasi PaddleOCR sekali saja agar tidak reload model tiap gambar."""
    if not hasattr(_get_ocr_engine, '_engine'):
        from paddleocr import PaddleOCR
        _get_ocr_engine._engine = PaddleOCR(
            use_angle_cls=False,
            lang='en',
            show_log=False,
        )
    return _get_ocr_engine._engine


def extract_mrtg_values(image_path):
    """Ekstrak nilai bandwidth dari gambar MRTG menggunakan PaddleOCR (Deep Learning).

    Mengembalikan dict berisi Inbound/Outbound × Current/Average/Maximum,
    atau None jika gagal.
    """
    try:
        ocr = _get_ocr_engine()
        results = ocr.ocr(image_path, cls=False)

        if not results or not results[0]:
            print(f"    PaddleOCR: Tidak ada teks terdeteksi di {image_path}")
            return None

        # Kumpulkan semua teks yang terdeteksi beserta posisi Y-nya
        detected = []
        for line in results[0]:
            bbox = line[0]           # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
            text = line[1][0]        # recognized text
            confidence = line[1][1]  # confidence score
            y_center = (bbox[0][1] + bbox[2][1]) / 2  # rata-rata Y atas dan bawah
            detected.append((y_center, text, confidence))

        # Gabungkan semua teks menjadi satu string untuk parsing
        full_text = ' '.join([t for _, t, _ in detected])

        # Cari baris Inbound dan Outbound
        # PaddleOCR mengembalikan teks per-box, jadi kita kelompokkan berdasarkan Y
        # Box yang Y-nya berdekatan (±15px) dianggap satu baris
        detected.sort(key=lambda x: x[0])  # sort by Y

        lines_grouped = []
        current_line = []
        current_y = None

        for y, text, conf in detected:
            if current_y is None or abs(y - current_y) < 15:
                current_line.append(text)
                current_y = y if current_y is None else (current_y + y) / 2
            else:
                lines_grouped.append(' '.join(current_line))
                current_line = [text]
                current_y = y
        if current_line:
            lines_grouped.append(' '.join(current_line))

        inbound_line = ""
        outbound_line = ""
        for line in lines_grouped:
            if 'Inbound' in line or 'inbound' in line:
                inbound_line = line
            if 'Outbound' in line or 'outbound' in line:
                outbound_line = line

        # Fallback: cari di full text
        if not inbound_line and 'Inbound' in full_text:
            inbound_line = full_text
        if not outbound_line and 'Outbound' in full_text:
            outbound_line = full_text

        def parse_line(line):
            """Ekstrak Current, Average, Maximum dari satu baris teks."""
            result = {'Current': 'N/A', 'Average': 'N/A', 'Maximum': 'N/A'}
            for keyword in ['Current', 'Average', 'Maximum']:
                # Pattern: keyword diikuti tanda : atau spasi, lalu angka + satuan
                pattern = rf'{keyword}\s*:?\s*([\d\.]+)\s*([MkGT]?)'
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    raw_val = match.group(1)
                    unit = match.group(2) if match.group(2) else ''
                    try:
                        val = float(raw_val)
                        if val > 100000:
                            result[keyword] = 'N/A'
                        else:
                            result[keyword] = f"{raw_val} {unit}" if unit else f"{raw_val} M"
                    except ValueError:
                        result[keyword] = 'N/A'
            return result

        inbound_vals = parse_line(inbound_line) if inbound_line else {'Current': 'N/A', 'Average': 'N/A', 'Maximum': 'N/A'}
        outbound_vals = parse_line(outbound_line) if outbound_line else {'Current': 'N/A', 'Average': 'N/A', 'Maximum': 'N/A'}

        result = {
            'Inbound_Current': inbound_vals['Current'],
            'Inbound_Average': inbound_vals['Average'],
            'Inbound_Maximum': inbound_vals['Maximum'],
            'Outbound_Current': outbound_vals['Current'],
            'Outbound_Average': outbound_vals['Average'],
            'Outbound_Maximum': outbound_vals['Maximum']
        }
        return result
    except Exception as e:
        print(f"OCR Error: {e}")
        return None


def tulis_nilai(sheet, entry, values):
    """Tulis nilai OCR ke sel Excel yang ditentukan."""
    for key, (row, col) in entry.items():
        if key in values and key != 'Image':
            sheet.cell(row=row, column=col, value=values[key])


def proses_tanggal_ocr(wb, tanggal_str, items, mapping):
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

    for nomor, tipe, id_val in items:
        id_clean = re.sub(r'^\(\d+\)\s*', '', id_val)
        if id_clean not in mapping:
            print(f"    Peringatan: ID '{id_clean}' tidak ada di mapping")
            continue
        entry = mapping[id_clean]

        path_gambar = cari_path_gambar(FOLDER_DATA, tanggal_str, tipe, id_val)
        if not os.path.exists(path_gambar):
            print(f"    Gambar tidak ditemukan: {path_gambar}")
            continue

        values = extract_mrtg_values(path_gambar)
        if not values:
            print(f"    Gagal OCR untuk {id_val}")
            continue

        tulis_nilai(sheet, entry, values)
        if 'Image' in entry:
            (start_row, start_col), (end_row, end_col) = entry['Image']
            tambah_gambar_di_area(sheet, path_gambar, start_row, start_col, end_row, end_col)


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

        for tgl in tanggal_list:
            print(f"Memproses tanggal: {tgl}")
            proses_tanggal_ocr(wb, tgl, items, mapping)

        wb.save(output_file)
        print("\n" + "=" * 60)
        print(f"  🎉 SELESAI! File Excel dengan {len(tanggal_list)} sheet telah dibuat.")
        print(f"  📁 File output: {output_file}")
        print("=" * 60)

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
