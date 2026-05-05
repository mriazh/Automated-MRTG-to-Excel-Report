import os
import re
import io
import cv2
import numpy as np
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string, get_column_letter
from PIL import Image as PILImage
import pytesseract

# ========== KONFIGURASI ==========
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

FOLDER_DATA = "MRTG-Data"
TEMPLATE_FILE = "Report on Internet Bandwidth Utilization by Telkom (MRTG).xlsx"
OUTPUT_FILE = "Complete_Monthly_Report.xlsx"
MAPPING_FILE = "sid-in-out-image-position-excel.txt"
DAFTAR_FILE = "list_mrtg_data.txt"
IMAGE_SCALE = 0.98

# ========== BACA MAPPING ==========
def baca_mapping(filepath):
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

# ========== BACA DAFTAR ITEM ==========
def baca_daftar(filepath):
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

# ========== EKSTRAK NILAI DARI GAMBAR (FINAL - ROBUST) ==========
def extract_mrtg_values(image_path):
    try:
        img = cv2.imread(image_path)
        if img is None:
            return None
        
        # Preprocessing
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)
        thresh = cv2.resize(thresh, None, fx=3, fy=3, interpolation=cv2.INTER_LINEAR)
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(thresh, config=custom_config)
        
        # Debug (aktifkan jika perlu)
        # with open("ocr_debug.txt", "a", encoding="utf-8") as f:
        #     f.write(f"File: {image_path}\n{text}\n{'-'*80}\n")
        
        # Cari baris yang mengandung Inbound dan Outbound
        lines = text.split('\n')
        inbound_line = ""
        outbound_line = ""
        for line in lines:
            if 'Inbound' in line:
                inbound_line = line
            if 'Outbound' in line:
                outbound_line = line
        
        # Jika tidak ketemu, coba gabung semua teks (mungkin dalam satu baris)
        if not inbound_line and 'Inbound' in text:
            inbound_line = text
        if not outbound_line and 'Outbound' in text:
            outbound_line = text
        
        def parse_line(line):
            """Ekstrak Current, Average, Maximum dari satu baris teks."""
            result = {'Current': 'N/A', 'Average': 'N/A', 'Maximum': 'N/A'}
            # Pola: kata kunci diikuti titik dua, spasi, lalu angka atau N/A, lalu spasi, lalu satuan opsional
            for keyword in ['Current', 'Average', 'Maximum']:
                # Pattern menangkap angka (dengan titik desimal) atau kata N/A (case-insensitive)
                pattern = rf'{keyword}\s*:\s*([\d\.]+|[Nn][Aa]|[Nn]/?[Aa]?|[Nn][Aa][Nn]?)\s*([Mk]?)'
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    raw_val = match.group(1)
                    unit = match.group(2) if match.group(2) else ''
                    # Normalisasi nilai
                    if re.match(r'^[Nn]', raw_val):  # N/A, NA, NaN, nan
                        result[keyword] = 'N/A'
                    else:
                        try:
                            val = float(raw_val)
                            # Validasi: nilai wajar (tidak terlalu besar dari 100000, yang muncul karena error OCR)
                            if val > 100000:
                                result[keyword] = 'N/A'
                            else:
                                if unit:
                                    result[keyword] = f"{raw_val} {unit}"
                                else:
                                    result[keyword] = f"{raw_val} M"
                        except:
                            result[keyword] = 'N/A'
                else:
                    # Fallback: cari angka di sekitar keyword (tanpa titik dua)
                    pattern2 = rf'{keyword}\s+([\d\.]+)\s*([Mk]?)'
                    match2 = re.search(pattern2, line, re.IGNORECASE)
                    if match2:
                        raw_val = match2.group(1)
                        unit = match2.group(2) if match2.group(2) else ''
                        try:
                            val = float(raw_val)
                            if val > 100000:
                                result[keyword] = 'N/A'
                            else:
                                result[keyword] = f"{raw_val} {unit}" if unit else f"{raw_val} M"
                        except:
                            result[keyword] = 'N/A'
            return result
        
        inbound_vals = parse_line(inbound_line) if inbound_line else {'Current':'N/A','Average':'N/A','Maximum':'N/A'}
        outbound_vals = parse_line(outbound_line) if outbound_line else {'Current':'N/A','Average':'N/A','Maximum':'N/A'}
        
        # Jika semua N/A, coba ekstrak angka berurutan dari seluruh teks (fallback terakhir)
        if all(v == 'N/A' for v in inbound_vals.values()) and all(v == 'N/A' for v in outbound_vals.values()):
            # Cari semua angka dengan satuan opsional
            all_matches = re.findall(r'([\d\.]+)\s*([Mk]?)', text)
            if len(all_matches) >= 6:
                # Urutan: Inbound Current, Inbound Average, Inbound Maximum, Outbound Current, Outbound Average, Outbound Maximum
                keys = ['Inbound_Current', 'Inbound_Average', 'Inbound_Maximum',
                        'Outbound_Current', 'Outbound_Average', 'Outbound_Maximum']
                for i, (val_str, unit) in enumerate(all_matches[:6]):
                    try:
                        val = float(val_str)
                        if val > 100000:
                            continue
                        unit_disp = unit if unit else 'M'
                        inbound_vals[keys[i].split('_')[1]] = f"{val_str} {unit_disp}"
                    except:
                        pass
        
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

# ========== FUNGSI LAINNYA ==========
def get_area_size_pixels(sheet, start_row, start_col, end_row, end_col):
    total_width = 0
    for col in range(start_col, end_col + 1):
        col_letter = get_column_letter(col)
        col_width = sheet.column_dimensions[col_letter].width
        if col_width is None:
            col_width = 8.43
        total_width += col_width * 7.4
    total_height = 0
    for row in range(start_row, end_row + 1):
        row_height = sheet.row_dimensions[row].height
        if row_height is None:
            row_height = 15
        total_height += row_height * 1.333
    return total_width, total_height

def resize_image_stretch(image_path, target_width, target_height):
    with PILImage.open(image_path) as img:
        if img.mode in ('RGBA', 'LA', 'P'):
            img = img.convert('RGB')
        img_resized = img.resize((int(target_width), int(target_height)), PILImage.Resampling.LANCZOS)
        output = io.BytesIO()
        img_resized.save(output, format='PNG')
        output.seek(0)
        return output

def tambah_gambar_di_area(sheet, image_path, start_row, start_col, end_row, end_col, scale=IMAGE_SCALE):
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

def tulis_nilai(sheet, entry, values):
    for key, (row, col) in entry.items():
        if key in values and key != 'Image':
            sheet.cell(row=row, column=col, value=values[key])

def proses_tanggal(wb, tanggal_str, items, mapping):
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
        
        if tipe == 'SID':
            nama_file = f"MRTG_{id_val}.png"
        else:
            nama_file = f"MRTG_{id_val}_{tanggal_str}.png"
        path_gambar = os.path.join(FOLDER_DATA, tanggal_str, nama_file)
        if not os.path.exists(path_gambar):
            folder_tgl = os.path.join(FOLDER_DATA, tanggal_str)
            if os.path.exists(folder_tgl):
                for f in os.listdir(folder_tgl):
                    if f.startswith(f"MRTG_{id_val}") and f.endswith(".png"):
                        path_gambar = os.path.join(folder_tgl, f)
                        break
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

# ========== MAIN ==========
def main():
    print("=" * 60)
    print("AUTOMATED MRTG TO EXCEL - FINAL ROBUST OCR")
    print("=" * 60)
    
    if not os.path.exists(MAPPING_FILE):
        print(f"File mapping '{MAPPING_FILE}' tidak ditemukan!")
        return
    mapping = baca_mapping(MAPPING_FILE)
    print(f"Mapping berisi {len(mapping)} entri.")
    
    items = baca_daftar(DAFTAR_FILE)
    print(f"Daftar berisi {len(items)} item.")
    
    if not os.path.exists(FOLDER_DATA):
        print(f"Folder '{FOLDER_DATA}' tidak ditemukan!")
        return
    tanggal_list = [d for d in os.listdir(FOLDER_DATA) if os.path.isdir(os.path.join(FOLDER_DATA, d)) and d.isdigit() and len(d) == 8]
    tanggal_list.sort()
    print(f"Ditemukan {len(tanggal_list)} folder tanggal.")
    
    if not os.path.exists(TEMPLATE_FILE):
        print(f"Template '{TEMPLATE_FILE}' tidak ditemukan!")
        return
    
    wb = load_workbook(TEMPLATE_FILE)
    print("Template berhasil dimuat.")
    
    for tgl in tanggal_list:
        print(f"\nMemproses tanggal: {tgl}")
        proses_tanggal(wb, tgl, items, mapping)
    
    wb.save(OUTPUT_FILE)
    print("\n" + "=" * 60)
    print(f"🎉 SELESAI! File Excel dengan {len(tanggal_list)} sheet telah dibuat.")
    print(f"📁 File output: {OUTPUT_FILE}")
    print("=" * 60)

if __name__ == "__main__":
    main()