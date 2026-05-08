"""Diagnostic script: cek format output PaddleOCR 3.5.0 predict()"""
import os
import sys

# Cari satu gambar MRTG sample
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_DATA = os.path.join(BASE_DIR, "MRTG-Data")

sample_img = None
for d in os.listdir(FOLDER_DATA):
    folder = os.path.join(FOLDER_DATA, d)
    if os.path.isdir(folder):
        for f in os.listdir(folder):
            if f.endswith('.png'):
                sample_img = os.path.join(folder, f)
                break
        if sample_img:
            break

if not sample_img:
    print("Tidak ditemukan gambar sample!")
    sys.exit(1)

print(f"Sample image: {sample_img}")
print(f"File exists: {os.path.exists(sample_img)}")
print(f"File size: {os.path.getsize(sample_img)} bytes")
print()

# Init PaddleOCR
from paddleocr import PaddleOCR
print("=== Inisialisasi PaddleOCR ===")
ocr = PaddleOCR(lang='en')
print()

# Test predict()
print("=== Calling ocr.predict() ===")
result = ocr.predict(sample_img)
print(f"predict() returned: {type(result)}")
print()

# Force evaluation
print("=== Converting to list ===")
result_list = list(result) if hasattr(result, '__iter__') else [result]
print(f"Number of results: {len(result_list)}")
print()

for i, res in enumerate(result_list):
    print(f"--- Result [{i}] ---")
    print(f"  Type: {type(res).__name__}")
    
    # List all non-private attributes
    attrs = [a for a in dir(res) if not a.startswith('_')]
    print(f"  Attributes: {attrs}")
    
    # Dump each attribute value
    for attr in attrs:
        try:
            val = getattr(res, attr)
            if callable(val):
                if attr == 'print':
                    print(f"\n  === res.print() output ===")
                    val()
                    print(f"  === end res.print() ===\n")
                elif attr == 'json':
                    pass  # skip json method
                continue
            val_str = str(val)[:500]
            print(f"  res.{attr} ({type(val).__name__}): {val_str}")
        except Exception as e:
            print(f"  res.{attr}: ERROR - {e}")

print("\n=== DONE ===")
