"""
從 PDF 擷取前三款產品圖片，建立圖庫資料夾結構。
用法：python3 extract_images.py
"""
import fitz
import os
from pathlib import Path

PDF_PATH  = "1. 鋁線槽系列P2~P15.pdf"
BASE      = Path("圖庫")
MIN_SIZE  = 120   # 過濾掉寬或高小於此值的雜圖

# (料號, PDF頁碼索引, 標籤)
# PDF index 0 = 型錄 P002，index 1 = P003
PRODUCTS = [
    ("HB128",  0, "P002"),
    ("HB128X", 0, "P002"),
    ("HB381A", 1, "P003"),
]

doc = fitz.open(PDF_PATH)

# 先把每頁需要的圖片提取出來存放，避免重複開檔
page_cache = {}

for model, page_idx, label in PRODUCTS:
    folder = BASE / model
    folder.mkdir(parents=True, exist_ok=True)
    print(f"\n[{model}] 建立資料夾：{folder}")

    page = doc[page_idx]

    # ── 1. 全頁高解析度截圖（2.5x ≈ 180 dpi） ──────────────────────────────
    if page_idx not in page_cache:
        mat  = fitz.Matrix(2.5, 2.5)
        pix  = page.get_pixmap(matrix=mat, alpha=False)
        page_cache[page_idx] = pix
    pix = page_cache[page_idx]
    ref = folder / "全頁參考.png"
    pix.save(str(ref))
    print(f"  ✓ 全頁參考截圖 ({pix.width}×{pix.height}px) → {ref.name}")

    # ── 2. 提取頁面內嵌圖片 ───────────────────────────────────────────────
    images   = page.get_images(full=True)
    seen_xref = set()
    saved    = 0
    skipped  = 0

    for img_info in images:
        xref = img_info[0]
        if xref in seen_xref:
            continue
        seen_xref.add(xref)

        try:
            base_img  = doc.extract_image(xref)
            img_bytes = base_img["image"]
            ext       = base_img["ext"]
            w, h      = base_img["width"], base_img["height"]

            if w < MIN_SIZE or h < MIN_SIZE:
                skipped += 1
                continue

            saved += 1
            # 以尺寸排序命名，供後續手動辨識
            out = folder / f"提取_{saved:02d}_{w}x{h}.{ext}"
            out.write_bytes(img_bytes)
        except Exception as e:
            print(f"  ⚠ xref={xref} 失敗：{e}")

    print(f"  ✓ 內嵌圖片：擷取 {saved} 張（跳過過小圖 {skipped} 張）")

doc.close()

print("\n" + "─" * 50)
print("完成！資料夾總覽：")
for model, _, _ in PRODUCTS:
    files = list((BASE / model).iterdir())
    print(f"  圖庫/{model}/ → {len(files)} 個檔案")
    for f in sorted(files):
        size_kb = f.stat().st_size // 1024
        print(f"    {f.name}  ({size_kb} KB)")
