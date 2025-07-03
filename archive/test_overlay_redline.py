from PIL import Image
import numpy as np
import os
import fitz  # PyMuPDF

# 互動式選擇 PDF 檔
print('請選擇 PDF 檔：')
pdf_dir = 'assets'
pdf_candidates = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
for idx, fname in enumerate(pdf_candidates):
    print(f'{idx+1}. {fname}')
pdf_idx = int(input('請輸入 PDF 編號：')) - 1
pdf_path = os.path.join(pdf_dir, pdf_candidates[pdf_idx])

# 轉 PDF 第一頁為圖片
doc = fitz.open(pdf_path)
page = doc.load_page(0)
mat = fitz.Matrix(300/72, 300/72)
pix = page.get_pixmap(matrix=mat)
base_img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples).convert('RGBA')
base_img.save('test_base.png')
print('已產生 test_base.png（PDF 第一頁轉圖）')

# 互動式選擇輔助線圖
print('\n請選擇輔助線圖檔：')
converted_dir = 'temp_images'
redline_candidates = [f for f in os.listdir(converted_dir) if 'support_line' in f and f.lower().endswith('.png')]
for idx, fname in enumerate(redline_candidates):
    print(f'{idx+1}. {fname}')
redline_idx = int(input('請輸入輔助線圖編號：')) - 1
redline_path = os.path.join(converted_dir, redline_candidates[redline_idx])

output_path = 'test_overlay.png'

# 表格區塊座標
x, y, w, h = 81, 1157, 1797, 747

# 讀取輔助線圖，只裁出表格區塊
redline_img = Image.open(redline_path).convert('RGBA')
redline_crop = redline_img.crop((x, y, x+w, y+h))

# 貼到底圖對應區塊
base_img.paste(redline_crop, (x, y), mask=redline_crop)

base_img.save(output_path)
print(f'已產生疊加紅線後的圖片：{output_path}')
print('請人工比對 test_base.png（原圖） 與 test_overlay.png（疊加紅線後）')
try:
    base_img.show()
except Exception as e:
    print('無法自動顯示圖片，請手動打開 test_base.png 與 test_overlay.png 檢查') 