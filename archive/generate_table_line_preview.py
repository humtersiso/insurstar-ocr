import os
from image_processing import ImageProcessing

PDF_PATH = 'assets/RAS-5879 要保書_任意+強制_無車體.pdf'  # 你可以改成其他 PDF 路徑
OUTPUT_FOLDER = 'temp_images'
TABLE_LINE_PATH = os.path.join(OUTPUT_FOLDER, 'table_line.png')

# 1. PDF 轉圖片
image_paths = ImageProcessing.pdf_to_images(PDF_PATH, OUTPUT_FOLDER)

# 2. 疊加 table_line.png
for img_path in image_paths:
    base_name = os.path.splitext(os.path.basename(img_path))[0]
    output_path = os.path.join(OUTPUT_FOLDER, f'{base_name}_table_lines_overlay.png')
    ImageProcessing.overlay_table_line(img_path, TABLE_LINE_PATH, output_path, alpha=0.5)
    print(f'已產生：{output_path}') 