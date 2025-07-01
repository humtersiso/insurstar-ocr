import cv2
import numpy as np
import os

# 紅線 template 路徑
template_path = 'static/table_preview/table_lines_template.png'
template = cv2.imread(template_path, cv2.IMREAD_UNCHANGED)
if template is None:
    raise FileNotFoundError(f'找不到 template: {template_path}')

h, w = template.shape[:2]
x, y = 85, 1156

src_dir = 'converted_images'
dst_dir = 'static/table_preview'
os.makedirs(dst_dir, exist_ok=True)

for fname in os.listdir(src_dir):
    if 'page_01' in fname and fname.lower().endswith('.png'):
        src_path = os.path.join(src_dir, fname)
        img = cv2.imread(src_path)
        if img is None:
            print(f'無法讀取圖片: {src_path}')
            continue
        # 檢查尺寸
        if img.shape[0] < y + h or img.shape[1] < x + w:
            print(f'圖片尺寸不足: {src_path}')
            continue
        # 疊加 template
        roi = img[y:y+h, x:x+w]
        if template.shape[2] == 4:
            alpha = template[:,:,3] / 255.0
            for c in range(3):
                roi[:,:,c] = (1 - alpha) * roi[:,:,c] + alpha * template[:,:,c]
        else:
            roi = cv2.addWeighted(roi, 0.7, template, 0.3, 0)
        img[y:y+h, x:x+w] = roi
        out_path = os.path.join(dst_dir, f'{os.path.splitext(fname)[0]}_table_lines_overlay.png')
        cv2.imwrite(out_path, img)
        print(f'已產生浮水印紅線: {out_path}') 