import cv2
import numpy as np
import os

# 來源圖片與輸出路徑
src_path = 'temp_images/table_line.png'
out_path = 'static/table_preview/table_lines_template.png'

# 裁切座標
x, y, w, h = 85, 1156, 1794, 743

# 讀取原圖
img = cv2.imread(src_path)
if img is None:
    raise FileNotFoundError(f'找不到圖片: {src_path}')

# 裁切區塊
roi = img[y:y+h, x:x+w]

# 建立 alpha 通道，預設全透明
alpha = np.zeros((h, w), dtype=np.uint8)

# 偵測紅線（BGR: 0,0,255，允許一點誤差）
red_mask = (roi[:,:,2] > 200) & (roi[:,:,0] < 50) & (roi[:,:,1] < 50)

# 紅線區域設為不透明
alpha[red_mask] = 255

# 合併 BGR 與 alpha
rgba = cv2.cvtColor(roi, cv2.COLOR_BGR2BGRA)
rgba[:,:,3] = alpha

# 輸出
os.makedirs(os.path.dirname(out_path), exist_ok=True)
cv2.imwrite(out_path, rgba)
print(f'已產生紅線 template: {out_path}') 