from PIL import Image
import numpy as np

input_path = 'temp_images/table_line.png'
output_path = 'temp_images/table_line_redline_only.png'

# 讀取圖片
img = Image.open(input_path).convert('RGBA')
data = np.array(img)

# 建立全透明背景
transparent = np.zeros_like(data)

# 要保留的區域
x, y, w, h = 81, 1157, 1797, 747
x2, y2 = x + w, y + h

# 只保留紅色線條（R>200, G<80, B<80）
region = data[y:y2, x:x2]
red_mask = (region[:, :, 0] > 200) & (region[:, :, 1] < 80) & (region[:, :, 2] < 80)

# 將紅線像素複製到透明背景
transparent[y:y2, x:x2][red_mask] = region[red_mask]

# 產生新圖片
img2 = Image.fromarray(transparent)
img2.save(output_path)

print(f'已產生只保留紅色線條的去背圖：{output_path}') 