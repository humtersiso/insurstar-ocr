from PIL import Image
import numpy as np

input_path = 'temp_images/support_line.png'
output_path = 'temp_images/support_line_transparent.png'

# 表格區塊座標
x, y, w, h = 81, 1157, 1797, 747

# 讀取圖片
img = Image.open(input_path).convert('RGBA')
data = np.array(img)

# 紅線遮罩
red_mask = (data[:, :, 0] > 180) & (data[:, :, 1] < 80) & (data[:, :, 2] < 80)

# 全部設為透明
data[:, :, 3] = 0
# 紅線設為純紅且完全不透明
data[red_mask, 0] = 255  # R
data[red_mask, 1] = 0    # G
data[red_mask, 2] = 0    # B
data[red_mask, 3] = 255  # A

# 只裁出表格區塊
crop = data[y:y+h, x:x+w]
img2 = Image.fromarray(crop)
img2.save(output_path)
print(f'已產生裁切後的透明紅線輔助線圖：{output_path}') 