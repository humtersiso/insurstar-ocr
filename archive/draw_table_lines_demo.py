from PIL import Image, ImageDraw

# 目標圖片路徑
base_path = 'static/table_preview/RAS-5879_page_01_table_lines_overlay.png'
output_path = 'static/table_preview/RAS-5879_page_01_table_lines_overlay_demo.png'

# 表格區域座標與大小
x, y, w, h = 81, 1157, 1797, 747
rows = 6  # 可依實際表格結構調整
cols = 4  # 可依實際表格結構調整
line_width = 3

base = Image.open(base_path).convert('RGB')
draw = ImageDraw.Draw(base)
# 畫外框
draw.rectangle([x, y, x+w, y+h], outline=(255,0,0), width=line_width)
# 畫橫線
row_height = h // rows
for i in range(1, rows):
    y_line = y + i * row_height
    draw.line([(x, y_line), (x+w, y_line)], fill=(255,0,0), width=line_width)
# 畫直線
col_width = w // cols
for i in range(1, cols):
    x_line = x + i * col_width
    draw.line([(x_line, y), (x_line, y+h)], fill=(255,0,0), width=line_width)
base.save(output_path)
print(f'已直接畫紅線：{output_path}') 