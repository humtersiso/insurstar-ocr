import os
import numpy as np
from PIL import Image

def make_color_watermark(input_path, output_path, white_thresh=220, only_blue=False):
    pil_img = Image.open(input_path).convert('RGBA')
    arr = np.array(pil_img)
    r, g, b, a = arr[:,:,0], arr[:,:,1], arr[:,:,2], arr[:,:,3]
    # 去除白色背景
    mask = (r > white_thresh) & (g > white_thresh) & (b > white_thresh)
    arr[:,:,3] = np.where(mask, 0, 255)
    # 只保留藍色（可選）
    if only_blue:
        blue_mask = (b > r + 20) & (b > g + 20) & (b > 80)
        arr[:,:,3] = np.where(blue_mask, arr[:,:,3], 0)
    out_img = Image.fromarray(arr)
    out_img.save(output_path)
    print(f'已產生浮水印: {output_path}')

if __name__ == '__main__':
    os.makedirs('assets/watermark', exist_ok=True)
    # 參數組合
    thresholds = [200, 210, 220, 230, 240]
    for t in thresholds:
        make_color_watermark('assets/templates/signature_company.JPG', f'assets/watermark/watermark_company_{t}.png', white_thresh=t)
        make_color_watermark('assets/templates/signature_name.jpg', f'assets/watermark/watermark_name_{t}.png', white_thresh=t)
    # 只保留藍色版本
    make_color_watermark('assets/templates/signature_company.JPG', 'assets/watermark/watermark_company_blue.png', white_thresh=220, only_blue=True)
    make_color_watermark('assets/templates/signature_name.jpg', 'assets/watermark/watermark_name_blue.png', white_thresh=220, only_blue=True) 