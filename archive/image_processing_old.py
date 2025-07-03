import fitz  # PyMuPDF
from PIL import Image
import os
import numpy as np

class ImageProcessing:
    @staticmethod
    def pdf_to_images(pdf_path, output_folder, dpi=300):
        """
        使用 PyMuPDF 將 PDF 每一頁轉成圖片，存到 output_folder，回傳圖片路徑 list。
        """
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        doc = fitz.open(pdf_path)
        image_paths = []
        for i in range(len(doc)):
            page = doc.load_page(i)
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat)
            mode = "RGBA" if pix.alpha else "RGB"
            base = Image.frombytes(mode, (pix.width, pix.height), pix.samples).convert('RGBA')
            img_path = os.path.join(output_folder, f"page_{i+1:02d}.png")
            base.save(img_path)
            image_paths.append(img_path)
        doc.close()
        return image_paths

    @staticmethod
    def overlay_support_line_on_pdf(pdf_path, output_folder, support_line_path, dpi=300, alpha=0.8):
        """
        將 PDF 每頁轉成圖片，並在 (81, 1157, 1797, 747) 疊加 support_line 圖片裁切區域，存到 output_folder。
        """
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        doc = fitz.open(pdf_path)
        support_line = Image.open(support_line_path).convert('RGBA')
        # 只裁切 (81, 1157, 1797, 747)
        x, y, w, h = 81, 1157, 1797, 747
        support_line_crop = support_line.crop((x, y, x + w, y + h))
        image_paths = []
        for i in range(len(doc)):
            page = doc.load_page(i)
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat)
            mode = "RGBA" if pix.alpha else "RGB"
            base = Image.frombytes(mode, (pix.width, pix.height), pix.samples).convert('RGBA')
            # 建立一個透明圖層，與 base 同大小
            overlay = Image.new('RGBA', base.size, (0,0,0,0))
            # 貼到 overlay 的指定區域
            overlay.paste(support_line_crop, (x, y), support_line_crop)
            # 疊加
            combined = Image.alpha_composite(base, overlay)
            combined = combined.convert('RGB')
            out_path = os.path.join(output_folder, f"page_{i+1:02d}_with_support_line.png")
            combined.save(out_path)
            image_paths.append(out_path)
        doc.close()
        return image_paths 

    @staticmethod
    def overlay_support_line_on_image(
        image_path, support_line_path, output_path,
        orig_size=(2481, 3508), crop_box=(81, 1157, 1797, 747), alpha=1.0
    ):
        """
        將 support_line_path 浮水印紅線依底圖尺寸自動縮放與對齊，疊加到底圖 image_path 上，存到 output_path。
        orig_size: 浮水印原圖尺寸（預設為A4 300dpi）
        crop_box: (x, y, w, h) 浮水印紅線區域在原圖的座標與大小
        alpha: 浮水印透明度（0~1）
        """
        base = Image.open(image_path).convert('RGBA')
        support_line = Image.open(support_line_path).convert('RGBA')
        base_w, base_h = base.size
        orig_w, orig_h = orig_size
        scale_x = base_w / orig_w
        scale_y = base_h / orig_h
        x, y, w, h = crop_box
        new_x = int(x * scale_x)
        new_y = int(y * scale_y)
        new_w = int(w * scale_x)
        new_h = int(h * scale_y)
        # 浮水印裁切並縮放
        support_line_crop = support_line.crop((x, y, x + w, y + h)).resize((new_w, new_h), Image.Resampling.LANCZOS)
        # 疊加
        overlay = Image.new('RGBA', base.size, (0,0,0,0))
        overlay.paste(support_line_crop, (new_x, new_y), support_line_crop)
        # 可調整透明度
        if alpha < 1.0:
            alpha_mask = overlay.split()[-1].point(lambda p: int(p * alpha))
            overlay.putalpha(alpha_mask)
        combined = Image.alpha_composite(base, overlay).convert('RGB')
        combined.save(output_path) 