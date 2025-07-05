import fitz  # PyMuPDF
from PIL import Image
import os
import numpy as np

class ImageProcessing:
    @staticmethod
    def pdf_to_images(pdf_path, output_folder, dpi=300):
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
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        # 檢查是PDF還是圖片檔案
        if pdf_path.lower().endswith('.pdf'):
            doc = fitz.open(pdf_path)
            support_line = Image.open(support_line_path).convert('RGBA')
            x, y, w, h = 81, 1157, 1797, 747
            support_line_crop = support_line.crop((x, y, x + w, y + h))
            image_paths = []
            for i in range(len(doc)):
                page = doc.load_page(i)
                mat = fitz.Matrix(dpi/72, dpi/72)
                pix = page.get_pixmap(matrix=mat)
                mode = "RGBA" if pix.alpha else "RGB"
                base = Image.frombytes(mode, (pix.width, pix.height), pix.samples).convert('RGBA')
                
                # 根據實際圖片尺寸縮放輔助線
                orig_w, orig_h = 2481, 3508  # 原始模板尺寸
                base_w, base_h = base.size
                scale_x = base_w / orig_w
                scale_y = base_h / orig_h
                
                new_x = int(x * scale_x)
                new_y = int(y * scale_y)
                new_w = int(w * scale_x)
                new_h = int(h * scale_y)
                
                # 重新裁剪和縮放輔助線
                scaled_support_line = support_line_crop.resize((new_w, new_h), Image.Resampling.LANCZOS)
                
                overlay = Image.new('RGBA', base.size, (0,0,0,0))
                overlay.paste(scaled_support_line, (new_x, new_y), scaled_support_line)
                combined = Image.alpha_composite(base, overlay)
                combined = combined.convert('RGB')
                out_path = os.path.join(output_folder, f"page_{i+1:02d}_with_support_line.png")
                combined.save(out_path)
                image_paths.append(out_path)
            doc.close()
            return image_paths
        else:
            # 如果是圖片檔案，使用 overlay_support_line_on_image 方法
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_path = os.path.join(output_folder, f"{base_name}_with_support_line.png")
            ImageProcessing.overlay_support_line_on_image(
                pdf_path, support_line_path, out_path,
                orig_size=(2481, 3508), crop_box=(81, 1157, 1797, 747), alpha=alpha
            )
            return [out_path] if os.path.exists(out_path) else []

    @staticmethod
    def overlay_support_line_on_image(
        image_path, support_line_path, output_path,
        orig_size=(2481, 3508), crop_box=(81, 1157, 1797, 747), alpha=1.0
    ):
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
        support_line_crop = support_line.crop((x, y, x + w, y + h)).resize((new_w, new_h), Image.Resampling.LANCZOS)
        overlay = Image.new('RGBA', base.size, (0,0,0,0))
        overlay.paste(support_line_crop, (new_x, new_y), support_line_crop)
        if alpha < 1.0:
            alpha_mask = overlay.split()[-1].point(lambda p: int(p * alpha))
            overlay.putalpha(alpha_mask)
        combined = Image.alpha_composite(base, overlay).convert('RGB')
        combined.save(output_path) 