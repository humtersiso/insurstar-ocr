#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 轉圖片工具
將保單 PDF 轉換為高品質圖片供 Gemini OCR 使用
"""

import os
import fitz  # PyMuPDF
from PIL import Image
import logging
from typing import List, Tuple, Dict, Optional

# 設定日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PDFToImageConverter:
    """PDF 轉圖片轉換器"""
    
    def __init__(self, output_dir: str = "converted_images"):
        """
        初始化轉換器
        
        Args:
            output_dir: 輸出目錄
        """
        self.output_dir = output_dir
        self.ensure_output_dir()
    
    def ensure_output_dir(self):
        """確保輸出目錄存在"""
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            logger.info(f"建立輸出目錄: {self.output_dir}")
    
    def convert_pdf_to_images(self, pdf_path: str, dpi: int = 300, 
                            pages: Optional[List[int]] = None, 
                            output_format: str = "PNG") -> List[str]:
        """
        將 PDF 轉換為圖片
        
        Args:
            pdf_path: PDF 檔案路徑
            dpi: 解析度 (預設 300 DPI)
            pages: 要轉換的頁面列表 (None 表示全部)
            output_format: 輸出格式 (PNG, JPEG, TIFF)
            
        Returns:
            生成的圖片檔案路徑列表
        """
        if not os.path.exists(pdf_path):
            logger.error(f"PDF 檔案不存在: {pdf_path}")
            return []
        
        try:
            # 開啟 PDF
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            
            logger.info(f"PDF 總頁數: {total_pages}")
            
            # 決定要轉換的頁面
            if pages is None:
                pages = list(range(total_pages))
            else:
                pages = [p for p in pages if 0 <= p < total_pages]
            
            if not pages:
                logger.error("沒有有效的頁面要轉換")
                return []
            
            # 取得 PDF 檔名（不含副檔名）
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            
            # 轉換每一頁
            image_paths = []
            for page_num in pages:
                try:
                    # 取得頁面
                    page = doc.load_page(page_num)
                    
                    # 設定縮放矩陣 (DPI 轉換)
                    mat = fitz.Matrix(dpi/72, dpi/72)
                    
                    # 渲染頁面為圖片
                    pix = page.get_pixmap(matrix=mat)
                    
                    # 建立輸出檔名
                    output_filename = f"{pdf_name}_page_{page_num + 1:02d}.{output_format.lower()}"
                    output_path = os.path.join(self.output_dir, output_filename)
                    
                    # 儲存圖片
                    pix.save(output_path)
                    
                    logger.info(f"已轉換頁面 {page_num + 1}: {output_filename}")
                    image_paths.append(output_path)
                    
                except Exception as e:
                    logger.error(f"轉換頁面 {page_num + 1} 失敗: {str(e)}")
            
            doc.close()
            return image_paths
            
        except Exception as e:
            logger.error(f"PDF 轉換失敗: {str(e)}")
            return []
    
    def convert_all_pdfs_in_directory(self, pdf_dir: str, dpi: int = 300,
                                    pages: Optional[List[int]] = None,
                                    output_format: str = "PNG") -> Dict[str, List[str]]:
        """
        轉換目錄中所有 PDF 檔案
        
        Args:
            pdf_dir: PDF 目錄路徑
            dpi: 解析度
            pages: 要轉換的頁面
            output_format: 輸出格式
            
        Returns:
            每個 PDF 對應的圖片路徑字典
        """
        if not os.path.exists(pdf_dir):
            logger.error(f"PDF 目錄不存在: {pdf_dir}")
            return {}
        
        # 尋找所有 PDF 檔案
        pdf_files = []
        for file in os.listdir(pdf_dir):
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(pdf_dir, file))
        
        if not pdf_files:
            logger.warning(f"在 {pdf_dir} 中沒有找到 PDF 檔案")
            return {}
        
        logger.info(f"找到 {len(pdf_files)} 個 PDF 檔案")
        
        # 轉換每個 PDF
        results = {}
        for pdf_path in pdf_files:
            pdf_name = os.path.basename(pdf_path)
            logger.info(f"處理 PDF: {pdf_name}")
            
            image_paths = self.convert_pdf_to_images(
                pdf_path, dpi, pages, output_format
            )
            
            results[pdf_name] = image_paths
        
        return results
    
    def optimize_image_for_ocr(self, image_path: str, 
                             target_dpi: int = 300,
                             enhance_contrast: bool = True) -> str:
        """
        優化圖片以提升 OCR 效果
        
        Args:
            image_path: 圖片路徑
            target_dpi: 目標解析度
            enhance_contrast: 是否增強對比度
            
        Returns:
            優化後的圖片路徑
        """
        try:
            # 開啟圖片
            image = Image.open(image_path)
            
            # 轉換為 RGB 模式（如果不是的話）
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # 增強對比度
            if enhance_contrast:
                from PIL import ImageEnhance
                enhancer = ImageEnhance.Contrast(image)
                image = enhancer.enhance(1.2)  # 增強 20%
            
            # 建立優化後的檔名
            base_name = os.path.splitext(image_path)[0]
            optimized_path = f"{base_name}_optimized.png"
            
            # 儲存優化後的圖片
            image.save(optimized_path, 'PNG', dpi=(target_dpi, target_dpi))
            
            logger.info(f"圖片已優化: {optimized_path}")
            return optimized_path
            
        except Exception as e:
            logger.error(f"圖片優化失敗: {str(e)}")
            return image_path

def main():
    """主函數"""
    print("📄 PDF 轉圖片工具")
    print("=" * 50)
    
    # 初始化轉換器
    converter = PDFToImageConverter()
    
    # 檢查 insurtech_data 目錄
    pdf_dir = "insurtech_data"
    if not os.path.exists(pdf_dir):
        print(f"❌ 找不到 {pdf_dir} 目錄")
        return
    
    # 轉換所有 PDF
    print(f"🔍 掃描 {pdf_dir} 目錄中的 PDF 檔案...")
    results = converter.convert_all_pdfs_in_directory(
        pdf_dir, 
        dpi=300,  # 高解析度
        pages=[0],  # 只轉換第一頁
        output_format="PNG"
    )
    
    if not results:
        print("❌ 沒有找到 PDF 檔案或轉換失敗")
        return
    
    # 顯示結果
    print(f"\n✅ 轉換完成！")
    total_images = sum(len(paths) for paths in results.values())
    print(f"📊 總共轉換了 {len(results)} 個 PDF，產生 {total_images} 個圖片")
    
    print(f"\n📁 圖片儲存在: {converter.output_dir}")
    
    # 列出生成的圖片
    print(f"\n📋 生成的圖片:")
    for pdf_name, image_paths in results.items():
        print(f"  📄 {pdf_name}:")
        for image_path in image_paths:
            print(f"    - {os.path.basename(image_path)}")
    
    print(f"\n💡 現在可以使用以下指令測試 Gemini OCR:")
    print(f"   python test_gemini_ocr.py")

if __name__ == "__main__":
    main() 