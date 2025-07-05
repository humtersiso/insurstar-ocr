#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批次處理腳本
處理所有轉換後的圖片並生成財產分析書
"""

import os
import glob
from pdf_filler import PDFFiller

def main():
    """主函數"""
    print("🚀 批次處理財產分析書生成")
    print("=" * 60)
    
    # 初始化 PDF 填寫系統
    pdf_filler = PDFFiller()
    
    # 取得所有轉換後的圖片
    image_extensions = ['*.png', '*.jpg', '*.jpeg', '*.bmp', '*.tiff']
    image_paths = []
    
    for ext in image_extensions:
        pattern = os.path.join('temp_images', ext)
        image_paths.extend(glob.glob(pattern))
    
    if not image_paths:
        print("❌ 沒有找到轉換後的圖片")
        print("請先執行 pdf_to_images.py 將 PDF 轉換為圖片")
        return
    
    print(f"📸 找到 {len(image_paths)} 個圖片檔案")
    
    # 顯示找到的圖片
    for i, path in enumerate(image_paths, 1):
        filename = os.path.basename(path)
        print(f"  {i}. {filename}")
    
    print("\n" + "=" * 60)
    
    # 批次處理
    results = pdf_filler.batch_process_documents(image_paths)
    
    # 顯示處理摘要
    summary = pdf_filler.get_processing_summary(results)
    
    print("\n📊 批次處理摘要")
    print("=" * 60)
    print(f"總文件數: {summary['total_documents']}")
    print(f"成功處理: {summary['successful_documents']}")
    print(f"處理失敗: {summary['failed_documents']}")
    print(f"成功率: {summary['success_rate']}")
    print(f"平均辨識率: {summary['average_extraction_rate']}")
    
    if summary['generated_pdfs']:
        print(f"\n📄 生成的 PDF 檔案:")
        for pdf in summary['generated_pdfs']:
            print(f"  ✅ {pdf}")
    
    if summary['errors']:
        print(f"\n❌ 錯誤訊息:")
        for error in summary['errors']:
            print(f"  - {error}")
    
    print(f"\n🎉 批次處理完成！")
    print(f"📁 所有 PDF 檔案已儲存至 property_reports/ 目錄")

if __name__ == "__main__":
    main() 