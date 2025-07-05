#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試 OCR 資料內容
"""

import json
from gemini_ocr_processor import GeminiOCRProcessor
from data_processor import DataProcessor
from word_template_processor_pure import WordTemplateProcessorPure

def test_ocr_data():
    """測試 OCR 資料處理"""
    print("🔍 測試 OCR 資料處理")
    print("=" * 50)
    
    # 初始化處理器
    ocr_processor = GeminiOCRProcessor()
    data_processor = DataProcessor()
    word_processor = WordTemplateProcessorPure()
    
    # 使用最新的 OCR 結果
    ocr_file = "ocr_results/gemini_ocr_output_20250705_061913.json"
    
    try:
        # 讀取 OCR 結果
        with open(ocr_file, 'r', encoding='utf-8') as f:
            raw_data = json.load(f)
        
        print(f"📄 原始 OCR 資料:")
        print("-" * 30)
        for key, value in raw_data.items():
            print(f"  {key}: {value}")
        
        # 資料處理
        processed_data = data_processor.process_insurance_data(raw_data)
        
        print(f"\n🔄 處理後的資料:")
        print("-" * 30)
        for key, value in processed_data.items():
            print(f"  {key}: {value}")
        
        # 測試 Word 模板處理
        print(f"\n📝 Word 模板處理:")
        print("-" * 30)
        word_data = word_processor.process_ocr_data(processed_data)
        
        print(f"Word 模板資料:")
        for key, value in word_data.items():
            if key not in ['watermark_name_blue', 'watermark_company_blue']:
                print(f"  {key}: {value}")
        
        # 生成 Word 檔案
        output_path = "test_outputs/test_ocr_data.docx"
        result = word_processor.fill_template(processed_data, output_path)
        
        if result:
            print(f"\n✅ Word 檔案生成成功: {result}")
        else:
            print(f"\n❌ Word 檔案生成失敗")
            
    except Exception as e:
        print(f"❌ 測試失敗: {str(e)}")

if __name__ == "__main__":
    test_ocr_data() 