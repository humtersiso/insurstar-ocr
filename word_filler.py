#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 填寫系統
整合 Gemini OCR 辨識結果與 Word 模板生成
"""

import os
import json
import uuid
from datetime import datetime
from typing import Dict, List, Any, Optional

from gemini_ocr_processor import GeminiOCRProcessor
from data_processor import DataProcessor
from word_template_generator import WordTemplateGenerator

class WordFiller:
    """Word 填寫系統"""
    
    def __init__(self):
        """初始化 Word 填寫系統"""
        self.ocr_processor = GeminiOCRProcessor()
        self.data_processor = DataProcessor()
        self.word_generator = WordTemplateGenerator()
        
        # 建立輸出目錄
        self.output_dir = 'outputs'
        os.makedirs(self.output_dir, exist_ok=True)
    
    def process_insurance_document(self, image_path: str) -> Dict:
        """
        處理保險文件：OCR 辨識 + Word 生成
        
        Args:
            image_path: 圖片路徑
            
        Returns:
            處理結果字典
        """
        try:
            print(f"🔍 開始處理文件: {image_path}")
            
            # 1. OCR 辨識
            print("📝 執行 OCR 辨識...")
            raw_data = self.ocr_processor.extract_insurance_data_with_gemini(image_path)
            
            if not raw_data:
                return {
                    'success': False,
                    'error': 'OCR 辨識失敗，無法提取資料'
                }
            
            print("✅ OCR 辨識完成")
            print(f"📊 原始資料: {json.dumps(raw_data, ensure_ascii=False, indent=2)}")
            
            # 2. 資料處理
            print("🔄 處理資料...")
            processed_data = self.data_processor.process_insurance_data(raw_data)
            validation_result = self.data_processor.validate_processed_data(processed_data)
            
            print("✅ 資料處理完成")
            print(f"📊 處理後資料: {json.dumps(processed_data, ensure_ascii=False, indent=2)}")
            
            # 3. 生成 Word 檔案
            print("📄 生成 Word 檔案...")
            file_id = str(uuid.uuid4())
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            word_filename = f"property_analysis_{file_id}_{timestamp}.docx"
            word_path = os.path.join(self.output_dir, word_filename)
            
            word_result = self.word_generator.create_property_analysis_docx(processed_data, word_path)
            
            if not word_result:
                return {
                    'success': False,
                    'error': 'Word 檔案生成失敗'
                }
            
            print("✅ Word 檔案生成完成")
            
            # 4. 準備回應資料
            result = {
                'success': True,
                'file_id': file_id,
                'original_image': image_path,
                'word_path': word_path,
                'word_filename': word_filename,
                'raw_data': raw_data,
                'processed_data': processed_data,
                'validation_result': validation_result,
                'data_summary': self.data_processor.get_data_summary(processed_data),
                'processing_time': datetime.now().isoformat(),
                'download_url': f"/download/{word_filename}"
            }
            
            print("🎉 文件處理完成！")
            return result
            
        except Exception as e:
            print(f"❌ 處理失敗: {str(e)}")
            return {
                'success': False,
                'error': f'處理失敗: {str(e)}'
            }
    
    def batch_process_documents(self, image_paths: List[str]) -> List[Dict]:
        """
        批次處理多個文件
        
        Args:
            image_paths: 圖片路徑列表
            
        Returns:
            處理結果列表
        """
        results = []
        
        print(f"🚀 開始批次處理 {len(image_paths)} 個文件")
        
        for i, image_path in enumerate(image_paths, 1):
            print(f"\n📄 處理第 {i}/{len(image_paths)} 個文件: {image_path}")
            
            result = self.process_insurance_document(image_path)
            results.append(result)
            
            if result['success']:
                print(f"✅ 第 {i} 個文件處理成功")
            else:
                print(f"❌ 第 {i} 個文件處理失敗: {result.get('error', '未知錯誤')}")
        
        print(f"\n🎉 批次處理完成！成功: {sum(1 for r in results if r['success'])}/{len(results)}")
        return results
    
    def get_processing_summary(self, results: List[Dict]) -> Dict:
        """
        取得處理摘要
        
        Args:
            results: 處理結果列表
            
        Returns:
            處理摘要
        """
        total = len(results)
        successful = sum(1 for r in results if r['success'])
        failed = total - successful
        
        # 統計欄位辨識率
        total_fields = 0
        filled_fields = 0
        
        for result in results:
            if result['success']:
                summary = result.get('data_summary', {})
                total_fields += summary.get('total_fields', 0)
                filled_fields += summary.get('filled_fields', 0)
        
        avg_extraction_rate = (filled_fields / total_fields * 100) if total_fields > 0 else 0
        
        return {
            'total_documents': total,
            'successful_documents': successful,
            'failed_documents': failed,
            'success_rate': f"{successful / total * 100:.1f}%" if total > 0 else "0%",
            'average_extraction_rate': f"{avg_extraction_rate:.1f}%",
            'generated_files': [r['word_filename'] for r in results if r['success']],
            'errors': [r['error'] for r in results if not r['success']]
        }

def main():
    """測試函數"""
    print("📄 Word 填寫系統測試")
    print("=" * 50)
    
    # 初始化系統
    word_filler = WordFiller()
    
    # 檢查測試圖片
    test_images = []
    for ext in ['jpg', 'jpeg', 'png', 'bmp', 'tiff']:
        for file in os.listdir('converted_images'):
            if file.lower().endswith(f'.{ext}'):
                test_images.append(os.path.join('converted_images', file))
    
    if not test_images:
        print("❌ 沒有找到測試圖片")
        return
    
    # 選擇第一個圖片進行測試
    test_image = test_images[0]
    print(f"📸 測試圖片: {test_image}")
    
    # 處理文件
    result = word_filler.process_insurance_document(test_image)
    
    if result['success']:
        print(f"✅ 處理成功！")
        print(f"📄 Word 檔案: {result['word_filename']}")
        print(f"📊 資料摘要: {result['data_summary']}")
        print(f"🔍 驗證結果: {result['validation_result']}")
        print(f"📝 請用 Microsoft Word 或 LibreOffice 開啟檢查內容")
    else:
        print(f"❌ 處理失敗: {result['error']}")

if __name__ == "__main__":
    main() 