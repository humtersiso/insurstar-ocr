#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 整合功能測試
測試完整的 OCR + Word 生成流程
"""

import os
import json
import requests
from datetime import datetime

def test_word_generation():
    """測試 Word 生成功能"""
    print("🧪 Word 整合功能測試")
    print("=" * 50)
    
    # 檢查測試圖片
    test_images = []
    for ext in ['jpg', 'jpeg', 'png', 'bmp', 'tiff']:
        for file in os.listdir('temp_images'):
            if file.lower().endswith(f'.{ext}'):
                test_images.append(os.path.join('temp_images', file))
    
    if not test_images:
        print("❌ 沒有找到測試圖片")
        return
    
    # 選擇第一個圖片進行測試
    test_image = test_images[0]
    print(f"📸 測試圖片: {test_image}")
    
    # 檢查 Web 服務是否運行
    try:
        response = requests.get('http://localhost:5000/api/health', timeout=5)
        if response.status_code != 200:
            print("❌ Web 服務未運行，請先啟動 app.py")
            return
        print("✅ Web 服務正常運行")
    except Exception as e:
        print(f"❌ 無法連接到 Web 服務: {str(e)}")
        print("請先執行: python app.py")
        return
    
    # 測試 Word 生成 API
    print("\n📄 測試 Word 生成 API...")
    
    try:
        # 準備請求資料
        payload = {
            'image_path': test_image
        }
        
        # 發送請求
        response = requests.post(
            'http://localhost:5000/api/generate-word',
            json=payload,
            timeout=60  # 設定較長的超時時間
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print("✅ Word 生成成功！")
                print(f"📄 檔案名稱: {result['word_filename']}")
                print(f"📁 檔案路徑: {result['word_path']}")
                print(f"🔗 下載連結: {result['download_url']}")
                
                # 檢查檔案是否存在
                if os.path.exists(result['word_path']):
                    file_size = os.path.getsize(result['word_path'])
                    print(f"📊 檔案大小: {file_size:,} bytes")
                    print("✅ 檔案已成功生成並保存")
                else:
                    print("❌ 檔案未找到")
                
                # 顯示資料摘要
                if 'data_summary' in result:
                    summary = result['data_summary']
                    print(f"📊 資料摘要: 完成率 {summary.get('completion_rate', 'N/A')}")
                
            else:
                print(f"❌ Word 生成失敗: {result.get('error', '未知錯誤')}")
        else:
            print(f"❌ API 請求失敗: {response.status_code}")
            print(f"錯誤訊息: {response.text}")
            
    except Exception as e:
        print(f"❌ 測試失敗: {str(e)}")

def test_batch_processing():
    """測試批次處理功能"""
    print("\n🚀 測試批次處理功能...")
    
    try:
        from word_filler import WordFiller
        
        # 初始化 Word 填寫系統
        word_filler = WordFiller()
        
        # 取得所有測試圖片
        test_images = []
        for ext in ['jpg', 'jpeg', 'png', 'bmp', 'tiff']:
            for file in os.listdir('temp_images'):
                if file.lower().endswith(f'.{ext}'):
                    test_images.append(os.path.join('temp_images', file))
        
        if len(test_images) > 3:
            test_images = test_images[:3]  # 只測試前3個檔案
        
        print(f"📸 批次處理 {len(test_images)} 個檔案...")
        
        # 執行批次處理
        results = word_filler.batch_process_documents(test_images)
        
        # 顯示結果摘要
        summary = word_filler.get_processing_summary(results)
        
        print(f"📊 批次處理結果:")
        print(f"   - 總檔案數: {summary['total_documents']}")
        print(f"   - 成功處理: {summary['successful_documents']}")
        print(f"   - 處理失敗: {summary['failed_documents']}")
        print(f"   - 成功率: {summary['success_rate']}")
        print(f"   - 平均辨識率: {summary['average_extraction_rate']}")
        
        if summary['generated_files']:
            print(f"   - 生成的 Word 檔案:")
            for filename in summary['generated_files']:
                print(f"     • {filename}")
        
        if summary['errors']:
            print(f"   - 錯誤訊息:")
            for error in summary['errors']:
                print(f"     • {error}")
        
    except Exception as e:
        print(f"❌ 批次處理測試失敗: {str(e)}")

def main():
    """主測試函數"""
    print("🎯 Word 整合功能完整測試")
    print("=" * 60)
    
    # 1. 測試單一檔案 Word 生成
    test_word_generation()
    
    # 2. 測試批次處理
    test_batch_processing()
    
    print("\n🎉 測試完成！")
    print("💡 提示:")
    print("   - 請用 Microsoft Word 或 LibreOffice 開啟生成的 .docx 檔案")
    print("   - 檢查中文顯示是否正常")
    print("   - 確認資料填入是否正確")

if __name__ == "__main__":
    main() 