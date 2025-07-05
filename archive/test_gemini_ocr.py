#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試優化後的 Gemini OCR 處理器
"""

import os
import json
from datetime import datetime
from gemini_ocr_processor import GeminiOCRProcessor

def test_optimized_gemini_ocr():
    """測試優化後的 Gemini OCR"""
    print("🔍 測試優化後的 Gemini OCR 處理器")
    print("=" * 60)
    
    # 初始化處理器
    processor = GeminiOCRProcessor()
    
    # 檢查轉換後的圖片
    converted_dir = "temp_images"
    if not os.path.exists(converted_dir):
        print(f"❌ 找不到 {converted_dir} 目錄")
        return
    
    # 取得所有 PNG 圖片
    png_files = [f for f in os.listdir(converted_dir) if f.endswith('.png')]
    if not png_files:
        print(f"❌ 在 {converted_dir} 中沒有找到 PNG 圖片")
        return
    
    print(f"📁 找到 {len(png_files)} 個測試圖片")
    
    # 測試結果
    results = []
    total_fields = 20  # 新的欄位總數（包含產險公司、強制險保期、任意險保期）
    
    for i, png_file in enumerate(png_files, 1):
        print(f"\n📄 測試 {i}/{len(png_files)}: {png_file}")
        print("-" * 40)
        
        image_path = os.path.join(converted_dir, png_file)
        
        try:
            # 提取保單資料
            print("🔄 正在提取保單資料...")
            insurance_data = processor.extract_insurance_data_with_gemini(image_path)
            
            if insurance_data:
                # 計算提取率
                filled_fields = sum(1 for v in insurance_data.values() if v and v != "")
                extraction_rate = (filled_fields / total_fields) * 100
                
                # 檢查特殊欄位
                birth_date_status = "✅ 有資料" if insurance_data.get('birth_date') and insurance_data.get('birth_date') != "無填寫" else "❌ 無填寫"
                gender_status = "✅ 有資料" if insurance_data.get('gender') and insurance_data.get('gender') != "無填寫" else "❌ 無填寫"
                insurance_company_status = "✅ 有資料" if insurance_data.get('insurance_company') and insurance_data.get('insurance_company') != "無填寫" else "❌ 無資料"
                coverage_items_status = "✅ 有資料" if insurance_data.get('coverage_items') and insurance_data.get('coverage_items') != "無填寫" else "❌ 無資料"
                
                result = {
                    'file': png_file,
                    'extraction_rate': f"{extraction_rate:.1f}%",
                    'filled_fields': filled_fields,
                    'total_fields': total_fields,
                    'birth_date_status': birth_date_status,
                    'gender_status': gender_status,
                    'insurance_company_status': insurance_company_status,
                    'coverage_items_status': coverage_items_status,
                    'insurance_company': insurance_data.get('insurance_company', ''),
                    'birth_date': insurance_data.get('birth_date', ''),
                    'gender': insurance_data.get('gender', ''),
                    'compulsory_insurance_period': insurance_data.get('compulsory_insurance_period', ''),
                    'optional_insurance_period': insurance_data.get('optional_insurance_period', ''),
                    'coverage_items': insurance_data.get('coverage_items', ''),
                    'data': insurance_data
                }
                
                results.append(result)
                
                print(f"✅ 提取完成 - 提取率: {extraction_rate:.1f}%")
                print(f"🏢 產險公司: {insurance_data.get('insurance_company', '無資料')}")
                print(f"📅 出生日期: {birth_date_status}")
                print(f"👤 性別: {gender_status}")
                print(f"🗓️ 強制險保期: {insurance_data.get('compulsory_insurance_period', '無資料')}")
                print(f"🗓️ 任意險保期: {insurance_data.get('optional_insurance_period', '無資料')}")
                # 顯示 coverage_items 前2筆摘要
                items = insurance_data.get('coverage_items', [])
                if isinstance(items, list):
                    for idx, item in enumerate(items[:2], 1):
                        print(f"🛡️ 承保內容{idx}: {item}")
                else:
                    print(f"🛡️ 承保內容: {str(items)[:100]}...")
                
            else:
                print("❌ 提取失敗")
                results.append({
                    'file': png_file,
                    'extraction_rate': "0%",
                    'filled_fields': 0,
                    'total_fields': total_fields,
                    'error': '提取失敗'
                })
                
        except Exception as e:
            print(f"❌ 處理錯誤: {str(e)}")
            results.append({
                'file': png_file,
                'extraction_rate': "0%",
                'filled_fields': 0,
                'total_fields': total_fields,
                'error': str(e)
            })
    
    # 生成報告
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_file = f"gemini_ocr_test_results_{timestamp}.json"
    
    # 計算統計資料
    successful_tests = [r for r in results if 'error' not in r]
    if successful_tests:
        avg_extraction_rate = sum(float(r['extraction_rate'].rstrip('%')) for r in successful_tests) / len(successful_tests)
        avg_filled_fields = sum(r['filled_fields'] for r in successful_tests) / len(successful_tests)
        
        # 統計特殊欄位
        birth_date_success = sum(1 for r in successful_tests if '有資料' in r['birth_date_status'])
        gender_success = sum(1 for r in successful_tests if '有資料' in r['gender_status'])
        insurance_company_success = sum(1 for r in successful_tests if '有資料' in r['insurance_company_status'])
        coverage_items_success = sum(1 for r in successful_tests if '有資料' in r['coverage_items_status'])
        
        summary = {
            'test_timestamp': timestamp,
            'total_tests': len(results),
            'successful_tests': len(successful_tests),
            'failed_tests': len(results) - len(successful_tests),
            'average_extraction_rate': f"{avg_extraction_rate:.1f}%",
            'average_filled_fields': f"{avg_filled_fields:.1f}",
            'birth_date_success_rate': f"{birth_date_success}/{len(successful_tests)} ({birth_date_success/len(successful_tests)*100:.1f}%)",
            'gender_success_rate': f"{gender_success}/{len(successful_tests)} ({gender_success/len(successful_tests)*100:.1f}%)",
            'insurance_company_success_rate': f"{insurance_company_success}/{len(successful_tests)} ({insurance_company_success/len(successful_tests)*100:.1f}%)",
            'coverage_items_success_rate': f"{coverage_items_success}/{len(successful_tests)} ({coverage_items_success/len(successful_tests)*100:.1f}%)",
            'results': results
        }
    else:
        summary = {
            'test_timestamp': timestamp,
            'total_tests': len(results),
            'successful_tests': 0,
            'failed_tests': len(results),
            'error': '所有測試都失敗'
        }
    
    # 儲存報告
    with open(report_file, 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    
    # 顯示總結
    print(f"\n📊 測試總結")
    print("=" * 60)
    print(f"📁 測試檔案: {report_file}")
    print(f"📄 總測試數: {len(results)}")
    print(f"✅ 成功數: {len(successful_tests) if 'successful_tests' in summary else 0}")
    print(f"❌ 失敗數: {summary.get('failed_tests', len(results))}")
    
    if successful_tests:
        print(f"📈 平均提取率: {summary['average_extraction_rate']}")
        print(f"📊 平均填寫欄位: {summary['average_filled_fields']}")
        print(f"📅 出生日期成功率: {summary['birth_date_success_rate']}")
        print(f"👤 性別成功率: {summary['gender_success_rate']}")
        print(f"🏢 產險公司成功率: {summary['insurance_company_success_rate']}")
        print(f"🛡️ 承保內容成功率: {summary['coverage_items_success_rate']}")
    
    print(f"\n💾 詳細結果已儲存至: {report_file}")
    
    return summary

if __name__ == "__main__":
    test_optimized_gemini_ocr() 