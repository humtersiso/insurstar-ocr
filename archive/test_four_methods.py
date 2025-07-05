#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試四種不同的標記填入方式
比較原始模板 vs 修復模板，以及 word_template_processor 處理 vs 直接填入
"""

import os
import json
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from word_template_processor_pure import WordTemplateProcessor

def load_ocr_data(json_path):
    """載入 OCR 結果資料"""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def create_context_from_ocr(ocr_data, doc=None):
    """從 OCR 資料建立 context"""
    # 先包含所有 OCR 原始資料（像 test_pure_docxtpl.py 一樣）
    context = dict(ocr_data)
    
    # 補齊 PCN 編號
    context['PCN'] = f"PCN-{ocr_data.get('license_number', '')}-{datetime.now().strftime('%Y%m%d')}"
    
    # 只有當 doc 存在時才加入 InlineImage
    if doc is not None:
        context['watermark_name_blue'] = InlineImage(
            doc, 
            "assets/watermark/watermark_name_blue.png",
            width=150
        )
        context['watermark_company_blue'] = InlineImage(
            doc, 
            "assets/watermark/watermark_company_blue.png",
            width=150
        )
    else:
        # 如果沒有 doc，就給圖片路徑
        context['watermark_name_blue'] = "assets/watermark/watermark_name_blue.png"
        context['watermark_company_blue'] = "assets/watermark/watermark_company_blue.png"
    
    return context

def method_1_original_template_direct():
    """方法1: 讀取原始模板，直接填入標記"""
    print("🔧 方法1: 原始模板 + 直接填入")
    print("-" * 50)
    
    template_path = "assets/templates/財產分析書.docx"
    output_path = "test_outputs/method1_original_direct.docx"
    
    try:
        # 載入模板
        doc = DocxTemplate(template_path)
        
        # 建立 context (傳入 doc 實例)
        context = create_context_from_ocr(ocr_data, doc)
        
        # 直接填入
        doc.render(context)
        
        # 儲存
        os.makedirs("test_outputs", exist_ok=True)
        doc.save(output_path)
        print(f"✅ 成功: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ 失敗: {str(e)}")
        return False

def method_2_fixed_template_processor():
    """方法2: 讀取修復模板，使用 word_template_processor 處理"""
    print("\n🔧 方法2: 修復模板 + word_template_processor")
    print("-" * 50)
    
    template_path = "assets/templates/財產分析書_fixed.docx"
    output_path = "test_outputs/method2_fixed_processor.docx"
    
    try:
        # 使用 word_template_processor
        processor = WordTemplateProcessor(template_path)
        
        # 直接填入 OCR 資料，processor 會自動處理
        result = processor.fill_template(ocr_data, output_path)
        
        if result:
            print(f"✅ 成功: {output_path}")
            return True
        else:
            print(f"❌ 失敗: processor.fill_template 返回 None")
            return False
        
    except Exception as e:
        print(f"❌ 失敗: {str(e)}")
        return False

def method_3_fixed_template_direct():
    """方法3: 讀取修復模板，直接填入標記"""
    print("\n🔧 方法3: 修復模板 + 直接填入")
    print("-" * 50)
    
    template_path = "assets/templates/財產分析書_fixed.docx"
    output_path = "test_outputs/method3_fixed_direct.docx"
    
    try:
        # 載入模板
        doc = DocxTemplate(template_path)
        
        # 建立 context (傳入 doc 實例)
        context = create_context_from_ocr(ocr_data, doc)
        
        # 直接填入
        doc.render(context)
        
        # 儲存
        os.makedirs("test_outputs", exist_ok=True)
        doc.save(output_path)
        print(f"✅ 成功: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ 失敗: {str(e)}")
        return False

def method_4_original_template_processor():
    """方法4: 讀取原始模板，使用 word_template_processor 處理"""
    print("\n🔧 方法4: 原始模板 + word_template_processor")
    print("-" * 50)
    
    template_path = "assets/templates/財產分析書.docx"
    output_path = "test_outputs/method4_original_processor.docx"
    
    try:
        # 使用 word_template_processor
        processor = WordTemplateProcessor(template_path)
        
        # 直接填入 OCR 資料，processor 會自動處理
        result = processor.fill_template(ocr_data, output_path)
        
        if result:
            print(f"✅ 成功: {output_path}")
            return True
        else:
            print(f"❌ 失敗: processor.fill_template 返回 None")
            return False
        
    except Exception as e:
        print(f"❌ 失敗: {str(e)}")
        return False

def analyze_template_markers(template_path):
    """分析模板中的標記"""
    print(f"\n📊 分析模板標記: {template_path}")
    print("-" * 30)
    
    try:
        doc = DocxTemplate(template_path)
        
        # 取得所有段落文字
        all_text = ""
        if doc.docx and doc.docx.paragraphs:
            for paragraph in doc.docx.paragraphs:
                all_text += paragraph.text + "\n"
        
        # 取得所有表格文字
        if doc.docx and doc.docx.tables:
            for table in doc.docx.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            all_text += paragraph.text + "\n"
        
        # 找出所有標記
        import re
        markers = re.findall(r'\{\{[^}]+\}\}', all_text)
        
        print(f"找到 {len(markers)} 個標記:")
        for marker in sorted(set(markers)):
            print(f"  - {marker}")
        
        # 檢查新標記
        new_markers = ['{{watermark_name_blue}}', '{{watermark_company_blue}}', '{{PCN}}']
        for marker in new_markers:
            if marker in markers:
                print(f"✅ {marker} - 存在")
            else:
                print(f"❌ {marker} - 不存在")
        
        # 檢查是否有被拆開的標記
        print(f"\n🔍 檢查被拆開的標記:")
        if "{{P" in all_text and "CN}}" in all_text:
            print(f"⚠️ 發現被拆開的 PCN 標記: {{P + CN}}")
        if "{{watermark_" in all_text and "name_blue}}" in all_text:
            print(f"⚠️ 發現被拆開的 watermark_name_blue 標記")
        if "{{watermark_" in all_text and "company_blue}}" in all_text:
            print(f"⚠️ 發現被拆開的 watermark_company_blue 標記")
        
        # 顯示包含這些標記的上下文
        print(f"\n📝 包含標記的上下文:")
        lines = all_text.split('\n')
        for i, line in enumerate(lines):
            if any(marker.replace('{{', '').replace('}}', '') in line for marker in new_markers):
                print(f"  第{i+1}行: {line.strip()}")
            elif "{{P" in line or "CN}}" in line:
                print(f"  第{i+1}行: {line.strip()}")
            elif "{{watermark_" in line:
                print(f"  第{i+1}行: {line.strip()}")
        
        return markers
        
    except Exception as e:
        print(f"❌ 分析失敗: {str(e)}")
        return []

def main():
    """主函數"""
    print("🧪 四種標記填入方式測試")
    print("=" * 80)
    
    # 載入 OCR 資料
    global ocr_data
    ocr_path = "ocr_results/gemini_ocr_output_20250704_144524.json"
    ocr_data = load_ocr_data(ocr_path)
    print(f"📄 載入 OCR 資料: {ocr_path}")
    
    # 分析模板標記
    print("\n" + "="*80)
    analyze_template_markers("assets/templates/財產分析書.docx")
    analyze_template_markers("assets/templates/財產分析書_fixed.docx")
    
    # 執行四種方法
    print("\n" + "="*80)
    results = []
    
    results.append(("方法1: 原始模板 + 直接填入", method_1_original_template_direct()))
    results.append(("方法2: 修復模板 + word_template_processor", method_2_fixed_template_processor()))
    results.append(("方法3: 修復模板 + 直接填入", method_3_fixed_template_direct()))
    results.append(("方法4: 原始模板 + word_template_processor", method_4_original_template_processor()))
    
    # 總結結果
    print("\n" + "="*80)
    print("📊 測試結果總結:")
    print("-" * 50)
    
    for method_name, success in results:
        status = "✅ 成功" if success else "❌ 失敗"
        print(f"{method_name}: {status}")
    
    print(f"\n📁 輸出檔案位置: test_outputs/")
    print("請檢查各方法的輸出檔案，比較差異和問題。")

if __name__ == "__main__":
    main() 