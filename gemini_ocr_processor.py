#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gemini 2.5 Flash OCR 處理器
使用 Google Gemini API 進行保單文字辨識
"""

import google.generativeai as genai
import base64
import json
import logging
from typing import Dict, List, Tuple, Optional
import cv2
import numpy as np
from PIL import Image
import io

# 設定日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class GeminiOCRProcessor:
    """Gemini OCR 處理器"""
    
    def __init__(self, api_key: str = "AIzaSyBRSCd0Tke2rOobZOv-HsgDq81CQKER1ks"):
        """
        初始化 Gemini OCR 處理器
        
        Args:
            api_key: Gemini API Key
        """
        self.api_key = api_key
        genai.configure(api_key=api_key)
        
        # 初始化 Gemini 模型
        try:
            self.model = genai.GenerativeModel('gemini-2.0-flash-exp')
            logger.info("Gemini 2.0 Flash 模型初始化成功")
        except Exception as e:
            logger.error(f"Gemini 模型初始化失敗: {str(e)}")
            self.model = None
    
    def image_to_base64(self, image_path: str) -> str:
        """
        將圖片轉換為 base64 編碼
        
        Args:
            image_path: 圖片路徑
            
        Returns:
            base64 編碼的字串
        """
        try:
            with open(image_path, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
            return encoded_string
        except Exception as e:
            logger.error(f"圖片轉 base64 失敗: {str(e)}")
            return ""
    
    def extract_text_with_gemini(self, image_path: str) -> List[Tuple[str, float]]:
        """
        使用 Gemini 提取圖片中的文字
        
        Args:
            image_path: 圖片路徑
            
        Returns:
            包含文字和信心度的列表
        """
        if not self.model:
            logger.error("Gemini 模型未初始化")
            return []
        
        try:
            # 轉換圖片為 base64
            image_base64 = self.image_to_base64(image_path)
            if not image_base64:
                return []
            
            # 建立 Gemini 提示詞
            prompt = """
            請分析這張保單圖片，提取所有可見的文字內容。
            
            請按照以下格式回傳 JSON：
            {
                "texts": [
                    {
                        "text": "提取的文字內容",
                        "confidence": 0.95,
                        "position": "描述文字在圖片中的位置"
                    }
                ],
                "summary": "整體文字內容摘要"
            }
            
            注意事項：
            1. 只提取實際的文字內容，不要添加解釋
            2. confidence 是 0-1 之間的數值，表示對該文字的信心度
            3. position 描述文字在圖片中的大致位置（如：左上角、中間、右下角等）
            4. 如果某些文字模糊或無法辨識，請標註 confidence 較低
            """
            
            # 建立圖片數據
            image_data = {
                "mime_type": "image/png",
                "data": image_base64
            }
            
            # 呼叫 Gemini API
            response = self.model.generate_content([prompt, image_data])
            
            # 解析回應
            if response.text:
                try:
                    # 清理回應文字，移除可能的 markdown 標記
                    cleaned_text = response.text.strip()
                    if cleaned_text.startswith('```json'):
                        cleaned_text = cleaned_text[7:]  # 移除 ```json
                    if cleaned_text.endswith('```'):
                        cleaned_text = cleaned_text[:-3]  # 移除結尾 ```
                    cleaned_text = cleaned_text.strip()
                    
                    # 嘗試解析 JSON
                    result = json.loads(cleaned_text)
                    
                    # 預期欄位
                    expected_fields = [
                        'policy_number', 'insured_name', 'insurance_company',
                        'premium_amount', 'coverage_amount',
                        'start_date', 'end_date', 'phone_number', 'id_number'
                    ]
                    insurance_data = {field: result.get(field, '') for field in expected_fields}
                    
                    logger.info(f"Gemini 提取保單資料成功")
                    return insurance_data
                    
                except json.JSONDecodeError:
                    logger.error("Gemini 回應不是有效的 JSON 格式")
                    print("\n[DEBUG] Gemini 原始回應內容：")
                    print(response.text)
                    return {}
            
            return {}
            
        except Exception as e:
            logger.error(f"Gemini OCR 處理失敗: {str(e)}")
            return []
    
    def extract_insurance_data_with_gemini(self, image_path: str) -> Dict:
        """
        使用 Gemini 提取保單相關資料
        
        Args:
            image_path: 圖片路徑
            
        Returns:
            包含保單資料的字典
        """
        if not self.model:
            logger.error("Gemini 模型未初始化")
            return {}
        
        try:
            # 轉換圖片為 base64
            image_base64 = self.image_to_base64(image_path)
            if not image_base64:
                return {}
            
            # 建立專門的保單資料提取提示詞
            prompt = """
            請分析這張保單圖片，按照以下順序提取關鍵欄位的資訊：

            1. 產險公司 (insurance_company) - 位於文件最上方
            2. 被保險人區塊 (insured_section)
            3. 被保險人 (insured_person)
            4. 法人之代表人 (legal_representative)
            5. 身分證字號 (統一編號) (id_number)
            6. 出生日期 (birth_date)
            7. 性別 (gender)
            8. 要保人區塊 (注意虛線分隔) (policyholder_section)
            9. 要保人 (policyholder)
            10. 與被保險人關係 (relationship)
            11. 法人之代表人 (policyholder_legal_representative)
            12. 性別 (policyholder_gender)
            13. 要保人身份證字號/統編 (policyholder_id)
            14. 出生日期 (policyholder_birth_date)
            15. 車輛種類 (vehicle_type)
            16. 牌照號碼 (license_number)
            17. 承保內容 (coverage_items)：請將保單「承保內容」區塊每一條分別提取為五個欄位：保險代號、保險種類、保險金額、自負額、簽單保費。每一筆主項（有數字代號）為一筆，若有子項（如每一個人傷害、每一意外事故之傷害等，前面有大空格），請以 sub_items 陣列放在對應主項下。每一欄若值為空白就留空，若明確寫「無」就顯示「無」。

            在判斷承保內容每一個數值應該屬於哪個欄位時，請參考以下傳統表格分割的做法：
            - 如果圖片中有表格線（橫線、直線），請以這些表格線將表格分割成多個儲存格，並根據每個儲存格在表格中的位置，對應到正確的欄位（如「保險金額」、「自負額」、「簽單保費」等）。
            - 若表格線不明顯，請根據每一行的欄位標題（如「保險金額」、「自負額」、「簽單保費」等）來判斷每個數值應該填在哪個欄位。
            - 若同一行同時有多個數值，請依照它們在表格中的相對位置，分別對應到正確的欄位。
            - 請避免將「自負額」的數字填到「保險金額」欄位，或將「保險金額」的數字填到「自負額」欄位。
            - 請參考上下行的欄位標題與內容，確保每個值都對應到正確欄位。
            例如：
            - 「保險金額 40.2萬 自負額 20,000 簽單保費 20,527」這一行，請將 40.2萬 填到「保險金額」，20,000 填到「自負額」，20,527 填到「簽單保費」。
            - 若表格有明顯的欄位分隔線，請以分隔線為準，將每個儲存格的內容對應到正確欄位。

            18. 總保險費 (total_premium)

            特別注意：
            - 「保險金額」和「自負額」欄位都可以是數字（如：40.2萬、20,000）、金額單位，也可以是文字說明（如：「同竊盜損失險」、「同第三人責任險」等），請如實填寫，不要省略。
            - 請務必將所有承保內容主項都列出，不要因為欄位內容型態不同或內容為文字就省略該主項。
            - 「保險金額」欄位只填真正的保險金額（如：40.2萬、300萬、60萬等），不要填自負額或其他數字。
            - 「自負額」欄位只填自負額（如：20,000、10%、同竊盜損失險等），不要填保險金額。
            - 若主項有子項（如下方有每一個人傷害、每一意外事故之傷害、每一個人體傷、每一個人死亡或失能等），主項的「保險金額」請留空，金額請分別填在子項的「保險金額」欄位。
            - 只有當主項沒有子項時，主項的「保險金額」才填寫實際金額。
            - 子項的「保險金額」要正確填寫。
            - 若值為空白就留空，明確寫「無」就顯示「無」。
            - 要特別注意主項有子項，主項的「保險金額」請留空
            - 注意每個值的欄位位置，尤其是保險金額跟自負額的欄位位置，不要填錯，可以參考上下行

            範例：
            {
                "insurance_company": "產險公司名稱",
                ...
                "coverage_items": [
                  {
                    "保險代號": "05",
                    "保險種類": "車體損失保險乙式(Q)",
                    "保險金額": "",
                    "自負額": "20,000",
                    "簽單保費": "20,527",
                    "sub_items": [
                      { "保險種類": "每一個人傷害", "保險金額": "300萬", "自負額": "", "簽單保費": "9,485" },
                      { "保險種類": "每一意外事故之傷害", "保險金額": "600萬", "自負額": "", "簽單保費": "" }
                    ]
                  },
                  {
                    "保險代號": "11",
                    "保險種類": "汽車竊盜損失保險(Q)",
                    "保險金額": "40.2萬",
                    "自負額": "10%",
                    "簽單保費": "2,247",
                    "sub_items": []
                  },
                  {
                    "保險代號": "17",
                    "保險種類": "竊盜險全損免折舊附加條款",
                    "保險金額": "同竊盜損失險",
                    "自負額": "同竊盜損失險",
                    "簽單保費": "236",
                    "sub_items": []
                  }
                ],
                ...
            }

            注意事項：
            1. 如果某個欄位找不到，請填入空字串 ""
            2. 出生日期和性別欄位如果偵測到空白或無資料，請回傳「無填寫」
            3. 日期格式請保持原始格式
            4. 金額請包含單位（如：元、萬元等）
            5. 只提取實際存在的資訊，不要推測
            6. 注意區分被保險人區塊和要保人區塊（通常有虛線分隔）
            7. 法人之代表人可能與被保險人/要保人相同
            8. 承保內容請務必以陣列方式結構化輸出，每一筆主項（有數字代號）為一筆，若有子項（如每一個人傷害等，前面有大空格），請以 sub_items 陣列放在對應主項下。每一欄若值為空白就留空，若明確寫「無」就顯示「無」。
            9. 產險公司通常位於文件最上方，請仔細辨識公司全名
            10. 對於沒有辨識出來的欄位，請在該欄位填入「無法辨識」並在備註中說明原因
            """
            
            # 建立圖片數據
            image_data = {
                "mime_type": "image/png",
                "data": image_base64
            }
            
            # 呼叫 Gemini API
            response = self.model.generate_content([prompt, image_data])
            
            # 解析回應
            if response.text:
                try:
                    # 清理回應文字，移除可能的 markdown 標記
                    cleaned_text = response.text.strip()
                    if cleaned_text.startswith('```json'):
                        cleaned_text = cleaned_text[7:]  # 移除 ```json
                    if cleaned_text.endswith('```'):
                        cleaned_text = cleaned_text[:-3]  # 移除結尾 ```
                    cleaned_text = cleaned_text.strip()
                    
                    # 嘗試解析 JSON
                    result = json.loads(cleaned_text)
                    
                    # 預期欄位（按照順序）
                    expected_fields = [
                        'insurance_company', 'insured_section', 'insured_person',
                        'legal_representative', 'id_number', 'birth_date', 'gender',
                        'policyholder_section', 'policyholder', 'relationship',
                        'policyholder_legal_representative', 'policyholder_gender',
                        'policyholder_id', 'policyholder_birth_date', 'vehicle_type',
                        'license_number', 'coverage_items', 'total_premium'
                    ]
                    insurance_data = {field: result.get(field, '') for field in expected_fields}
                    
                    logger.info(f"Gemini 提取保單資料成功")
                    return insurance_data
                    
                except json.JSONDecodeError:
                    logger.error("Gemini 回應不是有效的 JSON 格式")
                    print("\n[DEBUG] Gemini 原始回應內容：")
                    print(response.text)
                    return {}
            
            return {}
            
        except Exception as e:
            logger.error(f"Gemini 保單資料提取失敗: {str(e)}")
            return {}
    
    def analyze_image_quality(self, image_path: str) -> Dict:
        """
        分析圖片品質
        
        Args:
            image_path: 圖片路徑
            
        Returns:
            品質分析結果
        """
        try:
            image = cv2.imread(image_path)
            if image is None:
                return {'error': '無法讀取圖片'}
            
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            height, width = gray.shape
            
            # 計算品質指標
            dpi_estimate = min(width, height) / 8.5  # 假設A4紙張
            mean_brightness = np.mean(gray)
            contrast = np.std(gray)
            
            # 評估品質
            quality_score = 0
            suggestions = []
            
            if dpi_estimate >= 300:
                quality_score += 30
            elif dpi_estimate >= 150:
                quality_score += 20
                suggestions.append("建議提高解析度至300 DPI")
            else:
                suggestions.append("解析度過低，建議重新掃描")
            
            if 100 <= mean_brightness <= 200:
                quality_score += 30
            elif mean_brightness < 100:
                suggestions.append("圖片過暗，建議調整亮度")
            else:
                suggestions.append("圖片過亮，建議調整亮度")
            
            if contrast >= 30:
                quality_score += 40
            else:
                suggestions.append("對比度不足，建議增強對比度")
            
            return {
                'dpi_estimate': float(dpi_estimate),
                'brightness': float(mean_brightness),
                'contrast': float(contrast),
                'quality_score': quality_score,
                'suggestions': suggestions,
                'image_size': f"{width} x {height}"
            }
            
        except Exception as e:
            logger.error(f"圖片品質分析失敗: {str(e)}")
            return {'error': str(e)}
    
    def get_extraction_summary(self, data: Dict) -> Dict:
        """
        取得提取摘要
        
        Args:
            data: 提取的資料
            
        Returns:
            提取摘要
        """
        total_fields = len(data)
        filled_fields = sum(1 for v in data.values() if v)
        
        return {
            'total_fields': total_fields,
            'filled_fields': filled_fields,
            'empty_fields': total_fields - filled_fields,
            'extraction_rate': f"{filled_fields / total_fields * 100:.1f}%" if total_fields > 0 else "0%",
            'data': data
        }

def main():
    """測試函數"""
    print("🔍 Gemini OCR 處理器測試")
    print("=" * 50)
    
    # 初始化處理器
    processor = GeminiOCRProcessor()
    
    # 檢查測試圖片
    import os
    test_images = []
    for ext in ['jpg', 'jpeg', 'png', 'bmp', 'tiff']:
        for file in os.listdir('.'):
            if file.lower().endswith(f'.{ext}'):
                test_images.append(file)
    
    if not test_images:
        print("❌ 沒有找到測試圖片")
        return
    
    # 使用第一個圖片進行測試
    image_path = test_images[0]
    print(f"測試圖片: {image_path}")
    
    # 分析圖片品質
    quality_info = processor.analyze_image_quality(image_path)
    print(f"圖片品質: {quality_info}")
    
    # 提取保單資料
    print("\n開始提取保單資料...")
    data = processor.extract_insurance_data_with_gemini(image_path)
    
    # 顯示結果
    print("\n提取結果:")
    for field, value in data.items():
        status = "✅" if value else "❌"
        print(f"{status} {field}: {value or '未找到'}")
    
    # 顯示統計
    stats = processor.get_extraction_summary(data)
    print(f"\n📊 提取統計:")
    print(f"總欄位數: {stats['total_fields']}")
    print(f"成功提取: {stats['filled_fields']}")
    print(f"提取率: {stats['extraction_rate']}")

if __name__ == "__main__":
    main() 