#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gemini OCR 處理器 - 測試版本
支援自定義 prompt 進行測試
"""

import os
import json
import base64
import logging
from datetime import datetime
from typing import Dict, Optional
import google.generativeai as genai

# 設定日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class GeminiOCRProcessorTest:
    """Gemini OCR 處理器 - 測試版本"""
    
    def __init__(self, api_key: str = "AIzaSyBRSCd0Tke2rOobZOv-HsgDq81CQKER1ks"):
        """
        初始化 Gemini OCR 處理器
        
        Args:
            api_key: Gemini API 金鑰
        """
        try:
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel('gemini-2.0-flash-exp')
            logger.info("Gemini OCR 處理器初始化成功")
        except Exception as e:
            logger.error(f"Gemini OCR 處理器初始化失敗: {str(e)}")
            self.model = None
    
    def image_to_base64(self, image_path: str) -> str:
        """
        將圖片轉換為 base64 編碼
        
        Args:
            image_path: 圖片檔案路徑
            
        Returns:
            base64 編碼的字串
        """
        try:
            with open(image_path, 'rb') as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
            return encoded_string
        except Exception as e:
            logger.error(f"圖片轉 base64 失敗: {str(e)}")
            return ""
    
    def extract_insurance_data_with_gemini(self, image_path: str, file_id: str = None, custom_prompt: str = None) -> Dict:
        """
        使用 Gemini 提取保單相關資料 - 支援自定義 prompt
        
        Args:
            image_path: 圖片路徑
            file_id: 唯一檔案識別碼（用於存 OCR output）
            custom_prompt: 自定義的 prompt
            
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
            
            # 使用自定義 prompt 或預設 prompt
            if custom_prompt:
                prompt = custom_prompt
            else:
                # 預設的強化版 prompt
                prompt = """
                【角色定位】
                你是專業的保險人員，具有豐富的保單文件分析經驗。請以專業的角度仔細分析這張保單圖片，確保每個欄位都能準確辨識。

                【分析框架】
                請按照以下順序進行分析：
                1. 被保險人區塊分析
                2. 要保人區塊分析  
                3. 承保內容分析
                4. 總保險費分析

                【必要欄位清單】
                請提取以下 20 個關鍵欄位：

                === 被保險人區塊 ===
                1. 產險公司 (insurance_company) - 位於文件最上方
                2. 被保險人區塊 (insured_section) - 完整區塊內容
                3. 被保險人 (insured_person) - 被保險人姓名或公司名稱
                4. 法人之代表人 (legal_representative) - 法人代表姓名
                5. 身分證字號/統一編號 (id_number) - 身分證或統編
                6. 出生日期 (birth_date) - 在與被保險人關係的上方
                7. 性別 (gender) - 男/女

                === 要保人區塊 ===
                8. 要保人區塊 (policyholder_section) - 注意虛線分隔
                9. 要保人 (policyholder) - 要保人姓名或公司名稱
                10. 與被保險人關係 (relationship) - 本人/配偶/父母等
                11. 法人之代表人 (policyholder_legal_representative) - 在關係右側，性別下方
                12. 性別 (policyholder_gender) - 男/女
                13. 要保人身份證字號/統編 (policyholder_id) - 身分證或統編
                14. 出生日期 (policyholder_birth_date) - 出生日期

                === 車輛資訊 ===
                15. 車輛種類 (vehicle_type) - 車種描述
                16. 牌照號碼 (license_number) - 車牌號碼

                === 承保內容 ===
                17. 承保內容 (coverage_items) - 詳細分析如下

                === 保費與保期 ===
                18. 總保險費 (total_premium) - 總保費金額
                19. 強制險保期 (compulsory_insurance_period) - 格式：「自民國XXX年X月X日中午12時起至民國XXX年X月X日中午12時止」
                20. 任意險保期 (optional_insurance_period) - 格式同上

                【承保內容詳細分析】
                請將保單「承保內容」區塊每一條分別提取為五個欄位：
                - 保險代號：數字代號
                - 保險種類：保險項目名稱
                - 保險金額：實際保險金額
                - 自負額：自負額金額或比例
                - 簽單保費：保費金額

                【表格分析原則】
                在判斷承保內容每一個數值應該屬於哪個欄位時，請參考以下做法：
                - 如果圖片中有表格線（橫線、直線），請以這些表格線將表格分割成多個儲存格
                - 根據每個儲存格在表格中的位置，對應到正確的欄位
                - 若表格線不明顯，請根據每一行的欄位標題來判斷
                - 若同一行同時有多個數值，請依照它們在表格中的相對位置分別對應
                - 請避免將「自負額」的數字填到「保險金額」欄位，或將「保險金額」的數字填到「自負額」欄位

                【子項處理規則】
                - 若主項有子項（如每一個人傷害、每一意外事故之傷害等），主項的「保險金額」請留空
                - 金額請分別填在子項的「保險金額」欄位
                - 只有當主項沒有子項時，主項的「保險金額」才填寫實際金額
                - 每一筆主項（有數字代號）為一筆，子項請以 sub_items 陣列放在對應主項下

                【注意事項】
                1. 如果某個欄位找不到、空白或無法辨識，請填入「無填寫」
                2. 所有欄位如果偵測到空白、無資料或無法辨識，請回傳「無填寫」
                3. 日期格式請保持原始格式
                4. 金額請包含單位（如：元、萬元等）
                5. 只提取實際存在的資訊，不要推測
                6. 注意區分被保險人區塊和要保人區塊（通常有虛線分隔）
                7. 法人之代表人可能與被保險人/要保人相同
                8. 承保內容請務必以陣列方式結構化輸出
                9. 每一欄若值為空白就留空，若明確寫「無」請填空字串""，不要填入「無」字
                10. 產險公司通常位於文件最上方，請仔細辨識公司全名
                11. 對於沒有辨識出來的欄位，請在該欄位填入「無填寫」
                12. 承保內容任何欄位辨識到「無」時，請填空字串""，不要填入「無」字

                【特殊欄位位置提示】
                - 要保人法人之代表人 (policyholder_legal_representative)：
                  * 通常會出現在與被保險人關係的右側
                  * 在被保險人性別的下面
                  * 可能標示為「法人之代表人」、「代表人」或類似文字
                  * 請仔細尋找這個欄位，不要遺漏
                - 被保人出生日期 (birth_date)：
                  * 在與被保險人關係的上方
                  * 通常標示為「出生日期」或類似文字
                  * 請仔細尋找這個欄位，不要遺漏

                【重要觀察：被保險人與要保人關係】
                - 被保險人可能是公司名稱，但填寫人可能是公司老闆
                - 當 relationship（與被保險人關係）為「本人」時：
                  * 被保險人區塊的個人資料（出生日期、性別、身分證字號）通常與要保人區塊一致
                  * 被保險人可能是公司名稱，但個人資料欄位會填入老闆的個人資訊
                  * 請特別注意這種情況，不要因為被保險人是公司就忽略個人資料欄位
                - 當 relationship 為「本人」時，請交叉比對兩個區塊的個人資料是否一致
                - 如果發現不一致，請以要保人區塊的資料為準，因為要保人通常是實際的個人

                【範例輸出格式】
                {
                    "insurance_company": "新安東京海上產物保險股份有限公司",
                    "insured_section": "被保險人完整區塊內容",
                    "insured_person": "大欣租賃有限公司",
                    "legal_representative": "干慧俊",
                    "id_number": "96847309",
                    "birth_date": "無填寫",
                    "gender": "無填寫",
                    "policyholder_section": "要保人完整區塊內容",
                    "policyholder": "大欣租賃有限公司",
                    "relationship": "本人",
                    "policyholder_legal_representative": "干慧俊",
                    "policyholder_gender": "無填寫",
                    "policyholder_id": "96847309",
                    "policyholder_birth_date": "無填寫",
                    "vehicle_type": "14/租賃小客車",
                    "license_number": "RFR-0276",
                    "coverage_items": [
                        {
                            "保險代號": "31",
                            "保險種類": "第三人傷害責任險",
                            "保險金額": "",
                            "自負額": "無",
                            "簽單保費": "10,539",
                            "sub_items": [
                                {
                                    "保險種類": "每一個人傷害",
                                    "保險金額": "300萬",
                                    "自負額": "",
                                    "簽單保費": ""
                                },
                                {
                                    "保險種類": "每一意外事故之傷害",
                                    "保險金額": "600萬",
                                    "自負額": "",
                                    "簽單保費": ""
                                }
                            ]
                        }
                    ],
                    "total_premium": "NT$31, 373",
                    "compulsory_insurance_period": "無填寫",
                    "optional_insurance_period": "自民國114年5月19日中午12時起至民國115年5月19日中午12時止"
                }
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
                        'license_number', 'coverage_items', 'total_premium',
                        'compulsory_insurance_period', 'optional_insurance_period'
                    ]
                    
                    # 處理所有欄位，將空白、空字串、無法辨識等統一為「無填寫」
                    insurance_data = {}
                    for field in expected_fields:
                        value = result.get(field, '')
                        # 檢查是否為空白、空字串、無法辨識等
                        if not value or str(value).strip() == '' or str(value).strip() == '無法辨識':
                            insurance_data[field] = '無填寫'
                        else:
                            insurance_data[field] = value
                    
                    # 保期欄位已經在統一處理中處理過了，不需要重複處理
                    
                    logger.info(f"Gemini 提取保單資料成功")
                    # 新增：將 output 存成 json
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    output_path = os.path.join('ocr_results', f'gemini_ocr_output_{timestamp}.json')
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(result, f, ensure_ascii=False, indent=2)
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