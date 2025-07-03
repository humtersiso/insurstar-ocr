#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 模板處理系統
讀取創星保經財產分析書模板，填入OCR辨識結果，處理勾選標記
"""

import os
import re
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docxtpl import DocxTemplate
import json

class WordTemplateProcessor:
    """Word 模板處理器"""
    
    def __init__(self, template_path: Optional[str] = None):
        """
        初始化處理器
        
        Args:
            template_path: Word 模板檔案路徑，預設為創星保經財產分析書
        """
        if template_path is None:
            template_path = "assets/templates/財產分析書_fixed.docx"
        
        self.template_path = template_path
        self.document = None
        
        # 檢查模板檔案是否存在
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Word 模板檔案不存在: {template_path}")
        
        # OCR欄位對照表
        self.field_mapping = {
            "insurance_company": "{{insurance_company}}",
            "insured_section": "{{insured_section}}",
            "insured_person": "{{insured_person}}",
            "legal_representative": "{{legal_representative}}",
            "id_number": "{{id_number}}",
            "birth_date": "{{birth_date}}",
            "gender": "{{gender}}",
            "policyholder_section": "{{policyholder_section}}",
            "policyholder": "{{policyholder}}",
            "relationship": "{{relationship}}",
            "policyholder_legal_representative": "{{policyholder_legal_representative}}",
            "policyholder_gender": "{{policyholder_gender}}",
            "policyholder_id": "{{policyholder_id}}",
            "policyholder_birth_date": "{{policyholder_birth_date}}",
            "vehicle_type": "{{vehicle_type}}",
            "license_number": "{{license_number}}",
            "coverage_items": "{{coverage_items}}",
            "total_premium": "{{total_premium}}",
            "compulsory_insurance_period": "{{compulsory_insurance_period}}",
            "optional_insurance_period": "{{optional_insurance_period}}",
            "optional_insurance_amount": "{{optional_insurance_amount}}"
        }
        
        # 勾選標記對照
        self.checkbox_mapping = {
            "gender_male": "{{gender_male}}",
            "gender_female": "{{gender_female}}",
            "policyholder_gender_male": "{{policyholder_gender_male}}",
            "policyholder_gender_female": "{{policyholder_gender_female}}",
            "relationship_self": "{{CHECK_RELATIONSHIP_本人}}",
            "relationship_spouse": "{{CHECK_RELATIONSHIP_配偶}}",
            "relationship_parent": "{{CHECK_RELATIONSHIP_父母}}",
            "relationship_child": "{{CHECK_RELATIONSHIP_子女}}",
            "relationship_employee": "{{CHECK_RELATIONSHIP_雇傭}}",
            "relationship_grandparent": "{{CHECK_RELATIONSHIP_祖孫}}",
            "relationship_creditor": "{{CHECK_RELATIONSHIP_債權債務}}",
            "relationship_object": "{{CHECK_RELATIONSHIP_標的物}}",
            "check_1": "{{CHECK_1}}",
            "check_2": "{{CHECK_2}}"
        }
    
    def fix_template_issues(self, template_path: str) -> str:
        """
        修復模板檔案中的問題標記
        
        Args:
            template_path: 原始模板路徑
            
        Returns:
            修復後的模板路徑
        """
        try:
            # 載入原始模板
            doc = Document(template_path)
            
            # 修復空的標記
            fixed_count = 0
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if "{{}}" in para.text:
                                para.text = para.text.replace("{{}}", "{{gender_male}}")
                                fixed_count += 1
            
            # 儲存修復後的模板
            fixed_template_path = template_path.replace('.docx', '_fixed.docx')
            doc.save(fixed_template_path)
            
            if fixed_count > 0:
                print(f"✅ 修復了 {fixed_count} 個問題標記")
                return fixed_template_path
            else:
                print("✅ 模板檔案無需修復")
                return template_path
                
        except Exception as e:
            print(f"❌ 修復模板失敗: {str(e)}")
            return template_path
    
    def load_template(self):
        """載入 Word 模板"""
        try:
            # 先修復模板問題
            fixed_template_path = self.fix_template_issues(self.template_path)
            self.template_path = fixed_template_path
            
            self.document = Document(self.template_path)
            print(f"✅ Word 模板載入成功: {self.template_path}")
            return True
        except Exception as e:
            print(f"❌ Word 模板載入失敗: {str(e)}")
            return False
    
    def process_ocr_data(self, ocr_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        處理OCR辨識資料，轉換為Word模板所需的格式
        
        Args:
            ocr_data: OCR辨識結果字典
            
        Returns:
            處理後的資料字典
        """
        # 將所有值為 '無填寫' 的欄位轉為空字串
        ocr_data = {k: ("" if isinstance(v, str) and v.strip() == "無填寫" else v) for k, v in ocr_data.items()}

        processed_data = {}
        
        # 處理基本欄位
        for field, template_tag in self.field_mapping.items():
            if field in ["compulsory_insurance_period", "optional_insurance_period"]:
                # 若為空值則保留原本內容
                value = ocr_data.get(field)
                if value is None:
                    processed_data[field] = None
                else:
                    processed_data[field] = str(value)
            else:
                value = ocr_data.get(field, "")
                processed_data[field] = str(value) if value else ""
        
        # 性別勾選（被保險人）
        gender = ocr_data.get("gender", "")
        if gender == "男":
            processed_data["gender_male"] = "☑ "
            processed_data["gender_female"] = "□"
        elif gender == "女":
            processed_data["gender_male"] = "□"
            processed_data["gender_female"] = "☑ "
        else:
            processed_data["gender_male"] = "□"
            processed_data["gender_female"] = "□"
        processed_data["gender"] = gender
        
        # 性別勾選（要保人）
        policyholder_gender = ocr_data.get("policyholder_gender", "")
        if policyholder_gender == "男":
            processed_data["policyholder_gender_male"] = "☑ "
            processed_data["policyholder_gender_female"] = "□"
        elif policyholder_gender == "女":
            processed_data["policyholder_gender_male"] = "□"
            processed_data["policyholder_gender_female"] = "☑ "
        else:
            processed_data["policyholder_gender_male"] = "□"
            processed_data["policyholder_gender_female"] = "□"
        processed_data["policyholder_gender"] = policyholder_gender
        
        # 車種勾選
        vehicle_type = ocr_data.get("vehicle_type", "")
        if "機車" in vehicle_type:
            processed_data["vehicle_type_moto"] = "☑ "
            processed_data["vehicle_type_car"] = "□"
        else:
            processed_data["vehicle_type_moto"] = "□"
            processed_data["vehicle_type_car"] = "☑ "
        processed_data["vehicle_type"] = vehicle_type
        
        # 新版關係勾選（relationship_1 ~ relationship_8）
        relationship = ocr_data.get("relationship", "")
        relationship_map = {
            "本人": "relationship_1",
            "配偶": "relationship_2",
            "父母": "relationship_3",
            "子女": "relationship_4",
            "雇傭": "relationship_5",
            "祖孫": "relationship_6",
            "債權債務": "relationship_7",
            "標的物": "relationship_8"
        }
        for rel, tag in relationship_map.items():
            processed_data[tag] = "☑ " if relationship == rel else "□"
        
        # 舊版關係勾選（保留原有 CHECK_RELATIONSHIP_XXX 標記）
        relationship_options = [
            "本人", "配偶", "父母", "子女", "雇傭", "祖孫", "債權債務", "標的物"
        ]
        for option in relationship_options:
            key = f"CHECK_RELATIONSHIP_{option}"
            processed_data[key] = "☑ " if relationship == option else "□"
        
        # 固定勾選項目
        processed_data["CHECK_1"] = "☑ "
        processed_data["CHECK_2"] = "☑ "
        
        # 保期勾選
        compulsory_period = ocr_data.get("compulsory_insurance_period", None)
        optional_period = ocr_data.get("optional_insurance_period", None)
        processed_data["check_compulsory_insurance_period"] = "☑ " if compulsory_period not in [None, ""] else "□"
        processed_data["check_optional_insurance_period"] = "☑ " if optional_period not in [None, ""] else "□"
        
        # 若 policyholder_gender、relationship、gender 為空，補「□」
        for field in ["policyholder_gender", "relationship", "gender"]:
            if not processed_data.get(field):
                processed_data[field] = "□"
        
        # 判斷強制/任意險有無
        compulsory = ocr_data.get("compulsory_insurance_period")
        optional = ocr_data.get("optional_insurance_period")
        coverage_items = ocr_data.get("coverage_items", [])
        def find_car_damage_amount():
            for item in coverage_items:
                if "車體損失保險" in item.get("保險種類", ""):
                    return item.get("保險金額", "")
            return ""
        def find_third_party_personal_amount():
            for item in coverage_items:
                if "第三人傷害責任險" in item.get("保險種類", ""):
                    for sub in item.get("sub_items", []):
                        if "每一個人傷害" in sub.get("保險種類", ""):
                            return sub.get("保險金額", "")
            return ""
        optional_insurance_amount = ""
        if compulsory and not optional:
            # 只有強制險，金額不填
            pass
        elif optional:
            # 有任意險（不論強制險有無）
            car_damage = find_car_damage_amount()
            if car_damage:
                optional_insurance_amount = car_damage
            else:
                optional_insurance_amount = find_third_party_personal_amount()
        processed_data["optional_insurance_amount"] = optional_insurance_amount
        
        # 若有需要填入強制險金額欄位，這裡不填值（保留空白）
        processed_data["compulsory_insurance_amount"] = ""
        
        return processed_data
    
    def set_checkbox_font(self, docx_path: str, extra_fields: dict = {}):
        """將所有 '☑' 設為新細明體8pt粗體，'□' 設為新細明體8pt非粗體，指定欄位內容也設新細明體8pt非粗體"""
        doc = Document(docx_path)
        # 處理所有段落
        for para in doc.paragraphs:
            for run in para.runs:
                if "☑" in run.text:
                    run.font.name = "新細明體"
                    run.font.size = Pt(8)
                    run.bold = True
                elif "□" in run.text:
                    run.font.name = "新細明體"
                    run.font.size = Pt(8)
                    run.bold = False
                # 額外處理純文字欄位
                for value in extra_fields.values():
                    if value and value in run.text:
                        run.font.name = "新細明體"
                        run.font.size = Pt(8)
                        run.bold = False
        # 處理表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if "☑" in run.text:
                                run.font.name = "新細明體"
                                run.font.size = Pt(8)
                                run.bold = True
                            elif "□" in run.text:
                                run.font.name = "新細明體"
                                run.font.size = Pt(8)
                                run.bold = False
                            for value in extra_fields.values():
                                if value and value in run.text:
                                    run.font.name = "新細明體"
                                    run.font.size = Pt(8)
                                    run.bold = False
        doc.save(docx_path)
    
    def fill_template(self, ocr_data: Dict[str, Any], output_path: Optional[str] = None) -> Optional[Tuple[str, Optional[str]]]:
        """
        填入OCR資料到Word模板
        
        Args:
            ocr_data: OCR辨識結果
            output_path: 輸出檔案路徑，如果為None則自動生成
            
        Returns:
            生成的檔案路徑，失敗時返回None
        """
        try:
            print("🔄 開始處理Word模板...")
            
            # 處理OCR資料
            processed_data = self.process_ocr_data(ocr_data)
            
            # 生成輸出檔案路徑
            if output_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"property_reports/財產分析書_{timestamp}.docx"
            
            # 確保輸出目錄存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # 使用docxtpl填入資料
            template = DocxTemplate(self.template_path)
            template.render(processed_data)
            template.save(output_path)
            print(f"✅ Word 檔案生成成功: {output_path}")
            
            # 設定所有 '☑' 及 '□' 為新細明體8pt
            self.set_checkbox_font(
                output_path,
                extra_fields={
                    "policyholder_gender": processed_data.get("policyholder_gender", ""),
                    "relationship": processed_data.get("relationship", ""),
                    "gender": processed_data.get("gender", "")
                }
            )
            
            # 產生PDF
            pdf_path = output_path.replace('.docx', '.pdf')
            try:
                import docx2pdf
                docx2pdf.convert(output_path, pdf_path)
                print(f"✅ PDF 檔案生成成功: {pdf_path}")
            except Exception as e:
                print(f"⚠️ PDF 轉換失敗: {e}")
                pdf_path = None
            
            return (output_path, pdf_path)
        except Exception as e:
            print(f"❌ 填入資料失敗: {str(e)}")
            return None
    
    def get_template_info(self) -> Dict[str, Any]:
        """獲取模板資訊"""
        if not self.document:
            self.load_template()
        
        if self.document is None:
            return {
                "template_path": self.template_path,
                "paragraphs_count": 0,
                "tables_count": 0,
                "sections_count": 0,
                "field_mapping": self.field_mapping,
                "checkbox_mapping": self.checkbox_mapping
            }
        
        return {
            "template_path": self.template_path,
            "paragraphs_count": len(self.document.paragraphs),
            "tables_count": len(self.document.tables),
            "sections_count": len(self.document.sections),
            "field_mapping": self.field_mapping,
            "checkbox_mapping": self.checkbox_mapping
        }
    
    def validate_ocr_data(self, ocr_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        驗證OCR資料的完整性
        
        Args:
            ocr_data: OCR辨識結果
            
        Returns:
            驗證結果，包含錯誤和警告
        """
        errors: List[str] = []
        warnings: List[str] = []
        
        # 檢查必要欄位
        required_fields = ["insured_person", "policyholder", "vehicle_type"]
        for field in required_fields:
            if not ocr_data.get(field):
                errors.append(f"缺少必要欄位: {field}")
        
        # 檢查性別欄位格式
        gender = ocr_data.get("gender", "")
        if gender and gender not in ["男", "女"]:
            warnings.append(f"性別欄位格式異常: {gender}")
        
        policyholder_gender = ocr_data.get("policyholder_gender", "")
        if policyholder_gender and policyholder_gender not in ["男", "女"]:
            warnings.append(f"要保人性別欄位格式異常: {policyholder_gender}")
        
        # 檢查關係欄位
        relationship = ocr_data.get("relationship", "")
        valid_relationships = ["本人", "配偶", "父母", "子女", "雇傭", "祖孫", "債權債務", "標的物"]
        if relationship and relationship not in valid_relationships:
            warnings.append(f"關係欄位格式異常: {relationship}")
        
        return {
            "errors": errors,
            "warnings": warnings,
            "is_valid": len(errors) == 0
        }
    
    def save_processed_data(self, ocr_data: Dict[str, Any], output_path: str):
        """
        儲存處理後的資料為JSON檔案
        
        Args:
            ocr_data: OCR辨識結果
            output_path: 輸出檔案路徑
        """
        try:
            processed_data = self.process_ocr_data(ocr_data)
            validation_result = self.validate_ocr_data(ocr_data)
            
            output_data = {
                "original_ocr_data": ocr_data,
                "processed_data": processed_data,
                "validation_result": validation_result,
                "processed_at": datetime.now().isoformat(),
                "template_path": self.template_path
            }
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(output_data, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 處理資料已儲存: {output_path}")
            
        except Exception as e:
            print(f"❌ 儲存處理資料失敗: {str(e)}")

def main():
    """測試Word模板處理器"""
    print("📄 Word 模板處理器測試")
    print("=" * 50)
    
    try:
        # 初始化處理器
        processor = WordTemplateProcessor()
        
        # 顯示模板資訊
        template_info = processor.get_template_info()
        print(f"📋 模板資訊:")
        print(f"   - 模板路徑: {template_info['template_path']}")
        print(f"   - 段落數量: {template_info['paragraphs_count']}")
        print(f"   - 表格數量: {template_info['tables_count']}")
        print(f"   - 區段數量: {template_info['sections_count']}")
        
        # 測試OCR資料
        test_ocr_data = {
            "insurance_company": "新安東京海上產物保險股份有限公司",
            "insured_section": "被保險人區塊",
            "insured_person": "王小明",
            "legal_representative": "林經理",
            "id_number": "A123456789",
            "birth_date": "1990-01-01",
            "gender": "男",
            "policyholder_section": "要保人區塊",
            "policyholder": "王小明",
            "relationship": "本人",
            "policyholder_legal_representative": "林經理",
            "policyholder_gender": "男",
            "policyholder_id": "A123456789",
            "policyholder_birth_date": "1990-01-01",
            "vehicle_type": "小客車",
            "license_number": "ABC-1234",
            "coverage_items": "車體險、強制險",
            "total_premium": "27,644"
        }
        
        # 驗證資料
        validation_result = processor.validate_ocr_data(test_ocr_data)
        print(f"\n🔍 資料驗證結果:")
        print(f"   - 是否有效: {validation_result['is_valid']}")
        if validation_result['errors']:
            print(f"   - 錯誤: {validation_result['errors']}")
        if validation_result['warnings']:
            print(f"   - 警告: {validation_result['warnings']}")
        
        # 填入模板
        output_path, pdf_path = processor.fill_template(test_ocr_data)
        
        if output_path:
            # 儲存處理資料
            json_output_path = output_path.replace('.docx', '_data.json')
            processor.save_processed_data(test_ocr_data, json_output_path)
            
            print(f"\n✅ 測試完成！")
            print(f"   - Word檔案: {output_path}")
            print(f"   - PDF檔案: {pdf_path}")
            print(f"   - 資料檔案: {json_output_path}")
        
    except Exception as e:
        print(f"❌ 測試失敗: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 