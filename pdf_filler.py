import os
import json
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import black, blue
from reportlab.lib.units import inch
import io
from typing import Dict, Optional

class PDFFiller:
    """PDF表單填寫器"""
    
    def __init__(self):
        """初始化PDF填寫器"""
        # 註冊中文字體（如果可用）
        try:
            pdfmetrics.registerFont(TTFont('SimSun', 'simsun.ttc'))
            self.font_name = 'SimSun'
        except:
            self.font_name = 'Helvetica'
    
    def create_insurance_form(self, data: Dict, output_path: str) -> str:
        """
        建立保單表單PDF
        
        Args:
            data: 保單資料字典
            output_path: 輸出檔案路徑
            
        Returns:
            生成的PDF檔案路徑
        """
        # 建立PDF
        c = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        
        # 設定字體
        c.setFont(self.font_name, 16)
        
        # 標題
        c.drawString(width/2 - 100, height - 50, "保單資料表")
        c.line(width/2 - 100, height - 60, width/2 + 100, height - 60)
        
        # 設定欄位字體
        c.setFont(self.font_name, 12)
        
        # 定義欄位位置
        fields = [
            ('保單號碼', data.get('policy_number', ''), 50, height - 100),
            ('被保險人', data.get('insured_name', ''), 50, height - 130),
            ('保險公司', data.get('insurance_company', ''), 50, height - 160),
            ('保單類型', data.get('policy_type', ''), 50, height - 190),
            ('保費金額', data.get('premium_amount', ''), 50, height - 220),
            ('保額', data.get('coverage_amount', ''), 50, height - 250),
            ('生效日期', data.get('start_date', ''), 50, height - 280),
            ('到期日期', data.get('end_date', ''), 50, height - 310),
            ('業務員', data.get('agent_name', ''), 50, height - 340),
            ('聯絡電話', data.get('phone_number', ''), 50, height - 370),
            ('身分證號', data.get('id_number', ''), 50, height - 400),
            ('地址', data.get('address', ''), 50, height - 430)
        ]
        
        # 繪製欄位
        for label, value, x, y in fields:
            # 欄位標籤
            c.setFont(self.font_name, 10)
            c.setFillColor(black)
            c.drawString(x, y, f"{label}:")
            
            # 欄位值
            c.setFont(self.font_name, 12)
            c.setFillColor(blue)
            c.drawString(x + 80, y, str(value) if value else "未填寫")
            
            # 底線
            c.setStrokeColor(black)
            c.line(x + 80, y - 5, x + 300, y - 5)
        
        # 添加備註
        c.setFont(self.font_name, 10)
        c.setFillColor(black)
        c.drawString(50, height - 480, "備註:")
        c.drawString(50, height - 500, "1. 此表單由OCR系統自動生成")
        c.drawString(50, height - 520, "2. 請核對資料正確性")
        c.drawString(50, height - 540, "3. 如有疑問請聯繫相關人員")
        
        # 保存PDF
        c.save()
        return output_path
    
    def fill_existing_pdf(self, template_path: str, data: Dict, output_path: str) -> str:
        """
        填寫現有的PDF表單
        
        Args:
            template_path: 模板PDF路徑
            data: 要填入的資料
            output_path: 輸出檔案路徑
            
        Returns:
            填寫完成的PDF檔案路徑
        """
        try:
            # 讀取模板PDF
            reader = PdfReader(template_path)
            writer = PdfWriter()
            
            # 複製所有頁面
            for page in reader.pages:
                writer.add_page(page)
            
            # 建立覆蓋層
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=A4)
            width, height = A4
            
            # 設定字體
            c.setFont(self.font_name, 12)
            c.setFillColor(blue)
            
            # 定義欄位對應（需要根據實際PDF模板調整座標）
            field_mappings = {
                'policy_number': (150, height - 100),
                'insured_name': (150, height - 130),
                'insurance_company': (150, height - 160),
                'premium_amount': (150, height - 190),
                'coverage_amount': (150, height - 220),
                'start_date': (150, height - 250),
                'end_date': (150, height - 280)
            }
            
            # 填入資料
            for field, (x, y) in field_mappings.items():
                value = data.get(field, '')
                if value:
                    c.drawString(x, y, str(value))
            
            c.save()
            
            # 將覆蓋層合併到原PDF
            packet.seek(0)
            overlay = PdfReader(packet)
            overlay_page = overlay.pages[0]
            
            # 合併到第一頁
            writer.pages[0].merge_page(overlay_page)
            
            # 保存結果
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
            return output_path
            
        except Exception as e:
            print(f"PDF填寫錯誤: {str(e)}")
            # 如果填寫失敗，回退到建立新表單
            return self.create_insurance_form(data, output_path)
    
    def create_summary_report(self, extraction_data: Dict, validation_result: Dict, output_path: str) -> str:
        """
        建立處理摘要報告
        
        Args:
            extraction_data: 提取的資料
            validation_result: 驗證結果
            output_path: 輸出檔案路徑
            
        Returns:
            報告PDF檔案路徑
        """
        c = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        
        # 標題
        c.setFont(self.font_name, 18)
        c.drawString(width/2 - 100, height - 50, "OCR處理摘要報告")
        c.line(width/2 - 100, height - 60, width/2 + 100, height - 60)
        
        c.setFont(self.font_name, 12)
        
        # 處理統計
        y_position = height - 100
        c.drawString(50, y_position, f"總欄位數: {len(extraction_data)}")
        y_position -= 25
        
        filled_count = sum(1 for v in extraction_data.values() if v)
        c.drawString(50, y_position, f"成功提取: {filled_count}")
        y_position -= 25
        
        extraction_rate = filled_count / len(extraction_data) * 100 if extraction_data else 0
        c.drawString(50, y_position, f"提取率: {extraction_rate:.1f}%")
        y_position -= 40
        
        # 驗證結果
        c.drawString(50, y_position, "驗證結果:")
        y_position -= 25
        
        if validation_result.get('is_valid'):
            c.setFillColor(blue)
            c.drawString(50, y_position, "✓ 資料驗證通過")
        else:
            c.setFillColor(black)
            c.drawString(50, y_position, "✗ 資料驗證失敗")
            y_position -= 25
            c.drawString(50, y_position, f"缺失欄位: {', '.join(validation_result.get('missing_fields', []))}")
        
        y_position -= 40
        
        # 提取的資料
        c.setFillColor(black)
        c.drawString(50, y_position, "提取的資料:")
        y_position -= 25
        
        for field, value in extraction_data.items():
            if value:
                c.drawString(50, y_position, f"{field}: {value}")
                y_position -= 20
        
        # 建議
        if validation_result.get('suggestions'):
            y_position -= 20
            c.drawString(50, y_position, "建議:")
            y_position -= 25
            for suggestion in validation_result['suggestions']:
                c.drawString(50, y_position, f"• {suggestion}")
                y_position -= 20
        
        c.save()
        return output_path
    
    def merge_pdfs(self, pdf_files: list, output_path: str) -> str:
        """
        合併多個PDF檔案
        
        Args:
            pdf_files: PDF檔案路徑列表
            output_path: 輸出檔案路徑
            
        Returns:
            合併後的PDF檔案路徑
        """
        writer = PdfWriter()
        
        for pdf_file in pdf_files:
            if os.path.exists(pdf_file):
                reader = PdfReader(pdf_file)
                for page in reader.pages:
                    writer.add_page(page)
        
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return output_path 