#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
財產分析書生成器
使用 ReportLab 建立財產分析書 PDF
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
import os
from datetime import datetime
from typing import Dict, List, Any, Optional

class PropertyAnalysisGenerator:
    """財產分析書生成器"""
    
    def __init__(self):
        """初始化生成器"""
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        self._setup_fonts()
    
    def _setup_custom_styles(self):
        """設定自定義樣式"""
        # 標題樣式
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            spaceAfter=20,
            alignment=TA_CENTER
        ))
        
        # 小標題樣式
        self.styles.add(ParagraphStyle(
            name='CustomHeading',
            parent=self.styles['Heading2'],
            fontSize=12,
            spaceAfter=10
        ))
        
        # 正文樣式
        self.styles.add(ParagraphStyle(
            name='CustomBody',
            parent=self.styles['Normal'],
            fontSize=10,
            spaceAfter=6
        ))
        
        # 表格標題樣式
        self.styles.add(ParagraphStyle(
            name='TableHeader',
            parent=self.styles['Normal'],
            fontSize=10,
            alignment=TA_CENTER,
            textColor=colors.white
        ))
    
    def _setup_fonts(self):
        """設定字體"""
        try:
            # 使用 ReportLab 內建的中文字體支援
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
            
            # 更新樣式中的字體
            self.styles.add(ParagraphStyle(
                name='CustomTitle',
                parent=self.styles['Heading1'],
                fontSize=16,
                spaceAfter=20,
                alignment=TA_CENTER,
                fontName='STSong-Light'
            ))
            
            self.styles.add(ParagraphStyle(
                name='CustomHeading',
                parent=self.styles['Heading2'],
                fontSize=12,
                spaceAfter=10,
                fontName='STSong-Light'
            ))
            
            self.styles.add(ParagraphStyle(
                name='CustomBody',
                parent=self.styles['Normal'],
                fontSize=10,
                spaceAfter=6,
                fontName='STSong-Light'
            ))
            
            print("✅ 中文字體設定成功")
            
        except Exception as e:
            print(f"⚠️ 中文字體設定失敗: {str(e)}")
            print("使用預設字體")
    
    def generate_property_analysis(self, insurance_data: Dict, output_path: str) -> str:
        """
        生成財產分析書
        
        Args:
            insurance_data: 保險資料字典
            output_path: 輸出檔案路徑
            
        Returns:
            生成的檔案路徑
        """
        # 建立 PDF 文件
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            rightMargin=2*cm,
            leftMargin=2*cm,
            topMargin=2*cm,
            bottomMargin=2*cm
        )
        
        # 建立內容
        story = []
        
        # 1. 標題
        story.append(Paragraph("創星保險經紀人股份有限公司分析報告書", self.styles['CustomTitle']))
        story.append(Spacer(1, 20))
        
        # 2. 公司資訊
        company_info = [
            ["創星保險經紀人(股)公司", "總公司地址：台北市中山區民權東路二段46號3樓之1"]
        ]
        company_table = Table(company_info, colWidths=[doc.width/2]*2)
        company_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(company_table)
        story.append(Spacer(1, 20))
        
        # 3. 保險類型選擇
        insurance_type = self._create_insurance_type_section(doc)
        story.append(insurance_type)
        story.append(Spacer(1, 20))
        
        # 4. 客戶基本資料
        customer_info = self._create_customer_info_section(insurance_data, doc)
        story.append(customer_info)
        story.append(Spacer(1, 20))
        
        # 5. 財產保險契約分析報告書
        analysis_report = self._create_analysis_report_section(insurance_data, doc)
        story.append(analysis_report)
        
        # 生成 PDF
        doc.build(story)
        return output_path
    
    def _create_insurance_type_section(self, doc) -> Table:
        """建立保險類型選擇區塊"""
        data = [
            ["□人身保險", "□財產保險", "□旅行平安保險"]
        ]
        
        table = Table(data, colWidths=[doc.width/3]*3)
        table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        return table
    
    def _create_customer_info_section(self, data: Dict, doc) -> Table:
        """建立客戶基本資料區塊"""
        # 要保人資料
        policyholder_name = data.get('policyholder', '')
        policyholder_rep = data.get('policyholder_legal_representative', '')
        policyholder_gender = data.get('policyholder_gender', '')
        policyholder_id = data.get('policyholder_id', '')
        policyholder_birth = data.get('policyholder_birth_date', '')
        
        # 被保險人資料
        insured_name = data.get('insured_person', '')
        insured_rep = data.get('legal_representative', '')
        insured_gender = data.get('gender', '')
        insured_id = data.get('id_number', '')
        insured_birth = data.get('birth_date', '')
        
        # 車輛資料
        license_number = data.get('license_number', '')
        vehicle_type = data.get('vehicle_type', '')
        relationship = data.get('relationship', '')
        
        # 建立表格資料
        table_data = [
            ["一、客戶基本資料", "", "", "", ""],
            ["", "", "", "", ""],
            ["要保人", "姓名/公司行號", policyholder_name, "法人代表人", policyholder_rep],
            ["", "□男 □女", "✓" if policyholder_gender == "男" else "□" if policyholder_gender == "女" else "□", "", ""],
            ["", "身分證字號/統編", policyholder_id, "出生年月日", policyholder_birth],
            ["", "職業", "", "", ""],
            ["", "", "", "", ""],
            ["被保險人", "姓名/公司行號", insured_name, "法人代表人", insured_rep],
            ["", "□男 □女", "✓" if insured_gender == "男" else "□" if insured_gender == "女" else "□", "", ""],
            ["", "身分證字號/統編", insured_id, "出生年月日", insured_birth],
            ["", "職業", "", "", ""],
            ["", "", "", "", ""],
            ["投保車險必填", "車牌號碼", license_number, "", ""],
            ["", f"□汽車 □機車", "✓" if "汽車" in vehicle_type else "✓" if "機車" in vehicle_type else "□", "", ""],
            ["", "要／被保險人關係", relationship, "", ""],
        ]
        
        # 建立表格
        col_widths = [doc.width*0.15, doc.width*0.25, doc.width*0.25, doc.width*0.15, doc.width*0.2]
        table = Table(table_data, colWidths=col_widths)
        
        # 設定表格樣式
        table.setStyle(TableStyle([
            # 基本樣式
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            
            # 標題行
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            
            # 邊框
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOX', (0, 0), (-1, -1), 2, colors.black),
        ]))
        
        return table
    
    def _create_analysis_report_section(self, data: Dict, doc) -> Table:
        """建立財產保險契約分析報告書區塊"""
        # 取得承保內容
        coverage_items = data.get('coverage_items', [])
        total_premium = data.get('total_premium', '')
        
        # 建立表格資料
        table_data = [
            ["二、財產保險契約分析報告書", "", "", ""],
            ["", "", "", ""],
            ["本次投保之目的", "", "", ""],
            ["", "", "", ""],
            ["欲投保之保險種類", "保險金額", "自負額", "簽單保費"],
        ]
        
        # 加入承保內容
        for item in coverage_items:
            if isinstance(item, dict):
                insurance_type = item.get('保險種類', '')
                insurance_amount = item.get('保險金額', '')
                deductible = item.get('自負額', '')
                premium = item.get('簽單保費', '')
                
                table_data.append([insurance_type, insurance_amount, deductible, premium])
                
                # 加入子項目
                sub_items = item.get('sub_items', [])
                for sub_item in sub_items:
                    if isinstance(sub_item, dict):
                        sub_type = sub_item.get('保險種類', '')
                        sub_amount = sub_item.get('保險金額', '')
                        sub_deductible = sub_item.get('自負額', '')
                        sub_premium = sub_item.get('簽單保費', '')
                        
                        table_data.append([f"  └ {sub_type}", sub_amount, sub_deductible, sub_premium])
        
        # 加入總計
        table_data.append(["", "", "總保險費", total_premium])
        
        # 建立表格
        col_widths = [doc.width*0.4, doc.width*0.2, doc.width*0.2, doc.width*0.2]
        table = Table(table_data, colWidths=col_widths)
        
        # 設定表格樣式
        table.setStyle(TableStyle([
            # 基本樣式
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            
            # 標題行
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            
            # 欄位標題行
            ('BACKGROUND', (0, 4), (-1, 4), colors.lightgrey),
            ('FONTSIZE', (0, 4), (-1, 4), 10),
            ('ALIGN', (0, 4), (-1, 4), 'CENTER'),
            
            # 邊框
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOX', (0, 0), (-1, -1), 2, colors.black),
        ]))
        
        return table

def main():
    """測試函數"""
    print("🏠 財產分析書生成器測試")
    print("=" * 50)
    
    # 初始化生成器
    generator = PropertyAnalysisGenerator()
    
    # 測試資料
    test_data = {
        'policyholder': '張三',
        'policyholder_legal_representative': '',
        'policyholder_gender': '男',
        'policyholder_id': 'A123456789',
        'policyholder_birth_date': '1980/01/01',
        'insured_person': '張三',
        'legal_representative': '',
        'gender': '男',
        'id_number': 'A123456789',
        'birth_date': '1980/01/01',
        'license_number': 'ABC-1234',
        'vehicle_type': '汽車',
        'relationship': '本人',
        'coverage_items': [
            {
                '保險代號': '05',
                '保險種類': '車體損失保險乙式(Q)',
                '保險金額': '',
                '自負額': '20,000',
                '簽單保費': '20,527',
                'sub_items': [
                    {'保險種類': '每一個人傷害', '保險金額': '300萬', '自負額': '', '簽單保費': '9,485'},
                    {'保險種類': '每一意外事故之傷害', '保險金額': '600萬', '自負額': '', '簽單保費': ''}
                ]
            }
        ],
        'total_premium': '29,012'
    }
    
    # 生成財產分析書
    output_path = 'property_reports/test_property_analysis.pdf'
    os.makedirs('property_reports', exist_ok=True)
    
    try:
        result_path = generator.generate_property_analysis(test_data, output_path)
        print(f"✅ 財產分析書生成成功: {result_path}")
    except Exception as e:
        print(f"❌ 生成失敗: {str(e)}")

if __name__ == "__main__":
    main() 