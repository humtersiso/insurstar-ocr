#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
簡化 PDF 測試腳本
測試 ReportLab 基本功能
"""

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
import os

def test_simple_pdf():
    """測試簡單的 PDF 生成"""
    print("🧪 測試簡單 PDF 生成")
    
    # 建立輸出目錄
    os.makedirs('property_reports', exist_ok=True)
    
    # 建立 PDF 文件
    doc = SimpleDocTemplate(
        'property_reports/simple_test.pdf',
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    # 取得樣式
    styles = getSampleStyleSheet()
    
    # 建立內容
    story = []
    
    # 標題
    story.append(Paragraph("測試標題", styles['Heading1']))
    story.append(Spacer(1, 20))
    
    # 正文
    story.append(Paragraph("這是一個測試段落。", styles['Normal']))
    story.append(Spacer(1, 10))
    story.append(Paragraph("如果這個 PDF 能正常顯示，說明 ReportLab 基本功能正常。", styles['Normal']))
    
    # 生成 PDF
    try:
        doc.build(story)
        print("✅ 簡單 PDF 生成成功")
        return True
    except Exception as e:
        print(f"❌ 簡單 PDF 生成失敗: {str(e)}")
        return False

def test_chinese_pdf():
    """測試中文 PDF 生成"""
    print("\n🧪 測試中文 PDF 生成")
    
    # 建立 PDF 文件
    doc = SimpleDocTemplate(
        'property_reports/chinese_test.pdf',
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    # 取得樣式
    styles = getSampleStyleSheet()
    
    # 建立內容
    story = []
    
    # 中文標題
    story.append(Paragraph("創星保險經紀人股份有限公司", styles['Heading1']))
    story.append(Spacer(1, 20))
    
    # 中文正文
    story.append(Paragraph("這是一個中文測試段落。", styles['Normal']))
    story.append(Spacer(1, 10))
    story.append(Paragraph("測試中文字體顯示是否正常。", styles['Normal']))
    
    # 生成 PDF
    try:
        doc.build(story)
        print("✅ 中文 PDF 生成成功")
        return True
    except Exception as e:
        print(f"❌ 中文 PDF 生成失敗: {str(e)}")
        return False

def main():
    """主函數"""
    print("🔍 ReportLab PDF 測試")
    print("=" * 40)
    
    # 測試簡單 PDF
    simple_success = test_simple_pdf()
    
    # 測試中文 PDF
    chinese_success = test_chinese_pdf()
    
    print("\n📊 測試結果")
    print("=" * 40)
    print(f"簡單 PDF: {'✅ 成功' if simple_success else '❌ 失敗'}")
    print(f"中文 PDF: {'✅ 成功' if chinese_success else '❌ 失敗'}")
    
    if simple_success and chinese_success:
        print("\n🎉 所有測試通過！ReportLab 功能正常")
    else:
        print("\n⚠️ 部分測試失敗，需要檢查 ReportLab 設定")

if __name__ == "__main__":
    main() 