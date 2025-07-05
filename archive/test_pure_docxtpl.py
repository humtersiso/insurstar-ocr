from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import json
import os

# 設定路徑
TEMPLATE_PATH = 'assets/templates/財產分析書.docx'
OUTPUT_PATH = 'output_pure_docxtpl.docx'
OCR_JSON = 'ocr_results/gemini_ocr_output_20250704_144524.json'

# 載入 OCR 資料
with open(OCR_JSON, 'r', encoding='utf-8') as f:
    ocr_data = json.load(f)

# 載入模板
tpl = DocxTemplate(TEMPLATE_PATH)

# context: 先用 OCR 原始資料
context = dict(ocr_data)

# 補齊圖片欄位（如果模板有用到）
context['watermark_name_blue'] = InlineImage(tpl, 'assets/watermark/watermark_name_blue.png', width=Mm(30))
context['watermark_company_blue'] = InlineImage(tpl, 'assets/watermark/watermark_company_blue.png', width=Mm(30))

# 你可以再補其他欄位（如 PCN）
context['PCN'] = 'BB2H699299'

# 渲染並儲存
tpl.render(context)
tpl.save(OUTPUT_PATH)
print(f'已產生 {OUTPUT_PATH}') 