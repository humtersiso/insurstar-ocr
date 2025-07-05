import json
import os
from information_integration import replace_special_tags
from word_template_processor import WordTemplateProcessor

# 直接用原始模板
template_path = 'assets/templates/財產分析書.docx'

# 讀取 OCR 資料
ocr_json = 'ocr_results/gemini_ocr_output_20250704_144524.json'
with open(ocr_json, 'r', encoding='utf-8') as f:
    ocr_data = json.load(f)

# 主動補齊圖片與特殊欄位 - 修正：使用與 WATERMARK_MAP 一致的格式（無空格）
ocr_data['watermark_name_blue'] = '{{watermark_name_blue}}'
ocr_data['watermark_company_blue'] = '{{watermark_company_blue}}'
ocr_data['PCN'] = '{{PCN}}'

# 替換特殊標記
for k in list(ocr_data.keys()):
    if isinstance(ocr_data[k], str):
        ocr_data[k] = replace_special_tags(ocr_data[k])

print('=== 欄位實際值 ===')
for k, v in ocr_data.items():
    print(f'{k}: {v}')

# 產生 Word
processor = WordTemplateProcessor(template_path)
output_path = processor.fill_template(ocr_data)
print(f'產生 Word 檔案：{output_path}') 