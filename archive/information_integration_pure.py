import glob
import os
import json
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

def select_ocr_json_file():
    """讓使用者選擇 ocr_results 資料夾下的 OCR JSON 檔案"""
    json_files = glob.glob('ocr_results/*.json')
    if not json_files:
        print('❌ ocr_results 資料夾下沒有 json 檔案')
        return None
    print('\n請選擇要測試的 OCR JSON 檔案：')
    for idx, f in enumerate(json_files):
        print(f"  [{idx+1}] {os.path.basename(f)}")
    while True:
        try:
            sel = int(input('請輸入檔案編號：'))
            if 1 <= sel <= len(json_files):
                return json_files[sel-1]
        except Exception:
            pass
        print('輸入錯誤，請重新輸入。')

def main():
    TEMPLATE_PATH = 'assets/templates/財產分析書.docx'
    OUTPUT_PATH = 'output_pure_docxtpl_from_integration.docx'

    # 選擇 OCR JSON
    ocr_json_path = select_ocr_json_file()
    if not ocr_json_path:
        print('❌ 未選擇 OCR 結果檔案')
        return
    with open(ocr_json_path, 'r', encoding='utf-8') as f:
        ocr_data = json.load(f)

    tpl = DocxTemplate(TEMPLATE_PATH)
    context = dict(ocr_data)
    # 補齊圖片欄位
    context['watermark_name_blue'] = InlineImage(tpl, 'assets/watermark/watermark_name_blue.png', width=Mm(30))
    context['watermark_company_blue'] = InlineImage(tpl, 'assets/watermark/watermark_company_blue.png', width=Mm(30))
    context['PCN'] = 'BB2H699299'

    tpl.render(context)
    tpl.save(OUTPUT_PATH)
    print(f'✅ 已產生 {OUTPUT_PATH}')

if __name__ == '__main__':
    main() 