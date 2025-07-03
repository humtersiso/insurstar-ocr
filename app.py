import os
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
import tempfile
import shutil
import json
import time

from gemini_ocr_processor import GeminiOCRProcessor
from data_processor import DataProcessor
from image_processing import ImageProcessing
from word_filler import WordFiller
from word_template_processor import WordTemplateProcessor

app = Flask(__name__)
CORS(app)

# 設定
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'property_reports'
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'pdf', 'tiff', 'bmp'}

# 建立必要的資料夾
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 建立暫存圖片資料夾
TEMP_PREVIEW_FOLDER = os.path.join(os.path.dirname(__file__), 'temp_images')
os.makedirs(TEMP_PREVIEW_FOLDER, exist_ok=True)

# 啟動時自動清理 temp_images 資料夾下所有檔案
def cleanup_temp_previews_on_start(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        if os.path.isfile(file_path):
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"刪除暫存檔案失敗: {file_path} - {e}")

cleanup_temp_previews_on_start(TEMP_PREVIEW_FOLDER)

# 初始化處理器
ocr_processor = GeminiOCRProcessor()  # 使用 Gemini OCR
data_processor = DataProcessor()
image_processor = ImageProcessing
word_filler = WordFiller()  # Word 填寫系統
word_template_processor = WordTemplateProcessor()  # Word 模板處理器

def allowed_file(filename):
    """檢查檔案格式是否允許"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def cleanup_temp_files(file_paths):
    """清理暫存檔案"""
    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"清理檔案錯誤: {str(e)}")

@app.route('/')
def index():
    """首頁"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """處理檔案上傳和OCR辨識"""
    try:
        # 檢查是否有檔案
        if 'file' not in request.files:
            return jsonify({'error': '沒有選擇檔案'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '沒有選擇檔案'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': '不支援的檔案格式'}), 400
        
        # 生成唯一檔案名
        file_id = str(uuid.uuid4())
        filename = secure_filename(file.filename or 'unknown')
        file_ext = filename.rsplit('.', 1)[1].lower()
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{file_id}.{file_ext}")
        
        # 保存上傳的檔案
        file.save(upload_path)

        # 如果是 PDF，先轉成圖片（只取第一頁）
        ext = file_ext.lower()
        if ext == 'pdf':
            image_paths = ImageProcessing.pdf_to_images(upload_path, app.config['UPLOAD_FOLDER'], dpi=300)
            if image_paths:
                image_path_for_ocr = image_paths[0]
            else:
                return jsonify({'error': 'PDF 轉圖片失敗'}), 500
        else:
            image_path_for_ocr = upload_path

        # ====== Gemini OCR 與資料處理 ======
        # 產生加上輔助線的預覽圖（強化去背紅線版本）
        support_line_path = os.path.join('assets', 'watermark', 'table_line_redline_only.png')
        preview_images = ImageProcessing.overlay_support_line_on_pdf(
            image_path_for_ocr, TEMP_PREVIEW_FOLDER, support_line_path, dpi=300, alpha=1.0
        )
        if preview_images:
            overlay_preview_path = preview_images[0]
            overlay_preview_url = f"/temp_images/{os.path.basename(overlay_preview_path)}"
        else:
            overlay_preview_url = None

        if overlay_preview_path:
            # OCR
            raw_data = ocr_processor.extract_insurance_data_with_gemini(image_path_for_ocr, file_id=file_id)
            # 資料處理
            processed_data = data_processor.process_insurance_data(raw_data)

            # 摘要
            data_summary = data_processor.get_data_summary(processed_data)
            
            # 分析保期資訊和保險類型
            insurance_periods_info = {
                'compulsory_insurance_period': {
                    'raw': raw_data.get('compulsory_insurance_period', ''),
                    'processed': processed_data.get('compulsory_insurance_period', ''),
                    'status': '有資料' if processed_data.get('compulsory_insurance_period') and processed_data.get('compulsory_insurance_period') != '無填寫' else '無填寫'
                },
                'optional_insurance_period': {
                    'raw': raw_data.get('optional_insurance_period', ''),
                    'processed': processed_data.get('optional_insurance_period', ''),
                    'status': '有資料' if processed_data.get('optional_insurance_period') and processed_data.get('optional_insurance_period') != '無填寫' else '無填寫'
                }
            }
            
            # 分析保險類型
            coverage_items = processed_data.get('coverage_items', [])
            insurance_types = {
                'has_compulsory': False,
                'has_optional': False,
                'compulsory_items': [],
                'optional_items': []
            }
            
            if isinstance(coverage_items, list):
                for item in coverage_items:
                    if isinstance(item, dict):
                        insurance_type = item.get('保險種類', '')
                        if '強制' in insurance_type:
                            insurance_types['has_compulsory'] = True
                            insurance_types['compulsory_items'].append(insurance_type)
                        elif any(keyword in insurance_type for keyword in ['車體', '竊盜', '第三人', '超額', '駕駛人']):
                            insurance_types['has_optional'] = True
                            insurance_types['optional_items'].append(insurance_type)
            
            # 生成提醒訊息
            reminders = []
            
            # 根據保期判斷是否有強制險和任意險
            has_compulsory_period = insurance_periods_info['compulsory_insurance_period']['status'] == '有資料'
            has_optional_period = insurance_periods_info['optional_insurance_period']['status'] == '有資料'
            
            if not has_compulsory_period:
                reminders.append("沒有辨識到強制險")
            if not has_optional_period:
                reminders.append("沒有辨識到任意險")
            
            # 如果沒有提醒，添加成功訊息
            if not reminders:
                reminders.append("所有保險項目和保期都已成功辨識")
        else:
            processed_data = {}
            data_summary = {}
            insurance_periods_info = {}
            insurance_types = {}
            reminders = ["❌ 無法產生預覽圖，請檢查檔案格式"]

        response_data = {
            'file_id': file_id,
            'original_filename': filename,
            'overlay_preview_image_url': overlay_preview_url,
            'original_preview_image_url': image_path_for_ocr if os.path.exists(image_path_for_ocr) else None,
            'extracted_data': processed_data,

            'data_summary': data_summary,
            'insurance_periods_info': insurance_periods_info,
            'insurance_types': insurance_types,
            'reminders': reminders,
            'total_fields': 20,
            'filled_fields': sum(1 for v in processed_data.values() if v and v != "無填寫") if processed_data else 0,
            'extraction_rate': f"{sum(1 for v in processed_data.values() if v and v != '無填寫') / 20 * 100:.1f}%" if processed_data else "0.0%"
        }

        return jsonify(response_data)
        
    except Exception as e:
        print(f"處理錯誤: {str(e)}")
        return jsonify({'error': f'處理失敗: {str(e)}'}), 500

@app.route('/api/process', methods=['POST'])
def api_process():
    """API端點：處理OCR和PDF生成"""
    try:
        data = request.get_json()
        
        if not data or 'image_path' not in data:
            return jsonify({'error': '缺少必要參數'}), 400
        
        image_path = data['image_path']
        
        if not os.path.exists(image_path):
            return jsonify({'error': '圖片檔案不存在'}), 404
        
        # 執行OCR處理
        raw_data = ocr_processor.extract_insurance_data_with_gemini(image_path)
        processed_data = data_processor.process_insurance_data(raw_data)
        
        return jsonify({
            'success': True,
            'extracted_data': processed_data
        })
        
    except Exception as e:
        return jsonify({'error': f'API處理失敗: {str(e)}'}), 500

@app.route('/api/health')
def health_check():
    """健康檢查端點"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'version': '1.0.0'
    })

@app.route('/api/generate-word', methods=['POST'])
def generate_word():
    """生成 Word 檔案"""
    try:
        data = request.get_json()
        
        if not data or 'image_path' not in data:
            return jsonify({'error': '缺少必要參數'}), 400
        
        image_path = data['image_path']
        
        if not os.path.exists(image_path):
            return jsonify({'error': '圖片檔案不存在'}), 404
        
        # 使用 Word 填寫系統處理
        result = word_filler.process_insurance_document(image_path)
        
        if result['success']:
            return jsonify({
                'success': True,
                'word_filename': result['word_filename'],
                'word_path': result['word_path'],
                'download_url': f"/download/{result['word_filename']}",
                'data_summary': result['data_summary']
            })
        else:
            return jsonify({'error': result['error']}), 500
        
    except Exception as e:
        return jsonify({'error': f'Word 生成失敗: {str(e)}'}), 500

@app.route('/api/generate-word-template', methods=['POST'])
def generate_word_template():
    """使用Word模板生成財產分析書"""
    try:
        data = request.get_json()
        
        if not data or 'ocr_data' not in data:
            return jsonify({'error': '缺少必要參數'}), 400
        
        ocr_data = data['ocr_data']
        
        # 使用Word模板處理器
        output_path = word_template_processor.fill_template(ocr_data)
        
        if output_path:
            filename = os.path.basename(output_path)
            return jsonify({
                'success': True,
                'word_filename': filename,
                'word_path': output_path,
                'download_url': f"/download/{filename}",
                'message': '財產分析書生成成功'
            })
        else:
            return jsonify({'error': 'Word模板處理失敗'}), 500
        
    except Exception as e:
        return jsonify({'error': f'Word模板生成失敗: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下載檔案"""
    try:
        return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': f'下載失敗: {str(e)}'}), 404

@app.route('/api/analyze', methods=['POST'])
def analyze_image():
    """分析圖片品質並提供優化建議"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '沒有選擇檔案'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '沒有選擇檔案'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': '不支援的檔案格式'}), 400
        
        # 生成暫存檔案
        file_id = str(uuid.uuid4())
        filename = secure_filename(file.filename or 'unknown')
        file_ext = filename.rsplit('.', 1)[1].lower()
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{file_id}.{file_ext}")
        
        file.save(temp_path)
        
        try:
            # 分析圖片品質
            quality_info = ocr_processor.analyze_image_quality(temp_path)
            
            # 執行 Gemini OCR 測試
            extracted_data = ocr_processor.extract_insurance_data_with_gemini(temp_path)
            extraction_stats = ocr_processor.get_extraction_summary(extracted_data)
            
            return jsonify({
                'quality_analysis': quality_info,
                'extracted_data': extracted_data,
                'extraction_stats': extraction_stats
            })
            
        finally:
            # 清理暫存檔案
            cleanup_temp_files([temp_path])
        
    except Exception as e:
        return jsonify({'error': f'分析失敗: {str(e)}'}), 500

@app.route('/temp_images/<path:filename>')
def temp_images_preview(filename):
    return send_from_directory(TEMP_PREVIEW_FOLDER, filename)

@app.route('/api/preview', methods=['POST'])
def api_preview():
    """即時產生疊合輔助線的圖片預覽，檔名用 file_id"""
    try:
        data = request.get_json()
        image_path = data.get('image_path')
        file_id = data.get('file_id')
        if not image_path or not os.path.exists(image_path):
            return jsonify({'error': '圖片檔案不存在'}), 404
        if not file_id:
            return jsonify({'error': '缺少 file_id'}), 400
        # 產生唯一預覽圖路徑
        preview_path = os.path.join(TEMP_PREVIEW_FOLDER, f'{file_id}_preview.png')
        # 疊加輔助線（自動縮放座標與大小）
        support_line_path = os.path.join('assets', 'watermark', 'table_line_redline_only.png')
        ImageProcessing.overlay_support_line_on_image(
            image_path, support_line_path, preview_path,
            orig_size=(2481, 3508), crop_box=(81, 1157, 1797, 747), alpha=1.0
        )
        overlay_preview_url = f"/temp_images/{file_id}_preview.png"
        return jsonify({'preview_url': overlay_preview_url})
    except Exception as e:
        return jsonify({'error': f'預覽產生失敗: {str(e)}'}), 500

@app.route('/api/test-insurance-periods', methods=['POST'])
def test_insurance_periods():
    """測試保期欄位處理功能"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '沒有選擇檔案'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '沒有選擇檔案'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': '不支援的檔案格式'}), 400
        
        # 生成唯一檔案名
        file_id = str(uuid.uuid4())
        filename = secure_filename(file.filename or 'unknown')
        file_ext = filename.rsplit('.', 1)[1].lower()
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{file_id}.{file_ext}")
        
        # 保存上傳的檔案
        file.save(upload_path)

        # 如果是 PDF，先轉成圖片（只取第一頁）
        ext = file_ext.lower()
        if ext == 'pdf':
            image_paths = ImageProcessing.pdf_to_images(upload_path, app.config['UPLOAD_FOLDER'], dpi=300)
            if image_paths:
                image_path_for_ocr = image_paths[0]
            else:
                return jsonify({'error': 'PDF 轉圖片失敗'}), 500
        else:
            image_path_for_ocr = upload_path

        # 執行 Gemini OCR
        print("🔄 執行 Gemini OCR...")
        raw_data = ocr_processor.extract_insurance_data_with_gemini(image_path_for_ocr, file_id=file_id)
        
        if not raw_data:
            return jsonify({'error': 'OCR 辨識失敗'}), 500
        
        # 資料處理
        print("🔄 處理資料...")
        processed_data = data_processor.process_insurance_data(raw_data)
        
        # 準備回應資料
        response_data = {
            'file_id': file_id,
            'original_filename': filename,
            'raw_ocr_data': raw_data,
            'processed_data': processed_data,
            'insurance_periods_info': {
                'compulsory_insurance_period': {
                    'raw': raw_data.get('compulsory_insurance_period', ''),
                    'processed': processed_data.get('compulsory_insurance_period', ''),
                    'status': '有資料' if processed_data.get('compulsory_insurance_period') and processed_data.get('compulsory_insurance_period') != '無填寫' else '無填寫'
                },
                'optional_insurance_period': {
                    'raw': raw_data.get('optional_insurance_period', ''),
                    'processed': processed_data.get('optional_insurance_period', ''),
                    'status': '有資料' if processed_data.get('optional_insurance_period') and processed_data.get('optional_insurance_period') != '無填寫' else '無填寫'
                }
            },
            'total_fields': 20,
            'filled_fields': sum(1 for v in processed_data.values() if v and v != "無填寫"),
            'extraction_rate': f"{sum(1 for v in processed_data.values() if v and v != '無填寫') / 20 * 100:.1f}%"
        }

        return jsonify(response_data)
        
    except Exception as e:
        print(f"測試保期處理錯誤: {str(e)}")
        return jsonify({'error': f'測試失敗: {str(e)}'}), 500

@app.errorhandler(413)
def too_large(e):
    """檔案過大錯誤處理"""
    return jsonify({'error': '檔案大小超過限制 (最大16MB)'}), 413

@app.errorhandler(404)
def not_found(e):
    """404錯誤處理"""
    return jsonify({'error': '頁面不存在'}), 404

@app.errorhandler(500)
def internal_error(e):
    """500錯誤處理"""
    return jsonify({'error': '伺服器內部錯誤'}), 500

if __name__ == '__main__':
    print("啟動保單辨識系統...")
    print("請訪問: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000) 