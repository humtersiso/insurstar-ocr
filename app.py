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

app = Flask(__name__)
CORS(app)

# 設定
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'pdf', 'tiff', 'bmp'}

# 建立必要的資料夾
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 建立暫存圖片資料夾
TEMP_PREVIEW_FOLDER = os.path.join(os.path.dirname(__file__), 'temp_previews')
os.makedirs(TEMP_PREVIEW_FOLDER, exist_ok=True)

# 啟動時自動清理 temp_previews 資料夾下所有檔案
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
image_processor = ImageProcessing(output_dir=TEMP_PREVIEW_FOLDER)
word_filler = WordFiller()  # Word 填寫系統

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
        
        # 處理 PDF 轉圖片
        images = image_processor.pdf_to_images(upload_path, pages=[0])
        if images:
            image_path = images[0]
        else:
            cleanup_temp_files([upload_path])
            return jsonify({'error': 'PDF 轉換失敗'}), 500
        
        # 保存原始圖片到暫存資料夾（用於預覽）
        base_name = os.path.splitext(os.path.basename(image_path))[0]
        original_preview_path = os.path.join(TEMP_PREVIEW_FOLDER, f"{base_name}_original.png")
        shutil.copy2(image_path, original_preview_path)
        original_preview_url = f"/temp_previews/{os.path.basename(original_preview_path)}"
        
        # 疊加輔助線到圖片上（存到暫存資料夾，用於OCR）
        overlay_image_path = os.path.join(TEMP_PREVIEW_FOLDER, f"{base_name}_with_overlay.png")
        image_with_overlay = image_processor.overlay_table_lines(image_path, output_path=overlay_image_path)
        overlay_preview_url = f"/temp_previews/{os.path.basename(image_with_overlay)}"
        
        # 分析圖片品質
        quality_info = ocr_processor.analyze_image_quality(image_with_overlay)
        print(f"圖片品質分析: {quality_info}")
        
        # 使用 Gemini 提取保單資料
        raw_data = ocr_processor.extract_insurance_data_with_gemini(image_with_overlay)
        print('=== raw_data ===')
        print(json.dumps(raw_data, ensure_ascii=False, indent=2))
        print('================')
        
        # 資料處理和驗證
        processed_data = data_processor.process_insurance_data(raw_data)
        print('=== processed_data ===')
        print(json.dumps(processed_data, ensure_ascii=False, indent=2))
        print('======================')
        print('=== insurance_details (processed) ===')
        print(json.dumps(processed_data.get('insurance_details', ''), ensure_ascii=False, indent=2))
        print('====================================')
        validation_result = data_processor.validate_processed_data(processed_data)
        
        # 準備回應資料
        response_data = {
            'file_id': file_id,
            'original_filename': filename,
            'extracted_data': processed_data,
            'validation_result': validation_result,
            'data_summary': data_processor.get_data_summary(processed_data),
            'processing_time': datetime.now().isoformat(),
            'original_preview_image_url': original_preview_url,
            'overlay_preview_image_url': overlay_preview_url
        }
        
        # 清理暫存檔案（但不清理 overlay 圖片，讓前端可以預覽）
        temp_files_to_clean = [upload_path]
        if image_path != upload_path:
            temp_files_to_clean.append(image_path)
        # 暫時不清理 overlay 圖片，讓前端可以預覽
        # if image_with_overlay != image_path:
        #     temp_files_to_clean.append(image_with_overlay)
        
        cleanup_temp_files(temp_files_to_clean)
        
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
        validation_result = data_processor.validate_processed_data(processed_data)
        
        return jsonify({
            'success': True,
            'extracted_data': processed_data,
            'validation_result': validation_result
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
                'data_summary': result['data_summary'],
                'validation_result': result['validation_result']
            })
        else:
            return jsonify({'error': result['error']}), 500
        
    except Exception as e:
        return jsonify({'error': f'Word 生成失敗: {str(e)}'}), 500

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

@app.route('/temp_previews/<path:filename>')
def temp_previews(filename):
    return send_from_directory(TEMP_PREVIEW_FOLDER, filename)

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