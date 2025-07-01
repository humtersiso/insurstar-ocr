import os
import uuid
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
import tempfile
import shutil
import json

from gemini_ocr_processor import GeminiOCRProcessor
from pdf_filler import PDFFiller
from data_processor import DataProcessor

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

# 初始化處理器
ocr_processor = GeminiOCRProcessor()  # 使用 Gemini OCR
pdf_filler = PDFFiller()
data_processor = DataProcessor()

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
        
        # 執行OCR處理
        print(f"開始處理檔案: {upload_path}")
        
        # 處理 PDF 轉圖片
        image_path = upload_path
        if file_ext.lower() == 'pdf':
            print("檢測到 PDF 檔案，轉換為圖片...")
            from pdf_to_images import PDFToImageConverter
            
            try:
                # 初始化轉換器
                converter = PDFToImageConverter()
                
                # 轉換 PDF 第一頁為圖片
                images = converter.convert_pdf_to_images(upload_path, pages=[0])  # 只轉換第一頁
                if images:
                    image_path = images[0]  # 使用第一頁
                    print(f"PDF 轉換成功: {image_path}")
                else:
                    raise Exception("PDF 轉換失敗，沒有生成圖片")
            except Exception as e:
                print(f"PDF 轉換錯誤: {str(e)}")
                cleanup_temp_files([upload_path])
                return jsonify({'error': f'PDF 轉換失敗: {str(e)}'}), 500
        
        # 分析圖片品質
        quality_info = ocr_processor.analyze_image_quality(image_path)
        print(f"圖片品質分析: {quality_info}")
        
        # 使用 Gemini 提取保單資料
        raw_data = ocr_processor.extract_insurance_data_with_gemini(image_path)
        print('=== raw_data ===')
        print(json.dumps(raw_data, ensure_ascii=False, indent=2))
        print('================')
        
        # 如果是轉換的圖片，清理暫存圖片
        if image_path != upload_path:
            cleanup_temp_files([image_path])
        
        # 檢查是否成功提取資料
        if not raw_data:
            print("❌ 未能提取到任何資料")
            cleanup_temp_files([upload_path])
            return jsonify({
                'error': '無法從檔案中提取保單資料，請確認檔案品質或格式',
                'suggestions': [
                    '請確認檔案是清晰的保單圖片或 PDF',
                    '檔案解析度建議至少 300 DPI',
                    '文字應清晰可讀，避免模糊或傾斜',
                    '支援格式：PNG, JPG, PDF, TIFF, BMP'
                ]
            }), 400
        
        # 資料處理和驗證
        processed_data = data_processor.process_insurance_data(raw_data)
        print('=== processed_data ===')
        print(json.dumps(processed_data, ensure_ascii=False, indent=2))
        print('======================')
        print('=== insurance_details (processed) ===')
        print(json.dumps(processed_data.get('insurance_details', ''), ensure_ascii=False, indent=2))
        print('====================================')
        validation_result = data_processor.validate_processed_data(processed_data)
        
        # 生成PDF檔案
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_filename = f"insurance_form_{file_id}_{timestamp}.pdf"
        pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
        
        # 建立PDF表單
        pdf_filler.create_insurance_form(processed_data, pdf_path)
        
        # 生成摘要報告
        report_filename = f"summary_report_{file_id}_{timestamp}.pdf"
        report_path = os.path.join(app.config['OUTPUT_FOLDER'], report_filename)
        pdf_filler.create_summary_report(processed_data, validation_result, report_path)
        
        # 準備回應資料
        response_data = {
            'file_id': file_id,
            'original_filename': filename,
            'extracted_data': processed_data,
            'validation_result': validation_result,
            'data_summary': data_processor.get_data_summary(processed_data),
            'pdf_filename': pdf_filename,
            'report_filename': report_filename,
            'processing_time': datetime.now().isoformat()
        }
        
        # 清理上傳的檔案
        cleanup_temp_files([upload_path])
        
        return jsonify(response_data)
        
    except Exception as e:
        print(f"處理錯誤: {str(e)}")
        return jsonify({'error': f'處理失敗: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下載生成的PDF檔案"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({'error': '檔案不存在'}), 404
    except Exception as e:
        return jsonify({'error': f'下載失敗: {str(e)}'}), 500

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
        
        # 生成PDF
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_filename = f"insurance_form_{timestamp}.pdf"
        pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
        
        pdf_filler.create_insurance_form(processed_data, pdf_path)
        
        return jsonify({
            'success': True,
            'extracted_data': processed_data,
            'validation_result': validation_result,
            'pdf_path': pdf_path,
            'pdf_filename': pdf_filename
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