import os
import datetime
import threading
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
from word_template_processor import WordTemplateProcessorPure
from auto_cleanup_manager import start_auto_cleanup, stop_auto_cleanup, add_session_file, get_cleanup_manager

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

# 啟動時不清理 temp_images 資料夾，保留預覽圖片
# def cleanup_temp_previews_on_start(folder):
#     for filename in os.listdir(folder):
#         file_path = os.path.join(folder, filename)
#         if os.path.isfile(file_path):
#             try:
#                 os.remove(file_path)
#             except Exception as e:
#                 print(f"刪除暫存檔案失敗: {file_path} - {e}")

# cleanup_temp_previews_on_start(TEMP_PREVIEW_FOLDER)

# 初始化處理器
ocr_processor = GeminiOCRProcessor()  # 使用 Gemini OCR
data_processor = DataProcessor()
image_processor = ImageProcessing

# 初始化 Word 模板處理器
try:
    word_template_processor = WordTemplateProcessorPure()  # Word 模板處理器
    print("✅ Word 模板處理器初始化成功")
except Exception as e:
    print(f"❌ Word 模板處理器初始化失敗: {str(e)}")
    word_template_processor = None

# 啟動自動清理管理器
try:
    start_auto_cleanup()
    print("✅ 自動清理管理器已啟動")
except Exception as e:
    print(f"❌ 自動清理管理器啟動失敗: {str(e)}")

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
        filename = secure_filename(file.filename or 'unknown')
        file_ext = filename.rsplit('.', 1)[1].lower()
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{filename}.{file_ext}")
        
        # 保存上傳的檔案
        file.save(upload_path)
        
        # 註冊會話檔案（用於自動清理）
        add_session_file(upload_path)

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

        # OCR 處理（無論是否有預覽圖都執行）
        raw_data = ocr_processor.extract_insurance_data_with_gemini(image_path_for_ocr)
        # 資料處理
        processed_data = data_processor.process_insurance_data(raw_data)

        # ====== 背景產生財產分析書（Word+PDF） ======
        # 在主流程產生唯一檔名（只用時間戳）
        now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        word_filename = f'財產分析書_{now}.docx'
        word_path = os.path.join('property_reports', word_filename)
        # 啟動 thread 時，把 word_path 傳進去
        def generate_property_report(processed_data, word_path):
            try:
                word_template_processor.fill_template(processed_data, output_path=word_path)
                # 註冊產生的 Word 檔案（用於自動清理）
                add_session_file(word_path)
            except Exception as e:
                print(f'背景產生財產分析書失敗: {e}')
        if word_template_processor is not None:
            threading.Thread(target=generate_property_report, args=(processed_data, word_path)).start()
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
        
        # 如果無法產生預覽圖，添加提醒
        if not overlay_preview_path:
            reminders.append("❌ 無法產生預覽圖，請檢查檔案格式")

        # 自動清理舊資料（每次新的要保書辨識完成後）
        # auto_cleanup_old_data()  # 註解自動清理呼叫
        
        response_data = {
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
            'extraction_rate': f"{sum(1 for v in processed_data.values() if v and v != '無填寫') / 20 * 100:.1f}%" if processed_data else "0.0%",
            'word_filename': word_filename,
            'pdf_filename': word_filename.replace('.docx', '.pdf')
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
        
        # 先執行 OCR 取得資料
        raw_data = ocr_processor.extract_insurance_data_with_gemini(image_path)
        if not raw_data:
            return jsonify({'error': 'OCR 辨識失敗'}), 500
        
        # 處理資料
        processed_data = data_processor.process_insurance_data(raw_data)
        
        # 檢查 Word 模板處理器是否已初始化
        if word_template_processor is None:
            return jsonify({'error': 'Word 模板處理器未初始化'}), 500
        
        # 使用 Word 模板處理器生成檔案
        output_path = word_template_processor.fill_template(processed_data)
        
        # === 新增：自動轉 PDF ===
        if output_path and output_path.endswith('.docx'):
            pdf_path = output_path.replace('.docx', '.pdf')
            if not os.path.exists(pdf_path):
                try:
                    import pythoncom
                    from docx2pdf import convert
                    pythoncom.CoInitialize()
                    convert(output_path, pdf_path)
                    pythoncom.CoUninitialize()
                except Exception as e:
                    print(f"[自動轉 PDF 失敗] {e}")
        
        if output_path:
            filename = os.path.basename(output_path)
            return jsonify({
                'success': True,
                'word_filename': filename,
                'word_path': output_path,
                'download_url': f"/download/{filename}",
                'data_summary': data_processor.get_data_summary(processed_data)
            })
        else:
            return jsonify({'error': 'Word 檔案生成失敗'}), 500
        
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
        
        # 檢查 Word 模板處理器是否已初始化
        if word_template_processor is None:
            return jsonify({'error': 'Word 模板處理器未初始化'}), 500
        
        # 使用Word模板處理器
        output_path = word_template_processor.fill_template(ocr_data)
        
        # === 新增：自動轉 PDF ===
        if output_path and output_path.endswith('.docx'):
            pdf_path = output_path.replace('.docx', '.pdf')
            if not os.path.exists(pdf_path):
                try:
                    import pythoncom
                    from docx2pdf import convert
                    pythoncom.CoInitialize()
                    convert(output_path, pdf_path)
                    pythoncom.CoUninitialize()
                except Exception as e:
                    print(f"[自動轉 PDF 失敗] {e}")
        
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

@app.route('/api/generate-property-analysis', methods=['POST'])
def generate_property_analysis():
    """生成財產分析書（整合 information_integration 功能）"""
    try:
        data = request.get_json()
        
        if not data or 'extracted_data' not in data:
            return jsonify({'error': '缺少必要參數'}), 400
        
        extracted_data = data['extracted_data']
        
        # 檢查 Word 模板處理器是否已初始化
        if word_template_processor is None:
            return jsonify({'error': 'Word 模板處理器未初始化'}), 500
        
        # 使用Word模板處理器生成財產分析書
        output_path = word_template_processor.fill_template(extracted_data)
        
        # === 新增：自動轉 PDF ===
        if output_path and output_path.endswith('.docx'):
            pdf_path = output_path.replace('.docx', '.pdf')
            if not os.path.exists(pdf_path):
                try:
                    import pythoncom
                    from docx2pdf import convert
                    pythoncom.CoInitialize()
                    convert(output_path, pdf_path)
                    pythoncom.CoUninitialize()
                except Exception as e:
                    print(f"[自動轉 PDF 失敗] {e}")
        
        if output_path:
            filename = os.path.basename(output_path)
            response_data = {
                'success': True,
                'analysis_filename': filename,
                'analysis_path': output_path,
                'download_url': f"/download/{filename}",
                'message': '財產分析書生成成功'
            }
            
            return jsonify(response_data)
        else:
            return jsonify({'error': '財產分析書生成失敗'}), 500
        
    except Exception as e:
        return jsonify({'error': f'財產分析書生成失敗: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下載檔案"""
    try:
        return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': f'下載失敗: {str(e)}'}), 404

@app.route('/preview/<filename>')
def preview_file(filename):
    """預覽財產分析書"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if not os.path.exists(file_path):
            return jsonify({'error': '檔案不存在'}), 404
        
        # 如果是PDF檔案，直接返回
        if filename.lower().endswith('.pdf'):
            return send_from_directory(app.config['OUTPUT_FOLDER'], filename)
        
        # 如果是Word檔案，轉換為PDF後返回
        elif filename.lower().endswith('.docx'):
            pdf_filename = filename.replace('.docx', '.pdf')
            pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
            
            if os.path.exists(pdf_path):
                return send_from_directory(app.config['OUTPUT_FOLDER'], pdf_filename)
            else:
                # 如果PDF不存在，嘗試轉換
                try:
                    import docx2pdf
                    docx2pdf.convert(file_path, pdf_path)
                    return send_from_directory(app.config['OUTPUT_FOLDER'], pdf_filename)
                except Exception as e:
                    return jsonify({'error': f'PDF轉換失敗: {str(e)}'}), 500
        
        else:
            return jsonify({'error': '不支援的檔案格式'}), 400
            
    except Exception as e:
        return jsonify({'error': f'預覽失敗: {str(e)}'}), 500

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
        filename = secure_filename(file.filename or 'unknown')
        file_ext = filename.rsplit('.', 1)[1].lower()
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{filename}.{file_ext}")
        
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
            # cleanup_temp_files([temp_path])
            pass
        
    except Exception as e:
        return jsonify({'error': f'分析失敗: {str(e)}'}), 500

@app.route('/temp_images/<path:filename>')
def temp_images_preview(filename):
    return send_from_directory(TEMP_PREVIEW_FOLDER, filename)

@app.route('/api/preview', methods=['POST'])
def api_preview():
    """支援要保書疊合輔助線預覽與財產分析書 docx 轉 PDF 預覽"""
    try:
        data = request.get_json()
        print(f"[DEBUG] 收到 /api/preview 請求，data: {data}")
        # 財產分析書預覽
        if 'filename' in data:
            filename = data['filename']
            print(f"[DEBUG] filename: {filename}")
            file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
            print(f"[DEBUG] file_path: {file_path}")
            print(f"[DEBUG] 檢查 docx 檔案是否存在: {os.path.exists(file_path)}")
            if not os.path.exists(file_path):
                print(f"[DEBUG] 檔案不存在: {file_path}")
                return jsonify({'error': '檔案不存在'}), 404
            # 如果是 docx 檔案，先轉成 PDF
            if filename.lower().endswith('.docx'):
                print(f"[DEBUG] 檔案為 docx，準備進行 PDF 轉檔流程")
                pdf_filename = filename.replace('.docx', '.pdf')
                pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
                print(f"[DEBUG] 預期 PDF 路徑: {pdf_path}")
                # 如果 PDF 不存在，則轉換
                if not os.path.exists(pdf_path):
                    print(f"[DEBUG] PDF 不存在，開始轉檔...")
                    try:
                        import pythoncom
                        from docx2pdf import convert
                        pythoncom.CoInitialize()
                        print(f"🔄 準備將 {file_path} 轉成 PDF（docx2pdf）...")
                        convert(file_path, pdf_path)
                        print(f"✅ PDF 產生成功: {pdf_path}")
                    except Exception as e:
                        print(f"❌ PDF 轉換失敗: {e}")
                        return jsonify({'error': f'PDF 轉換失敗: {e}'}), 500
                    finally:
                        try:
                            pythoncom.CoUninitialize()
                        except Exception:
                            pass
                else:
                    print(f"[DEBUG] PDF 已存在: {pdf_path}")
            # 產生 PNG 預覽
            try:
                import fitz
                print(f"[DEBUG] 開始產生 PNG 預覽: {pdf_path}")
                doc = fitz.open(pdf_path)
                page = doc.load_page(0)
                pix = page.get_pixmap(dpi=150)
                os.makedirs('temp_images', exist_ok=True)
                preview_path = f'temp_images/preview_{os.path.splitext(filename)[0]}.png'
                pix.save(preview_path)
                print(f"✅ 預覽圖片產生成功: {preview_path}")
                return jsonify({'preview_url': f'/{preview_path}'})
            except Exception as e:
                print(f"❌ 預覽圖片產生失敗: {e}")
                return jsonify({'error': f'預覽圖片產生失敗: {e}'}), 500
        # 要保書疊合輔助線預覽
        image_path = data.get('image_path')
        if not image_path or not os.path.exists(image_path):
            return jsonify({'error': '圖片檔案不存在'}), 404
        preview_path = os.path.join(TEMP_PREVIEW_FOLDER, f'{filename}_preview.png')
        support_line_path = os.path.join('assets', 'watermark', 'table_line_redline_only.png')
        ImageProcessing.overlay_support_line_on_image(
            image_path, support_line_path, preview_path,
            orig_size=(2481, 3508), crop_box=(81, 1157, 1797, 747), alpha=1.0
        )
        overlay_preview_url = f"/temp_images/{filename}_preview.png"
        return jsonify({'preview_url': overlay_preview_url})
    except Exception as e:
        print(f"❌ /api/preview 發生未預期錯誤: {e}")
        return jsonify({'error': f'預覽發生未預期錯誤: {e}'}), 500

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
        filename = secure_filename(file.filename or 'unknown')
        file_ext = filename.rsplit('.', 1)[1].lower()
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{filename}.{file_ext}")
        
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
        raw_data = ocr_processor.extract_insurance_data_with_gemini(image_path_for_ocr)
        
        if not raw_data:
            return jsonify({'error': 'OCR 辨識失敗'}), 500
        
        # 資料處理
        print("🔄 處理資料...")
        processed_data = data_processor.process_insurance_data(raw_data)
        
        # 準備回應資料
        response_data = {
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

@app.route('/api/cleanup', methods=['POST'])
def cleanup_data():
    """清理舊資料"""
    try:
        data = request.get_json() or {}
        cleanup_type = data.get('type', 'all')  # all, uploads, ocr_results, property_reports
        
        cleaned_files = []
        total_size_freed = 0
        
        if cleanup_type in ['all', 'uploads']:
            # 清理uploads資料夾（保留最近5個檔案）
            # uploads_cleaned = cleanup_old_files(app.config['UPLOAD_FOLDER'], keep_count=5)
            # cleaned_files.extend(uploads_cleaned['files'])
            # total_size_freed += uploads_cleaned['size_freed']
            pass
        
        if cleanup_type in ['all', 'ocr_results']:
            # 清理ocr_results資料夾（保留最近10個檔案）
            # ocr_cleaned = cleanup_old_files('ocr_results', keep_count=10)
            # cleaned_files.extend(ocr_cleaned['files'])
            # total_size_freed += ocr_cleaned['size_freed']
            pass
        
        if cleanup_type in ['all', 'property_reports']:
            # 清理property_reports資料夾（保留最近5個檔案）
            # reports_cleaned = cleanup_old_files(app.config['OUTPUT_FOLDER'], keep_count=5)
            # cleaned_files.extend(reports_cleaned['files'])
            # total_size_freed += reports_cleaned['size_freed']
            pass
        
        # 清理temp_images資料夾
        # temp_cleaned = cleanup_temp_files_in_folder('temp_images')
        # cleaned_files.extend(temp_cleaned['files'])
        # total_size_freed += temp_cleaned['size_freed']
        
        return jsonify({
            'success': True,
            'cleaned_files': cleaned_files,
            'total_files_cleaned': len(cleaned_files),
            'total_size_freed_mb': round(total_size_freed / (1024 * 1024), 2),
            'message': f'成功清理 {len(cleaned_files)} 個檔案，釋放 {round(total_size_freed / (1024 * 1024), 2)} MB 空間'
        })
        
    except Exception as e:
        return jsonify({'error': f'清理失敗: {str(e)}'}), 500

@app.route('/api/cleanup/status', methods=['GET'])
def cleanup_status():
    """取得自動清理狀態"""
    try:
        manager = get_cleanup_manager()
        status = manager.get_status()
        
        return jsonify({
            'success': True,
            'status': status,
            'message': '自動清理狀態查詢成功'
        })
        
    except Exception as e:
        return jsonify({'error': f'查詢狀態失敗: {str(e)}'}), 500

def cleanup_old_files(folder_path: str, keep_count: int = 5) -> dict:
    """清理舊檔案，保留指定數量的最新檔案"""
    try:
        if not os.path.exists(folder_path):
            return {'files': [], 'size_freed': 0}
        
        # 獲取所有檔案及其修改時間
        files = []
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                mtime = os.path.getmtime(file_path)
                size = os.path.getsize(file_path)
                files.append((file_path, filename, mtime, size))
        
        # 按修改時間排序（最新的在前）
        files.sort(key=lambda x: x[2], reverse=True)
        
        # 保留最新的檔案，刪除其餘的
        files_to_delete = files[keep_count:]
        deleted_files = []
        total_size_freed = 0
        
        for file_path, filename, mtime, size in files_to_delete:
            try:
                os.remove(file_path)
                deleted_files.append(filename)
                total_size_freed += size
            except Exception as e:
                print(f"刪除檔案失敗 {filename}: {str(e)}")
        
        return {
            'files': deleted_files,
            'size_freed': total_size_freed
        }
        
    except Exception as e:
        print(f"清理資料夾失敗 {folder_path}: {str(e)}")
        return {'files': [], 'size_freed': 0}

def auto_cleanup_old_data():
    """自動清理舊資料（每次新的要保書辨識完成後執行）"""
    try:
        print("🧹 自動清理舊資料...")
        
        cleaned_files = []
        total_size_freed = 0
        
        # 清理uploads資料夾（保留最近2個檔案，清理之前的）
        # uploads_cleaned = cleanup_old_files(app.config['UPLOAD_FOLDER'], keep_count=2)
        # cleaned_files.extend(uploads_cleaned['files'])
        # total_size_freed += uploads_cleaned['size_freed']
        
        # 清理ocr_results資料夾（保留最近3個檔案，清理之前的）
        # ocr_cleaned = cleanup_old_files('ocr_results', keep_count=3)
        # cleaned_files.extend(ocr_cleaned['files'])
        # total_size_freed += ocr_cleaned['size_freed']
        
        # 清理property_reports資料夾（保留最近2個檔案，清理之前的）
        # reports_cleaned = cleanup_old_files(app.config['OUTPUT_FOLDER'], keep_count=2)
        # cleaned_files.extend(reports_cleaned['files'])
        # total_size_freed += reports_cleaned['size_freed']
        
        # 不清理temp_images資料夾，保留當前預覽圖片供前端使用
        # temp_cleaned = cleanup_temp_files_in_folder('temp_images')
        # cleaned_files.extend(temp_cleaned['files'])
        # total_size_freed += temp_cleaned['size_freed']
        
        if cleaned_files:
            print(f"✅ 自動清理完成：刪除 {len(cleaned_files)} 個檔案，釋放 {round(total_size_freed / (1024 * 1024), 2)} MB 空間")
        else:
            print("✅ 無需清理的檔案")
            
    except Exception as e:
        print(f"❌ 自動清理失敗: {str(e)}")

def cleanup_temp_files_in_folder(folder_path: str) -> dict:
    """清理暫存檔案資料夾"""
    try:
        if not os.path.exists(folder_path):
            return {'files': [], 'size_freed': 0}
        
        deleted_files = []
        total_size_freed = 0
        
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                try:
                    size = os.path.getsize(file_path)
                    os.remove(file_path)
                    deleted_files.append(filename)
                    total_size_freed += size
                except Exception as e:
                    print(f"刪除暫存檔案失敗 {filename}: {str(e)}")
        
        return {
            'files': deleted_files,
            'size_freed': total_size_freed
        }
        
    except Exception as e:
        print(f"清理暫存資料夾失敗 {folder_path}: {str(e)}")
        return {'files': [], 'size_freed': 0}

@app.errorhandler(500)
def internal_error(e):
    """500錯誤處理"""
    return jsonify({'error': '伺服器內部錯誤'}), 500

if __name__ == '__main__':
    print("啟動保單辨識系統...")
    print("請訪問: http://localhost:5000")
    
    try:
        # 在 Windows 上使用 threaded=True 和 use_reloader=False 來避免通訊端錯誤
        app.run(
            debug=True, 
            host='0.0.0.0', 
            port=5000,
            threaded=True,
            use_reloader=False  # 關閉自動重載以避免通訊端錯誤
        )
    except KeyboardInterrupt:
        print("\n正在關閉系統...")
    except OSError as e:
        if "嘗試操作的對象不是通訊端" in str(e):
            print("⚠️  檢測到通訊端錯誤，這通常是正常的關閉過程")
        else:
            print(f"❌ 系統錯誤: {str(e)}")
    finally:
        # 程式結束時停止自動清理
        try:
            stop_auto_cleanup()
            print("✅ 自動清理管理器已停止")
        except Exception as e:
            print(f"❌ 停止自動清理管理器失敗: {str(e)}") 