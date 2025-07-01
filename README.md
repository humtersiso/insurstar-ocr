# 保單辨識系統 (Insurance Document OCR System)

基於 Gemini 2.5 Flash 的智能保單文字辨識與 PDF 填寫系統

## 🚀 功能特色

- **AI 驅動辨識**: 使用 Google Gemini 2.5 Flash 進行高精度文字辨識
- **智能資料提取**: 自動提取保單關鍵資訊（保單號碼、被保險人、保險公司等）
- **PDF 自動填寫**: 將辨識結果自動填入標準化 PDF 表單
- **Web 介面**: 友善的瀏覽器操作介面
- **品質分析**: 自動分析圖片品質並提供改善建議
- **多格式支援**: 支援 PNG、JPG、PDF、TIFF 等格式

## 📁 專案結構

```
insurstar_OCR/
├── app.py                      # Flask Web 應用主程式
├── gemini_ocr_processor.py     # Gemini OCR 處理器
├── pdf_filler.py              # PDF 填寫功能
├── data_processor.py          # 資料處理與驗證
├── policy_layouts.py          # 保單版面配置
├── pdf_to_images.py           # PDF 轉圖片工具
├── test_gemini_ocr.py         # Gemini OCR 測試腳本
├── cleanup.py                 # 專案清理工具
├── requirements.txt           # Python 依賴套件
├── README.md                  # 專案說明文件
├── templates/                 # HTML 模板
├── static/                    # CSS/JS 靜態檔案
├── insurtech_data/            # 保單資料目錄
├── converted_images/          # 轉換後的圖片
├── uploads/                   # 上傳檔案暫存
└── outputs/                   # 生成的 PDF 檔案
```

## 🛠️ 安裝與設定

### 1. 安裝依賴套件

```bash
pip install -r requirements.txt
```

### 2. 設定 Gemini API Key

在 `gemini_ocr_processor.py` 中設定您的 Gemini API Key：

```python
# 預設已設定，如需更改請修改以下行
api_key = "YOUR_GEMINI_API_KEY"
```

### 3. 準備保單資料

將保單 PDF 檔案放入 `insurtech_data/` 目錄中。

## 🚀 使用方式

### 方法一：Web 介面

1. 啟動 Web 服務：
```bash
python app.py
```

2. 開啟瀏覽器訪問：`http://localhost:5000`

3. 上傳保單圖片或 PDF，系統會自動：
   - 分析圖片品質
   - 使用 Gemini 提取文字
   - 生成填寫完成的 PDF
   - 提供下載連結

### 方法二：PDF 轉圖片 + OCR 測試

1. 將 PDF 轉換為圖片：
```bash
python pdf_to_images.py
```

2. 測試 Gemini OCR 功能：
```bash
python test_gemini_ocr.py
```

### 方法三：單一圖片測試

```bash
python test_gemini_ocr.py "path/to/image.png"
```

## 🔧 API 端點

- `POST /upload` - 上傳檔案並進行 OCR 處理
- `GET /download/<filename>` - 下載生成的 PDF
- `POST /api/process` - API 處理端點
- `GET /api/health` - 健康檢查
- `POST /api/analyze` - 圖片品質分析

## 📊 辨識欄位

系統會自動提取以下保單資訊：

- 保單號碼 (policy_number)
- 被保險人姓名 (insured_name)
- 保險公司名稱 (insurance_company)
- 保費金額 (premium_amount)
- 保險金額 (coverage_amount)
- 生效日期 (start_date)
- 到期日期 (end_date)
- 聯絡電話 (phone_number)
- 身分證字號 (id_number)

## 🧹 專案清理

如需清理測試檔案和快取：

```bash
python cleanup.py
```

## 🔍 技術架構

- **OCR 引擎**: Google Gemini 2.5 Flash
- **Web 框架**: Flask
- **PDF 處理**: PyMuPDF + ReportLab
- **圖片處理**: OpenCV + Pillow
- **資料處理**: 自定義處理器

## 📝 注意事項

1. 確保圖片品質良好（建議 300 DPI 以上）
2. Gemini API 有使用限制，請注意請求頻率
3. 大型 PDF 檔案處理時間較長，請耐心等待
4. 建議在處理前先使用品質分析功能

## 🤝 貢獻

歡迎提交 Issue 和 Pull Request 來改善專案！

## 📄 授權

本專案採用 MIT 授權條款。 