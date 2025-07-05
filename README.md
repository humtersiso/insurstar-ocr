# 保險財產分析書系統 (Insurance Property Analysis System)

基於 Gemini 2.5 Flash 的智能保單文字辨識與財產分析書自動生成系統

## 🚀 功能特色

- **AI 驅動辨識**: 使用 Google Gemini 2.5 Flash 進行高精度文字辨識
- **專注版 OCR Prompt**: 特別優化被保險人和要保人出生日期及性別欄位辨識
- **智能資料提取**: 自動提取保單關鍵資訊（20個核心欄位）
- **Word 模板處理**: 將辨識結果自動填入 Word 模板生成財產分析書
- **PDF 自動填寫**: 將辨識結果自動填入標準化 PDF 表單
- **Web 介面**: 友善的瀏覽器操作介面
- **品質分析**: 自動分析圖片品質並提供改善建議
- **多格式支援**: 支援 PNG、JPG、PDF、TIFF 等格式
- **自動清理機制**: 智能管理暫存檔案，無需手動維護

## 📁 專案結構

```
insurstar-ocr/
├── app.py                      # Flask Web 應用主程式
├── auto_cleanup_manager.py     # 自動清理管理器
├── gemini_ocr_processor.py     # Gemini OCR 處理器 (專注版)
├── word_template_processor.py  # Word 模板處理器
├── data_processor.py          # 資料處理與驗證
├── image_processing.py        # 圖片處理功能
├── pdf_to_images.py           # PDF 轉圖片工具
├── requirements.txt           # Python 依賴套件
├── README.md                  # 專案說明文件
├── templates/                 # HTML 模板
├── static/                    # CSS/JS 靜態檔案
├── assets/                    # 模板與資源檔案
│   ├── templates/            # Word 模板檔案
│   ├── ocr_fields/          # OCR 欄位對照表
│   └── watermark/           # 浮水印資源
├── insurtech_data/            # 保單資料目錄
├── ocr_results/               # OCR 結果檔案 (自動清理)
├── property_reports/          # 財產分析書檔案 (自動清理)
├── temp_images/               # 暫存圖片 (自動清理)
├── uploads/                   # 上傳檔案暫存 (自動清理)
└── __pycache__/               # Python 快取 (自動清理)
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

### 方法一：Web 介面 (推薦)

1. 啟動 Web 服務：
```bash
python app.py
```

2. 開啟瀏覽器訪問：`http://localhost:5000`

3. 上傳保單圖片或 PDF，系統會自動：
   - 分析圖片品質
   - 使用專注版 Gemini OCR 提取文字
   - 生成財產分析書 Word 檔案
   - 轉換為 PDF 格式
   - 提供下載連結

### 方法二：PDF 轉圖片 + OCR 測試

1. 將 PDF 轉換為圖片：
```bash
python pdf_to_images.py
```

2. 測試專注版 Gemini OCR 功能：
```bash
python test_ocr_focus.py
```

### 方法三：單一圖片測試

```bash
python test_ocr_focus.py "path/to/image.png"
```

## 🔧 API 端點

- `POST /upload` - 上傳檔案並進行 OCR 處理
- `GET /download/<filename>` - 下載生成的 Word/PDF
- `POST /api/process` - API 處理端點
- `GET /api/health` - 健康檢查
- `POST /api/analyze` - 圖片品質分析
- `GET /api/cleanup/status` - 自動清理狀態查詢

## 📊 辨識欄位 (20個核心欄位)

系統使用專注版 OCR Prompt 自動提取以下保單資訊：

### 被保險人區塊
- 產險公司 (insurance_company)
- 被保險人區塊 (insured_section)
- 被保險人 (insured_person)
- 法人之代表人 (legal_representative)
- 身分證字號/統一編號 (id_number)
- **出生日期 (birth_date)** - 專注辨識
- **性別 (gender)** - 專注辨識

### 要保人區塊
- 要保人區塊 (policyholder_section)
- 要保人 (policyholder)
- 與被保險人關係 (relationship)
- 法人之代表人 (policyholder_legal_representative)
- **性別 (policyholder_gender)** - 專注辨識
- 要保人身份證字號/統編 (policyholder_id)
- **出生日期 (policyholder_birth_date)** - 專注辨識

### 車輛資訊
- 車輛種類 (vehicle_type)
- 牌照號碼 (license_number)

### 承保內容
- 承保內容 (coverage_items) - 結構化陣列

### 保費與保期
- 總保險費 (total_premium)
- 強制險保期 (compulsory_insurance_period)
- 任意險保期 (optional_insurance_period)

## 🎯 專注版 OCR 特色

### 重點觀察欄位
- **被保險人出生日期與性別**: 特別針對公司被保險人的個人資料欄位
- **要保人出生日期與性別**: 確保個人資料的完整辨識
- **關係交叉比對**: 當關係為「本人」時，自動比對兩個區塊的個人資料

### 智能辨識邏輯
- 被保險人可能是公司名稱，但個人資料欄位會填入實際負責人的資訊
- 當關係為「本人」時，被保險人和要保人的個人資料通常一致
- 出生日期和性別欄位可能標示為「出生日期」、「性別」或類似文字
- 自動處理空白欄位，統一填入「無填寫」

## 📄 Word 模板處理

### 財產分析書生成
- 使用 `assets/templates/` 中的 Word 模板
- 自動填入 OCR 辨識結果
- 支援表格線和浮水印處理
- 生成專業的財產分析書

### 模板欄位對應
- 詳細的欄位對照表位於 `assets/ocr_fields/OCR欄位對照表.md`
- 自動處理欄位格式轉換
- 支援複雜的表格結構

## 🧹 自動清理機制

系統內建智能自動清理機制，無需手動管理暫存檔案：

### 自動清理功能
- **磁碟監控**: 每 5 分鐘檢查暫存檔案大小
- **閒置清理**: 系統閒置 30 分鐘後自動清理
- **緊急清理**: 磁碟使用量超過 1GB 時觸發緊急清理
- **會話管理**: 程式結束時自動清理本次會話產生的檔案

### 清理策略
| 目錄 | 保留期限 | 說明 |
|------|----------|------|
| `ocr_results/` | 3 天 | OCR 結果 JSON 檔案 |
| `property_reports/` | 7 天 | 財產分析書 Word/PDF 檔案 |
| `temp_images/` | 1 天 | 暫存圖片檔案 |
| `uploads/` | 3 天 | 上傳的 PDF 檔案 |
| `__pycache__/` | 1 天 | Python 快取檔案 |

### 清理觸發條件
- **一般清理**: 磁碟使用量 > 500MB
- **緊急清理**: 磁碟使用量 > 1GB
- **閒置清理**: 系統閒置 > 30 分鐘
- **會話清理**: 程式結束或網頁關閉

### 監控與日誌
- 清理日誌: `auto_cleanup.log`
- 狀態查詢: `GET /api/cleanup/status`
- 自動啟動: 系統啟動時自動啟用

## 🔍 技術架構

- **OCR 引擎**: Google Gemini 2.5 Flash (專注版 Prompt)
- **Web 框架**: Flask
- **Word 處理**: python-docx + docxtpl
- **PDF 處理**: PyMuPDF
- **圖片處理**: OpenCV + Pillow
- **資料處理**: 自定義處理器
- **自動清理**: 內建清理管理器

## 📝 使用注意事項

1. **圖片品質**: 確保圖片品質良好（建議 300 DPI 以上）
2. **API 限制**: Gemini API 有使用限制，請注意請求頻率
3. **處理時間**: 大型 PDF 檔案處理時間較長，請耐心等待
4. **品質分析**: 建議在處理前先使用品質分析功能
5. **自動清理**: 系統會自動管理暫存檔案，無需手動清理
6. **重要報告**: 建議及時下載備份財產分析書
7. **專注辨識**: 系統特別優化被保險人和要保人的出生日期及性別欄位辨識

## 🆕 最新更新

### v2.0 - 專注版 OCR (2024-12-19)
- ✅ 整合專注版 OCR Prompt，特別優化被保險人和要保人欄位辨識
- ✅ 加強出生日期和性別欄位的辨識準確性
- ✅ 新增被保險人與要保人關係的交叉比對邏輯
- ✅ 優化 Word 模板處理功能
- ✅ 完善財產分析書生成流程

### v1.0 - 基礎版本
- ✅ 基礎 OCR 辨識功能
- ✅ Web 介面
- ✅ 自動清理機制
- ✅ Word 模板處理

## 🤝 貢獻

歡迎提交 Issue 和 Pull Request 來改善專案！

## 📄 授權

本專案採用 MIT 授權條款。 