# 自動清理機制完整指南

## 概述

保險財產分析書系統現在內建了智能自動清理機制，可以根據系統狀態自動管理暫存檔案，避免磁碟空間不足的問題。

## 核心功能

### 1. 自動監控
- **磁碟使用量監控**：每 5 分鐘檢查一次暫存檔案大小
- **閒置時間監控**：系統閒置 30 分鐘後自動清理
- **緊急清理**：當磁碟使用量超過 1GB 時觸發緊急清理

### 2. 會話檔案管理
- **自動追蹤**：系統自動記錄本次會話產生的檔案
- **程式結束清理**：當網頁關閉或程式結束時自動清理會話檔案
- **即時註冊**：上傳檔案和產生報告時自動註冊到會話

### 3. 智能清理策略
- **分級清理**：根據檔案類型和重要性設定不同的保留期限
- **安全清理**：只清理舊檔案，保留重要資料
- **可配置規則**：透過設定檔自訂清理規則

## 系統架構

```
自動清理管理器 (AutoCleanupManager)
├── 監控線程 (Monitoring Thread)
├── 會話檔案追蹤 (Session File Tracking)
├── 清理規則引擎 (Cleanup Rules Engine)
└── 日誌系統 (Logging System)
```

## 檔案結構

```
insurstar-ocr/
├── auto_cleanup_manager.py      # 自動清理管理器核心
├── cleanup_config.json          # 清理規則設定檔
├── cleanup_temp_files.py        # 手動清理工具
├── cleanup.bat                  # Windows 批次檔案
├── test_auto_cleanup.py         # 測試腳本
├── auto_cleanup.log             # 清理日誌
└── 暫存目錄/
    ├── ocr_results/             # OCR 結果 (保留 3 天)
    ├── property_reports/        # 財產分析書 (保留 7 天)
    ├── temp_images/             # 暫存圖片 (保留 1 天)
    ├── uploads/                 # 上傳檔案 (保留 3 天)
    └── __pycache__/             # Python 快取 (保留 1 天)
```

## 設定檔說明

### cleanup_config.json

```json
{
    "auto_cleanup": {
        "enabled": true,                    // 啟用自動清理
        "check_interval_seconds": 300,      // 檢查間隔 (5分鐘)
        "disk_threshold_mb": 500,           // 一般清理閾值 (500MB)
        "emergency_threshold_mb": 1000,     // 緊急清理閾值 (1GB)
        "session_cleanup": true,            // 啟用會話清理
        "idle_timeout_minutes": 30          // 閒置超時 (30分鐘)
    },
    "cleanup_rules": {
        "ocr_results": {
            "enabled": true,
            "keep_days": 3
        },
        "property_reports": {
            "enabled": true,
            "keep_days": 7
        },
        "temp_images": {
            "enabled": true,
            "keep_days": 1
        },
        "uploads": {
            "enabled": true,
            "keep_days": 3
        },
        "pycache": {
            "enabled": true,
            "keep_days": 1
        }
    }
}
```

## 使用方式

### 1. 自動啟動
系統啟動時會自動啟動清理管理器：

```python
# 在 app.py 中
from auto_cleanup_manager import start_auto_cleanup

# 啟動自動清理
start_auto_cleanup()
```

### 2. 手動控制
```python
from auto_cleanup_manager import (
    start_auto_cleanup, 
    stop_auto_cleanup, 
    add_session_file,
    get_cleanup_manager
)

# 啟動/停止監控
start_auto_cleanup()
stop_auto_cleanup()

# 註冊會話檔案
add_session_file("path/to/file.pdf")

# 取得狀態
manager = get_cleanup_manager()
status = manager.get_status()
```

### 3. API 端點
```bash
# 查看清理狀態
GET /api/cleanup/status

# 手動清理
POST /api/cleanup
```

### 4. 批次檔案
```bash
# 查看使用量
cleanup.bat --usage

# 預覽清理
cleanup.bat --preview

# 執行清理
cleanup.bat --execute

# 緊急清理
cleanup.bat --emergency
```

## 清理策略

### 1. 一般清理
- **觸發條件**：磁碟使用量 > 500MB
- **清理範圍**：根據設定檔規則清理舊檔案
- **保留期限**：按檔案類型設定 (1-7 天)

### 2. 緊急清理
- **觸發條件**：磁碟使用量 > 1GB
- **清理範圍**：清理所有超過 1 天的檔案
- **保留期限**：只保留當天檔案

### 3. 閒置清理
- **觸發條件**：系統閒置 > 30 分鐘
- **清理範圍**：清理超過 3 天的檔案
- **保留期限**：保留最近 3 天檔案

### 4. 會話清理
- **觸發條件**：程式結束或網頁關閉
- **清理範圍**：本次會話產生的所有檔案
- **保留期限**：立即清理

## 監控與日誌

### 日誌檔案
- **位置**：`auto_cleanup.log`
- **格式**：時間戳 - 等級 - 訊息
- **內容**：清理操作、錯誤訊息、狀態變更

### 監控指標
- 磁碟使用量
- 檔案數量統計
- 清理操作次數
- 錯誤發生次數

## 測試與驗證

### 執行測試
```bash
python test_auto_cleanup.py
```

### 測試內容
1. 自動清理機制測試
2. 設定檔功能測試
3. 會話檔案管理測試
4. 監控線程測試

## 故障排除

### 常見問題

**Q: 自動清理沒有啟動？**
A: 檢查 `cleanup_config.json` 中的 `enabled` 設定

**Q: 清理後檔案還在？**
A: 檢查檔案是否被其他程式鎖定，或檢查權限設定

**Q: 日誌檔案過大？**
A: 定期清理 `auto_cleanup.log` 檔案

**Q: 清理規則不生效？**
A: 檢查 `cleanup_rules` 設定，確認目錄名稱正確

### 除錯模式
```python
# 啟用詳細日誌
import logging
logging.getLogger('auto_cleanup_manager').setLevel(logging.DEBUG)
```

## 效能考量

### 記憶體使用
- 監控線程：約 1-2MB
- 會話檔案追蹤：根據檔案數量而定
- 日誌系統：約 100KB

### CPU 使用
- 檢查間隔：每 5 分鐘一次，影響極小
- 檔案掃描：只在需要時執行
- 清理操作：非阻塞式，不影響主程式

### 磁碟 I/O
- 日誌寫入：每次操作約 1KB
- 檔案刪除：只在清理時執行
- 狀態檢查：只讀取檔案資訊

## 安全注意事項

1. **備份重要檔案**：清理前請確認重要報告已備份
2. **權限檢查**：確保程式有足夠權限刪除檔案
3. **檔案鎖定**：避免刪除正在使用的檔案
4. **設定驗證**：定期檢查清理規則設定

## 未來擴展

### 計劃功能
- [ ] 雲端備份整合
- [ ] 清理統計報表
- [ ] 自訂清理腳本支援
- [ ] 多租戶隔離
- [ ] 清理策略學習

### 自訂開發
```python
# 自訂清理回調
def custom_cleanup_callback():
    # 執行自訂清理邏輯
    pass

register_cleanup_callback(custom_cleanup_callback)
```

## 版本記錄

- **v1.0**: 基礎自動清理功能
- **v1.1**: 會話檔案管理
- **v1.2**: 設定檔支援
- **v1.3**: API 整合
- **v1.4**: 監控與日誌系統 