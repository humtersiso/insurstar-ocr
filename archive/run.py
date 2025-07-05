#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
保單辨識系統啟動腳本
"""

import os
import sys
import subprocess
from pathlib import Path

def check_dependencies():
    """檢查依賴套件"""
    print("檢查依賴套件...")
    
    required_packages = [
        'flask',
        'paddlepaddle',
        'paddleocr',
        'opencv-python',
        'PyPDF2',
        'reportlab'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
            print(f"✓ {package}")
        except ImportError:
            missing_packages.append(package)
            print(f"✗ {package} (未安裝)")
    
    if missing_packages:
        print(f"\n缺少以下套件: {', '.join(missing_packages)}")
        print("請執行: pip install -r requirements.txt")
        return False
    
    return True

def create_directories():
    """建立必要的目錄"""
    directories = ['uploads', 'property_reports', 'assets/templates', 'static']
    
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)
        print(f"✓ 目錄已建立: {directory}")

def download_ocr_models():
    """下載OCR模型（如果需要）"""
    print("檢查OCR模型...")
    
    try:
        from paddleocr import PaddleOCR
        # 初始化PaddleOCR會自動下載模型
        ocr = PaddleOCR(use_angle_cls=True, lang='ch')
        print("✓ OCR模型已準備就緒")
        return True
    except Exception as e:
        print(f"⚠ OCR模型下載可能失敗: {str(e)}")
        print("首次使用時會自動下載，請確保網路連線正常")
        return True

def start_server():
    """啟動Web伺服器"""
    print("\n啟動保單辨識系統...")
    print("=" * 50)
    
    try:
        # 檢查app.py是否存在
        if not os.path.exists('app.py'):
            print("✗ 找不到 app.py 檔案")
            return False
        
        # 啟動Flask應用
        print("正在啟動Web伺服器...")
        print("請在瀏覽器中訪問: http://localhost:5000")
        print("按 Ctrl+C 停止伺服器")
        print("=" * 50)
        
        # 使用subprocess啟動Flask應用
        subprocess.run([sys.executable, 'app.py'])
        
    except KeyboardInterrupt:
        print("\n\n伺服器已停止")
    except Exception as e:
        print(f"啟動失敗: {str(e)}")
        return False
    
    return True

def main():
    """主函數"""
    print("保單辨識系統啟動器")
    print("=" * 50)
    
    # 檢查依賴
    if not check_dependencies():
        return
    
    # 建立目錄
    create_directories()
    
    # 檢查OCR模型
    download_ocr_models()
    
    # 啟動伺服器
    start_server()

if __name__ == "__main__":
    main() 