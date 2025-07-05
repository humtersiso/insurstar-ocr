#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自動清理管理器
內建在系統中的智能清理機制，根據系統狀態自動清理暫存檔案
"""

import os
import time
import threading
import json
import psutil
import atexit
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Callable
from pathlib import Path
import logging

class AutoCleanupManager:
    """自動清理管理器"""
    
    def __init__(self, config_path: str = "cleanup_config.json"):
        self.config_path = config_path
        self.config = self.load_config()
        self.cleanup_thread = None
        self.is_running = False
        self.session_files = set()  # 記錄本次會話產生的檔案
        self.cleanup_callbacks = []  # 清理回調函數
        
        # 設定日誌
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('auto_cleanup.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        # 註冊程式結束時的清理
        atexit.register(self.cleanup_on_exit)
        
    def load_config(self) -> Dict:
        """載入設定檔"""
        default_config = {
            "auto_cleanup": {
                "enabled": True,
                "check_interval_seconds": 300,  # 5分鐘檢查一次
                "disk_threshold_mb": 500,
                "emergency_threshold_mb": 1000,
                "session_cleanup": True,
                "idle_timeout_minutes": 30
            },
            "cleanup_rules": {
                "ocr_results": {"enabled": True, "keep_days": 3},
                "property_reports": {"enabled": True, "keep_days": 7},
                "temp_images": {"enabled": True, "keep_days": 1},
                "uploads": {"enabled": True, "keep_days": 3},
                "pycache": {"enabled": True, "keep_days": 1}
            }
        }
        
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 合併預設設定
                    for key, value in default_config.items():
                        if key not in config:
                            config[key] = value
                        elif isinstance(value, dict):
                            config[key].update(value)
                    return config
        except Exception as e:
            print(f"載入設定檔失敗: {e}")
        
        return default_config
    
    def get_disk_usage(self) -> Dict[str, Dict[str, int]]:
        """取得磁碟使用量"""
        usage = {}
        temp_dirs = {
            "ocr_results": "ocr_results",
            "property_reports": "property_reports",
            "temp_images": "temp_images", 
            "uploads": "uploads",
            "pycache": "__pycache__"
        }
        
        for dir_name, dir_path in temp_dirs.items():
            if os.path.exists(dir_path):
                total_size = 0
                file_count = 0
                
                for root, dirs, files in os.walk(dir_path):
                    for file in files:
                        try:
                            file_path = os.path.join(root, file)
                            total_size += os.path.getsize(file_path)
                            file_count += 1
                        except:
                            pass
                
                usage[dir_name] = {
                    "size": total_size,
                    "files": file_count
                }
        
        return usage
    
    def format_size(self, size_bytes: int) -> str:
        """格式化檔案大小"""
        size_float = float(size_bytes)
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_float < 1024.0:
                return f"{size_float:.1f} {unit}"
            size_float /= 1024.0
        return f"{size_float:.1f} TB"
    
    def cleanup_old_files(self, days_old: int, dry_run: bool = False) -> Dict:
        """清理舊檔案"""
        cutoff_date = datetime.now() - timedelta(days=days_old)
        results = {"deleted": [], "failed": [], "total_size": 0}
        
        temp_dirs = {
            "ocr_results": "ocr_results",
            "property_reports": "property_reports",
            "temp_images": "temp_images",
            "uploads": "uploads",
            "pycache": "__pycache__"
        }
        
        for dir_name, dir_path in temp_dirs.items():
            if not os.path.exists(dir_path):
                continue
                
            rule = self.config["cleanup_rules"].get(dir_name, {})
            if not rule.get("enabled", True):
                continue
            
            keep_days = rule.get("keep_days", 7)
            if days_old < keep_days:
                continue
            
            for root, dirs, files in os.walk(dir_path):
                for file in files:
                    try:
                        file_path = os.path.join(root, file)
                        mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
                        
                        if mtime < cutoff_date:
                            if not dry_run:
                                os.remove(file_path)
                                self.session_files.discard(file_path)
                            
                            file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                            results["deleted"].append({
                                "path": file_path,
                                "size": file_size,
                                "modified": mtime
                            })
                            results["total_size"] += file_size
                            
                    except Exception as e:
                        results["failed"].append({
                            "path": file_path,
                            "error": str(e)
                        })
        
        return results
    
    def cleanup_session_files(self):
        """清理本次會話產生的檔案"""
        if not self.config["auto_cleanup"]["session_cleanup"]:
            return
        
        self.logger.info("清理本次會話產生的檔案...")
        deleted_count = 0
        
        for file_path in self.session_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    deleted_count += 1
            except Exception as e:
                self.logger.warning(f"無法刪除會話檔案 {file_path}: {e}")
        
        self.session_files.clear()
        self.logger.info(f"已清理 {deleted_count} 個會話檔案")
    
    def cleanup_on_exit(self):
        """程式結束時的清理"""
        self.logger.info("程式結束，執行清理...")
        self.cleanup_session_files()
        
        # 執行註冊的清理回調
        for callback in self.cleanup_callbacks:
            try:
                callback()
            except Exception as e:
                self.logger.error(f"清理回調執行失敗: {e}")
    
    def register_cleanup_callback(self, callback: Callable):
        """註冊清理回調函數"""
        self.cleanup_callbacks.append(callback)
    
    def add_session_file(self, file_path: str):
        """新增會話檔案"""
        self.session_files.add(file_path)
    
    def check_disk_usage(self) -> bool:
        """檢查磁碟使用量，決定是否需要清理"""
        usage = self.get_disk_usage()
        total_size_mb = sum(info["size"] for info in usage.values()) / (1024 * 1024)
        
        config = self.config["auto_cleanup"]
        emergency_threshold = config["emergency_threshold_mb"]
        normal_threshold = config["disk_threshold_mb"]
        
        if total_size_mb > emergency_threshold:
            self.logger.warning(f"磁碟使用量過高: {self.format_size(int(total_size_mb * 1024 * 1024))}，執行緊急清理")
            self.emergency_cleanup()
            return True
        elif total_size_mb > normal_threshold:
            self.logger.info(f"磁碟使用量較高: {self.format_size(int(total_size_mb * 1024 * 1024))}，執行一般清理")
            self.normal_cleanup()
            return True
        
        return False
    
    def emergency_cleanup(self):
        """緊急清理：清理所有超過 1 天的檔案"""
        self.logger.info("執行緊急清理...")
        results = self.cleanup_old_files(days_old=1, dry_run=False)
        
        deleted_size = self.format_size(results["total_size"])
        self.logger.info(f"緊急清理完成: 刪除 {len(results['deleted'])} 個檔案 ({deleted_size})")
        
        if results["failed"]:
            self.logger.warning(f"清理失敗 {len(results['failed'])} 個檔案")
    
    def normal_cleanup(self):
        """一般清理：根據設定清理舊檔案"""
        self.logger.info("執行一般清理...")
        
        for dir_name, rule in self.config["cleanup_rules"].items():
            if rule.get("enabled", True):
                keep_days = rule.get("keep_days", 7)
                results = self.cleanup_old_files(days_old=keep_days, dry_run=False)
                
                if results["deleted"]:
                    deleted_size = self.format_size(results["total_size"])
                    self.logger.info(f"清理 {dir_name}: 刪除 {len(results['deleted'])} 個檔案 ({deleted_size})")
    
    def start_monitoring(self):
        """開始監控"""
        if self.is_running:
            return
        
        self.is_running = True
        self.cleanup_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.cleanup_thread.start()
        self.logger.info("自動清理監控已啟動")
    
    def stop_monitoring(self):
        """停止監控"""
        self.is_running = False
        if self.cleanup_thread:
            self.cleanup_thread.join()
        self.logger.info("自動清理監控已停止")
    
    def _monitor_loop(self):
        """監控迴圈"""
        check_interval = self.config["auto_cleanup"]["check_interval_seconds"]
        idle_timeout = self.config["auto_cleanup"]["idle_timeout_minutes"]
        last_activity = time.time()
        
        while self.is_running:
            try:
                # 檢查磁碟使用量
                if self.check_disk_usage():
                    last_activity = time.time()
                
                # 檢查閒置時間
                idle_time = (time.time() - last_activity) / 60  # 分鐘
                if idle_time > idle_timeout:
                    self.logger.info(f"系統閒置 {idle_time:.1f} 分鐘，執行閒置清理")
                    self.idle_cleanup()
                    last_activity = time.time()
                
                time.sleep(check_interval)
                
            except Exception as e:
                self.logger.error(f"監控迴圈錯誤: {e}")
                time.sleep(60)  # 錯誤時等待 1 分鐘
    
    def idle_cleanup(self):
        """閒置清理"""
        self.logger.info("執行閒置清理...")
        # 清理較舊的檔案
        results = self.cleanup_old_files(days_old=3, dry_run=False)
        
        if results["deleted"]:
            deleted_size = self.format_size(results["total_size"])
            self.logger.info(f"閒置清理完成: 刪除 {len(results['deleted'])} 個檔案 ({deleted_size})")
    
    def get_status(self) -> Dict:
        """取得清理狀態"""
        usage = self.get_disk_usage()
        total_size_mb = sum(info["size"] for info in usage.values()) / (1024 * 1024)
        
        return {
            "is_running": self.is_running,
            "session_files_count": len(self.session_files),
            "disk_usage_mb": total_size_mb,
            "disk_usage_formatted": self.format_size(int(total_size_mb * 1024 * 1024)),
            "config": self.config["auto_cleanup"]
        }

# 全域實例
cleanup_manager = AutoCleanupManager()

def get_cleanup_manager() -> AutoCleanupManager:
    """取得清理管理器實例"""
    return cleanup_manager

def start_auto_cleanup():
    """啟動自動清理"""
    cleanup_manager.start_monitoring()

def stop_auto_cleanup():
    """停止自動清理"""
    cleanup_manager.stop_monitoring()

def add_session_file(file_path: str):
    """新增會話檔案"""
    cleanup_manager.add_session_file(file_path)

def register_cleanup_callback(callback: Callable):
    """註冊清理回調"""
    cleanup_manager.register_cleanup_callback(callback)

if __name__ == "__main__":
    # 測試模式
    print("自動清理管理器測試")
    print("=" * 50)
    
    manager = AutoCleanupManager()
    
    # 顯示狀態
    status = manager.get_status()
    print(f"監控狀態: {'運行中' if status['is_running'] else '已停止'}")
    print(f"會話檔案: {status['session_files_count']} 個")
    print(f"磁碟使用: {status['disk_usage_formatted']}")
    
    # 啟動監控
    manager.start_monitoring()
    
    try:
        print("\n按 Ctrl+C 停止監控...")
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        print("\n停止監控...")
        manager.stop_monitoring() 