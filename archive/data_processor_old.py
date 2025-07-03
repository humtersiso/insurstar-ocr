import re
import json
from datetime import datetime
from typing import Dict, List, Any, Optional

class DataProcessor:
    """資料處理器"""
    
    def __init__(self):
        """初始化資料處理器"""
        self.processed_data = {}
    
    def clean_text(self, text: str) -> str:
        """
        清理文字資料
        
        Args:
            text: 原始文字
            
        Returns:
            清理後的文字
        """
        if not text:
            return ""
        
        # 移除多餘空白
        text = re.sub(r'\s+', ' ', text.strip())
        
        # 移除特殊字符（保留中文、英文、數字、常用符號）
        text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\s\-\.\,\:\/\(\)]', '', text)
        
        return text
    
    def format_policy_number(self, policy_number: str) -> str:
        """
        格式化保單號碼
        
        Args:
            policy_number: 原始保單號碼
            
        Returns:
            格式化後的保單號碼
        """
        if not policy_number:
            return ""
        
        # 移除所有非字母數字字符
        cleaned = re.sub(r'[^A-Za-z0-9]', '', policy_number)
        
        # 轉換為大寫
        cleaned = cleaned.upper()
        
        # 如果長度合理，返回格式化結果
        if 8 <= len(cleaned) <= 20:
            return cleaned
        
        return policy_number
    
    def format_amount(self, amount: str) -> str:
        """
        格式化金額
        
        Args:
            amount: 原始金額字串
            
        Returns:
            格式化後的金額
        """
        if not amount:
            return ""
        
        # 移除所有非數字字符
        cleaned = re.sub(r'[^\d]', '', amount)
        
        if cleaned:
            # 格式化為千分位
            try:
                num = int(cleaned)
                return f"{num:,}"
            except:
                return amount
        
        return amount
    
    def format_date(self, date_str: str) -> str:
        """
        格式化日期
        
        Args:
            date_str: 原始日期字串
            
        Returns:
            格式化後的日期 (YYYY/MM/DD)
        """
        if not date_str:
            return ""
        
        # 嘗試多種日期格式
        date_patterns = [
            r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})',
            r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})',
            r'(\d{4})(\d{2})(\d{2})'
        ]
        
        for pattern in date_patterns:
            match = re.search(pattern, date_str)
            if match:
                if len(match.group(1)) == 4:  # YYYY-MM-DD
                    year, month, day = match.groups()
                else:  # MM-DD-YYYY
                    month, day, year = match.groups()
                
                try:
                    # 驗證日期有效性
                    datetime(int(year), int(month), int(day))
                    return f"{year}/{month.zfill(2)}/{day.zfill(2)}"
                except ValueError:
                    continue
        
        return date_str
    
    def format_phone_number(self, phone: str) -> str:
        """
        格式化電話號碼
        
        Args:
            phone: 原始電話號碼
            
        Returns:
            格式化後的電話號碼
        """
        if not phone:
            return ""
        
        # 移除所有非數字字符
        cleaned = re.sub(r'[^\d]', '', phone)
        
        if len(cleaned) == 10:  # 手機號碼
            return f"{cleaned[:4]}-{cleaned[4:7]}-{cleaned[7:]}"
        elif len(cleaned) == 8:  # 市話
            return f"{cleaned[:4]}-{cleaned[4:]}"
        elif len(cleaned) == 9:  # 市話
            return f"{cleaned[:4]}-{cleaned[4:]}"
        
        return phone
    
    def format_id_number(self, id_number: str) -> str:
        """
        格式化身分證字號/統一編號
        
        Args:
            id_number: 原始身分證字號
            
        Returns:
            格式化後的身分證字號
        """
        if not id_number:
            return ""
        
        # 移除所有空白和特殊字元
        cleaned = re.sub(r'[^\w]', '', id_number)
        
        # 如果是統編（8位數字）
        if re.match(r'^\d{8}$', cleaned):
            return cleaned
        
        # 如果是身分證字號（1個字母+9位數字）
        if re.match(r'^[A-Z]\d{9}$', cleaned):
            return cleaned
        
        # 如果格式不正確，返回原始值
        return id_number
    
    def format_insurance_period(self, period: str) -> str:
        """
        直接回傳保險期間原始值，不做格式化
        """
        if not period or period.strip() == "":
            return "無填寫"
        return period.strip()
    
    def _format_single_date(self, date_str: str) -> str:
        """
        格式化單一日期（保留民國年格式）
        
        Args:
            date_str: 原始日期字串
            
        Returns:
            格式化後的日期
        """
        if not date_str or date_str.strip() == "":
            return "無填寫"
        
        date_str = date_str.strip()
        
        # 處理民國年格式 - 統一格式（包含時間）
        minguo_match = re.match(r'民國(\d{2,3})年(\d{1,2})月(\d{1,2})日', date_str)
        if minguo_match:
            year, month, day = minguo_match.groups()
            # 統一格式：民國114年5月20日
            return f"民國{year}年{month.zfill(2)}月{day.zfill(2)}日"
        
        # 處理包含時間的民國年格式
        minguo_time_match = re.match(r'民國(\d{2,3})年(\d{1,2})月(\d{1,2})日(.*)', date_str)
        if minguo_time_match:
            year, month, day, time_part = minguo_time_match.groups()
            # 統一格式：民國114年5月20日中午12時
            return f"民國{year}年{month.zfill(2)}月{day.zfill(2)}日{time_part}"
        
        # 處理西元年格式
        western_patterns = [
            r'(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})[日]?',
            r'(\d{4})(\d{2})(\d{2})'
        ]
        
        for pattern in western_patterns:
            match = re.search(pattern, date_str)
            if match:
                year, month, day = match.groups()
                try:
                    # 驗證日期有效性
                    datetime(int(year), int(month), int(day))
                    return f"{year}/{month.zfill(2)}/{day.zfill(2)}"
                except ValueError:
                    continue
        
        # 如果無法解析，返回原始值
        return date_str
    
    def parse_insurance_details(self, details_text: str) -> list:
        """
        解析承保內容文字為結構化資料
        
        Args:
            details_text: 承保內容文字
            
        Returns:
            結構化的承保內容列表
        """
        if not details_text:
            return []
        
        items = []
        lines = details_text.split('\n')
        current = {}
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # 檢查是否為新的保險項目
            if any(keyword in line for keyword in ['險種', '保險', '責任險', '車體險', '竊盜險']):
                if current:
                    items.append(current)
                current = {'name': line}
            elif '保險金額' in line:
                amt = line.split('保險金額')[-1].strip(' ：:').replace('萬', '').replace(':', '').replace('：', '').strip()
                if amt:
                    current['amount'] = amt + '萬' if '萬' not in amt else amt
            elif '自負額' in line:
                ded = line.split('自負額')[-1].strip(' ：:').replace('元', '').replace(':', '').replace('：', '').strip()
                if ded:
                    current['deductible'] = ded + '元' if '元' not in ded else ded
            elif '保費' in line or '簽單保費' in line:
                pfx = '簽單保費' if '簽單保費' in line else '保費'
                prem = line.split(pfx)[-1].strip(' ：:').replace('元', '').replace(':', '').replace('：', '').strip()
                if prem:
                    current['premium'] = prem + '元' if '元' not in prem else prem
        
        if current:
            items.append(current)
        return items

    def process_insurance_data(self, raw_data: Dict) -> Dict:
        """
        處理保單資料
        
        Args:
            raw_data: 原始OCR資料
            
        Returns:
            處理後的資料
        """
        processed_data = {}
        
        # 處理各個欄位
        field_processors = {
            'license_number': self.format_policy_number,
            'total_premium': self.format_amount,
            'birth_date': self.format_date,
            'policyholder_birth_date': self.format_date,
            'id_number': self.format_id_number,
            'policyholder_id': self.format_id_number,
            'compulsory_insurance_period': self.format_insurance_period,  # 強制險保期
            'optional_insurance_period': self.format_insurance_period     # 任意險保期
        }
        
        for field, value in raw_data.items():
            if field == 'coverage_items':
                processed_data[field] = value  # 已經是 list，直接存
            elif field in field_processors:
                processed_data[field] = field_processors[field](value)
            else:
                processed_data[field] = self.clean_text(value) if isinstance(value, str) and value else value
        
        # 將空值轉為「無填寫」
        for field, value in processed_data.items():
            if isinstance(value, str) and not value.strip():
                processed_data[field] = "無填寫"
        
        self.processed_data = processed_data
        return processed_data
    
    def get_data_summary(self, data: Dict) -> Dict:
        """
        取得資料摘要
        
        Args:
            data: 資料字典
            
        Returns:
            資料摘要
        """
        total_fields = len(data)
        filled_fields = sum(1 for v in data.values() if v)
        
        return {
            'total_fields': total_fields,
            'filled_fields': filled_fields,
            'empty_fields': total_fields - filled_fields,
            'completion_rate': f"{filled_fields / total_fields * 100:.1f}%" if total_fields > 0 else "0%",
            'field_details': {
                field: '已填寫' if value else '未填寫'
                for field, value in data.items()
            }
        } 