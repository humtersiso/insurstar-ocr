import re
from typing import Dict
from datetime import datetime

class DataProcessor:
    """資料處理器"""
    def __init__(self):
        self.processed_data = {}

    def clean_text(self, text: str) -> str:
        if not text:
            return ""
        text = re.sub(r'\s+', ' ', text.strip())
        text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\s\-\.\,\:\/\(\)]', '', text)
        return text

    def format_policy_number(self, policy_number: str) -> str:
        if not policy_number:
            return ""
        cleaned = re.sub(r'[^A-Za-z0-9]', '', policy_number)
        cleaned = cleaned.upper()
        if 8 <= len(cleaned) <= 20:
            return cleaned
        return policy_number

    def format_amount(self, amount: str) -> str:
        if not amount:
            return ""
        # 先移除所有空白字符
        amount = re.sub(r'\s+', '', amount)
        # 移除所有非數字字符，包括逗號
        cleaned = re.sub(r'[^\d]', '', amount)
        if cleaned:
            try:
                num = int(cleaned)
                return f"{num:,}"
            except:
                return amount
        return amount

    def format_date(self, date_str: str) -> str:
        if not date_str:
            return ""
        date_patterns = [
            r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})',
            r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})',
            r'(\d{4})(\d{2})(\d{2})'
        ]
        for pattern in date_patterns:
            match = re.search(pattern, date_str)
            if match:
                if len(match.group(1)) == 4:
                    year, month, day = match.groups()
                else:
                    month, day, year = match.groups()
                try:
                    datetime(int(year), int(month), int(day))
                    return f"{year}/{month.zfill(2)}/{day.zfill(2)}"
                except ValueError:
                    continue
        return date_str

    def format_id_number(self, id_number: str) -> str:
        if not id_number:
            return ""
        cleaned = re.sub(r'[^\w]', '', id_number)
        if re.match(r'^\d{8}$', cleaned):
            return cleaned
        if re.match(r'^[A-Z]\d{9}$', cleaned):
            return cleaned
        return id_number

    def format_insurance_period(self, period: str) -> str:
        if not period or period.strip() == "":
            return "無填寫"
        return period.strip()

    def process_insurance_data(self, raw_data: Dict) -> Dict:
        processed_data = {}
        field_processors = {
            'license_number': self.format_policy_number,
            'birth_date': self.format_date,
            'policyholder_birth_date': self.format_date,
            'id_number': self.format_id_number,
            'policyholder_id': self.format_id_number,
            'compulsory_insurance_period': self.format_insurance_period,
            'optional_insurance_period': self.format_insurance_period
        }
        for field, value in raw_data.items():
            if field == 'coverage_items':
                processed_data[field] = value
            elif field in field_processors:
                processed_data[field] = field_processors[field](value)
            elif field == 'total_premium':
                # total_premium 直接使用原始值，不進行任何處理
                processed_data[field] = value
            else:
                processed_data[field] = self.clean_text(value) if isinstance(value, str) and value else value
        for field, value in processed_data.items():
            if isinstance(value, str) and not value.strip():
                processed_data[field] = "無填寫"
        self.processed_data = processed_data
        return processed_data

    def get_data_summary(self, data: Dict) -> Dict:
        total_fields = len(data)
        filled_fields = sum(1 for v in data.values() if v)
        rate = f"{filled_fields / total_fields * 100:.1f}%" if total_fields > 0 else "0%"
        return {
            'total_fields': total_fields,
            'filled_fields': filled_fields,
            'empty_fields': total_fields - filled_fields,
            'extraction_rate': rate,
            'completion_rate': rate,
            'data': data
        } 