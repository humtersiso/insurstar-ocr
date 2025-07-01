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
        格式化身分證號碼
        
        Args:
            id_number: 原始身分證號碼
            
        Returns:
            格式化後的身分證號碼
        """
        if not id_number:
            return ""
        
        # 移除所有非字母數字字符
        cleaned = re.sub(r'[^A-Za-z0-9]', '', id_number)
        
        # 轉換為大寫
        cleaned = cleaned.upper()
        
        # 檢查格式
        if re.match(r'^[A-Z][0-9]{9}$', cleaned):
            return cleaned
        
        return id_number
    
    def parse_insurance_details(self, details_text: str) -> list:
        """
        將保險詳情字串解析為陣列，每筆含 name、amount、deductible、premium
        """
        items = []
        if not details_text:
            return items
        # 以換行分段
        lines = [line.strip() for line in details_text.split('\n') if line.strip()]
        current = None
        for line in lines:
            # 偵測新保險項目（有保險/責任險/附加條款等關鍵字）
            if any(kw in line for kw in ['保險', '責任險', '附加條款']):
                if current:
                    items.append(current)
                current = {'name': line, 'amount': '', 'deductible': '', 'premium': ''}
            elif current:
                if '保險金額' in line:
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
            'policyholder_id': self.format_id_number
        }
        
        for field, value in raw_data.items():
            if field == 'coverage_items':
                processed_data[field] = value  # 已經是 list，直接存
            elif field in field_processors:
                processed_data[field] = field_processors[field](value)
            else:
                processed_data[field] = self.clean_text(value) if isinstance(value, str) and value else value
        
        self.processed_data = processed_data
        return processed_data
    
    def validate_processed_data(self, data: Dict) -> Dict:
        """
        驗證處理後的資料
        
        Args:
            data: 處理後的資料
            
        Returns:
            驗證結果
        """
        validation_result = {
            'is_valid': True,
            'errors': [],
            'warnings': [],
            'missing_fields': [],
            'suggestions': []
        }
        
        # 驗證牌照號碼
        if data.get('license_number'):
            license_num = data['license_number'].replace('-', '')
            if not re.match(r'^[A-Z0-9]{6,10}$', license_num):
                validation_result['warnings'].append('牌照號碼格式可能不正確')
        
        # 驗證身分證號碼/統編
        if data.get('id_number'):
            id_num = data['id_number']
            # 檢查是否為身分證格式或統編格式
            if not (re.match(r'^[A-Z][0-9]{9}$', id_num) or re.match(r'^[0-9]{8}$', id_num)):
                validation_result['warnings'].append('身分證字號/統編格式可能不正確')
        
        if data.get('policyholder_id'):
            policyholder_id = data['policyholder_id']
            if not (re.match(r'^[A-Z][0-9]{9}$', policyholder_id) or re.match(r'^[0-9]{8}$', policyholder_id)):
                validation_result['warnings'].append('要保人身分證字號/統編格式可能不正確')
        
        # 驗證日期格式
        date_fields = ['birth_date', 'policyholder_birth_date']
        for date_field in date_fields:
            date_value = data.get(date_field)
            if date_value and date_value not in ['無填寫', '無法辨識', '']:
                # 檢查民國年格式或西元年格式
                if not (re.match(r'民國\d{2,3}年\d{1,2}月\d{1,2}日', date_value) or 
                       re.match(r'\d{4}[/-]\d{1,2}[/-]\d{1,2}', date_value)):
                    validation_result['warnings'].append(f'{date_field}日期格式可能不正確')
        
        # 檢查重要欄位是否存在
        important_fields = {
            'insurance_company': '產險公司',
            'insured_person': '被保險人',
            'license_number': '牌照號碼',
            'coverage_items': '承保內容',
            'total_premium': '總保險費'
        }
        
        for field, field_name in important_fields.items():
            value = data.get(field)
            if field == 'coverage_items':
                # 只有 coverage_items 為空或不是有效 list 才算未辨識
                if not value or not isinstance(value, list) or len(value) == 0:
                    validation_result['missing_fields'].append(field_name)
            else:
                if not value or value in ['無填寫', '無法辨識', '']:
                    validation_result['missing_fields'].append(field_name)
        
        # 根據缺失欄位數量決定驗證狀態
        total_important_fields = len(important_fields)
        missing_count = len(validation_result['missing_fields'])
        filled_count = total_important_fields - missing_count
        
        if missing_count == 0:
            validation_result['is_valid'] = True
            validation_result['suggestions'].append('所有重要欄位都已成功辨識')
        elif missing_count <= 2:
            validation_result['is_valid'] = True
            validation_result['suggestions'].append(f'已成功辨識 {filled_count}/{total_important_fields} 個重要欄位')
        else:
            validation_result['is_valid'] = False
            validation_result['suggestions'].append('建議重新上傳更清晰的檔案')
        
        return validation_result
    
    def export_to_json(self, data: Dict, file_path: str) -> str:
        """
        匯出資料為JSON格式
        
        Args:
            data: 要匯出的資料
            file_path: 檔案路徑
            
        Returns:
            匯出的檔案路徑
        """
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return file_path
        except Exception as e:
            print(f"JSON匯出錯誤: {str(e)}")
            return ""
    
    def import_from_json(self, file_path: str) -> Dict:
        """
        從JSON檔案匯入資料
        
        Args:
            file_path: JSON檔案路徑
            
        Returns:
            匯入的資料字典
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data
        except Exception as e:
            print(f"JSON匯入錯誤: {str(e)}")
            return {}
    
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