"""
Pattern definitions for text extraction
الگوهای regex برای استخراج اطلاعات از متن
"""

import re
from typing import List, Tuple


class Patterns:
    """کلاس حاوی الگوهای استخراج اطلاعات"""
    
    # الگوهای استخراج شماره صورت‌وضعیت
    INVOICE_PATTERNS = [
        # فارسی
        r'صورت وضعیت\s*[:؛]?\s*(\d+)',
        r'شماره\s*صورت وضعیت\s*[:؛]?\s*(\d+)',
        r'صورت وضعیت شماره\s*(\d+)',
        r'ش.\s*و.\s*(\d+)',
        r'شماره\s*[:؛]?\s*(\d+)',
        
        # انگلیسی
        r'Invoice\s*#?\s*(\d+)',
        r'INV\s*(\d+)',
        r'Invoice\s*Number\s*[:]?\s*(\d+)',
        r'Bill\s*#?\s*(\d+)',
    ]
    
    # الگوهای استخراج اطلاعات ارز
    CURRENCY_PATTERNS = [
        # فارسی - با نرخ
        r'(\d[\d,\.]*)\s*(یورو|دلار|یورو|ريال|ریال)\s*(?:نرخ|با نرخ|في|@)\s*(\d[\d,\.]*)',
        r'(\d[\d,\.]*)\s*(یورو|دلار|یورو|ريال|ریال)\s*(?:نرخ|با نرخ|في|@)\s*(\d[\d,\.]*)',
        
        # فارسی - بدون نرخ
        r'(\d[\d,\.]*)\s*(یورو|دلار|یورو|ريال|ریال)',
        
        # انگلیسی - با نرخ
        r'(\d[\d,\.]*)\s*(EUR|USD|IRR|Euro|Dollar|Rial)\s*(?:rate|@|at)\s*(\d[\d,\.]*)',
        
        # انگلیسی - بدون نرخ
        r'(\d[\d,\.]*)\s*(EUR|USD|IRR|Euro|Dollar|Rial)',
    ]
    
    # الگوهای شناسایی شرکت
    COMPANY_PATTERNS = [
        'ایران',
        'ایرایتک', 
        'پترو ساحل',
        'فرآب',
        'ناردیس',
        'خارک',
        'Iran',
        'Iratec',
        'Petro Sahel',
        'Farab',
        'Nardis',
        'Khark',
    ]
    
    # الگوهای شناسایی نوع سند
    DOCUMENT_TYPE_PATTERNS = [
        (r'تسعیر|نرخ', 'تسعیر ارز'),
        (r'صورت وضعیت|صورتوضعیت|Invoice', 'صورت وضعیت'),
        (r'چک|Check', 'چک'),
        (r'انتقال|مانده|Transfer', 'انتقال'),
        (r'سند متفرقه|Misc', 'سند متفرقه'),
    ]
    
    @classmethod
    def extract_invoice_number(cls, text: str) -> Tuple[Optional[str], float]:
        """استخراج شماره صورت‌وضعیت از متن"""
        if not text:
            return None, 0.0
        
        for pattern in cls.INVOICE_PATTERNS:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                confidence = 1.0 if 'صورت وضعیت' in pattern or 'Invoice' in pattern else 0.8
                return match.group(1), confidence
        
        return None, 0.0
    
    @classmethod
    def extract_currency_info(cls, text: str) -> Tuple[Optional[dict], float]:
        """استخراج اطلاعات ارز از متن"""
        if not text:
            return None, 0.0
        
        for pattern in cls.CURRENCY_PATTERNS:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                groups = match.groups()
                amount_str = groups[0].replace(',', '') if groups[0] else None
                currency = groups[1]
                rate = groups[2] if len(groups) > 2 else None
                
                try:
                    amount = float(amount_str) if amount_str else None
                    if rate:
                        rate = float(rate.replace(',', ''))
                except (ValueError, TypeError):
                    continue
                
                # محاسبه اطمینان
                confidence = 1.0 if rate else 0.8
                
                return {
                    'amount': amount,
                    'currency': currency,
                    'rate': rate,
                    'original_text': match.group(0)
                }, confidence
        
        return None, 0.0
    
    @classmethod
    def extract_company(cls, text: str) -> Tuple[Optional[str], float]:
        """استخراج نام شرکت از متن"""
        if not text:
            return None, 0.0
        
        for company in cls.COMPANY_PATTERNS:
            if company in text:
                confidence = 0.9 if company in ['ایران', 'Iran'] else 0.7
                return company, confidence
        
        return None, 0.0
    
    @classmethod
    def detect_document_type(cls, text: str) -> Tuple[Optional[str], float]:
        """تشخیص نوع سند از متن"""
        if not text:
            return 'سند متفرقه', 0.5
        
        text_lower = text.lower()
        
        for pattern, doc_type in cls.DOCUMENT_TYPE_PATTERNS:
            if re.search(pattern, text_lower, re.IGNORECASE):
                confidence = 0.9 if pattern in ['صورت وضعیت', 'Invoice'] else 0.7
                return doc_type, confidence
        
        return 'سند متفرقه', 0.5
