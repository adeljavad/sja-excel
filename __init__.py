"""
Smart Extractor - ماژول استخراج هوشمند اطلاعات از شرح اسناد
یک سیستم ماژولار برای استخراج اطلاعات از متن‌های فارسی و انگلیسی

ویژگی‌ها:
- استخراج شماره صورت‌وضعیت
- استخراج اطلاعات ارز (مبلغ، نرخ، نوع)
- پشتیبانی از فرمت‌های فارسی و انگلیسی
- طراحی ماژولار برای استفاده در اودوو، جنگو و پایتون
"""

__version__ = "1.0.0"
__author__ = "Smart Extractor Team"

from .core.extractors import SmartExtractor
from .core.models import ExtractionResult
from .processors.excel_processor import ExcelProcessor
from .utils.file_handler import FileHandler

__all__ = [
    'SmartExtractor',
    'ExtractionResult', 
    'ExcelProcessor',
    'FileHandler'
]
