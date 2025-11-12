"""
File handling utilities for Smart Extractor
ابزارهای مدیریت فایل برای سیستم استخراج هوشمند
"""

import os
import re
from pathlib import Path
from typing import Optional


class FileHandler:
    """کلاس مدیریت فایل و نام‌گذاری"""
    
    @staticmethod
    def generate_output_filename(input_path: str, suffix: str = "_extracted") -> str:
        """تولید نام فایل خروجی با اندیس"""
        input_path = Path(input_path)
        
        # استخراج نام فایل بدون پسوند
        filename = input_path.stem
        
        # حذف کاراکترهای غیرمجاز
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        
        # اضافه کردن اندیس در صورت وجود فایل تکراری
        counter = 1
        output_filename = f"{filename}{suffix}{input_path.suffix}"
        
        while os.path.exists(output_filename):
            output_filename = f"{filename}{suffix}_{counter}{input_path.suffix}"
            counter += 1
        
        return output_filename
    
    @staticmethod
    def validate_file_path(file_path: str) -> bool:
        """اعتبارسنجی مسیر فایل"""
        try:
            path = Path(file_path)
            return path.exists() and path.is_file()
        except Exception:
            return False
    
    @staticmethod
    def get_file_info(file_path: str) -> dict:
        """دریافت اطلاعات فایل"""
        path = Path(file_path)
        return {
            'filename': path.name,
            'stem': path.stem,
            'suffix': path.suffix,
            'parent': str(path.parent),
            'size_bytes': path.stat().st_size if path.exists() else 0,
            'exists': path.exists()
        }
    
    @staticmethod
    def create_backup_filename(original_path: str) -> str:
        """ایجاد نام فایل پشتیبان"""
        path = Path(original_path)
        return f"{path.stem}_backup{path.suffix}"
    
    @staticmethod
    def sanitize_filename(filename: str) -> str:
        """پاکسازی نام فایل از کاراکترهای غیرمجاز"""
        # حذف کاراکترهای غیرمجاز در ویندوز
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
        
        # حذف فاصله‌های اضافی
        sanitized = re.sub(r'\s+', ' ', sanitized).strip()
        
        # محدود کردن طول نام فایل
        if len(sanitized) > 255:
            sanitized = sanitized[:255]
        
        return sanitized
