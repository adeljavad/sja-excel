#!/usr/bin/env python3
"""
Standalone script for Smart Extractor
اسکریپت مستقل برای سیستم استخراج هوشمند
"""

import argparse
import sys
import os
from pathlib import Path

# اضافه کردن مسیر ماژول به sys.path
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

# استفاده از import مطلق
try:
    from smart_extractor.processors.excel_processor import ExcelProcessor
    from smart_extractor.utils.file_handler import FileHandler
except ImportError:
    # اگر import مطلق کار نکرد، از import نسبی استفاده کنیم
    from processors.excel_processor import ExcelProcessor
    from utils.file_handler import FileHandler


def main():
    """تابع اصلی اسکریپت مستقل"""
    parser = argparse.ArgumentParser(
        description='سیستم استخراج هوشمند اطلاعات از فایل‌های اکسل',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
نمونه استفاده:
  python standalone.py data.xlsx
  python standalone.py data.xlsx -o extracted_data
  python standalone.py data.xlsx --suffix "_processed"
        """
    )
    
    parser.add_argument('input_file', help='مسیر فایل اکسل ورودی')
    parser.add_argument('-o', '--output', help='پسوند نام فایل خروجی (اختیاری)', default='_extracted')
    parser.add_argument('--suffix', help='نام مستعار برای پسوند خروجی', default=None)
    
    args = parser.parse_args()
    
    # اعتبارسنجی فایل ورودی
    if not os.path.exists(args.input_file):
        print(f"❌ فایل {args.input_file} یافت نشد")
        return 1
    
    # استفاده از پسوند مشخص شده یا مقدار پیش‌فرض
    suffix = args.suffix if args.suffix else args.output
    
    try:
        print("🚀 سیستم استخراج هوشمند - نسخه مستقل")
        print("=" * 50)
        
        # پردازش فایل
        processor = ExcelProcessor()
        output_path = processor.process_excel_file(args.input_file, suffix)
        
        # نمایش خلاصه نتایج
        df = processor.read_excel_file(output_path)
        summary = processor.get_processing_summary(df)
        
        print(f"\n📊 خلاصه نتایج:")
        print(f"   کل رکوردها: {summary['total_records']}")
        print(f"   شماره‌های وضعیت استخراج شده: {summary['invoices_extracted']}")
        print(f"   اطلاعات ارز استخراج شده: {summary['currency_info_extracted']}")
        print(f"   شرکت‌های شناسایی شده: {summary['companies_identified']}")
        print(f"   میانگین اطمینان: {summary['average_confidence']}")
        print(f"   نرخ موفقیت: {summary['success_rate']}%")
        
        print(f"\n🎉 پردازش با موفقیت تکمیل شد!")
        print(f"📁 فایل خروجی: {output_path}")
        
        return 0
        
    except Exception as e:
        print(f"❌ خطا در پردازش فایل: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
