#!/usr/bin/env python3
"""
Test script for Smart Extractor
اسکریپت تست برای سیستم استخراج هوشمند
"""

import pandas as pd
import os
from pathlib import Path

# اضافه کردن مسیر ماژول به sys.path
import sys
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from core.extractors import SmartExtractor
from processors.excel_processor import ExcelProcessor


def test_extraction():
    """تست استخراج اطلاعات از متن"""
    print("🧪 تست استخراج اطلاعات از متن")
    print("=" * 40)
    
    extractor = SmartExtractor()
    
    # نمونه متن‌های تست
    test_texts = [
        "صورت وضعیت شماره 1234 - پرداخت از شرکت ایران - مبلغ 1000 یورو با نرخ 50000",
        "چک شماره 5678 - مبلغ 5000000 ریال",
        "تسعیر ارز 2000 دلار با نرخ 300000",
        "انتقال از حساب جاری - شرکت پترو ساحل",
        "سند متفرقه - پرداخت هزینه‌های اداری",
        "Invoice #999 - Payment to Iratec - 1500 EUR at rate 55000",
    ]
    
    for i, text in enumerate(test_texts, 1):
        print(f"\n📝 متن {i}: {text}")
        result = extractor.extract_from_text(text)
        
        print(f"   📄 شماره وضعیت: {result.invoice_number}")
        if result.currency_info:
            print(f"   💰 ارز: {result.currency_info.amount} {result.currency_info.currency}")
            print(f"   📊 نرخ: {result.currency_info.rate}")
        print(f"   🏢 شرکت: {result.company_name}")
        print(f"   📋 نوع سند: {result.document_type}")
        print(f"   ✅ اطمینان: {result.confidence:.2f}")


def create_sample_excel():
    """ایجاد فایل اکسل نمونه برای تست"""
    print("\n\n📁 ایجاد فایل اکسل نمونه")
    print("=" * 40)
    
    # داده‌های نمونه
    sample_data = [
        {
            'شرح': 'صورت وضعیت شماره 1001 - پرداخت از شرکت ایران - مبلغ 1,000,000 ریال',
            'مبلغ': 1000000,
            'تاریخ': '1402/01/15',
            'شماره سند': 'INV1001'
        },
        {
            'شرح': 'چک شماره 1234 - مبلغ 500,000 ریال',
            'مبلغ': 500000,
            'تاریخ': '1402/01/20',
            'شماره سند': 'CHK1234'
        },
        {
            'شرح': 'صورت وضعیت شماره 1002 - تسعیر ارز 2000 یورو با نرخ 50000',
            'مبلغ': 100000000,
            'تاریخ': '1402/02/01',
            'شماره سند': 'INV1002'
        },
        {
            'شرح': 'انتقال از حساب جاری - شرکت پترو ساحل - مبلغ 750,000 ریال',
            'مبلغ': 750000,
            'تاریخ': '1402/02/05',
            'شماره سند': 'TRF001'
        },
        {
            'شرح': 'Invoice #1003 - Payment to Farab - 1500 USD at rate 300000',
            'مبلغ': 450000000,
            'تاریخ': '2023/04/10',
            'شماره سند': 'INV1003'
        }
    ]
    
    # ایجاد DataFrame
    df = pd.DataFrame(sample_data)
    
    # ذخیره فایل نمونه
    sample_file = 'sample_data.xlsx'
    df.to_excel(sample_file, index=False)
    
    print(f"✅ فایل نمونه ایجاد شد: {sample_file}")
    print(f"📊 تعداد رکوردها: {len(df)}")
    
    return sample_file


def test_excel_processing():
    """تست پردازش فایل اکسل"""
    print("\n\n🔧 تست پردازش فایل اکسل")
    print("=" * 40)
    
    # ایجاد فایل نمونه
    sample_file = create_sample_excel()
    
    try:
        # پردازش فایل
        processor = ExcelProcessor()
        output_file = processor.process_excel_file(sample_file)
        
        # خواندن فایل خروجی و نمایش نتایج
        df_output = pd.read_excel(output_file)
        
        print(f"\n📈 نتایج پردازش:")
        print(f"   فایل خروجی: {output_file}")
        print(f"   تعداد رکوردها: {len(df_output)}")
        
        # نمایش خلاصه
        summary = processor.get_processing_summary(df_output)
        print(f"\n📊 خلاصه استخراج:")
        print(f"   شماره‌های وضعیت استخراج شده: {summary['invoices_extracted']}")
        print(f"   اطلاعات ارز استخراج شده: {summary['currency_info_extracted']}")
        print(f"   شرکت‌های شناسایی شده: {summary['companies_identified']}")
        print(f"   میانگین اطمینان: {summary['average_confidence']}")
        print(f"   نرخ موفقیت: {summary['success_rate']}%")
        
        # نمایش نمونه داده‌های استخراج شده
        print(f"\n📋 نمونه داده‌های استخراج شده:")
        columns_to_show = ['شرح', 'شماره_وضعیت', 'مبلغ_ارزی', 'نوع_ارز', 'نرخ_ارز', 'نام_شرکت', 'اطمینان_استخراج']
        available_columns = [col for col in columns_to_show if col in df_output.columns]
        print(df_output[available_columns].head(3).to_string())
        
    except Exception as e:
        print(f"❌ خطا در پردازش فایل: {str(e)}")
    
    finally:
        # حذف فایل‌های موقت
        if os.path.exists(sample_file):
            os.remove(sample_file)
            print(f"\n🗑️ فایل نمونه حذف شد: {sample_file}")


def test_standalone_script():
    """تست اسکریپت مستقل"""
    print("\n\n🚀 تست اسکریپت مستقل")
    print("=" * 40)
    
    # ایجاد فایل نمونه
    sample_file = create_sample_excel()
    
    try:
        # اجرای اسکریپت مستقل
        print("💻 اجرای اسکریپت مستقل...")
        os.system(f'python standalone.py {sample_file}')
        
    except Exception as e:
        print(f"❌ خطا در اجرای اسکریپت: {str(e)}")
    
    finally:
        # حذف فایل نمونه
        if os.path.exists(sample_file):
            os.remove(sample_file)
            print(f"\n🗑️ فایل نمونه حذف شد: {sample_file}")


if __name__ == "__main__":
    print("🧪 سیستم استخراج هوشمند - تست جامع")
    print("=" * 50)
    
    # اجرای تست‌ها
    test_extraction()
    test_excel_processing()
    test_standalone_script()
    
    print("\n🎉 تمام تست‌ها با موفقیت اجرا شد!")
