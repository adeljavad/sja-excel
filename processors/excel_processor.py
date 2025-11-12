"""
Excel processor for Smart Extractor
پردازش‌گر فایل‌های اکسل برای سیستم استخراج هوشمند
"""

import pandas as pd
from pathlib import Path
from typing import List, Dict, Any, Optional

# استفاده از import مطلق
try:
    from smart_extractor.core.extractors import SmartExtractor
    from smart_extractor.utils.file_handler import FileHandler
except ImportError:
    # اگر import مطلق کار نکرد، از import نسبی استفاده کنیم
    from ..core.extractors import SmartExtractor
    from ..utils.file_handler import FileHandler


class ExcelProcessor:
    """کلاس پردازش فایل‌های اکسل"""
    
    def __init__(self):
        self.extractor = SmartExtractor()
        self.file_handler = FileHandler()
        
        # نگاشت ستون‌های فارسی و انگلیسی
        self.column_mapping = {
            'شرح': 'description',
            'مبلغ': 'amount',
            'تاریخ': 'date',
            'شماره سند': 'document_number',
            'شماره حساب': 'account_number',
            'نام حساب': 'account_name',
            
            # انگلیسی
            'description': 'description',
            'amount': 'amount',
            'date': 'date',
            'document_number': 'document_number',
            'account_number': 'account_number',
            'account_name': 'account_name',
            
            # متغیرهای رایج
            'شرح عملیات': 'description',
            'شرح تراکنش': 'description',
            'مبلغ تراکنش': 'amount',
            'مبلغ عملیات': 'amount',
            'تاریخ تراکنش': 'date',
            'تاریخ عملیات': 'date',
        }
    
    def read_excel_file(self, file_path: str) -> pd.DataFrame:
        """خواندن فایل اکسل و ترکیب تمام شیت‌ها"""
        try:
            print(f"📖 خواندن فایل: {file_path}")
            
            # خواندن تمام شیت‌ها
            excel_file = pd.ExcelFile(file_path)
            all_sheets_data = []
            
            for sheet_name in excel_file.sheet_names:
                print(f"   📄 پردازش شیت: {sheet_name}")
                df_sheet = pd.read_excel(file_path, sheet_name=sheet_name)
                df_sheet['sheet_name'] = sheet_name
                all_sheets_data.append(df_sheet)
            
            # ترکیب تمام شیت‌ها
            combined_df = pd.concat(all_sheets_data, ignore_index=True)
            print(f"   ✅ کل رکوردها: {len(combined_df)}")
            
            # استانداردسازی ستون‌ها
            combined_df = self._standardize_columns(combined_df)
            
            return combined_df
            
        except Exception as e:
            print(f"❌ خطا در خواندن فایل اکسل: {str(e)}")
            raise
    
    def _standardize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """استانداردسازی نام ستون‌ها"""
        df.columns = [self.column_mapping.get(str(col).strip(), str(col).strip()) for col in df.columns]
        return df
    
    def extract_and_enrich(self, df: pd.DataFrame, description_column: str = 'description') -> pd.DataFrame:
        """استخراج اطلاعات از ستون شرح و افزودن ستون‌های جدید"""
        if description_column not in df.columns:
            raise ValueError(f"ستون '{description_column}' در داده‌ها یافت نشد")
        
        print(f"🔍 استخراج اطلاعات از ستون '{description_column}'...")
        
        # استخراج اطلاعات از شرح‌ها
        descriptions = df[description_column].astype(str).tolist()
        extracted_data = self.extractor.extract_from_description_column(descriptions)
        
        # افزودن ستون‌های جدید به DataFrame
        enriched_df = df.copy()
        
        # ستون‌های استخراج شده
        enriched_df['شماره_وضعیت'] = [data['invoice_number'] for data in extracted_data]
        enriched_df['مبلغ_ارزی'] = [data['currency_amount'] for data in extracted_data]
        enriched_df['نوع_ارز'] = [data['currency_type'] for data in extracted_data]
        enriched_df['نرخ_ارز'] = [data['exchange_rate'] for data in extracted_data]
        enriched_df['نام_شرکت'] = [data['company_name'] for data in extracted_data]
        enriched_df['نوع_سند'] = [data['document_type'] for data in extracted_data]
        enriched_df['اطمینان_استخراج'] = [data['extraction_confidence'] for data in extracted_data]
        
        print(f"✅ {len(enriched_df)} رکورد پردازش شد")
        
        return enriched_df
    
    def process_excel_file(self, input_path: str, output_suffix: str = "_extracted") -> str:
        """پردازش کامل فایل اکسل و ذخیره فایل جدید"""
        # اعتبارسنجی فایل
        if not self.file_handler.validate_file_path(input_path):
            raise ValueError(f"فایل {input_path} یافت نشد یا معتبر نیست")
        
        # خواندن فایل
        df = self.read_excel_file(input_path)
        
        # استخراج و غنی‌سازی داده‌ها
        enriched_df = self.extract_and_enrich(df)
        
        # تولید نام فایل خروجی
        output_path = self.file_handler.generate_output_filename(input_path, output_suffix)
        
        # ذخیره فایل جدید
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                enriched_df.to_excel(writer, sheet_name='داده‌های_استخراج_شده', index=False)
            
            print(f"💾 فایل خروجی ذخیره شد: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"❌ خطا در ذخیره فایل: {str(e)}")
            raise
    
    def get_processing_summary(self, df: pd.DataFrame) -> Dict[str, Any]:
        """خلاصه نتایج پردازش"""
        total_records = len(df)
        
        # آمار استخراج
        invoice_count = df['شماره_وضعیت'].notna().sum()
        currency_count = df['مبلغ_ارزی'].notna().sum()
        company_count = df['نام_شرکت'].notna().sum()
        
        avg_confidence = df['اطمینان_استخراج'].mean() if 'اطمینان_استخراج' in df.columns else 0
        
        return {
            'total_records': total_records,
            'invoices_extracted': invoice_count,
            'currency_info_extracted': currency_count,
            'companies_identified': company_count,
            'average_confidence': round(avg_confidence, 2),
            'success_rate': round((invoice_count + currency_count) / (total_records * 2) * 100, 1) if total_records > 0 else 0
        }
