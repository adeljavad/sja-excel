#!/usr/bin/env python3
"""
Simple standalone script for Smart Extractor
اسکریپت ساده مستقل برای سیستم استخراج هوشمند
"""

import argparse
import sys
import os
import pandas as pd
import re
from pathlib import Path
from pandas import ExcelFile, read_excel, concat, isna


class ExcelSheetCombiner:
    """کلاس ترکیب کننده شیت‌های اکسل"""
    
    def __init__(self):
        self.sheet_name_column = "نام_شیت"
        # ستون‌های مهم که باید همیشه حفظ شوند
        self.important_columns = [
            'تاریخ سند', 'شرح سند', 'بدهکار', 'بستانکار', 'بدهكار - ريالي', 
            'بستانكار - ريالي', 'بدهكار - ارزي', 'بستانكار - ارزي', 'نوع ارز',
            'نرخ تبديل ارز', 'شماره سند', 'رديف', 'كد حساب', 'شرح رديف سند',
            'صادر کننده سند', 'پروژه', 'پيمانكار/كارفرما'
        ]
        # ستون‌های کم اهمیت که باید حذف شوند اگر خالی هستند
        self.low_importance_columns = [
            'تسهيلات', 'مشخصات پرسنلي', 'تنخواه دار', 'انبار', 'اعتبارات اسنادی',
            'شماره مجوز پرداخت', 'شماره مجوز دريافت', 'صندوقدار', 'شماره چك يا رسيد',
            'تاريخ چك', 'سهامدار', 'حساب بانكي', 'شماره نامه اعلاميه', 
            'تاييد شده در سامانه مودیان', 'چاپ سند', 'شماره عطف', 'شماره پيگيري',
            'تاريخ پيگيري', 'نوع فعالیت/مرکز هزینه', 'محل ایجاد کننده هزینه',
            'قرارداد فروش', 'شماره صورتحساب فروش', 'قرارداد خريد',
            'مانده بستانکار', 'مانده بدهکار'
        ]
    
    def analyze_column_completeness(self, df, threshold=0.1):
        """تحلیل کامل بودن ستون‌ها و حذف ستون‌های خالی و تکراری"""
        total_rows = len(df)
        if total_rows == 0:
            return df
        
        columns_to_keep = []
        columns_to_remove = []
        
        # گروه‌بندی ستون‌های تکراری
        column_groups = {
            'بستانکار': ['بستانکار', 'بستانكار - ريالي', 'معادل ریالی بستانکار'],
            'بدهکار': ['بدهکار', 'بدهكار - ريالي', 'معادل ریالی بدهکار'],
            'مانده بستانکار': ['مانده بستانکار', 'معادل ریالی مانده بستانکار'],
            'مانده بدهکار': ['مانده بدهکار', 'معادل ریالی مانده بدهکار'],
            'بستانكار - ارزي': ['بستانكار - ارزي', 'بستانکار ارزی'],
            'بدهكار - ارزي': ['بدهكار - ارزي', 'بدهکار ارزی']
        }
        
        # ستون‌های انتخاب شده از هر گروه
        selected_columns = {}
        
        for column in df.columns:
            if column == self.sheet_name_column:
                columns_to_keep.append(column)
                continue
            
            # بررسی اینکه آیا ستون در گروه تکراری قرار دارد
            column_in_group = False
            for group_name, group_columns in column_groups.items():
                if column in group_columns:
                    column_in_group = True
                    # اگر هنوز ستونی از این گروه انتخاب نشده، این ستون را انتخاب کن
                    if group_name not in selected_columns:
                        selected_columns[group_name] = column
                        columns_to_keep.append(column)
                    else:
                        # ستون تکراری - حذف شود
                        columns_to_remove.append(column)
                    break
            
            if column_in_group:
                continue
                
            # ستون‌های مهم همیشه حفظ شوند
            if any(important_col in str(column) for important_col in self.important_columns):
                columns_to_keep.append(column)
                continue
            
            # محاسبه درصد داده‌های غیرخالی
            non_empty_count = df[column].notna().sum()
            completeness_ratio = non_empty_count / total_rows
            
            # اگر ستون کم اهمیت است و کمتر از آستانه داده دارد، حذف شود
            if any(low_col in str(column) for low_col in self.low_importance_columns):
                if completeness_ratio < threshold:
                    columns_to_remove.append(column)
                else:
                    columns_to_keep.append(column)
            else:
                # ستون‌های دیگر اگر کمتر از آستانه داده دارند حذف شوند
                if completeness_ratio < threshold:
                    columns_to_remove.append(column)
                else:
                    columns_to_keep.append(column)
        
        # حذف ستون‌های تکراری
        columns_to_keep = list(set(columns_to_keep))
        
        # حذف ستون‌های مشخص شده برای تحلیل حسابرسی
        audit_columns_to_remove = ['مانده بستانکار', 'مانده بدهکار', 'تاييد شده در سامانه مودیان', 'چاپ سند']
        columns_to_keep = [col for col in columns_to_keep if col not in audit_columns_to_remove]
        columns_to_remove.extend(audit_columns_to_remove)
        
        print(f"   📊 تحلیل ستون‌ها: {len(columns_to_keep)} ستون نگهداری شد، {len(columns_to_remove)} ستون حذف شد")
        if columns_to_remove:
            print(f"   🗑️ ستون‌های حذف شده: {', '.join(columns_to_remove[:5])}{'...' if len(columns_to_remove) > 5 else ''}")
        
        # نمایش ستون‌های انتخاب شده از گروه‌های تکراری
        if selected_columns:
            print(f"   🔄 ستون‌های یکسان ادغام شدند: {selected_columns}")
        
        return df[columns_to_keep]
    
    def combine_sheets_simple(self, input_path, output_suffix="_combined"):
        """ترکیب ساده و قابل اعتماد تمام شیت‌های اکسل"""
        print(f"🚀 شروع ترکیب شیت‌های فایل: {input_path}")
        
        try:
            # خواندن تمام شیت‌های فایل اکسل
            excel_file = ExcelFile(input_path)
            sheet_names = excel_file.sheet_names
            print(f"📋 شیت‌های شناسایی شده: {sheet_names}")
            
            if len(sheet_names) == 0:
                print("❌ هیچ شیتی در فایل یافت نشد")
                return None
            
            # لیست برای ذخیره تمام داده‌ها
            all_data = []
            
            # پردازش تمام شیت‌ها
            for sheet_name in sheet_names:
                print(f"📖 خواندن شیت: {sheet_name}")
                
                try:
                    # خواندن شیت بدون فرض سرستون
                    df_raw = read_excel(input_path, sheet_name=sheet_name, header=None)
                    
                    if len(df_raw) == 0:
                        print(f"   ⚠️ شیت {sheet_name} خالی است")
                        continue
                    
                    # پیدا کردن ردیف سرستون (ردیف اول حاوی کلمات کلیدی)
                    header_row = 0
                    for i in range(min(3, len(df_raw))):  # بررسی ۳ ردیف اول
                        row_text = df_raw.iloc[i].astype(str).str.lower()
                        header_keywords = ['شماره', 'تاریخ', 'سند', 'حساب', 'شرح', 'بدهکار', 'بستانکار']
                        header_count = sum(any(keyword in cell for keyword in header_keywords) for cell in row_text)
                        
                        if header_count >= 2:  # اگر حداقل ۲ کلمه سرستون پیدا شد
                            header_row = i
                            break
                    
                    # خواندن شیت با سرستون صحیح
                    df = read_excel(input_path, sheet_name=sheet_name, header=header_row)
                    
                    # حذف ردیف‌های خالی
                    df = df.dropna(how='all')
                    
                    # حذف ستون‌های Unnamed
                    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                    
                    # اضافه کردن ستون نام شیت
                    df[self.sheet_name_column] = sheet_name
                    
                    # اضافه کردن به لیست داده‌ها
                    all_data.append(df)
                    
                    print(f"   ✅ {len(df)} رکورد از شیت {sheet_name} خوانده شد")
                        
                except Exception as e:
                    print(f"   ⚠️ خطا در خواندن شیت {sheet_name}: {str(e)}")
                    continue
            
            if not all_data:
                print("❌ هیچ داده‌ای از شیت‌ها خوانده شد")
                return None
            
            # ترکیب تمام داده‌ها با pd.concat()
            print("🔗 ترکیب داده‌ها...")
            combined_df = concat(all_data, axis=0, ignore_index=True, sort=False)
            
            print(f"✅ ترکیب کامل شد: {len(combined_df)} رکورد در مجموع")
            
            # تحلیل و حذف ستون‌های خالی
            print("🔍 تحلیل کامل بودن ستون‌ها...")
            combined_df = self.analyze_column_completeness(combined_df, threshold=0.1)
            
            # تولید نام فایل خروجی
            input_path_obj = Path(input_path)
            output_filename = f"{input_path_obj.stem}{output_suffix}{input_path_obj.suffix}"
            
            counter = 1
            while os.path.exists(output_filename):
                output_filename = f"{input_path_obj.stem}{output_suffix}_{counter}{input_path_obj.suffix}"
                counter += 1
            
            # ذخیره فایل
            combined_df.to_excel(output_filename, index=False)
            print(f"💾 فایل ترکیب شده ذخیره شد: {output_filename}")
            
            # نمایش خلاصه
            print(f"\n📊 خلاصه ترکیب:")
            print(f"   تعداد شیت‌های ترکیب شده: {len(sheet_names)}")
            print(f"   کل رکوردها: {len(combined_df)}")
            print(f"   تعداد ستون‌ها: {len(combined_df.columns)}")
            
            # نمایش تعداد رکوردها در هر شیت
            for sheet_name in sheet_names:
                count = len(combined_df[combined_df[self.sheet_name_column] == sheet_name])
                print(f"   - {sheet_name}: {count} رکورد")
            
            return output_filename
            
        except Exception as e:
            print(f"❌ خطا در ترکیب شیت‌ها: {str(e)}")
            return None
    
    def combine_sheets(self, input_path, output_suffix="_combined"):
        """ترکیب تمام شیت‌های یک فایل اکسل در یک شیت واحد"""
        return self.combine_sheets_simple(input_path, output_suffix)



class SimpleSmartExtractor:
    """نسخه ساده استخراج کننده اطلاعات"""
    
    def __init__(self):
        self.column_mapping = {
            'شرح': 'description',
            'شرح سند': 'description',
            'شرح رديف سند': 'description',
            'مبلغ': 'amount',
            'تاریخ': 'date',
            'تاریخ سند': 'date',
            'شماره سند': 'document_number',
            'شماره حساب': 'account_number',
            'كد حساب': 'account_number',
            'نام حساب': 'account_name',
            'description': 'description',
            'amount': 'amount',
            'date': 'date',
            'document_number': 'document_number',
            'account_number': 'account_number',
            'account_name': 'account_name',
        }
    
    def extract_invoice_number(self, text):
        """استخراج شماره صورت‌وضعیت"""
        if not text:
            return None
        
        patterns = [
            r'صورت وضعیت\s*[:؛]?\s*(\d+)',
            r'شماره\s*صورت وضعیت\s*[:؛]?\s*(\d+)',
            r'صورت وضعیت شماره\s*(\d+)',
            r'صورت وضعيت\s*[:؛]?\s*(\d+)',
            r'ش.\s*و.\s*(\d+)',
            r'شماره\s*[:؛]?\s*(\d+)',
            r'Invoice\s*#?\s*(\d+)',
            r'INV\s*(\d+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, str(text), re.IGNORECASE)
            if match:
                return match.group(1)
        
        return None
    
    def extract_currency_info(self, text):
        """استخراج اطلاعات ارز"""
        if not text:
            return {'amount': None, 'currency': None, 'rate': None}
        
        # الگوهای بهبود یافته برای شناسایی دقیق مبلغ و نرخ
        patterns = [
            # فارسی - با نرخ (مثال: 8،276/74 یورو به نرخ 28500)
            r'(\d[\d،,\.\/]*)\s*(یورو|دلار|يورو|ريال|ریال)\s*(?:به نرخ|با نرخ|نرخ|في|@|ارزش)\s*(\d[\d،,\.]*)\s*(?:ريال|ریال)?',
            # فارسی - با نرخ (مثال: 8،276/74 یورو نرخ 28500)
            r'(\d[\d،,\.\/]*)\s*(یورو|دلار|يورو|ريال|ریال)\s*(?:نرخ)\s*(\d[\d،,\.]*)\s*(?:ريال|ریال)?',
            # فارسی - با نرخ (مثال: 8،276/74 یورو فی 28500)
            r'(\d[\d،,\.\/]*)\s*(یورو|دلار|يورو|ريال|ریال)\s*(?:في|@)\s*(\d[\d،,\.]*)\s*(?:ريال|ریال)?',
            # فارسی - با نرخ و خط تیره (مثال: 210154 يورو با نرخ- 16093 ريال)
            r'(\d[\d،,\.\/]*)\s*(یورو|دلار|يورو|ريال|ریال)\s*(?:با نرخ|نرخ)\s*[-–]\s*(\d[\d،,\.]*)\s*(?:ريال|ریال)?',
            # فارسی - نرخ بعد از ارز (مثال: 777635 يورو 14874 ريال)
            r'(\d[\d،,\.\/]*)\s*(یورو|دلار|يورو|ريال|ریال)\s+(\d[\d،,\.]*)\s*(?:ريال|ریال)',
            # فارسی - بدون نرخ
            r'(\d[\d،,\.\/]*)\s*(یورو|دلار|يورو|ريال|ریال)',
            # انگلیسی - با نرخ
            r'(\d[\d,\.]*)\s*(EUR|USD|IRR|Euro|Dollar|Rial)\s*(?:rate|@|at|value)\s*(\d[\d,\.]*)',
            # انگلیسی - بدون نرخ
            r'(\d[\d,\.]*)\s*(EUR|USD|IRR|Euro|Dollar|Rial)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, str(text))
            if match:
                groups = match.groups()
                amount_str = groups[0] if groups[0] else None
                currency = groups[1]
                rate = groups[2] if len(groups) > 2 else None
                
                try:
                    # تبدیل مبلغ ارزی - پردازش فرمت فارسی
                    if amount_str:
                        # تشخیص نوع فرمت عدد
                        if '/' in amount_str:
                            # فرمت فارسی با اسلش (32،368/44) - اسلش به عنوان ممیز
                            amount_str = amount_str.replace('،', '').replace(',', '').replace('.', '')
                            amount_str = amount_str.replace('/', '.')
                        elif '.' in amount_str:
                            # فرمت انگلیسی با نقطه
                            if amount_str.count('.') == 1:
                                # یک نقطه - ممیز ارز (28679.3)
                                amount_str = amount_str.replace('،', '').replace(',', '')
                            else:
                                # بیش از یک نقطه - جداکننده هزارگان (48.638)
                                amount_str = amount_str.replace('،', '').replace(',', '').replace('.', '')
                        else:
                            # فرمت با جداکننده‌های هزارگان
                            amount_str = amount_str.replace('،', '').replace(',', '').replace('.', '')
                        
                        # حذف کاراکترهای غیرعددی (به جز نقطه)
                        amount_str = re.sub(r'[^\d\.]', '', amount_str)
                        amount = float(amount_str) if amount_str else None
                    else:
                        amount = None
                    
                    # تبدیل نرخ
                    if rate:
                        # برای نرخ هم جداکننده‌ها را حذف کنیم
                        rate_str = rate.replace('،', '').replace(',', '').replace('.', '')
                        rate_str = re.sub(r'[^\d\.]', '', rate_str)
                        rate = float(rate_str) if rate_str else None
                except (ValueError, TypeError) as e:
                    print(f"⚠️ خطا در تبدیل عدد: {amount_str} یا {rate} - {str(e)}")
                    continue
                
                return {
                    'amount': amount,
                    'currency': currency,
                    'rate': rate
                }
        
        return {'amount': None, 'currency': None, 'rate': None}
    
    def extract_company(self, text):
        """استخراج نام شرکت"""
        if not text:
            return None
        
        # الگوهای شناسایی نام شرکت بعد از کلمه "شرکت"
        patterns = [
            r'شركت\s+([^\s،]+)',
            r'شرکت\s+([^\s،]+)',
            r'شركت\s+([^\s،]+)\s+([^\s،]+)?',
            r'شرکت\s+([^\s،]+)\s+([^\s،]+)?',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, str(text))
            if match:
                # ترکیب کلمات نام شرکت
                company_parts = [part for part in match.groups() if part]
                if company_parts:
                    return ' '.join(company_parts)
        
        # روش قدیمی برای پشتیبانی از شرکت‌های شناخته شده
        companies = ['ایران', 'ایرایتک', 'پترو ساحل', 'فرآب', 'ناردیس', 'خارک', 'پتروساحل', 'پترو ساحل خلیج فارس']
        for company in companies:
            if company in str(text):
                return company
        
        return None
    
    def detect_document_type(self, text):
        """تشخیص نوع سند"""
        if not text:
            return 'سند متفرقه'
        
        desc_lower = str(text).lower()
        
        if 'تسعیر' in desc_lower or 'نرخ' in desc_lower:
            return 'تسعیر ارز'
        elif 'صورت وضعیت' in desc_lower or 'صورتوضعیت' in desc_lower:
            return 'صورت وضعیت'
        elif 'چک' in desc_lower:
            return 'چک'
        elif 'انتقال' in desc_lower or 'مانده' in desc_lower:
            return 'انتقال'
        else:
            return 'سند متفرقه'
    
    def process_excel_file(self, input_path, output_suffix="_extracted"):
        """پردازش کامل فایل اکسل"""
        print(f"🚀 شروع پردازش فایل: {input_path}")
        
        # خواندن فایل
        try:
            df = read_excel(input_path)
            print(f"✅ فایل خوانده شد: {len(df)} رکورد")
        except Exception as e:
            print(f"❌ خطا در خواندن فایل: {str(e)}")
            return None
        
        # استانداردسازی ستون‌ها
        original_columns = df.columns.tolist()
        df.columns = [self.column_mapping.get(str(col).strip(), str(col).strip()) for col in df.columns]
        
        # حذف ستون‌های تکراری
        df = df.loc[:, ~df.columns.duplicated()]
        
        print(f"   📊 ستون‌ها پس از استانداردسازی: {list(df.columns)}")
        
        # استخراج اطلاعات
        print("🔍 استخراج اطلاعات از شرح...")
        
        if 'description' not in df.columns:
            print(f"   ⚠️ ستون 'description' یافت نشد. ستون‌های موجود: {list(df.columns)}")
            # سعی کنیم ستون شرح را پیدا کنیم
            description_columns = [col for col in df.columns if 'description' in col.lower() or 'شرح' in col]
            if description_columns:
                print(f"   🔍 ستون‌های شرح پیدا شده: {description_columns}")
                descriptions = df[description_columns[0]].astype(str).tolist()
            else:
                print("   ❌ هیچ ستون شرحی یافت نشد")
                descriptions = []
        else:
            descriptions = df['description'].astype(str).tolist()
        
        # ستون‌های جدید
        invoice_numbers = []
        currency_amounts = []
        currency_types = []
        exchange_rates = []
        company_names = []
        document_types = []
        
        for desc in descriptions:
            invoice_numbers.append(self.extract_invoice_number(desc))
            
            currency_info = self.extract_currency_info(desc)
            currency_amounts.append(currency_info['amount'])
            currency_types.append(currency_info['currency'])
            exchange_rates.append(currency_info['rate'])
            
            company_names.append(self.extract_company(desc))
            document_types.append(self.detect_document_type(desc))
        
        # افزودن ستون‌های جدید
        df['شماره_وضعیت'] = invoice_numbers
        df['مبلغ_ارزی'] = currency_amounts
        df['نوع_ارز'] = currency_types
        df['نرخ_ارز'] = exchange_rates
        df['نام_شرکت'] = company_names
        df['نوع_سند'] = document_types
        
        # تولید نام فایل خروجی
        input_path_obj = Path(input_path)
        output_filename = f"{input_path_obj.stem}{output_suffix}{input_path_obj.suffix}"
        
        counter = 1
        while os.path.exists(output_filename):
            output_filename = f"{input_path_obj.stem}{output_suffix}_{counter}{input_path_obj.suffix}"
            counter += 1
        
        # ذخیره فایل
        try:
            df.to_excel(output_filename, index=False)
            print(f"💾 فایل خروجی ذخیره شد: {output_filename}")
            
            # نمایش خلاصه
            invoice_count = df['شماره_وضعیت'].notna().sum()
            currency_count = df['مبلغ_ارزی'].notna().sum()
            company_count = df['نام_شرکت'].notna().sum()
            
            print(f"\n📊 خلاصه نتایج:")
            print(f"   کل رکوردها: {len(df)}")
            print(f"   شماره‌های وضعیت استخراج شده: {invoice_count}")
            print(f"   اطلاعات ارز استخراج شده: {currency_count}")
            print(f"   شرکت‌های شناسایی شده: {company_count}")
            
            return output_filename
            
        except Exception as e:
            print(f"❌ خطا در ذخیره فایل: {str(e)}")
            return None


def process_all_integration(input_path, output_suffix="_all"):
    """پردازش کامل یکپارچه: ترکیب شیت‌ها + استخراج اطلاعات + فیلتر ستون‌ها"""
    print(f"🚀 شروع پردازش یکپارچه کامل: {input_path}")
    
    # مرحله ۱: ترکیب شیت‌ها
    print("\n📋 مرحله ۱: ترکیب تمام شیت‌ها")
    combiner = ExcelSheetCombiner()
    combined_file = combiner.combine_sheets(input_path, "_combined_temp")
    
    if not combined_file:
        print("❌ خطا در ترکیب شیت‌ها")
        return None
    
    # مرحله ۲: استخراج اطلاعات
    print("\n🔍 مرحله ۲: استخراج اطلاعات هوشمند")
    extractor = SimpleSmartExtractor()
    final_file = extractor.process_excel_file(combined_file, output_suffix)
    
    if not final_file:
        print("❌ خطا در استخراج اطلاعات")
        return None
    
    # حذف فایل موقت
    try:
        os.remove(combined_file)
        print(f"🗑️ فایل موقت حذف شد: {combined_file}")
    except:
        pass
    
    return final_file


def main():
    """تابع اصلی"""
    parser = argparse.ArgumentParser(
        description='سیستم استخراج هوشمند اطلاعات از فایل‌های اکسل',
        epilog="""
نمونه استفاده:
  python simple_standalone.py data.xlsx
  python simple_standalone.py data.xlsx -o "_processed"
  python simple_standalone.py data.xlsx --combine-sheets
  python simple_standalone.py data.xlsx --combine-sheets -o "_combined"
  python simple_standalone.py data.xlsx all_integration -o "_all"
        """
    )
    
    parser.add_argument('input_file', help='مسیر فایل اکسل ورودی')
    parser.add_argument('operation', nargs='?', help='نوع عملیات (all_integration برای پردازش کامل)')
    parser.add_argument('-o', '--output', help='پسوند نام فایل خروجی', default='_extracted')
    parser.add_argument('--combine-sheets', action='store_true', 
                       help='ترکیب تمام شیت‌های فایل اکسل در یک شیت واحد')
    
    args = parser.parse_args()
    
    # اعتبارسنجی فایل
    if not os.path.exists(args.input_file):
        print(f"❌ فایل {args.input_file} یافت نشد")
        return 1
    
    try:
        if args.operation == 'all_integration':
            # پردازش یکپارچه کامل
            output_file = process_all_integration(args.input_file, args.output)
        elif args.combine_sheets:
            # استفاده از کلاس ترکیب کننده شیت‌ها
            combiner = ExcelSheetCombiner()
            output_file = combiner.combine_sheets(args.input_file, args.output)
        else:
            # استفاده از کلاس استخراج کننده اطلاعات
            extractor = SimpleSmartExtractor()
            output_file = extractor.process_excel_file(args.input_file, args.output)
        
        if output_file:
            print(f"\n🎉 پردازش با موفقیت تکمیل شد!")
            print(f"📁 فایل خروجی: {output_file}")
            return 0
        else:
            return 1
            
    except Exception as e:
        print(f"❌ خطا در پردازش: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
