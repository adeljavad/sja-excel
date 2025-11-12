#!/usr/bin/env python3
"""
تحلیل هوشمند شیت‌ها بر اساس مبالغ
Smart Sheet Analysis Based on Amounts

این اسکریپت به صورت هوشمند شیت‌ها را بر اساس مبالغ تطبیق می‌دهد
و از مپینگ ثابت استفاده نمی‌کند.
"""

import pandas as pd
import numpy as np
from difflib import SequenceMatcher
import argparse
import os


class SmartSheetAnalysis:
    """تحلیل هوشمند شیت‌ها بر اساس مبالغ"""
    
    def __init__(self):
        self.debit_keywords = ['بدهكار', 'دبیت', 'debit', 'بدهی', 'بدکار', 'بدهکار']
        self.credit_keywords = ['بستانكار', 'کریدیت', 'credit', 'بستانکاری', 'بستکار', 'بستانکار']
        # اولویت‌بندی: ابتدا ستون‌های ریالی را جستجو کنیم
        self.priority_keywords = ['ریالی', 'ریال', 'rial', 'ریال']
    
    def detect_amount_columns(self, df, file_label):
        """شناسایی ستون‌های مبلغی در فایل"""
        amount_columns = {}
        
        print(f"🔍 فایل {file_label} - بررسی ستون‌ها:")
        
        for col in df.columns:
            col_str = str(col).lower()
            
            # شناسایی ستون‌های بدهکار
            for keyword in self.debit_keywords:
                if keyword in col_str:
                    amount_columns['debit'] = col
                    print(f"   ✅ ستون بدهکار شناسایی شد: {col}")
                    break
            
            # شناسایی ستون‌های بستانکار
            for keyword in self.credit_keywords:
                if keyword in col_str:
                    amount_columns['credit'] = col
                    print(f"   ✅ ستون بستانکار شناسایی شد: {col}")
                    break
        
        # اگر ستون‌های مبلغی شناسایی نشدند، ستون‌های عددی را بررسی کنیم
        if not amount_columns:
            print(f"   ⚠️ ستون‌های مبلغی شناسایی نشد. بررسی ستون‌های عددی...")
            numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
            print(f"   ستون‌های عددی: {numeric_columns}")
            
            # اگر ستون‌های عددی وجود دارند، از آنها استفاده کنیم
            if numeric_columns:
                # فرض می‌کنیم اولین ستون عددی بدهکار است
                amount_columns['debit'] = numeric_columns[0]
                print(f"   🎯 استفاده از ستون عددی: {numeric_columns[0]}")
            else:
                # اگر ستون‌های عددی هم وجود ندارند، ستون‌های object را بررسی کنیم
                print(f"   🔍 بررسی ستون‌های object برای مقادیر عددی...")
                for col in df.columns:
                    col_str = str(col).lower()
                    # بررسی ستون‌های با نام‌های مشخص
                    if 'بدهکار' in col_str or 'دبیت' in col_str or 'debit' in col_str:
                        amount_columns['debit'] = col
                        print(f"   🎯 ستون بدهکار شناسایی شد (object): {col}")
                    elif 'بستانکار' in col_str or 'کریدیت' in col_str or 'credit' in col_str:
                        amount_columns['credit'] = col
                        print(f"   🎯 ستون بستانکار شناسایی شد (object): {col}")
        
        # اگر هنوز ستون‌ها شناسایی نشدند، از منطق پیشرفته‌تر استفاده کنیم
        if not amount_columns:
            print(f"   🔍 استفاده از منطق پیشرفته برای شناسایی ستون‌ها...")
            for col in df.columns:
                col_str = str(col)
                # بررسی ستون‌های Unnamed که حاوی مقادیر عددی هستند
                if 'unnamed' in col_str.lower():
                    # بررسی محتوای ستون
                    sample_values = df[col].dropna().head(5)
                    if len(sample_values) > 0:
                        # اگر مقادیر عددی هستند
                        if any(isinstance(val, (int, float)) for val in sample_values if val is not None):
                            if 'debit' not in amount_columns:
                                amount_columns['debit'] = col
                                print(f"   🎯 ستون بدهکار شناسایی شد (Unnamed): {col}")
                            elif 'credit' not in amount_columns:
                                amount_columns['credit'] = col
                                print(f"   🎯 ستون بستانکار شناسایی شد (Unnamed): {col}")
        
        print(f"🔍 فایل {file_label} - ستون‌های شناسایی شده: {amount_columns}")
        return amount_columns
    
    def convert_amount_columns(self, df, amount_cols):
        """تبدیل ستون‌های مبلغی به عدد"""
        for col_type, col_name in amount_cols.items():
            if col_name in df.columns:
                # تبدیل مقادیر به عدد
                df[col_name] = pd.to_numeric(df[col_name], errors='coerce')
                # جایگزینی مقادیر NaN با 0
                df[col_name] = df[col_name].fillna(0)
                print(f"   🔄 ستون {col_name} به عدد تبدیل شد")
        
        return df
    
    def group_by_sheet(self, df, file_label):
        """گروه‌بندی داده‌ها بر اساس نام شیت (با جمع‌بندی شیت‌های هم‌نام)"""
        if 'نام_شیت' not in df.columns:
            raise ValueError(f"ستون 'نام_شیت' در فایل {file_label} وجود ندارد")
        
        # شناسایی ستون‌های مبلغی
        amount_cols = self.detect_amount_columns(df, file_label)
        
        # اضافه کردن ستون نام نرمال‌شده
        df['نام_شیت_نرمال'] = df['نام_شیت'].apply(self._normalize_sheet_name)
        
        # جمع‌بندی تعداد رکوردها بر اساس نام نرمال‌شده
        sheet_summary = df.groupby('نام_شیت_نرمال').size().reset_index(name='تعداد_رکورد')
        
        # جمع‌بندی مبالغ بدهکار بر اساس نام نرمال‌شده
        if 'debit' in amount_cols:
            debit_sum = df.groupby('نام_شیت_نرمال')[amount_cols['debit']].sum().reset_index(name='جمع_بدهکار')
            sheet_summary = pd.merge(sheet_summary, debit_sum, on='نام_شیت_نرمال')
        
        # جمع‌بندی مبالغ بستانکار بر اساس نام نرمال‌شده
        if 'credit' in amount_cols:
            credit_sum = df.groupby('نام_شیت_نرمال')[amount_cols['credit']].sum().reset_index(name='جمع_بستانکار')
            sheet_summary = pd.merge(sheet_summary, credit_sum, on='نام_شیت_نرمال')
        
        # محاسبه مبلغ خالص
        if 'جمع_بدهکار' in sheet_summary.columns and 'جمع_بستانکار' in sheet_summary.columns:
            sheet_summary['مبلغ_خالص'] = sheet_summary['جمع_بدهکار'] - sheet_summary['جمع_بستانکار']
        elif 'جمع_بدهکار' in sheet_summary.columns:
            sheet_summary['مبلغ_خالص'] = sheet_summary['جمع_بدهکار']
        elif 'جمع_بستانکار' in sheet_summary.columns:
            sheet_summary['مبلغ_خالص'] = -sheet_summary['جمع_بستانکار']
        else:
            sheet_summary['مبلغ_خالص'] = 0
        
        # تغییر نام ستون به نام اصلی
        sheet_summary = sheet_summary.rename(columns={'نام_شیت_نرمال': 'نام_شیت'})
        
        print(f"   📊 شیت‌های {file_label} بر اساس نام نرمال‌شده گروه‌بندی شدند")
        
        return sheet_summary
    
    def find_amount_matches(self, summary_a, summary_b):
        """پیدا کردن تمام ترکیبات ممکن بین شیت‌ها"""
        matches = []
        
        print(f"   🔍 بررسی {len(summary_a)} × {len(summary_b)} = {len(summary_a) * len(summary_b)} ترکیب ممکن")
        
        for idx_a, row_a in summary_a.iterrows():
            sheet_a = row_a['نام_شیت']
            debit_a = row_a.get('جمع_بدهکار', 0)
            credit_a = row_a.get('جمع_بستانکار', 0)
            
            # بررسی تمام شیت‌های فایل B
            for idx_b, row_b in summary_b.iterrows():
                sheet_b = row_b['نام_شیت']
                debit_b = row_b.get('جمع_بدهکار', 0)
                credit_b = row_b.get('جمع_بستانکار', 0)
                
                # محاسبه تشابه‌های مختلف
                debit_to_debit = self._calculate_amount_similarity(debit_a, debit_b)
                credit_to_credit = self._calculate_amount_similarity(credit_a, credit_b)
                debit_to_credit = self._calculate_amount_similarity(debit_a, credit_b)  # بدهکار A با بستانکار B
                credit_to_debit = self._calculate_amount_similarity(credit_a, debit_b)  # بستانکار A با بدهکار B
                
                # انتخاب بهترین تشابه
                best_similarity = max(debit_to_debit, credit_to_credit, debit_to_credit, credit_to_debit)
                
                # تشخیص نوع تطابق
                match_type = "نامشخص"
                if best_similarity == debit_to_debit:
                    match_type = "بدهکار ↔ بدهکار"
                elif best_similarity == credit_to_credit:
                    match_type = "بستانکار ↔ بستانکار"
                elif best_similarity == debit_to_credit:
                    match_type = "بدهکار ↔ بستانکار"
                elif best_similarity == credit_to_debit:
                    match_type = "بستانکار ↔ بدهکار"
                
                # محاسبه تشابه نام
                name_similarity = self._calculate_name_similarity(sheet_a, sheet_b)
                
                # محاسبه امتیاز کلی (حتی اگر تشابه مبلغ کم باشد)
                overall_score = (best_similarity * 0.7) + (name_similarity * 0.3)
                
                matches.append({
                    'نام_شیت_A': sheet_a,
                    'نام_شیت_B': sheet_b,
                    'بدهکار_A': debit_a,
                    'بستانکار_A': credit_a,
                    'بدهکار_B': debit_b,
                    'بستانکار_B': credit_b,
                    'تشابه_مبلغ': best_similarity,
                    'تشابه_نام': name_similarity,
                    'نوع_تطابق': match_type,
                    'امتیاز_کلی': overall_score
                })
        
        # مرتب‌سازی بر اساس امتیاز کلی
        matches.sort(key=lambda x: x['امتیاز_کلی'], reverse=True)
        return matches
    
    def _calculate_amount_similarity(self, amount_a, amount_b):
        """محاسبه تشابه مبالغ"""
        if amount_a == 0 and amount_b == 0:
            return 100.0
        
        # بررسی تطابق بدهکار با بستانکار (مقادیر مخالف)
        if abs(amount_a + amount_b) < 0.01:
            return 100.0
        
        # بررسی تطابق مستقیم
        if abs(amount_a - amount_b) / max(abs(amount_a), abs(amount_b), 1) < 0.01:
            return 100.0
        
        return 0.0
    
    def _normalize_sheet_name(self, sheet_name):
        """نرمال‌سازی نام شیت با حذف نام شرکت از انتها"""
        if not sheet_name:
            return sheet_name
        
        name = str(sheet_name).strip()
        
        # حذف نام شرکت‌ها از انتهای نام شیت
        company_names = ['اير', 'پتروساحل', 'نارديس', 'شركت', 'شرکت']
        
        for company in company_names:
            if name.endswith(company):
                name = name[:-len(company)].strip()
            elif name.endswith(f"- {company}"):
                name = name[:-len(f"- {company}")].strip()
            elif name.endswith(f" - {company}"):
                name = name[:-len(f" - {company}")].strip()
        
        # حذف کاراکترهای اضافی از انتها
        name = name.rstrip(' -')
        
        return name
    
    def _calculate_name_similarity(self, text1, text2):
        """محاسبه تشابه نام"""
        if not text1 or not text2:
            return 0.0
        
        # نرمال‌سازی نام‌ها
        normalized1 = self._normalize_sheet_name(text1)
        normalized2 = self._normalize_sheet_name(text2)
        
        # محاسبه تشابه بر اساس نام‌های نرمال‌شده
        similarity = SequenceMatcher(None, normalized1.lower(), normalized2.lower()).ratio() * 100
        
        return similarity
    
    def generate_analysis_report(self, file_a_path, file_b_path, output_path):
        """ایجاد گزارش تحلیل کامل"""
        print("🧠 شروع تحلیل هوشمند شیت‌ها...")
        print("=" * 50)
        
        # خواندن فایل‌ها
        df_a = pd.read_excel(file_a_path)
        df_b = pd.read_excel(file_b_path)
        
        print(f"✅ فایل A خوانده شد: {len(df_a)} رکورد")
        print(f"✅ فایل B خوانده شد: {len(df_b)} رکورد")
        
        # گروه‌بندی داده‌ها بر اساس شیت
        print("\n📊 گروه‌بندی داده‌ها...")
        summary_a = self.group_by_sheet(df_a, 'A')
        summary_b = self.group_by_sheet(df_b, 'B')
        
        print(f"📈 فایل A: {len(summary_a)} شیت")
        print(f"📈 فایل B: {len(summary_b)} شیت")
        
        # پیدا کردن تطابق‌های مبلغی
        print("\n🔍 جستجوی تطابق‌های مبلغی...")
        matches = self.find_amount_matches(summary_a, summary_b)
        
        print(f"🎯 تعداد تطابق‌های یافت شده: {len(matches)}")
        
        # ایجاد گزارش
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. خلاصه شیت‌های فایل A
            summary_a.to_excel(writer, sheet_name='خلاصه_فایل_A', index=False)
            
            # 2. خلاصه شیت‌های فایل B
            summary_b.to_excel(writer, sheet_name='خلاصه_فایل_B', index=False)
            
            # 3. تطابق‌های یافت شده
            if matches:
                matches_df = pd.DataFrame(matches)
                matches_df.to_excel(writer, sheet_name='تطابق‌ها', index=False)
            
            # 4. آمار کلی
            stats_data = {
                'آمار': [
                    'تعداد شیت‌های فایل A',
                    'تعداد شیت‌های فایل B',
                    'تعداد تطابق‌های یافت شده',
                    'میانگین امتیاز تطابق',
                    'بیشترین امتیاز تطابق',
                    'کمترین امتیاز تطابق'
                ],
                'مقدار': [
                    len(summary_a),
                    len(summary_b),
                    len(matches),
                    f"{matches_df['امتیاز_کلی'].mean():.1f}" if matches else "0",
                    f"{matches_df['امتیاز_کلی'].max():.1f}" if matches else "0",
                    f"{matches_df['امتیاز_کلی'].min():.1f}" if matches else "0"
                ]
            }
            stats_df = pd.DataFrame(stats_data)
            stats_df.to_excel(writer, sheet_name='آمار_کلی', index=False)
        
        # نمایش نتایج
        self._display_results(summary_a, summary_b, matches)
        
        print(f"\n✅ گزارش تحلیل ایجاد شد: {output_path}")
        return matches
    
    def _display_results(self, summary_a, summary_b, matches):
        """نمایش نتایج در کنسول"""
        print(f"\n📊 خلاصه شیت‌ها:")
        print("=" * 40)
        print(f"فایل A: {len(summary_a)} شیت")
        print(f"فایل B: {len(summary_b)} شیت")
        
        if matches:
            print(f"\n🏆 بهترین تطابق‌ها:")
            for i, match in enumerate(matches[:5]):
                print(f"  {i+1}. {match['نام_شیت_A']} ↔ {match['نام_شیت_B']}")
                print(f"     بدهکار A: {match['بدهکار_A']:,.0f} | بستانکار A: {match['بستانکار_A']:,.0f}")
                print(f"     بدهکار B: {match['بدهکار_B']:,.0f} | بستانکار B: {match['بستانکار_B']:,.0f}")
                print(f"     تشابه مبلغ: {match['تشابه_مبلغ']:.1f}% | تشابه نام: {match['تشابه_نام']:.1f}%")
                print(f"     نوع تطابق: {match['نوع_تطابق']} | امتیاز کلی: {match['امتیاز_کلی']:.1f}")
                print()
        else:
            print("\n⚠️ هیچ تطابق مبلغی یافت نشد")


def main():
    """تابع اصلی برای اجرای تحلیل"""
    parser = argparse.ArgumentParser(description='تحلیل هوشمند شیت‌ها بر اساس مبالغ')
    parser.add_argument('file_a', help='مسیر فایل اکسل شرکت A')
    parser.add_argument('file_b', help='مسیر فایل اکسل شرکت B')
    parser.add_argument('-o', '--output', help='مسیر فایل خروجی', default='smart_sheet_analysis.xlsx')
    
    args = parser.parse_args()
    
    # بررسی وجود فایل‌ها
    if not os.path.exists(args.file_a):
        print(f"❌ فایل {args.file_a} یافت نشد")
        return
    
    if not os.path.exists(args.file_b):
        print(f"❌ فایل {args.file_b} یافت نشد")
        return
    
    # اجرای تحلیل
    analyzer = SmartSheetAnalysis()
    try:
        results = analyzer.generate_analysis_report(args.file_a, args.file_b, args.output)
        print(f"\n🎉 تحلیل هوشمند با موفقیت تکمیل شد!")
        print(f"📁 فایل گزارش: {args.output}")
    except Exception as e:
        print(f"❌ خطا در تحلیل: {str(e)}")
        raise


if __name__ == "__main__":
    main()
