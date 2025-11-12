#!/usr/bin/env python3
"""
Standalone Intelligent Reconciliation System
سیستم مغایرت‌گیری هوشمند مستقل

این اسکریپت قابلیت‌های ماژول مغایرت‌گیری هوشمند را به صورت مستقل ارائه می‌دهد.
می‌توانید دو فایل اکسل را به آن بدهید و نتایج مغایرت‌گیری را دریافت کنید.

Usage:
    python standalone_reconciliation.py file_a.xlsx file_b.xlsx [output_file.xlsx]
"""

import pandas as pd
import sys
import os
import re
import argparse
from pathlib import Path


class StandaloneReconciliation:
    """سیستم مغایرت‌گیری هوشمند مستقل"""
    
    def __init__(self):
        self.column_mapping = {
            # Persian column names
            'شرح': 'description',
            'مبلغ': 'amount',
            'تاریخ': 'date',
            'شماره سند': 'document_number',
            'شماره حساب': 'account_number',
            'نام حساب': 'account_name',
            
            # English column names
            'description': 'description',
            'amount': 'amount',
            'date': 'date',
            'document_number': 'document_number',
            'account_number': 'account_number',
            'account_name': 'account_name',
            
            # Common variations
            'شرح عملیات': 'description',
            'شرح تراکنش': 'description',
            'مبلغ تراکنش': 'amount',
            'مبلغ عملیات': 'amount',
            'تاریخ تراکنش': 'date',
            'تاریخ عملیات': 'date',
        }
    
    def extract_invoice_number(self, description):
        """استخراج شماره صورت‌وضعیت از شرح"""
        if not description:
            return None
        
        patterns = [
            r'صورت وضعیت\s*[:؛]?\s*(\d+)',
            r'شماره\s*صورت وضعیت\s*[:؛]?\s*(\d+)',
            r'شماره\s*[:؛]?\s*(\d+)',
            r'صورت وضعیت شماره\s*(\d+)',
            r'ش.\s*و.\s*(\d+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, str(description), re.IGNORECASE)
            if match:
                return match.group(1)
        
        return None
    
    def extract_check_number(self, description):
        """استخراج شماره چک از شرح"""
        if not description:
            return None
            
        patterns = [
            r'چک\s*شماره\s*(\d+)',
            r'شماره چک\s*(\d+)',
            r'چک\s*(\d+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, str(description), re.IGNORECASE)
            if match:
                return match.group(1)
        
        return None
    
    def extract_currency_info(self, description):
        """استخراج اطلاعات ارز از شرح"""
        if not description:
            return {'amount': None, 'currency': None, 'rate': None}
            
        patterns = [
            r'(\d[\d,\.]*)\s*(یورو|دلار|یورو|ريال)\s*(?:نرخ|با نرخ|في|@)\s*(\d[\d,\.]*)',
            r'(\d[\d,\.]*)\s*(یورو|دلار|یورو)',
            r'(\d[\d,\.]*)\s*(EUR|USD|IRR)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, str(description))
            if match:
                amount_str = match.group(1).replace(',', '') if match.group(1) else None
                amount = float(amount_str) if amount_str else None
                currency = match.group(2)
                rate = match.group(3) if len(match.groups()) > 2 else None
                if rate:
                    rate = float(rate.replace(',', ''))
                
                return {
                    'amount': amount,
                    'currency': currency,
                    'rate': rate
                }
        
        return {'amount': None, 'currency': None, 'rate': None}
    
    def extract_company(self, description):
        """استخراج نام شرکت از شرح"""
        if not description:
            return None
            
        companies = ['ایران', 'ایرایتک', 'پترو ساحل', 'فرآب', 'ناردیس', 'خارک']
        for company in companies:
            if company in str(description):
                return company
        return None
    
    def detect_document_type(self, description):
        """تشخیص نوع سند"""
        if not description:
            return 'سند متفرقه'
            
        desc_lower = str(description).lower()
        
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
    
    def _convert_to_float(self, value):
        """Convert string value to float, handling commas and Persian numbers"""
        if value is None:
            return 0.0
        
        try:
            # تبدیل به رشته و حذف فاصله و کاما
            value_str = str(value).strip()
            value_str = value_str.replace(',', '').replace(' ', '')
            
            # تبدیل اعداد فارسی به انگلیسی
            persian_digits = '۰۱۲۳۴۵۶۷۸۹'
            english_digits = '0123456789'
            for p, e in zip(persian_digits, english_digits):
                value_str = value_str.replace(p, e)
            
            # اگر رشته خالی شد
            if not value_str:
                return 0.0
            
            return float(value_str)
        except (ValueError, TypeError):
            return 0.0
    
    def _process_excel_file(self, file_path, company_label):
        """Process Excel file and extract data"""
        try:
            # Read all sheets and combine
            excel_file = pd.ExcelFile(file_path)
            all_sheets_data = []
            
            for sheet_name in excel_file.sheet_names:
                df_sheet = pd.read_excel(file_path, sheet_name=sheet_name)
                df_sheet['sheet_name'] = sheet_name
                all_sheets_data.append(df_sheet)
            
            # Combine all sheets
            combined_df = pd.concat(all_sheets_data, ignore_index=True)
            
            # Standardize column names
            combined_df = self._standardize_columns(combined_df)
            
            print(f"✅ فایل {company_label} پردازش شد: {len(combined_df)} رکورد")
            return combined_df
            
        except Exception as e:
            print(f"❌ خطا در پردازش فایل {company_label}: {str(e)}")
            raise
    
    def _standardize_columns(self, df):
        """Standardize column names for Persian and English"""
        # Rename columns
        df.columns = [self.column_mapping.get(str(col).strip(), str(col).strip()) for col in df.columns]
        return df
    
    def _calculate_similarity(self, text1, text2):
        """Calculate text similarity using simple algorithm"""
        if not text1 or not text2:
            return 0.0
        
        # تبدیل به حروف کوچک
        text1 = str(text1).lower()
        text2 = str(text2).lower()
        
        # محاسبه تشابه بر اساس کلمات مشترک
        words1 = set(text1.split())
        words2 = set(text2.split())
        
        if not words1 or not words2:
            return 0.0
        
        common_words = words1.intersection(words2)
        similarity = len(common_words) / max(len(words1), len(words2)) * 100
        
        return similarity
    
    def _extract_smart_data(self, description_a, description_b):
        """Extract smart data from descriptions"""
        # استفاده از توابع استخراج موجود
        invoice_number = self.extract_invoice_number(description_a) or self.extract_invoice_number(description_b)
        check_number = self.extract_check_number(description_a) or self.extract_check_number(description_b)
        currency_info = self.extract_currency_info(description_a) or self.extract_currency_info(description_b)
        company = self.extract_company(description_a) or self.extract_company(description_b)
        doc_type = self.detect_document_type(description_a) or self.detect_document_type(description_b)
        
        return {
            'invoice_number': invoice_number,
            'check_number': check_number,
            'currency': currency_info['currency'],
            'foreign_amount': currency_info['amount'],
            'exchange_rate': currency_info['rate'],
            'company_name': company,
            'document_type': doc_type,
        }
    
    def _find_exact_matches(self, df_a, df_b):
        """Find exact matches based on invoice number and amount"""
        matches = []
        
        for idx_a, row_a in df_a.iterrows():
            description_a = str(row_a.get('description', ''))
            amount_a = self._convert_to_float(row_a.get('amount', 0))
            invoice_number = self.extract_invoice_number(description_a)
            
            if invoice_number:
                # جستجوی تطبیق دقیق در فایل دوم
                for idx_b, row_b in df_b.iterrows():
                    description_b = str(row_b.get('description', ''))
                    amount_b = self._convert_to_float(row_b.get('amount', 0))
                    
                    if (self.extract_invoice_number(description_b) == invoice_number and 
                        abs(amount_a - amount_b) < 0.01):  # اختلاف کمتر از 0.01
                        
                        # استخراج اطلاعات هوشمند
                        extracted_info = self._extract_smart_data(description_a, description_b)
                        
                        matches.append({
                            'statement_number': f"INV{invoice_number}",
                            'amount_a': amount_a,
                            'amount_b': amount_b,
                            'description_a': description_a,
                            'description_b': description_b,
                            'state': 'matched',
                            'similarity_score': 100.0,
                            'match_type': 'exact',
                            **extracted_info
                        })
                        break
        
        return matches
    
    def _find_fuzzy_matches(self, df_a, df_b):
        """Find fuzzy matches based on description similarity"""
        matches = []
        
        for idx_a, row_a in df_a.iterrows():
            description_a = str(row_a.get('description', ''))
            amount_a = self._convert_to_float(row_a.get('amount', 0))
            
            best_match = None
            best_score = 0
            
            for idx_b, row_b in df_b.iterrows():
                description_b = str(row_b.get('description', ''))
                amount_b = self._convert_to_float(row_b.get('amount', 0))
                
                # محاسبه تشابه شرح
                similarity = self._calculate_similarity(description_a, description_b)
                
                # محاسبه تشابه مبلغ (اختلاف کمتر از 1%)
                amount_similarity = 100.0 if abs(amount_a - amount_b) / max(amount_a, 1) < 0.01 else 0
                
                # امتیاز کلی
                total_score = (similarity * 0.7) + (amount_similarity * 0.3)
                
                if total_score > best_score and total_score > 70:  # آستانه تشابه
                    best_score = total_score
                    best_match = (row_b, total_score)
            
            if best_match:
                row_b, score = best_match
                description_b = str(row_b.get('description', ''))
                amount_b = self._convert_to_float(row_b.get('amount', 0))
                
                extracted_info = self._extract_smart_data(description_a, description_b)
                
                matches.append({
                    'statement_number': f"FUZZY{idx_a}",
                    'amount_a': amount_a,
                    'amount_b': amount_b,
                    'description_a': description_a,
                    'description_b': description_b,
                    'state': 'matched',
                    'similarity_score': score,
                    'match_type': 'fuzzy',
                    **extracted_info
                })
        
        return matches
    
    def _find_missing_records(self, df_a, df_b, existing_matches):
        """Find records that exist in only one file"""
        missing_records = []
        
        # استخراج رکوردهای تطبیق شده
        matched_a_indices = set()
        matched_b_indices = set()
        
        for match in existing_matches:
            if match['description_a']:
                matched_a_indices.add(match['description_a'])
            if match['description_b']:
                matched_b_indices.add(match['description_b'])
        
        # شناسایی رکوردهای مفقود در فایل A
        for idx_b, row_b in df_b.iterrows():
            description_b = str(row_b.get('description', ''))
            if description_b and description_b not in matched_b_indices:
                amount_b = self._convert_to_float(row_b.get('amount', 0))
                
                extracted_info = self._extract_smart_data('', description_b)
                
                missing_records.append({
                    'statement_number': f"MISSING_A{idx_b}",
                    'amount_a': 0,
                    'amount_b': amount_b,
                    'description_a': '',
                    'description_b': description_b,
                    'state': 'missing_a',
                    'similarity_score': 0.0,
                    'match_type': 'none',
                    **extracted_info
                })
        
        # شناسایی رکوردهای مفقود در فایل B
        for idx_a, row_a in df_a.iterrows():
            description_a = str(row_a.get('description', ''))
            if description_a and description_a not in matched_a_indices:
                amount_a = self._convert_to_float(row_a.get('amount', 0))
                
                extracted_info = self._extract_smart_data(description_a, '')
                
                missing_records.append({
                    'statement_number': f"MISSING_B{idx_a}",
                    'amount_a': amount_a,
                    'amount_b': 0,
                    'description_a': description_a,
                    'description_b': '',
                    'state': 'missing_b',
                    'similarity_score': 0.0,
                    'match_type': 'none',
                    **extracted_info
                })
        
        return missing_records
    
    def run_reconciliation(self, file_a_path, file_b_path, output_path=None):
        """Run the complete reconciliation process"""
        print("🚀 شروع مغایرت‌گیری هوشمند...")
        print("=" * 50)
        
        # پردازش فایل‌های اکسل
        df_a = self._process_excel_file(file_a_path, 'A')
        df_b = self._process_excel_file(file_b_path, 'B')
        
        # اجرای الگوریتم‌های تطبیق
        print("🔍 اجرای الگوریتم‌های تطبیق...")
        
        # تطبیق دقیق - بر اساس شماره صورت‌وضعیت و مبلغ
        exact_matches = self._find_exact_matches(df_a, df_b)
        print(f"   تطبیق دقیق: {len(exact_matches)} رکورد")
        
        # تطبیق فازی - بر اساس تشابه شرح و مبلغ
        fuzzy_matches = self._find_fuzzy_matches(df_a, df_b)
        print(f"   تطبیق فازی: {len(fuzzy_matches)} رکورد")
        
        # شناسایی رکوردهای مفقود
        all_matches = exact_matches + fuzzy_matches
        missing_records = self._find_missing_records(df_a, df_b, all_matches)
        print(f"   رکوردهای مفقود: {len(missing_records)} رکورد")
        
        # ترکیب تمام نتایج
        analysis_lines = exact_matches + fuzzy_matches + missing_records
        
        # تولید فایل نتایج
        if output_path:
            self._generate_result_file(analysis_lines, output_path)
        
        # نمایش آمار خلاصه
        self._display_summary(analysis_lines)
        
        return analysis_lines
    
    def _generate_result_file(self, analysis_lines, output_path):
        """Generate result Excel file"""
        try:
            # ایجاد دیتافریم نتایج
            result_data = []
            for line in analysis_lines:
                result_data.append({
                    'Statement Number': line.get('statement_number', ''),
                    'Amount A': line.get('amount_a', 0),
                    'Amount B': line.get('amount_b', 0),
                    'Difference': line.get('amount_b', 0) - line.get('amount_a', 0),
                    'Status': line.get('state', ''),
                    'Similarity Score': line.get('similarity_score', 0),
                    'Match Type': line.get('match_type', ''),
                    'Invoice Number': line.get('invoice_number', ''),
                    'Check Number': line.get('check_number', ''),
                    'Company': line.get('company_name', ''),
                    'Document Type': line.get('document_type', ''),
                    'Description A': line.get('description_a', ''),
                    'Description B': line.get('description_b', ''),
                })
            
            result_df = pd.DataFrame(result_data)
            
            # آمار خلاصه
            summary_data = {
                'Metric': ['Total Records', 'Matched Records', 'Mismatch Records', 'Missing in A', 'Missing in B'],
                'Count': [
                    len(analysis_lines),
                    len([l for l in analysis_lines if l['state'] == 'matched']),
                    len([l for l in analysis_lines if l['state'] == 'mismatch']),
                    len([l for l in analysis_lines if l['state'] == 'missing_a']),
                    len([l for l in analysis_lines if l['state'] == 'missing_b']),
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            
            # ایجاد فایل اکسل
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name='Reconciliation Results', index=False)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            print(f"✅ فایل نتایج ایجاد شد: {output_path}")
            
        except Exception as e:
            print(f"❌ خطا در تولید فایل نتایج: {str(e)}")
            raise
    
    def _display_summary(self, analysis_lines):
        """Display summary statistics"""
        print("\n📊 آمار خلاصه مغایرت‌گیری:")
        print("=" * 40)
        print(f"   کل رکوردها: {len(analysis_lines)}")
        print(f"   رکوردهای تطبیق شده: {len([l for l in analysis_lines if l['state'] == 'matched'])}")
        print(f"   رکوردهای مغایرت: {len([l for l in analysis_lines if l['state'] == 'mismatch'])}")
        print(f"   مفقود در فایل A: {len([l for l in analysis_lines if l['state'] == 'missing_a'])}")
        print(f"   مفقود در فایل B: {len([l for l in analysis_lines if l['state'] == 'missing_b'])}")
        print("=" * 40)


def main():
    """Main function for command line usage"""
    parser = argparse.ArgumentParser(description='سیستم مغایرت‌گیری هوشمند مستقل')
    parser.add_argument('file_a', help='مسیر فایل اکسل شرکت A')
    parser.add_argument('file_b', help='مسیر فایل اکسل شرکت B')
    parser.add_argument('-o', '--output', help='مسیر فایل خروجی (اختیاری)', default='reconciliation_results.xlsx')
    
    args = parser.parse_args()
    
    # بررسی وجود فایل‌ها
    if not os.path.exists(args.file_a):
        print(f"❌ فایل {args.file_a} یافت نشد")
        return
    
    if not os.path.exists(args.file_b):
        print(f"❌ فایل {args.file_b} یافت نشد")
        return
    
    # اجرای مغایرت‌گیری
    reconciliation = StandaloneReconciliation()
    try:
        results = reconciliation.run_reconciliation(args.file_a, args.file_b, args.output)
        print(f"\n🎉 مغایرت‌گیری با موفقیت تکمیل شد!")
        print(f"📁 فایل نتایج: {args.output}")
    except Exception as e:
        print(f"❌ خطا در مغایرت‌گیری: {str(e)}")


if __name__ == "__main__":
    main()
