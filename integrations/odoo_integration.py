"""
Odoo integration adapter for Smart Extractor
آداپتور یکپارچه‌سازی با اودوو برای سیستم استخراج هوشمند
"""

from typing import List, Dict, Any, Optional
from ..core.extractors import SmartExtractor
from ..core.models import ExtractionResult, BatchExtractionResult


class OdooIntegration:
    """کلاس یکپارچه‌سازی با اودوو"""
    
    def __init__(self):
        self.extractor = SmartExtractor()
    
    def extract_from_account_move_lines(self, move_lines: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """استخراج اطلاعات از خطوط سند حسابداری اودوو"""
        extracted_data = []
        
        for line in move_lines:
            description = line.get('name', '') or line.get('description', '')
            
            if description:
                result = self.extractor.extract_from_text(description)
                
                extracted_line = {
                    'move_line_id': line.get('id'),
                    'original_description': description,
                    'extracted_data': result.to_dict(),
                    'invoice_number': result.invoice_number,
                    'currency_amount': result.currency_info.amount if result.currency_info else None,
                    'currency_type': result.currency_info.currency if result.currency_info else None,
                    'exchange_rate': result.currency_info.rate if result.currency_info else None,
                    'company_name': result.company_name,
                    'document_type': result.document_type,
                    'confidence': result.confidence
                }
                
                extracted_data.append(extracted_line)
        
        return extracted_data
    
    def create_extracted_fields(self, model: str) -> List[Dict[str, Any]]:
        """ایجاد فیلدهای جدید برای مدل اودوو"""
        fields = [
            {
                'name': 'invoice_number_extracted',
                'field_description': 'شماره وضعیت استخراج شده',
                'type': 'char',
                'string': 'شماره وضعیت'
            },
            {
                'name': 'currency_amount_extracted',
                'field_description': 'مبلغ ارزی استخراج شده',
                'type': 'float',
                'string': 'مبلغ ارزی'
            },
            {
                'name': 'currency_type_extracted',
                'field_description': 'نوع ارز استخراج شده',
                'type': 'char',
                'string': 'نوع ارز'
            },
            {
                'name': 'exchange_rate_extracted',
                'field_description': 'نرخ ارز استخراج شده',
                'type': 'float',
                'string': 'نرخ ارز'
            },
            {
                'name': 'company_name_extracted',
                'field_description': 'نام شرکت استخراج شده',
                'type': 'char',
                'string': 'نام شرکت'
            },
            {
                'name': 'document_type_extracted',
                'field_description': 'نوع سند استخراج شده',
                'type': 'char',
                'string': 'نوع سند'
            },
            {
                'name': 'extraction_confidence',
                'field_description': 'اطمینان استخراج',
                'type': 'float',
                'string': 'اطمینان'
            }
        ]
        
        return fields
    
    def update_records_with_extracted_data(self, model: str, records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """به‌روزرسانی رکوردها با داده‌های استخراج شده"""
        updated_records = []
        
        for record in records:
            description = record.get('name') or record.get('description') or ''
            
            if description:
                result = self.extractor.extract_from_text(description)
                
                updated_record = record.copy()
                updated_record.update({
                    'invoice_number_extracted': result.invoice_number,
                    'currency_amount_extracted': result.currency_info.amount if result.currency_info else None,
                    'currency_type_extracted': result.currency_info.currency if result.currency_info else None,
                    'exchange_rate_extracted': result.currency_info.rate if result.currency_info else None,
                    'company_name_extracted': result.company_name,
                    'document_type_extracted': result.document_type,
                    'extraction_confidence': result.confidence
                })
                
                updated_records.append(updated_record)
        
        return updated_records
    
    def batch_extract_from_model(self, model: str, domain: List = None) -> Dict[str, Any]:
        """استخراج دسته‌ای از یک مدل اودوو"""
        # این تابع نیاز به محیط اودوو دارد
        # در محیط واقعی، این تابع با استفاده از ORM اودوو اجرا می‌شود
        
        # نمونه کد برای محیط اودوو:
        """
        records = self.env[model].search(domain or [])
        descriptions = [rec.name or rec.description or '' for rec in records]
        
        batch_result = self.extractor.extract_batch(descriptions)
        
        # به‌روزرسانی رکوردها
        for record, result in zip(records, batch_result.results):
            record.write({
                'invoice_number_extracted': result.invoice_number,
                'currency_amount_extracted': result.currency_info.amount if result.currency_info else None,
                'currency_type_extracted': result.currency_info.currency if result.currency_info else None,
                'exchange_rate_extracted': result.currency_info.rate if result.currency_info else None,
                'company_name_extracted': result.company_name,
                'document_type_extracted': result.document_type,
                'extraction_confidence': result.confidence
            })
        
        return batch_result.to_dict()
        """
        
        # در این نسخه، فقط ساختار را برمی‌گردانیم
        return {
            'model': model,
            'domain': domain,
            'message': 'این تابع نیاز به محیط اودوو دارد'
        }
    
    def get_extraction_statistics(self, model: str) -> Dict[str, Any]:
        """دریافت آمار استخراج از یک مدل"""
        # این تابع نیز نیاز به محیط اودوو دارد
        return {
            'model': model,
            'total_records': 0,
            'records_with_extraction': 0,
            'average_confidence': 0.0,
            'message': 'این تابع نیاز به محیط اودوو دارد'
        }
