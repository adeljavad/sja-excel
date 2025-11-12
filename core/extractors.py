"""
Main extractor class for Smart Extractor
کلاس اصلی استخراج کننده اطلاعات
"""

from typing import List, Optional
from .models import ExtractionResult, CurrencyInfo, BatchExtractionResult
from .patterns import Patterns


class SmartExtractor:
    """کلاس اصلی استخراج کننده اطلاعات هوشمند"""
    
    def __init__(self):
        self.patterns = Patterns()
    
    def extract_from_text(self, text: str) -> ExtractionResult:
        """استخراج اطلاعات از یک متن"""
        if not text:
            return ExtractionResult(original_text=text, confidence=0.0)
        
        # استخراج شماره صورت‌وضعیت
        invoice_number, invoice_confidence = self.patterns.extract_invoice_number(text)
        
        # استخراج اطلاعات ارز
        currency_data, currency_confidence = self.patterns.extract_currency_info(text)
        currency_info = None
        if currency_data:
            currency_info = CurrencyInfo(
                amount=currency_data.get('amount'),
                currency=currency_data.get('currency'),
                rate=currency_data.get('rate'),
                original_text=currency_data.get('original_text')
            )
        
        # استخراج نام شرکت
        company_name, company_confidence = self.patterns.extract_company(text)
        
        # تشخیص نوع سند
        document_type, doc_confidence = self.patterns.detect_document_type(text)
        
        # محاسبه اطمینان کلی
        confidences = [invoice_confidence, currency_confidence, company_confidence, doc_confidence]
        non_zero_confidences = [c for c in confidences if c > 0]
        overall_confidence = sum(non_zero_confidences) / len(non_zero_confidences) if non_zero_confidences else 0.0
        
        return ExtractionResult(
            original_text=text,
            invoice_number=invoice_number,
            currency_info=currency_info,
            company_name=company_name,
            document_type=document_type,
            confidence=overall_confidence
        )
    
    def extract_batch(self, texts: List[str]) -> BatchExtractionResult:
        """استخراج اطلاعات از لیستی از متون"""
        results = []
        successful = 0
        failed = 0
        
        for text in texts:
            try:
                result = self.extract_from_text(text)
                results.append(result)
                if result.confidence > 0.5:  # آستانه موفقیت
                    successful += 1
                else:
                    failed += 1
            except Exception:
                failed += 1
                results.append(ExtractionResult(original_text=text, confidence=0.0))
        
        return BatchExtractionResult(
            results=results,
            total_records=len(texts),
            successful_extractions=successful,
            failed_extractions=failed
        )
    
    def extract_from_description_column(self, descriptions: List[str]) -> List[dict]:
        """استخراج اطلاعات از ستون شرح و تبدیل به لیست دیکشنری"""
        batch_result = self.extract_batch(descriptions)
        
        extracted_data = []
        for result in batch_result.results:
            data = {
                'original_description': result.original_text,
                'invoice_number': result.invoice_number,
                'currency_amount': result.currency_info.amount if result.currency_info else None,
                'currency_type': result.currency_info.currency if result.currency_info else None,
                'exchange_rate': result.currency_info.rate if result.currency_info else None,
                'company_name': result.company_name,
                'document_type': result.document_type,
                'extraction_confidence': result.confidence
            }
            extracted_data.append(data)
        
        return extracted_data
    
    def get_extraction_summary(self, batch_result: BatchExtractionResult) -> dict:
        """خلاصه نتایج استخراج"""
        return {
            'total_records': batch_result.total_records,
            'successful_extractions': batch_result.successful_extractions,
            'failed_extractions': batch_result.failed_extractions,
            'success_rate': batch_result.successful_extractions / batch_result.total_records if batch_result.total_records > 0 else 0,
            'average_confidence': sum(r.confidence for r in batch_result.results) / len(batch_result.results) if batch_result.results else 0
        }
