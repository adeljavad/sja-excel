"""
Data models for Smart Extractor
مدل‌های داده برای سیستم استخراج هوشمند
"""

from dataclasses import dataclass
from typing import Optional, List, Dict, Any


@dataclass
class CurrencyInfo:
    """اطلاعات ارز استخراج شده"""
    amount: Optional[float] = None
    currency: Optional[str] = None
    rate: Optional[float] = None
    original_text: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """تبدیل به دیکشنری"""
        return {
            'amount': self.amount,
            'currency': self.currency,
            'rate': self.rate,
            'original_text': self.original_text
        }


@dataclass
class ExtractionResult:
    """نتیجه استخراج اطلاعات از یک متن"""
    original_text: str
    invoice_number: Optional[str] = None
    currency_info: Optional[CurrencyInfo] = None
    company_name: Optional[str] = None
    document_type: Optional[str] = None
    confidence: float = 0.0
    
    def to_dict(self) -> Dict[str, Any]:
        """تبدیل به دیکشنری"""
        return {
            'original_text': self.original_text,
            'invoice_number': self.invoice_number,
            'currency_info': self.currency_info.to_dict() if self.currency_info else None,
            'company_name': self.company_name,
            'document_type': self.document_type,
            'confidence': self.confidence
        }


@dataclass
class BatchExtractionResult:
    """نتیجه استخراج دسته‌ای"""
    results: List[ExtractionResult]
    total_records: int
    successful_extractions: int
    failed_extractions: int
    
    def to_dict(self) -> Dict[str, Any]:
        """تبدیل به دیکشنری"""
        return {
            'results': [result.to_dict() for result in self.results],
            'total_records': self.total_records,
            'successful_extractions': self.successful_extractions,
            'failed_extractions': self.failed_extractions,
            'success_rate': self.successful_extractions / self.total_records if self.total_records > 0 else 0
        }
