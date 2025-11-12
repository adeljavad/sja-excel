"""
Django integration adapter for Smart Extractor
آداپتور یکپارچه‌سازی با جنگو برای سیستم استخراج هوشمند
"""

from typing import List, Dict, Any, Optional
from ..core.extractors import SmartExtractor
from ..core.models import ExtractionResult, BatchExtractionResult


class DjangoIntegration:
    """کلاس یکپارچه‌سازی با جنگو"""
    
    def __init__(self):
        self.extractor = SmartExtractor()
    
    def extract_from_model_instances(self, instances: List[Any], description_field: str = 'description') -> List[Dict[str, Any]]:
        """استخراج اطلاعات از نمونه‌های مدل جنگو"""
        extracted_data = []
        
        for instance in instances:
            description = getattr(instance, description_field, '') or ''
            
            if description:
                result = self.extractor.extract_from_text(description)
                
                extracted_instance = {
                    'instance_id': instance.id,
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
                
                extracted_data.append(extracted_instance)
        
        return extracted_data
    
    def create_extraction_mixin(self) -> str:
        """ایجاد میکسین برای اضافه کردن فیلدهای استخراج شده به مدل‌های جنگو"""
        mixin_code = '''
class SmartExtractionMixin(models.Model):
    """میکسین برای اضافه کردن فیلدهای استخراج شده"""
    
    invoice_number_extracted = models.CharField(
        max_length=50,
        blank=True,
        null=True,
        verbose_name='شماره وضعیت استخراج شده'
    )
    
    currency_amount_extracted = models.FloatField(
        blank=True,
        null=True,
        verbose_name='مبلغ ارزی استخراج شده'
    )
    
    currency_type_extracted = models.CharField(
        max_length=20,
        blank=True,
        null=True,
        verbose_name='نوع ارز استخراج شده'
    )
    
    exchange_rate_extracted = models.FloatField(
        blank=True,
        null=True,
        verbose_name='نرخ ارز استخراج شده'
    )
    
    company_name_extracted = models.CharField(
        max_length=100,
        blank=True,
        null=True,
        verbose_name='نام شرکت استخراج شده'
    )
    
    document_type_extracted = models.CharField(
        max_length=50,
        blank=True,
        null=True,
        verbose_name='نوع سند استخراج شده'
    )
    
    extraction_confidence = models.FloatField(
        default=0.0,
        verbose_name='اطمینان استخراج'
    )
    
    extraction_timestamp = models.DateTimeField(
        auto_now=True,
        verbose_name='زمان استخراج'
    )
    
    class Meta:
        abstract = True
    
    def extract_from_description(self, description_field: str = 'description'):
        """استخراج اطلاعات از فیلد شرح"""
        description = getattr(self, description_field, '') or ''
        
        if description:
            from smart_extractor.core.extractors import SmartExtractor
            extractor = SmartExtractor()
            result = extractor.extract_from_text(description)
            
            self.invoice_number_extracted = result.invoice_number
            self.currency_amount_extracted = result.currency_info.amount if result.currency_info else None
            self.currency_type_extracted = result.currency_info.currency if result.currency_info else None
            self.exchange_rate_extracted = result.currency_info.rate if result.currency_info else None
            self.company_name_extracted = result.company_name
            self.document_type_extracted = result.document_type
            self.extraction_confidence = result.confidence
            
            self.save()
            
            return True
        
        return False
'''
        return mixin_code
    
    def create_management_command(self) -> str:
        """ایجاد دستور مدیریت برای استخراج دسته‌ای"""
        command_code = '''
from django.core.management.base import BaseCommand
from django.apps import apps
from smart_extractor.core.extractors import SmartExtractor


class Command(BaseCommand):
    help = 'استخراج هوشمند اطلاعات از مدل‌های جنگو'
    
    def add_arguments(self, parser):
        parser.add_argument(
            '--model',
            type=str,
            required=True,
            help='نام مدل (به فرمت app_label.ModelName)'
        )
        parser.add_argument(
            '--field',
            type=str,
            default='description',
            help='نام فیلد شرح (پیش‌فرض: description)'
        )
        parser.add_argument(
            '--limit',
            type=int,
            default=None,
            help='محدودیت تعداد رکوردها'
        )
    
    def handle(self, *args, **options):
        model_name = options['model']
        field_name = options['field']
        limit = options['limit']
        
        try:
            # دریافت مدل
            Model = apps.get_model(model_name)
            
            # دریافت رکوردها
            queryset = Model.objects.all()
            if limit:
                queryset = queryset[:limit]
            
            self.stdout.write(f'استخراج اطلاعات از {queryset.count()} رکورد...')
            
            extractor = SmartExtractor()
            updated_count = 0
            
            for instance in queryset:
                description = getattr(instance, field_name, '') or ''
                
                if description:
                    result = extractor.extract_from_text(description)
                    
                    # به‌روزرسانی فیلدها
                    instance.invoice_number_extracted = result.invoice_number
                    instance.currency_amount_extracted = result.currency_info.amount if result.currency_info else None
                    instance.currency_type_extracted = result.currency_info.currency if result.currency_info else None
                    instance.exchange_rate_extracted = result.currency_info.rate if result.currency_info else None
                    instance.company_name_extracted = result.company_name
                    instance.document_type_extracted = result.document_type
                    instance.extraction_confidence = result.confidence
                    
                    instance.save()
                    updated_count += 1
            
            self.stdout.write(
                self.style.SUCCESS(
                    f'استخراج تکمیل شد! {updated_count} رکورد به‌روزرسانی شد.'
                )
            )
            
        except Exception as e:
            self.stderr.write(
                self.style.ERROR(f'خطا در استخراج: {str(e)}')
            )
'''
        return command_code
    
    def create_admin_action(self) -> str:
        """ایجاد اکشن ادمین برای استخراج دسته‌ای"""
        action_code = '''
def extract_smart_data(modeladmin, request, queryset):
    """اکشن ادمین برای استخراج هوشمند اطلاعات"""
    from smart_extractor.core.extractors import SmartExtractor
    
    extractor = SmartExtractor()
    updated_count = 0
    
    for instance in queryset:
        description = instance.description or ''
        
        if description:
            result = extractor.extract_from_text(description)
            
            instance.invoice_number_extracted = result.invoice_number
            instance.currency_amount_extracted = result.currency_info.amount if result.currency_info else None
            instance.currency_type_extracted = result.currency_info.currency if result.currency_info else None
            instance.exchange_rate_extracted = result.currency_info.rate if result.currency_info else None
            instance.company_name_extracted = result.company_name
            instance.document_type_extracted = result.document_type
            instance.extraction_confidence = result.confidence
            
            instance.save()
            updated_count += 1
    
    messages.success(
        request,
        f'استخراج تکمیل شد! {updated_count} رکورد به‌روزرسانی شد.'
    )

extract_smart_data.short_description = 'استخراج هوشمند اطلاعات از شرح'
'''
        return action_code
