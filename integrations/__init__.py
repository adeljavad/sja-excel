"""
Integration adapters for Smart Extractor
آداپتورهای یکپارچه‌سازی برای سیستم استخراج هوشمند
"""

from .odoo_integration import OdooIntegration
from .django_integration import DjangoIntegration

__all__ = ['OdooIntegration', 'DjangoIntegration']
