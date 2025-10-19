from django.contrib import admin
from .models import Product
# Register your models here.

@admin.register(Product)
class ProductAdmin(admin.ModelAdmin):
    list_display = ('sku','name', 'category','price', 'quantity', 'tx_date')
    search_fields = ('sku','name')
    list_filter = ('category','tx_date')

