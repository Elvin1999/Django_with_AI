from django.db import models

# Create your models here.

class Product(models.Model):
    sku = models.CharField(max_length=64,unique=True)
    name = models.CharField(max_length=255)
    category = models.CharField(max_length=120,blank=True,null=True)

    price = models.DecimalField(max_digits=12,decimal_places=2)
    quantity = models.PositiveIntegerField(default=0)

    #satish tarixi
    tx_date=models.DateField()

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def revenue(self):
        return self.price * self.quantity

    def __str__(self):
        return f"{self.sku} - {self.name}"