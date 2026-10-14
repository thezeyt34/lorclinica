from django.db import models
from django.contrib.auth.models import AbstractUser

class User(AbstractUser):
    ROLE_CHOICES = (
        ('admin', 'Администратор'),
        ('worker', 'Работник'),
    )
    role = models.CharField(max_length=10, choices=ROLE_CHOICES, default='worker')

    def is_admin(self):
        return self.role == 'admin' or self.is_superuser

    def __str__(self):
        return f"{self.username} ({self.get_role_display()})"

class Kirim(models.Model):
    client_name = models.CharField(max_length=255, verbose_name="Имя клиента")
    amount = models.DecimalField(max_digits=12, decimal_places=2, verbose_name="Сумма")
    comment = models.TextField(blank=True, null=True, verbose_name="Комментарий")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата и время")
    author = models.ForeignKey(User, on_delete=models.CASCADE, related_name='kirims', verbose_name="Автор")

    class Meta:
        ordering = ['-created_at']
        verbose_name = "Приход"
        verbose_name_plural = "Приходы"

    def __str__(self):
        return f"Приход: {self.client_name} - {self.amount} сум"

class Chiqim(models.Model):
    item_name = models.CharField(max_length=255, verbose_name="Название товара")
    quantity = models.PositiveIntegerField(verbose_name="Количество")
    price = models.DecimalField(max_digits=12, decimal_places=2, verbose_name="Цена")
    total_price = models.DecimalField(max_digits=12, decimal_places=2, verbose_name="Общая сумма", blank=True)
    comment = models.TextField(blank=True, null=True, verbose_name="Комментарий")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата и время")
    author = models.ForeignKey(User, on_delete=models.CASCADE, related_name='chiqims', verbose_name="Автор")

    class Meta:
        ordering = ['-created_at']
        verbose_name = "Расход"
        verbose_name_plural = "Расходы"

    def save(self, *args, **kwargs):
        self.total_price = self.quantity * self.price
        super().save(*args, **kwargs)

    def __str__(self):
        return f"Расход: {self.item_name} - {self.total_price} сум"
