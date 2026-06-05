from django.db import models
from django.utils import timezone

class PresentationHistory(models.Model):
    name = models.CharField(max_length=255)
    file_path = models.CharField(max_length=500)
    payload = models.JSONField(default=dict)
    created_at = models.DateTimeField(default=timezone.now)
    def __str__(self):
        return self.name