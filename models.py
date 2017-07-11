from django.db import models

# Create your models here.
class Document(models.Model):
    file = models.FileField(upload_to='documents')
    title = models.CharField(max_length=100, null=False)

    def __str__(self):
        return self.title