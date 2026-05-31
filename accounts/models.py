
# Create your models here.
from django.contrib.auth.models import AbstractUser
from django.db import models
from django.conf import settings
import random, string
from django.utils import timezone

class User(AbstractUser):
    # By default username max_length is 150; if you really want varchar(30), override:
    username = models.CharField(max_length=150, unique=True)
    email = models.EmailField(unique=True)
    profile_picture = models.ImageField(upload_to='profile_pictures/', null=True, blank=True)
    researcher = models.BooleanField(default=True)
    secretariat = models.BooleanField(default=False)
    stakeholder = models.BooleanField(default=False)
    admin = models.BooleanField(default=False)  # This can be used for admin role if needed
    user = models.BooleanField(default=False)  # This can be used for user role if needed
    
    # Add a role field
    
    # `is_active`, `date_joined`, `last_login` come from AbstractUser
    # password is hashed by Django
    

    def __str__(self):
        return self.username


class VerificationCode(models.Model):
    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE)
    code = models.CharField(max_length=6)
    created_at = models.DateTimeField(auto_now_add=True)

    def is_expired(self):
        # Code expires after 24 hours
        return timezone.now() > self.created_at + timezone.timedelta(hours=24)

    @staticmethod
    def generate_code():
        return ''.join(random.choices(string.digits, k=6))

    def __str__(self):
        return f"{self.user.username} - {self.code}"