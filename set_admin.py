import os
import django
from django.core.management.base import BaseCommand
from django.contrib.auth import get_user_model

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'DocumentSystem.settings')
django.setup()

def make_user_admin(username):
    User = get_user_model()
    try:
        user = User.objects.get(username=username)
        user.admin = True
        user.is_staff = True
        user.is_superuser = True
        user.save()
        print(f"User '{username}' is now an admin.")
    except User.DoesNotExist:
        print(f"User '{username}' not found.")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        make_user_admin(sys.argv[1])
    else:
        print("Please provide a username: python set_admin.py <username>")
