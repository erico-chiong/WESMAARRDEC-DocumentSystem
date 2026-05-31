# settings.py (Final, Corrected Version)

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = 'django-insecure-mijfrs)tj^e2fc$$30k-*+orurrnbdo0^2yn%6f%^(#38do5lw'

DEBUG = True

ALLOWED_HOSTS = ['127.0.0.1', '192.168.1.169', 'localhost']

INSTALLED_APPS = [
    'ocrtesting',
    'accounts',  # Your custom user app

    # Default Django apps
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'django.contrib.sites',
    
    'allauth',
    'allauth.account',
    'allauth.socialaccount',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    'allauth.account.middleware.AccountMiddleware',
]

AUTHENTICATION_BACKENDS = [
    'django.contrib.auth.backends.ModelBackend',
    'allauth.account.auth_backends.AuthenticationBackend',
]

SITE_ID = 1

ACCOUNT_EMAIL_VERIFICATION = 'none'
ACCOUNT_LOGIN_METHODS = {'email', 'username'}
ACCOUNT_SIGNUP_FIELDS = ['email*', 'username*', 'password1*', 'password2*', 'first_name*', 'last_name*']
ACCOUNT_FORMS = {'signup': 'accounts.forms.CustomSignupForm'}
ACCOUNT_ADAPTER = 'accounts.adapter.AccountAdapter'
ACCOUNT_AUTHENTICATE_ON_SIGNUP = True

ROOT_URLCONF = 'DocumentSystem.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [os.path.join(BASE_DIR, 'templates')],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'DocumentSystem.wsgi.application'

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
]

# This is correct and points to your custom user model.
AUTH_USER_MODEL = 'accounts.User'

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'Asia/Manila'
USE_I18N = True
USE_TZ = True

STATIC_URL = '/static/'
STATICFILES_DIRS = [ BASE_DIR / "static" ]
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'
MEDIA_URL = '/media/'
MEDIA_ROOT = os.path.join(BASE_DIR, 'images')

EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True
EMAIL_HOST_USER = 'julahmadnur3@gmail.com'
EMAIL_HOST_PASSWORD = 'wobr wnbx bxxt suao'
DEFAULT_FROM_EMAIL = 'WESMAARDEC <julahmadnur3@gmail.com>'

# --- CORRECTED REDIRECT URLS ---
# LOGIN_URL now correctly points to the login view in your 'accounts' app
LOGIN_URL = 'accounts:login'
# LOGIN_REDIRECT_URL should point to a valid URL NAME.
# Based on your app's urls.py, 'dashboard' is the correct name.
LOGIN_REDIRECT_URL = 'ocrtesting:dashboard'
# LOGOUT_REDIRECT_URL is also correct
LOGOUT_REDIRECT_URL = 'accounts:login'