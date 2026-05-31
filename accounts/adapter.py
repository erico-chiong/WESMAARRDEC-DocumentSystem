from allauth.account.adapter import DefaultAccountAdapter
from django.urls import reverse

class AccountAdapter(DefaultAccountAdapter):
    def get_signup_redirect_url(self, request):
        return reverse('accounts:verify_code')
