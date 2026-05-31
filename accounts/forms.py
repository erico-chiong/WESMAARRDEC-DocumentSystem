from django import forms
from django.contrib.auth.forms import UserCreationForm
from django.contrib.auth import get_user_model

User = get_user_model()

from allauth.account.forms import SignupForm

from .models import VerificationCode
from django.core.mail import send_mail
from django.conf import settings

class CustomSignupForm(SignupForm):
    first_name = forms.CharField(max_length=150, label='First Name', widget=forms.TextInput(attrs={'placeholder': 'Enter first name', 'class': 'form-control'}))
    last_name = forms.CharField(max_length=150, label='Last Name', widget=forms.TextInput(attrs={'placeholder': 'Enter last name', 'class': 'form-control'}))

    def save(self, request):
        user = super().save(request)
        user.first_name = self.cleaned_data.get('first_name')
        user.last_name = self.cleaned_data.get('last_name')
        # user.is_active is True by default
        # Roles are handled by model defaults (researcher=True)
        user.save()

        # Generate and save verification code
        code = VerificationCode.generate_code()
        VerificationCode.objects.update_or_create(
            user=user,
            defaults={'code': code}
        )

        # Send email with code
        try:
            send_mail(
                subject='Your Account Verification Code',
                message=(
                    f'Hi {user.username or user.email},\n\n'
                    f'Thank you for signing up. Use the code below to verify and activate your account:\n\n'
                    f'  Verification Code: {code}\n\n'
                    f'This code expires in 24 hours.\n\n'
                    f'If you did not expect this email, please ignore it.'
                ),
                from_email=settings.DEFAULT_FROM_EMAIL,
                recipient_list=[user.email],
                fail_silently=False,
            )
        except Exception:
            pass # Email failure handled by user not being able to verify

        return user

class CustomUserCreationForm(UserCreationForm):
    first_name = forms.CharField(max_length=150, required=True, label="First Name")
    last_name = forms.CharField(max_length=150, required=True, label="Last Name")
    email = forms.EmailField(required=True, label="Email Address")
    
    class Meta(UserCreationForm.Meta):
        model = User
        fields = UserCreationForm.Meta.fields + (
            'first_name', 'last_name', 'email', 'admin', 'secretariat', 'stakeholder', 'user', 'researcher'
        )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field_name, field in self.fields.items():
            if isinstance(field.widget, forms.CheckboxInput):
                field.widget.attrs.update({'class': 'h-4 w-4 text-indigo-600 focus:ring-indigo-500 border-gray-300 rounded'})
            else:
                field.widget.attrs.update({'class': 'mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm'})
            
            if 'password' in field_name:
                field.widget.attrs.update({'class': 'mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm'})


User = get_user_model()

class UserEditForm(forms.ModelForm):
    class Meta:
        model = User
        fields = ['first_name', 'last_name', 'email', 'admin', 'secretariat', 'stakeholder', 'user', 'researcher']
        widgets = {
            'first_name': forms.TextInput(attrs={'class': 'mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm'}),
            'last_name': forms.TextInput(attrs={'class': 'mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm'}),
            'email': forms.EmailInput(attrs={'class': 'mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm'}),
            'admin': forms.CheckboxInput(attrs={'class': 'h-4 w-4 text-indigo-600 focus:ring-indigo-500 border-gray-300 rounded'}),
            'secretariat': forms.CheckboxInput(attrs={'class': 'h-4 w-4 text-indigo-600 focus:ring-indigo-500 border-gray-300 rounded'}),
            'stakeholder': forms.CheckboxInput(attrs={'class': 'h-4 w-4 text-indigo-600 focus:ring-indigo-500 border-gray-300 rounded'}),
            'user': forms.CheckboxInput(attrs={'class': 'h-4 w-4 text-indigo-600 focus:ring-indigo-500 border-gray-300 rounded'}),
            'researcher': forms.CheckboxInput(attrs={'class': 'h-4 w-4 text-indigo-600 focus:ring-indigo-500 border-gray-300 rounded'}),
        }
