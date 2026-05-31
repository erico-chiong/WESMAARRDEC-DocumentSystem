from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.decorators import user_passes_test
from django.contrib.auth import get_user_model
from django.core.paginator import Paginator
from .forms import CustomUserCreationForm
from django.contrib import messages
from django.contrib.auth.models import User
from django.shortcuts import redirect, get_object_or_404
from django.contrib import messages
from .forms import UserEditForm
from .models import VerificationCode
from django.core.mail import send_mail
from django.conf import settings
from django.utils import timezone

app_name = 'accounts'
# Create your views here.
def login_user(request):
    
    #if user is already authenticated, redirect to the dashboard
    if request.user.is_authenticated:
        return redirect('ocrtesting:dashboard')
    
    #if the request method is POST, process the login form
    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            login(request, user)
            return redirect('ocrtesting:dashboard') # Redirect to the dashboard after successful login
    else:
        form = AuthenticationForm()
    return render(request, 'account/login.html', {'form': form})

def password_reset_request(request):
    if request.method == 'POST':
        email = request.POST.get('email')
        try:
            user = User.objects.get(email=email)
            code = VerificationCode.generate_code()
            VerificationCode.objects.update_or_create(
                user=user,
                defaults={'code': code, 'created_at': timezone.now()}
            )
            
            send_mail(
                subject='Password Reset Verification Code',
                message=f'Your password reset verification code is: {code}\n\nThis code expires in 24 hours.',
                from_email=settings.DEFAULT_FROM_EMAIL,
                recipient_list=[email],
                fail_silently=False,
            )
            
            request.session['reset_email'] = email
            messages.success(request, f'A verification code has been sent to {email}.')
            return redirect('accounts:password_reset_verify')
            
        except User.DoesNotExist:
            messages.error(request, 'No account found with this email address.')
            
    return render(request, 'account/password_reset.html')

def password_reset_verify(request):
    email = request.session.get('reset_email')
    if not email:
        return redirect('accounts:password_reset_request')
        
    if request.method == 'POST':
        code = request.POST.get('code')
        try:
            user = User.objects.get(email=email)
            verification = VerificationCode.objects.get(user=user, code=code)
            
            if verification.is_expired():
                messages.error(request, 'This code has expired.')
            else:
                request.session['reset_verified'] = True
                return redirect('accounts:password_reset_confirm')
                
        except (User.DoesNotExist, VerificationCode.DoesNotExist):
            messages.error(request, 'Invalid verification code.')
            
    return render(request, 'account/password_reset_verify.html', {'email': email})

def password_reset_confirm(request):
    email = request.session.get('reset_email')
    verified = request.session.get('reset_verified')
    
    if not email or not verified:
        return redirect('accounts:password_reset_request')
        
    if request.method == 'POST':
        password = request.POST.get('password')
        confirm_password = request.POST.get('confirm_password')
        
        if password != confirm_password:
            messages.error(request, 'Passwords do not match.')
        elif len(password) < 8:
            messages.error(request, 'Password must be at least 8 characters long.')
        else:
            user = User.objects.get(email=email)
            user.set_password(password)
            user.save()
            
            # Clean up session and verification code
            VerificationCode.objects.filter(user=user).delete()
            del request.session['reset_email']
            del request.session['reset_verified']
            
            messages.success(request, 'Your password has been successfully reset. You can now log in.')
            return redirect('accounts:login')
            
    return render(request, 'account/password_reset_confirm.html')

def logout_user(request):
    logout(request)
    return redirect('accounts:login')  # Redirect to the login page after logout

User = get_user_model()

# --- Helper function for decorator ---
def is_superuser(user):
    return user.is_authenticated and user.is_superuser

@user_passes_test(is_superuser, login_url='/login/')
def user_list_view(request):
    """
    Displays a paginated list of all users in the system.
    Handles searching/filtering by username or role.
    """
    user_list = User.objects.all().order_by('username')

    # Optional: Add searching functionality
    search_query = request.GET.get('q', '')
    if search_query:
        user_list = user_list.filter(username__icontains=search_query)

    # Pagination
    paginator = Paginator(user_list, 10) # Show 10 users per page
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    query_params = request.GET.copy()
    query_params.pop('page', None)
    querystring = query_params.urlencode()

    if hasattr(paginator, 'get_elided_page_range'):
        page_range = list(paginator.get_elided_page_range(page_obj.number))
    else:
        page_range = list(paginator.page_range)

    context = {
        'users_page': page_obj,
        'search_query': search_query,
        'page_range': page_range,
        'querystring': querystring,
    }
    return render(request, 'manage/accounts/user_list.html', context)


# --- View to Add a New User (for Admins) ---
@user_passes_test(lambda u: u.is_superuser, login_url='/login/')
def add_user_view(request):
    if request.method == 'POST':
        form = CustomUserCreationForm(request.POST)
        if form.is_valid():
            # Save user and activate immediately
            user = form.save(commit=False)
            user.is_active = True
            user.save()
            form.save_m2m() # Ensure any many-to-many fields are saved

            messages.success(request, f"User '{user.username}' created successfully.")
            return redirect('accounts:user_list')
        else:
            messages.error(request, "Please correct the errors below.")
    else:
        form = CustomUserCreationForm()

    return render(request, 'manage/accounts/add_user.html', {'form': form})

# --- View to Verify User Account with Code ---
def verify_code(request):
    if request.method == 'POST':
        code = request.POST.get('code')

        try:
            # Search for the verification record by code
            verification = VerificationCode.objects.get(code=code)
            user = verification.user

            if verification.is_expired():
                messages.error(request, 'This code has expired. Please contact your administrator.')
            else:
                user.is_active = True
                user.save()
                verification.delete()
                messages.success(request, 'Your account has been verified! You can now log in.')
                return redirect('accounts:login')

        except VerificationCode.DoesNotExist:
            messages.error(request, 'Invalid verification code.')

    return render(request, 'account/verify_code.html')


@user_passes_test(is_superuser, login_url='/login/')
def edit_user(request, user_id):
    user = get_object_or_404(User, pk=user_id)
    if request.method == 'POST':
        form = UserEditForm(request.POST, instance=user)
        if form.is_valid():
            form.save()
            messages.success(request, f"User '{user.username}' updated successfully.")
            return redirect('accounts:user_list')
        else:
            messages.error(request, "Please correct the errors in the form.")
    else:
        form = UserEditForm(instance=user)
    return render(request, 'manage/accounts/edit_user.html', {'form': form, 'user': user})

@user_passes_test(is_superuser, login_url='/login/')
def delete_user_view(request, pk):
    user = get_object_or_404(User, pk=pk)
    username = user.username
    user.delete()
    messages.success(request, f'User {username} has been deleted.')
    return redirect('accounts:user_list')

@user_passes_test(is_superuser, login_url='/login/')
def deactivate_user(request, user_id):
    user = get_object_or_404(User, pk=user_id)
    user.is_active = False
    user.save()
    messages.success(request, f'{user.username} has been deactivated.')
    return redirect('accounts:user_list')

@user_passes_test(is_superuser, login_url='/login/')
def activate_user(request, user_id):
    user = get_object_or_404(User, pk=user_id)
    user.is_active = True
    user.save()
    messages.success(request, f'{user.username} has been activated.')
    return redirect('accounts:user_list')
