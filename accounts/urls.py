from django.urls import path
from . import views

app_name = 'accounts'

urlpatterns = [
    path('', views.user_list_view, name='user_list'),
    path('add/', views.add_user_view, name='add_user'),
    path('<int:user_id>/edit/', views.edit_user, name='edit_user'),
    path('<int:user_id>/deactivate/', views.deactivate_user, name='deactivate_user'),
    path('<int:user_id>/activate/', views.activate_user, name='activate_user'),
    path('login/', views.login_user, name='login'), 
    path('logout/', views.logout_user, name='logout'),
    path('password-reset/', views.password_reset_request, name='password_reset_request'),
    path('password-reset/verify/', views.password_reset_verify, name='password_reset_verify'),
    path('password-reset/confirm/', views.password_reset_confirm, name='password_reset_confirm'),
 # Make this unique
    path('user-management/<int:pk>/delete/', views.delete_user_view, name='delete_user'),
    path('verify/', views.verify_code, name='verify_code'),
]
