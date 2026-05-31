from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from ocrtesting.views import dashboard, dashboard_cards_data, dashboard_direction_data

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', dashboard, name='dashboard'),
    path('dashboard/cards/', dashboard_cards_data, name='dashboard_cards_data'),
    path('dashboard/direction/', dashboard_direction_data, name='dashboard_direction_data'),
    path('', include('ocrtesting.urls')), 
     # points to your app-level urls
    path('users/', include('accounts.urls')),
    path('accounts/', include('allauth.urls')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)