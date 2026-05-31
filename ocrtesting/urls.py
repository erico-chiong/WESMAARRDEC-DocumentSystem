# urls.py

from django.urls import path, include
# --- MODIFIED: Import the new edit_memorandum view ---
from .views import (
    create_special_order, create_travel_order, create_memorandum, create_moau, create_communication_letter, create_other_document,
    view_special_order, view_travel_order, view_memorandum, view_moau,view_communication_letter, view_other_document,
    edit_special_order, edit_travel_order, edit_memorandum, edit_moau, edit_communication_letter, edit_other_document,
    delete_document, download_document_pdf, send_document_email, view_documents, export_documents_pdf, export_documents_excel, dashboard,
    dashboard_cards_data, dashboard_direction_data,
    manage_recipients, add_recipient, edit_recipient, delete_recipient, manage_archive,
    archive_documents, restore_documents, view_archived_file,
    view_logs,
)
from . import views


app_name = 'ocrtesting'

urlpatterns = [



    path('accounts/', include('accounts.urls')),
     # Dashboard URL
    path('dashboard/', dashboard, name='dashboard'),  # Ensure this line is present
    path('dashboard/cards/', dashboard_cards_data, name='dashboard_cards_data'),
    path('dashboard/direction/', dashboard_direction_data, name='dashboard_direction_data'),


    # Create URLs
    path('create/special-order/', create_special_order, name='create_special_order'),
    path('create/travel-order/', create_travel_order, name='create_travel_order'),
    path('create/communication-letter/<str:category>/', create_communication_letter, name='create_communication_letter'),
    path('create/moau/', create_moau, name='create_moau'),
    path('create/memorandum/<str:category>/', create_memorandum, name='create_memorandum'),
    path('create/other-document/', create_other_document, name='create_other_document'),


    # View URLs
    path('special-order/<str:pk>/', view_special_order, name='view_special_order'),
    path('travel-order/<str:pk>/', view_travel_order, name='view_travel_order'),
    path('memorandum/<str:pk>/', view_memorandum, name='view_memorandum'),
    path('moau/<int:pk>/', view_moau, name='view_moau'),
    path('communication-letter/<int:pk>/', view_communication_letter, name='view_communication_letter'),
    path('other-document/<int:pk>/', view_other_document, name='view_other_document'),


    # Edit URLs
    path('special-order/<str:pk>/edit/', edit_special_order, name='edit_special_order'),
    path('travel-order/<str:pk>/edit/', edit_travel_order, name='edit_travel_order'),
    path('memorandum/<str:pk>/edit/', edit_memorandum, name='edit_memorandum'),
    path('moau/<int:pk>/edit/', edit_moau, name='edit_moau'),
    path('communication-letter/<int:pk>/edit/', edit_communication_letter, name='edit_communication_letter'),
    path('other-document/<int:pk>/edit/', edit_other_document, name='edit_other_document'),


    # Other utility URLs
    
    path('scan-document/', views.unified_scan_view, name='scan_document'),
    path('documents/<str:doc_type>/<str:pk>/download/', download_document_pdf, name='download_document_pdf'),
    path('documents/<str:doc_type>/<str:pk>/send-email/', send_document_email, name='send_document_email'),
    path('documents/export/pdf/', export_documents_pdf, name='export_documents_pdf'),
    path('documents/export/excel/', export_documents_excel, name='export_documents_excel'),
    path('logs/', view_logs, name='document_logs'),




    
    # --- NEW URL FOR THE LIST VIEW ---
    path('documents/', view_documents, name='document_list'),


    # --- NEW URLS FOR RECIPIENT MANAGEMENT ---
    path('recipients/', manage_recipients, name='manage_recipients'),
    path('recipients/add/', add_recipient, name='add_recipient'),
    path('recipients/<int:pk>/edit/', edit_recipient, name='edit_recipient'),
    path('recipients/<int:pk>/delete/', delete_recipient, name='delete_recipient'),
    path('doc/<slug:doc_type>/<path:pk>/delete/', delete_document, name='delete_document'),




]