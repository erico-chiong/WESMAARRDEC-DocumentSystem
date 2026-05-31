# views.py
from django.shortcuts import render, redirect, get_object_or_404
from django.utils.timezone import now
from .forms import SpecialOrderForm, TravelOrderForm, MemorandumForm, MOAUForm, CommunicationLetterForm, OtherDocumentForm, RecipientForm
from .models import DocumentFile, Recipient, SpecialOrder, TravelOrder, Memorandum, MOAU, Signatory, CommunicationLetter, OtherDocument
import os
from django.conf import settings
import json
import re
from datetime import datetime, timedelta
# Add these imports at the top
import requests
import json
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from .forms import SpecialOrderForm
import io
import pdf2image
from pdf2image import convert_from_bytes
from django.http import FileResponse, JsonResponse, HttpResponse, Http404
from django.core.mail import EmailMessage
from django.utils.text import slugify
from PIL import Image # Make sure you have Pillow installed: pip install Pillow
# --- Add all necessary imports ---
import traceback
from django.db.models import Q
from django.utils import timezone
from django.urls import reverse
from datetime import date
from itertools import chain
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import calendar
from .logs_utils import create_log
from difflib import get_close_matches
from django.urls import reverse_lazy
from django.contrib import messages
from datetime import date
import uuid
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from django.core.files.base import ContentFile













MODEL_MAP = {
    'special order': SpecialOrder,
    'travel order': TravelOrder,
    'memorandum': Memorandum,
    'moau': MOAU,
    'communication letter': CommunicationLetter,
    'other document': OtherDocument,
    # 'memorandum': Memorandum,  # Add other models here as you create them
    # 'travel order': TravelOrder,
}













def parse_special_order_memorandum_travel(raw_text):
    """
    Parses raw text to extract details for Special Orders, Memorandums,
    and Travel Orders.
    
    Args:
        raw_text (str): The OCR-extracted text from the document.

    Returns:
        dict: A dictionary containing the parsed data.
    """
    lines = raw_text.splitlines()
    result = {
        'document_id': '',
        'date': '',
        'subject': '',
        'recipient_ids': []
    }

    # Pre-fetch recipient names and create a mapping to their IDs
    all_recipients = {r.name.lower(): r.id for r in Recipient.objects.all()}
    recipient_names = list(all_recipients.keys())

    # --- REGEX PATTERNS ---
    # Matches dates like "Month Day, Year" (e.g., "July 13, 2025")
    date_pattern = re.compile(
        r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b',
        re.IGNORECASE
    )
    # Matches potential 6-digit document IDs
    id_pattern = re.compile(r'\b\d{6}\b')
    # Matches lines indicating a subject
    subject_keyword_pattern = re.compile(r'\b(subject|re)\s*:', re.IGNORECASE)

    # --- PARSING LOGIC ---
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        if not stripped_line:
            continue

        lower_line = stripped_line.lower()

        # 1. Extract Document ID (first 6-digit number found)
        if not result['document_id']:
            match = id_pattern.search(stripped_line)
            if match:
                result['document_id'] = match.group()

        # 2. Extract Date
        if not result['date']:
            match = date_pattern.search(stripped_line)
            if match:
                dt = datetime.strptime(match.group(0), "%B %d, %Y")
                result['date'] = dt.strftime("%Y-%m-%d") # Format for HTML date input

        # 3. Extract Subject
        if not result['subject'] and subject_keyword_pattern.search(lower_line):
            # Extract text after "SUBJECT:" or "RE:"
            subject_text = subject_keyword_pattern.sub('', stripped_line).strip()
            
            # Check if the next line is part of the subject (if it's all caps or starts with a capital)
            if (i + 1) < len(lines):
                next_line = lines[i + 1].strip()
                if next_line and (next_line.isupper() or next_line[0].isupper()):
                    subject_text += f" {next_line}"
            
            result['subject'] = subject_text

        # 4. Extract Recipients
        for name, recipient_id in all_recipients.items():
            if name in lower_line and recipient_id not in result['recipient_ids']:
                result['recipient_ids'].append(recipient_id)

    return result


















def parse_memorandum(raw_text):
    """
    Parses OCR text specifically for an OUTGOING Memorandum, following a strict,
    rule-based order of operations for high accuracy. It integrates fuzzy name
    matching for FOR, FROM, and THRU fields while preserving existing working logic.
    """
    # --- 1. SETUP AND PRE-COMPUTATION ---
    result = {
        'document_id': '', 'date': '', 'subject': '',
        'for_field': '', 'from_field': '', 'thru': '',
        'recipient_ids': []
    }
    
    lines = raw_text.splitlines()
    
    # Pre-fetch recipient data for efficient matching
    all_recipients_qs = Recipient.objects.all()
    recipient_names_for_matching = [r.name.upper() for r in all_recipients_qs]
    recipient_name_to_id_map = {r.name.upper(): r.id for r in all_recipients_qs}

    # Regex to find a date pattern
    date_pattern = re.compile(
        r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b',
        re.IGNORECASE
    )

    # --- 2. INTEGRATED HELPER FUNCTION FOR NAME STANDARDIZATION ---
    def extract_closest_name(raw_line_text):
        """Finds the best matching recipient name from a line of text."""
        cleaned_text = re.sub(r'\b(DR|MS|ENGR|MR)\.\s*', '', raw_line_text.strip(), flags=re.IGNORECASE).upper()
        matches = get_close_matches(cleaned_text, recipient_names_for_matching, n=1, cutoff=0.6)
        return matches[0] if matches else raw_line_text.strip()

    # --- 3. PARSING LOGIC IN STRICT ORDER ---
    subject_start_index = -1
    
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        upper_line = stripped_line.upper()

        # Document ID (Preserved Logic)
        if not result['document_id']:
            id_match = re.search(r'Memorandum\s+(?:Order\s+)?No\.\s*([\d-]+)', stripped_line, re.IGNORECASE)
            if id_match:
                result['document_id'] = id_match.group(1).strip()
        
        # Date (REUSED LOGIC from parse_special_order_memorandum_travel)
        if not result['date']:
            date_match = date_pattern.search(stripped_line)
            if date_match:
                try:
                    dt = datetime.strptime(date_match.group(0), "%B %d, %Y")
                    result['date'] = dt.strftime("%Y-%m-%d")
                except ValueError:
                    pass

        # FOR field (applies same logic as THRU)
        if upper_line.startswith("FOR") and ":" in stripped_line:
            content = stripped_line.split(":", 1)[-1].strip()
            if "OCHOTORENA" in content.upper():
                result['for_field'] = "DR. MA. CARLA A. OCHOTORENA"
            else:
                result['for_field'] = extract_closest_name(content)

        # FROM field (same logic)
        elif upper_line.startswith("FROM") and ":" in stripped_line:
            content = stripped_line.split(":", 1)[-1].strip()
            result['from_field'] = extract_closest_name(content)

        # THRU field (existing + standardized)
        elif upper_line.startswith("THRU") and ":" in stripped_line:
            content = stripped_line.split(":", 1)[-1].strip()
            result['thru'] = extract_closest_name(content)

        # SUBJECT start index
        if subject_start_index == -1 and upper_line.startswith("SUBJECT") and ":" in stripped_line:
            subject_start_index = i

    # --- STAGE 2: Subject Block (Multi-line UPPERCASE style) ---
    if subject_start_index != -1:
        subject_lines = []
        initial_content = lines[subject_start_index].split(":", 1)[-1].strip()
        if initial_content:
            subject_lines.append(initial_content)

        for i in range(subject_start_index + 1, len(lines)):
            line = lines[i].strip()
            if not line or date_pattern.search(line) or not line.isupper():
                break
            subject_lines.append(line)

        result['subject'] = ' '.join(subject_lines)

    # --- Final: Add FOR, FROM, THRU to recipient_ids ---
    for field_key in ['for_field', 'from_field', 'thru']:
        standardized = result[field_key].upper()
        if standardized in recipient_name_to_id_map:
            recipient_id = recipient_name_to_id_map[standardized]
            if recipient_id not in result['recipient_ids']:
                result['recipient_ids'].append(recipient_id)

    # --- Final Pass: Fuzzy match any remaining recipient names ---
    for line in lines:
        lower_line = line.strip().lower()
        if not lower_line:
            continue
        for name, recipient_id in {r.name.lower(): r.id for r in all_recipients_qs}.items():
            if name in lower_line and recipient_id not in result['recipient_ids']:
                result['recipient_ids'].append(recipient_id)

    return result













def parse_moau(first_page_text, last_page_text):
    """
    Parses OCR text from an MOA/U with improved accuracy for Philippine government documents.
    Follows strict patterns for subject, parties, and signatories extraction.
    """
    result = {
        'subject': '', 'date': '', 
        'first_party_agency': '', 'first_party_representative': '',
        'second_party_agency': '', 'second_party_representative': '',
        'signatory_ids': [], 'recipient_ids': []
    }
    
    # Pre-fetch data from the database
    all_signatories = {s.name.upper(): s.id for s in Signatory.objects.all()}
    all_recipients = {r.name.lower(): r.id for r in Recipient.objects.all()}
    
    # --- Phase 1: Process First Page ---
    if first_page_text:
        # Combine lines for paragraph-level processing
        full_text_first_page = '\n'.join(first_page_text.splitlines())

        # 1. Subject/Project Title Extraction
        subject_match = re.search(
            r'(Memorandum of Agreement on|Project Title[:\s*])\s*["“](.*?)["”]|(Memorandum of Agreement on|Project Title[:\s*])(.*?)(?=KNOW ALL MEN BY THESE PRESENTS)',
            full_text_first_page, 
            re.IGNORECASE | re.DOTALL
        )
        
        if subject_match:
            # Use quoted text if available, otherwise use text before termination phrase
            result['subject'] = subject_match.group(2) if subject_match.group(2) else subject_match.group(4)
            result['subject'] = result['subject'].replace('\n', ' ').strip()

        # 2. First and Second Party Identification
        party_match = re.search(
            r'entered into by and between(.*?)\s+and\s+(.*?)\s+(?:represented|;)',
            full_text_first_page, 
            re.IGNORECASE | re.DOTALL
        )
        
        if party_match:
            # First party extraction
            first_party_block = party_match.group(1).strip()
            # Find agency (first ALL CAPS line matching organization pattern)
            first_agency_match = re.search(
                r'(DEPARTMENT OF [A-Z\s]+|WESTERN MINDANAO [A-Z\s]+|PHILIPPINE COCONUT AUTHORITY[A-Z\s]*)',
                first_party_block
            )
            if first_agency_match:
                result['first_party_agency'] = first_agency_match.group(0)
            
            # Find representative (last ALL CAPS name in paragraph)
            first_names = re.findall(
                r'((?:DR\.|MS\.|ENGR\.)?\s*[A-Z][A-Z\s\.]+[A-Z])(?:\s*,[^a-z]*)?$', 
                first_party_block
            )
            if first_names:
                result['first_party_representative'] = first_names[-1].strip()

            # Second party extraction
            second_party_block = party_match.group(2).strip()
            # Find agency (first ALL CAPS line matching organization pattern)
            second_agency_match = re.search(
                r'(DEPARTMENT OF [A-Z\s]+|WESTERN MINDANAO [A-Z\s]+|PHILIPPINE COCONUT AUTHORITY[A-Z\s]*)',
                second_party_block
            )
            if second_agency_match:
                result['second_party_agency'] = second_agency_match.group(0)
            
            # Find representative (last ALL CAPS name in paragraph)
            second_names = re.findall(
                r'((?:DR\.|MS\.|ENGR\.)?\s*[A-Z][A-Z\s\.]+[A-Z])(?:\s*,[^a-z]*)?$', 
                second_party_block
            )
            if second_names:
                result['second_party_representative'] = second_names[-1].strip()

        # 3. Date Extraction
        date_match = re.search(
            r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b',
            full_text_first_page, 
            re.IGNORECASE
        )
        if date_match:
            try:
                dt = datetime.strptime(date_match.group(0), "%B %d, %Y")
                result['date'] = dt.strftime("%Y-%m-%d")
            except ValueError:
                pass

        # 4. Recipients (from first page)
        for line in first_page_text.splitlines():
            lower_line = line.strip().lower()
            for name, recipient_id in all_recipients.items():
                if name in lower_line and recipient_id not in result['recipient_ids']:
                    result['recipient_ids'].append(recipient_id)

    # --- Phase 2: Process Last Page ---
    if last_page_text:
        # 5. Signatory Extraction
        signatory_blocks = []
        current_block = []
        lines = last_page_text.splitlines()
        
        # Identify signature blocks by structure
        for i, line in enumerate(lines):
            stripped_line = line.strip()
            # Skip witness/certified sections
            if "WITNESS" in stripped_line.upper() or "CERTIFIED" in stripped_line.upper():
                continue
                
            # Signature block pattern: [ORG] -> "By:" -> [NAME] -> [TITLE]
            if stripped_line and stripped_line.isupper() and len(stripped_line) > 5:
                # Potential organization line
                if i+2 < len(lines) and lines[i+1].strip().upper() == "BY:":
                    # Found signature block structure
                    org = stripped_line
                    name_line = lines[i+2].strip()
                    title_line = lines[i+3].strip() if i+3 < len(lines) else ""
                    
                    # Validate name format
                    if re.match(r'^(?:DR\.|MS\.|ENGR\.)?\s*[A-Z][A-Z\s\.]+[A-Z]$', name_line):
                        signatory_blocks.append({
                            'org': org,
                            'name': name_line,
                            'title': title_line
                        })
        
        # Process identified signature blocks
        for block in signatory_blocks:
            name_upper = block['name'].upper()
            if name_upper in all_signatories:
                result['signatory_ids'].append(all_signatories[name_upper])
        
        # 6. Recipients (from last page)
        for line in last_page_text.splitlines():
            lower_line = line.strip().lower()
            for name, recipient_id in all_recipients.items():
                if name in lower_line and recipient_id not in result['recipient_ids']:
                    result['recipient_ids'].append(recipient_id)

    return result




# --- COMPLETE AND CORRECTED UNIFIED SCAN VIEW ---
# In your views.py file:

# --- THIS IS THE COMPLETE, CORRECTED UNIFIED SCAN VIEW ---
# --- THIS IS THE COMPLETE, CORRECTED UNIFIED SCAN VIEW ---
@csrf_exempt
@require_POST
def unified_scan_view(request):
    """
    Handles file upload, OCR, and calls the correct parser based on doc_type.
    This version correctly handles file uploads from BOTH the create page ('files')
    and the edit page ('file'), preserving all original logic.
    """
    try:
        # --- THIS IS THE ROBUST FIX ---
        # First, try to get a list of files (for the create pages).
        uploaded_files = request.FILES.getlist('files')
        
        # If the list is empty, it might be a single file upload from the edit page.
        if not uploaded_files:
            single_file = request.FILES.get('file')
            if single_file:
                # If we found a single file, put it into a list to be processed uniformly.
                uploaded_files = [single_file]

        # Now, check if we have any files at all.
        if not uploaded_files:
            return JsonResponse({'error': 'No file(s) were uploaded.'}, status=400)
        
        # --- THE REST OF THE LOGIC IS PRESERVED ---
        doc_type = request.POST.get('doc_type')
        category = request.POST.get('category')
        api_key = 'K84532793088957'
        ocr_pages_text = []
        
        first_file = uploaded_files[0]
        is_pdf = first_file.content_type == 'application/pdf'

        # PDF Processing Branch (Unchanged)
        if is_pdf:
            if len(uploaded_files) > 1:
                return JsonResponse({'error': 'Please upload only one PDF file at a time.'}, status=400)
            
            pdf_bytes = first_file.read()
            poppler_path = r"C:\Users\julah\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"
            pil_images = convert_from_bytes(pdf_bytes, poppler_path=poppler_path)
            
            pages_to_ocr = []
            if doc_type == 'moau' and len(pil_images) > 1:
                pages_to_ocr.append(pil_images[0])
                pages_to_ocr.append(pil_images[-1])
            elif pil_images:
                pages_to_ocr.append(pil_images[0])

            for image in pages_to_ocr:
                with io.BytesIO() as output:
                    image.save(output, format="JPEG")
                    ocr_response = requests.post('https://api.ocr.space/parse/image', files={'file': ('image.jpg', output.getvalue(), 'image/jpeg')}, data={'isOverlayRequired': False, 'apikey': api_key, 'language': 'eng', 'scale': True, 'OCREngine': 2})
                    ocr_response.raise_for_status()
                    ocr_result = ocr_response.json()
                    if ocr_result.get('IsErroredOnProcessing'): return JsonResponse({'error': f"OCR Failed: {ocr_result.get('ErrorMessage', ['Unknown'])[0]}"}, status=400)
                    if ocr_result.get('ParsedResults'): ocr_pages_text.append(ocr_result['ParsedResults'][0]['ParsedText'])

        # Image Processing Branch (Unchanged)
        else:
            files_to_ocr = []
            if doc_type == 'moau' and len(uploaded_files) > 1:
                files_to_ocr.append(uploaded_files[0])
                files_to_ocr.append(uploaded_files[-1])
            else:
                files_to_ocr.append(uploaded_files[0])

            for file_obj in files_to_ocr:
                ocr_response = requests.post('https://api.ocr.space/parse/image', data={'isOverlayRequired': False, 'apikey': api_key, 'language': 'eng', 'scale': True, 'OCREngine': 2}, files={'file': (file_obj.name, file_obj.read(), file_obj.content_type)})
                ocr_response.raise_for_status()
                ocr_result = ocr_response.json()
                if ocr_result.get('IsErroredOnProcessing'): return JsonResponse({'error': f"OCR Failed: {ocr_result.get('ErrorMessage', ['Unknown'])[0]}"}, status=400)
                if ocr_result.get('ParsedResults'): ocr_pages_text.append(ocr_result['ParsedResults'][0]['ParsedText'])

        if not ocr_pages_text:
             return JsonResponse({'error': 'OCR could not extract any text from the document.'}, status=400)

        # PARSING DISPATCHER (Unchanged and Correct)
        parsed_data = {}
        first_page_text = ocr_pages_text[0] if ocr_pages_text else ""
        
        if doc_type in ['special_order', 'travel_order']:
            parsed_data = parse_special_order_memorandum_travel(first_page_text)
        elif doc_type == 'memorandum':
            if category == 'incoming':
                parsed_data = parse_special_order_memorandum_travel(first_page_text)
            elif category == 'outgoing':
                parsed_data = parse_memorandum(first_page_text)
            else:
                return JsonResponse({'error': 'Memorandum category (incoming/outgoing) is required for parsing.'}, status=400)
        elif doc_type == 'moau':
            last_page_text = ocr_pages_text[-1] if len(ocr_pages_text) > 1 else first_page_text
            parsed_data = parse_moau(first_page_text, last_page_text)
        elif doc_type == 'communication_letter':
            parsed_data = {}
        else:
            return JsonResponse({'error': f'Unsupported document type for parsing: {doc_type}'}, status=400)

        return JsonResponse(parsed_data)

    except requests.exceptions.RequestException as e:
        return JsonResponse({'error': f'Network error communicating with OCR service: {e}'}, status=500)
    except Exception as e:
        traceback.print_exc()
        return JsonResponse({'error': f'An unexpected server error occurred: {str(e)}'}, status=500)










# --- THIS IS THE FINAL, WORKING DELETE DOCUMENT FUNCTION ---

def delete_document(request, doc_type, pk):
    # Map URL slugs to model classes
    DOC_MODELS = {
        'memorandum':        Memorandum,
        'communication-letter': CommunicationLetter,
        'special-order':     SpecialOrder,
        'travel-order':      TravelOrder,
        'moau':              MOAU,
        'other-document':    OtherDocument,
    }

    Model = DOC_MODELS.get(doc_type)
    if not Model:
        raise Http404("Unknown document type")

    # Lookup by primary key
    obj = get_object_or_404(Model, pk=pk)

    if request.method == 'POST':
        # Delete associated files from storage
        for f in obj.documentfile_set.all():
            if f.file:
                f.file.delete(save=False)
        # Delete the record (cascades to DocumentFile in DB)
        obj.delete()
        return redirect('document_list')

    # GET → show confirmation page
    return render(request, 'confirm_delete.html', {
        'object':    obj,
        'doc_type':  doc_type,
    })



def edit_special_order(request, pk):
    so = get_object_or_404(SpecialOrder, document_id=pk)

    if request.method == 'POST':
        form = SpecialOrderForm(request.POST, request.FILES, instance=so)
        if form.is_valid():
            so = form.save(commit=False)
            so.edit_date = now()
            so.save()
            so.recipients.clear()
            form.save_m2m()

            if request.POST.get('clear_files') == 'true':
                for f in so.documentfile_set.all():
                    if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                        os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                    f.delete()

            for upload in request.FILES.getlist('files'):
                DocumentFile.objects.create(file=upload, content_type=upload.content_type, special_order=so)
            
            return redirect('view_special_order', pk=so.pk)
    else:
        form = SpecialOrderForm(instance=so)

    existing_recipient_ids = list(so.recipients.values_list('id', flat=True))
    existing_files = [{'url': f.file.url, 'name': os.path.basename(f.file.name), 'content_type': f.content_type} for f in so.documentfile_set.all()]
    
    # MODIFIED: Render the generic template
    return render(request, 'edit_order.html', {
        'form': form,
        'document': so, # Pass the instance itself for the URL
        'document_type': 'special order', # Specify the document type
        'recipient_list': Recipient.objects.all(),
        'existing_recipient_ids': existing_recipient_ids,
        'existing_files_json': json.dumps(existing_files),
    })





def view_special_order(request, pk):
    """
    Displays the details for a specific Special Order.
    """
    special_order = get_object_or_404(SpecialOrder, pk=pk)
    files = DocumentFile.objects.filter(special_order=special_order).order_by('order')
    
    # We now render a generic template, passing the SO instance as 'document'
    return render(request, 'view_order.html', {
        'document':      special_order,
        'files':         files,
        'document_type': 'special order',
    })






def create_travel_order(request):
    if request.method == 'POST':
        form = TravelOrderForm(request.POST, request.FILES) # Use TravelOrderForm
        if form.is_valid():
            to = form.save(commit=False) # Changed variable to 'to' for clarity
            to.fiscal_year = now().year
            to.save()
            form.save_m2m()  # Save recipients

            # Handle file uploads and link them to the TravelOrder instance
            for idx, upload in enumerate(request.FILES.getlist('files')):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    travel_order=to,  # CRITICAL: Link to travel_order, not special_order
                    order=idx + 1
                )

            # Redirect to a 'view_travel_order' URL (we will create this later)
            return redirect('view_travel_order', pk=to.pk)
    else:
        form = TravelOrderForm() # Use TravelOrderForm

    recipient_list = Recipient.objects.all() 

    # We reuse the exact same template, just pass a different 'document_type'
    return render(request, 'create_order.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'travel order', # This tells the template to display "Travel Order"
    })




def view_travel_order(request, pk):
    """
    Displays the details for a specific Travel Order.
    """
    travel_order = get_object_or_404(TravelOrder, pk=pk)
    files = DocumentFile.objects.filter(travel_order=travel_order).order_by('order')

    # This view ALSO renders the same generic template, passing the TO instance as 'document'
    return render(request, 'view_order.html', {
        'document':      travel_order,
        'files':         files,
        'document_type': 'travel order',
    })







def edit_travel_order(request, pk):
    to = get_object_or_404(TravelOrder, document_id=pk)

    if request.method == 'POST':
        # Use TravelOrderForm
        form = TravelOrderForm(request.POST, request.FILES, instance=to)
        if form.is_valid():
            to = form.save(commit=False)
            to.edit_date = now()
            to.save()
            to.recipients.clear()
            form.save_m2m()

            if request.POST.get('clear_files') == 'true':
                for f in to.documentfile_set.all():
                    if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                        os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                    f.delete()

            for upload in request.FILES.getlist('files'):
                # CRITICAL: Link new files to the travel_order instance
                DocumentFile.objects.create(file=upload, content_type=upload.content_type, travel_order=to)
            
            # Redirect to the correct view
            return redirect('view_travel_order', pk=to.pk)
    else:
        # Use TravelOrderForm
        form = TravelOrderForm(instance=to)

    existing_recipient_ids = list(to.recipients.values_list('id', flat=True))
    existing_files = [{'url': f.file.url, 'name': os.path.basename(f.file.name), 'content_type': f.content_type} for f in to.documentfile_set.all()]

    # Use the SAME generic template
    return render(request, 'edit_order.html', {
        'form': form,
        'document': to, # Pass the instance
        'document_type': 'travel order', # Specify the type
        'recipient_list': Recipient.objects.all(),
        'existing_recipient_ids': existing_recipient_ids,
        'existing_files_json': json.dumps(existing_files),
    })






def create_memorandum(request, category):
    if category not in ['incoming', 'outgoing']:
        # Handle invalid category, e.g., show an error page
        return redirect('some_error_page_or_dashboard')

    if request.method == 'POST':
        form = MemorandumForm(request.POST, request.FILES)
        if form.is_valid():
            memo = form.save(commit=False)
            memo.fiscal_year = now().year
            # The category is already set from the hidden form field
            memo.save()
            form.save_m2m()

            for idx, upload in enumerate(request.FILES.getlist('files')):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    memorandum=memo,
                    order=idx + 1
                )
            # You will need to create 'view_memorandum' later
            return redirect('view_memorandum', pk=memo.pk)
    else:
        # Pre-fill the form with the category from the URL
        form = MemorandumForm(initial={'category': category})

    recipient_list = Recipient.objects.all()

    return render(request, 'create_memorandum.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'memorandum',
        'category': category, # Pass category for JS and hidden fields
    })



def view_memorandum(request, pk):
    """
    Displays the details for a specific Memorandum.
    """
    memorandum = get_object_or_404(Memorandum, pk=pk)
    # Query files linked to the memorandum instance
    files = DocumentFile.objects.filter(memorandum=memorandum).order_by('order')

    # Render the same generic template with memorandum-specific context
    return render(request, 'view_order.html', {
        'document':      memorandum,
        'files':         files,
        'document_type': 'memorandum',
    })



# --- NEW, DEDICATED EDIT MEMORANDUM VIEW ---
def edit_memorandum(request, pk):
    """
    Handles editing a Memorandum document.
    """
    memo = get_object_or_404(Memorandum, pk=pk)

    if request.method == 'POST':
        form = MemorandumForm(request.POST, request.FILES, instance=memo)
        if form.is_valid():
            memo = form.save(commit=False)
            memo.edit_date = now()
            memo.save()
            memo.recipients.clear()
            form.save_m2m()

            # --- PRESERVED FILE DELETION LOGIC ---
            if request.POST.get('clear_files') == 'true':
                for f in memo.documentfile_set.all():
                    if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                        os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                    f.delete()

            # --- PRESERVED FILE ADDITION LOGIC ---
            for upload in request.FILES.getlist('files'):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    memorandum=memo  # Link to the memorandum instance
                )
            
            return redirect('view_memorandum', pk=memo.pk)
    else:
        form = MemorandumForm(instance=memo)

    existing_recipient_ids = list(memo.recipients.values_list('id', flat=True))
    existing_files = [{'url': f.file.url, 'name': os.path.basename(f.file.name), 'content_type': f.content_type} for f in memo.documentfile_set.all()]
    
    # Render the new, dedicated edit_memorandum.html template
    return render(request, 'edit_memorandum.html', {
        'form': form,
        'document': memo, # Pass the instance itself for context
        'recipient_list': Recipient.objects.all(),
        'existing_recipient_ids': existing_recipient_ids,
        'existing_files_json': json.dumps(existing_files),
    })











def create_moau(request):
    """
    Handles the creation of a new MOA/U document, including creating new
    signatories on-the-fly from a tags input.
    """
    if request.method == 'POST':
        form = MOAUForm(request.POST, request.FILES)
        if form.is_valid():
            # Save the main MOAU object, but don't handle m2m fields yet.
            moau = form.save(commit=False)
            moau.fiscal_year = now().year
            moau.save() 
            
            # Manually save recipients since we removed the default m2m save
            recipients = form.cleaned_data.get('recipients', [])
            moau.recipients.set(recipients)

            # --- NEW LOGIC: Process the string of signatory names ---
            signatories_str = form.cleaned_data.get('signatories_str', '')
            if signatories_str:
                # Split the string by commas and strip whitespace from each name
                signatory_names = [name.strip() for name in signatories_str.split(',') if name.strip()]
                
                signatory_objects = []
                for name in signatory_names:
                    # For each name, get the Signatory object if it exists,
                    # or create a new one if it does not.
                    signatory, created = Signatory.objects.get_or_create(name=name)
                    signatory_objects.append(signatory)
                
                # Associate the final list of signatory objects with the MOAU
                moau.signatories.set(signatory_objects)

            # --- PRESERVED LOGIC: Handle file uploads ---
            for idx, upload in enumerate(request.FILES.getlist('files')):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    moau=moau,
                    order=idx + 1
                )
            
            # Redirect to the view page upon success
            return redirect('view_moau', pk=moau.pk)
    else:
        form = MOAUForm()

    recipient_list = Recipient.objects.all()
    # We no longer need to pass the signatory_list as it's a text input now.

    return render(request, 'create_moau.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'moau',
    })



def view_moau(request, pk):
    """
    Displays the details for a specific MOA/U document.
    """
    moau = get_object_or_404(MOAU, pk=pk)
    files = DocumentFile.objects.filter(moau=moau).order_by('order')
    
    # Renders the new, dedicated view_moau.html template
    return render(request, 'view_moau.html', {
        'document':      moau,
        'files':         files,
        'document_type': 'moau', # Pass doc_type for consistency in includes/links
    })



# --- NEW, DEDICATED EDIT MOAU VIEW ---
def edit_moau(request, pk):
    """
    Handles editing an MOA/U document, including its typable signatories.
    """
    moau = get_object_or_404(MOAU, pk=pk)

    if request.method == 'POST':
        form = MOAUForm(request.POST, request.FILES, instance=moau)
        if form.is_valid():
            moau = form.save(commit=False)
            moau.edit_date = now()
            moau.save() # Save the main MOAU object first

            # Manually handle recipients
            moau.recipients.clear()
            recipients = form.cleaned_data.get('recipients', [])
            moau.recipients.set(recipients)

            # Manually handle signatories from the text input
            moau.signatories.clear() # Clear existing signatories before setting new ones
            signatories_str = form.cleaned_data.get('signatories_str', '')
            if signatories_str:
                signatory_names = [name.strip() for name in signatories_str.split(',') if name.strip()]
                signatory_objects = []
                for name in signatory_names:
                    signatory, created = Signatory.objects.get_or_create(name=name)
                    signatory_objects.append(signatory)
                moau.signatories.set(signatory_objects)

            # --- PRESERVED FILE DELETION LOGIC ---
            if request.POST.get('clear_files') == 'true':
                for f in moau.documentfile_set.all():
                    if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                        os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                    f.delete()
            
            # --- PRESERVED FILE ADDITION LOGIC ---
            for upload in request.FILES.getlist('files'):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    moau=moau # Link to the moau instance
                )
            
            return redirect('view_moau', pk=moau.pk)
    else:
        # For a GET request, pre-fill the signatories_str field for the tags input
        initial_signatories = moau.signatories.all()
        initial_signatories_str = ",".join([s.name for s in initial_signatories])
        form = MOAUForm(instance=moau, initial={'signatories_str': initial_signatories_str})

    # Prepare context for the template
    existing_recipient_ids = list(moau.recipients.values_list('id', flat=True))
    existing_files = [{'url': f.file.url, 'name': os.path.basename(f.file.name), 'content_type': f.content_type} for f in moau.documentfile_set.all()]
    
    # Render the new, dedicated edit_moau.html template
    return render(request, 'edit_moau.html', {
        'form': form,
        'document': moau,
        'recipient_list': Recipient.objects.all(),
        'existing_recipient_ids': existing_recipient_ids,
        'existing_files_json': json.dumps(existing_files),
    })











def create_communication_letter(request, category):
    """
    Handles the creation of a new Communication Letter document.
    """
    if category not in ['incoming', 'outgoing']:
        # Handle invalid category, e.g., show an error page
        return redirect('some_error_page_or_dashboard')

    if request.method == 'POST':
        form = CommunicationLetterForm(request.POST, request.FILES)
        if form.is_valid():
            comm_letter = form.save(commit=False)
            comm_letter.fiscal_year = now().year
            # The category is already set from the hidden form field
            comm_letter.save()
            form.save_m2m() # Saves recipients

            for idx, upload in enumerate(request.FILES.getlist('files')):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    communication_letter=comm_letter,
                    order=idx + 1
                )
            # You will need to create 'view_communication_letter' later
            return redirect('view_communication_letter', pk=comm_letter.pk) # Placeholder redirect
    else:
        # Pre-fill the form with the category from the URL
        form = CommunicationLetterForm(initial={'category': category})

    recipient_list = Recipient.objects.all()

    # Render the new, dedicated Communication Letter template
    return render(request, 'create_communication_letter.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'communication letter',
        'category': category, # Pass category for JS and hidden fields
    })




def view_communication_letter(request, pk):
    """
    Displays the details for a specific Communication Letter document.
    """
    comm_letter = get_object_or_404(CommunicationLetter, pk=pk)
    files = DocumentFile.objects.filter(communication_letter=comm_letter).order_by('order')
    
    # Renders the new, dedicated view_communication_letter.html template
    return render(request, 'view_communication_letter.html', {
        'document':      comm_letter,
        'files':         files,
        'document_type': 'communication letter', # Pass doc_type for consistency
    })






def edit_communication_letter(request, pk):
    """
    Handles editing a Communication Letter document.
    """
    comm_letter = get_object_or_404(CommunicationLetter, pk=pk)

    if request.method == 'POST':
        form = CommunicationLetterForm(request.POST, request.FILES, instance=comm_letter)
        if form.is_valid():
            comm_letter = form.save(commit=False)
            comm_letter.edit_date = now()
            comm_letter.save()

            # Manually save recipients
            comm_letter.recipients.clear()
            form.save_m2m()

            # --- PRESERVED FILE DELETION LOGIC ---
            if request.POST.get('clear_files') == 'true':
                for f in comm_letter.documentfile_set.all():
                    if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                        os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                    f.delete()

            # --- PRESERVED FILE ADDITION LOGIC ---
            for upload in request.FILES.getlist('files'):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    communication_letter=comm_letter  # Link to the correct instance
                )
            
            return redirect('view_communication_letter', pk=comm_letter.pk)
    else:
        form = CommunicationLetterForm(instance=comm_letter)

    existing_recipient_ids = list(comm_letter.recipients.values_list('id', flat=True))
    existing_files = [{'url': f.file.url, 'name': os.path.basename(f.file.name), 'content_type': f.content_type} for f in comm_letter.documentfile_set.all()]
    
    # Render the new, dedicated edit_communication_letter.html template
    return render(request, 'edit_communication_letter.html', {
        'form': form,
        'document': comm_letter,
        'recipient_list': Recipient.objects.all(),
        'existing_recipient_ids': existing_recipient_ids,
        'existing_files_json': json.dumps(existing_files),
    })







def create_other_document(request):
    """
    Handles the creation of a new Other Document.
    This view does not have a scanning feature.
    """
    if request.method == 'POST':
        form = OtherDocumentForm(request.POST, request.FILES)
        if form.is_valid():
            other_doc = form.save(commit=False)
            other_doc.fiscal_year = now().year
            other_doc.save()
            form.save_m2m() # Saves recipients

            for idx, upload in enumerate(request.FILES.getlist('files')):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    other_document=other_doc, # Link to the other_document instance
                    order=idx + 1
                )
            # You will need to create 'view_other_document' later
            return redirect('view_other_document', pk=other_doc.pk) # Placeholder redirect
    else:
        form = OtherDocumentForm()

    recipient_list = Recipient.objects.all()

    # Render the new, dedicated Other Document template
    return render(request, 'create_other_document.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'other document', # For consistency
    })



def view_other_document(request, pk):
    """
    Displays the details for a specific Other Document.
    """
    other_doc = get_object_or_404(OtherDocument, pk=pk)
    files = DocumentFile.objects.filter(other_document=other_doc).order_by('order')
    
    # Renders the new, dedicated view_other_document.html template
    return render(request, 'view_other_document.html', {
        'document':      other_doc,
        'files':         files,
        'document_type': 'other document', # Pass doc_type for consistency
    })


def edit_other_document(request, pk):
    """
    Handles editing an Other Document.
    This view does not have a scanning feature.
    """
    other_doc = get_object_or_404(OtherDocument, pk=pk)

    if request.method == 'POST':
        form = OtherDocumentForm(request.POST, request.FILES, instance=other_doc)
        if form.is_valid():
            other_doc = form.save(commit=False)
            other_doc.edit_date = now()
            other_doc.save()

            # Manually save recipients
            other_doc.recipients.clear()
            form.save_m2m()

            # --- PRESERVED FILE DELETION LOGIC ---
            if request.POST.get('clear_files') == 'true':
                for f in other_doc.documentfile_set.all():
                    if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                        os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                    f.delete()

            # --- PRESERVED FILE ADDITION LOGIC ---
            for upload in request.FILES.getlist('files'):
                DocumentFile.objects.create(
                    file=upload,
                    content_type=upload.content_type,
                    other_document=other_doc  # Link to the correct instance
                )
            
            return redirect('view_other_document', pk=other_doc.pk)
    else:
        form = OtherDocumentForm(instance=other_doc)

    existing_recipient_ids = list(other_doc.recipients.values_list('id', flat=True))
    existing_files = [{'url': f.file.url, 'name': os.path.basename(f.file.name), 'content_type': f.content_type} for f in other_doc.documentfile_set.all()]
    
    # Render the new, dedicated edit_other_document.html template
    return render(request, 'edit_other_document.html', {
        'form': form,
        'document': other_doc,
        'recipient_list': Recipient.objects.all(),
        'existing_recipient_ids': existing_recipient_ids,
        'existing_files_json': json.dumps(existing_files),
    })



















# --- REFACTORED HELPER FUNCTION TO GET DATA ---
def _get_filtered_documents(request):
    """
    Contains all the logic to fetch and filter documents from multiple models.
    MODIFIED: This version now filters out archived documents, ensuring only
    "live" documents are shown on the main document list.
    """
    today = timezone.localdate()
    params = request.GET

    # 1. Get Filter Parameters (Unchanged)
    doc_type = params.get('doc_type', '')
    category = params.get('category', '')
    recipient_id = params.get('recipient', '')
    search_query = params.get('search', '')
    start_date_str = params.get('start_date', '')
    end_date_str = params.get('end_date', '')

    # 2. Date Filtering Logic (Unchanged)
    start_date = None
    end_date = None
    if start_date_str and end_date_str:
        start_date = date.fromisoformat(start_date_str)
        end_date = date.fromisoformat(end_date_str)
    elif not any([doc_type, category, recipient_id, search_query]):
        start_date = today.replace(day=1)
        _, last_day = calendar.monthrange(today.year, today.month)
        end_date = today.replace(day=last_day)
    
    # 3. Query and Combine Documents
    models_to_query = {
        'special_order': SpecialOrder, 'travel_order': TravelOrder, 'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }
    models_to_fetch = {doc_type: models_to_query[doc_type]} if doc_type and doc_type in models_to_query else models_to_query

    querysets = []
    for type_slug, model_class in models_to_fetch.items():
        # --- THIS IS THE ONLY CHANGE IN THIS FUNCTION ---
        # We add .filter(is_archived=False) to every query to get only live documents.
        qs = model_class.objects.prefetch_related('recipients').filter(is_archived=False)

        # Apply other filters (All this logic is preserved and unchanged)
        if start_date: qs = qs.filter(date__gte=start_date)
        if end_date: qs = qs.filter(date__lte=end_date)
        if recipient_id: qs = qs.filter(recipients__id=recipient_id)
        if category and hasattr(model_class, 'category'): qs = qs.filter(category=category)
        if search_query:
            search_fields = Q(subject__icontains=search_query)
            if hasattr(model_class, 'document_id'): search_fields |= Q(document_id__icontains=search_query)
            if hasattr(model_class, 'remarks'): search_fields |= Q(remarks__icontains=search_query)
            qs = qs.filter(search_fields)

        # URL Generation Logic (All this logic is preserved and unchanged)
        for doc in qs:
            pk_as_string = str(doc.pk)
            doc.doc_type_slug = type_slug
            doc.doc_type_display = type_slug.replace('_', ' ').title()
            
            view_kwargs = {'pk': pk_as_string}
            edit_kwargs = {'pk': pk_as_string}
            
            if type_slug in ['memorandum', 'communication_letter']:
                pass 
            
            doc.view_url = reverse(f'view_{type_slug}', kwargs=view_kwargs)
            doc.edit_url = reverse(f'edit_{type_slug}', kwargs=edit_kwargs)
            delete_slug = type_slug.replace('_', '-')
            doc.delete_url = reverse('delete_document', kwargs={'doc_type': delete_slug, 'pk': pk_as_string})
            
            recipient_names = [r.name for r in doc.recipients.all()]
            doc.recipient_list_display = ", ".join(recipient_names) if recipient_names else "None"
            
            if doc.date:
                querysets.append(doc)

    # 4. Sort and return (Unchanged)
    sorted_documents = sorted(querysets, key=lambda x: x.date, reverse=True)
    page_title = _generate_document_list_title(params, today)
    
    return sorted_documents, page_title






# This is a helper function to generate the dynamic page title.
def _generate_document_list_title(params, today):
    """Generates a dynamic title based on the applied filters."""
    doc_type = params.get('doc_type')
    category = params.get('category')
    start_date_str = params.get('start_date')
    end_date_str = params.get('end_date')
    search = params.get('search')

    if search:
        return f"Search results for: '{search}'"

    title_parts = []
    
    # Category first
    if category:
        title_parts.append(category.title())

    # Then Document Type
    if doc_type:
        # Replace underscore with space and capitalize
        type_name = doc_type.replace('_', ' ').title()
        # Handle pluralization simply
        title_parts.append(f"{type_name}s")
    else:
        title_parts.append("Documents")
        
    # Then Date Range
    if start_date_str and end_date_str:
        start_date = date.fromisoformat(start_date_str)
        end_date = date.fromisoformat(end_date_str)
        if start_date == end_date:
            title_parts.append(f"on {start_date.strftime('%B %d, %Y')}")
        else:
            title_parts.append(f"from {start_date.strftime('%b %d, %Y')} to {end_date.strftime('%b %d, %Y')}")
    elif start_date_str:
         start_date = date.fromisoformat(start_date_str)
         title_parts.append(f"since {start_date.strftime('%B %d, %Y')}")
    elif end_date_str:
        end_date = date.fromisoformat(end_date_str)
        title_parts.append(f"until {end_date.strftime('%B %d, %Y')}")
    else:
        # Default to current month if no date is specified
        title_parts.append(f"for {today.strftime('%B %Y')}")

    return " ".join(title_parts)


# --- UPDATED MAIN VIEW ---
def view_documents(request):
    """
    Displays the document list page. Now uses the helper to get data.
    """
    sorted_documents, page_title = _get_filtered_documents(request)
    
    context = {
        'documents': sorted_documents,
        'page_title': page_title,
        'all_recipients': Recipient.objects.all().order_by('name'),
        'filter_values': request.GET, # Pass the entire GET dictionary
    }
    
    return render(request, 'document_list.html', context)

















def export_documents_pdf(request):
    """
    Exports the filtered document list as a PDF file, forcing a 'Save As' dialog.
    The filename is dynamically generated from the report's title.
    """
    documents, page_title = _get_filtered_documents(request)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=0.5*inch, leftMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
    
    elements = []
    styles = getSampleStyleSheet()
    styles['h1'].alignment = 1 # Center align the title
    elements.append(Paragraph(page_title, styles['h1']))
    elements.append(Paragraph("<br/><br/>", styles['Normal']))
    
    data = [ ["Type", "Subject", "Doc ID", "Recipients", "Doc Date"] ]
    for d in documents:
        doc_id_text = d.document_id if hasattr(d, 'document_id') and d.document_id else 'N/A'
        data.append([
            d.doc_type_display,
            Paragraph(d.subject or '(No Subject)', styles['Normal']),
            doc_id_text,
            Paragraph(d.recipient_list_display, styles['Normal']),
            d.date.strftime('%b. %d, %Y')
        ])

    table = Table(data, colWidths=[1.5*inch, 3*inch, 1*inch, 3.5*inch, 1*inch])
    style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4A5568')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 10),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F7FAFC')),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('FONTSIZE', (0,1), (-1,-1), 9),
    ])
    table.setStyle(style)
    
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    # --- THIS IS THE KEY CHANGE FOR PDF ---
    # 1. Create a clean filename from the page title.
    filename = f"{slugify(page_title)}.pdf"
    
    # 2. Create the HttpResponse with the correct content type.
    response = HttpResponse(buffer, content_type='application/pdf')
    
    # 3. Set the Content-Disposition header to force download.
    response['Content-Disposition'] = f'attachment; filename="{filename}"'

    return response


# --- THIS IS THE UPDATED EXCEL EXPORT FUNCTION ---

def export_documents_excel(request):
    """
    Exports the filtered document list as an XLSX Excel file.
    The filename is dynamically generated from the report's title.
    """
    documents, page_title = _get_filtered_documents(request)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Document Report"

    sheet.merge_cells('A1:E1') # Merge across all 5 columns
    title_cell = sheet['A1']
    title_cell.value = page_title
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal='center')

    headers = ["Type", "Subject", "Document ID", "Recipients", "Document Date"]
    sheet.append(headers)
    for cell in sheet[2]: # Headers are now on row 2
        cell.font = Font(bold=True)

    for doc in documents:
        sheet.append([
            doc.doc_type_display,
            doc.subject or '',
            doc.document_id if hasattr(doc, 'document_id') else '',
            doc.recipient_list_display,
            doc.date,
        ])
    
    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['B'].width = 50
    sheet.column_dimensions['C'].width = 15
    sheet.column_dimensions['D'].width = 40
    sheet.column_dimensions['E'].width = 15

    # --- THIS IS THE KEY CHANGE FOR EXCEL ---
    # 1. Create a clean filename from the page title.
    filename = f"{slugify(page_title)}.xlsx"

    # 2. Create the HttpResponse.
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    # 3. Set the Content-Disposition header with the new filename.
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    
    workbook.save(response)

    return response















def generate_document_pdf(document_instance):
    """
    Generates a single PDF in memory from all image files of ANY document instance.
    
    Args:
        document_instance: An instance of any document model (e.g., SpecialOrder).

    Returns:
        io.BytesIO: A byte stream of the generated PDF, or None if no images.
    """
    # This works generically because of Django's reverse relationship manager ('documentfile_set')
    image_files = document_instance.documentfile_set.filter(content_type__startswith='image')
    if not image_files:
        return None

    pil_images = []
    for doc_file in image_files:
        try:
            img = Image.open(doc_file.file.path).convert('RGB')
            pil_images.append(img)
        except Exception:
            continue

    if not pil_images:
        return None

    pdf_stream = io.BytesIO()
    pil_images[0].save(
        pdf_stream, format='PDF', resolution=100.0, save_all=True, append_images=pil_images[1:]
    )
    pdf_stream.seek(0)
    return pdf_stream


# --- 3. GENERIC VIEW FUNCTIONS ---

def download_document_pdf(request, doc_type, pk):
    """
    Handles downloading a document as a PDF.

    - If the document has an uploaded PDF, it serves that file directly.
    - Otherwise, it generates a new PDF from the document's associated images.
    - The downloaded filename is based on the document's subject.
    """
    model_class = MODEL_MAP.get(doc_type)
    if not model_class:
        return HttpResponse("Invalid document type specified.", status=400)
    
    try:
        # 1. Look up the model and fetch the document instance
        document = get_object_or_404(model_class, pk=pk)

        # 2. Generate a clean filename from the subject (used in both cases)
        # The 'or' provides a fallback if the subject is empty or only has special characters.
        slug = slugify(document.subject) or f"{doc_type.replace(' ','-')}-{document.pk}"
        filename = f"{slug[:50]}.pdf"
        
        # 3. Prioritize serving an existing, uploaded PDF file.
        # Convert doc_type with space to field name with underscore (e.g., 'special order' -> 'special_order')
        document_field_name = doc_type.replace(' ', '_')
        filter_kwargs = {
            document_field_name: document,
            'content_type': 'application/pdf'
        }
        
        existing_pdf = DocumentFile.objects.filter(**filter_kwargs).first()

        if existing_pdf:
            # CASE A: A PDF was found. Serve the original file directly.
            # FileResponse is smart enough to handle the file from storage.
            return FileResponse(existing_pdf.file, as_attachment=True, filename=filename)

        # 4. If no PDF exists, fall back to generating one from images.
        else:
            # CASE B: No uploaded PDF found, so we try to create one.
            pdf_stream = generate_document_pdf(document)

            if pdf_stream:
                # Images were found and converted to a PDF stream.
                return FileResponse(pdf_stream, as_attachment=True, filename=filename)
            else:
                # No existing PDF and no images to convert.
                return HttpResponse("This document has no PDF or image content to download.", status=404)

    except model_class.DoesNotExist:
        return HttpResponse("Document not found.", status=404)
    except Exception as e:
        # Catch any other unexpected errors during file processing.
        return HttpResponse(f"An unexpected error occurred: {e}", status=500)



# --- NEW AND IMPROVED EMAIL FUNCTION ---

# --- CORRECTED EMAIL FUNCTION ---
@require_POST
def send_document_email(request, doc_type, pk):
    """
    Finds an existing PDF or generates one from images, then sends it via email.
    """
    # Use the doc_type directly from the URL. No more slug conversion.
    model_class = MODEL_MAP.get(doc_type)
    if not model_class:
        return JsonResponse({'error': 'Invalid document type specified.'}, status=400)
        
    try:
        # 1. Fetch the main document object (using pk=pk is more robust)
        document = get_object_or_404(model_class, pk=pk)

        # 2. Check for recipient emails
        emails = [r.email for r in document.recipients.all() if r.email]
        if not emails:
            return JsonResponse({'error': 'No recipient email addresses found for this document.'}, status=400)

        # 3. Prepare email shell
        email = EmailMessage(
            subject=f'Document: {document.subject or "Untitled"}',
            body='Please find the attached document.',
            to=emails,
        )

        # --- THE CORE FIX IS HERE ---
        # Directly convert the 'doc_type' string with a space to a valid field name with an underscore.
        document_field_name = doc_type.replace(' ', '_')
        
        # 4. Logic to find or generate the PDF
        filter_kwargs = {
            document_field_name: document, 
            'content_type': 'application/pdf'
        }
        first_pdf_file = DocumentFile.objects.filter(**filter_kwargs).first()

        if first_pdf_file:
            # CASE A: An uploaded PDF was found
            filename = os.path.basename(first_pdf_file.file.name)
            email.attach(filename, first_pdf_file.file.read(), 'application/pdf')
            success_message = 'Email with attached PDF sent successfully!'
        else:
            # CASE B: No uploaded PDF, fall back to generating from images
            pdf_stream = generate_document_pdf(document)

            if pdf_stream:
                # Images were found and converted
                slug = slugify(document.subject) or f"{doc_type.replace(' ','-')}-{document.pk}"
                filename = f"{slug[:50]}.pdf"
                email.attach(filename, pdf_stream.getvalue(), 'application/pdf')
                success_message = 'Email with generated PDF sent successfully!'
            else:
                # No PDF and no images were found
                return JsonResponse({'error': 'No content to send. The document has no uploaded PDF or images.'}, status=400)

        # 5. Send the email
        email.send()
        return JsonResponse({'status': 'ok', 'message': success_message})

    except Exception as e:
        # General error handler
        return JsonResponse({'error': f'An unexpected error occurred: {str(e)}'}, status=500)
#
# The entire second try...except block that started with
# "try: so = get_object_or_404(SpecialOrder, ...)"
# should be completely deleted.
#

























#DASHBOARD

# --- UPDATED HELPER FUNCTION to prioritize date range ---
def _get_period_range(params):
    """
    Computes (start_date, end_date) for the given period.
    Prioritizes start/end date over month/year.
    """
    start_date_str = params.get('start_date')
    end_date_str = params.get('end_date')
    
    # If a date range is provided, use it.
    if start_date_str and end_date_str:
        try:
            start = date.fromisoformat(start_date_str)
            end = date.fromisoformat(end_date_str)
            return start, end
        except (ValueError, TypeError):
            # Fallback in case of invalid date format
            pass

    # Otherwise, fall back to month and year.
    today = timezone.localdate()
    y = int(params.get('year')) if params.get('year') else today.year
    m = int(params.get('month')) if params.get('month') else today.month
    
    month_start = date(y, m, 1)
    last_day = calendar.monthrange(y, m)[1]
    month_end = date(y, m, last_day)
    return month_start, month_end



# --- UPDATED DASHBOARD VIEW ---
def dashboard(request):
    """
    Renders the dashboard page and provides default filter values.
    """
    today = timezone.localdate()
    
    # Provide lists for dropdowns and set the default to the current month/year.
    context = {
        'default_month': today.month,
        'default_year': today.year,
        'month_list': [(i, calendar.month_name[i]) for i in range(1, 13)],
        # Generate a list of recent years for the dropdown
        'year_list': list(range(today.year, today.year - 10, -1)),
    }
    return render(request, 'dashboard.html', context)


# --- UPDATED API VIEW FOR CARDS ---
def dashboard_cards_data(request):
    """
    JSON response with total + counts per document_type for the selected period.
    """
    start, end = _get_period_range(request.GET)
    
    models_to_count = {
        'memorandum': Memorandum, 'travel_order': TravelOrder, 'special_order': SpecialOrder,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }
    
    counts = {}
    total = 0
    for name, model in models_to_count.items():
        count = model.objects.filter(upload_date__date__gte=start, upload_date__date__lte=end).count()
        counts[name] = count
        total += count
        
    return JsonResponse({'total': total, 'counts': counts})


# --- UPDATED API VIEW FOR PIE CHART ---
def dashboard_direction_data(request):
    """
    JSON counts of incoming vs outgoing for the selected period.
    """
    start, end = _get_period_range(request.GET)
    
    models_with_category = [Memorandum, CommunicationLetter]
    
    incoming = 0
    outgoing = 0
    for model in models_with_category:
        incoming += model.objects.filter(upload_date__date__gte=start, upload_date__date__lte=end, category='incoming').count()
        outgoing += model.objects.filter(upload_date__date__lte=end, upload_date__date__gte=start, category='outgoing').count()
        
    return JsonResponse({'incoming': incoming, 'outgoing': outgoing})

















#RECIPIENT MANAGEMENT
def manage_recipients(request):
    """
    Displays a list of all recipients with options to add, edit, or delete.
    """
    recipients = Recipient.objects.all().order_by('name')
    context = {
        'recipients': recipients,
    }
    return render(request, 'manage_recipients.html', context)

def add_recipient(request):
    """
    Handles the creation of a new recipient.
    """
    if request.method == 'POST':
        form = RecipientForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('manage_recipients')
    else:
        form = RecipientForm()
    
    context = {
        'form': form,
        'form_title': 'Add New Recipient',
        'button_text': 'Save Recipient'
    }
    return render(request, 'recipient_form.html', context)

def edit_recipient(request, pk):
    """
    Handles editing an existing recipient.
    """
    recipient = get_object_or_404(Recipient, pk=pk)
    if request.method == 'POST':
        form = RecipientForm(request.POST, instance=recipient)
        if form.is_valid():
            form.save()
            return redirect('manage_recipients')
    else:
        form = RecipientForm(instance=recipient)

    context = {
        'form': form,
        'form_title': f'Edit Recipient: {recipient.name}',
        'button_text': 'Update Recipient'
    }
    return render(request, 'recipient_form.html', context)

@require_POST # This ensures the view can only be accessed via a POST request for safety
def delete_recipient(request, pk):
    """
    Handles the deletion of a recipient.
    """
    recipient = get_object_or_404(Recipient, pk=pk)
    recipient.delete()
    return redirect('manage_recipients')


















#GDRIVE ARCHIVE FUNCTIONALITY
# --- NEW HELPER FUNCTION TO GET THE GOOGLE DRIVE SERVICE ---
def get_gdrive_service():
    """
    Authenticates using the service account BUT impersonates the folder owner
    to bypass the storage quota issue.
    """
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
    # Load the service account credentials from the file
    creds = service_account.Credentials.from_service_account_file(
        settings.GOOGLE_DRIVE_SERVICE_ACCOUNT_KEY_PATH, scopes=SCOPES
    )

    # --- THIS IS THE FIX ---
    # Find the email of the person who owns the folder. You must share your
    # archive folder with your own @gmail.com account if you created it with another.
    # We will assume your personal email is the one to impersonate.
    # IMPORTANT: Go to your Google Cloud Project -> IAM & Admin -> Service Accounts
    # -> Click your service account -> Permissions -> Grant Access -> Add your personal email
    # as a member with the "Service Account Token Creator" role.
    
    # You MUST put your personal Google account email here.
    DELEGATED_USER_EMAIL = 'julahmadnur3@gmail.com' 

    # Create new credentials delegated to your personal email
    delegated_creds = creds.with_subject(DELEGATED_USER_EMAIL)
    
    service = build('drive', 'v3', credentials=delegated_creds)
    return service

# Your manage_archive view is correct and does not need to be changed.





def manage_archive(request):
    """
    Renders the main page for managing the document archive.
    This version correctly adds the necessary 'doc_type_slug' to each document
    and handles all filtering as intended.
    """
    mode = request.GET.get('mode', 'live')

    # Get Filter Parameters
    doc_type = request.GET.get('doc_type', '')
    fiscal_year_filter = request.GET.get('fiscal_year', '')
    month_filter = request.GET.get('month', '')

    # Define Models to Query
    models_to_query = {
        'special_order': SpecialOrder, 'travel_order': TravelOrder, 'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }
    models_to_fetch = {doc_type: models_to_query[doc_type]} if doc_type and doc_type in models_to_query else models_to_query

    # Fetch Documents Based on Mode
    document_list = []
    
    if mode == 'live':
        page_title = "Archive Live Documents"
        for type_slug, model_class in models_to_fetch.items():
            qs = model_class.objects.filter(is_archived=False)
            if fiscal_year_filter:
                qs = qs.filter(fiscal_year=fiscal_year_filter)
            if month_filter:
                qs = qs.filter(date__month=month_filter)

            # --- THIS IS THE CRITICAL FIX ---
            # We must add the slug and display name to each document object here.
            for doc in qs:
                doc.doc_type_slug = type_slug # e.g., 'special_order'
                doc.doc_type_display = type_slug.replace('_', ' ').title()
                document_list.append(doc)
            
    else: # mode == 'archived'
        page_title = "Manage Archived Documents"
        for type_slug, model_class in models_to_fetch.items():
            qs = model_class.objects.filter(is_archived=True)
            if fiscal_year_filter:
                qs = qs.filter(fiscal_year=fiscal_year_filter)
            if month_filter:
                qs = qs.filter(date__month=month_filter)
            
            # --- THIS FIX IS ALSO APPLIED HERE FOR CONSISTENCY ---
            for doc in qs:
                doc.doc_type_slug = type_slug # e.g., 'special_order'
                doc.doc_type_display = type_slug.replace('_', ' ').title()
                document_list.append(doc)

    # Sort the final combined list
    sorted_documents = sorted(document_list, key=lambda x: x.date if x.date else date.min, reverse=True)

    # Prepare data for filter dropdowns
    all_years = set()
    for model in models_to_query.values():
        years = model.objects.values_list('fiscal_year', flat=True).distinct()
        all_years.update(years)
    fiscal_years_for_dropdown = sorted(list(all_years), reverse=True)
    month_list = [(i, calendar.month_name[i]) for i in range(1, 13)]

    context = {
        'documents': sorted_documents,
        'page_title': page_title,
        'current_mode': mode,
        'fiscal_years_for_dropdown': fiscal_years_for_dropdown,
        'month_list': month_list,
        'filter_values': request.GET,
    }
    return render(request, 'manage_archive.html', context)


@require_POST
def archive_documents(request):
    """
    Handles batch archiving. Now works with standard Google accounts.
    """
    document_pks_by_type = {}
    for key, value in request.POST.items():
        if key.startswith('doc_'):
            parts = value.rsplit('_', 1)
            if len(parts) == 2:
                doc_type_slug, pk = parts
                if doc_type_slug not in document_pks_by_type:
                    document_pks_by_type[doc_type_slug] = []
                document_pks_by_type[doc_type_slug].append(pk)
    
    if not document_pks_by_type:
        messages.warning(request, 'No documents were selected to archive.')
        return redirect('manage_archive')

    models_map = {
        'special_order': SpecialOrder, 'travel_order': TravelOrder, 'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }

    try:
        service = get_gdrive_service()
    except Exception as e:
        messages.error(request, f"Could not connect to Google Drive. Check credentials and configuration: {e}")
        return redirect('manage_archive')

    archived_count = 0
    error_count = 0
    for doc_type_slug, pks in document_pks_by_type.items():
        model_class = models_map.get(doc_type_slug)
        if not model_class: continue

        documents_to_archive = model_class.objects.filter(pk__in=pks, is_archived=False)
        for document in documents_to_archive:
            all_files_successful = True
            for doc_file in document.documentfile_set.all():
                if doc_file.file and hasattr(doc_file.file, 'path') and os.path.exists(doc_file.file.path):
                    try:
                        file_metadata = {
                            'name': os.path.basename(doc_file.file.name),
                            'parents': [settings.GOOGLE_DRIVE_ARCHIVE_FOLDER_ID]
                        }
                        media = MediaFileUpload(doc_file.file.path, mimetype=doc_file.content_type, resumable=True)
                        # No extra flags are needed here because the service is already impersonating you
                        gdrive_file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                        gdrive_id = gdrive_file.get('id')

                        doc_file.gdrive_file_id = gdrive_id
                        os.remove(doc_file.file.path)
                        doc_file.file = None
                        doc_file.save()
                    except Exception as e:
                        print(f"ERROR: Failed to upload file for Document PK {document.pk}. Reason: {e}")
                        all_files_successful = False
                        error_count += 1
                        break
            
            if all_files_successful:
                document.is_archived = True
                document.save()
                archived_count += 1

    if archived_count > 0:
        messages.success(request, f'Successfully archived {archived_count} document(s).')
    if error_count > 0:
        messages.warning(request, f'Failed to archive {error_count} document(s). Please check server logs.')

    return redirect('manage_archive')








@require_POST
def restore_documents(request):
    """
    Handles batch restoration. Now works with standard Google accounts.
    """
    document_pks_by_type = {}
    for key, value in request.POST.items():
        if key.startswith('doc_'):
            parts = value.rsplit('_', 1)
            if len(parts) == 2:
                doc_type_slug, pk = parts
                if doc_type_slug not in document_pks_by_type:
                    document_pks_by_type[doc_type_slug] = []
                document_pks_by_type[doc_type_slug].append(pk)

    if not document_pks_by_type:
        messages.warning(request, 'No documents were selected to restore.')
        return redirect(f"{reverse('manage_archive')}?mode=archived")

    models_map = {
        'special_order': SpecialOrder, 'travel_order': TravelOrder, 'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }
    
    try:
        service = get_gdrive_service()
    except Exception as e:
        messages.error(request, f"Could not connect to Google Drive. Check credentials and configuration: {e}")
        return redirect(f"{reverse('manage_archive')}?mode=archived")
        
    restored_count = 0
    error_count = 0
    for doc_type_slug, pks in document_pks_by_type.items():
        model_class = models_map.get(doc_type_slug)
        if not model_class: continue

        documents_to_restore = model_class.objects.filter(pk__in=pks, is_archived=True)
        for document in documents_to_restore:
            all_files_successful = True
            for doc_file in document.documentfile_set.all():
                if doc_file.gdrive_file_id:
                    try:
                        gdrive_request = service.files().get_media(fileId=doc_file.gdrive_file_id)
                        fh = io.BytesIO()
                        downloader = MediaIoBaseDownload(fh, gdrive_request)
                        done = False
                        while done is False:
                            status, done = downloader.next_chunk()
                        
                        fh.seek(0)
                        filename = f"{uuid.uuid4()}"
                        doc_file.file.save(filename, ContentFile(fh.read()), save=True)

                        service.files().delete(fileId=doc_file.gdrive_file_id).execute()
                        doc_file.gdrive_file_id = None
                        doc_file.save()
                    except Exception as e:
                        print(f"ERROR: Failed to restore file for Document PK {document.pk}. Reason: {e}")
                        all_files_successful = False
                        error_count += 1
                        break
            
            if all_files_successful:
                document.is_archived = False
                document.save()
                restored_count += 1
            
    if restored_count > 0:
        messages.success(request, f'Successfully restored {restored_count} document(s).')
    if error_count > 0:
         messages.warning(request, f'Failed to restore {error_count} document(s). Please check server logs.')

    return redirect(f"{reverse('manage_archive')}?mode=archived")








# --- THIS IS THE FINAL, WORKING VIEW ARCHIVED FILE FUNCTION ---
def view_archived_file(request, doc_type, pk):
    models_map = {
        'special_order': SpecialOrder, 'travel_order': TravelOrder, 'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }
    model_class = models_map.get(doc_type.replace('-', '_'))
    if not model_class:
        raise Http404

    document = get_object_or_404(model_class, pk=pk)
    doc_file = document.documentfile_set.first()

    if not doc_file or not doc_file.gdrive_file_id:
        return HttpResponse("Archived file not found.", status=404)

    try:
        service = get_gdrive_service()
        request_gdrive = service.files().get_media(fileId=doc_file.gdrive_file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_gdrive)
        done = False
        while done is False:
            status, done = downloader.next_chunk()

        fh.seek(0)
        response = HttpResponse(fh.getvalue(), content_type=doc_file.content_type)
        # The following header makes it display inline in the browser if possible
        response['Content-Disposition'] = f'inline; filename="archived_file"'
        return response

    except Exception as e:
        return HttpResponse(f"Error viewing archived file: {e}", status=500)


