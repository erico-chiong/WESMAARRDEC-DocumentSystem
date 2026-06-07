import os
import io
import json
import re
import uuid
import calendar
import traceback
import difflib
from datetime import date, datetime, timedelta
from difflib import get_close_matches, SequenceMatcher
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.core.files.base import ContentFile
from django.core.mail import EmailMessage
from django.db.models import Q
from django.http import FileResponse, JsonResponse, HttpResponse, Http404
from django.shortcuts import render, redirect, get_object_or_404
from django.template.defaultfilters import slugify # Corrected from django.utils.text
from django.urls import reverse, reverse_lazy
from django.utils import timezone
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from django.utils.timezone import now
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from .logs_utils import create_log
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from pdf2image import convert_from_bytes
from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
import requests
# -- 3. Local Application Imports ---
from .forms import (
    SpecialOrderForm, TravelOrderForm, MemorandumForm, MOAUForm, 
    CommunicationLetterForm, OtherDocumentForm, RecipientForm
)
from .models import (
    DocumentFile, Recipient, SpecialOrder, TravelOrder, Memorandum, 
    MOAU, Signatory, CommunicationLetter, OtherDocument, Log
)
from reportlab.lib.units import inch
from PyPDF2 import PdfReader, PdfWriter
from django.utils.text import slugify





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
    and Travel Orders with improved subject extraction and fuzzy recipient matching.
    Uses parse_memorandum's subject detection logic with capitalization-based cutoff.
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
    
    # Regex patterns
    date_pattern = re.compile(
        r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b',
        re.IGNORECASE
    )
    id_pattern = re.compile(r'\b\d{6}\b')
    
    # Subject detection - using parse_memorandum's logic
    current_field = None
    collecting_subject = False
    subject_lines = []
    
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        if not stripped_line:
            continue

        lower_line = stripped_line.lower()

        # 1. Extract Document ID
        if not result['document_id']:
            match = id_pattern.search(stripped_line)
            if match:
                result['document_id'] = match.group()

        # 2. Extract Date
        if not result['date']:
            match = date_pattern.search(stripped_line)
            if match:
                try:
                    dt = datetime.strptime(match.group(0), "%B %d, %Y")
                    result['date'] = dt.strftime("%Y-%m-%d")
                except ValueError:
                    pass

        # 3. Subject Extraction - copied from parse_memorandum with capitalization cutoff
        # Check for SUBJECT field starter (using parse_memorandum's pattern)
        if re.match(r'^SUBJECT\s*:?', stripped_line, re.IGNORECASE):
            match = re.match(r'^SUBJECT\s*:?\s*(.*)', stripped_line, re.IGNORECASE)
            if match:
                value = match.group(1).strip()
                
                # Finalize previous subject if needed
                if collecting_subject:
                    result['subject'] = ' '.join(subject_lines).lstrip(':').lstrip()
                    collecting_subject = False
                
                # Start new subject collection
                current_field = 'SUBJECT'
                
                # Handle value on same line
                if value:
                    subject_lines = [value]
                    collecting_subject = True
                else:
                    # Prepare to collect from next lines
                    subject_lines = []
                    collecting_subject = True
        
        # Collect multi-line SUBJECT values with capitalization cutoff
        elif current_field == 'SUBJECT' and collecting_subject:
            # Stop at date field or non-text lines (from parse_memorandum)
            if re.match(r'DATE\s*:', stripped_line, re.IGNORECASE) or not any(c.isalpha() for c in stripped_line):
                result['subject'] = ' '.join(subject_lines).lstrip(':').lstrip()
                collecting_subject = False
                current_field = None
            else:
                # Add capitalization-based cutoff logic
                # Check if line is mostly capitalized (like original function)
                if stripped_line.isupper():
                    subject_lines.append(stripped_line)
                else:
                    # Calculate uppercase ratio for mixed case lines
                    alpha_chars = [c for c in stripped_line if c.isalpha()]
                    if alpha_chars:
                        uppercase_ratio = sum(1 for c in alpha_chars if c.isupper()) / len(alpha_chars)
                        if uppercase_ratio > 0.8:
                            subject_lines.append(stripped_line)
                        else:
                            # End of subject block - words are no longer capitalized
                            result['subject'] = ' '.join(subject_lines).lstrip(':').lstrip()
                            collecting_subject = False
                            current_field = None
                    else:
                        # No letters - end subject collection
                        result['subject'] = ' '.join(subject_lines).lstrip(':').lstrip()
                        collecting_subject = False
                        current_field = None

        # 4. Extract Recipients with fuzzy matching
        for name, recipient_id in all_recipients.items():
            # Exact match
            if name in lower_line:
                if recipient_id not in result['recipient_ids']:
                    result['recipient_ids'].append(recipient_id)
            # Fuzzy match (60% similarity)
            else:
                for word in lower_line.split():
                    if len(word) > 3:  # Only check meaningful words
                        if name.startswith(word) or name.endswith(word):
                            similarity = difflib.SequenceMatcher(None, name, word).ratio()
                            if similarity >= 0.6 and recipient_id not in result['recipient_ids']:
                                result['recipient_ids'].append(recipient_id)

    # Handle case where subject block goes until end of text
    if collecting_subject and subject_lines:
        result['subject'] = ' '.join(subject_lines).lstrip(':').lstrip()

    return result


def parse_travel_order(raw_text):
    """
    Parses raw text to extract details for Travel Orders with custom subject extraction:
    - Subject comes from the second occurrence of a line starting with "To" until first period
    - Includes the word "To" in the subject value
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
    
    # Regex patterns
    date_pattern = re.compile(
        r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b',
        re.IGNORECASE
    )
    id_pattern = re.compile(r'\b\d{6}\b')
    to_pattern = re.compile(r'^to\b\s*:?\s*', re.IGNORECASE)

    # Flags for subject extraction
    subject_extracted = False
    to_count = 0  # Counter for "To" occurrences
    
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        if not stripped_line:
            continue

        lower_line = stripped_line.lower()

        # 1. Extract Document ID
        if not result['document_id']:
            match = id_pattern.search(stripped_line)
            if match:
                result['document_id'] = match.group()

        # 2. Extract Date
        if not result['date']:
            match = date_pattern.search(stripped_line)
            if match:
                try:
                    dt = datetime.strptime(match.group(0), "%B %d, %Y")
                    result['date'] = dt.strftime("%Y-%m-%d")
                except ValueError:
                    pass

        # 3. Subject Extraction - SECOND "To" occurrence including "To"
        if not subject_extracted:
            # Check for "To" at line start
            if to_pattern.match(stripped_line):
                to_count += 1
                
                # Only process the SECOND occurrence of "To"
                if to_count == 2:
                    subject_parts = []
                    # Keep the entire line including "To"
                    subject_text = stripped_line
                    subject_parts.append(subject_text)
                    
                    # Check if period exists in current line
                    if '.' in subject_text:
                        result['subject'] = subject_text.split('.', 1)[0]
                        subject_extracted = True
                    else:
                        # Continue through next lines until period (max 5 lines)
                        for j in range(i+1, min(i+6, len(lines))):
                            next_line = lines[j].strip()
                            if not next_line:
                                continue
                                
                            if '.' in next_line:
                                subject_parts.append(next_line.split('.', 1)[0])
                                result['subject'] = ' '.join(subject_parts)
                                subject_extracted = True
                                break
                            else:
                                subject_parts.append(next_line)
                        else:
                            # If we didn't find a period, take all we have
                            result['subject'] = ' '.join(subject_parts)
                            subject_extracted = True

        # 4. Extract Recipients with fuzzy matching
        for name, recipient_id in all_recipients.items():
            # Exact match
            if name in lower_line:
                if recipient_id not in result['recipient_ids']:
                    result['recipient_ids'].append(recipient_id)
            # Fuzzy match (60% similarity)
            else:
                # Split line into words and check for partial matches
                for word in lower_line.split():
                    if len(word) > 3:  # Only check meaningful words
                        if name.startswith(word) or name.endswith(word):
                            similarity = difflib.SequenceMatcher(None, name, word).ratio()
                            if similarity >= 0.6 and recipient_id not in result['recipient_ids']:
                                result['recipient_ids'].append(recipient_id)

    return result



def parse_moau(first_page_text, last_page_text, second_to_last_page_text=None):
    """
    Optimized MOA/U parser (updated):
    - Removed DB fuzzy matching for signatories (now pure pattern-based name extraction)
    - Replaced recipient logic with the parse_travel_order style (exact + fuzzy)
    - Preserved all other original logic (agencies, representatives detection, date, subject rules)
    
    FIXED ONLY THESE 4 ISSUES (keeping everything else unchanged):
    1. Subject - Better handling for OCR variations
    2. First Party Agency - Fixed regex to capture full name  
    3. First Party Representative - Better name extraction
    4. Signatories - Better filtering to avoid boilerplate
    """
    result = {
        'subject': '', 'date': timezone.now().strftime("%Y-%m-%d"),
        'first_party_agency': '', 'first_party_representative': '',
        'second_party_agency': '', 'second_party_representative': '',
        'signatory_names': [], 
        'recipient_ids': []
    }
    
    # --- 1. PRE-COMPUTATION AND SETUP (unchanged) ---
    all_recipients_qs = Recipient.objects.all()
    all_recipients = {r.name.lower(): r.id for r in all_recipients_qs}
    
    full_text = first_page_text or ""
    if second_to_last_page_text:
        full_text += "\n" + second_to_last_page_text
    if last_page_text:
        full_text += "\n" + last_page_text

    # --- Helper: clean name pattern (unchanged) ---
    def extract_clean_name(raw_line_text):
        """
        Clean an OCR'd name-like string and return a normalized uppercase name.
        - Strips common titles (Dr, Mr, Ms, Engr, etc.)
        - Removes stray punctuation
        - Collapses whitespace
        - Keeps plausible name parts (2-5 words)
        """
        if not raw_line_text:
            return ''
        # Remove prefixes/titles
        cleaned = re.sub(r'\b(DR|MR|MRS|MS|ENGR|MISS)\.?\s+', '', raw_line_text, flags=re.IGNORECASE)
        # Remove undesirable chars but keep letters, spaces, dots, commas, hyphens
        cleaned = re.sub(r'[^A-Za-z\s\.,\-]', ' ', cleaned)
        # Collapse whitespace and trim
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        # Prefer a multi-word capitalized name if present
        name_match = re.search(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,4})\b', cleaned)
        if name_match:
            return name_match.group(1).upper()
        return cleaned.upper()

    # --- 2. FIRST PAGE PROCESSING ---
    if first_page_text:
        # --- SUBJECT extraction: FIXED for better OCR handling ---
        lines = first_page_text.splitlines()
        collecting_subject = False
        subject_lines = []
        
        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            
            # Start trigger 1: MEMORANDUM OF AGREEMENT ON
            if re.search(r'MEMORANDUM OF AGREEMENT ON', stripped, re.IGNORECASE):
                collecting_subject = True
                after_phrase = re.split(r'MEMORANDUM OF AGREEMENT ON', stripped, flags=re.IGNORECASE)[-1].strip()
                if after_phrase:
                    subject_lines.append(after_phrase)
                continue
            
            # Start trigger 2: Project Title
            if re.search(r'Project Title', stripped, re.IGNORECASE):
                collecting_subject = True
                after_phrase = re.split(r'Project Title', stripped, flags=re.IGNORECASE)[-1].strip(':').strip()
                if after_phrase:
                    subject_lines.append(after_phrase)
                continue
            
            # End triggers
            if re.search(r'KNOW ALL MEN BY THESE PRESENTS', stripped, re.IGNORECASE):
                collecting_subject = False
                break
            if re.search(r'Project Site', stripped, re.IGNORECASE):
                collecting_subject = False
                break
            
            if collecting_subject:
                subject_lines.append(stripped)

        if subject_lines:
            result['subject'] = ' '.join(subject_lines).strip()
        else:
            # ENHANCED fallback patterns for better OCR handling
            subject_patterns = [
                r'MEMORANDUM OF AGREEMENT ON\s*[""]?(.*?)[""]?\s*KNOW ALL MEN BY THESE PRESENTS:',
                r'Project Title[:\s*]\s*[""]?(.*?)[""]?\s*Project Site',
                # Additional patterns for OCR variations
                r'MEMORANDUM OF AGREEMENT\s*ON\s*[""]?(.*?)[""]?\s*(?=KNOW ALL MEN|Project Site|WMSU)',
                r'Project Title\s*(.*?)(?=Project Site|WMSU|DEPARTMENT)'
            ]
            for pattern in subject_patterns:
                match = re.search(pattern, first_page_text, re.DOTALL | re.IGNORECASE)
                if match:
                    subject_text = match.group(1).replace('\n', ' ').strip()
                    # Clean up quotes and extra whitespace
                    subject_text = re.sub(r'^[""\'\s]+|[""\'\s]+$', '', subject_text)
                    subject_text = re.sub(r'\s+', ' ', subject_text)
                    if subject_text:
                        result['subject'] = subject_text
                        break
        
        # --- FIRST PARTY AGENCY: FIXED to capture full name ---
        first_party_agency_match = re.search(r'((?:The )?DEPARTMENT OF [A-Z\s&,-]+)', first_page_text)
        if first_party_agency_match:
            agency_text = first_party_agency_match.group(1)
            
            # Continue capturing until we hit lowercase or specific patterns
            remaining_text = first_page_text[first_party_agency_match.end():]
            continuation_match = re.match(r'([A-Z\s&,-]+?)(?=\s+hereinafter|\s*,|\s+represented)', remaining_text)
            if continuation_match:
                continuation = continuation_match.group(1).strip()
                # Only add if it looks like part of agency name
                if continuation and not re.search(r'\bhereinafter\b', continuation, re.IGNORECASE):
                    agency_text += ' ' + continuation
            
            result['first_party_agency'] = agency_text.strip()
            
            # --- FIRST PARTY REPRESENTATIVE: FIXED for better name extraction ---
            # Look for the representative after the agency
            agency_section_end = first_party_agency_match.end() + (len(continuation_match.group(1)) if continuation_match else 0)
            remaining_text = first_page_text[agency_section_end:]
            
            # Find representative pattern
            name_match = re.search(
                r'represented by[^,]+,\s*((?:(?:DR|MS|ENGR|MR|MRS)\.?\s*)?(?:[A-Z][A-Z\s.,-]*)+)(?=\s*[a-z]|;|,)',
                remaining_text,
                re.IGNORECASE
            )
            if name_match:
                raw_name = name_match.group(1).strip()
                # Extract only capitalized name parts, remove titles
                name_parts = []
                for word in raw_name.split():
                    word_clean = re.sub(r'[,.;]$', '', word)  # Remove trailing punctuation
                    if (word_clean.upper() not in ['DR', 'DR.', 'MR', 'MR.', 'MRS', 'MRS.', 'MS', 'MS.', 'ENGR', 'ENGR.'] and 
                        len(word_clean) > 1 and word_clean[0].isupper()):
                        name_parts.append(word_clean.upper())
                
                if name_parts:
                    result['first_party_representative'] = ' '.join(name_parts)
        
        # --- SECOND PARTY AGENCY (unchanged - working) ---
         # --- SECOND PARTY AGENCY (unchanged) ---
        second_party_agency_match = re.search(r'WESTERN MINDANAO STATE UNIVERSITY', first_page_text, re.IGNORECASE)
        if second_party_agency_match:
            result['second_party_agency'] = second_party_agency_match.group(0)
            name_match = re.search(
                r'represented by[^,]+,\s*((?:DR\.|MS\.|ENGR\.|MR\.|MRS\.)?\s*MA\.?\s*CARLA\s*A\.?\s*O[CK]HOTORENA[A-Z\s]*)',
                first_page_text[second_party_agency_match.end():],
                re.IGNORECASE
            )
            if name_match:
                raw_name = re.sub(r'\s+', ' ', name_match.group(1)).strip()
                result['second_party_representative'] = extract_clean_name(raw_name)
        
    
    # --- 3. GLOBAL RECIPIENT PROCESSING (unchanged - working) ---
    recipient_ids = []
    for line in full_text.splitlines():
        stripped_line = line.strip()
        if not stripped_line:
            continue
        lower_line = stripped_line.lower()

        for name, recipient_id in all_recipients.items():
            # Exact match
            if name in lower_line:
                if recipient_id not in recipient_ids:
                    recipient_ids.append(recipient_id)
            else:
                # Fuzzy match (word-level)
                for word in lower_line.split():
                    if len(word) > 3:
                        if name.startswith(word) or name.endswith(word):
                            similarity = difflib.SequenceMatcher(None, name, word).ratio()
                            if similarity >= 0.6 and recipient_id not in recipient_ids:
                                recipient_ids.append(recipient_id)
    result['recipient_ids'] = recipient_ids

    # --- 4. SIGNATORY PROCESSING: FIXED to avoid boilerplate text ---
    signatory_names = set()
    
    # Add representatives automatically (they were cleaned earlier)
    for field in ['first_party_representative', 'second_party_representative']:
        if rep := result.get(field):
            signatory_names.add(rep.strip().upper())
    
    # Find witness block with better filtering
    witness_text = ""
    witness_patterns = [
        r'(?:IN\s+)?WITNESS\s+WHEREOF.*?(?=\n\n|\nACKNOWLEDGMENT|\nCERTIFIED|$)',
        r'WITNESSES?:\s*.*?(?=\n\n|\nACKNOWLEDGMENT|\nCERTIFIED|$)',
        r'Signed in the Presence of:.*?(?=\n\n|\nACKNOWLEDGMENT|\nCERTIFIED|$)'
    ]
    
    for pattern in witness_patterns:
        match = re.search(pattern, full_text, re.DOTALL | re.IGNORECASE)
        if match:
            witness_text = match.group(0)
            break

    # If no specific witness block found, fall back to original simple search
    if not witness_text:
        for phrase in ["WITNESS WHEREOF", "WITNESSETH", "SIGNED IN THE PRESENCE OF"]:
            if phrase.lower() in full_text.lower():
                start_idx = full_text.lower().index(phrase.lower())
                witness_text = full_text[start_idx:]
                break

    # Extract candidate names from witness block with strict filtering
    if witness_text:
        # Skip obvious document boilerplate patterns
        lines = witness_text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Skip boilerplate text
            skip_patterns = [
                r'WITNESS\s+WHEREOF', r'IN\s+WITNESS', r'WITNESSES?:', 
                r'set their hands', r'above mentioned', r'By:', 
                r'Signed in the Presence', r'CERTIFIED', r'Adequate available',
                r'hereto have', r'date and at', r'DEPARTMENT OF', r'UNIVERSITY'
            ]
            
            should_skip = False
            for skip_pattern in skip_patterns:
                if re.search(skip_pattern, line, re.IGNORECASE):
                    should_skip = True
                    break
                    
            if should_skip:
                continue
            
            # Look for name-like patterns
            name_candidates = re.findall(r'\b([A-Z][A-Z\s\.]{4,}[A-Z])\b', line)
            for candidate in name_candidates:
                candidate = candidate.strip()
                
                # Filter out obvious non-names
                if (8 <= len(candidate) <= 50 and  # Reasonable length for full names
                    len(candidate.split()) >= 2 and  # At least 2 words
                    not re.search(r'\b(DEPARTMENT|OFFICE|UNIVERSITY|REGIONAL|DIRECTOR|PRESIDENT|COUNCIL|TECHNOLOGY)\b', candidate) and
                    candidate.count('.') <= 3):  # Not too many periods
                    
                    cleaned = extract_clean_name(candidate)
                    if cleaned and len(cleaned.split()) >= 2:
                        signatory_names.add(cleaned.upper())

    # Always set this, even if no witnesses were found
    result['signatory_names'] = sorted(signatory_names)

    return result


def parse_memorandum(raw_text):
    """
    Parses OCR text for OUTGOING Memorandums with improved DATE and FROM field extraction
    """
    result = {
        'document_id': '', 'date': '', 'subject': '',
        'for_field': '', 'from_field': '', 'thru': '',
        'recipient_ids': []
    }
    
    lines = raw_text.splitlines()
    
    # Pre-fetch recipient data
    all_recipients_qs = Recipient.objects.all()
    recipient_name_to_id_map = {r.name.upper(): r.id for r in all_recipients_qs}

    # Regex patterns
    date_pattern = re.compile(
        r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b|\b\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b',
        re.IGNORECASE
    )
    id_pattern = re.compile(r'Memorandum\s+(?:Order\s+)?No\.\s*([\d-]+)', re.IGNORECASE)
    
    # Helper function for fuzzy matching
    def extract_closest_name(raw_line_text):
        """Fuzzy match names using SequenceMatcher"""
        best_match = None
        best_similarity = 0
        cleaned_text = re.sub(r'\b(DR|MS|ENGR|MR)\.\s*', '', raw_line_text.strip(), flags=re.IGNORECASE).upper()
        
        for recipient in all_recipients_qs:
            similarity = difflib.SequenceMatcher(None, cleaned_text, recipient.name.upper()).ratio()
            if similarity > best_similarity and similarity >= 0.6:
                best_similarity = similarity
                best_match = recipient.name
        return best_match or cleaned_text

    # Document ID extraction
    for line in lines:
        if not result['document_id']:
            match = id_pattern.search(line)
            if match:
                result['document_id'] = match.group(1).strip()
                break

    # Field extraction with multi-line support
    fields = {
        'FOR': {'value': '', 'found': False},
        'FROM': {'value': '', 'found': False},
        'THRU': {'value': '', 'found': False},
        'SUBJECT': {'value': '', 'found': False, 'lines': []},
        'DATE': {'value': '', 'found': False}
    }
    
    current_field = None
    collecting_subject = False
    
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        if not stripped_line:
            continue
            
        # Check for field starters
        if re.match(r'^(FOR|FROM|THRU|SUBJECT|DATE)\s*:?', stripped_line, re.IGNORECASE):
            match = re.match(r'^(FOR|FROM|THRU|SUBJECT|DATE)\s*:?\s*(.*)', stripped_line, re.IGNORECASE)
            if match:
                field_type = match.group(1).upper()
                value = match.group(2).strip()
                
                # Finalize previous subject if needed
                if collecting_subject:
                    fields['SUBJECT']['value'] = ' '.join(fields['SUBJECT']['lines']).lstrip(':').lstrip()
                    collecting_subject = False
                
                # Start new field
                current_field = field_type
                fields[field_type]['found'] = True
                
                # Special handling for FOR field
                if field_type == 'FOR':
                    fields['FOR']['value'] = "DR. MA. CARLA A. OCHOTORENA"
                
                # Handle values on same line
                elif value:
                    if field_type == 'SUBJECT':
                        fields['SUBJECT']['lines'] = [value]
                        collecting_subject = True
                    elif field_type == 'DATE':
                        # Extract date from value
                        date_match = date_pattern.search(value)
                        if date_match:
                            try:
                                date_str = date_match.group(0)
                                # Handle different date formats
                                if ',' in date_str:
                                    dt = datetime.strptime(date_str, "%B %d, %Y")
                                else:
                                    dt = datetime.strptime(date_str, "%d %B %Y")
                                fields['DATE']['value'] = dt.strftime("%Y-%m-%d")
                            except:
                                pass
                    else:
                        fields[field_type]['value'] = extract_closest_name(value)
                # Prepare to collect from next lines if no value
                elif field_type == 'SUBJECT':
                    fields['SUBJECT']['lines'] = []
                    collecting_subject = True
                elif field_type == 'FROM' or field_type == 'THRU':
                    # Expect value on next line
                    current_field = field_type
        
        # Collect multi-line values
        elif current_field:
            # FROM/THRU value collection (single line)
            if current_field in ['FROM', 'THRU'] and not fields[current_field]['value']:
                fields[current_field]['value'] = extract_closest_name(stripped_line)
                current_field = None
            
            # SUBJECT value collection
            elif current_field == 'SUBJECT' and collecting_subject:
                # Stop at date field or non-text lines
                if re.match(r'DATE\s*:', stripped_line, re.IGNORECASE) or not any(c.isalpha() for c in stripped_line):
                    fields['SUBJECT']['value'] = ' '.join(fields['SUBJECT']['lines']).lstrip(':').lstrip()
                    collecting_subject = False
                    current_field = None
                else:
                    fields['SUBJECT']['lines'].append(stripped_line)
            
            # DATE value collection
            elif current_field == 'DATE' and not fields['DATE']['value']:
                date_match = date_pattern.search(stripped_line)
                if date_match:
                    try:
                        date_str = date_match.group(0)
                        # Handle different date formats
                        if ',' in date_str:
                            dt = datetime.strptime(date_str, "%B %d, %Y")
                        else:
                            dt = datetime.strptime(date_str, "%d %B %Y")
                        fields['DATE']['value'] = dt.strftime("%Y-%m-%d")
                        current_field = None
                    except:
                        pass
    
    # Finalize subject if still collecting
    if collecting_subject and fields['SUBJECT']['lines']:
        fields['SUBJECT']['value'] = ' '.join(fields['SUBJECT']['lines']).lstrip(':').lstrip()
    
    # Map fields to result
    result['for_field'] = fields['FOR']['value']
    result['from_field'] = fields['FROM']['value']
    result['thru'] = fields['THRU']['value']
    result['subject'] = fields['SUBJECT']['value']
    result['date'] = fields['DATE']['value']

    # Extract recipients with fuzzy matching
    for line in lines:
        lower_line = line.lower()
        for name, recipient_id in {r.name.lower(): r.id for r in all_recipients_qs}.items():
            # Exact match
            if name in lower_line:
                if recipient_id not in result['recipient_ids']:
                    result['recipient_ids'].append(recipient_id)
            # Fuzzy match using SequenceMatcher
            else:
                for word in lower_line.split():
                    if len(word) > 3:
                        if name.startswith(word) or name.endswith(word):
                            similarity = difflib.SequenceMatcher(None, name, word).ratio()
                            if similarity >= 0.6 and recipient_id not in result['recipient_ids']:
                                result['recipient_ids'].append(recipient_id)

    # Add FOR/FROM/THRU to recipient_ids
    for field in ['for_field', 'from_field', 'thru']:
        name = result[field].upper()
        if name in recipient_name_to_id_map:
            recipient_id = recipient_name_to_id_map[name]
            if recipient_id not in result['recipient_ids']:
                result['recipient_ids'].append(recipient_id)

    return result


def parse_communication_letter(raw_text, category):
    """
    Parses OCR text for Communication Letters with optimized TO/THRU and FROM handling.
    Rules:
    - For INCOMING: 
        • FROM = external (raw text) 
        • TO/THRU = internal (cross-match with recipients)
    - For OUTGOING:
        • FROM = internal (cross-match with recipients)  # Uses incoming TO/THRU logic
        • TO/THRU = external (raw text)  # Uses incoming FROM logic
    - Position constraints:
        • TO/THRU always before Subject
        • FROM always in signature block at bottom
    """
    result = {
        'subject': '',
        'from_field': '',
        'to_or_thru': '',  # Changed from 'to_field' to match HTML form
        'date': '',
        'recipient_ids': []
    }
    
    # --- 1. SETUP AND PRE-COMPUTATION ---
    all_recipients_qs = Recipient.objects.all()
    recipient_names = [r.name for r in all_recipients_qs]
    name_to_id = {r.name.upper(): r.id for r in all_recipients_qs}
    lines = [ln.strip() for ln in raw_text.splitlines()]
    date_pattern = re.compile(r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}', re.IGNORECASE)

    # Helper for internal name matching
    def match_internal_name(text):
        cleaned = re.sub(r'\b(DR|MS|ENGR|MR|MRS)\.?\s*', '', text, flags=re.IGNORECASE)
        matches = get_close_matches(cleaned, recipient_names, n=1, cutoff=0.7)
        return matches[0] if matches else text

    # --- 2. FIND DATE ---
    for line in lines:
        date_match = date_pattern.search(line)
        if date_match:
            try:
                dt = datetime.strptime(date_match.group(), "%B %d, %Y")
                result['date'] = dt.strftime("%Y-%m-%d")
                break
            except ValueError:
                pass

    # --- 3. FIND SUBJECT ---
    subject_idx = -1
    dear_found = False
    subject_lines = []
    
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        
        if not dear_found and re.search(r'\bdear\b', stripped_line, re.IGNORECASE):
            dear_found = True
            continue
        
        if dear_found and stripped_line:
            if any(skip in stripped_line.upper() for skip in ['GREETINGS', 'GOOD DAY', 'HOPE THIS LETTER']):
                continue
            subject_lines.append(stripped_line)
            if re.search(r'[.!?]\s*$', stripped_line):
                break
            if len(subject_lines) >= 5:
                break
    
    if subject_lines:
        result['subject'] = ' '.join(subject_lines)

    # --- 4. FIND TO/THRU ---
    if category == 'incoming':
        # INCOMING TO/THRU: internal matching (recipient database lookup)
        for line in lines[:15]:  # Same as recipient but first 15 lines only
            lower_line = line.lower()
            for name, recipient_id in name_to_id.items():
                if name.lower() in lower_line:
                    if result['to_or_thru'] == '':  # Only set if empty (like recipient_ids check)
                        result['to_or_thru'] = [r.name for r in all_recipients_qs if r.id == recipient_id][0]
                        break  # Break after first match
                else:
                    for word in lower_line.split():
                        if len(word) > 3:
                            if name.lower().startswith(word) or name.lower().endswith(word):
                                similarity = difflib.SequenceMatcher(None, name.lower(), word).ratio()
                                if similarity >= 0.6 and result['to_or_thru'] == '':
                                    result['to_or_thru'] = [r.name for r in all_recipients_qs if r.id == recipient_id][0]
                                    break
                        if result['to_or_thru'] != '':
                            break
                if result['to_or_thru'] != '':
                    break
            if result['to_or_thru'] != '':
                break

        if not result['to_or_thru']:
            result['to_or_thru'] = "DR. MA. CARLA A. OCHOTORENA"

    else:
        # OUTGOING TO/THRU: external (raw text) - using incoming FROM logic
        header = lines[:subject_idx] if subject_idx != -1 else lines[:15]
        found_to = False
        for i, line in enumerate(header):
            if re.match(r'^\s*(TO|THRU)\s*:', line, re.IGNORECASE):
                name_part = line.split(':', 1)[1].strip()
                if not name_part and i+1 < len(header):
                    name_part = header[i+1].strip()
                result['to_or_thru'] = name_part  # Keep raw text (like incoming FROM)
                found_to = True
                break
        if not found_to:
            for line in header:
                if not line or date_pattern.search(line):
                    continue
                if len(line) > 5 and len(line.split()) >= 2:
                    result['to_or_thru'] = line  # Keep raw text (like incoming FROM)
                    break

    # --- 5. FIND FROM ---
    footer = lines[-12:]
    found_from = False
    closing_phrases = ["sincerely", "truly yours", "respectfully", "very truly"]
    
    if category == 'incoming':
        # INCOMING FROM: external (raw text)
        for i, line in enumerate(footer):
            if re.match(r'^\s*FROM\s*:', line, re.IGNORECASE):
                name_part = line.split(':', 1)[1].strip()
                if not name_part and i+1 < len(footer):
                    name_part = footer[i+1].strip()
                result['from_field'] = name_part  # Keep raw text
                found_from = True
                break
        
        if not found_from:
            closing_idx = -1
            for i, line in enumerate(footer):
                if any(phrase in line.lower() for phrase in closing_phrases):
                    closing_idx = i
                    break
            
            if closing_idx != -1:
                for line in footer[closing_idx+1:closing_idx+4]:
                    if line and len(line) > 4:
                        result['from_field'] = line  # Keep raw text
                        break
    else:
        # OUTGOING FROM: internal matching (recipient database lookup) - using incoming TO/THRU logic
        for line in footer:
            lower_line = line.lower()
            for name, recipient_id in name_to_id.items():
                if name.lower() in lower_line:
                    if result['from_field'] == '':  # Only set if empty
                        result['from_field'] = [r.name for r in all_recipients_qs if r.id == recipient_id][0]
                        break  # Break after first match
                else:
                    for word in lower_line.split():
                        if len(word) > 3:
                            if name.lower().startswith(word) or name.lower().endswith(word):
                                similarity = difflib.SequenceMatcher(None, name.lower(), word).ratio()
                                if similarity >= 0.6 and result['from_field'] == '':
                                    result['from_field'] = [r.name for r in all_recipients_qs if r.id == recipient_id][0]
                                    break
                        if result['from_field'] != '':
                            break
                if result['from_field'] != '':
                    break
            if result['from_field'] != '':
                break

    # --- 6. RECIPIENT ID MATCHING (restored from old logic) ---
    for line in lines:
        lower_line = line.lower()
        for name, recipient_id in name_to_id.items():
            if name.lower() in lower_line:
                if recipient_id not in result['recipient_ids']:
                    result['recipient_ids'].append(recipient_id)
            else:
                for word in lower_line.split():
                    if len(word) > 3:
                        if name.lower().startswith(word) or name.lower().endswith(word):
                            similarity = difflib.SequenceMatcher(None, name.lower(), word).ratio()
                            if similarity >= 0.6 and recipient_id not in result['recipient_ids']:
                                result['recipient_ids'].append(recipient_id)

    for field in ['to_or_thru', 'from_field']:
        name = result[field].upper()
        if name in name_to_id:
            recipient_id = name_to_id[name]
            if recipient_id not in result['recipient_ids']:
                result['recipient_ids'].append(recipient_id)

    return result


#unify scan for testing only
@csrf_exempt
@require_POST
def unified_scan_view(request):
    """
    Handles file upload, OCR, and calls the correct parser based on doc_type.
    This version correctly handles single/multiple PDFs and single/multiple images,
    including the new 3-page logic for MOAUs.
    """
    try:
        # --- PRESERVED LOGIC for getting files from create or edit pages ---
        uploaded_files = request.FILES.getlist('files')
        if not uploaded_files:
            single_file = request.FILES.get('file')
            if single_file:
                uploaded_files = [single_file]
        if not uploaded_files:
            return JsonResponse({'error': 'No file(s) were uploaded.'}, status=400)
        
        doc_type = request.POST.get('doc_type')
        category = request.POST.get('category')
        api_key = 'K84532793088957'
        ocr_pages_text = []
        
        first_file = uploaded_files[0]
        is_pdf = first_file.content_type == 'application/pdf'

        # --- PRESERVED PDF Processing Logic (with updated MOAU page selection) ---
        if is_pdf:
            if len(uploaded_files) > 1:
                return JsonResponse({'error': 'Please upload only one PDF file at a time.'}, status=400)
            pdf_bytes = first_file.read()

            # Try to render PDF pages locally (requires Poppler). If Poppler is not available,
            # fall back to submitting the PDF directly to OCR.space.
            pil_images = None
            try:
                poppler_path = getattr(settings, 'POPPLER_PATH', None) or os.environ.get('POPPLER_PATH')
                if poppler_path:
                    poppler_path = os.path.expandvars(poppler_path)
                    if not os.path.isdir(poppler_path):
                        poppler_path = None

                if poppler_path:
                    pil_images = convert_from_bytes(pdf_bytes, poppler_path=poppler_path)
                else:
                    pil_images = convert_from_bytes(pdf_bytes)
            except (PDFInfoNotInstalledError, PDFPageCountError, FileNotFoundError, OSError):
                pil_images = None

            if pil_images is None:
                # Fallback: submit the PDF directly to OCR.space.
                # Note: Depending on your OCR.space plan, PDFs may have a page limit.
                pdf_page_count = None
                try:
                    pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
                    pdf_page_count = len(pdf_reader.pages)
                except Exception:
                    pdf_page_count = None

                ocr_response = requests.post(
                    'https://api.ocr.space/parse/image',
                    files={'file': (first_file.name, pdf_bytes, 'application/pdf')},
                    data={
                        'isOverlayRequired': False,
                        'apikey': api_key,
                        'language': 'eng',
                        'scale': True,
                        'OCREngine': 2,
                        'filetype': 'PDF',
                    },
                )
                ocr_response.raise_for_status()
                ocr_result = ocr_response.json()
                if ocr_result.get('IsErroredOnProcessing'):
                    return JsonResponse({'error': f"OCR Failed: {ocr_result.get('ErrorMessage', ['Unknown'])[0]}"}, status=400)

                parsed_results = ocr_result.get('ParsedResults') or []
                if not parsed_results:
                    return JsonResponse({'error': 'OCR could not extract any text from the PDF.'}, status=400)

                # If we need last-page text (MOAU) but the OCR API didn’t return all pages,
                # we can't reliably pick the last/second-to-last page.
                if doc_type == 'moau' and pdf_page_count and len(parsed_results) < pdf_page_count:
                    return JsonResponse(
                        {
                            'error': 'Poppler is required to OCR the last pages of this PDF (MOAU parsing).',
                            'details': 'Install Poppler and either add its bin folder to PATH or set POPPLER_PATH (points to the folder containing pdfinfo.exe).',
                            'pdf_page_count': pdf_page_count,
                            'ocr_returned_pages': len(parsed_results),
                        },
                        status=500,
                    )

                def _parsed_text(idx):
                    try:
                        return (parsed_results[idx] or {}).get('ParsedText') or ''
                    except Exception:
                        return ''

                if doc_type == 'moau' and len(parsed_results) >= 3:
                    ocr_pages_text.extend([_parsed_text(0), _parsed_text(-2), _parsed_text(-1)])
                elif doc_type == 'moau' and len(parsed_results) == 2:
                    ocr_pages_text.extend([_parsed_text(0), _parsed_text(1)])
                else:
                    ocr_pages_text.append(_parsed_text(0))

            else:
                pages_to_ocr = []
                # --- THIS IS THE UPDATED PDF PAGE SELECTION LOGIC ---
                if doc_type == 'moau' and len(pil_images) > 2:
                    pages_to_ocr.append(pil_images[0])      # First page
                    pages_to_ocr.append(pil_images[-2])     # Second-to-last page
                    pages_to_ocr.append(pil_images[-1])     # Last page
                elif doc_type == 'moau' and len(pil_images) == 2:
                    pages_to_ocr.append(pil_images[0])      # First page
                    pages_to_ocr.append(pil_images[1])      # Last page
                elif pil_images:
                    pages_to_ocr.append(pil_images[0]) # Default for all other cases

                for image in pages_to_ocr:
                    with io.BytesIO() as output:
                        image.save(output, format="JPEG")
                        ocr_response = requests.post('https://api.ocr.space/parse/image', files={'file': ('image.jpg', output.getvalue(), 'image/jpeg')}, data={'isOverlayRequired': False, 'apikey': api_key, 'language': 'eng', 'scale': True, 'OCREngine': 2})
                        ocr_response.raise_for_status()
                        ocr_result = ocr_response.json()
                        if ocr_result.get('IsErroredOnProcessing'): return JsonResponse({'error': f"OCR Failed: {ocr_result.get('ErrorMessage', ['Unknown'])[0]}"}, status=400)
                        if ocr_result.get('ParsedResults'): ocr_pages_text.append(ocr_result['ParsedResults'][0]['ParsedText'])

        # --- PRESERVED Image Processing Logic ---
        # The frontend sends the correct 1, 2, or 3 images, so we just process all of them.
        else:
            for file_obj in uploaded_files:
                ocr_response = requests.post('https://api.ocr.space/parse/image', data={'isOverlayRequired': False, 'apikey': api_key, 'language': 'eng', 'scale': True, 'OCREngine': 2}, files={'file': (file_obj.name, file_obj.read(), file_obj.content_type)})
                ocr_response.raise_for_status()
                ocr_result = ocr_response.json()
                if ocr_result.get('IsErroredOnProcessing'): return JsonResponse({'error': f"OCR Failed: {ocr_result.get('ErrorMessage', ['Unknown'])[0]}"}, status=400)
                if ocr_result.get('ParsedResults'): ocr_pages_text.append(ocr_result['ParsedResults'][0]['ParsedText'])

        if not ocr_pages_text:
             return JsonResponse({'error': 'OCR could not extract any text from the document.'}, status=400)

        # --- PARSING DISPATCHER (with updated MOAU call) ---
        parsed_data = {}
        first_page_text = ocr_pages_text[0] if ocr_pages_text else ""
        
        if doc_type in ['special_order']:
            parsed_data = parse_special_order_memorandum_travel(first_page_text)
        elif doc_type == 'memorandum':
            if category == 'incoming':
                parsed_data = parse_special_order_memorandum_travel(first_page_text)
            elif category == 'outgoing':
                parsed_data = parse_memorandum(first_page_text)
            else:
                return JsonResponse({'error': 'Memorandum category (incoming/outgoing) is required for parsing.'}, status=400)
        
        elif doc_type == 'moau':
            # --- THIS IS THE UPDATED MOAU PARSING LOGIC ---
            if len(ocr_pages_text) >= 3:
                # Case for 3+ pages/images (first, second-to-last, last)
                # The frontend sends them in the order: [first, second-to-last, last]
                parsed_data = parse_moau(
                    first_page_text=ocr_pages_text[0], 
                    last_page_text=ocr_pages_text[2], 
                    second_to_last_page_text=ocr_pages_text[1]
                )
            elif len(ocr_pages_text) == 2:
                # Case for 2 pages/images (first, last)
                parsed_data = parse_moau(
                    first_page_text=ocr_pages_text[0], 
                    last_page_text=ocr_pages_text[1]
                )
            else:
                # Case for 1 page/image
                parsed_data = parse_moau(
                    first_page_text=first_page_text, 
                    last_page_text=first_page_text
                )
        
        elif doc_type == 'communication_letter':
            parsed_data = parse_communication_letter(first_page_text, category)
        elif doc_type == 'travel_order':
            parsed_data = parse_travel_order(first_page_text)
        else:
            return JsonResponse({'error': f'Unsupported document type for parsing: {doc_type}'}, status=400)

        parsed_data['raw_ocr_text'] = "\n---\n".join(ocr_pages_text)
        return JsonResponse(parsed_data)

    except requests.exceptions.RequestException as e:
        return JsonResponse({'error': f'Network error communicating with OCR service: {e}'}, status=500)
    except Exception as e:
        traceback.print_exc()
        return JsonResponse({'error': f'An unexpected server error occurred: {str(e)}'}, status=500)



 # Protect the view and ensure we have a user
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

    obj = get_object_or_404(Model, pk=pk)

    if request.method == 'POST':
        # --- ADDED: Create a log entry BEFORE deleting the object ---
        # We must do this first, otherwise the document details will be gone.
        create_log(request.user, f"Deleted {obj.__class__.__name__}", obj)

        # Your existing deletion logic is preserved
        for f in obj.documentfile_set.all():
            if f.file:
                f.file.delete(save=False)
        obj.delete()
        
        # Add a success message after deletion
        messages.success(request, f"Successfully deleted the document.")
        return redirect('ocrtesting:document_list')

    # GET request logic is unchanged
    return render(request, 'confirm_delete.html', {
        'object':    obj,
        'doc_type':  doc_type,
    })


@login_required
def create_special_order(request):
    if request.method == 'POST':
        form = SpecialOrderForm(request.POST, request.FILES)
        if form.is_valid():
            so = form.save(commit=False)
            so.fiscal_year = now().year
            so.save()
            form.save_m2m()

            uploaded_files = request.FILES.getlist('files')
            if uploaded_files:
                combined_pdf = combine_files_to_pdf(uploaded_files)
                DocumentFile.objects.create(
                    file=combined_pdf,
                    content_type='application/pdf',
                    special_order=so
                )
            
            # --- ADDED: Log entry ---
            create_log(request.user, "Created Special Order", so)

            return redirect('ocrtesting:view_special_order', pk=so.pk)
    else:
        form = SpecialOrderForm()

    return render(request, 'create_order.html', {
        'form': form, 'recipient_list': Recipient.objects.all(),
        'document_type':  'special_order',
    })


@login_required
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

            new_files = request.FILES.getlist('files')
            clear_files = request.POST.get('clear_files') == 'true'
            
            if new_files or clear_files:
                existing_files = []
                if not clear_files:
                    # Get existing files if we are not clearing them
                    existing_files = [f.file for f in so.documentfile_set.all() if f.file]
                
                # Combine existing (if any) and new files
                all_to_combine = existing_files + new_files
                
                if all_to_combine:
                    combined_pdf = combine_files_to_pdf(all_to_combine)
                    
                    # Delete old records
                    for f in so.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
                    
                    # Create new combined record
                    DocumentFile.objects.create(
                        file=combined_pdf,
                        content_type='application/pdf',
                        special_order=so
                    )
                elif clear_files:
                    # Just delete old records if we cleared and didn't add new
                    for f in so.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
            
            create_log(request.user, "Edited Special Order", so)

            return redirect('ocrtesting:view_special_order', pk=so.pk)
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


@login_required
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


@login_required
def create_travel_order(request):
    if request.method == 'POST':
        form = TravelOrderForm(request.POST, request.FILES) # Use TravelOrderForm
        if form.is_valid():
            to = form.save(commit=False) # Changed variable to 'to' for clarity
            to.fiscal_year = now().year
            to.save()
            form.save_m2m()  # Save recipients

            # Combine uploaded files into exactly ONE DocumentFile with the resulting PDF
            uploaded_files = request.FILES.getlist('files')
            if uploaded_files:
                combined_pdf = combine_files_to_pdf(uploaded_files)
                DocumentFile.objects.create(
                    file=combined_pdf,
                    content_type='application/pdf',
                    travel_order=to,
                )

            # --- ADDED: Create a log entry for the creation action ---
            create_log(request.user, "Created Travel Order", to)
            # Redirect to a 'view_travel_order' URL (we will create this later)
            return redirect('ocrtesting:view_travel_order', pk=to.pk)
    else:
        form = TravelOrderForm() # Use TravelOrderForm

    recipient_list = Recipient.objects.all() 

    # We reuse the exact same template, just pass a different 'document_type'
    return render(request, 'create_order.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'travel_order', # This tells the template to display "Travel Order"
    })




@login_required
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


@login_required
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

            new_files = request.FILES.getlist('files')
            clear_files = request.POST.get('clear_files') == 'true'
            
            if new_files or clear_files:
                existing_files = []
                if not clear_files:
                    # Get existing files if we are not clearing them
                    existing_files = [f.file for f in to.documentfile_set.all() if f.file]
                
                # Combine existing (if any) and new files
                all_to_combine = existing_files + new_files
                
                if all_to_combine:
                    combined_pdf = combine_files_to_pdf(all_to_combine)
                    
                    # Delete old records
                    for f in to.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
                    
                    # Create new combined record
                    DocumentFile.objects.create(
                        file=combined_pdf,
                        content_type='application/pdf',
                        travel_order=to
                    )
                elif clear_files:
                    # Just delete old records if we cleared and didn't add new
                    for f in to.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
            
            # --- ADDED: Log entry ---
            create_log(request.user, "Edited Travel Order", to)
            messages.success(request, "Travel Order updated successfully.")
            # Redirect to the correct view
            return redirect('ocrtesting:view_travel_order', pk=to.pk)
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


 # Protect the view and ensure we have a user
@login_required
def create_memorandum(request, category):
    if category not in ['incoming', 'outgoing']:
        return redirect('some_error_page_or_dashboard') # Your existing logic

    if request.method == 'POST':
        form = MemorandumForm(request.POST, request.FILES)
        if form.is_valid():
            memo = form.save(commit=False)
            memo.fiscal_year = now().year
            memo.save()
            form.save_m2m()

            uploaded_files = request.FILES.getlist('files')
            if uploaded_files:
                combined_pdf = combine_files_to_pdf(uploaded_files)
                DocumentFile.objects.create(
                    file=combined_pdf,
                    content_type='application/pdf',
                    memorandum=memo,
                )
            
            # --- ADDED: Create a log entry for the creation action ---
            create_log(request.user, "Created Memorandum", memo)
            messages.success(request, "Memorandum created successfully.")
            
            return redirect('ocrtesting:view_memorandum', pk=memo.pk)
    else:
        form = MemorandumForm(initial={'category': category})

    recipient_list = Recipient.objects.all()

    return render(request, 'create_memorandum.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'memorandum',
        'category': category,
    })



@login_required
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
 # Protect the view and ensure we have a user
@login_required
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

            new_files = request.FILES.getlist('files')
            clear_files = request.POST.get('clear_files') == 'true'
            
            if new_files or clear_files:
                existing_files = []
                if not clear_files:
                    # Get existing files if we are not clearing them
                    existing_files = [f.file for f in memo.documentfile_set.all() if f.file]
                
                # Combine existing (if any) and new files
                all_to_combine = existing_files + new_files
                
                if all_to_combine:
                    combined_pdf = combine_files_to_pdf(all_to_combine)
                    
                    # Delete old records
                    for f in memo.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
                    
                    # Create new combined record
                    DocumentFile.objects.create(
                        file=combined_pdf,
                        content_type='application/pdf',
                        memorandum=memo
                    )
                elif clear_files:
                    # Just delete old records if we cleared and didn't add new
                    for f in memo.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
            
            # --- ADDED: Create a log entry for the edit action ---
            create_log(request.user, "Edited Memorandum", memo)
            messages.success(request, "Memorandum updated successfully.")
            
            return redirect('ocrtesting:view_memorandum', pk=memo.pk)
    else:
        form = MemorandumForm(instance=memo)

    existing_recipient_ids = list(memo.recipients.values_list('id', flat=True))
    existing_files = [{'url': f.file.url, 'name': os.path.basename(f.file.name), 'content_type': f.content_type} for f in memo.documentfile_set.all()]
    
    return render(request, 'edit_memorandum.html', {
        'form': form,
        'document': memo,
        'recipient_list': Recipient.objects.all(),
        'existing_recipient_ids': existing_recipient_ids,
        'existing_files_json': json.dumps(existing_files),
    })


@login_required
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
            uploaded_files = request.FILES.getlist('files')
            if uploaded_files:
                combined_pdf = combine_files_to_pdf(uploaded_files)
                DocumentFile.objects.create(
                    file=combined_pdf,
                    content_type='application/pdf',
                    moau=moau,
                )
            
            # --- ADDED: Log entry ---
            create_log(request.user, "Created MOAU", moau)
            # Redirect to the view page upon success
            return redirect('ocrtesting:view_moau', pk=moau.pk)
    else:
        form = MOAUForm()

    recipient_list = Recipient.objects.all()
    # We no longer need to pass the signatory_list as it's a text input now.

    return render(request, 'create_moau.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'moau',
    })



@login_required
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
@login_required
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

            new_files = request.FILES.getlist('files')
            clear_files = request.POST.get('clear_files') == 'true'
            
            if new_files or clear_files:
                existing_files = []
                if not clear_files:
                    # Get existing files if we are not clearing them
                    existing_files = [f.file for f in moau.documentfile_set.all() if f.file]
                
                # Combine existing (if any) and new files
                all_to_combine = existing_files + new_files
                
                if all_to_combine:
                    combined_pdf = combine_files_to_pdf(all_to_combine)
                    
                    # Delete old records
                    for f in moau.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
                    
                    # Create new combined record
                    DocumentFile.objects.create(
                        file=combined_pdf,
                        content_type='application/pdf',
                        moau=moau
                    )
                elif clear_files:
                    # Just delete old records if we cleared and didn't add new
                    for f in moau.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
            
            # --- ADDED: Log entry ---
            create_log(request.user, "Edited MOAU", moau)

            return redirect('ocrtesting:view_moau', pk=moau.pk)
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


@login_required
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

            uploaded_files = request.FILES.getlist('files')
            if uploaded_files:
                combined_pdf = combine_files_to_pdf(uploaded_files)
                DocumentFile.objects.create(
                    file=combined_pdf,
                    content_type='application/pdf',
                    communication_letter=comm_letter,
                )

            # --- ADDED: Log entry ---
            create_log(request.user, "Created Communication Letter", comm_letter)

            # You will need to create 'view_communication_letter' later
            return redirect('ocrtesting:view_communication_letter', pk=comm_letter.pk) # Placeholder redirect
    else:
        # Pre-fill the form with the category from the URL
        form = CommunicationLetterForm(initial={'category': category})

    recipient_list = Recipient.objects.all()

    # Render the new, dedicated Communication Letter template
    return render(request, 'create_communication_letter.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'communication_letter',
        'category': category, # Pass category for JS and hidden fields
    })


@login_required
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


@login_required
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

            new_files = request.FILES.getlist('files')
            clear_files = request.POST.get('clear_files') == 'true'
            
            if new_files or clear_files:
                existing_files = []
                if not clear_files:
                    # Get existing files if we are not clearing them
                    existing_files = [f.file for f in comm_letter.documentfile_set.all() if f.file]
                
                # Combine existing (if any) and new files
                all_to_combine = existing_files + new_files
                
                if all_to_combine:
                    combined_pdf = combine_files_to_pdf(all_to_combine)
                    
                    # Delete old records
                    for f in comm_letter.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
                    
                    # Create new combined record
                    DocumentFile.objects.create(
                        file=combined_pdf,
                        content_type='application/pdf',
                        communication_letter=comm_letter
                    )
                elif clear_files:
                    # Just delete old records if we cleared and didn't add new
                    for f in comm_letter.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
            
            # --- ADDED: Log entry ---
            create_log(request.user, "Edited Communication Letter", comm_letter)
            messages.success(request, "Communication Letter updated successfully.")

            return redirect('ocrtesting:view_communication_letter', pk=comm_letter.pk)
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


@login_required
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

            uploaded_files = request.FILES.getlist('files')
            if uploaded_files:
                combined_pdf = combine_files_to_pdf(uploaded_files)
                DocumentFile.objects.create(
                    file=combined_pdf,
                    content_type='application/pdf',
                    other_document=other_doc,
                )
            # --- ADDED: Log entry ---
            create_log(request.user, "Created Other Document", other_doc)
            messages.success(request, "Document created successfully.")
            # You will need to create 'view_other_document' later
            return redirect('ocrtesting:view_other_document', pk=other_doc.pk) # Placeholder redirect
    else:
        form = OtherDocumentForm()

    recipient_list = Recipient.objects.all()

    # Render the new, dedicated Other Document template
    return render(request, 'create_other_document.html', {
        'form': form,
        'recipient_list': recipient_list,
        'document_type': 'other document', # For consistency
    })



@login_required
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


@login_required
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

            new_files = request.FILES.getlist('files')
            clear_files = request.POST.get('clear_files') == 'true'
            
            if new_files or clear_files:
                existing_files = []
                if not clear_files:
                    # Get existing files if we are not clearing them
                    existing_files = [f.file for f in other_doc.documentfile_set.all() if f.file]
                
                # Combine existing (if any) and new files
                all_to_combine = existing_files + new_files
                
                if all_to_combine:
                    combined_pdf = combine_files_to_pdf(all_to_combine)
                    
                    # Delete old records
                    for f in other_doc.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
                    
                    # Create new combined record
                    DocumentFile.objects.create(
                        file=combined_pdf,
                        content_type='application/pdf',
                        other_document=other_doc
                    )
                elif clear_files:
                    # Just delete old records if we cleared and didn't add new
                    for f in other_doc.documentfile_set.all():
                        if f.file and os.path.isfile(os.path.join(settings.MEDIA_ROOT, f.file.name)):
                            os.remove(os.path.join(settings.MEDIA_ROOT, f.file.name))
                        f.delete()
            # --- ADDED: Log entry ---
            create_log(request.user, "Edited Other Document", other_doc)
            
            return redirect('ocrtesting:view_other_document', pk=other_doc.pk)
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
# --- REFACTORED HELPER FUNCTION TO GET DATA ---
def _get_filtered_documents(request):
    """
    Contains all the logic to fetch and filter documents from multiple models.
    UPDATED: Now properly handles documents filtered by dashboard clicks,
    including all time, year-only, all months for a year, and preserves existing functionality.
    FIXED: All Types filter now shows all documents across all time periods.
    """
    today = timezone.localdate()
    params = request.GET

    # 1. Get Filter Parameters
    doc_type = params.get('doc_type', '')
    category = params.get('category', '')
    recipient_id = params.get('recipient', '')
    search_query = params.get('search', '')
    start_date_str = params.get('start_date', '')
    end_date_str = params.get('end_date', '')

    # 2. Enhanced Date Filtering Logic
    start_date = None
    end_date = None
    
    # Check for all time filter first
    if params.get('all_time') == 'true':
        # No date filtering for all time
        pass
    # If explicit date range is provided, use it
    elif start_date_str and end_date_str:
        try:
            start_date = date.fromisoformat(start_date_str)
            end_date = date.fromisoformat(end_date_str)
        except (ValueError, TypeError):
            pass
    # Check for year and month parameters (from dashboard)
    else:
        year_str = params.get('year')
        month_str = params.get('month')
        
        if year_str and month_str == 'all':
            # All months for a specific year
            try:
                year = int(year_str)
                start_date = date(year, 1, 1)
                end_date = date(year, 12, 31)
            except (ValueError, TypeError):
                pass
        elif year_str and month_str and month_str != 'all':
            # Both year and specific month provided
            try:
                year = int(year_str)
                month = int(month_str)
                start_date = date(year, month, 1)
                _, last_day = calendar.monthrange(year, month)
                end_date = date(year, month, last_day)
            except (ValueError, TypeError):
                pass
        elif year_str and not month_str:
            # Only year provided (year-only filter)
            try:
                year = int(year_str)
                start_date = date(year, 1, 1)
                end_date = date(year, 12, 31)
            except (ValueError, TypeError):
                pass
        else:
            # Check if this is a fresh page load with no meaningful filters
            # Only default to current month if NO filters are applied AND no URL parameters exist
            # that indicate user interaction (like clicking "All Types")
            has_filter_params = any([
                params.get('doc_type') is not None,  # Check if doc_type parameter exists (even if empty)
                category, recipient_id, search_query, 
                start_date_str, end_date_str, year_str, month_str
            ])
            
            if not has_filter_params:
                # Fresh page load with no filters - default to current month
                start_date = today.replace(day=1)
                _, last_day = calendar.monthrange(today.year, today.month)
                end_date = today.replace(day=last_day)
            # If has_filter_params is True but no date filters, show all documents (no date restriction)
    
    # 3. Query and Combine Documents
    models_to_query = {
        'special_order': SpecialOrder, 
        'travel_order': TravelOrder, 
        'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 
        'moau': MOAU, 
        'other_document': OtherDocument,
    }
    
    # If doc_type is specified, only query that specific model
    models_to_fetch = {doc_type: models_to_query[doc_type]} if doc_type and doc_type in models_to_query else models_to_query

    querysets = []
    for type_slug, model_class in models_to_fetch.items():
        # Start with non-archived documents
        qs = model_class.objects.prefetch_related('recipients').filter(is_archived=False)

        # Apply date filters (if any)
        if start_date: 
            qs = qs.filter(date__gte=start_date)
        if end_date: 
            qs = qs.filter(date__lte=end_date)
            
        # Apply other filters
        if recipient_id: 
            qs = qs.filter(recipients__id=recipient_id)
            
        # FIXED: Category filter - only apply to models that have category field
        # and skip models that don't have category when category filter is active
        if category:
            if hasattr(model_class, 'category'):
                qs = qs.filter(category=category)
            else:
                # Skip this model entirely if category filter is applied but model doesn't have category field
                continue
                
        if search_query:
            search_fields = Q(subject__icontains=search_query)
            if hasattr(model_class, 'document_id'): 
                search_fields |= Q(document_id__icontains=search_query)
            if hasattr(model_class, 'remarks'): 
                search_fields |= Q(remarks__icontains=search_query)
            qs = qs.filter(search_fields)

        # Process each document for URL generation and display
        for doc in qs:
            pk_as_string = str(doc.pk)
            doc.doc_type_slug = type_slug
            doc.doc_type_display = type_slug.replace('_', ' ').title()
            
            # Generate URLs
            view_kwargs = {'pk': pk_as_string}
            edit_kwargs = {'pk': pk_as_string}
            
            doc.view_url = reverse(f'ocrtesting:view_{type_slug}', kwargs=view_kwargs)
            doc.edit_url = reverse(f'ocrtesting:edit_{type_slug}', kwargs=edit_kwargs)
            delete_slug = type_slug.replace('_', '-')
            doc.delete_url = reverse('ocrtesting:delete_document', kwargs={'doc_type': delete_slug, 'pk': pk_as_string})

            # Format recipient display
            recipient_names = [r.name for r in doc.recipients.all()]
            doc.recipient_list_display = ", ".join(recipient_names) if recipient_names else "None"
            
            # Only add documents with valid dates
            if doc.date:
                querysets.append(doc)

    # 4. Sort and generate title
    sorted_documents = sorted(querysets, key=lambda x: x.date, reverse=True)
    page_title = _generate_document_list_title(params, today)
    
    return sorted_documents, page_title


# This is a helper function to generate the dynamic page title.
def _generate_document_list_title(params, today):
    """
    Generates a dynamic title based on the applied filters.
    UPDATED: Enhanced to handle dashboard-originated filters including "All Months".
    FIXED: Now properly handles "All Types" filter to not show current month.
    """
    doc_type = params.get('doc_type')
    category = params.get('category')
    recipient = params.get('recipient')
    start_date_str = params.get('start_date')
    end_date_str = params.get('end_date')
    search = params.get('search')
    year_str = params.get('year')
    month_str = params.get('month')

    if search:
        return f"Search results for: '{search}'"

    title_parts = []
    
    # Category first (for incoming/outgoing from pie chart clicks)
    if category:
        title_parts.append(category.title())

    # Then Document Type (from card clicks)
    if doc_type:
        # Convert underscore to space and handle pluralization
        type_name = doc_type.replace('_', ' ').title()
        if type_name.endswith('y'):
            # Handle cases like "MOAU" -> "MOAUs"
            title_parts.append(f"{type_name}s")
        elif type_name == "Moau":
            title_parts.append("MOAUs")
        else:
            title_parts.append(f"{type_name}s")
    else:
        title_parts.append("All Documents")
        
    # Then Date Range
    if start_date_str and end_date_str:
        try:
            start_date = date.fromisoformat(start_date_str)
            end_date = date.fromisoformat(end_date_str)
            
            # Check if the range exactly matches a full month
            import calendar
            if (start_date.day == 1 and 
                end_date.day == calendar.monthrange(end_date.year, end_date.month)[1] and
                start_date.year == end_date.year and 
                start_date.month == end_date.month):
                # It's a full month range
                title_parts.append(f"for {start_date.strftime('%B %Y')}")
            # Check if it's a full year range
            elif (start_date.month == 1 and start_date.day == 1 and
                  end_date.month == 12 and end_date.day == 31 and
                  start_date.year == end_date.year):
                # It's a full year range
                title_parts.append(f"for {start_date.year}")
            elif start_date == end_date:
                title_parts.append(f"on {start_date.strftime('%B %d, %Y')}")
            else:
                title_parts.append(f"from {start_date.strftime('%b %d, %Y')} to {end_date.strftime('%b %d, %Y')}")
        except (ValueError, TypeError):
            pass
    elif year_str and month_str == 'all':
        # All months for a specific year (from "All Months" selection)
        try:
            year = int(year_str)
            title_parts.append(f"for {year}")
        except (ValueError, TypeError):
            pass
    elif year_str and month_str and month_str != 'all':
        # Specific year and month
        try:
            year = int(year_str)
            month = int(month_str)
            month_names = [
                'January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'
            ]
            title_parts.append(f"for {month_names[month - 1]} {year}")
        except (ValueError, TypeError, IndexError):
            pass
    elif start_date_str:
        try:
            start_date = date.fromisoformat(start_date_str)
            title_parts.append(f"since {start_date.strftime('%B %d, %Y')}")
        except (ValueError, TypeError):
            pass
    elif end_date_str:
        try:
            end_date = date.fromisoformat(end_date_str)
            title_parts.append(f"until {end_date.strftime('%B %d, %Y')}")
        except (ValueError, TypeError):
            pass
    else:
        # Check if this is a fresh page load with no meaningful filters
        # Only show current month if NO filters are applied AND no URL parameters exist
        has_filter_params = any([
            params.get('doc_type') is not None,  # Check if doc_type parameter exists (even if empty)
            category, recipient, search, 
            start_date_str, end_date_str, year_str, month_str
        ])
        
        if not has_filter_params:
            # Fresh page load with no filters - show current month in title
            title_parts.append(f"for {today.strftime('%B %Y')}")
        # If has_filter_params is True but no date filters, don't add date to title
        # This handles the "All Types" case - just shows "Documents" without date

    return " ".join(title_parts)


def _paginate(request, object_list, *, per_page=10):
    paginator = Paginator(object_list, per_page)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    query_params = request.GET.copy()
    query_params.pop('page', None)
    querystring = query_params.urlencode()

    if hasattr(paginator, 'get_elided_page_range'):
        page_range = list(paginator.get_elided_page_range(page_obj.number))
    else:
        page_range = list(paginator.page_range)

    return page_obj, page_range, querystring


# --- UPDATED MAIN VIEW ---
@login_required
def view_documents(request):
    """
    Displays the document list page. Now uses the helper to get data.
    """
    sorted_documents, page_title = _get_filtered_documents(request)

    documents_page, page_range, querystring = _paginate(request, sorted_documents, per_page=10)
    
    context = {
        'documents': documents_page,
        'page_obj': documents_page,
        'page_range': page_range,
        'querystring': querystring,
        'page_title': page_title,
        'all_recipients': Recipient.objects.all().order_by('name'),
        'filter_values': request.GET, # Pass the entire GET dictionary
    }
    
    return render(request, 'document_list.html', context)


def export_documents_pdf(request):
    documents, page_title = _get_filtered_documents(request)
    
    output = PdfWriter()
    styles = getSampleStyleSheet()
    
    PAGE_WIDTH, PAGE_HEIGHT = letter
    MARGIN = 0.5 * inch
    
    # Create title/summary page
    packet = io.BytesIO()
    overlay = canvas.Canvas(packet, pagesize=letter)
    
    # Load template for title page
    template_path = "images/logo/template.pdf"
    try:
        template_reader = PdfReader(template_path)
        template_page = template_reader.pages[0]
    except:
        template_page = None
    
    # Add title
    overlay.setFont("Helvetica-Bold", 16)
    overlay.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 1 * inch, page_title)
    
    # Add report summary
    overlay.setFont("Helvetica", 10)
    overlay.drawString(MARGIN, PAGE_HEIGHT - 1.5 * inch, f"Report Date: {date.today().strftime('%B %d, %Y')}")
    overlay.drawString(MARGIN, PAGE_HEIGHT - 1.8 * inch, f"Total Documents: {len(documents)}")
    
    # Build summary table
    data = [["Type", "Subject", "Doc ID", "Recipients", "Doc Date"]]
    for d in documents:
        doc_id_text = d.document_id if hasattr(d, 'document_id') and d.document_id else 'N/A'
        data.append([
            Paragraph(d.doc_type_display, styles['Normal']),
            Paragraph(d.subject or '(No Subject)', styles['Normal']),
            doc_id_text,
            Paragraph(d.recipient_list_display, styles['Normal']),
            d.date.strftime('%b. %d, %Y')
        ])
    
    usable_width = PAGE_WIDTH - (2 * MARGIN)
    col_widths = [
        1.5 * inch,  # Type
        1.8 * inch,  # Subject
        0.6 * inch,  # Doc ID
        2.0 * inch,  # Recipients
        0.9 * inch   # Date
    ]
    total_width = sum(col_widths)
    if total_width > usable_width:
        scale = usable_width / total_width
        col_widths = [w * scale for w in col_widths]
    
    table = Table(data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4A5568')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F7FAFC')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9FAFB')]),
    ]))
    
    table_width, table_height = table.wrap(usable_width, 0)
    y_position = PAGE_HEIGHT - 2.3 * inch
    table.drawOn(overlay, MARGIN, y_position - table_height)
    
    overlay.save()
    packet.seek(0)
    
    # Add title page to output
    overlay_reader = PdfReader(packet)
    if template_page:
        template_page.merge_page(overlay_reader.pages[0])
        output.add_page(template_page)
    else:
        output.add_page(overlay_reader.pages[0])
    
    # Add detailed page for each document
    for idx, doc in enumerate(documents, 1):
        packet = io.BytesIO()
        doc_canvas = canvas.Canvas(packet, pagesize=letter)
        
        # Page header
        doc_canvas.setFont("Helvetica-Bold", 12)
        doc_canvas.drawString(MARGIN, PAGE_HEIGHT - 0.5 * inch, f"Document {idx} of {len(documents)}")
        
        # Document details
        y = PAGE_HEIGHT - 1 * inch
        doc_canvas.setFont("Helvetica-Bold", 10)
        doc_canvas.drawString(MARGIN, y, "DOCUMENT DETAILS")
        
        y -= 0.25 * inch
        doc_canvas.setFont("Helvetica", 9)
        
        detail_data = []
        detail_data.append(["Type:", doc.doc_type_display])
        detail_data.append(["Document ID:", doc.document_id if hasattr(doc, 'document_id') and doc.document_id else 'N/A'])
        detail_data.append(["Subject:", doc.subject or '(No Subject)'])
        detail_data.append(["Date:", doc.date.strftime('%B %d, %Y')])
        
        # Add category if available
        if hasattr(doc, 'category'):
            detail_data.append(["Category:", doc.category.title()])
        
        # Add from/for fields if available
        if hasattr(doc, 'from_field') and doc.from_field:
            detail_data.append(["From:", doc.from_field])
        if hasattr(doc, 'for_field') and doc.for_field:
            detail_data.append(["For:", doc.for_field])
        if hasattr(doc, 'thru') and doc.thru:
            detail_data.append(["Thru:", doc.thru])
        if hasattr(doc, 'to_or_thru') and doc.to_or_thru:
            detail_data.append(["To/Thru:", doc.to_or_thru])
        
        # Add MOAU specific fields
        if hasattr(doc, 'first_party_agency') and doc.first_party_agency:
            detail_data.append(["First Party:", f"{doc.first_party_agency} - {doc.first_party_representative or 'N/A'}"])
        if hasattr(doc, 'second_party_agency') and doc.second_party_agency:
            detail_data.append(["Second Party:", f"{doc.second_party_agency} - {doc.second_party_representative or 'N/A'}"])
        
        # Add OtherDocument type if available
        if hasattr(doc, 'doc_type') and isinstance(doc, OtherDocument):
            detail_data.append(["Document Type:", doc.doc_type])
        
        detail_data.append(["Recipients:", doc.recipient_list_display or '(None)'])
        detail_data.append(["Received By:", getattr(doc, 'received_by', '') or '(Not specified)'])
        
        if hasattr(doc, 'approved_date') and doc.approved_date:
            detail_data.append(["Approved Date:", doc.approved_date.strftime('%B %d, %Y')])
        
        # Create detail table
        detail_table = Table(detail_data, colWidths=[1.5 * inch, 4.5 * inch])
        detail_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.white, colors.HexColor('#F9FAFB')]),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ]))
        
        detail_table_width, detail_table_height = detail_table.wrap(usable_width, 0)
        detail_table.drawOn(doc_canvas, MARGIN, y - detail_table_height - 0.3 * inch)
        
        # Add remarks if available
        remarks_y = y - detail_table_height - 0.8 * inch
        if hasattr(doc, 'remarks') and doc.remarks:
            doc_canvas.setFont("Helvetica-Bold", 10)
            doc_canvas.drawString(MARGIN, remarks_y, "REMARKS")
            doc_canvas.setFont("Helvetica", 9)
            
            # Word wrap remarks
            remarks_text = doc.remarks
            remark_lines = []
            max_chars = 90
            for i in range(0, len(remarks_text), max_chars):
                remark_lines.append(remarks_text[i:i + max_chars])
            
            remarks_y -= 0.25 * inch
            for line in remark_lines[:5]:  # Limit to 5 lines
                doc_canvas.drawString(MARGIN + 0.2 * inch, remarks_y, line)
                remarks_y -= 0.2 * inch
        
        # Add page number
        doc_canvas.setFont("Helvetica", 8)
        doc_canvas.drawRightString(PAGE_WIDTH - MARGIN, MARGIN * 0.5, f"Page {idx + 1} of {len(documents) + 1}")
        
        doc_canvas.save()
        packet.seek(0)
        
        doc_reader = PdfReader(packet)
        output.add_page(doc_reader.pages[0])
    
    final_pdf = io.BytesIO()
    output.write(final_pdf)
    final_pdf.seek(0)
    
    filename = f"{slugify(page_title)}.pdf"
    response = HttpResponse(final_pdf, content_type="application/pdf")
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
    Generates a single combined PDF in memory from ALL files attached to a document.

    - Merges uploaded PDFs (in order)
    - Converts uploaded images to PDF pages and appends them (in order)

    Returns:
        io.BytesIO: A byte stream of the generated PDF, or None if no file content.
    """
    doc_files = list(
        document_instance.documentfile_set.exclude(file__isnull=True).order_by('order', 'uploaded_at', 'id')
    )
    if not doc_files:
        return None

    image_exts = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.webp'}

    def _is_pdf(doc_file):
        ct = (getattr(doc_file, 'content_type', '') or '').lower()
        if ct == 'application/pdf' or ct.endswith('/pdf') or ct.endswith('pdf'):
            return True
        name = (getattr(getattr(doc_file, 'file', None), 'name', '') or '').lower()
        return name.endswith('.pdf')

    def _is_image(doc_file):
        ct = (getattr(doc_file, 'content_type', '') or '').lower()
        if ct.startswith('image/'):
            return True
        name = (getattr(getattr(doc_file, 'file', None), 'name', '') or '').lower()
        _, ext = os.path.splitext(name)
        return ext in image_exts

    def _read_bytes(field_file):
        field_file.open('rb')
        try:
            return field_file.read()
        finally:
            try:
                field_file.close()
            except Exception:
                pass

    writer = PdfWriter()
    kept_readers = []
    kept_streams = []
    pages_added = 0

    for doc_file in doc_files:
        if not doc_file.file:
            continue

        try:
            raw = _read_bytes(doc_file.file)
        except Exception:
            continue

        if not raw:
            continue

        if _is_pdf(doc_file):
            try:
                stream = io.BytesIO(raw)
                kept_streams.append(stream)
                reader = PdfReader(stream)
                kept_readers.append(reader)
                for page in reader.pages:
                    writer.add_page(page)
                    pages_added += 1
            except Exception:
                continue

        elif _is_image(doc_file):
            try:
                with Image.open(io.BytesIO(raw)) as img:
                    img = img.convert('RGB')
                    img_pdf = io.BytesIO()
                    img.save(img_pdf, format='PDF', resolution=100.0)

                img_pdf.seek(0)
                kept_streams.append(img_pdf)
                reader = PdfReader(img_pdf)
                kept_readers.append(reader)
                for page in reader.pages:
                    writer.add_page(page)
                    pages_added += 1
            except Exception:
                continue

    if pages_added == 0:
        return None

    pdf_stream = io.BytesIO()
    writer.write(pdf_stream)
    pdf_stream.seek(0)
    return pdf_stream


# --- 3. GENERIC VIEW FUNCTIONS ---

def download_document_pdf(request, doc_type, pk):
    """
    Handles downloading a document as a PDF.

    - If the document has exactly one uploaded PDF, it serves that file directly.
    - Otherwise, it generates a single combined PDF from all attached PDFs + images.
    - The downloaded filename is always the document's document_id.
    """
    model_class = MODEL_MAP.get(doc_type)
    if not model_class:
        return HttpResponse("Invalid document type specified.", status=400)
    
    try:
        # 1. Look up the model and fetch the document instance
        document = get_object_or_404(model_class, pk=pk)

        # 2. Filename is ALWAYS the document_id (fallback to pk)
        raw_doc_id = getattr(document, 'document_id', None) or str(getattr(document, 'pk', 'document'))
        safe_doc_id = re.sub(r'[^A-Za-z0-9._-]+', '_', str(raw_doc_id)).strip('._-') or 'document'
        filename = f"{safe_doc_id}.pdf"

        # 3. If there's exactly one PDF file attached, serve it directly.
        attached_files = list(
            document.documentfile_set.exclude(file__isnull=True).order_by('order', 'uploaded_at', 'id')
        )

        if len(attached_files) == 1 and attached_files[0].file:
            ct = (attached_files[0].content_type or '').lower()
            name = (getattr(attached_files[0].file, 'name', '') or '').lower()
            if ct == 'application/pdf' or name.endswith('.pdf'):
                return FileResponse(attached_files[0].file, as_attachment=True, filename=filename)

        # 4. Otherwise, build a combined PDF from PDFs + images.
        pdf_stream = generate_document_pdf(document)
        if pdf_stream:
            return FileResponse(pdf_stream, as_attachment=True, filename=filename)

        return HttpResponse("This document has no file content to download.", status=404)

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
            subject=f"A new {doc_type.replace('_', ' ').title()} for you!",
            body='Please find the attached document.',
            to=emails,
        )

        # 4. Attachment filename is ALWAYS the document_id (fallback to pk)
        raw_doc_id = getattr(document, 'document_id', None) or str(getattr(document, 'pk', 'document'))
        safe_doc_id = re.sub(r'[^A-Za-z0-9._-]+', '_', str(raw_doc_id)).strip('._-') or 'document'
        filename = f"{safe_doc_id}.pdf"

        # 5. Attach either the single PDF, or a combined PDF built from PDFs + images.
        attached_files = list(
            document.documentfile_set.exclude(file__isnull=True).order_by('order', 'uploaded_at', 'id')
        )

        if len(attached_files) == 1 and attached_files[0].file:
            ct = (attached_files[0].content_type or '').lower()
            name = (getattr(attached_files[0].file, 'name', '') or '').lower()
            if ct == 'application/pdf' or name.endswith('.pdf'):
                attached_files[0].file.open('rb')
                try:
                    email.attach(filename, attached_files[0].file.read(), 'application/pdf')
                finally:
                    try:
                        attached_files[0].file.close()
                    except Exception:
                        pass
                success_message = 'Email with attached PDF sent successfully!'
            else:
                pdf_stream = generate_document_pdf(document)
                if not pdf_stream:
                    return JsonResponse({'error': 'No content to send. The document has no uploaded PDF or images.'}, status=400)
                email.attach(filename, pdf_stream.getvalue(), 'application/pdf')
                success_message = 'Email with generated PDF sent successfully!'
        else:
            pdf_stream = generate_document_pdf(document)
            if not pdf_stream:
                return JsonResponse({'error': 'No content to send. The document has no uploaded PDF or images.'}, status=400)
            email.attach(filename, pdf_stream.getvalue(), 'application/pdf')
            success_message = 'Email with combined PDF sent successfully!'

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
    FIXED: Now properly handles all cases without breaking date import
    """
    from datetime import date
    import calendar
    from django.utils import timezone
    
    start_date_str = params.get('start_date')
    end_date_str = params.get('end_date')
    
    # Handle all_time parameter
    if params.get('all_time') == 'true':
        # Return a very wide date range to capture all documents
        return date(1900, 1, 1), date(2100, 12, 31)
    
    # If a date range is provided, use it.
    if start_date_str and end_date_str:
        try:
            start = date.fromisoformat(start_date_str)
            end = date.fromisoformat(end_date_str)
            return start, end
        except (ValueError, TypeError):
            # Fallback in case of invalid date format
            pass

    # Handle year with "all months" case
    year_str = params.get('year')
    month_str = params.get('month')
    
    if year_str and month_str == 'all':
        try:
            year = int(year_str)
            return date(year, 1, 1), date(year, 12, 31)
        except (ValueError, TypeError):
            pass
    
    # Otherwise, fall back to month and year.
    today = timezone.localdate()
    
    try:
        y = int(year_str) if year_str else today.year
        m = int(month_str) if month_str and month_str != 'all' else today.month
    except (ValueError, TypeError):
        y = today.year
        m = today.month
    
    month_start = date(y, m, 1)
    last_day = calendar.monthrange(y, m)[1]
    month_end = date(y, m, last_day)
    return month_start, month_end




# --- UPDATED DASHBOARD VIEW ---
@login_required
def dashboard(request):
    """
    Renders the dashboard page and provides default filter values.
    """
    from django.utils import timezone
    import calendar
    
    today = timezone.localdate()
    
    # Provide lists for dropdowns and set the default to All Months / All Time.
    context = {
        'default_month': 'all',
        'default_year': '',
        'month_list': [(i, calendar.month_name[i]) for i in range(1, 13)],
        # Generate a list of recent years for the dropdown
        'year_list': list(range(today.year, today.year - 10, -1)),
    }
    return render(request, 'dashboard.html', context)


# --- UPDATED API VIEW FOR CARDS ---
# --- UPDATED API VIEW FOR CARDS ---
def dashboard_cards_data(request):
    """
    JSON response with total + counts per document_type for the selected period.
    FIXED: Now uses document date instead of upload_date
    """
    start, end = _get_period_range(request.GET)
    
    models_to_count = {
        'memorandum': Memorandum, 
        'travel_order': TravelOrder, 
        'special_order': SpecialOrder,
        'communication_letter': CommunicationLetter, 
        'moau': MOAU, 
        'other_document': OtherDocument,
    }
    
    counts = {}
    total = 0
    for name, model in models_to_count.items():
        # FIXED: Changed from upload_date to date
        count = model.objects.filter(date__gte=start, date__lte=end).count()
        counts[name] = count
        total += count
        
    return JsonResponse({'total': total, 'counts': counts})


# --- UPDATED API VIEW FOR PIE CHART ---
# FIXED: dashboard_direction_data function
def dashboard_direction_data(request):
    """
    JSON counts of incoming vs outgoing for the selected period.
    FIXED: Now only counts documents that actually have category field
    and returns totals that match the pie chart display logic
    """
    start, end = _get_period_range(request.GET)
    
    # Only count models that have the category field
    models_with_category = [Memorandum, CommunicationLetter]
    
    incoming = 0
    outgoing = 0
    total_categorizable = 0  # Total docs that can be categorized
    
    for model in models_with_category:
        model_incoming = model.objects.filter(date__gte=start, date__lte=end, category='incoming').count()
        model_outgoing = model.objects.filter(date__gte=start, date__lte=end, category='outgoing').count()
        
        incoming += model_incoming
        outgoing += model_outgoing
        total_categorizable += model_incoming + model_outgoing
    
    return JsonResponse({
        'incoming': incoming, 
        'outgoing': outgoing,
        'total_categorizable': total_categorizable  # Optional: for debugging
    })





#RECIPIENT MANAGEMENT
def manage_recipients(request):
    """
    Displays a list of all recipients with options to add, edit, or delete.
    """
    recipients = Recipient.objects.all().order_by('name')
    recipients_page, page_range, querystring = _paginate(request, recipients, per_page=10)
    context = {
        'recipients': recipients_page,
        'page_obj': recipients_page,
        'page_range': page_range,
        'querystring': querystring,
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
            messages.success(request, 'Recipient added successfully.')
            return redirect('ocrtesting:manage_recipients')
    else:
        form = RecipientForm()
    
    context = {
        'form': form,
        'form_title': 'Add New Recipient',
        'button_text': 'Save Recipient'
    }
    return render(request, 'recipient_form.html', context)


@login_required
def edit_recipient(request, pk):
    """
    Handles editing an existing recipient.
    """
    recipient = get_object_or_404(Recipient, pk=pk)
    if request.method == 'POST':
        form = RecipientForm(request.POST, instance=recipient)
        if form.is_valid():
            form.save()
            messages.success(request, 'Recipient updated successfully.')
            return redirect('ocrtesting:manage_recipients')
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
    messages.success(request, 'Recipient deleted successfully.')
    return redirect('ocrtesting:manage_recipients')




# --- HELPER FUNCTION TO GET THE GOOGLE DRIVE SERVICE ---
def get_gdrive_service():
    """
    Authenticates using the Service Account credentials and returns a
    Google Drive service object.
    """
    SCOPES = ['https://www.googleapis.com/auth/drive']
    try:
        creds = service_account.Credentials.from_service_account_file(
            settings.GOOGLE_DRIVE_SERVICE_ACCOUNT_KEY_PATH, scopes=SCOPES
        )
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        print(f"ERROR: Could not build Google Drive service. Check settings.py and credentials file. Details: {e}")
        raise

# --- ARCHIVE ACTION VIEW ---
# --- THIS IS THE FINAL, CORRECTED ARCHIVE FUNCTION ---

@require_POST
def archive_documents(request):
    """
    Handles the batch archiving of selected documents using a Service Account.
    Includes the 'supportsAllDrives=True' flag to work with shared folders.
    """
    document_pks_by_type = {}
    for key, value in request.POST.items():
        if key.startswith('doc_'):
            parts = value.rsplit('_', 1)
            if len(parts) == 2:
                doc_type_slug, pk = parts
                document_pks_by_type.setdefault(doc_type_slug, []).append(pk)
    
    if not document_pks_by_type:
        messages.warning(request, 'No documents were selected to archive.')
        return redirect('ocrtesting:manage_archive')

    models_map = {
        'special_order': SpecialOrder, 'travel_order': TravelOrder, 'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }

    try:
        service = get_gdrive_service()
    except Exception as e:
        messages.error(request, f"Could not connect to Google Drive. Please check server configuration.")
        return redirect('ocrtesting:manage_archive')

    archived_count, error_count = 0, 0
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
                        
                        # --- THIS IS THE CRITICAL FIX ---
                        gdrive_file = service.files().create(
                            body=file_metadata,
                            media_body=media,
                            fields='id',
                            supportsAllDrives=True  # <-- THIS FLAG IS REQUIRED
                        ).execute()
                        
                        doc_file.gdrive_file_id = gdrive_file.get('id')
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

    return redirect('ocrtesting:manage_archive')

# --- THIS IS THE FINAL, CORRECTED RESTORE FUNCTION ---

@require_POST
def restore_documents(request):
    """
    Handles the batch restoration of selected documents using a Service Account.
    Includes the 'supportsAllDrives=True' flag to work with shared folders.
    """
    document_pks_by_type = {}
    for key, value in request.POST.items():
        if key.startswith('doc_'):
            parts = value.rsplit('_', 1)
            if len(parts) == 2:
                doc_type_slug, pk = parts
                document_pks_by_type.setdefault(doc_type_slug, []).append(pk)

    if not document_pks_by_type:
        messages.warning(request, 'No documents were selected to restore.')
        return redirect(f"{reverse('ocrtesting:manage_archive')}?mode=archived")

    models_map = {
        'special_order': SpecialOrder, 'travel_order': TravelOrder, 'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }
    
    try:
        service = get_gdrive_service()
    except Exception as e:
        messages.error(request, f"Could not connect to Google Drive. Please check server configuration.")
        return redirect(f"{reverse('ocrtesting:manage_archive')}?mode=archived")
        
    restored_count, error_count = 0, 0
    for doc_type_slug, pks in document_pks_by_type.items():
        model_class = models_map.get(doc_type_slug)
        if not model_class: continue

        documents_to_restore = model_class.objects.filter(pk__in=pks, is_archived=True)
        for document in documents_to_restore:
            all_files_successful = True
            for doc_file in document.documentfile_set.all():
                if doc_file.gdrive_file_id:
                    try:
                        # Get, Get_Media, and Download do not need the flag, but Delete does.
                        file_metadata = service.files().get(fileId=doc_file.gdrive_file_id, fields='name').execute()
                        original_filename = file_metadata.get('name', f"{uuid.uuid4()}")

                        gdrive_request = service.files().get_media(fileId=doc_file.gdrive_file_id)
                        fh = io.BytesIO()
                        downloader = MediaIoBaseDownload(fh, gdrive_request)
                        done = False
                        while done is False:
                            status, done = downloader.next_chunk()
                        fh.seek(0)
                        
                        doc_file.file.save(original_filename, ContentFile(fh.read()), save=True)

                        # --- THIS IS THE CRITICAL FIX ---
                        service.files().delete(
                            fileId=doc_file.gdrive_file_id,
                            supportsAllDrives=True # <-- THIS FLAG IS REQUIRED
                        ).execute()
                        
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

    return redirect(f"{reverse('ocrtesting:manage_archive')}?mode=archived")


# --- VIEW ARCHIVED FILE PROXY ---

@login_required
def view_archived_file(request, doc_type, pk):
    models_map = {
        'special_order': SpecialOrder, 'travel_order': TravelOrder, 'memorandum': Memorandum,
        'communication_letter': CommunicationLetter, 'moau': MOAU, 'other_document': OtherDocument,
    }
    model_class = models_map.get(doc_type.replace('-', '_'))
    if not model_class: raise Http404

    document = get_object_or_404(model_class, pk=pk)
    doc_file = document.documentfile_set.first()

    if not doc_file or not doc_file.gdrive_file_id:
        return HttpResponse("Archived file not found.", status=404)

    try:
        service = get_gdrive_service()
        file_metadata = service.files().get(fileId=doc_file.gdrive_file_id, fields='name').execute()
        filename = file_metadata.get('name', 'archived_file')
        gdrive_request = service.files().get_media(fileId=doc_file.gdrive_file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, gdrive_request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        fh.seek(0)
        
        response = HttpResponse(fh.getvalue(), content_type=doc_file.content_type)
        response['Content-Disposition'] = f'inline; filename="{filename}"'
        return response
    except Exception as e:
        return HttpResponse(f"Error viewing archived file: {e}", status=500)




# #GDRIVE ARCHIVE FUNCTIONALITY
# # --- NEW HELPER FUNCTION TO GET THE GOOGLE DRIVE SERVICE ---
# def get_gdrive_service():
#     """
#     Authenticates using the Service Account credentials and returns a
#     Google Drive service object. This is the correct method for an
#     application acting on its own behalf.
#     """
#     SCOPES = ['https://www.googleapis.com/auth/drive']
#     try:
#         creds = service_account.Credentials.from_service_account_file(
#             settings.GOOGLE_DRIVE_SERVICE_ACCOUNT_KEY_PATH, scopes=SCOPES
#         )
#         service = build('drive', 'v3', credentials=creds)
#         return service
#     except Exception as e:
#         # This will catch errors like the JSON file not being found or being invalid.
#         print(f"ERROR: Could not build Google Drive service. Check settings and credentials file. Details: {e}")
#         raise





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

    documents_page, page_range, querystring = _paginate(request, sorted_documents, per_page=10)

    # Prepare data for filter dropdowns
    all_years = set()
    for model in models_to_query.values():
        years = model.objects.values_list('fiscal_year', flat=True).distinct()
        all_years.update(years)
    fiscal_years_for_dropdown = sorted(list(all_years), reverse=True)
    month_list = [(i, calendar.month_name[i]) for i in range(1, 13)]

    context = {
        'documents': documents_page,
        'page_obj': documents_page,
        'page_range': page_range,
        'querystring': querystring,
        'page_title': page_title,
        'current_mode': mode,
        'fiscal_years_for_dropdown': fiscal_years_for_dropdown,
        'month_list': month_list,
        'filter_values': request.GET,
    }
    return render(request, 'manage_archive.html', context)



@login_required
def view_logs(request):
    """
    Displays a searchable list of all document activity logs.
    """
    # Start with all log entries, ordered by most recent first
    log_list = Log.objects.all()

    # Get the search query from the URL, if it exists
    search_query = request.GET.get('search', '')

    if search_query:
        # Filter the logs based on the search query.
        # This searches the action, document type, subject, and the user's username.
        log_list = log_list.filter(
            Q(action__icontains=search_query) |
            Q(document_type__icontains=search_query) |
            Q(document_subject__icontains=search_query) |
            Q(user__username__icontains=search_query)
        )

    logs_page, page_range, querystring = _paginate(request, log_list, per_page=10)

    context = {
        'logs': logs_page,
        'page_obj': logs_page,
        'page_range': page_range,
        'querystring': querystring,
        'search_query': search_query,
    }
    
    return render(request, 'logs.html', context)


def combine_files_to_pdf(files):
    """
    Combines multiple images and PDF files into a single PDF.
    Returns a ContentFile containing the merged PDF data.
    """
    from PyPDF2 import PdfWriter, PdfReader
    from PIL import Image
    import io
    from django.core.files.base import ContentFile

    writer = PdfWriter()

    for file in files:
        file.seek(0)
        content_type = getattr(file, 'content_type', '')
        
        # Check by content type or file extension
        is_pdf = content_type == 'application/pdf' or file.name.lower().endswith('.pdf')
        is_image = content_type.startswith('image/') or any(file.name.lower().endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.bmp'])

        if is_pdf:
            try:
                reader = PdfReader(file)
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                print(f"Error reading PDF {file.name}: {e}")
        elif is_image:
            try:
                img = Image.open(file).convert('RGB')
                pdf_bytes = io.BytesIO()
                img.save(pdf_bytes, format='PDF')
                pdf_bytes.seek(0)
                img_reader = PdfReader(pdf_bytes)
                writer.add_page(img_reader.pages[0])
            except Exception as e:
                print(f"Error converting image {file.name} to PDF: {e}")
    
    output_pdf = io.BytesIO()
    writer.write(output_pdf)
    output_pdf.seek(0)
    
    return ContentFile(output_pdf.read(), name="combined_document.pdf")