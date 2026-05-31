from .models import Log

def create_log(user, action, document):
    """
    A helper function to create a Log entry.
    This centralizes the logging logic and is robust against anonymous users.
    """
    # Get the display name for the document type
    # This handles names like "SpecialOrder" -> "Special Order"
    doc_type_display = document.__class__.__name__.replace(
        'Order', ' Order'
    ).replace(
        'Letter', ' Letter'
    )

    # Prefer the human-readable document_id if available; otherwise fall back to PK.
    doc_identifier = getattr(document, 'document_id', None) or document.pk
    doc_pk = str(doc_identifier)
    
    # Get the subject, providing a fallback if it doesn't exist or is empty
    doc_subject = getattr(document, 'subject', '(No Subject)') or '(No Subject)'
    
    # --- THIS IS THE ONLY CHANGE ---
    # Check if the user is a real, logged-in user. If not, set user to None.
    # This prevents errors if an AnonymousUser object is passed.
    user_to_log = user if user.is_authenticated else None

    Log.objects.create(
        user=user_to_log,
        action=action,
        document_type=doc_type_display,
        document_id=doc_pk,
        document_subject=doc_subject
    )