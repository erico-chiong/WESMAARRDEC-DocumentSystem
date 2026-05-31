from django.core.mail import EmailMessage
from django.utils.text import slugify
from .pdf_export import generate_image_pdf

def send_document_email(document, recipient_email):
    # Generate PDF from the document
    pdf_stream = generate_image_pdf(document)

    # Create a clean filename from the document's subject
    raw_name = document.subject or 'document'
    slug = slugify(raw_name)
    if not slug:
        slug = f"doc-{document.pk}"
    slug = slug[:50]
    filename = f"{slug}.pdf"

    # Prepare the email
    email = EmailMessage(
        subject=f"Document: {document.subject or 'Untitled'}",
        body="Please find the attached document.",
        to=[recipient_email]
    )

    # Attach PDF to the email
    email.attach(filename, pdf_stream.read(), 'application/pdf')

    # Send email
    email.send()
