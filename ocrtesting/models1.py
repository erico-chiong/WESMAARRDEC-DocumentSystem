# ocrtesting/models.py

import os
import uuid
from django.db import models
from django.core.validators import RegexValidator
from django.core.exceptions import ValidationError
from django.utils import timezone

def document_image_upload_to(instance, filename):
    """
    Generates a unique filename using a UUID and stores the file in the 'documents/' directory.
    """
    ext = filename.split('.')[-1]
    new_filename = f"{uuid.uuid4()}.{ext}"
    return os.path.join('documents', new_filename)

class Recipient(models.Model):
    name  = models.CharField(max_length=200)
    email = models.EmailField()

    def __str__(self):
        return self.name

class Document(models.Model):
    DOCUMENT_TYPES = [
        ('memorandum', 'Memorandum'),
        ('travel_order', 'Travel Order'),
        ('special_order', 'Special Order'),
        ('communication_letter', 'Communication Letter'),
        ('moau', 'MOAU'),
        ('other_document', 'Other Document'),
    ]

    CATEGORIES = [
        ('incoming', 'Incoming'),
        ('outgoing', 'Outgoing'),
    ]

    document_type = models.CharField(max_length=30, choices=DOCUMENT_TYPES)
    category      = models.CharField(max_length=30, choices=CATEGORIES, blank=True, null=True)
    doc_id        = models.CharField(
                        max_length=6,
                        validators=[RegexValidator(r'^\d{6}$', message='ID must be exactly 6 digits')],
                        blank=True
                    )
    subject       = models.CharField(max_length=500)
    recipients    = models.ManyToManyField(Recipient, blank=True)
    date          = models.DateField()
    approved_date = models.DateField(blank=True, null=True)

    # Keep this for backward-compatibility during migration:
    image         = models.ImageField(upload_to=document_image_upload_to)

    cancelled     = models.BooleanField(default=False)

    # Phase 1 fields
    from_field   = models.CharField("From", max_length=200, blank=True, null=True)
    to           = models.CharField(max_length=200, blank=True, null=True)
    upload_date  = models.DateTimeField(auto_now_add=True, editable=False)
    edit_date    = models.DateTimeField(auto_now=True)
    fiscal_year  = models.CharField(max_length=9, blank=True, null=True)
    remarks      = models.TextField(blank=True, null=True)

    class Meta:
        unique_together = ('document_type', 'doc_id')

    def __str__(self):
        return f"{self.get_document_type_display()} - {self.doc_id}"

    def clean(self):
        if self.document_type in ['memorandum', 'communication_letter'] and not self.category:
            raise ValidationError("Category is required for Memorandum and Communication Letter.")

    def save(self, *args, **kwargs):
        # Auto-generate the doc_id if not provided.
        if not self.doc_id:
            last = Document.objects.filter(document_type=self.document_type)\
                                   .order_by('id')\
                                   .last()
            if last and last.doc_id:
                try:
                    new_id = int(last.doc_id) + 1
                except ValueError:
                    new_id = 1
            else:
                new_id = 1
            self.doc_id = str(new_id).zfill(6)

        # Auto-set fiscal_year to current year if not already provided.
        if not self.fiscal_year:
            self.fiscal_year = str(timezone.now().year)

        super().save(*args, **kwargs)

class DocumentPage(models.Model):
    document = models.ForeignKey(
        Document,
        related_name='pages',
        on_delete=models.CASCADE
    )
    image = models.ImageField(
        upload_to=document_image_upload_to
    )
    order = models.PositiveIntegerField(
        default=1,
        help_text='Determines page position within the document'
    )

    class Meta:
        unique_together = ('document', 'order')
        ordering = ['order']

    def __str__(self):
        return f"{self.document} – page {self.order}"