import os
import uuid
from django.db import models
from django.utils import timezone
from django.conf import settings
from django.db.utils import IntegrityError


def document_file_upload_to(instance, filename):
    ext = filename.split('.')[-1]
    new_filename = f"{uuid.uuid4()}.{ext}"
    return os.path.join('documents', new_filename)


def _get_document_year(instance) -> int:
    date_value = getattr(instance, 'date', None)
    if date_value:
        try:
            return int(date_value.year)
        except Exception:
            pass

    fiscal_year_value = getattr(instance, 'fiscal_year', None)
    if fiscal_year_value:
        try:
            return int(fiscal_year_value)
        except (TypeError, ValueError):
            pass

    return int(timezone.now().year)


def _generate_document_id(prefix: str, year: int | None = None) -> str:
    now = timezone.now()
    timestamp = now.strftime('%Y%m%d%H%M%S%f')
    doc_year = int(year) if year else int(now.year)
    safe_prefix = (prefix or 'DOC').upper().replace(' ', '')
    return f"{safe_prefix}_{timestamp}_{doc_year}"


def _generate_unique_document_id(model_class, prefix: str, year: int | None = None) -> str:
    """
    Generates a unique ID for the given model.

    Base format: PREFIX_<timestamp>_<year>
    Example: MEMO_20260519153012123456_2026
    """
    base = _generate_document_id(prefix=prefix, year=year)
    candidate = base
    suffix = 1
    while model_class.objects.filter(document_id=candidate).exists():
        candidate = f"{base}_{suffix}"
        suffix += 1
    return candidate


def _ensure_document_id(instance, prefix: str, *, force: bool = False) -> None:
    if not force and getattr(instance, 'document_id', None):
        return

    doc_year = _get_document_year(instance)
    instance.document_id = _generate_unique_document_id(instance.__class__, prefix=prefix, year=doc_year)

# ─── Recipients ──────────────────────────────

class Recipient(models.Model):
    name = models.CharField(max_length=200)
    email = models.EmailField()

    def __str__(self):
        return self.name

# ─── Signatories ─────────────────────────────

class Signatory(models.Model):
    name = models.CharField(max_length=200)
    designation = models.CharField(max_length=200, blank=True, null=True)

    def __str__(self):
        return self.name

# ─── Document Models ─────────────────────────

class Memorandum(models.Model):
    DOCUMENT_ID_PREFIX = 'MEMO'

    fiscal_year = models.IntegerField(default=timezone.now().year)
    document_id = models.CharField(max_length=50, primary_key=True, editable=False)
    subject = models.CharField(max_length=500)
    category = models.CharField(max_length=20, choices=[('incoming', 'Incoming'), ('outgoing', 'Outgoing')])
    for_field = models.CharField(max_length=200, blank=True, null=True)
    from_field = models.CharField(max_length=200, blank=True, null=True)
    thru = models.CharField(max_length=200, blank=True, null=True)
    date = models.DateField()
    approved_date = models.DateField(blank=True, null=True)
    upload_date = models.DateTimeField(auto_now_add=True)
    edit_date = models.DateTimeField(auto_now=True)
    recipients = models.ManyToManyField(Recipient, blank=True)
    is_archived = models.BooleanField(default=False)

    # --- NEW FIELDS ---
    received_by = models.CharField(max_length=255, blank=True)
    remarks = models.TextField(blank=True, null=True)

    def save(self, *args, **kwargs):
        if not self.document_id:
            _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX)

        # Safety net for rare collisions under concurrency
        for _ in range(3):
            try:
                return super().save(*args, **kwargs)
            except IntegrityError:
                _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX, force=True)
        return super().save(*args, **kwargs)

    def __str__(self):
        return f"Memorandum {self.document_id}"

class CommunicationLetter(models.Model):
    DOCUMENT_ID_PREFIX = 'COMM'

    fiscal_year = models.IntegerField(default=timezone.now().year)
    document_id = models.CharField(max_length=50, unique=True, editable=False, null=True, blank=True)
    id = models.AutoField(primary_key=True)
    subject = models.CharField(max_length=500)
    category = models.CharField(max_length=20, choices=[('incoming', 'Incoming'), ('outgoing', 'Outgoing')])
    date = models.DateField()
    from_field = models.CharField(max_length=200)
    to_or_thru = models.CharField(max_length=200)
    upload_date = models.DateTimeField(auto_now_add=True)
    edit_date = models.DateTimeField(auto_now=True)
    recipients = models.ManyToManyField(Recipient, blank=True)
    is_archived = models.BooleanField(default=False)

    # --- NEW FIELDS ---
    received_by = models.CharField(max_length=255, blank=True)
    remarks = models.TextField(blank=True, null=True)

    def save(self, *args, **kwargs):
        if not self.document_id:
            _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX)

        for _ in range(3):
            try:
                return super().save(*args, **kwargs)
            except IntegrityError:
                _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX, force=True)
        return super().save(*args, **kwargs)

    def __str__(self):
        return f"CommLetter {self.document_id or self.id}"

class SpecialOrder(models.Model):
    DOCUMENT_ID_PREFIX = 'SO'

    fiscal_year = models.IntegerField(default=timezone.now().year)
    document_id = models.CharField(max_length=50, primary_key=True, editable=False)
    subject = models.CharField(max_length=500)
    date = models.DateField()
    upload_date = models.DateTimeField(auto_now_add=True)
    edit_date = models.DateTimeField(auto_now=True)
    recipients = models.ManyToManyField(Recipient, blank=True)
    is_archived = models.BooleanField(default=False)

    # --- NEW FIELDS ---
    received_by = models.CharField(max_length=255, blank=True)
    remarks = models.TextField(blank=True, null=True)

    def save(self, *args, **kwargs):
        if not self.document_id:
            _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX)

        for _ in range(3):
            try:
                return super().save(*args, **kwargs)
            except IntegrityError:
                _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX, force=True)
        return super().save(*args, **kwargs)

    def __str__(self):
        return f"SpecialOrder {self.document_id}"

class TravelOrder(models.Model):
    DOCUMENT_ID_PREFIX = 'TO'

    fiscal_year = models.IntegerField(default=timezone.now().year)
    document_id = models.CharField(max_length=50, primary_key=True, editable=False)
    subject = models.CharField(max_length=500, blank=True)
    date = models.DateField()
    upload_date = models.DateTimeField(auto_now_add=True)
    edit_date = models.DateTimeField(auto_now=True)
    recipients = models.ManyToManyField(Recipient, blank=True)
    is_archived = models.BooleanField(default=False)

    # --- NEW FIELDS ---
    received_by = models.CharField(max_length=255, blank=True)
    remarks = models.TextField(blank=True, null=True)

    def save(self, *args, **kwargs):
        if not self.document_id:
            _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX)

        for _ in range(3):
            try:
                return super().save(*args, **kwargs)
            except IntegrityError:
                _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX, force=True)
        return super().save(*args, **kwargs)

    def __str__(self):
        return f"TravelOrder {self.document_id}"

class MOAU(models.Model):
    DOCUMENT_ID_PREFIX = 'MOAU'

    fiscal_year = models.IntegerField(default=timezone.now().year)
    document_id = models.CharField(max_length=50, unique=True, editable=False, null=True, blank=True)
    id = models.AutoField(primary_key=True)
    subject = models.CharField(max_length=500)
    first_party_agency = models.CharField(max_length=200, blank=True, null=True)
    first_party_representative = models.CharField(max_length=200, blank=True, null=True)
    second_party_agency = models.CharField(max_length=200, blank=True, null=True)
    second_party_representative = models.CharField(max_length=200, blank=True, null=True)
    date = models.DateField(blank=True, null=True)
    signatories = models.ManyToManyField(Signatory, blank=True)
    upload_date = models.DateTimeField(auto_now_add=True)
    edit_date = models.DateTimeField(auto_now=True)
    recipients = models.ManyToManyField(Recipient, blank=True)
    is_archived = models.BooleanField(default=False)

    # --- NEW FIELDS ---
    received_by = models.CharField(max_length=255, blank=True)
    remarks = models.TextField(blank=True, null=True)

    def save(self, *args, **kwargs):
        if not self.document_id:
            _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX)

        for _ in range(3):
            try:
                return super().save(*args, **kwargs)
            except IntegrityError:
                _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX, force=True)
        return super().save(*args, **kwargs)

    def __str__(self):
        return f"MOAU {self.document_id or self.id}"

class OtherDocument(models.Model):
    DOCUMENT_ID_PREFIX = 'OTHR'

    fiscal_year = models.IntegerField(default=timezone.now().year)
    document_id = models.CharField(max_length=50, unique=True, editable=False, null=True, blank=True)
    id = models.AutoField(primary_key=True)
    subject = models.CharField(max_length=500)
    doc_type = models.CharField(max_length=100)
    date = models.DateField()
    remarks = models.TextField(blank=True, null=True) # This field already existed
    upload_date = models.DateTimeField(auto_now_add=True)
    edit_date = models.DateTimeField(auto_now=True)
    recipients = models.ManyToManyField(Recipient, blank=True)
    is_archived = models.BooleanField(default=False)

    # --- NEW FIELD ---
    received_by = models.CharField(max_length=255, blank=True)

    def save(self, *args, **kwargs):
        if not self.document_id:
            _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX)

        for _ in range(3):
            try:
                return super().save(*args, **kwargs)
            except IntegrityError:
                _ensure_document_id(self, prefix=self.DOCUMENT_ID_PREFIX, force=True)
        return super().save(*args, **kwargs)

    def __str__(self):
        return f"OtherDocument {self.document_id or self.id}"

# ─── Document Files ──────────────────────────

class DocumentFile(models.Model):
    # This model is unchanged
    file = models.FileField(upload_to=document_file_upload_to, blank=True, null=True)
    content_type = models.CharField(max_length=100)
    order = models.IntegerField(default=1)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    gdrive_file_id = models.CharField(max_length=100, blank=True, null=True)
    memorandum = models.ForeignKey(Memorandum, blank=True, null=True, on_delete=models.CASCADE)
    communication_letter = models.ForeignKey(CommunicationLetter, blank=True, null=True, on_delete=models.CASCADE)
    special_order = models.ForeignKey(SpecialOrder, blank=True, null=True, on_delete=models.CASCADE)
    travel_order = models.ForeignKey(TravelOrder, blank=True, null=True, on_delete=models.CASCADE)
    moau = models.ForeignKey(MOAU, blank=True, null=True, on_delete=models.CASCADE)
    other_document = models.ForeignKey(OtherDocument, blank=True, null=True, on_delete=models.CASCADE)

    def __str__(self):
        return f"File for document ({self.id})"

# ─── Logging ─────────────────────────────────

class Log(models.Model):
    # This model is unchanged
    user = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True)
    action = models.CharField(max_length=255)
    timestamp = models.DateTimeField(auto_now_add=True)
    document_type = models.CharField(max_length=100)
    document_id = models.CharField(max_length=50)
    document_subject = models.CharField(max_length=500, blank=True, null=True)

    def __str__(self):
        user_name = self.user.username if self.user else "System"
        return f"{self.timestamp.strftime('%Y-%m-%d %H:%M')} - {user_name} - {self.action}"

    class Meta:
        ordering = ['-timestamp']