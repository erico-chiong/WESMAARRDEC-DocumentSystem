from django import forms
from .models import SpecialOrder, TravelOrder, Memorandum, MOAU, Recipient, CommunicationLetter, OtherDocument
from django.forms import HiddenInput

# --- SPECIAL ORDER FORM ---
class SpecialOrderForm(forms.ModelForm):
    class Meta:
        model  = SpecialOrder
        # ADDED new fields
        fields = ['subject', 'date', 'recipients', 'received_by', 'remarks']
        widgets = {
            'subject': forms.Textarea(attrs={'rows': 2, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter subject'}),
            'date': forms.DateInput(attrs={'type': 'date', 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'recipients': forms.SelectMultiple(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            # ADDED widgets for new fields
            'received_by': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Name of person who received it'}),
            'remarks': forms.Textarea(attrs={'rows': 3, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter any additional notes'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # ADDED logic for required/optional fields
        self.fields['received_by'].required = True
        self.fields['remarks'].required = False


# --- TRAVEL ORDER FORM ---
class TravelOrderForm(forms.ModelForm):
    class Meta:
        model  = TravelOrder
        # ADDED new fields
        fields = ['subject', 'date', 'recipients', 'received_by', 'remarks']
        widgets = {
            'subject': forms.Textarea(attrs={'rows': 2, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter subject (optional)'}),
            'date': forms.DateInput(attrs={'type': 'date', 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'recipients': forms.SelectMultiple(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            # ADDED widgets for new fields
            'received_by': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Name of person who received it'}),
            'remarks': forms.Textarea(attrs={'rows': 3, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter any additional notes'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # ADDED logic for required/optional fields
        self.fields['received_by'].required = True
        self.fields['remarks'].required = False


# --- MEMORANDUM FORM ---
class MemorandumForm(forms.ModelForm):
    class Meta:
        model = Memorandum
        fields = [
            'subject', 'date', 'recipients', 
            'approved_date', 'category', 'for_field', 'from_field', 'thru',
            'received_by', 'remarks'
        ]
        widgets = {
            'subject': forms.Textarea(attrs={'rows': 2, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter subject'}),
            'date': forms.DateInput(attrs={'type': 'date', 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'recipients': forms.SelectMultiple(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'approved_date': forms.DateInput(attrs={'type': 'date', 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'category': HiddenInput(),
            'for_field': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'e.g., ALL CONCERNED PERSONNEL'}),
            'from_field': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'e.g., The Regional Director'}),
            'thru': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'e.g., The OIC-Director (Optional)'}),
            'received_by': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Name of person who received it'}),
            'remarks': forms.Textarea(attrs={'rows': 3, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter any additional notes'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Set fields as optional by default in the form's init.
        # The 'clean' method below will enforce the requirement conditionally.
        self.fields['approved_date'].required = False
        self.fields['for_field'].required = False
        self.fields['from_field'].required = False
        self.fields['thru'].required = False
        self.fields['received_by'].required = False # This is now always False here
        self.fields['remarks'].required = False

    def clean(self):
        cleaned_data = super().clean()
        category = cleaned_data.get("category")
        received_by = cleaned_data.get("received_by")

        # This is the correct server-side validation logic.
        # If the category is 'incoming' and the 'received_by' field is empty, raise an error.
        if category == 'incoming' and not received_by:
            self.add_error('received_by', 'This field is required for incoming memorandums.')
        
        return cleaned_data



# --- MOAU FORM ---
class MOAUForm(forms.ModelForm):
    signatories_str = forms.CharField(label="Signatories", required=False, widget=forms.HiddenInput())

    class Meta:
        model = MOAU
        # ADDED new fields
        fields = [
            'subject', 'date', 'recipients', 'first_party_agency', 
            'first_party_representative', 'second_party_agency', 
            'second_party_representative', 'received_by', 'remarks'
        ]
        widgets = {
            'subject': forms.Textarea(attrs={'rows': 2, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter subject of the MOA/U'}),
            'date': forms.DateInput(attrs={'type': 'date', 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'recipients': forms.SelectMultiple(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'first_party_agency': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'e.g., Department of Agriculture'}),
            'first_party_representative': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Name of the first representative'}),
            'second_party_agency': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'e.g., Local Government Unit of City'}),
            'second_party_representative': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Name of the second representative'}),
            # ADDED widgets for new fields
            'received_by': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Name of person who received it'}),
            'remarks': forms.Textarea(attrs={'rows': 3, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter any additional notes'}),
        }
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Preserved existing logic
        self.fields['first_party_agency'].required = False
        self.fields['first_party_representative'].required = False
        self.fields['second_party_agency'].required = False
        self.fields['second_party_representative'].required = False
        self.fields['date'].required = False
        # ADDED logic for new fields
        self.fields['received_by'].required = True
        self.fields['remarks'].required = False


# --- COMMUNICATION LETTER FORM ---
class CommunicationLetterForm(forms.ModelForm):
    class Meta:
        model = CommunicationLetter
        # ADDED new fields
        fields = [
            'subject', 'category', 'date', 'from_field',
            'to_or_thru', 'recipients', 'received_by', 'remarks'
        ]
        widgets = {
            'subject': forms.Textarea(attrs={'rows': 2, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter subject of the letter'}),
            'category': HiddenInput(),
            'date': forms.DateInput(attrs={'type': 'date', 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'from_field': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'e.g., Juan Dela Cruz, Company XYZ'}),
            'to_or_thru': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'e.g., The Regional Director'}),
            'recipients': forms.SelectMultiple(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            # ADDED widgets for new fields
            'received_by': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Name of person who received it'}),
            'remarks': forms.Textarea(attrs={'rows': 3, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter any additional notes'}),
        }
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # ADDED logic for new fields
        self.fields['remarks'].required = False
        # Conditional logic for received_by based on category
        if self.instance and self.instance.category == 'outgoing':
            self.fields['received_by'].required = False
        elif 'initial' in kwargs and kwargs['initial'].get('category') == 'outgoing':
            self.fields['received_by'].required = False
        else:
            self.fields['received_by'].required = True


# --- OTHER DOCUMENT FORM ---
class OtherDocumentForm(forms.ModelForm):
    class Meta:
        model = OtherDocument
        # ADDED new field
        fields = [
            'subject', 'doc_type', 'date', 'remarks',
            'recipients', 'received_by'
        ]
        widgets = {
            'subject': forms.Textarea(attrs={'rows': 2, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter the main subject or title'}),
            'doc_type': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'e.g., Certification, Report, Notice'}),
            'date': forms.DateInput(attrs={'type': 'date', 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            'remarks': forms.Textarea(attrs={'rows': 3, 'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter any additional notes or remarks'}),
            'recipients': forms.SelectMultiple(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500'}),
            # ADDED widget for new field
            'received_by': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Name of person who received it'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Preserved existing logic
        self.fields['remarks'].required = False
        # ADDED logic for new field
        self.fields['received_by'].required = True


# --- RECIPIENT FORM (UNCHANGED) ---
class RecipientForm(forms.ModelForm):
    class Meta:
        model = Recipient
        fields = ['name', 'email']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded-md shadow-sm px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter full name'}),
            'email': forms.EmailInput(attrs={'class': 'mt-1 block w-full border-gray-300 rounded-md shadow-sm px-3 py-2 focus:ring-blue-500 focus:border-blue-500', 'placeholder': 'Enter email address'}),
        }