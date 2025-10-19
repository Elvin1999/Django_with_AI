from django import forms

class UploadForm(forms.Form):
    file=forms.FileField(help_text="Excel/CSV (xlsx,xls,csv)")
    sheet_name=forms.CharField(required=False)

class DateFilterForm(forms.Form):
    date_from=forms.DateField(required=False,widget=forms.DateInput(attrs={'type':'date'}))
    date_to=forms.DateField(required=False,widget=forms.DateInput(attrs={'type':'date'}))
    category=forms.CharField(required=False)
