from django import forms
from django_select2.forms import Select2MultipleWidget, Select2Widget

class Dados(forms.Form):
    dataInicio = forms.CharField(label="Your name", max_length=100)