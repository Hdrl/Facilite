# app_name/forms.py
from django import forms
from .models import Despesa

class ImagemForm(forms.ModelForm):
    class Meta:
        model = Despesa
        fields = ['imagem']