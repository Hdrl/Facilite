# app_name/forms.py
from django import forms
from .models import TransacaoFinanceira

class ImagemForm(forms.ModelForm):
    class Meta:
        model = TransacaoFinanceira
        fields = ['imagem']