from django.shortcuts import render
from django.http import HttpResponse
from django.db.models import Sum
from django.template import loader
from django.http import FileResponse
from io import StringIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font, Alignment

from .models import Viagem
from .models import TransacaoFinanceira
from .forms import ImagemForm
from django.urls import reverse
from django.http import HttpResponseRedirect
from bs4 import BeautifulSoup as bs4
import locale
import datetime
import csv
import os
import zipfile
import tempfile

# Create your views here.
        
def index(request):
    viagens = Viagem.objects.all()
    template = loader.get_template("viagens/index.html")
    context = {
        'viagens_list': viagens
    }
    return HttpResponse(template.render(context, request))
    
def detalhe(request, viagen_id):
    viagem = Viagem.objects.get(id=viagen_id)
    
    if request.method == "POST":
        imagens =  request.FILES.getlist('imagem')
        for img in imagens:
            form = ImagemForm(request.POST, {'imagem': img})
            if form.is_valid():
                despesa = form.save(commit=False)
                despesa.viagem = viagem
                despesa.imagem = img
                despesa.user = request.user
                despesa.save()
    else:
        form = ImagemForm()
    valor_total = TransacaoFinanceira.objects.filter(viagem=viagem, tipo='S').aggregate(Sum('valor'))['valor__sum'] or 0.0
    context = {
        'viagem': viagem,
        'despesas': viagem.despesas.filter(tipo='S'),
        'form' : form,
        'sum' : valor_total
    }
    return HttpResponse(loader.get_template("viagens/detalhe.html").render(context, request))