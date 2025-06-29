from django.shortcuts import render
from django.http import HttpResponse
from django.db.models import Sum
from django.template import loader
from django.http import FileResponse
from io import StringIO
from datetime import datetime

from .models import Viagen
from .models import Despesa
from .forms import ImagemForm

import locale
import csv
import os
import zipfile
import tempfile
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


# Create your views here.
def index(request):
    viagens = Viagen.objects.all()
    template = loader.get_template("viagens/index.html")
    context = {
        'viagens_list': viagens
    }
    return HttpResponse(template.render(context, request))
    

def exportar_zip(request, viagen_id):
    tmp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(tmp_dir, 'exportacao.zip')
    despesas = Despesa.objects.filter(viagem=Viagen.objects.get(id=viagen_id))
    
    with zipfile.ZipFile(zip_path, 'w')as zipf:
        csv_buffer = StringIO()
        writer = csv.writer(csv_buffer)
        writer.writerow(['Relatório de Despesas de Viagem'])
        writer.writerow(['Descricao', 'Data', 'Nota Fiscal', 'Valor'])  # cabeçalhos
        
        for despesa in despesas:
            writer.writerow([despesa.descricao.lower(), despesa.data.strftime('%d/%m/%Y %H:%M'), f"=hiperlink(\"notas fiscal/{os.path.basename(despesa.imagem.path)}\"; \"{os.path.basename(despesa.imagem.path)}\")", f"{despesa.valor:.2f}".replace('.', ',')])  # substitua pelos campos reais
            
            if despesa.imagem:
                caminho = despesa.imagem.path
                nome_arquivo = os.path.basename(caminho)
                zipf.write(caminho, arcname=F"notas fiscal/{nome_arquivo}")
        writer.writerow(["TOTAL", "", "", "=SOMA()"])
        zipf.writestr('despesas.csv', csv_buffer.getvalue())
    return FileResponse(open(zip_path, 'rb'), as_attachment=True, filename='despesas.zip')


def detalhe(request, viagen_id):
    viagem = Viagen.objects.get(id=viagen_id)
    if request.method == "POST":
        form = ImagemForm(request.POST, request.FILES)
        if form.is_valid():
            despesa = form.save(commit=False)
            despesa.viagem = viagem
            despesa.save()
    else:
        form = ImagemForm()
    
    context = {
        'viagem': viagem,
        'despesas': viagem.despesas.all(),
        'form' : form,
        'sum' : locale.currency(Despesa.objects.filter(viagem=viagem).aggregate(Sum('valor'))['valor__sum'], grouping=True)
    }
    return HttpResponse(loader.get_template("viagens/detalhe.html").render(context, request))