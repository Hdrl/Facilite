from django.db import models
from datetime import datetime
import locale

TIPO_CHOICES = [
    ('E', 'Entrada'),
    ('S', 'Saida'),
]

# Create your models here.
class Viagen(models.Model):
    destino=models.CharField(max_length=20) 
    empresa=models.CharField(max_length=50, default = "Smart Picking Soluções Digitais Ltda")
    setor=models.CharField(max_length=10, default="Smart")
    colaborador=models.CharField(max_length=50, default="Fernando Pereira da Trindade")
    retorno=models.DateField(blank = True, null = True)
    motivo=models.CharField(max_length=50)
    saida=models.CharField(max_length=20, default="Apucarana") 
    
    def __str__(self):
        return f"{self.destino} - {self.motivo}"
    
class Despesa(models.Model):
    valor = models.DecimalField(max_digits=10, decimal_places=2, blank = True, null = True)
    descricao = models.CharField(max_length=100, blank = True, null = True)
    data = models.DateTimeField(blank = True, null = True)
    nota_fiscal = models.URLField(max_length=200, blank = True, null = True)
    imagem = models.ImageField(upload_to='despesas/')
    viagem = models.ForeignKey(Viagen, on_delete=models.CASCADE, related_name='despesas')
    tipo = models.CharField(max_length=1, choices=TIPO_CHOICES, default='S')
    
    def __str__(self):
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        return f"{locale.currency(self.valor, grouping=True)} - {self.descricao.upper()} - {self.data.strftime('%d/%m/%Y %H:%M')}"