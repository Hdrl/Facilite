from django.urls import path

from . import views

urlpatterns = [
	path("", views.index, name="viagens"),
    path("<int:viagen_id>/", views.detalhe, name="detalhe"),
    path("<int:viagen_id>/exportar", views.exportar_zip, name="exportar_zip"),
]