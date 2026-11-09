from django.urls import path

from . import views

app_name = "web"

urlpatterns = [
    path("", views.inicio, name="inicio"),
    path("servicios/", views.servicios, name="servicios"),
    path("estilistas/", views.estilistas, name="estilistas"),
    path("productos/", views.productos, name="productos"),
    path("sucursales/", views.sucursales, name="sucursales"),
    path("reservar/", views.reservar, name="reservar"),
    path("contacto/", views.contacto, name="contacto"),
]
