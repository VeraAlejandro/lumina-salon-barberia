from django.contrib import messages
from django.db.models import Count
from django.shortcuts import redirect, render

from .forms import ReservaForm
from .models import Estilista, Producto, Reserva, Servicio, Sucursal


def inicio(request):
    servicios = Servicio.objects.filter(activo=True, destacado=True)[:6]
    if not servicios:
        servicios = Servicio.objects.filter(activo=True)[:6]
    contexto = {
        "titulo": "Inicio",
        "servicios": servicios,
        "estilistas": Estilista.objects.filter(activo=True).select_related("sucursal")[:4],
        "productos": Producto.objects.filter(activo=True)[:4],
        "sucursales": Sucursal.objects.filter(activa=True)[:3],
        "total_servicios": Servicio.objects.filter(activo=True).count(),
        "total_estilistas": Estilista.objects.filter(activo=True).count(),
    }
    return render(request, "web/inicio.html", contexto)


def servicios(request):
    categoria = request.GET.get("categoria", "").strip()
    consulta = Servicio.objects.filter(activo=True)
    if categoria:
        consulta = consulta.filter(categoria=categoria)
    contexto = {
        "titulo": "Servicios",
        "servicios": consulta,
        "categorias": Servicio.Categoria.choices,
        "categoria_activa": categoria,
    }
    return render(request, "web/servicios.html", contexto)


def estilistas(request):
    consulta = (
        Estilista.objects.filter(activo=True)
        .select_related("sucursal")
        .order_by("nombre")
    )
    contexto = {
        "titulo": "Estilistas",
        "estilistas": consulta,
    }
    return render(request, "web/estilistas.html", contexto)


def productos(request):
    categoria = request.GET.get("categoria", "").strip()
    busqueda = request.GET.get("q", "").strip()
    consulta = Producto.objects.filter(activo=True)
    if categoria:
        consulta = consulta.filter(categoria=categoria)
    if busqueda:
        consulta = consulta.filter(nombre__icontains=busqueda) | consulta.filter(
            marca__icontains=busqueda
        )
    contexto = {
        "titulo": "Productos",
        "productos": consulta.distinct(),
        "categorias": Producto.Categoria.choices,
        "categoria_activa": categoria,
        "busqueda": busqueda,
    }
    return render(request, "web/productos.html", contexto)


def sucursales(request):
    consulta = Sucursal.objects.filter(activa=True).annotate(
        total_estilistas=Count("estilistas", filter=None)
    )
    contexto = {
        "titulo": "Sucursales",
        "sucursales": consulta,
    }
    return render(request, "web/sucursales.html", contexto)


def reservar(request):
    if request.method == "POST":
        formulario = ReservaForm(request.POST)
        if formulario.is_valid():
            reserva = formulario.save()
            messages.success(
                request,
                f"¡Listo, {reserva.nombre.split()[0]}! Recibimos tu solicitud para el "
                f"{reserva.fecha:%d/%m/%Y} a las {reserva.hora:%H:%M}. Te contactaremos para confirmar.",
            )
            return redirect("web:reservar")
    else:
        formulario = ReservaForm()

    contexto = {
        "titulo": "Reserva tu cita",
        "formulario": formulario,
        "estilistas": Estilista.objects.filter(activo=True).select_related("sucursal"),
        "total_reservas": Reserva.objects.count(),
    }
    return render(request, "web/reservar.html", contexto)


def contacto(request):
    contexto = {
        "titulo": "Contacto",
        "sucursales": Sucursal.objects.filter(activa=True),
    }
    return render(request, "web/contacto.html", contexto)
