"""
Administración de Lumina.

Además de registrar los modelos, se personaliza la experiencia:
- Vistas previas de imágenes para identificar registros rápido.
- Filtros, búsqueda y edición en línea para operar sin entrar a cada ficha.
- Redirección amable para el rol Supervisor: al guardar vuelve al formulario
  de alta, porque su permiso es únicamente "add" y no puede ver los listados.
"""

from django.contrib import admin, messages
from django.db.models import Count
from django.http import HttpResponseRedirect
from django.utils.html import format_html

from .models import Cliente, Estilista, Producto, Reserva, Servicio, Sucursal

admin.site.site_header = "Lumina · Administración"
admin.site.site_title = "Lumina"
admin.site.index_title = "Panel de gestión del salón y barbería"
admin.site.empty_value_display = "—"


class ImagenAdminMixin:
    """Agrega una columna con la miniatura del registro."""

    campo_imagen = "imagen"
    campos_auditoria = ("creado_en", "actualizado_en")

    @admin.display(description="Vista previa")
    def vista_previa(self, obj):
        imagen = getattr(obj, self.campo_imagen, None)
        if not imagen:
            return "Sin imagen"
        return format_html(
            '<img src="{}" alt="Vista previa" '
            'style="height:46px;width:64px;object-fit:cover;border-radius:8px;'
            'border:1px solid #d8d2c4;" />',
            imagen.url,
        )


class RedireccionSupervisorMixin:
    """
    El Supervisor solo tiene permiso "add". Django, después de guardar, redirige
    al listado (donde recibiría 403). Aquí se le devuelve al formulario de alta
    con un mensaje de éxito para que pueda seguir capturando.
    """

    def response_add(self, request, obj, post_url_continue=None):
        es_supervisor = request.user.groups.filter(name="Supervisor").exists()
        if es_supervisor and not request.user.is_superuser:
            self.message_user(
                request,
                f'Registro "{obj}" guardado correctamente. Puede capturar el siguiente.',
                messages.SUCCESS,
            )
            return HttpResponseRedirect(request.path)
        return super().response_add(request, obj, post_url_continue)


@admin.register(Sucursal)
class SucursalAdmin(ImagenAdminMixin, RedireccionSupervisorMixin, admin.ModelAdmin):
    list_display = ("nombre", "ciudad", "telefono", "horario", "activa", "num_estilistas", "vista_previa")
    list_filter = ("activa", "ciudad")
    search_fields = ("nombre", "direccion", "telefono")
    readonly_fields = ("vista_previa",) + ImagenAdminMixin.campos_auditoria
    fieldsets = (
        ("Datos de la sucursal", {"fields": ("nombre", "direccion", "ciudad", "telefono", "correo")}),
        ("Operación", {"fields": ("horario", "mapa_url", "activa")}),
        ("Imagen", {"fields": ("imagen", "vista_previa")}),
        ("Auditoría", {"fields": ImagenAdminMixin.campos_auditoria, "classes": ("collapse",)}),
    )

    def get_queryset(self, request):
        return super().get_queryset(request).annotate(total_estilistas=Count("estilistas"))

    @admin.display(description="Estilistas", ordering="total_estilistas")
    def num_estilistas(self, obj):
        return obj.total_estilistas


@admin.register(Servicio)
class ServicioAdmin(ImagenAdminMixin, RedireccionSupervisorMixin, admin.ModelAdmin):
    list_display = (
        "nombre", "categoria", "precio", "duracion_minutos",
        "destacado", "activo", "vista_previa",
    )
    list_editable = ("destacado", "activo")
    list_filter = ("categoria", "activo", "destacado")
    search_fields = ("nombre", "descripcion")
    list_per_page = 15
    readonly_fields = ("vista_previa",) + ImagenAdminMixin.campos_auditoria
    fieldsets = (
        ("Información del servicio", {"fields": ("nombre", "categoria", "descripcion")}),
        ("Precio y duración", {"fields": ("precio", "duracion_minutos")}),
        ("Publicación", {"fields": ("destacado", "activo")}),
        ("Imagen", {"fields": ("imagen", "vista_previa")}),
        ("Auditoría", {"fields": ImagenAdminMixin.campos_auditoria, "classes": ("collapse",)}),
    )


@admin.register(Estilista)
class EstilistaAdmin(ImagenAdminMixin, RedireccionSupervisorMixin, admin.ModelAdmin):
    campo_imagen = "foto"
    list_display = (
        "nombre_completo", "especialidad", "sucursal",
        "anios_experiencia", "activo", "vista_previa",
    )
    list_filter = ("activo", "sucursal", "especialidad")
    search_fields = ("nombre", "apellidos", "especialidad")
    list_select_related = ("sucursal",)
    readonly_fields = ("vista_previa",) + ImagenAdminMixin.campos_auditoria
    fieldsets = (
        ("Datos personales", {"fields": ("nombre", "apellidos", "especialidad", "biografia")}),
        ("Contacto y adscripción", {"fields": ("sucursal", "correo", "telefono", "anios_experiencia")}),
        ("Publicación", {"fields": ("activo",)}),
        ("Fotografía", {"fields": ("foto", "vista_previa")}),
        ("Auditoría", {"fields": ImagenAdminMixin.campos_auditoria, "classes": ("collapse",)}),
    )

    @admin.display(description="Nombre completo", ordering="nombre")
    def nombre_completo(self, obj):
        return obj.nombre_completo


@admin.register(Producto)
class ProductoAdmin(ImagenAdminMixin, RedireccionSupervisorMixin, admin.ModelAdmin):
    list_display = (
        "nombre", "marca", "categoria", "precio", "stock",
        "estado_stock", "activo", "vista_previa",
    )
    list_filter = ("categoria", "activo", "marca")
    search_fields = ("nombre", "marca", "descripcion")
    list_editable = ("activo",)
    readonly_fields = ("vista_previa",) + ImagenAdminMixin.campos_auditoria
    fieldsets = (
        ("Información del producto", {"fields": ("nombre", "marca", "categoria", "descripcion")}),
        ("Precio e inventario", {"fields": ("precio", "stock")}),
        ("Publicación", {"fields": ("activo",)}),
        ("Imagen", {"fields": ("imagen", "vista_previa")}),
        ("Auditoría", {"fields": ImagenAdminMixin.campos_auditoria, "classes": ("collapse",)}),
    )

    @admin.display(description="Disponibilidad")
    def estado_stock(self, obj):
        if obj.stock == 0:
            return format_html('<span style="color:#a12622;font-weight:600;">Agotado</span>')
        if obj.stock <= 5:
            return format_html('<span style="color:#9a6700;font-weight:600;">Bajo ({})</span>', obj.stock)
        return format_html('<span style="color:#1a7f37;font-weight:600;">Disponible</span>')


@admin.register(Cliente)
class ClienteAdmin(RedireccionSupervisorMixin, admin.ModelAdmin):
    list_display = ("nombre", "apellidos", "telefono", "correo", "sucursal_preferida", "activo")
    list_filter = ("activo", "sucursal_preferida")
    search_fields = ("nombre", "apellidos", "telefono", "correo")
    list_select_related = ("sucursal_preferida",)
    readonly_fields = ImagenAdminMixin.campos_auditoria
    fieldsets = (
        ("Datos del cliente", {"fields": ("nombre", "apellidos", "telefono", "correo", "fecha_nacimiento")}),
        ("Preferencias", {"fields": ("sucursal_preferida", "notas", "activo")}),
        ("Auditoría", {"fields": ImagenAdminMixin.campos_auditoria, "classes": ("collapse",)}),
    )


@admin.register(Reserva)
class ReservaAdmin(RedireccionSupervisorMixin, admin.ModelAdmin):
    list_display = ("nombre", "telefono", "servicio", "estilista", "sucursal", "fecha", "hora", "estado")
    list_filter = ("estado", "sucursal", "fecha")
    search_fields = ("nombre", "telefono", "correo")
    list_select_related = ("servicio", "estilista", "sucursal")
    date_hierarchy = "fecha"
    list_editable = ("estado",)
    readonly_fields = ImagenAdminMixin.campos_auditoria
    fieldsets = (
        ("Cliente", {"fields": ("nombre", "telefono", "correo")}),
        ("Cita", {"fields": ("servicio", "estilista", "sucursal", "fecha", "hora")}),
        ("Seguimiento", {"fields": ("mensaje", "estado")}),
        ("Auditoría", {"fields": ImagenAdminMixin.campos_auditoria, "classes": ("collapse",)}),
    )
