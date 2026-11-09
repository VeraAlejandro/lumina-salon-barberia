"""
Modelos de Lumina — Salón & Barbería.

Cinco entidades son las solicitadas en el PIA (servicios, estilistas, productos,
sucursales y clientes). Se agrega "Reserva" para conectar el sitio público con
la administración: el visitante solicita una cita y el personal la gestiona.
"""

from django.db import models


class RegistroBase(models.Model):
    """Campos comunes de auditoría para todos los catálogos."""

    creado_en = models.DateTimeField("creado en", auto_now_add=True)
    actualizado_en = models.DateTimeField("actualizado en", auto_now=True)

    class Meta:
        abstract = True


class Sucursal(RegistroBase):
    nombre = models.CharField("nombre", max_length=100, unique=True)
    direccion = models.CharField("dirección", max_length=200)
    ciudad = models.CharField("ciudad", max_length=80, default="San Nicolás de los Garza")
    telefono = models.CharField("teléfono", max_length=20)
    correo = models.EmailField("correo electrónico", blank=True)
    horario = models.CharField(
        "horario",
        max_length=120,
        default="Lun a Sáb · 9:00 a 20:00",
        help_text="Texto que se muestra al público, por ejemplo: Lun a Sáb · 9:00 a 20:00",
    )
    imagen = models.ImageField("fotografía", upload_to="sucursales/", blank=True)
    mapa_url = models.URLField("enlace de mapa", blank=True)
    activa = models.BooleanField("activa", default=True)

    class Meta:
        verbose_name = "sucursal"
        verbose_name_plural = "sucursales"
        ordering = ["nombre"]

    def __str__(self):
        return self.nombre


class Servicio(RegistroBase):
    class Categoria(models.TextChoices):
        CABELLO = "cabello", "Cabello"
        BARBERIA = "barberia", "Barbería"
        UNAS = "unas", "Uñas"
        FACIAL = "facial", "Facial y skincare"
        MAQUILLAJE = "maquillaje", "Maquillaje"

    nombre = models.CharField("nombre", max_length=120, unique=True)
    categoria = models.CharField(
        "categoría", max_length=20, choices=Categoria.choices, default=Categoria.CABELLO
    )
    descripcion = models.TextField("descripción", max_length=400)
    precio = models.DecimalField("precio (MXN)", max_digits=8, decimal_places=2)
    duracion_minutos = models.PositiveIntegerField("duración (minutos)", default=45)
    imagen = models.ImageField("imagen", upload_to="servicios/", blank=True)
    destacado = models.BooleanField("destacado en inicio", default=False)
    activo = models.BooleanField("activo", default=True)

    class Meta:
        verbose_name = "servicio"
        verbose_name_plural = "servicios"
        ordering = ["categoria", "nombre"]

    def __str__(self):
        return self.nombre


class Estilista(RegistroBase):
    nombre = models.CharField("nombre", max_length=80)
    apellidos = models.CharField("apellidos", max_length=80)
    especialidad = models.CharField("especialidad", max_length=120)
    biografia = models.TextField("biografía", max_length=500, blank=True)
    foto = models.ImageField("fotografía", upload_to="estilistas/", blank=True)
    sucursal = models.ForeignKey(
        Sucursal,
        on_delete=models.PROTECT,
        related_name="estilistas",
        verbose_name="sucursal",
    )
    correo = models.EmailField("correo electrónico", blank=True)
    telefono = models.CharField("teléfono", max_length=20, blank=True)
    anios_experiencia = models.PositiveIntegerField("años de experiencia", default=1)
    activo = models.BooleanField("activo", default=True)

    class Meta:
        verbose_name = "estilista"
        verbose_name_plural = "estilistas"
        ordering = ["nombre", "apellidos"]

    def __str__(self):
        return f"{self.nombre} {self.apellidos}"

    @property
    def nombre_completo(self):
        return f"{self.nombre} {self.apellidos}"


class Producto(RegistroBase):
    class Categoria(models.TextChoices):
        CABELLO = "cabello", "Cuidado del cabello"
        BARBA = "barba", "Cuidado de la barba"
        PIEL = "piel", "Cuidado de la piel"
        UNAS = "unas", "Uñas"
        HERRAMIENTA = "herramienta", "Herramientas"

    nombre = models.CharField("nombre", max_length=120)
    marca = models.CharField("marca", max_length=80)
    categoria = models.CharField(
        "categoría", max_length=20, choices=Categoria.choices, default=Categoria.CABELLO
    )
    descripcion = models.TextField("descripción", max_length=400, blank=True)
    precio = models.DecimalField("precio (MXN)", max_digits=8, decimal_places=2)
    stock = models.PositiveIntegerField("existencias", default=0)
    imagen = models.ImageField("imagen", upload_to="productos/", blank=True)
    activo = models.BooleanField("activo", default=True)

    class Meta:
        verbose_name = "producto"
        verbose_name_plural = "productos"
        ordering = ["marca", "nombre"]

    def __str__(self):
        return f"{self.nombre} · {self.marca}"

    @property
    def disponible(self):
        return self.stock > 0


class Cliente(RegistroBase):
    nombre = models.CharField("nombre", max_length=80)
    apellidos = models.CharField("apellidos", max_length=80)
    telefono = models.CharField("teléfono", max_length=20)
    correo = models.EmailField("correo electrónico", blank=True)
    fecha_nacimiento = models.DateField("fecha de nacimiento", null=True, blank=True)
    sucursal_preferida = models.ForeignKey(
        Sucursal,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="clientes",
        verbose_name="sucursal preferida",
    )
    notas = models.TextField(
        "notas",
        max_length=500,
        blank=True,
        help_text="Preferencias, alergias o detalles del servicio.",
    )
    activo = models.BooleanField("activo", default=True)

    class Meta:
        verbose_name = "cliente"
        verbose_name_plural = "clientes"
        ordering = ["apellidos", "nombre"]

    def __str__(self):
        return f"{self.nombre} {self.apellidos}"


class Reserva(RegistroBase):
    class Estado(models.TextChoices):
        PENDIENTE = "pendiente", "Pendiente"
        CONFIRMADA = "confirmada", "Confirmada"
        ATENDIDA = "atendida", "Atendida"
        CANCELADA = "cancelada", "Cancelada"

    nombre = models.CharField("nombre del cliente", max_length=120)
    telefono = models.CharField("teléfono", max_length=20)
    correo = models.EmailField("correo electrónico", blank=True)
    servicio = models.ForeignKey(
        Servicio, on_delete=models.PROTECT, related_name="reservas", verbose_name="servicio"
    )
    estilista = models.ForeignKey(
        Estilista,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="reservas",
        verbose_name="estilista",
    )
    sucursal = models.ForeignKey(
        Sucursal, on_delete=models.PROTECT, related_name="reservas", verbose_name="sucursal"
    )
    fecha = models.DateField("fecha de la cita")
    hora = models.TimeField("hora de la cita")
    mensaje = models.TextField("mensaje", max_length=400, blank=True)
    estado = models.CharField(
        "estado", max_length=20, choices=Estado.choices, default=Estado.PENDIENTE
    )

    class Meta:
        verbose_name = "reserva"
        verbose_name_plural = "reservas"
        ordering = ["fecha", "hora"]

    def __str__(self):
        return f"{self.nombre} · {self.fecha:%d/%m/%Y} {self.hora:%H:%M}"
