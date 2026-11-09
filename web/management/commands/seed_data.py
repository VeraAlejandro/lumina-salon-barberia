"""
Carga datos de demostración de forma idempotente.

Se usa update_or_create para que el comando pueda ejecutarse varias veces sin
duplicar registros. Las imágenes viven en /media y se enlazan por su ruta.
"""

from datetime import date, time, timedelta
from decimal import Decimal

from django.core.management.base import BaseCommand

from web.models import Cliente, Estilista, Producto, Reserva, Servicio, Sucursal


class Command(BaseCommand):
    help = "Carga sucursales, servicios, estilistas, productos, clientes y reservas de ejemplo."

    def handle(self, *args, **options):
        sucursales = self._sucursales()
        servicios = self._servicios()
        estilistas = self._estilistas(sucursales)
        self._productos()
        self._clientes(sucursales)
        self._reservas(servicios, estilistas, sucursales)

        self.stdout.write(self.style.SUCCESS("Datos de demostración cargados correctamente."))
        self.stdout.write(
            f"  Sucursales: {Sucursal.objects.count()} | "
            f"Servicios: {Servicio.objects.count()} | "
            f"Estilistas: {Estilista.objects.count()} | "
            f"Productos: {Producto.objects.count()} | "
            f"Clientes: {Cliente.objects.count()} | "
            f"Reservas: {Reserva.objects.count()}"
        )

    def _sucursales(self):
        datos = [
            {
                "nombre": "Lumina Universidad",
                "direccion": "Av. Universidad 1200, Col. Universidad",
                "ciudad": "San Nicolás de los Garza",
                "telefono": "81 1234 5678",
                "correo": "universidad@lumina.mx",
                "horario": "Lun a Sáb · 9:00 a 20:00",
                "imagen": "sucursales/sucursal-1.jpg",
                "mapa_url": "https://maps.google.com/?q=Av.+Universidad+1200+San+Nicolas",
            },
            {
                "nombre": "Lumina Centro",
                "direccion": "Padre Mier 450, Centro",
                "ciudad": "Monterrey",
                "telefono": "81 2345 6789",
                "correo": "centro@lumina.mx",
                "horario": "Lun a Sáb · 10:00 a 20:00",
                "imagen": "sucursales/sucursal-2.jpg",
                "mapa_url": "https://maps.google.com/?q=Padre+Mier+450+Monterrey",
            },
            {
                "nombre": "Lumina Cumbres",
                "direccion": "Paseo de los Leones 2100, Col. Cumbres",
                "ciudad": "Monterrey",
                "telefono": "81 3456 7890",
                "correo": "cumbres@lumina.mx",
                "horario": "Lun a Sáb · 9:00 a 19:00",
                "imagen": "sucursales/sucursal-3.jpg",
                "mapa_url": "https://maps.google.com/?q=Paseo+de+los+Leones+2100+Monterrey",
            },
        ]
        resultado = {}
        for item in datos:
            obj, _ = Sucursal.objects.update_or_create(nombre=item["nombre"], defaults=item)
            resultado[obj.nombre] = obj
        return resultado

    def _servicios(self):
        datos = [
            ("Corte de autor", "cabello", "Corte personalizado con diagnóstico de textura, lavado y peinado final.", "450", 50, "servicios/servicio-corte.jpg", True),
            ("Color y balayage", "cabello", "Diseño de color a medida con técnicas de iluminación y matización profesional.", "1650", 150, "servicios/servicio-color.jpg", True),
            ("Peinado y styling", "cabello", "Peinado para evento con herramientas profesionales y fijación de larga duración.", "550", 45, "servicios/servicio-peinado.jpg", True),
            ("Tratamiento de keratina", "cabello", "Alisado progresivo que reduce el frizz y devuelve brillo sin maltratar la fibra.", "1250", 120, "servicios/servicio-keratina.jpg", False),
            ("Corte y barba clásica", "barberia", "Corte masculino, perfilado de barba con navaja y toalla caliente.", "350", 45, "servicios/servicio-barba.jpg", True),
            ("Afeitado ritual", "barberia", "Afeitado tradicional con aceites, espuma caliente y bálsamo calmante.", "300", 35, "servicios/servicio-afeitado.jpg", False),
            ("Manicure y pedicure", "unas", "Cuidado completo de manos y pies con esmaltado semipermanente opcional.", "480", 75, "servicios/servicio-unas.jpg", True),
            ("Facial hidratante", "facial", "Limpieza profunda, exfoliación suave y mascarilla adaptada a tu tipo de piel.", "750", 60, "servicios/servicio-facial.jpg", True),
            ("Maquillaje profesional", "maquillaje", "Maquillaje social o de novia con productos de larga duración y prueba previa.", "850", 60, "servicios/servicio-maquillaje.jpg", False),
        ]
        resultado = {}
        for nombre, categoria, descripcion, precio, duracion, imagen, destacado in datos:
            obj, _ = Servicio.objects.update_or_create(
                nombre=nombre,
                defaults={
                    "categoria": categoria,
                    "descripcion": descripcion,
                    "precio": Decimal(precio),
                    "duracion_minutos": duracion,
                    "imagen": imagen,
                    "destacado": destacado,
                    "activo": True,
                },
            )
            resultado[obj.nombre] = obj
        return resultado

    def _estilistas(self, sucursales):
        datos = [
            {
                "nombre": "Valeria", "apellidos": "Ríos", "especialidad": "Color y balayage",
                "biografia": "Especialista en colorimetría y corrección de tono. Formación en Madrid y CDMX.",
                "foto": "estilistas/estilista-1.jpg", "sucursal": sucursales["Lumina Cumbres"],
                "correo": "valeria@lumina.mx", "telefono": "81 1111 2222", "anios_experiencia": 9,
            },
            {
                "nombre": "Diego", "apellidos": "Salinas", "especialidad": "Corte y barbería",
                "biografia": "Barbero y colorista masculino. Le obsesionan los fades limpios y las barbas bien perfiladas.",
                "foto": "estilistas/estilista-2.jpg", "sucursal": sucursales["Lumina Universidad"],
                "correo": "diego@lumina.mx", "telefono": "81 2222 3333", "anios_experiencia": 7,
            },
            {
                "nombre": "Mariana", "apellidos": "Treviño", "especialidad": "Uñas y estética",
                "biografia": "Manicurista certificada en técnicas rusas y cuidado de piel. Detallista y paciente.",
                "foto": "estilistas/estilista-3.jpg", "sucursal": sucursales["Lumina Centro"],
                "correo": "mariana@lumina.mx", "telefono": "81 3333 4444", "anios_experiencia": 5,
            },
            {
                "nombre": "André", "apellidos": "Castillo", "especialidad": "Barbería clásica",
                "biografia": "Maestro barbero con técnica tradicional de navaja. Formador del equipo junior.",
                "foto": "estilistas/estilista-4.jpg", "sucursal": sucursales["Lumina Universidad"],
                "correo": "andre@lumina.mx", "telefono": "81 4444 5555", "anios_experiencia": 11,
            },
        ]
        resultado = []
        for item in datos:
            obj, _ = Estilista.objects.update_or_create(
                nombre=item["nombre"], apellidos=item["apellidos"], defaults=item
            )
            resultado.append(obj)
        return resultado

    def _productos(self):
        datos = [
            ("Shampoo hidratante", "Lumière Pro", "cabello", "Limpia sin resecar y prepara el cabello para el tratamiento.", "420", 24, "productos/producto-1.jpg"),
            ("Acondicionador nutriente", "Lumière Pro", "cabello", "Suaviza la fibra y facilita el desenredo en cabello largo.", "390", 18, "productos/producto-1.jpg"),
            ("Mascarilla de reparación", "Lumière Pro", "cabello", "Tratamiento intensivo semanal para cabello con decoloración.", "480", 0, "productos/producto-5.jpg"),
            ("Pomada de fijación mate", "Nórdica", "barba", "Fijación flexible con acabado natural y aroma amaderado.", "320", 30, "productos/producto-2.jpg"),
            ("Aceite para barba", "Nórdica", "barba", "Hidrata la piel bajo la barba y reduce la comezón.", "350", 12, "productos/producto-2.jpg"),
            ("Sérum facial vitamina C", "Atelier Botánico", "piel", "Ilumina y unifica el tono. Uso diario por la mañana.", "560", 9, "productos/producto-3.jpg"),
            ("Crema hidratante ligera", "Atelier Botánico", "piel", "Textura no grasa para piel mixta o sensible.", "490", 15, "productos/producto-3.jpg"),
            ("Kit de brochas profesional", "Casa Verde", "herramienta", "Set de ocho brochas de maquillaje con estuche de viaje.", "680", 6, "productos/producto-4.jpg"),
        ]
        for nombre, marca, categoria, descripcion, precio, stock, imagen in datos:
            Producto.objects.update_or_create(
                nombre=nombre,
                marca=marca,
                defaults={
                    "categoria": categoria,
                    "descripcion": descripcion,
                    "precio": Decimal(precio),
                    "stock": stock,
                    "imagen": imagen,
                    "activo": True,
                },
            )

    def _clientes(self, sucursales):
        datos = [
            ("Ana Sofía", "García", "81 5500 1100", "ana.garcia@example.com", date(1994, 3, 18), "Lumina Universidad", "Prefiere citas por la tarde."),
            ("Carlos", "Mendoza", "81 5500 2200", "carlos.mendoza@example.com", date(1988, 7, 2), "Lumina Cumbres", "Alergia a tintes con amoniaco."),
            ("Regina", "Lozano", "81 5500 3300", "regina.lozano@example.com", date(2000, 11, 25), "Lumina Centro", "Clienta de manicure cada quince días."),
            ("Jorge", "Treviño", "81 5500 4400", "jorge.trevino@example.com", date(1991, 1, 9), "Lumina Universidad", "Corte cada tres semanas."),
            ("Paulina", "Cantú", "81 5500 5500", "paulina.cantu@example.com", date(1996, 5, 14), "Lumina Cumbres", "Cabello decolorado, usa mascarilla."),
            ("Luis Ángel", "Ramos", "81 5500 6600", "luis.ramos@example.com", date(1985, 9, 30), "Lumina Centro", "Barba cerrada, prefiere navaja."),
            ("Fernanda", "Villarreal", "81 5500 7700", "fernanda.villarreal@example.com", date(1999, 12, 6), "Lumina Universidad", "Maquillaje para eventos."),
            ("Ricardo", "Elizondo", "81 5500 8800", "ricardo.elizondo@example.com", date(1979, 4, 21), "Lumina Cumbres", "Cliente frecuente de afeitado ritual."),
        ]
        for nombre, apellidos, telefono, correo, nacimiento, sucursal, notas in datos:
            Cliente.objects.update_or_create(
                telefono=telefono,
                defaults={
                    "nombre": nombre,
                    "apellidos": apellidos,
                    "correo": correo,
                    "fecha_nacimiento": nacimiento,
                    "sucursal_preferida": sucursales[sucursal],
                    "notas": notas,
                    "activo": True,
                },
            )

    def _reservas(self, servicios, estilistas, sucursales):
        hoy = date.today()
        datos = [
            ("Ana Sofía García", "81 5500 1100", "ana.garcia@example.com", "Color y balayage", "Valeria Ríos", "Lumina Cumbres", hoy + timedelta(days=2), time(11, 0), "Busco un tono cenizo.", "confirmada"),
            ("Jorge Treviño", "81 5500 4400", "jorge.trevino@example.com", "Corte y barba clásica", "André Castillo", "Lumina Universidad", hoy + timedelta(days=3), time(17, 30), "", "pendiente"),
            ("Regina Lozano", "81 5500 3300", "regina.lozano@example.com", "Manicure y pedicure", "Mariana Treviño", "Lumina Centro", hoy + timedelta(days=4), time(12, 0), "Esmaltado en tono nude.", "confirmada"),
            ("Luis Ángel Ramos", "81 5500 6600", "luis.ramos@example.com", "Afeitado ritual", "Diego Salinas", "Lumina Centro", hoy + timedelta(days=1), time(19, 0), "", "pendiente"),
        ]
        for nombre, telefono, correo, servicio, estilista, sucursal, fecha, hora, mensaje, estado in datos:
            Reserva.objects.update_or_create(
                nombre=nombre,
                fecha=fecha,
                defaults={
                    "telefono": telefono,
                    "correo": correo,
                    "servicio": servicios[servicio],
                    "estilista": next(e for e in estilistas if e.nombre_completo == estilista),
                    "sucursal": sucursales[sucursal],
                    "hora": hora,
                    "mensaje": mensaje,
                    "estado": estado,
                },
            )
