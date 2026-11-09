"""
Verifica en vivo que los permisos del modelo RBAC se apliquen.

Usa el cliente de pruebas de Django para iniciar sesión con cada usuario y
comparar el código HTTP obtenido contra el esperado. Sirve como evidencia para
el documento y como prueba rápida después de configurar el RBAC.
"""

from django.conf import settings
from django.contrib.auth import get_user_model
from django.test import Client
from django.core.management.base import BaseCommand

from web.models import Servicio

Usuario = get_user_model()


class Command(BaseCommand):
    help = "Comprueba que administrador, supervisor y empleado respeten la matriz RBAC."

    def handle(self, *args, **options):
        # Sin DEBUG para evitar que Django registre tracebacks de los 403 esperados.
        debug_original = settings.DEBUG
        settings.DEBUG = False
        try:
            self._verificar()
        finally:
            settings.DEBUG = debug_original

    def _verificar(self):
        if not Servicio.objects.exists():
            self.stdout.write(self.style.WARNING(
                "No hay servicios registrados. Ejecuta primero: python manage.py seed_data"
            ))
            return

        servicio = Servicio.objects.first()
        rutas = {
            "Panel /admin/": "/admin/",
            "Listado servicios": "/admin/web/servicio/",
            "Alta servicio": "/admin/web/servicio/add/",
            f"Editar servicio #{servicio.pk}": f"/admin/web/servicio/{servicio.pk}/change/",
            f"Eliminar servicio #{servicio.pk}": f"/admin/web/servicio/{servicio.pk}/delete/",
            "Listado clientes": "/admin/web/cliente/",
            "Alta cliente": "/admin/web/cliente/add/",
        }

        esperado = {
            "admin": {ruta: 200 for ruta in rutas},
            "supervisor": {
                "Panel /admin/": 200,
                "Listado servicios": 403,
                "Alta servicio": 200,
                f"Editar servicio #{servicio.pk}": 403,
                f"Eliminar servicio #{servicio.pk}": 403,
                "Listado clientes": 403,
                "Alta cliente": 200,
            },
            "empleado": {
                "Panel /admin/": 200,
                "Listado servicios": 200,
                "Alta servicio": 403,
                f"Editar servicio #{servicio.pk}": 200,
                f"Eliminar servicio #{servicio.pk}": 403,
                "Listado clientes": 200,
                "Alta cliente": 403,
            },
        }

        fallos = 0
        for username, expectativas in esperado.items():
            self.stdout.write("")
            self.stdout.write(self.style.MIGRATE_HEADING(f"Usuario: {username}"))
            cliente = Client(SERVER_NAME="localhost")
            if not cliente.login(username=username, password="123"):
                self.stdout.write(self.style.ERROR("  No fue posible iniciar sesión."))
                fallos += 1
                continue

            for etiqueta, codigo_esperado in expectativas.items():
                respuesta = cliente.get(rutas[etiqueta])
                correcto = respuesta.status_code == codigo_esperado
                if not correcto:
                    fallos += 1
                marca = "OK " if correcto else "FALLO"
                estilo = self.style.SUCCESS if correcto else self.style.ERROR
                self.stdout.write(
                    estilo(
                        f"  [{marca}] {etiqueta:<30} esperado {codigo_esperado:>3} | obtenido {respuesta.status_code:>3}"
                    )
                )
            cliente.logout()

        self.stdout.write("")
        if fallos:
            self.stdout.write(self.style.ERROR(f"Verificación terminada con {fallos} fallo(s)."))
            raise SystemExit(1)
        self.stdout.write(self.style.SUCCESS("Verificación correcta: la matriz RBAC se aplica en su totalidad."))
