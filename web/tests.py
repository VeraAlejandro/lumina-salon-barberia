"""
Pruebas del sitio público y de la matriz RBAC.

Se ejecutan con: python manage.py test
"""

from datetime import date, time, timedelta
from decimal import Decimal

from django.contrib.auth import get_user_model
from django.core.management import call_command
from django.test import TestCase
from django.urls import reverse

from .models import Cliente, Reserva, Servicio, Sucursal

Usuario = get_user_model()


class BaseDatosTest(TestCase):
    @classmethod
    def setUpTestData(cls):
        call_command("setup_rbac", verbosity=0)
        cls.sucursal = Sucursal.objects.create(
            nombre="Lumina Test",
            direccion="Calle de prueba 123",
            ciudad="Monterrey",
            telefono="81 0000 0000",
        )
        cls.servicio = Servicio.objects.create(
            nombre="Corte de prueba",
            categoria=Servicio.Categoria.CABELLO,
            descripcion="Servicio creado para las pruebas automatizadas.",
            precio=Decimal("450.00"),
            duracion_minutos=50,
            destacado=True,
        )


class SitioPublicoTest(BaseDatosTest):
    def test_paginas_publicas_responden(self):
        nombres = ["inicio", "servicios", "estilistas", "productos", "sucursales", "reservar", "contacto"]
        for nombre in nombres:
            with self.subTest(pagina=nombre):
                respuesta = self.client.get(reverse(f"web:{nombre}"))
                self.assertEqual(respuesta.status_code, 200)

    def test_servicio_de_la_base_se_muestra_en_el_sitio(self):
        respuesta = self.client.get(reverse("web:servicios"))
        self.assertContains(respuesta, "Corte de prueba")

    def test_inicio_muestra_servicios_destacados(self):
        respuesta = self.client.get(reverse("web:inicio"))
        self.assertContains(respuesta, "Corte de prueba")

    def test_reserva_publica_crea_registro(self):
        fecha = date.today() + timedelta(days=2)
        respuesta = self.client.post(
            reverse("web:reservar"),
            {
                "nombre": "Cliente de Prueba",
                "telefono": "81 9999 9999",
                "correo": "cliente@example.com",
                "sucursal": self.sucursal.pk,
                "servicio": self.servicio.pk,
                "estilista": "",
                "fecha": fecha.isoformat(),
                "hora": "11:00",
                "mensaje": "Prueba automatizada",
            },
        )
        self.assertRedirects(respuesta, reverse("web:reservar"))
        self.assertTrue(
            Reserva.objects.filter(nombre="Cliente de Prueba", fecha=fecha).exists()
        )

    def test_reserva_rechaza_fechas_pasadas(self):
        ayer = date.today() - timedelta(days=1)
        respuesta = self.client.post(
            reverse("web:reservar"),
            {
                "nombre": "Cliente Pasado",
                "telefono": "81 9999 9999",
                "sucursal": self.sucursal.pk,
                "servicio": self.servicio.pk,
                "fecha": ayer.isoformat(),
                "hora": "11:00",
            },
        )
        self.assertEqual(respuesta.status_code, 200)
        self.assertFalse(Reserva.objects.filter(nombre="Cliente Pasado").exists())


class RBACTest(BaseDatosTest):
    def _login(self, username):
        self.assertTrue(self.client.login(username=username, password="123"))

    def test_administrador_tiene_acceso_total(self):
        self._login("admin")
        for ruta in [
            "/admin/",
            "/admin/web/servicio/",
            "/admin/web/servicio/add/",
            f"/admin/web/servicio/{self.servicio.pk}/change/",
            f"/admin/web/servicio/{self.servicio.pk}/delete/",
            "/admin/web/cliente/",
        ]:
            with self.subTest(ruta=ruta):
                self.assertEqual(self.client.get(ruta).status_code, 200)

    def test_supervisor_solo_puede_agregar(self):
        self._login("supervisor")
        self.assertEqual(self.client.get("/admin/").status_code, 200)
        self.assertEqual(self.client.get("/admin/web/servicio/").status_code, 403)
        self.assertEqual(self.client.get("/admin/web/servicio/add/").status_code, 200)
        self.assertEqual(
            self.client.get(f"/admin/web/servicio/{self.servicio.pk}/change/").status_code,
            403,
        )
        self.assertEqual(
            self.client.get(f"/admin/web/servicio/{self.servicio.pk}/delete/").status_code,
            403,
        )

    def test_supervisor_puede_guardar_y_regresa_al_alta(self):
        self._login("supervisor")
        respuesta = self.client.post(
            "/admin/web/servicio/add/",
            {
                "nombre": "Servicio del supervisor",
                "categoria": "barberia",
                "descripcion": "Creado desde la prueba de permisos.",
                "precio": "300",
                "duracion_minutos": "30",
                "destacado": "",
                "activo": "on",
            },
        )
        self.assertRedirects(respuesta, "/admin/web/servicio/add/")
        self.assertTrue(Servicio.objects.filter(nombre="Servicio del supervisor").exists())

    def test_empleado_solo_puede_consultar(self):
        self._login("empleado")
        self.assertEqual(self.client.get("/admin/").status_code, 200)
        self.assertEqual(self.client.get("/admin/web/servicio/").status_code, 200)
        self.assertEqual(
            self.client.get(f"/admin/web/servicio/{self.servicio.pk}/change/").status_code,
            200,
        )
        self.assertEqual(self.client.get("/admin/web/servicio/add/").status_code, 403)
        self.assertEqual(
            self.client.get(f"/admin/web/servicio/{self.servicio.pk}/delete/").status_code,
            403,
        )

    def test_empleado_no_puede_crear_por_post(self):
        self._login("empleado")
        respuesta = self.client.post(
            "/admin/web/servicio/add/",
            {
                "nombre": "Intento no permitido",
                "categoria": "cabello",
                "descripcion": "No debe crearse.",
                "precio": "100",
                "duracion_minutos": "20",
            },
        )
        self.assertEqual(respuesta.status_code, 403)
        self.assertFalse(Servicio.objects.filter(nombre="Intento no permitido").exists())

    def test_anonimo_es_redirigido_al_login(self):
        respuesta = self.client.get("/admin/")
        self.assertEqual(respuesta.status_code, 302)
        self.assertIn("/admin/login/", respuesta.url)

    def test_usuarios_tienen_password_123(self):
        for username in ["admin", "supervisor", "empleado"]:
            with self.subTest(usuario=username):
                usuario = Usuario.objects.get(username=username)
                self.assertTrue(usuario.check_password("123"))
                self.assertTrue(usuario.is_staff)

    def test_supervisor_y_empleado_no_son_superusuarios(self):
        self.assertFalse(Usuario.objects.get(username="supervisor").is_superuser)
        self.assertFalse(Usuario.objects.get(username="empleado").is_superuser)


class ModelosTest(BaseDatosTest):
    def test_str_de_modelos(self):
        self.assertEqual(str(self.sucursal), "Lumina Test")
        self.assertEqual(str(self.servicio), "Corte de prueba")

    def test_cliente_str(self):
        cliente = Cliente.objects.create(
            nombre="Ana", apellidos="García", telefono="81 1111 1111"
        )
        self.assertEqual(str(cliente), "Ana García")
