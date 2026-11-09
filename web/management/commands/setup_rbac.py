"""
Configura los tres usuarios y los grupos del modelo RBAC.

Reglas del PIA:
- Administrador: superusuario con acceso total.
- Supervisor: únicamente permisos "add" en todos los modelos.
- Empleado: únicamente permisos "view" en todos los modelos.

El comando es idempotente: puede ejecutarse las veces necesarias y siempre
deja los permisos en el estado esperado.
"""

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group, Permission
from django.core.management.base import BaseCommand

Usuario = get_user_model()

MODELOS = ["servicio", "estilista", "producto", "sucursal", "cliente", "reserva"]
ACCIONES_ADMIN = ["add", "change", "delete", "view"]
ACCIONES_SUPERVISOR = ["add"]
ACCIONES_EMPLEADO = ["view"]

CREDENCIALES = [
    {
        "username": "admin",
        "password": "123",
        "first_name": "Administrador",
        "last_name": "General",
        "email": "admin@lumina.mx",
        "is_superuser": True,
        "grupo": "Administrador",
    },
    {
        "username": "supervisor",
        "password": "123",
        "first_name": "Supervisor",
        "last_name": "de Operaciones",
        "email": "supervisor@lumina.mx",
        "is_superuser": False,
        "grupo": "Supervisor",
    },
    {
        "username": "empleado",
        "password": "123",
        "first_name": "Empleado",
        "last_name": "de Recepción",
        "email": "empleado@lumina.mx",
        "is_superuser": False,
        "grupo": "Empleado",
    },
]


class Command(BaseCommand):
    help = "Crea los grupos y usuarios del modelo RBAC (contraseña 123)."

    def handle(self, *args, **options):
        self.verboso = options.get("verbosity", 1) > 0
        permisos = self._permisos()

        grupos = {
            "Administrador": self._crear_grupo(
                "Administrador", permisos, MODELOS, ACCIONES_ADMIN,
                "Acceso total al panel (el usuario administrador es superusuario).",
            ),
            "Supervisor": self._crear_grupo(
                "Supervisor", permisos, MODELOS, ACCIONES_SUPERVISOR,
                "Solo puede agregar registros en todos los modelos.",
            ),
            "Empleado": self._crear_grupo(
                "Empleado", permisos, MODELOS, ACCIONES_EMPLEADO,
                "Solo puede consultar registros de todos los modelos.",
            ),
        }

        for datos in CREDENCIALES:
            usuario, creado = Usuario.objects.update_or_create(
                username=datos["username"],
                defaults={
                    "first_name": datos["first_name"],
                    "last_name": datos["last_name"],
                    "email": datos["email"],
                    "is_staff": True,
                    "is_active": True,
                    "is_superuser": datos["is_superuser"],
                },
            )
            usuario.set_password(datos["password"])
            usuario.save()
            usuario.groups.set([grupos[datos["grupo"]]])
            estado = "creado" if creado else "actualizado"
            if self.verboso:
                self.stdout.write(f"  Usuario {usuario.username:<12} {estado:<12} grupo: {datos['grupo']}")

        if self.verboso:
            self.stdout.write("")
            self.stdout.write(self.style.SUCCESS("RBAC configurado. Credenciales (todas con contraseña 123):"))
            for datos in CREDENCIALES:
                self.stdout.write(f"  {datos['username']:<12} -> {datos['grupo']}")
            self.stdout.write("")

    def _permisos(self):
        """Devuelve un índice {(accion, modelo): Permission} para la app web."""
        consulta = Permission.objects.filter(content_type__app_label="web")
        indice = {}
        for permiso in consulta:
            accion, _, modelo = permiso.codename.partition("_")
            indice[(accion, modelo)] = permiso
        return indice

    def _crear_grupo(self, nombre, permisos, modelos, acciones, descripcion):
        grupo, _ = Group.objects.get_or_create(name=nombre)
        seleccionados = [
            permisos[(accion, modelo)]
            for modelo in modelos
            for accion in acciones
            if (accion, modelo) in permisos
        ]
        grupo.permissions.set(seleccionados)
        if self.verboso:
            self.stdout.write(
                f"  Grupo {nombre:<12} {len(seleccionados):>2} permisos | {descripcion}"
            )
        return grupo
