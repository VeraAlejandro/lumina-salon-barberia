"""
Genera las capturas que se insertan en el documento Word.

Requiere el servidor de desarrollo corriendo en http://127.0.0.1:8000.
Usa Microsoft Edge en modo headless (no descarga navegadores adicionales).
"""

import os
import sys
from pathlib import Path

import django

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "salon.settings")
# Playwright sync crea un event loop; se permite el ORM para preparar las capturas.
os.environ["DJANGO_ALLOW_ASYNC_UNSAFE"] = "true"
django.setup()

from django.contrib.auth import get_user_model  # noqa: E402
from django.contrib.auth.models import Group  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

from web.models import Cliente, Estilista, Producto, Reserva, Servicio, Sucursal  # noqa: E402

BASE = "http://127.0.0.1:8000"
SALIDA = Path("capturas")
VIEWPORT = {"width": 1440, "height": 900}
ESCALA = 1.25


def cargar_imagenes_perezosas(pagina):
    pagina.evaluate(
        """async () => {
            await new Promise((resolve) => {
                let recorrido = 0;
                const paso = 350;
                const temporizador = setInterval(() => {
                    window.scrollBy(0, paso);
                    recorrido += paso;
                    if (recorrido >= document.body.scrollHeight) {
                        clearInterval(temporizador);
                        resolve();
                    }
                }, 35);
            });
        }"""
    )
    pagina.wait_for_timeout(500)
    pagina.evaluate("window.scrollTo(0, 0)")
    pagina.wait_for_timeout(200)


def captura(pagina, nombre, completa=False):
    pagina.wait_for_timeout(350)
    pagina.screenshot(
        path=str(SALIDA / f"{nombre}.png"),
        full_page=completa,
    )
    print("  ", nombre)


def iniciar_sesion(pagina, usuario):
    pagina.goto(f"{BASE}/admin/login/", wait_until="networkidle")
    pagina.fill("#id_username", usuario)
    pagina.fill("#id_password", "123")
    pagina.click("input[type=submit]")
    pagina.wait_for_url("**/admin/", timeout=15000)
    pagina.wait_for_timeout(400)


def capturas_publicas(navegador):
    print("Sitio público")
    contexto = navegador.new_context(viewport=VIEWPORT, device_scale_factor=ESCALA)
    pagina = contexto.new_page()
    paginas = [
        ("publico_01_inicio", "/"),
        ("publico_02_servicios", "/servicios/"),
        ("publico_03_estilistas", "/estilistas/"),
        ("publico_04_productos", "/productos/"),
        ("publico_05_sucursales", "/sucursales/"),
        ("publico_06_reservar", "/reservar/"),
        ("publico_07_contacto", "/contacto/"),
    ]
    for nombre, ruta in paginas:
        pagina.goto(BASE + ruta, wait_until="networkidle")
        cargar_imagenes_perezosas(pagina)
        captura(pagina, nombre)
    contexto.close()

    contexto_movil = navegador.new_context(
        viewport={"width": 390, "height": 844}, device_scale_factor=2
    )
    pagina_movil = contexto_movil.new_page()
    pagina_movil.goto(BASE + "/", wait_until="networkidle")
    cargar_imagenes_perezosas(pagina_movil)
    captura(pagina_movil, "publico_08_inicio_movil")
    contexto_movil.close()


def capturas_admin(navegador):
    print("Administración y RBAC")
    User = get_user_model()
    servicio = Servicio.objects.first()
    estilista = Estilista.objects.first()
    producto = Producto.objects.first()
    sucursal = Sucursal.objects.first()
    cliente = Cliente.objects.first()
    reserva = Reserva.objects.first()
    grupo_supervisor = Group.objects.get(name="Supervisor")
    usuario_supervisor = User.objects.get(username="supervisor")

    objetos = [
        ("servicios", "servicio", servicio),
        ("estilistas", "estilista", estilista),
        ("productos", "producto", producto),
        ("sucursales", "sucursal", sucursal),
        ("clientes", "cliente", cliente),
        ("reservas", "reserva", reserva),
    ]

    # --- Administrador ---
    contexto = navegador.new_context(viewport=VIEWPORT, device_scale_factor=ESCALA)
    pagina = contexto.new_page()
    pagina.goto(f"{BASE}/admin/login/", wait_until="networkidle")
    captura(pagina, "admin_00_login")

    iniciar_sesion(pagina, "admin")
    pagina.wait_for_selector("#content")
    captura(pagina, "admin_01_panel")

    pagina.goto(f"{BASE}/admin/auth/user/", wait_until="networkidle")
    pagina.wait_for_selector("#result_list")
    captura(pagina, "admin_02_usuarios")

    pagina.goto(f"{BASE}/admin/auth/group/", wait_until="networkidle")
    pagina.wait_for_selector("#result_list")
    captura(pagina, "admin_03_grupos")

    pagina.goto(f"{BASE}/admin/auth/group/{grupo_supervisor.pk}/change/", wait_until="networkidle")
    pagina.wait_for_selector("form")
    captura(pagina, "admin_04_permisos_supervisor", completa=True)

    pagina.goto(f"{BASE}/admin/auth/user/{usuario_supervisor.pk}/change/", wait_until="networkidle")
    pagina.wait_for_selector("form")
    captura(pagina, "admin_05_usuario_supervisor", completa=True)

    numero = 6
    for etiqueta, modelo, obj in objetos:
        pagina.goto(f"{BASE}/admin/web/{modelo}/", wait_until="networkidle")
        pagina.wait_for_selector("#result_list")
        captura(pagina, f"admin_{numero:02d}_{etiqueta}_lista")
        numero += 1

        pagina.goto(f"{BASE}/admin/web/{modelo}/add/", wait_until="networkidle")
        pagina.wait_for_selector("form")
        captura(pagina, f"admin_{numero:02d}_{etiqueta}_alta", completa=True)
        numero += 1

        pagina.goto(f"{BASE}/admin/web/{modelo}/{obj.pk}/change/", wait_until="networkidle")
        pagina.wait_for_selector("form")
        captura(pagina, f"admin_{numero:02d}_{etiquetas_editar(etiqueta)}", completa=True)
        numero += 1

        pagina.goto(f"{BASE}/admin/web/{modelo}/{obj.pk}/delete/", wait_until="networkidle")
        captura(pagina, f"admin_{numero:02d}_{etiquetas_eliminar(etiqueta)}")
        numero += 1

    contexto.close()

    # --- Supervisor (solo alta) ---
    contexto = navegador.new_context(viewport=VIEWPORT, device_scale_factor=ESCALA)
    pagina = contexto.new_page()
    iniciar_sesion(pagina, "supervisor")
    pagina.wait_for_selector("#content")
    captura(pagina, "rbac_01_supervisor_panel")
    pagina.goto(f"{BASE}/admin/web/servicio/add/", wait_until="networkidle")
    pagina.wait_for_selector("form")
    captura(pagina, "rbac_02_supervisor_alta", completa=True)
    contexto.close()

    # --- Empleado (solo consulta) ---
    contexto = navegador.new_context(viewport=VIEWPORT, device_scale_factor=ESCALA)
    pagina = contexto.new_page()
    iniciar_sesion(pagina, "empleado")
    pagina.wait_for_selector("#content")
    captura(pagina, "rbac_03_empleado_panel")
    pagina.goto(f"{BASE}/admin/web/servicio/", wait_until="networkidle")
    pagina.wait_for_selector("#result_list")
    captura(pagina, "rbac_04_empleado_lista")
    pagina.goto(f"{BASE}/admin/web/servicio/{servicio.pk}/change/", wait_until="networkidle")
    pagina.wait_for_selector("form")
    captura(pagina, "rbac_05_empleado_detalle", completa=True)
    contexto.close()


def etiquetas_editar(etiqueta):
    singular = {
        "servicios": "servicio",
        "estilistas": "estilista",
        "productos": "producto",
        "sucursales": "sucursal",
        "clientes": "cliente",
        "reservas": "reserva",
    }[etiqueta]
    return f"{singular}_editar"


def etiquetas_eliminar(etiqueta):
    singular = {
        "servicios": "servicio",
        "estilistas": "estilista",
        "productos": "producto",
        "sucursales": "sucursal",
        "clientes": "cliente",
        "reservas": "reserva",
    }[etiqueta]
    return f"{singular}_eliminar"


def main():
    SALIDA.mkdir(exist_ok=True)
    with sync_playwright() as p:
        navegador = p.chromium.launch(channel="msedge", headless=True)
        capturas_publicas(navegador)
        capturas_admin(navegador)
        navegador.close()
    print("Capturas generadas en /capturas")


if __name__ == "__main__":
    main()
