"""Revisión visual rápida del sitio público (uso interno durante el desarrollo)."""

from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = "http://127.0.0.1:8000"
SALIDA = Path("capturas/revision")


def cargar_imagenes_perezosas(pagina):
    """Recorre la página para que las imágenes con loading=lazy se descarguen."""
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
                }, 40);
            });
        }"""
    )
    pagina.wait_for_timeout(700)
    pagina.evaluate("window.scrollTo(0, 0)")
    pagina.wait_for_timeout(300)


def main():
    SALIDA.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        navegador = p.chromium.launch(channel="msedge", headless=True)
        pagina = navegador.new_page(viewport={"width": 1440, "height": 950}, device_scale_factor=1)
        for nombre, ruta in [
            ("inicio", "/"),
            ("servicios", "/servicios/"),
            ("estilistas", "/estilistas/"),
            ("productos", "/productos/"),
            ("sucursales", "/sucursales/"),
            ("reservar", "/reservar/"),
            ("contacto", "/contacto/"),
        ]:
            pagina.goto(BASE + ruta, wait_until="networkidle")
            cargar_imagenes_perezosas(pagina)
            pagina.screenshot(path=str(SALIDA / f"{nombre}.jpg"), full_page=True, type="jpeg", quality=80)
            print("capturado", nombre)

        pagina_movil = navegador.new_page(viewport={"width": 390, "height": 844}, device_scale_factor=1)
        pagina_movil.goto(BASE + "/", wait_until="networkidle")
        cargar_imagenes_perezosas(pagina_movil)
        pagina_movil.screenshot(path=str(SALIDA / "inicio_movil.jpg"), full_page=True, type="jpeg", quality=80)
        print("capturado inicio_movil")
        navegador.close()


if __name__ == "__main__":
    main()
