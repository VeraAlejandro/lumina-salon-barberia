"""
Genera el documento Word del Producto Integrador de Aprendizaje de Lumina.

Uso:
    .\\.venv\\Scripts\\python.exe scripts\\generar_documento.py

El resultado es docs/PIA_Lumina.docx. Después se actualiza el índice y se
exporta el PDF con scripts/actualizar_indice.py (requiere Microsoft Word).
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

import docx_util as u
from docx_util import (
    CAPTURAS_DOC,
    caja,
    codigo,
    configurar_estilos,
    espacio,
    figura,
    h1,
    h2,
    h3,
    indice,
    iniciar_cuerpo,
    metadatos,
    par,
    pasos,
    portada,
    preparar_capturas,
    tabla,
    vinetas,
)

INTEGRANTES = [
    ("Diego Alejandro Vera Reyes", "[Matrícula]"),
    ("[Integrante 2 — Nombre completo]", "[Matrícula]"),
    ("[Integrante 3 — Nombre completo]", "[Matrícula]"),
    ("[Integrante 4 — Nombre completo]", "[Matrícula]"),
    ("[Integrante 5 — Nombre completo]", "[Matrícula]"),
    ("[Integrante 6 — Nombre completo]", "[Matrícula]"),
]


# ---------------------------------------------------------------------------
# 1. Introducción
# ---------------------------------------------------------------------------

def seccion_1(doc):
    h1(doc, "1. Introducción", salto=False)

    h2(doc, "1.1 Contexto y propósito")
    par(doc, (
        "El sector de la belleza y el cuidado personal se apoya cada vez más en canales digitales: "
        "el cliente descubre servicios en redes sociales, compara precios, revisa reseñas y espera "
        "reservar en pocos pasos. Sin embargo, muchos salones pequeños y medianos siguen operando con "
        "mensajes dispersos, agendas en papel y hojas de cálculo. Esa combinación provoca citas duplicadas, "
        "pérdida de clientes, poca visibilidad del inventario y una experiencia inconsistente entre sucursales."
    ))
    par(doc, (
        "**Lumina** se plantea como un salón de belleza y barbería con tres sucursales en el área "
        "metropolitana de Monterrey. Para este Producto Integrador de Aprendizaje (PIA) se desarrolló un "
        "sitio web dinámico con Django que cumple dos funciones complementarias. La primera es una vitrina "
        "pública de siete páginas donde el visitante conoce los servicios, al equipo, los productos y las "
        "sucursales, y solicita una cita en línea. La segunda es un panel administrativo protegido por "
        "autenticación y permisos, desde el cual el personal administra los seis catálogos del negocio y da "
        "seguimiento a las reservas."
    ))
    par(doc, (
        "El proyecto no se limita a “mostrar páginas”: conecta el sitio público con la base de datos. Los "
        "servicios, estilistas, productos, sucursales y reservas que se ven en el sitio provienen de SQLite y "
        "se actualizan desde el panel. Además, el acceso se segmenta con un modelo de control basado en roles "
        "(RBAC, por sus siglas en inglés) que distingue tres perfiles: Administrador, Supervisor y Empleado. "
        "Los datos, nombres, precios e imágenes incluidos son de demostración y pueden reemplazarse sin "
        "modificar el código."
    ))

    h2(doc, "1.2 Objetivos")
    par(doc, (
        "**Objetivo general.** Diseñar e implementar un sitio web dinámico con Django para Lumina que integre "
        "una presencia digital profesional, la captación de citas en línea y un sistema administrativo con "
        "control de acceso basado en roles, aplicando buenas prácticas de desarrollo web y documentando su "
        "operación."
    ))
    par(doc, "**Objetivos específicos.**", despues=3)
    vinetas(doc, [
        "Publicar un sitio público de siete páginas con diseño responsivo, identidad visual propia y contenido alimentado desde la base de datos.",
        "Modelar las entidades del negocio —servicios, estilistas, productos, sucursales y clientes— y una entidad adicional de reservas que conecta la captación pública con la operación interna.",
        "Implementar un panel administrativo que permita crear, consultar, editar y eliminar registros, con validaciones y vistas previas de imágenes.",
        "Configurar tres roles con privilegios distintos y verificar su comportamiento mediante pruebas automatizadas y una matriz de permisos reproducible.",
        "Documentar los requisitos, la oportunidad de negocio, la arquitectura, el modelo RBAC y el manual de usuario final con evidencia visual.",
    ])

    h2(doc, "1.3 Justificación")
    par(doc, (
        "Se eligió Django porque resuelve de forma integrada el enrutamiento, el acceso a datos, la "
        "autenticación, la administración y la protección contra ataques comunes como CSRF. Esta decisión "
        "reduce el código repetitivo y permite concentrar el esfuerzo en las reglas del negocio. SQLite se "
        "utilizó como motor de base de datos porque la entrega es académica y debe ejecutarse con un solo "
        "comando; el proyecto está preparado para migrar a PostgreSQL sin reescribir los modelos."
    ))
    par(doc, (
        "La separación de roles no es un adorno: en un negocio real, no todos los colaboradores deben poder "
        "borrar información. El Supervisor captura registros, pero no consulta ni modifica los existentes; el "
        "Empleado consulta la operación, pero no altera datos; el Administrador conserva el control total. "
        "Esta segmentación se implementó con grupos y permisos nativos de Django y se comprobó con pruebas que "
        "esperan respuestas HTTP 200 o 403 según corresponda."
    ))
    par(doc, (
        "Finalmente, el sitio se diseñó con una dirección de arte consistente —carbón, dorado champán y crema, "
        "con tipografías Playfair Display e Inter— para que la experiencia digital transmita la misma calidad "
        "que el servicio presencial. La estética no sustituye la funcionalidad: cada página tiene un objetivo "
        "comercial claro y cada formulario valida los datos antes de guardarlos."
    ))

    h2(doc, "1.4 Alcance del documento")
    par(doc, (
        "El documento se organiza en ocho capítulos y tres anexos. El capítulo 2 enumera los requisitos "
        "funcionales y no funcionales, incluyendo hardware y software. El capítulo 3 desarrolla la oportunidad "
        "de negocio, el alcance y la solución propuesta. El capítulo 4 describe el diseño del sitio, la "
        "identidad visual y el modelo de datos. El capítulo 5 explica la autenticación, la matriz RBAC y su "
        "verificación. El capítulo 6 es el manual de usuario final, con el procedimiento de alta, consulta, "
        "edición y eliminación para cada modelo. Los capítulos 7 y 8 presentan conclusiones y referencias."
    ))
    caja(doc, "Nota sobre los datos", (
        "Los registros, precios, teléfonos, correos y fotografías son de demostración. La marca Lumina y las "
        "tres sucursales se utilizan como caso de estudio académico. Antes de un uso comercial deben sustituirse "
        "por información real, revisarse los avisos de privacidad y configurarse un servidor de producción."
    ))

    h2(doc, "1.5 Metodología y organización del equipo")
    par(doc, (
        "El trabajo se desarrolló en iteraciones cortas: modelo de datos, sitio público, panel administrativo y "
        "RBAC; al final, pruebas, capturas y documentación. Cada iteración cerró con una verificación ejecutable: "
        "migraciones, carga de datos, pruebas y revisión visual. Las capturas se automatizaron con Playwright y se "
        "revisaron en Edge; ninguna de esas herramientas interviene en la operación del sitio."
    ))
    par(doc, (
        "El equipo se organizó por responsabilidades, con revisión cruzada. La tabla usa roles sugeridos y nombres "
        "editables que deben reemplazarse antes de entregar."
    ))
    tabla(
        doc,
        ["Integrante", "Rol sugerido", "Actividades principales"],
        [
            ["Diego Alejandro Vera Reyes", "Backend y liderazgo", "Django, modelos, vistas, RBAC y documento."],
            ["[Integrante 2 — Nombre completo]", "Análisis y documentación", "Requisitos, negocio y revisión de redacción."],
            ["[Integrante 3 — Nombre completo]", "Diseño UX/UI", "Identidad visual, plantillas y diseño responsivo."],
            ["[Integrante 4 — Nombre completo]", "Datos y administración", "Modelo de datos, panel y validaciones."],
            ["[Integrante 5 — Nombre completo]", "Pruebas y calidad", "Pruebas, verificación RBAC y capturas."],
            ["[Integrante 6 — Nombre completo]", "Investigación de negocio", "Contexto del sector y propuesta de valor."],
        ],
        anchos=[2.35, 1.35, 2.85],
        titulo="Organización sugerida del equipo",
        primera_negrita=True,
        fuente="Elaboración propia. Los nombres y matrículas son editables.",
    )



# ---------------------------------------------------------------------------
# 2. Requisitos
# ---------------------------------------------------------------------------

def seccion_2(doc):
    h1(doc, "2. Requisitos del sistema")

    par(doc, (
        "Los requisitos se redactaron pensando en un negocio real y en la evidencia académica. Cada requisito "
        "funcional tiene un identificador para poder rastrearlo en el manual y en las pruebas. Los no "
        "funcionales describen atributos de calidad: usabilidad, seguridad, mantenibilidad, portabilidad y "
        "accesibilidad."
    ))

    h2(doc, "2.1 Requisitos funcionales")
    par(doc, (
        "La siguiente tabla concentra las funciones que el sistema debe cumplir. La prioridad “Alta” indica que sin esa "
        "función el flujo principal del negocio se interrumpe; “Media” son mejoras necesarias para la operación; "
        "“Baja” son detalles de presentación o apoyo."
    ))
    filas_rf = [
        ["RF-01", "Inicio dinámico", "Mostrar servicios destacados, equipo, productos, sucursales y llamadas a la acción; el contenido se consulta desde la base de datos.", "Alta"],
        ["RF-02", "Catálogo de servicios", "Listar servicios activos y filtrarlos por categoría: cabello, barbería, uñas, facial y maquillaje.", "Alta"],
        ["RF-03", "Directorio de estilistas", "Mostrar nombre, especialidad, experiencia, fotografía y sucursal de cada estilista activo.", "Alta"],
        ["RF-04", "Catálogo de productos", "Listar productos activos, filtrar por categoría y buscar por nombre o marca.", "Alta"],
        ["RF-05", "Directorio de sucursales", "Mostrar dirección, teléfono, horario, fotografía y enlace de mapa de cada sucursal activa.", "Alta"],
        ["RF-06", "Contacto y preguntas frecuentes", "Concentrar medios de contacto, ubicaciones y respuestas a dudas recurrentes.", "Media"],
        ["RF-07", "Reserva en línea", "Capturar nombre, teléfono, correo, sucursal, servicio, estilista, fecha, hora y mensaje; guardar la solicitud en la base de datos.", "Alta"],
        ["RF-08", "Validación de reservas", "Rechazar fechas pasadas o del mismo día y evitar que se asigne un estilista de otra sucursal.", "Alta"],
        ["RF-09", "Confirmación de reserva", "Mostrar un mensaje claro con fecha, hora y nombre después de guardar la solicitud.", "Media"],
        ["RF-10", "Autenticación", "Permitir el acceso al panel solo a usuarios activos del personal mediante usuario y contraseña.", "Alta"],
        ["RF-11", "CRUD de servicios", "Crear, consultar, editar y eliminar servicios, con categoría, precio, duración, imagen y estado.", "Alta"],
        ["RF-12", "CRUD de estilistas", "Crear, consultar, editar y eliminar estilistas vinculados a una sucursal.", "Alta"],
        ["RF-13", "CRUD de productos", "Crear, consultar, editar y eliminar productos con marca, categoría, precio, existencias e imagen.", "Alta"],
        ["RF-14", "CRUD de sucursales", "Crear, consultar, editar y eliminar sucursales con dirección, contacto, horario e imagen.", "Alta"],
        ["RF-15", "CRUD de clientes", "Crear, consultar, editar y eliminar clientes con datos de contacto, nacimiento, sucursal preferida y notas.", "Alta"],
        ["RF-16", "CRUD de reservas", "Crear, consultar, editar y eliminar reservas, y cambiar su estado: pendiente, confirmada, atendida o cancelada.", "Media"],
        ["RF-17", "Control por roles", "Aplicar permisos diferenciados para Administrador, Supervisor y Empleado en todos los modelos.", "Alta"],
        ["RF-18", "Alta exclusiva del Supervisor", "Permitir al Supervisor agregar registros y regresarlo al formulario de alta sin exponer listados ni ediciones.", "Alta"],
        ["RF-19", "Consulta exclusiva del Empleado", "Permitir al Empleado consultar listados y fichas, sin botones ni permisos de alta, edición o borrado.", "Alta"],
        ["RF-20", "Panel con identidad", "Aplicar la identidad de Lumina al panel, con encabezado, títulos y tablas legibles.", "Baja"],
        ["RF-21", "Gestión de imágenes", "Permitir la carga de imágenes y mostrar una vista previa en los formularios y listados.", "Media"],
        ["RF-22", "Datos reproducibles", "Incluir comandos para cargar datos de demostración y configurar usuarios, grupos y permisos sin pasos manuales.", "Alta"],
        ["RF-23", "Pruebas automatizadas", "Cubrir páginas públicas, reservas y RBAC con pruebas ejecutables mediante `python manage.py test`.", "Media"],
        ["RF-24", "Diseño responsivo", "Adaptar la navegación y las rejillas a escritorio, tableta y móvil desde 360 px.", "Alta"],
    ]
    tabla(
        doc,
        ["ID", "Requisito", "Descripción", "Prioridad"],
        filas_rf,
        anchos=[0.65, 1.45, 3.9, 0.7],
        titulo="Requisitos funcionales del sistema",
        centrar=(0, 3),
        fuente="Elaboración propia con base en el PIA de Fundamentos de Desarrollo Web.",
    )

    h2(doc, "2.2 Requisitos no funcionales")
    filas_rnf = [
        ["RNF-01", "Usabilidad", "La navegación principal debe ser comprensible sin capacitación y operable con teclado; los formularios deben mostrar errores junto al campo correspondiente."],
        ["RNF-02", "Rendimiento", "Las páginas deben cargar en menos de dos segundos en el entorno local; las imágenes fuera de pantalla usan carga diferida."],
        ["RNF-03", "Seguridad", "Contraseñas almacenadas con hash salteado, sesiones protegidas, validación CSRF y permisos verificados en cada vista del panel."],
        ["RNF-04", "Disponibilidad", "El sistema debe iniciar con un comando y funcionar sin conexión a servicios externos, salvo los enlaces de mapas."],
        ["RNF-05", "Mantenibilidad", "Código organizado por responsabilidades (modelos, vistas, formularios, administración), con nombres claros y comentarios en español."],
        ["RNF-06", "Portabilidad", "Ejecutable en Windows, macOS o Linux con Python 3.12 o superior; la base SQLite no requiere instalación adicional."],
        ["RNF-07", "Accesibilidad", "HTML semántico, textos alternativos en imágenes, contraste suficiente, foco visible y estilos de movimiento reducido."],
        ["RNF-08", "Compatibilidad", "Funcionamiento en las versiones actuales de Edge, Chrome y Firefox, y en pantallas desde 360 px."],
        ["RNF-09", "Escalabilidad", "El modelo y las consultas permiten migrar a PostgreSQL y a un almacenamiento externo de imágenes sin cambiar la interfaz."],
        ["RNF-10", "Trazabilidad", "Cada registro conserva fecha de creación y de última actualización; los cambios quedan registrados en el historial del panel."],
    ]
    tabla(
        doc,
        ["ID", "Atributo", "Criterio de aceptación"],
        filas_rnf,
        anchos=[0.7, 1.3, 4.7],
        titulo="Requisitos no funcionales",
        centrar=(0,),
    )

    h2(doc, "2.3 Requisitos de hardware")
    par(doc, (
        "El sistema es ligero porque Django y SQLite se ejecutan en un solo proceso. Aun así, la generación de "
        "capturas, la edición del documento y la carga de imágenes requieren más recursos que la operación "
        "normal. La tabla distingue el equipo de desarrollo, un servidor de demostración y el equipo del usuario final."
    ))
    tabla(
        doc,
        ["Componente", "Mínimo", "Recomendado", "Observaciones"],
        [
            ["Equipo de desarrollo", "2 núcleos · 4 GB RAM · 3 GB libres", "4 núcleos · 8 GB RAM · SSD con 10 GB libres", "Debe permitir instalar Python, el entorno virtual y un navegador moderno."],
            ["Servidor de demostración", "1 vCPU · 1 GB RAM · 5 GB", "2 vCPU · 2 GB RAM · 20 GB", "Solo para un despliegue futuro; la entrega se ejecuta en local."],
            ["Equipo del usuario final", "Navegador moderno · 2 GB RAM", "4 GB RAM · pantalla de 1366 × 768 o superior", "El sitio funciona en móvil; el panel se recomienda en escritorio."],
            ["Conectividad", "No indispensable en local", "10 Mbps para abrir enlaces de mapas", "Fuentes, estilos e imágenes se sirven desde el propio proyecto."],
        ],
        anchos=[1.35, 1.45, 1.7, 2.1],
        titulo="Requisitos de hardware",
    )

    h2(doc, "2.4 Requisitos de software")
    par(doc, (
        "La tabla separa el software necesario para ejecutar el sitio del software utilizado solo para generar "
        "evidencia o documentación. Esa distinción evita que el profesor tenga que instalar herramientas que no "
        "intervienen en la operación."
    ))
    tabla(
        doc,
        ["Categoría", "Software", "Versión", "Uso"],
        [
            ["Ejecución", "Python", "3.12 o superior (probado en 3.14.7)", "Intérprete del proyecto."],
            ["Ejecución", "Django", "5.2.19", "Framework web, ORM, autenticación y panel administrativo."],
            ["Ejecución", "SQLite", "3", "Base de datos incluida en Python; no requiere servidor."],
            ["Ejecución", "Pillow", "13.0.0", "Procesamiento y validación de imágenes."],
            ["Ejecución", "Navegador web", "Edge, Chrome o Firefox actuales", "Acceso al sitio y al panel."],
            ["Desarrollo", "Visual Studio Code", "Versión actual", "Edición de código y terminal integrada."],
            ["Desarrollo", "Git", "Opcional", "Control de versiones si el equipo lo utiliza."],
            ["Documentación", "python-docx", "1.2.0", "Generación programática de este documento."],
            ["Documentación", "Playwright", "1.64.1", "Automatización de capturas con Microsoft Edge."],
            ["Documentación", "pywin32", "312", "Actualización del índice y exportación a PDF con Word."],
            ["Documentación", "Microsoft Word", "16.0", "Revisión final, actualización de campos y PDF."],
        ],
        anchos=[1.1, 1.45, 1.7, 2.35],
        titulo="Requisitos de software",
        primera_negrita=False,
    )

    h2(doc, "2.5 Arquitectura y stack tecnológico")
    par(doc, (
        "El proyecto sigue el patrón Modelo-Vista-Plantilla (MVT) de Django. El **modelo** define las entidades y "
        "sus relaciones; la **vista** recibe la petición, consulta el ORM y prepara el contexto; la **plantilla** "
        "renderiza HTML con herencia y parciales reutilizables. El panel administrativo se apoya en las mismas "
        "entidades, por lo que un cambio en el modelo se refleja en el sitio público y en la administración."
    ))
    tabla(
        doc,
        ["Capa", "Tecnología", "Responsabilidad en Lumina"],
        [
            ["Presentación", "HTML5 semántico, CSS3, JavaScript", "Estructura accesible, identidad visual, menú móvil y animaciones discretas."],
            ["Aplicación", "Django 5.2 (vistas, formularios, URLs)", "Enrutamiento, validación, mensajes, renderizado y control de acceso."],
            ["Datos", "Django ORM + SQLite", "Seis modelos, relaciones con PROTECT/SET_NULL y campos de auditoría."],
            ["Administración", "Django Admin personalizado", "CRUD de catálogos, filtros, búsqueda, vistas previas e historial."],
            ["Seguridad", "Auth, sesiones, CSRF, permisos", "Login, hash de contraseñas, grupos y verificación por vista."],
            ["Evidencia", "Playwright + python-docx", "Capturas automatizadas y documento reproducible."],
        ],
        anchos=[1.2, 2.0, 3.4],
        titulo="Stack tecnológico y responsabilidades",
    )
    par(doc, (
        "La aplicación se dividió en el proyecto `salon` (configuración global) y la app `web` (dominio del "
        "negocio). Las rutas públicas viven en `web/urls.py` con el espacio de nombres `web`, mientras que el "
        "panel se monta en `/admin/`. Los archivos multimedia se sirven desde `/media/` en desarrollo y las "
        "tipografías, estilos e imágenes de marca desde `/static/`."
    ))


# ---------------------------------------------------------------------------
# 3. Oportunidad de negocio
# ---------------------------------------------------------------------------

def seccion_3(doc):
    h1(doc, "3. Oportunidad de negocio", salto=False)

    h2(doc, "3.1 Contexto del sector")
    par(doc, (
        "Los salones de belleza y las barberías combinan servicios de alto contacto humano con una operación "
        "intensiva en agenda. La decisión de compra depende de la confianza, la recomendación y la claridad de "
        "precios. En ese contexto, un sitio web bien construido no compite contra las redes sociales: las "
        "complementa. Mientras Instagram genera descubrimiento, el sitio concentra información estable, "
        "catálogos actualizados y un canal formal de reserva."
    ))
    par(doc, (
        "El área metropolitana de Monterrey tiene una oferta amplia y clientes acostumbrados a comparar. Los "
        "negocios pequeños suelen perder oportunidades por tres motivos: no publican precios, no muestran al "
        "equipo y no ofrecen un mecanismo de cita distinto de los mensajes directos. Cuando el volumen crece, "
        "la operación se vuelve reactiva: se atiende primero al cliente que escribe más, no al que reservó "
        "primero. La digitalización básica resuelve ese problema sin requerir una inversión enterprise."
    ))

    h2(doc, "3.2 Problema y oportunidad")
    par(doc, "**Problemas identificados en el escenario de Lumina.**", despues=3)
    vinetas(doc, [
        "Información fragmentada: servicios, precios y horarios viven en publicaciones que se pierden con el tiempo.",
        "Reservas sin registro central: las citas se acuerdan por mensajes y no queda una fuente única de verdad.",
        "Inventario sin visibilidad: no se sabe qué productos están agotados o cuáles requieren reposición.",
        "Accesos homogéneos: cualquier persona con la contraseña podría modificar o borrar información sensible.",
        "Experiencia inconsistente entre sucursales: cada una comunica horarios y servicios de forma distinta.",
    ])
    par(doc, "**Oportunidad.**", despues=3)
    vinetas(doc, [
        "Convertir el sitio en el punto oficial de consulta y reserva, con enlaces desde redes sociales y Google Business Profile.",
        "Centralizar los seis catálogos y la agenda en una sola base de datos operada por el equipo.",
        "Reducir el tiempo de respuesta y los errores de agenda al recibir solicitudes estructuradas.",
        "Segmentar el acceso para que cada rol haga solo lo que le corresponde, sin depender de la confianza informal.",
        "Preparar el negocio para métricas: conversión de visitas a reservas, servicios más solicitados y clientes recurrentes.",
    ])

    h2(doc, "3.3 Alcance de la solución")
    par(doc, (
        "El alcance se definió para que el proyecto fuera terminable en el periodo académico y, al mismo tiempo, "
        "útiles para una operación real. La tabla siguiente separa con claridad lo que se entregó de lo que se "
        "propone como siguiente fase; esa honestidad evita prometer funciones que no están implementadas."
    ))
    tabla(
        doc,
        ["Dentro del alcance (implementado)", "Fuera del alcance (trabajo futuro)"],
        [
            ["Sitio público de siete páginas con contenido dinámico.", "Pago en línea y anticipos con pasarela."],
            ["Catálogos de servicios, estilistas, productos, sucursales y clientes.", "Notificaciones automáticas por SMS, WhatsApp o correo."],
            ["Formulario de reserva con validaciones y persistencia.", "Sincronización bidireccional con Google Calendar."],
            ["Panel administrativo con CRUD, filtros, búsqueda e imágenes.", "Facturación electrónica y punto de venta."],
            ["Tres roles RBAC y verificación automatizada.", "Programa de lealtad y cupones."],
            ["Diseño responsivo y accesibilidad básica.", "Aplicación móvil nativa."],
            ["Datos de demostración y comandos reproducibles.", "Despliegue productivo con HTTPS, respaldos y monitoreo."],
        ],
        anchos=[3.3, 3.3],
        titulo="Alcance del proyecto",
    )

    h2(doc, "3.4 Propuesta de valor")
    par(doc, (
        "La propuesta de valor de Lumina se resume en una idea: **el cliente decide con información clara y el "
        "negocio opera con datos ordenados**. El sitio no solo presenta la marca; reduce la fricción entre el "
        "deseo de un servicio y la cita confirmada. Para el equipo interno, el panel sustituye hojas sueltas por "
        "un flujo común con permisos."
    ))
    tabla(
        doc,
        ["Bloque", "Contenido para Lumina"],
        [
            ["Propuesta de valor", "Salón y barbería con servicios profesionales, precios visibles y reserva en línea en tres sucursales."],
            ["Segmentos de clientes", "Mujeres y hombres de 18 a 55 años que valoran técnica, trato cercano y conveniencia digital; clientes recurrentes de las tres sucursales."],
            ["Canales", "Sitio web, Instagram, Facebook, TikTok, Google Business Profile y recomendación boca a boca."],
            ["Relación con el cliente", "Atención personalizada, confirmación telefónica, historial de preferencias y contenido educativo."],
            ["Recursos clave", "Equipo de estilistas certificados, productos profesionales, tres ubicaciones y la plataforma web."],
            ["Actividades clave", "Gestión de citas, ejecución de servicios, control de inventario, contenido y seguimiento."],
            ["Socios clave", "Proveedores de productos profesionales, arrendadores, plataformas de mapas y hosting."],
            ["Fuentes de ingreso", "Servicios de cabello, barbería, uñas, faciales y maquillaje; venta de productos; paquetes y eventos."],
            ["Estructura de costos", "Rentas, nómina, productos, mantenimiento, hosting, marketing y capacitación."],
        ],
        anchos=[1.55, 5.0],
        titulo="Lienzo de propuesta de valor (adaptación de Business Model Canvas)",
        primera_negrita=True,
    )

    h2(doc, "3.5 Modelo de negocio y métricas")
    par(doc, (
        "El modelo es mixto: ingresos por servicios y por venta de productos. La reserva en línea alimenta el "
        "flujo principal y permite medir el interés real, no solo los “me gusta”. Las métricas propuestas son "
        "sencillas y accionables; no requieren una herramienta de analítica avanzada para empezar."
    ))
    tabla(
        doc,
        ["Métrica", "Definición", "Meta inicial", "Fuente"],
        [
            ["Visitas al sitio", "Sesiones únicas mensuales.", "1,500 en tres meses", "Analítica web (fase futura)."],
            ["Solicitudes de reserva", "Formularios enviados correctamente.", "80 al mes", "Modelo Reserva."],
            ["Conversión", "Reservas / visitas.", "5 %", "Cálculo con analítica y base de datos."],
            ["Tasa de confirmación", "Reservas confirmadas / solicitudes.", "70 %", "Estado de la reserva."],
            ["Clientes recurrentes", "Clientes con más de una reserva en 90 días.", "35 %", "Historial de reservas."],
            ["Ticket promedio", "Ingreso total / servicios atendidos.", "$650 MXN", "Reporte interno (fase futura)."],
        ],
        anchos=[1.35, 2.35, 1.5, 1.5],
        titulo="Indicadores propuestos para medir la oportunidad",
        centrar=(2,),
    )

    h2(doc, "3.6 Viabilidad y riesgos")
    par(doc, (
        "**Viabilidad técnica.** El proyecto usa tecnologías gratuitas y maduras. Un equipo de seis personas puede "
        "mantenerlo con conocimientos de Python, HTML, CSS y Django. **Viabilidad económica.** El costo marginal "
        "de operación local es cero; un despliegue pequeño puede costar menos que una campaña mensual de publicidad. "
        "**Viabilidad operativa.** El panel replica tareas cotidianas —capturar, consultar, editar— y el RBAC "
        "reduce el riesgo de errores humanos."
    ))
    tabla(
        doc,
        ["Riesgo", "Impacto", "Mitigación"],
        [
            ["Baja adopción del personal", "Alto", "Capacitación corta, manual con capturas y flujo de alta simplificado para el Supervisor."],
            ["Datos incompletos o duplicados", "Medio", "Validaciones, campos obligatorios, búsqueda y convención de captura."],
            ["Pérdida de información", "Alto", "Respaldo periódico de `db.sqlite3` y de la carpeta `media`; migración a PostgreSQL en producción."],
            ["Contraseñas débiles", "Alto", "La contraseña `123` es solo académica; en producción se exige una política robusta y segundo factor."],
            ["Dependencia de una sola persona", "Medio", "Documentación, comandos reproducibles y revisión cruzada."],
            ["Imágenes sin derechos claros", "Medio", "Créditos de Unsplash y sustitución por fotografía propia antes de uso comercial."],
        ],
        anchos=[2.0, 0.85, 3.7],
        titulo="Matriz de riesgos y mitigaciones",
        centrar=(1,),
    )

    h2(doc, "3.7 Trabajo futuro")
    par(doc, (
        "La siguiente fase debería priorizar tres frentes. Primero, **operación**: notificaciones automáticas, "
        "confirmación de asistencia y sincronización con calendario. Segundo, **negocio**: paquetes, lealtad, "
        "cupones y reportes de ingresos. Tercero, **plataforma**: despliegue con HTTPS, PostgreSQL, almacenamiento "
        "externo de imágenes, respaldos automáticos y monitoreo. Ninguna de esas funciones es requisito del PIA, "
        "pero el modelo actual está preparado para incorporarlas sin rehacer la arquitectura."
    ))
    par(doc, (
        "Un siguiente paso natural es convertir el catálogo en una experiencia de compra con carrito y pago en "
        "línia. Antes de hacerlo conviene medir la demanda real de productos y resolver la logística de entrega "
        "o recolección en sucursal. La tecnología no es el cuello de botella; la operación sí."
    ))


# ---------------------------------------------------------------------------
# 4. Diseño del sitio y del sistema
# ---------------------------------------------------------------------------

def seccion_4(doc):
    h1(doc, "4. Diseño del sitio y del sistema")

    h2(doc, "4.1 Arquitectura de información")
    par(doc, (
        "El sitio público se organiza en siete páginas. Cada una tiene un objetivo y una fuente de datos "
        "identificable, de modo que el contenido no se escriba dos veces ni se desactualice. La navegación "
        "principal prioriza Servicios, Estilistas, Productos y Sucursales; la acción comercial “Reservar cita” "
        "aparece siempre como botón dorado y, en móvil, como botón flotante."
    ))
    tabla(
        doc,
        ["Página", "Ruta", "Objetivo", "Datos que utiliza"],
        [
            ["Inicio", "/", "Presentar la marca y dirigir a la reserva.", "Servicio, Estilista, Producto, Sucursal"],
            ["Servicios", "/servicios/", "Mostrar catálogo, precios y duración con filtros.", "Servicio"],
            ["Estilistas", "/estilistas/", "Generar confianza mostrando al equipo.", "Estilista, Sucursal"],
            ["Productos", "/productos/", "Exhibir productos y disponibilidad.", "Producto"],
            ["Sucursales", "/sucursales/", "Facilitar la visita física y el contacto.", "Sucursal, Estilista"],
            ["Reservar", "/reservar/", "Capturar una solicitud de cita validada.", "Servicio, Estilista, Sucursal, Reserva"],
            ["Contacto", "/contacto/", "Resolver dudas y concentrar medios de contacto.", "Sucursal"],
            ["Administración", "/admin/", "Gestionar catálogos, reservas, usuarios y permisos.", "Los seis modelos + Auth"],
        ],
        anchos=[1.0, 1.15, 2.6, 1.8],
        titulo="Mapa del sitio y fuentes de datos",
        primera_negrita=True,
    )

    h2(doc, "4.2 Identidad visual")
    par(doc, (
        "La dirección de arte busca un equilibrio entre elegancia y calidez. El carbón aporta sobriedad; el "
        "dorado champán se usa con moderación para acentos y llamadas a la acción; la crema reduce la dureza del "
        "fondo y mejora la lectura. Las tipografías se descargaron al proyecto para que el sitio no dependa de "
        "Google Fonts en tiempo de ejecución."
    ))
    tabla(
        doc,
        ["Color", "Hexadecimal", "Uso principal"],
        [
            ["Carbón", "#171512", "Encabezado, fondos oscuros, texto principal y botones secundarios."],
            ["Carbón suave", "#2A241D", "Bloques de apoyo, tablas del panel y degradados."],
            ["Dorado champán", "#C6A15B", "Acentos, botones principales, líneas y estados activos."],
            ["Dorado claro", "#E3C98F", "Texto sobre fondo oscuro y realces."],
            ["Crema", "#F7F2E9", "Fondo general y tarjetas del panel."],
            ["Blanco cálido", "#FFFDF9", "Tarjetas y superficies de contenido."],
        ],
        anchos=[1.35, 1.15, 4.0],
        titulo="Paleta de color de Lumina",
        primera_negrita=True,
        centrar=(1,),
    )
    par(doc, (
        "**Tipografía.** Playfair Display se utiliza en títulos y en la marca; Inter se reserva para párrafos, "
        "formularios y datos. Ambas son variables y se sirven localmente desde `web/static/fonts/`. El tamaño de "
        "los títulos usa `clamp()` para adaptarse al ancho de pantalla sin saltos bruscos."
    ))
    par(doc, (
        "**Componentes.** Las tarjetas de servicio, estilista, producto y sucursal se implementaron como "
        "parciales reutilizables de Django. Esa decisión garantiza que la página de inicio y las páginas de "
        "catálogo muestren exactamente el mismo componente, con las mismas proporciones de imagen y la misma "
        "jerarquía tipográfica."
    ))

    h2(doc, "4.3 Modelo de datos")
    par(doc, (
        "El modelo consta de seis entidades. Las cinco primeras cubren los catálogos solicitados; Reserva se "
        "agregó para conectar el sitio público con la operación. Todas heredan de `RegistroBase`, que aporta "
        "`creado_en` y `actualizado_en` para auditoría. Las relaciones usan `PROTECT` cuando borrar el padre "
        "dejaría registros huérfanos y `SET_NULL` cuando la información puede sobrevivir sin el padre."
    ))
    tabla(
        doc,
        ["Modelo", "Propósito", "Campos clave", "Relaciones"],
        [
            ["Sucursal", "Ubicaciones físicas del negocio.", "nombre, dirección, ciudad, teléfono, correo, horario, imagen, mapa_url, activa", "1 → N estilistas, clientes y reservas."],
            ["Servicio", "Catálogo de servicios con precio y duración.", "nombre, categoría, descripción, precio, duración_minutos, imagen, destacado, activo", "1 → N reservas (PROTECT)."],
            ["Estilista", "Equipo que atiende los servicios.", "nombre, apellidos, especialidad, biografía, foto, correo, teléfono, años, activo", "N → 1 sucursal (PROTECT); 1 → N reservas."],
            ["Producto", "Inventario de productos para venta.", "nombre, marca, categoría, descripción, precio, stock, imagen, activo", "Sin relaciones; catálogo independiente."],
            ["Cliente", "Datos de contacto y preferencias.", "nombre, apellidos, teléfono, correo, nacimiento, notas, activo", "N → 1 sucursal preferida (SET_NULL)."],
            ["Reserva", "Solicitud de cita del sitio público.", "nombre, teléfono, correo, fecha, hora, mensaje, estado", "N → 1 servicio (PROTECT), estilista (SET_NULL), sucursal (PROTECT)."],
        ],
        anchos=[0.9, 1.65, 2.7, 1.3],
        titulo="Entidades del modelo de datos",
        primera_negrita=True,
    )
    par(doc, (
        "La relación más importante es la de Reserva: conecta un servicio, una sucursal y, opcionalmente, un "
        "estilista. La validación del formulario impide elegir un estilista que pertenezca a otra sucursal, lo "
        "que evita un error frecuente en agendas multisede. La fecha mínima es el día siguiente, de modo que la "
        "solicitud no compita con citas del mismo día que requieren confirmación inmediata."
    ))

    h2(doc, "4.4 Flujo de reserva en línea")
    par(doc, (
        "El flujo público se diseñó para que un visitante complete la solicitud en menos de dos minutos. El "
        "formulario agrupa los campos en orden lógico y muestra los errores junto a cada campo, sin recargar la "
        "página con mensajes genéricos."
    ))
    pasos(doc, [
        "El visitante entra a **Reservar** desde el menú, el botón del hero o el botón flotante.",
        "Captura nombre, teléfono y correo; el teléfono y el nombre son obligatorios.",
        "Selecciona sucursal, servicio y, si lo desea, un estilista. La lista de estilistas se valida contra la sucursal elegida.",
        "Indica fecha y hora. El campo de fecha bloquea días pasados y el servidor vuelve a validar la regla.",
        "Envía la solicitud. Si es válida, se guarda con estado **Pendiente** y aparece un mensaje de confirmación.",
        "El personal revisa la reserva en el panel, la confirma por teléfono y actualiza su estado a **Confirmada**.",
    ])
    par(doc, (
        "El formulario aplica tres reglas de negocio: la fecha debe ser posterior al día actual; el estilista, si "
        "se elige, debe pertenecer a la sucursal seleccionada; y los campos obligatorios se validan tanto en el "
        "navegador como en el servidor. La validación del servidor es la que garantiza la integridad, porque la "
        "del navegador puede desactivarse."
    ))

    h2(doc, "4.5 Evidencia del sitio público")
    par(doc, (
        "Las siguientes capturas se tomaron con Playwright sobre el servidor local, a 1440 × 900 px y con las "
        "imágenes cargadas. Muestran el resultado real del HTML, CSS y los datos de demostración, no un mockup."
    ))
    figura(doc, "publico_01_inicio.png", "Página de inicio con hero, servicios destacados, equipo, productos y sucursales", ancho=6.3)
    par(doc, (
        "La página de inicio funciona como portada comercial: un mensaje claro, dos llamadas a la acción y "
        "prueba social. Debajo del hero se muestran contadores calculados con consultas a la base de datos y "
        "cuatro bloques de contenido que enlazan con los catálogos."
    ))
    figura(doc, "publico_02_servicios.png", "Catálogo de servicios con filtros por categoría, precio y duración", ancho=6.3)
    par(doc, (
        "El catálogo de servicios permite filtrar sin recargar la página completa mediante parámetros GET. Cada "
        "tarjeta incluye categoría, descripción, precio en pesos y duración. El filtro se refleja en la URL, por "
        "lo que el enlace puede compartirse."
    ))
    figura(doc, "publico_03_estilistas.png", "Directorio del equipo con especialidad, experiencia y sucursal", ancho=6.3)
    figura(doc, "publico_04_productos.png", "Catálogo de productos con búsqueda por nombre o marca y filtro por categoría", ancho=6.3)
    figura(doc, "publico_05_sucursales.png", "Directorio de sucursales con dirección, teléfono, horario y enlace de mapa", ancho=6.3)
    figura(doc, "publico_06_reservar.png", "Formulario público de reserva con validaciones y resumen del proceso", ancho=6.3)
    par(doc, (
        "El formulario de reserva es la pieza que conecta el sitio con la operación. La columna izquierda explica "
        "el proceso y resuelve dudas; la derecha contiene los campos. En la captura se observa el botón flotante "
        "oculto, decisión intencional para no tapar el formulario."
    ))
    figura(doc, "publico_07_contacto.png", "Página de contacto con medios directos, ubicaciones y preguntas frecuentes", ancho=6.3)

    h2(doc, "4.6 Diseño responsivo")
    par(doc, (
        "El sitio se probó en escritorio y en un viewport móvil de 390 × 844 px. En pantallas menores a 900 px el "
        "menú se convierte en un panel desplegable, las rejillas pasan de tres o cuatro columnas a dos y luego a "
        "una, y el botón flotante reduce su texto a un icono. Las imágenes usan `aspect-ratio` y `object-fit` para "
        "evitar saltos de diseño."
    ))
    figura(doc, "publico_08_inicio_movil.png", "Vista móvil de la página de inicio (390 × 844 px)", ancho=2.6, max_alto=7.0)
    par(doc, (
        "La versión móvil conserva la jerarquía: marca, menú, hero, llamada a la acción y contenido. El menú "
        "cierra al seleccionar una opción y el botón flotante permanece accesible sin cubrir el contenido "
        "principal. Las tipografías y los contrastes se mantienen legibles."
    ))


# ---------------------------------------------------------------------------
# 5. Autenticación y RBAC
# ---------------------------------------------------------------------------

def seccion_5(doc):
    h1(doc, "5. Autenticación y control de acceso (RBAC)")

    h2(doc, "5.1 Autenticación implementada")
    par(doc, (
        "La autenticación se apoya en el sistema nativo de Django. Cada usuario se registra en la tabla `auth_user` "
        "con una contraseña transformada mediante un hash salteado; la contraseña original nunca se almacena ni se "
        "recupera. Al iniciar sesión, Django valida las credenciales, crea una sesión y adjunta el usuario a cada "
        "petición. El panel comprueba dos condiciones: que la cuenta esté activa y que pertenezca al personal "
        "(`is_staff=True`). Además, cada vista verifica el permiso específico del modelo antes de ejecutar una acción."
    ))
    pasos(doc, [
        "El usuario abre `http://127.0.0.1:8000/admin/`; si no tiene sesión, Django lo redirige al formulario de acceso.",
        "Captura usuario y contraseña. El servidor compara el hash y, si coincide, crea la sesión.",
        "El middleware de autenticación carga el usuario y sus grupos en cada petición.",
        "El panel consulta `has_perm()` antes de mostrar un módulo o permitir una operación.",
        "Si el usuario no tiene permiso, responde con HTTP 403; si no tiene sesión, redirige al login con HTTP 302.",
        "Al cerrar sesión, el identificador de sesión se invalida y el navegador vuelve al formulario de acceso.",
    ])
    codigo(doc, (
        "ACCIONES_ADMIN = [\"add\", \"change\", \"delete\", \"view\"]\n"
        "ACCIONES_SUPERVISOR = [\"add\"]\n"
        "ACCIONES_EMPLEADO = [\"view\"]\n\n"
        "usuario.groups.set([grupos[datos[\"grupo\"]]])\n"
        "usuario.set_password(datos[\"password\"])"
    ), titulo="Extracto de web/management/commands/setup_rbac.py")
    par(doc, (
        "Los tres usuarios son cuentas de personal, pero solo el Administrador es superusuario. Esa diferencia es "
        "deliberada: un superusuario omite la comprobación de permisos, mientras que Supervisor y Empleado dependen "
        "exclusivamente de los permisos otorgados a su grupo. Así se puede demostrar que el RBAC funciona y no que "
        "el acceso se debe a un privilegio global."
    ))
    caja(doc, "Contraseña de la práctica", (
        "Los tres usuarios usan la contraseña `123` porque así lo solicita la actividad. Esa clave no cumple una "
        "política de seguridad real. En producción debe sustituirse por una contraseña larga y única, activar los "
        "validadores de Django y, preferentemente, un segundo factor de autenticación."
    ))

    h2(doc, "5.2 Roles del sistema")
    tabla(
        doc,
        ["Rol", "Usuario", "¿Superusuario?", "Permisos efectivos", "Propósito"],
        [
            ["Administrador", "admin", "Sí", "24 permisos (CRUD completo en los seis modelos) + gestión de usuarios y grupos", "Control total del sistema y de la configuración."],
            ["Supervisor", "supervisor", "No", "6 permisos de alta (`add`) en los seis modelos", "Capturar registros sin consultar, editar ni borrar los existentes."],
            ["Empleado", "empleado", "No", "6 permisos de consulta (`view`) en los seis modelos", "Consultar la operación sin modificar la información."],
        ],
        anchos=[1.05, 0.9, 0.95, 2.15, 1.65],
        titulo="Roles, usuarios y alcance de permisos",
        primera_negrita=True,
        centrar=(2,),
    )
    par(doc, (
        "Los permisos se asignan por grupo, no usuario por usuario. Esa práctica simplifica el mantenimiento: si "
        "mañana se agrega una séptima entidad, basta con incluirla en la lista de modelos y volver a ejecutar "
        "`setup_rbac`; los tres usuarios heredan la configuración correcta. El comando es idempotente, por lo que "
        "puede ejecutarse varias veces sin duplicar grupos ni permisos."
    ))

    h2(doc, "5.3 Matriz de permisos")
    par(doc, (
        "La matriz siguiente desglosa las 24 combinaciones posibles entre los seis modelos y las cuatro acciones "
        "básicas. “Sí” indica que el rol puede ejecutar la acción; “—” indica que el panel responde con HTTP 403. "
        "La lectura incluye tanto el listado como la ficha de detalle."
    ))
    filas_matriz = []
    for modelo in ("Servicios", "Estilistas", "Productos", "Sucursales", "Clientes", "Reservas"):
        for accion in ("Crear", "Consultar", "Editar", "Eliminar"):
            filas_matriz.append([
                modelo,
                accion,
                "Sí",
                "Sí" if accion == "Crear" else "—",
                "Sí" if accion == "Consultar" else "—",
            ])
    tabla(
        doc,
        ["Módulo", "Acción", "Administrador", "Supervisor", "Empleado"],
        filas_matriz,
        anchos=[1.45, 1.3, 1.25, 1.15, 1.1],
        titulo="Matriz RBAC: módulos, acciones y roles",
        primera_negrita=True,
        centrar=(1, 2, 3, 4),
    )
    par(doc, (
        "El Administrador también hereda el permiso de gestionar usuarios, grupos y permisos porque es "
        "superusuario. El Supervisor y el Empleado no pueden entrar a esas pantallas. La matriz se limita a los "
        "seis modelos del negocio, que son los que evalúa el comando de verificación."
    ))

    h2(doc, "5.4 Implementación técnica")
    par(doc, (
        "El comando `setup_rbac` crea los grupos y asigna los permisos consultando los `codename` nativos de "
        "Django (`add_servicio`, `view_servicio`, etc.). También crea o actualiza los tres usuarios y fuerza la "
        "contraseña indicada. El comando `verificar_rbac` inicia sesión con cada cuenta y compara el código HTTP "
        "obtenido contra el esperado; si algo falla, termina con código de salida distinto de cero."
    ))
    par(doc, (
        "El caso más delicado es el Supervisor. Django, después de guardar un alta, redirige al listado del modelo. "
        "Como el Supervisor no tiene permiso de consulta, ese listado devolvería 403 y la experiencia sería "
        "frustrante. Para resolverlo se creó `RedireccionSupervisorMixin`, que detecta al usuario del grupo "
        "Supervisor y lo devuelve al formulario de alta con un mensaje de éxito."
    ))
    codigo(doc, (
        "class RedireccionSupervisorMixin:\n"
        "    def response_add(self, request, obj, post_url_continue=None):\n"
        "        es_supervisor = request.user.groups.filter(name=\"Supervisor\").exists()\n"
        "        if es_supervisor and not request.user.is_superuser:\n"
        "            self.message_user(request, \"Registro guardado correctamente.\", messages.SUCCESS)\n"
        "            return HttpResponseRedirect(request.path)\n"
        "        return super().response_add(request, obj, post_url_continue)"
    ), titulo="Extracto de web/admin.py")
    par(doc, (
        "El panel también se personalizó con la identidad de Lumina: encabezado carbón, títulos dorados, tablas "
        "legibles y el nombre “Lumina · Administración”. La personalización no altera los permisos; solo mejora la "
        "coherencia visual y la orientación del usuario."
    ))

    h2(doc, "5.5 Verificación y pruebas")
    par(doc, (
        "La verificación se realizó en dos niveles. Primero, el comando `verificar_rbac` recorre rutas reales del "
        "panel con cada usuario y compara códigos HTTP. Segundo, la suite de pruebas automatizadas valida el sitio "
        "público, las reservas, los modelos y el comportamiento de los tres roles. La tabla resume los códigos "
        "obtenidos en la última ejecución."
    ))
    tabla(
        doc,
        ["Ruta verificada", "Administrador", "Supervisor", "Empleado", "Resultado"],
        [
            ["Panel principal `/admin/`", "200", "200", "200", "Conforme"],
            ["Listado de servicios", "200", "403", "200", "Conforme"],
            ["Alta de servicio", "200", "200", "403", "Conforme"],
            ["Edición de servicio", "200", "403", "200", "Conforme"],
            ["Eliminación de servicio", "200", "403", "403", "Conforme"],
            ["Listado de clientes", "200", "403", "200", "Conforme"],
            ["Alta de cliente", "200", "200", "403", "Conforme"],
            ["Panel sin sesión iniciada", "302", "302", "302", "Conforme"],
        ],
        anchos=[2.35, 0.95, 0.9, 0.85, 0.95],
        titulo="Resultados de la verificación RBAC sobre rutas reales",
        centrar=(1, 2, 3, 4),
        fuente="Salida de `python manage.py verificar_rbac`.",
    )
    tabla(
        doc,
        ["Grupo de pruebas", "Casos", "Qué valida"],
        [
            ["Sitio público", "5", "Las siete páginas responden, el contenido de la base se muestra y las reservas válidas se guardan mientras las inválidas se rechazan."],
            ["RBAC", "8", "Acceso total del administrador, alta exclusiva del supervisor, consulta exclusiva del empleado, rechazo de altas por POST y redirección de anónimos."],
            ["Modelos", "2", "Representación textual de sucursal, servicio y cliente."],
            ["Total", "15", "Suite ejecutada con `python manage.py test`."],
        ],
        anchos=[1.5, 0.7, 4.35],
        titulo="Cobertura de la suite de pruebas",
        centrar=(1,),
    )
    codigo(doc, (
        "Found 15 test(s).\n"
        "System check identified no issues (0 silenced).\n"
        ".................\n"
        "----------------------------------------------------------------------\n"
        "Ran 15 tests in 0.9s\n\n"
        "OK"
    ), titulo="Resultado de python manage.py test")
    par(doc, (
        "Las pruebas se ejecutan sobre una base temporal en memoria, por lo que no alteran los datos de "
        "demostración. El comando `seed_data` es idempotente y el comando `setup_rbac` puede repetirse sin efectos "
        "secundarios. Esa combinación permite reconstruir el entorno desde cero en cualquier equipo."
    ))

    h2(doc, "5.6 Evidencia de los roles")
    par(doc, (
        "Las capturas siguientes muestran la configuración real de usuarios, grupos y permisos, así como la "
        "experiencia de cada rol en el panel. Se generaron con Playwright después de aplicar el tema visual."
    ))
    figura(doc, "admin_02_usuarios.png", "Listado de usuarios: admin es superusuario; supervisor y empleado no lo son", ancho=6.1)
    figura(doc, "admin_03_grupos.png", "Grupos Administrador, Supervisor y Empleado creados por setup_rbac", ancho=6.1)
    figura(doc, "admin_04_permisos_supervisor.png", "Grupo Supervisor con los seis permisos de alta seleccionados", ancho=6.1)
    figura(doc, "admin_05_usuario_supervisor.png", "Ficha del usuario supervisor: cuenta de personal, sin privilegios de superusuario", ancho=6.0, max_alto=6.9)
    figura(doc, "rbac_01_supervisor_panel.png", "Panel del Supervisor: solo aparecen las acciones de alta", ancho=6.1)
    figura(doc, "rbac_02_supervisor_alta.png", "Formulario de alta disponible para el Supervisor", ancho=6.0, max_alto=6.9)
    figura(doc, "rbac_03_empleado_panel.png", "Panel del Empleado: acceso de consulta a los módulos", ancho=6.1)
    figura(doc, "rbac_04_empleado_lista.png", "Listado de servicios visible para el Empleado, sin acciones de edición o borrado", ancho=6.1)
    figura(doc, "rbac_05_empleado_detalle.png", "Ficha de servicio en modo consulta para el Empleado", ancho=6.0, max_alto=6.9)
    par(doc, (
        "La comparación visual confirma la matriz: el Supervisor ve “Agregar” pero no listados; el Empleado ve "
        "listados y fichas, pero no botones de alta ni de eliminación. El Administrador conserva todas las "
        "acciones. La evidencia no depende de interpretaciones: los códigos HTTP de la tabla anterior respaldan "
        "cada captura."
    ))

    h2(doc, "5.7 Consideraciones de seguridad")
    vinetas(doc, [
        "**Contraseñas:** se almacenan con el hash predeterminado de Django; la clave `123` es solo académica y debe cambiarse en cualquier entorno real.",
        "**Sesiones:** la cookie de sesión es HttpOnly por configuración de Django; la duración se fijó en ocho horas para evitar cierres durante la operación.",
        "**CSRF:** todos los formularios incluyen `{% csrf_token %}` y el middleware rechaza peticiones sin token válido.",
        "**Permisos:** se verifican en el servidor, no solo se ocultan botones. Ocultar un botón mejora la experiencia, pero no protege los datos.",
        "**Producción:** antes de publicar se debe establecer `DEBUG=False`, mover `SECRET_KEY` a una variable de entorno, configurar `ALLOWED_HOSTS`, usar HTTPS y respaldos.",
        "**Privacidad:** la base contiene teléfonos y correos; se debe informar al cliente el uso de sus datos y permitir su eliminación cuando lo solicite.",
    ])


# ---------------------------------------------------------------------------
# 6. Manual de usuario
# ---------------------------------------------------------------------------

def _manual_modelo(doc, nombre, singular, proposito, campos, lista_id, alta_id, editar_id,
                   eliminar_id, notas_eliminar, buenas_practicas, con_imagen=True,
                   publicacion=None):
    femenino = singular in {"sucursal", "reserva"}
    articulo = "la" if femenino else "el"
    contraccion = "de la" if femenino else "del"
    indefinido = "una" if femenino else "un"
    etiqueta = nombre.split(" ", 1)[1] if nombre[:1].isdigit() else nombre

    h2(doc, nombre)
    par(doc, proposito)
    par(doc, f"**Campos principales:** {campos}")

    par(doc, "**Consultar los registros.**", despues=2)
    pasos(doc, [
        f"Inicie sesión con una cuenta que tenga permiso de consulta (Administrador o Empleado).",
        f"En el menú **Gestión Lumina**, seleccione **{etiqueta}**.",
        "Revise la tabla: puede ordenar por una columna, usar la búsqueda o aplicar los filtros laterales.",
        f"Para ver el detalle, haga clic en el nombre {contraccion} {singular}.",
    ])
    figura(doc, lista_id, f"Listado de {etiqueta.lower()} en el panel administrativo", ancho=6.0, max_alto=6.9)

    par(doc, f"**Crear {indefinido} {singular}.**", despues=2)
    pasos_alta = [
        f"Pulse **Agregar {singular}** en la esquina superior derecha.",
        "Complete los campos obligatorios; los que tienen ayuda contextual explican su formato.",
    ]
    if con_imagen:
        pasos_alta.append(
            "Cargue una imagen si desea mostrarla en el sitio público. El panel muestra una vista previa."
        )
    if publicacion:
        pasos_alta.append(publicacion)
    pasos_alta.append(
        "Pulse **Guardar**. El sistema valida y registra el alta; si falta un dato, lo indica junto al campo."
    )
    pasos(doc, pasos_alta)
    figura(doc, alta_id, f"Formulario de alta de {singular}", ancho=6.0, max_alto=6.9)

    par(doc, f"**Editar {indefinido} {singular}.**", despues=2)
    pasos(doc, [
        f"En el listado, haga clic sobre el nombre {contraccion} {singular} o use el enlace **Modificar**.",
        "Cambie los campos necesarios; los cambios se guardan al confirmar.",
        "Pulse **Guardar**. El campo de auditoría conserva la fecha de la última actualización.",
    ])
    figura(doc, editar_id, f"Edición de {indefinido} {singular} existente", ancho=6.0, max_alto=6.9)

    par(doc, f"**Eliminar {indefinido} {singular}.**", despues=2)
    pasos(doc, [
        f"Con una cuenta con permiso de eliminación (Administrador), abra {articulo} {singular} y pulse **Eliminar** en la esquina inferior derecha.",
        "Confirme la operación en la pantalla de advertencia. Esta acción no se puede deshacer.",
    ])
    caja(doc, "Antes de eliminar", notas_eliminar)
    figura(doc, eliminar_id, f"Confirmación de eliminación de {singular}", ancho=6.0, max_alto=6.9)
    par(doc, f"**Buenas prácticas.** {buenas_practicas}")


def seccion_6(doc):
    h1(doc, "6. Manual de usuario", salto=False)

    par(doc, (
        "Este manual está dirigido al personal de Lumina. Explica cómo entrar al panel y cómo realizar las cuatro "
        "operaciones básicas —consultar, crear, editar y eliminar— en cada uno de los seis módulos. Las capturas "
        "corresponden a la versión entregada y muestran datos de demostración. Los procedimientos aplican igual "
        "para escritorio y para tableta; en móvil el panel de Django es funcional, aunque se recomienda operar en "
        "pantalla grande."
    ))
    caja(doc, "Regla de oro", (
        "Si un botón no aparece, no es un error: su rol no tiene ese permiso. El Administrador puede todo; el "
        "Supervisor solo agrega; el Empleado solo consulta. Nunca comparta credenciales ni use la cuenta de otra "
        "persona para “salir del paso”."
    ))

    h2(doc, "6.1 Acceso al panel administrativo")
    par(doc, (
        "El panel se encuentra en `http://127.0.0.1:8000/admin/` cuando el servidor se ejecuta en local. En un "
        "despliegue real se usaría el dominio del negocio con HTTPS. La tabla siguiente concentra las credenciales "
        "de la práctica."
    ))
    tabla(
        doc,
        ["Usuario", "Contraseña", "Rol", "Puede hacer"],
        [
            ["admin", "123", "Administrador", "Todo: consultar, crear, editar, eliminar y administrar usuarios."],
            ["supervisor", "123", "Supervisor", "Solo crear registros en los seis módulos."],
            ["empleado", "123", "Empleado", "Solo consultar listados y fichas."],
        ],
        anchos=[1.1, 0.95, 1.3, 3.2],
        titulo="Credenciales de acceso al panel",
        centrar=(0, 1),
        fuente="Credenciales académicas definidas por la actividad.",
    )
    pasos(doc, [
        "Abra el navegador e ingrese a `http://127.0.0.1:8000/admin/`.",
        "Escriba el nombre de usuario y la contraseña. Respete mayúsculas y minúsculas.",
        "Pulse **Iniciar sesión**. Si los datos son correctos, verá el panel de gestión.",
        "Use el menú **Gestión Lumina** para ir a los módulos y la esquina superior derecha para cerrar sesión.",
    ])
    figura(doc, "admin_00_login.png", "Formulario de acceso al panel administrativo", ancho=6.1)
    figura(doc, "admin_01_panel.png", "Panel principal con los módulos de gestión y las acciones disponibles", ancho=6.1)
    par(doc, "**Anatomía del panel.**", despues=2)
    vinetas(doc, [
        "**Barra superior:** nombre del sitio, usuario conectado, enlace al sitio público, cambio de contraseña y cierre de sesión.",
        "**Migas de pan:** indican la ubicación actual y permiten volver a un nivel anterior.",
        "**Menú lateral:** agrupa Autenticación y Autorización (usuarios y grupos) y Gestión Lumina (los seis módulos).",
        "**Zona de contenido:** listado, formulario o confirmación según la acción. Los botones de acción se ubican en la parte inferior.",
        "**Filtros y búsqueda:** aparecen a la derecha o en la parte superior del listado, según el modelo.",
    ])

    _manual_modelo(
        doc,
        "6.2 Servicios",
        "servicio",
        (
            "El módulo de servicios alimenta la página pública de Servicios y los bloques destacados de Inicio. "
            "Cada registro incluye precio, duración y una imagen que se muestra en la tarjeta del catálogo."
        ),
        "nombre, categoría, descripción, precio en MXN, duración en minutos, imagen, destacado y activo.",
        "admin_06_servicios_lista.png",
        "admin_07_servicios_alta.png",
        "admin_08_servicio_editar.png",
        "admin_09_servicio_eliminar.png",
        (
            "Si el servicio ya tiene reservas asociadas, Django impedirá borrarlo para no romper el historial. "
            "En ese caso, desmarque **activo** para ocultarlo del sitio sin perder los datos."
        ),
        (
            "Use imágenes horizontales de 1200 × 900 px; escriba precios sin símbolo ni comas; destaque solo "
            "cuatro o seis servicios para no saturar el inicio; en lugar de borrar, desactive los servicios "
            "temporales."
        ),
        publicacion=(
            "Revise las opciones de publicación: **destacado** para mostrarlo en Inicio y **activo** para "
            "mantenerlo visible en el catálogo."
        ),
    )

    _manual_modelo(
        doc,
        "6.3 Estilistas",
        "estilista",
        (
            "El módulo de estilistas alimenta el directorio del equipo y la selección de estilista en el formulario "
            "de reserva. Cada estilista debe pertenecer a una sucursal; esa relación permite validar que la cita "
            "sea coherente con la ubicación elegida."
        ),
        "nombre, apellidos, especialidad, biografía, sucursal, correo, teléfono, años de experiencia, fotografía y activo.",
        "admin_10_estilistas_lista.png",
        "admin_11_estilistas_alta.png",
        "admin_12_estilista_editar.png",
        "admin_13_estilista_eliminar.png",
        (
            "Un estilista con reservas asociadas no se puede eliminar. Primero reasigne o cancele las reservas "
            "pendientes; si solo desea ocultarlo del sitio, desmarque **activo**."
        ),
        (
            "Use retratos verticales de 900 × 1200 px; escriba una biografía breve de dos o tres líneas; revise "
            "que la sucursal sea la correcta antes de guardar; mantenga actualizada la especialidad."
        ),
        publicacion="Marque **activo** para publicar el perfil en el directorio del sitio.",
    )

    _manual_modelo(
        doc,
        "6.4 Productos",
        "producto",
        (
            "El módulo de productos controla el catálogo de venta y las existencias. El sitio público muestra "
            "“Disponible” cuando hay stock y “Agotado” cuando llega a cero. El listado permite activar o "
            "desactivar un producto sin abrir su ficha."
        ),
        "nombre, marca, categoría, descripción, precio, existencias, imagen y activo.",
        "admin_14_productos_lista.png",
        "admin_15_productos_alta.png",
        "admin_16_producto_editar.png",
        "admin_17_producto_eliminar.png",
        (
            "Si un producto ya no se vende, desactívelo en lugar de borrarlo: conserva la referencia para reportes "
            "y evita perder la relación con imágenes o documentos internos."
        ),
        (
            "Use imágenes cuadradas de 1000 × 1000 px; actualice el stock después de cada venta; revise los "
            "productos con existencias bajas; escriba la marca tal como aparece en el empaque para facilitar la "
            "búsqueda."
        ),
        publicacion=(
            "Actualice **existencias** y marque **destacado** o **activo** según corresponda; con cero "
            "existencias el sitio muestra “Agotado”."
        ),
    )

    _manual_modelo(
        doc,
        "6.5 Sucursales",
        "sucursal",
        (
            "El módulo de sucursales concentra direcciones, teléfonos, horarios, fotografías y enlaces de mapa. "
            "Es la base para asignar estilistas y para que el visitante elija dónde atenderse."
        ),
        "nombre, dirección, ciudad, teléfono, correo, horario, imagen, enlace de mapa y activa.",
        "admin_18_sucursales_lista.png",
        "admin_19_sucursales_alta.png",
        "admin_20_sucursal_editar.png",
        "admin_21_sucursal_eliminar.png",
        (
            "Una sucursal con estilistas o reservas asociadas no se puede eliminar. Desactívela para ocultarla "
            "del sitio y conserve el historial de operación."
        ),
        (
            "Use fotografías horizontales de 1200 × 750 px; verifique el enlace de mapa antes de publicarlo; "
            "escriba el horario en el mismo formato en todas las sucursales; mantenga un teléfono directo "
            "actualizado."
        ),
        publicacion="Marque **activa** para mostrarla en el sitio y verifique el enlace de mapa.",
    )

    _manual_modelo(
        doc,
        "6.6 Clientes",
        "cliente",
        (
            "El módulo de clientes guarda datos de contacto, fecha de nacimiento, sucursal preferida y notas de "
            "atención. Es información sensible: solo debe consultarse para fines del servicio y no debe "
            "compartirse fuera del negocio."
        ),
        "nombre, apellidos, teléfono, correo, fecha de nacimiento, sucursal preferida, notas y activo.",
        "admin_22_clientes_lista.png",
        "admin_23_clientes_alta.png",
        "admin_24_cliente_editar.png",
        "admin_25_cliente_eliminar.png",
        (
            "Antes de eliminar un cliente, confirme que no exista una obligación legal o comercial de conservar "
            "sus datos. Si solo dejó de asistir, desactívelo para conservar el historial."
        ),
        (
            "Capture el teléfono con lada; use las notas para preferencias y alergias, no para juicios personales; "
            "solicite consentimiento para comunicaciones promocionales; limite el acceso a quien realmente lo "
            "necesite."
        ),
        con_imagen=False,
        publicacion=(
            "Revise la **sucursal preferida** y marque **activo** para conservar el historial de atención."
        ),
    )

    _manual_modelo(
        doc,
        "6.7 Reservas",
        "reserva",
        (
            "El módulo de reservas es el puente entre el sitio público y la operación. Cada solicitud entra con "
            "estado **Pendiente**; el personal la revisa, confirma por teléfono y actualiza el estado. El listado "
            "se puede filtrar por estado, sucursal y fecha, y permite cambiar el estado en línea."
        ),
        "nombre del cliente, teléfono, correo, servicio, estilista, sucursal, fecha, hora, mensaje y estado.",
        "admin_26_reservas_lista.png",
        "admin_27_reservas_alta.png",
        "admin_28_reserva_editar.png",
        "admin_29_reserva_eliminar.png",
        (
            "No elimine una reserva cancelada si necesita conservar el historial. Cámbiela a **Cancelada** y "
            "explique el motivo en el mensaje. Elimine solo duplicados evidentes o pruebas internas."
        ),
        (
            "Confirme las reservas pendientes el mismo día; actualice el estado después de la llamada; use el "
            "filtro por fecha para preparar la agenda; no prometa un horario sin revisar la disponibilidad real "
            "del estilista."
        ),
        con_imagen=False,
        publicacion=(
            "Asigne servicio, estilista, sucursal, fecha y hora; el sistema valida que la combinación sea "
            "coherente."
        ),
    )

    h2(doc, "6.8 Reserva en línea para visitantes")
    par(doc, (
        "El visitante no necesita crear una cuenta. El formulario público solicita los datos mínimos para "
        "contactarlo y guarda la solicitud en la misma base que utiliza el panel. La reserva no se confirma "
        "automáticamente: primero debe validarla el personal."
    ))
    pasos(doc, [
        "Entrar a **Reservar** desde el menú, el botón dorado o el botón flotante.",
        "Escribir nombre completo, teléfono y, opcionalmente, correo electrónico.",
        "Elegir sucursal, servicio y, si se desea, un estilista de esa misma sucursal.",
        "Seleccionar una fecha posterior al día actual y una hora disponible.",
        "Agregar un mensaje con preferencias o alergias y pulsar **Enviar solicitud**.",
        "Leer el mensaje de confirmación; el equipo se comunicará por teléfono para validar la cita.",
    ])
    par(doc, (
        "Si el formulario muestra un error, revise el campo señalado. Los errores más comunes son una fecha "
        "pasada, un teléfono vacío o un estilista de otra sucursal. El sistema conserva los datos ya escritos "
        "para no obligar a capturarlos de nuevo."
    ))
    caja(doc, "Qué ocurre detrás del formulario", (
        "La solicitud se guarda con estado Pendiente y aparece de inmediato en **Gestión Lumina → Reservas**. "
        "Si el visitante elige “Cualquier estilista disponible”, el campo se guarda vacío y el personal asigna "
        "al especialista según la agenda."
    ))

    h2(doc, "6.9 Preguntas frecuentes y solución de problemas")
    tabla(
        doc,
        ["Situación", "Causa probable", "Solución"],
        [
            ["No puedo iniciar sesión.", "Usuario o contraseña incorrectos, o cuenta inactiva.", "Verifique mayúsculas; pida al Administrador que revise la cuenta en Usuarios."],
            ["Veo el panel, pero no aparecen módulos.", "El usuario no pertenece a un grupo con permisos.", "Ejecute `python manage.py setup_rbac` o asigne el grupo correcto."],
            ["El botón Agregar no aparece.", "Su rol es Empleado (solo consulta).", "Solicite apoyo a un Administrador o Supervisor; no comparta cuentas."],
            ["No puedo eliminar una sucursal.", "Tiene estilistas o reservas asociadas (PROTECT).", "Desactívela para ocultarla del sitio; conserve el historial."],
            ["Una imagen no se ve en el sitio.", "El archivo no se cargó o la ruta es incorrecta.", "Edite el registro, vuelva a cargar la imagen y guarde; revise `/media/`."],
            ["El índice del documento no coincide.", "Los campos de Word no se actualizaron.", "Abra el documento, seleccione el índice y pulse F9, o ejecute el script de actualización."],
            ["El servidor no inicia.", "Dependencias o migraciones pendientes.", "Active el entorno virtual, ejecute `pip install -r requirements.txt` y `python manage.py migrate`."],
            ["La reserva no aparece.", "El formulario no se envió o falló una validación.", "Revise el mensaje de error; confirme que la fecha sea futura."],
        ],
        anchos=[1.7, 2.05, 2.8],
        titulo="Solución de problemas frecuentes",
        primera_negrita=True,
    )

    h2(doc, "6.10 Respaldo y buenas prácticas")
    par(doc, (
        "La base de datos es el archivo `db.sqlite3` y las imágenes viven en la carpeta `media/`. Para respaldar "
        "el sistema completo basta con copiar ambos elementos cuando el servidor esté detenido. En producción se "
        "recomienda programar respaldos automáticos, almacenar las copias fuera del servidor y probar la "
        "restauración al menos una vez al mes."
    ))
    vinetas(doc, [
        "No edite `db.sqlite3` con herramientas externas mientras el servidor esté en ejecución.",
        "Cambie la contraseña `123` antes de publicar el sitio y no reutilice la cuenta de Administrador para tareas cotidianas.",
        "Desactive registros en lugar de borrarlos cuando necesite conservar historial.",
        "Revise los permisos después de agregar un nuevo modelo o usuario.",
        "Mantenga actualizadas las dependencias y realice una copia antes de cada actualización.",
        "Documente los cambios relevantes en el README del proyecto.",
    ])


# ---------------------------------------------------------------------------
# 7. Conclusiones
# ---------------------------------------------------------------------------

def seccion_7(doc):
    h1(doc, "7. Conclusiones")

    h2(doc, "7.1 Resultados alcanzados")
    par(doc, (
        "El proyecto cumplió el objetivo general: Lumina cuenta con un sitio público funcional, un panel "
        "administrativo con seis módulos y un modelo de permisos verificable. La tabla siguiente relaciona cada "
        "objetivo específico con la evidencia que lo respalda."
    ))
    tabla(
        doc,
        ["Objetivo específico", "Evidencia en el proyecto", "Estado"],
        [
            ["Sitio público de siete páginas, responsivo y con identidad propia.", "Capturas de las secciones 4.5 y 4.6; pruebas de las siete rutas.", "Cumplido"],
            ["Modelo de datos con las cinco entidades solicitadas y reservas.", "Migración `0001_initial`, modelos y panel administrativo.", "Cumplido"],
            ["CRUD completo en el panel para los cinco catálogos.", "Manual de la sección 6 con listado, alta, edición y eliminación.", "Cumplido"],
            ["Tres roles con permisos diferenciados.", "Matriz RBAC, comando `verificar_rbac` y ocho pruebas de acceso.", "Cumplido"],
            ["Captación de citas en línea conectada a la base.", "Formulario público, validaciones y pruebas de reserva.", "Cumplido"],
            ["Documentación de requisitos, negocio, RBAC y manual.", "Este documento y el README del repositorio.", "Cumplido"],
            ["Despliegue productivo con HTTPS y respaldos.", "Propuesto como trabajo futuro; no forma parte del alcance.", "No incluido"],
        ],
        anchos=[2.1, 3.65, 0.8],
        titulo="Trazabilidad de objetivos y evidencia",
        centrar=(2,),
    )

    h2(doc, "7.2 Aprendizajes del equipo")
    vinetas(doc, [
        "**Django reduce la complejidad, pero no decide por nosotros.** El framework aporta autenticación y administración; las reglas del negocio —validar una fecha, impedir un estilista de otra sucursal, redirigir al Supervisor— siguen siendo responsabilidad del equipo.",
        "**El RBAC se demuestra con pruebas, no con capturas.** Ocultar un botón no es seguridad. La verificación con códigos HTTP y la suite automatizada obligaron a revisar cada permiso en el servidor.",
        "**La experiencia del rol importa.** El primer diseño del Supervisor terminaba en una pantalla 403 después de guardar. El mixin de redirección convirtió ese error técnico en un flujo comprensible.",
        "**La identidad visual también es funcional.** Una paleta consistente, jerarquía tipográfica y tarjetas reutilizables reducen la carga cognitiva y hacen que el contenido se lea mejor.",
        "**Los datos de demostración son parte del producto.** Sin registros coherentes, el sitio se ve vacío y las capturas no comunican nada. El comando `seed_data` resolvió esa reproducibilidad.",
        "**Documentar es diseñar.** Escribir el manual reveló campos confusos y acciones que necesitaban mejor texto de ayuda; la documentación se convirtió en una prueba de usabilidad.",
    ])

    h2(doc, "7.3 Trabajo futuro")
    par(doc, (
        "La siguiente versión debería enfocarse en operación real, no en agregar páginas. Los tres frentes "
        "prioritarios son: notificaciones automáticas y confirmación de asistencia; reportes de ingresos, "
        "servicios más solicitados y clientes recurrentes; y despliegue con PostgreSQL, HTTPS, respaldos y "
        "monitoreo. También conviene sustituir las fotografías de stock por fotografía propia y validar el aviso "
        "de privacidad con asesoría legal."
    ))
    par(doc, (
        "Desde la perspectiva académica, el proyecto deja una base sólida para seguir aprendiendo: la arquitectura "
        "MVT, el ORM, las migraciones, los formularios, la administración y el control de acceso son conceptos "
        "que se reutilizan en cualquier aplicación web profesional. La diferencia entre un ejercicio y un producto "
        "es la consistencia: datos, permisos, diseño y documentación deben contar la misma historia."
    ))


# ---------------------------------------------------------------------------
# 8. Referencias
# ---------------------------------------------------------------------------

def _referencia(doc, texto: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = u.Cm(1.0)
    p.paragraph_format.first_line_indent = u.Cm(-1.0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    u._agregar_runs(p, texto, 10)


def seccion_8(doc):
    h1(doc, "8. Referencias")
    par(doc, (
        "Las referencias se presentan en estilo APA 7. Se incluyen fuentes técnicas, estándares de accesibilidad "
        "y el material del curso. Las fotografías se acreditan en el anexo C."
    ))
    referencias = [
        "Django Software Foundation. (2026). *Django documentation* (Versión 5.2). https://docs.djangoproject.com/en/5.2/",
        "Google. (2026). *Playwright documentation*. https://playwright.dev/docs/intro",
        "INEGI. (2025). *Directorio Estadístico Nacional de Unidades Económicas*. https://www.inegi.org.mx/",
        "MDN Web Docs. (2026). *HTML: HyperText Markup Language*. https://developer.mozilla.org/es/docs/Web/HTML",
        "MDN Web Docs. (2026). *CSS: Cascading Style Sheets*. https://developer.mozilla.org/es/docs/Web/CSS",
        "OWASP Foundation. (2025). *Password storage cheat sheet*. https://cheatsheetseries.owasp.org/cheatsheets/Password_Storage_Cheat_Sheet.html",
        "Python Software Foundation. (2026). *Python 3 documentation*. https://docs.python.org/3/",
        "Universidad Autónoma de Nuevo León. (2026). *Fundamentos de Desarrollo Web: Producto Integrador de Aprendizaje* [Material del curso]. FACPYA.",
        "Unsplash. (2026). *Unsplash license*. https://unsplash.com/license",
        "W3C. (2023). *Web Content Accessibility Guidelines (WCAG) 2.2*. https://www.w3.org/TR/WCAG22/",
    ]
    for referencia in referencias:
        _referencia(doc, referencia)


# ---------------------------------------------------------------------------
# Anexos
# ---------------------------------------------------------------------------

def anexos(doc):
    h1(doc, "Anexos")

    h2(doc, "Anexo A. Credenciales y comandos de ejecución")
    par(doc, (
        "Los siguientes comandos permiten reconstruir el entorno desde cero en Windows. En macOS o Linux cambia "
        "la ruta del entorno virtual por `source .venv/bin/activate`."
    ))
    codigo(doc, (
        "python -m venv .venv\n"
        ".venv\\Scripts\\activate\n"
        "pip install -r requirements.txt\n"
        "python manage.py migrate\n"
        "python manage.py seed_data\n"
        "python manage.py setup_rbac\n"
        "python manage.py runserver\n\n"
        "# Verificación\n"
        "python manage.py test\n"
        "python manage.py verificar_rbac"
    ), titulo="Instalación y arranque")
    tabla(
        doc,
        ["Recurso", "URL o valor"],
        [
            ["Sitio público", "http://127.0.0.1:8000/"],
            ["Panel administrativo", "http://127.0.0.1:8000/admin/"],
            ["Usuario administrador", "admin / 123"],
            ["Usuario supervisor", "supervisor / 123"],
            ["Usuario empleado", "empleado / 123"],
            ["Base de datos", "db.sqlite3 (SQLite)"],
            ["Archivos multimedia", "media/ (imágenes de servicios, estilistas, productos y sucursales)"],
        ],
        anchos=[2.0, 4.55],
        titulo="Recursos de la entrega",
        primera_negrita=True,
    )
    caja(doc, "Advertencia", (
        "Las credenciales son de práctica. Cambie las contraseñas, establezca `DEBUG=False` y use HTTPS antes de "
        "publicar el sitio en internet."
    ))

    h2(doc, "Anexo B. Estructura del proyecto")
    codigo(doc, (
        "PIA/\n"
        "├── manage.py\n"
        "├── db.sqlite3\n"
        "├── requirements.txt\n"
        "├── README.md\n"
        "├── salon/            (settings, urls, wsgi)\n"
        "├── web/\n"
        "│   ├── admin.py, forms.py, models.py, views.py, urls.py\n"
        "│   ├── migrations/\n"
        "│   ├── management/commands/  (seed_data, setup_rbac, verificar_rbac)\n"
        "│   ├── static/       (css, fonts, img, admin)\n"
        "│   └── templates/    (admin, web)\n"
        "├── media/\n"
        "├── docs/             (PIA_Lumina.docx, assets)\n"
        "├── capturas/\n"
        "└── scripts/          (capturas.py, generar_documento.py, docx_util.py)"
    ), titulo="Árbol de directorios principal")

    h2(doc, "Anexo C. Créditos de imágenes")
    par(doc, (
        "Las fotografías del sitio provienen de Unsplash y se recortaron localmente para las proporciones de cada "
        "componente. La licencia de Unsplash permite el uso gratuito, incluido el comercial, sin obligación de "
        "atribución; aun así, se documentan por transparencia académica. Los logotipos institucionales son "
        "propiedad de la UANL y de la FACPYA y se utilizan únicamente en la portada de este trabajo."
    ))
    tabla(
        doc,
        ["Uso en el sitio", "Archivos", "Fotografías de Unsplash"],
        [
            ["Portada, ambiente y llamada a la acción", "hero.jpg, about.jpg, cta.jpg", "photo-1633681926022-84c23e8cb2d6 · photo-1562322140-8baeececf3df · photo-1621605815971-fbc98d665033"],
            ["Servicios de cabello", "servicio-corte.jpg, servicio-color.jpg, servicio-peinado.jpg, servicio-keratina.jpg", "photo-1622286342621-4bd786c2447c · photo-1595476108010-b4d1f102b1b1 · photo-1522337360788-8b13dee7a37e · photo-1519699047748-de8e457a634e"],
            ["Barbería", "servicio-barba.jpg, servicio-afeitado.jpg", "photo-1503951914875-452162b0f3f1 · photo-1585747860715-2ba37e788b70"],
            ["Uñas, facial y maquillaje", "servicio-unas.jpg, servicio-facial.jpg, servicio-maquillaje.jpg", "photo-1610992015732-2449b76344bc · photo-1570172619644-dfd03ed5d881 · photo-1487412947147-5cebf100ffc2"],
            ["Retratos del equipo", "estilista-1..4.jpg", "photo-1544005313-94ddf0286df2 · photo-1507003211169-0a1dd7228f2d · photo-1494790108377-be9c29b29330 · photo-1506794778202-cad84cf45f1d"],
            ["Productos", "producto-1..5.jpg", "photo-1571781926291-c477ebfd024b · photo-1556228720-195a672e8a03 · photo-1608248543803-ba4f8c70ae0b · photo-1596462502278-27bfdc403348 · photo-1585751119414-ef2636f8aede"],
            ["Sucursales", "sucursal-1..3.jpg", "photo-1560066984-138dadb4c035 · photo-1521590832167-7bcbfaa6381f · photo-1600948836101-f9ffda59d250"],
            ["Logotipos institucionales", "logo_uanl.png, logo_facpya.png", "Propiedad de la Universidad Autónoma de Nuevo León y la FACPYA."],
        ],
        anchos=[1.55, 1.9, 3.1],
        titulo="Créditos y procedencia de las imágenes",
        fuente="Unsplash (https://unsplash.com) y archivos institucionales.",
    )
    par(doc, (
        "Las capturas de pantalla del sitio y del panel fueron generadas por el equipo con Playwright sobre la "
        "aplicación en ejecución. No son imágenes de banco ni montajes: reflejan el estado real de la entrega."
    ))


# ---------------------------------------------------------------------------
# Ensamblado
# ---------------------------------------------------------------------------

def main():
    preparar_capturas()
    doc = Document()
    configurar_estilos(doc)
    metadatos(
        doc,
        "Lumina · Producto Integrador de Aprendizaje",
        "Diego Alejandro Vera Reyes y equipo",
        "Sitio web y sistema de gestión para un salón de belleza y barbería",
        "Django, RBAC, sitio web, salón de belleza, PIA, FACPYA",
    )

    portada(doc, INTEGRANTES)
    iniciar_cuerpo(doc)
    indice(doc)
    doc.add_page_break()

    seccion_1(doc)
    seccion_2(doc)
    seccion_3(doc)
    seccion_4(doc)
    seccion_5(doc)
    seccion_6(doc)
    seccion_7(doc)
    seccion_8(doc)
    anexos(doc)

    salida = u.DOCS / "PIA_Lumina.docx"
    salida.parent.mkdir(parents=True, exist_ok=True)
    doc.save(salida)
    print(f"Documento generado: {salida}")
    print(f"Figuras insertadas: {u._figura_actual} | Tablas: {u._tabla_actual}")


if __name__ == "__main__":
    main()
