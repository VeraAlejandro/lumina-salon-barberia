"""
Configuración del proyecto Lumina — Salón & Barbería.

La configuración funciona sin pasos adicionales en desarrollo (SQLite y DEBUG
activo), pero los valores sensibles se pueden sobrescribir con variables de
entorno para un despliegue real.
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# Clave de desarrollo para que el proyecto arranque sin configuración extra.
# En producción define DJANGO_SECRET_KEY con un valor largo y aleatorio.
SECRET_KEY = os.environ.get(
    "DJANGO_SECRET_KEY",
    "django-insecure-lumina-pia-solo-desarrollo",
)

# DEBUG activo por defecto para facilitar la evaluación local.
# En producción define DJANGO_DEBUG=False.
DEBUG = os.environ.get("DJANGO_DEBUG", "True").strip().lower() in {
    "1",
    "true",
    "yes",
    "on",
}

ALLOWED_HOSTS = [
    host.strip()
    for host in os.environ.get(
        "DJANGO_ALLOWED_HOSTS",
        "localhost,127.0.0.1,[::1]",
    ).split(",")
    if host.strip()
]

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "web",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "salon.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [BASE_DIR / "web" / "templates"],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

WSGI_APPLICATION = "salon.wsgi.application"

DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": BASE_DIR / "db.sqlite3",
    }
}

AUTH_PASSWORD_VALIDATORS = []  # La práctica exige la contraseña "123".

LANGUAGE_CODE = "es-mx"
TIME_ZONE = "America/Monterrey"
USE_I18N = True
USE_TZ = True

STATIC_URL = "static/"
STATICFILES_DIRS = [BASE_DIR / "web" / "static"]
STATIC_ROOT = BASE_DIR / "staticfiles"

MEDIA_URL = "media/"
MEDIA_ROOT = BASE_DIR / "media"

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# Sesión de administración: evita que el usuario pierda el trabajo por inactividad
# mientras se toman las capturas del manual.
SESSION_COOKIE_AGE = 60 * 60 * 8

MESSAGE_STORAGE = "django.contrib.messages.storage.session.SessionStorage"
