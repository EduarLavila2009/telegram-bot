"""Carga de configuración desde variables de entorno."""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

_base = Path(__file__).resolve().parent
_ENV_PATH = _base / ".env"
# Cargar .env junto a este paquete (funciona aunque ejecutes desde la raíz del repo).
load_dotenv(_ENV_PATH)


def _env_help() -> str:
    return (
        f"Archivo esperado: {_ENV_PATH}\n"
        "1) Abre ese archivo y guarda los cambios (Ctrl+S).\n"
        "2) TELEGRAM_BOT_TOKEN: copialo del mensaje de @BotFather al crear el bot "
        "(formato: numeros, dos puntos y letras; ejemplo 123456789:ABC...).\n"
        "3) TELEGRAM_ALLOWED_USER_ID: tu Id numerico (escribe a @userinfobot o @getidsbot).\n"
        "4) Vuelve a ejecutar: python -m bot_financiero_telegram desde la raiz del repo, "
        "con el venv activado o usando .\\bot_financiero_telegram\\.venv\\Scripts\\python."
    )


def _require_int(name: str) -> int:
    raw = os.environ.get(name)
    if raw is None or not str(raw).strip():
        raise RuntimeError(
            f"Falta la variable de entorno obligatoria: {name}\n{_env_help()}"
        )
    return int(str(raw).strip())


def _require_str(name: str) -> str:
    raw = os.environ.get(name)
    if raw is None or not str(raw).strip():
        raise RuntimeError(
            f"Falta la variable de entorno obligatoria: {name}\n{_env_help()}"
        )
    return str(raw).strip()


def _optional_int(name: str) -> int | None:
    raw = os.environ.get(name)
    if raw is None or not str(raw).strip():
        return None
    return int(str(raw).strip())


_DATA_DIR = Path("/data")
_db_root = _DATA_DIR if _DATA_DIR.is_dir() else _base


def _resolve_path_env(name: str, default_relative: str) -> Path:
    """
    Resuelve rutas relativas a la carpeta raíz de datos (`_db_root`).
    """
    raw = os.environ.get(name)
    if raw is None or not str(raw).strip():
        return (_db_root / default_relative).expanduser()
    p = Path(str(raw).strip()).expanduser()
    if not p.is_absolute():
        return (_db_root / p).expanduser()
    if _db_root == _DATA_DIR and not p.exists():
        # Redirigir archivo absoluto no existente al volumen persistente /data
        return (_DATA_DIR / p.name).expanduser()
    if p.exists():
        return p
    by_name = _db_root / p.name
    if by_name.exists():
        return by_name
    fallback = (_db_root / default_relative).expanduser()
    if fallback.exists():
        return fallback
    return p


def _validate_bot_token(token: str) -> str:
    token = token.strip()
    if ":" not in token:
        raise RuntimeError(
            "TELEGRAM_BOT_TOKEN no tiene el formato esperado "
            "(debe incluir ':' como en el token que envia @BotFather).\n"
            f"{_env_help()}"
        )
    return token


BOT_TOKEN: str = _validate_bot_token(_require_str("TELEGRAM_BOT_TOKEN"))
ALLOWED_USER_ID: int = _require_int("TELEGRAM_ALLOWED_USER_ID")

# Canal cuyas publicaciones (fotos) se procesan. Suele ser un id tipo -100xxxxxxxxxx.
SOURCE_CHANNEL_ID: int | None = _optional_int("TELEGRAM_SOURCE_CHANNEL_ID")

EXCEL_PATH = _resolve_path_env("EXCEL_PATH", "RETEN-REC.xlsx")

# Excel histórico para facturas de compra (si se usa).
FACTURAS_COMPRA_PATH = _resolve_path_env(
    "FACTURAS_COMPRA_XLSX",
    "facturas_compra_recibidas.xlsx",
)

# Nuevo Excel oficial para facturas recibidas (compra).
FACTURAS_RECIBIDAS_PATH = _resolve_path_env(
    "FACTURAS_RECIBIDAS_XLSX",
    "FACTURAS-RECIBIDAS-NUEVO.xlsx",
)

# Excel oficial para facturas de venta / emitidas.
FACTURAS_EMITIDAS_PATH = _resolve_path_env(
    "FACTURAS_EMITIDAS_XLSX",
    "FACTURAS-EMITIDAS.xlsx",
)

# Excel oficial para reportes Z de ventas.
REPORTES_Z_PATH = _resolve_path_env(
    "REPORTES_Z_XLSX",
    "REPORTES-Z-NUEVO.xlsx",
)

# Excel de inventario/productos para cotizaciones y notas de entrega.
PRODUCTOS_PATH = _resolve_path_env(
    "PRODUCTOS_XLSX",
    "inventario.xlsx",
)

# Carpeta base para libros mensuales de retenciones emitidas.
RETENCIONES_EMITIDAS_DIR = _resolve_path_env(
    "RETENCIONES_EMITIDAS_DIR",
    "RETENCIONES-EMITIDAS-NUEVO",
)

# Imagen PNG (preferible transparente) para firma+sello en comprobantes PDF emitidos.
FIRMA_SELLO_PATH = _resolve_path_env(
    "FIRMA_SELLO_PATH",
    "firma_sello_transparente.png",
)

OPENAI_API_KEY: str | None = os.environ.get("OPENAI_API_KEY") or None
GEMINI_API_KEY: str | None = os.environ.get("GEMINI_API_KEY") or None
GEMINI_MODEL: str = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash").strip()

SMTP_SERVER: str | None = os.environ.get("SMTP_SERVER") or None
SMTP_PORT: int = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER: str | None = os.environ.get("SMTP_USER") or None
SMTP_PASSWORD: str | None = os.environ.get("SMTP_PASSWORD") or None
DEFAULT_ACCOUNTANT_EMAIL: str | None = os.environ.get("DEFAULT_ACCOUNTANT_EMAIL") or None

# Variables para despliegue en la nube (Render)
PORT: int = int(os.environ.get("PORT", "8000"))
RENDER_EXTERNAL_URL: str | None = os.environ.get("RENDER_EXTERNAL_URL") or None

# RIF del Agente de Retención (nuestra empresa)
EMITTER_RIF: str = os.environ.get("EMITTER_RIF", "J-40194130-3").strip()

# Carpeta para libros mensuales de retenciones de ISLR
RETENCIONES_ISLR_DIR = _resolve_path_env(
    "RETENCIONES_ISLR_DIR",
    "RETENCIONES-ISLR-EMITIDAS",
)

# Evitar ejecución local accidental (polling) que rompa el webhook de producción
FORCE_LOCAL_POLLING: bool = os.environ.get("FORCE_LOCAL_POLLING", "false").lower() == "true"

