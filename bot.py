"""Aplicacion Telegram: texto/voz, Excel y restriccion por usuario."""

from __future__ import annotations

import asyncio
import logging
from logging.handlers import RotatingFileHandler
import re
import tempfile
import time
import unicodedata
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup, Update
from telegram.error import Conflict, NetworkError, TelegramError, TimedOut
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters

from . import config
from . import excel_store
from . import tributario_engine
from .factura_compra_parse import parse_factura_compra_text, _take_eol_label, _normalize_rif
from .transcription import transcribe_audio_file
from openpyxl import load_workbook

# Configuración inicial del logger raíz. Se sobreescribirá en _setup_logging() con el archivo de log.
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


def _setup_logging() -> None:
    """Configura el logging para escribir tanto en la consola como en un archivo rotativo."""
    log_dir = Path(__file__).resolve().parent
    log_file = log_dir / "bot.log"
    
    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    
    # Handler para consola
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.INFO)
    
    # Handler para archivo rotativo (5MB máximo por archivo, manteniendo hasta 5 respaldos)
    try:
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=5 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8"
        )
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)
    except Exception as e:
        logger.error("No se pudo configurar el RotatingFileHandler para el archivo de logs: %s", e)
        file_handler = None
        
    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(console_handler)
    if file_handler:
        root_logger.addHandler(file_handler)
        
    # Silenciar el ruido de logs de getUpdates constantes de la librería httpx
    logging.getLogger("httpx").setLevel(logging.WARNING)
    
    logger.info("Sistema de logs configurado. Archivo de logs: %s", log_file)

VOICE_BUTTON = "🎤 Activar comando de voz"
VOICE_CANCEL_BUTTON = "❌ Cancelar voz"
COTI_BUTTON = "📋 Nueva Cotización"
NOTA_BUTTON = "📦 Nueva Nota de Entrega"
RETENTION_RATE = Decimal("0.75")


def _main_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(VOICE_BUTTON), KeyboardButton(VOICE_CANCEL_BUTTON)],
            [KeyboardButton(COTI_BUTTON), KeyboardButton(NOTA_BUTTON)],
        ],
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def _allowed(update: Update) -> bool:
    u = update.effective_user
    return u is not None and u.id == config.ALLOWED_USER_ID


async def _deny(update: Update) -> None:
    msg = update.effective_message
    u = update.effective_user
    if msg:
        if u is not None:
            await msg.reply_text(
                "Acceso no autorizado.\n"
                f"Tu id de Telegram es: {u.id}\n"
                "Si este bot es tuyo, copia ese número en TELEGRAM_ALLOWED_USER_ID del .env "
                "y reinicia el bot."
            )
        else:
            await msg.reply_text("Acceso no autorizado.")

async def _notify_same_source_channel(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    text: str,
) -> None:
    """Notifica en SUFEVICA cuando el mensaje viene de ese chat."""
    msg = update.effective_message
    if not msg:
        return
    is_channel = update.channel_post is not None or update.edited_channel_post is not None
    if _is_sufevica_chat(update) or is_channel:
        try:
            await context.bot.send_message(chat_id=msg.chat_id, text=text)
        except Exception:  # noqa: BLE001
            # Si el bot no tiene permiso para publicar en el canal,
            # al menos intentamos responder como fallback.
            await msg.reply_text(text)
        return
    await msg.reply_text(text)


def _extract_labeled_value(text: str, labels: tuple[str, ...]) -> str:
    escaped = [re.escape(lbl) for lbl in labels]
    pattern = (
        r"(?:^|\n|[,;])\s*(?:" + "|".join(escaped) + r")\s*[:=]\s*([^\n,;]+)"
    )
    m = re.search(pattern, text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""


def _extract_labeled_value_eol(text: str, labels: tuple[str, ...]) -> str:
    """Busca etiqueta: valor hasta fin de línea (permite comas en montos)."""
    labels_norm = tuple(_normalize_text(lbl).strip() for lbl in labels)
    for raw in text.splitlines():
        line = raw.strip()
        if ":" not in line:
            continue
        key, _, rest = line.partition(":")
        key_norm = _normalize_text(re.sub(r"^[\s*•\-\d.\)\(]+", "", key)).strip()
        if not key_norm:
            continue
        if any(
            key_norm == lbl
            or key_norm.startswith(lbl + " ")
            or lbl in key_norm
            for lbl in labels_norm
        ):
            val = re.sub(r"\*+", "", rest).strip()
            if val:
                return val
    return ""


def _parse_retencion_from_any_text(text: str) -> dict[str, str] | None:
    """
    Parseo flexible para textos del canal SUFEVICA sin exigir verbo de intención.
    """
    t = text.strip()
    data = {
        "fecha_emision": _extract_labeled_value_eol(
            t,
            ("fecha_emision", "fecha emision", "fecha de emision", "fecha de emisión", "fecha"),
        ),
        "numero_comprobante": _extract_labeled_value_eol(
            t,
            (
                "numero_comprobante",
                "nro_comprobante",
                "nro comprobante",
                "numero de comprobante",
                "número de comprobante",
                "comprobante",
            ),
        ),
        "rif": _extract_labeled_value_eol(t, ("rif",)),
        "fechas_facturas": _extract_labeled_value_eol(
            t,
            ("fechas_facturas", "fecha_factura", "fecha factura", "fechas facturas"),
        ),
        "numeros_facturas": _extract_labeled_value_eol(
            t,
            (
                "numeros_facturas",
                "numero_factura",
                "nro_factura",
                "nro factura",
                "número factura",
                "número de factura",
                "numero de factura",
                "factura afectada",
                "factura",
            ),
        ),
        "controles_facturas": _extract_labeled_value_eol(
            t,
            (
                "controles_facturas",
                "control_factura",
                "control",
                "nro control",
                "numero de control",
                "número de control",
                "numero de control",
            ),
        ),
        "total_compra_con_iva": _extract_labeled_value_eol(
            t,
            (
                "total_compra_con_iva",
                "total compra",
                "total compra con iva",
                "total",
                "total factura",
                "total general",
            ),
        ),
        "base_imponible": _extract_labeled_value_eol(
            t,
            ("base_imponible", "base imponible", "base gravable", "base"),
        ),
        "iva_retenido": _extract_labeled_value_eol(
            t,
            (
                "iva_retenido",
                "iva retenido",
                "monto iva retenido",
                "iva retenido total",
                "impuesto retenido",
                "monto retenido",
                "retenido",
            ),
        ),
    }
    required = ("fecha_emision", "numero_comprobante", "rif", "iva_retenido")
    if any(not data[k] for k in required):
        return None
    return data


def _parse_retencion_entry_request(text: str) -> dict[str, str] | None:
    t = text.strip()
    t_norm = _normalize_text(t)
    # Acepta "registrar una retencion", "quiero guardar retencion", etc.
    has_ret = "retencion" in t_norm or "retenciones" in t_norm
    intent_cmd = any(
        k in t_norm
        for k in (
            "registrar retencion",
            "guardar retencion",
            "nueva retencion",
            "registro retencion",
        )
    )
    intent_natural = has_ret and any(
        w in t_norm
        for w in (
            "registrar",
            "guardar",
            "nueva",
            "agregar",
            "anadir",
            "ingresar",
            "cargar",
            "reportar",
        )
    )
    if not (intent_cmd or intent_natural):
        return None

    data = {
        "fecha_emision": _extract_labeled_value(
            t,
            ("fecha_emision", "fecha emision", "fecha"),
        ),
        "numero_comprobante": _extract_labeled_value(
            t,
            ("numero_comprobante", "nro_comprobante", "nro comprobante", "comprobante"),
        ),
        "rif": _extract_labeled_value(t, ("rif",)),
        "fechas_facturas": _extract_labeled_value(
            t,
            ("fechas_facturas", "fecha_factura", "fecha factura"),
        ),
        "numeros_facturas": _extract_labeled_value(
            t,
            (
                "numeros_facturas",
                "numero_factura",
                "nro_factura",
                "nro factura",
                "número de factura",
                "numero de factura",
                "factura afectada",
                "factura",
            ),
        ),
        "controles_facturas": _extract_labeled_value(
            t,
            (
                "controles_facturas",
                "control_factura",
                "control",
                "nro control",
                "número de control",
                "numero de control",
            ),
        ),
        "total_compra_con_iva": _extract_labeled_value(
            t,
            (
                "total_compra_con_iva",
                "total compra",
                "total",
                "total factura",
                "total general",
            ),
        ),
        "base_imponible": _extract_labeled_value(
            t,
            ("base_imponible", "base imponible", "base"),
        ),
        "iva_retenido": _extract_labeled_value(
            t,
            (
                "iva_retenido",
                "iva retenido",
                "monto iva retenido",
                "impuesto retenido",
                "monto retenido",
                "retenido",
            ),
        ),
    }
    required = ("fecha_emision", "numero_comprobante", "rif", "iva_retenido")
    if any(not data[k] for k in required):
        return None
    return data


def _parse_retenciones_report_request(
    text: str,
) -> tuple[date, date, str] | None:
    t = text.lower().strip()
    t_norm = _normalize_text(t)
    if "retencion" not in t_norm and "retenciones" not in t_norm:
        return None
    fmt = "excel"
    if " pdf" in f" {t_norm} ":
        fmt = "pdf"
    elif " excel" in f" {t_norm} ":
        fmt = "excel"
    m = re.search(
        r"(\d{1,2}[\/\-\s]\d{1,2}[\/\-\s]\d{2,4}).*?(\d{1,2}[\/\-\s]\d{1,2}[\/\-\s]\d{2,4})",
        t,
    )
    if not m:
        return None
    d1 = _parse_user_date(m.group(1))
    d2 = _parse_user_date(m.group(2))
    if d1 is None or d2 is None:
        return None
    date_from, date_to = (d1, d2) if d1 <= d2 else (d2, d1)
    return date_from, date_to, fmt


def _normalize_text(text: str) -> str:
    base = unicodedata.normalize("NFKD", text)
    plain = "".join(ch for ch in base if not unicodedata.combining(ch))
    return plain.lower().strip()


def _parse_user_date(s: str) -> date | None:
    s = re.sub(r"\s+", "/", s.strip()).replace("-", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _parse_emitir_retencion_request(text: str) -> list[str] | None:
    """
    Ejemplo: Emitir retencion de facturas:06949950|06949951
    """
    m = re.search(
        r"(?is)\bemitir\s+retencion\s+de\s+facturas?\s*:\s*(.+)$",
        text,
    )
    if not m:
        return None
    raw = m.group(1)
    docs = [x.strip() for x in re.split(r"[|,]", raw) if x.strip()]
    # Acepta números o alfanuméricos (ej. B00093101). Normaliza removiendo espacios.
    docs = [re.sub(r"\s+", "", x).upper() for x in docs]
    docs = [x for x in docs if re.search(r"[A-Z0-9]", x)]
    if not docs:
        return None
    # preservar orden sin duplicados
    seen: set[str] = set()
    out: list[str] = []
    for d in docs:
        if d in seen:
            continue
        seen.add(d)
        out.append(d)
    return out


def _periodo_fiscal(d: date) -> str:
    return d.strftime("%Y-%m")


def _reten_emit_monthly_path(emission_date: date) -> Path:
    base_dir = config.RETENCIONES_EMITIDAS_DIR
    return excel_store.monthly_retencion_emitida_path(base_dir, emission_date)


def _totals_for_items(items: list[excel_store.FacturaCompraRow]) -> tuple[Decimal, Decimal, Decimal]:
    base_total = Decimal("0")
    iva_total = Decimal("0")
    for it in items:
        base_total += it.base_imponible or Decimal("0")
        iva_total += it.monto_iva or Decimal("0")
    retenido = (iva_total * RETENTION_RATE).quantize(Decimal("0.01"))
    return base_total, iva_total, retenido


async def _start_emitir_retencion_flow(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    doc_numbers: list[str],
) -> None:
    msg = update.effective_message
    if not msg:
        return
    items = excel_store.load_facturas_by_document_numbers(
        config.FACTURAS_RECIBIDAS_PATH,
        doc_numbers,
    )
    if not items:
        await msg.reply_text(
            "No encontré facturas en FACTURAS-RECIBIDAS para esos números de documento."
        )
        return
    found_docs = {it.numero_documento for it in items}
    missing = [d for d in doc_numbers if d not in found_docs]
    if missing:
        await msg.reply_text(
            "Faltan estas facturas en el Excel de recibidas: " + ", ".join(missing)
        )
        return
    rifs = {it.proveedor_rif for it in items if it.proveedor_rif}
    if len(rifs) > 1:
        await msg.reply_text(
            "Las facturas seleccionadas tienen proveedores distintos. Emite por proveedor."
        )
        return
    provider = (items[0].proveedor or "").strip()
    provider_rif = (items[0].proveedor_rif or "").strip()
    provider_phone = (items[0].proveedor_telefono or "").strip()
    provider_address = (items[0].direccion_fiscal_proveedor or "").strip()
    base_total, iva_total, retenido = _totals_for_items(items)
    emission_date = date.today()
    monthly_path = _reten_emit_monthly_path(emission_date)
    next_num = excel_store.next_retencion_emitida_number(
        monthly_path,
        emission_date=emission_date,
    )
    context.user_data["pending_emit_ret"] = {
        "docs": doc_numbers,
        "provider": provider,
        "provider_rif": provider_rif,
        "provider_phone": provider_phone,
        "provider_address": provider_address,
        "base_total": str(base_total),
        "iva_total": str(iva_total),
        "retenido": str(retenido),
        "seq_mode": "auto",
    }
    kb = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("SI, seguir secuencia", callback_data="emit_seq_yes"),
                InlineKeyboardButton("NO, colocar manual", callback_data="emit_seq_no"),
            ]
        ]
    )
    await msg.reply_text(
        "Comprobante de retención listo para generar.\n"
        f"Proveedor: {provider or '—'} ({provider_rif or '—'})\n"
        f"Facturas: {' | '.join(doc_numbers)}\n"
        f"IVA total: {iva_total} | IVA retenido (75%): {retenido}\n"
        f"Siguiente correlativo según Excel ({monthly_path.parent.name}): {next_num}\n\n"
        "¿Deseas seguir la secuencia automática del Excel?",
        reply_markup=kb,
    )

def _parse_reporte_z_nuevo(text: str) -> dict[str, str] | None:
    t = text.strip()
    # Verificar si tiene el encabezado de Reporte Z y campos característicos
    if "reporte z" not in t.lower():
        return None
    if "desglose financiero" not in t.lower() and "sub-total" not in t.lower():
        return None

    # Extraer el número del reporte Z
    m_num = re.search(r"reporte z\s*:\s*(\d+)", t, flags=re.IGNORECASE)
    if not m_num:
        m_num = re.search(r"reporte z\s+(\d+)", t, flags=re.IGNORECASE)
        if not m_num:
            return None
    num_rep = m_num.group(1)

    # Extraer fecha de emisión
    m_fecha = re.search(r"fecha de emisi[oó]n\s*:\s*([^\n]+)", t, flags=re.IGNORECASE)
    if not m_fecha:
        m_fecha = re.search(r"fecha\s*:\s*([^\n]+)", t, flags=re.IGNORECASE)
        if not m_fecha:
            return None
    fecha_val = m_fecha.group(1).strip()

    # Extraer montos
    m_sub = re.search(r"sub-total\s*:\s*(?:bs\.?\s*)?([^\n]+)", t, flags=re.IGNORECASE)
    m_base = re.search(r"base imponible\s*:\s*(?:bs\.?\s*)?([^\n]+)", t, flags=re.IGNORECASE)
    m_exento = re.search(r"monto exento\s*:\s*(?:bs\.?\s*)?([^\n]+)", t, flags=re.IGNORECASE)
    m_iva = re.search(r"iva\s*(?:\(16%\))?\s*:\s*(?:bs\.?\s*)?([^\n]+)", t, flags=re.IGNORECASE)
    m_total = re.search(r"total general\s*(?:bs\.?\s*)?\s*:\s*([^\n]+)", t, flags=re.IGNORECASE)

    sub_str = m_sub.group(1).strip() if m_sub else ""
    base_str = m_base.group(1).strip() if m_base else ""
    exento_str = m_exento.group(1).strip() if m_exento else "0,00"
    iva_str = m_iva.group(1).strip() if m_iva else ""
    tot_str = m_total.group(1).strip() if m_total else ""

    # Validar que al menos tengamos base o total
    if not (base_str or tot_str):
        return None

    # Parsear y limpiar montos a strings decimales estándar
    try:
        sub_dec = excel_store.parse_amount_ves_string(sub_str)
        base_dec = excel_store.parse_amount_ves_string(base_str)
        exento_dec = excel_store.parse_amount_ves_string(exento_str) or Decimal("0")
        iva_dec = excel_store.parse_amount_ves_string(iva_str)
        tot_dec = excel_store.parse_amount_ves_string(tot_str)

        # Autocompletado si faltan campos
        if base_dec is not None and iva_dec is None:
            iva_dec = (base_dec * Decimal("0.16")).quantize(Decimal("0.01"))
        if base_dec is not None and tot_dec is None:
            tot_dec = base_dec + exento_dec + (iva_dec or Decimal("0"))
        if tot_dec is not None and base_dec is None:
            base_dec = ((tot_dec - exento_dec) / Decimal("1.16")).quantize(Decimal("0.01"))
            iva_dec = tot_dec - exento_dec - base_dec

        return {
            "numero_reporte": num_rep,
            "fecha_emision": fecha_val,
            "sub_total": str(sub_dec) if sub_dec is not None else (str(base_dec) if base_dec is not None else ""),
            "base_imponible": str(base_dec) if base_dec is not None else "",
            "monto_exento": str(exento_dec),
            "iva": str(iva_dec) if iva_dec is not None else "",
            "total": str(tot_dec) if tot_dec is not None else "",
        }
    except Exception:
        return None


def _parse_venta_o_reportez(text: str) -> dict[str, str] | None:
    t = text.strip()
    t_norm = _normalize_text(t)
    
    is_venta = (
        "venta" in t_norm
        or "factura emitida" in t_norm
        or "facturas emitidas" in t_norm
        or "factura manual" in t_norm
        or "facturas manuales" in t_norm
        or "factura de venta" in t_norm
    )
    is_z = (
        "reporte z" in t_norm
        or "reportez" in t_norm
        or "impresora fiscal" in t_norm
        or "resumen diario" in t_norm
        or "resumen diario z" in t_norm
        or "cierre z" in t_norm
        or "cierre diario" in t_norm
    )
    
    if not (is_venta or is_z):
        return None
        
    clasif = "Reporte Z" if is_z else "Factura Emitida"
    
    # 1. Intentar extracción con etiquetas estándar (con dos puntos o igual)
    fecha = _extract_labeled_value(t, ("fecha", "fecha_emision", "fecha emision"))
    num_doc = _extract_labeled_value(t, ("factura", "numero", "nro", "numero_documento", "documento"))
    cliente = _extract_labeled_value(t, ("cliente", "razon_social", "razon social", "comprador"))
    rif = _extract_labeled_value(t, ("rif", "rif_cliente", "rif cliente"))
    base = _extract_labeled_value(t, ("base", "base_imponible", "base imponible"))
    iva = _extract_labeled_value(t, ("iva", "impuesto"))
    total = _extract_labeled_value(t, ("total", "total_venta", "monto_total"))
    
    if not fecha:
        fecha = _extract_labeled_value_eol(t, ("fecha", "fecha emision"))
    if not num_doc:
        num_doc = _extract_labeled_value_eol(t, ("factura", "numero", "nro", "documento"))
    if not cliente:
        cliente = _extract_labeled_value_eol(t, ("cliente", "razon social"))
    if not rif:
        rif = _extract_labeled_value_eol(t, ("rif",))
    if not base:
        base = _extract_labeled_value_eol(t, ("base", "base imponible"))
    if not iva:
        iva = _extract_labeled_value_eol(t, ("iva",))
    if not total:
        total = _extract_labeled_value_eol(t, ("total",))

    # 2. Robustez: Extracción de fecha sin dos puntos (patrones DD/MM/YYYY o DD-MM-YYYY o YYYY-MM-DD)
    if not fecha:
        m_date = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", t)
        if m_date:
            day_str, month_str, year_str = m_date.groups()
            if len(year_str) == 2:
                year_str = "20" + year_str
            fecha = f"{day_str.zfill(2)}/{month_str.zfill(2)}/{year_str}"
        else:
            fecha = date.today().strftime("%d/%m/%Y")
            
    # 3. Robustez: Extracción de número de documento / reporte
    if not num_doc:
        if is_z:
            m_num = re.search(r"(?:z|reporte|nro|numero|#)\s*(\d+)", t, flags=re.IGNORECASE)
            if m_num:
                num_doc = m_num.group(1)
            else:
                parsed_date = tributario_engine._parse_row_date(fecha) or date.today()
                num_doc = f"Z-{parsed_date.strftime('%Y%m%d')}"
        else:
            num_doc = "S/N"
            
    # 4. Robustez: Extracción de montos sin dos puntos
    if not base:
        m_base = re.search(r"\b(?:base|imponible)\s+(\d+(?:[\.,]\d+)?)", t, flags=re.IGNORECASE)
        if m_base:
            base = m_base.group(1)
    if not iva:
        m_iva = re.search(r"\b(?:iva|impuesto)\s+(\d+(?:[\.,]\d+)?)", t, flags=re.IGNORECASE)
        if m_iva:
            iva = m_iva.group(1)
    if not total:
        m_total = re.search(r"\b(?:total)\s+(\d+(?:[\.,]\d+)?)", t, flags=re.IGNORECASE)
        if m_total:
            total = m_total.group(1)

    # 5. Valores predeterminados para Reporte Z
    if is_z:
        if not cliente:
            cliente = "CONSUMIDOR FINAL (VENTAS DIARIAS)"
        if not rif:
            rif = "V-00000000-0"

    # Verificar que al menos tengamos algún monto
    if not (base or iva or total):
        return None
        
    return {
        "clasificacion": clasif,
        "estado": "REGISTRADO",
        "fecha": fecha,
        "numero_documento": num_doc,
        "razon_social": cliente,
        "rif": rif,
        "base_imponible": base or "",
        "iva": iva or "",
        "total": total or "",
    }


def _match_intent(text: str) -> str | None:
    t = text.lower().strip()
    t_norm = _normalize_text(t)
    if any(
        p in t_norm
        for p in (
            "excel facturas compra",
            "excel de facturas compra",
            "enviar excel facturas compra",
            "mandar excel facturas compra",
            "descargar excel facturas compra",
            "facturas compra xlsx",
            "excel facturas recibidas",
        )
    ):
        return "send_facturas_compra_excel"
    if any(
        p in t_norm
        for p in (
            "enviar excel",
            "manda el excel",
            "mandar excel",
            "excel consolidado",
            "descargar excel",
            "envíame el excel",
            "enviame el excel",
        )
    ):
        return "send_excel"
    if any(
        p in t_norm
        for p in (
            "resumen de hoy",
            "resumen hoy",
            "dame el resumen",
            "el resumen de hoy",
        )
    ):
        return "summary_today"
    if any(
        p in t_norm
        for p in (
            "tributos",
            "compromisos tributarios",
            "iva por pagar",
            "reporte quincenal",
            "quincena actual",
            "corte de quincena",
        )
    ):
        return "tributos_report"
    if _parse_retenciones_report_request(text) is not None:
        return "retenciones_report"
    return None


def _is_sufevica_chat(update: Update) -> bool:
    msg = update.effective_message
    if not msg or not msg.chat:
        return False
    if config.SOURCE_CHANNEL_ID is not None and msg.chat_id == config.SOURCE_CHANNEL_ID:
        return True
    # Fallback por nombre para evitar caída si cambió el chat_id del canal/grupo.
    title = _normalize_text(getattr(msg.chat, "title", "") or "")
    username = _normalize_text(getattr(msg.chat, "username", "") or "")
    if "sufevica" in title or "sufevica" in username:
        return True
    return False


def _parse_document_text_explicit(text: str, forced_type: str) -> dict | None:
    doc_data = _parse_document_text(text)
    if doc_data is None:
        doc_data = _parse_document_text_relaxed(text, forced_type)
    if doc_data is not None:
        doc_data["docType"] = forced_type
    return doc_data


def _parse_document_items_master(text: str) -> list:
    lines = text.split('\n')
    if not lines:
        return []
    
    stitched_lines = [lines[0].strip()]
    for i in range(1, len(lines)):
        prev_line = stitched_lines[-1].strip()
        curr_line = lines[i].strip()
        
        if not prev_line or not curr_line:
            if curr_line:
                stitched_lines.append(curr_line)
            continue
            
        prev_words = prev_line.split()
        curr_words = curr_line.split()
        
        if not prev_words or not curr_words:
            stitched_lines.append(curr_line)
            continue
            
        last_word = prev_words[-1]
        first_word = curr_words[0]
        
        should_stitch = False
        
        # Case 1: Purely alphabetic word split (e.g. "TUBER", "IA" or "CAN", "T")
        if last_word.isalpha() and first_word.isalpha():
            if len(last_word) <= 2 or len(first_word) <= 2:
                should_stitch = True
                
        # Case 2: Split inside fractions or decimals (e.g. "1", "-1/2" or "33,", "05")
        elif re.match(r'.*\d$', last_word) and re.match(r'^[-/]', first_word):
            should_stitch = True
        elif re.match(r'.*[-/]$', last_word) and re.match(r'^\d', first_word):
            should_stitch = True
        elif re.match(r'.*[,.]$', last_word) and re.match(r'^\d', first_word):
            should_stitch = True
        elif re.match(r'.*\d$', last_word) and re.match(r'^[,.]', first_word):
            should_stitch = True
            
        if should_stitch:
            prev_line_rest = " ".join(prev_words[:-1])
            stitched_word = last_word + first_word
            
            if prev_line_rest:
                stitched_lines[-1] = prev_line_rest + " " + stitched_word
            else:
                stitched_lines[-1] = stitched_word
                
            curr_line_rest = " ".join(curr_words[1:])
            if curr_line_rest:
                stitched_lines.append(curr_line_rest)
        else:
            stitched_lines.append(curr_line)
            
    cleaned_unified = " ".join([l for l in stitched_lines if l.strip()])
    sp_parts = cleaned_unified.split()
    
    def parse_token_val(tok):
        clean = re.sub(r"[^\d.,\-]", "", tok)
        if "," in clean and "." in clean:
            clean = clean.replace(".", "").replace(",", ".")
        elif "," in clean:
            clean = clean.replace(",", ".")
        return float(clean) if clean else 0.0

    matches = []
    j = 2
    while j < len(sp_parts):
        try:
            qty_val = parse_token_val(sp_parts[j-2])
            price_val = parse_token_val(sp_parts[j-1])
            total_val = parse_token_val(sp_parts[j])
            
            if qty_val > 0 and price_val > 0 and total_val > 0:
                calc = qty_val * price_val
                if abs(calc - total_val) < 0.05 * total_val:
                    matches.append(j)
                    j += 3 # skip to avoid overlapping matches
                    continue
        except Exception:
            pass
        j += 1

    items = []
    headers_to_strip = {"codigo", "producto", "cant", "p.u.", "total", "pu", "cant.", "total."}
    common_desc_words = {"de", "para", "con", "en", "tubo", "codo", "anillo", "brida", "niple", "tee", "af", "pvc", "hg", "galvanizado", "soldadura"}
    
    for k in range(len(matches)):
        idx = matches[k]
        start = matches[k-1] + 1 if k > 0 else 0
        
        qty_val = parse_token_val(sp_parts[idx-2])
        price_val = parse_token_val(sp_parts[idx-1])
        
        text_tokens = sp_parts[start : idx-2]
        
        if k == 0:
            text_tokens = [t for t in text_tokens if t.lower() not in headers_to_strip]
            
        if not text_tokens:
            continue
            
        first_tok = text_tokens[0]
        is_code = len(first_tok) <= 15 and (
            any(c.isdigit() for c in first_tok) or 
            first_tok.lower() not in common_desc_words
        )
        
        if is_code and len(text_tokens) > 1:
            code = first_tok.upper()
            desc = " ".join(text_tokens[1:])
        else:
            code = ""
            desc = " ".join(text_tokens)
            
        items.append({
            "code": code,
            "desc": desc.strip(),
            "qty": qty_val,
            "priceUsd": price_val
        })
        
    return items


_bcv_rate_cache = {
    "rate": 550.00,
    "last_fetched": 0.0
}

def _fetch_bcv_usd_rate() -> float:
    """Obtiene la tasa oficial de cambio USD/VES del BCV desde DolarAPI con fallback raspando la web del BCV."""
    url_api = "https://ve.dolarapi.com/v1/dolares/oficial"
    headers = {'User-Agent': 'Mozilla/5.0'}
    
    import urllib.request
    import json
    import ssl
    import re
    
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    
    try:
        req = urllib.request.Request(url_api, headers=headers)
        with urllib.request.urlopen(req, context=ctx, timeout=5) as response:
            data = json.loads(response.read().decode('utf-8'))
            if data and "promedio" in data:
                return float(data["promedio"])
    except Exception as e:
        logger.warning(f"Error al obtener tasa BCV de DolarAPI: {e}. Intentando raspado directo...")
        
    url_bcv = "https://www.bcv.org.ve"
    try:
        req = urllib.request.Request(url_bcv, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        with urllib.request.urlopen(req, context=ctx, timeout=8) as response:
            html = response.read().decode('utf-8')
            match = re.search(r'id="dolar"[^>]*>.*?<strong[^>]*>\s*([\d,.]+)\s*</strong>', html, re.DOTALL)
            if match:
                val = match.group(1).strip()
                val_float = float(val.replace('.', '').replace(',', '.'))
                return val_float
    except Exception as ex:
        logger.error(f"Error al raspar página oficial del BCV: {ex}")
        
    return 550.00

def get_current_bcv_rate() -> float:
    import time
    now = time.time()
    if now - _bcv_rate_cache["last_fetched"] > 21600:
        try:
            rate = _fetch_bcv_usd_rate()
            _bcv_rate_cache["rate"] = rate
            _bcv_rate_cache["last_fetched"] = now
        except Exception as e:
            logger.error(f"Error en get_current_bcv_rate: {e}")
    return _bcv_rate_cache["rate"]


SYNC_FILES = {
    "reten_rec": "RETEN-REC.xlsx",
    "facturas_recibidas": "FACTURAS-RECIBIDAS-NUEVO.xlsx",
    "facturas_emitidas": "FACTURAS-EMITIDAS.xlsx",
    "reportes_z": "REPORTES-Z-NUEVO.xlsx"
}

def get_sync_file_path(key: str) -> Path | None:
    import config
    if key == "reten_rec":
        return config.EXCEL_PATH
    elif key == "facturas_recibidas":
        return config.FACTURAS_RECIBIDAS_PATH
    elif key == "facturas_emitidas":
        return config.FACTURAS_EMITIDAS_PATH
    elif key == "reportes_z":
        return config.REPORTES_Z_PATH
    return None

_last_mtime_cache = {}
PIN_PREFIX = "[SUFEVICA_BACKUP_STATE] "

async def _get_pinned_state(bot) -> dict:
    import config
    import json
    try:
        chat = await bot.get_chat(chat_id=config.ALLOWED_USER_ID)
        if chat.pinned_message and chat.pinned_message.text and chat.pinned_message.text.startswith(PIN_PREFIX):
            json_str = chat.pinned_message.text[len(PIN_PREFIX):].strip()
            return json.loads(json_str)
    except Exception as e:
        logger.error(f"Error al obtener el estado fijado de respaldos: {e}")
    return {}

async def _save_pinned_state(bot, state: dict) -> None:
    import config
    import json
    text = f"{PIN_PREFIX}{json.dumps(state)}"
    try:
        chat = await bot.get_chat(chat_id=config.ALLOWED_USER_ID)
        if chat.pinned_message and chat.pinned_message.text and chat.pinned_message.text.startswith(PIN_PREFIX):
            try:
                await bot.edit_message_text(
                    chat_id=config.ALLOWED_USER_ID,
                    message_id=chat.pinned_message.message_id,
                    text=text
                )
                return
            except Exception:
                pass
        msg = await bot.send_message(chat_id=config.ALLOWED_USER_ID, text=text)
        await bot.pin_chat_message(chat_id=config.ALLOWED_USER_ID, message_id=msg.message_id, disable_notification=True)
    except Exception as e:
        logger.error(f"Error al guardar el estado fijado de respaldos: {e}")

async def check_and_sync_files(context: ContextTypes.DEFAULT_TYPE) -> None:
    import os
    import config
    changed = False
    current_state = await _get_pinned_state(context.bot)
    
    for key, filename in SYNC_FILES.items():
        path = get_sync_file_path(key)
        if not path or not path.exists():
            continue
        mtime = os.path.getmtime(path)
        last_mtime = _last_mtime_cache.get(key)
        if last_mtime is None:
            _last_mtime_cache[key] = mtime
            continue
        if mtime > last_mtime:
            logger.info(f"Detectado cambio en {filename}. Subiendo respaldo...")
            try:
                with open(path, "rb") as f:
                    sent_msg = await context.bot.send_document(
                        chat_id=config.ALLOWED_USER_ID,
                        document=f,
                        filename=filename,
                        caption=f"📦 Respaldo automático de {filename}",
                        disable_notification=True
                    )
                new_file_id = sent_msg.document.file_id
                current_state[key] = new_file_id
                _last_mtime_cache[key] = mtime
                changed = True
            except Exception as e:
                logger.error(f"Error subiendo respaldo de {filename}: {e}")
    if changed:
        await _save_pinned_state(context.bot, current_state)

async def restore_files_from_backup(bot) -> None:
    import os
    current_state = await _get_pinned_state(bot)
    if not current_state:
        logger.info("No se encontró ningún respaldo fijado. Iniciando con archivos locales actuales.")
        return
    logger.info("Restaurando archivos de Excel desde el respaldo de Telegram...")
    for key, file_id in current_state.items():
        path = get_sync_file_path(key)
        if not path:
            continue
        filename = SYNC_FILES[key]
        try:
            logger.info(f"Descargando {filename} desde Telegram...")
            tg_file = await bot.get_file(file_id)
            path.parent.mkdir(parents=True, exist_ok=True)
            await tg_file.download_to_drive(custom_path=str(path))
            logger.info(f"Archivo {filename} restaurado con éxito.")
            _last_mtime_cache[key] = os.path.getmtime(path)
        except Exception as e:
            logger.error(f"Error descargando respaldo de {filename}: {e}")

async def post_init(application: Application) -> None:
    # 1. Restaurar archivos desde el respaldo de Telegram al iniciar
    await restore_files_from_backup(application.bot)
    # 2. Iniciar job periódico de monitoreo cada 15 segundos
    if application.job_queue:
        application.job_queue.run_repeating(check_and_sync_files, interval=15, first=10)

async def descargar_excel_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _allowed(update):
        await _deny(update)
        return
    msg = update.effective_message
    if not msg:
        return
    sent_any = False
    for key, filename in SYNC_FILES.items():
        path = get_sync_file_path(key)
        if path and path.exists():
            try:
                with open(path, "rb") as f:
                    await msg.reply_document(
                        document=f,
                        filename=filename,
                        caption=f"📊 Archivo: {filename}"
                    )
                sent_any = True
            except Exception as e:
                logger.error(f"Error al enviar {filename}: {e}")
    if not sent_any:
        await msg.reply_text("⚠️ No se encontraron archivos de Excel locales en el servidor.")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if not msg or not msg.document:
        return
    if not _allowed(update):
        await _deny(update)
        return
    filename = msg.document.file_name
    matched_key = None
    for key, name in SYNC_FILES.items():
        if filename.lower() == name.lower():
            matched_key = key
            break
    if matched_key:
        path = get_sync_file_path(matched_key)
        status_msg = await msg.reply_text(f"📥 *Recibido {filename}. Procesando y reemplazando archivo local...*", parse_mode="Markdown")
        try:
            tg_file = await context.bot.get_file(msg.document.file_id)
            path.parent.mkdir(parents=True, exist_ok=True)
            await tg_file.download_to_drive(custom_path=str(path))
            import os
            _last_mtime_cache[matched_key] = os.path.getmtime(path)
            current_state = await _get_pinned_state(context.bot)
            current_state[matched_key] = msg.document.file_id
            await _save_pinned_state(context.bot, current_state)
            await status_msg.edit_text(f"✅ *¡Archivo `{filename}` actualizado con éxito!* El bot trabajará con esta nueva versión.", parse_mode="Markdown")
        except Exception as e:
            logger.error(f"Error al descargar archivo subido {filename}: {e}")
            await status_msg.edit_text(f"❌ *Error al procesar el archivo:* `{e}`", parse_mode="Markdown")


def _parse_document_text_relaxed(text: str, doc_type: str) -> dict | None:
    client_name = _take_eol_label(text, ("cliente", "razon social", "razón social", "nombre", "comprador", "dirigido a"))
    rif = _take_eol_label(text, ("rif", "c.i.", "ci", "cedula", "cédula", "identificacion", "rif cliente"))
    address = _take_eol_label(text, ("direccion", "dirección", "direccion fiscal", "dirección fiscal"))
    phone = _take_eol_label(text, ("telefono", "teléfono", "telefonos", "teléfonos", "telf", "celular"))
    doc_number = _take_eol_label(text, ("numero", "nro", "número", "documento", "cotizacion nro", "cotización nro", "nota nro"))
    salesman = _take_eol_label(text, ("vendedor", "ejecutivo", "atendido por")) or "FREDDY LOPEZ"
    sale_type = _take_eol_label(text, ("tipo", "tipo de venta", "condición de pago", "condicion de pago")) or "Contado"
    
    if not client_name:
        m_cli = re.search(r"(?im)(?:cliente|razon social|sres|señores)\s*[:\-]\s*(.+)$", text)
        if m_cli:
            client_name = m_cli.group(1).strip()
    if not rif:
        m_rif = re.search(r"\b([VEJPGvejpg][\s.\-]*\d{6,12}(?:[\s.\-]\d)?)\b", text)
        if m_rif:
            rif = re.sub(r"\s+", "", m_rif.group(1).strip())
            
    if rif:
        rif = _normalize_rif(rif)
        
    if not doc_number:
        m_num = re.search(r"(?im)(?:nro|numero|#)\s*[:\-]?\s*(\d+)\b", text)
        if m_num:
            doc_number = m_num.group(1).strip()
        else:
            doc_number = ""
            
    # Try the master stream parser first
    items = _parse_document_items_master(text)
    if items:
        return {
            "docType": doc_type,
            "currency": "usd",
            "exchangeRate": get_current_bcv_rate(),
            "docNumber": doc_number,
            "docDate": date.today().strftime("%Y-%m-%d"),
            "client": {
                "name": client_name or "CLIENTE REGISTRADO POR TEXTO",
                "address": address or "DIRECCIÓN DE CLIENTE",
                "rif": rif or "V-00000000-0",
                "phone": phone or "—",
                "salesman": salesman,
                "saleType": sale_type
            },
            "items": items
        }
            
    items = []
    lines = text.splitlines()
    
    for line in lines:
        line_clean = line.strip()
        if not line_clean:
            continue
            
        t_line = _normalize_text(line_clean)
        if any(w in t_line for w in ("total", "subtotal", "sub-total", "iva", "cliente", "rif", "direccion", "vendedor", "fecha", "forma libre")):
            continue
            
        parts = [p.strip() for p in re.split(r"\s{2,}|\t", line_clean) if p.strip()]
        if len(parts) >= 3:
            nums = []
            for i, p in enumerate(parts):
                clean_num = re.sub(r"[^\d.,\-]", "", p)
                try:
                    val = float(clean_num.replace(".", "").replace(",", "."))
                    nums.append((i, val))
                except ValueError:
                    try:
                        val = float(clean_num.replace(",", "."))
                        nums.append((i, val))
                    except ValueError:
                        pass
            
            qty = None
            price = None
            code = ""
            desc = ""
            matched_math = False
            
            if len(nums) >= 2:
                for a in range(len(nums)):
                    for b in range(len(nums)):
                        if a == b: continue
                        for c in range(len(nums)):
                            if c == a or c == b: continue
                            val_a = nums[a][1]
                            val_b = nums[b][1]
                            val_c = nums[c][1]
                            if val_a > 0 and val_b > 0 and abs(val_a * val_b - val_c) < 0.05 * val_c:
                                qty = val_a
                                price = val_b
                                matched_math = True
                                break
                        if matched_math: break
                    if matched_math: break
                    
                if matched_math:
                    num_indices = {nums[a][0], nums[b][0], nums[c][0]}
                    text_parts = [parts[i] for i in range(len(parts)) if i not in num_indices]
                    if len(text_parts) >= 2:
                        code = text_parts[0].upper()
                        desc = " ".join(text_parts[1:])
                    elif len(text_parts) == 1:
                        desc = text_parts[0]
                else:
                    try:
                        first_clean = re.sub(r"[^\d.,\-]", "", parts[0])
                        qty = float(first_clean.replace(".", "").replace(",", "."))
                        
                        last_clean = re.sub(r"[^\d.,\-]", "", parts[-1])
                        price = float(last_clean.replace(".", "").replace(",", "."))
                        
                        if len(parts) >= 4:
                            penult_clean = re.sub(r"[^\d.,\-]", "", parts[-2])
                            try:
                                price_p = float(penult_clean.replace(".", "").replace(",", "."))
                                if abs(qty * price_p - price) < 0.1 * price:
                                    price = price_p
                            except ValueError:
                                pass
                                
                        text_parts = parts[1:-1] if len(parts) >= 3 else [parts[1]]
                        if len(text_parts) >= 2:
                            code = text_parts[0].upper()
                            desc = " ".join(text_parts[1:])
                        else:
                            desc = text_parts[0]
                    except Exception:
                        pass
                        
                if qty is not None and price is not None and qty > 0 and price >= 0:
                    items.append({
                        "code": code,
                        "desc": desc.strip(),
                        "qty": qty,
                        "priceUsd": price
                    })
                    continue

        # 1.5 Estrategia de Columnas al Final (para tablas copiadas de Excel con formato: CODIGO PRODUCTO CANT P.U. TOTAL)
        sp_parts = [p.strip() for p in line_clean.split() if p.strip()]
        if len(sp_parts) >= 4:
            try:
                def parse_token_val(tok):
                    clean = re.sub(r"[^\d.,\-]", "", tok)
                    if "," in clean and "." in clean:
                        clean = clean.replace(".", "").replace(",", ".")
                    elif "," in clean:
                        clean = clean.replace(",", ".")
                    return float(clean)

                qty_val = parse_token_val(sp_parts[-3])
                price_val = parse_token_val(sp_parts[-2])
                total_val = parse_token_val(sp_parts[-1])
                
                if qty_val > 0 and price_val >= 0 and total_val >= 0 and abs(qty_val * price_val - total_val) < 0.05 * total_val:
                    remaining_tokens = sp_parts[:-3]
                    first_tok = remaining_tokens[0]
                    
                    is_code = len(first_tok) <= 15 and (
                        any(c.isdigit() for c in first_tok) or 
                        not any(v in first_tok.lower() for v in ("de", "para", "con", "en", "tubo", "codo", "anillo", "brida", "niple", "tee"))
                    )
                    
                    if is_code and len(remaining_tokens) > 1:
                        code = first_tok.upper()
                        desc = " ".join(remaining_tokens[1:])
                    else:
                        code = ""
                        desc = " ".join(remaining_tokens)
                        
                    items.append({
                        "code": code,
                        "desc": desc.strip(),
                        "qty": qty_val,
                        "priceUsd": price_val
                    })
                    continue
            except Exception:
                pass

        m = re.match(
            r"^\s*[-*•]?\s*(\d+(?:[\.,]\d+)?)\s*(?:x|unidades|uds|unds|pzs|unid)?\s+([A-Z0-9\-]{3,15})?\s+(.+?)\s*(?:a|precio|costo)?\s+(\d+(?:[\.,]\d+)?)\s*$",
            line_clean,
            flags=re.IGNORECASE
        )
        if m:
            qty_str, code_str, desc_str, price_str = m.groups()
            try:
                qty = float(qty_str.replace(",", "."))
                price = float(price_str.replace(",", "."))
                items.append({
                    "code": code_str.upper() if code_str else "",
                    "desc": desc_str.strip(),
                    "qty": qty,
                    "priceUsd": price
                })
                continue
            except Exception:
                pass
                
        m_simple = re.match(
            r"^\s*[-*•]?\s*(\d+(?:[\.,]\d+)?)\s*(?:x|unidades|uds|unds|pzs|unid)?\s+(.+?)\s*(?:a|precio|costo)?\s+(\d+(?:[\.,]\d+)?)\s*$",
            line_clean,
            flags=re.IGNORECASE
        )
        if m_simple:
            qty_str, desc_str, price_str = m_simple.groups()
            try:
                qty = float(qty_str.replace(",", "."))
                price = float(price_str.replace(",", "."))
                items.append({
                    "code": "",
                    "desc": desc_str.strip(),
                    "qty": qty,
                    "priceUsd": price
                })
                continue
            except Exception:
                pass

    if not items:
        bracket_items = re.findall(r"\[\s*([^\]]+)\s*\]", text)
        for bit in bracket_items:
            parts = [p.strip() for p in bit.split("|")]
            if len(parts) >= 2:
                try:
                    qty = float(parts[0].replace(",", "."))
                    code = parts[1].upper() if len(parts) == 4 else ""
                    desc = parts[2] if len(parts) == 4 else (parts[1] if len(parts) == 3 else parts[0])
                    price = float(parts[-1].replace(",", "."))
                    items.append({
                        "code": code,
                        "desc": desc,
                        "qty": qty,
                        "priceUsd": price
                    })
                except Exception:
                    pass
                    
    if not items:
        return None
        
    return {
        "docType": doc_type,
        "currency": "usd",
        "exchangeRate": get_current_bcv_rate(),
        "docNumber": doc_number,
        "docDate": date.today().strftime("%Y-%m-%d"),
        "client": {
            "name": client_name or "CLIENTE REGISTRADO POR TEXTO",
            "address": address or "DIRECCIÓN DE CLIENTE",
            "rif": rif or "V-00000000-0",
            "phone": phone or "—",
            "salesman": salesman,
            "saleType": sale_type
        },
        "items": items
    }


def _parse_document_text(text: str) -> dict | None:
    t_norm = _normalize_text(text)
    
    # 1. Identificar tipo de documento
    is_cotizacion = "cotizacion" in t_norm or "presupuesto" in t_norm
    is_nota = "nota" in t_norm or "entrega" in t_norm
    
    if not (is_cotizacion or is_nota):
        return None
        
    # 2. Evitar procesar como cotización si parece un comprobante de retención o reporte tributario
    if any(k in t_norm for k in ("comprobante de retencion", "numero de comprobante", "iva retenido", "tributos", "anticipo islr", "debito fiscal")):
        return None
        
    # 3. Si parece una factura de compra recibida de un proveedor externo, no lo procesamos aquí (irá al flujo de facturas)
    is_compra = "proveedor" in t_norm or "emisor" in t_norm or "compra" in t_norm
    has_sufevica_emisor = "sufevica" in t_norm or "vittoria" in t_norm
    if is_compra and not has_sufevica_emisor and ("factura" in t_norm or "compras" in t_norm):
        return None

    # Debe ser una intención de creación explícita o implícita
    has_explicit = any(w in t_norm for w in ("crear", "nueva", "nuevo", "generar", "hacer", "registrar"))
    
    if not has_explicit:
        keywords_doc = ("cotizacion", "cotización", "presupuesto", "nota de entrega", "nota de despacho")
        if not any(k in t_norm for k in keywords_doc):
            return None
            
    doc_type = "cotizacion" if is_cotizacion else "nota"
    
    # Extraer campos cliente
    client_name = _take_eol_label(text, ("cliente", "razon social", "razón social", "nombre", "comprador", "dirigido a"))
    rif = _take_eol_label(text, ("rif", "c.i.", "ci", "cedula", "cédula", "identificacion", "rif cliente"))
    address = _take_eol_label(text, ("direccion", "dirección", "direccion fiscal", "dirección fiscal"))
    phone = _take_eol_label(text, ("telefono", "teléfono", "telefonos", "teléfonos", "telf", "celular"))
    doc_number = _take_eol_label(text, ("numero", "nro", "número", "documento", "cotizacion nro", "cotización nro", "nota nro"))
    salesman = _take_eol_label(text, ("vendedor", "ejecutivo", "atendido por")) or "FREDDY LOPEZ"
    sale_type = _take_eol_label(text, ("tipo", "tipo de venta", "condición de pago", "condicion de pago")) or "Contado"
    
    # Fallbacks de extracción de cliente en caso de texto desestructurado o pegado
    if not client_name:
        m_cli = re.search(r"(?im)(?:cliente|razon social|sres|señores)\s*[:\-]\s*(.+)$", text)
        if m_cli:
            client_name = m_cli.group(1).strip()
    if not rif:
        m_rif = re.search(r"\b([VEJPGvejpg][\s.\-]*\d{6,12}(?:[\s.\-]\d)?)\b", text)
        if m_rif:
            rif = re.sub(r"\s+", "", m_rif.group(1).strip())
            
    if rif:
        rif = _normalize_rif(rif)
        
    if not doc_number:
        m_num = re.search(r"(?im)(?:nro|numero|#|cotizacion|nota)\s*[:\-]?\s*(\d+)\b", text)
        if m_num:
            doc_number = m_num.group(1).strip()
        else:
            import random
            doc_number = f"{random.randint(1, 999):03d}{random.randint(1, 999):03d}"
            
    # Try the master stream parser first
    items = _parse_document_items_master(text)
    if items:
        return {
            "docType": doc_type,
            "currency": "usd",
            "exchangeRate": get_current_bcv_rate(),
            "docNumber": doc_number,
            "docDate": date.today().strftime("%Y-%m-%d"),
            "client": {
                "name": client_name or "CLIENTE REGISTRADO POR TEXTO",
                "address": address or "DIRECCIÓN DE CLIENTE",
                "rif": rif or "V-00000000-0",
                "phone": phone or "—",
                "salesman": salesman,
                "saleType": sale_type
            },
            "items": items
        }
            
    # Extraer ítems
    items = []
    lines = text.splitlines()
    
    for line in lines:
        line_clean = line.strip()
        if not line_clean:
            continue
            
        t_line = _normalize_text(line_clean)
        if any(w in t_line for w in ("total", "subtotal", "sub-total", "iva", "cliente", "rif", "direccion", "vendedor", "fecha", "cotizacion", "nota de entrega", "forma libre")):
            continue
            
        # 1. Estrategia Tabular (Múltiples espacios o tabulaciones)
        parts = [p.strip() for p in re.split(r"\s{2,}|\t", line_clean) if p.strip()]
        if len(parts) >= 3:
            nums = []
            for i, p in enumerate(parts):
                clean_num = re.sub(r"[^\d.,\-]", "", p)
                try:
                    val = float(clean_num.replace(".", "").replace(",", "."))
                    nums.append((i, val))
                except ValueError:
                    try:
                        val = float(clean_num.replace(",", "."))
                        nums.append((i, val))
                    except ValueError:
                        pass
            
            qty = None
            price = None
            code = ""
            desc = ""
            matched_math = False
            
            if len(nums) >= 2:
                for a in range(len(nums)):
                    for b in range(len(nums)):
                        if a == b: continue
                        for c in range(len(nums)):
                            if c == a or c == b: continue
                            val_a = nums[a][1]
                            val_b = nums[b][1]
                            val_c = nums[c][1]
                            if val_a > 0 and val_b > 0 and abs(val_a * val_b - val_c) < 0.05 * val_c:
                                qty = val_a
                                price = val_b
                                matched_math = True
                                break
                        if matched_math: break
                    if matched_math: break
                    
                if matched_math:
                    num_indices = {nums[a][0], nums[b][0], nums[c][0]}
                    text_parts = [parts[i] for i in range(len(parts)) if i not in num_indices]
                    if len(text_parts) >= 2:
                        code = text_parts[0].upper()
                        desc = " ".join(text_parts[1:])
                    elif len(text_parts) == 1:
                        desc = text_parts[0]
                else:
                    try:
                        first_clean = re.sub(r"[^\d.,\-]", "", parts[0])
                        qty = float(first_clean.replace(".", "").replace(",", "."))
                        
                        last_clean = re.sub(r"[^\d.,\-]", "", parts[-1])
                        price = float(last_clean.replace(".", "").replace(",", "."))
                        
                        if len(parts) >= 4:
                            penult_clean = re.sub(r"[^\d.,\-]", "", parts[-2])
                            try:
                                price_p = float(penult_clean.replace(".", "").replace(",", "."))
                                if abs(qty * price_p - price) < 0.1 * price:
                                    price = price_p
                            except ValueError:
                                pass
                                
                        text_parts = parts[1:-1] if len(parts) >= 3 else [parts[1]]
                        if len(text_parts) >= 2:
                            code = text_parts[0].upper()
                            desc = " ".join(text_parts[1:])
                        else:
                            desc = text_parts[0]
                    except Exception:
                        pass
                        
                if qty is not None and price is not None and qty > 0 and price >= 0:
                    items.append({
                        "code": code,
                        "desc": desc.strip(),
                        "qty": qty,
                        "priceUsd": price
                    })
                    continue

        # 1.5 Estrategia de Columnas al Final (para tablas copiadas de Excel con formato: CODIGO PRODUCTO CANT P.U. TOTAL)
        sp_parts = [p.strip() for p in line_clean.split() if p.strip()]
        if len(sp_parts) >= 4:
            try:
                def parse_token_val(tok):
                    clean = re.sub(r"[^\d.,\-]", "", tok)
                    if "," in clean and "." in clean:
                        clean = clean.replace(".", "").replace(",", ".")
                    elif "," in clean:
                        clean = clean.replace(",", ".")
                    return float(clean)

                qty_val = parse_token_val(sp_parts[-3])
                price_val = parse_token_val(sp_parts[-2])
                total_val = parse_token_val(sp_parts[-1])
                
                if qty_val > 0 and price_val >= 0 and total_val >= 0 and abs(qty_val * price_val - total_val) < 0.05 * total_val:
                    remaining_tokens = sp_parts[:-3]
                    first_tok = remaining_tokens[0]
                    
                    is_code = len(first_tok) <= 15 and (
                        any(c.isdigit() for c in first_tok) or 
                        not any(v in first_tok.lower() for v in ("de", "para", "con", "en", "tubo", "codo", "anillo", "brida", "niple", "tee"))
                    )
                    
                    if is_code and len(remaining_tokens) > 1:
                        code = first_tok.upper()
                        desc = " ".join(remaining_tokens[1:])
                    else:
                        code = ""
                        desc = " ".join(remaining_tokens)
                        
                    items.append({
                        "code": code,
                        "desc": desc.strip(),
                        "qty": qty_val,
                        "priceUsd": price_val
                    })
                    continue
            except Exception:
                pass

        # 2. Estrategia de Expresiones Regulares para líneas de texto simples
        m = re.match(
            r"^\s*[-*•]?\s*(\d+(?:[\.,]\d+)?)\s*(?:x|unidades|uds|unds|pzs|unid)?\s+([A-Z0-9\-]{3,15})?\s+(.+?)\s*(?:a|precio|costo)?\s+(\d+(?:[\.,]\d+)?)\s*$",
            line_clean,
            flags=re.IGNORECASE
        )
        if m:
            qty_str, code_str, desc_str, price_str = m.groups()
            try:
                qty = float(qty_str.replace(",", "."))
                price = float(price_str.replace(",", "."))
                items.append({
                    "code": code_str.upper() if code_str else "",
                    "desc": desc_str.strip(),
                    "qty": qty,
                    "priceUsd": price
                })
                continue
            except Exception:
                pass
                
        m_simple = re.match(
            r"^\s*[-*•]?\s*(\d+(?:[\.,]\d+)?)\s*(?:x|unidades|uds|unds|pzs|unid)?\s+(.+?)\s*(?:a|precio|costo)?\s+(\d+(?:[\.,]\d+)?)\s*$",
            line_clean,
            flags=re.IGNORECASE
        )
        if m_simple:
            qty_str, desc_str, price_str = m_simple.groups()
            try:
                qty = float(qty_str.replace(",", "."))
                price = float(price_str.replace(",", "."))
                items.append({
                    "code": "",
                    "desc": desc_str.strip(),
                    "qty": qty,
                    "priceUsd": price
                })
                continue
            except Exception:
                pass

    if not items:
        bracket_items = re.findall(r"\[\s*([^\]]+)\s*\]", text)
        for bit in bracket_items:
            parts = [p.strip() for p in bit.split("|")]
            if len(parts) >= 2:
                try:
                    qty = float(parts[0].replace(",", "."))
                    code = parts[1].upper() if len(parts) == 4 else ""
                    desc = parts[2] if len(parts) == 4 else (parts[1] if len(parts) == 3 else parts[0])
                    price = float(parts[-1].replace(",", "."))
                    items.append({
                        "code": code,
                        "desc": desc,
                        "qty": qty,
                        "priceUsd": price
                    })
                except Exception:
                    pass
                    
    if not items:
        return None
        
    return {
        "docType": doc_type,
        "currency": "usd",
        "exchangeRate": get_current_bcv_rate(),
        "docNumber": doc_number,
        "docDate": date.today().strftime("%Y-%m-%d"),
        "client": {
            "name": client_name or "CLIENTE REGISTRADO POR TEXTO",
            "address": address or "DIRECCIÓN DE CLIENTE",
            "rif": rif or "V-00000000-0",
            "phone": phone or "—",
            "salesman": salesman,
            "saleType": sale_type
        },
        "items": items
    }


def _normalize_phone_for_whatsapp(phone: str) -> str:
    # Remover todo lo que no sea dígito
    digits = re.sub(r"\D", "", phone)
    if not digits:
        return ""
    if len(digits) == 11 and digits.startswith("0"):
        digits = "58" + digits[1:]
    elif len(digits) == 10 and not digits.startswith("58"):
        digits = "58" + digits
    return digits


async def _generate_document_from_parsed_data(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    doc_data: dict,
) -> None:
    msg = update.effective_message
    if not msg:
        return
        
    doc_type = doc_data["docType"]
    doc_num = doc_data["docNumber"]
    title_up = "COTIZACIÓN" if doc_type == "cotizacion" else "NOTA DE ENTREGA"
    emoji = "📋" if doc_type == "cotizacion" else "📦"
    
    try:
        generados_dir = Path(__file__).resolve().parent / "modulo_cotizaciones" / "generados"
        generados_dir.mkdir(parents=True, exist_ok=True)
        
        # Generar el PDF oficial y estético en segundo plano automáticamente
        from .pdf_generator import generate_document_pdf
        pdf_filename = f"{title_up}_{doc_num}.pdf"
        pdf_output_path = generados_dir / pdf_filename
        generate_document_pdf(doc_data, pdf_output_path)
        
        # Calcular montos para el mensaje pre-completado de WhatsApp
        client_name = doc_data['client'].get('name', 'Cliente')
        total_amount = 0.0
        for it in doc_data['items']:
            qty = float(it.get('qty', 1.0))
            price = float(it.get('priceUsd', 0.0))
            total_amount += (qty * price)
            
        currency = doc_data.get("currency", "usd")
        rate = float(doc_data.get("exchangeRate", get_current_bcv_rate())) if currency == "ves" else 1.0
        total_conv = total_amount * rate
        
        formatted_total = f"{total_conv:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        symbol = "$" if currency == "usd" else "Bs."
        
        # Construir mensaje de WhatsApp
        import urllib.parse
        wa_msg = (
            f"Estimado/a *{client_name}*,\n\n"
            f"Le adjunto su *{title_up} Nro {doc_num}* por un monto total de *{symbol} {formatted_total}*.\n\n"
            f"Quedo a su entera disposición.\n\n"
            f"Atentamente,\n"
            f"*FREDDY LOPEZ* (SUFEVICA)"
        )
        encoded_text = urllib.parse.quote(wa_msg)
        normalized_phone = _normalize_phone_for_whatsapp(doc_data['client'].get('phone', ''))
        
        if normalized_phone:
            wa_url = f"https://api.whatsapp.com/send?phone={normalized_phone}&text={encoded_text}"
        else:
            wa_url = f"https://api.whatsapp.com/send?text={encoded_text}"
            
        # Almacenar en user_data para el flujo de envío de correo
        context.user_data["share_doc"] = {
            "pdf_path": str(pdf_output_path),
            "pdf_filename": pdf_filename,
            "title": title_up,
            "doc_number": doc_num,
            "client_name": client_name,
            "total_amount": f"{symbol} {formatted_total}",
            "awaiting": None
        }
        
        # Crear los botones de compartir
        kb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("🟢 Enviar por WhatsApp", url=wa_url),
                InlineKeyboardButton("📧 Enviar por Correo", callback_data="share_email")
            ]
        ])
        
        text = (
            f"✅ *{title_up} GENERADA CON ÉXITO* {emoji}\n\n"
            f"👤 *Cliente:* {doc_data['client']['name']}\n"
            f"🆔 *RIF:* {doc_data['client']['rif']}\n"
            f"🔢 *Número:* {doc_num}\n"
            f"🛒 *Productos:* {len(doc_data['items'])} ítems cargados\n\n"
            f"🔗 *Enlace en tu Servidor PC (Local):*\n"
            f"• *PDF:* `file:///c:/Users/Freddy%20Lopez/Documents/Telegram%20bot/bot_financiero_telegram/modulo_cotizaciones/generados/{pdf_filename}`\n\n"
            f"💡 *¿Cómo compartir el archivo PDF por WhatsApp?*\n"
            f"Debido a limitaciones de WhatsApp, los enlaces web no pueden adjuntar archivos locales. Para que el PDF le llegue al cliente por WhatsApp:\n"
            f"1. Mantén presionado el archivo PDF que el bot envió arriba (o haz clic derecho en PC).\n"
            f"2. Selecciona *Reenviar* o *Compartir* y elije *WhatsApp* para enviárselo directamente.\n"
            f"_(El botón verde de abajo te ayuda abriendo el chat de WhatsApp con el saludo de texto pre-escrito)._\n\n"
            f"📧 También puedes enviarlo automáticamente por Correo SMTP presionando el botón de abajo.\n\n"
            f"👇 *Elige una opción:*"
        )
        
        await msg.reply_text(text, parse_mode="Markdown")
        
        # Enviar PDF oficial (diseño corporativo premium) con los botones de compartir adjuntos
        await msg.reply_document(
            document=str(pdf_output_path),
            filename=pdf_filename,
            caption=f"📄 PDF Oficial {title_up} Nro {doc_num} - SUFEVICA",
            reply_markup=kb
        )
        
    except Exception as e:
        logger.exception("Error al generar el documento desde texto")
        await msg.reply_text(f"❌ Ocurrió un error al generar el documento: {e!s}")


async def _process_intent(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    text: str,
) -> None:
    msg = update.effective_message
    if not msg:
        return
    is_channel = update.channel_post is not None or update.edited_channel_post is not None
    emit_docs = _parse_emitir_retencion_request(text)
    if emit_docs is not None:
        await _start_emitir_retencion_flow(update, context, emit_docs)
        return
    ret_data = _parse_retencion_entry_request(text)
    # Registrar automáticamente si el texto ya contiene los campos en cualquier chat (privado o grupal).
    if ret_data is None:
        ret_data = _parse_retencion_from_any_text(text)
    if ret_data is not None:
        try:
            inserted = excel_store.append_record(
                config.EXCEL_PATH,
                fecha_emision=ret_data["fecha_emision"],
                numero_comprobante=ret_data["numero_comprobante"],
                rif=ret_data["rif"],
                fechas_facturas=ret_data["fechas_facturas"],
                numeros_facturas=ret_data["numeros_facturas"],
                controles_facturas=ret_data["controles_facturas"],
                total_compra_con_iva=ret_data["total_compra_con_iva"],
                base_imponible=ret_data["base_imponible"],
                iva_retenido=ret_data["iva_retenido"],
                ocr_snippet=(
                    "Comprobante de retencion (registro manual texto/voz). "
                    f"{text[:420]}"
                ),
            )
        except ValueError as e:
            await msg.reply_text(
                "No pude guardar en Excel por estructura incompatible.\n"
                f"Detalle: {e!s}"
            )
            return
        if not inserted:
            await _notify_same_source_channel(
                update,
                context,
                "Esa retencion parece duplicada; no se agregó de nuevo.",
            )
            return
        await _notify_same_source_channel(
            update,
            context,
            f"✅ Datos registrados correctamente en {config.EXCEL_PATH.name}.",
        )
        return

    # Intentar primero el nuevo formato de Reporte Z
    z_nuevo = _parse_reporte_z_nuevo(text)
    if z_nuevo is not None:
        try:
            inserted = excel_store.append_reporte_z_nuevo(
                config.REPORTES_Z_PATH,
                numero_reporte=z_nuevo["numero_reporte"],
                fecha_emision=z_nuevo["fecha_emision"],
                sub_total=z_nuevo["sub_total"],
                base_imponible=z_nuevo["base_imponible"],
                monto_exento=z_nuevo["monto_exento"],
                iva=z_nuevo["iva"],
                total=z_nuevo["total"],
                texto_origen=text,
            )
        except Exception as e:
            await msg.reply_text(
                "No pude registrar el nuevo reporte Z en Excel.\n"
                f"Detalle: {e!s}"
            )
            return
            
        if not inserted:
            await _notify_same_source_channel(
                update,
                context,
                f"⚠️ El Reporte Z Nro {z_nuevo['numero_reporte']} ya se encuentra registrado en el sistema; no se agregó de nuevo.",
            )
            return

        await _notify_same_source_channel(
            update,
            context,
            f"✅ Reporte Z Nro {z_nuevo['numero_reporte']} registrado correctamente en {config.REPORTES_Z_PATH.name}.",
        )
        return

    v_data = _parse_venta_o_reportez(text)
    if v_data is not None:
        try:
            b_dec = excel_store.parse_amount_ves_string(v_data["base_imponible"])
            i_dec = excel_store.parse_amount_ves_string(v_data["iva"])
            t_dec = excel_store.parse_amount_ves_string(v_data["total"])
            
            from decimal import Decimal
            if b_dec is not None and i_dec is None:
                i_dec = (b_dec * Decimal("0.16")).quantize(Decimal("0.01"))
            if b_dec is not None and t_dec is None:
                t_dec = b_dec + (i_dec or Decimal("0"))
            if t_dec is not None and b_dec is None:
                b_dec = (t_dec / Decimal("1.16")).quantize(Decimal("0.01"))
                i_dec = t_dec - b_dec
                
            base_str = str(b_dec) if b_dec is not None else v_data["base_imponible"]
            iva_str = str(i_dec) if i_dec is not None else v_data["iva"]
            tot_str = str(t_dec) if t_dec is not None else v_data["total"]
            
            if v_data["clasificacion"] == "Reporte Z":
                inserted = excel_store.append_reporte_z_nuevo(
                    config.REPORTES_Z_PATH,
                    numero_reporte=v_data["numero_documento"],
                    fecha_emision=v_data["fecha"],
                    sub_total=base_str,
                    base_imponible=base_str,
                    monto_exento="0,00",
                    iva=iva_str,
                    total=tot_str,
                    texto_origen=text,
                )
            else:
                inserted = excel_store.append_venta_record(
                    config.FACTURAS_EMITIDAS_PATH,
                    clasificacion=v_data["clasificacion"],
                    estado=v_data["estado"],
                    fecha=v_data["fecha"],
                    numero_documento=v_data["numero_documento"],
                    razon_social=v_data["razon_social"],
                    rif=v_data["rif"],
                    base_imponible=base_str,
                    iva=iva_str,
                    total=tot_str,
                    texto_origen=text,
                )
        except Exception as e:
            await msg.reply_text(
                "No pude registrar la venta en Excel.\n"
                f"Detalle: {e!s}"
            )
            return
            
        if not inserted:
            await _notify_same_source_channel(
                update,
                context,
                f"⚠️ Esa {v_data['clasificacion']} Nro {v_data['numero_documento']} ya se encuentra registrada; no se agregó de nuevo.",
            )
            return
            
        path_name = config.REPORTES_Z_PATH.name if v_data["clasificacion"] == "Reporte Z" else config.FACTURAS_EMITIDAS_PATH.name
        await _notify_same_source_channel(
            update,
            context,
            f"✅ {v_data['clasificacion']} Nro {v_data['numero_documento']} registrado correctamente en {path_name}.",
        )
        return

    fc = parse_factura_compra_text(text)
    if fc is not None:
        is_sale = False
        if fc.proveedor_rif:
            clean_rif = re.sub(r"\D", "", str(fc.proveedor_rif))
            if "40194130" in clean_rif:
                is_sale = True
        if not is_sale and fc.proveedor and "SUFEVICA" in str(fc.proveedor).upper():
            is_sale = True

        if is_sale:
            # Factura emitida por SUFEVICA -> Venta
            try:
                inserted = excel_store.append_venta_record(
                    config.FACTURAS_EMITIDAS_PATH,
                    clasificacion="Factura Emitida",
                    estado="REGISTRADO",
                    fecha=fc.fecha_emision,
                    numero_documento=fc.numero_documento,
                    razon_social=fc.receptor,
                    rif=fc.receptor_rif,
                    base_imponible=fc.base_imponible or fc.subtotal,
                    iva=fc.monto_iva,
                    total=fc.total,
                    texto_origen=text,
                )
            except Exception as e:
                await msg.reply_text(
                    f"No pude guardar la factura de venta en Excel: {e!s}"
                )
                return
                
            if not inserted:
                await _notify_same_source_channel(
                    update,
                    context,
                    f"⚠️ La factura de venta Nro {fc.numero_documento} ya se encuentra registrada; no se agregó de nuevo.",
                )
                return
            await _notify_same_source_channel(
                update,
                context,
                f"✅ Factura de venta Nro {fc.numero_documento or '—'} registrada correctamente en {config.FACTURAS_EMITIDAS_PATH.name}.",
            )
            return
        else:
            # Factura recibida de proveedor -> Compra
            warn = ""
            try:
                base = excel_store.parse_amount_ves_string(fc.base_imponible or fc.subtotal)
                exento = excel_store.parse_amount_ves_string(fc.monto_exento) or Decimal("0")
                iva = excel_store.parse_amount_ves_string(fc.monto_iva)
                total = excel_store.parse_amount_ves_string(fc.total)
                if base is not None and iva is not None and total is not None:
                    diff = (base + exento + iva - total).copy_abs()
                    if diff > Decimal("0.02"):
                        warn = (
                            f" [ADVERTENCIA: inconsistencia matemática "
                            f"base({base})+exento({exento})+iva({iva})!=total({total})]"
                        )
            except Exception:
                warn = ""
            try:
                inserted = excel_store.append_factura_compra(
                    config.FACTURAS_RECIBIDAS_PATH,
                    tipo_documento=fc.tipo_documento,
                    fecha_emision=fc.fecha_emision,
                    fecha_vencimiento=fc.fecha_vencimiento,
                    numero_documento=fc.numero_documento,
                    numero_control=fc.numero_control,
                    proveedor=fc.proveedor,
                    proveedor_rif=fc.proveedor_rif,
                    proveedor_telefono=fc.proveedor_telefono,
                    direccion_fiscal_proveedor=fc.direccion_fiscal_proveedor,
                    receptor=fc.receptor,
                    receptor_rif=fc.receptor_rif,
                    subtotal=fc.subtotal,
                    monto_exento=fc.monto_exento,
                    base_imponible=fc.base_imponible,
                    monto_iva=fc.monto_iva,
                    total=fc.total,
                    texto_resumen=(text[:420] + warn)[:500],
                )
            except ValueError as e:
                await msg.reply_text(
                    f"No pude guardar la factura de compra: {e!s}"
                )
                return
                
            if not inserted:
                await _notify_same_source_channel(
                    update,
                    context,
                    f"⚠️ La factura recibida Nro {fc.numero_documento} del proveedor {fc.proveedor or ''} ({fc.proveedor_rif or ''}) ya se encuentra registrada; no se agregó de nuevo.",
                )
                return
                
            await _notify_same_source_channel(
                update,
                context,
                "✅ Datos registrados correctamente en "
                f"{config.FACTURAS_RECIBIDAS_PATH.name} (Doc {fc.numero_documento or '—'}).",
            )
            return

    intent = _match_intent(text)
    if intent == "send_facturas_compra_excel":
        excel_store.ensure_factura_compra_workbook(config.FACTURAS_RECIBIDAS_PATH)
        await msg.reply_document(
            document=str(config.FACTURAS_RECIBIDAS_PATH),
            filename=config.FACTURAS_RECIBIDAS_PATH.name,
            caption="Facturas de compra / recibidas (Subtotal, IVA, Total, etc.).",
        )
        return
    if intent == "send_excel":
        if not config.EXCEL_PATH.exists():
            await msg.reply_text(
                "Todavía no existe el archivo Excel. Registra una retención primero "
                "o revisa EXCEL_PATH en .env."
            )
            return
        await msg.reply_document(
            document=str(config.EXCEL_PATH),
            filename=config.EXCEL_PATH.name,
            caption="consolidado_financiero.xlsx (retenciones; no es el de facturas compra).",
        )
        return
    if intent == "summary_today":
        today = date.today()
        n, total = excel_store.summary_for_date(config.EXCEL_PATH, today)
        await msg.reply_text(
            f"Resumen del {today.strftime('%d/%m/%Y')}: {n} registro(s). "
            f"Suma de montos: {total}"
        )
        return
    if intent == "tributos_report":
        today = date.today()
        fortnight = 1 if today.day <= 15 else 2
        report = tributario_engine.get_compromiso_tributario_report(today.year, today.month, fortnight)
        text = format_tributos_report(report)
        kb = _tributos_keyboard(today.year, today.month, fortnight, _generate_short_summary(report))
        await msg.reply_text(text, reply_markup=kb, parse_mode="Markdown")
        return
    if intent == "retenciones_report":
        parsed = _parse_retenciones_report_request(text)
        if parsed is None:
            await msg.reply_text(
                "Formato no reconocido. Usa por ejemplo: "
                "«retenciones recibidas del 01/05/2026 al 31/05/2026 en excel» "
                "o «... en pdf»."
            )
            return
        date_from, date_to, out_fmt = parsed
        records = excel_store.retenciones_by_document_date(
            config.EXCEL_PATH,
            date_from=date_from,
            date_to=date_to,
        )
        if not records:
            await msg.reply_text(
                "No encontré retenciones en ese rango por fecha del documento."
            )
            return
        suffix = ".xlsx" if out_fmt == "excel" else ".pdf"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            out_path = Path(tmp.name)
        try:
            if out_fmt == "excel":
                excel_store.export_retenciones_excel(records, out_path)
            else:
                excel_store.export_retenciones_pdf(
                    records,
                    out_path,
                    date_from=date_from,
                    date_to=date_to,
                )
            if out_fmt == "excel":
                filename = "RETEN-REC.xlsx"
            else:
                report_name = (
                    "detallado_de_retenciones_recibidas_de_clientes_del_"
                    f"{date_from.strftime('%d-%m-%Y')}_al_{date_to.strftime('%d-%m-%Y')}"
                )
                filename = f"{report_name}{suffix}"
            await msg.reply_document(
                document=str(out_path),
                filename=filename,
                caption=(
                    "Reporte generado: "
                    "detallado de retenciones recibidas de clientes del "
                    f"{date_from.strftime('%d/%m/%Y')} al {date_to.strftime('%d/%m/%Y')} "
                    "(filtrado por fecha del documento)."
                ),
            )
        finally:
            out_path.unlink(missing_ok=True)
        return
    await msg.reply_text(
        "No reconoci la orden.\n"
        "Ejemplos:\n"
        "1) «resumen de hoy»\n"
        "2) «enviar excel» (retenciones consolidado) o «enviar excel facturas compra»\n"
        "3) «retenciones recibidas del 01/05/2026 al 31/05/2026 en excel»\n"
        "4) «registrar retencion, fecha:24/04/2026, comprobante:20260400000381, "
        "rif:J-41278020-4, fecha factura:23/04/2026, nro factura:00007553, "
        "control:Z7C7018762, total:1.400,00, base imponible:1.206,90, iva retenido:144,83»\n"
        "5) Pega texto de factura de compra / factura recibida (Febeca, totales en Bs, etc.).\n"
        "6) Emitir retencion de facturas:06949950|06949951"
    )


async def _emitir_retencion_generate(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    msg = update.effective_message
    pending = context.user_data.get("pending_emit_ret")
    if not msg or not pending:
        return
    docs = list(pending.get("docs", []))
    items = excel_store.load_facturas_by_document_numbers(config.FACTURAS_RECIBIDAS_PATH, docs)
    if not items:
        await msg.reply_text("No encontré las facturas solicitadas al momento de emitir.")
        context.user_data.pop("pending_emit_ret", None)
        return
    emission_date_str = str(pending.get("emission_date") or "").strip()
    emission_date = _parse_user_date(emission_date_str) if emission_date_str else None
    if emission_date is None:
        emission_date = date.today()
        emission_date_str = emission_date.strftime("%d/%m/%Y")
    monthly_path = _reten_emit_monthly_path(emission_date)
    seq_mode = str(pending.get("seq_mode") or "auto")
    if seq_mode == "manual":
        num_comp = str(pending.get("manual_num") or "").strip()
        if not (len(num_comp) == 14 and num_comp.isdigit()):
            await msg.reply_text(
                "Número manual inválido. Debe tener 14 dígitos (YYYYMM + 8 secuencial)."
            )
            return
    else:
        num_comp = excel_store.next_retencion_emitida_number(
            monthly_path,
            emission_date=emission_date,
        )
        logger.info(
            "Correlativo asignado desde Excel %s: %s",
            monthly_path.parent,
            num_comp,
        )
    out_fmt = str(pending.get("format") or "pdf").strip().lower()
    provider = str(pending.get("provider") or "").strip()
    provider_rif = str(pending.get("provider_rif") or "").strip()
    provider_phone = str(pending.get("provider_phone") or "").strip()
    provider_address = str(pending.get("provider_address") or "").strip()
    base_total, iva_total, retenido = _totals_for_items(items)
    excel_store.append_retencion_emitida(
        monthly_path,
        numero_comprobante=num_comp,
        fecha_emision=emission_date_str,
        periodo_fiscal=_periodo_fiscal(emission_date),
        proveedor=provider,
        proveedor_rif=provider_rif,
        direccion_fiscal_prov=provider_address,
        documentos="|".join(docs),
        controles="|".join(str(it.numero_control or "") for it in items),
        base_imponible_total=base_total,
        iva_total=iva_total,
        porcentaje_retencion=RETENTION_RATE,
        iva_retenido_total=retenido,
        formato_salida=out_fmt,
    )
    suffix = ".xlsx" if out_fmt == "excel" else ".pdf"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        out_path = Path(tmp.name)
    try:
        if out_fmt == "excel":
            excel_store.export_comprobante_emitido_excel(
                out_path=out_path,
                numero_comprobante=num_comp,
                fecha_emision=emission_date_str,
                periodo_fiscal=_periodo_fiscal(emission_date),
                proveedor=provider,
                proveedor_rif=provider_rif,
                proveedor_telefono=provider_phone,
                direccion_fiscal_prov=provider_address,
                items=items,
                porcentaje_retencion=RETENTION_RATE,
            )
        else:
            excel_store.export_comprobante_emitido_pdf(
                out_path=out_path,
                numero_comprobante=num_comp,
                fecha_emision=emission_date_str,
                periodo_fiscal=_periodo_fiscal(emission_date),
                proveedor=provider,
                proveedor_rif=provider_rif,
                proveedor_telefono=provider_phone,
                direccion_fiscal_prov=provider_address,
                items=items,
                porcentaje_retencion=RETENTION_RATE,
                firma_sello_path=config.FIRMA_SELLO_PATH,
            )
        await msg.reply_document(
            document=str(out_path),
            filename=f"COMPROBANTE-RET-{num_comp}{suffix}",
            caption=(
                f"✅ Comprobante emitido ({num_comp}) y guardado en {monthly_path.name}."
            ),
        )
    finally:
        out_path.unlink(missing_ok=True)
    context.user_data.pop("pending_emit_ret", None)


async def handle_emit_retention_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    q = update.callback_query
    if not q:
        return
    await q.answer()
    data = (q.data or "").strip()
    msg = q.message
    if not msg:
        return
    pending = context.user_data.get("pending_emit_ret")
    if not pending:
        await msg.reply_text("No hay una emisión pendiente. Envía el comando nuevamente.")
        return
    if data == "emit_seq_yes":
        pending["seq_mode"] = "auto"
        emission_date = _parse_user_date(str(pending.get("emission_date") or ""))
        if emission_date is None:
            emission_date = date.today()
        monthly_path = _reten_emit_monthly_path(emission_date)
        next_num = excel_store.next_retencion_emitida_number(
            monthly_path,
            emission_date=emission_date,
        )
        kb = InlineKeyboardMarkup(
            [[
                InlineKeyboardButton("Fecha de hoy", callback_data="emit_date_today"),
                InlineKeyboardButton("Fecha manual", callback_data="emit_date_manual"),
            ]]
        )
        await msg.reply_text(
            f"Secuencia automática desde Excel. Correlativo a asignar: {next_num}\n"
            "¿Qué fecha de emisión deseas?",
            reply_markup=kb,
        )
        return
    if data == "emit_seq_no":
        pending["seq_mode"] = "manual"
        pending["awaiting"] = "manual_number"
        await msg.reply_text(
            "Envía el número de comprobante manual (14 dígitos, formato YYYYMM########)."
        )
        return
    if data == "emit_date_today":
        pending["emission_date"] = date.today().strftime("%d/%m/%Y")
        kb = InlineKeyboardMarkup(
            [[
                InlineKeyboardButton("PDF", callback_data="emit_fmt_pdf"),
                InlineKeyboardButton("Excel", callback_data="emit_fmt_excel"),
            ]]
        )
        await msg.reply_text("Formato del comprobante a generar:", reply_markup=kb)
        return
    if data == "emit_date_manual":
        pending["awaiting"] = "manual_date"
        await msg.reply_text("Envía la fecha de emisión en formato DD/MM/AAAA.")
        return
    if data in ("emit_fmt_pdf", "emit_fmt_excel"):
        pending["format"] = "pdf" if data.endswith("pdf") else "excel"
        await _emitir_retencion_generate(update, context)
        return


async def mi_id_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Para comprobar el id que Telegram usa (cualquier usuario)."""
    del context
    msg = update.effective_message
    u = update.effective_user
    if not msg or not u:
        return
    if _allowed(update):
        await msg.reply_text(
            f"Estas autorizado. Tu id: {u.id}. TELEGRAM_ALLOWED_USER_ID en .env debe ser ese número."
        )
    else:
        await msg.reply_text(
            f"Tu id de Telegram es: {u.id}\n"
            "Ponlo en TELEGRAM_ALLOWED_USER_ID del archivo .env del bot y reinicia."
        )


async def start_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _allowed(update):
        await _deny(update)
        return
    context.user_data["voice_mode"] = False
    if update.message:
        await update.message.reply_text(
            "Bot financiero listo.\n"
            "• /tributos: Consulta interactiva en tiempo real del IVA por pagar, retenciones y anticipos de la quincena.\n"
            "• /cotizacion: Abre el Módulo de Cotizaciones/Presupuestos en modo Cotización.\n"
            "• /nota: Abre el Módulo de Notas de Entrega/Despachos en modo Nota de Entrega.\n"
            "• Texto o nota de voz: reconozco ordenes y registro en Excel.\n"
            "• Botón: pulsa «Activar comando de voz» y luego envía tu nota de voz.\n"
            "• Escribe: «resumen de hoy», «enviar excel» (retenciones), "
            "«enviar excel facturas compra»,\n"
            "  «retenciones recibidas del 01/05/2026 al 31/05/2026 en pdf/excel»,\n"
            "  «Emitir retencion de facturas:06949950|06949951»,\n"
            "  «tributos» o «iva por pagar»,\n"
            "  o «registrar retencion, fecha:..., comprobante:..., rif:..., iva retenido:...».\n"
            "• Registrar Ventas o Reportes Z: Escribe o di por voz:\n"
            "  «registrar venta, factura:12, fecha:25/05/2026, cliente:ABC, rif:J-12345678-9, base:1000, iva:160»\n"
            "  «reporte z, fecha:26/05/2026, numero:0023, base:5000»\n"
            "• Facturas de compra (proveedor -> SUFEVICA): pegar el texto; se guardan en "
            "FACTURAS-RECIBIDAS-NUEVO.xlsx (Subtotal, IVA, Total en columnas propias).\n"
            "• Si ves «Acceso no autorizado», envia /mi_id y revisa .env.",
            reply_markup=_main_keyboard(),
        )


async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _allowed(update):
        await _deny(update)
        return
    context.user_data.setdefault("voice_mode", False)
    if not update.message or not update.message.voice:
        return
    voice = update.message.voice
    tg_file = await context.bot.get_file(voice.file_id)
    suffix = ".ogg"
    if tg_file.file_path:
        sfx = Path(tg_file.file_path).suffix
        if sfx:
            suffix = sfx
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        out_path = Path(tmp.name)
    try:
        await tg_file.download_to_drive(out_path)
        text = await asyncio.to_thread(transcribe_audio_file, out_path)
    except Exception as e:  # noqa: BLE001
        logger.exception("Transcripción fallida")
        await update.message.reply_text(
            "No pude transcribir el audio. Si no usas OpenAI, instala ffmpeg y configura "
            "OPENAI_API_KEY para Whisper, o revisa el micrófono/formato.\n"
            f"Detalle: {e!s}"
        )
        return
    finally:
        out_path.unlink(missing_ok=True)
    if update.message:
        await update.message.reply_text(f"Transcripción: «{text}»")
    await _process_intent(update, context, text)
    if context.user_data.get("voice_mode"):
        context.user_data["voice_mode"] = False
        if update.message:
            await update.message.reply_text(
                "Comando de voz procesado. Si quieres otro, pulsa de nuevo «Activar comando de voz».",
                reply_markup=_main_keyboard(),
            )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if not msg:
        return
    text = (msg.text or msg.caption or "").strip()
    if not text:
        return
    is_channel = update.channel_post is not None or update.edited_channel_post is not None

    # Procesar siempre publicaciones de canal y además chat SUFEVICA detectado.
    if is_channel or _is_sufevica_chat(update):
        logger.info("Procesando texto de canal/chat detectado (chat_id=%s).", msg.chat_id)
        await _process_intent(update, context, text)
        return

    # Mensajes privados: aplicar restricción por usuario.
    if not _allowed(update):
        await _deny(update)
        return

    context.user_data.setdefault("voice_mode", False)
    
    share_doc = context.user_data.get("share_doc")
    if share_doc and share_doc.get("awaiting") == "client_email":
        email_input = text.strip()
        if email_input.lower() in ("cancelar", "cancel"):
            context.user_data.pop("share_doc", None)
            await msg.reply_text("❌ Envío de correo cancelado.")
            return
            
        email_regex = r"^[\w\.-]+@[\w\.-]+\.\w+$"
        if not re.match(email_regex, email_input):
            await msg.reply_text(
                "❌ Dirección de correo electrónico inválida.\n"
                "Por favor, escribe una dirección de correo válida (ej. cliente@correo.com) o escribe `cancelar`:"
            )
            return
            
        share_doc["awaiting"] = None
        status_msg = await msg.reply_text(
            f"⏳ *Enviando {share_doc['title']} por correo SMTP a `{email_input}`...*",
            parse_mode="Markdown"
        )
        
        asyncio.create_task(_send_document_email_async(update, context, email_input, status_msg))
        return
        
    pending_doc = context.user_data.get("pending_doc")
    if pending_doc:
        state = pending_doc.get("awaiting")
        if state == "text_data":
            doc_type = pending_doc["type"]
            doc_data = _parse_document_text_explicit(text, doc_type)
            if doc_data is not None:
                # Eliminar el mensaje de prompt inicial
                start_prompt_id = pending_doc.pop("start_prompt_message_id", None)
                if start_prompt_id:
                    try:
                        await context.bot.delete_message(chat_id=msg.chat_id, message_id=start_prompt_id)
                    except Exception:
                        pass
                
                # Eliminar el mensaje del usuario con el texto pegado
                try:
                    await msg.delete()
                except Exception:
                    pass
                
                pending_doc["parsed_data"] = doc_data
                pending_doc["awaiting"] = "edit_card"
                await _send_client_data_card(update, context, first_time=True)
                return
            else:
                await msg.reply_text(
                    "⚠️ No pude identificar productos en el texto pegado.\n\n"
                    "Por favor, asegúrate de incluir cantidades y precios (ej: `10 x PARP-220 Perno a 2.50`), o escríbelo en columnas limpias.\n\n"
                    "Vuelve a intentarlo pegando el texto aquí, o envía `/start` para cancelar.",
                    parse_mode="Markdown"
                )
                return
                
        elif state in ("input_name", "input_rif", "input_address", "input_phone", "input_salesman", "input_note", "input_rate"):
            doc_data = pending_doc["parsed_data"]
            client = doc_data["client"]
            
            val = text.strip()
            
            if state == "input_rate":
                try:
                    rate_val = float(val.replace(",", "."))
                    if rate_val <= 0:
                        raise ValueError()
                    doc_data["exchangeRate"] = rate_val
                except ValueError:
                    await msg.reply_text("⚠️ Por favor ingresa una tasa de cambio válida mayor a cero (ej. 40.50 o 40,50).")
                    return
            else:
                # Si el usuario responde omitir, vacio, ninguno, etc., dejarlo vacío (se omitirá el dato en el PDF/HTML)
                if val.lower() in ("omitir", "ninguno", "vacio", "vacío", "omitido", "cancelar"):
                    val = ""
                    
                if state == "input_name":
                    client["name"] = val
                elif state == "input_rif":
                    client["rif"] = _normalize_rif(val) if val else ""
                elif state == "input_address":
                    client["address"] = val
                elif state == "input_phone":
                    client["phone"] = val
                elif state == "input_salesman":
                    client["salesman"] = val
                elif state == "input_note":
                    client["note"] = val
                
            # Eliminar el mensaje de prompt anterior
            prompt_id = pending_doc.pop("prompt_message_id", None)
            if prompt_id:
                try:
                    await context.bot.delete_message(chat_id=msg.chat_id, message_id=prompt_id)
                except Exception:
                    pass
            
            # Eliminar el mensaje de texto enviado por el usuario para mantener limpio el chat
            try:
                await msg.delete()
            except Exception:
                pass
                
            pending_doc["awaiting"] = "edit_card"
            await _send_client_data_card(update, context, first_time=False)
            return

    pending_email = context.user_data.get("pending_email_report")
    if pending_email and pending_email.get("awaiting") == "email_address":
        email_input = text.strip()
        if email_input.lower() in ("cancelar", "cancel"):
            context.user_data.pop("pending_email_report", None)
            await msg.reply_text("❌ Envío de reportes por correo cancelado.")
            return
            
        if email_input.lower() == "contador":
            if config.DEFAULT_ACCOUNTANT_EMAIL:
                email_input = config.DEFAULT_ACCOUNTANT_EMAIL
            else:
                await msg.reply_text(
                    "⚠️ No has configurado `DEFAULT_ACCOUNTANT_EMAIL` en tu archivo `.env`.\n"
                    "Por favor, escribe una dirección de correo válida para continuar:"
                )
                return
                
        email_regex = r"^[\w\.-]+@[\w\.-]+\.\w+$"
        if not re.match(email_regex, email_input):
            await msg.reply_text(
                "❌ Dirección de correo electrónico inválida.\n"
                "Por favor, escribe una dirección de correo válida (ej. nombre@servidor.com):"
            )
            return
            
        y = pending_email["year"]
        m = pending_email["month"]
        f = pending_email["fortnight"]
        context.user_data.pop("pending_email_report", None)
        
        status_msg = await msg.reply_text(
            f"⏳ *Generando reportes y conectando al servidor SMTP...*\n"
            f"Despachando correo a: `{email_input}`",
            parse_mode="Markdown"
        )
        
        asyncio.create_task(_send_reportes_async_workflow(update, context, y, m, f, email_input, status_msg))
        return

    pending = context.user_data.get("pending_emit_ret")
    if pending and pending.get("awaiting") == "manual_number":
        manual_num = re.sub(r"\D", "", text)
        if not (len(manual_num) == 14 and manual_num.isdigit()):
            await msg.reply_text(
                "Número inválido. Debe tener 14 dígitos (YYYYMM########)."
            )
            return
        pending["manual_num"] = manual_num
        pending["awaiting"] = None
        kb = InlineKeyboardMarkup(
            [[
                InlineKeyboardButton("Fecha de hoy", callback_data="emit_date_today"),
                InlineKeyboardButton("Fecha manual", callback_data="emit_date_manual"),
            ]]
        )
        await msg.reply_text("Número manual guardado. Ahora selecciona fecha de emisión.", reply_markup=kb)
        return
    if pending and pending.get("awaiting") == "manual_date":
        d = _parse_user_date(text)
        if d is None:
            await msg.reply_text("Fecha inválida. Usa DD/MM/AAAA.")
            return
        pending["emission_date"] = d.strftime("%d/%m/%Y")
        pending["awaiting"] = None
        kb = InlineKeyboardMarkup(
            [[
                InlineKeyboardButton("PDF", callback_data="emit_fmt_pdf"),
                InlineKeyboardButton("Excel", callback_data="emit_fmt_excel"),
            ]]
        )
        await msg.reply_text("Fecha guardada. Selecciona formato del comprobante.", reply_markup=kb)
        return

    if text == VOICE_BUTTON:
        context.user_data["voice_mode"] = True
        await msg.reply_text(
            "Modo voz activado. Envía ahora tu nota de voz con el requerimiento.",
            reply_markup=_main_keyboard(),
        )
        return
    if text == VOICE_CANCEL_BUTTON:
        context.user_data["voice_mode"] = False
        await msg.reply_text(
            "Modo voz desactivado.",
            reply_markup=_main_keyboard(),
        )
        return
    if text == COTI_BUTTON:
        await _start_document_flow(update, context, "cotizacion")
        return
    if text == NOTA_BUTTON:
        await _start_document_flow(update, context, "nota")
        return
    if context.user_data.get("voice_mode"):
        await msg.reply_text(
            "Modo voz activo: envía una nota de voz o pulsa «Cancelar voz».",
            reply_markup=_main_keyboard(),
        )
        return
    await _process_intent(update, context, text)


async def _send_reportes_async_workflow(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    year: int,
    month: int,
    fortnight: int,
    recipient_email: str,
    status_msg,
) -> None:
    from .tributario_engine import get_fortnight_range, get_compromiso_tributario_report
    from .email_sender import send_report_email
    import tempfile
    
    start_date, end_date = get_fortnight_range(year, month, fortnight)
    
    # 1. Generar datos y reportes quincenales
    records = excel_store.retenciones_by_document_date(
        config.EXCEL_PATH,
        date_from=start_date,
        date_to=end_date,
    )
    
    # Archivos temporales para adjuntar
    temp_files: list[Path] = []
    attachments: list[Path] = []
    
    try:

        month_names = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
        }
        period_str = f"Q{fortnight} - {month_names.get(month, str(month))} / {year}"

        # 1. Excel de Retenciones Recibidas
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            xls_path = Path(tmp.name)
            temp_files.append(xls_path)
        ret_headers = ["#", "Fecha Emisión", "Nro Comprobante", "Cliente RIF", "Factura Afectada", "Control Factura", "Base (Bs)", "IVA Retenido (Bs)"]
        ret_rows = []
        for idx, rec in enumerate(records):
            ret_rows.append([
                idx + 1,
                rec.fecha_emision.strftime("%d/%m/%Y"),
                rec.numero_comprobante,
                rec.rif,
                rec.numeros_facturas,
                rec.controles_facturas,
                float(rec.base_imponible or 0),
                float(rec.iva_retenido)
            ])
        excel_store.generate_premium_report_excel(
            xls_path,
            title="Resumen Quincenal de Retenciones de IVA Recibidas",
            period_str=period_str,
            headers=ret_headers,
            rows=ret_rows,
            numeric_cols=[6, 7],
            sum_cols=[6, 7]
        )
        prof_xls = xls_path.with_name(f"RETENCIONES-RECIBIDAS-Q{fortnight}-{month}-{year}.xlsx")
        xls_path.rename(prof_xls)
        attachments.append(prof_xls)
        temp_files.append(prof_xls)
        
        # 2. Excel de Facturas Recibidas / Compras con Retención Emitida
        purchases_rows = excel_store.load_purchases_by_date_range(
            config.RETENCIONES_EMITIDAS_DIR,
            date_from=start_date,
            date_to=end_date
        )
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            purchases_xls_path = Path(tmp.name)
            temp_files.append(purchases_xls_path)
        excel_store.generate_premium_report_excel(
            purchases_xls_path,
            title="Resumen Quincenal de Facturas Recibidas (Compras)",
            period_str=period_str,
            headers=["#", "Fecha", "Correlativo", "Proveedor", "RIF", "Nro Factura", "Nro Control", "Base (Bs)", "IVA (Bs)", "Monto (Bs)", "Retención (Bs)"],
            rows=purchases_rows,
            numeric_cols=[7, 8, 9, 10],
            sum_cols=[7, 8, 9, 10]
        )
        prof_purchases_xls = purchases_xls_path.with_name(f"FACTURAS-RECIBIDAS-Q{fortnight}-{month}-{year}.xlsx")
        purchases_xls_path.rename(prof_purchases_xls)
        attachments.append(prof_purchases_xls)
        temp_files.append(prof_purchases_xls)

        # 3. Excel de Facturas Emitidas / Ventas
        sales_rows = excel_store.load_sales_by_date_range(
            config.FACTURAS_EMITIDAS_PATH,
            date_from=start_date,
            date_to=end_date
        )
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            sales_xls_path = Path(tmp.name)
            temp_files.append(sales_xls_path)
        excel_store.generate_premium_report_excel(
            sales_xls_path,
            title="Resumen Quincenal de Facturas Emitidas (Ventas)",
            period_str=period_str,
            headers=["#", "Fecha", "Nro Factura", "Cliente", "RIF", "Base (Bs)", "IVA (Bs)", "Monto (Bs)"],
            rows=sales_rows,
            numeric_cols=[5, 6, 7],
            sum_cols=[5, 6, 7]
        )
        prof_sales_xls = sales_xls_path.with_name(f"FACTURAS-EMITIDAS-Q{fortnight}-{month}-{year}.xlsx")
        sales_xls_path.rename(prof_sales_xls)
        attachments.append(prof_sales_xls)
        temp_files.append(prof_sales_xls)

        # 4. Excel de Reportes Z / Ventas Diarias
        z_rows = excel_store.load_reportes_z_by_date_range(
            config.REPORTES_Z_PATH,
            date_from=start_date,
            date_to=end_date
        )
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            z_xls_path = Path(tmp.name)
            temp_files.append(z_xls_path)
        excel_store.generate_premium_report_excel(
            z_xls_path,
            title="Resumen Quincenal de Cierres Z (Ventas Diarias)",
            period_str=period_str,
            headers=["#", "Fecha", "Nro Reporte Z", "Subtotal (Bs)", "Exento (Bs)", "Base (Bs)", "IVA (Bs)", "Total (Bs)"],
            rows=z_rows,
            numeric_cols=[3, 4, 5, 6, 7],
            sum_cols=[3, 4, 5, 6, 7]
        )
        prof_z_xls = z_xls_path.with_name(f"REPORTES-Z-Q{fortnight}-{month}-{year}.xlsx")
        z_xls_path.rename(prof_z_xls)
        attachments.append(prof_z_xls)
        temp_files.append(prof_z_xls)
            
        # 2. Generar cuerpo del correo
        report = get_compromiso_tributario_report(year, month, fortnight)
        report_text = format_tributos_report(report)
        
        subject = f"Reporte Tributario Quincenal SUFEVICA - Q{fortnight} {month}/{year}"
        
        body = (
            f"Estimado destinatario,\n\n"
            f"Se adjuntan los reportes financieros y las planillas fiscales de la empresa "
            f"SUMINISTROS FERRETEROS VITTORIA (SUFEVICA), C.A. para la quincena evaluada:\n\n"
            f"--------------------------------------------------\n"
            f"{report_text}\n"
            f"--------------------------------------------------\n\n"
            f"Atentamente,\n"
            f"Bot Financiero Automatizado (SUFEVICA)"
        )
        
        # 3. Despachar el correo electrónico en segundo plano de forma no bloqueante
        await asyncio.to_thread(
            send_report_email,
            recipient_email,
            subject,
            body,
            attachments,
        )
        
        # Actualizar estado en Telegram
        await status_msg.edit_text(
            f"✅ *¡Reportes enviados exitosamente!*\n\n"
            f"Se despacharon todas las planillas quincenales (PDF y Excels) de la *Quincena {fortnight}* a la dirección: `{recipient_email}`",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.exception("Error en el envío asíncrono de correo")
        await status_msg.edit_text(
            f"❌ *Error al enviar los reportes por correo SMTP*\n\n"
            f"• *Destinatario:* `{recipient_email}`\n"
            f"• *Detalle del error:* `{e!s}`\n\n"
            f"Por favor, revisa tus credenciales SMTP en el archivo `.env` o el estado de la conexión a internet.",
            parse_mode="Markdown"
        )
    finally:
        # Eliminar archivos temporales creados para evitar fugas de espacio en disco
        for path in temp_files:
            try:
                path.unlink(missing_ok=True)
            except Exception:
                pass


def _generate_short_summary(report: dict[str, object]) -> str:
    month_names = {
        1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"
    }
    y = report["year"]
    m = report["month"]
    f = report["fortnight"]
    m_name = month_names.get(m, str(m))
    
    iva = report["iva_neto_pagar_efectivo"]
    ret = report["retenciones_emitidas"]
    anticipo = report["anticipo_islr"]
    total = report["total_compromisos_a_pagar"]
    
    return (
        f"*Resumen Tributario SUFEVICA - Q{f} {m_name} {y}*\n"
        f"• IVA Neto: {excel_store._format_monto_ves(iva)} Bs\n"
        f"• Retenciones IVA: {excel_store._format_monto_ves(ret)} Bs\n"
        f"• Anticipo ISLR: {excel_store._format_monto_ves(anticipo)} Bs\n"
        f"👉 *Total a Pagar:* {excel_store._format_monto_ves(total)} Bs\n"
        f"_Generado por el Bot Financiero SUFEVICA._"
    )


def format_tributos_report(report: dict[str, object]) -> str:
    month_names = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
        7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }
    y = report["year"]
    m = report["month"]
    f = report["fortnight"]
    m_name = month_names.get(m, str(m))
    
    start_str = report["start_date"].strftime("%d/%m/%Y")
    end_str = report["end_date"].strftime("%d/%m/%Y")
    
    v_base = report["ventas_base"]
    v_iva = report["ventas_iva"]
    v_cnt = report["ventas_count"]
    
    c_base = report["compras_base"]
    c_iva = report["compras_iva"]
    c_cnt = report["compras_count"]
    
    ret_rec = report["retenciones_recibidas"]
    ret_rec_cnt = report["retenciones_recibidas_count"]
    
    ret_emi = report["retenciones_emitidas"]
    ret_emi_cnt = report["retenciones_emitidas_count"]
    
    iva_neto = report["iva_neto_pagar"]
    pago_iva = report["iva_neto_pagar_efectivo"]
    anticipo = report["anticipo_islr"]
    total = report["total_compromisos_a_pagar"]
    
    due_date_str = report["due_date"].strftime("%d/%m/%Y")
    
    text = (
        f"🏛️ *CONTROL TRIBUTARIO - SUFEVICA* 🏛️\n"
        f"📅 *Período:* {f}ra Quincena de {m_name} {y}\n"
        f"🕒 *Rango:* {start_str} al {end_str}\n"
        f"⚠️ *Límite de Pago (RIF termina en 3):* `{due_date_str}`\n\n"
        f"--------------------------------------------------\n\n"
        f"📊 *Impuesto al Valor Agregado (IVA)*:\n"
        f" 🔸 *Ventas (Débito Fiscal):* {excel_store._format_monto_ves(v_iva)} Bs ({v_cnt} doc)\n"
        f" 🔸 *Compras (Crédito Fiscal):* {excel_store._format_monto_ves(c_iva)} Bs ({c_cnt} doc)\n"
        f" 🔸 *Retenciones Recibidas:* {excel_store._format_monto_ves(ret_rec)} Bs ({ret_rec_cnt} doc)\n"
    )
    
    if iva_neto >= 0:
        text += f" 👉 *IVA Neto a Pagar:* `{excel_store._format_monto_ves(iva_neto)}` Bs\n\n"
    else:
        text += f" 👉 *Excedente a Favor:* `{excel_store._format_monto_ves(abs(iva_neto))}` Bs (Crédito Fiscal)\n\n"
        
    text += (
        f"📤 *Retenciones de IVA a Enterar (Proveedores)*:\n"
        f" 🔸 *Retenido en Compras:* `{excel_store._format_monto_ves(ret_emi)}` Bs ({ret_emi_cnt} doc)\n"
        f"    _(Este monto se paga en su totalidad al SENIAT)_\n\n"
        f"💸 *Anticipo de ISLR (1% sobre Ventas)*:\n"
        f" 🔸 *Base imponible:* {excel_store._format_monto_ves(v_base)} Bs\n"
        f" 👉 *Anticipo a pagar:* `{excel_store._format_monto_ves(anticipo)}` Bs\n\n"
        f"--------------------------------------------------\n\n"
        f"💰 *TOTAL COMPROMISOS ESTIMADOS:* `{excel_store._format_monto_ves(total)}` Bs\n"
        f"_(Monto a pagar en banco/SENIAT para este corte)_"
    )
    return text


def _tributos_keyboard(year: int, month: int, fortnight: int, report_text: str = "") -> InlineKeyboardMarkup:
    prev_y, prev_m, prev_f = year, month, fortnight
    if fortnight == 2:
        prev_f = 1
    else:
        prev_f = 2
        if month == 1:
            prev_m = 12
            prev_y -= 1
        else:
            prev_m -= 1
            
    next_y, next_m, next_f = year, month, fortnight
    if fortnight == 1:
        next_f = 2
    else:
        next_f = 1
        if month == 12:
            next_m = 1
            next_y += 1
        else:
            next_m += 1
            
    keyboard = [
        [
            InlineKeyboardButton("📊 Detalle IVA", callback_data=f"tributos_detiva_{year}_{month}_{fortnight}"),
            InlineKeyboardButton("💸 Detalle ISLR", callback_data=f"tributos_detislr_{year}_{month}_{fortnight}"),
        ],
    ]
    
    if report_text:
        import urllib.parse
        # Codificar texto para compartir
        encoded_text = urllib.parse.quote(report_text)
        wa_url = f"https://api.whatsapp.com/send?text={encoded_text}"
        
        keyboard.append([
            InlineKeyboardButton("🟢 Compartir por WhatsApp", url=wa_url),
        ])
        
    keyboard.extend([
        [
            InlineKeyboardButton("📤 Enviar Reportes por Correo (SMTP)", callback_data=f"tributos_sendemail_{year}_{month}_{fortnight}")
        ],
        [
            InlineKeyboardButton("◀️ Quincena Ant.", callback_data=f"tributos_period_{prev_y}_{prev_m}_{prev_f}"),
            InlineKeyboardButton("Quincena Sig. ▶️", callback_data=f"tributos_period_{next_y}_{next_m}_{next_f}"),
        ],
        [
            InlineKeyboardButton("📅 Seleccionar Otro Mes", callback_data="tributos_selmonth"),
        ],
    ])
    return InlineKeyboardMarkup(keyboard)


def _months_selection_keyboard() -> InlineKeyboardMarkup:
    keyboard = []
    month_names = {
        1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"
    }
    today = date.today()
    current_year = today.year
    current_month = today.month
    
    row = []
    for i in range(6):
        m = current_month - i
        y = current_year
        if m <= 0:
            m += 12
            y -= 1
        btn_text = f"{month_names[m]} {y}"
        row.append(InlineKeyboardButton(f"Q1 {btn_text}", callback_data=f"tributos_period_{y}_{m}_1"))
        row.append(InlineKeyboardButton(f"Q2 {btn_text}", callback_data=f"tributos_period_{y}_{m}_2"))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    keyboard.append([InlineKeyboardButton("❌ Volver al menú principal", callback_data=f"tributos_period_{today.year}_{today.month}_1")])
    return InlineKeyboardMarkup(keyboard)


def _format_iva_details(year: int, month: int, fortnight: int) -> str:
    start_date, end_date = tributario_engine.get_fortnight_range(year, month, fortnight)
    from .tributario_engine import _parse_row_date
    
    compras_details = []
    path_c = config.FACTURAS_RECIBIDAS_PATH
    if path_c.exists():
        try:
            wb = load_workbook(path_c, read_only=True, data_only=True)
            ws = wb.active
            headers = excel_store._headers_index(ws)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                f_emi = _parse_row_date(excel_store._cell(row, headers, "Fecha_emision", None))
                if f_emi and start_date <= f_emi <= end_date:
                    prov = str(excel_store._cell(row, headers, "Proveedor", "-"))[:15]
                    rif = str(excel_store._cell(row, headers, "Proveedor_RIF", "-"))
                    iva_val = excel_store._parse_monto_cell(excel_store._cell(row, headers, "Monto_IVA", None)) or Decimal("0")
                    compras_details.append(f"• {f_emi.strftime('%d/%m')}: {prov} ({rif}) | IVA: {excel_store._format_monto_ves(iva_val)} Bs")
            wb.close()
        except Exception as e:
            compras_details.append(f"Error cargando compras: {e}")
            
    ventas_details = []
    for path_v in [config.FACTURAS_EMITIDAS_PATH, config.REPORTES_Z_PATH]:
        if not path_v.exists():
            continue
        try:
            wb = load_workbook(path_v, read_only=True, data_only=True)
            ws = wb.active
            headers = excel_store._headers_index(ws)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                fecha_cell = excel_store._cell(row, headers, "Fecha_emision", None) or excel_store._cell(row, headers, "Fecha", None)
                f_doc = _parse_row_date(fecha_cell)
                if f_doc and start_date <= f_doc <= end_date:
                    doc = str(excel_store._cell(row, headers, "Numero_reporte", None) or excel_store._cell(row, headers, "Numero_documento", "-"))
                    cli = str(excel_store._cell(row, headers, "Razon_social", None) or "VENTAS DIARIAS")[:15]
                    iva_val = excel_store._parse_monto_cell(excel_store._cell(row, headers, "IVA", None)) or Decimal("0")
                    file_label = "Fact" if "EMITIDAS" in path_v.name else "RepZ"
                    ventas_details.append(f"• {f_doc.strftime('%d/%m')} [{file_label} {doc}]: {cli} | IVA: {excel_store._format_monto_ves(iva_val)} Bs")
            wb.close()
        except Exception as e:
            ventas_details.append(f"Error cargando ventas: {e}")
            
    ret_details = []
    path_r = config.EXCEL_PATH
    if path_r.exists():
        try:
            wb = load_workbook(path_r, read_only=True, data_only=True)
            ws = wb.active
            headers = excel_store._headers_index(ws)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                f_emi = _parse_row_date(excel_store._cell(row, headers, "Fecha_emision", None))
                if f_emi and start_date <= f_emi <= end_date:
                    comp = str(excel_store._cell(row, headers, "Numero_comprobante", "-"))[:10]
                    cli = str(excel_store._cell(row, headers, "RIF", "-"))
                    ret_val = excel_store._parse_monto_cell(excel_store._cell(row, headers, "IVA_retenido", None)) or Decimal("0")
                    ret_details.append(f"• {f_emi.strftime('%d/%m')} [Comp {comp}]: {cli} | Ret: {excel_store._format_monto_ves(ret_val)} Bs")
            wb.close()
        except Exception as e:
            ret_details.append(f"Error cargando retenciones: {e}")
            
    text = f"📊 *DETALLE IVA QUINCENAL* ({start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')})\n\n"
    
    text += "*🛒 compras / crédito fiscal:*\n"
    if compras_details:
        text += "\n".join(compras_details) + "\n\n"
    else:
        text += "No hay compras registradas en este período.\n\n"
        
    text += "*📈 ventas / débito fiscal:*\n"
    if ventas_details:
        text += "\n".join(ventas_details) + "\n\n"
    else:
        text += "No hay ventas registradas en este período.\n\n"
        
    text += "*📥 retenciones recibidas:*\n"
    if ret_details:
        text += "\n".join(ret_details) + "\n"
    else:
        text += "No hay retenciones de clientes en este período.\n"
        
    return text


def _format_islr_details(year: int, month: int, fortnight: int) -> str:
    start_date, end_date = tributario_engine.get_fortnight_range(year, month, fortnight)
    from .tributario_engine import _parse_row_date, ALICUOTA_ANTICIPO_ISLR
    
    ventas_list = []
    for path_v in [config.FACTURAS_EMITIDAS_PATH, config.REPORTES_Z_PATH]:
        if not path_v.exists():
            continue
        try:
            wb = load_workbook(path_v, read_only=True, data_only=True)
            ws = wb.active
            headers = excel_store._headers_index(ws)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                fecha_cell = excel_store._cell(row, headers, "Fecha_emision", None) or excel_store._cell(row, headers, "Fecha", None)
                f_doc = _parse_row_date(fecha_cell)
                if f_doc and start_date <= f_doc <= end_date:
                    doc = str(excel_store._cell(row, headers, "Numero_reporte", None) or excel_store._cell(row, headers, "Numero_documento", "-"))
                    base_val = excel_store._parse_monto_cell(excel_store._cell(row, headers, "Base_imponible", None)) or Decimal("0")
                    file_label = "Fact" if "EMITIDAS" in path_v.name else "RepZ"
                    ventas_list.append(f"• {f_doc.strftime('%d/%m')} [{file_label} {doc}]: Base {excel_store._format_monto_ves(base_val)} Bs")
            wb.close()
        except Exception as e:
            ventas_list.append(f"Error cargando ventas: {e}")
            
    v_base, _, _ = tributario_engine.get_sales_totals(start_date, end_date)
    anticipo = v_base * ALICUOTA_ANTICIPO_ISLR
    
    text = f"💸 *DETALLE ANTICIPO ISLR* ({start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')})\n\n"
    text += f"*Fórmula:* Base de Ventas x Alícuota ({(ALICUOTA_ANTICIPO_ISLR * 100):.0f}%)\n\n"
    
    text += "*📈 desglose de bases de ventas:*\n"
    if ventas_list:
        text += "\n".join(ventas_list) + "\n\n"
    else:
        text += "No hay ventas en este período.\n\n"
        
    text += f"💵 *Base total acumulada:* {excel_store._format_monto_ves(v_base)} Bs\n"
    text += f"📊 *Alícuota anticipo:* {(ALICUOTA_ANTICIPO_ISLR * 100):.0f}%\n"
    text += f"👉 *Monto del Anticipo ISLR a pagar:* `{excel_store._format_monto_ves(anticipo)}` Bs"
    
    return text


async def handle_share_email_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q:
        return
    await q.answer()
    msg = q.message
    if not msg:
        return
        
    # Check if we have SMTP configured
    if not config.SMTP_SERVER or not config.SMTP_USER or not config.SMTP_PASSWORD:
        await msg.reply_text(
            "❌ *Configuración de correo incompleta*\n\n"
            "Para poder enviar cotizaciones por correo, debes configurar las variables SMTP (`SMTP_SERVER`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`) en tu archivo `.env`.",
            parse_mode="Markdown"
        )
        return
        
    share_doc = context.user_data.get("share_doc")
    if not share_doc:
        await msg.reply_text("❌ No encontré información del documento para compartir. Inicia el proceso nuevamente.")
        return
        
    share_doc["awaiting"] = "client_email"
    
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancelar", callback_data="share_cancel")]])
    await msg.reply_text(
        f"📧 *Enviar {share_doc['title']} Nro {share_doc['doc_number']} por Correo*\n\n"
        f"Por favor, escribe la dirección de correo electrónico del cliente a donde deseas enviar el PDF:",
        reply_markup=kb,
        parse_mode="Markdown"
    )

async def handle_share_cancel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q:
        return
    await q.answer()
    msg = q.message
    if not msg:
        return
    context.user_data.pop("share_doc", None)
    try:
        await q.delete_message()
    except Exception:
        pass
    await msg.reply_text("❌ Envío de correo cancelado.")


async def _send_document_email_async(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    recipient_email: str,
    status_msg,
) -> None:
    from .email_sender import send_report_email
    
    share_doc = context.user_data.get("share_doc")
    if not share_doc:
        await status_msg.edit_text("❌ Error: Información del documento perdida.")
        return
        
    pdf_path = Path(share_doc["pdf_path"])
    title = share_doc["title"]
    doc_num = share_doc["doc_number"]
    client_name = share_doc["client_name"]
    total_str = share_doc["total_amount"]
    
    subject = f"{title} Nro {doc_num} - SUFEVICA"
    body = (
        f"Estimado/a {client_name},\n\n"
        f"Adjunto a este correo encontrará su {title} Nro {doc_num} por un monto de {total_str}.\n\n"
        f"Agradecemos su preferencia.\n\n"
        f"Atentamente,\n"
        f"FREDDY LOPEZ (SUFEVICA)"
    )
    
    try:
        await asyncio.to_thread(
            send_report_email,
            recipient_email,
            subject,
            body,
            [pdf_path]
        )
        await status_msg.edit_text(
            f"✅ *¡Correo enviado con éxito!*\n\n"
            f"La {title.lower()} Nro *{doc_num}* ha sido enviada exitosamente a la dirección: `{recipient_email}`",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.exception("Error al enviar documento por correo")
        await status_msg.edit_text(
            f"❌ *Error al enviar el correo*\n\n"
            f"• *Destinatario:* `{recipient_email}`\n"
            f"• *Detalle:* `{e!s}`\n\n"
            f"Por favor verifica la conexión o tus credenciales SMTP en el `.env`.",
            parse_mode="Markdown"
        )
    finally:
        context.user_data.pop("share_doc", None)


def _get_and_increment_correlativo(doc_type: str) -> str:
    correlativos_path = Path(__file__).resolve().parent / "modulo_cotizaciones" / "correlativos.json"
    data = {"cotizacion": 1, "nota": 1}
    if correlativos_path.exists():
        try:
            import json
            data = json.loads(correlativos_path.read_text(encoding="utf-8"))
        except Exception:
            pass
    num = data.get(doc_type, 1)
    data[doc_type] = num + 1
    try:
        import json
        correlativos_path.write_text(json.dumps(data, indent=4), encoding="utf-8")
    except Exception:
        pass
    return f"{num:06d}"


async def _start_document_flow(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    doc_type: str,
) -> None:
    msg = update.effective_message
    if not msg:
        return
    
    emoji = "📋" if doc_type == "cotizacion" else "📦"
    title_up = "COTIZACIÓN" if doc_type == "cotizacion" else "NOTA DE ENTREGA"
    
    context.user_data["pending_doc"] = {
        "type": doc_type,
        "awaiting": "text_data"
    }
    
    prompt = await msg.reply_text(
        f"{emoji} *NUEVA {title_up}* {emoji}\n\n"
        f"Por favor, *pega o escribe aquí el texto con los datos* del cliente y los productos "
        f"(puedes copiarlo directamente desde Excel, WhatsApp o cualquier factura).\n\n"
        f"Yo me encargaré de organizarlos y calcularlos automáticamente.",
        parse_mode="Markdown"
    )
    context.user_data["pending_doc"]["start_prompt_message_id"] = prompt.message_id


async def _process_document_text_immediate(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    doc_type: str,
    text_content: str,
) -> None:
    msg = update.effective_message
    if not msg:
        return
        
    doc_data = _parse_document_text_explicit(text_content, doc_type)
    if doc_data is not None:
        context.user_data["pending_doc"] = {
            "type": doc_type,
            "awaiting": "edit_card",
            "parsed_data": doc_data
        }
        try:
            await msg.delete()
        except Exception:
            pass
        await _send_client_data_card(update, context, first_time=True)
    else:
        context.user_data["pending_doc"] = {
            "type": doc_type,
            "awaiting": "text_data"
        }
        await msg.reply_text(
            "⚠️ No pude identificar productos en el texto que incluiste después del comando.\n\n"
            "Por favor, vuelve a intentarlo *pegando o escribiendo solo el texto* con los datos aquí (asegúrate de incluir cantidades y precios, ej: `10 x PARP-220 Perno a 2.50`):\n\n"
            "O envía `/start` para cancelar.",
            parse_mode="Markdown"
        )


async def cotizacion_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _allowed(update):
        await _deny(update)
        return
    msg = update.effective_message
    if not msg:
        return
        
    text_content = ""
    if msg.text:
        parts = msg.text.split(None, 1)
        if len(parts) > 1:
            text_content = parts[1].strip()
            
    if text_content:
        await _process_document_text_immediate(update, context, "cotizacion", text_content)
    else:
        await _start_document_flow(update, context, "cotizacion")


async def nota_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _allowed(update):
        await _deny(update)
        return
    msg = update.effective_message
    if not msg:
        return
        
    text_content = ""
    if msg.text:
        parts = msg.text.split(None, 1)
        if len(parts) > 1:
            text_content = parts[1].strip()
            
    if text_content:
        await _process_document_text_immediate(update, context, "nota", text_content)
    else:
        await _start_document_flow(update, context, "nota")


async def _send_client_data_card(update: Update, context: ContextTypes.DEFAULT_TYPE, first_time: bool = False) -> None:
    msg = update.effective_message
    if not msg:
        return
    pending_doc = context.user_data.get("pending_doc")
    if not pending_doc or "parsed_data" not in pending_doc:
        return
    doc_data = pending_doc["parsed_data"]
    client = doc_data["client"]
    doc_type = doc_data["docType"]
    title_up = "COTIZACIÓN" if doc_type == "cotizacion" else "NOTA DE ENTREGA"
    emoji = "📋" if doc_type == "cotizacion" else "📦"
    
    rate = doc_data.get("exchangeRate", get_current_bcv_rate())
    text = (
        f"📝 *DATOS DEL CLIENTE DETECTADOS ({title_up})* {emoji}\n\n"
        f"👤 *Cliente:* {client.get('name') or '_[Omitido / Vacío]_'}\n"
        f"🆔 *RIF/CI:* {client.get('rif') or '_[Omitido / Vacío]_'}\n"
        f"📍 *Dirección:* {client.get('address') or '_[Omitido / Vacío]_'}\n"
        f"📞 *Teléfono:* {client.get('phone') or '_[Omitido / Vacío]_'}\n"
        f"👔 *Vendedor:* {client.get('salesman') or '_[Omitido / Vacío]_'}\n"
        f"💳 *Condición de Pago:* {client.get('saleType') or 'Contado'}\n"
        f"📝 *Nota:* {client.get('note') or '_[Nota por defecto]_'}\n"
        f"💵 *Tasa BCV:* Bs. {rate:,.2f}\n\n"
        f"¿Deseas modificar o añadir alguno de estos datos antes de continuar?\n"
        f"_(Cualquiera de estos datos puede ser omitido para dejarlo en blanco)_"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 Cliente", callback_data="coti_edit_name"), InlineKeyboardButton("🆔 RIF/CI", callback_data="coti_edit_rif")],
        [InlineKeyboardButton("📍 Dirección", callback_data="coti_edit_address"), InlineKeyboardButton("📞 Teléfono", callback_data="coti_edit_phone")],
        [InlineKeyboardButton("👔 Vendedor", callback_data="coti_edit_salesman"), InlineKeyboardButton("💳 Condición: " + client.get("saleType", "Contado"), callback_data="coti_edit_saletype")],
        [InlineKeyboardButton("📝 Nota", callback_data="coti_edit_note"), InlineKeyboardButton("💵 Tasa BCV", callback_data="coti_edit_rate")],
        [InlineKeyboardButton("⏩ CONTINUAR A LA MONEDA", callback_data="coti_edit_done")]
    ])
    menu_message_id = pending_doc.get("menu_message_id") if pending_doc else None
    chat_id = update.effective_chat.id if update.effective_chat else None
    
    if first_time or not menu_message_id or not chat_id:
        sent_msg = await context.bot.send_message(
            chat_id=chat_id,
            text=text,
            reply_markup=kb,
            parse_mode="Markdown"
        )
        if pending_doc:
            pending_doc["menu_message_id"] = sent_msg.message_id
    else:
        try:
            await context.bot.edit_message_text(
                chat_id=chat_id,
                message_id=menu_message_id,
                text=text,
                reply_markup=kb,
                parse_mode="Markdown"
            )
        except Exception:
            sent_msg = await context.bot.send_message(
                chat_id=chat_id,
                text=text,
                reply_markup=kb,
                parse_mode="Markdown"
            )
            if pending_doc:
                pending_doc["menu_message_id"] = sent_msg.message_id


async def handle_cotizaciones_edit_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    if not q:
        return
    await q.answer()
    data = (q.data or "").strip()
    msg = q.message
    if not msg:
        return
    pending_doc = context.user_data.get("pending_doc")
    if not pending_doc or "parsed_data" not in pending_doc:
        await msg.reply_text("❌ No hay ningún documento pendiente de procesar.")
        return
    doc_data = pending_doc["parsed_data"]
    client = doc_data["client"]
    
    prompt = None
    if data == "coti_edit_name":
        pending_doc["awaiting"] = "input_name"
        prompt = await msg.reply_text("👤 Envía el *Nombre del Cliente* (o escribe `omitir` para dejarlo en blanco):", parse_mode="Markdown")
    elif data == "coti_edit_rif":
        pending_doc["awaiting"] = "input_rif"
        prompt = await msg.reply_text("🆔 Envía el *RIF o C.I. del Cliente* (o escribe `omitir` para dejarlo en blanco):", parse_mode="Markdown")
    elif data == "coti_edit_address":
        pending_doc["awaiting"] = "input_address"
        prompt = await msg.reply_text("📍 Envía la *Dirección Fiscal del Cliente* (o escribe `omitir` para dejarlo en blanco):", parse_mode="Markdown")
    elif data == "coti_edit_phone":
        pending_doc["awaiting"] = "input_phone"
        prompt = await msg.reply_text("📞 Envía el *Teléfono del Cliente* (o escribe `omitir` para dejarlo en blanco):", parse_mode="Markdown")
    elif data == "coti_edit_salesman":
        pending_doc["awaiting"] = "input_salesman"
        prompt = await msg.reply_text("👔 Envía el *Nombre del Vendedor* (o escribe `omitir` para dejarlo en blanco):", parse_mode="Markdown")
    elif data == "coti_edit_note":
        pending_doc["awaiting"] = "input_note"
        prompt = await msg.reply_text("📝 Envía el texto de la *Nota* personalizada para el pie de página (o escribe `omitir` para usar la nota por defecto):", parse_mode="Markdown")
    elif data == "coti_edit_rate":
        pending_doc["awaiting"] = "input_rate"
        prompt = await msg.reply_text("💵 Envía el valor de la *Tasa de Cambio BCV* (ej. `40.50`):", parse_mode="Markdown")
        
    if prompt:
        pending_doc["prompt_message_id"] = prompt.message_id
    elif data == "coti_edit_saletype":
        current = client.get("saleType", "Contado")
        client["saleType"] = "Crédito" if current == "Contado" else "Contado"
        await _send_client_data_card(update, context)
    elif data == "coti_edit_done":
        pending_doc["awaiting"] = "currency"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("Dólares Americanos ($)", callback_data="coti_curr_usd"), InlineKeyboardButton("Bolívares (Bs.)", callback_data="coti_curr_ves")]
        ])
        try:
            await q.delete_message()
        except Exception:
            pass
        await msg.reply_text("💵 ¿En qué moneda deseas que se exprese el documento por defecto al abrirse?", reply_markup=kb, parse_mode="Markdown")


async def handle_cotizaciones_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    q = update.callback_query
    if not q:
        return
    await q.answer()
    data = (q.data or "").strip()
    msg = q.message
    if not msg:
        return
        
    pending_doc = context.user_data.get("pending_doc")
    if not pending_doc or "parsed_data" not in pending_doc:
        await msg.reply_text("❌ No hay ningún documento pendiente de procesar. Inicia el flujo nuevamente.")
        return
        
    currency = "ves" if "ves" in data else "usd"
    doc_data = pending_doc["parsed_data"]
    doc_data["currency"] = currency
    
    # Asignar correlativo automático secuencial
    doc_type = doc_data["docType"]
    correlativo = _get_and_increment_correlativo(doc_type)
    doc_data["docNumber"] = correlativo
    
    # Eliminar mensaje de espera
    try:
        await q.delete_message()
    except Exception:
        pass
        
    await _generate_document_from_parsed_data(update, context, doc_data)
    context.user_data.pop("pending_doc", None)


async def tributos_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _allowed(update):
        await _deny(update)
        return
    today = date.today()
    fortnight = 1 if today.day <= 15 else 2
    report = tributario_engine.get_compromiso_tributario_report(today.year, today.month, fortnight)
    text = format_tributos_report(report)
    kb = _tributos_keyboard(today.year, today.month, fortnight, _generate_short_summary(report))
    
    msg = update.effective_message
    if msg:
        await msg.reply_text(text, reply_markup=kb, parse_mode="Markdown")


async def handle_tributos_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    q = update.callback_query
    if not q:
        return
    await q.answer()
    data = (q.data or "").strip()
    msg = q.message
    if not msg:
        return
        
    if data.startswith("tributos_period_"):
        parts = data.split("_")
        y = int(parts[2])
        m = int(parts[3])
        f = int(parts[4])
        report = tributario_engine.get_compromiso_tributario_report(y, m, f)
        text = format_tributos_report(report)
        kb = _tributos_keyboard(y, m, f, _generate_short_summary(report))
        await msg.edit_text(text, reply_markup=kb, parse_mode="Markdown")
        
    elif data.startswith("tributos_detiva_"):
        parts = data.split("_")
        y = int(parts[2])
        m = int(parts[3])
        f = int(parts[4])
        text = _format_iva_details(y, m, f)
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Volver al Reporte", callback_data=f"tributos_period_{y}_{m}_{f}")]])
        await msg.edit_text(text, reply_markup=kb, parse_mode="Markdown")
        
    elif data.startswith("tributos_detislr_"):
        parts = data.split("_")
        y = int(parts[2])
        m = int(parts[3])
        f = int(parts[4])
        text = _format_islr_details(y, m, f)
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Volver al Reporte", callback_data=f"tributos_period_{y}_{m}_{f}")]])
        await msg.edit_text(text, reply_markup=kb, parse_mode="Markdown")
        
    elif data == "tributos_selmonth":
        text = "📅 *Selecciona la quincena y el mes que deseas consultar:*"
        kb = _months_selection_keyboard()
        await msg.edit_text(text, reply_markup=kb, parse_mode="Markdown")
        
    elif data.startswith("tributos_sendemail_"):
        parts = data.split("_")
        y = int(parts[2])
        m = int(parts[3])
        f = int(parts[4])
        
        # Verificar si SMTP está configurado
        if not config.SMTP_SERVER or not config.SMTP_USER or not config.SMTP_PASSWORD:
            await msg.reply_text(
                "❌ *Función de Correo SMTP no configurada*\n\n"
                "Para poder enviar reportes directamente en segundo plano con archivos adjuntos, "
                "debes definir las siguientes variables en tu archivo `.env`:\n\n"
                "• `SMTP_SERVER` (ej: smtp.gmail.com)\n"
                "• `SMTP_PORT` (ej: 587)\n"
                "• `SMTP_USER` (tu dirección de correo)\n"
                "• `SMTP_PASSWORD` (tu contraseña de aplicación)\n\n"
                "_Mientras tanto, puedes presionar el botón '📧 Correo Rápido' de arriba para enviar desde tu propio cliente._",
                parse_mode="Markdown"
            )
            return
            
        context.user_data["pending_email_report"] = {
            "year": y,
            "month": m,
            "fortnight": f,
            "awaiting": "email_address"
        }
        
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ Cancelar Envío", callback_data="tributos_cancelemail")]])
        default_lbl = f"`{config.DEFAULT_ACCOUNTANT_EMAIL}`" if config.DEFAULT_ACCOUNTANT_EMAIL else "_no configurado_"
        await msg.reply_text(
            "📧 *Envío de Reportes por Correo (SMTP)*\n\n"
            "Por favor, escribe la dirección de correo electrónico del destinatario a donde deseas enviar los reportes:\n\n"
            f"👉 *O escribe la palabra *`contador`* para enviar al destinatario predeterminado:* {default_lbl}",
            reply_markup=kb,
            parse_mode="Markdown"
        )
        
    elif data == "tributos_cancelemail":
        context.user_data.pop("pending_email_report", None)
        await msg.reply_text("❌ Envío de reportes por correo cancelado.")


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Manejador global de errores para capturar y registrar excepciones no controladas."""
    logger.error("Excepción capturada mientras se procesaba una actualización:", exc_info=context.error)
    
    # Intentar responder al usuario si es posible
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(
                "⚠️ Ocurrió un error inesperado al procesar tu solicitud.\n"
                "El detalle técnico ha sido registrado en los logs del servidor."
            )
        except Exception:
            pass


def build_application() -> Application:
    # Configurar un cliente HTTP con tiempos de espera mayores (30 segundos) para conexiones lentas o inestables
    from telegram.request import HTTPXRequest
    req = HTTPXRequest(connect_timeout=30.0, read_timeout=30.0)
    
    app = Application.builder().token(config.BOT_TOKEN).request(req).post_init(post_init).build()
    app.add_handler(CommandHandler("start", start_cmd))
    app.add_handler(CommandHandler("mi_id", mi_id_cmd))
    app.add_handler(CommandHandler("tributos", tributos_cmd))
    app.add_handler(CommandHandler("cotizacion", cotizacion_cmd))
    app.add_handler(CommandHandler("nota", nota_cmd))
    app.add_handler(CommandHandler("descargar_excel", descargar_excel_cmd))
    app.add_handler(CallbackQueryHandler(handle_emit_retention_callback, pattern=r"^emit_"))
    app.add_handler(CallbackQueryHandler(handle_tributos_callback, pattern=r"^tributos_"))
    app.add_handler(CallbackQueryHandler(handle_cotizaciones_edit_callback, pattern=r"^coti_edit_"))
    app.add_handler(CallbackQueryHandler(handle_cotizaciones_callback, pattern=r"^coti_curr_"))
    app.add_handler(CallbackQueryHandler(handle_share_email_callback, pattern=r"^share_email$"))
    app.add_handler(CallbackQueryHandler(handle_share_cancel_callback, pattern=r"^share_cancel$"))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    # Acepta texto normal y también caption de fotos/documentos.
    app.add_handler(MessageHandler((filters.TEXT | filters.CAPTION) & ~filters.COMMAND, handle_text))
    
    # Registrar el manejador global de errores
    app.add_error_handler(error_handler)
    return app


def main() -> None:
    # 1. Inicializar el sistema de logs (tanto consola como archivo rotativo bot.log)
    _setup_logging()
    logger.info("Iniciando el bot financiero de Telegram...")
    
    # 2. Bucle de ejecución resiliente con autoreconexión y detección de conflictos
    reconnect_delay = 15
    while True:
        try:
            app = build_application()
            if config.RENDER_EXTERNAL_URL:
                logger.info("Bot en marcha (Webhook en Render). Usuario permitido: %s", config.ALLOWED_USER_ID)
                webhook_url = f"{config.RENDER_EXTERNAL_URL}/{config.BOT_TOKEN}"
                logger.info("Configurando webhook en: %s en el puerto %d", webhook_url, config.PORT)
                app.run_webhook(
                    listen="0.0.0.0",
                    port=config.PORT,
                    url_path=config.BOT_TOKEN,
                    webhook_url=webhook_url,
                    allowed_updates=Update.ALL_TYPES,
                    close_loop=False
                )
            else:
                logger.info("Bot en marcha (polling). Usuario permitido: %s", config.ALLOWED_USER_ID)
                # Ejecutar el bot. Este método bloquea hasta que recibe señal de parada (Ctrl+C, SIGINT, SIGTERM)
                app.run_polling(allowed_updates=Update.ALL_TYPES, close_loop=False)
            logger.info("El bot se ha detenido de manera limpia.")
            break
            
        except Conflict as e:
            logger.critical(
                "¡CONFLICTO DE TOKEN DETECTADO! Hay otra instancia de este bot ejecutándose con el mismo token en otro lugar.\n"
                "Por favor, detén todas las otras ventanas de terminal o procesos de python.exe que estén usando este token.\n"
                "Detalle del error: %s\n"
                "Reintentando conexión en 60 segundos...", e
            )
            time.sleep(60)
            
        except (NetworkError, TimedOut) as e:
            logger.warning(
                "Falla temporal de red detectada: %s.\n"
                "Reintentando establecer conexión en %d segundos...", e, reconnect_delay
            )
            time.sleep(reconnect_delay)
            
        except TelegramError as e:
            logger.error(
                "Error en la API de Telegram: %s.\n"
                "Reintentando en 30 segundos...", e
            )
            time.sleep(30)
            
        except Exception as e:
            logger.exception(
                "Ocurrió un error inesperado e inusual en el bucle principal: %s.\n"
                "Reintentando arranque automático en 30 segundos...", e
            )
            time.sleep(30)
