"""Aplicacion Telegram: texto/voz, Excel y restriccion por usuario."""

from __future__ import annotations

import asyncio
import logging
import re
import tempfile
import unicodedata
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup, Update
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters

from . import config
from . import excel_store
from .factura_compra_parse import parse_factura_compra_text
from .transcription import transcribe_audio_file

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

VOICE_BUTTON = "🎤 Activar comando de voz"
VOICE_CANCEL_BUTTON = "❌ Cancelar voz"
RETENTION_RATE = Decimal("0.75")


def _main_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(VOICE_BUTTON), KeyboardButton(VOICE_CANCEL_BUTTON)],
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
            ("numeros_facturas", "numero_factura", "nro_factura", "nro factura", "número factura"),
        ),
        "controles_facturas": _extract_labeled_value_eol(
            t,
            ("controles_facturas", "control_factura", "control", "nro control", "numero de control"),
        ),
        "total_compra_con_iva": _extract_labeled_value_eol(
            t,
            ("total_compra_con_iva", "total compra", "total compra con iva", "total"),
        ),
        "base_imponible": _extract_labeled_value_eol(
            t,
            ("base_imponible", "base imponible", "base gravable", "base"),
        ),
        "iva_retenido": _extract_labeled_value_eol(
            t,
            ("iva_retenido", "iva retenido", "monto iva retenido", "iva retenido total"),
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
            ("numeros_facturas", "numero_factura", "nro_factura", "nro factura"),
        ),
        "controles_facturas": _extract_labeled_value(
            t,
            ("controles_facturas", "control_factura", "control", "nro control"),
        ),
        "total_compra_con_iva": _extract_labeled_value(
            t,
            ("total_compra_con_iva", "total compra", "total"),
        ),
        "base_imponible": _extract_labeled_value(
            t,
            ("base_imponible", "base imponible", "base"),
        ),
        "iva_retenido": _extract_labeled_value(
            t,
            ("iva_retenido", "iva retenido", "monto iva retenido"),
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
    # En SUFEVICA registrar automáticamente si el texto ya contiene los campos.
    if ret_data is None and (_is_sufevica_chat(update) or is_channel):
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

    fc = parse_factura_compra_text(text)
    if fc is not None:
        # Siempre se guarda aunque haya inconsistencia matemática. Solo se anota advertencia.
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
            excel_store.append_factura_compra(
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
            "• Texto o nota de voz: reconozco ordenes y registro retenciones en Excel.\n"
            "• Botón: pulsa «Activar comando de voz» y luego envía tu nota de voz.\n"
            "• Escribe: «resumen de hoy», «enviar excel» (retenciones), "
            "«enviar excel facturas compra»,\n"
            "  «retenciones recibidas del 01/05/2026 al 31/05/2026 en pdf/excel»,\n"
            "  «Emitir retencion de facturas:06949950|06949951»,\n"
            "  o «registrar retencion, fecha:..., comprobante:..., rif:..., iva retenido:...».\n"
            "• Facturas de compra (proveedor -> SUFEVICA): pegar el texto; se guardan en "
            "facturas_compra_recibidas.xlsx (Subtotal, IVA, Total en columnas propias).\n"
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
    if context.user_data.get("voice_mode"):
        await msg.reply_text(
            "Modo voz activo: envía una nota de voz o pulsa «Cancelar voz».",
            reply_markup=_main_keyboard(),
        )
        return
    await _process_intent(update, context, text)


def build_application() -> Application:
    app = Application.builder().token(config.BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start_cmd))
    app.add_handler(CommandHandler("help", start_cmd))
    app.add_handler(CommandHandler("mi_id", mi_id_cmd))
    app.add_handler(CallbackQueryHandler(handle_emit_retention_callback, pattern=r"^emit_"))
    app.add_handler(MessageHandler(filters.VOICE, handle_voice))
    # Acepta texto normal y también caption de fotos/documentos.
    app.add_handler(MessageHandler((filters.TEXT | filters.Caption) & ~filters.COMMAND, handle_text))
    return app


def main() -> None:
    app = build_application()
    logger.info("Bot en marcha (polling). Usuario permitido: %s", config.ALLOWED_USER_ID)
    app.run_polling(allowed_updates=Update.ALL_TYPES)
