"""Extraccion de datos de comprobantes usando Gemini Vision."""

from __future__ import annotations

import json
from dataclasses import dataclass
from io import BytesIO
import time
import logging

from google import genai
from PIL import Image

from . import config

logger = logging.getLogger(__name__)

def _generate_content_with_retry(client, model: str, contents: list, max_retries: int = 5, initial_delay: float = 2.0) -> object:
    delay = initial_delay
    for attempt in range(max_retries):
        try:
            return client.models.generate_content(model=model, contents=contents)
        except Exception as e:
            logger.warning(
                f"Error al llamar a Gemini (intento {attempt + 1}/{max_retries}): {e}"
            )
            if attempt == max_retries - 1:
                raise e
            time.sleep(delay)
            delay *= 2



@dataclass
class Extracted:
    rif: str
    fecha_emision: str
    numero_comprobante: str
    fechas_facturas: str
    numeros_facturas: str
    controles_facturas: str
    total_compra_iva: str
    base_imponible: str
    iva_retenido: str
    raw_text: str


SYSTEM_PROMPT = """
Analiza esta imagen de un comprobante de retencion de IVA (SENIAT).
Devuelve UNICAMENTE un JSON valido con estas claves exactas:
- rif
- fecha_emision
- numero_comprobante
- fechas_facturas
- numeros_facturas
- controles_facturas
- total_compra_iva
- base_imponible
- iva_retenido
- raw_text

Reglas:
1) Si un dato no aparece con claridad, usa cadena vacia.
2) No agregues texto fuera del JSON.
3) Mantiene fechas como DD/MM/YYYY cuando sea posible.
4) En campos de listas, usa elementos separados por coma y espacio.
"""


def _image_to_bytes(image: Image.Image) -> bytes:
    out = BytesIO()
    image.convert("RGB").save(out, format="JPEG", quality=92)
    return out.getvalue()


def _safe_str(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _extract_json_payload(text: str) -> dict[str, object]:
    t = text.strip()
    if t.startswith("```"):
        t = t.strip("`").replace("json\n", "", 1).strip()
    start = t.find("{")
    end = t.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("La respuesta del modelo no contiene JSON valido.")
    return json.loads(t[start : end + 1])


def extract_from_image(image: Image.Image) -> Extracted:
    if not config.GEMINI_API_KEY:
        raise RuntimeError(
            "Falta GEMINI_API_KEY en .env. Configuralo para extraer datos de comprobantes."
        )

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    response = _generate_content_with_retry(
        client,
        model=config.GEMINI_MODEL,
        contents=[
            SYSTEM_PROMPT,
            genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
        ],
    )
    response_text = _safe_str(getattr(response, "text", ""))
    data = _extract_json_payload(response_text)

    return Extracted(
        rif=_safe_str(data.get("rif")),
        fecha_emision=_safe_str(data.get("fecha_emision")),
        numero_comprobante=_safe_str(data.get("numero_comprobante")),
        fechas_facturas=_safe_str(data.get("fechas_facturas")),
        numeros_facturas=_safe_str(data.get("numeros_facturas")),
        controles_facturas=_safe_str(data.get("controles_facturas")),
        total_compra_iva=_safe_str(data.get("total_compra_iva")),
        base_imponible=_safe_str(data.get("base_imponible")),
        iva_retenido=_safe_str(data.get("iva_retenido")),
        raw_text=_safe_str(data.get("raw_text")) or response_text,
    )


import re
from .factura_compra_parse import FacturaCompraParsed

INVOICE_PROMPT = """
Analiza esta imagen de una factura de compra (proveedor emitiendo a SUFEVICA/Suministros Ferreteros Vittoria).
Devuelve UNICAMENTE un JSON valido con estas claves exactas:
- tipo_documento (ej: "Factura recibida / compra", "Nota de Credito", "Nota de Debito")
- fecha_emision (formato DD/MM/YYYY)
- fecha_vencimiento (formato DD/MM/YYYY o vacio "")
- numero_documento (numero de factura)
- numero_control (numero de control de factura)
- proveedor (nombre o razon social del emisor)
- proveedor_rif (RIF del emisor, ej: J-12345678-9)
- proveedor_telefono (telefono del emisor si esta visible, o vacio "")
- direccion_fiscal_proveedor (direccion del emisor si esta visible, o vacio "")
- receptor (nombre o razon social del receptor, ej: SUMINISTROS FERRETEROS VITTORIA, C.A.)
- receptor_rif (RIF del receptor, ej: J-40194130-3)
- subtotal (monto subtotal en la moneda original de los montos extraidos, ej: 2000.00)
- monto_exento (monto exento de IVA en la moneda original de los montos extraidos, ej: 0.00)
- base_imponible (base imponible gravada en la moneda original de los montos extraidos, ej: 2000.00)
- monto_iva (monto del IVA en la moneda original de los montos extraidos, ej: 320.00)
- total (monto total en la moneda original de los montos extraidos, ej: 2320.00)
- contribuyente_tipo (determina si el emisor es "Especial", "Ordinario" o "Formal" segun las leyendas de la factura, o vacio "")
- tasa_cambio (tasa de cambio de la factura si esta presente en el documento, ej: 39.58, o vacio "")
- moneda_original (la moneda de los montos extraidos: "VES" para Bolívares o "USD" para Dólares)

Reglas:
1) Moneda y Selección de Datos:
   - Si la factura muestra montos tanto en Bolívares (VES/Bs.) como en Dólares (USD) (factura de doble columna o de pago referencial), DEBES extraer los montos en BOLÍVARES (Bs.) e indicar 'moneda_original': "VES". NO realices ninguna conversión matemática ni multiplicación por tasa de cambio.
   - Solo si los montos están expresados única y exclusivamente en Dólares (USD) sin conversión alguna a Bolívares en el documento, extrae los montos en USD e indica 'moneda_original': "USD".
2) Formato de montos: numérico sin comas ni puntos de miles, utilizando el punto (.) como separador decimal (ej: 12500.50).
3) Formato de tasa_cambio: si está presente, devuélvela como número decimal utilizando el punto como separador decimal (ej: 39.5858), NO elimines el punto o coma decimal original del valor.
4) Si un dato no aparece, usa cadena vacia "".
5) No agregues texto explicativo fuera del JSON.
"""

ISLR_PROMPT = """
Analiza esta imagen de un comprobante de retención de ISLR (Impuesto sobre la Renta) en Venezuela.
Devuelve UNICAMENTE un JSON valido con estas claves exactas:
- numero_comprobante (numero de comprobante de retencion)
- fecha_emision (formato DD/MM/YYYY)
- proveedor (nombre o razon social del proveedor/retenido)
- proveedor_rif (RIF del proveedor/retenido, ej: J-12345678-9)
- concepto_retencion (ej. "Honorarios Profesionales", "Servicios de Mantenimiento", "Fletes", "Publicidad", etc.)
- numero_documento (numero de factura afectada)
- numero_control (numero de control de la factura)
- base_imponible (base imponible gravada en bolivares Bs, sin separadores de miles y con punto decimal, ej: 1000.00)
- porcentaje_retencion (porcentaje aplicado, ej. 2.00 para 2%)
- islr_retenido (monto del ISLR retenido en bolivares Bs, ej: 20.00)
- total_factura (monto total de la factura en bolivares Bs, ej: 1160.00)

Reglas:
1) Todos los montos deben expresarse en Bolívares (Bs.). Si los montos en la imagen están expresados en USD y no hay conversión impresa, realiza la conversión utilizando la tasa de cambio del comprobante. Si ya están en Bs., extrae directamente los montos en bolívares.
2) Formato de montos: numérico sin comas ni puntos de miles, usando punto (.) como separador decimal (ej: 1000.00).
3) Si un dato no aparece, usa cadena vacia "".
4) No agregues texto explicativo fuera del JSON.
"""

CLASSIFY_PROMPT = """
Analiza esta imagen comercial y clasifícala en una de las siguientes categorías exactas:
- "factura" (si es una factura de compra, nota de entrega, presupuesto, etc.)
- "retencion_iva" (si es un comprobante de retención de IVA del SENIAT)
- "retencion_islr" (si es un comprobante de retención de ISLR / Impuesto sobre la Renta)
- "desconocido" (si es cualquier otra cosa)

Devuelve UNICAMENTE la palabra de la categoría correspondiente, en minúsculas y sin comillas.
"""

def classify_image_type(image: Image.Image) -> str:
    if not config.GEMINI_API_KEY:
        return "desconocido"

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    response = _generate_content_with_retry(
        client,
        model=config.GEMINI_MODEL,
        contents=[
            CLASSIFY_PROMPT,
            genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
        ],
    )
    result = _safe_str(getattr(response, "text", "")).strip().lower()
    for cat in ("factura", "retencion_iva", "retencion_islr"):
        if cat in result:
            return cat
    return "desconocido"

def extract_invoice_from_image(image: Image.Image) -> FacturaCompraParsed:
    if not config.GEMINI_API_KEY:
        raise RuntimeError("Falta GEMINI_API_KEY en .env.")

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    response = _generate_content_with_retry(
        client,
        model=config.GEMINI_MODEL,
        contents=[
            INVOICE_PROMPT,
            genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
        ],
    )
    response_text = _safe_str(getattr(response, "text", ""))
    data = _extract_json_payload(response_text)

    return FacturaCompraParsed(
        tipo_documento=_safe_str(data.get("tipo_documento")) or "Factura recibida / compra",
        fecha_emision=_safe_str(data.get("fecha_emision")),
        fecha_vencimiento=_safe_str(data.get("fecha_vencimiento")),
        numero_documento=_safe_str(data.get("numero_documento")),
        numero_control=_safe_str(data.get("numero_control")),
        proveedor=_safe_str(data.get("proveedor")),
        proveedor_rif=_safe_str(data.get("proveedor_rif")),
        proveedor_telefono=_safe_str(data.get("proveedor_telefono")),
        direccion_fiscal_proveedor=_safe_str(data.get("direccion_fiscal_proveedor")),
        receptor=_safe_str(data.get("receptor")),
        receptor_rif=_safe_str(data.get("receptor_rif")),
        subtotal=_safe_str(data.get("subtotal")),
        monto_exento=_safe_str(data.get("monto_exento")) or "0.00",
        base_imponible=_safe_str(data.get("base_imponible")),
        monto_iva=_safe_str(data.get("monto_iva")),
        total=_safe_str(data.get("total")),
        contribuyente_tipo=_safe_str(data.get("contribuyente_tipo")),
        tasa_cambio=_safe_str(data.get("tasa_cambio")),
        moneda_original=_safe_str(data.get("moneda_original")) or "VES",
    )

def extract_islr_from_image(image: Image.Image) -> dict:
    if not config.GEMINI_API_KEY:
        raise RuntimeError("Falta GEMINI_API_KEY en .env.")

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    response = _generate_content_with_retry(
        client,
        model=config.GEMINI_MODEL,
        contents=[
            ISLR_PROMPT,
            genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
        ],
    )
    response_text = _safe_str(getattr(response, "text", ""))
    data = _extract_json_payload(response_text)

    return {
        "numero_comprobante": _safe_str(data.get("numero_comprobante")),
        "fecha_emision": _safe_str(data.get("fecha_emision")),
        "proveedor": _safe_str(data.get("proveedor")),
        "proveedor_rif": _safe_str(data.get("proveedor_rif")),
        "concepto_retencion": _safe_str(data.get("concepto_retencion")),
        "numero_documento": _safe_str(data.get("numero_documento")),
        "numero_control": _safe_str(data.get("numero_control")),
        "base_imponible": _safe_str(data.get("base_imponible")),
        "porcentaje_retencion": _safe_str(data.get("porcentaje_retencion")),
        "islr_retenido": _safe_str(data.get("islr_retenido")),
        "total_factura": _safe_str(data.get("total_factura")),
    }


BARCODE_PROMPT = """
Analiza esta imagen y busca cualquier código de barras lineal (ej. UPC, EAN, Code 128) o código QR que corresponda a un producto.
Devuelve UNICAMENTE el valor del código leído en texto plano, sin espacios ni caracteres adicionales.
Si no hay un código de barras o código QR visible o legible en la imagen, responde estrictamente con la palabra "NONE" en mayúsculas.
No agregues texto explicativo o formato adicional fuera del código decodificado.
"""

def extract_barcode_from_image(image: Image.Image) -> str:
    """
    Usa la API de Gemini para leer un código de barras de la imagen y retornar su valor.
    """
    if not config.GEMINI_API_KEY:
        return "NONE"

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    try:
        response = _generate_content_with_retry(
            client,
            model=config.GEMINI_MODEL,
            contents=[
                BARCODE_PROMPT,
                genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
            ],
        )
        result = _safe_str(getattr(response, "text", "")).strip().upper()
        # Eliminar posible formato markdown de texto plano si Gemini lo incluye
        result = re.sub(r"[`'\"]", "", result).strip()
        if not result or "NONE" in result:
            return "NONE"
        return result
    except Exception as e:
        logger.error(f"Error al extraer codigo de barras con Gemini: {e}")
        return "NONE"


PRODUCT_OCR_PROMPT = """
Analiza esta imagen (que puede ser una etiqueta de producto, caja, lista de precios, factura, empaque o catálogo).
Identifica el producto principal que se muestra o se describe en el texto de la imagen.
Extrae e indica UNICAMENTE el término de búsqueda más relevante (el código del producto o el nombre/descripción principal) para buscarlo en una base de datos de inventario.
Devuelve el resultado en texto plano, en una sola línea, sin ningún formato (sin negritas, sin comillas), explicaciones ni caracteres adicionales.
Si no puedes identificar ningún producto o texto relevante, responde estrictamente con la palabra "NONE" en mayúsculas.
"""

def extract_product_query_from_image(image: Image.Image) -> str:
    """
    Usa la API de Gemini para extraer el texto de búsqueda del producto (código o nombre) de la imagen.
    """
    if not config.GEMINI_API_KEY:
        return "NONE"

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    try:
        response = _generate_content_with_retry(
            client,
            model=config.GEMINI_MODEL,
            contents=[
                PRODUCT_OCR_PROMPT,
                genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
            ],
        )
        result = _safe_str(getattr(response, "text", "")).strip()
        result = re.sub(r"[`'\"]", "", result).strip()
        if not result or "NONE" in result.upper():
            return "NONE"
        return result
    except Exception as e:
        logger.error(f"Error al extraer busqueda de producto con Gemini OCR: {e}")
        return "NONE"


DOCUMENT_PARSER_PROMPT = """
Analiza esta imagen de una nota de entrega, cotización, factura o pedido.
Extrae la información del cliente y la lista de todos los productos/ítems.

Devuelve UNICAMENTE un JSON válido con esta estructura exacta:
{
  "client_name": "nombre del cliente o razón social",
  "client_rif": "RIF del cliente, ej: J-12345678-9",
  "client_address": "dirección fiscal del cliente",
  "client_phone": "teléfono del cliente",
  "items": [
    {
      "code": "código del producto si está visible",
      "desc": "descripción o nombre del producto",
      "qty": 10.0,
      "priceUsd": 1.50
    }
  ]
}

Reglas:
1) Si algún dato del cliente no es visible, usa cadena vacía "".
2) Para cada ítem, extrae la cantidad (qty) y precio unitario en USD (priceUsd). Si los montos están en Bolívares (Bs) y la tasa de cambio está visible, conviértelos a USD o mantén USD si el documento original los indica.
3) Formato numérico para qty y priceUsd: numérico simple (ej: 5.0, 12.50) sin comas ni otros caracteres.
4) No agregues texto explicativo o formato fuera del JSON.
"""

def extract_document_data_from_image(image: Image.Image) -> dict:
    """
    Usa la API de Gemini para extraer datos de cliente y productos de una foto de documento.
    """
    if not config.GEMINI_API_KEY:
        return {}

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    try:
        response = _generate_content_with_retry(
            client,
            model=config.GEMINI_MODEL,
            contents=[
                DOCUMENT_PARSER_PROMPT,
                genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
            ],
        )
        response_text = _safe_str(getattr(response, "text", ""))
        return _extract_json_payload(response_text)
    except Exception as e:
        logger.error(f"Error al extraer datos de documento con Gemini OCR: {e}")
        return {}


DOCUMENT_TEXT_PARSER_PROMPT = """
Analiza este texto que contiene una solicitud de pedido, cotización, factura o nota de entrega.
Extrae la información del cliente y la lista de todos los productos/ítems.

Devuelve UNICAMENTE un JSON válido con esta estructura exacta:
{
  "client_name": "nombre del cliente o razón social",
  "client_rif": "RIF del cliente, ej: J-12345678-9",
  "client_address": "dirección fiscal del cliente",
  "client_phone": "teléfono del cliente",
  "items": [
    {
      "code": "código del producto si está visible o deducible",
      "desc": "descripción o nombre del producto",
      "qty": 10.0,
      "priceUsd": 1.50
    }
  ]
}

Reglas:
1) Si algún dato del cliente no está explícito o no se menciona, usa cadena vacía "".
2) Para cada ítem, extrae la cantidad (qty) y precio unitario en USD (priceUsd). Si los montos están en Bolívares (Bs) y la tasa de cambio está visible o implícita, conviértelos a USD o mantén USD si el original los indica.
3) Formato numérico para qty y priceUsd: numérico simple (ej: 5.0, 12.50) sin comas ni otros caracteres.
4) No agregues texto explicativo o formato fuera del JSON.
"""

def parse_document_text_with_gemini(text: str) -> dict:
    """
    Usa la API de Gemini para extraer datos de cliente y productos a partir de un texto libre.
    """
    if not config.GEMINI_API_KEY:
        return {}

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    try:
        response = _generate_content_with_retry(
            client,
            model=config.GEMINI_MODEL,
            contents=[
                DOCUMENT_TEXT_PARSER_PROMPT,
                text,
            ],
        )
        response_text = _safe_str(getattr(response, "text", ""))
        return _extract_json_payload(response_text)
    except Exception as e:
        logger.error(f"Error al extraer datos de texto con Gemini: {e}")
        return {}
