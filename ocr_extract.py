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
Analiza esta imagen de una factura, nota de crédito o nota de débito comercial (puede ser de compra o de venta, física o fiscal).
Devuelve UNICAMENTE un JSON válido con estas claves exactas:
- tipo_documento (ej: "Factura", "Nota de Credito", "Nota de Debito")
- fecha_emision (formato DD/MM/YYYY)
- fecha_vencimiento (formato DD/MM/YYYY o vacío "")
- numero_documento (número del documento, ej. número de factura o de nota de crédito)
- numero_control (número de control de la factura o nota de crédito)
- proveedor (nombre o razón social del emisor/vendedor del documento)
- proveedor_rif (RIF del emisor/vendedor, ej: J-12345678-9)
- proveedor_telefono (teléfono del emisor/vendedor si está visible, o vacío "")
- direccion_fiscal_proveedor (dirección del emisor/vendedor si está visible, o vacío "")
- receptor (nombre o razón social del receptor/cliente/comprador)
- receptor_rif (RIF del receptor/cliente/comprador, ej: J-40194130-3)
- subtotal (monto subtotal en la moneda original de los montos extraídos, ej: 2000.00)
- monto_exento (monto exento de IVA en la moneda original de los montos extraídos, ej: 0.00)
- base_imponible (base imponible gravada en la moneda original de los montos extraídos, ej: 2000.00)
- monto_iva (monto del IVA en la moneda original de los montos extraídos, ej: 320.00)
- total (monto total en la moneda original de los montos extraídos, ej: 2320.00)
- contribuyente_tipo (determina si el emisor es "Especial", "Ordinario" o "Formal" según las leyendas del documento, o vacío "")
- tasa_cambio (tasa de cambio presente en el documento, ej: 39.58, o vacío "")
- moneda_original (la moneda de los montos numéricos que colocas en las claves de arriba: "VES" para Bolívares o "USD" para Dólares)

Reglas de Extracción de Montos y Moneda:
1) Identifica claramente quién es el emisor/vendedor (proveedor) y quién es el receptor/cliente. Lee atentamente el membrete del documento comercial. No asumas que la empresa del usuario es siempre el receptor o el emisor; deduce los roles de forma inteligente a partir del membrete y datos fiscales.
2) Si el documento contiene montos expresados en Dólares (USD) (ya sea como moneda principal, en una columna de doble columna, o de forma referencial), y posee una tasa de cambio o tasa BCV referencial anotada:
   - Extrae los valores numéricos correspondientes a la columna de Dólares (USD).
   - Coloca la tasa de cambio en 'tasa_cambio' como número decimal (ej: 39.5858).
   - Asigna 'moneda_original': "USD".
   - El sistema convertirá automáticamente estos valores a Bolívares usando la tasa de cambio provista.
3) Si los montos están expresados única y exclusivamente en Bolívares (VES/Bs.) sin indicar dólares ni tasa referencial:
   - Extrae los valores numéricos en Bolívares.
   - Asigna 'moneda_original': "VES" y 'tasa_cambio': "".
4) Formato de montos: numérico puro, sin separadores de miles y con punto decimal (.) (ej: 12500.50). Si no hay monto exento, responde estrictamente con 0.00.
5) Realiza una validación matemática básica: total = base_imponible + monto_iva + monto_exento. Si hay discrepancias menores de centavos, usa los valores impresos.
6) Si un dato no aparece, responde con cadena vacía "".
7) No agregues texto explicativo o formato fuera del JSON.
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
- "factura" (si es una factura formal de compra o venta, o factura fiscal, NO nota de entrega ni presupuesto ni cotización)
- "nota_credito" (si es una nota de crédito comercial que modifica o anula una factura y contiene texto explícito como "NOTA DE CREDITO")
- "reporte_z" (si es un Reporte Z de una máquina o impresora fiscal, que son tiras de papel largas y angostas de caja registradora con resúmenes diarios de venta)
- "documento_comercial" (si es una nota de entrega, presupuesto, cotización, orden de compra o una captura de pantalla/imagen de tabla de Excel conteniendo productos o listado comercial)
- "retencion_iva" (si es un comprobante de retención de IVA del SENIAT)
- "retencion_islr" (si es un comprobante de retención de ISLR / Impuesto sobre la Renta)
- "desconocido" (si es cualquier otra cosa)

Devuelve UNICAMENTE la palabra de la categoría correspondiente, en minúsculas y sin comillas.
"""

REPORTE_Z_PROMPT = """
Analiza esta imagen de un Reporte Z de una máquina fiscal (resumen diario de ventas de un comercio).
Devuelve UNICAMENTE un JSON válido con estas claves exactas:
- numero_reporte (ej: "0125" o "125")
- fecha_emision (formato DD/MM/YYYY, ej: 15/05/2026)
- sub_total (monto del subtotal o ventas gravadas antes del IVA en bolívares, ej: 1500.50)
- base_imponible (monto total de las ventas gravadas en bolívares, ej: 1500.50)
- monto_exento (monto de ventas exentas o no gravadas en bolívares, ej: 0.00)
- iva (monto total del IVA del día en bolívares, ej: 240.08)
- total (monto total de ventas del día en bolívares, ej: 1740.58)

Reglas:
1) Todos los montos deben ser numéricos sin comas ni puntos de miles, usando punto (.) como separador decimal (ej: 1500.50). Si un monto no aparece, usa "0.00".
2) Si el número de reporte no está claro, busca la palabra "REPORTE Z" o "CIERRE Z" seguida de un número.
3) Si no se encuentra un dato, usa una cadena vacía "".
4) No agregues texto explicativo o formato fuera del JSON.
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
    for cat in ("factura", "nota_credito", "reporte_z", "documento_comercial", "retencion_iva", "retencion_islr"):
        if cat in result:
            return cat
    return "desconocido"

def extract_reporte_z_from_image(image: Image.Image) -> dict:
    if not config.GEMINI_API_KEY:
        raise RuntimeError("Falta GEMINI_API_KEY en .env.")

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    response = _generate_content_with_retry(
        client,
        model=config.GEMINI_MODEL,
        contents=[
            REPORTE_Z_PROMPT,
            genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
        ],
    )
    response_text = _safe_str(getattr(response, "text", ""))
    data = _extract_json_payload(response_text)
    return {
        "numero_reporte": _safe_str(data.get("numero_reporte")),
        "fecha_emision": _safe_str(data.get("fecha_emision")),
        "sub_total": _safe_str(data.get("sub_total")),
        "base_imponible": _safe_str(data.get("base_imponible")),
        "monto_exento": _safe_str(data.get("monto_exento")) or "0.00",
        "iva": _safe_str(data.get("iva")),
        "total": _safe_str(data.get("total")),
    }

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
Extrae la lista de todos los productos/ítems (código, descripción, cantidad, precio unitario).

Devuelve UNICAMENTE un JSON válido con esta estructura exacta:
{
  "client_name": "",
  "client_rif": "",
  "client_address": "",
  "client_phone": "",
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
1) Deja los campos "client_name", "client_rif", "client_address" y "client_phone" siempre vacíos como "".
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
Extrae la lista de todos los productos/ítems (código, descripción, cantidad, precio unitario).

Devuelve UNICAMENTE un JSON válido con esta estructura exacta:
{
  "client_name": "",
  "client_rif": "",
  "client_address": "",
  "client_phone": "",
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
1) Deja los campos "client_name", "client_rif", "client_address" y "client_phone" siempre vacíos como "".
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


def extract_rif_data_from_image(image: Image.Image) -> dict:
    if not config.GEMINI_API_KEY:
        raise RuntimeError("Falta GEMINI_API_KEY en .env.")

    client = genai.Client(api_key=config.GEMINI_API_KEY)
    image_bytes = _image_to_bytes(image)
    
    prompt = """
    Analiza esta imagen del RIF (Registro de Información Fiscal) de Venezuela.
    Extrae la Razón Social (nombre del contribuyente o empresa) y el RIF.
    Devuelve UNICAMENTE un JSON válido con estas claves exactas:
    - razon_social
    - rif
    
    Reglas:
    1) RIF debe tener el formato J-12345678-9 (o V-, G-, etc.).
    2) Si no encuentras alguno de los campos, usa cadena vacía "".
    3) No agregues texto explicativo fuera del JSON.
    """
    
    response = _generate_content_with_retry(
        client,
        model=config.GEMINI_MODEL,
        contents=[
            prompt,
            genai.types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
        ],
    )
    response_text = _safe_str(getattr(response, "text", ""))
    return _extract_json_payload(response_text)
