"""Extraccion de datos de comprobantes usando Gemini Vision."""

from __future__ import annotations

import json
from dataclasses import dataclass
from io import BytesIO

from google import genai
from PIL import Image

from . import config


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
    response = client.models.generate_content(
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
