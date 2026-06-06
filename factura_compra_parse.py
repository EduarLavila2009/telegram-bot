"""Extraccion de datos de texto libre tipo factura de compra (proveedor -> SUFEVICA)."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation

from .excel_store import parse_amount_ves_string


def _norm(s: str) -> str:
    base = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in base if not unicodedata.combining(c)).lower()


def _first_ves_amount_in(text: str) -> str | None:
    """Montos tipo 74.009,12 u 84823,40."""
    m = re.search(
        r"\b(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\b",
        text,
    )
    return m.group(1) if m else None


def _take_eol_label(text: str, label_variants: tuple[str, ...]) -> str | None:
    """Busca 'Etiqueta: valor' hasta fin de linea (incluye comas en montos)."""
    lines = text.splitlines()
    ordered = [lv.lower().strip() for lv in label_variants]
    for raw in lines:
        line = raw.strip()
        if ":" not in line:
            continue
        key, _, rest = line.partition(":")
        key_clean = re.sub(r"^[\s*•\-\d.\)\(]+", "", key).strip().lower()
        key_clean = re.sub(r"\s+", " ", key_clean)
        for lv in ordered:
            if key_clean == lv or key_clean.startswith(lv + " ") or lv in key_clean:
                val = re.sub(r"\*+", "", rest).strip()
                if val:
                    return val
    return None


def _take_total_label(text: str) -> str | None:
    """
    Extrae el TOTAL real evitando confundirlo con SUBTOTAL.
    Prioridad: total general > total a pagar > total factura > monto total > total.
    """
    priorities = (
        "total general",
        "total a pagar",
        "total factura",
        "monto total",
        "total",
    )
    lines = text.splitlines()
    for label in priorities:
        for raw in lines:
            line = raw.strip()
            if ":" not in line:
                continue
            key, _, rest = line.partition(":")
            key_clean = re.sub(r"^[\s*•\-\d.\)\(]+", "", key).strip().lower()
            key_clean = re.sub(r"\s+", " ", key_clean)
            # Nunca tratar "sub-total/subtotal/sub total" como TOTAL.
            if key_clean in {"sub-total", "subtotal", "sub total"}:
                continue
            if key_clean == label or key_clean.startswith(label + " "):
                val = re.sub(r"\*+", "", rest).strip()
                if val:
                    return val
    return None


def _extract_rif_from_segment(seg: str) -> str:
    m = re.search(
        r"RIF\s*:\s*([VEJPGvejpg][\s.\-]*\d[\d.\-]*)\b",
        seg,
        re.IGNORECASE,
    )
    if m:
        raw = m.group(1).strip()
        return re.sub(r"\s+", "", raw)
    m = re.search(
        r"\b([VEJPGvejpg][\s.\-]*\d{6,12}(?:[\s.\-]\d)?)\b",
        seg,
    )
    if m:
        return re.sub(r"\s+", "", m.group(1).strip())
    return ""


def _normalize_rif(rif: str) -> str:
    s = re.sub(r"\s+", "", str(rif or "").strip()).upper()
    if not s:
        return ""
    if "-" not in s and len(s) > 1 and s[0] in "VEJPG":
        s = f"{s[0]}-{s[1:]}"
    return s


def _extract_razon_before_rif(seg: str) -> str:
    """Valor tipo 'FEBECA, C.A. (RIF: J-000033927)' o linea con ':' previo."""
    seg = seg.strip()
    m = re.match(r"(.+?)\s*\(\s*RIF\s*:", seg, re.IGNORECASE | re.DOTALL)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip(" ,.-")
    m = re.search(
        r"[:\-]\s*(.+?)\s*\(?\s*RIF\s*:",
        seg,
        re.IGNORECASE | re.DOTALL,
    )
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip(" ,.-")
    return ""


def looks_like_factura_compra(text: str) -> bool:
    t = _norm(text)
    # Exclusión explícita de comprobantes de retención para evitar que se clasifiquen como compras
    if any(k in t for k in ("comprobante de retencion", "numero de comprobante", "iva retenido", "impuesto retenido")):
        return False
    if "factura recibida" in t:
        return True
    if "factura de compra" in t:
        return True
    # Heurística amplia para textos parciales del canal:
    # si contiene suficientes "keywords" típicas de factura (aunque no diga "factura").
    keywords = (
        "factura",
        "compra",
        "proveedor",
        "emisor",
        "receptor",
        "cliente",
        "numero de control",
        "número de control",
        "numero de documento",
        "número de documento",
        "sub-total",
        "subtotal",
        "base imponible",
        "iva",
        "monto total",
        "total",
        "condicion de pago",
        "condición de pago",
        "sufevica",
        "vittoria",
        "rif:",
    )
    hits = sum(1 for k in keywords if k in t)
    if hits >= 3:
        return True
    return False


@dataclass
class FacturaCompraParsed:
    tipo_documento: str
    fecha_emision: str
    fecha_vencimiento: str
    numero_documento: str
    numero_control: str
    proveedor: str
    proveedor_rif: str
    proveedor_telefono: str
    direccion_fiscal_proveedor: str
    receptor: str
    receptor_rif: str
    subtotal: str
    monto_exento: str
    base_imponible: str
    monto_iva: str
    total: str
    contribuyente_tipo: str = ""
    tasa_cambio: str = ""


def parse_factura_compra_text(text: str) -> FacturaCompraParsed | None:
    if not looks_like_factura_compra(text):
        return None

    fecha_emi = (
        _take_eol_label(
            text,
            (
                "fecha de emision",
                "fecha de emisión",
                "fecha emision",
            ),
        )
        or ""
    )
    fecha_venc = (
        _take_eol_label(
            text,
            ("fecha de vencimiento",),
        )
        or ""
    )

    def _norm_date(s: str) -> str:
        from datetime import datetime

        s = s.strip()
        for a, fmt in (
            (s.replace("/", "-"), "%d-%m-%Y"),
            (s, "%d-%m-%Y"),
            (s.replace("-", "/"), "%d/%m/%Y"),
            (s, "%d/%m/%Y"),
        ):
            try:
                dt = datetime.strptime(a.strip(), fmt)
                return dt.strftime("%d/%m/%Y")
            except ValueError:
                continue
        return s

    if fecha_emi:
        fecha_emi = _norm_date(fecha_emi)
    if fecha_venc:
        fecha_venc = _norm_date(fecha_venc)

    nro_doc = (
        _take_eol_label(
            text,
            (
                "numero de documento",
                "número de documento",
                "nro de documento",
            ),
        )
        or ""
    )
    nro_ctrl_raw = (
        _take_eol_label(
            text,
            (
                "numero de control",
                "número de control",
                "nro de control",
            ),
        )
        or ""
    )
    nro_ctrl = re.sub(r"\s+", "", nro_ctrl_raw) if nro_ctrl_raw else ""

    proveedor_line = _take_eol_label(
        text,
        (
            "emisor (proveedor)",
            "emisor",
            "proveedor",
        ),
    )
    if not proveedor_line:
        m = re.search(
            r"(?im)emisor\s*\(?proveedor\)?\s*:\s*(.+)$",
            text,
        )
        proveedor_line = m.group(1).strip() if m else ""

    receptor_line = _take_eol_label(
        text,
        (
            "receptor (cliente)",
            "receptor",
            "cliente",
        ),
    )
    if not receptor_line:
        m = re.search(
            r"(?im)receptor\s*\(?cliente\)?\s*:\s*(.+)$",
            text,
        )
        receptor_line = m.group(1).strip() if m else ""

    prov_rif = _extract_rif_from_segment(proveedor_line or "")
    if not prov_rif:
        # Fallback: RIF en línea separada, común en textos del canal.
        prov_rif = (
            _take_eol_label(
                text,
                (
                    "rif del proveedor",
                    "rif proveedor",
                    "rif emisor",
                    "rif del emisor",
                    "proveedor rif",
                ),
            )
            or ""
        )
        if prov_rif:
            prov_rif = _extract_rif_from_segment(prov_rif) or prov_rif

    prov_rif = _normalize_rif(prov_rif)
    rec_rif = _extract_rif_from_segment(receptor_line or "")
    rec_rif = _normalize_rif(rec_rif)
    prov_name = _extract_razon_before_rif(proveedor_line) if proveedor_line else ""
    if not prov_name and proveedor_line:
        prov_name = re.sub(r"\(.*?\)", "", proveedor_line).strip(" ,.-")
    rec_name = _extract_razon_before_rif(receptor_line) if receptor_line else ""
    if not rec_name and receptor_line:
        rec_name = re.sub(r"\(.*?\)", "", receptor_line).strip(" ,.-")

    tel_prov = (
        _take_eol_label(
            text,
            (
                "telefono",
                "teléfono",
                "telefono del proveedor",
                "teléfono del proveedor",
                "tlf",
            ),
        )
        or ""
    )
    if not tel_prov and proveedor_line:
        m_tel = re.search(r"(?i)\b(?:tel(?:e|é)fono|tlf)\s*[:\-]?\s*([0-9().+\-\s]{7,})", proveedor_line)
        if m_tel:
            tel_prov = re.sub(r"\s+", "", m_tel.group(1)).strip()

    dir_prov = (
        _take_eol_label(
            text,
            (
                "direccion fiscal del proveedor",
                "dirección fiscal del proveedor",
                "direccion fiscal proveedor",
                "dirección fiscal proveedor",
                "direccion fiscal",
                "dirección fiscal",
            ),
        )
        or ""
    )

    sub_raw = _take_eol_label(
        text,
        ("sub-total", "subtotal", "sub total"),
    )
    exento_raw = _take_eol_label(
        text,
        ("monto exento",),
    )
    base_raw = _take_eol_label(
        text,
        ("base imponible (16.00%)", "base imponible", "base gravable"),
    )
    iva_line = None
    for raw in text.splitlines():
        if re.search(r"^\s*\*?\s*IVA\s*\(", raw, re.I) or re.search(
            r"IVA\s*\(\s*16", raw, re.I
        ):
            iva_line = raw
            break
    iva_raw = None
    if iva_line and ":" in iva_line:
        iva_raw = iva_line.split(":", 1)[-1].strip()
    if not iva_raw:
        iva_raw = _take_eol_label(text, ("iva (16.00%)", "iva"))

    total_raw = None
    m_tot = re.search(
        r"(?im)total\s+a\s+pagar[\s:]+(?:Bs\.?|bol[ií]vares?)?\s*:?\s*([\d.]+,\d{2})",
        text,
    )
    if m_tot:
        total_raw = m_tot.group(1)
    if not total_raw:
        total_raw = _take_total_label(text)
    if total_raw:
        total_raw = re.sub(r"\*+", "", total_raw).strip()

    def _fmt_money_label(label_val: str | None) -> str:
        if not label_val:
            return ""
        amt = _first_ves_amount_in(label_val) or label_val.strip()
        d = parse_amount_ves_string(amt)
        return "" if d is None else format(d, "f")

    sub = _fmt_money_label(sub_raw)
    ex = _fmt_money_label(exento_raw)
    base = _fmt_money_label(base_raw)
    iva = _fmt_money_label(iva_raw)
    tot = _fmt_money_label(total_raw)

    # Permite guardar textos parciales del canal siempre que haya algún identificador/monto:
    # - número de documento o
    # - total/subtotal/base/iva o
    # - RIF del proveedor
    if not (nro_doc or tot or sub or base or iva or prov_rif):
        return None

    return FacturaCompraParsed(
        tipo_documento="Factura recibida / compra",
        fecha_emision=fecha_emi,
        fecha_vencimiento=fecha_venc,
        numero_documento=nro_doc.strip(),
        numero_control=nro_ctrl,
        proveedor=prov_name,
        proveedor_rif=prov_rif,
        proveedor_telefono=tel_prov,
        direccion_fiscal_proveedor=dir_prov,
        receptor=rec_name,
        receptor_rif=rec_rif,
        subtotal=sub,
        monto_exento=ex,
        base_imponible=base,
        monto_iva=iva,
        total=tot,
    )
