"""Motor tributario para el cálculo de IVA por pagar y compromisos quincenales (SENIAT)."""

from __future__ import annotations

import logging
import re
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from . import config
from . import excel_store

logger = logging.getLogger(__name__)

# Alícuota estándar de anticipo de ISLR en Venezuela para la mayoría de contribuyentes especiales (1%)
ALICUOTA_ANTICIPO_ISLR = Decimal("0.01")


def get_fortnight_range(year: int, month: int, fortnight: int) -> tuple[date, date]:
    """
    Retorna el rango de fechas (desde, hasta) para una quincena específica de un mes/año.
    
    fortnight = 1: días 01 al 15
    fortnight = 2: días 16 al último día del mes
    """
    if fortnight == 1:
        start = date(year, month, 1)
        end = date(year, month, 15)
        return start, end
    elif fortnight == 2:
        start = date(year, month, 16)
        # Calcular el último día del mes
        if month == 12:
            end = date(year, month, 31)
        else:
            # Siguiente mes día 1 menos 1 día
            next_month = date(year, month + 1, 1)
            import datetime as dt
            end = next_month - dt.timedelta(days=1)
        return start, end
    else:
        raise ValueError("La quincena debe ser 1 o 2.")


def _parse_row_date(val: object) -> date | None:
    """Intenta parsear un valor de celda de fecha con máxima flexibilidad."""
    if val is None:
        return None
    if isinstance(val, (date, datetime)):
        return val.date() if isinstance(val, datetime) else val
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def get_sales_totals(start_date: date, end_date: date) -> tuple[Decimal, Decimal, int]:
    """
    Calcula totales de ventas (Base Imponible, IVA, número de registros)
    a partir de FACTURAS-EMITIDAS.xlsx y REPORTES-Z.xlsx en el rango dado.
    """
    total_base = Decimal("0")
    total_iva = Decimal("0")
    total_count = 0

    files_to_read = [config.FACTURAS_EMITIDAS_PATH, config.REPORTES_Z_PATH]

    for path in files_to_read:
        if not path.exists():
            continue
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            headers = excel_store._headers_index(ws)
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                # Se asume columnas estándar de facturas emitidas y fallbacks para reporte Z nuevo
                fecha_cell = excel_store._cell(row, headers, "Fecha_emision", None) or excel_store._cell(row, headers, "Fecha", None)
                fecha_doc = _parse_row_date(fecha_cell)
                
                if fecha_doc is None or fecha_doc < start_date or fecha_doc > end_date:
                    continue
                
                base_cell = excel_store._cell(row, headers, "Base_imponible", None)
                iva_cell = excel_store._cell(row, headers, "IVA", None)
                
                base = excel_store._parse_monto_cell(base_cell) or Decimal("0")
                iva = excel_store._parse_monto_cell(iva_cell) or Decimal("0")
                
                total_base += base
                total_iva += iva
                total_count += 1
            wb.close()
        except Exception as e:
            logger.exception("Error leyendo ventas desde %s", path)

    return total_base, total_iva, total_count


def get_purchases_totals(start_date: date, end_date: date) -> tuple[Decimal, Decimal, int]:
    """
    Calcula totales de compras (Base Imponible, IVA, número de registros)
    a partir de los Excels mensuales de retenciones emitidas (RETEN-EMIT-*.xlsx)
    en el rango dado, de acuerdo a la instrucción de obtener el crédito fiscal
    directamente de este libro en lugar del Excel de facturas recibidas.
    """
    total_base = Decimal("0")
    total_iva = Decimal("0")
    total_count = 0
    
    base_dir = config.RETENCIONES_EMITIDAS_DIR
    if not base_dir.is_dir():
        return total_base, total_iva, total_count
        
    for path in base_dir.glob("RETEN-EMIT-*.xlsx"):
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            headers = excel_store._headers_index(ws)
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                fecha_cell = excel_store._cell(row, headers, "Fecha_emision", None)
                fecha_doc = _parse_row_date(fecha_cell)
                
                if fecha_doc is None or fecha_doc < start_date or fecha_doc > end_date:
                    continue
                
                base_cell = excel_store._cell(row, headers, "Base_imponible_total", None)
                iva_cell = excel_store._cell(row, headers, "IVA_total", None)
                
                base = excel_store._parse_monto_cell(base_cell) or Decimal("0")
                iva = excel_store._parse_monto_cell(iva_cell) or Decimal("0")
                
                total_base += base
                total_iva += iva
                total_count += 1
            wb.close()
        except Exception as e:
            logger.exception("Error leyendo compras (crédito fiscal) desde %s", path)
            
    return total_base, total_iva, total_count


def get_withholdings_received_totals(start_date: date, end_date: date) -> tuple[Decimal, int]:
    """
    Calcula el total de retenciones de IVA recibidas de clientes en el rango dado.
    Se extraen de RETEN-REC.xlsx (config.EXCEL_PATH).
    """
    total_ret = Decimal("0")
    total_count = 0
    
    path = config.EXCEL_PATH
    if not path.exists():
        return total_ret, total_count
        
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        headers = excel_store._headers_index(ws)
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            fecha_cell = excel_store._cell(row, headers, "Fecha_emision", None)
            fecha_doc = _parse_row_date(fecha_cell)
            
            if fecha_doc is None or fecha_doc < start_date or fecha_doc > end_date:
                continue
            
            ret_cell = excel_store._cell(row, headers, "IVA_retenido", None)
            ret = excel_store._parse_monto_cell(ret_cell) or Decimal("0")
            
            total_ret += ret
            total_count += 1
        wb.close()
    except Exception as e:
        logger.exception("Error leyendo retenciones recibidas desde %s", path)
        
    return total_ret, total_count


def get_withholdings_issued_totals(start_date: date, end_date: date) -> tuple[Decimal, int]:
    """
    Calcula el total de retenciones de IVA emitidas a proveedores en el rango dado.
    Busca los Excels mensuales en config.RETENCIONES_EMITIDAS_DIR.
    """
    total_ret = Decimal("0")
    total_count = 0
    
    base_dir = config.RETENCIONES_EMITIDAS_DIR
    if not base_dir.is_dir():
        return total_ret, total_count
        
    # Buscar todos los libros de retenciones emitidas del directorio
    for path in base_dir.glob("RETEN-EMIT-*.xlsx"):
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            headers = excel_store._headers_index(ws)
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue
                fecha_cell = excel_store._cell(row, headers, "Fecha_emision", None)
                fecha_doc = _parse_row_date(fecha_cell)
                
                if fecha_doc is None or fecha_doc < start_date or fecha_doc > end_date:
                    continue
                
                ret_cell = excel_store._cell(row, headers, "IVA_retenido_total", None)
                ret = excel_store._parse_monto_cell(ret_cell) or Decimal("0")
                
                total_ret += ret
                total_count += 1
            wb.close()
        except Exception as e:
            logger.exception("Error leyendo retenciones emitidas desde %s", path)
            
    return total_ret, total_count


def get_seniat_due_date(year: int, month: int, fortnight: int) -> date:
    """
    Calcula la fecha de declaración y pago del SENIAT para RIF terminado en 3.
    Utiliza el Calendario Fiscal Oficial del SENIAT para el año 2026 (Gaceta Oficial 43.273).
    """
    # Calendario RIF terminado en 3 para el año 2026
    # Tabla 1: "Entre el día 01 y 15" (1ra Quincena, se declara el mismo mes)
    CALENDARIO_Q1_2026 = {
        1: 30,  # Enero
        2: 18,  # Febrero
        3: 23,  # Marzo
        4: 30,  # Abril
        5: 22,  # Mayo
        6: 18,  # Junio
        7: 23,  # Julio
        8: 18,  # Agosto
        9: 21,  # Septiembre
        10: 23, # Octubre
        11: 23, # Noviembre
        12: 28, # Diciembre
    }

    # Tabla 2: "Entre el día 16 y último" (2da Quincena, se declara el mes de declaración)
    # Mapeado por el mes de declaración (mes siguiente al de las operaciones)
    CALENDARIO_Q2_2026 = {
        1: 16,  # Enero
        2: 12,  # Febrero
        3: 4,   # Marzo
        4: 16,  # Abril
        5: 7,   # Mayo
        6: 10,  # Junio
        7: 7,   # Julio
        8: 5,   # Agosto
        9: 2,   # Septiembre
        10: 7,  # Octubre
        11: 9,  # Noviembre
        12: 11, # Diciembre
    }

    if fortnight == 1:
        # Se declara en el mismo mes de las operaciones
        day = CALENDARIO_Q1_2026.get(month, 22)
        return date(year, month, day)
    else:
        # Se declara en el mes siguiente
        if month == 12:
            decl_year = year + 1
            decl_month = 1
        else:
            decl_year = year
            decl_month = month + 1
        
        day = CALENDARIO_Q2_2026.get(decl_month, 5)
        return date(decl_year, decl_month, day)


def get_compromiso_tributario_report(year: int, month: int, fortnight: int) -> dict[str, object]:
    """
    Genera un informe completo estructurado de todos los compromisos tributarios
    de la quincena indicada.
    """
    start_date, end_date = get_fortnight_range(year, month, fortnight)
    
    # 1. Ventas
    v_base, v_iva, v_count = get_sales_totals(start_date, end_date)
    
    # 2. Compras
    c_base, c_iva, c_count = get_purchases_totals(start_date, end_date)
    
    # 3. Retenciones Recibidas
    ret_rec, ret_rec_count = get_withholdings_received_totals(start_date, end_date)
    
    # 4. Retenciones Emitidas
    ret_emi, ret_emi_count = get_withholdings_issued_totals(start_date, end_date)
    
    # 5. Cálculos netos de IVA por pagar
    # IVA a pagar = IVA Débito (Ventas) - IVA Crédito (Compras) - Retenciones Recibidas (Clientes)
    iva_neto = v_iva - c_iva - ret_rec
    
    # 6. Anticipo de ISLR
    anticipo_islr = v_base * ALICUOTA_ANTICIPO_ISLR
    
    # 7. Total general de compromisos de la quincena
    # IVA Neto + Retenciones Emitidas a Enterar + Anticipo de ISLR
    # Nota: Si el IVA Neto es negativo (excedente), para el pago al fisco se toma como 0.00
    pago_iva = max(Decimal("0"), iva_neto)
    total_pagos = pago_iva + ret_emi + anticipo_islr
    
    # 8. Fecha límite del SENIAT
    due_date = get_seniat_due_date(year, month, fortnight)
    
    return {
        "year": year,
        "month": month,
        "fortnight": fortnight,
        "start_date": start_date,
        "end_date": end_date,
        "due_date": due_date,
        "ventas_base": v_base,
        "ventas_iva": v_iva,
        "ventas_count": v_count,
        "compras_base": c_base,
        "compras_iva": c_iva,
        "compras_count": c_count,
        "retenciones_recibidas": ret_rec,
        "retenciones_recibidas_count": ret_rec_count,
        "retenciones_emitidas": ret_emi,
        "retenciones_emitidas_count": ret_emi_count,
        "iva_neto_pagar": iva_neto,
        "iva_neto_pagar_efectivo": pago_iva,
        "anticipo_islr": anticipo_islr,
        "anticipo_islr_alicuota": ALICUOTA_ANTICIPO_ISLR,
        "total_compromisos_a_pagar": total_pagos,
    }
