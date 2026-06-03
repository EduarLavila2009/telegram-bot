import os
from decimal import Decimal
from datetime import date

from . import config
from . import excel_store
from . import tributario_engine

def test_tributario():
    print("Iniciando pruebas con Directorio raiz para Retenciones Emitidas...")
    
    # 1. Definir fechas de prueba: Mayo 2026 - 1ra quincena (01/05/2026 - 15/05/2026)
    start_date = date(2026, 5, 1)
    end_date = date(2026, 5, 15)
    
    # Limpiar y regenerar archivos vacíos de ventas y retenciones recibidas para la prueba
    # pero guardando copias si ya existieran.
    for path in [config.FACTURAS_EMITIDAS_PATH, config.REPORTES_Z_PATH, config.EXCEL_PATH]:
        if path.exists():
            backup_path = path.with_suffix(".xlsx.bak")
            if not backup_path.exists():
                os.rename(path, backup_path)
                print(f"Respaldado {path.name} a {backup_path.name}")
            else:
                os.remove(path)
    
    # 2. Agregar registros de ventas de prueba (Mayo 2026 Q1)
    print("Insertando ventas de prueba en FACTURAS-EMITIDAS.xlsx y REPORTES-Z.xlsx...")
    
    # Venta 1: Factura emitida
    excel_store.append_venta_record(
        config.FACTURAS_EMITIDAS_PATH,
        clasificacion="Factura Emitida",
        estado="REGISTRADO",
        fecha="05/05/2026",
        numero_documento="000101",
        razon_social="DISTRIBUIDORA DE ACERO, C.A.",
        rif="J-30495831-2",
        base_imponible="10000.00",
        iva="1600.00",
        total="11600.00",
        emisor="SUMINISTROS FERRETEROS VITTORIA (SUFEVICA), C.A.",
        texto_origen="Test manual venta",
    )
    
    # Venta 2: Reporte Z
    excel_store.append_venta_record(
        config.REPORTES_Z_PATH,
        clasificacion="Reporte Z",
        estado="REGISTRADO",
        fecha="10/05/2026",
        numero_documento="000099",
        razon_social="CONSUMIDOR FINAL (VENTAS DIARIAS)",
        rif="V-00000000-0",
        base_imponible="20000.00",
        iva="3200.00",
        total="23200.00",
        emisor="SUMINISTROS FERRETEROS VITTORIA (SUFEVICA), C.A.",
        texto_origen="Test manual Reporte Z",
    )
    
    # 3. Agregar retención de IVA recibida (Mayo 2026 Q1)
    print("Insertando retenciones de clientes recibidas en RETEN-REC.xlsx...")
    excel_store.append_record(
        config.EXCEL_PATH,
        fecha_emision="12/05/2026",
        numero_comprobante="20260500000042",
        rif="J-30495831-2",
        fechas_facturas="05/05/2026",
        numeros_facturas="000101",
        controles_facturas="00-0010",
        total_compra_con_iva="11600.00",
        base_imponible="10000.00",
        iva_retenido="1200.00", # Retenido por cliente al 75%
        ocr_snippet="Registro de prueba retencion recibida",
    )

    # 4. Calcular los totales del motor tributario para 1ra Quincena Mayo 2026
    print("\nEjecutando calculos del motor para 1ra Quincena de Mayo 2026...")
    report = tributario_engine.get_compromiso_tributario_report(2026, 5, 1)
    
    print("\n--- REPORT RESULT ---")
    for k, v in report.items():
        print(f"{k}: {v}")
    print("---------------------\n")
    
    # Restaurar los archivos originales
    for path in [config.FACTURAS_EMITIDAS_PATH, config.REPORTES_Z_PATH, config.EXCEL_PATH]:
        backup_path = path.with_suffix(".xlsx.bak")
        if backup_path.exists():
            if path.exists():
                os.remove(path)
            os.rename(backup_path, path)
            print(f"Restaurado archivo original {path.name}")

if __name__ == "__main__":
    test_tributario()
