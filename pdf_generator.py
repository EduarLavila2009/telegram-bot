"""Generador premium de archivos PDF para Cotizaciones y Notas de Entrega de SUFEVICA.
El diseño emula con alta precisión el formato clásico de factura Forma Libre provisto en la imagen de Allprot.
"""

from pathlib import Path
from decimal import Decimal
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

def format_amount(amount, currency):
    symbol = "$" if currency == "usd" else "Bs."
    try:
        val = float(amount)
        formatted = f"{val:,.2f}"
        # Convertir formato inglés (1,234.56) a formato venezolano (1.234,56)
        formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{symbol} {formatted}"
    except Exception:
        return f"{symbol} {amount}"

def generate_document_pdf(doc_data: dict, out_path: Path) -> Path:
    # Configuración del documento en tamaño carta vertical con márgenes sutiles de 36pt (0.5 in)
    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    
    # Paleta de colores oficial de SUFEVICA extraída de su logo (azul marino profundo)
    PRIMARY_COLOR = colors.HexColor("#1A1B54")   # Azul marino oficial SUFEVICA
    SECONDARY_COLOR = colors.HexColor("#475569") # Gris oscuro para detalles de contacto
    TEXT_DARK = colors.HexColor("#1f2937")       # Carbón oscuro para lectura
    BORDER_COLOR = colors.HexColor("#475569")    # Gris pizarra para líneas divisorias principales
    BG_LIGHT = colors.HexColor("#f8fafc")        # Fondo sutil para recuadros
    
    # Definición de estilos de texto personalizados
    styles.add(ParagraphStyle(
        name='FormaLibreTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        alignment=TA_RIGHT,
        textColor=PRIMARY_COLOR
    ))
    
    styles.add(ParagraphStyle(
        name='HeaderAddress',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=TA_RIGHT,
        textColor=SECONDARY_COLOR
    ))
    
    styles.add(ParagraphStyle(
        name='DocTitleRight',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        alignment=TA_LEFT,
        textColor=PRIMARY_COLOR
    ))
    
    styles.add(ParagraphStyle(
        name='DocMetaRight',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        alignment=TA_LEFT,
        textColor=TEXT_DARK
    ))
    
    styles.add(ParagraphStyle(
        name='ClientFieldLabel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=PRIMARY_COLOR
    ))
    
    styles.add(ParagraphStyle(
        name='ClientFieldValue',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=TEXT_DARK
    ))
    
    styles.add(ParagraphStyle(
        name='TableHeaderFL',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=PRIMARY_COLOR
    ))
    
    styles.add(ParagraphStyle(
        name='TableCellFL',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=TEXT_DARK
    ))
    
    styles.add(ParagraphStyle(
        name='TableCellBoldFL',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=TEXT_DARK
    ))
    
    story = []
    
    # 1. ENCABEZADO SUPERIOR (Logo de SUFEVICA izquierda | Datos de Contacto y logo derecha)
    # Cargar el logo compuesto (Icono + Texto, Texto es 50% más grande que el Icono)
    logo_icono_path = Path(__file__).resolve().parent / "logo_icono.jpg"
    logo_texto_path = Path(__file__).resolve().parent / "logo_texto.jpg"
    
    if logo_icono_path.exists() and logo_texto_path.exists():
        # Proporciones:
        # Icono: 636 x 598 (~1.06). Base: Ancho = 42pt, Alto = 39.5pt
        # Texto: 709 x 309 (~2.29). Ancho es 50% más grande que el icono (63pt) y se incrementa en otro 50% a petición del usuario -> 94.5pt
        # Alto correspondiente para el texto con su proporción: 94.5 / 2.29 = 41.25pt
        img_icono = Image(str(logo_icono_path), width=42, height=39.5)
        img_texto = Image(str(logo_texto_path), width=94.5, height=41.25)
        
        # Colocar ambos en una sub-tabla horizontal limpia para alinearlos perfectamente
        logo_subtable_data = [[img_icono, img_texto]]
        logo_element = Table(logo_subtable_data, colWidths=[46, 100])
        logo_element.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
        ]))
    else:
        logo_element = Paragraph("<b>SUMINISTROS FERRETEROS VITTORIA, C.A.</b>", ParagraphStyle(name='FallbackTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, textColor=PRIMARY_COLOR))
        
    address_text = """Av. Juan de Urpín, CC Res. Vittoria III, Edif. F, Nivel PB, Local Nro. 2<br/>
    Barcelona - Edo. Anzoátegui<br/>
    Teléfono: 0424-890.61.68<br/>
    E-mail: sufevica@gmail.com
    """
    p_address = Paragraph(address_text.replace('\n', ''), styles['HeaderAddress'])
    
    header_table_data = [
        [logo_element, "", p_address]
    ]
    # Ancho disponible: 540pt. Logo (200), Separador (10), Dirección (330)
    header_table = Table(header_table_data, colWidths=[200, 10, 330])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    
    story.append(header_table)
    story.append(Spacer(1, 15))
    
    # 2. RECUADRO DE DATOS DE VENTA (Cliente izquierda | Metadatos Documento derecha)
    # Extraer campos de datos
    client = doc_data.get("client", {})
    client_name = client.get("name", "—")
    client_rif = client.get("rif", "—")
    client_address = client.get("address", "—")
    client_phone = client.get("phone", "—")
    client_salesman = client.get("salesman", "FREDDY LOPEZ")
    client_saletype = client.get("saleType", "Contado")
    
    doc_type = doc_data.get("docType", "cotizacion")
    title_text = "COTIZACIÓN Nº" if doc_type == "cotizacion" else "NOTA DE ENTREGA Nº"
    doc_num = doc_data.get("docNumber", "S/N")
    
    doc_date_raw = doc_data.get("docDate", "")
    try:
        dt = datetime.strptime(doc_date_raw, "%Y-%m-%d")
        doc_date = dt.strftime("%d/%m/%Y")
    except Exception:
        doc_date = doc_date_raw
        
    # Lado Izquierdo: Datos de Cliente en rejilla sutil
    left_client_data = [
        [
            Paragraph("Cliente", styles['ClientFieldLabel']),
            Paragraph(client_name, styles['ClientFieldValue']),
            Paragraph("Vendedor", styles['ClientFieldLabel']),
            Paragraph(client_salesman, styles['ClientFieldValue'])
        ],
        [
            Paragraph("Dirección", styles['ClientFieldLabel']),
            Paragraph(client_address, styles['ClientFieldValue']),
            "", "" # Span
        ],
        [
            Paragraph("R.I.F/C.I", styles['ClientFieldLabel']),
            Paragraph(client_rif, styles['ClientFieldValue']),
            Paragraph("Tipo Venta", styles['ClientFieldLabel']),
            Paragraph(client_saletype, styles['ClientFieldValue'])
        ],
        [
            Paragraph("Telefonos", styles['ClientFieldLabel']),
            Paragraph(client_phone, styles['ClientFieldValue']),
            "", "" # Vacío
        ]
    ]
    # Ancho total sección izquierda: 370pt. Etiquetas (60), Valores (125), Vendedor Et (65), Vendedor Val (120)
    left_client_table = Table(left_client_data, colWidths=[55, 130, 65, 120])
    left_client_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('SPAN', (1,1), (3,1)), # Dirección abarca todo
        ('SPAN', (1,3), (3,3)), # Teléfonos abarca resto de fila
    ]))
    
    # Lado Derecho: Recuadro del Documento
    right_doc_data = [
        [Paragraph(f"<b>{title_text}</b>", styles['DocTitleRight'])],
        [Paragraph(f"<b>{doc_num}</b>", ParagraphStyle(name='MetaNumHighlight', parent=styles['DocMetaRight'], fontSize=11, textColor=PRIMARY_COLOR))],
        [Paragraph(f"FECHA: {doc_date}", styles['DocMetaRight'])]
    ]
    right_doc_table = Table(right_doc_data, colWidths=[140])
    right_doc_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('LINEBELOW', (0,0), (-1,0), 0.5, BORDER_COLOR),
        ('LINEBELOW', (0,1), (-1,1), 0.5, BORDER_COLOR),
    ]))
    
    # Contenedor principal de datos (combina izquierda y derecha con línea vertical divisoria)
    outer_card_data = [
        [left_client_table, "", right_doc_table]
    ]
    # Ancho total: 540pt. Izquierda (375), Línea Divisora (5), Derecha (160)
    outer_card_table = Table(outer_card_data, colWidths=[370, 10, 160])
    outer_card_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOX', (0,0), (-1,-1), 1, BORDER_COLOR), # Línea de borde externo negro/gris
        ('LINEBEFORE', (2,0), (2,-1), 1, BORDER_COLOR), # Línea vertical que divide cliente de metadatos
    ]))
    
    story.append(outer_card_table)
    story.append(Spacer(1, 15))
    
    # 3. TABLA DE PRODUCTOS (Estilo Forma Libre: Líneas horizontales limpias, sin verticales)
    currency = doc_data.get("currency", "usd")
    rate = float(doc_data.get("exchangeRate", 550.00)) if currency == "ves" else 1.0
    cur_symbol = "$" if currency == "usd" else "Bs."
    
    col_headers = [
        Paragraph("<b>Codigo</b>", styles['TableHeaderFL']),
        Paragraph("<b>Descripción</b>", styles['TableHeaderFL']),
        Paragraph("<b>Cantidad</b>", ParagraphStyle(name='THCant', parent=styles['TableHeaderFL'], alignment=TA_CENTER)),
        Paragraph(f"<b>Precio ({cur_symbol})</b>", ParagraphStyle(name='THPrecio', parent=styles['TableHeaderFL'], alignment=TA_RIGHT)),
        Paragraph(f"<b>Total ({cur_symbol})</b>", ParagraphStyle(name='THTotal', parent=styles['TableHeaderFL'], alignment=TA_RIGHT))
    ]
    
    table_items_data = [col_headers]
    
    subtotal_usd = 0.0
    items = doc_data.get("items", [])
    
    for it in items:
        code = it.get("code", "") or "—"
        desc = it.get("desc", "")
        qty = float(it.get("qty", 1.0))
        price_usd = float(it.get("priceUsd", 0.0))
        
        price_conv = price_usd * rate
        total_conv = qty * price_conv
        
        subtotal_usd += (qty * price_usd)
        
        # Formatear montos quitando el símbolo monetario de las celdas ya que está en la cabecera
        price_formatted = f"{price_conv:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        total_formatted = f"{total_conv:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        row = [
            Paragraph(code, styles['TableCellBoldFL']),
            Paragraph(desc, styles['TableCellFL']),
            Paragraph(f"{qty:g}", ParagraphStyle(name='TDCant', parent=styles['TableCellFL'], alignment=TA_CENTER)),
            Paragraph(price_formatted, ParagraphStyle(name='TDPrecio', parent=styles['TableCellFL'], alignment=TA_RIGHT)),
            Paragraph(total_formatted, ParagraphStyle(name='TDTotal', parent=styles['TableCellBoldFL'], alignment=TA_RIGHT))
        ]
        table_items_data.append(row)
        
    # Anchos de columnas en pulgadas (Total 540pt): Código (80), Descripción (260), Cantidad (50), Precio (75), Total (75)
    products_table = Table(table_items_data, colWidths=[80, 260, 50, 75, 75])
    
    products_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
        # Líneas de estilo de Forma Libre idénticas al modelo:
        ('LINEABOVE', (0,0), (-1,0), 1, BORDER_COLOR),  # Línea arriba de la cabecera
        ('LINEBELOW', (0,0), (-1,0), 1, BORDER_COLOR),  # Línea abajo de la cabecera
        ('LINEBELOW', (0,-1), (-1,-1), 1, BORDER_COLOR), # Línea abajo de la última fila de productos
    ]))
    
    story.append(products_table)
    story.append(Spacer(1, 15))
    
    # 4. TOTALES & FIRMA DE RECIBIDO CONFORME
    total_conv = subtotal_usd * rate
    
    # Formatear montos sin el símbolo aquí porque el recuadro muestra la unidad
    total_raw = f"{total_conv:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    # Texto informativo de la tasa si es VES
    rate_text = ""
    if currency == "ves":
        rate_text = f"Tasa de cambio: Bs. {rate:,.2f}  |  "
        
    # Recuadro de Totales de la derecha: solo muestra el TOTAL $ o TOTAL Bs
    label_currency = "Bs" if currency == "ves" else "$"
    totales_box_data = [
        [Paragraph(f"<b>TOTAL {label_currency}</b>", styles['ClientFieldLabel']), Paragraph(total_raw, ParagraphStyle(name='TotFL', parent=styles['TableCellBoldFL'], fontName='Helvetica-Bold', fontSize=10, alignment=TA_RIGHT))]
    ]
    
    totales_box_table = Table(totales_box_data, colWidths=[110, 80])
    totales_box_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('BOX', (0,0), (-1,-1), 1, BORDER_COLOR),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#e2e8f0")), # Destacado completo
    ]))
    
    # Firma "Recibi Conforme" a la izquierda en el mismo nivel
    p_recibi = Paragraph("<br/><br/><b>Recibi Conforme</b>  ________________________________________", ParagraphStyle(name='RecibiFL', parent=styles['Normal'], fontName='Helvetica', fontSize=9, textColor=PRIMARY_COLOR))
    
    bottom_row_data = [
        [p_recibi, "", totales_box_table]
    ]
    # Ancho total: 540pt. Firma (310), Separador (40), Totales (190)
    bottom_row_table = Table(bottom_row_data, colWidths=[310, 40, 190])
    bottom_row_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    
    story.append(KeepTogether([
        bottom_row_table,
        Spacer(1, 20),
        Paragraph(f"<i>{rate_text}Precios expresados en {'Bolívares (Bs.)' if currency == 'ves' else 'Dólares ($)'}</i>", ParagraphStyle(name='FooterRate', parent=styles['Normal'], fontSize=7.5, fontName='Helvetica-Oblique', textColor=SECONDARY_COLOR)),
        Spacer(1, 15)
    ]))
    
    # 5. PIE DE PÁGINA (Texto en azul en la parte inferior de forma libre)
    default_note = (
        "<b>NOTA:</b> Precios sujetos a cambio sin previo aviso. Esta cotización representa un presupuesto informativo. ¡Gracias por su preferencia!"
        if doc_type == "cotizacion" else
        "<b>NOTA:</b> La mercancía viaja por cuenta y riesgo del cliente. Documento de despacho informativo. ¡Gracias por su preferencia!"
    )
    
    footer_elements = [Spacer(1, 10)]
    p_default = Paragraph(default_note, ParagraphStyle(name='FooterFL', parent=styles['Normal'], fontSize=8, leading=10, alignment=TA_CENTER, textColor=PRIMARY_COLOR))
    footer_elements.append(p_default)
    
    custom_note = doc_data.get("client", {}).get("note")
    if custom_note:
        # Nota adicional en negrilla y con fuente 50% más grande (8 * 1.5 = 12pt)
        p_custom = Paragraph(f"<b>{custom_note}</b>", ParagraphStyle(
            name='CustomNoteFL',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=12,
            leading=15,
            alignment=TA_CENTER,
            textColor=PRIMARY_COLOR
        ))
        footer_elements.append(Spacer(1, 5))
        footer_elements.append(p_custom)
        
    story.append(KeepTogether(footer_elements))
    
    # Construcción final del documento PDF
    doc.build(story)
    return out_path
