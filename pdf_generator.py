"""
Generador de cotizaciones PDF — Ingram Micro IBM S&S
Estructura idéntica a la hoja Cotización del Excel generado:
  1. Encabezado (Canal, Cliente, Fabricante, Fecha, SBID)
  2. Resumen Propuesta Económica (tabla líneas)
  3. Total + Nota de descuento
  4. Detalle de incentivos (OnTime + Renew por línea)
  5. Condiciones generales + firma
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from datetime import datetime, timedelta
from typing import List, Dict, Any
import re

# ── Paleta ────────────────────────────────────────────────────────────────────
C_DARK  = colors.HexColor('#003366')
C_BLUE  = colors.HexColor('#0066CC')
C_LBLUE = colors.HexColor('#EBF3FF')
C_LGRAY = colors.HexColor('#F8FAFC')
C_MGRAY = colors.HexColor('#CBD5E1')
C_DGRAY = colors.HexColor('#64748B')
C_TEXT  = colors.HexColor('#1A2B3C')
C_WHITE = colors.white
C_IBM   = colors.HexColor('#054ADA')
C_GREEN_BG = colors.HexColor('#F0FDF4')
C_GREEN_TX = colors.HexColor('#166534')
C_YELLOW   = colors.HexColor('#FFF9C4')


def _extract_reseller_name(s):
    m = re.match(r'^\d+\s+(.+)', str(s or ''))
    return m.group(1).strip() if m else str(s or '')

def _extract_customer_name(s):
    m = re.match(r'^\d+-(.+?)(?:,|$)', str(s or ''))
    return m.group(1).strip() if m else str(s or '')

def _usd(v):
    try:
        return f'${float(v):,.2f}'
    except:
        return str(v)

def _pct(v):
    try:
        return f'{float(v)*100:.1f}%'
    except:
        return str(v)

def _parse_cov(cov, part):
    if ' To ' in str(cov):
        return str(cov).split(' To ')[part].strip()
    return ''


# ── Estilos ───────────────────────────────────────────────────────────────────
def _styles():
    def ps(name, **kw):
        defaults = dict(fontName='Helvetica', fontSize=9, textColor=C_TEXT,
                        leading=12, spaceAfter=0, spaceBefore=0)
        defaults.update(kw)
        return ParagraphStyle(name, **defaults)

    return {
        'title':     ps('title',    fontName='Helvetica-Bold', fontSize=13,
                        textColor=C_WHITE,  alignment=TA_LEFT),
        'subtitle':  ps('subtitle', fontName='Helvetica-Bold', fontSize=10,
                        textColor=C_WHITE,  alignment=TA_LEFT),
        'sec_hdr':   ps('sec_hdr',  fontName='Helvetica-Bold', fontSize=9,
                        textColor=C_WHITE,  alignment=TA_LEFT),
        'lbl':       ps('lbl',      fontName='Helvetica-Bold', fontSize=8,
                        textColor=C_DGRAY),
        'val':       ps('val',      fontName='Helvetica',      fontSize=9,
                        textColor=C_TEXT),
        'val_b':     ps('val_b',    fontName='Helvetica-Bold', fontSize=9,
                        textColor=C_TEXT),
        'th':        ps('th',       fontName='Helvetica-Bold', fontSize=7.5,
                        textColor=C_WHITE,  alignment=TA_CENTER, leading=10),
        'td':        ps('td',       fontName='Helvetica',      fontSize=7.5,
                        textColor=C_TEXT,   leading=10),
        'td_r':      ps('td_r',     fontName='Helvetica',      fontSize=7.5,
                        textColor=C_TEXT,   alignment=TA_RIGHT, leading=10),
        'td_b':      ps('td_b',     fontName='Helvetica-Bold', fontSize=7.5,
                        textColor=C_TEXT,   alignment=TA_RIGHT, leading=10),
        'td_dark':   ps('td_dark',  fontName='Helvetica-Bold', fontSize=8,
                        textColor=C_WHITE,  alignment=TA_RIGHT, leading=10),
        'td_dark_l': ps('td_dark_l',fontName='Helvetica-Bold', fontSize=8,
                        textColor=C_WHITE,  alignment=TA_LEFT,  leading=10),
        'note':      ps('note',     fontName='Helvetica',      fontSize=7.5,
                        textColor=C_DGRAY,  leading=11),
        'cond':      ps('cond',     fontName='Helvetica',      fontSize=7,
                        textColor=C_DGRAY,  leading=10),
        'firma':     ps('firma',    fontName='Helvetica-Bold', fontSize=8,
                        textColor=C_BLUE),
        'total_lbl': ps('total_lbl',fontName='Helvetica-Bold', fontSize=9,
                        textColor=C_WHITE,  alignment=TA_RIGHT),
        'total_val': ps('total_val',fontName='Helvetica-Bold', fontSize=10,
                        textColor=C_WHITE,  alignment=TA_RIGHT),
        'nd_lbl':    ps('nd_lbl',   fontName='Helvetica',      fontSize=8,
                        textColor=C_DGRAY,  alignment=TA_LEFT),
        'nd_val':    ps('nd_val',   fontName='Helvetica-Bold', fontSize=9,
                        textColor=C_TEXT,   alignment=TA_RIGHT),
    }


def _ts(*cmds):
    return TableStyle(list(cmds))


# ── Sección 1: Encabezado ─────────────────────────────────────────────────────
def _build_header(renewals, customer_name, quote_number, s):
    first   = renewals[0] if renewals else {}
    reseller = _extract_reseller_name(first.get('reseller', ''))
    sbid     = quote_number or first.get('quote_number', '')
    due      = first.get('renewal_due_date', '')
    today    = datetime.now().strftime('%d/%m/%Y')
    valid_dt = (datetime.now() + timedelta(days=30)).strftime('%d %b %Y')

    # Bloque logo + título
    logo_data = [[
        Table([[
            [Paragraph('INGRAM <font color="#003366">MICRO</font>', ParagraphStyle(
                'logo', fontName='Helvetica-Bold', fontSize=20, textColor=C_BLUE))],
            [Paragraph('Technology Solutions', ParagraphStyle(
                'logsub', fontName='Helvetica', fontSize=8, textColor=C_DGRAY))],
            [Spacer(1, 3)],
            [Paragraph('  IBM Authorized Reseller  ', ParagraphStyle(
                'ibmbadge', fontName='Helvetica-Bold', fontSize=7.5,
                textColor=C_WHITE, backColor=C_IBM))],
        ]], colWidths=[2.8*inch]),
        Table([[
            [Paragraph('COTIZACIÓN IBM S&amp;S', ParagraphStyle(
                'cot_hdr', fontName='Helvetica-Bold', fontSize=10,
                textColor=C_DARK, alignment=TA_RIGHT))],
            [Paragraph(f'Fecha: {today}', ParagraphStyle(
                'cot_f', fontName='Helvetica', fontSize=8,
                textColor=C_DGRAY, alignment=TA_RIGHT))],
            [Paragraph(f'SBID / Quote: <b>{sbid}</b>', ParagraphStyle(
                'cot_q', fontName='Helvetica', fontSize=8.5,
                textColor=C_TEXT, alignment=TA_RIGHT))],
            [Paragraph(f'Canal: <b>{reseller}</b>', ParagraphStyle(
                'cot_c', fontName='Helvetica', fontSize=8.5,
                textColor=C_TEXT, alignment=TA_RIGHT))],
            [Paragraph(f'Cliente: <b>{customer_name}</b>', ParagraphStyle(
                'cot_cl', fontName='Helvetica', fontSize=8.5,
                textColor=C_TEXT, alignment=TA_RIGHT))],
        ]], colWidths=[4.0*inch]),
    ]]
    hdr_tbl = Table(logo_data, colWidths=[2.8*inch, 4.0*inch])
    hdr_tbl.setStyle(_ts(
        ('VALIGN',       (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING',  (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING',   (0,0), (-1,-1), 0),
        ('BOTTOMPADDING',(0,0), (-1,-1), 0),
    ))

    # Banda azul con título
    title_tbl = Table(
        [[Paragraph('IBM Software Subscription &amp; Support (S&amp;S) — Renovación', s['subtitle'])]],
        colWidths=[6.8*inch]
    )
    title_tbl.setStyle(_ts(
        ('BACKGROUND',   (0,0), (-1,-1), C_DARK),
        ('TOPPADDING',   (0,0), (-1,-1), 7),
        ('BOTTOMPADDING',(0,0), (-1,-1), 7),
        ('LEFTPADDING',  (0,0), (-1,-1), 10),
    ))

    # Info meta (4 celdas: cliente, canal, moneda, válida hasta)
    meta_data = [[
        _meta_block('CLIENTE',      customer_name,  s),
        _meta_block('CANAL',        reseller,       s),
        _meta_block('MONEDA',       'USD',          s),
        _meta_block('VÁLIDA HASTA', valid_dt,       s),
    ]]
    meta_tbl = Table(meta_data, colWidths=[1.9*inch, 1.9*inch, 1.1*inch, 1.9*inch])
    meta_tbl.setStyle(_ts(
        ('BACKGROUND',   (0,0), (-1,-1), C_LBLUE),
        ('TOPPADDING',   (0,0), (-1,-1), 7),
        ('BOTTOMPADDING',(0,0), (-1,-1), 7),
        ('LEFTPADDING',  (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('VALIGN',       (0,0), (-1,-1), 'TOP'),
        ('LINEAFTER',    (0,0), (2,0), 0.4, C_MGRAY),
    ))

    # Texto SBID
    sbid_txt = Paragraph(
        f'A continuación se comparten precios aprobados bajo SBID <b>{sbid}</b> para el cliente <b>{customer_name}</b>.',
        ParagraphStyle('sbid_p', fontName='Helvetica', fontSize=8.5,
                       textColor=C_TEXT, leading=12)
    )

    return [hdr_tbl, Spacer(1, 6), title_tbl, Spacer(1, 6), meta_tbl, Spacer(1, 8), sbid_txt]


def _meta_block(label, value, s):
    return Table([[
        Paragraph(label, ParagraphStyle('ml', fontName='Helvetica', fontSize=7,
                                        textColor=C_BLUE, spaceAfter=2)),
        Paragraph(value, ParagraphStyle('mv', fontName='Helvetica-Bold', fontSize=8.5,
                                        textColor=C_TEXT)),
    ]], colWidths=[None])


# ── Sección 2: Resumen Propuesta Económica ────────────────────────────────────
def _build_tabla_lineas(renewals, bp_margin_percent, cvr_renew_bp, cvr_ontime_bp, s):
    """
    Idéntico a la hoja Cotización del Excel:
    Fecha Inicio | Fecha Fin | Parte Número | Descripción | Qty |
    Precio Cliente Final (MEP) | Margen Canal | V/unit Canal | Costo Total Canal
    """
    # Encabezado de sección
    sec = Table(
        [[Paragraph('RESUMEN PROPUESTA ECONÓMICA', s['sec_hdr'])]],
        colWidths=[6.8*inch]
    )
    sec.setStyle(_ts(
        ('BACKGROUND',   (0,0), (-1,-1), C_DARK),
        ('TOPPADDING',   (0,0), (-1,-1), 6),
        ('BOTTOMPADDING',(0,0), (-1,-1), 6),
        ('LEFTPADDING',  (0,0), (-1,-1), 10),
    ))

    hdrs = ['Fecha\nInicio', 'Fecha\nFin', 'Parte Número', 'Descripción',
            'Qty', 'Precio Cliente\nFinal (MEP)', 'Margen\nCanal',
            'V/unit\nCanal', 'Costo Total\nCanal']
    col_w = [0.7*inch, 0.7*inch, 0.85*inch, 2.0*inch,
             0.35*inch, 0.85*inch, 0.55*inch, 0.75*inch, 0.9*inch]

    rows = [[Paragraph(h, s['th']) for h in hdrs]]

    total_mep   = 0.0
    total_canal = 0.0

    for item in renewals:
        mep  = float(item.get('extended_price', 0))
        qty  = int(item.get('quantity', 1))
        cc   = float(item.get('costo_canal', round(mep * (1 - bp_margin_percent/100), 2)))
        v_unit = round(cc / qty, 2) if qty else 0

        cov = str(item.get('coverage_dates', ''))
        inicio = _parse_cov(cov, 0) or str(item.get('renewal_line_item_start_date', ''))
        fin    = _parse_cov(cov, 1) or str(item.get('renewal_due_date', ''))

        desc = str(item.get('part_description', ''))

        rows.append([
            Paragraph(inicio, s['td']),
            Paragraph(fin,    s['td']),
            Paragraph(str(item.get('part_number', '')), s['td']),
            Paragraph(desc,   s['td']),
            Paragraph(str(qty), ParagraphStyle('ctr', fontName='Helvetica', fontSize=7.5,
                                               textColor=C_TEXT, alignment=TA_CENTER)),
            Paragraph(_usd(mep),    s['td_r']),
            Paragraph(_pct(bp_margin_percent/100), s['td_r']),
            Paragraph(_usd(v_unit), s['td_r']),
            Paragraph(_usd(cc),     s['td_r']),
        ])
        total_mep   += mep
        total_canal += cc

    # Fila TOTAL
    rows.append([
        Paragraph('', s['td']),
        Paragraph('', s['td']),
        Paragraph('', s['td']),
        Paragraph('TOTAL:', s['td_dark_l']),
        Paragraph('', s['td']),
        Paragraph(_usd(total_mep),   s['td_dark']),
        Paragraph('', s['td']),
        Paragraph('', s['td']),
        Paragraph(_usd(total_canal), s['td_dark']),
    ])

    n = len(rows)
    tbl = Table(rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_ts(
        # Header
        ('BACKGROUND',   (0,0), (-1,0), C_BLUE),
        ('TOPPADDING',   (0,0), (-1,0), 5),
        ('BOTTOMPADDING',(0,0), (-1,0), 5),
        # Data rows alternadas
        ('ROWBACKGROUNDS',(0,1), (-1,n-2), [C_WHITE, C_LGRAY]),
        ('TOPPADDING',   (0,1), (-1,n-2), 4),
        ('BOTTOMPADDING',(0,1), (-1,n-2), 4),
        # Total row
        ('BACKGROUND',   (0,-1), (-1,-1), C_DARK),
        ('TOPPADDING',   (0,-1), (-1,-1), 6),
        ('BOTTOMPADDING',(0,-1), (-1,-1), 6),
        # General
        ('LEFTPADDING',  (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('GRID',         (0,0), (-1,-1), 0.3, C_MGRAY),
        ('LINEABOVE',    (0,-1), (-1,-1), 1.5, C_BLUE),
    ))

    return [sec, Spacer(1, 4), tbl], total_mep, total_canal


# ── Sección 3: Nota de descuento ──────────────────────────────────────────────
def _build_nota_descuento(renewals, total_canal, s):
    total_nd = sum(float(l.get('cvr_renew_bp_amt', 0)) + float(l.get('cvr_ontime_bp_amt', 0))
                   for l in renewals)

    nd_data = [
        [Paragraph('Nota de descuento:', s['nd_lbl']),
         Paragraph(_usd(total_nd), s['nd_val'])],
    ]
    nd_tbl = Table(nd_data, colWidths=[5.9*inch, 0.9*inch])
    nd_tbl.setStyle(_ts(
        ('BACKGROUND',   (0,0), (-1,-1), C_YELLOW),
        ('TOPPADDING',   (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 5),
        ('LEFTPADDING',  (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('BOX',          (0,0), (-1,-1), 0.5, C_MGRAY),
    ))

    nota_txt = Paragraph(
        'Los precios aquí expresados son los precios a los que <b>INGRAM MICRO SAS</b> facturará al Canal. '
        'El precio al Cliente debe ser definido por el Canal o por el fabricante en su propuesta según las condiciones del negocio.',
        s['note']
    )
    return [Spacer(1, 3), nd_tbl, Spacer(1, 6), nota_txt]


# ── Sección 4: Detalle de incentivos ─────────────────────────────────────────
def _build_incentivos(renewals, cvr_renew_bp, cvr_ontime_bp, s):
    total_nd = sum(float(l.get('cvr_renew_bp_amt', 0)) + float(l.get('cvr_ontime_bp_amt', 0))
                   for l in renewals)

    # Margen transaccional aproximado
    total_mep   = sum(float(l.get('extended_price', 0)) for l in renewals)
    total_canal = sum(float(l.get('costo_canal', 0)) for l in renewals)
    pct_mg = (total_nd / total_canal * 100 + (total_mep - total_canal) / total_mep * 100) if total_mep else 0

    mg_txt = Paragraph(
        f'Margen transaccional aproximado incluyendo nota de descuento: <b>{pct_mg:.1f}%</b>  ·  Detalle de incentivos:',
        ParagraphStyle('mg_t', fontName='Helvetica', fontSize=8.5,
                       textColor=C_TEXT, leading=12, spaceBefore=4)
    )

    # Encabezado de sección
    sec = Table(
        [[Paragraph('DETALLE DE INCENTIVOS IBM', s['sec_hdr'])]],
        colWidths=[6.8*inch]
    )
    sec.setStyle(_ts(
        ('BACKGROUND',   (0,0), (-1,-1), C_BLUE),
        ('TOPPADDING',   (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 5),
        ('LEFTPADDING',  (0,0), (-1,-1), 10),
    ))

    hdrs = ['Parte Número', 'Descripción', 'Cantidad',
            f'Total Incentivo\n(CVR {cvr_renew_bp:.0f}%+{cvr_ontime_bp:.0f}%)',
            f'On Time\n({cvr_ontime_bp:.0f}%)',
            f'Renew\n({cvr_renew_bp:.0f}%)']
    col_w = [0.9*inch, 2.95*inch, 0.55*inch, 1.1*inch, 0.85*inch, 0.85*inch]

    rows = [[Paragraph(h, s['th']) for h in hdrs]]

    total_inc   = 0.0
    total_ontime = 0.0
    total_renew  = 0.0

    for item in renewals:
        inc    = float(item.get('cvr_renew_bp_amt', 0)) + float(item.get('cvr_ontime_bp_amt', 0))
        ontime = float(item.get('cvr_ontime_bp_amt', 0))
        renew  = float(item.get('cvr_renew_bp_amt', 0))
        desc   = str(item.get('part_description', ''))

        rows.append([
            Paragraph(str(item.get('part_number', '')), s['td']),
            Paragraph(desc, s['td']),
            Paragraph(str(item.get('quantity', 1)),
                      ParagraphStyle('ctr2', fontName='Helvetica', fontSize=7.5,
                                     textColor=C_TEXT, alignment=TA_CENTER)),
            Paragraph(_usd(inc),    s['td_r']),
            Paragraph(_usd(ontime) if ontime else '—', s['td_r']),
            Paragraph(_usd(renew),  s['td_r']),
        ])
        total_inc    += inc
        total_ontime += ontime
        total_renew  += renew

    # Fila Total
    rows.append([
        Paragraph('', s['td']),
        Paragraph('Total', s['td_dark_l']),
        Paragraph('', s['td']),
        Paragraph(_usd(total_inc),    s['td_dark']),
        Paragraph(_usd(total_ontime), s['td_dark']),
        Paragraph(_usd(total_renew),  s['td_dark']),
    ])

    n = len(rows)
    tbl = Table(rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_ts(
        ('BACKGROUND',   (0,0), (-1,0), C_BLUE),
        ('TOPPADDING',   (0,0), (-1,0), 5),
        ('BOTTOMPADDING',(0,0), (-1,0), 5),
        ('ROWBACKGROUNDS',(0,1), (-1,n-2), [C_WHITE, C_LGRAY]),
        ('TOPPADDING',   (0,1), (-1,n-2), 4),
        ('BOTTOMPADDING',(0,1), (-1,n-2), 4),
        ('BACKGROUND',   (0,-1), (-1,-1), C_DARK),
        ('TOPPADDING',   (0,-1), (-1,-1), 5),
        ('BOTTOMPADDING',(0,-1), (-1,-1), 5),
        ('LEFTPADDING',  (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('GRID',         (0,0), (-1,-1), 0.3, C_MGRAY),
        ('LINEABOVE',    (0,-1), (-1,-1), 1.5, C_BLUE),
    ))

    return [mg_txt, Spacer(1, 5), sec, Spacer(1, 4), tbl]


# ── Sección 5: Condiciones y firma ────────────────────────────────────────────
def _build_condiciones(s):
    condiciones = [
        'Incentivos sujetos a cambio de programa Partner Plus de IBM.',
        'Los precios enviados en esta propuesta están aprobados por IBM bajo el SBID.',
        'El manejo de los SBID está regulado por lo establecido en el BPA firmado con IBM.',
        'Para procesar este SBID, Ingram Micro podrá solicitar la Orden de Compra (OC) enviada por el cliente final.',
        'Las notas de descuento se generan de manera independiente; se emiten excluidas de IVA a TRM fecha de emisión.',
        'Los precios están expresados en USD y sujetos a cambios según las políticas de IBM.',
        'Se respetará el Precio Máximo para el Usuario Final (MEP), según lo indicado por IBM.',
        'Por favor recuerde que se factura a la TRM fecha de emisión de la factura.',
        'Un SBID puede ser revocado o ajustado en caso de cambios en las políticas de IBM.',
        'Los valores aquí mencionados no incluyen IVA; dicho impuesto será cargado en la factura.',
        f'Propuesta válida por 30 días a partir de la fecha de emisión.',
    ]

    sec = Table(
        [[Paragraph('Condiciones Generales de Cotización – INGRAM MICRO S.A.S.', s['sec_hdr'])]],
        colWidths=[6.8*inch]
    )
    sec.setStyle(_ts(
        ('BACKGROUND',   (0,0), (-1,-1), C_DARK),
        ('TOPPADDING',   (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 5),
        ('LEFTPADDING',  (0,0), (-1,-1), 10),
    ))

    items = [sec, Spacer(1, 4)]
    for c in condiciones:
        items.append(Paragraph(f'• {c}', s['cond']))
        items.append(Spacer(1, 1))

    items += [
        Spacer(1, 8),
        HRFlowable(width='100%', thickness=0.5, color=C_MGRAY),
        Spacer(1, 5),
        Paragraph(
            'Angela Herrera · <b>Business Development Manager IBM</b> · '
            'angela.herrera@ingrammicro.com · +57 316 696 91 38 · Ingram Micro Colombia',
            s['firma']
        ),
    ]
    return items


# ── Función principal ─────────────────────────────────────────────────────────
def generate_quote_pdf(
    renewals: List[Dict[str, Any]],
    customer_name: str,
    quote_title: str,
    bp_margin_percent: float,
    output_path: str,
    quote_number: str = None,
    ingram_margin_percent: float = 0.0,
    cvr_renew_bp: float = 0.0,
    cvr_ontime_bp: float = 0.0,
):
    # Si customer_name viene del site completo, limpiarlo
    customer_name = _extract_customer_name(customer_name) or customer_name

    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        rightMargin=0.6*inch,
        leftMargin=0.6*inch,
        topMargin=0.5*inch,
        bottomMargin=0.6*inch,
    )

    s = _styles()
    story = []

    # 1. Encabezado
    story += _build_header(renewals, customer_name, quote_number, s)
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width='100%', thickness=1.5, color=C_BLUE))
    story.append(Spacer(1, 8))

    # 2. Tabla de líneas
    tabla_elements, total_mep, total_canal = _build_tabla_lineas(
        renewals, bp_margin_percent, cvr_renew_bp, cvr_ontime_bp, s)
    story += tabla_elements
    story.append(Spacer(1, 4))

    # 3. Nota de descuento
    story += _build_nota_descuento(renewals, total_canal, s)
    story.append(Spacer(1, 10))

    # 4. Detalle de incentivos
    story += _build_incentivos(renewals, cvr_renew_bp, cvr_ontime_bp, s)
    story.append(Spacer(1, 10))

    # 5. Condiciones y firma
    story += _build_condiciones(s)

    doc.build(story)
