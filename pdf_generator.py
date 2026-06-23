"""
Generador de cotizaciones PDF — Ingram Micro IBM S&S
Replica exactamente el formato del PDF de referencia:
- Logo Ingram Micro (izquierda) + Logo IBM Distributor (derecha) en cada página
- Tabla info: Canal, Cliente Final, Fabricante, Fecha / Solicitante, Celular, Mail
- Resumen Propuesta Económica
- Nota de descuento
- Detalle de incentivos
- Condiciones generales + página 2 con texto legal completo
"""

import os
import re
from datetime import datetime, timedelta
from typing import List, Dict, Any

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, Image as RLImage
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus.flowables import HRFlowable

# ── Rutas de logos ────────────────────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
INGRAM_LOGO = os.path.join(_HERE, 'ingram_logo.png')
IBM_LOGO    = os.path.join(_HERE, 'ibm_logo.png')

# ── Paleta ────────────────────────────────────────────────────────────────────
C_DARK   = colors.HexColor('#003366')
C_BLUE   = colors.HexColor('#0066CC')
C_RED    = colors.HexColor('#CC0000')
C_LBLUE  = colors.HexColor('#EBF3FF')
C_LGRAY  = colors.HexColor('#F8FAFC')
C_MGRAY  = colors.HexColor('#CBD5E1')
C_DGRAY  = colors.HexColor('#64748B')
C_TEXT   = colors.HexColor('#1A2B3C')
C_WHITE  = colors.white
C_BLACK  = colors.black
C_BORDER = colors.HexColor('#C0C0C0')

PAGE_W, PAGE_H = letter
MARGIN_L = MARGIN_R = 0.6 * inch
MARGIN_T = MARGIN_B = 0.5 * inch
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R

# ── Helpers ───────────────────────────────────────────────────────────────────
def _extract_reseller_name(s):
    m = re.match(r'^\d+\s+(.+)', str(s or ''))
    return m.group(1).strip() if m else str(s or '')

def _extract_customer_name(s):
    m = re.match(r'^\d+-(.+?)(?:,|$)', str(s or ''))
    return m.group(1).strip() if m else str(s or '')

def _usd(v):
    try:    return f'$ {float(v):,.2f}'
    except: return str(v)

def _pct(v):
    try:    return f'{float(v):.0f}%'
    except: return str(v)

def _parse_cov(cov, part):
    if ' To ' in str(cov):
        raw = str(cov).split(' To ')[part].strip()
        try:
            from datetime import datetime as dt
            d = dt.strptime(raw, '%d %b %Y')
            return d.strftime('%-d/%m/%Y')
        except:
            return raw
    return ''

def _ts(*cmds):
    return TableStyle(list(cmds))

def _ps(name, **kw):
    defaults = dict(fontName='Helvetica', fontSize=9, textColor=C_TEXT,
                    leading=12, spaceAfter=0, spaceBefore=0)
    defaults.update(kw)
    return ParagraphStyle(name, **defaults)

# ── Estilos ───────────────────────────────────────────────────────────────────
def _styles():
    return {
        'sec_red':   _ps('sec_red',  fontName='Helvetica-Bold', fontSize=9, textColor=C_RED),
        'th':        _ps('th',       fontName='Helvetica-Bold', fontSize=7.5,
                         textColor=C_WHITE, alignment=TA_CENTER, leading=10),
        'td':        _ps('td',       fontSize=7.5, leading=10),
        'td_c':      _ps('td_c',     fontSize=7.5, leading=10, alignment=TA_CENTER),
        'td_r':      _ps('td_r',     fontSize=7.5, leading=10, alignment=TA_RIGHT),
        'td_dark':   _ps('td_dark',  fontName='Helvetica-Bold', fontSize=8,
                         textColor=C_WHITE, alignment=TA_RIGHT, leading=10),
        'td_dark_l': _ps('td_dark_l',fontName='Helvetica-Bold', fontSize=8,
                         textColor=C_WHITE, alignment=TA_LEFT, leading=10),
        'cond':      _ps('cond',     fontSize=7.5, textColor=C_TEXT, leading=11),
    }

# ── Bloque de logos ───────────────────────────────────────────────────────────
def _logo_header():
    ingram_img = RLImage(INGRAM_LOGO, width=1.8*inch, height=0.38*inch)
    ibm_img    = RLImage(IBM_LOGO,    width=0.9*inch, height=0.53*inch)
    tbl = Table(
        [[ingram_img, '', ibm_img]],
        colWidths=[2.0*inch, 3.5*inch, 1.2*inch]
    )
    tbl.setStyle(_ts(
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN',        (2,0), (2,0),   'RIGHT'),
        ('LEFTPADDING',  (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING',   (0,0), (-1,-1), 0),
        ('BOTTOMPADDING',(0,0), (-1,-1), 0),
    ))
    return tbl

# ── Tabla de info ─────────────────────────────────────────────────────────────
def _info_table(reseller, customer, fecha, solicitante, celular, mail):
    def lbl(txt):
        return Paragraph(f'<b>{txt}</b>', _ps('lbl2', fontName='Helvetica-Bold',
                                               fontSize=8, textColor=C_TEXT))
    def val(txt, link=False):
        if link and txt:
            return Paragraph(f'<font color="#0066CC">{txt}</font>',
                             _ps('v2', fontSize=8, textColor=C_BLUE))
        return Paragraph(str(txt or ''), _ps('v2', fontSize=8, textColor=C_TEXT))

    data = [
        [lbl('Canal:'),         val(reseller),   lbl('Solicitante:'), val(solicitante)],
        [lbl('Cliente Final:'), val(customer),   lbl('Celular:'),     val(celular)],
        [lbl('Fabricante:'),    val('IBM - International Business Mach'),
                                                 lbl('Mail:'),        val(mail, link=True)],
        [lbl('Fecha:'),         val(fecha),      '',                  ''],
    ]
    tbl = Table(data, colWidths=[1.0*inch, 3.0*inch, 0.9*inch, 2.1*inch])
    tbl.setStyle(_ts(
        ('BOX',          (0,0), (-1,-1), 0.8, C_BORDER),
        ('INNERGRID',    (0,0), (-1,-1), 0.3, C_MGRAY),
        ('TOPPADDING',   (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 5),
        ('LEFTPADDING',  (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN',         (2,3), (3,3)),
    ))
    return tbl

# ── Tabla principal de líneas ─────────────────────────────────────────────────
def _tabla_lineas(renewals, bp_margin_percent, s):
    hdrs   = ['Fecha de Inicio', 'Fecha de Fin', 'Parte Número', 'Descripción',
              'Cantidad', 'Precio Cliente\nFinal (MEP)', 'Margen Canal',
              'V/unit Canal', 'Costo Total\nCanal']
    col_w  = [0.65*inch, 0.65*inch, 0.8*inch, 2.1*inch,
              0.45*inch, 0.8*inch, 0.55*inch, 0.7*inch, 0.8*inch]
    rows   = [[Paragraph(h, s['th']) for h in hdrs]]
    total_mep = total_canal = 0.0

    for item in renewals:
        mep    = float(item.get('extended_price', 0))
        qty    = int(item.get('quantity', 1))
        cc     = float(item.get('costo_canal', round(mep * (1 - bp_margin_percent/100), 2)))
        v_unit = round(cc / qty, 2) if qty else 0
        cov    = str(item.get('coverage_dates', ''))
        inicio = _parse_cov(cov, 0) or str(item.get('renewal_line_item_start_date', ''))
        fin    = _parse_cov(cov, 1) or str(item.get('renewal_due_date', ''))

        rows.append([
            Paragraph(inicio,  s['td']),
            Paragraph(fin,     s['td']),
            Paragraph(str(item.get('part_number', '')), s['td']),
            Paragraph(str(item.get('part_description', '')), s['td']),
            Paragraph(str(qty), s['td_c']),
            Paragraph(_usd(mep),    s['td_r']),
            Paragraph(_pct(bp_margin_percent), s['td_c']),
            Paragraph(_usd(v_unit), s['td_r']),
            Paragraph(_usd(cc),     s['td_r']),
        ])
        total_mep   += mep
        total_canal += cc

    rows.append([
        Paragraph('', s['td']), Paragraph('', s['td']), Paragraph('', s['td']),
        Paragraph('TOTAL:', s['td_dark_l']),
        Paragraph('', s['td']),
        Paragraph(_usd(total_mep),   s['td_dark']),
        Paragraph('', s['td']), Paragraph('', s['td']),
        Paragraph(_usd(total_canal), s['td_dark']),
    ])

    n   = len(rows)
    tbl = Table(rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_ts(
        ('BACKGROUND',    (0,0),  (-1,0),  C_BLACK),
        ('TOPPADDING',    (0,0),  (-1,0),  5),
        ('BOTTOMPADDING', (0,0),  (-1,0),  5),
        ('ROWBACKGROUNDS',(0,1),  (-1,n-2),[C_WHITE, C_LGRAY]),
        ('TOPPADDING',    (0,1),  (-1,n-2),4),
        ('BOTTOMPADDING', (0,1),  (-1,n-2),4),
        ('BACKGROUND',    (0,-1), (-1,-1), C_BLACK),
        ('TOPPADDING',    (0,-1), (-1,-1), 5),
        ('BOTTOMPADDING', (0,-1), (-1,-1), 5),
        ('LEFTPADDING',   (0,0),  (-1,-1), 5),
        ('RIGHTPADDING',  (0,0),  (-1,-1), 5),
        ('VALIGN',        (0,0),  (-1,-1), 'MIDDLE'),
        ('GRID',          (0,0),  (-1,-1), 0.3, C_MGRAY),
        ('BOX',           (0,0),  (-1,-1), 0.8, C_BORDER),
    ))
    return tbl, total_mep, total_canal

# ── Nota de descuento ─────────────────────────────────────────────────────────
def _nota_descuento(renewals):
    total_nd = sum(
        float(l.get('cvr_renew_bp_amt', 0)) + float(l.get('cvr_ontime_bp_amt', 0))
        for l in renewals
    )
    data = [[
        Paragraph('Nota de descuento', _ps('nd_l', fontName='Helvetica-Bold',
                                            fontSize=8, textColor=C_TEXT)),
        Paragraph(_usd(total_nd), _ps('nd_r', fontName='Helvetica-Bold',
                                       fontSize=8, textColor=C_TEXT, alignment=TA_RIGHT)),
    ]]
    tbl = Table(data, colWidths=[6.2*inch, 1.1*inch])
    tbl.setStyle(_ts(
        ('BOX',          (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING',   (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 5),
        ('LEFTPADDING',  (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
    ))
    return tbl, total_nd

# ── Tabla de incentivos ───────────────────────────────────────────────────────
def _tabla_incentivos(renewals, cvr_renew_bp, cvr_ontime_bp, s):
    hdrs  = ['Parte Número', 'Descripción', 'Cantidad',
             'Total\nIncentivo', 'On time', 'Renew']
    col_w = [0.85*inch, 3.0*inch, 0.55*inch, 1.0*inch, 0.85*inch, 0.85*inch]
    rows  = [[Paragraph(h, s['th']) for h in hdrs]]
    total_inc = total_ontime = total_renew = 0.0

    for item in renewals:
        ontime = float(item.get('cvr_ontime_bp_amt', 0))
        renew  = float(item.get('cvr_renew_bp_amt',  0))
        inc    = ontime + renew
        rows.append([
            Paragraph(str(item.get('part_number', '')), s['td']),
            Paragraph(str(item.get('part_description', '')), s['td']),
            Paragraph(str(item.get('quantity', 1)), s['td_c']),
            Paragraph(_usd(inc),    s['td_r']),
            Paragraph(_usd(ontime) if ontime else '$ —', s['td_r']),
            Paragraph(_usd(renew),  s['td_r']),
        ])
        total_inc    += inc
        total_ontime += ontime
        total_renew  += renew

    rows.append([
        Paragraph('', s['td']),
        Paragraph('Total', s['td_dark_l']),
        Paragraph('', s['td']),
        Paragraph(_usd(total_inc),    s['td_dark']),
        Paragraph(_usd(total_ontime), s['td_dark']),
        Paragraph(_usd(total_renew),  s['td_dark']),
    ])

    n   = len(rows)
    tbl = Table(rows, colWidths=col_w, repeatRows=1)
    tbl.setStyle(_ts(
        ('BACKGROUND',    (0,0),  (-1,0),  C_BLACK),
        ('TOPPADDING',    (0,0),  (-1,0),  5),
        ('BOTTOMPADDING', (0,0),  (-1,0),  5),
        ('ROWBACKGROUNDS',(0,1),  (-1,n-2),[C_WHITE, C_LGRAY]),
        ('TOPPADDING',    (0,1),  (-1,n-2),4),
        ('BOTTOMPADDING', (0,1),  (-1,n-2),4),
        ('BACKGROUND',    (0,-1), (-1,-1), C_BLACK),
        ('TOPPADDING',    (0,-1), (-1,-1), 5),
        ('BOTTOMPADDING', (0,-1), (-1,-1), 5),
        ('LEFTPADDING',   (0,0),  (-1,-1), 5),
        ('RIGHTPADDING',  (0,0),  (-1,-1), 5),
        ('VALIGN',        (0,0),  (-1,-1), 'MIDDLE'),
        ('GRID',          (0,0),  (-1,-1), 0.3, C_MGRAY),
        ('BOX',           (0,0),  (-1,-1), 0.8, C_BORDER),
    ))
    return tbl

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
    solicitante: str = '',
    celular: str = '',
    mail: str = '',
):
    customer_name = _extract_customer_name(customer_name) or customer_name
    first     = renewals[0] if renewals else {}
    reseller  = _extract_reseller_name(first.get('reseller', ''))
    sbid      = quote_number or first.get('quote_number', '')
    today     = datetime.now().strftime('%-d/%m/%Y')
    valid_dt  = (datetime.now() + timedelta(days=30)).strftime('%d %B de %Y')

    doc = SimpleDocTemplate(
        output_path, pagesize=letter,
        rightMargin=MARGIN_R, leftMargin=MARGIN_L,
        topMargin=MARGIN_T,   bottomMargin=MARGIN_B,
    )
    s     = _styles()
    story = []

    # ── PÁGINA 1 ──────────────────────────────────────────────────────────────
    story.append(_logo_header())
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width='100%', thickness=0.5, color=C_MGRAY))
    story.append(Spacer(1, 8))

    story.append(_info_table(reseller, customer_name, today, solicitante, celular, mail))
    story.append(Spacer(1, 10))

    story.append(Paragraph('DESCRIPCIÓN DE LA SOLUCIÓN', s['sec_red']))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        f'A continuación, se comparten precios aprobados bajo SBID {sbid} '
        f'para el cliente {customer_name}',
        _ps('desc', fontSize=8.5, textColor=C_TEXT, leading=12)
    ))
    story.append(Spacer(1, 10))

    story.append(Paragraph('RESUMEN PROPUESTA ECONÓMICA', s['sec_red']))
    story.append(Spacer(1, 4))

    tbl_lineas, total_mep, total_canal = _tabla_lineas(renewals, bp_margin_percent, s)
    story.append(tbl_lineas)
    story.append(Spacer(1, 2))

    nd_tbl, total_nd = _nota_descuento(renewals)
    story.append(nd_tbl)
    story.append(Spacer(1, 8))

    story.append(Paragraph(
        'Los precios aquí expresados son los precios a los que <b>INGRAM MICRO SAS</b> facturará al Canal. '
        'El precio al Cliente debe ser definido por el Canal o por el fabricante en su propuesta '
        'según las condiciones del negocio.',
        _ps('nota', fontSize=7.5, textColor=C_TEXT, leading=11)
    ))
    story.append(Spacer(1, 10))

    cond_box = Table(
        [[Paragraph('Condiciones Generales de Cotización – INGRAM MICRO S.A.S.',
                    _ps('cb', fontName='Helvetica-Bold', fontSize=8.5,
                        textColor=C_TEXT, alignment=TA_CENTER))]],
        colWidths=[CONTENT_W]
    )
    cond_box.setStyle(_ts(
        ('BOX',          (0,0), (-1,-1), 0.8, C_BORDER),
        ('TOPPADDING',   (0,0), (-1,-1), 7),
        ('BOTTOMPADDING',(0,0), (-1,-1), 7),
        ('LEFTPADDING',  (0,0), (-1,-1), 8),
    ))
    story.append(cond_box)
    story.append(Spacer(1, 8))

    pct_mg = (total_nd / total_canal * 100) if total_canal else 0
    story.append(Paragraph(
        f'<b>Margen transaccional aproximado incluyendo nota de descuento: {pct_mg:.1f}%</b>',
        _ps('mg', fontSize=8.5, textColor=C_TEXT, leading=12)
    ))
    story.append(Paragraph(
        '<font color="#CC0000"><b>Detalle de incentivos:</b></font>',
        _ps('inc_hdr', fontSize=8.5, leading=12)
    ))
    story.append(Spacer(1, 6))
    story.append(_tabla_incentivos(renewals, cvr_renew_bp, cvr_ontime_bp, s))
    story.append(Spacer(1, 14))

    condiciones = [
        'Incentivos sujetos a cambio de programa Partner Plus de IBM.',
        f'Los precios enviados en esta propuesta están aprobados por IBM bajo el SBID {sbid}.',
        'El manejo de los SBID está regulado por lo establecido en el BPA firmado con IBM.',
        'Para procesar este SBID, Ingram Micro podrá solicitar la Orden de Compra (OC) enviada por el cliente final.',
        'Las notas de descuento se generan de manera independiente; se emiten excluidas de IVA a TRM fecha de emisión.',
        'Los precios están expresados en USD y sujetos a cambios según las políticas de IBM.',
        'Se respetará el Precio Máximo para el Usuario Final (MEP), según lo indicado por IBM.',
        'Por favor recuerde que se factura a la TRM fecha de emisión de la factura.',
        'Un SBID puede ser revocado o ajustado en caso de cambios en las políticas de IBM.',
        'Los valores aquí mencionados no incluyen IVA; dicho impuesto será cargado en la factura.',
        f'Propuesta válida hasta {valid_dt}.',
    ]
    for cond in condiciones:
        story.append(Paragraph(cond, _ps('ci', fontSize=7.5, textColor=C_TEXT, leading=11)))
        story.append(Spacer(1, 1))

    # ── PÁGINA 2 ──────────────────────────────────────────────────────────────
    story.append(PageBreak())
    story.append(_logo_header())
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width='100%', thickness=0.5, color=C_MGRAY))
    story.append(Spacer(1, 10))

    legal = [
        ('bold',   'El presente servicio corresponde a la renovación de Software Subscription and Support (S&S) de IBM, '
                   'el cual proporciona al cliente los siguientes beneficios durante el período contratado:'),
        ('bullet', 'Acceso a soporte técnico de IBM durante el horario laboral para resolver dudas de instalación, uso o problemas técnicos.'),
        ('bullet', 'Atención 24/7 para incidentes críticos (Severidad 1).'),
        ('bullet', 'Actualizaciones, parches y correcciones de seguridad para el software licenciado.'),
        ('bullet', 'Costos de mantenimiento predecibles, facilitando la planeación financiera.'),
        ('bullet', 'Posibilidad de extender el soporte hasta por 3 años desde la compra inicial.'),
        ('normal', ''),
        ('bold',   'Exclusiones:'),
        ('normal', 'Este servicio no cubre actividades de diseño o desarrollo de aplicaciones, soporte en ambientes '
                   'no autorizados por IBM, ni incidentes causados por software o hardware no perteneciente a IBM.'),
        ('normal', ''),
        ('bold',   'Condiciones comerciales:'),
        ('normal', 'Los precios enviados en esta propuesta son precios referenciales, están dados en dólares y '
                   'sujetos a cambio dependiendo de las políticas del Fabricante.'),
        ('normal', 'Por favor recuerde facturamos a la TRM de la facturación de Ingram Micro.'),
        ('normal', 'No incluyen IVA.'),
        ('normal', ''),
        ('normal', 'Como Asociado de Negocios de IBM, usted ha solicitado, a través de Ingram Micro, un precio '
                   'especial para una oportunidad de negocio específica con el Usuario Final indicado en esta '
                   'cotización, a través del proceso de Special Bid.'),
        ('normal', ''),
        ('normal', 'Al aceptar esta cotización, usted asume la responsabilidad de trasladar íntegramente el '
                   'beneficio financiero del Special Bid al Usuario Final y garantizar que el precio facturado no '
                   'exceda el Precio Máximo para el Usuario Final ("Maximum End User Price") estipulado en esta cotización.'),
        ('normal', ''),
        ('normal', 'Esta oferta de precio especial es válida únicamente durante el plazo indicado y, en caso de '
                   'aceptación, quedará sujeta a los términos y condiciones de esta cotización y al Contrato de '
                   'Asociado de Negocios IBM que usted tenga suscrito con IBM.'),
        ('normal', ''),
        ('normal', 'Además, usted reconoce y acepta que el Maximum End User Price no podrá verse afectado en '
                   'las facturas al Usuario Final por cargos adicionales no relacionados directamente con la '
                   'transacción, tales como costos generales de ventas, infraestructura operativa u otros gastos '
                   'indirectos. Cualquier cargo de este tipo que incremente el precio final será considerado un '
                   'sobreprecio, constituyendo una violación al acuerdo de Special Bid y al Contrato de Asociado '
                   'de Negocios IBM.'),
        ('normal', ''),
        ('normal', 'Asimismo, usted acepta que Ingram Micro tiene el derecho de verificar el cumplimiento de '
                   'estas condiciones. En caso de que se detecte un incumplimiento, o si no se proporciona '
                   'información confiable y precisa para validar la correcta aplicación del precio aprobado, '
                   'Ingram Micro podrá reclamarle la diferencia entre el Precio aprobado para el Usuario Final '
                   'y el precio efectivamente facturado.'),
        ('normal', ''),
        ('normal', 'El no proporcionar la información solicitada por Ingram Micro para acreditar el cumplimiento '
                   'del Special Bid podrá resultar en la invalidación de la oferta especial y en la recuperación, '
                   'por parte de Ingram Micro, de cualquier descuento aplicado.'),
        ('normal', ''),
        ('bold',   'Su organización es responsable de:'),
        ('numbered','1. Asegurar que el nombre del cliente final corresponde a quien se presenta esta oferta.'),
        ('numbered','2. Asegurar que se cobrará como valor máximo el valor indicado en el MEP (Maximum End User Price).'),
        ('numbered','3. Que la Ruta de Mercado se encuentra aprobada. Contrato BPA activo, Entidades financieras, '
                    'si hay un tercero debe quedar documentado por medio de un Transaccional Agreement o el proceso que corresponda.'),
        ('numbered','4. Que la política de Tasa de Cambio (TRM) que utilice con su cliente esté alineada con la '
                    'entregada en esta propuesta o que utilice una fuente autorizada para estimar dicha Tasa de Cambio.'),
        ('normal', ''),
        ('normal', 'Cordialmente,'),
    ]

    for tipo, txt in legal:
        if not txt:
            story.append(Spacer(1, 4))
            continue
        if tipo == 'bold':
            story.append(Paragraph(txt, _ps('l_b', fontName='Helvetica-Bold',
                                            fontSize=7.5, textColor=C_TEXT, leading=11)))
        elif tipo == 'bullet':
            story.append(Paragraph(f'• {txt}', _ps('l_bul', fontSize=7.5,
                                                     textColor=C_TEXT, leading=11, leftIndent=12)))
        elif tipo == 'numbered':
            story.append(Paragraph(txt, _ps('l_num', fontSize=7.5,
                                             textColor=C_BLUE, leading=11, leftIndent=12)))
        else:
            story.append(Paragraph(txt, _ps('l_n', fontSize=7.5,
                                            textColor=C_TEXT, leading=11)))
        story.append(Spacer(1, 2))

    doc.build(story)
    return output_path
