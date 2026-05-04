"""
IBM S&S Renewal Module - Ingram Micro
Flask Blueprint — se integra al servidor Flask existente con 2 líneas.

INTEGRACIÓN:
    from ibm_renewals_blueprint import ibm_renewals_bp
    app.register_blueprint(ibm_renewals_bp, url_prefix='/renewals')
"""

from flask import Blueprint, request, jsonify, send_file
import os
import json
import zipfile
from datetime import datetime, date
from collections import defaultdict
import re

from xml_parser import parse_renewal_xml
from planilla_generator import check_coverage_errors, generate_planilla
from excel_generator import generate_internal_excel
from pdf_generator import generate_quote_pdf
from ui import register_ui_route

ibm_renewals_bp = Blueprint('ibm_renewals', __name__)
register_ui_route(ibm_renewals_bp)

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_file_from_request():
    if 'file' not in request.files:
        return None, jsonify({"success": False, "error": "No se envió ningún archivo"}), 400
    f = request.files['file']
    if not f.filename.endswith('.xml'):
        return None, jsonify({"success": False, "error": "El archivo debe ser XML (.xml)"}), 400
    return f, None, None


def _group_by_quote(renewals):
    groups = defaultdict(list)
    for item in renewals:
        quote_num = item.get('quote_number', 'SIN-QUOTE').strip() or 'SIN-QUOTE'
        groups[quote_num].append(item)
    return dict(groups)


def _is_ontime(renewal_due_date_str):
    try:
        for fmt in ('%d %b %Y', '%Y-%m-%d', '%m/%d/%Y'):
            try:
                due = datetime.strptime(renewal_due_date_str.strip(), fmt).date()
                return due >= date.today()
            except ValueError:
                continue
    except Exception:
        pass
    return False


def _apply_margins(renewals, bp_margin, ingram_margin, cvr_renew_bp, cvr_renew_ingram,
                   cvr_ontime_bp, cvr_ontime_ingram):
    for item in renewals:
        mep      = float(item.get('extended_price', 0))
        qty      = float(item.get('quantity', 1)) or 1
        unit_mep = float(item.get('unit_price', 0)) or round(mep / qty, 4)
        ontime   = _is_ontime(str(item.get('renewal_due_date', '')))

        costo_canal           = round(mep * (1 - bp_margin / 100), 2)
        cvr_renew_bp_amt      = round(mep * (cvr_renew_bp / 100), 2)
        cvr_renew_ingram_amt  = round(mep * (cvr_renew_ingram / 100), 2)
        cvr_ontime_bp_amt     = round(mep * (cvr_ontime_bp / 100), 2) if ontime else 0.0
        cvr_ontime_ingram_amt = round(mep * (cvr_ontime_ingram / 100), 2) if ontime else 0.0
        precio_final_bp       = round(costo_canal - cvr_renew_bp_amt - cvr_ontime_bp_amt, 2)
        costo_ingram          = round(costo_canal * (1 - ingram_margin / 100)
                                      - cvr_renew_ingram_amt - cvr_ontime_ingram_amt, 2)

        item.update({
            'mep':                   mep,
            'unit_mep':              unit_mep,
            'costo_canal':           costo_canal,
            'cvr_renew_bp_amt':      cvr_renew_bp_amt,
            'cvr_ontime_bp_amt':     cvr_ontime_bp_amt,
            'cvr_renew_ingram_amt':  cvr_renew_ingram_amt,
            'cvr_ontime_ingram_amt': cvr_ontime_ingram_amt,
            'precio_final_bp':       precio_final_bp,
            'costo_ingram':          costo_ingram,
            'ontime':                ontime,
            'final_price':           precio_final_bp,
            'bp_margin':             bp_margin,
            'ingram_margin':         ingram_margin,
            'cvr_renew_bp':          cvr_renew_bp,
            'cvr_ontime_bp':         cvr_ontime_bp,
        })
    return renewals


def _safe_filename(text):
    return re.sub(r'[^\w\-. ]', '_', str(text)).strip()


def _extract_reseller_name(reseller_str):
    if not reseller_str:
        return 'SIN_RESELLER'
    match = re.match(r'^\d+\s+(.+)', str(reseller_str))
    return match.group(1).strip() if match else str(reseller_str).strip()


def _extract_customer_short(site_str):
    if not site_str:
        return 'SIN_CLIENTE'
    match = re.match(r'^\d+-(.+?)(?:,|$)', str(site_str))
    return match.group(1).strip() if match else str(site_str).strip()


def _build_quote_files(lines, quote_num, bp_margin, ingram_margin,
                       cvr_renew_bp, cvr_renew_ingram,
                       cvr_ontime_bp, cvr_ontime_ingram):
    """Genera PDF + Excel + Planilla para un quote y retorna sus rutas."""
    year           = datetime.now().year
    reseller_name  = _extract_reseller_name(lines[0].get('reseller', ''))
    customer_short = _extract_customer_short(lines[0].get('site', ''))
    customer_full  = lines[0].get('site', 'Cliente')
    base_name      = f"S&S - {customer_short} - {reseller_name} - {year}"

    pdf_path = os.path.join(OUTPUT_DIR, f"{_safe_filename(base_name)}.pdf")
    generate_quote_pdf(lines, customer_full,
                       f"Renovación S&S IBM — {customer_short}",
                       bp_margin, pdf_path,
                       quote_number=quote_num,
                       cvr_renew_bp=cvr_renew_bp,
                       cvr_ontime_bp=cvr_ontime_bp)

    excel_path = os.path.join(OUTPUT_DIR, f"{_safe_filename(base_name)}.xlsx")
    generate_internal_excel(lines, quote_num, excel_path,
                            bp_margin=bp_margin,
                            ingram_margin=ingram_margin,
                            cvr_renew_bp=cvr_renew_bp,
                            cvr_renew_ingram=cvr_renew_ingram,
                            cvr_ontime_bp=cvr_ontime_bp,
                            cvr_ontime_ingram=cvr_ontime_ingram,)

    planilla_path = None
    if check_coverage_errors(lines):
        planilla_name = (f"Planilla de renovación SSA y MX - "
                         f"{customer_short} - {reseller_name} - 12 Meses - {year}")
        planilla_path = os.path.join(OUTPUT_DIR, f"{_safe_filename(planilla_name)}.xlsx")
        generate_planilla(lines, quote_num, planilla_path)

    return reseller_name, customer_short, pdf_path, excel_path, planilla_path


# ── Endpoints ─────────────────────────────────────────────────────────────────

@ibm_renewals_bp.get('/health')
def health():
    return jsonify({"status": "ok", "module": "IBM S&S Renewals",
                    "timestamp": datetime.now().isoformat()})


@ibm_renewals_bp.post('/parse-renewal-report')
def parse_renewal_report():
    f, err, code = _get_file_from_request()
    if err:
        return err, code
    try:
        renewals = parse_renewal_xml(f.read())
        groups   = _group_by_quote(renewals)
        quotes   = {}
        for qn, lines in groups.items():
            quotes[qn] = {
                "lines":      lines,
                "customer":   lines[0].get('site', 'Cliente'),
                "total_ibm":  sum(float(l.get('extended_price', 0)) for l in lines),
                "line_count": len(lines),
                "ontime":     _is_ontime(str(lines[0].get('renewal_due_date', ''))),
            }
        coverage_alerts = {}
        for qn, qdata in quotes.items():
            errors = check_coverage_errors(qdata['lines'])
            if errors:
                coverage_alerts[qn] = [{
                    'part_number': e['part_number'],
                    'months':      e['months'],
                    'start':       e['start_str'],
                    'end':         e['end_str'],
                    'correct_end': e['correct_end'],
                } for e in errors]

        return jsonify({"success": True, "total_lines": len(renewals),
                        "total_quotes": len(groups), "quotes": quotes,
                        "coverage_alerts": coverage_alerts})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@ibm_renewals_bp.post('/generate-planilla')
def generate_planilla_endpoint():
    f, err, code = _get_file_from_request()
    if err:
        return err, code
    quote_filter = request.form.get('quote_filter', '').strip() or None
    try:
        renewals = parse_renewal_xml(f.read())
        groups   = _group_by_quote(renewals)
        lines    = groups.get(quote_filter, renewals) if quote_filter else renewals
        timestamp   = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(OUTPUT_DIR,
                        f"Planilla_IBM_{_safe_filename(quote_filter or 'all')}_{timestamp}.xlsx")
        result = generate_planilla(lines, quote_filter or 'all', output_path)
        if not result:
            return jsonify({"success": False,
                            "error": "No hay errores de coverage en este quote"}), 400
        return send_file(output_path,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name=os.path.basename(output_path))
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@ibm_renewals_bp.post('/generate-quote')
def generate_quote():
    """Genera ZIP para un quote individual."""
    f, err, code = _get_file_from_request()
    if err:
        return err, code
    try:
        bp_margin         = float(request.form.get('bp_margin', 3))
        ingram_margin     = float(request.form.get('ingram_margin', 2.1))
        cvr_renew_bp      = float(request.form.get('cvr_renew_bp', 2))
        cvr_renew_ingram  = float(request.form.get('cvr_renew_ingram', 2))
        cvr_ontime_bp     = float(request.form.get('cvr_ontime_bp', 2))
        cvr_ontime_ingram = float(request.form.get('cvr_ontime_ingram', 1))
    except (ValueError, TypeError):
        return jsonify({"success": False, "error": "Los márgenes deben ser números"}), 400

    quote_filter = request.form.get('quote_filter', '').strip() or None

    try:
        renewals = parse_renewal_xml(f.read())
        if not renewals:
            return jsonify({"success": False,
                            "error": "No se encontraron líneas de renovación"}), 400

        renewals = _apply_margins(renewals, bp_margin, ingram_margin,
                                  cvr_renew_bp, cvr_renew_ingram,
                                  cvr_ontime_bp, cvr_ontime_ingram)
        groups   = _group_by_quote(renewals)
        if quote_filter and quote_filter in groups:
            groups = {quote_filter: groups[quote_filter]}

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        year      = datetime.now().year
        entries   = []

        for quote_num, lines in groups.items():
            reseller_name, customer_short, pdf_p, excel_p, planilla_p = \
                _build_quote_files(lines, quote_num, bp_margin, ingram_margin,
                                   cvr_renew_bp, cvr_renew_ingram,
                                   cvr_ontime_bp, cvr_ontime_ingram)
            entries.append((reseller_name, customer_short, pdf_p, excel_p, planilla_p))

        q0       = list(groups.keys())[0]
        zip_name = (f"Quote_{_safe_filename(q0)}_{timestamp}.zip"
                    if len(entries) == 1
                    else f"Cotizaciones_IBM_SS_{year}_{timestamp}.zip")
        zip_path = os.path.join(OUTPUT_DIR, zip_name)

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for reseller, customer_short, pdf_p, excel_p, planilla_p in entries:
                folder = f"{year}/{reseller}/"
                zf.write(pdf_p,   folder + os.path.basename(pdf_p))
                zf.write(excel_p, folder + os.path.basename(excel_p))
                if planilla_p:
                    zf.write(planilla_p, folder + os.path.basename(planilla_p))

        return send_file(zip_path, mimetype='application/zip',
                         as_attachment=True, download_name=zip_name)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@ibm_renewals_bp.post('/generate-bulk')
def generate_bulk():
    """
    Genera un ZIP maestro con todos los quotes chequeados en la UI.
    Estructura interna del ZIP:
        {CANAL}/{AÑO}/  →  PDF + Excel + Planilla (si aplica)

    Recibe:
        file           — el XML original (igual que los otros endpoints)
        quote_numbers  — JSON array con los quote numbers a incluir
        quote_params   — JSON object con parámetros individuales por quote
                         (si el usuario ajustó márgenes en el editor)
        bp_margin, ingram_margin, cvr_* — valores globales de fallback
    """
    f, err, code = _get_file_from_request()
    if err:
        return err, code

    # Parámetros globales (fallback)
    try:
        bp_margin         = float(request.form.get('bp_margin', 3))
        ingram_margin     = float(request.form.get('ingram_margin', 2.1))
        cvr_renew_bp      = float(request.form.get('cvr_renew_bp', 2))
        cvr_renew_ingram  = float(request.form.get('cvr_renew_ingram', 2))
        cvr_ontime_bp     = float(request.form.get('cvr_ontime_bp', 2))
        cvr_ontime_ingram = float(request.form.get('cvr_ontime_ingram', 1))
    except (ValueError, TypeError):
        return jsonify({"success": False, "error": "Márgenes inválidos"}), 400

    # Lista de quotes chequeados y sus parámetros individuales
    try:
        quote_numbers = json.loads(request.form.get('quote_numbers', '[]'))
        quote_params  = json.loads(request.form.get('quote_params',  '{}'))
    except Exception:
        return jsonify({"success": False,
                        "error": "Formato inválido en quote_numbers o quote_params"}), 400

    if not quote_numbers:
        return jsonify({"success": False,
                        "error": "No se enviaron quotes para generar"}), 400

    try:
        xml_content  = f.read()
        renewals_all = parse_renewal_xml(xml_content)
        groups       = _group_by_quote(renewals_all)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    year      = datetime.now().year
    zip_name  = f"IBM_SS_Renovaciones_{timestamp}_{len(quote_numbers)}quotes.zip"
    zip_path  = os.path.join(OUTPUT_DIR, zip_name)

    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for qn in quote_numbers:
                if qn not in groups:
                    continue

                lines = list(groups[qn])  # copia para no mutar el original

                # Parámetros específicos del quote (ajustados en el editor)
                p          = quote_params.get(str(qn), {})
                q_bp       = float(p.get('bp',           bp_margin))
                q_ingram   = float(p.get('ingram',        ingram_margin))
                q_cvr_r    = float(p.get('cvrRenew',      cvr_renew_bp))
                q_cvr_ot   = float(p.get('cvrOntime',     cvr_ontime_bp))
                q_cvr_ri   = float(p.get('cvrRenewIng',   cvr_renew_ingram))
                q_cvr_oi   = float(p.get('cvrOntimeIng',  cvr_ontime_ingram))

                # Si useRenew/useOntime fue desmarcado en el editor, usar 0
                if p.get('useRenew') is False:
                    q_cvr_r  = 0.0
                    q_cvr_ri = 0.0
                if p.get('useOntime') is False:
                    q_cvr_ot = 0.0
                    q_cvr_oi = 0.0

                lines = _apply_margins(lines, q_bp, q_ingram,
                                       q_cvr_r, q_cvr_ri,
                                       q_cvr_ot, q_cvr_oi)

                reseller_name, customer_short, pdf_p, excel_p, planilla_p = \
                    _build_quote_files(lines, qn, q_bp, q_ingram,
                                       q_cvr_r, q_cvr_ri,
                                       q_cvr_ot, q_cvr_oi)

                # Carpeta: {CANAL}/{AÑO}/
                folder = f"{_safe_filename(reseller_name)}/{year}/"
                zf.write(pdf_p,   folder + os.path.basename(pdf_p))
                zf.write(excel_p, folder + os.path.basename(excel_p))
                if planilla_p:
                    zf.write(planilla_p, folder + os.path.basename(planilla_p))

        return send_file(zip_path, mimetype='application/zip',
                         as_attachment=True, download_name=zip_name)

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
