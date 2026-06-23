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
from rates_config import get_rates

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
        mep = float(item.get('extended_price', 0))
        qty = float(item.get('quantity', 1)) or 1
        unit_mep = float(item.get('unit_price', 0)) or round(mep / qty, 4)
        ontime = _is_ontime(str(item.get('renewal_due_date', '')))

        costo_canal          = round(mep * (1 - bp_margin / 100), 2)
        cvr_renew_bp_amt     = round(mep * (cvr_renew_bp / 100), 2)
        cvr_renew_ingram_amt = round(mep * (cvr_renew_ingram / 100), 2)
        cvr_ontime_bp_amt    = round(mep * (cvr_ontime_bp / 100), 2) if ontime else 0.0
        cvr_ontime_ingram_amt= round(mep * (cvr_ontime_ingram / 100), 2) if ontime else 0.0
        precio_final_bp      = round(costo_canal - cvr_renew_bp_amt - cvr_ontime_bp_amt, 2)
        costo_ingram         = round(costo_canal * (1 - ingram_margin / 100)
                                     - cvr_renew_ingram_amt - cvr_ontime_ingram_amt, 2)

        item.update({
            'mep': mep, 'unit_mep': unit_mep,
            'costo_canal': costo_canal,
            'cvr_renew_bp_amt': cvr_renew_bp_amt,
            'cvr_ontime_bp_amt': cvr_ontime_bp_amt,
            'cvr_renew_ingram_amt': cvr_renew_ingram_amt,
            'cvr_ontime_ingram_amt': cvr_ontime_ingram_amt,
            'precio_final_bp': precio_final_bp,
            'costo_ingram': costo_ingram,
