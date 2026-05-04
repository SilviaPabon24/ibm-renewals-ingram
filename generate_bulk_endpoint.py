# Agregar este endpoint en ibm_renewals_blueprint.py
# Pegar antes del último bloque de la función health()

import json as _json

@ibm_renewals_bp.post('/generate-bulk')
def generate_bulk():
    """
    Genera un ZIP maestro con todos los quotes chequeados.
    Estructura: {canal}/{año}/{reseller}/  →  PDF + Excel + Planilla
    """
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
        return jsonify({"success": False, "error": "Márgenes inválidos"}), 400

    # Quotes a incluir
    try:
        quote_numbers = _json.loads(request.form.get('quote_numbers', '[]'))
        quote_params  = _json.loads(request.form.get('quote_params',  '{}'))
    except Exception:
        quote_numbers = []
        quote_params  = {}

    if not quote_numbers:
        return jsonify({"success": False, "error": "No se enviaron quotes para generar"}), 400

    xml_content = f.read()
    try:
        renewals_all = parse_renewal_xml(xml_content)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

    groups = _group_by_quote(renewals_all)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    year      = datetime.now().year

    # ZIP maestro nombrado por fecha
    zip_name = f'IBM_SS_Renovaciones_{timestamp}_{len(quote_numbers)}quotes.zip'
    zip_path = os.path.join(OUTPUT_DIR, zip_name)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for qn in quote_numbers:
            if qn not in groups:
                continue
            lines = groups[qn]

            # Parámetros específicos del quote (si fue ajustado en el editor)
            p = quote_params.get(qn, {})
            q_bp          = float(p.get('bp',          bp_margin))
            q_ingram      = float(p.get('ingram',       ingram_margin))
            q_cvr_renew   = float(p.get('cvrRenew',     cvr_renew_bp))
            q_cvr_ontime  = float(p.get('cvrOntime',    cvr_ontime_bp))
            q_cvr_ri      = float(p.get('cvrRenewIng',  cvr_renew_ingram))
            q_cvr_oi      = float(p.get('cvrOntimeIng', cvr_ontime_ingram))

            lines = _apply_margins(lines, q_bp, q_ingram,
                                   q_cvr_renew, q_cvr_ri,
                                   q_cvr_ontime, q_cvr_oi)

            reseller_name  = _extract_reseller_name(lines[0].get('reseller', ''))
            customer_short = _extract_customer_short(lines[0].get('site', ''))
            customer_full  = lines[0].get('site', 'Cliente')

            base_name = f'S&S - {customer_short} - {reseller_name} - {year}'

            # Carpeta: {canal}/{año}/
            folder = f'{_safe_filename(reseller_name)}/{year}/'

            # PDF
            pdf_path = os.path.join(OUTPUT_DIR, f'{_safe_filename(base_name)}.pdf')
            generate_quote_pdf(lines, customer_full,
                               f'Renovación S&S IBM — {customer_short}',
                               q_bp, pdf_path,
                               quote_number=qn,
                               cvr_renew_bp=q_cvr_renew,
                               cvr_ontime_bp=q_cvr_ontime)
            zf.write(pdf_path, folder + os.path.basename(pdf_path))

            # Excel
            excel_path = os.path.join(OUTPUT_DIR, f'{_safe_filename(base_name)}.xlsx')
            generate_internal_excel(lines, qn, excel_path)
            zf.write(excel_path, folder + os.path.basename(excel_path))

            # Planilla (solo si hay errores de coverage)
            coverage_errors = check_coverage_errors(lines)
            if coverage_errors:
                planilla_name = (f'Planilla de renovación SSA y MX - '
                                 f'{customer_short} - {reseller_name} - 12 Meses - {year}')
                planilla_path = os.path.join(OUTPUT_DIR, f'{_safe_filename(planilla_name)}.xlsx')
                generate_planilla(lines, qn, planilla_path)
                zf.write(planilla_path, folder + os.path.basename(planilla_path))

    return send_file(zip_path, mimetype='application/zip',
                     as_attachment=True, download_name=zip_name)