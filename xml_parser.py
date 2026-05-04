"""
Parser para el reporte XML de IBM Passport Advantage Online (PAO)
Formato: XML for Microsoft Excel (SpreadsheetML)
"""

import xml.etree.ElementTree as ET
from datetime import datetime
from typing import List, Dict, Any


NS = {
    'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
    'o':  'urn:schemas-microsoft-com:office:office',
    'x':  'urn:schemas-microsoft-com:office:excel',
}

# Columnas del reporte PAO (índice basado en 1, según el XML)
COLUMN_MAP = {
    1:  "renewal_due_date",
    2:  "part_number",
    3:  "part_description",
    4:  "quantity",
    5:  "extended_price",
    6:  "currency",
    7:  "quote_number",
    8:  "quote_modified_date",
    9:  "coverage_dates",
    10: "reseller_authorization",
    11: "customer_set_designation",
    12: "select_territory_accelerator",
    13: "eligible_government_program",
    14: "site",
    15: "ibm_customer_number",
    16: "site_contact",
    17: "reseller",
    18: "ibm_renewal_contact",
    19: "full_view_indicator",
    20: "processor_vendor",
    21: "processor_brand",
    22: "processor_description",
    23: "processor_quantity",
    24: "vus_per_processor",
    25: "total_vus",
    26: "pvu_qty_change_description",
    27: "renewal_line_item_start_date",
}

HEADER_KEYWORDS = {
    "renewal due date", "part number", "part description",
    "quantity", "extended price", "currency", "quote number",
    "site", "ibm customer number", "coverage dates"
}


def parse_renewal_xml(content: bytes) -> List[Dict[str, Any]]:
    """
    Parsea el XML SpreadsheetML descargado de IBM PAO.
    Retorna lista de dicts con los datos de cada línea de renovación.
    """
    root = ET.fromstring(content)

    # Buscar el worksheet de renovaciones
    worksheet = None
    for ws in root.findall('.//ss:Worksheet', NS):
        name = ws.get('{urn:schemas-microsoft-com:office:spreadsheet}Name', '')
        if 'renewal' in name.lower() or 'active' in name.lower():
            worksheet = ws
            break

    if worksheet is None:
        # Tomar el primero disponible
        worksheet = root.find('.//ss:Worksheet', NS)

    if worksheet is None:
        raise ValueError("No se encontró ningún Worksheet en el XML")

    table = worksheet.find('.//ss:Table', NS)
    if table is None:
        raise ValueError("No se encontró la tabla de datos")

    rows = table.findall('ss:Row', NS)

    # Encontrar la fila de encabezados
    header_row_idx = None
    for i, row in enumerate(rows):
        cells = row.findall('ss:Cell', NS)
        texts = [_cell_value(c).lower() for c in cells]
        matches = sum(1 for t in texts if any(k in t for k in HEADER_KEYWORDS))
        if matches >= 3:
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("No se encontraron los encabezados del reporte")

    # Parsear encabezados dinámicamente
    header_cells = rows[header_row_idx].findall('ss:Cell', NS)
    headers = []
    col_idx = 0
    for cell in header_cells:
        ss_index = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
        if ss_index:
            col_idx = int(ss_index) - 1
        headers.append((col_idx, _cell_value(cell).strip()))
        col_idx += 1

    # Parsear filas de datos
    renewals = []
    for row in rows[header_row_idx + 1:]:
        cells = row.findall('ss:Cell', NS)
        if not cells:
            continue

        row_data = {}
        col_idx = 0
        for cell in cells:
            ss_index = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
            if ss_index:
                col_idx = int(ss_index) - 1

            value = _cell_value(cell)
            col_idx += 1

            # Buscar el nombre del campo
            for hdr_col, hdr_name in headers:
                if hdr_col == col_idx - 1 and hdr_name:
                    field = _normalize_header(hdr_name)
                    row_data[field] = value
                    break

        # Filtrar filas vacías o de metadatos
        if not row_data:
            continue
        if _is_metadata_row(row_data):
            continue

        # Castear tipos
        row_data = _cast_types(row_data)
        renewals.append(row_data)

    return renewals


def _cell_value(cell) -> str:
    data = cell.find('ss:Data', NS)
    if data is not None and data.text:
        return data.text.strip()
    return ""


def _normalize_header(header: str) -> str:
    return (header.lower()
            .replace(' ', '_')
            .replace('/', '_')
            .replace('®', '')
            .replace('(', '')
            .replace(')', '')
            .strip('_'))


def _is_metadata_row(row: Dict) -> bool:
    values = list(row.values())
    if not any(values):
        return True
    text_values = ' '.join(str(v).lower() for v in values)
    skip_keywords = ['end of report', 'confidential', 'title', 'sort by', 'date range']
    return any(kw in text_values for kw in skip_keywords)


def _cast_types(row: Dict) -> Dict:
    numeric_fields = ['quantity', 'extended_price', 'processor_quantity', 'vus_per_processor', 'total_vus']
    for field in numeric_fields:
        if field in row:
            try:
                val = str(row[field]).replace(',', '').replace('$', '').strip()
                row[field] = float(val) if '.' in val else int(val)
            except (ValueError, TypeError):
                row[field] = 0
    return row
