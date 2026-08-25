#!/usr/bin/env python3
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

xlsx_path = r"C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"

with zipfile.ZipFile(xlsx_path, 'r') as z:
    # Read the workbook to find sheet names
    workbook_xml = z.read('xl/workbook.xml')
    wb_tree = ET.fromstring(workbook_xml)

    # Find Facturas sheet
    ns = {'rel': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    sheets = wb_tree.findall('.//ns:sheet', {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'})

    # Read sheet1.xml (Facturas sheet)
    sheet_xml = z.read('xl/worksheets/sheet1.xml')
    sheet_tree = ET.fromstring(sheet_xml)

    # Parse rows
    ns_sheet = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rows = sheet_tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')

    # Read shared strings for proper cell values
    strings_xml = z.read('xl/sharedStrings.xml')
    strings_tree = ET.fromstring(strings_xml)
    shared_strings = []
    for si in strings_tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
        t_elem = si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
        if t_elem is not None:
            shared_strings.append(t_elem.text or "")

    revisar_providers = defaultdict(lambda: {'count': 0, 'sum': 0})
    total_revisar = 0

    for row in rows:
        cells = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')

        # Build a dict of column letter -> value
        cell_values = {}
        for cell in cells:
            ref = cell.get('r')  # e.g., 'A2', 'B2'
            col_letter = ''.join([c for c in ref if c.isalpha()])
            col_num = ord(col_letter) - ord('A') + 1 if col_letter else 0

            # Get cell value
            v_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            if v_elem is not None:
                v = v_elem.text
                # If type is 's' (shared string), look up the string
                if cell.get('t') == 's':
                    try:
                        cell_values[col_num] = shared_strings[int(v)]
                    except:
                        cell_values[col_num] = v
                else:
                    cell_values[col_num] = v

        # Column mapping (1-indexed):
        # Col 4 = Proveedor, Col 16 = Total, Col 17 = Categoria
        categoria = cell_values.get(17, "")

        if categoria == "REVISAR":
            total_revisar += 1
            proveedor = cell_values.get(4, "UNKNOWN")
            total = float(cell_values.get(16, 0) or 0)

            revisar_providers[proveedor]['count'] += 1
            revisar_providers[proveedor]['sum'] += total

print(f"Total REVISAR: {total_revisar}\n")
print("Distribution by Provider:")
for prov in sorted(revisar_providers.keys(), key=lambda x: revisar_providers[x]['count'], reverse=True)[:25]:
    data = revisar_providers[prov]
    print(f"{prov}: {data['count']} items, ${data['sum']:,.0f} CLP")
