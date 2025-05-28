import xml.etree.ElementTree as ET
import csv
import os

def parse_cfdi(xml_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Define namespaces (these might vary slightly)
        ns = {
            'cfdi': 'http://www.sat.gob.mx/cfd/4', # o /3 (CFDI 3.3) o /3.3
            'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
        }

        # For CFDI 4.0, namespaces might be different.
        # Example: if root.tag is '{http://www.sat.gob.mx/cfd/4}Comprobante', then ns['cfdi'] is correct.
        # Adjust if you're using CFDI 3.3.

        # --- Extracting data ---
        # If namespaces are not automatically handled by your ET version for findall,
        # you might need to prefix them like: root.find('cfdi:Emisor', ns)

        # Check for CFDI version (example for finding root tag and its namespace)
        # This part can be complex due to different CFDI versions.
        # Assuming CFDI 4.0 for simplicity in path definitions here.
        # You might need to adapt paths like 'cfdi:Complemento/tfd:TimbreFiscalDigital'

        comprobante_tag = root.tag
        if 'cfd/3' in comprobante_tag: # CFDI 3.3
             ns['cfdi'] = 'http://www.sat.gob.mx/cfd/3'
        elif 'cfd/4' in comprobante_tag: # CFDI 4.0
             ns['cfdi'] = 'http://www.sat.gob.mx/cfd/4'
        else: # Older or unknown, default to common one or raise error
             ns['cfdi'] = 'http://www.sat.gob.mx/cfd/3' # Default or handle
             # print(f"Warning: Unknown CFDI version for {xml_file}")


        emisor_node = root.find('cfdi:Emisor', ns)
        receptor_node = root.find('cfdi:Receptor', ns)
        conceptos_node = root.find('cfdi:Conceptos', ns)
        timbre_node = root.find('cfdi:Complemento', ns).find('tfd:TimbreFiscalDigital', ns) if root.find('cfdi:Complemento', ns) is not None else None


        proveedor_rfc = emisor_node.get('Rfc') if emisor_node is not None else ''
        proveedor_nombre = emisor_node.get('Nombre') if emisor_node is not None else ''

        # Asumiendo un solo concepto para simplificar, o puedes iterar
        primer_concepto = conceptos_node.find('cfdi:Concepto', ns) if conceptos_node is not None else None
        descripcion = primer_concepto.get('Descripcion') if primer_concepto is not None else ''
        cantidad = float(primer_concepto.get('Cantidad', 0)) if primer_concepto is not None else 0
        valor_unitario = float(primer_concepto.get('ValorUnitario', 0)) if primer_concepto is not None else 0

        subtotal = float(root.get('SubTotal', 0))
        total = float(root.get('Total', 0))
        fecha = root.get('Fecha')
        folio_factura = root.get('Folio') # Folio interno, puede no estar
        serie_factura = root.get('Serie') # Serie interna, puede no estar
        uuid = timbre_node.get('UUID') if timbre_node is not None else ''

        # Extracción de impuestos (simplificado, puede ser más complejo)
        # Deberás sumar los IVAs, por ejemplo.
        iva_total = 0
        impuestos_node = root.find('cfdi:Impuestos', ns)
        if impuestos_node is not None:
            traslados_node = impuestos_node.find('cfdi:Traslados', ns)
            if traslados_node is not None:
                for traslado in traslados_node.findall('cfdi:Traslado', ns):
                    if traslado.get('Impuesto') == '002': # IVA
                        iva_total += float(traslado.get('Importe', 0))

        return {
            'Descripcion Gasto': descripcion,
            'Producto (Nombre)': descripcion, # O mapear a un producto específico
            'Cantidad': cantidad,
            'Precio Unitario': valor_unitario, # Precio antes de IVA para el concepto
            'Proveedor (Nombre)': proveedor_nombre,
            'Proveedor (RFC)': proveedor_rfc, # Usar para buscar/crear partner en Odoo
            'Fecha Factura': fecha.split('T')[0] if fecha else '', # Solo fecha
            'Numero Factura': f"{serie_factura}{folio_factura}" if serie_factura or folio_factura else uuid, # O solo folio si existe
            'Folio Fiscal (UUID)': uuid,
            'Total': total,
            'IVA (Monto)': iva_total,
            'Subtotal (Antes de Impuestos)': subtotal
            # Añade más campos según la plantilla de importación de Odoo Gastos
        }
    except Exception as e:
        print(f"Error procesando {xml_file}: {e}")
        return None

# --- Script Principal ---
xml_folder = r'C:\Users\axelg\Documents\CFiles\PBScripts\MisFacturasCFDI'
output_csv = 'gastos_para_odoo.csv'

# Define los encabezados del CSV según la plantilla de importación de Odoo
# Esto es un ejemplo, ajústalo!
# Puedes obtener los encabezados exactos exportando un gasto desde Odoo.
# Algunos campos comunes para importar gastos (Expense en inglés):
# 'name' (Descripción), 'product_id/name' (Producto), 'unit_amount' (Precio Unitario), 
# 'quantity', 'date', 'partner_id/name' (Proveedor), 'reference' (Número Factura),
# 'tax_ids/name' (Impuestos), 'analytic_account_id/name' (Cuenta Analítica)

fieldnames = [
    'name', # Descripción del gasto
    'product_id/name', # Nombre del producto/servicio
    'unit_amount', # Precio Unitario (costo del producto para el gasto)
    'quantity',
    'date', # Fecha del gasto
    'partner_id/name', # Nombre del Proveedor
    'partner_id/vat', # RFC del Proveedor (para buscar o crear)
    'reference', # Número de Factura o Folio Fiscal
    'l10n_mx_edi_cfdi_uuid', # Para el UUID en localización mexicana
    'amount_total', # Total (informativo, Odoo lo recalcula)
    # Considera cómo manejar los impuestos. Odoo puede calcularlos si el producto tiene impuestos configurados
    # o puedes especificar los impuestos por su nombre o ID.
]

all_data = []
for filename in os.listdir(xml_folder):
    if filename.endswith('.xml'):
        xml_path = os.path.join(xml_folder, filename)
        data = parse_cfdi(xml_path)
        if data:
            # Mapear los datos extraídos a los nombres de columna de Odoo
            odoo_expense_record = {
                'name': data['Descripcion Gasto'],
                'product_id/name': data['Producto (Nombre)'], # Asegúrate que este producto exista en Odoo o permita creación
                'unit_amount': data['Precio Unitario'], # O el total si es un gasto simple
                'quantity': data['Cantidad'] if data['Cantidad'] > 0 else 1,
                'date': data['Fecha Factura'],
                'partner_id/name': data['Proveedor (Nombre)'],
                'partner_id/vat': data['Proveedor (RFC)'],
                'reference': data['Numero Factura'] if data['Numero Factura'] != data['Folio Fiscal (UUID)'] else '', # Si es distinto al UUID
                'l10n_mx_edi_cfdi_uuid': data['Folio Fiscal (UUID)'],
                'amount_total': data['Total'] # Odoo usualmente recalcula esto basado en precio unitario, cantidad e impuestos.
            }
            all_data.append(odoo_expense_record)

if all_data:
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_data)
    print(f"Archivo CSV '{output_csv}' generado con {len(all_data)} registros.")
else:
    print("No se procesaron datos.")