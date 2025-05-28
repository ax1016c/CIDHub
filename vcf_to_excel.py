import vobject
import openpyxl
from openpyxl.styles import Font

def export_vcf_to_excel(vcf_file_path, excel_file_path):
    """
    Exports contacts from a VCF file to an Excel spreadsheet.

    Args:
        vcf_file_path (str): The path to the input VCF file.
        excel_file_path (str): The path to save the output Excel file.
    """
    try:
        # Intenta leer con utf-8, pero si falla, prueba con latin-1 u otro encoding común
        try:
            with open(vcf_file_path, 'r', encoding='utf-8') as f:
                vcf_data = f.read()
        except UnicodeDecodeError:
            print("UTF-8 decoding failed, trying latin-1...")
            with open(vcf_file_path, 'r', encoding='latin-1') as f:
                vcf_data = f.read()
    except FileNotFoundError:
        print(f"Error: Archivo VCF no encontrado en {vcf_file_path}")
        return
    except Exception as e:
        print(f"Error leyendo el archivo VCF: {e}")
        return

    # Crear un nuevo libro de Excel y seleccionar la hoja activa
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Contactos"

    # Definir la fila de encabezado
    headers = [
        "Nombre Completo",
        "Nombre",
        "Apellido",
        "Organización",
        "Cargo",
        "Teléfono (Móvil)",
        "Teléfono (Casa)",
        "Teléfono (Trabajo)",
        "Teléfono (Otro)",
        "Email (Casa)",
        "Email (Trabajo)",
        "Email (Otro)",
        "Dirección (Casa)",
        "Dirección (Trabajo)",
        "Cumpleaños",
        "Notas"
    ]
    sheet.append(headers)

    # Poner el encabezado en negrita
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    processed_contacts = 0
    # Procesar cada vCard en el archivo VCF
    for vcard in vobject.readComponents(vcf_data):
        contact_data = {}

        # Nombre Completo (FN) y Nombre Formateado (N)
        contact_data["Nombre Completo"] = vcard.fn.value if hasattr(vcard, 'fn') else ""
        if hasattr(vcard, 'n'):
            n_obj = vcard.n.value
            contact_data["Nombre"] = n_obj.given if hasattr(n_obj, 'given') else ""
            contact_data["Apellido"] = n_obj.family if hasattr(n_obj, 'family') else ""
        else:
            contact_data["Nombre"] = ""
            contact_data["Apellido"] = ""
            # Intentar derivar de FN si N está ausente
            if contact_data["Nombre Completo"]:
                parts = contact_data["Nombre Completo"].split(" ", 1)
                contact_data["Nombre"] = parts[0]
                if len(parts) > 1:
                    contact_data["Apellido"] = parts[1]

        # Organización (ORG)
        contact_data["Organización"] = vcard.org.value[0] if hasattr(vcard, 'org') and vcard.org.value else ""

        # Cargo (TITLE)
        contact_data["Cargo"] = vcard.title.value if hasattr(vcard, 'title') else ""

        # Números de Teléfono (TEL)
        phone_numbers = {"Móvil": [], "Casa": [], "Trabajo": [], "Otro": []}
        if hasattr(vcard, 'tel_list'):
            for tel in vcard.tel_list:
                tel_type_params = [p.upper() for p in tel.params.get('TYPE', [])]
                number = tel.value
                if "CELL" in tel_type_params:
                    phone_numbers["Móvil"].append(number)
                elif "VOICE" in tel_type_params and not any(pt in ["HOME", "WORK", "CELL"] for pt in tel_type_params): # Voz general, a menudo móvil si no es casa/trabajo
                    phone_numbers["Móvil"].append(number)
                elif "HOME" in tel_type_params:
                    phone_numbers["Casa"].append(number)
                elif "WORK" in tel_type_params:
                    phone_numbers["Trabajo"].append(number)
                else:
                    phone_numbers["Otro"].append(number)
        contact_data["Teléfono (Móvil)"] = "; ".join(phone_numbers["Móvil"])
        contact_data["Teléfono (Casa)"] = "; ".join(phone_numbers["Casa"])
        contact_data["Teléfono (Trabajo)"] = "; ".join(phone_numbers["Trabajo"])
        contact_data["Teléfono (Otro)"] = "; ".join(phone_numbers["Otro"])

        # Emails (EMAIL)
        emails = {"Casa": [], "Trabajo": [], "Otro": []}
        if hasattr(vcard, 'email_list'):
            for email_entry in vcard.email_list:
                email_type_params = [p.upper() for p in email_entry.params.get('TYPE', [])]
                email_address = email_entry.value
                if "HOME" in email_type_params:
                    emails["Casa"].append(email_address)
                elif "INTERNET" in email_type_params and not any(pt in ["WORK"] for pt in email_type_params): # A menudo por defecto o casa
                    emails["Casa"].append(email_address)
                elif "WORK" in email_type_params:
                    emails["Trabajo"].append(email_address)
                else:
                    emails["Otro"].append(email_address)

        contact_data["Email (Casa)"] = "; ".join(emails["Casa"])
        contact_data["Email (Trabajo)"] = "; ".join(emails["Trabajo"])
        contact_data["Email (Otro)"] = "; ".join(emails["Otro"])

        # Direcciones (ADR)
        addresses = {"Casa": [], "Trabajo": []}
        if hasattr(vcard, 'adr_list'):
            for adr in vcard.adr_list:
                adr_type_params = [p.upper() for p in adr.params.get('TYPE', [])]
                adr_parts = [
                    getattr(adr.value, 'box', '') or "",
                    getattr(adr.value, 'extended', '') or "",
                    getattr(adr.value, 'street', '') or "",
                    getattr(adr.value, 'city', '') or "",
                    getattr(adr.value, 'region', '') or "",
                    getattr(adr.value, 'code', '') or "",
                    getattr(adr.value, 'country', '') or ""
                ]
                full_address = ", ".join(filter(None, adr_parts))

                if "HOME" in adr_type_params:
                    addresses["Casa"].append(full_address)
                elif "WORK" in adr_type_params:
                    addresses["Trabajo"].append(full_address)
                else: # Si no hay tipo, o es otro tipo, considerarlo general o añadir a una categoría específica si es necesario
                     addresses["Casa"].append(full_address) # Por defecto a Casa para los no tipificados

        contact_data["Dirección (Casa)"] = "; ".join(addresses["Casa"])
        contact_data["Dirección (Trabajo)"] = "; ".join(addresses["Trabajo"])

        # Cumpleaños (BDAY)
        contact_data["Cumpleaños"] = vcard.bday.value if hasattr(vcard, 'bday') else ""

        # Notas (NOTE)
        contact_data["Notas"] = vcard.note.value if hasattr(vcard, 'note') else ""

        # Añadir fila a la hoja
        row_to_add = [contact_data.get(header, "") for header in headers]
        sheet.append(row_to_add)
        processed_contacts += 1

    # Auto-ajustar columnas para mejor legibilidad
    for col in sheet.columns:
        max_length = 0
        column_letter = col[0].column_letter # Obtener la letra de la columna
        for cell in col:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            except:
                pass
        adjusted_width = (max_length + 2) if max_length > 0 else 12 # Ancho mínimo
        sheet.column_dimensions[column_letter].width = adjusted_width

    # Guardar el libro
    try:
        workbook.save(excel_file_path)
        print(f"¡Éxito! {processed_contacts} contactos exportados a {excel_file_path}")
    except Exception as e:
        print(f"Error guardando el archivo Excel: {e}")

# --- Cómo usar ---
if __name__ == "__main__":
    # Nombre de tu archivo VCF. Asegúrate que esté en la misma carpeta que el script
    # o proporciona la ruta completa, por ejemplo: r"C:\Ruta\A\Tu\Archivo\MICROREC 3DD.vcf"
    vcf_file = "1.vcf"

    # Nombre deseado para tu archivo Excel de salida
    excel_file = "contactos_exportados.xlsx"

    # --- Inicio de la creación del VCF de ejemplo (COMENTADO) ---
    # No necesitas esta parte si estás usando tu propio archivo VCF.
    # dummy_vcf_content = """BEGIN:VCARD
# VERSION:3.0
# N:Doe;John;;;
# FN:John Doe
# ORG:Example Corp.
# TITLE:Engineer
# TEL;TYPE=WORK,VOICE:(111) 555-1212
# TEL;TYPE=HOME,VOICE:(222) 555-1212
# TEL;TYPE=CELL:(333) 555-1212
# ADR;TYPE=WORK:;;123 Main St;Anytown;CA;91234;USA
# ADR;TYPE=HOME:;;456 Oak Ln;Otherville;NY;10001;USA
# EMAIL;TYPE=INTERNET,PREF:john.doe@example.com
# EMAIL;TYPE=HOME:jdoe@example.org
# BDAY:19800101
# NOTE:Un contacto de prueba.
# END:VCARD

# BEGIN:VCARD
# VERSION:3.0
# N:Smith;Jane;;;
# FN:Jane Smith
# ORG:Another Company
# TEL;TYPE=CELL:(444) 555-5678
# EMAIL;TYPE=WORK:jane.smith@another.co
# NOTE:Segundo contacto.
# END:VCARD

# BEGIN:VCARD
# VERSION:3.0
# FN:Solo un Nombre
# TEL:(555) 123-4567
# EMAIL:nombre@dominio.tld
# END:VCARD
# """
    # print(f"Intentando crear un archivo VCF de prueba: {vcf_file} (Esto no debería pasar si estás usando tu propio archivo)")
    # with open(vcf_file, 'w', encoding='utf-8') as f: # ASEGÚRATE QUE ESTO ESTÉ COMENTADO SI USAS TU PROPIO VCF_FILE
    #     f.write(dummy_vcf_content)                   # ASEGÚRATE QUE ESTO ESTÉ COMENTADO
    # print(f"Creado archivo VCF de prueba: {vcf_file} para testeo.") # ASEGÚRATE QUE ESTO ESTÉ COMENTADO
    # --- Fin de la creación del VCF de ejemplo ---

    export_vcf_to_excel(vcf_file, excel_file)