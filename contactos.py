import pandas as pd
import re

# Lista de títulos conocidos para ayudar en la identificación
KNOWN_TITLES_LIST = ["dr.", "dra.", "dr", "dra", "lic.", "ing.", "lic", "ing", "sr.", "sra.", "srta."]
# Normalizar títulos para comparación (minúsculas, sin punto al final)
NORMALIZED_TITLES = {t.lower().rstrip('.') for t in KNOWN_TITLES_LIST}

# Lista de vendedores conocidos (ejemplo)
KNOWN_VENDORS_LIST = ["ami"]
NORMALIZED_VENDORS = {v.lower() for v in KNOWN_VENDORS_LIST}

def is_placeholder_zero(text):
    """Verifica si un texto es un cero placeholder como '0.0', '.0', o '0'."""
    if not isinstance(text, str):
        return False
    return text in ["0.0", ".0", "0"]

def is_phone_like(text):
    """Verifica si un texto parece un número de teléfono."""
    if not isinstance(text, str):
        return False
    # Si el texto termina en ".0" (común si pandas leyó un float), quitarlo
    if text.endswith(".0"):
        text = text[:-2]
    cleaned_text = re.sub(r'[\s()-]', '', text) # Eliminar espacios, paréntesis, guiones
    return cleaned_text.isdigit() and len(cleaned_text) >= 7

def consolidate_buffered_rows(buffered_rows, num_cols_in_df):
    """
    Consolida múltiples filas (que se asume pertenecen a una misma entrada lógica)
    en un único diccionario estructurado.
    """
    record = {'Serial': '', 'Título': '', 'Contacto': '', 'Teléfono': '', 'Vendedor': '', 'Etapa': ''}

    if not buffered_rows:
        return record

    # --- Procesar la primera fila del buffer para información primaria ---
    first_row = buffered_rows[0]
    # Acceder a las celdas usando iloc, verificando que el índice exista en la fila
    b_val_first = str(first_row.iloc[1]).strip() if num_cols_in_df > 1 and pd.notna(first_row.iloc[1]) else ""
    c_val_first = str(first_row.iloc[2]).strip() if num_cols_in_df > 2 and pd.notna(first_row.iloc[2]) else ""
    # Asumiendo Etapa está en columna F (índice 5) basado en la imagen de Excel.
    # Si Etapa está en G (índice 6) como el código original sugiere para g_val_first, ajustar aquí.
    # Por ahora, el código original usa iloc[6] (G) para Etapa más adelante.
    # Mantendremos la extracción de Etapa de G para consistencia con la lógica original de `record['Etapa'] = g_val_first`
    # pero es importante que el usuario verifique que esto coincide con su estructura de archivo.
    # `g_val_first` se usa para Etapa.
    g_val_first = str(first_row.iloc[6]).strip() if num_cols_in_df > 6 and pd.notna(first_row.iloc[6]) else ""


    # 1. Extraer Serial y Título (principalmente de la columna B)
    if b_val_first:
        match_num_text = re.match(r'^\s*(\d+)\s*(.*)$', b_val_first) # Ej: "1 Dr."
        match_just_num = re.match(r'^\s*(\d+)\s*$', b_val_first)     # Ej: "3"
        
        if match_num_text:
            num_part, text_part = match_num_text.groups()
            record['Serial'] = num_part.strip()
            text_part = text_part.strip()
            if text_part:
                if text_part.lower().rstrip('.') in NORMALIZED_TITLES:
                    record['Título'] = text_part
                elif not is_placeholder_zero(text_part): # MODIFIED
                    record['Contacto'] = text_part
        elif match_just_num:
            record['Serial'] = b_val_first.strip()
        else: # b_val_first es solo texto (ej: "Dr. Nombre" o "Nombre")
            parts = b_val_first.split(maxsplit=1)
            first_word = parts[0]
            if first_word.lower().rstrip('.') in NORMALIZED_TITLES:
                record['Título'] = first_word
                if len(parts) > 1:
                    potential_contact = parts[1].strip()
                    if not is_placeholder_zero(potential_contact): # MODIFIED
                        record['Contacto'] = potential_contact
            elif not is_placeholder_zero(b_val_first): # MODIFIED
                record['Contacto'] = b_val_first # Asumir que es nombre de contacto

    # 2. Extraer Título y Contacto (de la columna C, si no se llenó antes o para complementar)
    if c_val_first:
        parts = c_val_first.split(maxsplit=1)
        first_word = parts[0]
        rest_of_c = parts[1].strip() if len(parts) > 1 else ""

        if first_word.lower().rstrip('.') in NORMALIZED_TITLES:
            if not record['Título']: # Si Título no vino de columna B
                record['Título'] = first_word
            # Si hay más texto y Contacto está vacío, y no es placeholder
            if rest_of_c and not record['Contacto'] and not is_placeholder_zero(rest_of_c): # MODIFIED
                record['Contacto'] = rest_of_c
        # Si no es título, Contacto está vacío y c_val_first no es placeholder
        elif not record['Contacto'] and not is_placeholder_zero(c_val_first): # MODIFIED
            record['Contacto'] = c_val_first

    # 3. Extraer Etapa (de la columna G de la primera fila)
    # Note: La imagen de ejemplo muestra "Etapa" en columna F.
    # El código original busca Etapa en G (iloc[6]). Si es F (iloc[5]), g_val_first debería ser ajustado arriba.
    if g_val_first and not is_placeholder_zero(g_val_first): # Added placeholder check for Etapa too
        record['Etapa'] = g_val_first

    # --- Procesar todas las filas del buffer para Teléfono, Vendedor y rellenar huecos ---
    for row_data in buffered_rows:
        # iloc[1] es Col B, iloc[2] es Col C, etc.
        c_val = str(row_data.iloc[2]).strip() if num_cols_in_df > 2 and pd.notna(row_data.iloc[2]) else ""
        d_val = str(row_data.iloc[3]).strip() if num_cols_in_df > 3 and pd.notna(row_data.iloc[3]) else ""
        e_val = str(row_data.iloc[4]).strip() if num_cols_in_df > 4 and pd.notna(row_data.iloc[4]) else ""
        f_val = str(row_data.iloc[5]).strip() if num_cols_in_df > 5 and pd.notna(row_data.iloc[5]) else ""
        # Si Etapa está en F, entonces g_val_row (para Etapa en filas subsecuentes) sería f_val.
        # Si Etapa está en G (iloc[6]), extraerlo:
        g_val_row = str(row_data.iloc[6]).strip() if num_cols_in_df > 6 and pd.notna(row_data.iloc[6]) else ""

        # 4. Teléfono (buscar en D, luego E, luego C)
        if not record['Teléfono']:
            if d_val and is_phone_like(d_val): record['Teléfono'] = d_val
            elif e_val and is_phone_like(e_val): record['Teléfono'] = e_val
            elif c_val and is_phone_like(c_val): record['Teléfono'] = c_val
        
        # 5. Vendedor (buscar en E si no es teléfono, luego F, luego D si es conocido y no teléfono)
        if not record['Vendedor']:
            if e_val and not is_phone_like(e_val) and e_val and not is_placeholder_zero(e_val): record['Vendedor'] = e_val
            elif f_val and not is_phone_like(f_val) and f_val and not is_placeholder_zero(f_val): # Asumiendo F no es Etapa o es secundario para Vendedor
                # CAUTION: If F is exclusively Etapa, this line might be problematic.
                # The original code allows F to be a Vendedor.
                record['Vendedor'] = f_val
            elif d_val and not is_phone_like(d_val) and d_val.lower() in NORMALIZED_VENDORS and not is_placeholder_zero(d_val):
                record['Vendedor'] = d_val
        
        # 6. Rellenar Contacto si aún está vacío y D parece un nombre
        if not record['Contacto'] and d_val and not is_phone_like(d_val) and d_val.lower() not in NORMALIZED_VENDORS:
            if not is_placeholder_zero(d_val): # MODIFIED
                record['Contacto'] = d_val
        
        # 7. Rellenar Etapa si aún está vacía (usando g_val_row de Col G, o f_val si Etapa es Col F)
        # This assumes Etapa might appear in subsequent rows of a buffer if not in the first.
        if not record['Etapa'] and g_val_row and not is_placeholder_zero(g_val_row):
            record['Etapa'] = g_val_row
        # If Etapa is truly in column F (iloc[5]) and g_val_first was for G:
        # elif not record['Etapa'] and f_val and not is_placeholder_zero(f_val) and f_val not in NORMALIZED_VENDORS and not is_phone_like(f_val):
            # (This condition gets complex, depends on whether F can also be vendor)
            # Safest is to ensure Etapa is consistently read from its correct column (F or G).
            # The current code primarily sets Etapa from g_val_first (col G of first row).


    # --- Limpieza final ---
    # Si Título está en Contacto pero no en el campo Título
    if record['Contacto'] and not record['Título']:
        contact_parts = record['Contacto'].split(maxsplit=1)
        first_word_contact = contact_parts[0]
        if first_word_contact.lower().rstrip('.') in NORMALIZED_TITLES:
            record['Título'] = first_word_contact
            potential_new_contact = contact_parts[1].strip() if len(contact_parts) > 1 else ""
            if not is_placeholder_zero(potential_new_contact): # MODIFIED
                 record['Contacto'] = potential_new_contact
            else:
                 record['Contacto'] = "" # Clear if it was a placeholder
    
    # Estandarizar Títulos (ej. "dr" a "Dr.")
    if record['Título']:
        normalized_title_val = record['Título'].lower().rstrip('.')
        found_title = False
        # Buscar en la lista original para mantener el formato preferido (ej. con punto)
        for known_title_original in KNOWN_TITLES_LIST:
            if known_title_original.lower().rstrip('.') == normalized_title_val:
                record['Título'] = known_title_original 
                found_title = True
                break
        # Si no se encontró un match exacto en KNOWN_TITLES_LIST pero es una forma válida sin punto
        if not found_title and normalized_title_val in ["dr", "dra", "lic", "ing", "sr", "sra", "srta"]:
             # Capitalize and ensure dot, handling if original already had one then got capitalized
             temp_title = record['Título'].lower().rstrip('.') # e.g. "dr"
             capitalized_title = temp_title.capitalize() # e.g. "Dr"
             record['Título'] = capitalized_title + "." # e.g. "Dr."


    return record

def process_excel_data(df):
    """
    Procesa el DataFrame de entrada, agrupando filas y consolidándolas.
    """
    processed_records = []
    current_record_rows_buffer = []
    num_cols = df.shape[1]

    for index, row in df.iterrows():
        col_b_val = str(row.iloc[1]).strip() if num_cols > 1 and pd.notna(row.iloc[1]) else ""
        col_c_val = str(row.iloc[2]).strip() if num_cols > 2 and pd.notna(row.iloc[2]) else ""

        is_new_entry_signal = False
        # Consider placeholder zero as not a strong signal for new entry if it's the only content in C
        is_c_val_placeholder = is_placeholder_zero(col_c_val)

        if col_b_val and not is_placeholder_zero(col_b_val): # Serial/Title in B is a strong signal unless it's "0"
            is_new_entry_signal = True
        elif col_c_val and not is_c_val_placeholder: # Content in C is a signal if B is empty AND C is not just a placeholder
            if not is_phone_like(col_c_val) and col_c_val.lower() not in NORMALIZED_VENDORS:
                is_new_entry_signal = True
        
        if is_new_entry_signal and current_record_rows_buffer:
            consolidated = consolidate_buffered_rows(current_record_rows_buffer, num_cols)
            if consolidated.get('Contacto') or consolidated.get('Serial') or consolidated.get('Título'): # Added Título
                processed_records.append(consolidated)
            current_record_rows_buffer = []
        
        current_record_rows_buffer.append(row)

    if current_record_rows_buffer:
        consolidated = consolidate_buffered_rows(current_record_rows_buffer, num_cols)
        if consolidated.get('Contacto') or consolidated.get('Serial') or consolidated.get('Título'): # Added Título
            processed_records.append(consolidated)
            
    return pd.DataFrame(processed_records)

def run_formatter(file_path, sheet_name):
    """
    Función principal para cargar, procesar y devolver los datos formateados.
    """
    try:
        # La imagen muestra encabezados en la fila 2 de Excel.
        # Pandas usa indexación 0, así que header=1 significa que la fila 2 es el encabezado.
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    except FileNotFoundError:
        print(f"Error: El archivo '{file_path}' no fue encontrado.")
        return None
    except Exception as e: 
        print(f"Error al leer el archivo Excel o la hoja especificada: {e}")
        return None

    if df.empty:
        print(f"Advertencia: La hoja '{sheet_name}' del archivo '{file_path}' está vacía o no se pudo leer correctamente con los encabezados especificados.")
        return pd.DataFrame() 
        
    print(f"Procesando hoja '{sheet_name}' del archivo '{file_path}'...")
    # Drop rows where all values are NaN because these can interfere with logic if they have ".0" like strings
    df.dropna(how='all', inplace=True)
    if df.empty:
        print(f"Advertencia: La hoja '{sheet_name}' está vacía después de eliminar filas completamente vacías.")
        return pd.DataFrame(columns=['Serial', 'Título', 'Contacto', 'Teléfono', 'Vendedor', 'Etapa'])

    processed_df = process_excel_data(df.copy()) 

    output_columns = ['Serial', 'Título', 'Contacto', 'Teléfono', 'Vendedor', 'Etapa']
    
    if not processed_df.empty:
        final_df = processed_df.reindex(columns=output_columns).fillna('')
    else:
        final_df = pd.DataFrame(columns=output_columns)

    print("Procesamiento completado.")
    return final_df

if __name__ == '__main__':
    # --- CONFIGURACIÓN ---
    archivo_excel = "1.xlsx" 
    nombre_hoja = "one"            
    
    formatted_data = run_formatter(archivo_excel, nombre_hoja)
    
    if formatted_data is not None:
        if not formatted_data.empty:
            print("\nDatos Procesados:")
            print(formatted_data.to_string()) 
            
            try:
                output_filename = "datos_procesados.xlsx"
                formatted_data.to_excel(output_filename, index=False)
                print(f"\nLos datos procesados se han guardado en '{output_filename}'")
            except Exception as e:
                print(f"Error al guardar el archivo de salida: {e}")
        else:
            print("\nEl procesamiento resultó en una tabla vacía. Verifica los datos de entrada y la lógica del script.")
    else:
        print("No se generaron datos procesados debido a un error previo (revisa los mensajes).")