import fitz  # PyMuPDF
import os
import sys
import re

def limpiar_nombre_archivo(nombre):
    """
    Limpia un nombre para que sea válido como nombre de archivo.
    Reemplaza '/' con '_' y elimina otros caracteres no válidos.
    """
    if not nombre:
        return "nombre_no_encontrado"
    # Reemplazar / con _
    nombre = nombre.replace('/', '_')
    # Eliminar caracteres no válidos para nombres de archivo en Windows/Unix
    nombre = re.sub(r'[\\:*?"<>|]', '', nombre)
    # Reemplazar múltiples espacios o puntos con uno solo
    nombre = re.sub(r'\s+', ' ', nombre).strip()
    nombre = re.sub(r'\.+', '.', nombre).strip('.')
    # Limitar la longitud del nombre de archivo (opcional, pero buena práctica)
    if len(nombre) > 100:
        nombre = nombre[:100]
    if not nombre: # Si después de limpiar queda vacío
        return "nombre_invalido_o_vacio"
    return nombre

def procesar_pdf(pdf_path):
    """
    Procesa el archivo PDF, extrae páginas y las guarda con el nombre encontrado.
    """
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error al abrir el archivo PDF: {e}")
        input("Presiona Enter para salir.")
        return

    # Crear un directorio para los PDFs de salida
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_dir = os.path.join(os.path.dirname(pdf_path), f"{base_name}_paginas_exportadas")

    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"Directorio de salida creado: {output_dir}")
        except Exception as e:
            print(f"Error al crear el directorio de salida: {e}")
            input("Presiona Enter para salir.")
            return
    else:
        print(f"Directorio de salida ya existe: {output_dir}")

    print(f"Procesando {doc.page_count} páginas...")

    nombres_usados = {} # Para manejar nombres de archivo duplicados

    for i in range(doc.page_count):
        page = doc.load_page(i)
        text = page.get_text("text") # Extraer texto plano

        nombre_extraido = None
        # Intentar encontrar "Nombre" y extraer el texto que sigue
        # Se buscan variaciones comunes como "Nombre:", "Nombre ", etc.
        match = re.search(r"Nombre[:\s]+([^\n]+)", text, re.IGNORECASE)
        if match:
            nombre_extraido = match.group(1).strip()
        else:
            # Intento alternativo si "Nombre" está al final de una línea y el nombre en la siguiente
            # Esto es más complejo y requeriría un análisis más profundo del formato del PDF
            # Por ahora, nos enfocamos en el caso más simple
            print(f"Página {i+1}: No se encontró 'Nombre:' seguido de texto en la misma línea.")
            # Si no se encuentra un nombre, se puede usar un nombre genérico o pedir al usuario.
            # Aquí usamos un nombre genérico.
            nombre_extraido = f"pagina_{i+1}_sin_nombre_identificado"


        if nombre_extraido:
            nombre_limpio = limpiar_nombre_archivo(nombre_extraido)
            
            # Manejo de nombres duplicados
            contador = nombres_usados.get(nombre_limpio, 0) + 1
            nombres_usados[nombre_limpio] = contador
            
            nombre_archivo_final = nombre_limpio
            if contador > 1:
                nombre_archivo_final = f"{nombre_limpio}_{contador-1}" # El primer archivo no lleva sufijo, el segundo _1, etc.

            output_pdf_path = os.path.join(output_dir, f"{nombre_archivo_final}.pdf")

            # Crear un nuevo PDF con solo esta página
            new_doc = fitz.open() # Documento PDF vacío
            new_doc.insert_pdf(doc, from_page=i, to_page=i) # Insertar la página actual
            
            try:
                new_doc.save(output_pdf_path)
                print(f"Página {i+1} guardada como: {output_pdf_path}")
            except Exception as e:
                print(f"Error al guardar la página {i+1} ({output_pdf_path}): {e}")
            finally:
                new_doc.close()
        else:
            # Esto no debería ocurrir si se usa el nombre genérico anterior, pero por si acaso
            print(f"Página {i+1}: No se pudo extraer un nombre.")
            # Podrías guardar con un nombre por defecto aquí también
            nombre_archivo_final = limpiar_nombre_archivo(f"pagina_{i+1}_error_extraccion")
            output_pdf_path = os.path.join(output_dir, f"{nombre_archivo_final}.pdf")
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
            try:
                new_doc.save(output_pdf_path)
                print(f"Página {i+1} guardada con nombre por defecto: {output_pdf_path}")
            except Exception as e:
                print(f"Error al guardar la página {i+1} con nombre por defecto ({output_pdf_path}): {e}")
            finally:
                new_doc.close()


    doc.close()
    print("\nProceso completado.")
    input("Presiona Enter para salir.")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_file_path = sys.argv[1]
        print(f"Archivo PDF recibido: {pdf_file_path}")
        if pdf_file_path.lower().endswith(".pdf"):
            procesar_pdf(pdf_file_path)
        else:
            print("El archivo arrastrado no es un PDF.")
            input("Presiona Enter para salir.")
    else:
        print("Por favor, arrastra un archivo PDF sobre el ejecutable.")
        # También podrías permitir que el usuario ingrese la ruta si no se arrastra ningún archivo
        # pdf_file_path = input("O ingresa la ruta del archivo PDF: ")
        # if os.path.exists(pdf_file_path) and pdf_file_path.lower().endswith(".pdf"):
        #    procesar_pdf(pdf_file_path)
        # else:
        #    print("Ruta de archivo no válida o no es un PDF.")
        input("Presiona Enter para salir.")