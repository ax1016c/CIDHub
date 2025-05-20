import os
import subprocess
import sys

# --- Configuración ---
# Puedes pre-rellenar esta variable si siempre usas la misma ruta para CoreConverter.exe
# Ejemplo: DBPOWERAMP_CORECONVERTER_PATH = r"C:\Program Files\dBpoweramp\CoreConverter.exe"
DBPOWERAMP_CORECONVERTER_PATH = ""

# Configuración del formato de salida
OUTPUT_FORMAT_NAME = "FLAC"  # Formato reconocido por dbpoweramp
OUTPUT_EXTENSION = ".flac"
# Nivel de compresión FLAC (0-8, 5 es el predeterminado, 8 es la mejor compresión pero más lento)
FLAC_COMPRESSION_LEVEL = "5"  # Se usa como cadena para la línea de comandos
# Opciones adicionales para dbpoweramp, por ejemplo, para verificar el archivo después de la conversión.
# Si quieres que se borre el archivo original después de una conversión exitosa,
# puedes añadir (CON MUCHO CUIDADO): "-delete_source"
# EXTRA_DB_OPTIONS = ["-verify", "-delete_source"]
EXTRA_DB_OPTIONS = ["-verify"]

def main():
    global DBPOWERAMP_CORECONVERTER_PATH

    print("Conversor de música de M4A a FLAC con dBpoweramp")
    print("================================================")
    print(f"Este script convertirá archivos .m4a a {OUTPUT_FORMAT_NAME} manteniendo la estructura de carpetas.")
    print("Necesitarás tener dBpoweramp Reference instalado.")
    print("-" * 40)

    # 1. Obtener la ruta de CoreConverter.exe
    if not DBPOWERAMP_CORECONVERTER_PATH:
        default_db_path = ""
        if os.name == 'nt': # Asumir ruta común en Windows
            program_files = os.environ.get("ProgramFiles", "C:\\Program Files")
            program_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
            possible_paths = [
                os.path.join(program_files, "dBpoweramp", "CoreConverter.exe"),
                os.path.join(program_files_x86, "dBpoweramp", "CoreConverter.exe")
            ]
            for p_path in possible_paths:
                if os.path.exists(p_path):
                    default_db_path = p_path
                    break
        
        if default_db_path:
            user_path = input(f"Introduce la ruta a 'CoreConverter.exe' de dBpoweramp [{default_db_path}]: ").strip().strip('"')
            DBPOWERAMP_CORECONVERTER_PATH = user_path if user_path else default_db_path
        else:
            DBPOWERAMP_CORECONVERTER_PATH = input("Introduce la ruta completa a 'CoreConverter.exe' de dBpoweramp: ").strip().strip('"')


    if not os.path.isfile(DBPOWERAMP_CORECONVERTER_PATH) or not DBPOWERAMP_CORECONVERTER_PATH.lower().endswith("coreconverter.exe"):
        print(f"\nError: 'CoreConverter.exe' no encontrado en la ruta especificada: '{DBPOWERAMP_CORECONVERTER_PATH}'")
        print("Asegúrate de que dBpoweramp esté instalado y la ruta sea correcta.")
        print(r"Ejemplo de ruta: C:\Program Files\dBpoweramp\CoreConverter.exe")
        sys.exit(1)
    
    print(f"Usando CoreConverter.exe de: {DBPOWERAMP_CORECONVERTER_PATH}")

    # 2. Obtener el directorio de la librería de iTunes (entrada)
    input_music_library = ""
    while not os.path.isdir(input_music_library):
        input_music_library = input("Introduce la ruta a tu librería de música M4A (ej: D:\\iTunes\\Music): ").strip().strip('"')
        if not os.path.isdir(input_music_library):
            print(f"Error: El directorio de entrada no existe o no es válido: '{input_music_library}'")

    # 3. Obtener el directorio de salida para los archivos convertidos
    output_converted_library = ""
    while not output_converted_library:
        output_converted_library = input(f"Introduce la ruta para guardar los archivos convertidos a {OUTPUT_FORMAT_NAME} (ej: D:\\Music_FLAC): ").strip().strip('"')
        if not output_converted_library:
            print("Error: Debes especificar un directorio de salida.")
        elif os.path.abspath(input_music_library).lower() == os.path.abspath(output_converted_library).lower():
            print("Error: El directorio de entrada y salida no pueden ser el mismo. Elige un directorio de salida diferente.")
            output_converted_library = "" # Forzar a pedir de nuevo

    try:
        os.makedirs(output_converted_library, exist_ok=True)
    except OSError as e:
        print(f"Error creando el directorio de salida '{output_converted_library}': {e}")
        sys.exit(1)

    print(f"\nBuscando archivos .m4a en: {input_music_library}")
    print(f"Los archivos convertidos se guardarán en: {output_converted_library}")
    print("-" * 40)

    processed_count = 0
    success_count = 0
    error_count = 0
    skipped_count = 0

    for root_dir, _, files in os.walk(input_music_library):
        for filename in files:
            if filename.lower().endswith(".m4a"):
                m4a_file_path = os.path.join(root_dir, filename)
                processed_count += 1
                
                print(f"\n[{processed_count}] Procesando: {m4a_file_path}")

                # Calcular la ruta relativa para mantener la estructura de carpetas
                relative_dir_path = os.path.relpath(root_dir, input_music_library)
                
                # Construir la ruta de salida completa
                output_file_basename = os.path.splitext(filename)[0] + OUTPUT_EXTENSION
                
                if relative_dir_path == ".": # Archivos en la raíz del directorio de entrada
                    output_target_dir = output_converted_library
                else:
                    output_target_dir = os.path.join(output_converted_library, relative_dir_path)

                output_full_path = os.path.join(output_target_dir, output_file_basename)

                # Crear el subdirectorio de salida si no existe
                try:
                    os.makedirs(output_target_dir, exist_ok=True)
                except OSError as e:
                    print(f"  Error creando el subdirectorio de salida '{output_target_dir}': {e}")
                    error_count += 1
                    continue # Saltar este archivo

                # Verificar si el archivo de salida ya existe
                if os.path.exists(output_full_path):
                    print(f"  Saltando, el archivo de salida ya existe: {output_full_path}")
                    skipped_count +=1
                    continue

                # Construir el comando para dbpoweramp
                cmd = [
                    DBPOWERAMP_CORECONVERTER_PATH,
                    f"-infile={m4a_file_path}",
                    f"-outfile={output_full_path}",
                    f"-convert_to={OUTPUT_FORMAT_NAME}"
                ]
                if OUTPUT_FORMAT_NAME == "FLAC":
                    cmd.append(f"-compression-level-{FLAC_COMPRESSION_LEVEL}")
                
                cmd.extend(EXTRA_DB_OPTIONS)

                print(f"  Comando: {' '.join(cmd)}")

                try:
                    # Ejecutar el comando de conversión
                    # CREATE_NO_WINDOW oculta la ventana de la consola en Windows
                    creation_flags = 0
                    if os.name == 'nt':
                        creation_flags = subprocess.CREATE_NO_WINDOW
                    
                    result = subprocess.run(cmd, capture_output=True, text=True, check=False, 
                                            encoding='utf-8', errors='replace',
                                            creationflags=creation_flags)
                    
                    if result.returncode == 0:
                        print(f"  Éxito: Convertido a -> {output_full_path}")
                        if result.stdout and result.stdout.strip():
                            # dBpoweramp a veces no produce salida en stdout en éxito si no hay problemas
                            pass # print(f"    Salida dBpoweramp (stdout):\n{result.stdout.strip()}")
                        success_count += 1
                    else:
                        print(f"  Error convirtiendo {m4a_file_path}:")
                        print(f"    Código de retorno: {result.returncode}")
                        if result.stdout and result.stdout.strip():
                            print(f"    Salida dBpoweramp (stdout):\n{result.stdout.strip()}")
                        if result.stderr and result.stderr.strip():
                            print(f"    Salida dBpoweramp (stderr):\n{result.stderr.strip()}")
                        error_count += 1
                except FileNotFoundError:
                    # Esto no debería ocurrir si ya verificamos DBPOWERAMP_CORECONVERTER_PATH
                    print(f"  Error Crítico: No se pudo encontrar CoreConverter.exe en '{DBPOWERAMP_CORECONVERTER_PATH}'.")
                    print("  Por favor, verifica la ruta e inténtalo de nuevo.")
                    sys.exit(1) 
                except Exception as e:
                    print(f"  Excepción durante la conversión de {m4a_file_path}: {e}")
                    error_count += 1

    print("\n" + "-" * 40)
    print("--- Resumen de Conversión ---")
    print(f"Archivos .m4a encontrados: {processed_count}")
    print(f"Conversiones exitosas: {success_count}")
    print(f"Archivos omitidos (ya existían): {skipped_count}")
    print(f"Conversiones fallidas: {error_count}")
    print(f"Archivos convertidos guardados en: {output_converted_library}")
    print("-" * 40)

if __name__ == "__main__":
    main()