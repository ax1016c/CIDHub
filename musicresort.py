import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import re # Para sanitizar nombres de archivo
try:
    import mutagen
    from mutagen.flac import FLACNoHeaderError
    MUTAGEN_AVAILABLE = True
except ImportError:
    MUTAGEN_AVAILABLE = False

def seleccionar_carpeta(titulo_ventana):
    """Abre un diálogo para seleccionar una carpeta."""
    ruta = filedialog.askdirectory(title=titulo_ventana)
    if not ruta:
        # No se cierra el root de tkinter aquí para que el mensaje de advertencia pueda tener un padre
        # messagebox.showwarning("Advertencia", "No se seleccionó ninguna carpeta.")
        return None
    return ruta

def sanitize_filename_part(text_original):
    """
    Limpia una cadena de texto para que sea segura como parte de un nombre de archivo.
    Reemplaza espacios con guiones bajos y elimina caracteres no permitidos.
    Limita la longitud.
    """
    if not text_original:
        return ""
    text = str(text_original)
    text = re.sub(r'[\\/*?:"<>|]', '', text)
    text = re.sub(r'[\s._-]+', '_', text)
    text = text.strip('_')
    return text[:40]

def obtener_prefijo_flac(ruta_archivo):
    """
    Intenta extraer Artista del Álbum, Artista o Álbum de un archivo FLAC.
    Devuelve un prefijo sanitizado o un genérico si no se encuentra/error.
    """
    if not MUTAGEN_AVAILABLE:
        return "FLAC_AUDIO"

    try:
        audio = mutagen.File(ruta_archivo, easy=True)
        if not audio:
            return "FLAC_AUDIO"

        prefijo_extraido = None
        if 'albumartist' in audio and audio['albumartist'][0]:
            prefijo_extraido = audio['albumartist'][0]
        elif 'artist' in audio and audio['artist'][0]:
            prefijo_extraido = audio['artist'][0]
        elif 'album' in audio and audio['album'][0]:
            prefijo_extraido = audio['album'][0]
        
        if prefijo_extraido:
            sanitized = sanitize_filename_part(prefijo_extraido)
            return sanitized if sanitized else "FLAC_AUDIO"
        else:
            return "FLAC_AUDIO"

    except FLACNoHeaderError:
        return "FLAC_CORRUPTO"
    except Exception:
        return "FLAC_AUDIO"

def organizar_y_mover_archivos():
    """
    Función principal para seleccionar carpetas, renombrar automáticamente
    y mover archivos (buscando recursivamente) a la raíz de una USB.
    """
    if not MUTAGEN_AVAILABLE:
        # Crear una ventana raíz temporal para el messagebox si no existe
        temp_root_for_msg = None
        if not tk._default_root:
            temp_root_for_msg = tk.Tk()
            temp_root_for_msg.withdraw()
        messagebox.showerror("Dependencia Faltante",
                             "La biblioteca 'mutagen' es necesaria para leer metadatos de FLAC.\n"
                             "Por favor, instálala con: pip install mutagen")
        if temp_root_for_msg:
            temp_root_for_msg.destroy()
        return

    root = tk.Tk()
    root.withdraw() # Escondemos la ventana principal de Tkinter

    messagebox.showinfo("Información", "Por favor, selecciona la CARPETA DE ORIGEN de los archivos (se buscará en subcarpetas).", parent=root)
    carpeta_origen = seleccionar_carpeta("Selecciona la Carpeta de Origen (Recursivo)")
    if not carpeta_origen:
        messagebox.showwarning("Cancelado", "No se seleccionó carpeta de origen. Operación cancelada.", parent=root)
        root.destroy()
        return

    messagebox.showinfo("Información", "Ahora, selecciona la RAÍZ de tu unidad USB de DESTINO.", parent=root)
    usb_raiz_destino = seleccionar_carpeta("Selecciona la Raíz de la USB de Destino")
    if not usb_raiz_destino:
        messagebox.showwarning("Cancelado", "No se seleccionó USB de destino. Operación cancelada.", parent=root)
        root.destroy()
        return

    try:
        print("Escaneando archivos en la carpeta de origen (esto puede tardar un momento)...")
        lista_rutas_archivos_origen = []
        for dirpath, _, filenames in os.walk(carpeta_origen):
            for filename in filenames:
                lista_rutas_archivos_origen.append(os.path.join(dirpath, filename))
        
        if not lista_rutas_archivos_origen:
            messagebox.showinfo("Información", "No se encontraron archivos (ni en subcarpetas) en la carpeta de origen.", parent=root)
            root.destroy()
            return
        
        print(f"Se encontraron {len(lista_rutas_archivos_origen)} archivos para procesar.")

        num_digitos = len(str(len(lista_rutas_archivos_origen)))
        archivos_movidos_contador = 0
        archivos_fallidos = []
        log_operaciones = [f"Origen: {carpeta_origen}", f"Destino: {usb_raiz_destino}", "---"]


        print(f"Moviendo archivos de '{carpeta_origen}' (y subcarpetas) a '{usb_raiz_destino}'...")
        print("--------------------------------------------------")

        for i, ruta_origen_completa in enumerate(lista_rutas_archivos_origen):
            nombre_archivo_original = os.path.basename(ruta_origen_completa)
            nombre_base, extension = os.path.splitext(nombre_archivo_original)
            
            prefijo_descriptivo = ""
            if extension.lower() == '.flac':
                prefijo_descriptivo = obtener_prefijo_flac(ruta_origen_completa)
            else:
                ext_limpia = extension[1:] if extension and len(extension) > 1 else ""
                prefijo_descriptivo = sanitize_filename_part(ext_limpia.upper())
                if not prefijo_descriptivo:
                    prefijo_descriptivo = "ARCHIVO"
            
            if not prefijo_descriptivo:
                prefijo_descriptivo = "GENERAL"

            numero_formateado = f"{i + 1:0{num_digitos}d}"
            
            nuevo_nombre_archivo = f"{numero_formateado}-{prefijo_descriptivo}-{nombre_base.lstrip('-_')}{extension}"
            ruta_destino_completa = os.path.join(usb_raiz_destino, nuevo_nombre_archivo)

            # Evitar mover un archivo sobre sí mismo si origen y destino son iguales (poco probable aquí)
            if os.path.abspath(ruta_origen_completa) == os.path.abspath(ruta_destino_completa):
                msg = f"OMITIDO (mismo origen y destino): '{nombre_archivo_original}'"
                log_operaciones.append(msg)
                print(msg)
                continue

            try:
                shutil.move(ruta_origen_completa, ruta_destino_completa)
                msg = f"OK: '{nombre_archivo_original}' (de '{os.path.dirname(ruta_origen_completa)}') -> '{nuevo_nombre_archivo}'"
                log_operaciones.append(msg)
                print(msg)
                archivos_movidos_contador += 1
            except Exception as e:
                error_msg = f"ERROR al mover '{nombre_archivo_original}': {e}"
                log_operaciones.append(error_msg)
                print(error_msg)
                archivos_fallidos.append(f"{nombre_archivo_original} (Error: {e})")
        
        print("--------------------------------------------------")
        
        resumen_final = f"Proceso completado.\n\nArchivos encontrados: {len(lista_rutas_archivos_origen)}\nArchivos movidos exitosamente: {archivos_movidos_contador}"
        if archivos_fallidos:
            resumen_final += f"\nArchivos que no se pudieron mover: {len(archivos_fallidos)}"
            messagebox.showwarning("Proceso Completado con Errores", resumen_final, parent=root)
        else:
            messagebox.showinfo("Proceso Completado", resumen_final, parent=root)
        
        # Guardar log detallado en la raíz de la USB
        try:
            log_filename = f"log_movimiento_{os.path.basename(carpeta_origen)}_{numero_formateado}.txt"
            with open(os.path.join(usb_raiz_destino, log_filename), "w", encoding="utf-8") as logfile:
                logfile.write("Resumen de operaciones:\n")
                for linea in log_operaciones:
                    logfile.write(linea + "\n")
            print(f"Log detallado guardado en: {os.path.join(usb_raiz_destino, log_filename)}")
        except Exception as e:
            print(f"No se pudo guardar el log detallado: {e}")


    except FileNotFoundError: # Esto es menos probable ahora con las comprobaciones iniciales
        messagebox.showerror("Error de Ruta", "La carpeta de origen o destino no fue encontrada o no es accesible.", parent=root)
    except Exception as e:
        messagebox.showerror("Error Inesperado", f"Ocurrió un error general: {e}", parent=root)
    finally:
        if root: # Asegurarse de que la ventana de Tkinter se cierre
            root.destroy()

if __name__ == "__main__":
    organizar_y_mover_archivos()