import os
import pandas as pd
from tkinter import Tk, filedialog
import collections

def seleccionar_carpeta():
    root = Tk()
    root.withdraw()  # Ocultar la ventana principal
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con archivos XLS")
    return carpeta

def obtener_archivos_xls(carpeta):
    archivos = []
    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith(('.xls', '.xlsx')):
            archivos.append(os.path.join(carpeta, archivo))
    return archivos

def contar_columnas_con_datos(archivo):
    # Intentar primero detectar si es un archivo de texto (común en exportaciones de SAP)
    try:
        # Leer los primeros bytes para detectar formato
        with open(archivo, 'rb') as f:
            primeros_bytes = f.read(2)
        
        # Si empieza con BOM UTF-16 LE (\xff\xfe), es un archivo de texto
        if primeros_bytes == b'\xff\xfe':
            # Es un archivo de texto con codificación UTF-16 LE (exportado de SAP)
            try:
                df = pd.read_csv(archivo, header=None, sep='\t', encoding='utf-16', engine='python')
                df = df.dropna(axis=1, how='all')
                return df.shape[1]
            except:
                # Intentar con otros separadores
                df = pd.read_csv(archivo, header=None, sep=None, encoding='utf-16', engine='python')
                df = df.dropna(axis=1, how='all')
                return df.shape[1]
        else:
            # Intentar leer como Excel real
            if archivo.lower().endswith('.xls'):
                df = pd.read_excel(archivo, header=None, engine='xlrd')
            elif archivo.lower().endswith('.xlsx'):
                df = pd.read_excel(archivo, header=None, engine='openpyxl')
            else:
                raise ValueError("Formato de archivo no soportado")
            df = df.dropna(axis=1, how='all')
            return df.shape[1]
    except Exception as e:
        print(f"Error al leer {os.path.basename(archivo)}: {e}")
        return None

def main():
    carpeta = seleccionar_carpeta()
    if not carpeta:
        print("No se seleccionó ninguna carpeta.")
        return

    archivos = obtener_archivos_xls(carpeta)
    if not archivos:
        print("No se encontraron archivos XLS en la carpeta seleccionada.")
        print("Archivos en la carpeta:", os.listdir(carpeta))
        return

    columnas_por_archivo = {}
    for archivo in archivos:
        num_columnas = contar_columnas_con_datos(archivo)
        if num_columnas is not None:
            columnas_por_archivo[archivo] = num_columnas

    if not columnas_por_archivo:
        print("No se pudieron leer ningún archivo.")
        return

    # Verificar si todos tienen la misma cantidad
    valores = list(columnas_por_archivo.values())
    if len(set(valores)) == 1:
        print("Todos los archivos tienen la misma cantidad de columnas con datos:", valores[0])
    else:
        print("Se encontraron diferencias en la cantidad de columnas con datos:")
        for archivo, num in columnas_por_archivo.items():
            print(f"{os.path.basename(archivo)}: {num} columnas")
        
        # Encontrar el número más común
        counter = collections.Counter(valores)
        num_comun = counter.most_common(1)[0][0]
        
        # Archivos que difieren del número más común
        archivos_diferentes = [archivo for archivo, num in columnas_por_archivo.items() if num != num_comun]
        
        print("\nArchivos con diferencias (no coinciden con el número más común de {} columnas):".format(num_comun))
        for archivo in archivos_diferentes:
            num = columnas_por_archivo[archivo]
            print(f"{os.path.basename(archivo)}: {num} columnas")

if __name__ == "__main__":
    main()