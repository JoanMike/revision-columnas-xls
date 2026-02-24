# Verificador de Columnas en Archivos XLS

Este script en Python permite seleccionar una carpeta y verificar si todos los archivos XLS (o XLSX) de reportes descargados de SAP, tienen la misma cantidad de columnas con datos. Si hay diferencias, indica claramente el nombre del archivo y la cantidad de columnas, resaltados en color para ubicarlos rápidamente.

## Características

- **Barra de progreso** en tiempo real para visualizar el avance del análisis.
- **Tabla formateada** con colores: archivos con diferencias se muestran primero y en rojo.
- **Panel destacado** con los nombres exactos de los archivos que difieren.
- **Procesamiento paralelo** (multi-hilo) para mayor velocidad con muchos archivos.
- **Comparación de headers** para detectar columnas renombradas o reordenadas.
- **Escaneo recursivo** de subcarpetas.
- **Exportación** de resultados a CSV o JSON.
- **Modos quiet/verbose** para automatización o inspección detallada.
- Soporte para archivos Excel nativos (.xls, .xlsx) y exportaciones de texto de SAP (UTF-16).

## Requisitos

- Python 3.12+
- Librerías: pandas, openpyxl, xlrd, rich, tkinter (tkinter viene incluido con Python)

## Instalación de dependencias

```
pip install -r requirements.txt
```

O manualmente:

```
pip install pandas openpyxl xlrd rich
```

Nota: Para archivos .xls (Excel 97-2003), se usa xlrd. Para .xlsx (Excel 2007+), se usa openpyxl.

## Uso

### Modo interactivo (diálogo gráfico)

```
python check_columns.py
```

Se abrirá un diálogo para seleccionar la carpeta que contiene los archivos XLS.

### Modo línea de comandos (CLI)

```
python check_columns.py --carpeta "C:\ruta\a\mi\carpeta"
python check_columns.py -c "C:\ruta\a\mi\carpeta"
```

### Opciones disponibles

| Opción | Corta | Descripción |
|--------|-------|-------------|
| `--carpeta` | `-c` | Ruta a la carpeta con archivos XLS |
| `--recursivo` | `-r` | Buscar también en subcarpetas |
| `--comparar-headers` | | Comparar nombres de columnas además de la cantidad |
| `--exportar` | `-e` | Exportar resultados a archivo (.csv o .json) |
| `--verbose` | `-v` | Mostrar metadatos adicionales (filas, tamaño, formato, tiempo) |
| `--quiet` | `-q` | Modo silencioso: solo una línea resumen + archivos con diferencias |
| `--workers` | `-w` | Hilos de procesamiento paralelo (por defecto: 4) |

### Ejemplos

```bash
# Análisis básico con diálogo gráfico
python check_columns.py

# Análisis con barra de progreso y detalles adicionales
python check_columns.py -c "C:\SAP\Reportes" -v

# Escaneo recursivo comparando headers
python check_columns.py -c "C:\SAP\Reportes" -r --comparar-headers

# Exportar resultados a JSON
python check_columns.py -c "C:\SAP\Reportes" -e resultados.json

# Modo silencioso para scripts batch
python check_columns.py -c "C:\SAP\Reportes" -q
```

### Códigos de salida

| Código | Significado |
|--------|-------------|
| 0 | Éxito: todos los archivos tienen la misma cantidad de columnas |
| 1 | Error: carpeta no seleccionada, sin archivos, o fallo de lectura |
| 2 | Diferencias encontradas entre archivos |

## Notas

- Se considera "columnas con datos" a las columnas que tienen al menos un valor no vacío en alguna fila.
- Si un archivo no se puede leer como Excel, el script intenta leerlo como CSV (común en exportaciones de SAP).
- Si un archivo no se puede leer, se mostrará un error y se omitirá.
- El script lee la primera hoja del archivo Excel por defecto, o el archivo CSV completo.
- Los archivos con diferencias se muestran **primero** en la tabla y en un panel separado en rojo para ubicarlos fácilmente.