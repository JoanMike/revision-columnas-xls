# Verificador de Columnas en Archivos XLS

Este script en Python permite seleccionar una carpeta y verificar si todos los archivos XLS (o XLSX) de reportes descargados de SAP, tienen la misma cantidad de columnas con datos. Si hay diferencias, indica claramente el nombre del archivo y la cantidad de columnas, resaltados en color para ubicarlos rápidamente.

<div align="center">
	<img width="743" height="716" alt="Captura de pantalla 2026-02-24 153522" src="https://github.com/user-attachments/assets/7b95ff85-74c7-4b4e-adce-54f0b1a3073a" />
</div>

## Características

- **Barra de progreso** en tiempo real para visualizar el avance del análisis.
- **Tabla formateada** con colores: archivos con diferencias se muestran primero y en rojo; los que no se pueden leer muestran el motivo del error.
- **Panel destacado** con los nombres exactos de los archivos que difieren.
- **Procesamiento paralelo** (multi-hilo) para mayor velocidad con muchos archivos.
- **Robustez ante archivos corruptos**: un archivo dañado se reporta con su error y no detiene el análisis del resto.
- **Comparación de headers** para detectar columnas renombradas, reordenadas o duplicadas.
- **Escaneo recursivo** de subcarpetas (los archivos se muestran con ruta relativa para distinguir nombres repetidos).
- **Exportación** de resultados a CSV o JSON.
- **Modos quiet/verbose** para automatización o inspección detallada.
- Soporte para archivos Excel nativos (.xls, .xlsx) y exportaciones de texto de SAP (UTF-16 LE y BE).

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
| `--workers` | `-w` | Hilos de procesamiento paralelo (entero >= 1, por defecto: 4) |

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
- Si un archivo no se puede leer, se mostrará el motivo del error en la tabla y el archivo se omitirá del análisis, sin detener el resto.
- El script lee la primera hoja del archivo Excel por defecto, o el archivo CSV completo.
- Los archivos con diferencias se muestran **primero** en la tabla y en un panel separado en rojo para ubicarlos fácilmente.
- En consolas Windows antiguas (símbolo del sistema legacy con cp1252/cp850), los emojis y símbolos se degradan a `?` en lugar de fallar; en Windows Terminal o cualquier terminal moderna se ven correctamente.

## Desarrollo

Para contribuir o modificar el script, instala también las dependencias de desarrollo:

```
pip install -r requirements-dev.txt
```

Ejecutar los tests (generan archivos Excel de prueba al vuelo, sin datos externos):

```
pytest tests/ -v
```

Auditar vulnerabilidades de las dependencias:

```
pip-audit -r requirements.txt
```

## Licencia

Este proyecto está bajo la licencia MIT. Ver el archivo [LICENSE](LICENSE).
