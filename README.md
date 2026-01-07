# Verificador de Columnas en Archivos XLS

Este script en Python permite seleccionar una carpeta y verificar si todos los archivos XLS (o XLSX) de reportes descargados de SAP, tienen la misma cantidad de columnas con datos. Si hay diferencias, indica el nombre del archivo y la cantidad de columnas.

## Requisitos

- Python 3.x
- Librerías: pandas, openpyxl, xlrd, tkinter (tkinter viene incluido con Python)

## Instalación de dependencias

Si no tienes las librerías instaladas, ejecuta:

```
pip install -r requirements.txt
```

O manualmente:

```
pip install pandas openpyxl xlrd
```

Nota: Para archivos .xls (Excel 97-2003), se usa xlrd. Para .xlsx (Excel 2007+), se usa openpyxl.

## Uso

1. Ejecuta el script:
   ```
   python check_columns.py
   ```

2. Se abrirá un diálogo para seleccionar la carpeta que contiene los archivos XLS.

3. El script analizará todos los archivos .xls y .xlsx en la carpeta.

4. Mostrará en la consola si todos tienen la misma cantidad de columnas o listará las diferencias.

## Notas

- Se considera "columnas con datos" a las columnas que tienen al menos un valor no vacío en alguna fila.
- Si un archivo no se puede leer como Excel, el script intenta leerlo como CSV (común en exportaciones de SAP).
- Si un archivo no se puede leer, se mostrará un error y se omitirá.
- El script lee la primera hoja del archivo Excel por defecto, o el archivo CSV completo.