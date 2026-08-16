"""Verificador de columnas en archivos XLS/XLSX.

Analiza una carpeta con archivos Excel (o exportaciones de texto de SAP)
y verifica que todos tengan la misma cantidad de columnas con datos.
Muestra progreso en tiempo real, tabla de resultados con colores y
resalta claramente los archivos que presentan diferencias.
"""

import argparse
import collections
import csv
import json
import logging
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
)
from rich.logging import RichHandler
from rich.table import Table
from rich.text import Text

logger = logging.getLogger(__name__)

# En consolas Windows legacy (cp1252/cp850) los emojis y símbolos no son
# codificables y provocarían UnicodeEncodeError; se degradan a '?' en su lugar.
if sys.platform == "win32":
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError, OSError):
            pass

console = Console()

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
EXTENSIONES_VALIDAS = (".xls", ".xlsx")
# BOMs UTF-16 (LE y BE); pandas con encoding="utf-16" maneja ambos.
BOMS_UTF16 = (b"\xff\xfe", b"\xfe\xff")


# ---------------------------------------------------------------------------
# Modelo de datos
# ---------------------------------------------------------------------------
@dataclass
class ResultadoArchivo:
    """Resultado del análisis de un archivo individual."""

    nombre: str
    ruta: Path
    num_columnas: int | None = None
    num_filas: int | None = None
    tamano_bytes: int = 0
    formato_detectado: str = ""
    nombres_columnas: list[str] = field(default_factory=list)
    tiempo_proceso: float = 0.0
    error: str | None = None

    @property
    def tamano_legible(self) -> str:
        """Retorna el tamaño del archivo en formato legible (B, KB, MB)."""
        if self.tamano_bytes < 1024:
            return f"{self.tamano_bytes} B"
        if self.tamano_bytes < 1024 * 1024:
            return f"{self.tamano_bytes / 1024:.1f} KB"
        return f"{self.tamano_bytes / (1024 * 1024):.1f} MB"


# ---------------------------------------------------------------------------
# Selección de carpeta
# ---------------------------------------------------------------------------
def seleccionar_carpeta(ruta_cli: str | None = None) -> str:
    """Obtiene la ruta de la carpeta a analizar.

    Si se proporciona una ruta por CLI, se usa directamente.
    En caso contrario se abre un diálogo gráfico con tkinter.

    Args:
        ruta_cli: Ruta proporcionada por argumento de línea de comandos.

    Returns:
        Ruta absoluta a la carpeta seleccionada, o cadena vacía si se cancela.
    """
    if ruta_cli:
        ruta = Path(ruta_cli)
        if not ruta.is_dir():
            logger.error("La ruta proporcionada no es un directorio válido: %s", ruta_cli)
            return ""
        return str(ruta.resolve())

    # Importar tkinter solo cuando se necesita (evitar fallo en entornos headless)
    try:
        from tkinter import Tk, filedialog

        root = Tk()
        root.withdraw()
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta con archivos XLS")
        root.destroy()
        return carpeta
    except ImportError:
        logger.error("tkinter no disponible. Proporciona la ruta con --carpeta.")
        return ""


# ---------------------------------------------------------------------------
# Búsqueda de archivos
# ---------------------------------------------------------------------------
def obtener_archivos_xls(carpeta: str, recursivo: bool = False) -> list[Path]:
    """Retorna las rutas de archivos .xls/.xlsx en la carpeta indicada.

    Args:
        carpeta: Ruta absoluta de la carpeta a escanear.
        recursivo: Si ``True``, busca también en subcarpetas.

    Returns:
        Lista de objetos Path apuntando a cada archivo encontrado.
    """
    raiz = Path(carpeta)
    iterador = raiz.rglob("*") if recursivo else raiz.iterdir()
    return [
        p for p in iterador
        if p.is_file() and p.suffix.lower() in EXTENSIONES_VALIDAS
    ]


# ---------------------------------------------------------------------------
# Detección de formato y lectura
# ---------------------------------------------------------------------------
def detectar_formato_archivo(ruta_archivo: Path) -> str:
    """Detecta si el archivo es texto UTF-16 (SAP) o Excel nativo.

    Args:
        ruta_archivo: Ruta al archivo a inspeccionar.

    Returns:
        ``'utf16'`` si comienza con BOM UTF-16 (LE o BE), ``'excel'`` en otro caso.
    """
    with open(ruta_archivo, "rb") as f:
        primeros_bytes = f.read(2)
    return "utf16" if primeros_bytes in BOMS_UTF16 else "excel"


def leer_dataframe(ruta_archivo: Path, formato: str) -> pd.DataFrame:
    """Lee el archivo según su formato y retorna un DataFrame.

    Args:
        ruta_archivo: Ruta al archivo a leer.
        formato: ``'utf16'`` para texto exportado de SAP, ``'excel'`` para Excel nativo.

    Returns:
        DataFrame con el contenido sin procesar del archivo.

    Raises:
        ValueError: Si la extensión no es soportada.
        pd.errors.ParserError: Si falla el parseo del CSV.
    """
    if formato == "utf16":
        try:
            return pd.read_csv(
                ruta_archivo, header=None, sep="\t", encoding="utf-16", engine="python",
            )
        except (pd.errors.ParserError, ValueError) as e:
            logger.warning("Fallo con separador tab, reintentando con detección automática: %s", e)
            return pd.read_csv(
                ruta_archivo, header=None, sep=None, encoding="utf-16", engine="python",
            )

    # Formato Excel nativo
    sufijo = ruta_archivo.suffix.lower()
    if sufijo == ".xls":
        return pd.read_excel(ruta_archivo, header=None, engine="xlrd")
    if sufijo == ".xlsx":
        return pd.read_excel(ruta_archivo, header=None, engine="openpyxl")

    raise ValueError(f"Formato de archivo no soportado: {sufijo}")


def contar_columnas_con_datos(df: pd.DataFrame) -> int:
    """Cuenta columnas que tienen al menos un valor no vacío.

    Args:
        df: DataFrame a evaluar.

    Returns:
        Número de columnas con al menos un dato.
    """
    return df.dropna(axis=1, how="all").shape[1]


def obtener_nombres_columnas(df: pd.DataFrame) -> list[str]:
    """Extrae los nombres de columnas de la primera fila del DataFrame.

    Útil porque el DataFrame se lee sin encabezado (``header=None``),
    por lo que la primera fila suele contener los headers reales del reporte.

    Args:
        df: DataFrame a evaluar.

    Returns:
        Lista con los valores de la primera fila convertidos a cadena.
    """
    if df.empty:
        return []
    return [str(val).strip() for val in df.iloc[0].tolist() if pd.notna(val)]


def procesar_archivo(
    ruta_archivo: Path,
    comparar_headers: bool = False,
) -> ResultadoArchivo:
    """Detecta el formato, lee el archivo y recopila metadatos.

    Args:
        ruta_archivo: Ruta al archivo a procesar.
        comparar_headers: Si ``True``, extrae los nombres de las columnas.

    Returns:
        ``ResultadoArchivo`` con los metadatos del procesamiento.
    """
    resultado = ResultadoArchivo(
        nombre=ruta_archivo.name,
        ruta=ruta_archivo,
    )

    inicio = time.perf_counter()
    try:
        resultado.tamano_bytes = ruta_archivo.stat().st_size
        formato = detectar_formato_archivo(ruta_archivo)
        resultado.formato_detectado = "SAP (UTF-16)" if formato == "utf16" else "Excel"

        df = leer_dataframe(ruta_archivo, formato)
        resultado.num_columnas = contar_columnas_con_datos(df)
        resultado.num_filas = len(df)

        if comparar_headers:
            resultado.nombres_columnas = obtener_nombres_columnas(df)

    except (ValueError, pd.errors.ParserError, OSError) as e:
        resultado.error = str(e)
        logger.error("Error al leer %s: %s", ruta_archivo.name, e)
    except Exception as e:
        # xlrd (XLRDError), openpyxl (InvalidFileException) y zipfile
        # (BadZipFile) lanzan excepciones que no heredan de las anteriores;
        # un archivo corrupto no debe abortar el análisis del resto.
        resultado.error = f"{type(e).__name__}: {e}"
        logger.error("Error al leer %s: %s", ruta_archivo.name, resultado.error)
    finally:
        resultado.tiempo_proceso = time.perf_counter() - inicio

    return resultado


# ---------------------------------------------------------------------------
# Análisis con progreso y paralelismo
# ---------------------------------------------------------------------------
def analizar_carpeta(
    archivos: list[Path],
    comparar_headers: bool = False,
    max_workers: int = 4,
    silencioso: bool = False,
) -> list[ResultadoArchivo]:
    """Procesa archivos en paralelo con barra de progreso visual.

    Usa ``ThreadPoolExecutor`` porque la lectura de archivos es I/O-bound
    y los hilos liberan el GIL durante las operaciones de disco.

    Args:
        archivos: Lista de archivos a analizar.
        comparar_headers: Si ``True``, extrae nombres de columnas.
        max_workers: Número máximo de hilos para procesamiento paralelo.
        silencioso: Si ``True``, no muestra barra de progreso.

    Returns:
        Lista de ``ResultadoArchivo`` con los resultados de cada archivo.
    """
    resultados: list[ResultadoArchivo] = []

    if silencioso:
        for archivo in archivos:
            resultados.append(procesar_archivo(archivo, comparar_headers))
        return resultados

    with Progress(
        SpinnerColumn(),
        TextColumn("[bold blue]{task.description}"),
        BarColumn(bar_width=40),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TextColumn("({task.completed}/{task.total})"),
        TimeElapsedColumn(),
        console=console,
    ) as progreso:
        tarea = progreso.add_task("Analizando archivos", total=len(archivos))

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futuros = {
                executor.submit(procesar_archivo, a, comparar_headers): a
                for a in archivos
            }
            for futuro in as_completed(futuros):
                try:
                    resultados.append(futuro.result())
                except Exception as e:
                    # Última línea de defensa: un fallo inesperado en un hilo
                    # se registra como error del archivo y no aborta el lote.
                    archivo = futuros[futuro]
                    logger.error("Error inesperado procesando %s: %s", archivo.name, e)
                    resultados.append(ResultadoArchivo(
                        nombre=archivo.name,
                        ruta=archivo,
                        error=f"{type(e).__name__}: {e}",
                    ))
                progreso.advance(tarea)

    return resultados


# ---------------------------------------------------------------------------
# Reporte visual con Rich
# ---------------------------------------------------------------------------
def mostrar_resumen(
    carpeta: str,
    resultados: list[ResultadoArchivo],
    tiempo_total: float,
    silencioso: bool = False,
) -> None:
    """Muestra un panel resumen del análisis.

    Args:
        carpeta: Ruta de la carpeta analizada.
        resultados: Lista de resultados del análisis.
        tiempo_total: Tiempo total de ejecución en segundos.
        silencioso: Si ``True``, muestra solo una línea resumen.
    """
    exitosos = [r for r in resultados if r.error is None]
    fallidos = [r for r in resultados if r.error is not None]

    if silencioso:
        valores = [r.num_columnas for r in exitosos]
        iguales = len(set(valores)) <= 1
        estado = "OK" if iguales else "DIFERENCIAS"
        console.print(f"{estado} | {len(exitosos)} archivos | {tiempo_total:.1f}s")
        return

    resumen = (
        f"[bold]📂 Carpeta:[/bold] {carpeta}\n"
        f"[bold]📄 Archivos encontrados:[/bold] {len(resultados)}\n"
        f"[bold]✅ Procesados correctamente:[/bold] {len(exitosos)}\n"
        f"[bold]❌ Con errores:[/bold] {len(fallidos)}\n"
        f"[bold]⏱  Tiempo de análisis:[/bold] {tiempo_total:.2f}s"
    )
    console.print()
    console.print(Panel(resumen, title="📊 Resumen", border_style="blue"))


def nombre_visible(resultado: ResultadoArchivo, carpeta: str) -> str:
    """Retorna la ruta del archivo relativa a la carpeta analizada.

    En modo recursivo varios archivos pueden compartir nombre; mostrar la
    ruta relativa los hace distinguibles en los reportes.

    Args:
        resultado: Resultado del archivo.
        carpeta: Carpeta base del análisis.

    Returns:
        Ruta relativa a ``carpeta``, o el nombre simple si no es posible.
    """
    try:
        return str(resultado.ruta.relative_to(carpeta))
    except ValueError:
        return resultado.nombre


def mostrar_tabla_detallada(
    resultados: list[ResultadoArchivo],
    num_comun: int | None,
    verbose: bool = False,
    carpeta: str = "",
) -> None:
    """Muestra una tabla formateada con los resultados de cada archivo.

    Los archivos con diferencias se muestran **primero** y en rojo para
    facilitar su identificación rápida.

    Args:
        resultados: Lista de resultados del análisis.
        num_comun: Número de columnas más común (para resaltar diferencias).
        verbose: Si ``True``, muestra columnas adicionales de metadatos.
        carpeta: Carpeta base del análisis (para rutas relativas).
    """
    tabla = Table(
        title="📋 Detalle por archivo",
        show_lines=True,
        header_style="bold cyan",
    )
    tabla.add_column("#", justify="right", style="dim", width=4)
    tabla.add_column("Archivo", style="bold", max_width=50)
    tabla.add_column("Columnas", justify="right")
    tabla.add_column("Estado", justify="center")

    if verbose:
        tabla.add_column("Filas", justify="right")
        tabla.add_column("Tamaño", justify="right")
        tabla.add_column("Formato", justify="center")
        tabla.add_column("Tiempo", justify="right")

    exitosos = [r for r in resultados if r.error is None]
    fallidos = [r for r in resultados if r.error is not None]

    # Ordenar: primero los que difieren (para ubicarlos fácilmente)
    ordenados = sorted(exitosos, key=lambda r: (r.num_columnas == num_comun, r.nombre))

    indice = 1
    for resultado in ordenados:
        coincide = resultado.num_columnas == num_comun

        if coincide:
            estilo_nombre = ""
            estado = Text("✔ OK", style="green")
        else:
            estilo_nombre = "bold red"
            estado = Text("✘ DIFERENTE", style="bold red")

        fila: list[str | Text] = [
            str(indice),
            Text(nombre_visible(resultado, carpeta), style=estilo_nombre),
            str(resultado.num_columnas),
            estado,
        ]

        if verbose:
            fila.extend([
                str(resultado.num_filas or "-"),
                resultado.tamano_legible,
                resultado.formato_detectado,
                f"{resultado.tiempo_proceso:.2f}s",
            ])

        tabla.add_row(*fila)
        indice += 1

    for resultado in fallidos:
        detalle = resultado.error or ""
        if len(detalle) > 70:
            detalle = detalle[:67] + "..."
        fila = [
            str(indice),
            Text(nombre_visible(resultado, carpeta), style="dim"),
            "-",
            Text(f"⚠ {detalle}", style="yellow"),
        ]
        if verbose:
            fila.extend([
                "-",
                resultado.tamano_legible,
                "-",
                f"{resultado.tiempo_proceso:.2f}s",
            ])
        tabla.add_row(*fila)
        indice += 1

    console.print()
    console.print(tabla)


def mostrar_archivos_diferentes(
    resultados: list[ResultadoArchivo],
    num_comun: int,
    carpeta: str = "",
) -> None:
    """Muestra un bloque dedicado con los nombres de archivos que difieren.

    Esta sección destaca los archivos problemáticos para que el usuario
    pueda ubicarlos rápidamente sin recorrer toda la tabla.

    Args:
        resultados: Lista de resultados del análisis.
        num_comun: Número de columnas esperado (más común).
        carpeta: Carpeta base del análisis (para rutas relativas).
    """
    diferentes = [
        r for r in resultados
        if r.error is None and r.num_columnas != num_comun
    ]
    if not diferentes:
        return

    lineas = [
        f"[bold red]{nombre_visible(r, carpeta)}[/bold red]  →  {r.num_columnas} columnas "
        f"(esperado: {num_comun})"
        for r in sorted(diferentes, key=lambda r: nombre_visible(r, carpeta))
    ]
    contenido = "\n".join(lineas)

    console.print()
    console.print(Panel(
        contenido,
        title=f"🔴 {len(diferentes)} archivo(s) con diferencias",
        border_style="red",
        subtitle="Revisar manualmente estos archivos",
    ))


def mostrar_diferencias_headers(
    resultados: list[ResultadoArchivo],
    num_comun: int | None,
    carpeta: str = "",
) -> None:
    """Compara los nombres de columnas entre archivos y reporta diferencias.

    Toma como referencia el primer archivo que tenga la cantidad de columnas
    más común, y compara los demás contra él. La comparación es ordenada:
    detecta columnas renombradas, reordenadas y duplicadas.

    Args:
        resultados: Lista de resultados del análisis.
        num_comun: Número de columnas más común.
        carpeta: Carpeta base del análisis (para rutas relativas).
    """
    exitosos = [r for r in resultados if r.error is None and r.nombres_columnas]
    if not exitosos:
        return

    # Archivo de referencia: el primero con la cantidad de columnas más común
    referencia = next(
        (r for r in exitosos if r.num_columnas == num_comun),
        exitosos[0],
    )
    headers_ref = referencia.nombres_columnas
    set_ref = set(headers_ref)

    diferencias_encontradas = False
    for resultado in exitosos:
        if resultado is referencia:
            continue
        headers_actual = resultado.nombres_columnas
        # Comparación como listas: detecta renombradas, reordenadas y duplicadas
        if headers_actual != headers_ref:
            if not diferencias_encontradas:
                console.print()
                console.print("[bold yellow]⚠ Diferencias en nombres de columnas (headers):[/bold yellow]")
                console.print(f"  Referencia: [cyan]{nombre_visible(referencia, carpeta)}[/cyan]")
                diferencias_encontradas = True

            set_actual = set(headers_actual)
            solo_en_ref = sorted(set_ref - set_actual)
            solo_en_actual = sorted(set_actual - set_ref)
            console.print(f"\n  [bold red]{nombre_visible(resultado, carpeta)}:[/bold red]")
            if solo_en_ref:
                console.print(f"    Faltan:  {', '.join(solo_en_ref)}")
            if solo_en_actual:
                console.print(f"    Sobran:  {', '.join(solo_en_actual)}")
            if not solo_en_ref and not solo_en_actual:
                if len(headers_actual) != len(headers_ref):
                    console.print("    Mismos nombres, pero hay columnas duplicadas.")
                else:
                    console.print("    Mismos nombres, pero en distinto orden.")

    if not diferencias_encontradas:
        console.print()
        console.print("[green]✔ Todos los archivos tienen los mismos nombres de columnas.[/green]")


def reportar_resultados(
    carpeta: str,
    resultados: list[ResultadoArchivo],
    tiempo_total: float,
    verbose: bool = False,
    silencioso: bool = False,
    comparar_headers: bool = False,
) -> bool:
    """Orquesta la presentación visual completa de resultados.

    Args:
        carpeta: Ruta de la carpeta analizada.
        resultados: Lista de resultados.
        tiempo_total: Tiempo total de ejecución.
        verbose: Modo detallado.
        silencioso: Modo silencioso (una línea).
        comparar_headers: Si se deben comparar los headers.

    Returns:
        ``True`` si todos coinciden, ``False`` si hay diferencias.
    """
    exitosos = [r for r in resultados if r.error is None]
    valores = [r.num_columnas for r in exitosos]

    if not valores:
        return False

    counter = collections.Counter(valores)
    num_comun = counter.most_common(1)[0][0]
    todos_iguales = len(set(valores)) == 1

    # --- Modo silencioso: una línea + nombres de archivos con diferencias ---
    mostrar_resumen(carpeta, resultados, tiempo_total, silencioso)

    if silencioso:
        if not todos_iguales:
            for r in exitosos:
                if r.num_columnas != num_comun:
                    console.print(f"  ✘ {nombre_visible(r, carpeta)}: {r.num_columnas} columnas")
        return todos_iguales

    # --- Modo normal / verbose ---
    if todos_iguales:
        console.print(
            f"\n[bold green]✔ Todos los archivos tienen la misma cantidad de "
            f"columnas con datos: {valores[0]}[/bold green]",
        )
    else:
        archivos_dif = [r for r in exitosos if r.num_columnas != num_comun]
        console.print(
            f"\n[bold red]✘ Se encontraron {len(archivos_dif)} archivo(s) con "
            f"diferencias (esperado: {num_comun} columnas)[/bold red]",
        )

    mostrar_tabla_detallada(resultados, num_comun, verbose, carpeta)

    # Panel destacado con los nombres de archivos que difieren
    if not todos_iguales:
        mostrar_archivos_diferentes(resultados, num_comun, carpeta)

    if comparar_headers:
        mostrar_diferencias_headers(resultados, num_comun, carpeta)

    return todos_iguales


# ---------------------------------------------------------------------------
# Exportación de resultados
# ---------------------------------------------------------------------------
def exportar_resultados(
    resultados: list[ResultadoArchivo],
    ruta_salida: str,
) -> None:
    """Exporta los resultados a CSV o JSON según la extensión del archivo.

    Args:
        resultados: Lista de resultados a exportar.
        ruta_salida: Ruta del archivo de salida (``.csv`` o ``.json``).
    """
    salida = Path(ruta_salida)
    datos = [
        {
            "archivo": r.nombre,
            "ruta": str(r.ruta),
            "columnas": r.num_columnas,
            "filas": r.num_filas,
            "tamano_bytes": r.tamano_bytes,
            "formato": r.formato_detectado,
            "tiempo_proceso_s": round(r.tiempo_proceso, 3),
            "error": r.error,
            "headers": r.nombres_columnas if r.nombres_columnas else None,
        }
        for r in resultados
    ]

    if salida.suffix.lower() == ".json":
        with open(salida, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    elif salida.suffix.lower() == ".csv":
        campos = [
            "archivo", "ruta", "columnas", "filas",
            "tamano_bytes", "formato", "tiempo_proceso_s", "error",
        ]
        with open(salida, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=campos, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(datos)
    else:
        logger.warning("Formato de exportación no soportado: %s. Use .csv o .json", salida.suffix)
        return

    console.print(f"\n[green]📁 Resultados exportados a:[/green] {salida.resolve()}")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------
def _entero_positivo(valor: str) -> int:
    """Tipo de argparse: entero mayor o igual que 1.

    Args:
        valor: Valor en texto recibido por línea de comandos.

    Returns:
        El valor convertido a ``int``.

    Raises:
        argparse.ArgumentTypeError: Si no es un entero positivo.
    """
    try:
        numero = int(valor)
    except ValueError:
        raise argparse.ArgumentTypeError(f"se esperaba un entero, se recibió: {valor!r}") from None
    if numero < 1:
        raise argparse.ArgumentTypeError(f"el valor debe ser >= 1, se recibió: {numero}")
    return numero


def parsear_argumentos() -> argparse.Namespace:
    """Parsea los argumentos de línea de comandos.

    Returns:
        Namespace con los argumentos parseados.
    """
    parser = argparse.ArgumentParser(
        description="Verifica que todos los archivos XLS/XLSX de una carpeta tengan la misma cantidad de columnas.",
    )
    parser.add_argument(
        "--carpeta", "-c",
        help="Ruta a la carpeta con archivos XLS. Si se omite, se abre un diálogo gráfico.",
    )
    parser.add_argument(
        "--recursivo", "-r",
        action="store_true",
        help="Buscar archivos también en subcarpetas.",
    )
    parser.add_argument(
        "--comparar-headers",
        action="store_true",
        help="Comparar los nombres de las columnas además de la cantidad.",
    )
    parser.add_argument(
        "--exportar", "-e",
        metavar="ARCHIVO",
        help="Exportar resultados a un archivo (.csv o .json).",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Mostrar metadatos adicionales por archivo (filas, tamaño, formato, tiempo).",
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Modo silencioso: solo imprime una línea resumen.",
    )
    parser.add_argument(
        "--workers", "-w",
        type=_entero_positivo,
        default=4,
        metavar="N",
        help="Número de hilos para procesamiento paralelo (entero >= 1, por defecto: 4).",
    )
    return parser.parse_args()


def main() -> int:
    """Flujo principal del script.

    Returns:
        Código de salida: 0 = éxito, 1 = error, 2 = diferencias encontradas.
    """
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[RichHandler(console=console, show_path=False)],
    )

    args = parsear_argumentos()

    if args.quiet and args.comparar_headers:
        logger.warning("--comparar-headers no se evalúa en modo --quiet; se ignorará.")

    carpeta = seleccionar_carpeta(args.carpeta)

    if not carpeta:
        logger.warning("No se seleccionó ninguna carpeta.")
        return 1

    archivos = obtener_archivos_xls(carpeta, recursivo=args.recursivo)
    if not archivos:
        logger.warning("No se encontraron archivos XLS en la carpeta seleccionada.")
        logger.info("Archivos en la carpeta: %s", [p.name for p in Path(carpeta).iterdir()])
        return 1

    if not args.quiet:
        console.print(
            f"\n[bold blue]🔍 Iniciando análisis de {len(archivos)} archivo(s)...[/bold blue]\n",
        )

    inicio = time.perf_counter()

    resultados = analizar_carpeta(
        archivos,
        comparar_headers=args.comparar_headers,
        max_workers=args.workers,
        silencioso=args.quiet,
    )

    tiempo_total = time.perf_counter() - inicio

    exitosos = [r for r in resultados if r.error is None]
    if not exitosos:
        logger.error("No se pudo leer ningún archivo.")
        return 1

    todos_iguales = reportar_resultados(
        carpeta,
        resultados,
        tiempo_total,
        verbose=args.verbose,
        silencioso=args.quiet,
        comparar_headers=args.comparar_headers,
    )

    if args.exportar:
        exportar_resultados(resultados, args.exportar)

    return 0 if todos_iguales else 2


if __name__ == "__main__":
    sys.exit(main())