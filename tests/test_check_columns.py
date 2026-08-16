"""Tests para check_columns.py.

Genera archivos Excel y de texto al vuelo en carpetas temporales,
sin depender de datos externos.
"""

import sys
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import check_columns as cc


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def crear_xlsx(ruta: Path, filas: list[list]) -> None:
    """Crea un archivo .xlsx con las filas indicadas."""
    wb = Workbook()
    ws = wb.active
    for fila in filas:
        ws.append(fila)
    wb.save(ruta)


def crear_utf16(ruta: Path, lineas: list[str], encoding: str = "utf-16-le") -> None:
    """Crea un archivo de texto UTF-16 con BOM (exportación típica de SAP)."""
    bom = b"\xff\xfe" if encoding == "utf-16-le" else b"\xfe\xff"
    ruta.write_bytes(bom + "\n".join(lineas).encode(encoding))


# ---------------------------------------------------------------------------
# contar_columnas_con_datos
# ---------------------------------------------------------------------------
def test_contar_columnas_ignora_columnas_vacias():
    df = pd.DataFrame({"a": [1, 2], "b": [None, None], "c": ["x", "y"]})
    assert cc.contar_columnas_con_datos(df) == 2


def test_contar_columnas_dataframe_vacio():
    assert cc.contar_columnas_con_datos(pd.DataFrame()) == 0


# ---------------------------------------------------------------------------
# obtener_archivos_xls
# ---------------------------------------------------------------------------
def test_obtener_archivos_no_recursivo(tmp_path):
    (tmp_path / "a.xlsx").touch()
    (tmp_path / "b.xls").touch()
    (tmp_path / "ignorar.txt").touch()
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "c.xlsx").touch()

    archivos = cc.obtener_archivos_xls(str(tmp_path), recursivo=False)
    assert sorted(p.name for p in archivos) == ["a.xlsx", "b.xls"]


def test_obtener_archivos_recursivo(tmp_path):
    (tmp_path / "a.xlsx").touch()
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "c.xlsx").touch()

    archivos = cc.obtener_archivos_xls(str(tmp_path), recursivo=True)
    assert sorted(p.name for p in archivos) == ["a.xlsx", "c.xlsx"]


# ---------------------------------------------------------------------------
# detectar_formato_archivo
# ---------------------------------------------------------------------------
def test_detectar_utf16_le(tmp_path):
    ruta = tmp_path / "sap.xls"
    crear_utf16(ruta, ["A\tB", "1\t2"])
    assert cc.detectar_formato_archivo(ruta) == "utf16"


def test_detectar_utf16_be(tmp_path):
    ruta = tmp_path / "sap_be.xls"
    crear_utf16(ruta, ["A\tB", "1\t2"], encoding="utf-16-be")
    assert cc.detectar_formato_archivo(ruta) == "utf16"


def test_detectar_excel(tmp_path):
    ruta = tmp_path / "real.xlsx"
    crear_xlsx(ruta, [["A", "B"], [1, 2]])
    assert cc.detectar_formato_archivo(ruta) == "excel"


# ---------------------------------------------------------------------------
# procesar_archivo
# ---------------------------------------------------------------------------
def test_procesar_xlsx_valido(tmp_path):
    ruta = tmp_path / "ok.xlsx"
    crear_xlsx(ruta, [["H1", "H2", "H3"], [1, 2, 3]])
    resultado = cc.procesar_archivo(ruta, comparar_headers=True)

    assert resultado.error is None
    assert resultado.num_columnas == 3
    assert resultado.num_filas == 2
    assert resultado.nombres_columnas == ["H1", "H2", "H3"]
    assert resultado.tamano_bytes > 0


def test_procesar_utf16_sap(tmp_path):
    ruta = tmp_path / "sap.xls"
    crear_utf16(ruta, ["H1\tH2", "1\t2"])
    resultado = cc.procesar_archivo(ruta)

    assert resultado.error is None
    assert resultado.formato_detectado == "SAP (UTF-16)"
    assert resultado.num_columnas == 2


def test_procesar_archivo_corrupto_no_revienta(tmp_path):
    """Un .xlsx corrupto debe producir error en el resultado, no una excepción."""
    ruta = tmp_path / "corrupto.xlsx"
    ruta.write_bytes(b"esto no es un zip valido")

    resultado = cc.procesar_archivo(ruta)
    assert resultado.error is not None
    assert resultado.num_columnas is None


def test_procesar_archivo_desaparecido_no_revienta(tmp_path):
    """Un archivo borrado entre el listado y la lectura da error, no crash."""
    ruta = tmp_path / "fantasma.xlsx"
    resultado = cc.procesar_archivo(ruta)
    assert resultado.error is not None


# ---------------------------------------------------------------------------
# analizar_carpeta (paralelo)
# ---------------------------------------------------------------------------
def test_analizar_carpeta_con_corrupto_continua(tmp_path):
    crear_xlsx(tmp_path / "ok.xlsx", [["A"], [1]])
    (tmp_path / "malo.xlsx").write_bytes(b"basura")

    archivos = cc.obtener_archivos_xls(str(tmp_path))
    resultados = cc.analizar_carpeta(archivos, silencioso=True)

    assert len(resultados) == 2
    por_nombre = {r.nombre: r for r in resultados}
    assert por_nombre["ok.xlsx"].error is None
    assert por_nombre["malo.xlsx"].error is not None


# ---------------------------------------------------------------------------
# nombre_visible
# ---------------------------------------------------------------------------
def test_nombre_visible_relativo(tmp_path):
    sub = tmp_path / "sub"
    sub.mkdir()
    ruta = sub / "a.xlsx"
    ruta.touch()
    resultado = cc.ResultadoArchivo(nombre=ruta.name, ruta=ruta)
    assert cc.nombre_visible(resultado, str(tmp_path)) == str(Path("sub") / "a.xlsx")


def test_nombre_visible_fuera_de_carpeta(tmp_path):
    ruta = tmp_path / "a.xlsx"
    resultado = cc.ResultadoArchivo(nombre=ruta.name, ruta=ruta)
    assert cc.nombre_visible(resultado, "C:\\otra\\carpeta") == "a.xlsx"


# ---------------------------------------------------------------------------
# _entero_positivo
# ---------------------------------------------------------------------------
def test_entero_positivo_valido():
    assert cc._entero_positivo("4") == 4


@pytest.mark.parametrize("valor", ["0", "-2", "abc", "2.5"])
def test_entero_positivo_invalido(valor):
    with pytest.raises(cc.argparse.ArgumentTypeError):
        cc._entero_positivo(valor)


# ---------------------------------------------------------------------------
# Comparación de headers ordenada
# ---------------------------------------------------------------------------
def _resultado(nombre, headers, columnas=None):
    return cc.ResultadoArchivo(
        nombre=nombre,
        ruta=Path(nombre),
        num_columnas=columnas if columnas is not None else len(headers),
        nombres_columnas=headers,
    )


def test_headers_reordenados_se_detectan(capsys):
    resultados = [
        _resultado("ref.xlsx", ["A", "B", "C"]),
        _resultado("reordenado.xlsx", ["B", "A", "C"]),
    ]
    cc.mostrar_diferencias_headers(resultados, num_comun=3)
    salida = capsys.readouterr().out
    assert "distinto orden" in salida


def test_headers_duplicados_se_detectan(capsys):
    resultados = [
        _resultado("ref.xlsx", ["A", "B", "C"]),
        _resultado("dup.xlsx", ["A", "B", "C", "C"], columnas=3),
    ]
    cc.mostrar_diferencias_headers(resultados, num_comun=3)
    salida = capsys.readouterr().out
    assert "duplicadas" in salida


def test_headers_iguales_no_reportan(capsys):
    resultados = [
        _resultado("a.xlsx", ["A", "B"]),
        _resultado("b.xlsx", ["A", "B"]),
    ]
    cc.mostrar_diferencias_headers(resultados, num_comun=2)
    salida = capsys.readouterr().out
    assert "mismos nombres de columnas" in salida


# ---------------------------------------------------------------------------
# main() end-to-end
# ---------------------------------------------------------------------------
def test_main_todos_iguales_devuelve_0(tmp_path, monkeypatch):
    crear_xlsx(tmp_path / "a.xlsx", [["H1", "H2"], [1, 2]])
    crear_xlsx(tmp_path / "b.xlsx", [["H1", "H2"], [3, 4]])
    monkeypatch.setattr(sys, "argv", ["prog", "-c", str(tmp_path), "-q"])
    assert cc.main() == 0


def test_main_con_diferencias_devuelve_2(tmp_path, monkeypatch):
    crear_xlsx(tmp_path / "a.xlsx", [["H1", "H2"], [1, 2]])
    crear_xlsx(tmp_path / "b.xlsx", [["H1", "H2", "H3"], [1, 2, 3]])
    monkeypatch.setattr(sys, "argv", ["prog", "-c", str(tmp_path), "-q"])
    assert cc.main() == 2


def test_main_carpeta_vacia_devuelve_1(tmp_path, monkeypatch):
    monkeypatch.setattr(sys, "argv", ["prog", "-c", str(tmp_path), "-q"])
    assert cc.main() == 1


def test_main_exporta_json(tmp_path, monkeypatch):
    crear_xlsx(tmp_path / "a.xlsx", [["H1"], [1]])
    salida = tmp_path / "out.json"
    monkeypatch.setattr(
        sys, "argv", ["prog", "-c", str(tmp_path), "-q", "-e", str(salida)],
    )
    assert cc.main() == 0
    assert salida.exists()


def test_main_workers_invalido_falla_argparse(tmp_path, monkeypatch):
    monkeypatch.setattr(sys, "argv", ["prog", "-c", str(tmp_path), "-w", "0"])
    with pytest.raises(SystemExit):
        cc.main()
