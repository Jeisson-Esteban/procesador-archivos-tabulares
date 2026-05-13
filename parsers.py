"""
Decodificacion y lectura tolerante de archivos CSV/Excel.

Los archivos reales suelen venir mal formados: BOM utf-16, delimitadores mezclados,
filas vacias antes de la cabecera. Aqui esta la pelea con esos casos para que el
resto de la app reciba siempre un DataFrame "razonable" o un error claro.
"""

import io
import os

import magic
import pandas as pd

CSV_ENCODINGS = ["utf-8", "utf-16", "latin-1", "cp1252"]
CSV_DELIMITERS = [",", ";", "\t", "|"]
EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}


def decode_file_content(file_bytes: bytes):
    """Prueba BOM, utf-8, utf-16, latin-1 y cp1252 hasta que alguno funcione.

    Devuelve (texto_decodificado, encoding_usado) o (None, None) si todo falla.
    """
    if not file_bytes:
        return "", "empty"
    if file_bytes.startswith(b"\xff\xfe") or file_bytes.startswith(b"\xfe\xff"):
        try:
            data = file_bytes + b"\x00" if len(file_bytes) % 2 else file_bytes
            return data.decode("utf-16"), "utf-16"
        except Exception:
            pass
    if file_bytes.startswith(b"\xef\xbb\xbf"):
        try:
            return file_bytes.decode("utf-8-sig"), "utf-8-sig"
        except Exception:
            pass
    for encoding in CSV_ENCODINGS:
        try:
            return file_bytes.decode(encoding), encoding
        except Exception:
            continue
    return None, None


def detect_csv_properties(text: str):
    """Estima el delimitador mas probable contando columnas por linea en una muestra."""
    lines = text.splitlines()
    if not lines:
        return ",", 1
    sample = lines[:50]
    best_delimiter, max_cols = ",", 1
    for sep in CSV_DELIMITERS:
        current_max = max((len(line.split(sep)) for line in sample), default=0)
        if current_max > max_cols:
            max_cols, best_delimiter = current_max, sep
    return best_delimiter, max_cols


def try_read_csv(file_bytes: bytes):
    """Intenta parsear bytes como CSV con varias estrategias de delimitador."""
    text, _encoding = decode_file_content(file_bytes)
    if text is None:
        return None, "decoding failed"
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    best_delimiter, max_cols = detect_csv_properties(text)
    col_names = range(max_cols)
    strategies = [
        {"sep": best_delimiter, "names": col_names, "header": None},
        {"sep": None, "header": None},
        {"sep": "\t", "names": col_names, "header": None},
        {"sep": best_delimiter, "names": col_names, "header": None, "quoting": 3},
    ]
    for kwargs in strategies:
        try:
            return pd.read_csv(io.StringIO(text), engine="python", **kwargs), None
        except Exception:
            continue
    return None, "all csv parsing strategies failed"


def read_file_to_df(file_bytes: bytes, filename: str):
    """Decide si el archivo es Excel o CSV y devuelve un DataFrame crudo (header=None).

    Combina deteccion de mime real (python-magic) + extension + fallbacks. No
    aplica deteccion de cabecera ni limpieza de celdas — eso lo hace transformers.py.
    """
    mime = magic.from_buffer(file_bytes, mime=True)
    extension = os.path.splitext(filename)[1].lower()
    looks_excel = (
        "spreadsheet" in mime
        or "excel" in mime
        or "zip" in mime
        or "octet-stream" in mime
        or extension in EXCEL_EXTENSIONS
    )
    errors = []

    # UTF-16 puro: pandas a veces se atraganta, mejor decodificar antes.
    if file_bytes.startswith(b"\xff\xfe") or file_bytes.startswith(b"\xfe\xff"):
        try:
            data = file_bytes + b"\x00" if len(file_bytes) % 2 else file_bytes
            text = data.decode("utf-16").replace("\x00", "")
            text = text.replace("\r\n", "\n").replace("\r", "\n")
            best_sep, max_cols = detect_csv_properties(text)
            col_names = range(max_cols)
            for sep in [best_sep, "\t", ",", ";", "|"]:
                try:
                    df = pd.read_csv(io.StringIO(text), sep=sep, names=col_names, header=None, engine="c")
                    if df is not None and df.shape[1] > 1 and not df.empty:
                        return df, None
                except Exception:
                    continue
        except Exception as e:
            errors.append(f"utf16: {e}")

    if looks_excel:
        for engine in ["openpyxl", "xlrd", None]:
            try:
                kwargs = {"header": None}
                if engine:
                    kwargs["engine"] = engine
                df = pd.read_excel(io.BytesIO(file_bytes), **kwargs)
                if df is not None and not df.empty:
                    return df, None
            except Exception as e:
                errors.append(f"excel({engine}): {e}")

    df, error = try_read_csv(file_bytes)
    if df is not None and not df.empty:
        return df, None
    if error:
        errors.append(f"csv: {error}")

    return None, f"unable to decode. mime: {mime}, errors: {'; '.join(errors)}"
