"""
Limpieza y normalizacion de DataFrames ya parseados.

Una vez `parsers.py` devuelve un DataFrame con `header=None`, estas funciones
detectan donde empieza la cabecera real, promueven esa fila a columnas y
normalizan las celdas (strip, descartar nulls equivalentes a vacio, etc.).
"""

import pandas as pd


def detect_header_row_index(df: pd.DataFrame) -> int:
    """Busca la fila que probablemente contiene los encabezados.

    Algunos archivos vienen con metadata arriba (titulos, fechas, blancos) antes
    de la cabecera real. Se asume que la cabecera es la fila con mas celdas no
    vacias dentro de las primeras 150.
    """
    if df.empty:
        return 0
    col_counts = []
    max_valid_cols = 0
    for i, row in df.head(150).iterrows():
        valid_in_row = row.apply(lambda x: pd.notnull(x) and str(x).strip() != "").sum()
        col_counts.append((i, valid_in_row))
        if valid_in_row > max_valid_cols:
            max_valid_cols = valid_in_row
    if max_valid_cols <= 1:
        return 0
    threshold = max_valid_cols * 0.8
    for idx, count in col_counts:
        if count >= threshold:
            return idx
    return 0


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica deteccion de cabecera + normalizacion de celdas.

    El resultado tiene columnas tomadas de la fila detectada como cabecera y
    cuerpo desde la siguiente fila en adelante. Celdas vacias quedan como None.
    """
    if df is None or df.empty:
        return df

    header_idx = detect_header_row_index(df)
    raw_headers = df.iloc[header_idx].tolist()
    headers = []
    for i, h in enumerate(raw_headers):
        if pd.notnull(h):
            h_str = str(h).replace("\x00", "").strip().replace("\n", " ").replace("\r", " ")
        else:
            h_str = ""
        if not h_str or h_str.lower() in {"nan", "none", "unknown"}:
            headers.append(f"Column_{i + 1}")
        else:
            headers.append(h_str)

    body = df.iloc[header_idx + 1:].copy()
    if not body.empty:
        body.columns = headers[: body.shape[1]]
    body.reset_index(drop=True, inplace=True)

    def cell_clean(val):
        if pd.isnull(val):
            return None
        s = str(val).replace("\x00", "").replace("\n", " ").replace("\r", " ").strip()
        # Excel a veces serializa numeros como texto prefijado con apostrofo.
        if s.startswith("'"):
            s = s[1:]
        return s or None

    for col in body.columns:
        body[col] = body[col].apply(cell_clean)
    return body


def drop_unnamed_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina columnas que pandas auto-nombra como 'Unnamed: N' y normaliza NaN a None."""
    df = df[[c for c in df.columns if "Unnamed" not in str(c)]]
    return df.where(pd.notnull(df), None)
