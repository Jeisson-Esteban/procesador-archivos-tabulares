"""
procesador-archivos-tabulares - utilidades sobre archivos CSV y Excel.

Endpoints expuestos desde un solo Flask app. La logica esta separada en dos
familias internas:

- `*_csv_*`: parser tolerante para CSV con encoding y delimitador raros.
- `simple_*`: lectura directa de Excel/CSV detectando la fila de cabecera.

Ojo: muchos archivos reales en el mundo no son CSV bien formados (BOM utf-16,
delimitadores mezclados, filas vacias antes de la cabecera). El parser intenta
varias estrategias antes de rendirse.
"""

import io
import logging
import os
import zipfile

import magic
import pandas as pd
from flask import Flask, Response, jsonify, request, send_file

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
)
log = logging.getLogger("procesador-archivos-tabulares")

app = Flask(__name__)

DEFAULT_CHUNK_SIZE = int(os.environ.get("DEFAULT_CHUNK_SIZE", 5000))
MAX_CONTENT_LENGTH_MB = int(os.environ.get("MAX_CONTENT_LENGTH_MB", 50))
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH_MB * 1024 * 1024

CSV_ENCODINGS = ["utf-8", "utf-16", "latin-1", "cp1252"]
CSV_DELIMITERS = [",", ";", "\t", "|"]
EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}


# ---------------------------------------------------------------------------
# Parser tolerante para CSV / Excel (deteccion automatica de cabecera)
# ---------------------------------------------------------------------------

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
    """Aplica deteccion de cabecera + normalizacion de celdas."""
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


def decode_file_content(file_bytes: bytes):
    """Prueba BOM, utf-8, utf-16, latin-1 y cp1252 hasta que alguno funcione."""
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
    """Estima el delimitador mas probable contando columnas por linea."""
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
    """Decide si el archivo es Excel o CSV y devuelve un DataFrame crudo (sin cabecera)."""
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


def get_uploaded_file():
    """Validacion comun de uploads. Devuelve (file, error_response, status)."""
    if "file" not in request.files:
        return None, jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if not file or file.filename == "":
        return None, jsonify({"error": "Empty filename"}), 400
    return file, None, 200


def process_upload_to_clean_df():
    file, err, status = get_uploaded_file()
    if err:
        return None, err, status
    file_bytes = file.read()
    df, error = read_file_to_df(file_bytes, file.filename)
    if df is None:
        return None, jsonify({"error": f"Processing failed: {error}"}), 422
    return clean_dataframe(df), None, 200


# ---------------------------------------------------------------------------
# Variante simple para /extract-data
# ---------------------------------------------------------------------------

def simple_find_header_index(df: pd.DataFrame, scan_limit: int = 50) -> int:
    max_cols, header_idx = 0, 0
    for i in range(min(len(df), scan_limit)):
        count = sum(1 for v in df.iloc[i] if pd.notna(v) and str(v).strip() != "")
        if count > max_cols:
            max_cols, header_idx = count, i
    return header_idx


def drop_unnamed_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df[[c for c in df.columns if "Unnamed" not in str(c)]]
    return df.where(pd.notnull(df), None)


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "procesador-archivos-tabulares"}), 200


@app.route("/to-json", methods=["POST"])
def convert_to_json():
    """Lee CSV/Excel y devuelve un JSON array de filas."""
    clean_df, err, status = process_upload_to_clean_df()
    if err:
        return err, status
    return jsonify(clean_df.to_dict(orient="records")), 200


@app.route("/to-csv", methods=["POST"])
def convert_to_csv():
    """Lee CSV/Excel y devuelve un CSV limpio con cabecera detectada."""
    clean_df, err, status = process_upload_to_clean_df()
    if err:
        return err, status
    output = io.StringIO()
    clean_df.to_csv(output, index=False)
    original = request.files["file"].filename
    filename_no_ext = os.path.splitext(original)[0]
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment;filename={filename_no_ext}.csv"},
    )


@app.route("/split-file", methods=["POST"])
def split_file():
    """Parte un CSV/Excel grande en lotes de N filas y devuelve un ZIP."""
    try:
        chunk_size = int(request.args.get("chunk_size", DEFAULT_CHUNK_SIZE))
        if chunk_size <= 0:
            return jsonify({"error": "chunk_size debe ser positivo"}), 400
    except ValueError:
        return jsonify({"error": "chunk_size invalido"}), 400

    file, err, status = get_uploaded_file()
    if err:
        return err, status

    # Algunos servidores envuelven el stream sin metodo seekable; toca parcharlo.
    if not hasattr(file.stream, "seekable"):
        file.stream.seekable = lambda: True

    ext = os.path.splitext(file.filename.lower())[1]
    if ext not in {".xlsx", ".csv"}:
        return jsonify({"error": "Formato no soportado"}), 400

    zip_buffer = io.BytesIO()
    try:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            # Muestra rapida (primeras 4 filas) para inspeccion visual sin abrir todo el zip.
            sample_buffer = io.BytesIO()
            if ext == ".csv":
                file.stream.seek(0)
                lines = []
                for _ in range(4):
                    line = file.stream.readline()
                    if not line:
                        break
                    lines.append(line.decode("utf-8", errors="replace"))
                sample_buffer.write("".join(lines).encode("utf-8"))
                file.stream.seek(0)
            else:
                sample_df = pd.read_excel(file.stream, nrows=4)
                sample_df.to_csv(sample_buffer, index=False)
                file.stream.seek(0)
            zipf.writestr("sample_data.csv", sample_buffer.getvalue())

            if ext == ".csv":
                for i, chunk in enumerate(pd.read_csv(file, chunksize=chunk_size)):
                    buf = io.BytesIO()
                    chunk.to_csv(buf, index=False)
                    zipf.writestr(f"part_{i + 1}.csv", buf.getvalue())
            else:
                df = pd.read_excel(file)
                for i, start in enumerate(range(0, len(df), chunk_size)):
                    buf = io.BytesIO()
                    df.iloc[start:start + chunk_size].to_csv(buf, index=False)
                    zipf.writestr(f"part_{i + 1}.csv", buf.getvalue())
    except Exception as e:
        log.exception("split-file failed")
        return jsonify({"error": str(e)}), 500

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name="lotes_divididos.zip",
    )


@app.route("/extract-headers", methods=["POST"])
def extract_headers():
    """Devuelve la primera fila como lista de literales + el contenido completo."""
    file, err, status = get_uploaded_file()
    if err:
        return err, status
    if not file.filename.lower().endswith(".csv"):
        return jsonify({"error": "Solo CSV"}), 400
    content = file.read().decode("utf-8-sig")
    headers = content.splitlines()[0].split(",")
    return jsonify({
        "input_column_literals": ", ".join(f"'{h.strip()}'" for h in headers),
        "CSV_content_file": content,
    })


@app.route("/extractjson", methods=["POST"])
def extract_json_from_zip_or_csv():
    """Convierte un CSV (o todos los CSV dentro de un ZIP) a JSON.

    Las celdas vacias se omiten en el dict resultante.
    """
    file, err, status = get_uploaded_file()
    if err:
        return err, status
    ext = os.path.splitext(file.filename.lower())[1]

    def clean_row(d):
        return {k: v.strip() for k, v in d.items() if str(v).strip() != ""}

    def csv_bytes_to_json(b):
        df = pd.read_csv(io.BytesIO(b), dtype=str, keep_default_na=False)
        return [clean_row(r) for r in df.to_dict(orient="records")]

    data = file.read()
    if ext == ".csv":
        return jsonify(csv_bytes_to_json(data))
    if ext == ".zip":
        result = {}
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for name in z.namelist():
                if name.endswith(".csv"):
                    result[name] = csv_bytes_to_json(z.read(name))
        return jsonify(result)
    return jsonify({"error": "Formato no soportado"}), 400


@app.route("/extract-data", methods=["POST"])
def extract_data():
    """Lee un CSV/Excel y devuelve metadata previa a la cabecera + filas."""
    file, err, status = get_uploaded_file()
    if err:
        return err, status
    content = file.read()
    name = file.filename.lower()
    try:
        if name.endswith(".csv"):
            df_raw = pd.read_csv(io.BytesIO(content), header=None, sep=None, engine="python")
        else:
            df_raw = pd.read_excel(io.BytesIO(content), header=None)

        header_idx = simple_find_header_index(df_raw)
        metadata = []
        for _, row in df_raw.iloc[:header_idx].iterrows():
            cleaned = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
            if cleaned:
                metadata.append(cleaned)

        if name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(content), header=header_idx)
        else:
            df = pd.read_excel(io.BytesIO(content), header=header_idx)
        df = drop_unnamed_columns(df)

        return jsonify({
            "filename": file.filename,
            "detected_header_row": header_idx + 1,
            "metadata_header": metadata,
            "headers": list(df.columns),
            "data": df.to_dict(orient="records"),
        })
    except Exception as e:
        log.exception("extract-data failed")
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8010))
    host = os.environ.get("HOST", "0.0.0.0")
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    log.info("starting procesador-archivos-tabulares on %s:%s", host, port)
    app.run(host=host, port=port, debug=debug)
