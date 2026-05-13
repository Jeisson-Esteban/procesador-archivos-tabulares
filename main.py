"""
procesador-archivos-tabulares - API HTTP para procesar archivos CSV y Excel.

Este archivo solo contiene la app Flask y los endpoints. La logica pesada esta
en `parsers.py` (decoding/lectura) y `transformers.py` (limpieza de DataFrames).
"""

import io
import logging
import os
import zipfile

import pandas as pd
from flask import Flask, Response, jsonify, request, send_file

from parsers import read_file_to_df
from transformers import clean_dataframe, detect_header_row_index, drop_unnamed_columns

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
)
log = logging.getLogger("procesador-archivos-tabulares")

app = Flask(__name__)

DEFAULT_CHUNK_SIZE = int(os.environ.get("DEFAULT_CHUNK_SIZE", 5000))
MAX_CONTENT_LENGTH_MB = int(os.environ.get("MAX_CONTENT_LENGTH_MB", 50))
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH_MB * 1024 * 1024


# ---------------------------------------------------------------------------
# Helpers de request
# ---------------------------------------------------------------------------

def get_uploaded_file():
    """Validacion comun de uploads. Devuelve (file, error_response, status)."""
    if "file" not in request.files:
        return None, jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if not file or file.filename == "":
        return None, jsonify({"error": "Empty filename"}), 400
    return file, None, 200


def process_upload_to_clean_df():
    """Lee el archivo subido y devuelve un DataFrame ya limpio (cabecera detectada)."""
    file, err, status = get_uploaded_file()
    if err:
        return None, err, status
    file_bytes = file.read()
    df, error = read_file_to_df(file_bytes, file.filename)
    if df is None:
        return None, jsonify({"error": f"Processing failed: {error}"}), 422
    return clean_dataframe(df), None, 200


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
        return {k: (v.strip() if isinstance(v, str) else v)
                for k, v in d.items()
                if v is not None and str(v).strip() != ""}

    def csv_bytes_to_json(b, name="archivo.csv"):
        # Usa el parser tolerante para soportar CSV con metadata arriba, BOM raros, etc.
        df, error = read_file_to_df(b, name)
        if df is None:
            return {"error": f"CSV parse failed: {error}"}
        clean = clean_dataframe(df)
        return [clean_row(r) for r in clean.to_dict(orient="records")]

    data = file.read()
    if ext == ".csv":
        return jsonify(csv_bytes_to_json(data, file.filename))
    if ext == ".zip":
        result = {}
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for name in z.namelist():
                if name.endswith(".csv"):
                    result[name] = csv_bytes_to_json(z.read(name), name)
        return jsonify(result)
    return jsonify({"error": "Formato no soportado"}), 400


@app.route("/extract-data", methods=["POST"])
def extract_data():
    """Lee un CSV/Excel y devuelve metadata previa a la cabecera + filas."""
    file, err, status = get_uploaded_file()
    if err:
        return err, status
    content = file.read()
    try:
        df_raw, error = read_file_to_df(content, file.filename)
        if df_raw is None:
            return jsonify({"error": f"Processing failed: {error}"}), 422

        header_idx = detect_header_row_index(df_raw)

        metadata = []
        for _, row in df_raw.iloc[:header_idx].iterrows():
            cleaned = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
            if cleaned:
                metadata.append(cleaned)

        raw_headers = df_raw.iloc[header_idx].tolist()
        headers = [str(h).strip() if pd.notna(h) and str(h).strip() else f"Column_{i + 1}"
                   for i, h in enumerate(raw_headers)]
        body = df_raw.iloc[header_idx + 1:].copy()
        body.columns = headers[: body.shape[1]]
        body = drop_unnamed_columns(body)
        body = body.where(pd.notnull(body), None)

        return jsonify({
            "filename": file.filename,
            "detected_header_row": header_idx + 1,
            "metadata_header": metadata,
            "headers": list(body.columns),
            "data": body.to_dict(orient="records"),
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
