# tests

Tests manuales con curl contra los endpoints. Idealmente migrar a `pytest` + `requests` mas adelante.

## Prerequisitos

La API corriendo localmente en `http://localhost:8010` (`python app.py` o `docker compose up`).

## Casos

### Health

```bash
curl http://localhost:8010/health
# Esperado: {"status":"ok","service":"file-tools-api"}
```

### /to-json con CSV limpio

```bash
curl -X POST http://localhost:8010/to-json \
  -F "file=@../ejemplos/ventas_demo.csv"
```

Esperado: array de 10 objetos con keys `order_id`, `product`, `qty`, `unit_price`, `total`, `customer`.

### /to-json con CSV que tiene metadata arriba

```bash
curl -X POST http://localhost:8010/to-json \
  -F "file=@../ejemplos/reporte_con_metadata.csv"
```

Esperado: array de 8 objetos. La cabecera detectada debe ser la fila con `sku,description,...`, no las lineas de "Reporte mensual".

### /to-csv

```bash
curl -X POST http://localhost:8010/to-csv \
  -F "file=@../ejemplos/reporte_con_metadata.csv" \
  --output limpio.csv
```

Esperado: archivo `limpio.csv` con solo las filas de datos, sin la metadata.

### /split-file

```bash
curl -X POST "http://localhost:8010/split-file?chunk_size=3" \
  -F "file=@../ejemplos/ventas_demo.csv" \
  --output lotes.zip
```

Esperado: ZIP con `sample_data.csv` + `part_1.csv` (3 filas), `part_2.csv` (3), `part_3.csv` (3), `part_4.csv` (1).

### /extract-headers

```bash
curl -X POST http://localhost:8010/extract-headers \
  -F "file=@../ejemplos/ventas_demo.csv"
```

Esperado: JSON con `input_column_literals` y `CSV_content_file`.

### /extract-data

```bash
curl -X POST http://localhost:8010/extract-data \
  -F "file=@../ejemplos/reporte_con_metadata.csv"
```

Esperado: JSON con `detected_header_row` apuntando a la fila correcta y `metadata_header` conteniendo las lineas previas.
