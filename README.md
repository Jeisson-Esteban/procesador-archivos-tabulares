# procesador-archivos-tabulares

> Servicio HTTP que normaliza archivos CSV y Excel desordenados: BOM raros, delimitadores mezclados, cabeceras escondidas. Devuelve JSON, CSV limpio o lotes en ZIP.

## Que es esto

- **Que es**: un servicio HTTP que recibe archivos CSV o Excel y devuelve datos limpios (JSON, CSV o lotes ZIP), sin que el cliente tenga que saber que codificacion, separador o cuantas filas de basura trae el archivo arriba de la cabecera.
- **Por que existe**: los archivos reales (los que mandan los proveedores, los que exporta un ERP viejo, los que llegan por correo de SAP) raramente son CSV bonitos. Vienen en UTF-16 con BOM, con titulos y fechas antes de la fila de cabecera, con celdas que parecen numeros pero son strings con apostrofo prefijado, con separadores mezclados. Esta API encapsula esa pelea para que un workflow de n8n, un script de ETL o un backend cualquiera no tengan que reescribirla cada vez.
- **Para que sirve / a quien le sirve**: util para automatizaciones que reciben archivos heterogeneos (logistica, e-commerce, finanzas operativas), para devs que necesitan partir un Excel de 200k filas en lotes manejables, o para flujos no-code (n8n, Make, Zapier) que necesitan un endpoint estable que se trague cualquier formato.

## Stack

- **Lenguaje**: Python 3.11 (corre en 3.9+).
- **Frameworks / libs clave**: Flask, pandas, openpyxl, xlrd, python-magic, gunicorn.
- **Servicios externos**: ninguno. Es stateless, no necesita base de datos.

## Casos de uso

- **Caso 1 — Normalizar archivos heterogeneos en un pipeline n8n**: el nodo HTTP de n8n pasa el archivo a `/to-json` y recibe un array limpio, sin importar si era CSV, XLS o XLSX. Toca un solo nodo en lugar de cinco.
- **Caso 2 — Partir un Excel gigante para procesar en lotes**: `/split-file?chunk_size=5000` devuelve un ZIP con N archivos de 5000 filas, listos para meter en una cola de trabajos o subir a un servicio que tiene limite de tamano.
- **Caso 3 — Extraer la metadata previa a la cabecera**: muchos reportes traen el nombre del cliente, fechas y filtros antes de la tabla real. `/extract-data` separa esa metadata de los datos tabulares.
- **Caso 4 — Endpoint generico para SaaS de scraping/RPA**: los bots descargan reportes en formatos inconsistentes; pasarlos por esta API les da una sola interfaz estable.
- **Caso 5 — Convertir ZIP de CSVs a un solo JSON por archivo**: `/extractjson` desempaqueta el ZIP, parsea cada CSV y devuelve un diccionario `{nombre_archivo: [...filas]}`.

## Requisitos previos

- Python 3.9 o superior (recomendado 3.11).
- `libmagic` instalado en el sistema (Linux: `apt install libmagic1`, Mac: `brew install libmagic`, Windows: viene con `python-magic-bin` o usar el contenedor Docker).
- Docker y Docker Compose (opcional, recomendado para correrlo aislado).

## Como usarla

### Instalacion

```bash
# 1. Clonar
git clone https://github.com/Jeisson-Esteban/procesador-archivos-tabulares.git
cd procesador-archivos-tabulares

# 2. Crear y activar virtualenv
python -m venv .venv
source .venv/bin/activate  # en Windows: .venv\Scripts\activate

# 3. Dependencias
pip install -r requirements.txt

# 4. Variables de entorno
cp .env.example .env
# editar .env si necesitas cambiar puerto, log level, etc.
```

> **Ojo con python-magic en Windows / Mac**: la libreria `python-magic` necesita `libmagic` instalada en el sistema operativo. El Dockerfile ya la trae, pero si corres local:
> - **Linux**: `sudo apt install libmagic1`
> - **Mac**: `brew install libmagic`
> - **Windows**: en lugar de `python-magic` instala `pip install python-magic-bin` (trae el binario empaquetado).


### Correrla

```bash
# Modo desarrollo (Flask):
python app.py

# Modo produccion (gunicorn):
gunicorn --bind 0.0.0.0:8010 --workers 2 --timeout 60 app:app

# Con Docker Compose:
docker compose up --build
```

Por defecto escucha en `http://localhost:8010`.

### Ejemplo rapido

Convertir un Excel a JSON:

```bash
curl -X POST http://localhost:8010/to-json \
  -F "file=@ejemplos/ventas_demo.csv"
```

Respuesta (formato):

```json
[
  {"order_id": "1001", "product": "Widget A", "qty": "3", "total": "29.97"},
  {"order_id": "1002", "product": "Widget B", "qty": "1", "total": "15.50"}
]
```

Partir un archivo grande en lotes de 1000 filas:

```bash
curl -X POST "http://localhost:8010/split-file?chunk_size=1000" \
  -F "file=@archivo_grande.xlsx" \
  --output lotes.zip
```

Health check:

```bash
curl http://localhost:8010/health
# {"status":"ok","service":"procesador-archivos-tabulares"}
```

### Endpoints disponibles

| Endpoint | Metodo | Descripcion |
|----------|--------|-------------|
| `/health` | GET | Verifica que el servicio esta vivo. |
| `/to-json` | POST | Convierte CSV/Excel a JSON array. |
| `/to-csv` | POST | Devuelve un CSV limpio con cabecera detectada. |
| `/split-file` | POST | Parte un archivo en lotes y devuelve ZIP. `chunk_size` por query param. |
| `/extract-headers` | POST | Devuelve cabeceras + contenido completo del CSV. |
| `/extractjson` | POST | Acepta un CSV o ZIP de CSVs y devuelve JSON. |
| `/extract-data` | POST | Devuelve metadata previa a la cabecera + headers + data. |

### Recomendaciones y tips

Cosas que aprendi a las malas y te ahorran tiempo:

- **Tip 1**: si el archivo es Excel viejo (`.xls`) y `openpyxl` falla, la API ya hace fallback a `xlrd`. No tienes que preocuparte por que motor usar.
- **Tip 2**: para archivos UTF-16 con BOM (los famosos exportes de SAP), el endpoint los detecta y decodifica antes de pasarlos a pandas. Si vas a usar pandas directo en otro lado, copia la logica de `decode_file_content`.
- **Tip 3**: cuando partas archivos muy grandes con `/split-file`, ajusta `MAX_CONTENT_LENGTH_MB` en `.env` y el timeout de gunicorn. El default de 60s puede quedarse corto para 100MB+.
- **Ojo con**: la deteccion de cabecera asume que la fila "buena" tiene >=80% de las columnas no vacias. Si tus archivos tienen cabeceras muy sparse, ajusta el threshold en `detect_header_row_index`.
- **Ojo con**: `python-magic` necesita `libmagic` instalado en el SO. Si te falla la importacion, ese suele ser el problema. El Dockerfile ya lo instala.

## Estructura del proyecto

```
procesador-archivos-tabulares/
├── app.py                <- toda la API en un solo archivo (Flask + parsers)
├── ejemplos/             <- archivos de muestra para probar los endpoints
│   ├── ventas_demo.csv
│   └── reporte_con_metadata.csv
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md
└── LICENSE
```

## Variables de entorno

Ver [.env.example](.env.example) para la lista completa. Las clave:

- `PORT` — puerto del servidor (default 8010).
- `LOG_LEVEL` — nivel de log (DEBUG/INFO/WARNING/ERROR).
- `DEFAULT_CHUNK_SIZE` — tamano de lote por defecto en `/split-file`.
- `MAX_CONTENT_LENGTH_MB` — limite de tamano por upload.

## Contribuir / ideas para mejorarlo

Este repo es publico justamente porque siempre se puede hacer mejor. Si lo usaste, lo rompiste, o se te ocurrio algo:

- **Issues bienvenidos** — bugs, dudas, propuestas. No hay issue tonto.
- **PRs tambien** — para cambios grandes, mejor abrir un issue primero y discutimos el approach.
- **Ideas en el aire**:
  - Streaming real en `/split-file` para no cargar el Excel completo en memoria.
  - Endpoint de validacion contra un schema (declarado en JSON Schema o YAML).
  - Soporte para `.parquet` y `.feather` como input/output.

## Agradecimientos

Gracias por pasarte por aca. Si te sirvio aunque sea para sacar una idea, ya valio la pena publicarlo.

Inspirado en el dia a dia automatizando flujos de archivos que nadie quiere parsear a mano.

## Licencia

MIT — ver [LICENSE](LICENSE). Usalo, modifica, distribuye, lo que necesites.

---

Hecho con cafe en Bogota por [Jeisson](https://github.com/Jeisson-Esteban) — dev backend / automatizacion.
