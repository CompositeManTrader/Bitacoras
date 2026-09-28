import streamlit as st
import pandas as pd
import openpyxl
from datetime import date, time, datetime
import io
import re
import unicodedata
import zipfile

st.set_page_config(page_title="Bitácoras por Promotor", page_icon="📋", layout="centered")
st.title("📋 Generador de Bitácoras por Promotor")

# ─────────────────────────────────────────────────────────────────
# CONFIGURACIÓN FIJA
# ─────────────────────────────────────────────────────────────────
CONTRATO_PROMOTOR = {
    0:      "MIGUEL ANGEL TEBAR PEDROZA",   # ORDEN GLOBAL O PAQUETE
    9890:   "MIGUEL ANGEL TEBAR PEDROZA",   # CAPITALES FACILITATION
    100320: "MIGUEL ANGEL TEBAR PEDROZA",   # INDUSTRIAS CH
    100321: "MIGUEL ANGEL TEBAR PEDROZA",   # GRUPO SIMEC
    104351: "MIGUEL ANGEL TEBAR PEDROZA",   # BANCO AZTECA (H2H)
    105433: "MIGUEL ANGEL TEBAR PEDROZA",   # SEGUROS AZTECA
    105862: "MIGUEL ANGEL TEBAR PEDROZA",   # COMPASS INVESTMENTS
    106043: "MIGUEL ANGEL TEBAR PEDROZA",
    105741: "MIGUEL ANGEL TEBAR PEDROZA",   # FEX CAPITAL
    105806: "MIGUEL ANGEL TEBAR PEDROZA",
    105813: "MIGUEL ANGEL TEBAR PEDROZA",
    104871: "JOSE LUIS ALCAINE",            # FONDO DE PROMOCION B
    105775: "JOSE LUIS ALCAINE",            # SKANDIA LIFE
    100844: "JOSE LUIS ALCAINE",            # HDI SEGUROS
    105434: "JOSE LUIS ALCAINE",
    105777: "JOSE LUIS ALCAINE",
    106044: "GERARDO PEREZ CRUZ",
}

OP_TO_NAME = {
    "CB1074134":  "GERARDO PEREZ CRUZ",
    "CB1059258":  "MIGUEL ANGEL TEBAR PEDROZA",
    "CBP1059258": "MIGUEL ANGEL TEBAR PEDROZA",
    "1059258":    "MIGUEL ANGEL TEBAR PEDROZA",
    "CLCB178007": "ALBERTO ALARCON GONZALEZ",
    "CB331177":   "CB331177",
    "H2H":        "H2H",
}

LAYOUT_MAP = {
    "Fecha de la instrucción \n-recepción-":        ("Fecha Registro",      None),
    "Hora de la instrucción \n-recepción-":          ("Hora Registro",       None),
    "Nombre de la persona que gira la instrucción":  ("Nombre",              None),
    "Persona facultada para girar instrucciones":    (None,                  "SI"),
    "Contrato se encuentra vigente":                 (None,                  "SI"),
    "Instrucción registrada como orden":             (None,                  "SI"),
    "Contrato":                                      ("Contrato",            None),
    "Tipo de servicio":                              ("Servicio Contratado", None),
    "Cliente":                                       ("Nombre",              None),
    "Sentido de la operación":                       ("Operación",           None),
    "Emisora":                                       ("Emisora",             None),
    "Serie":                                         ("Serie",               None),
    "Títulos":                                       ("Títulos Ordenados",   None),
    "Precio fijado":                                 ("Precio asignado",     None),
    "Precio a mercado":                              ("Mdo",                 None),
    "Tipo Orden":                                    ("Tipo Orden",          None),
    "Vigencia":                                      ("Vigencia Original",   None),
    "Medio de instrucción":                          ("Medio Instruccion",   None),
    "Clave del promotor que atendió":                ("Operador",            None),
    "Nombre del promotor que atendió":               ("Operador",            "PROMOTOR"),
    "Promotor asignado al contrato":                 ("Contrato",            "CONTRATO_PROMOTOR"),
    "Hora de la captura\n-registro-":                ("Hora Registro",       None),
    "Folio Orden":                                   ("Folio Orden",         None),
    "Comentarios":                                   (None,                  ""),
}

# Columnas que usa la bitácora (nombres canónicos del archivo fuente)
SOURCE_COLS = sorted({src for src, _ in LAYOUT_MAP.values() if src})

HEADER_ROW = 2
DATA_START  = 3
LAYOUT_PATH = "Layout Bitácora Promotor.xlsx"

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# Caracteres dañados por una mala codificación del export (ï¿½ = carácter perdido)
MOJIBAKE_FIX = {
    "telï¿½fono": "teléfono",
    "Instrucciï¿½n": "Instrucción",
    "instrucciï¿½n": "instrucción",
    "Dï¿½a": "Día",
    "Tï¿½tulos": "Títulos",
    "Operaciï¿½n": "Operación",
}

# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def parse_date(x):
    if x is None or (not isinstance(x, (date, time)) and pd.isna(x)): return None
    if isinstance(x, datetime): return x.date()
    if isinstance(x, date): return x
    dt = pd.to_datetime(x, errors="coerce")
    return None if pd.isna(dt) else dt.date()

def parse_time(x):
    if x is None or (not isinstance(x, (date, time)) and pd.isna(x)): return None
    if isinstance(x, time): return x
    if isinstance(x, datetime): return x.time()
    dt = pd.to_datetime(str(x), errors="coerce")
    return None if pd.isna(dt) else dt.time()

def fix_text(x):
    if not isinstance(x, str): return x
    for bad, good in MOJIBAKE_FIX.items():
        x = x.replace(bad, good)
    return x

def norm_header(h) -> str:
    """Normaliza un encabezado: sin acentos, sin caracteres dañados, minúsculas."""
    h = fix_text(str(h))
    h = unicodedata.normalize("NFKD", h).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", h).strip().lower()

def clean_fragment(x) -> str:
    """Quita la basura de comas/comillas que deja un CSV mal convertido: ',CB1074134"' -> 'CB1074134'."""
    return str(x).strip().strip(',"').strip()

def to_number(s):
    s = str(s).replace(",", "").replace('"', "").replace(" ", "")
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f.is_integer() else f

def get_promotor(row):
    try:
        contrato = int(row["Contrato"])
    except (ValueError, TypeError):
        contrato = None
    return CONTRATO_PROMOTOR.get(contrato, "SIN ASIGNAR")

def etiqueta_mes(periodo: pd.Period) -> str:
    return f"{MESES[periodo.month - 1]} {periodo.year}"

# ─────────────────────────────────────────────────────────────────
# LECTURA DEL ARCHIVO FUENTE
# ─────────────────────────────────────────────────────────────────
BROKEN_CELL = re.compile(r'^,.*"$')

def is_broken(rows) -> bool:
    """El export trae celdas tipo ',CB1074134"' cuando el CSV se abrió mal en Excel
    y las columnas quedaron recorridas respecto al encabezado."""
    sample = rows[:50]
    return any(isinstance(v, str) and BROKEN_CELL.match(v) for r in sample for v in r)

def parse_broken_row(vals):
    """Reconstruye un renglón desalineado usando anclas fijas:
    las dos fechas (Fecha Registro y Fecha de Operación Sentra) y el par Elegible/Calificado (SI/NO).
    Orden real de los datos:
      Contrato, Nombre, Hora, Folio, Emisora, Serie, Tipo Orden, Precio Ord.,
      Precio asignado (1 o 2 celdas), Operador, Títulos Originales, FECHA REGISTRO,
      Títulos Ordenados, ..., Vigencia Original, Vigencia Faltante, Elegible, Calificado,
      Mdo/MPL, Hora Vigencia, ..., Bolsa, Libro, FECHA SENTRA, Hora Sentra, Medio, Servicio
    """
    dts = [i for i, v in enumerate(vals) if isinstance(v, datetime)]
    if len(dts) < 2:
        raise ValueError("no se encontraron las dos fechas del renglón")
    f1, f2 = dts[0], dts[1]
    e = next(i for i in range(f1 + 1, f2) if vals[i] in ("SI", "NO"))

    def at(i):
        return vals[i] if i < len(vals) else None

    return {
        "Contrato":            vals[0],
        "Nombre":              vals[1],
        "Hora Registro":       vals[2],
        "Folio Orden":         vals[3],
        "Emisora":             vals[4],
        "Serie":               vals[5],
        "Tipo Orden":          vals[6],
        "Precio asignado":     to_number("".join(str(v) for v in vals[8:f1 - 2])),
        "Operador":            clean_fragment(vals[f1 - 2]),
        "Fecha Registro":      vals[f1],
        "Títulos Ordenados":   vals[f1 + 1],
        "Vigencia Original":   vals[e - 2],
        "Mdo":                 "Mdo" if "Mdo" in str(vals[e + 2]) else "",
        "Medio Instruccion":   fix_text(at(f2 + 2)),
        "Servicio Contratado": fix_text(at(f2 + 3)),
        "Operación":           None,   # este export no trae el sentido (Compra/Venta)
    }

def load_source(file):
    """Devuelve (DataFrame con columnas canónicas, lista de avisos)."""
    ws = openpyxl.load_workbook(file, data_only=True, read_only=True).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    header, body = rows[0], [r for r in rows[1:] if any(v not in (None, "") for v in r)]
    avisos = []

    if is_broken(body):
        avisos.append(
            "El archivo viene **desalineado** (parece un CSV mal convertido a Excel). "
            "Se reconstruyeron las columnas automáticamente, pero **la columna Operación "
            "(Compra/Venta) no viene en el archivo**, así que “Sentido de la operación” quedará vacío. "
            "Para tenerla, exporta el reporte directo a Excel o abre el CSV con *Datos → Desde texto/CSV*."
        )
        parsed, errores = [], []
        for n, r in enumerate(body, start=2):
            try:
                parsed.append(parse_broken_row(r))
            except (ValueError, StopIteration, IndexError) as ex:
                errores.append(f"fila {n}: {ex}")
        if errores:
            avisos.append(f"No se pudieron leer {len(errores)} renglones: {errores[:10]}")
        df = pd.DataFrame(parsed, columns=SOURCE_COLS)
    else:
        df = pd.DataFrame(body, columns=[str(h) if h is not None else "" for h in header])
        # Empata encabezados sin importar acentos ("Operacion" = "Operación", etc.)
        by_norm = {norm_header(c): c for c in df.columns}
        faltan = []
        for canon in SOURCE_COLS:
            real = by_norm.get(norm_header(canon))
            if real is None:
                faltan.append(canon)
                df[canon] = None
            elif real != canon:
                df[canon] = df[real]
        if faltan:
            avisos.append(f"Columnas no encontradas en el archivo (quedarán vacías): {faltan}")
        df = df[SOURCE_COLS].copy()
        for c in ("Medio Instruccion", "Servicio Contratado", "Nombre"):
            df[c] = df[c].map(fix_text)

    df["Operador"] = df["Operador"].map(lambda v: "" if v is None or pd.isna(v) else str(v).strip())
    df["__Fecha__"] = pd.to_datetime(df["Fecha Registro"].map(parse_date), errors="coerce")
    return df, avisos

# ─────────────────────────────────────────────────────────────────
# GENERACIÓN DE LA BITÁCORA
# ─────────────────────────────────────────────────────────────────
def build_bitacora(df_p: pd.DataFrame, layout_bytes: bytes) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(layout_bytes))
    ws = wb[wb.sheetnames[0]]

    header_map = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(HEADER_ROW, c).value
        if isinstance(val, str):
            header_map[val.strip()] = c

    df_p = df_p.sort_values(["__Fecha__", "Hora Registro"], key=lambda s: s.astype(str))
    for i, row in enumerate(df_p.to_dict("records")):
        r = DATA_START + i
        for header, (src_col, rule) in LAYOUT_MAP.items():
            h_key = header.strip()
            if h_key not in header_map:
                continue
            c = header_map[h_key]

            if rule == "PROMOTOR":
                op = str(row["Operador"]).strip()
                val = OP_TO_NAME.get(op, op)
            elif rule == "CONTRATO_PROMOTOR":
                val = get_promotor(row)
            else:
                val = row.get(src_col) if src_col else rule

            if header.startswith("Fecha"):
                val = parse_date(val)
            elif "Hora" in header:
                val = parse_time(val)
            elif val is not None and not isinstance(val, str) and pd.isna(val):
                val = None

            ws.cell(r, c).value = val

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────
src_file = st.file_uploader("Sube el archivo de Bitácoras (.xlsx)", type=["xlsx"])

if not src_file:
    st.stop()

try:
    with open(LAYOUT_PATH, "rb") as f:
        layout_bytes = f.read()
except FileNotFoundError:
    st.error(f"No se encontró `{LAYOUT_PATH}` en la misma carpeta que este script.")
    st.stop()

src, avisos = load_source(src_file)
for a in avisos:
    st.warning(a)

src = src[src["__Fecha__"].notna()].copy()
src["__Mes__"] = src["__Fecha__"].dt.to_period("M")
src["__Promotor__"] = src.apply(get_promotor, axis=1)

# ── Periodo ───────────────────────────────────────────────────
meses_archivo = sorted(src["__Mes__"].unique())
if not meses_archivo:
    st.error("El archivo no tiene renglones con Fecha Registro válida.")
    st.stop()

meses_sel = st.multiselect(
    "Meses a generar (una bitácora por promotor y mes)",
    options=meses_archivo,
    default=meses_archivo[-3:],       # últimos 3 meses disponibles
    format_func=etiqueta_mes,
)
if not meses_sel:
    st.info("Selecciona al menos un mes.")
    st.stop()
if len(meses_archivo) < 3:
    st.caption(f"El archivo solo trae {len(meses_archivo)} mes(es): "
               f"{', '.join(etiqueta_mes(m) for m in meses_archivo)}.")

src = src[src["__Mes__"].isin(meses_sel)]

# ── Comprobaciones ────────────────────────────────────────────
st.subheader("Comprobaciones")

total       = len(src)
sin_asignar = src[src["__Promotor__"] == "SIN ASIGNAR"]

c1, c2, c3 = st.columns(3)
c1.metric("Total operaciones", total)
c2.metric("Asignadas", total - len(sin_asignar))
c3.metric("Sin asignar", len(sin_asignar),
          delta=f"-{len(sin_asignar)}" if len(sin_asignar) else None,
          delta_color="inverse")

dist = (
    src[src["__Promotor__"] != "SIN ASIGNAR"]
    .assign(Mes=lambda d: d["__Mes__"].map(etiqueta_mes))
    .groupby(["__Promotor__", "__Mes__", "Mes"])
    .size()
    .reset_index(name="Operaciones")
    .sort_values(["__Promotor__", "__Mes__"])
    .rename(columns={"__Promotor__": "Promotor"})
    [["Promotor", "Mes", "Operaciones"]]
)
st.dataframe(dist, width="stretch", hide_index=True)

if len(sin_asignar) > 0:
    st.warning(f"Contratos no reconocidos: {sorted(str(c) for c in sin_asignar['Contrato'].unique())}")

sin_clave = sorted(set(src["Operador"]) - set(OP_TO_NAME) - {""})
if sin_clave:
    st.info(f"Claves de operador sin nombre en OP_TO_NAME (se escribirá la clave tal cual): {sin_clave}")

# ── Descargas ─────────────────────────────────────────────────
st.subheader("Descargar bitácoras")

grupos = sorted(
    (promotor, mes)
    for promotor, mes in src[["__Promotor__", "__Mes__"]].drop_duplicates().itertuples(index=False)
    if promotor != "SIN ASIGNAR"
)
archivos = []
for promotor, mes in grupos:
    df_p = src[(src["__Promotor__"] == promotor) & (src["__Mes__"] == mes)]
    archivos.append((f"{promotor} {etiqueta_mes(mes)}.xlsx", len(df_p), build_bitacora(df_p, layout_bytes)))

# ZIP con todos
zip_buf = io.BytesIO()
with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
    for nombre, _, contenido in archivos:
        zf.writestr(nombre, contenido)
zip_buf.seek(0)

etiqueta_periodo = "_".join(etiqueta_mes(m).replace(" ", "_") for m in sorted(meses_sel))
st.download_button(
    label=f"⬇️ Descargar todos en ZIP ({len(archivos)} archivos)",
    data=zip_buf,
    file_name=f"Bitacoras_{etiqueta_periodo}.zip",
    mime="application/zip",
    type="primary",
)

# Individuales
st.caption("O descarga por separado:")
for nombre, n_ops, contenido in archivos:
    st.download_button(
        label=f"📄 {nombre[:-5]}  ({n_ops} operaciones)",
        data=contenido,
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=nombre,
    )
