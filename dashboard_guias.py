import io
import re
import sys
from typing import Optional

import pandas as pd
import streamlit as st


try:
    from streamlit.runtime.scriptrunner import get_script_run_ctx
except Exception:
    get_script_run_ctx = None

if get_script_run_ctx is not None and get_script_run_ctx() is None:
    print("Esta es una aplicación de Streamlit.")
    print("Ejecuta el dashboard con:")
    print("    streamlit run dashboard_guias.py")
    sys.exit(0)


st.set_page_config(
    page_title="Dashboard de guías",
    page_icon="📦",
    layout="wide",
)


GUIDE_CANDIDATES = ["STR_REM_NUMERO", "GUIA", "GUÍA", "NUMERO_GUIA", "NRO_GUIA", "REMESA"]
STATUS_CANDIDATES = ["ESTADO", "ESTADO_GUIA", "STR_ESTADO"]
POPULATION_CANDIDATES = ["STR_CIU_ZONA", "POBLACION", "POBLACIÓN", "CIUDAD_DESTINO", "DESTINO", "ZONA"]
CITY_CANDIDATES = ["CIUDAD_ORIGEN", "CIUDAD", "ORIGEN"]
CUSTOMER_CANDIDATES = ["CLIENTE", "REMITENTE", "DESTINATARIO"]
DATE_CANDIDATES = ["FEC_REM_FECHA", "FECHA", "FEC_PRE_FECHA_ENTREGA"]
VALUE_CANDIDATES = ["NUM_REM_VALOR_TOTAL", "VALOR_TOTAL", "VALOR"]
WEIGHT_CANDIDATES = ["NUM_REM_PESO_COBRADO", "PESO", "PESO_COBRADO"]
UNITS_CANDIDATES = ["NUM_REM_UNIDADES", "UNIDADES"]


def normalize_col(col: object) -> str:
    """Convierte nombres de columnas en texto limpio."""
    col = str(col).strip()
    col = re.sub(r"\s+", "_", col)
    return col.upper()


def looks_like_wrong_header(df: pd.DataFrame) -> bool:
    """
    Detecta cuando pandas dejó columnas 0,1,2... y los nombres reales
    quedaron en la primera fila.
    """
    if df.empty:
        return False

    numeric_column_names = 0
    for col in df.columns:
        if isinstance(col, int) or str(col).strip().isdigit():
            numeric_column_names += 1

    first_row_values = [normalize_col(v) for v in df.iloc[0].tolist()]
    expected_headers = {
        "UID",
        "UID_REM",
        "STR_REM_NUMERO",
        "CLIENTE",
        "STR_CIU_ZONA",
        "ESTADO",
        "FEC_REM_FECHA",
        "FEC_PRE_FECHA_ENTREGA",
    }

    matches = sum(1 for value in first_row_values if value in expected_headers)

    return numeric_column_names >= max(3, len(df.columns) // 2) and matches >= 2


def promote_first_row_to_header(df: pd.DataFrame) -> pd.DataFrame:
    """Usa la primera fila como encabezado y elimina esa fila del dataframe."""
    new_columns = [normalize_col(c) for c in df.iloc[0].tolist()]
    df = df.iloc[1:].copy()
    df.columns = new_columns

    # Evitar columnas vacías o duplicadas.
    final_columns = []
    seen = {}
    for idx, col in enumerate(df.columns):
        if not col or col == "NAN":
            col = f"COLUMNA_{idx + 1}"

        if col in seen:
            seen[col] += 1
            col = f"{col}_{seen[col]}"
        else:
            seen[col] = 1

        final_columns.append(col)

    df.columns = final_columns
    return df


def find_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    """Busca una columna por nombre exacto o parcial."""
    normalized_map = {normalize_col(c): c for c in df.columns}

    for candidate in candidates:
        key = normalize_col(candidate)
        if key in normalized_map:
            return normalized_map[key]

    for candidate in candidates:
        key = normalize_col(candidate)
        for normalized_name, original_name in normalized_map.items():
            if key in normalized_name:
                return original_name

    return None


@st.cache_data(show_spinner=False)
def load_file(uploaded_file) -> pd.DataFrame:
    """
    Lee archivos Excel reales o archivos .xls exportados como HTML.
    Muchos sistemas generan .xls que internamente son tablas HTML.
    """
    raw = uploaded_file.read()
    suffix = uploaded_file.name.lower().split(".")[-1]

    # 1) Intentar Excel estándar.
    try:
        if suffix == "xlsx":
            df = pd.read_excel(io.BytesIO(raw), engine="openpyxl")
        elif suffix == "xls":
            df = pd.read_excel(io.BytesIO(raw), engine="xlrd")
        else:
            df = pd.read_excel(io.BytesIO(raw))
    except Exception:
        # 2) Intentar HTML exportado como Excel.
        # header=None es clave para poder detectar si la primera fila trae los encabezados.
        tables = pd.read_html(io.BytesIO(raw), header=None)
        if not tables:
            raise ValueError("No se encontraron tablas en el archivo.")
        df = tables[0]

    df = df.dropna(how="all").copy()

    # Si pandas dejó columnas 0,1,2... y los encabezados quedaron como fila 0.
    if looks_like_wrong_header(df):
        df = promote_first_row_to_header(df)
    else:
        df.columns = [normalize_col(c) for c in df.columns]

    df = df.dropna(how="all").copy()

    # Eliminar filas repetidas de encabezado que a veces aparecen en exportaciones HTML.
    if "UID" in df.columns:
        df = df[df["UID"].astype(str).str.upper() != "UID"]

    # Limpieza de texto.
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(["nan", "NaN", "None", ""]), col] = pd.NA

    # Convertir columnas numéricas conocidas cuando existan.
    for col in df.columns:
        if any(token in col for token in ["VALOR", "PESO", "UNIDADES", "PUNTOS"]):
            cleaned = (
                df[col]
                .astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.replace("$", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(cleaned, errors="coerce")

    # Convertir fechas probables.
    for col in df.columns:
        if col.startswith("FEC") or "FECHA" in col:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    return df


def multi_filter(label: str, df: pd.DataFrame, col: Optional[str]):
    if not col:
        return None

    values = (
        df[col]
        .dropna()
        .astype(str)
        .sort_values()
        .unique()
        .tolist()
    )

    return st.sidebar.multiselect(
        label,
        options=values,
        default=[],
        placeholder=f"Todas las opciones de {label.lower()}",
    )


def apply_text_filter(df: pd.DataFrame, search_text: str) -> pd.DataFrame:
    if not search_text.strip():
        return df

    text = search_text.strip().lower()
    mask = pd.Series(False, index=df.index)

    for col in df.columns:
        mask |= df[col].astype(str).str.lower().str.contains(text, na=False)

    return df[mask]


st.title("📦 Dashboard de guías y estados")
st.caption("Sube tu archivo Excel para filtrar por población, ciudad, cliente y estado.")

uploaded_file = st.file_uploader(
    "Archivo Excel",
    type=["xls", "xlsx", "html"],
    help="Acepta .xlsx, .xls real y .xls exportado como HTML.",
)

if uploaded_file is None:
    st.info("Sube un archivo para iniciar el dashboard.")
    st.stop()

try:
    df = load_file(uploaded_file)
except Exception as exc:
    st.error("No pude leer el archivo.")
    st.exception(exc)
    st.stop()

if df.empty:
    st.warning("El archivo no tiene registros para mostrar.")
    st.stop()

guide_col = find_column(df, GUIDE_CANDIDATES)
status_col = find_column(df, STATUS_CANDIDATES)
population_col = find_column(df, POPULATION_CANDIDATES)
city_col = find_column(df, CITY_CANDIDATES)
customer_col = find_column(df, CUSTOMER_CANDIDATES)
date_col = find_column(df, DATE_CANDIDATES)
value_col = find_column(df, VALUE_CANDIDATES)
weight_col = find_column(df, WEIGHT_CANDIDATES)
units_col = find_column(df, UNITS_CANDIDATES)

st.sidebar.header("Filtros")

population_filter = multi_filter("Población / zona", df, population_col)
city_filter = multi_filter("Ciudad origen", df, city_col)
status_filter = multi_filter("Estado", df, status_col)
customer_filter = multi_filter("Cliente / remitente", df, customer_col)

search_text = st.sidebar.text_input(
    "Buscar guía, destinatario, dirección, etc.",
    placeholder="Ej: 123456, Pérez, Calle 10...",
)

filtered = df.copy()

if population_col and population_filter:
    filtered = filtered[filtered[population_col].astype(str).isin(population_filter)]

if city_col and city_filter:
    filtered = filtered[filtered[city_col].astype(str).isin(city_filter)]

if status_col and status_filter:
    filtered = filtered[filtered[status_col].astype(str).isin(status_filter)]

if customer_col and customer_filter:
    filtered = filtered[filtered[customer_col].astype(str).isin(customer_filter)]

filtered = apply_text_filter(filtered, search_text)

# KPIs
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

total_guides = filtered[guide_col].nunique() if guide_col else len(filtered)
kpi1.metric("Guías", f"{total_guides:,.0f}".replace(",", "."))

if units_col and pd.api.types.is_numeric_dtype(filtered[units_col]):
    kpi2.metric("Unidades", f"{filtered[units_col].sum():,.0f}".replace(",", "."))
else:
    kpi2.metric("Registros", f"{len(filtered):,.0f}".replace(",", "."))

if weight_col and pd.api.types.is_numeric_dtype(filtered[weight_col]):
    kpi3.metric("Peso cobrado", f"{filtered[weight_col].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
else:
    kpi3.metric("Estados", filtered[status_col].nunique() if status_col else "N/D")

if value_col and pd.api.types.is_numeric_dtype(filtered[value_col]):
    kpi4.metric("Valor total", f"$ {filtered[value_col].sum():,.0f}".replace(",", "."))
else:
    kpi4.metric("Poblaciones", filtered[population_col].nunique() if population_col else "N/D")

st.divider()

left, right = st.columns(2)

with left:
    st.subheader("Guías por estado")
    if status_col:
        status_chart = (
            filtered[status_col]
            .fillna("SIN ESTADO")
            .astype(str)
            .value_counts()
            .reset_index()
        )
        status_chart.columns = ["Estado", "Cantidad"]
        st.bar_chart(status_chart, x="Estado", y="Cantidad")
    else:
        st.warning("No encontré una columna de estado.")

with right:
    st.subheader("Guías por población / zona")
    if population_col:
        population_chart = (
            filtered[population_col]
            .fillna("SIN POBLACIÓN")
            .astype(str)
            .value_counts()
            .head(20)
            .reset_index()
        )
        population_chart.columns = ["Población / zona", "Cantidad"]
        st.bar_chart(population_chart, x="Población / zona", y="Cantidad")
    else:
        st.warning("No encontré una columna de población/zona.")

st.divider()

st.subheader("Detalle de guías")

# Orden solicitado para el detalle:
# 1. STR_REM_NUMERO
# 2. ESTADO
# 3. STR_CIU_ZONA
# 4. CIUDAD_ORIGEN
# 5. DESTINATARIO
# 6. DIRECCION
# 7. Resto de columnas
priority_columns = [
    guide_col,
    status_col,
    population_col,
    city_col,
    "DESTINATARIO" if "DESTINATARIO" in filtered.columns else None,
    "DIRECCION" if "DIRECCION" in filtered.columns else None,
]

preferred_columns = []
for col in priority_columns:
    if col and col in filtered.columns and col not in preferred_columns:
        preferred_columns.append(col)

for col in filtered.columns:
    if col not in preferred_columns:
        preferred_columns.append(col)

st.dataframe(
    filtered[preferred_columns],
    use_container_width=True,
    hide_index=True,
)

csv = filtered.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "Descargar resultado filtrado en CSV",
    data=csv,
    file_name="guias_filtradas.csv",
    mime="text/csv",
)

with st.expander("Columnas detectadas"):
    st.write(
        {
            "guía": guide_col,
            "estado": status_col,
            "población/zona": population_col,
            "ciudad origen": city_col,
            "cliente/remitente": customer_col,
            "fecha": date_col,
            "valor": value_col,
            "peso": weight_col,
            "unidades": units_col,
        }
    )
    st.dataframe(pd.DataFrame({"columnas": df.columns}), use_container_width=True, hide_index=True)
