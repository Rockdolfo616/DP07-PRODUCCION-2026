import gc
import os
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path

import psutil

import pandas as pd
import streamlit as st

# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(page_title="Producción - Consultas y Atenciones", layout="wide")
st.title("TOTAL DE CONSULTAS Y ATENCIONES")
st.caption("Versión optimizada para bases Excel grandes")

# Columnas del dashboard
COL_ESP = "PROF_ESP_ATE"
COL_CIE10 = "ATEMED_CIE10"
COL_ATENCION = "ATEMED_ID"
COL_TIPO = "ATEMED_TIP_DIAG"
COL_CRON = "ATEMED_CRON_DIAG"
COL_CON_DIAG = "ATEMED_CON_DIAG"
COL_DES_CIE10 = "ATEMED_DES_CIE10"
COL_FECHA = "ATEMED_FEC_INI"
COL_PROV = "ENT_DES_PROV"
COL_CANTON = "ENT_DES_CANT"
COL_PARR = "ENT_DES_PARR"
COL_DIST = "ENT_DIST_DIS"
COL_EST = "ENT_NOM"
COL_NIVEL = "ENT_NIV"
COL_NAC = "PCTE_NACIONALIDAD"
COL_SEXO = "PCTE_SEXO"
COL_ETNIA = "PCTE_AUTID_ETN"
COL_GRP_PRI = "PCTE_GRP_PRI"
COL_PCTE_ID = "PCTE_IDE"
COL_EDAD = "PCTE_ANIOS"

COLUMNAS_USAR = [
    COL_ESP, COL_CIE10, COL_ATENCION, COL_TIPO, COL_CRON, COL_CON_DIAG, COL_DES_CIE10, COL_FECHA,
    COL_PROV, COL_CANTON, COL_PARR, COL_DIST, COL_EST, COL_NIVEL, COL_NAC, COL_SEXO, COL_ETNIA,
    COL_GRP_PRI, COL_PCTE_ID, COL_EDAD,
]

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

VACIOS = ["", "NAN", "NONE", "NULL", "<NA>"]

# ============================================================
# FUNCIONES
# ============================================================
def limpiar_texto(serie):
    s = serie.astype("string").str.strip()
    return s.mask(s.str.upper().isin(VACIOS))


def opciones(serie):
    s = limpiar_texto(serie).dropna()
    return sorted(s.astype(str).unique().tolist())


def formato_numero(x):
    try:
        return f"{int(x):,}".replace(",", ".")
    except (ValueError, TypeError):
        return x


def estilo_tabla(df):
    return (
        df.style
        .format(formato_numero)
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("font-weight", "bold"), ("text-align", "center"),
                ("vertical-align", "middle"), ("background-color", "#FFFFFF"),
                ("border-bottom", "1px solid #000000")
            ]},
            {"selector": "tbody th", "props": [("text-align", "left")]},
            {"selector": "tbody td", "props": [("text-align", "center")]},
            {"selector": "tbody tr:last-child th", "props": [
                ("font-weight", "bold"), ("border-top", "1px solid #000000")
            ]},
            {"selector": "tbody tr:last-child td", "props": [
                ("font-weight", "bold"), ("border-top", "1px solid #000000")
            ]},
        ])
    )


def preparar_base(df):
    # Texto necesario para filtros y agrupaciones
    texto_cols = [COL_ESP, COL_CIE10, COL_ATENCION, COL_TIPO, COL_CRON,
                  COL_PROV, COL_CANTON, COL_PARR, COL_DIST, COL_EST, COL_NIVEL, COL_NAC, COL_SEXO, COL_ETNIA,
                  COL_GRP_PRI, COL_PCTE_ID]
    for c in texto_cols:
        df[c] = limpiar_texto(df[c])

    # Fecha se procesa UNA sola vez
    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], errors="coerce", dayfirst=True)
    df["_ANIO"] = df[COL_FECHA].dt.year.astype("Int16")
    df["_MES_NUM"] = df[COL_FECHA].dt.month.astype("Int8")

    # Edad numérica para tablas de pacientes únicos
    df[COL_EDAD] = pd.to_numeric(df[COL_EDAD], errors="coerce").astype("Float32")

    # Categorías de baja cardinalidad: menor RAM
    categoricas = [COL_ESP, COL_TIPO, COL_CRON, COL_PROV, COL_CANTON,
                   COL_PARR, COL_DIST, COL_EST, COL_NIVEL, COL_NAC, COL_ETNIA]
    for c in categoricas:
        df[c] = df[c].fillna("No Especificado").astype("category")
    return df


def motor_excel_disponible():
    """Prefiere Calamine (Rust) si está instalado; conserva openpyxl como fallback."""
    try:
        import python_calamine  # noqa: F401
        return "calamine"
    except ImportError:
        return "openpyxl"


@st.cache_resource(show_spinner=False, max_entries=1)
def cargar_excel_desde_ruta(ruta, tamano, mtime):
    """Cachea la base consolidada. tamano/mtime invalidan caché si cambia el archivo."""
    del tamano, mtime
    motor = motor_excel_disponible()
    xls = pd.ExcelFile(ruta, engine=motor)
    hojas = [h for h in xls.sheet_names if h.strip().lower() != "catalogo"]

    # La hoja Catalogo se excluye: el dashboard usa ATEMED_DES_CIE10 directamente.

    partes = []
    progreso = st.progress(0, text="Leyendo hojas del Excel...")

    for i, hoja in enumerate(hojas, start=1):
        # usecols callable evita cargar al DataFrame las columnas que el dashboard no usa
        temp = pd.read_excel(
            xls,
            sheet_name=hoja,
            usecols=lambda c: str(c).strip() in COLUMNAS_USAR,
        )
        temp.columns = temp.columns.astype(str).str.strip()
        if len(temp.columns):
            partes.append(temp)
        progreso.progress(i / len(hojas), text=f"Procesando hoja {i} de {len(hojas)}: {hoja}")

    progreso.empty()
    xls.close()
    if not partes:
        raise ValueError("No se encontraron hojas con las columnas requeridas.")

    df = pd.concat(partes, ignore_index=True, copy=False)
    del partes
    gc.collect()

    faltantes = [c for c in COLUMNAS_USAR if c not in df.columns]
    if faltantes:
        raise ValueError("Faltan columnas necesarias en el Excel: " + ", ".join(faltantes))

    df = df[COLUMNAS_USAR]
    df = preparar_base(df)
    gc.collect()
    return df, hojas


def guardar_upload_temporal(uploaded):
    """Guarda el upload en disco sin usar getvalue(), evitando una copia gigante en RAM."""
    clave = f"{uploaded.name}_{uploaded.size}"
    if st.session_state.get("upload_clave") == clave:
        ruta = st.session_state.get("upload_ruta")
        if ruta and os.path.exists(ruta):
            return ruta

    sufijo = Path(uploaded.name).suffix or ".xlsx"
    fd, ruta = tempfile.mkstemp(prefix="produ_", suffix=sufijo)
    os.close(fd)
    uploaded.seek(0)
    with open(ruta, "wb") as salida:
        while True:
            bloque = uploaded.read(8 * 1024 * 1024)
            if not bloque:
                break
            salida.write(bloque)
    st.session_state["upload_clave"] = clave
    st.session_state["upload_ruta"] = ruta
    return ruta


def aplicar_filtro(df, columna, valor):
    if valor == "Todos":
        return df
    return df.loc[df[columna] == valor]


def tabla_mensual_unicos(df_consultas, columna_fila, nombre_indice, orden=None):
    base = df_consultas.loc[
        df_consultas["_MES_NUM"].notna(),
        [columna_fila, COL_ATENCION, "_MES_NUM"]
    ]

    tabla = (
        base.groupby([columna_fila, "_MES_NUM"], observed=True)[COL_ATENCION]
        .nunique()
        .unstack(fill_value=0)
        .reindex(columns=range(1, 13), fill_value=0)
    )
    tabla.columns = [MESES[m] for m in tabla.columns]

    if orden is not None:
        tabla = tabla.reindex(orden, fill_value=0)
    else:
        tabla = tabla.sort_index()

    total_fila = base.groupby(columna_fila, observed=True)[COL_ATENCION].nunique()
    tabla["Total"] = tabla.index.to_series().map(total_fila).fillna(0).astype(int)

    totales_mes = base.groupby("_MES_NUM", observed=True)[COL_ATENCION].nunique()
    total_general = {MESES[m]: int(totales_mes.get(m, 0)) for m in range(1, 13)}
    total_general["Total"] = int(base[COL_ATENCION].nunique())
    tabla.loc["Total general"] = total_general
    tabla = tabla.fillna(0).astype(int)
    tabla.index.name = nombre_indice
    return tabla


# ============================================================
# CARGA DEL ARCHIVO
# ============================================================
st.info(
    "Para archivos cercanos a 1 GB se recomienda **Ruta local**. Así Streamlit no necesita "
    "mantener una copia adicional del Excel dentro del cargador."
)

modo = st.radio(
    "Origen del archivo",
    ["Ruta local (recomendado para archivos grandes)", "Cargar archivo"],
    horizontal=True,
)

ruta_excel = None
if modo.startswith("Ruta local"):
    ruta_ingresada = st.text_input(
        "Ruta completa del Excel",
        placeholder=r"C:\Users\D01-ESTADISTICA\Desktop\PRODU\DP_07_Produccion.xlsx",
    ).strip().strip('"')
    if not ruta_ingresada:
        st.stop()
    if not os.path.isfile(ruta_ingresada):
        st.error("No se encontró el archivo en esa ruta.")
        st.stop()
    ruta_excel = ruta_ingresada
else:
    archivo_excel = st.file_uploader("Seleccione el archivo Excel", type=["xlsx", "xls"])
    if archivo_excel is None:
        st.stop()
    ruta_excel = guardar_upload_temporal(archivo_excel)

t_total_inicio = time.perf_counter()
proceso = psutil.Process(os.getpid())
ram_inicio_mb = proceso.memory_info().rss / 1024**2

try:
    stat = os.stat(ruta_excel)
    t_lectura_inicio = time.perf_counter()
    with st.spinner("Cargando únicamente las columnas necesarias. La primera carga puede tardar..."):
        df_original, hojas_procesadas = cargar_excel_desde_ruta(
            ruta_excel, stat.st_size, stat.st_mtime_ns
        )
    tiempo_lectura = time.perf_counter() - t_lectura_inicio
except Exception as e:
    st.error(f"Error al procesar el Excel: {e}")
    st.stop()

memoria_mb = df_original.memory_usage(deep=True).sum() / 1024**2
st.success(
    f"Base preparada: {len(df_original):,} registros | "
    f"{len(df_original.columns)} columnas | ~{memoria_mb:,.0f} MB en memoria".replace(",", ".")
)

# ============================================================
# FILTROS EN CASCADA
# ============================================================
st.sidebar.title("FILTRO")
df_filtrado = df_original

anios = sorted(df_filtrado["_ANIO"].dropna().astype(int).unique().tolist())
anio = st.sidebar.selectbox("Año", ["Todos"] + anios)
if anio != "Todos":
    df_filtrado = df_filtrado.loc[df_filtrado["_ANIO"] == int(anio)]

filtros = [
    ("Provincia", COL_PROV),
    ("Cantón", COL_CANTON),
    ("Parroquia", COL_PARR),
    ("Distrito", COL_DIST),
    ("Establecimiento", COL_EST),
    ("Nivel de Atención", COL_NIVEL),
]
selecciones = {"Año": anio}
for etiqueta, columna in filtros:
    valor = st.sidebar.selectbox(etiqueta, ["Todos"] + opciones(df_filtrado[columna]), key=f"f_{columna}")
    selecciones[etiqueta] = valor
    df_filtrado = aplicar_filtro(df_filtrado, columna, valor)

if df_filtrado.empty:
    st.warning("No existen registros para los filtros seleccionados.")
    st.stop()

st.sidebar.caption(f"Registros seleccionados: {len(df_filtrado):,}".replace(",", "."))
with st.sidebar.expander("Filtros aplicados"):
    for k, v in selecciones.items():
        st.write(f"**{k}:** {v}")

# ============================================================
# BASE DE CONSULTAS / ATENCIONES
# ============================================================
# No se copia toda la base: se seleccionan únicamente columnas necesarias para las tablas.
cols_consultas = [
    COL_ESP, COL_CIE10, COL_ATENCION, COL_TIPO, COL_CRON,
    COL_NIVEL, COL_NAC, COL_CANTON, COL_SEXO, "_ANIO", "_MES_NUM"
]
df_consultas = df_filtrado.loc[df_filtrado[COL_CIE10].notna(), cols_consultas].copy()

total_consultas = len(df_consultas)
total_atenciones_general = df_consultas[COL_ATENCION].nunique()

# Clasificaciones vectorizadas (más rápidas que .apply() fila por fila)
tipo = df_consultas[COL_TIPO].astype("string").str.upper()
cron = df_consultas[COL_CRON].astype("string").str.upper()

df_consultas["GRUPO_CONSULTA"] = "No Clasificado"
df_consultas.loc[tipo.str.contains("MORBIL", na=False), "GRUPO_CONSULTA"] = "Morbilidad"
df_consultas.loc[tipo.str.contains("PREVEN", na=False), "GRUPO_CONSULTA"] = "Prevención"

df_consultas["CRON_DIAG_CLASIFICACION"] = "No Clasificado"
df_consultas.loc[cron.str.contains("PRIMERA", na=False), "CRON_DIAG_CLASIFICACION"] = "Primera"
df_consultas.loc[cron.str.contains("SUBSECUENTE", na=False), "CRON_DIAG_CLASIFICACION"] = "Subsecuente"

# ============================================================
# RESUMEN GENERAL
# ============================================================
st.subheader("RESUMEN GENERAL")
c1, c2 = st.columns(2)
c1.metric("Total de Consultas", formato_numero(total_consultas))
c2.metric("Total de Atenciones", formato_numero(total_atenciones_general))

activos = [f"{k}: {v}" for k, v in selecciones.items() if v != "Todos"]
if activos:
    st.caption(" | ".join(activos))

# ============================================================
# TABLA 1: CONSULTAS Y ATENCIONES POR ESPECIALIDAD
# ============================================================
st.markdown("---")
st.subheader("Total de consultas y atenciones por Especialidad")

clasificables = df_consultas.loc[
    df_consultas["GRUPO_CONSULTA"].isin(["Morbilidad", "Prevención"])
]

pivot = pd.pivot_table(
    clasificables,
    index=COL_ESP,
    columns=["GRUPO_CONSULTA", "CRON_DIAG_CLASIFICACION"],
    values=COL_CIE10,
    aggfunc="count",
    fill_value=0,
    observed=True,
)

base_cols = pd.MultiIndex.from_product([
    ["Morbilidad", "Prevención"], ["Primera", "Subsecuente"]
])
pivot = pivot.reindex(columns=base_cols, fill_value=0)
pivot[("Morbilidad", "Total Morbilidad")] = pivot["Morbilidad"].sum(axis=1)
pivot[("Prevención", "Total Prevención")] = pivot["Prevención"].sum(axis=1)
pivot[("Total", "Total general")] = (
    pivot[("Morbilidad", "Total Morbilidad")] + pivot[("Prevención", "Total Prevención")]
)

orden = pd.MultiIndex.from_tuples([
    ("Morbilidad", "Primera"), ("Morbilidad", "Subsecuente"),
    ("Morbilidad", "Total Morbilidad"), ("Prevención", "Primera"),
    ("Prevención", "Subsecuente"), ("Prevención", "Total Prevención"),
    ("Total", "Total general"), ("Total", "Atenciones"),
])
pivot = pivot.reindex(columns=orden, fill_value=0)

atenciones_esp = df_consultas.groupby(COL_ESP, observed=True)[COL_ATENCION].nunique()
pivot[("Total", "Atenciones")] = pivot.index.to_series().map(atenciones_esp).fillna(0)
fila = pivot.sum(axis=0)
fila.name = "Total general"
tabla_especialidad = pd.concat([pivot, fila.to_frame().T])
tabla_especialidad.loc["Total general", ("Total", "Atenciones")] = total_atenciones_general
tabla_especialidad = tabla_especialidad.fillna(0).astype(int)
tabla_especialidad.index.name = "Especialidad"

st.dataframe(estilo_tabla(tabla_especialidad), use_container_width=True)
st.caption("Fuente: PRAS 2026")

total_tabla = int(tabla_especialidad.loc["Total general", ("Total", "Total general")])
diferencia = total_consultas - total_tabla

# ============================================================
# TABLA 2: ATENCIONES POR NIVEL Y MES
# ============================================================
st.markdown("---")
st.subheader("Total de Atenciones por Nivel de Atención y Mes")
tabla_nivel = tabla_mensual_unicos(df_consultas, COL_NIVEL, "Nivel de Atención")
st.dataframe(estilo_tabla(tabla_nivel).set_properties(subset=["Total"], **{"font-weight": "bold"}), use_container_width=True)
st.caption("Fuente: PRAS 2026")

# ============================================================
# TABLA 3: ATENCIONES POR NACIONALIDAD Y MES
# ============================================================
st.markdown("---")
st.subheader("Total de Atenciones por Nacionalidad y Mes")

nac = df_consultas.loc[:, [COL_NAC, COL_ATENCION, "_MES_NUM"]].copy()
nac_txt = nac[COL_NAC].astype("string").str.strip().str.upper()
nac["NACIONALIDAD_AGRUPADA"] = "NAC_OTROS/AS"
nac.loc[nac_txt.isin(["ECUATORIANO/A", "ECUATOGUINEANO/NA"]), "NACIONALIDAD_AGRUPADA"] = "ECUATORIANO/A"
for valor in ["COLOMBIANO/A", "PERUANO/A", "CUBANO/A", "VENEZOLANO/A"]:
    nac.loc[nac_txt == valor, "NACIONALIDAD_AGRUPADA"] = valor

orden_nac = [
    "ECUATORIANO/A", "COLOMBIANO/A", "PERUANO/A",
    "CUBANO/A", "VENEZOLANO/A", "NAC_OTROS/AS"
]
tabla_nacionalidad = tabla_mensual_unicos(nac, "NACIONALIDAD_AGRUPADA", "Nacionalidad", orden_nac)
st.dataframe(estilo_tabla(tabla_nacionalidad).set_properties(subset=["Total"], **{"font-weight": "bold"}), use_container_width=True)
st.caption("Fuente: PRAS 2026")

# ============================================================
# TABLA 4: ATENCIONES POR CANTÓN Y MES
# ============================================================
st.markdown("---")
st.subheader("Total de Atenciones por Cantón y Mes")
tabla_canton = tabla_mensual_unicos(df_consultas, COL_CANTON, "Cantón")
st.dataframe(estilo_tabla(tabla_canton).set_properties(subset=["Total"], **{"font-weight": "bold"}), use_container_width=True)
st.caption("Fuente: PRAS 2026")

# ============================================================
# TABLA 5: NUMERO DE EMBARAZADAS POR CANTÓN Y GRUPO DE EDAD
# ============================================================
st.markdown("---")
st.subheader("Numero de Embarazadas")

# Esta tabla trabaja con pacientes únicos (PCTE_IDE), no con consultas.
# Para evitar que el filtro de establecimiento asigne una misma paciente
# a distintos establecimientos, se reconstruye la base con todos los filtros
# activos EXCEPTO ENT_NOM. Después cada paciente se asigna al establecimiento
# donde registra el mayor número de filas. Si el usuario seleccionó un
# establecimiento, el filtro se aplica sobre esa asignación final.

def normalizar_identificacion_serie(serie):
    """Versión vectorizada equivalente a normalizar_identificacion()."""
    s = serie.astype("string").str.strip().str.upper()
    s = s.mask(s.isna() | s.isin(VACIOS))

    # Equivalente a quitar el .0 terminal cuando el contenido previo es numérico.
    prefijo = s.str[:-2].str.replace(".", "", regex=False)
    quitar_decimal = s.str.endswith(".0", na=False) & prefijo.str.fullmatch(r"\d+", na=False)
    s = s.mask(quitar_decimal, s.str[:-2])

    # Si solo contiene dígitos, espacios, puntos o guiones, conservar únicamente dígitos.
    compacto = s.str.replace(r"[ .-]", "", regex=True)
    solo_numerico = compacto.str.fullmatch(r"\d+", na=False)
    digitos = s.str.replace(r"\D", "", regex=True)
    resultado = s.mask(solo_numerico, digitos)

    nueve = solo_numerico & resultado.str.len().eq(9)
    resultado = resultado.mask(nueve, "0" + resultado)
    return resultado


# Reconstruir el universo aplicando los filtros actuales excepto Establecimiento.
df_emb = df_original
if selecciones.get("Año") != "Todos":
    df_emb = df_emb.loc[df_emb["_ANIO"] == int(selecciones["Año"])]

for etiqueta, columna in [
    ("Provincia", COL_PROV),
    ("Cantón", COL_CANTON),
    ("Parroquia", COL_PARR),
    ("Distrito", COL_DIST),
    ("Nivel de Atención", COL_NIVEL),
]:
    valor = selecciones.get(etiqueta, "Todos")
    if valor != "Todos":
        df_emb = aplicar_filtro(df_emb, columna, valor)

# Criterios solicitados:
# 1) ATEMED_CIE10 informado.
# 2) PCTE_GRP_PRI contiene la palabra "Embarazadas".
# 3) PCTE_IDE informado para poder identificar paciente único.
grp_pri_txt = df_emb[COL_GRP_PRI].astype("string").str.strip().str.upper()
mask_emb = (
    df_emb[COL_CIE10].notna()
    & grp_pri_txt.str.contains("EMBARAZADAS", na=False)
)

base_emb = df_emb.loc[
    mask_emb,
    [COL_PCTE_ID, COL_EDAD, COL_EST, COL_CANTON]
].copy()

base_emb["PACIENTE_ID"] = normalizar_identificacion_serie(base_emb[COL_PCTE_ID])
base_emb = base_emb.loc[base_emb["PACIENTE_ID"].notna()].copy()
base_emb[COL_EDAD] = pd.to_numeric(base_emb[COL_EDAD], errors="coerce")

if base_emb.empty:
    st.info("No existen registros de embarazadas para los filtros seleccionados.")
else:
    # --------------------------------------------------------
    # ESTABLECIMIENTO PRINCIPAL DE CADA PACIENTE
    # --------------------------------------------------------
    # Se cuenta cuántos registros tiene cada paciente en cada
    # establecimiento/cantón. Se elige el establecimiento con más registros.
    # En empate, el orden alfabético hace la selección reproducible.
    conteo_est = (
        base_emb
        .groupby(["PACIENTE_ID", COL_EST, COL_CANTON], observed=True, dropna=False)
        .size()
        .reset_index(name="N_REGISTROS")
    )
    conteo_est[COL_EST] = conteo_est[COL_EST].astype("string").fillna("No Especificado")
    conteo_est[COL_CANTON] = conteo_est[COL_CANTON].astype("string").fillna("No Especificado")
    conteo_est = conteo_est.sort_values(
        ["PACIENTE_ID", "N_REGISTROS", COL_EST, COL_CANTON],
        ascending=[True, False, True, True]
    )
    principal = conteo_est.drop_duplicates("PACIENTE_ID", keep="first").copy()

    # Edad de ubicación = edad MÁS ALTA registrada para la paciente.
    edad_max = (
        base_emb.groupby("PACIENTE_ID", observed=True)[COL_EDAD]
        .max()
        .rename("EDAD_MAX")
    )
    pacientes = principal.merge(edad_max, on="PACIENTE_ID", how="left")

    # Si se solicitó un establecimiento específico, se cuenta únicamente
    # a las pacientes cuyo establecimiento principal es el seleccionado.
    establecimiento_sel = selecciones.get("Establecimiento", "Todos")
    if establecimiento_sel != "Todos":
        pacientes = pacientes.loc[
            pacientes[COL_EST].astype("string") == str(establecimiento_sel)
        ].copy()

    # --------------------------------------------------------
    # GRUPOS DE EDAD
    # --------------------------------------------------------
    pacientes["GRUPO_EDAD"] = pd.cut(
        pacientes["EDAD_MAX"],
        bins=[9, 14, 19, float("inf")],
        labels=[
            "Adolescentes 10 a 14 años",
            "Adolescentes 15 a 19 años",
            "20 en adelante",
        ],
        right=True,
    )
    pacientes = pacientes.loc[pacientes["GRUPO_EDAD"].notna()].copy()

    orden_edades = [
        "Adolescentes 10 a 14 años",
        "Adolescentes 15 a 19 años",
        "20 en adelante",
    ]

    if pacientes.empty:
        st.info("No existen pacientes únicas dentro de los grupos de edad solicitados.")
    else:
        # Cada fila de 'pacientes' ya representa UNA paciente única.
        tabla_embarazadas = (
            pacientes
            .groupby([COL_CANTON, "GRUPO_EDAD"], observed=True)["PACIENTE_ID"]
            .nunique()
            .unstack(fill_value=0)
            .reindex(columns=orden_edades, fill_value=0)
            .sort_index()
        )

        # Total por cantón = pacientes únicas, no suma ciega de registros.
        total_canton_emb = pacientes.groupby(COL_CANTON, observed=True)["PACIENTE_ID"].nunique()
        tabla_embarazadas["Total"] = (
            tabla_embarazadas.index.to_series()
            .map(total_canton_emb)
            .fillna(0)
            .astype(int)
        )

        # Total general por grupo y total de pacientes únicas.
        fila_total_emb = {
            grupo: pacientes.loc[pacientes["GRUPO_EDAD"] == grupo, "PACIENTE_ID"].nunique()
            for grupo in orden_edades
        }
        fila_total_emb["Total"] = pacientes["PACIENTE_ID"].nunique()
        tabla_embarazadas.loc["Total general"] = fila_total_emb
        tabla_embarazadas = tabla_embarazadas.fillna(0).astype(int)
        tabla_embarazadas.index.name = "Cantón"

        st.dataframe(
            estilo_tabla(tabla_embarazadas).set_properties(
                subset=["Total"], **{"font-weight": "bold"}
            ),
            use_container_width=True
        )
        st.caption("Fuente: PRAS 2026")
        st.caption(
            "Criterio: pacientes únicas por PCTE_IDE; cédulas de 9 dígitos se homologan "
            "anteponiendo 0; edad = máxima PCTE_ANIOS; establecimiento = el de mayor "
            "número de registros por paciente."
        )


# ============================================================
# TABLAS 6-9: PERFILES DE MORBILIDAD
# ============================================================
st.markdown("---")
st.header("PERFILES DE MORBILIDAD")

numero_causas = st.selectbox(
    "Filtro",
    [10, 20, 30],
    index=0,
    format_func=lambda x: f"{x} Primeras Causas",
    key="numero_causas_morbilidad",
)

ESP_EXCLUIDAS_MORBILIDAD = {
    "AUDILOGÍA / FONIATRÍA", "AUDIOLOGÍA / FONIATRÍA",
    "FISIATRÍA", "MEDICINA DEL TRABAJO", "MEDICINA OCUPACIONAL",
    "NUTRICIÓN", "NUTRICIÓN CLÍNICA", "ODONTOLOGÍA", "ODONTOLOGIA",
    "ODONTOLOGIA RURAL", "ODONTOLOGÍA RURAL",
    "PSICOREHABILITACION", "PSICOREHABILITACIÓN",
}

ESP_ODONTOLOGIA = {
    "ODONTOLOGÍA", "ODONTOLOGIA", "ODONTOLOGIA RURAL", "ODONTOLOGÍA RURAL"
}

ETNIAS_ORDEN = [
    "Afrodescendiente", "Blanco", "Indígena", "Mestizos", "Montubio",
    "Mulato", "Negro", "No Aplica", "No sabe/ no responde", "Otro"
]


def normalizar_mayusculas(serie):
    return serie.astype("string").str.strip().str.upper()


def sexo_h_m(serie):
    s = normalizar_mayusculas(serie)
    salida = pd.Series(pd.NA, index=serie.index, dtype="string")
    salida.loc[s.isin(["H", "HOMBRE", "MASCULINO", "MASCULINO/A"])] = "H"
    salida.loc[s.isin(["M", "MUJER", "F", "FEMENINO", "FEMENINO/A"])] = "M"
    return salida


def etnia_agrupada(serie):
    """Normaliza PCTE_AUTID_ETN a las columnas mostradas en el formato adjunto."""
    s = normalizar_mayusculas(serie)
    out = pd.Series("Otro", index=serie.index, dtype="string")
    out.loc[s.str.contains("AFRO", na=False)] = "Afrodescendiente"
    out.loc[s.str.contains("BLANCO", na=False)] = "Blanco"
    out.loc[s.str.contains("IND[IÍ]GEN", regex=True, na=False)] = "Indígena"
    out.loc[s.str.contains("MESTIZ", na=False)] = "Mestizos"
    out.loc[s.str.contains("MONTUB", na=False)] = "Montubio"
    out.loc[s.str.contains("MULAT", na=False)] = "Mulato"
    out.loc[s.str.contains("NEGRO", na=False)] = "Negro"
    out.loc[s.str.contains("NO APLICA", na=False)] = "No Aplica"
    out.loc[s.str.contains("NO SABE|NO RESPONDE|NO SABE/NO RESPONDE", regex=True, na=False)] = "No sabe/ no responde"
    return out


def perfil_morbilidad(base, nombre, modo):
    """
    Perfil por CONSULTAS aplicando, en TODOS los perfiles:
      - ATEMED_CIE10 informado.
      - ATEMED_TIP_DIAG = Morbilidad.
      - ATEMED_CRON_DIAG = Primera.
      - ATEMED_CON_DIAG contiene "Definitivo".
      - Descripción tomada de ATEMED_DES_CIE10.

    El porcentaje SIEMPRE es:
        TOTAL (H + M) del CIE10 / TOTAL GENERAL (H + M) de esta tabla * 100.
    """
    columnas = [
        COL_CIE10, COL_DES_CIE10, COL_ESP, COL_SEXO, COL_ETNIA,
        COL_TIPO, COL_CRON, COL_CON_DIAG,
    ]
    trabajo = base.loc[base[COL_CIE10].notna(), columnas].copy()

    trabajo["CIE10"] = normalizar_mayusculas(trabajo[COL_CIE10])
    trabajo["DESCRIPCION"] = limpiar_texto(trabajo[COL_DES_CIE10])
    trabajo["SEXO_HM"] = sexo_h_m(trabajo[COL_SEXO])
    trabajo["ETNIA"] = etnia_agrupada(trabajo[COL_ETNIA])

    # --------------------------------------------------------
    # CRITERIOS COMUNES OBLIGATORIOS PARA LOS 4 PERFILES
    # --------------------------------------------------------
    tipo_diag = normalizar_mayusculas(trabajo[COL_TIPO])
    cron_diag = normalizar_mayusculas(trabajo[COL_CRON])
    con_diag = normalizar_mayusculas(trabajo[COL_CON_DIAG])

    trabajo = trabajo.loc[
        trabajo["CIE10"].notna()
        & (trabajo["CIE10"] != "")
        & tipo_diag.eq("MORBILIDAD")
        & cron_diag.eq("PRIMERA")
        & con_diag.str.contains("DEFINITIVO", na=False)
    ].copy()

    # --------------------------------------------------------
    # CRITERIOS ESPECÍFICOS DE CADA PERFIL
    # --------------------------------------------------------
    esp = normalizar_mayusculas(trabajo[COL_ESP])

    if modo in {"general", "materno", "mental"}:
        trabajo = trabajo.loc[~esp.isin(ESP_EXCLUIDAS_MORBILIDAD)].copy()
        esp = normalizar_mayusculas(trabajo[COL_ESP])

    if modo == "general":
        trabajo = trabajo.loc[~trabajo["CIE10"].str.startswith("R", na=False)].copy()
    elif modo == "materno":
        trabajo = trabajo.loc[trabajo["CIE10"].str.startswith("O", na=False)].copy()
    elif modo == "mental":
        trabajo = trabajo.loc[trabajo["CIE10"].str.startswith("F", na=False)].copy()
    elif modo == "odontologico":
        trabajo = trabajo.loc[esp.isin(ESP_ODONTOLOGIA)].copy()

    # Para HOMBRE/MUJER/TOTAL se consideran sexos reconocidos como H o M.
    trabajo = trabajo.loc[trabajo["SEXO_HM"].isin(["H", "M"])].copy()

    st.subheader(nombre)
    if trabajo.empty:
        st.info("No existen registros para los filtros y criterios seleccionados.")
        return

    # --------------------------------------------------------
    # DESCRIPCIÓN CIE-10 DESDE ATEMED_DES_CIE10
    # Si un mismo código tiene más de una descripción, se usa
    # la descripción no vacía más frecuente para ese CIE-10.
    # --------------------------------------------------------
    desc_validas = trabajo.loc[
        trabajo["DESCRIPCION"].notna() & (trabajo["DESCRIPCION"] != ""),
        ["CIE10", "DESCRIPCION"]
    ]

    if desc_validas.empty:
        mapa_descripcion = {}
    else:
        conteo_desc = (
            desc_validas.groupby(["CIE10", "DESCRIPCION"], observed=True)
            .size()
            .reset_index(name="N")
            .sort_values(["CIE10", "N"], ascending=[True, False], kind="stable")
        )
        mapa_descripcion = (
            conteo_desc.drop_duplicates("CIE10", keep="first")
            .set_index("CIE10")["DESCRIPCION"]
            .to_dict()
        )

    # HOMBRE / MUJER
    sexo = (
        trabajo.groupby(["CIE10", "SEXO_HM"], observed=True)
        .size().unstack(fill_value=0)
        .reindex(columns=["H", "M"], fill_value=0)
    )
    sexo.columns = ["HOMBRE", "MUJER"]
    sexo["TOTAL (H + M)"] = sexo["HOMBRE"] + sexo["MUJER"]

    # Etnia por CIE10
    etnia = (
        trabajo.groupby(["CIE10", "ETNIA"], observed=True)
        .size().unstack(fill_value=0)
        .reindex(columns=ETNIAS_ORDEN, fill_value=0)
    )

    conteo = sexo.join(etnia, how="left").fillna(0)
    conteo = conteo.sort_values(["TOTAL (H + M)"], ascending=False, kind="stable")

    # Denominador específico de CADA perfil.
    total_general = int(conteo["TOTAL (H + M)"].sum())

    top = conteo.head(numero_causas).copy()
    resto = conteo.iloc[numero_causas:]

    if not resto.empty:
        top.loc["OTRAS CAUSAS"] = resto.sum(numeric_only=True)

    # TOTAL (H+M) de cada CIE-10 / TOTAL GENERAL (H+M) de la tabla.
    top["%"] = (
        top["TOTAL (H + M)"].astype(float) / total_general * 100
        if total_general else 0.0
    )

    salida = top.reset_index().rename(columns={"CIE10": "CIE 10"})
    salida.insert(0, "Nro", range(1, len(salida) + 1))
    salida.insert(
        2,
        "DESCRIPCION CIE 10",
        salida["CIE 10"].map(mapa_descripcion).fillna("")
    )

    # OTRAS CAUSAS no lleva número ni descripción.
    mask_otras = salida["CIE 10"].eq("OTRAS CAUSAS")
    salida.loc[mask_otras, "Nro"] = pd.NA
    salida.loc[mask_otras, "DESCRIPCION CIE 10"] = ""

    # TOTAL GENERAL calculado sobre todos los CIE-10 incluidos en el perfil.
    fila_total = {
        "Nro": pd.NA,
        "CIE 10": "TOTAL GENERAL",
        "DESCRIPCION CIE 10": "",
        "HOMBRE": int(conteo["HOMBRE"].sum()),
        "MUJER": int(conteo["MUJER"].sum()),
        "TOTAL (H + M)": total_general,
        "%": 100.0 if total_general else 0.0,
    }
    for e in ETNIAS_ORDEN:
        fila_total[e] = int(conteo[e].sum())

    salida = pd.concat([salida, pd.DataFrame([fila_total])], ignore_index=True)

    orden_cols = [
        "Nro", "CIE 10", "DESCRIPCION CIE 10", "HOMBRE", "MUJER",
        "TOTAL (H + M)", "%"
    ] + ETNIAS_ORDEN
    salida = salida[orden_cols]

    filas_causas = salida["CIE 10"] != "TOTAL GENERAL"
    suma_pct = float(salida.loc[filas_causas, "%"].sum())

    formato = {
        "Nro": lambda x: "" if pd.isna(x) else f"{int(x)}",
        "HOMBRE": formato_numero,
        "MUJER": formato_numero,
        "TOTAL (H + M)": formato_numero,
        "%": lambda x: f"{float(x):.2f}%",
    }
    for e in ETNIAS_ORDEN:
        formato[e] = formato_numero

    sty = (
        salida.style
        .format(formato, na_rep="")
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("font-weight", "bold"), ("text-align", "center"),
                ("vertical-align", "middle"), ("white-space", "normal"),
                ("border-bottom", "1px solid #000")
            ]},
            {"selector": "tbody td", "props": [
                ("text-align", "center"), ("vertical-align", "middle")
            ]},
            {"selector": "tbody td:nth-child(3)", "props": [("text-align", "left")]},
            {"selector": "tbody tr:last-child td", "props": [
                ("font-weight", "bold"), ("border-top", "1px solid #000")
            ]},
        ])
    )

    st.dataframe(sty, use_container_width=True, hide_index=True)
    st.caption("Fuente: PRAS 2026")
    st.caption("Elaborado: Unidad Provincial de Estadística y Análisis de la Información del Sistema Nacional de Salud")
    st.caption(
        f"Verificación: Morbilidad + Primera + Definitivo | "
        f"denominador TOTAL GENERAL (H + M) = {formato_numero(total_general)} | "
        f"Suma de porcentajes de causas + OTRAS CAUSAS = {suma_pct:.2f}%"
    )


perfil_morbilidad(df_filtrado, "PERFIL DE MORBILIDAD GENERAL", "general")
perfil_morbilidad(df_filtrado, "PERFIL DE MORBILIDAD MATERNO", "materno")
perfil_morbilidad(df_filtrado, "PERFIL DE MORBILIDAD SALUD MENTAL", "mental")
perfil_morbilidad(df_filtrado, "PERFIL DE MORBILIDAD ODONTOLOGICO", "odontologico")


# ============================================================
# MÉTRICAS DE RENDIMIENTO
# ============================================================
tiempo_total = time.perf_counter() - t_total_inicio
ram_actual_mb = proceso.memory_info().rss / 1024**2
with st.expander("Rendimiento"):
    m1, m2, m3 = st.columns(3)
    m1.metric("Lectura + preparación", f"{tiempo_lectura:.2f} s")
    m2.metric("Tiempo total", f"{tiempo_total:.2f} s")
    m3.metric("RAM proceso", f"{ram_actual_mb:,.0f} MB")
    st.caption(
        f"RAM al inicio: {ram_inicio_mb:,.0f} MB | "
        f"DataFrame consolidado: {memoria_mb:,.0f} MB | "
        f"Motor Excel: {motor_excel_disponible()}"
    )

# ============================================================
# VERIFICACIONES
# ============================================================
with st.expander("Verificación de resultados"):
    st.write("**Registros después de filtros:**", formato_numero(len(df_filtrado)))
    st.write("**Total de consultas con CIE10:**", formato_numero(total_consultas))
    st.write("**Consultas clasificadas en Morbilidad/Prevención:**", formato_numero(total_tabla))
    st.write("**Consultas no clasificadas:**", formato_numero(diferencia))
    st.write("**Total de Atenciones (ATEMED_ID únicos con CIE10):**", formato_numero(total_atenciones_general))
    st.write("**RAM aproximada de la base consolidada:**", f"{memoria_mb:,.0f} MB")

with st.expander("Hojas procesadas"):
    for hoja in hojas_procesadas:
        st.write(f"• {hoja}")
