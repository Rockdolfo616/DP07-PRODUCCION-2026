"""
app.py
======
Aplicación pública de Streamlit.

Lee exclusivamente "datos_publicos.db", generado previamente por
GENERAR_DATOS_PUBLICOS.py en el entorno institucional.

NO carga Excel, NO contiene PCTE_IDE, NO contiene ATEMED_ID y NO necesita
acceso a la base confidencial.
"""

import io
import json
import sqlite3
from pathlib import Path

import pandas as pd
import streamlit as st

DB = Path(__file__).with_name("datos_publicos.db")

FILTROS = [
    ("Año", "anio"),
    ("Provincia", "provincia"),
    ("Cantón", "canton"),
    ("Parroquia", "parroquia"),
    ("Distrito", "distrito"),
    ("Establecimiento", "establecimiento"),
    ("Nivel de Atención", "nivel"),
]

ETNIAS_ORDEN = [
    "Afrodescendiente", "Blanco", "Indígena", "Mestizos", "Montubio",
    "Mulato", "Negro", "No Aplica", "No sabe/ no responde", "Otro"
]

st.set_page_config(
    page_title="DP El Oro - Consultas y Atenciones",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
.block-container {padding-top: 2rem; padding-bottom: 3rem;}
div[data-testid="stMetric"] {
    border: 1px solid rgba(49, 51, 63, 0.18);
    padding: 1rem;
    border-radius: 0.6rem;
}
</style>
""", unsafe_allow_html=True)

st.title("TOTAL DE CONSULTAS Y ATENCIONES")
st.markdown("**DIRECCIÓN PROVINCIAL DE EL ORO**")
st.caption(
    "Producción de establecimientos de salud — resultados estadísticos agregados. "
    "Fuente: PRAS 2026."
)

def formato_numero(x):
    try:
        return f"{int(x):,}".replace(",", ".")
    except (ValueError, TypeError):
        return x

def estilo_tabla(df):
    return (
        df.style.format(formato_numero)
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

@st.cache_resource
def conexion():
    if not DB.exists():
        return None
    return sqlite3.connect(f"file:{DB}?mode=ro", uri=True, check_same_thread=False)

def condicion(selecciones, hasta=None):
    partes, params = [], []
    keys = [k for _, k in FILTROS]
    if hasta is not None:
        keys = keys[:hasta]
    for k in keys:
        v = selecciones.get(k)
        if v == "Todos" or v is None:
            partes.append(f"{k} IS NULL")
        else:
            partes.append(f"{k} = ?")
            params.append(str(v))
    return " AND ".join(partes) if partes else "1=1", params

def opciones_siguiente(conn, selecciones, indice):
    _, key = FILTROS[indice]
    where, params = condicion(selecciones, hasta=indice)
    # Los campos posteriores pueden tener cualquier valor; para obtener las
    # opciones válidas basta buscar estados con el prefijo actual.
    sql = f"""
        SELECT DISTINCT {key}
        FROM estados
        WHERE {where} AND {key} IS NOT NULL
        ORDER BY {key}
    """
    vals = [r[0] for r in conn.execute(sql, params).fetchall()]
    if key == "anio":
        vals = sorted(vals, key=lambda x: int(x))
    return vals

def cargar_resultado(conn, selecciones):
    partes, params = [], []
    for _, k in FILTROS:
        v = selecciones.get(k)
        if v == "Todos" or v is None:
            partes.append(f"{k} IS NULL")
        else:
            partes.append(f"{k} = ?")
            params.append(str(v))
    row = conn.execute(
        "SELECT resultado_json FROM estados WHERE " + " AND ".join(partes) + " LIMIT 1",
        params
    ).fetchone()
    return json.loads(row[0]) if row else None

def payload_a_df(payload, index_col=None):
    if not payload:
        return pd.DataFrame()
    df = pd.DataFrame(payload["records"], columns=payload["columns"])
    if index_col and index_col in df.columns:
        df = df.set_index(index_col)
    return df

def tabla_especialidad(payload):
    df = payload_a_df(payload)
    if df.empty:
        return df
    # El generador aplana MultiIndex con "|||".
    idx = "Especialidad"
    if idx in df.columns:
        df = df.set_index(idx)
    columnas = []
    for c in df.columns:
        p = c.split("|||")
        columnas.append(tuple(p) if len(p) > 1 else (p[0], ""))
    df.columns = pd.MultiIndex.from_tuples(columnas)
    return df

def dataframe_a_excel_bytes(df, nombre_hoja="Tabla"):
    """Genera un XLSX en memoria únicamente a partir de una tabla agregada."""
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=nombre_hoja[:31])
    salida.seek(0)
    return salida.getvalue()

def libro_completo_excel(tablas):
    """Genera un libro XLSX con todas las tablas agregadas visibles."""
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        for nombre, df in tablas.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=nombre[:31])
    salida.seek(0)
    return salida.getvalue()

def boton_descarga_excel(df, nombre_archivo, nombre_hoja, key):
    if df is not None and not df.empty:
        st.download_button(
            "⬇️ Descargar tabla en Excel",
            data=dataframe_a_excel_bytes(df, nombre_hoja),
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=key,
        )

def mostrar_perfil(payload, nombre, numero_causas):
    st.subheader(nombre)
    base = payload_a_df(payload)
    if base.empty:
        st.info("No existen registros para los filtros y criterios seleccionados.")
        return

    if "CIE 10" in base.columns:
        base = base.set_index("CIE 10")

    numericas = ["HOMBRE", "MUJER", "TOTAL (H + M)"] + ETNIAS_ORDEN
    for c in numericas:
        if c in base.columns:
            base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

    total_general = int(base["TOTAL (H + M)"].sum())
    top = base.head(numero_causas).copy()
    resto = base.iloc[numero_causas:]

    if not resto.empty:
        fila = {c: int(resto[c].sum()) for c in numericas}
        fila["DESCRIPCION CIE 10"] = ""
        top.loc["OTRAS CAUSAS"] = fila

    top["%"] = (
        top["TOTAL (H + M)"].astype(float) / total_general * 100
        if total_general else 0.0
    )

    salida = top.reset_index()
    salida.insert(0, "Nro", range(1, len(salida) + 1))
    mask_otras = salida["CIE 10"].eq("OTRAS CAUSAS")
    salida.loc[mask_otras, "Nro"] = pd.NA
    salida.loc[mask_otras, "DESCRIPCION CIE 10"] = ""

    fila_total = {
        "Nro": pd.NA, "CIE 10": "TOTAL GENERAL", "DESCRIPCION CIE 10": "",
        "HOMBRE": int(base["HOMBRE"].sum()),
        "MUJER": int(base["MUJER"].sum()),
        "TOTAL (H + M)": total_general,
        "%": 100.0 if total_general else 0.0,
    }
    for e in ETNIAS_ORDEN:
        fila_total[e] = int(base[e].sum())

    salida = pd.concat([salida, pd.DataFrame([fila_total])], ignore_index=True)
    orden = [
        "Nro", "CIE 10", "DESCRIPCION CIE 10",
        "HOMBRE", "MUJER", "TOTAL (H + M)", "%"
    ] + ETNIAS_ORDEN
    salida = salida[orden]

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
        salida.style.format(formato, na_rep="")
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

conn = conexion()
if conn is None:
    st.error(
        "No se encontró datos_publicos.db. Ejecute GENERAR_DATOS_PUBLICOS.py "
        "en el equipo institucional y publique únicamente el archivo agregado resultante."
    )
    st.stop()

# ---------------------------------------------------------------------
# FILTROS EN CASCADA
# ---------------------------------------------------------------------
st.sidebar.title("FILTRO")
selecciones = {}
for i, (etiqueta, key) in enumerate(FILTROS):
    vals = opciones_siguiente(conn, selecciones, i)
    valor = st.sidebar.selectbox(etiqueta, ["Todos"] + vals, key=f"f_{key}")
    selecciones[key] = valor

resultado = cargar_resultado(conn, selecciones)
if resultado is None:
    st.warning("No existen resultados precalculados para los filtros seleccionados.")
    st.stop()

st.sidebar.caption(f"Registros seleccionados: {formato_numero(resultado['registros'])}")
with st.sidebar.expander("Filtros aplicados"):
    for etiqueta, key in FILTROS:
        st.write(f"**{etiqueta}:** {selecciones[key]}")

# ---------------------------------------------------------------------
# RESUMEN GENERAL
# ---------------------------------------------------------------------
st.subheader("RESUMEN GENERAL")
c1, c2 = st.columns(2)
c1.metric("Total de Consultas", formato_numero(resultado["total_consultas"]))
c2.metric("Total de Atenciones", formato_numero(resultado["total_atenciones"]))

activos = [
    f"{etiqueta}: {selecciones[key]}"
    for etiqueta, key in FILTROS if selecciones[key] != "Todos"
]
if activos:
    st.caption(" | ".join(activos))

# ---------------------------------------------------------------------
# TABLAS
# ---------------------------------------------------------------------
st.markdown("---")
st.subheader("Total de consultas y atenciones por Especialidad")
tabla_esp = tabla_especialidad(resultado["especialidad"])
st.dataframe(estilo_tabla(tabla_esp), use_container_width=True)
boton_descarga_excel(tabla_esp, "consultas_atenciones_especialidad.xlsx", "Especialidad", "dl_esp")
st.caption("Fuente: PRAS 2026")

st.markdown("---")
st.subheader("Total de Atenciones por Nivel de Atención y Mes")
tabla_nivel = payload_a_df(resultado["nivel_mes"], "Nivel de Atención")
st.dataframe(estilo_tabla(tabla_nivel).set_properties(subset=["Total"], **{"font-weight": "bold"}), use_container_width=True)
boton_descarga_excel(tabla_nivel, "atenciones_nivel_mes.xlsx", "Nivel por mes", "dl_nivel")
st.caption("Fuente: PRAS 2026")

st.markdown("---")
st.subheader("Total de Atenciones por Nacionalidad y Mes")
tabla_nac = payload_a_df(resultado["nacionalidad_mes"], "Nacionalidad")
st.dataframe(estilo_tabla(tabla_nac).set_properties(subset=["Total"], **{"font-weight": "bold"}), use_container_width=True)
boton_descarga_excel(tabla_nac, "atenciones_nacionalidad_mes.xlsx", "Nacionalidad por mes", "dl_nac")
st.caption("Fuente: PRAS 2026")

st.markdown("---")
st.subheader("Total de Atenciones por Cantón y Mes")
tabla_canton = payload_a_df(resultado["canton_mes"], "Cantón")
st.dataframe(estilo_tabla(tabla_canton).set_properties(subset=["Total"], **{"font-weight": "bold"}), use_container_width=True)
boton_descarga_excel(tabla_canton, "atenciones_canton_mes.xlsx", "Cantón por mes", "dl_canton")
st.caption("Fuente: PRAS 2026")

st.markdown("---")
st.subheader("Numero de Embarazadas")
emb = payload_a_df(resultado["embarazadas"], "Cantón")
if emb.empty:
    st.info("No existen registros de embarazadas para los filtros seleccionados.")
else:
    st.dataframe(
        estilo_tabla(emb).set_properties(subset=["Total"], **{"font-weight": "bold"}),
        use_container_width=True
    )
    boton_descarga_excel(emb, "numero_embarazadas.xlsx", "Embarazadas", "dl_emb")
    st.caption("Fuente: PRAS 2026")
    st.caption(
        "Criterio: pacientes únicas por PCTE_IDE; cédulas de 9 dígitos se homologan "
        "anteponiendo 0; edad = máxima PCTE_ANIOS; establecimiento = el de mayor "
        "número de registros por paciente. Los identificadores no forman parte "
        "de la base publicada."
    )

st.markdown("---")
st.subheader("DESCARGA DE RESULTADOS AGREGADOS")
tablas_exportar = {
    "Especialidad": tabla_esp,
    "Nivel_mes": tabla_nivel,
    "Nacionalidad_mes": tabla_nac,
    "Canton_mes": tabla_canton,
}
if not emb.empty:
    tablas_exportar["Embarazadas"] = emb

st.download_button(
    "📥 Descargar tablas principales en un solo Excel",
    data=libro_completo_excel(tablas_exportar),
    file_name="DP07_resultados_agregados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="dl_todo",
)
st.caption("El archivo descargado contiene únicamente resultados estadísticos agregados.")

# ---------------------------------------------------------------------
# PERFILES DE MORBILIDAD
# ---------------------------------------------------------------------
st.markdown("---")
st.header("PERFILES DE MORBILIDAD")
numero_causas = st.selectbox(
    "Filtro", [10, 20, 30], index=0,
    format_func=lambda x: f"{x} Primeras Causas",
    key="numero_causas_morbilidad"
)
mostrar_perfil(resultado["perfil_general"], "PERFIL DE MORBILIDAD GENERAL", numero_causas)
mostrar_perfil(resultado["perfil_materno"], "PERFIL DE MORBILIDAD MATERNO", numero_causas)
mostrar_perfil(resultado["perfil_mental"], "PERFIL DE MORBILIDAD SALUD MENTAL", numero_causas)
mostrar_perfil(resultado["perfil_odontologico"], "PERFIL DE MORBILIDAD ODONTOLOGICO", numero_causas)

# ---------------------------------------------------------------------
# VERIFICACIONES
# ---------------------------------------------------------------------
with st.expander("Verificación de resultados"):
    st.write("**Registros después de filtros:**", formato_numero(resultado["registros"]))
    st.write("**Total de consultas con CIE10:**", formato_numero(resultado["total_consultas"]))
    st.write("**Consultas clasificadas en Morbilidad/Prevención:**", formato_numero(resultado["consultas_clasificadas"]))
    st.write("**Consultas no clasificadas:**", formato_numero(resultado["consultas_no_clasificadas"]))
    st.write("**Total de Atenciones (ATEMED_ID únicos con CIE10):**", formato_numero(resultado["total_atenciones"]))
    st.caption("Los ATEMED_ID se utilizaron únicamente durante el procesamiento local; no están en datos_publicos.db.")

st.markdown("---")
st.caption(
    "Fuente: PRAS 2026 | Elaborado: Unidad Provincial de Estadística y Análisis "
    "de la Información del Sistema Nacional de Salud."
)
