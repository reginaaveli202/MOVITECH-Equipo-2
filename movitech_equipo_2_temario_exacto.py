# -*- coding: utf-8 -*-
"""
MOVITECH | Caso práctico de Tecnología Informática
Aplicación Streamlit con temas oficiales, tablas editables y gráficas dinámicas.

Instalar:
    py -m pip install streamlit pandas plotly

Ejecutar:
    py -m streamlit run movitech_caso_practico.py
"""

import json
import sqlite3
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

DB = Path(__file__).with_name("movitech_caso_practico.db")

TEMARIO = {
    "II": {
        "nombre": "Sistemas de integración empresarial para la cadena de suministro",
        "temas": [
            ("2.1", "Sistemas de CRM"),
            ("2.2", "Sistemas ERP"),
            ("2.3", "Sistemas SRM"),
            ("2.3.1", "Proceso de SRM"),
            ("2.4", "Sistemas de SCM"),
        ],
    },
    "III": {
        "nombre": "Sistemas integrales para la planeación industrial",
        "temas": [
            ("3.2", "Herramientas de información para la planeación de la demanda"),
            ("3.3", "Herramientas de información para la planeación del Plan Maestro de Producción (MPS)"),
            ("3.4", "Sistema de información para la Planeación de Recursos de Distribución (DRP)"),
            ("3.5", "Sistema de información para la Planeación de Requerimientos de Materiales (MRP)"),
            ("3.6", "Sistema de información para la Planeación de Capacidad de Planta"),
        ],
    },
}

st.set_page_config(page_title="MOVITECH | Tecnología Informática", page_icon=None, layout="wide")

def init_db():
    with sqlite3.connect(DB) as con:
        con.execute("""CREATE TABLE IF NOT EXISTS ejemplos (
            clave TEXT PRIMARY KEY, datos TEXT NOT NULL)""")

def datos_iniciales(unidad, codigo, titulo):
    t = (codigo + " " + titulo).lower()

    if any(x in t for x in ["crm", "cliente", "filosofía", "tecnicas de crm"]):
        return pd.DataFrame({
            "Mes": ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio"],
            "Clientes potenciales": [120, 135, 142, 160, 175, 190],
            "Ventas concretadas": [36, 42, 45, 52, 61, 68],
            "Satisfacción (%)": [82, 84, 83, 87, 89, 92],
        })
    if any(x in t for x in ["srm", "proveedor"]):
        return pd.DataFrame({
            "Proveedor": ["Acero MX", "Baterías del Centro", "Ruedas Industriales", "ElectroPartes"],
            "Entregas a tiempo (%)": [94, 88, 97, 91],
            "Defectos recibidos": [12, 25, 7, 18],
            "Costo (miles $)": [85, 120, 48, 95],
        })
    if any(x in t for x in ["scm", "supply chain", "cadena", "distribución", "logística", "drp", "deployment"]):
        return pd.DataFrame({
            "Centro / Ruta": ["CDMX", "Toluca", "Puebla", "Querétaro"],
            "Demanda (unidades)": [180, 120, 145, 110],
            "Unidades enviadas": [175, 120, 140, 105],
            "Días de entrega": [2, 1, 3, 2],
        })
    if any(x in t for x in ["erp", "corporativ", "organización", "información en las organizaciones", "sistemas de información"]):
        return pd.DataFrame({
            "Área": ["Producción", "Ventas", "Compras", "Finanzas", "Almacén"],
            "Registros procesados": [420, 310, 185, 160, 275],
            "Errores detectados": [8, 5, 9, 2, 6],
            "Tiempo promedio (min)": [4.2, 2.8, 5.1, 3.4, 3.9],
        })
    if any(x in t for x in ["base de datos", "modelo entidad", "modelo relacional", "gestor", "sql"]):
        return pd.DataFrame({
            "Tabla del sistema": ["Clientes", "Productos", "Pedidos", "Proveedores", "Inventario"],
            "Registros": [245, 38, 680, 16, 38],
            "Campos principales": [5, 6, 7, 5, 6],
            "Registros incompletos": [3, 0, 8, 1, 2],
        })
    if any(x in t for x in ["pronóstico", "estimación", "tendencia", "demanda"]):
        return pd.DataFrame({
            "Mes": ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio"],
            "Demanda real": [95, 110, 108, 125, 140, 152],
            "Demanda pronosticada": [100, 105, 115, 120, 135, 145],
        })
    if any(x in t for x in ["mps", "maestro de producción", "planificación"]):
        return pd.DataFrame({
            "Semana": ["1", "2", "3", "4", "5", "6"],
            "Demanda": [90, 105, 115, 120, 135, 150],
            "Producción planeada": [100, 100, 120, 120, 140, 140],
            "Inventario proyectado": [30, 25, 30, 30, 35, 25],
        })
    if "mrp" in t or "materiales" in t or "bom" in t or "requerimientos" in t:
        return pd.DataFrame({
            "Componente": ["Motor", "Batería", "Chasis", "Rueda", "Controlador"],
            "Requerimiento bruto": [100, 100, 100, 400, 100],
            "Inventario disponible": [20, 15, 35, 100, 10],
            "Requerimiento neto": [80, 85, 65, 300, 90],
        })
    if any(x in t for x in ["capacidad", "planta"]):
        return pd.DataFrame({
            "Estación": ["Corte", "Soldadura", "Ensamble", "Pruebas", "Empaque"],
            "Capacidad (unid.)": [150, 130, 110, 125, 160],
            "Producción real": [140, 118, 105, 120, 150],
        })
    if any(x in t for x in ["inventario", "stock", "almacén", "tableros de control"]):
        return pd.DataFrame({
            "Producto": ["Motor", "Batería", "Chasis", "Ruedas", "Controlador"],
            "Existencias": [28, 18, 42, 160, 14],
            "Stock mínimo": [20, 25, 30, 80, 15],
            "Punto de reorden": [25, 30, 35, 100, 20],
        })
    if any(x in t for x in ["calidad", "defectos", "control y la toma"]):
        return pd.DataFrame({
            "Lote": ["L-01", "L-02", "L-03", "L-04", "L-05"],
            "Unidades inspeccionadas": [100, 100, 100, 100, 100],
            "Unidades defectuosas": [5, 8, 3, 6, 2],
            "Retrabajos": [3, 5, 1, 4, 1],
        })
    if any(x in t for x in ["cnc", "automatizados", "producción", "serie", "células flexibles"]):
        return pd.DataFrame({
            "Línea / Máquina": ["Línea A", "Línea B", "CNC 01", "CNC 02", "Célula flexible"],
            "Unidades producidas": [120, 110, 85, 92, 105],
            "Tiempo ciclo (min)": [6.2, 6.8, 8.1, 7.7, 6.5],
            "Paros (min)": [25, 32, 18, 22, 15],
        })
    if any(x in t for x in ["big data", "visualización", "analítica", "transformación digital", "tecnologías"]):
        return pd.DataFrame({
            "Mes": ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio"],
            "Datos procesados (GB)": [120, 145, 160, 190, 225, 260],
            "Eficiencia (%)": [78, 80, 82, 84, 87, 90],
            "Incidencias": [18, 16, 15, 12, 9, 7],
        })
    if any(x in t for x in ["sustentabilidad", "sostenibilidad"]):
        return pd.DataFrame({
            "Mes": ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio"],
            "Energía (kWh)": [5200, 5100, 4950, 4800, 4650, 4500],
            "Residuos (kg)": [420, 400, 390, 370, 350, 330],
            "Material reciclado (kg)": [110, 120, 135, 145, 160, 175],
        })

    return pd.DataFrame({
        "Área": ["Producción", "Almacén", "Compras", "Ventas"],
        "Indicador principal": [82, 74, 91, 88],
        "Meta (%)": [90, 85, 95, 90],
        "Incidencias": [8, 12, 4, 6],
    })

def cargar_datos(clave, base):
    with sqlite3.connect(DB) as con:
        fila = con.execute("SELECT datos FROM ejemplos WHERE clave=?", (clave,)).fetchone()
    if fila:
        try:
            return pd.DataFrame(json.loads(fila[0]))
        except Exception:
            pass
    return base.copy()

def guardar_datos(clave, df):
    with sqlite3.connect(DB) as con:
        con.execute(
            "INSERT OR REPLACE INTO ejemplos(clave, datos) VALUES(?, ?)",
            (clave, df.to_json(orient="records", force_ascii=False))
        )

def explicacion(unidad, codigo, titulo):
    return (
        f"MOVITECH fabrica scooters eléctricos. En este ejercicio, el área responsable "
        f"aplica **{titulo}** para organizar, controlar o analizar una parte de sus operaciones. "
        f"Los datos de ejemplo representan una situación simulada de la empresa y permiten "
        f"observar cómo la información apoya una decisión industrial."
    )

init_db()

st.title("MOVITECH | Caso práctico empresarial")
st.caption("Tecnología Informática | Ingeniería Industrial | Equipo 2")
st.info("Caso académico desarrollado por el Equipo 2. MOVITECH es una empresa ficticia; los datos son simulados y editables.")

# Navegación: únicamente las cinco unidades y sus temas oficiales.
unidad_sel = st.sidebar.selectbox(
    "UNIDAD",
    list(TEMARIO.keys()),
    format_func=lambda k: f"Unidad {k} · {TEMARIO[k]['nombre']}"
)
temas_unidad = TEMARIO[unidad_sel]["temas"]
opciones = {f"{c} — {n}": (c, n) for c, n in temas_unidad}
tema_sel = st.sidebar.selectbox("TEMA Y SUBTEMA", list(opciones.keys()))
codigo, titulo = opciones[tema_sel]

st.sidebar.divider()
st.sidebar.caption(f"Unidad {unidad_sel}")
st.sidebar.caption(f"{len(temas_unidad)} temas y subtemas disponibles")
st.sidebar.caption("Selecciona otro tema para cambiar el caso práctico.")

st.header(f"Unidad {unidad_sel}. {TEMARIO[unidad_sel]['nombre']}")
st.subheader(f"{codigo} · {titulo}")
st.markdown(explicacion(unidad_sel, codigo, titulo))

clave = f"{unidad_sel}|{codigo}"
base = datos_iniciales(unidad_sel, codigo, titulo)
df = cargar_datos(clave, base)

st.markdown("### Datos del caso práctico")
st.write("La tabla es editable. Los indicadores y las gráficas se actualizan al modificar los datos.")

editado = st.data_editor(
    df,
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    key=f"editor_{clave}",
)
col_guardar, col_descarga = st.columns([1, 1])
with col_guardar:
    if st.button("Guardar cambios", type="primary", use_container_width=True):
        guardar_datos(clave, editado)
        st.success("Los datos de este tema quedaron guardados en tu computadora.")
with col_descarga:
    st.download_button(
        "Descargar tabla CSV",
        data=editado.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"MOVITECH_{codigo.replace('.', '_')}.csv",
        mime="text/csv",
        use_container_width=True,
    )

st.markdown("### Indicadores")
numericas = editado.select_dtypes(include="number").columns.tolist()
if numericas:
    cols = st.columns(min(4, len(numericas)))
    for i, col in enumerate(numericas[:4]):
        valor = editado[col].mean()
        cols[i].metric(f"Promedio: {col}", f"{valor:,.1f}")
else:
    st.caption("Agrega columnas numéricas para visualizar indicadores.")

st.markdown("### Análisis gráfico")
if len(editado.columns) >= 2 and not editado.empty:
    categoricas = [c for c in editado.columns if c not in numericas]
    x_default = categoricas[0] if categoricas else editado.columns[0]
    y_default = numericas[0] if numericas else editado.columns[-1]

    g1, g2 = st.columns(2)
    with g1:
        x_col = st.selectbox("Eje horizontal", editado.columns, index=list(editado.columns).index(x_default))
    with g2:
        y_col = st.selectbox("Indicador para comparar", editado.columns, index=list(editado.columns).index(y_default))

    if y_col in numericas:
        tipo = st.radio("Tipo de gráfica", ["Barras", "Línea", "Dispersión"], horizontal=True)
        if tipo == "Barras":
            fig = px.bar(editado, x=x_col, y=y_col, color=x_col, title=f"{y_col} por {x_col}")
        elif tipo == "Línea":
            fig = px.line(editado, x=x_col, y=y_col, markers=True, title=f"Evolución de {y_col}")
        else:
            fig = px.scatter(editado, x=x_col, y=y_col, title=f"Relación entre {x_col} y {y_col}")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Selecciona una columna numérica como indicador para generar la gráfica.")
else:
    st.info("La tabla necesita al menos dos columnas y un registro para generar gráficas.")

st.markdown("### Interpretación de resultados")
if numericas and not editado.empty:
    principal = numericas[0]
    max_idx = editado[principal].idxmax()
    min_idx = editado[principal].idxmin()
    etiqueta = editado.columns[0]
    st.write(
        f"En los datos actuales, el valor más alto de **{principal}** es "
        f"**{editado.loc[max_idx, principal]}** en **{editado.loc[max_idx, etiqueta]}**. "
        f"El valor más bajo es **{editado.loc[min_idx, principal]}** en "
        f"**{editado.loc[min_idx, etiqueta]}**. La empresa puede utilizar esta comparación "
        f"para identificar diferencias y evaluar decisiones relacionadas con el tema."
    )
else:
    st.write("Agrega datos numéricos para obtener una interpretación automática.")

st.caption("Nota: los resultados corresponden a datos simulados con fines académicos. El Equipo 2 puede modificar los registros para analizar distintos escenarios.")
