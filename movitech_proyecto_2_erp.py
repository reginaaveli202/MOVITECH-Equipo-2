# -*- coding: utf-8 -*-
"""
MOVITECH | ERP INTEGRAL - CASO PRÁCTICO
Basado en la estructura del archivo de simulación de capacidad de planta:
supuestos, demanda, CRM, SRM, inventarios/MRP, MPS, capacidad, DRP/SCM,
costos, riesgos y dashboard.

Instalar:
    py -m pip install -r requirements_movitech.txt

Ejecutar:
    py -m streamlit run movitech_proyecto_2_erp.py

Los datos son hipotéticos y editables. Se guardan en movitech_erp_local.db.
"""

from pathlib import Path
import io
from datetime import datetime
import sqlite3
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference

# ---------------------------- CONFIGURACIÓN ----------------------------
BASE = Path(__file__).resolve().parent
DB = BASE / "movitech_erp_local.db"
st.set_page_config(page_title="MOVITECH | ERP integral", page_icon="⚙️", layout="wide")

st.markdown("""
<style>
.stApp {background:#F4F8F4}
[data-testid="stSidebar"] {background:#E5F0E7}
h1,h2,h3 {color:#52796F}
div[data-testid="stMetric"] {background:white;border:1px solid #C9DDCE;padding:12px;border-radius:12px}
.hero {background:linear-gradient(110deg,#52796F,#84A98C);padding:22px;border-radius:16px;color:white}
.hero h1,.hero p {color:white!important}
.tip {background:#E4F0E6;border-left:5px solid #84A98C;padding:12px 15px;border-radius:8px}
</style>
""", unsafe_allow_html=True)

# ---------------------------- DATOS BASE ----------------------------
SUPUESTOS = pd.DataFrame([
    ["Demanda total", 200, "scooters", "Pedidos confirmados que se deben atender"],
    ["Días laborables", 20, "días", "Horizonte del plan"],
    ["Horas por turno", 8, "h/día", "Duración del turno"],
    ["Turnos por día", 1, "turnos", "Turnos programados"],
    ["Eficiencia planeada", 0.85, "factor", "Disponibilidad/ritmo/calidad estimados"],
    ["Horas extra por día", 0, "h/día", "Horas extra disponibles por estación"],
    ["Costo hora operador", 85, "MXN/h", "Costo hipotético de mano de obra"],
    ["Costo energía", 2.8, "MXN/kWh", "Tarifa simulada"],
    ["Consumo por scooter", 1.8, "kWh/u", "Consumo promedio en ensamble y pruebas"],
    ["Costo material por scooter", 7200, "MXN/u", "Costo agregado hipotético de componentes"],
    ["Costo indirecto diario", 1800, "MXN/día", "Supervisión, servicios y gastos indirectos"],
    ["Inventario inicial de scooters", 10, "unidades", "Producto terminado disponible al inicio"],
], columns=["Variable", "Valor", "Unidad", "Descripción"])

PROBLEMA = pd.DataFrame([
    ["P-01","Información fragmentada","Ventas, compras, almacén y producción trabajan con registros separados.","Datos duplicados o desactualizados.","ERP con base de datos central.","% de registros completos"],
    ["P-02","Faltantes de componentes","El inventario no se compara oportunamente con la lista de materiales y el plan.","Compras urgentes y riesgo de detener producción.","MRP conectado con inventario, MPS y compras.","Componentes faltantes"],
    ["P-03","Proveedores sin evaluación integral","No se concentra el historial de entregas, defectos, costos y tiempos.","Retrasos o recepción de componentes defectuosos.","SRM para evaluar proveedores y dar seguimiento a órdenes.","% entregas a tiempo"],
    ["P-04","Producción sin visibilidad de capacidad","La programación no siempre compara demanda, tiempos y recursos disponibles.","Cuello de botella y retrasos de fabricación.","MPS y planeación de capacidad conectados.","Utilización del cuello de botella"],
    ["P-05","Distribución descoordinada","No se consideran simultáneamente pedidos, inventario por zona y mercancía en tránsito.","Pedidos incompletos o entregas tardías.","DRP y SCM para planear reabastecimiento y entregas.","% pedidos completos"],
    ["P-06","Costos dispersos","Los costos de materiales, mano de obra, energía e indirectos se calculan por separado.","Dificultad para estimar costo por scooter.","ERP para consolidar costos operativos.","Costo estimado por scooter"],
], columns=["ID","Problema","Causa / evidencia","Consecuencia","Solución ERP propuesta","Indicador"])

CRM = pd.DataFrame([
    ["V-001","Cliente Centro","CDMX",70,"Confirmado","Prioridad alta"],
    ["V-002","Movilidad Toluca","Toluca",45,"Confirmado","Prioridad media"],
    ["V-003","Distribuidora Puebla","Puebla",50,"Confirmado","Prioridad alta"],
    ["V-004","Scooters Bajío","Querétaro",35,"Confirmado","Prioridad media"],
], columns=["Pedido","Cliente","Zona","Unidades","Estado","Observaciones"])

SRM = pd.DataFrame([
    ["PROV-01","Motores del Centro","Motor",20,92,8,1800,5],
    ["PROV-02","Energía y Baterías MX","Batería",15,88,18,2600,7],
    ["PROV-03","Estructuras Industriales","Chasis",35,96,5,1200,4],
    ["PROV-04","Ruedas y Rodamientos","Rueda",50,94,7,350,3],
    ["PROV-05","Electrónica Aplicada","Controlador",10,90,12,950,6],
], columns=["ID proveedor","Proveedor","Componente","Stock disponible","Entregas a tiempo (%)","Defectos recibidos","Costo unitario (MXN)","Lead time (días)"])

BOM = pd.DataFrame([
    ["Motor",1,20,0,"Motores del Centro"],
    ["Batería",1,15,10,"Energía y Baterías MX"],
    ["Chasis",1,35,0,"Estructuras Industriales"],
    ["Rueda",2,50,20,"Ruedas y Rodamientos"],
    ["Controlador",1,10,0,"Electrónica Aplicada"],
], columns=["Componente","Cantidad por scooter","Inventario disponible","Recepciones programadas","Proveedor"])

DEMANDA = pd.DataFrame([
    ["Semana 1",45,45],["Semana 2",50,50],["Semana 3",50,50],["Semana 4",55,55]
], columns=["Periodo","Demanda pronosticada","Pedidos confirmados"])

MPS = pd.DataFrame([
    ["Semana 1",45,45,10],["Semana 2",50,50,10],["Semana 3",50,50,10],["Semana 4",55,55,10]
], columns=["Periodo","Demanda","Producción planeada","Inventario inicial"])

RUTA = pd.DataFrame([
    [1,"Preparación e inspección del chasis",25,1,1,"Banco de trabajo","Inspección visual y dimensional"],
    [2,"Montaje de motor y batería",35,1,2,"Banco y herramienta de montaje","Revisión de fijaciones"],
    [3,"Cableado y conexiones eléctricas",20,1,1,"Herramienta aislada y multímetro","Continuidad eléctrica"],
    [4,"Montaje de ruedas y frenos",20,1,1,"Herramienta manual","Alineación y ajuste"],
    [5,"Prueba funcional y de seguridad",15,1,1,"Equipo de prueba","Checklist funcional"],
    [6,"Inspección final y empaque",10,1,1,"Mesa de empaque","Etiquetado y revisión final"],
], columns=["Orden","Operación","Tiempo estándar (min/u)","Estaciones","Operadores por estación","Equipo","Control de calidad"])

DRP = pd.DataFrame([
    ["CDMX",70,10,5,2],["Toluca",45,5,5,1],["Puebla",50,8,0,3],["Querétaro",35,4,5,2]
], columns=["Centro / zona","Demanda","Inventario disponible","En tránsito","Lead time (días)"])

RIESGOS = pd.DataFrame([
    ["Retraso en baterías","Media","Alto","Confirmar fechas con proveedor y revisar stock de seguridad","Compras / SRM","Entregas a tiempo (%)"],
    ["Falla en prueba eléctrica","Media","Alto","Aplicar checklist y registrar resultados por unidad","Calidad / Producción","Aprobación a la primera (%)"],
    ["Ausencia de operador clave","Media","Medio","Capacitación cruzada y suplente","Producción / RH","Ausentismo diario"],
    ["Tiempo real mayor al estándar","Media","Alto","Tomar tiempos y balancear estaciones","Ingeniería de procesos","Tiempo real vs. estándar"],
    ["Daño de componentes","Baja","Alto","Manejo y almacenamiento controlados","Almacén","Merma de componentes"],
    ["Paro de equipo","Baja","Medio","Inspección y mantenimiento preventivo","Mantenimiento","Minutos de paro"],
], columns=["Riesgo","Probabilidad","Impacto","Acción preventiva","Responsable","Indicador de seguimiento"])

# ---------------------------- BASE LOCAL ----------------------------
def init_db():
    with sqlite3.connect(DB) as con:
        con.execute("CREATE TABLE IF NOT EXISTS tablas (nombre TEXT PRIMARY KEY, datos TEXT NOT NULL, actualizado TEXT NOT NULL)")
        con.execute("CREATE TABLE IF NOT EXISTS textos (nombre TEXT PRIMARY KEY, valor TEXT NOT NULL)")
init_db()

def load_df(name, default):
    with sqlite3.connect(DB) as con:
        row = con.execute("SELECT datos FROM tablas WHERE nombre=?",(name,)).fetchone()
    if row is None:
        return default.copy()
    try:
        return pd.read_json(row[0], orient="records")
    except Exception:
        return default.copy()

def save_df(name, df):
    with sqlite3.connect(DB) as con:
        con.execute("INSERT OR REPLACE INTO tablas VALUES (?,?,?)",
                    (name,df.to_json(orient="records",force_ascii=False),datetime.now().isoformat(timespec="seconds")))

def load_text(name, default):
    with sqlite3.connect(DB) as con:
        row=con.execute("SELECT valor FROM textos WHERE nombre=?",(name,)).fetchone()
    return row[0] if row else default

def save_text(name, value):
    with sqlite3.connect(DB) as con:
        con.execute("INSERT OR REPLACE INTO textos VALUES (?,?)",(name,value))

def num(df,col):
    return pd.to_numeric(df[col],errors="coerce").fillna(0) if col in df else pd.Series([0]*len(df),index=df.index,dtype=float)

def edit_table(name, default, key, help_text=""):
    st.caption(help_text or "Edita las celdas, agrega o elimina filas y pulsa Guardar para conservar los cambios.")
    df=load_df(name,default)
    edited=st.data_editor(df,use_container_width=True,hide_index=True,num_rows="dynamic",key=key)
    c1,c2=st.columns([1,1])
    with c1:
        if st.button("Guardar cambios",key=f"save_{key}",type="primary",use_container_width=True):
            save_df(name,edited)
            st.success("Cambios guardados en la base local.")
            st.rerun()
    with c2:
        if st.button("Restaurar ejemplo",key=f"reset_{key}",use_container_width=True):
            save_df(name,default)
            st.success("Datos de ejemplo restaurados.")
            st.rerun()
    st.download_button("Descargar CSV",edited.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"MOVITECH_{key}.csv",mime="text/csv",key=f"csv_{key}")
    return edited

def get_all():
    names={"Supuestos":SUPUESTOS,"Diagnóstico":PROBLEMA,"CRM":CRM,"SRM":SRM,"BOM_MRP":BOM,
           "Demanda":DEMANDA,"MPS":MPS,"Ruta":RUTA,"DRP":DRP,"Riesgos":RIESGOS}
    return {k:load_df(k,v) for k,v in names.items()}

def val(sup, label, default):
    row=sup.loc[sup["Variable"]==label,"Valor"]
    if row.empty: return default
    try: return float(row.iloc[0])
    except Exception: return default

def demanda_integrada(sup):
    # Los pedidos CRM son la fuente principal de demanda; Supuestos sirve como respaldo.
    pedidos = load_df("CRM", CRM)
    if not pedidos.empty and "Unidades" in pedidos.columns:
        total = pd.to_numeric(pedidos["Unidades"], errors="coerce").fillna(0).sum()
        return float(total)
    return val(sup, "Demanda total", 200)

def calc_capacity(sup,ruta):
    days=val(sup,"Días laborables",20)
    hours=val(sup,"Horas por turno",8)
    shifts=val(sup,"Turnos por día",1)
    eff=val(sup,"Eficiencia planeada",.85)
    overtime=val(sup,"Horas extra por día",0)
    x=ruta.copy()
    mins=(days*hours*shifts*60 + days*overtime*60)*eff
    x["Minutos efectivos por estación"]=mins
    x["Capacidad por proceso (u)"]=(
        mins*num(x,"Estaciones")/num(x,"Tiempo estándar (min/u)").replace(0,pd.NA)
    ).fillna(0).round(1)
    x["Carga requerida (min)"]=num(x,"Tiempo estándar (min/u)")*demanda_integrada(sup)
    denom=mins*num(x,"Estaciones")
    x["Utilización estimada (%)"]=(x["Carga requerida (min)"]/denom.replace(0,pd.NA)*100).fillna(0).round(1)
    return x

def calc_mrp(sup,bom):
    x=bom.copy()
    gross=demanda_integrada(sup)*num(x,"Cantidad por scooter")
    x["Requerimiento bruto"]=gross
    x["Requerimiento neto"]=(gross-num(x,"Inventario disponible")-num(x,"Recepciones programadas")).clip(lower=0)
    return x

def calc_drp(drp):
    x=drp.copy()
    x["Reabastecimiento sugerido"]=(num(x,"Demanda")-num(x,"Inventario disponible")-num(x,"En tránsito")).clip(lower=0)
    x["Cobertura actual (%)"]=((num(x,"Inventario disponible")+num(x,"En tránsito"))/num(x,"Demanda").replace(0,pd.NA)*100).fillna(0).round(1)
    return x

def calc_mps(mps):
    x=mps.copy()
    x["Inventario final proyectado"]=num(x,"Inventario inicial")+num(x,"Producción planeada")-num(x,"Demanda")
    x["Producción acumulada"]=num(x,"Producción planeada").cumsum()
    return x

def calc_costs(sup,ruta):
    demand=demanda_integrada(sup)
    days=val(sup,"Días laborables",20)
    hours=val(sup,"Horas por turno",8)
    shifts=val(sup,"Turnos por día",1)
    labor=val(sup,"Costo hora operador",85)
    staff=(num(ruta,"Estaciones")*num(ruta,"Operadores por estación")).sum()+2
    material=demand*val(sup,"Costo material por scooter",7200)
    labor_cost=staff*hours*shifts*days*labor
    energy=demand*val(sup,"Consumo por scooter",1.8)*val(sup,"Costo energía",2.8)
    overhead=days*val(sup,"Costo indirecto diario",1800)
    return pd.DataFrame([
        ["Materiales",demand,val(sup,"Costo material por scooter",7200),material],
        ["Mano de obra",staff*hours*shifts*days,labor,labor_cost],
        ["Energía",demand*val(sup,"Consumo por scooter",1.8),val(sup,"Costo energía",2.8),energy],
        ["Costos indirectos",days,val(sup,"Costo indirecto diario",1800),overhead],
    ],columns=["Concepto","Base de cálculo","Costo unitario (MXN)","Costo estimado (MXN)"])


# ---------------------------- EXPORTACIÓN EXCEL ----------------------------
def build_excel_bytes(selected_sheets=None):
    """Genera un libro integral o por práctica con los datos actuales guardados en SQLite."""
    data = get_all()
    sup = data["Supuestos"]
    ruta = data["Ruta"]
    demand = demanda_integrada(sup)
    mrp = calc_mrp(sup, data["BOM_MRP"])
    capacity = calc_capacity(sup, ruta)
    drp_calc = calc_drp(data["DRP"])
    mps_calc = calc_mps(data["MPS"])
    costs = calc_costs(sup, ruta)
    costs.loc[len(costs)] = ["TOTAL", None, None, num(costs, "Costo estimado (MXN)").sum()]
    diagnosis = data["Diagnóstico"]
    solution_default = pd.DataFrame([
        ["CRM","Clientes y pedidos","Pedidos confirmados por zona","Concentrar pedidos y fechas requeridas","Demanda para planeación"],
        ["ERP","Información de áreas","Ventas, compras, almacén, producción y costos","Unificar catálogos y movimientos","Datos centrales consistentes"],
        ["SRM","Proveedores","Precios, tiempos, calidad y entregas","Evaluar proveedores y controlar abastecimiento","Órdenes y fechas de recepción"],
        ["SCM","Cadena de suministro","Pedidos, materiales, fabricación y envíos","Coordinar el flujo completo","Visibilidad de suministro y entrega"],
        ["Demanda","Pronóstico","Historial y pedidos confirmados","Estimar demanda por periodo","Pronóstico para MPS"],
        ["MPS","Plan maestro","Demanda, inventario y calendario","Definir unidades por periodo","Programa de producción"],
        ["MRP","Materiales","MPS, lista de materiales y existencias","Calcular requerimiento bruto y neto","Compras de componentes"],
        ["Capacidad","Capacidad de planta","Tiempos, estaciones, turnos y eficiencia","Calcular capacidad y cuello de botella","Validación del programa"],
        ["DRP","Distribución","Demanda por zona, inventario y tránsito","Calcular reabastecimiento","Plan de envío por centro"],
    ], columns=["Módulo","Proceso","Datos de entrada","Actividad del sistema","Resultado compartido"])
    solution = load_df("Solución ERP", solution_default)
    sheets = {
        "Resumen": pd.DataFrame([
            ["Empresa","MOVITECH (caso académico ficticio)"],
            ["Demanda integrada (scooters)", demand],
            ["Capacidad de línea (scooters)", capacity["Capacidad por proceso (u)"].min() if not capacity.empty else 0],
            ["Margen de capacidad", (capacity["Capacidad por proceso (u)"].min()-demand) if not capacity.empty else 0],
            ["Componentes por abastecer", num(mrp,"Requerimiento neto").sum()],
            ["Reabastecimiento DRP", num(drp_calc,"Reabastecimiento sugerido").sum()],
            ["Costo total estimado (MXN)", num(costs,"Costo estimado (MXN)").sum()],
            ["Costo estimado por scooter (MXN)", (num(costs,"Costo estimado (MXN)").sum()/demand if demand else 0)],
            ["Nota","Datos hipotéticos y editables; validar antes de tomar decisiones reales."]
        ], columns=["Indicador","Resultado"]),
        "Diagnostico": diagnosis,
        "Solucion_ERP": solution,
        "Supuestos": sup,
        "CRM_Pedidos": data["CRM"],
        "SRM_Proveedores": data["SRM"],
        "MRP_Materiales": mrp,
        "Demanda": data["Demanda"],
        "MPS": mps_calc,
        "Capacidad": capacity,
        "DRP_Distribucion": drp_calc,
        "Costos": costs,
        "Riesgos": data["Riesgos"],
    }
    if selected_sheets:
        sheets = {name: frame for name, frame in sheets.items() if name in selected_sheets}
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    output.seek(0)
    wb = load_workbook(output)
    green = "52796F"
    pale = "E5F0E7"
    light = "F4F8F4"
    thin = Side(style="thin", color="D6E4D8")
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_view.showGridLines = False
        for cell in ws[1]:
            cell.font = Font(name="Arial", bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=green)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 30
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name="Arial", color="26352B", size=10)
                cell.fill = PatternFill("solid", fgColor=("FFFFFF" if cell.row % 2 == 0 else light))
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(bottom=thin)
        for col_cells in ws.columns:
            letter = get_column_letter(col_cells[0].column)
            max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells[:100])
            ws.column_dimensions[letter].width = min(max(max_len + 3, 13), 42)
    # Gráfica de capacidad dentro de la hoja Resumen.
    if "Capacidad" in wb.sheetnames and "Resumen" in wb.sheetnames:
        ws = wb["Capacidad"]
        headers = [c.value for c in ws[1]]
        if "Operación" in headers and "Capacidad por proceso (u)" in headers:
            cat_col = headers.index("Operación") + 1
            val_col = headers.index("Capacidad por proceso (u)") + 1
            chart = BarChart()
            chart.title = "Capacidad estimada por operación"
            chart.y_axis.title = "Unidades"
            chart.x_axis.title = "Operación"
            chart.height, chart.width = 8, 15
            chart.add_data(Reference(ws, min_col=val_col, min_row=1, max_row=ws.max_row), titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=cat_col, min_row=2, max_row=ws.max_row))
            wb["Resumen"].add_chart(chart, "D2")
    wb.calculation.fullCalcOnLoad = True
    final = io.BytesIO()
    wb.save(final)
    return final.getvalue()

# ---------------------------- NAVEGACIÓN ----------------------------
st.markdown("""
<div class="hero">
<h1>MOVITECH | PROYECTO 2: ERP INTEGRAL</h1>
<p>Caso práctico: integración de procesos para fabricar y distribuir scooters eléctricos</p>
<p><b>Ingeniería Industrial · Tecnología Informática · Equipo 2</b></p>
</div>
""",unsafe_allow_html=True)
st.markdown('<div class="tip"><b>Modelo práctico editable:</b> los datos iniciales son supuestos académicos. Modifica las tablas, guarda y revisa cómo cambian los cálculos y el dashboard.</div>',unsafe_allow_html=True)

PAGES=[
"1. Caso empresarial y problemática","2. Diagnóstico con datos","3. Solución ERP y módulos",
"4. CRM: clientes y pedidos","5. SRM: proveedores","6. MRP: materiales e inventario",
"7. Demanda y MPS","8. Capacidad de planta","9. DRP y SCM: distribución",
"10. Costos y riesgos","11. Dashboard integral","12. Reportes y respaldo",
"13. Archivos Excel y código fuente"
]
page=st.sidebar.radio("CONTENIDO DEL PROYECTO",PAGES)
st.sidebar.markdown("---")
st.sidebar.caption("Base local SQLite · No requiere conexión a internet para operar después de instalar librerías.")
st.sidebar.caption(f"Archivo de datos: {DB.name}")

# ---------------------------- 1. CASO ----------------------------
if page==PAGES[0]:
    st.header("1. Caso empresarial y problemática")
    st.subheader("Contexto de MOVITECH")
    st.write("MOVITECH es una empresa ficticia que fabrica y comercializa scooters eléctricos. Para este caso práctico debe atender una demanda de referencia de 200 unidades en 20 días laborables. La empresa necesita coordinar pedidos, proveedores, materiales, producción, capacidad, costos y distribución.")
    sup=edit_table("Supuestos",SUPUESTOS,"supuestos","Estos supuestos alimentan los cálculos de capacidad y costos. La demanda efectiva se toma de la suma de unidades registradas en CRM; este valor queda como respaldo si CRM está vacío. Modifica días, turnos, eficiencia y costos.")
    st.subheader("Problema principal (editable)")
    default_problem=("MOVITECH administra pedidos, compras, inventario, producción y distribución en registros separados. "
                     "Esto dificulta conocer en un mismo momento cuántos scooters se deben fabricar, qué componentes faltan, "
                     "qué proveedor debe surtirlos, si la planta tiene capacidad suficiente y qué pedidos pueden entregarse a tiempo.")
    problem=st.text_area("Redacción de la problemática",load_text("problema",default_problem),height=125)
    if st.button("Guardar problemática",type="primary"):
        save_text("problema",problem); st.success("Problemática guardada.")
    st.subheader("Objetivo de la propuesta")
    default_goal=("Diseñar una propuesta de ERP que integre la información comercial, de proveedores, inventarios, "
                  "producción, capacidad, costos y distribución para planear la fabricación de 200 scooters y apoyar "
                  "decisiones con datos compartidos e indicadores.")
    goal=st.text_area("Objetivo general (editable)",load_text("objetivo",default_goal),height=100)
    if st.button("Guardar objetivo"):
        save_text("objetivo",goal); st.success("Objetivo guardado.")
    st.subheader("Flujo general")
    st.write("CRM → ERP central → Demanda → MPS → MRP / SRM → Producción y capacidad → DRP / SCM → Cliente")

# ---------------------------- 2. DIAGNÓSTICO ----------------------------
elif page==PAGES[1]:
    st.header("2. Diagnóstico de la problemática con datos")
    st.write("Aquí la problemática se convierte en un registro de trabajo: cada problema tiene evidencia, consecuencia, solución e indicador.")
    diag=edit_table("Diagnóstico",PROBLEMA,"diagnostico","Agrega o modifica problemas detectados. Relaciona cada causa con una acción del ERP y un indicador que permita verificarla.")
    all_data=get_all()
    sup=all_data["Supuestos"]
    cap=calc_capacity(sup,all_data["Ruta"])
    mrp=calc_mrp(sup,all_data["BOM_MRP"])
    drp=calc_drp(all_data["DRP"])
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Demanda del caso",f"{demanda_integrada(sup):,.0f} scooters")
    c2.metric("Capacidad de línea",f"{cap['Capacidad por proceso (u)'].min():,.1f} u.")
    c3.metric("Componentes por abastecer",f"{mrp['Requerimiento neto'].sum():,.0f} u.")
    c4.metric("Reabastecimiento DRP",f"{drp['Reabastecimiento sugerido'].sum():,.0f} u.")
    st.subheader("Evidencia numérica para explicar el problema")
    left,right=st.columns(2)
    with left:
        st.write("**Capacidad por operación**")
        st.dataframe(cap[["Operación","Capacidad por proceso (u)","Utilización estimada (%)"]],use_container_width=True,hide_index=True)
        fig=px.bar(cap,x="Operación",y="Capacidad por proceso (u)",title="Capacidad estimada por operación")
        st.plotly_chart(fig,use_container_width=True)
    with right:
        st.write("**Materiales que requiere abastecer MRP**")
        st.dataframe(mrp[["Componente","Requerimiento bruto","Inventario disponible","Recepciones programadas","Requerimiento neto"]],use_container_width=True,hide_index=True)
        fig=px.bar(mrp,x="Componente",y="Requerimiento neto",title="Requerimiento neto por componente")
        st.plotly_chart(fig,use_container_width=True)
    st.info("Interpretación: el ERP se justifica porque la demanda debe relacionarse con inventarios, proveedores y capacidad. Si cada área conserva datos aislados, no se identifica oportunamente qué comprar, qué producir y qué distribuir.")

# ---------------------------- 3. SOLUCIÓN ----------------------------
elif page==PAGES[2]:
    st.header("3. Solución propuesta: ERP integrado")
    st.write("El ERP será el sistema central que conecte los módulos. Esta tabla explica qué información entra, qué proceso se realiza y qué resultado debe utilizar otra área.")
    solution=pd.DataFrame([
        ["CRM","Clientes y pedidos","Pedidos confirmados por zona","Concentrar pedidos y fechas requeridas","Demanda para planeación"],
        ["ERP","Información de áreas","Registros de ventas, compras, almacén, producción y costos","Unificar catálogos y movimientos","Datos centrales consistentes"],
        ["SRM","Proveedores","Precios, tiempos, calidad y entregas","Evaluar proveedores y controlar abastecimiento","Órdenes y fechas de recepción"],
        ["SCM","Cadena de suministro","Pedidos, materiales, fabricación y envíos","Coordinar el flujo completo","Visibilidad de suministro y entrega"],
        ["Demanda","Pronóstico","Historial y pedidos confirmados","Estimar demanda por periodo","Pronóstico para MPS"],
        ["MPS","Plan maestro","Demanda, inventario y calendario","Definir unidades por periodo","Programa de producción"],
        ["MRP","Materiales","MPS, lista de materiales y existencias","Calcular requerimiento bruto y neto","Compras de componentes"],
        ["Capacidad","Capacidad de planta","Tiempos, estaciones, turnos y eficiencia","Calcular capacidad y cuello de botella","Validación del programa"],
        ["DRP","Distribución","Demanda por zona, inventario y tránsito","Calcular reabastecimiento","Plan de envío por centro"],
    ],columns=["Módulo","Proceso","Datos de entrada","Actividad del sistema","Resultado compartido"])
    sol=edit_table("Solución ERP",solution,"solucion_erp","Puedes adaptar la función de cada módulo a la propuesta de tu equipo.")
    st.subheader("Mapa de integración")
    st.write("CRM → ERP → Demanda → MPS → MRP → SRM / Compras → Producción → DRP → SCM → Cliente")
    st.write("Ejemplo de integración: un pedido registrado en CRM aumenta la demanda; MPS define cuántos scooters fabricar; MRP calcula componentes; SRM ayuda a comprarlos; capacidad verifica si pueden fabricarse; DRP organiza la distribución.")

# ---------------------------- 4. CRM ----------------------------
elif page==PAGES[3]:
    st.header("4. CRM: clientes y pedidos")
    st.write("Práctica: registrar los pedidos que originan la demanda de producción. La suma de unidades de los pedidos es la demanda integrada que alimenta MRP, capacidad y costos.")
    df=edit_table("CRM",CRM,"crm","Cada fila es un pedido. Cambia unidades o zona y guarda para reflejar la demanda comercial.")
    if not df.empty:
        total=num(df,"Unidades").sum()
        a,b,c=st.columns(3)
        a.metric("Pedidos registrados",len(df)); b.metric("Unidades solicitadas",f"{total:,.0f}"); c.metric("Clientes",df["Cliente"].nunique() if "Cliente" in df else 0)
        if "Zona" in df and "Unidades" in df:
            resumen=df.groupby("Zona",dropna=False)["Unidades"].sum().reset_index()
            st.dataframe(resumen,use_container_width=True,hide_index=True)
            st.plotly_chart(px.bar(resumen,x="Zona",y="Unidades",title="Demanda por zona"),use_container_width=True)

# ---------------------------- 5. SRM ----------------------------
elif page==PAGES[4]:
    st.header("5. SRM: relación con proveedores")
    st.write("Práctica: comparar proveedores por componente, existencias, puntualidad, defectos, costo y tiempo de entrega.")
    df=edit_table("SRM",SRM,"srm","Actualiza las condiciones de los proveedores. La información sirve para decidir a quién comprar y qué riesgo revisar.")
    if not df.empty:
        a,b,c=st.columns(3)
        a.metric("Proveedores",len(df))
        b.metric("Puntualidad promedio",f"{num(df,'Entregas a tiempo (%)').mean():.1f}%")
        c.metric("Defectos recibidos",f"{num(df,'Defectos recibidos').sum():,.0f}")
        st.plotly_chart(px.bar(df,x="Proveedor",y="Entregas a tiempo (%)",title="Puntualidad por proveedor"),use_container_width=True)
        st.caption("El proveedor con menor puntualidad requiere seguimiento; no debe seleccionarse solo por precio.")

# ---------------------------- 6. MRP ----------------------------
elif page==PAGES[5]:
    st.header("6. MRP: requerimientos de materiales")
    st.write("Práctica: para fabricar cada scooter se necesita una lista de materiales. MRP calcula cuánto componente falta después de descontar inventario y recepciones programadas.")
    bom=edit_table("BOM_MRP",BOM,"mrp","Edita cantidad por scooter, inventario disponible y recepciones programadas. El cálculo utiliza la suma de unidades registradas en CRM; si CRM está vacío, utiliza la demanda de Supuestos.")
    sup=load_df("Supuestos",SUPUESTOS)
    calc=calc_mrp(sup,bom)
    st.subheader("Cálculo de requerimientos")
    st.dataframe(calc,use_container_width=True,hide_index=True)
    st.latex(r"\text{Requerimiento neto}=\max(0,\text{Requerimiento bruto}-\text{Inventario}-\text{Recepciones})")
    st.metric("Total de componentes por abastecer",f"{num(calc,'Requerimiento neto').sum():,.0f} unidades")
    st.plotly_chart(px.bar(calc,x="Componente",y="Requerimiento neto",title="Componentes que deben comprarse"),use_container_width=True)
    st.info("La cantidad calculada es una necesidad de compra estimada. Antes de emitir una orden real se deben validar lotes, mermas, fechas y políticas de inventario.")

# ---------------------------- 7. DEMANDA / MPS ----------------------------
elif page==PAGES[6]:
    st.header("7. Planeación de demanda y MPS")
    st.subheader("7.1 Datos de demanda")
    d=edit_table("Demanda",DEMANDA,"demanda","Compara pronóstico y pedidos confirmados por periodo.")
    if not d.empty:
        d["Diferencia pedidos-pronóstico"]=num(d,"Pedidos confirmados")-num(d,"Demanda pronosticada")
        st.dataframe(d,use_container_width=True,hide_index=True)
        st.plotly_chart(px.bar(d,x="Periodo",y=["Demanda pronosticada","Pedidos confirmados"],barmode="group",title="Demanda pronosticada vs. pedidos"),use_container_width=True)
    st.divider()
    st.subheader("7.2 Plan Maestro de Producción")
    m=edit_table("MPS",MPS,"mps","Edita demanda, producción planeada e inventario inicial por periodo.")
    mc=calc_mps(m)
    st.dataframe(mc,use_container_width=True,hide_index=True)
    if not mc.empty:
        st.metric("Producción total planeada",f"{num(mc,'Producción planeada').sum():,.0f} scooters")
        st.plotly_chart(px.line(mc,x="Periodo",y=["Demanda","Producción planeada"],markers=True,title="Plan de producción vs. demanda"),use_container_width=True)
    st.caption("El MPS debe comprobarse contra la disponibilidad de materiales (MRP) y la capacidad de planta.")

# ---------------------------- 8. CAPACIDAD ----------------------------
elif page==PAGES[7]:
    st.header("8. Planeación de capacidad de planta")
    st.write("Esta práctica toma la estructura del archivo que compartiste: ruta de proceso, tiempos, estaciones, operadores, capacidad, utilización y cuello de botella.")
    ruta=edit_table("Ruta",RUTA,"ruta","Los tiempos y recursos son supuestos didácticos. Modifícalos para analizar cómo cambia la capacidad.")
    sup=load_df("Supuestos",SUPUESTOS)
    c=calc_capacity(sup,ruta)
    st.subheader("Capacidad por operación")
    st.dataframe(c,use_container_width=True,hide_index=True)
    if not c.empty:
        bottleneck=c.loc[c["Capacidad por proceso (u)"].idxmin()]
        capline=c["Capacidad por proceso (u)"].min()
        demand=demanda_integrada(sup)
        a,b,d=st.columns(3)
        a.metric("Capacidad de línea",f"{capline:,.1f} scooters")
        b.metric("Demanda",f"{demand:,.0f} scooters")
        d.metric("Margen de capacidad",f"{capline-demand:,.1f} scooters")
        st.write(f"**Cuello de botella estimado:** {bottleneck['Operación']} ({bottleneck['Capacidad por proceso (u)']:.1f} unidades por periodo).")
        st.plotly_chart(px.bar(c,x="Operación",y="Capacidad por proceso (u)",title="Capacidad por operación"),use_container_width=True)
        st.plotly_chart(px.bar(c,x="Operación",y="Utilización estimada (%)",title="Utilización estimada por operación"),use_container_width=True)
        if capline>=demand: st.success("Con los supuestos actuales, la capacidad estimada cubre la demanda.")
        else: st.warning("La capacidad estimada no cubre la demanda. Se deben revisar tiempos, estaciones, turnos o eficiencia.")
    st.caption("El cálculo es una simulación académica. No sustituye un estudio de tiempos ni una validación de seguridad y calidad.")

# ---------------------------- 9. DRP / SCM ----------------------------
elif page==PAGES[8]:
    st.header("9. DRP y SCM: distribución")
    st.write("Práctica: revisar demanda por zona, inventario disponible y unidades en tránsito para calcular el reabastecimiento sugerido.")
    df=edit_table("DRP",DRP,"drp","Edita las necesidades de cada centro, existencias, tránsito y tiempo de reabastecimiento.")
    calc=calc_drp(df)
    st.subheader("Plan de distribución calculado")
    st.dataframe(calc,use_container_width=True,hide_index=True)
    a,b=st.columns(2)
    a.metric("Demanda total por zonas",f"{num(calc,'Demanda').sum():,.0f}")
    b.metric("Reabastecimiento sugerido",f"{num(calc,'Reabastecimiento sugerido').sum():,.0f}")
    st.plotly_chart(px.bar(calc,x="Centro / zona",y="Reabastecimiento sugerido",title="Reabastecimiento sugerido por zona"),use_container_width=True)
    st.caption("DRP calcula necesidades por ubicación. SCM coordina el flujo completo con compras, inventario, producción y transporte.")

# ---------------------------- 10. COSTOS / RIESGOS ----------------------------
elif page==PAGES[9]:
    st.header("10. Costos y riesgos del proyecto")
    st.subheader("10.1 Estimación de costos")
    all_data=get_all()
    costs=calc_costs(all_data["Supuestos"],all_data["Ruta"])
    costs.loc[len(costs)] = ["TOTAL",None,None,num(costs,"Costo estimado (MXN)").sum()]
    st.dataframe(costs,use_container_width=True,hide_index=True)
    total=num(costs,"Costo estimado (MXN)").sum()
    demand=demanda_integrada(all_data["Supuestos"])
    st.metric("Costo estimado por scooter",f"${(total/demand if demand else 0):,.2f} MXN")
    st.caption("Estimación académica basada en los supuestos editables; no incluye impuestos, inversión inicial ni costos comerciales.")
    st.divider()
    st.subheader("10.2 Registro de riesgos")
    risks=edit_table("Riesgos",RIESGOS,"riesgos","Actualiza la probabilidad, impacto, acciones, responsables e indicadores preventivos.")
    if not risks.empty and "Impacto" in risks:
        st.dataframe(risks.groupby(["Probabilidad","Impacto"],dropna=False).size().reset_index(name="Número de riesgos"),use_container_width=True,hide_index=True)

# ---------------------------- 11. DASHBOARD ----------------------------
elif page==PAGES[10]:
    st.header("11. Dashboard integral del ERP")
    data=get_all()
    sup=data["Supuestos"]
    orders=data["CRM"]
    suppliers=data["SRM"]
    mrp=calc_mrp(sup,data["BOM_MRP"])
    capacity=calc_capacity(sup,data["Ruta"])
    drp=calc_drp(data["DRP"])
    mps=calc_mps(data["MPS"])
    demand=demanda_integrada(sup)
    capline=capacity["Capacidad por proceso (u)"].min() if not capacity.empty else 0
    total_cost=calc_costs(sup,data["Ruta"])["Costo estimado (MXN)"].sum()
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Demanda planificada",f"{demand:,.0f} scooters")
    c2.metric("Pedidos CRM",f"{num(orders,'Unidades').sum():,.0f} unidades")
    c3.metric("Capacidad de línea",f"{capline:,.1f} unidades")
    c4.metric("Margen de capacidad",f"{capline-demand:,.1f} unidades")
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Componentes por comprar",f"{num(mrp,'Requerimiento neto').sum():,.0f}")
    c2.metric("Puntualidad proveedores",f"{num(suppliers,'Entregas a tiempo (%)').mean():.1f}%")
    c3.metric("Reabastecimiento DRP",f"{num(drp,'Reabastecimiento sugerido').sum():,.0f}")
    c4.metric("Costo estimado",f"${total_cost:,.0f} MXN")
    left,right=st.columns(2)
    with left:
        st.plotly_chart(px.bar(capacity,x="Operación",y="Capacidad por proceso (u)",title="Capacidad por operación"),use_container_width=True)
        st.plotly_chart(px.bar(mrp,x="Componente",y="Requerimiento neto",title="Requerimientos de materiales"),use_container_width=True)
    with right:
        st.plotly_chart(px.bar(drp,x="Centro / zona",y="Reabastecimiento sugerido",title="Reabastecimiento por centro"),use_container_width=True)
        if not suppliers.empty:
            st.plotly_chart(px.bar(suppliers,x="Proveedor",y="Entregas a tiempo (%)",title="Desempeño de proveedores"),use_container_width=True)
    st.subheader("Conclusión automática del escenario")
    if capline>=demand:
        st.success(f"La capacidad estimada cubre la demanda de {demand:.0f} scooters con los supuestos actuales. El análisis debe concentrarse en asegurar los materiales requeridos y coordinar la distribución.")
    else:
        st.warning(f"La capacidad estimada ({capline:.1f}) es menor que la demanda ({demand:.0f}). Antes de comprometer entregas, MOVITECH debe ajustar la operación cuello de botella.")
    st.info("Los indicadores cambian al guardar las tablas de cada módulo. Los datos iniciales son simulados.")

# ---------------------------- 12. REPORTES ----------------------------
elif page==PAGES[11]:
    st.header("12. Reportes y respaldo")
    bases={"Supuestos":SUPUESTOS,"Diagnóstico":PROBLEMA,"CRM":CRM,"SRM":SRM,"BOM_MRP":BOM,"Demanda":DEMANDA,"MPS":MPS,"Ruta":RUTA,"DRP":DRP,"Riesgos":RIESGOS}
    name=st.selectbox("Tabla para consultar",list(bases.keys()))
    df=load_df(name,bases[name])
    st.dataframe(df,use_container_width=True,hide_index=True)
    st.download_button("Descargar tabla CSV",df.to_csv(index=False).encode("utf-8-sig"),file_name=f"MOVITECH_{name}.csv",mime="text/csv")
    st.subheader("Exportación consolidada")
    pieces=[]
    for n,b in bases.items():
        t=load_df(n,b).copy()
        t.insert(0,"Módulo / tabla",n)
        pieces.append(t)
    combined=pd.concat(pieces,ignore_index=True,sort=False)
    st.download_button("Descargar todos los datos CSV",combined.to_csv(index=False).encode("utf-8-sig"),file_name="MOVITECH_datos_ERP.csv",mime="text/csv")
    if DB.exists():
        st.download_button("Descargar respaldo de base SQLite",DB.read_bytes(),file_name="movitech_erp_respaldo.db",mime="application/octet-stream")
    st.caption("Conserva el archivo movitech_erp_local.db junto al código para mantener los cambios.")


# ---------------------------- 13. EXCEL Y CÓDIGO ----------------------------
elif page==PAGES[12]:
    st.header("13. Archivos Excel y código fuente")
    st.write("Desde esta página puedes descargar el libro integral actualizado con los datos guardados en el ERP y consultar el código Python que ejecuta la aplicación.")
    st.subheader("Libros de Excel")
    excel_data = build_excel_bytes()
    st.download_button(
        "⬇️ Descargar Excel integral de MOVITECH",
        data=excel_data,
        file_name="MOVITECH_ERP_Integral.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    st.caption("El archivo reúne diagnóstico, solución ERP, supuestos, CRM, SRM, MRP, demanda, MPS, capacidad, DRP, costos y riesgos. Se genera con los datos guardados actualmente.")
    st.write("**Descargas por práctica:**")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Excel 1 · Diagnóstico y solución",
            data=build_excel_bytes(["Diagnostico","Solucion_ERP","Supuestos"]),
            file_name="MOVITECH_01_Diagnostico_y_Solucion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
        st.download_button("Excel 2 · CRM y SRM",
            data=build_excel_bytes(["CRM_Pedidos","SRM_Proveedores"]),
            file_name="MOVITECH_02_CRM_y_SRM.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with c2:
        st.download_button("Excel 3 · Demanda, MPS, MRP y capacidad",
            data=build_excel_bytes(["Demanda","MPS","MRP_Materiales","Capacidad","Supuestos"]),
            file_name="MOVITECH_03_Demanda_MPS_MRP_Capacidad.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
        st.download_button("Excel 4 · DRP, costos y riesgos",
            data=build_excel_bytes(["DRP_Distribucion","Costos","Riesgos"]),
            file_name="MOVITECH_04_DRP_Costos_Riesgos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    st.subheader("Código fuente de la aplicación")
    code_text = Path(__file__).read_text(encoding="utf-8")
    st.download_button(
        "⬇️ Descargar código Python (.py)",
        data=code_text.encode("utf-8"),
        file_name="movitech_proyecto_2_erp.py",
        mime="text/x-python"
    )
    st.caption("También puedes desplegar el código aquí mismo para revisarlo sin salir de la aplicación.")
    with st.expander("👩‍💻 Visualizar código Python completo", expanded=True):
        st.caption(f"Código fuente completo · {len(code_text.splitlines()):,} líneas")
        st.code(code_text, language="python", line_numbers=True)
    st.subheader("Respaldo de datos")
    if DB.exists():
        st.download_button("Descargar base de datos SQLite", DB.read_bytes(),
                           file_name="movitech_erp_respaldo.db", mime="application/octet-stream")
    st.info("Para conservar los cambios, mantén movitech_erp_local.db en la misma carpeta que el archivo Python.")

st.markdown("---")
st.caption("MOVITECH | Proyecto 2 · Caso académico simulado · Los resultados dependen de los datos y supuestos que se ingresen.")
