# -*- coding: utf-8 -*-
"""
Created on Tue May 12 16:43:59 2026

@author: ISAHISURISADAYIBARRA
"""
"""
@author: ISAHISURISADAYIBARRA
"""
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import psycopg2
import io
import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(layout="wide", page_title="Sistema Biomédico HDLM")

# --- DISEÑO AZUL Y NEGRO ---
st.markdown("""
    <style>
    .stApp { background-color: #003366; }
    .stForm { 
        background-color: #1a1a1a !important; 
        border: 2px solid #000000; 
        border-radius: 10px; 
        padding: 20px; 
    }
    label, p, h1, h2, h3, div { color: #ffffff !important; }
    [data-testid="stSidebar"] { background-color: #001a33; }
    </style>
""", unsafe_allow_html=True)

# --- FUNCIONES ---
def get_connection(): 
    return psycopg2.connect(**st.secrets["database"])

def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def generate_pdf(df, titulo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    elements = [Paragraph(titulo, styles['Title']), Spacer(1, 12)]
    data = [df.columns.tolist()] + df.values.tolist()
    t = Table(data)
    t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.grey)]))
    elements.append(t); doc.build(elements)
    return buffer.getvalue()

# --- SEGURIDAD ---
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.title("Acceso al Sistema Biomédico")
    pwd = st.text_input("Contraseña:", type="password")
    if st.button("Ingresar"):
        if "auth" in st.secrets and pwd == st.secrets["auth"]["password"]:
            st.session_state.authenticated = True
            st.rerun()
    st.stop()

# --- INTERFAZ PRINCIPAL ---
choice = st.sidebar.selectbox("Módulo", ["Inventario", "Mantenimiento", "Bajas"])

if choice == "Inventario":
    st.header("Inventario de Equipos")
    with st.form("reg", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1: eq = st.text_input("Nombre del Equipo"); ma = st.text_input("Marca"); mo = st.text_input("Modelo")
        with c2: se = st.text_input("Serie"); ub = st.text_input("Ubicación"); es = st.text_input("Estado")
        if st.form_submit_button("Guardar"):
            conn = get_connection(); cur = conn.cursor()
            cur.execute("INSERT INTO inventario (equipo, marca, modelo, serie, ubicacion, estado) VALUES (%s,%s,%s,%s,%s,%s)", (eq, ma, mo, se, ub, es))
            conn.commit(); conn.close(); st.success("Guardado")
    
    conn = get_connection(); df = pd.read_sql("SELECT * FROM inventario WHERE estado != 'Baja'", conn); conn.close()
    st.dataframe(df, use_container_width=True)
    
    st.divider()
    st.download_button("Descargar Excel", data=generate_excel(df), file_name="inventario.xlsx")

elif choice == "Mantenimiento":
    st.header("Registro de Mantenimiento")
    with st.form("form_manto", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1: equipo = st.text_input("Nombre del Equipo"); serie = st.text_input("Serie del Equipo"); fecha = st.text_input("Fecha")
        with c2: tipo = st.text_input("Tipo de Mantenimiento"); tec = st.text_input("Técnico"); costo = st.text_input("Costo")
        prox = st.text_input("Próximo mantenimiento"); desc = st.text_area("Descripción detallada del trabajo")
        if st.form_submit_button("Guardar Mantenimiento"):
            conn = get_connection(); cur = conn.cursor()
            cur.execute("INSERT INTO mantenimientos (equipo_info, fecha_mantenimiento, tipo, tecnico, costo, descripcion, proximo_mantenimiento) VALUES (%s,%s,%s,%s,%s,%s,%s)", (f"{equipo} (S/N: {serie})", fecha, tipo, tec, costo, desc, prox))
            conn.commit(); conn.close(); st.success("Guardado")
    conn = get_connection(); df = pd.read_sql("SELECT * FROM mantenimientos", conn); conn.close()
    st.dataframe(df, use_container_width=True)
    
    st.divider()
    st.download_button("Descargar Excel", data=generate_excel(df), file_name="mantenimientos.xlsx")

elif choice == "Bajas":
    st.header("Control de Bajas")
    conn = get_connection()
    df = pd.read_sql("SELECT * FROM inventario WHERE estado != 'Baja'", conn)
    conn.close()

    if not df.empty:
        with st.form("form_baja_completo", clear_on_submit=True):
            st.subheader("Datos del equipo a dar de baja")
            seleccion = st.selectbox("Seleccione el equipo", df["equipo"] + " - " + df["serie"])
            equipo_sel = df[df["equipo"] + " - " + df["serie"] == seleccion].iloc[0]
            
            c1, c2 = st.columns(2)
            with c1: motivo = st.text_input("Motivo de la baja"); obs = st.text_area("Descripción detallada"); autorizado = st.text_input("Autorizado por")
            with c2: destino = st.text_input("Destino"); folio = st.text_input("Folio de acta"); fecha_acta = st.date_input("Fecha del acta"); valor_res = st.number_input("Valor residual", min_value=0.0)

            if st.form_submit_button("Confirmar Baja Definitiva"):
                conn = get_connection(); cur = conn.cursor()
                cur.execute("UPDATE inventario SET estado='Baja' WHERE id=%s", (int(equipo_sel['id']),))
                cur.execute("INSERT INTO bajas (id_equipo, fecha_baja, motivo, descripcion_motivo, quien_autorizo, destino, folio_acta, fecha_acta, valor_residual) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)", (int(equipo_sel['id']), datetime.date.today(), motivo, obs, autorizado, destino, folio, fecha_acta, valor_res))
                conn.commit(); conn.close(); st.success("Baja procesada")
                st.rerun()

    st.subheader("Historial de Bajas")
    conn = get_connection()
    df_h = pd.read_sql("SELECT b.*, i.equipo FROM bajas b JOIN inventario i ON b.id_equipo = i.id", conn)
    conn.close()
    st.dataframe(df_h, use_container_width=True)
    
    st.divider()
    st.download_button("Descargar Reporte Bajas (PDF)", data=generate_pdf(df_h, "Reporte de Bajas"), file_name="bajas.pdf")No hay equipos disponibles para dar de baja.")
