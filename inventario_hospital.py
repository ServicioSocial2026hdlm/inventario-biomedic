# -*- coding: utf-8 -*-
"""
Created on Tue May 12 16:43:59 2026

@author: ISAHISURISADAYIBARRA
"""
import streamlit as st
import pandas as pd
import psycopg2
import io
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Sistema Biomédico - Hospital de la Mujer", layout="wide")

# Seguridad
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.title("Acceso al Sistema Biomédico")
    pwd = st.text_input("Contraseña:", type="password")
    if st.button("Entrar"):
        if "auth" in st.secrets and pwd == st.secrets["auth"]["password"]:
            st.session_state.authenticated = True
            st.rerun()
    st.stop()

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
    t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.grey), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke)]))
    elements.append(t); doc.build(elements)
    return buffer.getvalue()

# Interfaz
choice = st.sidebar.selectbox("Módulo", ["Inventario", "Mantenimiento", "Bajas"])

if choice == "Inventario":
    with st.expander("Registrar Equipo"):
        with st.form("reg"):
            c1, c2 = st.columns(2)
            with c1: eq = st.text_input("Equipo"); ma = st.text_input("Marca"); mo = st.text_input("Modelo")
            with c2: se = st.text_input("Serie"); ub = st.selectbox("Ubicación", ["Tococirugía", "Quirófano", "UCIN"]); es = st.selectbox("Estado", ["Operativo", "En Mantenimiento", "Baja"])
            if st.form_submit_button("Guardar"):
                conn = get_connection(); cur = conn.cursor()
                cur.execute("INSERT INTO inventario (equipo, marca, modelo, serie, ubicacion, estado) VALUES (%s,%s,%s,%s,%s,%s)", (eq, ma, mo, se, ub, es))
                conn.commit(); conn.close(); st.rerun()
    
    conn = get_connection(); df = pd.read_sql("SELECT * FROM inventario", conn); conn.close()
    st.dataframe(df, use_container_width=True)
    st.download_button("Descargar Excel", generate_excel(df), "inventario.xlsx")

elif choice == "Mantenimiento":
    st.header("Registro de Mantenimiento")
    conn = get_connection()
    equipos = pd.read_sql("SELECT id, equipo || ' (' || serie || ')' as label FROM inventario", conn)
    with st.form("manto"):
        sel = st.selectbox("Equipo", equipos['label'])
        fecha = st.date_input("Fecha"); tipo = st.selectbox("Tipo", ["Preventivo", "Correctivo"]); desc = st.text_area("Descripción")
        if st.form_submit_button("Guardar"): st.success("Registrado")
    conn.close()

elif choice == "Bajas":
    st.header("Acta de Baja")
    conn = get_connection()
    equipos = pd.read_sql("SELECT id, equipo || ' (' || serie || ')' as label FROM inventario WHERE estado != 'Baja'", conn)
    with st.form("baja"):
        sel = st.selectbox("Equipo a dar de baja", equipos['label'])
        fecha = st.date_input("Fecha de baja"); motivo = st.selectbox("Motivo", ["Obsolescencia", "Daño irreparable", "Robo", "Otro"])
        obs = st.text_area("Descripción del motivo"); aut = st.text_input("Quién autoriza"); folio = st.text_input("Folio del Acta")
        if st.form_submit_button("Confirmar Baja"): st.warning("Baja documentada")
    conn.close()
