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

# --- CONFIGURACION DE LA PAGINA ---
st.set_page_config(page_title="Sistema de Inventario - Hospital de la Mujer", layout="wide")

# --- SEGURIDAD ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_password():
    if not st.session_state.authenticated:
        st.title("Acceso al Sistema Biomedico")
        pwd = st.text_input("Ingrese la contraseña de acceso:", type="password")
        if st.button("Ingresar"):
            if "auth" in st.secrets and pwd == st.secrets["auth"]["password"]:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta")
        return False
    return True

if not check_password():
    st.stop()

# --- CONEXION A BASE DE DATOS ---
def get_connection():
    try:
        return psycopg2.connect(
            host=st.secrets["database"]["host"],
            database=st.secrets["database"]["dbname"],
            user=st.secrets["database"]["user"],
            password=st.secrets["database"]["password"],
            port=st.secrets["database"]["port"]
        )
    except Exception as e:
        st.error(f"Error de conexion: {e}")
        return None

# --- FUNCION PARA GENERAR PDF PROFESIONAL ---
def generate_pdf(df):
    buffer = io.BytesIO()
    # Usamos landscape (horizontal) para que la tabla quepa bien como en tu imagen
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter))
    elements = []
    
    styles = getSampleStyleSheet()
    title = Paragraph("HOSPITAL DE LA MUJER - REPORTE DE INVENTARIO BIOMEDICO", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1, 12))
    
    # Preparar datos de la tabla (Encabezados + Filas)
    data = [df.columns.tolist()] + df.values.tolist()
    
    # Crear la tabla
    t = Table(data)
    
    # Estilo de la tabla similar a tu imagen 2
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
    ])
    t.setStyle(style)
    
    elements.append(t)
    doc.build(elements)
    return buffer.getvalue()

def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- INTERFAZ PRINCIPAL ---
st.title("Gestion de Inventario Biomedico")
menu = ["Inventario y Consultas", "Mantenimiento", "Bajas"]
choice = st.sidebar.selectbox("Modulo", menu)

ubicaciones_lista = ["Hospitalización"," Alto Riesgo", "Tococirugía", "Quirófano", "Expulsión", "Labor", "UCIN", "Crecimiento y Desarrollo", "Terapia Intensiva", "Imagenología", "Urgencias", "Consulta Externa", "CEYE"]

if choice == "Inventario y Consultas":
    with st.expander("Registrar Nuevo Equipo"):
        with st.form("registro_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                equipo = st.text_input("Equipo")
                marca = st.text_input("Marca")
                modelo = st.text_input("Modelo")
            with c2:
                serie = st.text_input("Serie")
                ubicacion = st.selectbox("Ubicacion", ubicaciones_lista)
                estado = st.selectbox("Estado", ["Operativo", "En Mantenimiento", "Baja"])
            if st.form_submit_button("Guardar"):
                conn = get_connection()
                if conn:
                    cur = conn.cursor()
                    cur.execute("INSERT INTO inventario (equipo, marca, modelo, serie, ubicacion, estado) VALUES (%s,%s,%s,%s,%s,%s)", (equipo, marca, modelo, serie, ubicacion, estado))
                    conn.commit()
                    st.success("Guardado")
                    st.rerun()

    st.header("Consulta de Equipos")
    conn = get_connection()
    if conn:
        df = pd.read_sql("SELECT equipo, marca, modelo, serie, ubicacion, estado FROM inventario ORDER BY equipo ASC", conn)
        
        # Contador y Buscador
        st.metric("Total de equipos", len(df))
        search = st.text_input("Buscar equipo, marca o modelo:")
        if search:
            df = df[df.apply(lambda row: search.lower() in row.astype(str).str.lower().values, axis=1)]

        st.dataframe(df, use_container_width=True)

        st.subheader("Exportar Informacion")
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            st.download_button("Descargar en Excel", generate_excel(df), "inventario.xlsx")
        with col_btn2:
            # NUEVO BOTÓN PARA PDF PROFESIONAL
            pdf_data = generate_pdf(df)
            st.download_button("Descargar en PDF", pdf_data, f"reporte_{datetime.now().strftime('%Y%m%d')}.pdf", "application/pdf")
