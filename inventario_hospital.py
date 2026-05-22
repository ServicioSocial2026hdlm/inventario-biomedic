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

# --- CONFIGURACION DE LA PAGINA ---
st.set_page_config(
    page_title="Sistema de Inventario - Hospital de la Mujer",
    layout="wide"
)

# --- SEGURIDAD Y ACCESO ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_password():
    if not st.session_state.authenticated:
        st.title("Acceso al Sistema Biomedico")
        st.write("Hospital de la Mujer - Aguascalientes")
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

# --- FUNCION PARA EXCEL ---
def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Inventario')
    return output.getvalue()

# --- INTERFAZ PRINCIPAL ---
st.title("Gestion de Inventario Biomedico")
st.write("Departamento de Ingenieria Clinica")

menu = ["Inventario y Consultas", "Mantenimiento", "Bajas"]
choice = st.sidebar.selectbox("Seleccione Modulo", menu)

# Unico campo desplegable solicitado
ubicaciones_lista = [
    "Hospitalización", "Alto Riesgo", "Tococirugía", "Quirófano", "Expulsión", 
    "Labor", "UCIN", "Crecimiento y Desarrollo", "Terapia Intensiva", 
    "Imagenología", "Urgencias", "Consulta Externa", "CEYE"
]

if choice == "Inventario y Consultas":
    
    # --- FORMULARIO DE REGISTRO ---
    with st.expander("Registrar Nuevo Equipo", expanded=False):
        with st.form("registro_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                equipo = st.text_input("Equipo (ej. Incubadora)")
                marca = st.text_input("Marca")
                modelo = st.text_input("Modelo")
            with col2:
                serie = st.text_input("Numero de Serie")
                ubicacion = st.selectbox("Ubicacion", ubicaciones_lista)
                estado = st.selectbox("Estado Actual", ["Operativo", "En Mantenimiento", "Fuera de Servicio"])
            
            if st.form_submit_button("Guardar Registro"):
                if equipo and serie:
                    conn = get_connection()
                    if conn:
                        cur = conn.cursor()
                        query = "INSERT INTO inventario (equipo, marca, modelo, serie, ubicacion, estado) VALUES (%s, %s, %s, %s, %s, %s)"
                        cur.execute(query, (equipo, marca, modelo, serie, ubicacion, estado))
                        conn.commit()
                        cur.close()
                        conn.close()
                        st.success("Registro guardado exitosamente")
                        st.rerun()
                else:
                    st.warning("Complete el nombre del equipo y serie para continuar")

    st.divider()

    # --- TABLA DE CONSULTAS Y BUSQUEDA ---
    st.header("Consulta de Equipos")
    
    conn = get_connection()
    if conn:
        # Ordenar por equipo y marca de la A a la Z
        query_sql = "SELECT equipo, marca, modelo, serie, ubicacion, estado FROM inventario ORDER BY equipo ASC, marca ASC"
        df = pd.read_sql(query_sql, conn)
        conn.close()

        if not df.empty:
            # Contador de equipos
            st.metric("Total de equipos registrados", len(df))
            
            # Buscador
            search = st.text_input("Buscar equipo, marca o modelo:")
            if search:
                df = df[df.apply(lambda row: search.lower() in row.astype(str).str.lower().values, axis=1)]

            # Tabla principal ordenada
            st.dataframe(df, use_container_width=True)

            # Botones de descarga
            st.subheader("Exportar Informacion")
            c1, c2 = st.columns(2)
            with c1:
                excel_file = generate_excel(df)
                st.download_button(
                    label="Descargar en Excel",
                    data=excel_file,
                    file_name=f"inventario_hospital_{datetime.now().strftime('%Y%m%d')}.xlsx"
                )
            with c2:
                st.write("Para PDF: Presione Ctrl+P y guarde como archivo PDF desde el navegador.")
        else:
            st.info("No hay datos registrados en el sistema.")

elif choice == "Mantenimiento":
    st.header("Modulo de Mantenimiento")
    st.write("Espacio destinado para el historial de servicios preventivos y correctivos.")

elif choice == "Bajas":
    st.header("Control de Bajas")
    st.write("Registro de equipos desincorporados por obsolescencia o daño.")
