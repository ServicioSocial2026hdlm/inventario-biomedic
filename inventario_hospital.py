# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import psycopg2
import io
import xlsxwriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, letter

# Configuración de la página
st.set_page_config(layout="wide", page_title="Sistema Biomédico HDLM")

# --- DISEÑO Y ESTILO DE PÁGINA ---
st.markdown("""
    <style>
    .stApp { background-color: #f4f7f6; }
    .stForm { background-color: #ffffff; border: 1px solid #dcdcdc; border-radius: 8px; padding: 20px; }
    [data-testid="stSidebar"] { background-color: #e6eaf1; }
    </style>
""", unsafe_allow_html=True)

# --- FUNCIONES DE BASE DE DATOS ---
def get_connection(): 
    return psycopg2.connect(**st.secrets["database"])

# --- FUNCIONES DE REPORTES ---
def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def generate_pdf_custom(df, titulo):
    buffer = io.BytesIO()
    # 792 puntos de ancho x 612 puntos de alto (Carta Horizontal)
    c = canvas.Canvas(buffer, pagesize=(792, 612))
    
    # 1. INSERTAR IMAGEN (Logo ISSEA)
    # Debe estar en la misma carpeta que este archivo .py
    try:
        c.drawImage("issea.png", 50, 520, width=100, height=50)
    except:
        pass 

    # 2. ENCABEZADO
    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(450, 550, "INVENTARIO DE EQUIPO MEDICO")
    
    # 3. ENCABEZADOS DE TABLA
    y = 480
    columnas = ["ID", "NOMBRE", "MARCA", "MODELO", "SERIE", "UBICACIÓN", "ESTADO", "FECHA"]
    pos_x = [50, 120, 250, 350, 450, 550, 650, 720]
    
    c.setFont("Helvetica-Bold", 10)
    for i, col in enumerate(columnas):
        c.drawString(pos_x[i], y, col)
    
    c.line(50, 475, 780, 475)
    
    # 4. CONTENIDO
    y -= 25
    c.setFont("Helvetica", 9)
    for _, row in df.iterrows():
        datos = [
            str(row.get('id', '')),
            str(row.get('equipo', '')),
            str(row.get('marca', '')),
            str(row.get('modelo', '')),
            str(row.get('serie', '')),
            str(row.get('ubicacion', '')),
            str(row.get('estado', '')),
            str(row.get('fecha', ''))
        ]
        
        for i, val in enumerate(datos):
            c.drawString(pos_x[i], y, val)
        y -= 20
        
        if y < 50:
            c.showPage()
            y = 550
            
    c.save()
    return buffer.getvalue()

def export_module(df, nombre):
    st.write("---")
    st.subheader("📤 Exportar Datos")
    indices = st.multiselect("Selecciona registros para exportar:", df.index.tolist())
    df_f = df.loc[indices] if indices else df
    
    if not df_f.empty:
        c1, c2 = st.columns(2)
        c1.download_button("📥 Exportar a Excel", generate_excel(df_f), f"{nombre}.xlsx", "application/vnd.ms-excel")
        c2.download_button("📄 Exportar a PDF", generate_pdf_custom(df_f, nombre), f"{nombre}.pdf", "application/pdf")
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

def get_connection(): 
    return psycopg2.connect(**st.secrets["database"])

if choice == "Inventario":
    st.header("Inventario de Equipos")
    with st.form("reg", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1: eq = st.text_input("Nombre del Equipo"); ma = st.text_input("Marca"); mo = st.text_input("Modelo")
        with c2: se = st.text_input("Serie"); ub = st.text_input("Ubicación"); es = st.text_input("Estado")
        if st.form_submit_button("Guardar"):
            conn = get_connection(); cur = conn.cursor()
            cur.execute("INSERT INTO inventario (equipo, marca, modelo, serie, ubicacion, estado) VALUES (%s,%s,%s,%s,%s,%s)", (eq, ma, mo, se, ub, es))
            conn.commit(); conn.close(); st.rerun()
    conn = get_connection(); df = pd.read_sql("SELECT * FROM inventario", conn); conn.close()
    st.dataframe(df, use_container_width=True)
    export_module(df, "Inventario_Equipos")

elif choice == "Mantenimiento":
    st.header("Registro de Mantenimiento")
    # Cambiamos el nombre a "form_manto" para que sea único y no choque con Inventario
    with st.form("form_manto", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1: 
            equipo = st.text_input("Nombre del Equipo")
            serie = st.text_input("Serie del Equipo")
            fecha = st.text_input("Fecha")
        with c2: 
            tipo = st.text_input("Tipo de Mantenimiento")
            tec = st.text_input("Técnico")
            costo = st.text_input("Costo")
        prox = st.text_input("Próximo mantenimiento")
        desc = st.text_area("Descripción detallada del trabajo")
        
        if st.form_submit_button("Guardar Mantenimiento"):
            # Tu lógica de base de datos aquí...
            conn = get_connection(); cur = conn.cursor()
            cur.execute("INSERT INTO mantenimientos (...) VALUES (...)", (...))
            conn.commit(); conn.close(); st.success("Guardado")
            st.rerun() # Esto recarga y limpia el formulario gracias al clear_on_submit

elif choice == "Bajas":
    st.header("Control de Bajas")
    conn = get_connection(); df = pd.read_sql("SELECT * FROM inventario WHERE estado != 'Baja'", conn); conn.close()
    
    if not df.empty:
        # Cambiamos "reg" por "form_baja" para que sea único
        with st.form("form_baja", clear_on_submit=True):
            seleccion = st.selectbox("Seleccione el equipo", df["equipo"] + " - " + df["serie"])
            # Nota: Al usar clear_on_submit, debemos asegurarnos de que la lógica de guardado procese los datos antes de limpiar
            c1, c2 = st.columns(2)
            with c1: 
                motivo = st.text_input("Motivo")
                obs = st.text_area("Descripción")
                autor = st.text_input("Autorizado por")
            with c2: 
                destino = st.text_input("Destino")
                folio = st.text_input("Folio")
                f_acta = st.date_input("Fecha acta")
                val = st.number_input("Valor residual", format="%.2f")
            
            if st.form_submit_button("Confirmar Baja"):
                # Obtenemos el equipo seleccionado justo antes de procesar
                equipo_sel = df[df["equipo"] + " - " + df["serie"] == seleccion].iloc[0]
                
                conn = get_connection(); cur = conn.cursor()
                cur.execute("UPDATE inventario SET estado='Baja' WHERE id=%s", (int(equipo_sel['id']),))
                cur.execute("INSERT INTO bajas (id_equipo, fecha_baja, motivo, descripcion_motivo, quien_autorizo, destino, folio_acta, fecha_acta, valor_residual) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)", 
                            (int(equipo_sel['id']), datetime.date.today(), motivo, obs, autor, destino, folio, f_acta, val))
                conn.commit(); cur.close(); conn.close()
                st.success("Baja procesada")
                st.rerun() # Esto recargará la página y el formulario se limpiará
    else: 
        st.info("No hay equipos para dar de baja.")
    
    # Mostrar tabla de bajas históricas
    conn = get_connection(); df_bajas = pd.read_sql("SELECT * FROM bajas", conn); conn.close()
    st.dataframe(df_bajas, use_container_width=True)
    export_module(df_bajas, "Reporte_Bajas")
