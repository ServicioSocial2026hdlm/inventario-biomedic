# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import psycopg2
import io
import datetime
import xlsxwriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, letter

# Configuración de la página
st.set_page_config(layout="wide", page_title="Sistema Biomédico HDLM")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f6; }
    .stForm { background-color: #ffffff; border: 1px solid #dcdcdc; border-radius: 8px; padding: 20px; }
    [data-testid="stSidebar"] { background-color: #e6eaf1; }
    </style>
""", unsafe_allow_html=True)

# --- CONEXIÓN ---
def get_connection():
    return psycopg2.connect(**st.secrets["database"])

# ==============================================================
# FUNCIONES DE EXPORTACIÓN — INVENTARIO
# ==============================================================
def generate_excel(df):
    output = io.BytesIO()
    columnas_deseadas = ['id', 'equipo', 'marca', 'modelo', 'serie', 'ubicacion', 'estado', 'valor_adquisicion']
    df_filtrado = df[columnas_deseadas].copy()
    df_filtrado.columns = ["ID", "NOMBRE", "MARCA", "MODELO", "SERIE", "UBICACIÓN", "ESTADO", "VALOR ADQ."]

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Inventario', startrow=6)
        wb = writer.book
        ws = writer.sheets['Inventario']

        header_format = wb.add_format({'bold': True, 'align': 'center', 'valign': 'center', 'border': 1})
        title_format   = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        subtitle_format = wb.add_format({'bold': True, 'font_size': 12, 'align': 'center'})

        ws.merge_range('A2:H2', 'HOSPITAL DE LA MUJER', title_format)
        ws.merge_range('A3:H3', 'INGENIERÍA BIOMÉDICA', subtitle_format)
        ws.merge_range('A4:H4', 'INVENTARIO DE EQUIPO MÉDICO', title_format)
        ws.merge_range('A5:H5', '(F-HM-BM-01)', subtitle_format)

        try:
            ws.insert_image('A1', 'issea.png', {'x_scale': 0.5, 'y_scale': 0.5})
        except:
            pass

        for i, col in enumerate(df_filtrado.columns):
            ws.set_column(i, i, 18)
            ws.write(6, i, col, header_format)

    return output.getvalue()


def generate_pdf_custom(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))

    try:
        c.drawImage("issea.png", 50, 520, width=120, height=60)
    except:
        pass

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(450, 560, "HOSPITAL DE LA MUJER")
    c.drawCentredString(450, 540, "INGENIERÍA BIOMÉDICA")
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(450, 500, "INVENTARIO DE EQUIPO MEDICO")
    c.setFont("Helvetica", 10)
    c.drawCentredString(450, 480, "(F-HM-BM-01)")

    y = 440
    pos_x = [50, 90, 200, 290, 380, 470, 570, 670]
    headers = ["ID", "NOMBRE", "MARCA", "MODELO", "SERIE", "UBICACIÓN", "ESTADO", "VALOR"]
    c.rect(45, y, 705, 30)
    c.setFont("Helvetica-Bold", 10)
    for i, h in enumerate(headers):
        c.drawString(pos_x[i], y + 10, h)

    def draw_footer(canvas_obj):
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.drawRightString(750, 20, "REV-01")

    draw_footer(c)
    y -= 25
    c.setFont("Helvetica", 9)

    for _, row in df.iterrows():
        c.line(45, y + 20, 750, y + 20)
        datos = [
            str(row.get('id', '')), str(row.get('equipo', ''))[:20], str(row.get('marca', ''))[:15],
            str(row.get('modelo', ''))[:15], str(row.get('serie', ''))[:15], str(row.get('ubicacion', ''))[:18],
            str(row.get('estado', ''))[:15], str(row.get('valor_adquisicion', ''))
        ]
        for i, val in enumerate(datos):
            c.drawString(pos_x[i], y + 5, val)
        y -= 25
        if y < 50:
            c.showPage()
            draw_footer(c)
            y = 450
            c.rect(45, y, 705, 30)
            c.setFont("Helvetica-Bold", 10)
            for i, h in enumerate(headers):
                c.drawString(pos_x[i], y + 10, h)
            y -= 25
            c.setFont("Helvetica", 9)

    c.save()
    return buffer.getvalue()


# ==============================================================
# FUNCIONES DE EXPORTACIÓN — MANTENIMIENTO
# ==============================================================
def generate_excel_mtto(df):
    output = io.BytesIO()
    cols = ['equipo_info', 'fecha_mantenimiento', 'tipo', 'tecnico', 'costo', 'descripcion', 'proximo_mantenimiento']
    df_f = df[cols].copy()
    df_f.columns = ["EQUIPO/SERIE", "FECHA", "TIPO", "TÉCNICO", "COSTO", "DESCRIPCIÓN", "PRÓX. MTTO"]

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_f.to_excel(writer, index=False, sheet_name='Mtto', startrow=6)
        wb = writer.book
        ws = writer.sheets['Mtto']
        h_f = wb.add_format({'bold': True, 'align': 'center', 'valign': 'center', 'border': 1})
        t_f = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center'})

        ws.merge_range('A2:G2', 'HOSPITAL DE LA MUJER', t_f)
        ws.merge_range('A3:G3', 'INGENIERÍA BIOMÉDICA', t_f)
        ws.merge_range('A4:G4', 'REGISTRO DE MANTENIMIENTO (F-HM-BM-02)', t_f)

        try:
            ws.insert_image('A1', 'issea.png', {'x_scale': 0.5, 'y_scale': 0.5})
        except:
            pass

        for i, col in enumerate(df_f.columns):
            ws.set_column(i, i, 18)
            ws.write(6, i, col, h_f)

    return output.getvalue()


def generate_pdf_mtto(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))

    try:
        c.drawImage("issea.png", 50, 520, width=120, height=60)
    except:
        pass

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(450, 560, "HOSPITAL DE LA MUJER")
    c.drawCentredString(450, 540, "INGENIERÍA BIOMÉDICA")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(450, 515, "REGISTRO DE MANTENIMIENTO")
    c.setFont("Helvetica", 11)
    c.drawCentredString(450, 498, "(F-HM-BM-02)")

    y = 460
    pos_x = [50, 160, 255, 330, 405, 475, 620]
    headers = ["EQUIPO/SERIE", "FECHA", "TIPO", "TÉCNICO", "COSTO", "DESCRIPCIÓN", "PRÓX. MTTO"]
    c.rect(45, y, 705, 28)
    c.setFont("Helvetica-Bold", 9)
    for i, h in enumerate(headers):
        c.drawString(pos_x[i], y + 9, h)

    def draw_footer(canvas_obj):
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.drawRightString(750, 20, "REV-01")

    draw_footer(c)
    y -= 22
    c.setFont("Helvetica", 8)

    for _, row in df.iterrows():
        c.line(45, y + 18, 750, y + 18)
        datos = [
            str(row.get('equipo_info', ''))[:28],
            str(row.get('fecha_mantenimiento', '')),
            str(row.get('tipo', ''))[:14],
            str(row.get('tecnico', ''))[:14],
            str(row.get('costo', '')),
            str(row.get('descripcion', ''))[:22],
            str(row.get('proximo_mantenimiento', '')),
        ]
        for i, val in enumerate(datos):
            c.drawString(pos_x[i], y + 4, val)
        y -= 22
        if y < 50:
            c.showPage()
            draw_footer(c)
            y = 470
            c.rect(45, y, 705, 28)
            c.setFont("Helvetica-Bold", 9)
            for i, h in enumerate(headers):
                c.drawString(pos_x[i], y + 9, h)
            y -= 22
            c.setFont("Helvetica", 8)

    c.save()
    return buffer.getvalue()


# ==============================================================
# FUNCIONES DE EXPORTACIÓN — BAJAS
# ==============================================================
def generate_excel_bajas(df):
    output = io.BytesIO()
    cols = ['equipo_info', 'fecha_baja', 'motivo', 'quien_autorizo', 'destino', 'folio_acta', 'valor_residual']
    df_f = df[cols].copy()
    df_f.columns = ["EQUIPO/SERIE", "FECHA BAJA", "MOTIVO", "AUTORIZÓ", "DESTINO", "FOLIO ACTA", "VALOR"]

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_f.to_excel(writer, index=False, sheet_name='Bajas', startrow=6)
        wb = writer.book
        ws = writer.sheets['Bajas']
        h_f = wb.add_format({'bold': True, 'align': 'center', 'valign': 'center', 'border': 1})
        t_f = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center'})

        ws.merge_range('A2:G2', 'HOSPITAL DE LA MUJER', t_f)
        ws.merge_range('A3:G3', 'INGENIERÍA BIOMÉDICA', t_f)
        ws.merge_range('A4:G4', 'REGISTRO DE BAJAS (F-HM-BM-03)', t_f)

        try:
            ws.insert_image('A1', 'issea.png', {'x_scale': 0.5, 'y_scale': 0.5})
        except:
            pass

        for i, col in enumerate(df_f.columns):
            ws.set_column(i, i, 18)
            ws.write(6, i, col, h_f)

    return output.getvalue()


def generate_pdf_bajas(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))

    try:
        c.drawImage("issea.png", 50, 520, width=120, height=60)
    except:
        pass

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(450, 560, "HOSPITAL DE LA MUJER")
    c.drawCentredString(450, 540, "INGENIERÍA BIOMÉDICA")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(450, 515, "REGISTRO DE BAJAS")
    c.setFont("Helvetica", 11)
    c.drawCentredString(450, 498, "(F-HM-BM-03)")

    y = 460
    pos_x = [50, 130, 220, 330, 430, 540, 650]
    headers = ["EQUIPO/SERIE", "FECHA BAJA", "MOTIVO", "AUTORIZÓ", "DESTINO", "FOLIO ACTA", "VALOR"]
    c.rect(45, y, 705, 28)
    c.setFont("Helvetica-Bold", 9)
    for i, h in enumerate(headers):
        c.drawString(pos_x[i], y + 9, h)

    def draw_footer(canvas_obj):
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.drawRightString(750, 20, "REV-01")

    draw_footer(c)
    y -= 22
    c.setFont("Helvetica", 8)

    for _, row in df.iterrows():
        c.line(45, y + 18, 750, y + 18)
        datos = [
            str(row.get('equipo_info', ''))[:22],
            str(row.get('fecha_baja', '')),
            str(row.get('motivo', ''))[:18],
            str(row.get('quien_autorizo', ''))[:16],
            str(row.get('destino', ''))[:16],
            str(row.get('folio_acta', ''))[:14],
            str(row.get('valor_residual', '')),
        ]
        for i, val in enumerate(datos):
            c.drawString(pos_x[i], y + 4, val)
        y -= 22
        if y < 50:
            c.showPage()
            draw_footer(c)
            y = 470
            c.rect(45, y, 705, 28)
            c.setFont("Helvetica-Bold", 9)
            for i, h in enumerate(headers):
                c.drawString(pos_x[i], y + 9, h)
            y -= 22
            c.setFont("Helvetica", 8)

    c.save()
    return buffer.getvalue()


# ==============================================================
# FUNCIONES DE EXPORTACIÓN — CONSUMIBLES
# ==============================================================
def generate_excel_consumibles(df):
    output = io.BytesIO()
    cols = ['id', 'descripcion', 'cantidad', 'equipo_compatible']
    df_f = df[cols].copy()
    df_f.columns = ["ID", "DESCRIPCIÓN", "CANTIDAD", "EQUIPO COMPATIBLE"]

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_f.to_excel(writer, index=False, sheet_name='Consumibles', startrow=6)
        wb = writer.book
        ws = writer.sheets['Consumibles']
        h_f = wb.add_format({'bold': True, 'align': 'center', 'valign': 'center', 'border': 1})
        t_f = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center'})

        ws.merge_range('A2:E2', 'HOSPITAL DE LA MUJER', t_f)
        ws.merge_range('A3:E3', 'INGENIERÍA BIOMÉDICA', t_f)
        ws.merge_range('A4:E4', 'CONTROL DE CONSUMIBLES (F-HM-BM-04)', t_f)

        try:
            ws.insert_image('A1', 'issea.png', {'x_scale': 0.5, 'y_scale': 0.5})
        except:
            pass

        for i, col in enumerate(df_f.columns):
            ws.set_column(i, i, 22)
            ws.write(6, i, col, h_f)

    return output.getvalue()


def generate_pdf_consumibles(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(letter))

    try:
        c.drawImage("issea.png", 50, 520, width=120, height=60)
    except:
        pass

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(450, 560, "HOSPITAL DE LA MUJER")
    c.drawCentredString(450, 540, "INGENIERÍA BIOMÉDICA")
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(450, 515, "CONTROL DE CONSUMIBLES")
    c.setFont("Helvetica", 11)
    c.drawCentredString(450, 498, "(F-HM-BM-04)")

    y = 460
    pos_x = [60, 120, 340, 460, 640]
    headers = ["ID", "DESCRIPCIÓN", "CANTIDAD", "EQUIPO COMPATIBLE", "VALOR"]
    c.rect(45, y, 705, 28)
    c.setFont("Helvetica-Bold", 10)
    for i, h in enumerate(headers):
        c.drawString(pos_x[i], y + 9, h)

    def draw_footer(canvas_obj):
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.drawRightString(750, 20, "REV-01")

    draw_footer(c)
    y -= 22
    c.setFont("Helvetica", 9)

    for _, row in df.iterrows():
        c.line(45, y + 18, 750, y + 18)
        datos = [
            str(row.get('id', '')),
            str(row.get('descripcion', ''))[:35],
            str(row.get('cantidad', ''))[:15],
            str(row.get('equipo_compatible', ''))[:25]
            
        ]
        for i, val in enumerate(datos):
            c.drawString(pos_x[i], y + 4, val)
        y -= 22
        if y < 50:
            c.showPage()
            draw_footer(c)
            y = 470
            c.rect(45, y, 705, 28)
            c.setFont("Helvetica-Bold", 10)
            for i, h in enumerate(headers):
                c.drawString(pos_x[i], y + 9, h)
            y -= 22
            c.setFont("Helvetica", 9)

    c.save()
    return buffer.getvalue()


# ==============================================================
# MÓDULO DE EXPORTACIÓN GENÉRICO
# ==============================================================
def export_module(df, nombre, excel_fn, pdf_fn):
    st.write("---")
    st.subheader("📤 Exportar Datos")

    # --- Buscador ---
    col_busq, col_campo = st.columns([3, 1])
    texto = col_busq.text_input("🔍 Buscar en " + nombre.replace("_", " ") + ":", placeholder="Escribe para filtrar...", label_visibility="collapsed")
    columnas_texto = df.select_dtypes(include="object").columns.tolist()
    campo = col_campo.selectbox("Campo en " + nombre, ["Todos"] + columnas_texto, label_visibility="collapsed")

    if texto:
        if campo == "Todos":
            mask = df[columnas_texto].apply(lambda col: col.astype(str).str.contains(texto, case=False, na=False)).any(axis=1)
        else:
            mask = df[campo].astype(str).str.contains(texto, case=False, na=False)
        df_filtrado = df[mask]
    else:
        df_filtrado = df

    # --- Multiselect sobre resultados filtrados ---
    opciones = df_filtrado.index.tolist()
    indices = st.multiselect("Selecciona registros para exportar en " + nombre.replace("_", " ") + " (vacío = todos los filtrados):", opciones,
                             format_func=lambda i: " | ".join(str(df_filtrado.loc[i, c]) for c in columnas_texto[:3] if c in df_filtrado.columns))
    df_f = df_filtrado.loc[indices] if indices else df_filtrado

    if not df_f.empty:
        st.caption(f"{len(df_f)} registro(s) seleccionados para exportar")
        c1, c2 = st.columns(2)
        c1.download_button("📥 Exportar a Excel", excel_fn(df_f), f"{nombre}.xlsx", "application/vnd.ms-excel")
        c2.download_button("📄 Exportar a PDF",   pdf_fn(df_f),   f"{nombre}.pdf",  "application/pdf")


# ==============================================================
# SEGURIDAD
# ==============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("Acceso al Sistema Biomédico")
    pwd = st.text_input("Contraseña:", type="password")
    
    if st.button("Ingresar"):
        if "auth" not in st.secrets:
            st.error("❌ El sistema no detecta el archivo secrets.toml. Revisa que esté en la carpeta .streamlit y no termine en .txt")
        elif pwd != st.secrets["auth"]["password"]:
            st.error("❌ Contraseña incorrecta. Intenta de nuevo.")
        else:
            st.session_state.authenticated = True
            st.rerun()
            
    st.stop()

# ==============================================================
# INTERFAZ PRINCIPAL
# ==============================================================
choice = st.sidebar.selectbox("Módulo", ["Inventario", "Mantenimiento", "Consumibles", "Bajas"])

# ------ INVENTARIO ------
if choice == "Inventario":
    st.header("Inventario de Equipos")
    with st.form("reg", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            eq = st.text_input("Nombre del Equipo")
            ma = st.text_input("Marca")
            mo = st.text_input("Modelo")
        with c2:
            se = st.text_input("Serie")
            ub = st.text_input("Ubicación")
            es = st.text_input("Estado")
        val_adq = st.number_input("Valor de Adquisición", min_value=0.0, format="%.2f")
        
        if st.form_submit_button("Guardar"):
            conn = get_connection()
            cur  = conn.cursor()
            cur.execute(
                "INSERT INTO inventario (equipo, marca, modelo, serie, ubicacion, estado, valor_adquisicion) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (eq, ma, mo, se, ub, es, val_adq)
            )
            conn.commit()
            conn.close()
            st.rerun()

    conn = get_connection()
    df   = pd.read_sql("SELECT * FROM inventario", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)
    export_module(df, "Inventario_Equipos", generate_excel, generate_pdf_custom)


# ------ MANTENIMIENTO ------
elif choice == "Mantenimiento":
    st.header("Registro de Mantenimiento")
    with st.form("form_manto", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            equipo = st.text_input("Nombre del Equipo")
            serie  = st.text_input("Serie del Equipo")
            fecha  = st.date_input("Fecha de Mantenimiento")
        with c2:
            tipo  = st.text_input("Tipo de Mantenimiento")
            tec   = st.text_input("Técnico")
            costo = st.text_input("Costo")
        prox = st.date_input("Próximo mantenimiento")
        desc = st.text_area("Descripción detallada del trabajo")

        if st.form_submit_button("Guardar Mantenimiento"):
            conn = get_connection()
            cur  = conn.cursor()
            try:
                cur.execute(
                    "INSERT INTO mantenimientos (equipo_info, fecha_mantenimiento, tipo, tecnico, costo, descripcion, proximo_mantenimiento) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                    (f"{equipo} - Serie: {serie}", fecha, tipo, tec, float(costo) if costo else 0.0, desc, prox)
                )
                conn.commit()
                st.success("Guardado correctamente")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
            finally:
                cur.close()
                conn.close()

    conn  = get_connection()
    df_mt = pd.read_sql("SELECT * FROM mantenimientos", conn)
    conn.close()
    st.dataframe(df_mt, use_container_width=True)
    export_module(df_mt, "Registro_Mantenimiento", generate_excel_mtto, generate_pdf_mtto)


# ------ CONSUMIBLES ------
elif choice == "Consumibles":
    st.header("Control de Consumibles")
    with st.form("form_consumibles", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            desc_cons = st.text_area("Descripción del Consumible")
            cant_cons = st.text_input("Cantidad")
        with c2:
            eq_comp = st.text_input("Equipo Compatible")
            val_adq_cons = st.number_input("Valor de Adquisición", min_value=0.0, format="%.2f")

        if st.form_submit_button("Guardar Consumible"):
            conn = get_connection()
            cur  = conn.cursor()
            try:
                cur.execute(
                    "INSERT INTO consumibles (descripcion, cantidad, equipo_compatible, valor_adquisicion) VALUES (%s,%s,%s,%s)",
                    (desc_cons, cant_cons, eq_comp, val_adq_cons)
                )
                conn.commit()
                st.success("Consumible guardado exitosamente")
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
            finally:
                cur.close()
                conn.close()

    conn = get_connection()
    df_cs = pd.read_sql("SELECT * FROM consumibles", conn)
    conn.close()
    st.dataframe(df_cs, use_container_width=True)
    export_module(df_cs, "Control_Consumibles", generate_excel_consumibles, generate_pdf_consumibles)


# ------ BAJAS ------
elif choice == "Bajas":
    st.header("Control de Bajas")
    conn = get_connection()
    df   = pd.read_sql("SELECT * FROM inventario WHERE estado != 'Baja'", conn)
    conn.close()

    if not df.empty:
        with st.form("form_baja", clear_on_submit=True):
            seleccion = st.selectbox("Seleccione el equipo", df["equipo"] + " - " + df["serie"])
            c1, c2 = st.columns(2)
            with c1:
                motivo = st.text_input("Motivo")
                obs    = st.text_area("Descripción")
                autor  = st.text_input("Autorizado por")
            with c2:
                destino = st.text_input("Destino")
                folio   = st.text_input("Folio")
                f_acta  = st.date_input("Fecha acta")
                val     = st.number_input("Valor residual", format="%.2f")

            if st.form_submit_button("Confirmar Baja"):
                equipo_sel = df[df["equipo"] + " - " + df["serie"] == seleccion].iloc[0]
                conn = get_connection()
                cur  = conn.cursor()
                try:
                    cur.execute("UPDATE inventario SET estado='Baja' WHERE id=%s", (int(equipo_sel['id']),))
                    cur.execute(
                        "INSERT INTO bajas (equipo_info, fecha_baja, motivo, descripcion_motivo, quien_autorizo, destino, folio_acta, fecha_acta, valor_residual) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                        (f"{equipo_sel['equipo']} - Serie: {equipo_sel['serie']}", datetime.date.today(), motivo, obs, autor, destino, folio, f_acta, val)
                    )
                    conn.commit()
                    st.success("Baja procesada")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    cur.close()
                    conn.close()
    else:
        st.info("No hay equipos para dar de baja.")

    conn      = get_connection()
    df_bajas  = pd.read_sql("SELECT * FROM bajas", conn)
    conn.close()
    st.dataframe(df_bajas, use_container_width=True)
    export_module(df_bajas, "Reporte_Bajas", generate_excel_bajas, generate_pdf_bajas)
