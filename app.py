import pandas as pd
import streamlit as st
import io
import re
import requests # LIBRERÍA PARA ENVIAR DATOS A POWER AUTOMATE

st.set_page_config(page_title="Formateador SharePoint", page_icon="⚡", layout="wide")

# --- INICIO DEL FORMATO VISUAL (ESTILOS DEL PORTAL) ---
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        :root {
            --app-fondo: #1B222C;        /* Azul Noche */
            --app-borde: #F6C128;        /* Amarillo Oro */
            --app-texto-titulos: #FFFFFF;/* Blanco */
            --app-texto-datos: #CBD5E1;  /* Gris Claro */
        }

        .stApp {
            background-color: var(--app-fondo);
            color: var(--app-texto-datos);
        }
        
        [data-testid="stHeader"] {
            background-color: rgba(0,0,0,0) !important;
        }

        h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {
            color: var(--app-texto-titulos) !important;
        }

        .stButton > button, .stDownloadButton > button {
            background-color: transparent !important;
            border: 1px solid var(--app-borde) !important;
            color: var(--app-borde) !important;
            border-radius: 8px !important;
            transition: all 0.3s ease !important;
        }

        .stButton > button:hover, .stDownloadButton > button:hover {
            background-color: rgba(246, 193, 40, 0.1) !important;
            color: var(--app-borde) !important;
            border-color: var(--app-borde) !important;
        }

        [data-testid="stFileUploadDropzone"] {
            background-color: rgba(0, 0, 0, 0.2) !important;
            border: 1px dashed var(--app-borde) !important;
        }

        .stTextArea textarea {
            background-color: rgba(0, 0, 0, 0.2) !important;
            border: 1px solid var(--app-borde) !important;
            color: var(--app-texto-titulos) !important;
        }
    </style>
""", unsafe_allow_html=True)
# --- FIN DEL FORMATO VISUAL ---

# Botón de limpieza rápida en la parte superior
col1, col2 = st.columns([8, 2])
with col1:
    st.title("⚡ Centro de Herramientas ETOP")
    st.markdown("Selecciona la herramienta que necesitas usar en las pestañas de abajo.")
with col2:
    if st.button("🧹 Limpiar Todo"):
        st.rerun()

# --- CREACIÓN DE LAS DOS PESTAÑAS ---
tab1, tab2 = st.tabs(["📋 Formateador Nóminas SharePoint", "📧 Generador Correos Bienvenida"])

# =====================================================================
# PESTAÑA 1: FORMATEADOR NÓMINAS (CON FILTRO NRC)
# =====================================================================
with tab1:
    st.header("Formateador Universal de Estudiantes")
    st.write("Sube archivos o pega los datos directamente de la web.")
    
    tipo_formato = st.radio(
        "Selecciona el método de entrada:",
        [
            "📊 Sistema Diario (Sube archivo con RUT, NRC, PERIODO...)", 
            "🌐 Descarga Manual Web (Sube archivos Excel/CSV)",
            "📝 Pegar texto directamente (Copia de la web y pega aquí)"
        ]
    )

    archivos = None
    texto_pegado = ""
    fecha_filtro = None
    filtro_nrc = ""

    if "Sistema Diario" in tipo_formato:
        st.info("💡 Selecciona la fecha y los códigos NRC (opcional) a filtrar.")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            fecha_filtro = st.date_input("📅 Fecha a filtrar", format="DD/MM/YYYY")
        with col_f2:
            filtro_nrc = st.text_input("🔍 Filtrar por NRC_COD (Separar con comas)")
            
        archivos = st.file_uploader("📥 Sube los archivos", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    elif "Descarga Manual Web" in tipo_formato:
        archivos = st.file_uploader("📥 Sube los archivos Excel o CSV", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    else:
        st.info("💡 Ve a la web de tus cursos, selecciona todo el texto y pégalo abajo.")
        texto_pegado = st.text_area("📋 Pega aquí los datos de la web:", height=300)

    if archivos or texto_pegado:
        if st.button("⚡ Procesar Datos"):
            with st.spinner('Procesando datos...'):
                try:
                    filas_destino_globales = []
                    if "Sistema Diario" in tipo_formato and archivos:
                        for archivo in archivos:
                            df = pd.read_csv(archivo, sep=None, engine='python') if archivo.name.endswith('.csv') else pd.read_excel(archivo)
                            df.columns = df.columns.astype(str).str.strip().str.upper()
                            
                            if 'ROL' in df.columns:
                                df['ROL'] = df['ROL'].astype(str).str.strip().str.upper()
                                df = df[df['ROL'] == 'ESTUDIANTE'].copy()
                            
                            if 'FECHA_INICIO' in df.columns:
                                def limpiar_fecha(val):
                                    val = str(val).split(' ')[0].replace('.0', '')
                                    return f"{val[:4]}-{val[4:6]}-{val[6:]}" if len(val) == 8 and val.isdigit() else val
                                
                                df['FECHA_CLEAN'] = df['FECHA_INICIO'].apply(limpiar_fecha)
                                df['FECHA_FORMATEADA'] = pd.to_datetime(df['FECHA_CLEAN'], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df['FECHA_CLEAN'])
                                fecha_str = fecha_filtro.strftime('%Y-%m-%d') if fecha_filtro else ""
                                df = df[df['FECHA_FORMATEADA'] == fecha_str].copy()
                            
                            if 'RUT' in df.columns:
                                df['RUT'] = df['RUT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                df = df[~df['RUT'].isin(["nan", ""])]
                            
                            if all(col in df.columns for col in ['CURSO', 'MATERIA', 'NRC']):
                                df['CURSO_FMT'] = df['CURSO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.zfill(3)
                                df['NRC_FMT'] = df['NRC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                df['NRC_COD'] = df['NRC_FMT'] + "_" + df['MATERIA'].astype(str).str.strip() + df['CURSO_FMT']
                            
                            df_res = pd.DataFrame()
                            df_res['PERIODO'] = df.get('PERIODO', pd.Series(dtype=str)).astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            df_res['NRC_COD'] = df.get('NRC_COD', pd.Series(dtype=str))
                            df_res['NOMBRE_CURSO'] = df.get('NOMBRE_CURSO', pd.Series(dtype=str)).astype(str).str.strip()
                            df_res['FECHA_INICIO'] = df.get('FECHA_FORMATEADA', pd.Series(dtype=str))
                            df_res['RUT'] = df.get('RUT', pd.Series(dtype=str))
                            df_res['NOMBRE'] = df.get('NOMBRE', pd.Series(dtype=str)).astype(str).str.strip()
                            df_res['CORREO_UNAB'] = df.get('CORREO_UNAB', pd.Series(dtype=str)).astype(str).str.strip()
                            df_res['Columna7'] = df.get('CORREO_PERSONAL', pd.Series(dtype=str)).astype(str).str.strip()
                            filas_destino_globales.extend(df_res.values.tolist())
                    else:
                        # Lógica de texto pegado
                        texto_total_unido = texto_pegado if "Pegar" in tipo_formato else ""
                        if not texto_total_unido and archivos:
                            lista_lineas = []
                            for archivo in archivos:
                                df_raw = (pd.read_csv(archivo, sep=None, engine='python') if archivo.name.endswith('.csv') else pd.read_excel(archivo, header=None)).fillna("")
                                for _, row in df_raw.iterrows():
                                    lista_lineas.append(" ".join([str(x).strip() for x in row.tolist() if x]))
                            texto_total_unido = "\n".join(lista_lineas)

                        bloques_cursos = re.split(r'(?i)NOMBRE:', texto_total_unido)
                        for bloque in bloques_cursos:
                            if not bloque.strip(): continue
                            match_codigo = re.search(r'C[ÓO]DIGO DE CURSO:\s*([A-Za-z]+)\s*(\d+)\.(\d+)\.(\d+)', bloque, re.IGNORECASE)
                            match_inicio = re.search(r'INICIO:\s*(\d{2}[/-]\d{2}[/-]\d{4})', bloque, re.IGNORECASE)
                            materia = match_codigo.group(1).upper() if match_codigo else ""
                            nrc = match_codigo.group(4) if match_codigo else ""
                            nrc_cod = f"{nrc}_{materia}{match_codigo.group(2).zfill(3)}" if nrc and materia else ""
                            
                            for m in re.finditer(r'(\d{7,8}[0-9Kk])(.*?)((?:[a-zA-Z0-9._\-]+)@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})\s*(\d{4,8})\s*(Estudiante)', bloque, re.IGNORECASE | re.DOTALL):
                                filas_destino_globales.append([match_codigo.group(3) if match_codigo else "", nrc_cod, "", "", m.group(1).upper(), m.group(2).strip().upper(), m.group(3).lower(), ""])

                    df_final_maestro = pd.DataFrame(filas_destino_globales, columns=['PERIODO', 'NRC_COD', 'NOMBRE_CURSO', 'FECHA_INICIO', 'RUT', 'NOMBRE', 'CORREO_UNAB', 'Columna7']).replace("nan", "")
                    if filtro_nrc.strip():
                        df_final_maestro = df_final_maestro[df_final_maestro['NRC_COD'].str.upper().isin([x.strip().upper() for x in filtro_nrc.split(',')])]
                    
                    st.success(f"¡Éxito! {len(df_final_maestro)} estudiantes procesados.")
                    st.dataframe(df_final_maestro, use_container_width=True)
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as writer: df_final_maestro.to_excel(writer, index=False)
                    st.download_button("✅ Descargar Tabla Maestra", out.getvalue(), "Estudiantes_Formateados.xlsx")
                except Exception as e: st.error(f"Error: {e}")

# =====================================================================
# PESTAÑA 2: GENERADOR CORREOS (CORRECCIÓN DE RUT)
# =====================================================================
with tab2:
    st.header("Generador de Base para Correos de Bienvenida")
    archivo_correos = st.file_uploader("1. Sube la nómina original", type=["xlsx", "xls", "csv"], key="uploader_correos")
    
    if archivo_correos:
        try:
            df_raw = pd.read_csv(archivo_correos, sep=None, engine='python', header=None) if archivo_correos.name.endswith('.csv') else pd.read_excel(archivo_correos, header=None)
            df_raw = df_raw.fillna("")
            
            st.subheader("2. Ingresa los datos de la Inducción")
            col_a, col_b = st.columns(2)
            with col_a:
                tutor_nombre = st.text_input("NOMBRE TUTOR", key="t_nom")
                tutor_correo = st.text_input("CORREO TUTOR", key="t_cor")
                tutor_anexo  = st.text_input("ANEXO TUTOR", key="t_anx")
                programa     = st.text_input("NOMBRE DE PROGRAMA", key="t_prog")
            with col_b:
                nrc_induc = st.text_input("NRC Y NOMBRE DE CURSO INDUCC", key="t_nrc")
                curso_1   = st.text_input("NOMBRE PRIMER CURSO", key="t_cur1")
                fecha_ini = st.date_input("FECHA INICIO INDUCCION", format="DD/MM/YYYY", key="t_fini")
                fecha_fin = st.date_input("FECHA TERMINO INDUCCION", format="DD/MM/YYYY", key="t_ffin")

            if st.button("🪄 Construir Plantilla Exacta", key="btn_correos"):
                header_idx = -1
                for i in range(min(20, len(df_raw))):
                    fila_str = " ".join([str(v).upper() for v in df_raw.iloc[i].values])
                    if ('RUT' in fila_str or 'ID' in fila_str) and ('CORREO' in fila_str or 'EMAIL' in fila_str):
                        header_idx = i
                        break
                
                if header_idx != -1:
                    cols_crudas = df_raw.iloc[header_idx].astype(str).str.upper().str.strip().tolist()
                    df_alumnos = df_raw.iloc[header_idx+1:].copy()
                    vistas = set(); nuevas = []
                    for c in cols_crudas:
                        nuevas.append(c + "_REP" if c in vistas else c); vistas.add(c)
                    df_alumnos.columns = nuevas
                else:
                    df_alumnos = df_raw.copy()
                    df_alumnos.columns = [f"COL_{i}" for i in range(len(df_alumnos.columns))]

                rut_col = next((c for c in df_alumnos.columns if 'RUT' in str(c) or 'ID' in str(c)), None)
                mail_col = next((c for c in df_alumnos.columns if 'CORREO' in str(c) or 'EMAIL' in str(c)), None)
                nom_cs = [c for c in df_alumnos.columns if 'NOMBRE' in str(c) and 'TUTOR' not in str(c) and 'PROGRAMA' not in str(c)]
                ape_cs = [c for c in df_alumnos.columns if 'APELLIDO' in str(c)]

                salida_datos = []
                regex_rut_estricto = re.compile(r'^\b(\d{7,9}[-]*[0-9Kk])\b$', re.IGNORECASE)

                for _, row in df_alumnos.iterrows():
                    row_str = " ".join([str(x).strip() for x in row.dropna().values])
                    if '@' not in row_str: continue 

                    # LÓGICA REFORZADA DE RUT
                    rut_final = ""
                    if rut_col and pd.notna(row.get(rut_col)):
                        v_rut = str(row[rut_col]).replace('.0', '').strip().upper()
                        if regex_rut_estricto.match(v_rut): rut_final = v_rut
                    
                    if not rut_final:
                        rut_m = re.search(r'\b(\d{7,8}[-]*[0-9Kk])\b', row_str, re.IGNORECASE)
                        rut_final = rut_m.group(1).upper() if rut_m else ""

                    mail_final = str(row[mail_col]).strip().lower() if mail_col and pd.notna(row.get(mail_col)) else (re.search(r'[a-zA-Z0-9._\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', row_str).group(0).lower() if re.search(r'[a-zA-Z0-9._\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', row_str) else "")
                    
                    n_p = [str(row[nom_cs[0]]).strip() if nom_cs else ""]
                    if ape_cs: n_p.append(str(row[ape_cs[0]]).strip())
                    
                    if "@" in mail_final:
                        salida_datos.append({
                            "Columna1": len(salida_datos)+1, "NOMBRES": " ".join([p for p in n_p if p.lower() != 'nan']).strip().upper(),
                            "Rut o ID": rut_final, "CORREO": mail_final, "NOMBRE TUTOR": tutor_nombre, "CORREO TUTOR": tutor_correo,
                            "ANEXO TUTOR": tutor_anexo, "NOMBRE DE PROGRAMA": programa, "NRC Y NOMBRE DE CURSO INDUCC": nrc_induc,
                            "NOMBRE PRIMER CURSO": curso_1, "FECHA INICIO INDUCCION": fecha_ini.strftime("%Y-%m-%d"),
                            "FECHA TERMINO INDUCCION": fecha_fin.strftime("%Y-%m-%d"), "FECHA_INICIO_CALCULADA": fecha_ini.strftime("%d-%m-%Y"),
                            "FECHA FIN CALCULADA": fecha_fin.strftime("%d-%m-%Y"), "Columna2": ""
                        })
                st.session_state['datos_p'] = pd.DataFrame(salida_datos)

            if 'datos_p' in st.session_state:
                st.success("🎉 Plantilla generada con éxito.")
                st.dataframe(st.session_state['datos_p'], use_container_width=True)
                out_c = io.BytesIO()
                with pd.ExcelWriter(out_c, engine='xlsxwriter') as writer: st.session_state['datos_p'].to_excel(writer, index=False)
                st.download_button("📥 Descargar Excel", out_c.getvalue(), "Plantilla_Bienvenida.xlsx")
                
                if st.button("📨 Enviar directamente a Power Automate", type="primary"):
                    webhook_url = "https://default8fbed393d03b49f8be79cd5e1f590f.b2.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/116abc48d2194f8e925b7b3d57b15d85/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=1zArAdSz1df7OFOaYP6Xu8AOH3wrqTnxOADV7G5inks"
                    requests.post(webhook_url, json=st.session_state['datos_p'].to_dict(orient="records"))
                    st.success("✅ ¡Transferencia exitosa!")
                    st.balloons()
        except Exception as e: st.error(f"Error: {e}")
