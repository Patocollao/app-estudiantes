import pandas as pd
import streamlit as st
import io
import re

st.set_page_config(page_title="Formateador SharePoint", page_icon="⚡", layout="wide")

# --- INICIO DEL FORMATO VISUAL (ESTILOS DEL PORTAL) ---
st.markdown("""
    <style>
        /* Variables de color de tu portal */
        :root {
            --app-fondo: #1B222C;        /* Azul Noche */
            --app-borde: #F6C128;        /* Amarillo Oro */
            --app-texto-titulos: #FFFFFF;/* Blanco */
            --app-texto-datos: #CBD5E1;  /* Gris Claro */
        }

        /* Fondo de la app */
        .stApp {
            background-color: var(--app-fondo);
            color: var(--app-texto-datos);
        }
        
        /* Cabecera transparente */
        [data-testid="stHeader"] {
            background-color: rgba(0,0,0,0) !important;
        }

        /* Títulos y textos generales */
        h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {
            color: var(--app-texto-titulos) !important;
        }

        /* Estilo para TODOS los botones (Limpiar, Procesar, Descargar) */
        .stButton > button, .stDownloadButton > button {
            background-color: transparent !important;
            border: 1px solid var(--app-borde) !important;
            color: var(--app-borde) !important;
            border-radius: 8px !important;
            transition: all 0.3s ease !important;
        }

        /* Efecto al pasar el mouse por los botones */
        .stButton > button:hover, .stDownloadButton > button:hover {
            background-color: rgba(246, 193, 40, 0.1) !important;
            color: var(--app-borde) !important;
            border-color: var(--app-borde) !important;
        }

        /* Estilo para las cajas de arrastrar archivos (Uploader) */
        [data-testid="stFileUploadDropzone"] {
            background-color: rgba(0, 0, 0, 0.2) !important;
            border: 1px dashed var(--app-borde) !important;
        }

        /* Estilo para la caja de texto (Pegar texto web) */
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
# PESTAÑA 1: TU CÓDIGO ORIGINAL (NÓMINAS)
# =====================================================================
with tab1:
    st.header("Formateador Universal de Estudiantes")
    st.write("Sube archivos o pega los datos directamente de la web para obtener tu tabla lista para Power Automate.")
    
    # SELECTOR DE TRES OPCIONES
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

    if "Sistema Diario" in tipo_formato:
        st.info("💡 Selecciona la fecha a filtrar.")
        fecha_filtro = st.date_input("📅 Fecha a filtrar")
        archivos = st.file_uploader("📥 Sube los archivos", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

    elif "Descarga Manual Web" in tipo_formato:
        archivos = st.file_uploader("📥 Sube los archivos Excel o CSV", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

    else:
        st.info("💡 Ve a la web de tus cursos, selecciona todo el texto, presiona Ctrl+C y pégalo abajo con Ctrl+V. Puedes pegar varios cursos uno debajo del otro.")
        texto_pegado = st.text_area("📋 Pega aquí los datos de la web:", height=300)

    if archivos or texto_pegado:
        if st.button("⚡ Procesar Datos"):
            with st.spinner('Procesando datos...'):
                try:
                    filas_destino_globales = []
                    
                    # --- LÓGICA A: SISTEMA DIARIO ---
                    if "Sistema Diario" in tipo_formato and archivos:
                        for archivo in archivos:
                            if archivo.name.endswith('.csv'):
                                df = pd.read_csv(archivo, sep=None, engine='python')
                            else:
                                df = pd.read_excel(archivo)
                                
                            df.columns = df.columns.astype(str).str.strip().str.upper()
                            
                            if 'ROL' in df.columns:
                                df['ROL'] = df['ROL'].astype(str).str.strip().str.upper()
                                df = df[df['ROL'] == 'ESTUDIANTE'].copy()
                            
                            def limpiar_fecha(val):
                                val = str(val).split(' ')[0].replace('.0', '')
                                if len(val) == 8 and val.isdigit(): 
                                    return f"{val[:4]}-{val[4:6]}-{val[6:]}"
                                return val
                                
                            if 'FECHA_INICIO' in df.columns:
                                df['FECHA_CLEAN'] = df['FECHA_INICIO'].apply(limpiar_fecha)
                                # Formato YYYY-MM-DD que exige Power Automate
                                df['FECHA_FORMATEADA'] = pd.to_datetime(df['FECHA_CLEAN'], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df['FECHA_CLEAN'])
                                fecha_str = fecha_filtro.strftime('%Y-%m-%d') if fecha_filtro else ""
                                df = df[df['FECHA_FORMATEADA'] == fecha_str].copy()
                            
                            if 'RUT' in df.columns:
                                df['RUT'] = df['RUT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                df = df[df['RUT'] != "nan"] 
                                df = df[df['RUT'] != ""] 
                            if 'NRC' in df.columns and 'RUT' in df.columns:
                                df = df.drop_duplicates(subset=['NRC', 'RUT'])
                            
                            if all(col in df.columns for col in ['CURSO', 'MATERIA', 'NRC']):
                                df['CURSO_FMT'] = df['CURSO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.zfill(3)
                                df['MATERIA_FMT'] = df['MATERIA'].astype(str).str.strip()
                                df['NRC_FMT'] = df['NRC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                df['NRC_COD'] = df['NRC_FMT'] + "_" + df['MATERIA_FMT'] + df['CURSO_FMT']
                            
                            df_final_temp = pd.DataFrame()
                            df_final_temp['PERIODO'] = df.get('PERIODO', pd.Series(dtype=str)).astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            df_final_temp['NRC_COD'] = df.get('NRC_COD', pd.Series(dtype=str))
                            df_final_temp['NOMBRE_CURSO'] = df.get('NOMBRE_CURSO', pd.Series(dtype=str)).astype(str).str.strip()
                            df_final_temp['FECHA_INICIO'] = df.get('FECHA_FORMATEADA', pd.Series(dtype=str))
                            df_final_temp['RUT'] = df.get('RUT', pd.Series(dtype=str))
                            df_final_temp['NOMBRE'] = df.get('NOMBRE', pd.Series(dtype=str)).astype(str).str.strip()
                            df_final_temp['CORREO_UNAB'] = df.get('CORREO_UNAB', pd.Series(dtype=str)).astype(str).str.strip()
                            df_final_temp['Columna7'] = df.get('CORREO_PERSONAL', pd.Series(dtype=str)).astype(str).str.strip()
                            
                            filas_destino_globales.extend(df_final_temp.values.tolist())

                    # --- LÓGICA B: LA TRITURADORA DEFINITIVA (Web/Texto Pegado) ---
                    else:
                        texto_total_unido = ""
                        
                        if "Pegar texto" in tipo_formato and texto_pegado:
                            texto_total_unido = texto_pegado
                        elif "Descarga Manual Web" in tipo_formato and archivos:
                            lista_lineas = []
                            for archivo in archivos:
                                if archivo.name.endswith('.csv'):
                                    df_raw = pd.read_csv(archivo, sep=None, engine='python').fillna("")
                                else:
                                    df_raw = pd.read_excel(archivo, header=None).fillna("")
                                for index, row in df_raw.iterrows():
                                    fila = [str(x).strip() for x in row.tolist()]
                                    lista_lineas.append(" ".join([x for x in fila if x]))
                            texto_total_unido = "\n".join(lista_lineas)

                        bloques_cursos = re.split(r'(?i)NOMBRE:', texto_total_unido)
                        
                        patron_estudiante = re.compile(
                            r'(\d{7,8}[0-9Kk])(.*?)((?:[a-zA-Z0-9._\-]+)@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})\s*(\d{4,8})\s*(Estudiante|Docente[a-zA-Z\s\(\)]*)', 
                            re.IGNORECASE | re.DOTALL
                        )

                        for bloque in bloques_cursos:
                            if not bloque.strip(): continue
                            
                            match_curso = re.search(r'^\s*(.*?)(?=\s*C[ÓO]DIGO DE CURSO:)', bloque, re.IGNORECASE | re.DOTALL)
                            match_codigo = re.search(r'C[ÓO]DIGO DE CURSO:\s*([A-Za-z]+)\s*(\d+)\.(\d+)\.(\d+)', bloque, re.IGNORECASE)
                            match_inicio = re.search(r'INICIO:\s*(\d{2}[/-]\d{2}[/-]\d{4})', bloque, re.IGNORECASE)
                            
                            nombre_curso = match_curso.group(1).strip() if match_curso else ""
                            materia = match_codigo.group(1).upper() if match_codigo else ""
                            curso = match_codigo.group(2).zfill(3) if match_codigo else ""
                            periodo = match_codigo.group(3) if match_codigo else ""
                            nrc = match_codigo.group(4) if match_codigo else ""
                            
                            # 🗓️ CAMBIO APLICADO: Formato YYYY-MM-DD para satisfacer a Power Automate
                            fecha_inicio = ""
                            if match_inicio:
                                f_raw = match_inicio.group(1).replace('/', '-')
                                partes = f_raw.split('-')
                                if len(partes) == 3:
                                    # partes[0] = DD, partes[1] = MM, partes[2] = YYYY
                                    fecha_inicio = f"{partes[2]}-{partes[1].zfill(2)}-{partes[0].zfill(2)}"
                            
                            nrc_cod = f"{nrc}_{materia}{curso}" if nrc and materia else ""
                            rut_por_nrc = set()
                            
                            for m in patron_estudiante.finditer(bloque):
                                rut_raw = m.group(1).upper()
                                crudo_nombre = m.group(2)
                                crudo_email = m.group(3)
                                pidm = m.group(4)
                                rol_est = m.group(5).strip()
                                
                                if rol_est and "ESTUDIANTE" not in rol_est.upper():
                                    continue
                                    
                                match_separacion = re.match(r'^([A-ZÁÉÍÓÚÑÜ]*)([a-z0-9._\-]+@.*)$', crudo_email)
                                if match_separacion:
                                    apellido = match_separacion.group(1)
                                    email_est = match_separacion.group(2)
                                    nombre_est = " ".join((crudo_nombre + apellido).split())
                                else:
                                    nombre_est = " ".join(crudo_nombre.split())
                                    email_est = crudo_email.strip()
                                    
                                clave_dup = f"{nrc}_{rut_raw}"
                                if clave_dup in rut_por_nrc: continue
                                rut_por_nrc.add(clave_dup)
                                
                                filas_destino_globales.append([
                                    periodo, nrc_cod, nombre_curso, fecha_inicio, 
                                    rut_raw, nombre_est, email_est, ""
                                ])

                    # --- EXPORTACIÓN MAESTRA ---
                    df_final_maestro = pd.DataFrame(filas_destino_globales, columns=[
                        'PERIODO', 'NRC_COD', 'NOMBRE_CURSO', 'FECHA_INICIO', 
                        'RUT', 'NOMBRE', 'CORREO_UNAB', 'Columna7'
                    ])
                    df_final_maestro = df_final_maestro.replace("nan", "")
                    
                    if df_final_maestro.empty:
                        st.warning("⚠️ El proceso terminó, pero no se extrajo ningún estudiante. Asegúrate de estar copiando desde la palabra 'NOMBRE:' hasta el final de la tabla.")
                        st.stop()
                    
                    output = io.BytesIO()
                    writer = pd.ExcelWriter(output, engine='xlsxwriter')
                    df_final_maestro.to_excel(writer, sheet_name='Resultado', index=False)
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Resultado']
                    
                    max_row, max_col = df_final_maestro.shape
                    if max_row > 0:
                        column_settings = [{'header': column} for column in df_final_maestro.columns]
                        worksheet.add_table(0, 0, max_row, max_col - 1, {
                            'columns': column_settings,
                            'name': 'TablaEstudiantes',
                            'style': 'Table Style Medium 2' 
                        })
                        
                    worksheet.autofit()
                    writer.close()
                    output.seek(0)
                    
                    st.success(f"¡Éxito total! Se procesaron y consolidaron {max_row} estudiantes.")
                    
                    # --- VISTA PREVIA ---
                    st.markdown("### 👀 Vista previa de los datos")
                    st.dataframe(df_final_maestro, use_container_width=True)
                    
                    # --- DESCARGA CON NOMBRE ÚNICO ---
                    marca_tiempo = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                    nombre_unico = f"Estudiantes_Formateados_{marca_tiempo}.xlsx"
                    
                    st.download_button(
                        label="✅ Descargar Tabla Maestra",
                        data=output,
                        file_name=nombre_unico,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"Hubo un error procesando los datos: {e}")


# =====================================================================
# PESTAÑA 2: NUEVA HERRAMIENTA (CORREOS DE BIENVENIDA)
# =====================================================================
with tab2:
    st.header("Generador de Base para Correos de Bienvenida")
    st.write("Sube la lista de estudiantes y llena los datos del curso. El sistema copiará tu información para todos los alumnos automáticamente.")
    
    # 1. SUBIR EL ARCHIVO ORIGINAL
    archivo_correos = st.file_uploader("1. Sube la nómina original de estudiantes (Excel o CSV)", type=["xlsx", "xls", "csv"], key="uploader_correos")
    
    if archivo_correos:
        try:
            # Leemos el archivo
            if archivo_correos.name.endswith('.csv'):
                df_alumnos = pd.read_csv(archivo_correos, sep=None, engine='python')
            else:
                df_alumnos = pd.read_excel(archivo_correos)
                
            st.success(f"✅ Archivo cargado correctamente con {len(df_alumnos)} estudiantes.")
            
            # 2. FORMULARIO PARA LOS DATOS ESTÁTICOS
            st.subheader("2. Ingresa los datos de la Inducción")
            st.write("Llena estos campos UNA SOLA VEZ. Se aplicarán a todos los estudiantes.")
            
            # Usamos columnas para que el formulario se vea ordenado y no ocupe tanta pantalla
            col_a, col_b = st.columns(2)
            
            with col_a:
                tutor_nombre = st.text_input("NOMBRE TUTOR", key="t_nom")
                tutor_correo = st.text_input("CORREO TUTOR", key="t_cor")
                tutor_anexo  = st.text_input("ANEXO TUTOR", key="t_anx")
                programa     = st.text_input("NOMBRE DE PROGRAMA", key="t_prog")
                
            with col_b:
                nrc_induc = st.text_input("NRC Y NOMBRE DE CURSO INDUCC", key="t_nrc")
                curso_1   = st.text_input("NOMBRE PRIMER CURSO", key="t_cur1")
                # Usamos selectores de fecha nativos de Streamlit
                fecha_ini = st.date_input("FECHA INICIO INDUCCION", key="t_fini")
                fecha_fin = st.date_input("FECHA TERMINO INDUCCION", key="t_ffin")
            
            # 3. BOTÓN DE PROCESAMIENTO
            if st.button("🪄 Generar Excel para Power Automate", key="btn_correos"):
                
                # Creamos una copia para no alterar el original
                df_final_correos = df_alumnos.copy()
                
                # Asignamos el valor a la columna y Pandas lo repite en TODAS las filas
                df_final_correos["NOMBRE TUTOR"] = tutor_nombre
                df_final_correos["CORREO TUTOR"] = tutor_correo
                df_final_correos["ANEXO TUTOR"] = tutor_anexo
                df_final_correos["NOMBRE DE PROGRAMA"] = programa
                df_final_correos["NRC Y NOMBRE DE CURSO INDUCC"] = nrc_induc
                df_final_correos["NOMBRE PRIMER CURSO"] = curso_1
                
                # Formateamos las fechas a texto (DD-MM-YYYY)
                df_final_correos["FECHA INICIO INDUCCION"] = fecha_ini.strftime("%d-%m-%Y")
                df_final_correos["FECHA TERMINO INDUCCION"] = fecha_fin.strftime("%d-%m-%Y")
                
                # Vista previa opcional para el tutor
                st.write("👀 Vista previa de cómo quedó tu archivo:")
                st.dataframe(df_final_correos.head(3), use_container_width=True) # Muestra solo las 3 primeras filas
                
                # 4. PREPARAR DESCARGA
                output_correos = io.BytesIO()
                with pd.ExcelWriter(output_correos, engine='xlsxwriter') as writer:
                    df_final_correos.to_excel(writer, index=False, sheet_name='Base_Correos')
                output_correos.seek(0)
                
                st.success("🎉 ¡Archivo generado con éxito! Haz clic abajo para descargarlo.")
                st.download_button(
                    label="📥 Descargar Base Formateada",
                    data=output_correos,
                    file_name="Base_Correos_Bienvenida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_correos"
                )
                
        except Exception as e:
            st.error(f"Hubo un error al procesar el archivo de correos: {e}")
