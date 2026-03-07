import pandas as pd
import streamlit as st
import io
import re

st.set_page_config(page_title="Formateador SharePoint", page_icon="⚡")

st.title("⚡ Formateador Universal de Estudiantes")
st.markdown("Sube archivos o pega los datos directamente de la web para obtener tu tabla lista para Power Automate.")

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
                        df['FECHA_FORMATEADA'] = pd.to_datetime(df['FECHA_CLEAN'], errors='ignore').astype(str).str.split(' ').str[0]
                        df = df[df['FECHA_FORMATEADA'] == str(fecha_filtro)].copy()
                    
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

            # --- LÓGICA B: EL ESCÁNER "LÁSER" (Web/Texto Pegado) ---
            else:
                lista_textos_a_procesar = []
                
                if "Pegar texto" in tipo_formato and texto_pegado:
                    lista_textos_a_procesar = texto_pegado.split('\n')
                elif "Descarga Manual Web" in tipo_formato and archivos:
                    for archivo in archivos:
                        if archivo.name.endswith('.csv'):
                            df_raw = pd.read_csv(archivo, sep=None, engine='python').fillna("")
                        else:
                            df_raw = pd.read_excel(archivo, header=None).fillna("")
                        for index, row in df_raw.iterrows():
                            fila = [str(x).strip() for x in row.tolist()]
                            texto_fila = " ".join([x for x in fila if x])
                            lista_textos_a_procesar.append(texto_fila)

                nombre_curso, periodo, nrc, materia, curso, fecha_inicio = "", "", "", "", "", ""
                rut_por_nrc = set()
                
                # LA MAGIA: Este patrón extrae correos correctamente y permite buscar múltiples estudiantes en 1 sola línea infinita
                patron_estudiante = re.compile(
                    r'(\d{7,8}[0-9Kk])\s*(.*?)\s*([a-zA-Z0-9._\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})\s*\d*\s*(Estudiante|Docente[^\d]*)?', 
                    re.IGNORECASE
                )

                for linea in lista_textos_a_procesar:
                    linea_str = linea.strip()
                    if not linea_str: continue
                    
                    linea_upper = linea_str.upper()
                    
                    if linea_upper.startswith("NOMBRE:"):
                        nombre_curso = linea_str[7:].strip()
                        continue
                        
                    if linea_upper.startswith("CÓDIGO DE CURSO:"):
                        codigo = linea_str[16:].strip()
                        partes = codigo.split(".")
                        if len(partes) >= 3:
                            materia_completa = partes[0].strip()
                            periodo = partes[1].strip()
                            nrc = partes[2].strip()
                            
                            match = re.match(r'^([A-Za-z]+)(\d+)$', materia_completa)
                            if match:
                                materia = match.group(1).upper()
                                curso = match.group(2).zfill(3)
                            else:
                                materia = materia_completa.upper()
                                curso = ""
                        continue
                        
                    if linea_upper.startswith("INICIO:"):
                        fecha_raw = linea_str[7:].strip()
                        partes_fecha = fecha_raw.split("/")
                        if len(partes_fecha) == 3:
                            fecha_inicio = f"{partes_fecha[2]}-{partes_fecha[1].zfill(2)}-{partes_fecha[0].zfill(2)}"
                        else:
                            fecha_inicio = fecha_raw
                        continue

                    # EL CAMBIO CLAVE: finditer() escanea toda la línea y atrapa a todos los estudiantes sin importar qué tan pegados estén
                    for match_est in patron_estudiante.finditer(linea_str):
                        rut_raw = match_est.group(1).upper()
                        nombre_est = match_est.group(2).strip()
                        email_est = match_est.group(3).strip()
                        rol_est = match_est.group(4)

                        if rol_est and "ESTUDIANTE" not in rol_est.upper():
                            continue
                            
                        clave_dup = f"{nrc}_{rut_raw}"
                        if clave_dup in rut_por_nrc: continue
                        rut_por_nrc.add(clave_dup)
                        
                        nrc_cod = f"{nrc}_{materia}{curso}"
                        
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
            
            st.download_button(
                label="✅ Descargar Tabla Maestra",
                data=output,
                file_name="Estudiantes_Formateados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Hubo un error procesando los datos: {e}")
