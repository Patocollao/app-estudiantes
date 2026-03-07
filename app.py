import pandas as pd
import streamlit as st
import io
import re

st.set_page_config(page_title="Formateador SharePoint", page_icon="⚡")

st.title("⚡ Formateador Universal de Estudiantes")
st.markdown("Sube archivos o pega los datos directamente de la web para obtener tu tabla lista para Power Automate.")

# 1. SELECTOR DE TRES OPCIONES
tipo_formato = st.radio(
    "Selecciona el método de entrada:",
    [
        "📊 Sistema Diario (Sube archivo con RUT, NRC, PERIODO...)", 
        "🌐 Descarga Manual Web (Sube archivos Excel/CSV de la web)",
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

# Botón de procesamiento dinámico
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

            # --- LÓGICA B: FORMATO MANUAL WEB (Por Archivo o Texto Pegado) ---
            else:
                lista_dataframes_raw = []
                
                if "Pegar texto" in tipo_formato and texto_pegado:
                    # Convertir el texto pegado en una cuadrícula (como si fuera un Excel)
                    filas_texto = [linea.split('\t') for linea in texto_pegado.strip().split('\n')]
                    lista_dataframes_raw.append(pd.DataFrame(filas_texto).fillna(""))
                    
                elif "Descarga Manual Web" in tipo_formato and archivos:
                    for archivo in archivos:
                        if archivo.name.endswith('.csv'):
                            lista_dataframes_raw.append(pd.read_csv(archivo, sep=None, engine='python').fillna(""))
                        else:
                            lista_dataframes_raw.append(pd.read_excel(archivo, header=None).fillna(""))

                # Procesar todos los bloques de datos (ya sea que vinieron de archivos o del cuadro de texto)
                for df_raw in lista_dataframes_raw:
                    nombre_curso, periodo, nrc, materia, curso, fecha_inicio = "", "", "", "", "", ""
                    en_tabla = False
                    idx_rut, idx_nombre, idx_email, idx_rol = -1, -1, -1, -1
                    rut_por_nrc = set()
                    
                    for index, row in df_raw.iterrows():
                        fila = [str(x).strip() for x in row.tolist()]
                        if all(x == "" for x in fila): continue
                        
                        celda0 = fila[0].upper()
                        
                        if celda0.startswith("NOMBRE:"):
                            nombre_curso = fila[0][7:].strip()
                            en_tabla = False
                            continue
                            
                        if celda0.startswith("CÓDIGO DE CURSO:"):
                            codigo = fila[0][16:].strip()
                            partes = codigo.split(".")
                            if len(partes) >= 3:
                                materia_completa = partes[0].strip()
                                periodo = partes[1].strip()
                                nrc = partes[2].strip()
                                
                                match = re.match(r'^([A-Za-z]+)(\d+)$', materia_completa)
                                if match:
                                    materia = match.group(1)
                                    curso = match.group(2).zfill(3)
                                else:
                                    materia = materia_completa
                                    curso = ""
                            continue
                            
                        if celda0.startswith("INICIO:"):
                            fecha_raw = fila[0][7:].strip()
                            partes_fecha = fecha_raw.split("/")
                            if len(partes_fecha) == 3:
                                fecha_inicio = f"{partes_fecha[2]}-{partes_fecha[1].zfill(2)}-{partes_fecha[0].zfill(2)}"
                            else:
                                fecha_inicio = fecha_raw
                            continue
                            
                        if celda0.lower() == "rut" or "rut" in [x.lower() for x in fila]:
                            en_tabla = True
                            fila_lower = [x.lower() for x in fila]
                            idx_rut = fila_lower.index("rut") if "rut" in fila_lower else -1
                            
                            if "nombre completo" in fila_lower: idx_nombre = fila_lower.index("nombre completo")
                            elif "nombres" in fila_lower: idx_nombre = fila_lower.index("nombres")
                            elif "nombre" in fila_lower: idx_nombre = fila_lower.index("nombre")
                            
                            idx_email = fila_lower.index("email") if "email" in fila_lower else -1
                            idx_rol = fila_lower.index("rol") if "rol" in fila_lower else -1
                            continue
                            
                        if not en_tabla: continue
                        if idx_rol != -1 and fila[idx_rol].lower() != "estudiante": continue
                        if idx_rut != -1 and idx_rut < len(fila):
                            rut_raw = fila[idx_rut].replace(".0", "")
                            if not rut_raw: continue
                            
                            clave_dup = f"{nrc}_{rut_raw}"
                            if clave_dup in rut_por_nrc: continue
                            rut_por_nrc.add(clave_dup)
                            
                            nrc_cod = f"{nrc}_{materia}{curso}"
                            nombre_estudiante = fila[idx_nombre] if idx_nombre != -1 and idx_nombre < len(fila) else ""
                            
                            if idx_nombre != -1 and "apellidos" in [str(x).lower() for x in df_raw.iloc[0].tolist()]:
                                idx_apellido = [str(x).lower() for x in df_raw.iloc[0].tolist()].index("apellidos")
                                if idx_apellido < len(fila):
                                    nombre_estudiante = f"{fila[idx_nombre]} {fila[idx_apellido]}"

                            email_estudiante = fila[idx_email] if idx_email != -1 and idx_email < len(fila) else ""
                            
                            filas_destino_globales.append([
                                periodo, nrc_cod, nombre_curso, fecha_inicio, 
                                rut_raw, nombre_estudiante, email_estudiante, ""
                            ])

            # --- EXPORTACIÓN MAESTRA ---
            df_final_maestro = pd.DataFrame(filas_destino_globales, columns=[
                'PERIODO', 'NRC_COD', 'NOMBRE_CURSO', 'FECHA_INICIO', 
                'RUT', 'NOMBRE', 'CORREO_UNAB', 'Columna7'
            ])
            df_final_maestro = df_final_maestro.replace("nan", "")
            
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
