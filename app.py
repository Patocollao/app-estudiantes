import pandas as pd
import streamlit as st
import io
import re

# Configuración de la página
st.set_page_config(page_title="Formateador SharePoint", page_icon="⚡")

st.title("⚡ Formateador Universal de Estudiantes")
st.markdown("Sube tu archivo y obtén la tabla estructurada para Power Automate.")

# 1. SELECTOR DE FORMATO
tipo_formato = st.radio(
    "Selecciona el origen del archivo:",
    [
        "📊 Sistema Diario (Columnas estándar: RUT, NRC, PERIODO...)", 
        "🌐 Descarga Manual Web (Con etiquetas 'NOMBRE:', 'CÓDIGO DE CURSO:', etc.)"
    ]
)

# 2. CALENDARIO (Solo aparece si seleccionan el Sistema Diario)
fecha_filtro = None
if "Sistema Diario" in tipo_formato:
    st.info("💡 Haz clic en la caja de abajo para abrir el calendario y seleccionar la fecha.")
    fecha_filtro = st.date_input("📅 Fecha de Inicio a filtrar")

archivo = st.file_uploader("📥 Sube el archivo Excel", type=["xlsx", "xls"])

if archivo:
    if st.button("Procesar Archivo"):
        try:
            # ---------------------------------------------------------
            # LÓGICA 1: FORMATO SISTEMA DIARIO
            # ---------------------------------------------------------
            if "Sistema Diario" in tipo_formato:
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
                    
                    # Filtro exacto usando el calendario
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
                
                df_final = pd.DataFrame()
                df_final['PERIODO'] = df.get('PERIODO', pd.Series(dtype=str)).astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_final['NRC_COD'] = df.get('NRC_COD', pd.Series(dtype=str))
                df_final['NOMBRE_CURSO'] = df.get('NOMBRE_CURSO', pd.Series(dtype=str)).astype(str).str.strip()
                df_final['FECHA_INICIO'] = df.get('FECHA_FORMATEADA', pd.Series(dtype=str))
                df_final['RUT'] = df.get('RUT', pd.Series(dtype=str))
                df_final['NOMBRE'] = df.get('NOMBRE', pd.Series(dtype=str)).astype(str).str.strip()
                df_final['CORREO_UNAB'] = df.get('CORREO_UNAB', pd.Series(dtype=str)).astype(str).str.strip()
                df_final['Columna7'] = df.get('CORREO_PERSONAL', pd.Series(dtype=str)).astype(str).str.strip()

            # ---------------------------------------------------------
            # LÓGICA 2: FORMATO MANUAL WEB (Con etiquetas)
            # ---------------------------------------------------------
            else:
                df_raw = pd.read_excel(archivo, header=None).fillna("")
                filas_destino = []
                
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
                        
                    if celda0.lower() == "rut":
                        en_tabla = True
                        fila_lower = [x.lower() for x in fila]
                        idx_rut = fila_lower.index("rut") if "rut" in fila_lower else -1
                        idx_nombre = fila_lower.index("nombre completo") if "nombre completo" in fila_lower else -1
                        idx_email = fila_lower.index("email") if "email" in fila_lower else -1
                        idx_rol = fila_lower.index("rol") if "rol" in fila_lower else -1
                        continue
                        
                    if not en_tabla: continue
                    if idx_rol != -1 and fila[idx_rol].lower() != "estudiante": continue
                    if idx_rut != -1:
                        rut_raw = fila[idx_rut].replace(".0", "")
                        if not rut_raw: continue
                        
                        clave_dup = f"{nrc}_{rut_raw}"
                        if clave_dup in rut_por_nrc: continue
                        rut_por_nrc.add(clave_dup)
                        
                        nrc_cod = f"{nrc}_{materia}{curso}"
                        nombre_estudiante = fila[idx_nombre] if idx_nombre != -1 else ""
                        email_estudiante = fila[idx_email] if idx_email != -1 else ""
                        
                        filas_destino.append({
                            "PERIODO": periodo,
                            "NRC_COD": nrc_cod,
                            "NOMBRE_CURSO": nombre_curso,
                            "FECHA_INICIO": fecha_inicio,
                            "RUT": rut_raw,
                            "NOMBRE": nombre_estudiante,
                            "CORREO_UNAB": email_estudiante,
                            "Columna7": ""
                        })
                        
                df_final = pd.DataFrame(filas_destino, columns=[
                    'PERIODO', 'NRC_COD', 'NOMBRE_CURSO', 'FECHA_INICIO', 
                    'RUT', 'NOMBRE', 'CORREO_UNAB', 'Columna7'
                ])

            # ---------------------------------------------------------
            # EXPORTACIÓN (Común para ambas lógicas)
            # ---------------------------------------------------------
            df_final = df_final.replace("nan", "")
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df_final.to_excel(writer, sheet_name='Resultado', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Resultado']
            
            max_row, max_col = df_final.shape
            if max_row > 0:
                column_settings = [{'header': column} for column in df_final.columns]
                worksheet.add_table(0, 0, max_row, max_col - 1, {
                    'columns': column_settings,
                    'name': 'TablaEstudiantes',
                    'style': 'Table Style Medium 2' 
                })
                
            worksheet.autofit()
            writer.close()
            output.seek(0)
            
            # Mensaje de éxito dinámico
            origen_usado = "Sistema Diario" if "Sistema Diario" in tipo_formato else "Descarga Manual Web"
            st.success(f"¡Éxito! Se encontraron y procesaron {max_row} estudiantes usando la lógica de {origen_usado}.")
            
            # Nombre del archivo dinámico
            nombre_descarga = f"Estudiantes_{fecha_filtro}.xlsx" if "Sistema" in tipo_formato else "Estudiantes_Web.xlsx"
            
            st.download_button(
                label="✅ Descargar Tabla Limpia para SharePoint",
                data=output,
                file_name=nombre_descarga,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Hubo un error al procesar el archivo: {e}")
