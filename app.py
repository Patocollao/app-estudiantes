import pandas as pd
import streamlit as st
import io

st.set_page_config(page_title="Formateador SharePoint", page_icon="⚡")

st.title("⚡ Formateador de Estudiantes para SharePoint")
st.markdown("1. Escribe la fecha que deseas procesar.\n2. Sube el archivo base.\nEl sistema te entregará la tabla limpia lista para Power Automate.")

fecha_filtro = st.text_input("📅 Fecha a filtrar (Formato YYYY-MM-DD)", placeholder="Ejemplo: 2026-03-16")
archivo = st.file_uploader("📥 Sube el archivo original", type=["xlsx", "xls"])

if archivo and fecha_filtro:
    if st.button("Procesar Archivo"):
        try:
            df = pd.read_excel(archivo)
            df.columns = df.columns.astype(str).str.strip().str.upper()
            
            if 'ROL' in df.columns:
                df['ROL'] = df['ROL'].astype(str).str.strip().str.upper()
                df = df[df['ROL'] == 'ESTUDIANTE'].copy()
            
            def limpiar_fecha(val):
                val = str(val).split(' ')[0] 
                val = val.replace('.0', '')
                if len(val) == 8 and val.isdigit(): 
                    return f"{val[:4]}-{val[4:6]}-{val[6:]}"
                return val
                
            if 'FECHA_INICIO' in df.columns:
                df['FECHA_CLEAN'] = df['FECHA_INICIO'].apply(limpiar_fecha)
                df['FECHA_FORMATEADA'] = pd.to_datetime(df['FECHA_CLEAN'], errors='ignore').astype(str).str.split(' ').str[0]
                df = df[df['FECHA_FORMATEADA'] == fecha_filtro.strip()].copy()
            
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
            
            st.success(f"¡Éxito! Se encontraron y procesaron {max_row} estudiantes.")
            
            st.download_button(
                label="✅ Descargar Tabla Limpia para SharePoint",
                data=output,
                file_name=f"Estudiantes_{fecha_filtro}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Hubo un error al procesar el archivo: {e}")
