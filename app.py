import streamlit as st
import pandas as pd

# 1. Configuración básica de la página (Pestaña del navegador)
st.set_page_config(
    page_title="Portal ETOP - App",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 2. Inyección de los estilos CSS de tu portal
st.markdown("""
    <style>
        /* Paleta de colores principal */
        :root {
            --app-fondo: #1B222C;        /* Azul noche */
            --app-borde: #F6C128;        /* Amarillo oro */
            --app-texto-titulos: #FFFFFF;/* Blanco */
            --app-texto-datos: #CBD5E1;  /* Gris claro */
        }

        /* Cambiar el fondo de toda la aplicación */
        .stApp {
            background-color: var(--app-fondo);
            color: var(--app-texto-datos);
        }

        /* Hacer transparente la barra superior de Streamlit */
        [data-testid="stHeader"] {
            background-color: rgba(0,0,0,0) !important;
        }

        /* Forzar el color de todos los textos y títulos */
        h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {
            color: var(--app-texto-titulos) !important;
        }

        /* Estilizar los botones con el estilo "Ghost" amarillo */
        .stButton > button {
            background-color: transparent !important;
            border: 1px solid var(--app-borde) !important;
            color: var(--app-borde) !important;
            border-radius: 8px !important;
            transition: all 0.3s ease !important;
            width: 100%;
        }

        .stButton > button:hover {
            background-color: rgba(246, 193, 40, 0.1) !important; /* Brillo amarillo */
            color: var(--app-borde) !important;
            border-color: var(--app-borde) !important;
        }

        /* Estilizar las cajas de entrada de texto (Inputs) */
        div[data-baseweb="input"] > div {
            background-color: rgba(0, 0, 0, 0.2) !important; /* Fondo más oscuro */
            border: 1px solid rgba(203, 213, 225, 0.3) !important; /* Borde gris sutil */
            color: var(--app-texto-titulos) !important;
            border-radius: 6px !important;
        }
        
        div[data-baseweb="input"] > div:focus-within {
            border-color: var(--app-borde) !important; /* Borde amarillo al hacer clic */
        }

        /* Cambiar el color de los números grandes en las métricas */
        [data-testid="stMetricValue"] {
            color: var(--app-borde) !important;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. CONSTRUCCIÓN DE LA INTERFAZ DE LA APP
# ==========================================

st.title("Gestión Tutorial")
st.write("Monitor de estudiantes, riesgos, fichas académicas e incidencias.")

st.markdown("---") # Línea divisoria

# Ejemplo de métricas superiores
col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label="Estudiantes Activos", value="1,245")
with col2:
    st.metric(label="Riesgos Detectados", value="12", delta="-3 desde ayer")
with col3:
    st.metric(label="Incidencias Abiertas", value="5")

st.markdown("<br>", unsafe_allow_html=True) # Espacio en blanco

# Ejemplo de controles de búsqueda y botones
col_busqueda, col_boton = st.columns([3, 1])

with col_busqueda:
    estudiante_buscar = st.text_input("🔍 Buscar estudiante por nombre o correo:")

with col_boton:
    st.markdown("<br>", unsafe_allow_html=True) # Alinear botón con el input
    if st.button("Buscar ficha"):
        if estudiante_buscar:
            st.success(f"Buscando a: {estudiante_buscar}...")
        else:
            st.warning("Por favor, ingresa un nombre.")

# Ejemplo de cómo se vería una tabla de datos
st.subheader("Últimos registros")
datos_ejemplo = pd.DataFrame({
    "Estudiante": ["Ana Pérez", "Carlos Gómez", "María Silva"],
    "Estado": ["Al día", "En riesgo", "Al día"],
    "Último contacto": ["05/03/2026", "01/03/2026", "04/03/2026"]
})
st.dataframe(datos_ejemplo, use_container_width=True)
