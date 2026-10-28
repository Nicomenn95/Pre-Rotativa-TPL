import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador Pre Rotativa", layout="wide", page_icon="🚌")

st.title("🚌 Generador de Reporte Pre Rotativa")
st.markdown("Sube el reporte de ocupación descargado de Terra. El sistema cruzará los datos con los tiempos de viaje y generará el formato exacto para la Pre Rotativa.")

# --- CARGADOR DE ARCHIVOS ---
archivo_terra = st.file_uploader("📥 Subir Reporte de Terra (Excel)", type=["xlsx", "xls"])

# --- FUNCIONES ---
@st.cache_data
def cargar_tiempos():
    """Carga el archivo maestro de tiempos que debe estar subido en el repositorio de GitHub."""
    try:
        return pd.read_excel("Tiempos de llegada a destino TPL.xlsx")
    except Exception as e:
        st.error(f"❌ No se encontró el archivo 'Tiempos de llegada a destino TPL.xlsx'. Por favor, asegúrate de haberlo subido a GitHub junto al código. Detalle: {e}")
        return pd.DataFrame()

def calcular_hora_llegada(hora_salida, duracion):
    """Suma la hora de salida y el tiempo de viaje para obtener la hora de llegada."""
    if pd.isna(hora_salida) or pd.isna(duracion):
        return None
    try:
        # Convertir la hora de salida (texto) a un objeto de tiempo
        h_s = pd.to_datetime(str(hora_salida), format='%H:%M').time()
        
        # Convertir la duración a un timedelta (lapso de tiempo)
        if isinstance(duracion, str):
            h, m, s = map(int, duracion.split(':'))
            td = timedelta(hours=h, minutes=m, seconds=s)
        else:
            td = timedelta(hours=duracion.hour, minutes=duracion.minute, seconds=duracion.second)
            
        # Sumar los tiempos creando una fecha base
        dt_salida = datetime.combine(datetime.today(), h_s)
        dt_llegada = dt_salida + td
        return dt_llegada.strftime('%H:%M')
    except Exception:
        return None

# --- PROCESAMIENTO PRINCIPAL ---
if archivo_terra:
    try:
        # 1. Leer el reporte de Terra (saltando la primera fila que es un encabezado inútil)
        df_terra = pd.read_excel(archivo_terra, skiprows=1)
        
        # 2. Cargar tabla de tiempos desde GitHub
        df_tiempos = cargar_tiempos()
        
        if not df_tiempos.empty:
            with st.spinner("Procesando y calculando tiempos de llegada..."):
                
                # 3. Cruzar datos con la tabla de tiempos usando la columna 'Ruta'
                df_merged = pd.merge(df_terra, df_tiempos, on='Ruta', how='left')
                
                # 4. Calcular la Hora de Llegada
                df_merged['Hora de llegada'] = df_merged.apply(
                    lambda row: calcular_hora_llegada(row['Hora salida'], row.get('TIEMPO DE LLEGADA')), axis=1
                )
                
                # 5. Separar la Ocupación Relativa
                # Toma algo como "2/42 (4.76%)" y lo corta donde hay un espacio.
                temp = df_merged['Ocupación Relativa'].astype(str).str.split(' ', n=1, expand=True)
                df_merged['Ocupacion'] = temp[0] # Guarda el "2/42"
                
                if temp.shape[1] > 1:
                    # Limpia los paréntesis para dejar solo "4.76%"
                    df_merged['% ocupacion'] = temp[1].str.replace(r'\(|\)', '', regex=True)
                else:
                    df_merged['% ocupacion'] = ""
                    
                # 6. Seleccionar el orden exacto de las columnas que pediste
                columnas_finales = [
                    'Fecha salida', 'Hora salida', 'Hora de llegada', 
                    'Origen (ciudad)', 'Destino (ciudad)', 'Ocupacion', '% ocupacion'
                ]
                
                df_final = df_merged[columnas_finales]
                
                st.success("✅ ¡Reporte procesado correctamente!")
                
                # 7. Mostrar Vista Previa
                st.subheader("Vista Previa del Reporte Resultante")
                st.dataframe(df_final, use_container_width=True)
                
                # 8. Crear Excel para descarga
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, sheet_name='Pre Rotativa', index=False)
                
                # Botón de Descarga
                fecha_hoy = datetime.now().strftime("%d%m%Y")
                st.download_button(
                    label="📥 Descargar Reporte Pre Rotativa (Excel)",
                    data=output.getvalue(),
                    file_name=f"Reporte_Pre_Rotativa_{fecha_hoy}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
    except Exception as e:
        st.error(f"❌ Ocurrió un error al leer el archivo de Terra: {e}")
