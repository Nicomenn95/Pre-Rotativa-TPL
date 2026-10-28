import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador Pre Rotativa", layout="wide", page_icon="🚌")

st.title("🚌 Generador de Reporte Pre Rotativa")
st.markdown("Sube el reporte de ocupación descargado de Terra. El sistema cruzará los datos con los tiempos de viaje integrados y generará el formato exacto para la Pre Rotativa.")

# --- DICCIONARIO DE TIEMPOS INTEGRADO ---
# Aquí están pre-guardados todos los tiempos de llegada por ruta
TIEMPOS_RUTAS = {
    '001 TAL-STGO-': '03:00:00',
    '002 STGO - TALCA': '03:00:00',
    '003 TALCA - CHILLAN - CONCEPCION': '03:30:00',
    '004 CONCEPCION - CHILLAN- TALCA': '03:30:00',
    '004 CONCEPCION - CHILLAN- TALCA 07:45': '03:40:00',
    '005 TALCA - CHILLAN': '02:00:00',
    '006 CHILLAN - TALCA': '02:00:00',
    '007 CAUQUENES - SANTIAGO 1:15': '04:30:00',
    '007 CAUQUENES - SANTIAGO 1:30': '04:45:00',
    '008 SANTIAGO - CAUQUENES': '04:45:00',
    '009 TALCA-CHILLAN-LOS ANGELES': '03:30:00',
    '010 LOS ANGELES - CHILLAN - TALCA': '04:30:00',
    '011 CURANIPE-PELLUHUE-CAUQUENES-TALCA-SANTIAGO': '06:30:00',
    '012 SANTIAGO-TALCA-CAUQUENES-PELLUHUE-CURANIPE': '06:15:00',
    '013 TALCA - CAUQUENES': '01:30:00',
    '014 CAUQUENES - TALCA': '01:30:00',
    '015 TAL-STGO EXPRESO': '02:45:00',
    '016 STGO - TALCA EXPRESO': '02:45:00',
    '017 TALCA CONCEPCION DIRECTO': '03:00:00',
    'CONCEPCION - CHILLAN': '01:30:00',
    'TALCA BICENTENARIO': '03:00:00',
    'TALCA BUIN ZOO': '03:00:00',
    'TALCA MONUMENTAL': '03:00:00',
    'TALCA NACIONAL': '03:00:00',
    'TALCA SANTANDER': '03:00:00'
}

# --- CARGADOR DE ARCHIVOS ---
archivo_terra = st.file_uploader("📥 Subir Reporte de Terra (Excel)", type=["xlsx", "xls"])

# --- FUNCIONES ---
def calcular_hora_llegada(hora_salida, ruta):
    """Suma la hora de salida y el tiempo pre-guardado para obtener la hora de llegada."""
    duracion = TIEMPOS_RUTAS.get(ruta) # Busca el tiempo exacto usando la ruta
    
    if pd.isna(hora_salida) or not duracion:
        return None # Si no hay hora de salida o la ruta no existe en el diccionario
        
    try:
        # Convertir la hora de salida (texto) a un objeto de tiempo
        h_s = pd.to_datetime(str(hora_salida), format='%H:%M').time()
        
        # Extraer horas, minutos y segundos del diccionario
        h, m, s = map(int, duracion.split(':'))
        td = timedelta(hours=h, minutes=m, seconds=s)
            
        # Sumar los tiempos creando una fecha base
        dt_salida = datetime.combine(datetime.today(), h_s)
        dt_llegada = dt_salida + td
        return dt_llegada.strftime('%H:%M')
    except Exception:
        return None

def resaltar_cauquenes_santiago(row):
    """Aplica color #ff9900 si la ruta es Cauquenes-Santiago o viceversa."""
    origen = str(row['Origen (ciudad)']).strip().lower()
    destino = str(row['Destino (ciudad)']).strip().lower()
    
    if (origen == 'santiago' and destino == 'cauquenes') or (origen == 'cauquenes' and destino == 'santiago'):
        return ['background-color: #ff9900; color: black; font-weight: bold;'] * len(row)
    return [''] * len(row)

# --- PROCESAMIENTO PRINCIPAL ---
if archivo_terra:
    try:
        # 1. Leer el reporte de Terra (saltando la primera fila que es un encabezado inútil)
        df_terra = pd.read_excel(archivo_terra, skiprows=1)
        
        with st.spinner("Procesando y calculando tiempos de llegada..."):
            
            # 2. Calcular la Hora de Llegada consultando el diccionario interno
            df_terra['Hora de llegada'] = df_terra.apply(
                lambda row: calcular_hora_llegada(row['Hora salida'], row['Ruta']), axis=1
            )
            
            # 3. Separar la Ocupación Relativa
            # Toma algo como "2/42 (4.76%)" y lo corta donde hay un espacio.
            temp = df_terra['Ocupación Relativa'].astype(str).str.split(' ', n=1, expand=True)
            df_terra['Ocupacion'] = temp[0] # Guarda el "2/42"
            
            if temp.shape[1] > 1:
                # Limpia los paréntesis para dejar solo "4.76%"
                df_terra['% ocupacion'] = temp[1].str.replace(r'\(|\)', '', regex=True)
            else:
                df_terra['% ocupacion'] = ""
                
            # 4. Seleccionar el orden exacto de las columnas que pediste
            columnas_finales = [
                'Fecha salida', 'Hora salida', 'Hora de llegada', 
                'Origen (ciudad)', 'Destino (ciudad)', 'Ocupacion', '% ocupacion'
            ]
            
            df_final = df_terra[columnas_finales]
            
            # 5. Aplicar Estilos de Color
            # Aplicamos la función para resaltar las filas según el origen y destino
            df_final_estilizado = df_final.style.apply(resaltar_cauquenes_santiago, axis=1)
            
            st.success("✅ ¡Reporte procesado correctamente!")
            
            # 6. Mostrar Vista Previa
            st.subheader("Vista Previa del Reporte Resultante")
            st.dataframe(df_final_estilizado, use_container_width=True)
            
            # 7. Crear Excel para descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Al exportar el dataframe estilizado (df_final_estilizado), pandas guarda los colores en el Excel
                df_final_estilizado.to_excel(writer, sheet_name='Pre Rotativa', index=False)
            
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
