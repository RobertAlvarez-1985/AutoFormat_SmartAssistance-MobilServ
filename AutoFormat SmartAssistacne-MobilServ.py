import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
from PIL import Image
import base64

# Configuración de la página
st.set_page_config(
    page_title="Smart Assistance to MobilServ Converter",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS personalizado para un diseño moderno y profesional
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    }
    
    .process-flow {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 2rem 0;
        gap: 2rem;
    }
    
    .arrow {
        font-size: 2rem;
        color: #667eea;
        font-weight: bold;
    }
    
    .success-message {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
        text-align: center;
        font-weight: 500;
    }
    
    .instruction-box {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1rem 0;
    }
    
    .footer {
        text-align: center;
        padding: 2rem;
        border-top: 1px solid #e9ecef;
        margin-top: 3rem;
        background-color: #f8f9fa;
        border-radius: 8px;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        font-weight: 600;
        transition: transform 0.2s;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
    }
</style>
""", unsafe_allow_html=True)

# Header principal
st.markdown("""
<div class="main-header">
    <h1>🔄 Smart Assistance to MobilServ Converter</h1>
    <p>Transformador profesional de datos de análisis de lubricantes</p>
</div>
""", unsafe_allow_html=True)

# Función para mostrar las imágenes en el flujo del proceso
def display_process_flow():
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col1:
        st.markdown("### 📤 Smart Assistance")
        # Placeholder para la imagen Smart Assistance
        st.markdown("""
        <div style="background: linear-gradient(135deg, #ff6b6b, #ee5a52); 
                    color: white; padding: 2rem; border-radius: 10px; 
                    text-align: center; height: 120px; display: flex; 
                    align-items: center; justify-content: center;">
            <div>
                <h4>Smart Assistance</h4>
                <p>Formato Original</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("### ⚡ Conversión")
        st.markdown("""
        <div style="text-align: center; padding: 2rem;">
            <div style="font-size: 3rem; color: #667eea;">→</div>
            <p style="margin-top: 1rem; font-weight: 600;">Procesando...</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("### 📥 MobilServ")
        # Placeholder para la imagen MobilServ
        st.markdown("""
        <div style="background: linear-gradient(135deg, #4ecdc4, #44a08d); 
                    color: white; padding: 2rem; border-radius: 10px; 
                    text-align: center; height: 120px; display: flex; 
                    align-items: center; justify-content: center;">
            <div>
                <h4>MobilServ</h4>
                <p>Formato Convertido</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

# Instrucciones de uso
st.markdown("""
<div class="instruction-box">
    <h3>📋 Instrucciones de Uso</h3>
    <ol>
        <li><strong>Cargue el archivo:</strong> Seleccione y cargue el archivo Excel descargado de Smart Assistance</li>
        <li><strong>Procesamiento automático:</strong> La aplicación reordenará las columnas y convertirá los datos automáticamente</li>
        <li><strong>Descarga:</strong> Una vez completada la conversión, descargue el archivo transformado</li>
        <li><strong>Verificación:</strong> El archivo convertido estará listo para usar en MobilServ</li>
    </ol>
</div>
""", unsafe_allow_html=True)

# Mostrar flujo del proceso
display_process_flow()

# Definir mapeos de columnas y encabezados
def get_column_mappings():
    """Define los mapeos de columnas según la macro original"""
    
    # Grupo 1: Información de la Muestra
    grupo1 = {
        'A': 'W', 'Y': 'B', 'H': 'C', 'U': 'E', 'X': 'F', 'Z': 'J', 
        'V': 'L', 'W': 'O', 'E': 'AA', 'F': 'AB', 'C': 'K', 'D': 'AH', 
        'G': 'AC', 'I': 'BB', 'J': 'BC', 'K': 'BD', 'L': 'BE', 'M': 'BF', 
        'N': 'BG', 'O': 'I', 'B': 'R'
    }
    
    # Grupo 2: Elementos desgaste + aditivos
    grupo2 = {
        'IO': 'FW', 'MI': 'CC', 'AJ': 'CG', 'FK': 'CY', 'BV': 'DA', 
        'IE': 'DS', 'OZ': 'GT', 'MK': 'FS', 'JQ': 'ES', 'JJ': 'EM', 
        'HK': 'GJ', 'OB': 'GH', 'OG': 'EQ', 'MM': 'EE', 'PD': 'GX',
        'BI': 'CK', 'BD': 'CM', 'BM': 'CO', 'BL': 'CQ', 'JE': 'EI', 
        'JF': 'EK', 'HQ': 'FA', 'PO': 'HN', 'BZ': 'FK', 'FB': 'FM', 
        'FC': 'FO', 'FA': 'FQ'
    }
    
    # Grupo 3: Variables de análisis
    grupo3 = {
        'KB': 'EW', 'JR': 'EU', 'JU': 'GN', 'IY': 'GP', 'JV': 'GR',
        'IG': 'GL', 'GO': 'DY', 'AE': 'HH', 'CS': 'HJ', 'EF': 'PI',
        'PG': 'GZ', 'CE': 'EO', 'PC': 'GV', 'AD': 'HD', 'PH': 'HB'
    }
    
    return {**grupo1, **grupo2, **grupo3}

def get_new_headers():
    """Define los nuevos encabezados según la macro"""
    headers = [
        "Sample Status", "Report Status", "Date Reported", "Asset ID", "Unit ID", 
        "Unit Description", "Asset Class", "Position", "Tested Lubricant", "Service Level",
        "Sample Bottle ID", "Manufacturer", "Alt Manufacturer", "Model", "Alt Model",
        "Model Year", "Serial Number", "Account Name", "Account ID", "Oil Rating",
        "Contamination Rating", "Equipment Rating", "Parent Account Name", "Parent Account ID",
        "ERP Account Number", "Days Since Sampled", "Date Sampled", "Date Registered",
        "Date Received", "Country", "Laboratory", "Business Lines", "Fully Qualified",
        "LIMS Sample ID", "Schedule", "Tested Lubricant ID", "Registered Lubricant",
        "Registered Lubricant ID", "Zone", "Work ID", "Sampler", "IMO No", "Service Type",
        "Component Type", "Fuel Type", "RPM", "Cycles", "Pressure", "kW Rating",
        "Cylinder Number", "Target PC 4", "Target PC 6", "Target PC 14", "Equipment Age",
        "Equipment UOM", "Oil Age", "Oil Age UOM", "Makeup Volume", "MakeUp Volume UOM",
        "Oil Changed", "Filter Changed", "Oil Temp In", "Oil Temp Out", "Oil Temp UOM",
        "Coolant Temp In", "Coolant Temp Out", "Coolant Temp UOM", "Reservoir Temp",
        "Reservoir Temp UOM", "Total Engine Hours", "Hrs. Since Last Overhaul",
        "Oil Service Hours", "Used Oil Volume", "Used Oil Volume UOM", "Oil Used in Last 24Hrs",
        "Oil Used in Last 24Hrs UOM", "Sulphur %", "Engine Power at Sampling", "Date Landed",
        "Port Landed", "Ag (Silver)", "RESULT_Ag", "Air Release @50 (min)", 
        "RESULT_Air Release @50 (min)", "Al (Aluminum)", "RESULT_Al", "Appearance", 
        "RESULT_Appearance", "B (Boron)", "RESULT_ B", "Ba (Barium)", "RESULT_Ba", 
        "Ca (Calcium)", "RESULT_Ca", "Cd (Cadmium)", "RESULT_Cd", "Cl (Chlorine ppm - Xray)",
        "RESULT_Cl (Chlorine ppm - Xray)", "Compatibility", "RESULT_Compatibility",
        "Coolant Indicator", "RESULT_Coolant Indicator", "Cr (Chromium)", "RESULT_Cr",
        "Cu (Copper)", "RESULT_Cu", "DAC(%Asphalt.)", "RESULT_DAC(%Asphalt.)"
        # ... (continuar con todos los encabezados de la macro)
    ]
    return headers

def convert_dates(df, date_columns):
    """Convierte columnas de fecha del formato texto a fecha"""
    for col in date_columns:
        if col in df.columns:
            for idx in df.index:
                cell_value = str(df.at[idx, col])
                if '-' in cell_value and ':' in cell_value:
                    try:
                        # Separar fecha y hora, tomar solo la fecha
                        date_part = cell_value.split(' ')[0]
                        date_components = date_part.split('-')
                        if len(date_components) == 3:
                            year, month, day = date_components
                            df.at[idx, col] = f"{month}/{day}/{year}"
                    except:
                        continue
    return df

def apply_number_formatting(df):
    """Aplica formato numérico a columnas específicas"""
    # Columnas que deben ser números enteros
    integer_columns = ['BB', 'BD', 'BF', 'CC', 'CG', 'CK', 'CM', 'CO', 'CQ', 'CY', 
                      'DA', 'DS', 'EE', 'EI', 'EK', 'EM', 'EQ', 'ES', 'EW', 'FA', 
                      'FM', 'FO', 'FQ', 'FS', 'FW', 'GH', 'GJ', 'GT', 'GX', 'HN']
    
    # Columnas que deben ser decimales
    decimal_columns = ['DY', 'GL', 'GN', 'GP', 'GR', 'GZ', 'HB', 'HH', 'HJ']
    
    # Convertir columnas enteras
    for col in integer_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    
    # Convertir columnas decimales
    for col in decimal_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).round(2)
    
    return df

def complete_status_column(df):
    """Completa la columna A con 'Completed' donde hay datos en columna B"""
    if 'B' in df.columns:
        mask = df['B'].notna() & (df['B'] != '')
        df.loc[mask, 'A'] = 'Completed'
    return df

def process_excel_file(uploaded_file):
    """Procesa el archivo Excel aplicando todas las transformaciones"""
    try:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        # Obtener mapeos de columnas
        column_mappings = get_column_mappings()
        
        # Crear DataFrame resultado con el orden correcto de columnas
        result_df = pd.DataFrame()
        
        # Aplicar mapeos de columnas
        for old_col, new_col in column_mappings.items():
            if old_col in df.columns:
                result_df[new_col] = df[old_col]
        
        # Aplicar nuevos encabezados
        new_headers = get_new_headers()
        if len(new_headers) <= len(result_df.columns):
            result_df.columns = new_headers[:len(result_df.columns)]
        
        # Convertir fechas
        date_columns = ['C', 'AA', 'AB', 'AC']  # Date Reported, Date Sampled, Date Registered, Date Received
        result_df = convert_dates(result_df, date_columns)
        
        # Aplicar formato numérico
        result_df = apply_number_formatting(result_df)
        
        # Completar columna de estado
        result_df = complete_status_column(result_df)
        
        return result_df, None
        
    except Exception as e:
        return None, str(e)

# Sección principal de carga de archivos
st.markdown("## 📤 Cargar Archivo Smart Assistance")

uploaded_file = st.file_uploader(
    "Seleccione el archivo Excel de Smart Assistance",
    type=['xlsx', 'xls'],
    help="Cargue el archivo descargado directamente de la plataforma Smart Assistance"
)

if uploaded_file is not None:
    with st.spinner('🔄 Procesando archivo...'):
        # Procesar el archivo
        result_df, error = process_excel_file(uploaded_file)
        
        if error:
            st.error(f"❌ Error al procesar el archivo: {error}")
        else:
            # Mostrar mensaje de éxito
            st.markdown("""
            <div class="success-message">
                ✅ Transformación y reordenamiento de datos a plantilla MobilServ, completado con éxito
            </div>
            """, unsafe_allow_html=True)
            
            # Mostrar vista previa de los datos
            st.markdown("## 📊 Vista Previa del Archivo Convertido")
            st.dataframe(result_df.head(10), use_container_width=True)
            
            # Generar nombre del archivo con fecha y hora
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"ExportMobilServ_{timestamp}.xlsx"
            
            # Crear buffer para el archivo Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='MobilServ')
            
            buffer.seek(0)
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Archivo Convertido",
                data=buffer.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help=f"Descargar como: {filename}"
            )
            
            # Información adicional
            st.info(f"📋 **Estadísticas del procesamiento:**\n"
                   f"- Filas procesadas: {len(result_df)}\n"
                   f"- Columnas en archivo final: {len(result_df.columns)}\n"
                   f"- Nombre del archivo: {filename}")

# Footer con información del desarrollador
st.markdown("""
<div class="footer">
    <h4>👨‍💻 Desarrollado por Ingeniero Roberto Alvarez</h4>
    <p><strong>RCA Smart Tools</strong> | Soluciones Inteligentes para Análisis de Lubricantes</p>
    <p style="font-size: 0.9em; color: #666;">
        © 2024 Todos los derechos reservados. Esta aplicación está protegida por derechos de autor.
    </p>
    <p style="font-size: 0.8em; color: #888;">
        🔧 Versión 1.0 - Transformador Smart Assistance ➜ MobilServ
    </p>
</div>
""", unsafe_allow_html=True)
