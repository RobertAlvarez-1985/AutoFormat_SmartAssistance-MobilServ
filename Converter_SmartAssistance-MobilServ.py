import streamlit as st
import pandas as pd
import numpy as np
import io
from PIL import Image

# --- Configuración de la Página de Streamlit ---
st.set_page_config(
    page_title="Conversor Smart Assistance a MobilServ",
    page_icon="🔄",
    layout="wide"
)

# --- Funciones de Lógica de Conversión ---

def letter_to_index(letter):
    """Convierte una letra de columna de Excel a un índice numérico (base 0)."""
    letter = letter.upper()
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A')) + 1
    return result - 1

def process_excel_file(df):
    """
    Función principal que orquesta toda la lógica de conversión del archivo Excel.
    """
    
    # 1. Mapeo de columnas: (Origen, Destino)
    movimientos = [
        ("A", "W"), ("Y", "B"), ("H", "C"), ("U", "E"), ("X", "F"), ("Z", "J"),
        ("V", "L"), ("W", "O"), ("E", "AA"), ("F", "AB"), ("C", "K"), ("D", "AH"),
        ("G", "AC"), ("I", "BB"), ("J", "BC"), ("K", "BD"), ("L", "BE"), ("M", "BF"),
        ("N", "BG"), ("O", "I"), ("B", "R"),
        ("IP", "FW"), ("MI", "CC"), ("AJ", "CG"), ("FL", "CY"), ("BW", "DA"),
        ("IE", "DS"), ("PA", "GT"), ("MM", "FS"), ("JR", "ES"), ("JL", "EM"),
        ("HK", "GJ"), ("OD", "GH"), ("OG", "EQ"), ("MO", "EE"), ("PE", "GX"),
        ("BJ", "CK"), ("BD", "CM"), ("BN", "CO"), ("BL", "CQ"), ("JF", "EI"),
        ("JG", "EK"), ("HQ", "FA"), ("PP", "HN"), ("BZ", "FK"), ("FB", "FM"),
        ("FC", "FO"), ("FA", "FQ"),
        ("KC", "EW"), ("JS", "EU"), ("JV", "GN"), ("JW", "GP"), ("JX", "GR"),
        ("IG", "GL"), ("GO", "DY"), ("AE", "HH"), ("CS", "HJ"), ("EF", "PI"),
        ("PH", "GZ"), ("CE", "EO"), ("PD", "GV"), ("AD", "HD"), ("PI", "HB")
    ]

    origen_indices = [letter_to_index(m[0]) for m in movimientos]
    destino_indices = [letter_to_index(m[1]) for m in movimientos]

    # Crear un DataFrame de destino lo suficientemente grande
    max_col_index = max(destino_indices) if destino_indices else df.shape[1]
    df_nuevo = pd.DataFrame(np.nan, index=df.index, columns=range(max_col_index + 1))

    # Mover las columnas del DataFrame original al nuevo
    for orig_idx, dest_idx in zip(origen_indices, destino_indices):
        if orig_idx < df.shape[1]:
            df_nuevo.iloc[:, dest_idx] = df.iloc[:, orig_idx].values

    # --- CORRECCIÓN CLAVE: Encabezados duplicados han sido renombrados para ser únicos ---
    # Se modificaron los nombres que se repetían, como "TBN (mg KOH/g)" y "Water (Vol %)"
    header_string = (
        "Sample Status,Report Status,Date Reported,Asset ID,Unit ID,Unit Description,Asset Class,Position,"
        "Tested Lubricant,Service Level,Sample Bottle ID,Manufacturer,Alt Manufacturer,Model,Alt Model,"
        "Model Year,Serial Number,Account Name,Account ID,Oil Rating,Contamination Rating,Equipment Rating,"
        "Parent Account Name,Parent Account ID,ERP Account Number,Days Since Sampled,Date Sampled,"
        "Date Registered,Date Received,Country,Laboratory,Business Lines,Fully Qualified,LIMS Sample ID,"
        "Schedule,Tested Lubricant ID,Registered Lubricant,Registered Lubricant ID,Zone,Work ID,Sampler,"
        "IMO No,Service Type,Component Type,Fuel Type,RPM,Cycles,Pressure,kW Rating,Cylinder Number,"
        "Target PC 4,Target PC 6,Target PC 14,Equipment Age,Equipment UOM,Oil Age,Oil Age UOM,Makeup Volume,"
        "MakeUp Volume UOM,Oil Changed,Filter Changed,Oil Temp In,Oil Temp Out,Oil Temp UOM,Coolant Temp In,"
        "Coolant Temp Out,Coolant Temp UOM,Reservoir Temp,Reservoir Temp UOM,Total Engine Hours,"
        "Hrs. Since Last Overhaul,Oil Service Hours,Used Oil Volume,Used Oil Volume UOM,"
        "Oil Used in Last 24Hrs,Oil Used in Last 24Hrs UOM,Sulphur %,Engine Power at Sampling,Date Landed,"
        "Port Landed,Ag (Silver),RESULT_Ag,Air Release @50 (min),RESULT_Air Release @50 (min),Al (Aluminum),"
        "RESULT_Al,Appearance,RESULT_Appearance,B (Boron),RESULT_ B,Ba (Barium),RESULT_Ba,Ca (Calcium),"
        "RESULT_Ca,Cd (Cadmium),RESULT_Cd,Cl (Chlorine ppm - Xray),RESULT_Cl (Chlorine ppm - Xray),"
        "Compatibility,RESULT_Compatibility,Coolant Indicator,RESULT_Coolant Indicator,Cr (Chromium),RESULT_Cr,"
        "Cu (Copper),RESULT_Cu,DAC(%Asphalt.),RESULT_DAC(%Asphalt.),Demul@54C  (min),RESULT_Demul@54C  (min),"
        "Demul@54C (min),RESULT_Demul@54C (min),Demul@54C (Oil/Water/Emul/Time),RESULT_Demul@54C (Oil/Water/E/T),"
        "Demulsibility @54C (time-min),RESULT_Demulsibility @54C (time-min),Demulsibility @54C (oil),"
        "RESULT_Demulsibility @54C (oil),Demulsibility @54C (water),RESULT_Demulsibility @54C (water),"
        "Demulsibility @54C (emulsion),RESULT_Demulsibility @54C (emulsion),Fe (Iron),RESULT_Fe (Iron),"
        "Foam Seq 1 - stability (ml),RESULT_Foam Seq 1 - stability (ml),Foam Seq 1 - tendency (ml),"
        "RESULT_Foam Seq 1 - tendency (ml),Fuel Dilut. (Vol%),RESULT_Fuel Dilut. (Vol%),Initial pH,"
        "RESULT_Initial pH,Insolubles 5u,RESULT_Insolubles 5u,K (Potassium),RESULT_K,MCR,RESULT_MCR,"
        "Mg (Magnesium),RESULT_Mg,Mn (Manganese),RESULT_Mn (Manganese),Mo (Molybdenum),RESULT_Mo,"
        "MPC delta E,RESULT_MPC delta E,Na (Sodium),RESULT_Na,Ni (Nickel),RESULT_Ni,Nitration (Ab/cm),"
        "RESULT_Nitration (Ab/cm),Oxidation (Ab/cm),RESULT_Oxidation (Ab/cm),P  (Phosphorus),RESULT_P  (Phosphorus),"
        "P (Phosphorus),RESULT_P (Phosphorus),ISO Code (4/6/14),RESULT_ISO Code (4/6/14),"
        "Particle Count  >4um,RESULT_Particle Count  >4um,Particle Count  >6um,RESULT_Particle Count  >6um,"
        "Particle Count>14um,RESULT_Particle Count>14um,Diluted ISO Code (4/6/14),RESULT_Diluted ISO Code (4/6/14),"
        "Particle Count (Diluted) >4um,RESULT_Particle Count (Diluted) >4um,Particle Count (Diluted) >6um,"
        "RESULT_Particle Count (Diluted) >6um,Particle Count (Diluted) >14um,RESULT_Particle Count (Diluted) >14um,"
        "Pb (Lead),RESULT_Pb,PM Flash Pt.(°C),RESULT_PM Flash Pt.(°C),PQ Index,RESULT_PQ Index,RESULT_Product,"
        "RPVOT (Min),RESULT_RPVOT (Min),RULER-Amine (% vs new),RESULT_RULER-Amine (% vs new),"
        "RULER-Phenol (% vs new),RESULT_RULER-Phenol (% vs new),SAE Viscosity Grade,RESULT_SAE Viscosity Grade,"
        "Si (Silicon),RESULT_Si,Sn (Tin),RESULT_Sn,Soot (Wt%),RESULT_Soot (Wt%),TAN (mg KOH/g),"
        "RESULT_TAN (mg KOH/g),TBN (mg KOH/g),RESULT_TBN (mg KOH/g),TBN (mg KOH/g) 2,RESULT_TBN (mg KOH/g) 2,"
        "Ti (Titanium),RESULT_Ti,UC Rating,RESULT_UC Rating,V (Vanadium),RESULT_V,Visc@100C (cSt),"
        "RESULT_Visc@100C (cSt),Visc@40C (cSt),RESULT_Visc@40C (cSt),Water (Hot Plate),RESULT_Water (Hot Plate),"
        "Water (Vol %),RESULT_Water (Vol%),Water (Vol%),RESULT_Water (Vol%) 2,Water (Vol.%),"
        "RESULT_Water (Vol%) 3,Water Free est.,RESULT_Water Free est.,Zn (Zinc),"
        "RESULT_Zn,Zn (Zinc) 2,RESULT_Zn 2,Soot (Wt%)- No Ref,RESULT_Soot (Wt%)- No Ref,Oxidation (Abs/cm)- no Ref,RESULT_Oxidation (Abs/cm)- no Ref,"
        "Nitration (Abs/cm)- no Ref,RESULT_Nitration (Abs/cm)- no Ref,Water (Abs/cm)- no Ref,RESULT_Water (Abs/cm) - no Ref,Aluminum - GR,RESULT_Aluminum - GR,"
        "Antimony - gr,RESULT_Antimony - gr,Appearance - gr,RESULT_Appearance - gr,Barium - GR,RESULT_Barium - GR,Boron - GR,RESULT_Boron - GR,"
        "Cadmium - gr,RESULT_Cadmium - gr,Calcium - GR,RESULT_Calcium - GR,Chromium - gr,RESULT_Chromium - gr,Copper - GR,RESULT_Copper - GR,"
        "IR Correlation - gr,RESULT_IR Correlation - gr,Ferrous Debris - gr,RESULT_Ferrous Debris - gr,Stress Index - Gr,RESULT_Stress Index - Gr,"
        "Grease Thief Video,RESULT_Grease Thief Video,Iron - GR,RESULT_Iron - GR,Lead - gr,RESULT_Lead - gr,Magnesium - GR,RESULT_Magnesium - GR,"
        "Manganese - gr,RESULT_Manganese - gr,Molybdenum -gr,RESULT_Molybdenum -gr,Nickel -gr,RESULT_Nickel -gr,Phosphorus - GR,RESULT_Phosphorus - GR,"
        "Potassium - Gr,RESULT_Potassium - Gr,Silicon - gr,RESULT_Silicon - gr,Silver - Grease,RESULT_Silver - Grease,Sodium - Gr,RESULT_Sodium - Gr,"
        "Tin - gr,RESULT_Tin - gr,Titanium - gr,RESULT_Titanium - gr,Vanadium - gr,RESULT_Vanadium - gr,Water - Gr,RESULT_Water - Gr,Zinc - gr,RESULT_Zinc - gr,"
        "Fuel Dilution - INDO,RESULT_Fuel Dilution - INDO,TBN - INDO,RESULT_TBN - INDO,Soot - INDO,RESULT_Soot - INDO,Water - INDO,RESULT_Water - INDO,"
        "Oxidation - INDO,RESULT_Oxidation - INDO,Nitration - INDO,RESULT_Nitration - INDO,Boron,RESULT_Boron,Barium,RESULT_Barium,Calcium,RESULT_Calcium,"
        "Magnesium,RESULT_Magnesium,Lithium -gr,RESULT_Lithium -gr,Color -gr,RESULT_Color -gr,Chlorine,RESULT_Chlorine,Lithium,RESULT_Lithium,"
        "Antimony,RESULT_Antimony,Sulfur,RESULT_Sulfur,Insolubles,RESULT_Insolubles,Aluminum - gr - ICP,RESULT_Aluminum - gr - ICP,Antimony - gr- ICP,"
        "RESULT_Antimony - gr- ICP,Barium - gr - ICP,RESULT_Barium - gr - ICP,Boron - gr - ICP,RESULT_Boron - gr - ICP,Cadmium - gr - ICP,RESULT_Cadmium - gr - ICP,"
        "Calcium - gr - ICP,RESULT_Calcium - gr - ICP,Chromium - gr - ICP,RESULT_Chromium - gr - ICP,Copper - gr - ICP,RESULT_Copper - gr - ICP,"
        "Iron - gr - ICP,RESULT_Iron - gr - ICP,Lead - gr - ICP,RESULT_Lead - gr - ICP,Lithium - gr - ICP,RESULT_Lithium - gr - ICP,Magnesium - gr - ICP,"
        "RESULT_Magnesium - gr - ICP,Manganese - gr - ICP,RESULT_Manganese - gr - ICP,Molybdneum - gr - ICP,RESULT_Molybdneum - gr - ICP,"
        "Nickel - gr - ICP,RESULT_Nickel - gr - ICP,Phosphorus - gr - ICP,RESULT_Phosphorus - gr - ICP,Potassium - gr - ICP,RESULT_Potassium - gr - ICP,"
        "Silicon - gr - ICP,RESULT_Silicon - gr - ICP,Silver - Grease ICP,RESULT_Silver - Grease ICP,Sodium - gr - ICP,RESULT_Sodium - gr - ICP,"
        "Tin - gr - ICP,RESULT_Tin - gr - ICP,Titanium - gr - ICP,RESULT_Titanium - gr - ICP,Vanadium - gr - ICP,RESULT_Vanadium - gr - ICP,"
        "Zinc - gr - ICP,RESULT_Zinc - gr - ICP,Water (Vol%) - KF - 3P,RESULT_Water (Vol%) - KF - 3P,Water - E2412,RESULT_Water - E2412,"
        "Sulfur by xray,RESULT_Sulfur by xray,Viscosity at 100C,RESULT_Viscosity at 100C,Viscosity at 40C,RESULT_Viscosity at 40C,"
        "Blotter test,RESULT_Blotter test,TrendAnalysis,Flashpoint D3828,RESULT_Flashpoint D3828,Foam Seq 2 tendency,RESULT_Foam Seq 2 tendency,"
        "Foam_Seq 2 stability,RESULT_Foam Seq 2 stability,Foam Seq 3 tendency,RESULT_Foam Seq 3 tendency,Foam Seq 3 stability,RESULT_Foam Seq 3 stability,"
        "Dielectric breakdown,RESULT_Dielectric breakdown,Serial Number ID,RESULT_Serial Number ID,Software Version,RESULT_Software Version,"
        "Sulfation abs/0.1mm,RESULT_Sulfation abs/0.1mm,Antiwear %,RESULT_Antiwear %,Total Fe < 100um ppm,RESULT_Total Fe < 100um ppm,"
        "Fe Wear Severity Index,RESULT_Fe Wear Severity Index,Large Fe ppm,RESULT_Large Fe ppm,Non-Metallic > 20 um,RESULT_Non-Metallic > 20 um,"
        "NAS particles 5-15um,RESULT_NAS particles 5-15um,NAS particles 15-25um,RESULT_NAS particles 15-25um,NAS particles 25-50um,RESULT_NAS particles 25-50um,"
        "NAS particles 50-100um,RESULT_NAS particles 50-100um,NAS particles > 100um,RESULT_NAS particles > 100um,Glycol %,RESULT_Glycol %,"
        "Blotter Spot C-Index,RESULT_Blotter Spot C-Index,Blotter Spot Diameter,RESULT_Blotter Spot Diameter,Blotter Spot Dispersancy,"
        "RESULT_Blotter Spot Dispersancy,Blotter Spot Opacity,RESULT_Blotter Spot Opacity,Blotter Spot Note,RESULT_Blotter Spot Note"
    )
    new_headers = header_string.split(',')
    
    # Asegurar que el DataFrame tenga el número correcto de columnas y asignar los nombres
    df_final = df_nuevo.reindex(columns=range(len(new_headers)))
    df_final.columns = new_headers

    # Convertir columnas de fecha
    columnas_fecha = ["Date Reported", "Date Sampled", "Date Registered", "Date Received"]
    for col in columnas_fecha:
        if col in df_final.columns:
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce')

    # Convertir columnas numéricas
    columnas_enteras_letras = ["BB", "BD", "BF", "CC", "CG", "CK", "CM", "CO", "CQ", "CY", "DA", "DS", "EE", "EI", "EK", "EM", "EQ", "ES", "EW", "FA", "FM", "FO", "FQ", "FS", "FW", "GH", "GJ", "GT", "GX", "HN"]
    columnas_decimales_letras = ["DY", "GL", "GN", "GP", "GR", "GZ", "HB", "HH", "HJ"]

    for col_letter in columnas_enteras_letras:
        col_idx = letter_to_index(col_letter)
        if col_idx < len(df_final.columns):
            col_name = df_final.columns[col_idx]
            df_final[col_name] = pd.to_numeric(df_final[col_name], errors='coerce')
            df_final[col_name] = df_final[col_name].astype('Int64') # Nullable Integer
    
    for col_letter in columnas_decimales_letras:
        col_idx = letter_to_index(col_letter)
        if col_idx < len(df_final.columns):
            col_name = df_final.columns[col_idx]
            df_final[col_name] = pd.to_numeric(df_final[col_name], errors='coerce')

    # Completar estado de la muestra
    if 'Report Status' in df_final.columns and 'Sample Status' in df_final.columns:
        mask = df_final['Report Status'].notna() & (df_final['Report Status'] != '')
        df_final.loc[mask, 'Sample Status'] = 'Completed'

    return df_final

def to_excel(df):
    """Convierte un DataFrame a un objeto de bytes en formato Excel con formato."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', datetime_format='MM/DD/YYYY') as writer:
        df.to_excel(writer, index=False, sheet_name='MobilServ_Data')
        
        workbook = writer.book
        worksheet = writer.sheets['MobilServ_Data']
        
        # Formatos numéricos
        formato_entero = '0'
        formato_decimal = '0.00'

        columnas_enteras_letras = ["BB", "BD", "BF", "CC", "CG", "CK", "CM", "CO", "CQ", "CY", "DA", "DS", "EE", "EI", "EK", "EM", "EQ", "ES", "EW", "FA", "FM", "FO", "FQ", "FS", "FW", "GH", "GJ", "GT", "GX", "HN"]
        columnas_decimales_letras = ["DY", "GL", "GN", "GP", "GR", "GZ", "HB", "HH", "HJ"]

        # Lógica de formato optimizada
        for col_letter, number_format in [(columnas_enteras_letras, formato_entero), (columnas_decimales_letras, formato_decimal)]:
            for letter in col_letter:
                col_idx_excel = letter_to_index(letter) + 1 # +1 para índice base 1 de Excel
                if col_idx_excel <= len(df.columns):
                    for row in range(2, worksheet.max_row + 1):
                        cell = worksheet.cell(row=row, column=col_idx_excel)
                        if cell.value is not None and isinstance(cell.value, (int, float)):
                            cell.number_format = number_format

    return output.getvalue()


# --- Interfaz de Usuario (UI) con Streamlit ---

st.title("🔄 Aplicación de Conversión de Formato")
st.markdown("Transforma archivos de **Smart Assistance** al formato requerido por **MobilServ**.")

with st.expander("📖 Manual de Uso (Haga clic para expandir)"):
    st.markdown("""
    Siga estos pasos para convertir su archivo:
    1.  **Cargar Archivo**: Use el botón "Browse files" para subir su Excel. Asegúrese de que la hoja de datos se llame **"Sheet"**.
    2.  **Iniciar Conversión**: Presione el botón "Iniciar Proceso de Conversión".
    3.  **Descargar Resultado**: Una vez finalizado, presione "Descargar archivo" para guardar el resultado.
    """)

# ... (El resto de la UI permanece igual, es funcional) ...
try:
    logo_smart = Image.open("Smart Assistance.png")
    logo_mobil = Image.open("MobilServ.png")
    col1, col2, col3 = st.columns([2, 1, 2])
    with col1:
        st.image(logo_smart, caption="Formato de Origen", use_container_width=True)
    with col2:
        st.markdown("<h1 style='text-align: center; color: #007bff; margin-top: 50px;'>➡️</h1>", unsafe_allow_html=True)
    with col3:
        st.image(logo_mobil, caption="Formato de Destino", use_container_width=True)
except FileNotFoundError:
    st.info("Logos no encontrados. La funcionalidad de la aplicación no se ve afectada.")

st.divider()

st.header("1. Adjunte el archivo de Excel")
uploaded_file = st.file_uploader(
    "El archivo debe contener una hoja llamada 'Sheet'",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    st.success(f"Archivo cargado: **{uploaded_file.name}**")
    
    st.header("2. Inicie la transformación")
    if st.button("Iniciar Proceso de Conversión", type="primary", use_container_width=True):
        with st.spinner("Procesando... Por favor espere."):
            try:
                # Se omite la primera fila (encabezado) del archivo original.
                input_df = pd.read_excel(uploaded_file, sheet_name="Sheet", header=None, skiprows=1)
                
                processed_df = process_excel_file(input_df)
                
                st.session_state.processed_data = to_excel(processed_df)
                st.session_state.processing_complete = True
                st.session_state.file_name = uploaded_file.name

            except Exception as e:
                st.error(f"Ocurrió un error inesperado durante el procesamiento: {e}")
                st.error("Asegúrese de que el archivo es válido y la hoja se llama 'Sheet'.")
                st.exception(e) # Muestra el error detallado para depuración
                st.session_state.processing_complete = False

if 'processing_complete' in st.session_state and st.session_state.get('processing_complete', False):
    st.balloons()
    st.success("✔️ ¡Proceso de Conversión completado exitosamente!")
    st.header("3. Descargue el archivo final")
    
    st.download_button(
        label="📥 Descargar archivo en formato MobilServ",
        data=st.session_state.processed_data,
        file_name=f"CONVERTIDO_{st.session_state.file_name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.divider()
st.markdown("""
<div style="text-align: center; margin-top: 30px; color: grey;">
    <p>© 2025 – Todos los derechos reservados.</p>
    <p>Desarrollado por: <strong>Roberto Alvarez / RCA Smart Tools.</strong></p>
</div>
""", unsafe_allow_html=True)

