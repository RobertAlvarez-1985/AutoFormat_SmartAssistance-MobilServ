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

# --- Funciones de Lógica de Conversión (Traducción de VBA a Python) ---

def letter_to_index(letter):
    """Convierte una letra de columna de Excel (A, B, ..., Z, AA, etc.) a un índice numérico (base 0)."""
    letter = letter.upper()
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A')) + 1
    return result - 1

def process_excel_file(df):
    """
    Función principal que orquesta toda la lógica de conversión del archivo Excel.
    Replica la funcionalidad de las macros de VBA en Python usando pandas.
    """
    
    # --- 1. Reubicar Columnas y Renombrar Encabezados (Macro: ReubicarColumnasYRenombrarEncabezados_Mejorado) ---
    
    # Definición de los movimientos de columnas: (Origen, Destino)
    movimientos = [
        # Grupo 1: Información de la Muestra
        ("A", "W"), ("Y", "B"), ("H", "C"), ("U", "E"), ("X", "F"), ("Z", "J"),
        ("V", "L"), ("W", "O"), ("E", "AA"), ("F", "AB"), ("C", "K"), ("D", "AH"),
        ("G", "AC"), ("I", "BB"), ("J", "BC"), ("K", "BD"), ("L", "BE"), ("M", "BF"),
        ("N", "BG"), ("O", "I"), ("B", "R"),
        # Grupo 2: Variables: IPQ+ elementos desgaste + aditivos + Cont. + Código ISO
        ("IO", "FW"), ("MI", "CC"), ("AJ", "CG"), ("FK", "CY"), ("BV", "DA"),
        ("IE", "DS"), ("OZ", "GT"), ("MK", "FS"), ("JQ", "ES"), ("JJ", "EM"),
        ("HK", "GJ"), ("OB", "GH"), ("OG", "EQ"), ("MM", "EE"), ("PD", "GX"),
        ("BI", "CK"), ("BD", "CM"), ("BM", "CO"), ("BL", "CQ"), ("JE", "EI"),
        ("JF", "EK"), ("HQ", "FA"), ("PO", "HN"), ("BZ", "FK"), ("FB", "FM"),
        ("FC", "FO"), ("FA", "FQ"),
        # Grupo 3: Variables: OXI, NIT, TAN, TBN, HOLLIN, FUEL DIL, AGUA, VISC
        ("KB", "EW"), ("JR", "EU"), ("JU", "GN"), ("IY", "GP"), ("JV", "GR"),
        ("IG", "GL"), ("GO", "DY"), ("AE", "HH"), ("CS", "HJ"), ("EF", "PI"),
        ("PG", "GZ"), ("CE", "EO"), ("PC", "GV"), ("AD", "HD"), ("PH", "HB")
    ]

    # Convertir letras a índices numéricos
    origen_indices = [letter_to_index(m[0]) for m in movimientos]
    destino_indices = [letter_to_index(m[1]) for m in movimientos]

    # Crear un nuevo DataFrame para los datos reordenados
    max_col_index = max(destino_indices)
    df_nuevo = pd.DataFrame(np.nan, index=df.index, columns=range(max_col_index + 1))

    # Copiar las columnas del DataFrame original al nuevo
    for i, (orig_idx, dest_idx) in enumerate(zip(origen_indices, destino_indices)):
        if orig_idx < len(df.columns):
            df_nuevo.iloc[:, dest_idx] = df.iloc[:, orig_idx].values
        else:
            # Advertencia si una columna de origen no existe en el archivo cargado
            st.warning(f"Advertencia: La columna de origen '{movimientos[i][0]}' no se encontró en el archivo. Se omitirá en el destino.")
    
    # Cadena completa de encabezados, extraída del VBA
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
        "Demul@54C (min),RESULT_Demul@54C (min),Demul@54C (Oil/Water/Emul/Time),RESULT_Demul@54C (Oil/Water/Emul/Time),"
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
        "RESULT_TAN (mg KOH/g),TBN (mg KOH/g),RESULT_TBN (mg KOH/g) 2,TBN (mg KOH/g),RESULT_TBN (mg KOH/g) 2,"
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
    
    # Asegurarse de que el DataFrame tenga suficientes columnas para los nuevos encabezados
    num_headers = len(new_headers)
    if num_headers > df_nuevo.shape[1]:
        # Añadir columnas vacías si es necesario
        num_cols_to_add = num_headers - df_nuevo.shape[1]
        for _ in range(num_cols_to_add):
            df_nuevo[df_nuevo.shape[1]] = np.nan
            
    # Asignar los nuevos encabezados a las primeras N columnas
    final_columns = new_headers + [f"Unnamed: {i}" for i in range(num_headers, df_nuevo.shape[1])]
    df_nuevo.columns = final_columns
    
    # Recortar el DataFrame para que solo contenga las columnas con los nuevos encabezados
    df_final = df_nuevo.iloc[:, :num_headers]

    # --- 2. Convertir y Formatear Fechas (Macro: ConvertirYFormatearFechas) ---
    columnas_fecha = ["Date Reported", "Date Sampled", "Date Registered", "Date Received"]
    for col in columnas_fecha:
        if col in df_final.columns:
            # errors='coerce' convierte los valores no válidos en NaT (Not a Time)
            df_final[col] = pd.to_datetime(df_final[col], errors='coerce')
            # Para la descarga en Excel, es mejor dejarlo como datetime y aplicar formato al escribir.
            # Para visualización o si se requiere como texto:
            # df_final[col] = df_final[col].dt.strftime('%m/%d/%Y')

    # --- 3. Formatear Números (Macro: Formato_NumeroEntero_Decimal_Columnas) ---
    # Columnas que deben ser números enteros
    columnas_enteras_letras = ["BB", "BD", "BF", "CC", "CG", "CK", "CM", "CO", "CQ", "CY", "DA",
                               "DS", "EE", "EI", "EK", "EM", "EQ", "ES", "EW", "FA", "FM",
                               "FO", "FQ", "FS", "FW", "GH", "GJ", "GT", "GX", "HN"]
    # Columnas que deben ser números decimales
    columnas_decimales_letras = ["DY", "GL", "GN", "GP", "GR", "GZ", "HB", "HH", "HJ"]

    # Mapear letras a los nuevos nombres de encabezado
    mapa_letras_a_nombres = {letra: new_headers[letter_to_index(letra)] for letra in columnas_enteras_letras + columnas_decimales_letras if letter_to_index(letra) < len(new_headers)}

    for letra, nombre_col in mapa_letras_a_nombres.items():
        if nombre_col in df_final.columns:
            df_final[nombre_col] = pd.to_numeric(df_final[nombre_col], errors='coerce')
            if letra in columnas_decimales_letras:
                # El formato se aplica al guardar en Excel, aquí solo nos aseguramos de que sea numérico (float)
                pass 
            elif letra in columnas_enteras_letras:
                # Convertir a tipo entero que admite nulos (NaN)
                df_final[nombre_col] = df_final[nombre_col].astype('Int64')

    # --- 4. Completar Estado (Macro: CompletarEstadoColumnaA) ---
    # La columna 'A' original se convierte en 'Sample Status'.
    # La columna 'B' original se convierte en 'Report Status'.
    if 'Report Status' in df_final.columns and 'Sample Status' in df_final.columns:
        # Si la celda en 'Report Status' no está vacía, poner "Completed" en 'Sample Status'
        df_final.loc[df_final['Report Status'].notna() & (df_final['Report Status'] != ''), 'Sample Status'] = 'Completed'

    return df_final

def to_excel(df):
    """Convierte un DataFrame a un objeto de bytes en formato Excel."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', datetime_format='MM/DD/YYYY') as writer:
        df.to_excel(writer, index=False, sheet_name='MobilServ_Data')
        
        # Aplicar formato numérico a las columnas (opcional, pero mejora el resultado)
        workbook = writer.book
        worksheet = writer.sheets['MobilServ_Data']
        
        # Formato para enteros
        formato_entero = '0'
        # Formato para decimales
        formato_decimal = '0.00'

        # Columnas enteras
        columnas_enteras_letras = ["BB", "BD", "BF", "CC", "CG", "CK", "CM", "CO", "CQ", "CY", "DA", "DS", "EE", "EI", "EK", "EM", "EQ", "ES", "EW", "FA", "FM", "FO", "FQ", "FS", "FW", "GH", "GJ", "GT", "GX", "HN"]
        # Columnas decimales
        columnas_decimales_letras = ["DY", "GL", "GN", "GP", "GR", "GZ", "HB", "HH", "HJ"]

        header_list = list(df.columns)
        for col_letter in columnas_enteras_letras:
            idx = letter_to_index(col_letter)
            if idx < len(header_list):
                col_name = header_list[idx]
                col_idx_in_df = df.columns.get_loc(col_name) + 1
                for cell in worksheet[chr(ord('A') + col_idx_in_df -1)]:
                    cell.number_format = formato_entero

        for col_letter in columnas_decimales_letras:
            idx = letter_to_index(col_letter)
            if idx < len(header_list):
                col_name = header_list[idx]
                col_idx_in_df = df.columns.get_loc(col_name) + 1
                for cell in worksheet[chr(ord('A') + col_idx_in_df -1)]:
                    cell.number_format = formato_decimal

    processed_data = output.getvalue()
    return processed_data

# --- Interfaz de Usuario (UI) con Streamlit ---

st.title("🔄 Aplicación de Conversión de Formato")
st.markdown("Transforma archivos de **Smart Assistance** al formato requerido por **MobilServ**.")

# Sección del Manual de Usuario
with st.expander("📖 Manual de Uso (Haga clic para expandir)"):
    st.markdown("""
    Esta aplicación le permite convertir archivos de Excel de forma rápida y sencilla. Siga estos pasos:

    1.  **Cargar el Archivo**:
        -   Haga clic en el botón **"Browse files"** o arrastre y suelte su archivo de Excel en el área designada.
        -   Asegúrese de que el archivo tenga el formato original de "Smart Assistance".
        -   El archivo debe tener los datos en una hoja llamada **"Plantilla"**.

    2.  **Iniciar la Conversión**:
        -   Una vez cargado el archivo, aparecerá un botón llamado **"Iniciar Proceso de Conversión"**.
        -   Haga clic en este botón para comenzar la transformación de los datos.

    3.  **Descargar el Resultado**:
        -   Después de unos segundos, el proceso finalizará y se mostrará un mensaje de éxito.
        -   Aparecerá un botón de **"Descargar archivo en formato MobilServ"**.
        -   Haga clic en él para guardar el archivo convertido en su computador.
    """)

# Elemento gráfico que representa la conversión
try:
    logo_smart = Image.open("Smart Assistance.png")
    logo_mobil = Image.open("MobilServ.png")
    
    col1, col2, col3 = st.columns([2, 1, 2])
    with col1:
        st.image(logo_smart, caption="Formato de Origen", use_column_width=True)
    with col2:
        st.markdown("<h1 style='text-align: center; color: #007bff; margin-top: 50px;'>➡️</h1>", unsafe_allow_html=True)
    with col3:
        st.image(logo_mobil, caption="Formato de Destino", use_column_width=True)
except FileNotFoundError:
    st.info("Logos no encontrados. Coloque 'Smart Assistance.png' y 'MobilServ.png' en la misma carpeta que la aplicación para una mejor visualización.")

st.divider()

# Carga de Archivo
st.header("1. Adjunte el archivo de Excel con formato Smart Assistance")
uploaded_file = st.file_uploader(
    "El archivo debe contener una hoja llamada 'Plantilla'",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    st.success(f"Archivo cargado: **{uploaded_file.name}**")
    
    # Botón de Inicio
    st.header("2. Inicie la transformación")
    if st.button("Iniciar Proceso de Conversión", type="primary", use_container_width=True):
        with st.spinner("Procesando... Por favor espere."):
            try:
                # Leer los datos de la hoja "Plantilla" sin encabezados
                input_df = pd.read_excel(uploaded_file, sheet_name="Plantilla", header=None)
                
                # Procesar el DataFrame
                processed_df = process_excel_file(input_df)
                
                # Guardar el resultado en el estado de la sesión para la descarga
                st.session_state.processed_data = to_excel(processed_df)
                st.session_state.processing_complete = True

            except Exception as e:
                st.error(f"Ocurrió un error durante el procesamiento: {e}")
                st.error("Asegúrese de que el archivo cargado sea válido y contenga una hoja llamada 'Plantilla'.")
                st.session_state.processing_complete = False

# Notificación y Descarga Post-Conversión
if 'processing_complete' in st.session_state and st.session_state.processing_complete:
    st.balloons()
    st.success("✔️ Proceso de Conversión de formato Exitoso :)")
    st.header("3. Descargue el archivo final")
    
    st.download_button(
        label="📥 Descargar archivo en formato MobilServ",
        data=st.session_state.processed_data,
        file_name=f"CONVERTIDO_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.divider()

# Información de Propiedad y Desarrollo
st.markdown("""
<div style="text-align: center; margin-top: 30px; color: grey;">
    <p>© 2025 – Todos los derechos reservados.</p>
    <p>Desarrollado por: <strong>Roberto Alvarez / RCA Smart Tools.</strong></p>
</div>
""", unsafe_allow_html=True)
