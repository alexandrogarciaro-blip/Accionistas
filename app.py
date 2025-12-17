import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import json
import tempfile
import os
import re

# --- 1. CONFIGURACIÓN ---
st.set_page_config(page_title="Extractor Accionariado Excel", page_icon="📊", layout="wide")

# Estilos CSS
st.markdown("""
    <style>
    section[data-testid="stSidebar"] {background-color: #101820;}
    section[data-testid="stSidebar"] * {color: #ffffff !important;}
    .stButton>button {width: 100%; background-color: #28a745; color: white; border: none;}
    .stButton>button:hover {background-color: #218838; color: white;}
    </style>
    """, unsafe_allow_html=True)

# Conexión API
try:
    api_key = st.secrets["GOOGLE_API_KEY"]
    genai.configure(api_key=api_key)
except:
    st.error("⚠️ Error: Configura la API Key en los Secrets.")
    st.stop()

# --- 2. FUNCIONES DE LÓGICA ---

def clean_json_response(text):
    """Limpia la respuesta de la IA para obtener solo el JSON válido"""
    text = re.sub(r'```json', '', text)
    text = re.sub(r'```', '', text)
    start = text.find('{')
    end = text.rfind('}') + 1
    if start != -1 and end != -1:
        return text[start:end]
    return text

def generate_excel(data_json):
    """
    Convierte el JSON de datos en el Excel con el formato exacto de la imagen.
    Estructura: Accionista 1 | TOTALES | Accionista 2 | Accionista 3 ...
    """
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet("Accionariado")

    # --- FORMATOS ---
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#f2f2f2', 'border': 1})
    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    total_header_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#d9ead3', 'border': 1, 'text_wrap': True})
    
    # Datos extraídos
    shareholders = data_json.get("accionistas", [])
    totals = data_json.get("totales", {"publica": 0, "privada": 100})

    # --- CABECERAS ---
    # Fila 1: Títulos grandes (ACCIONISTA 1, TOTALES, ACCIONISTA 2...)
    current_col = 0
    
    # 1. ACCIONISTA 1 (Si existe)
    if len(shareholders) > 0:
        worksheet.merge_range(0, current_col, 0, current_col + 5, "ACCIONISTA 1", header_format)
        # Subcabeceras Accionista
        headers = ["NIF", "RAZON SOCIAL", "%", "PAIS", "NATURALEZA", "PYME"]
        for i, h in enumerate(headers):
            worksheet.write(1, current_col + i, h, header_format)
        
        # Datos Accionista 1
        s = shareholders[0]
        row_data = [s.get("nif"), s.get("nombre"), s.get("porcentaje"), s.get("pais"), s.get("naturaleza"), s.get("pyme")]
        for i, d in enumerate(row_data):
            worksheet.write(2, current_col + i, d, cell_format)
        
        current_col += 6

    # 2. BLOQUE TOTALES (En medio, como en la imagen)
    worksheet.merge_range(0, current_col, 0, current_col + 1, "TOTALES", total_header_format)
    worksheet.write(1, current_col, "TOTAL PARTICIPACIÓN PÚBLICA", total_header_format)
    worksheet.write(1, current_col + 1, "TOTAL PARTICIPACIÓN PRIVADA", total_header_format)
    
    worksheet.write(2, current_col, f"{totals.get('publica')}%", cell_format)
    worksheet.write(2, current_col + 1, f"{totals.get('privada')}%", cell_format)
    
    current_col += 2

    # 3. RESTO DE ACCIONISTAS (2, 3, 4...)
    for idx, s in enumerate(shareholders[1:], start=2): # Empezamos en el 2
        worksheet.merge_range(0, current_col, 0, current_col + 5, f"ACCIONISTA {idx}", header_format)
        # Subcabeceras
        headers = ["NIF", "RAZON SOCIAL", "%", "PAIS", "NATURALEZA", "PYME"]
        for i, h in enumerate(headers):
            worksheet.write(1, current_col + i, h, header_format)
        
        # Datos
        row_data = [s.get("nif"), s.get("nombre"), s.get("porcentaje"), s.get("pais"), s.get("naturaleza"), s.get("pyme")]
        for i, d in enumerate(row_data):
            worksheet.write(2, current_col + i, d, cell_format)
        
        current_col += 6

    # Ajustar ancho de columnas
    worksheet.set_column(0, current_col, 15)

    writer.close()
    return output.getvalue()

# --- 3. INTERFAZ ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/732/732220.png", width=100)
    st.title("Panel de Control")
    uploaded_files = st.file_uploader("Sube las Escrituras (PDF)", type=['pdf'], accept_multiple_files=True)
    st.markdown("---")
    process_btn = st.button("GENERAR EXCEL 📗", type="primary")

st.title("📊 Extractor de Accionariado a Excel")
st.markdown("Esta herramienta lee las escrituras y genera el archivo `.xlsx` con el formato horizontal específico (Accionista 1 | Totales | Accionista 2...).")

if process_btn and uploaded_files:
    with st.spinner("Analizando participaciones y calculando estructura..."):
        try:
            # 1. Subir archivos
            gemini_files = []
            for f in uploaded_files:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(f.getvalue())
                    tmp_path = tmp.name
                g_file = genai.upload_file(path=tmp_path, display_name=f.name)
                gemini_files.append(g_file)
                os.remove(tmp_path)
            
            # Esperar a que procese
            time.sleep(2)

            # 2. PROMPT ESTRUCTURADO (JSON)
            SYSTEM_PROMPT = """
            ROL: Analista de Datos Societarios.
            OBJETIVO: Analizar las escrituras y extraer la estructura accionarial FINAL ACTUAL.

            SALIDA OBLIGATORIA: Debes devolver UNICAMENTE un objeto JSON válido con esta estructura exacta:
            {
                "accionistas": [
                    {
                        "nif": "...", 
                        "nombre": "...", 
                        "porcentaje": 50.00, 
                        "pais": "ESPAÑA", 
                        "naturaleza": "Persona Física" o "Persona Jurídica", 
                        "pyme": "SI" o "NO"
                    },
                    ... (resto de socios ordenados por % de mayor a menor)
                ],
                "totales": {
                    "publica": 0, 
                    "privada": 100
                }
            }
            
            REGLAS:
            1. 'naturaleza': Si es empresa pon "Persona Jurídica", si es humano "Persona Física".
            2. 'pyme': Si es persona física pon "NO". Si es empresa, intenta deducirlo, si no sabes pon "SI".
            3. 'totales': Calcula qué % suma el capital público (estado) y privado. Normalmente privada es 100%.
            4. Devuelve SOLO EL JSON. Sin texto antes ni después.
            """

            model = genai.GenerativeModel("gemini-2.5-flash", system_instruction=SYSTEM_PROMPT)
            response = model.generate_content(["Extrae los datos actuales.", *gemini_files])
            
            # 3. Procesar JSON
            json_str = clean_json_response(response.text)
            data = json.loads(json_str)
            
            # 4. Generar Excel
            excel_data = generate_excel(data)
            
            st.success("¡Datos extraídos correctamente!")
            
            st.download_button(
                label="📥 Descargar Excel (.xlsx)",
                data=excel_data,
                file_name="Estructura_Accionarial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
            st.warning("Detalle técnico: La IA no devolvió un JSON válido o hubo un error de conexión.")


