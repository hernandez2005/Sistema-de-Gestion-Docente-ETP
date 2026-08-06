import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import PyPDF2
import json
from google.api_core.exceptions import ResourceExhausted
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# --- ESTILOS CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #F8FAFC; color: #1E293B; }
    .main-header { font-size: 2.2rem; font-weight: 700; color: #0F172A; text-align: center; margin-bottom: 5px; line-height: 1.2; }
    .sub-header { text-align: center; color: #475569; font-size: 1.1rem; font-weight: 400; margin-bottom: 35px; }
    [data-testid="stForm"] { background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); padding: 30px; }
    .section-title { color: #1D4ED8; font-weight: 600; font-size: 1.2rem; border-bottom: 2px solid #DBEAFE; padding-bottom: 8px; margin-top: 20px; margin-bottom: 18px; }
    div.stButton > button:first-child, div.stFormSubmitButton > button:first-child { background-color: #2563EB !important; color: #FFFFFF !important; border: none !important; border-radius: 6px !important; font-weight: 600 !important; padding: 10px 24px !important; width: 100%; }
    div.stButton > button:first-child:hover, div.stFormSubmitButton > button:first-child:hover { background-color: #1D4ED8 !important; }
</style>
""", unsafe_allow_html=True)

# --- REINTENTO DE API ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_con_reintento(api_key, modelo, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(max_output_tokens=8192, temperature=0.0) 
    )
    return respuesta.text

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("⚡ Núcleo de Procesamiento")
    proveedor_ia = st.selectbox("Motor Analítico:", ["Google Gemini", "OpenAI (ChatGPT)"], key="prov_pond")
    
    if proveedor_ia == "Google Gemini":
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"], key="mod_pond")
    else:
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gpt-4o-mini", "gpt-3.5-turbo", "gpt-4o"], key="mod_pond2")
        
    api_key_usuario = st.text_input("Clave de Autenticación (API Key):", type="password", key="api_pond")

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Sistema de Ponderación Analítica por R.A.</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Extracción curricular estricta y generación de documento oficial MINERD</div>', unsafe_allow_html=True)

# --- FORMULARIO ---
with st.form("form_ponderacion", clear_on_submit=False):
    
    st.markdown('<div class="section-title">📄 1. Fuente Curricular (PDF)</div>', unsafe_allow_html=True)
    archivo_pdf = st.file_uploader("Cargue el documento PDF oficial del diseño curricular", type=["pdf"], help="RECOMENDACIÓN: Sube solo las páginas del módulo para evitar que la IA se sature.")
    
    st.markdown('<div class="section-title">🏛️ 2. Datos Institucionales</div>', unsafe_allow_html=True)
    col_inst1, col_inst2 = st.columns(2)
    with col_inst1:
        politecnico = st.text_input("Nombre del Politécnico", value="Politécnico Salesiano Arquides Calderón")
        docente = st.text_input("Nombre del Docente", value="Ing. Bernardo Antonio Hernández Batista")
    with col_inst2:
        ano_escolar = st.text_input("Año Escolar", placeholder="Ej: 2026-2027")
        
    st.markdown('<div class="section-title">📝 3. Parámetros de Ponderación</div>', unsafe_allow_html=True)
    modulo = st.text_input("Nombre del Módulo Formativo", placeholder="Ej: MF_358_3: Impuestos al consumo y a vehículos de motor")
    
    col1, col2 = st.columns(2)
    with col1:
        semanas_totales = st.number_input("Semanas Totales", min_value=1, max_value=50, value=38)
    with col2:
        cantidad_ra = st.number_input("Cantidad Exacta de R.A.", min_value=1, max_value=20, value=6, help="¿Cuántos R.A. tiene el módulo en total?")
    
    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Procesar Ponderación y Generar Documento")

# --- LÓGICA CORE ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la barra lateral.")
    elif not archivo_pdf or not modulo or not ano_escolar:
        st.warning("📝 Por favor, carga el PDF y completa el Módulo Formativo y el Año Escolar.")
    else:
        with st.spinner(f'🧠 Analizando currículo en formato JSON nativo con {modelo_seleccionado}...'):
            try:
                # 1. Extracción PDF
                pdf_reader = PyPDF2.PdfReader(archivo_pdf)
                texto_curriculo = "".join([pagina.extract_text() for pagina in pdf_reader.pages])
                if len(texto_curriculo) > 80000: texto_curriculo = texto_curriculo[:80000]

                # 2. PROMPT JSON NATIVO (Sin Elementos de Capacidad)
                prompt_maestro = f"""Actúa como un experto en Diseño Curricular y Planificación Didáctica (ETP). Estructura la ponderación del módulo de forma estricta y literal.

INSUMOS:
1. Módulo Formativo: {modulo}
2. Semanas disponibles: {semanas_totales}
3. Cantidad EXACTA de Resultados de Aprendizaje (R.A.) en el módulo: {cantidad_ra}

REGLAS INQUEBRANTABLES:
1. LEY DE EXTRACCIÓN EXACTA: Debes procesar exactamente {cantidad_ra} Resultados de Aprendizaje (R.A.).
2. LEY DE LITERALIDAD: Extrae TODOS los R.A., C.R., C.E. y Contenidos. NO RESUMAS. Cópialos exactamente.
3. CÁLCULO PROPORCIONAL: Asigna Valor (%) y Semanas a cada R.A. evaluando su carga (cantidad de CR y Contenidos). La suma total de % debe ser 100 y de semanas {semanas_totales}.
4. FORMATO: Utiliza un guion "-" seguido de un espacio para listar elementos dentro de cada bloque.

FORMATO DE SALIDA ESTRICTO (JSON NATIVO):
Devuelve ÚNICA Y EXCLUSIVAMENTE un objeto JSON válido. NO envuelvas en bloques markdown (```json). La estructura DEBE ser exactamente esta:

{{
  "RESUMEN": "[Resumen analítico detallando cómo distribuiste los porcentajes y semanas]",
  "TABLA_GENERAL": [
    {{
      "RA": "[RA Completo]",
      "VALOR": "[Valor %]",
      "SEMANAS": "[Semanas]"
    }}
  ],
  "MATRICES": [
    {{
      "TEXTO": "[Ej: RA.8.1: Texto del RA... (Nivel de Bloom: Aplicar)]",
      "CONTENIDOS": "[- Contenido 1...\\n- Contenido 2...]",
      "CR": "[- CR8.1.1...\\n- CR8.1.2...]",
      "CE": "[- CE8.1.1...\\n- CE8.1.2...]",
      "VALOR": "[Ej: 35%]",
      "SEMANAS": "[Ej: 13]"
    }}
  ]
}}
NOTA: El arreglo "TABLA_GENERAL" y el arreglo "MATRICES" deben tener exactamente {cantidad_ra} objetos cada uno.

Documento a analizar:
{texto_curriculo}
"""
                # 3. Petición a la IA 
                respuesta_ia = ""
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_con_reintento(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    client = OpenAI(api_key=api_key_usuario)
                    response = client.chat.completions.create(
                        model=modelo_seleccionado,
                        messages=[{"role": "user", "content": prompt_maestro}],
                        temperature=0.0
                    )
                    respuesta_ia = response.choices[0].message.content

                # Limpieza de Markdowns accidentales
                respuesta_limpia = respuesta_ia.strip()
                if respuesta_limpia.startswith("```json"): respuesta_limpia = respuesta_limpia[7:]
                elif respuesta_limpia.startswith("```"): respuesta_limpia = respuesta_limpia[3:]
                if respuesta_limpia.endswith("```"): respuesta_limpia = respuesta_limpia[:-3]
                respuesta_limpia = respuesta_limpia.strip()

                # 4. PARSEO JSON INFALIBLE
                try:
                    datos_json = json.loads(respuesta_limpia)
                    resumen_texto = datos_json.get("RESUMEN", "Resumen no generado.")
                    lineas_tabla_gen = datos_json.get("TABLA_GENERAL", [])
                    matrices_items = datos_json.get("MATRICES", [])
                    
                    if not matrices_items:
                        st.error("❌ El JSON se procesó pero la matriz de R.A. está vacía.")
                        raise ValueError("Matriz vacía")
                        
                except json.JSONDecodeError as e:
                    st.error("❌ Error grave: La IA no devolvió un formato JSON válido.")
                    with st.expander("🔍 Ver respuesta cruda de la IA (Para depuración)"):
                        st.write("Si el texto de abajo se corta abruptamente, recorta tu PDF a solo las páginas del módulo y vuelve a subirlo.")
                        st.text(respuesta_limpia)
                        st.text(f"Detalle del error JSON: {e}")
                    st.stop()

                if len(matrices_items) == cantidad_ra:
                    st.success(f"✅ ¡Perfecto! Se desglosaron exactamente los {cantidad_ra} Resultados de Aprendizaje mediante JSON.")
                else:
                    st.warning(f"⚠️ Se solicitaron {cantidad_ra} R.A., pero se extrajeron {len(matrices_items)}.")

                # 5. CONSTRUCCIÓN DEL DOCUMENTO WORD
                doc = Document()
                doc.styles['Normal'].font.name = 'Calibri'
                doc.styles['Normal'].font.size = Pt(11)

                sections = doc.sections
                for section in sections:
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)

                # Encabezado Oficial Institucional
                p_titulo = doc.add_paragraph()
                run_tit = p_titulo.add_run("Planificación Didáctica - Módulo Formativo\n")
                run_tit.bold = True
                run_tit.font.size = Pt(14)
                
                p_datos = doc.add_paragraph()
                p_datos.add_run(f"Centro Educativo: {politecnico}\n").bold = True
                p_datos.add_run(f"Docente: {docente}\n").bold = True
                p_datos.add_run(f"Año Escolar: {ano_escolar}\n").bold = True
                p_datos.add_run(f"Módulo Formativo: {modulo}\n").bold = True
                p_datos.add_run(f"Duración Total: {semanas_totales} Semanas").bold = True

                # Sección: Resumen Analítico
                doc.add_heading('Resumen Analítico', level=2)
                doc.add_paragraph(str(resumen_texto))

                # Sección: Tabla General
                doc.add_heading('Tabla General', level=2)
                if lineas_tabla_gen:
                    tabla_gen = doc.add_table(rows=1, cols=3)
                    tabla_gen.style = 'Table Grid'
                    hdr_cells = tabla_gen.rows[0].cells
                    
                    encabezados_gen = ["Resultado de Aprendizaje (R.A.)", "Valor (%)", "Semanas"]
                    for i, texto in enumerate(encabezados_gen):
                        hdr_cells[i].text = texto
                        hdr_cells[i].paragraphs[0].runs[0].bold = True

                    for item in lineas_tabla_gen:
                        row_cells = tabla_gen.add_row().cells
                        row_cells[0].text = str(item.get("RA", "")).replace("**", "")
                        row_cells[1].text = str(item.get("VALOR", "")).replace("**", "")
                        row_cells[2].text = str(item.get("SEMANAS", "")).replace("**", "")
                    
                    # Fila de Totales
                    row_cells = tabla_gen.add_row().cells
                    row_cells[0].text = "TOTAL"
                    row_cells[1].text = "100%"
                    row_cells[2].text = str(semanas_totales)
                    for i in range(3):
                        row_cells[i].paragraphs[0].runs[0].bold = True

                # Sección: Matriz de Desarrollo
                doc.add_heading('Matriz de Desarrollo', level=2)
                
                # Se eliminó "Elementos de Capacidad (EC creados)" de los encabezados (ahora son 6 columnas)
                encabezados_matriz = [
                    "R.A. (Literal y Nivel de Bloom)", 
                    "Contenidos Asociados (Literales)", 
                    "C.R. Literales Asociados", 
                    "C.E. Literales Asociados", 
                    "Valor (%)", 
                    "Semanas"
                ]

                for item in matrices_items:
                    try:
                        # Tabla ajustada a 6 columnas
                        tabla_m = doc.add_table(rows=2, cols=6)
                        tabla_m.style = 'Table Grid'
                        
                        for i, titulo in enumerate(encabezados_matriz):
                            tabla_m.cell(0, i).text = titulo
                            tabla_m.cell(0, i).paragraphs[0].runs[0].bold = True
                            
                        # Datos actualizados sin la columna EC
                        datos_fila = [
                            str(item.get("TEXTO", "")), 
                            str(item.get("CONTENIDOS", "")), 
                            str(item.get("CR", "")), 
                            str(item.get("CE", "")), 
                            str(item.get("VALOR", "")), 
                            str(item.get("SEMANAS", ""))
                        ]
                        
                        for i, dato in enumerate(datos_fila):
                            tabla_m.cell(1, i).text = dato.replace("**", "")
                            
                        doc.add_paragraph() 
                        
                    except Exception as e:
                        print(f"Error procesando un bloque RA: {e}")

                # --- FIRMA INSTITUCIONAL Y DOCENTE ---
                doc.add_paragraph("\n\n")
                p_firmas = doc.add_paragraph()
                p_firmas.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_firmas.add_run("______________________________________________________________\n")
                p_firmas.add_run(f"{docente}\n").bold = True
                p_firmas.add_run("Docente de Módulos Formativos Educación Técnico Profesional\n")
                p_firmas.add_run(f"{politecnico}\n")

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Descargar Documento Oficial de Ponderación (.docx)",
                    data=buffer,
                    file_name=f"Ponderacion_RA_{modulo[:15]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary" 
                )
                
            except ResourceExhausted:
                st.error("❌ Se alcanzó el límite de API. Espera unos momentos antes de reintentar.")
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")