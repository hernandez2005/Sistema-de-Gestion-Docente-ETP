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

# --- LLAMADAS A API CON JSON FORZADO ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_json(api_key, modelo, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    # response_mime_type fuerza a Gemini a responder 100% en JSON válido
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=8192, 
            temperature=0.0,
            response_mime_type="application/json"
        )
    )
    return respuesta.text

def solicitar_openai_json(api_key, modelo, prompt):
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=modelo,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.0,
        response_format={"type": "json_object"} # Fuerza a OpenAI a responder 100% en JSON
    )
    return response.choices[0].message.content

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("⚡ Núcleo de Procesamiento")
    proveedor_ia = st.selectbox("Motor Analítico:", ["Google Gemini", "OpenAI (ChatGPT)"], key="prov_items")
    
    if proveedor_ia == "Google Gemini":
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"], key="mod_items")
    else:
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gpt-4o-mini", "gpt-3.5-turbo", "gpt-4o"], key="mod_items2")
        
    api_key_usuario = st.text_input("Clave de Autenticación (API Key):", type="password", key="api_items")

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Generador de Banco de Ítems y Pruebas Teóricas</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Evaluación escrita objetiva con anclaje estricto al contenido curricular ETP</div>', unsafe_allow_html=True)

# --- FORMULARIO ---
with st.form("form_banco_items", clear_on_submit=False):
    
    st.markdown('<div class="section-title">📄 1. Fuente Curricular (PDF)</div>', unsafe_allow_html=True)
    archivo_pdf = st.file_uploader("Cargue el documento PDF oficial del módulo", type=["pdf"], help="Sube las páginas específicas del contenido.")
    
    st.markdown('<div class="section-title">🏛️ 2. Datos Institucionales y de la Prueba</div>', unsafe_allow_html=True)
    col_inst1, col_inst2 = st.columns(2)
    with col_inst1:
        politecnico = st.text_input("Centro Educativo", value="Politécnico Salesiano Arquides Calderón")
        docente = st.text_input("Nombre del Docente", value="Ing. Bernardo Antonio Hernández Batista")
        asignatura = st.text_input("Módulo / Asignatura", placeholder="Ej: Ofimática o Redes LAN")
    with col_inst2:
        titulo_prueba = st.text_input("Título de la Evaluación", value="Prueba Teórica de Validación de Competencias")
        valor_total = st.number_input("Puntuación Total de la Prueba", min_value=10, max_value=100, value=100)
        fecha = st.date_input("Fecha de Aplicación")
        
    st.markdown('<div class="section-title">📝 3. Estructura de Ítems a Generar</div>', unsafe_allow_html=True)
    col_q1, col_q2, col_q3, col_q4 = st.columns(4)
    with col_q1:
        cant_mc = st.number_input("Opción Múltiple", min_value=0, max_value=20, value=4)
    with col_q2:
        cant_ci = st.number_input("Correcto / Incorrecto (C/I)", min_value=0, max_value=20, value=4)
    with col_q3:
        cant_match = st.number_input("Apareamiento", min_value=0, max_value=10, value=3)
    with col_q4:
        cant_tech = st.number_input("Análisis Técnico", min_value=0, max_value=10, value=2)
        
    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Generar Banco de Ítems Estricto (Word)")

# --- LÓGICA CORE ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la barra lateral.")
    elif not archivo_pdf or not asignatura:
        st.warning("📝 Por favor, carga el PDF curricular y define la asignatura.")
    else:
        with st.spinner(f'🧠 Leyendo documento y aplicando anclaje estricto con {modelo_seleccionado}...'):
            try:
                # 1. Extracción del PDF
                pdf_reader = PyPDF2.PdfReader(archivo_pdf)
                texto_curriculo = "".join([pagina.extract_text() for pagina in pdf_reader.pages])
                if len(texto_curriculo) > 80000: texto_curriculo = texto_curriculo[:80000]

                # 2. PROMPT MAESTRO CON ANCLAJE TEXTUAL ESTRICTO (CERO ALUCINACIONES)
                prompt_maestro = f"""Actúa como un Evaluador Educativo Máster y Especialista en la Educación Técnico Profesional (ETP). 

OBJETIVO CRÍTICO:
Diseña una Prueba Teórica y Técnica basándote EXCLUSIVAMENTE Y DE FORMA LITERAL en el texto del currículo proporcionado abajo. 

REGLAS ABSOLUTAS (CERO ALUCINACIONES Y CERO CONOCIMIENTO EXTERNO):
1. PROHIBIDO INVENTAR: No utilices conocimiento general ni externo sobre la materia. Si un concepto, definición o término no se encuentra explícitamente en el texto del PDF adjunto, no lo utilices bajo ninguna circunstancia.
2. EXTRACCIÓN DE CONTENIDO: Extrae los conceptos, definiciones, normativas y procedimientos descritos en el documento y conviértelos directamente en ítems de evaluación.
3. FORMATO DE OPCIÓN MÚLTIPLE ({cant_mc} preguntas): Cada pregunta debe derivar de un párrafo real del texto. Incluye 4 opciones (A, B, C, D) y la letra de la respuesta correcta.
4. FORMATO CORRECTO / INCORRECTO (C/I) ({cant_ci} ítems): Las afirmaciones deben ser extractos parafraseados directamente del documento (algunas verdaderas y otras alteradas sutilmente para que sean falsas).
5. FORMATO DE APAREAMIENTO ({cant_match} ítems): Relaciona conceptos y definiciones reales extraídas textualmente del PDF.
6. FORMATO DE ANÁLISIS TÉCNICO ({cant_tech} preguntas): Plantea un caso o análisis basado exclusivamente en los procesos técnicos reales que aparecen en el documento.

FORMATO DE SALIDA ESTRICTO (JSON NATIVO OBLIGATORIO):
Devuelve un objeto JSON válido con la siguiente estructura exacta:
{{
  "MULTIPLE_CHOICE": [
    {{
      "PREGUNTA": "Texto basado estrictamente en el PDF...",
      "OPCIONES": ["A) Opción 1", "B) Opción 2", "C) Opción 3", "D) Opción 4"],
      "RESPUESTA_CORRECTA": "A"
    }}
  ],
  "CORRECT_INCORRECT": [
    {{
      "ENUNCIADO": "Afirmación técnica extraída del texto...",
      "ES_CORRECTO": true
    }}
  ],
  "MATCHING": [
    {{
      "PREMISA": "Concepto real del PDF...",
      "RESPUESTA": "Definición real del PDF..."
    }}
  ],
  "TECHNICAL_ANALYSIS": [
    "Pregunta de análisis basada en los procesos del documento..."
  ]
}}

DOCUMENTO CURRICULAR OFICIAL A ANALIZAR (BASE ÚNICA Y OBLIGATORIA):
{texto_curriculo}
"""
                # 3. Petición a la IA con JSON garantizado por API
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_json(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    respuesta_ia = solicitar_openai_json(api_key_usuario, modelo_seleccionado, prompt_maestro)

                # Parseo directo sin fallos gracias al motor de API
                datos = json.loads(respuesta_ia)

                mc_list = datos.get("MULTIPLE_CHOICE", [])
                ci_list = datos.get("CORRECT_INCORRECT", [])
                match_list = datos.get("MATCHING", [])
                tech_list = datos.get("TECHNICAL_ANALYSIS", [])

                # --- CONSTRUCCIÓN DEL DOCUMENTO WORD ---
                doc = Document()
                doc.styles['Normal'].font.name = 'Calibri'
                doc.styles['Normal'].font.size = Pt(11)

                sections = doc.sections
                for section in sections:
                    section.left_margin = Inches(0.75)
                    section.right_margin = Inches(0.75)

                # Encabezado Institucional de la Evaluación
                p_encabezado = doc.add_paragraph()
                p_encabezado.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_inst = p_encabezado.add_run(f"{politecnico}\n")
                run_inst.bold = True
                run_inst.font.size = Pt(13)
                
                run_tit = p_encabezado.add_run(f"{titulo_prueba.upper()}\n")
                run_tit.bold = True
                run_tit.font.size = Pt(12)
                
                doc.add_paragraph(f"Asignatura / Módulo: {asignatura} | Docente: {docente} | Fecha: {fecha.strftime('%d/%m/%Y')}")
                doc.add_paragraph(f"Valor Total: {valor_total} Puntos | Calificación Obtenida: _________")
                doc.add_paragraph("Nombre del Estudiante: _________________________________________________ | Sección: _______")
                doc.add_paragraph("_" * 70)

                # Instrucciones Generales
                p_inst = doc.add_paragraph()
                p_inst.add_run("Instrucciones Generales: ").bold = True
                p_inst.add_run("Responda cada sección basándose estrictamente en los contenidos y conceptos presentados en el material de estudio oficial del módulo.")

                # SECCIÓN I: OPCIÓN MÚLTIPLE
                if mc_list:
                    doc.add_heading("I. Selección Múltiple", level=2)
                    doc.add_paragraph("Instrucción: Seleccione la opción correcta para cada uno de los siguientes enunciados.")
                    for idx, item in enumerate(mc_list):
                        p_q = doc.add_paragraph()
                        p_q.add_run(f"1.{idx+1} ").bold = True
                        p_q.add_run(str(item.get("PREGUNTA", "")))
                        
                        opciones = item.get("OPCIONES", [])
                        for opc in opciones:
                            p_opc = doc.add_paragraph(str(opc))
                            p_opc.paragraph_format.left_indent = Inches(0.25)
                        doc.add_paragraph()

                # SECCIÓN II: CORRECTO E INCORRECTO (C/I)
                if ci_list:
                    doc.add_heading("II. Criterio de Correcto (C) e Incorrecto (I)", level=2)
                    doc.add_paragraph("Instrucción: Escriba una 'C' si la afirmación es correcta o una 'I' si es incorrecta en el espacio indicado a la izquierda.")
                    for idx, item in enumerate(ci_list):
                        p_ci = doc.add_paragraph()
                        p_ci.add_run("_____ ").bold = True
                        p_ci.add_run(f"2.{idx+1}. {str(item.get('ENUNCIADO', ''))}")
                    doc.add_paragraph()

                # SECCIÓN III: APAREAMIENTO
                if match_list:
                    doc.add_heading("III. Apareamiento", level=2)
                    doc.add_paragraph("Instrucción: Relacione los conceptos de la columna izquierda con la definición correspondiente de la derecha.")
                    
                    tabla_match = doc.add_table(rows=1, cols=2)
                    tabla_match.style = 'Table Grid'
                    hdr = tabla_match.rows[0].cells
                    hdr[0].text = "Premisas / Conceptos"
                    hdr[1].text = "Términos / Definiciones"
                    hdr[0].paragraphs[0].runs[0].bold = True
                    hdr[1].paragraphs[0].runs[0].bold = True

                    for item in match_list:
                        row = tabla_match.add_row().cells
                        row[0].text = str(item.get("PREMISA", ""))
                        row[1].text = str(item.get("RESPUESTA", ""))
                    
                    doc.add_paragraph()

                # SECCIÓN IV: ANÁLISIS TÉCNICO
                if tech_list:
                    doc.add_heading("IV. Análisis Técnico y Resolución de Casos", level=2)
                    doc.add_paragraph("Instrucción: Analice detalladamente cada planteamiento y desarrolle su respuesta fundamentada en los procesos del módulo.")
                    for idx, item in enumerate(tech_list):
                        p_t = doc.add_paragraph()
                        p_t.add_run(f"4.{idx+1} ").bold = True
                        p_t.add_run(str(item))
                        doc.add_paragraph("\n" + "_" * 75 + "\n" + "_" * 75 + "\n")

                # --- SOLUCIONARIO EXCLUSIVO PARA EL DOCENTE (PÁGINA FINAL) ---
                doc.add_page_break()
                p_sol = doc.add_paragraph()
                p_sol.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_sol_title = p_sol.add_run("🔑 SOLUCIONARIO OFICIAL (EXCLUSIVO PARA EL DOCENTE)\n")
                run_sol_title.bold = True
                run_sol_title.font.size = Pt(13)
                p_sol.add_run("Claves de corrección basadas estrictamente en el contenido cargado.\n")
                doc.add_paragraph("_" * 70)

                if mc_list:
                    doc.add_heading("Respuestas - Sección I: Selección Múltiple", level=3)
                    for idx, item in enumerate(mc_list):
                        p_ans = doc.add_paragraph()
                        p_ans.add_run(f"Pregunta {idx+1}: ").bold = True
                        p_ans.add_run(f"Respuesta Correcta -> {item.get('RESPUESTA_CORRECTA', 'N/A')}")

                if ci_list:
                    doc.add_heading("Respuestas - Sección II: Correcto (C) e Incorrecto (I)", level=3)
                    for idx, item in enumerate(ci_list):
                        p_ans_ci = doc.add_paragraph()
                        estado = "Correcto (C)" if item.get('ES_CORRECTO') else "Incorrecto (I)"
                        p_ans_ci.add_run(f"Afirmación {idx+1}: ").bold = True
                        p_ans_ci.add_run(f"Respuesta -> {estado}")

                if match_list:
                    doc.add_heading("Respuestas - Sección III: Apareamiento", level=3)
                    for idx, item in enumerate(match_list):
                        p_ans2 = doc.add_paragraph()
                        p_ans2.add_run(f"Ítem {idx+1}: ").bold = True
                        p_ans2.add_run(f"{item.get('PREMISA')}  <--->  {item.get('RESPUESTA')}")

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.success("✅ ¡Banco de Ítems estrictamente anclado generado con éxito!")
                
                st.download_button(
                    label="📥 Descargar Prueba Teórica (.docx)",
                    data=buffer,
                    file_name=f"Prueba_Teorica_Estricta_{asignatura[:10]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary" 
                )
                
            except ResourceExhausted:
                st.error("❌ Se alcanzó el límite de API.")
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")