import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import PyPDF2
import json
import re
from google.api_core.exceptions import ResourceExhausted
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# --- ESTILOS CSS AVANZADOS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #F8FAFC; color: #0F172A; }
    .main-header { font-size: 2.4rem; font-weight: 800; color: #1E293B; text-align: center; margin-bottom: 5px; line-height: 1.2; letter-spacing: -0.02em; }
    .sub-header { text-align: center; color: #475569; font-size: 1.15rem; font-weight: 400; margin-bottom: 40px; }
    [data-testid="stForm"], .st-key-panel_extraccion { background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05); padding: 35px; margin-bottom: 25px; }
    .section-title { color: #0F172A; font-weight: 700; font-size: 1.3rem; border-bottom: 2px solid #E2E8F0; padding-bottom: 10px; margin-top: 25px; margin-bottom: 20px; }
    .academic-note { background-color: #F0F9FF; border-left: 4px solid #0EA5E9; padding: 15px; border-radius: 4px; color: #0C4A6E; font-size: 0.95rem; margin-bottom: 20px; }
    div.stButton > button:first-child { background-color: #0F172A !important; color: #FFFFFF !important; border: none !important; border-radius: 6px !important; font-weight: 600 !important; padding: 12px 24px !important; width: 100%; transition: all 0.2s ease; }
    div.stButton > button:first-child:hover { background-color: #334155 !important; transform: translateY(-1px); }
</style>
""", unsafe_allow_html=True)

# --- 1. PARSEADOR INDESTRUCTIBLE PARA FASE 1 (Evita el error JSONDecodeError) ---
def parsear_texto_etiquetas(texto):
    """Lee el texto crudo de la IA usando etiquetas, 100% inmune a errores de formato."""
    resultado = {
        "COMPETENCIAS_ESPECIFICAS": [],
        "INDICADORES_LOGRO": [],
        "CONTENIDOS_CONCEPTUALES": [],
        "CONTENIDOS_PROCEDIMENTALES": [],
        "CONTENIDOS_ACTITUDINALES": []
    }
    
    seccion_actual = None
    
    for linea in texto.split('\n'):
        linea = linea.strip()
        if not linea:
            continue
            
        # Detectar delimitadores exactos
        if linea == "===COMPETENCIAS_ESPECIFICAS===":
            seccion_actual = "COMPETENCIAS_ESPECIFICAS"
            continue
        elif linea == "===INDICADORES_LOGRO===":
            seccion_actual = "INDICADORES_LOGRO"
            continue
        elif linea == "===CONTENIDOS_CONCEPTUALES===":
            seccion_actual = "CONTENIDOS_CONCEPTUALES"
            continue
        elif linea == "===CONTENIDOS_PROCEDIMENTALES===":
            seccion_actual = "CONTENIDOS_PROCEDIMENTALES"
            continue
        elif linea == "===CONTENIDOS_ACTITUDINALES===":
            seccion_actual = "CONTENIDOS_ACTITUDINALES"
            continue
        elif linea.startswith("==="):
            seccion_actual = None 
            continue
            
        # Guardar y limpiar viñetas
        if seccion_actual:
            item_limpio = re.sub(r'^[\-\*\•\d\.\s]+', '', linea).strip()
            if item_limpio:
                resultado[seccion_actual].append(item_limpio)
                
    return resultado

# --- 2. FUNCIÓN DE LIMPIEZA JSON PARA FASE 2 ---
def extraer_json_seguro(texto_ia):
    try: return json.loads(texto_ia, strict=False)
    except: pass 
    texto = texto_ia.strip()
    if texto.startswith("```json"): texto = texto[7:]
    elif texto.startswith("```"): texto = texto[3:]
    if texto.endswith("```"): texto = texto[:-3]
    match = re.search(r'(\{.*\})', texto, re.DOTALL)
    if match: texto = match.group(1)
    texto = re.sub(r',\s*([\]\}])', r'\1', texto)
    texto_plano = texto.replace('\n', ' ').replace('\r', '')
    return json.loads(texto_plano, strict=False)

# --- 3. MOTORES DE LLAMADA API (TEXTO LIBRE PARA FASE 1, JSON PARA FASE 2) ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini(api_key, modelo, prompt, es_json=False):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    config = genai.types.GenerationConfig(max_output_tokens=8192, temperature=0.15)
    if es_json: config.response_mime_type = "application/json"
    return model.generate_content(prompt, generation_config=config).text

def solicitar_openai(api_key, modelo, prompt, es_json=False):
    client = OpenAI(api_key=api_key)
    kwargs = {"model": modelo, "messages": [{"role": "user", "content": prompt}], "temperature": 0.15}
    if es_json: kwargs["response_format"] = {"type": "json_object"}
    return client.chat.completions.create(**kwargs).choices[0].message.content

# --- PANEL DE CONFIGURACIÓN LATERAL ---
with st.sidebar:
    st.title("🔬 Arquitectura Curricular")
    proveedor_ia = st.selectbox("Motor de Inferencia:", ["Google Gemini", "OpenAI (ChatGPT)"], key="prov_acad")
    
    if proveedor_ia == "Google Gemini":
        modelo_seleccionado = st.selectbox("Modelo Neuronal:", ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"], key="mod_acad")
    else:
        modelo_seleccionado = st.selectbox("Modelo Neuronal:", ["gpt-4o-mini", "gpt-3.5-turbo", "gpt-4o"], key="mod_acad2")
        
    api_key_usuario = st.text_input(
        "Clave de Autenticación (API Key):", 
        type="password", 
        value=st.session_state.get("api_key_global", ""), 
        key="api_acad_sinc"
    )
    st.session_state.api_key_global = api_key_usuario

# --- ENCABEZADO ACADÉMICO ---
st.markdown('<div class="main-header">Diseño Curricular por Competencias (MINERD)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Alineación constructiva, transposición didáctica y evaluación auténtica</div>', unsafe_allow_html=True)

# --- ESTADOS DE SESIÓN ---
if "malla_dividida" not in st.session_state:
    st.session_state.malla_dividida = {}
if "datos_malla_cargada" not in st.session_state:
    st.session_state.datos_malla_cargada = False

# --- FASE 1: PARAMETRIZACIÓN DEL CONTEXTO Y EXÉGESIS CURRICULAR ---
st.markdown('<div class="section-title">I. Parametrización del Contexto Escolar</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    docente = st.text_input("Docente Titular", value=st.session_state.get("docente_nombre", "Ing. Bernardo Antonio Hernández Batista"))
    materia = st.selectbox("Disciplina Académica", ["Ciencias Sociales", "Lengua Española", "Matemática", "Ciencias de la Naturaleza", "Inglés", "Francéss", "Educación Artística", "Educación Física","Formación Integral Humana y Religiosa"])
    centro_educativo = st.text_input("Institución Educativa", value=st.session_state.get("docente_politecnico", "Liceo Secundario Ejemplo"))
with col2:
    grado = st.text_input("Grado", value="2do de Secundaria")
    seccion = st.text_input("Sección", value="B")
    ano_escolar = st.text_input("Año Lectivo", value="2026-2027")
    duracion = st.text_input("Temporalidad de la Unidad", value="5 semanas")

st.markdown('<div class="section-title">II. Problematización y Análisis del Diseño Curricular</div>', unsafe_allow_html=True)

archivo_pdf = st.file_uploader("Documento Base (Malla Curricular Oficial MINERD en PDF)", type=["pdf"], help="Cargue el referente prescriptivo para garantizar la alineación normativa.")
contexto_situacion = st.text_area(
    "Conflicto Cognitivo / Contexto Socioeducativo (Situación de Aprendizaje)", 
    height=100, 
    placeholder="Describa la fenomenología o problemática del entorno que detonará el aprendizaje significativo (Ej. Alto índice de contaminación por plásticos en el río local, requiriendo un plan de mitigación desde la química y el civismo)..."
)

if st.button("🔍 Fase 1: Realizar Exégesis Curricular (Extraer del PDF)"):
    if not api_key_usuario:
        st.error("🔒 Autenticación requerida. Ingrese su API Key en la barra lateral.")
    elif not archivo_pdf:
        st.warning("📝 Se requiere la carga del referente curricular (PDF) para el anclaje prescriptivo.")
    else:
        with st.spinner("🧠 Ejecutando análisis heurístico con lectura anti-errores del PDF..."):
            try:
                pdf_reader = PyPDF2.PdfReader(archivo_pdf)
                texto_pdf = "".join([pagina.extract_text() for pagina in pdf_reader.pages])
                if len(texto_pdf) > 70000: texto_pdf = texto_pdf[:70000]
                
                # NUEVO PROMPT ESTRUCTURADO CON ETIQUETAS (ADIÓS JSON)
                prompt_extraccion = f"""Actúa como un Doctor en Educación y Especialista en Currículo del MINERD. 
Realiza una extracción rigurosa de la matriz curricular de la disciplina {materia}.

REGLA DE ORO: NO USES FORMATO JSON.
Debes extraer la información y agruparla usando EXACTAMENTE las siguientes etiquetas. No escribas introducciones, solo usa las etiquetas y los items en forma de viñeta (-):

===COMPETENCIAS_ESPECIFICAS===
- [Competencia 1 extraída]
- [Competencia 2 extraída]

===INDICADORES_LOGRO===
- [Indicador 1 extraído]
- [Indicador 2 extraído]

===CONTENIDOS_CONCEPTUALES===
- [Concepto 1 extraído]

===CONTENIDOS_PROCEDIMENTALES===
- [Procedimiento 1 extraído]

===CONTENIDOS_ACTITUDINALES===
- [Actitud 1 extraída]

REFERENTE CURRICULAR A LEER:
{texto_pdf}
"""
                if proveedor_ia == "Google Gemini":
                    resp = solicitar_gemini(api_key_usuario, modelo_seleccionado, prompt_extraccion, es_json=False)
                else:
                    resp = solicitar_openai(api_key_usuario, modelo_seleccionado, prompt_extraccion, es_json=False)

                # Procesamiento indestructible
                st.session_state.malla_dividida = parsear_texto_etiquetas(resp)
                
                # Verificación de seguridad
                if not st.session_state.malla_dividida.get("COMPETENCIAS_ESPECIFICAS"):
                    st.warning("⚠️ El modelo de IA no encontró datos bajo las etiquetas requeridas. Intenta nuevamente o cambia el modelo.")
                else:
                    st.session_state.datos_malla_cargada = True
                    st.success("✓ Análisis heurístico completado. Elementos curriculares listos para selección.")
                    st.rerun()

            except Exception as e:
                st.error(f"⚠️ Error sistémico en el procesamiento del documento: {e}")

# --- FASE 2: SELECCIÓN Y DISEÑO INSTRUCCIONAL ---
if st.session_state.datos_malla_cargada and st.session_state.malla_dividida:
    st.markdown('<div class="section-title">III. Selección Curricular (Alineación Constructiva)</div>', unsafe_allow_html=True)
    st.markdown('<div class="academic-note"><b>Criterio Pedagógico:</b> Seleccione los elementos garantizando la coherencia interna entre la competencia esperada, el saber abordado y el indicador que evidenciará el logro.</div>', unsafe_allow_html=True)

    with st.expander("🎯 Dimensión Competencial (Competencias Específicas)", expanded=True):
        comps_seleccionadas = []
        for idx, comp in enumerate(st.session_state.malla_dividida.get("COMPETENCIAS_ESPECIFICAS", [])):
            if st.checkbox(comp, value=True, key=f"ce_{idx}"): comps_seleccionadas.append(comp)

    with st.expander("📋 Dimensión Evaluativa (Indicadores de Logro)", expanded=True):
        inds_seleccionados = []
        for idx, ind in enumerate(st.session_state.malla_dividida.get("INDICADORES_LOGRO", [])):
            if st.checkbox(ind, value=True, key=f"il_{idx}"): inds_seleccionados.append(ind)

    with st.expander("📚 Dimensión Epistemológica (Contenidos)", expanded=True):
        st.markdown("**Saberes Conceptuales**")
        cont_conc_sel = [c for idx, c in enumerate(st.session_state.malla_dividida.get("CONTENIDOS_CONCEPTUALES", [])) if st.checkbox(c, value=True, key=f"cc_{idx}")]
        st.markdown("**Saberes Procedimentales**")
        cont_proc_sel = [p for idx, p in enumerate(st.session_state.malla_dividida.get("CONTENIDOS_PROCEDIMENTALES", [])) if st.checkbox(p, value=True, key=f"cp_{idx}")]
        st.markdown("**Saberes Actitudinales**")
        cont_act_sel = [a for idx, a in enumerate(st.session_state.malla_dividida.get("CONTENIDOS_ACTITUDINALES", [])) if st.checkbox(a, value=True, key=f"ca_{idx}")]

    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("⚙️ Fase 2: Generar Diseño Instruccional (Exportación Oficial MINERD)"):
        if not comps_seleccionadas or not inds_seleccionados:
            st.warning("⚠️ La validación académica requiere al menos una Competencia y un Indicador de Logro.")
        else:
            with st.spinner("🧠 Orquestando la transposición didáctica y generando el documento de planificación..."):
                try:
                    prompt_final = f"""Actúa como un Doctor en Educación, Experto en Diseño Instruccional y Asesor Curricular del MINERD de la República Dominicana.

DATOS DEL ECOSISTEMA EDUCATIVO:
- Institución: {centro_educativo}
- Docente Titular: {docente}
- Nivel/Grado: {grado} | Sección: {seccion}
- Disciplina: {materia}
- Temporalidad: {duracion}
- Conflicto Cognitivo / Contexto: {contexto_situacion}

SELECCIÓN CURRICULAR DEL DOCENTE (Garantizar Alineación Constructiva):
- Competencias Específicas: {json.dumps(comps_seleccionadas, ensure_ascii=False)}
- Indicadores de Logro: {json.dumps(inds_seleccionados, ensure_ascii=False)}
- Saberes Conceptuales: {json.dumps(cont_conc_sel, ensure_ascii=False)}
- Saberes Procedimentales: {json.dumps(cont_proc_sel, ensure_ascii=False)}
- Saberes Actitudinales: {json.dumps(cont_act_sel, ensure_ascii=False)}

DIRECTRICES PEDAGÓGICAS DE ALTO NIVEL:
1. SITUACIÓN DE APRENDIZAJE: Redacta utilizando el enfoque socio-constructivista. Debe contener: Contexto, Problematización, Estrategia y Producto.
2. SECUENCIA DIDÁCTICA: Aplica la teoría del "Andamiaje" (Bruner).
   - INICIO: Recuperación de saberes previos y motivación.
   - DESARROLLO: Metodologías activas y procesamiento de la información.
   - CIERRE: Metacognición y evaluación.
3. EVALUACIÓN AUTÉNTICA: Diseña la evaluación (Diagnóstica, Formativa y Sumativa).
4. ESTRUCTURA: Adhiérete estrictamente al estándar prescriptivo.

REGLAS ESTRICTAS DE FORMATO JSON:
- NO uses saltos de línea literales (Enters/\\n) dentro de los valores de texto. Únelos con espacios.
- NO uses comillas dobles internas dentro de los textos. Usa comillas simples (' ') si necesitas citar.

FORMATO DE SALIDA ESTRICTO (JSON NATIVO OBLIGATORIO):
{{
  "TITULO_UNIDAD": "Título académico y motivador",
  "EJE_TRANSVERSAL": {{
    "NOMBRE": "Eje transversal MINERD aplicable",
    "DESCRIPTOR": "Justificación del eje vinculada a la situación de aprendizaje"
  }},
  "EFEMERIDE": "Efeméride escolar relevante",
  "SITUACION_APRENDIZAJE": "Redacción académica profunda (párrafo integrado)",
  "COMPETENCIAS_FUNDAMENTALES": ["Competencia 1", "Competencia 2", "Competencia 3"],
  "COMPETENCIAS_ESPECIFICAS_TEXTO": "Síntesis integradora de las competencias seleccionadas",
  "CONTENIDOS": {{
    "CONCEPTUALES": "Integración lógica de saberes conceptuales",
    "PROCEDIMENTALES": "Integración lógica de saberes procedimentales",
    "ACTITUDINALES": "Integración lógica de saberes actitudinales"
  }},
  "INDICADORES_LOGRO": ["Indicador 1...", "Indicador 2..."],
  "SECUENCIA_DIDACTICA": [
    {{"MOMENTO": "Inicio", "ACTIVIDADES": "Descripción metodológica..."}},
    {{"MOMENTO": "Desarrollo", "ACTIVIDADES": "Descripción metodológica..."}},
    {{"MOMENTO": "Cierre", "ACTIVIDADES": "Descripción metodológica..."}}
  ],
  "RECURSOS": ["Recurso 1...", "Recurso 2..."],
  "EVALUACION": [
    {{"TIPO": "Diagnóstica", "TECNICA": "Ej. Lluvia de ideas", "MOMENTO": "Inicio", "PONDERACION": "No aplica"}},
    {{"TIPO": "Formativa", "TECNICA": "Ej. Guía de observación", "MOMENTO": "Desarrollo", "PONDERACION": "30%"}},
    {{"TIPO": "Sumativa", "TECNICA": "Ej. Rúbrica", "MOMENTO": "Cierre", "PONDERACION": "70%"}}
  ],
  "PRODUCTO_FINAL": "Descripción del producto final"
}}
"""
                    if proveedor_ia == "Google Gemini":
                        respuesta_ia = solicitar_gemini(api_key_usuario, modelo_seleccionado, prompt_final, es_json=True)
                    else:
                        respuesta_ia = solicitar_openai(api_key_usuario, modelo_seleccionado, prompt_final, es_json=True)

                    datos_plan = extraer_json_seguro(respuesta_ia)

                    # --- RENDERIZADO DOCUMENTAL (PYTHON-DOCX) ---
                    doc = Document()
                    doc.styles['Normal'].font.name = 'Calibri'
                    doc.styles['Normal'].font.size = Pt(11)

                    sections = doc.sections
                    for section in sections:
                        section.left_margin = Inches(0.75)
                        section.right_margin = Inches(0.75)

                    p_enc = doc.add_paragraph()
                    p_enc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_enc.add_run("MINISTERIO DE EDUCACIÓN DE LA REPÚBLICA DOMINICANA\n").bold = True
                    p_enc.add_run("PLANIFICACIÓN POR UNIDAD DE APRENDIZAJE — NIVEL SECUNDARIO\n").bold = True

                    t_datos = doc.add_table(rows=3, cols=4)
                    t_datos.style = 'Table Grid'
                    t_datos.rows[0].cells[0].text = "Centro Educativo"
                    t_datos.rows[0].cells[1].text = centro_educativo
                    t_datos.rows[0].cells[2].text = "Año Escolar"
                    t_datos.rows[0].cells[3].text = ano_escolar
                    t_datos.rows[1].cells[0].text = "Docente Titular"
                    t_datos.rows[1].cells[1].text = docente
                    t_datos.rows[1].cells[2].text = "Grado"
                    t_datos.rows[1].cells[3].text = f"Sección {seccion}"
                    t_datos.rows[2].cells[0].text = "Área Curricular"
                    t_datos.rows[2].cells[1].text = materia
                    t_datos.rows[2].cells[2].text = "Título de la Unidad"
                    t_datos.rows[2].cells[3].text = datos_plan.get("TITULO_UNIDAD", "")
                    doc.add_paragraph()

                    doc.add_heading("1. Eje Transversal", level=2)
                    eje = datos_plan.get("EJE_TRANSVERSAL", {})
                    doc.add_paragraph(f"{eje.get('NOMBRE', '')}.\nDescriptor: {eje.get('DESCRIPTOR', '')}")

                    doc.add_heading("2. Efeméride del período", level=2)
                    doc.add_paragraph(datos_plan.get("EFEMERIDE", ""))

                    doc.add_heading("3. Situación de Aprendizaje", level=2)
                    p_sit = doc.add_paragraph(datos_plan.get("SITUACION_APRENDIZAJE", ""))
                    p_sit.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                    doc.add_heading("4. Competencias Fundamentales", level=2)
                    for cf in datos_plan.get("COMPETENCIAS_FUNDAMENTALES", []):
                        doc.add_paragraph(f"• {cf}", style='List Bullet')

                    doc.add_heading(f"5. Competencia Específica del Área ({materia})", level=2)
                    doc.add_paragraph(datos_plan.get("COMPETENCIAS_ESPECIFICAS_TEXTO", ""))

                    doc.add_heading("6. Contenidos Curriculares", level=2)
                    cont = datos_plan.get("CONTENIDOS", {})
                    t_cont = doc.add_table(rows=2, cols=3)
                    t_cont.style = 'Table Grid'
                    t_cont.rows[0].cells[0].text = "Conceptuales"
                    t_cont.rows[0].cells[1].text = "Procedimentales"
                    t_cont.rows[0].cells[2].text = "Actitudinales"
                    t_cont.rows[1].cells[0].text = str(cont.get("CONCEPTUALES", ""))
                    t_cont.rows[1].cells[1].text = str(cont.get("PROCEDIMENTALES", ""))
                    t_cont.rows[1].cells[2].text = str(cont.get("ACTITUDINALES", ""))
                    doc.add_paragraph()

                    doc.add_heading("7. Indicadores de Logro", level=2)
                    for ind in datos_plan.get("INDICADORES_LOGRO", []):
                        doc.add_paragraph(f"• {ind}", style='List Bullet')

                    doc.add_heading("8. Secuencia Didáctica", level=2)
                    t_sec = doc.add_table(rows=1, cols=2)
                    t_sec.style = 'Table Grid'
                    t_sec.rows[0].cells[0].text = "Momento"
                    t_sec.rows[0].cells[1].text = "Estrategias y Actividades"
                    for sec_item in datos_plan.get("SECUENCIA_DIDACTICA", []):
                        row = t_sec.add_row().cells
                        row[0].text = str(sec_item.get("MOMENTO", ""))
                        row[1].text = str(sec_item.get("ACTIVIDADES", ""))
                    doc.add_paragraph()

                    doc.add_heading("9. Recursos Didácticos", level=2)
                    for rec in datos_plan.get("RECURSOS", []):
                        doc.add_paragraph(f"• {rec}", style='List Bullet')

                    doc.add_heading("10. Plan de Evaluación", level=2)
                    t_eval = doc.add_table(rows=1, cols=4)
                    t_eval.style = 'Table Grid'
                    t_eval.rows[0].cells[0].text = "Función / Tipo"
                    t_eval.rows[0].cells[1].text = "Técnica e Instrumento"
                    t_eval.rows[0].cells[2].text = "Momento"
                    t_eval.rows[0].cells[3].text = "Ponderación"
                    for ev in datos_plan.get("EVALUACION", []):
                        row = t_eval.add_row().cells
                        row[0].text = str(ev.get("TIPO", ""))
                        row[1].text = str(ev.get("TECNICA", ""))
                        row[2].text = str(ev.get("MOMENTO", ""))
                        row[3].text = str(ev.get("PONDERACION", ""))
                    doc.add_paragraph()

                    doc.add_heading("11. Desempeño / Producto Final", level=2)
                    doc.add_paragraph(datos_plan.get("PRODUCTO_FINAL", ""))

                    doc.add_paragraph("\n\n____________________________________       ____________________________________")
                    doc.add_paragraph("Firma del Docente Titular                                       Firma Coordinación / Dirección")

                    buffer = BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                    
                    st.success("✅ ¡Diseño Curricular y Planificación Oficial generada con rigor académico!")
                    st.download_button(
                        label="📥 Exportar Planificación (Formato Word MINERD)",
                        data=buffer,
                        file_name=f"Diseño_Curricular_{materia}_{grado[:3]}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )

                except json.JSONDecodeError as e:
                    st.error(f"⚠️ Error de Formato JSON en Fase 2. Intenta presionar el botón nuevamente.")
                except Exception as e:
                    st.error(f"⚠️ Error en la orquestación del documento: {e}")