import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
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

# --- FUNCIÓN DE LIMPIEZA JSON (100% FORTIFICADA) ---
def extraer_json_seguro(texto_ia):
    """Extrae y parsea JSON a prueba de balas, tolerando errores comunes de la IA."""
    try:
        return json.loads(texto_ia, strict=False)
    except json.JSONDecodeError:
        pass 

    texto = texto_ia.strip()
    if texto.startswith("```json"): texto = texto[7:]
    elif texto.startswith("```"): texto = texto[3:]
    if texto.endswith("```"): texto = texto[:-3]
    texto = texto.strip()

    match = re.search(r'(\{.*\})', texto, re.DOTALL)
    if match:
        texto = match.group(1)

    try:
        return json.loads(texto, strict=False)
    except json.JSONDecodeError:
        pass

    texto = re.sub(r',\s*([\]\}])', r'\1', texto)
    texto_plano = texto.replace('\n', ' ').replace('\r', '')
    
    return json.loads(texto_plano, strict=False)

# --- NÚCLEO DE INFERENCIA PEDAGÓGICA ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_json(api_key, modelo, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=8192, 
            temperature=0.15,
            response_mime_type="application/json"
        )
    )
    return respuesta.text

def solicitar_openai_json(api_key, modelo, prompt):
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=modelo,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.15,
        response_format={"type": "json_object"}
    )
    return response.choices[0].message.content

# --- PANEL DE CONFIGURACIÓN LATERAL ---
with st.sidebar:
    st.title("🔬 Arquitectura Curricular")
    proveedor_ia = st.selectbox("Motor de Inferencia:", ["Google Gemini", "OpenAI (ChatGPT)"], key="prov_dia")
    
    if proveedor_ia == "Google Gemini":
        modelo_seleccionado = st.selectbox("Modelo Neuronal:", ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"], key="mod_dia")
    else:
        modelo_seleccionado = st.selectbox("Modelo Neuronal:", ["gpt-4o-mini", "gpt-3.5-turbo", "gpt-4o"], key="mod_dia2")
        
    api_key_usuario = st.text_input(
        "Clave de Autenticación (API Key):", 
        type="password", 
        value=st.session_state.get("api_key_global", ""), 
        key="api_dia_sinc"
    )
    st.session_state.api_key_global = api_key_usuario

# --- ENCABEZADO ACADÉMICO ---
st.markdown('<div class="main-header">Planificación Diaria Académica (MINERD)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Desglose micro-curricular: Secuencias didácticas, tiempos y evaluación diaria</div>', unsafe_allow_html=True)

# --- FORMULARIO DE PARAMETRIZACIÓN ---
with st.form("form_diario_acad"):
    st.markdown('<div class="section-title">🏫 I. Contextualización Institucional</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        docente = st.text_input("Docente Titular", value=st.session_state.get("docente_nombre", ""))
        materia = st.selectbox("Área Curricular", ["Ciencias Sociales", "Lengua Española", "Matemática", "Ciencias de la Naturaleza", "Inglés", "Francés", "Educación Artística", "Educación Física", "Formación Integral Humana y Religiosa"])
        centro_educativo = st.text_input("Centro Educativo", value=st.session_state.get("docente_politecnico", ""))
    with col2:
        grado = st.text_input("Grado", value="2do de Secundaria")
        seccion = st.text_input("Sección", value="B")
        ano_escolar = st.text_input("Año Escolar", value="2026-2027")
    
    st.markdown('<div class="section-title">📚 II. Parámetros de la Unidad</div>', unsafe_allow_html=True)
    unidad_aprendizaje = st.text_input("Título de la Unidad de Aprendizaje", placeholder="Ej: Riesgos naturales en nuestra comunidad")
    
    col3, col4 = st.columns([1, 1])
    with col3:
        duracion_total = st.text_input("Duración total", placeholder="Ej: 5 semanas / 6 sesiones")
    with col4:
        cantidad_sesiones = st.number_input("Número de Sesiones a Generar", min_value=1, max_value=10, value=3, help="Cantidad de planes diarios que la IA diseñará.")

    situacion_aprendizaje = st.text_area(
        "Situación de Aprendizaje (Referencia de toda la unidad)", 
        height=100, 
        placeholder="Ej: La Defensa Civil de tu municipio ha solicitado a los centros educativos elaborar un mapa de riesgos..."
    )
    
    temas_indicadores = st.text_area(
        "Temas o Indicadores a trabajar (Insumo para la IA)", 
        height=80, 
        placeholder="Escribe brevemente los temas o indicadores que deben cubrirse en estas sesiones para guiar a la IA."
    )

    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Generar Plan Diario Oficial MINERD (Word)")

# --- NÚCLEO DE PROCESAMIENTO ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Autenticación requerida. Ingrese su API Key en la barra lateral.")
    elif not unidad_aprendizaje or not situacion_aprendizaje:
        st.warning("📝 Se requiere el título de la Unidad y la Situación de Aprendizaje.")
    else:
        with st.spinner(f"🧠 Orquestando el diseño instruccional de {cantidad_sesiones} sesiones didácticas..."):
            try:
                prompt_maestro = f"""Actúa como un Doctor en Educación, Experto en Diseño Instruccional y Asesor Curricular del MINERD de la República Dominicana.

DATOS DEL CONTEXTO:
- Área Curricular: {materia}
- Unidad de Aprendizaje: {unidad_aprendizaje}
- Duración total: {duracion_total}
- Situación de Aprendizaje (Anclaje): {situacion_aprendizaje}
- Temas/Indicadores Guía: {temas_indicadores}
- Cantidad de Sesiones a Desarrollar: {cantidad_sesiones}

DIRECTRICES PEDAGÓGICAS PARA EL PLAN DIARIO:
Debes desglosar la unidad en {cantidad_sesiones} sesiones de clase. 
Para CADA sesión, debes definir estrictamente:
1. El Propósito de la clase y el Indicador de logro trabajado.
2. Fase de la unidad (Inicio de la unidad, Desarrollo, o Cierre de la unidad).
3. Los momentos didácticos (Inicio, Desarrollo, Cierre) con actividades concretas y tiempos realistas que sumen el total del bloque (ej. 50 min).
4. Recursos a utilizar y el método de Evaluación del día (ej. Observación directa, Rúbrica, Guía de preguntas).

REGLAS ESTRICTAS DE FORMATO JSON:
- NO uses saltos de línea literales (Enters/\\n) dentro de los valores de texto. Únelos con espacios.
- NO uses comillas dobles internas dentro de los textos. Usa comillas simples (' ') si necesitas citar.

FORMATO DE SALIDA ESTRICTO (JSON NATIVO OBLIGATORIO):
{{
  "SESIONES": [
    {{
      "NUMERO_SESION": 1,
      "TITULO_SESION": "Título motivador de la clase",
      "SEMANA": "Semana 1",
      "FASE_UNIDAD": "Inicio de la unidad",
      "TIEMPO_CLASE": "50 min",
      "PROPOSITO_CLASE": "Propósito específico del día",
      "INDICADOR_LOGRO": "Indicador que se evaluará hoy",
      "MOMENTOS": {{
        "INICIO_TIEMPO": "10 min",
        "INICIO_ACTIVIDAD": "Descripción de la actividad de inicio",
        "DESARROLLO_TIEMPO": "30 min",
        "DESARROLLO_ACTIVIDAD": "Descripción de la actividad central",
        "CIERRE_TIEMPO": "10 min",
        "CIERRE_ACTIVIDAD": "Descripción del cierre y metacognición"
      }},
      "RECURSOS": "Materiales a utilizar",
      "EVALUACION_DIA": "Técnica e instrumento de hoy"
    }}
  ]
}}
"""
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_json(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    respuesta_ia = solicitar_openai_json(api_key_usuario, modelo_seleccionado, prompt_maestro)

                datos_diarios = extraer_json_seguro(respuesta_ia)

                # --- RENDERIZADO DOCUMENTAL (PYTHON-DOCX) ---
                doc = Document()
                doc.styles['Normal'].font.name = 'Calibri'
                doc.styles['Normal'].font.size = Pt(11)

                sections = doc.sections
                for section in sections:
                    section.left_margin = Inches(0.75)
                    section.right_margin = Inches(0.75)

                def shade_cell(cell, color_hex):
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
                    cell._tc.get_or_add_tcPr().append(shd)
                
                # Encabezado Institucional
                p_enc = doc.add_paragraph()
                p_enc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_enc.add_run("MINISTERIO DE EDUCACIÓN DE LA REPÚBLICA DOMINICANA\n").bold = True
                p_enc.add_run("PLANIFICACIÓN DIARIA — NIVEL SECUNDARIO\n").bold = True

                # 1. Matriz General de Datos
                t_datos = doc.add_table(rows=4, cols=4)
                t_datos.style = 'Table Grid'
                
                t_datos.rows[0].cells[0].text = "Centro Educativo"
                t_datos.rows[0].cells[1].text = centro_educativo
                t_datos.rows[0].cells[2].text = "Año Escolar"
                t_datos.rows[0].cells[3].text = ano_escolar
                
                t_datos.rows[1].cells[0].text = "Docente"
                t_datos.rows[1].cells[1].text = docente
                t_datos.rows[1].cells[2].text = "Grado"
                t_datos.rows[1].cells[3].text = grado
                
                t_datos.rows[2].cells[0].text = "Área Curricular"
                t_datos.rows[2].cells[1].text = materia
                t_datos.rows[2].cells[2].text = "Sección"
                t_datos.rows[2].cells[3].text = seccion
                
                t_datos.rows[3].cells[0].text = "Unidad de Aprendizaje"
                t_datos.rows[3].cells[1].text = unidad_aprendizaje
                t_datos.rows[3].cells[2].text = "Duración total"
                t_datos.rows[3].cells[3].text = duracion_total

                for row in t_datos.rows:
                    row.cells[0].paragraphs[0].runs[0].bold = True
                    row.cells[2].paragraphs[0].runs[0].bold = True

                doc.add_paragraph()

                # 2. Situación de Aprendizaje Global
                p_sit_title = doc.add_paragraph()
                p_sit_title.add_run("Situación de Aprendizaje (referencia de toda la unidad)").bold = True
                p_sit = doc.add_paragraph(situacion_aprendizaje)
                p_sit.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                doc.add_paragraph("_" * 70)

                # 3. Iteración de Sesiones Diarias
                sesiones = datos_diarios.get("SESIONES", [])
                for sesion in sesiones:
                    
                    titulo_sesion = f"Sesión {sesion.get('NUMERO_SESION', '')}: {sesion.get('TITULO_SESION', '')}"
                    doc.add_heading(titulo_sesion, level=2)
                    
                    t_sesion = doc.add_table(rows=2, cols=4)
                    t_sesion.style = 'Table Grid'
                    
                    t_sesion.rows[0].cells[0].text = "Sesión N.º"
                    t_sesion.rows[0].cells[1].text = str(sesion.get("NUMERO_SESION", ""))
                    t_sesion.rows[0].cells[2].text = "Semana"
                    t_sesion.rows[0].cells[3].text = str(sesion.get("SEMANA", ""))
                    
                    t_sesion.rows[1].cells[0].text = "Fase de la unidad"
                    t_sesion.rows[1].cells[1].text = str(sesion.get("FASE_UNIDAD", ""))
                    t_sesion.rows[1].cells[2].text = "Tiempo de clase"
                    t_sesion.rows[1].cells[3].text = str(sesion.get("TIEMPO_CLASE", ""))
                    
                    for row in t_sesion.rows:
                        row.cells[0].paragraphs[0].runs[0].bold = True
                        row.cells[2].paragraphs[0].runs[0].bold = True
                        shade_cell(row.cells[0], "F8FAFC")
                        shade_cell(row.cells[2], "F8FAFC")
                        
                    doc.add_paragraph()
                    
                    p_prop = doc.add_paragraph()
                    p_prop.add_run("Propósito de la clase: ").bold = True
                    p_prop.add_run(str(sesion.get("PROPOSITO_CLASE", "")))
                    
                    p_ind = doc.add_paragraph()
                    p_ind.add_run("Indicador de logro trabajado: ").bold = True
                    p_ind.add_run(str(sesion.get("INDICADOR_LOGRO", "")))
                    
                    momentos = sesion.get("MOMENTOS", {})
                    t_act = doc.add_table(rows=4, cols=2)
                    t_act.style = 'Table Grid'
                    
                    t_act.rows[0].cells[0].text = "Momento"
                    t_act.rows[0].cells[1].text = "Actividades"
                    t_act.rows[0].cells[0].paragraphs[0].runs[0].bold = True
                    t_act.rows[0].cells[1].paragraphs[0].runs[0].bold = True
                    shade_cell(t_act.rows[0].cells[0], "E2E8F0")
                    shade_cell(t_act.rows[0].cells[1], "E2E8F0")
                    
                    t_act.rows[1].cells[0].text = f"Inicio\n({momentos.get('INICIO_TIEMPO', '')})"
                    t_act.rows[1].cells[1].text = str(momentos.get("INICIO_ACTIVIDAD", ""))
                    
                    t_act.rows[2].cells[0].text = f"Desarrollo\n({momentos.get('DESARROLLO_TIEMPO', '')})"
                    t_act.rows[2].cells[1].text = str(momentos.get("DESARROLLO_ACTIVIDAD", ""))
                    
                    t_act.rows[3].cells[0].text = f"Cierre\n({momentos.get('CIERRE_TIEMPO', '')})"
                    t_act.rows[3].cells[1].text = str(momentos.get("CIERRE_ACTIVIDAD", ""))
                    
                    for row in t_act.rows[1:]:
                        row.cells[0].paragraphs[0].runs[0].bold = True
                        row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    
                    doc.add_paragraph()
                    
                    p_rec = doc.add_paragraph()
                    p_rec.add_run("Recursos: ").bold = True
                    p_rec.add_run(str(sesion.get("RECURSOS", "")))
                    
                    p_eval = doc.add_paragraph()
                    p_eval.add_run("Evaluación del día: ").bold = True
                    p_eval.add_run(str(sesion.get("EVALUACION_DIA", "")))
                    
                    doc.add_paragraph("_" * 70)

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.success(f"✅ ¡Planificación Diaria (Secuencia de {cantidad_sesiones} sesiones) generada con éxito!")
                st.download_button(
                    label="📥 Exportar Plan Diario Oficial (Word)",
                    data=buffer,
                    file_name=f"Plan_Diario_{materia}_{grado[:3]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

            except Exception as e:
                st.error(f"⚠️ Error en la orquestación del documento: {e}")