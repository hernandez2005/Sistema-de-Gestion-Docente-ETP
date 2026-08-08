import streamlit as st
import google.generativeai as genai
from openai import OpenAI, RateLimitError as OpenAIRateLimitError
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from io import BytesIO
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

# --- FUNCIONES DE API CON REINTENTOS ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_con_reintento(api_key, modelo, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=6000, 
            temperature=0.2,
            response_mime_type="application/json" # JSON Nativo
        )
    )
    return respuesta.text

@retry(retry=retry_if_exception_type(OpenAIRateLimitError), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_openai_con_reintento(api_key, modelo, prompt):
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=modelo,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=4000,
        response_format={"type": "json_object"} # JSON Nativo
    )
    return response.choices[0].message.content

# --- FUNCIÓN MODULAR PARA GENERAR WORD ---
def generar_documento_plandiario(datos, form_data):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)

    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_tit1 = p_titulo.add_run("MINISTERIO DE EDUCACIÓN DE LA REPÚBLICA DOMINICANA\n")
    run_tit1.bold = True
    run_tit1.font.size = Pt(14)
    run_tit2 = p_titulo.add_run("PLANIFICACIÓN DE CLASE DIARIA - Modalidad Técnico Profesional (ETP)")
    run_tit2.bold = True
    run_tit2.font.size = Pt(12)

    def shade_cell(cell, color):
        shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
        cell._tc.get_or_add_tcPr().append(shd)

    def add_table_header(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.color.rgb = RGBColor(37, 99, 235) 

    # --- TABLA 1: DATOS GENERALES ---
    add_table_header("Tabla 1: Datos Generales")
    t1 = doc.add_table(rows=6, cols=2)
    t1.style = 'Table Grid'
    t1.cell(0,0).text = f"Nombre de Docente: {form_data['docente']}"
    t1.cell(0,1).text = f"Cédula: {form_data['cedula']}"
    t1.cell(1,0).text = f"Regional: {form_data['regional']}"
    t1.cell(1,1).text = f"Distrito: {form_data['distrito']}"
    t1.cell(2,0).text = f"Centro Educativo: {form_data['centro']}"
    t1.cell(2,1).text = "Modalidad: Técnico Profesional (ETP)"
    t1.cell(3,0).text = "Nivel / Subsistema: Secundaria"
    t1.cell(3,1).text = f"Grado y Sección: {form_data['grado']}"
    t1.cell(4,0).text = f"Módulo Formativo: {form_data['modulo']}"
    t1.cell(4,1).text = "Tiempo Estimado: 50 minutos"
    t1.cell(5,0).text = f"Estrategias: {form_data['estrategias']}"
    t1.cell(5,1).text = f"Fecha: {form_data['fecha'].strftime('%d/%m/%Y')}"
    for row in t1.rows:
        for cell in row.cells:
            cell.paragraphs[0].runs[0].bold = True

    doc.add_paragraph()

    # --- TABLA 2: DATOS CURRICULARES ---
    add_table_header("Tabla 2: Datos Curriculares")
    t2 = doc.add_table(rows=6, cols=3)
    t2.style = 'Table Grid'
    
    def set_merged_row(table, row_idx, label, text):
        cell = table.cell(row_idx, 0)
        cell.merge(table.cell(row_idx, 1)).merge(table.cell(row_idx, 2))
        p = cell.paragraphs[0]
        p.add_run(label).bold = True
        p.add_run(f" {text}")

    comp = datos.get("COMPONENTES", {})
    act = datos.get("ACTIVIDAD", {})
    
    set_merged_row(t2, 0, "Resultado de Aprendizaje (RA):", str(form_data['ra']))
    set_merged_row(t2, 1, "Criterio de Evaluación (CE):", str(form_data['ce']))
    set_merged_row(t2, 2, "Elemento de Capacidad (EC):", str(form_data['ec']))
    
    t2.cell(3,0).text = f"Conceptuales:\n{comp.get('CONCEPTUALES', '')}"
    t2.cell(3,1).text = f"Procedimentales:\n{comp.get('PROCEDIMENTALES', '')}"
    t2.cell(3,2).text = f"Actitudinales:\n{comp.get('ACTITUDINALES', '')}"
    
    set_merged_row(t2, 4, "Enunciado de la Actividad:", str(act.get("ENUNCIADO", "")))
    set_merged_row(t2, 5, "Intención Educativa:", str(act.get("INTENCION", "")))

    doc.add_paragraph()

    # --- TABLA 3: MOMENTOS PEDAGÓGICOS ---
    add_table_header("Tabla 3: Momentos Pedagógicos y Tiempo")
    t3 = doc.add_table(rows=3, cols=1)
    t3.style = 'Table Grid'
    
    ini = datos.get("INICIO", {})
    des = datos.get("DESARROLLO", {})
    cie = datos.get("CIERRE", {})
    
    p_inicio = t3.cell(0,0).paragraphs[0]
    p_inicio.add_run("INICIO (8 min):\n").bold = True
    p_inicio.add_run(f"Fase 1 (Motivación): {ini.get('FASE1', '')}\nFase 2 (Saberes Previos): {ini.get('FASE2', '')}\nFase 3 (Intención): {ini.get('FASE3', '')}")
    
    p_desarrollo = t3.cell(1,0).paragraphs[0]
    p_desarrollo.add_run("DESARROLLO (32 min):\n").bold = True
    p_desarrollo.add_run(f"Fase 1 (Análisis/Introducción): {des.get('FASE1', '')}\nFase 2 (Aplicación colaborativa): {des.get('FASE2', '')}\nFase 3 (Socialización): {des.get('FASE3', '')}")
    
    p_cierre = t3.cell(2,0).paragraphs[0]
    p_cierre.add_run("CIERRE (10 min):\n").bold = True
    p_cierre.add_run(f"Fase 1 (Consolidación): {cie.get('FASE1', '')}\nFase 2 (Metacognición): {cie.get('FASE2', '')}")

    doc.add_paragraph()

    # --- TABLA 4: RECURSOS Y NEAE ---
    add_table_header("Tabla 4: Recursos y Adaptaciones NEAE")
    t4 = doc.add_table(rows=2, cols=1)
    t4.style = 'Table Grid'
    p_rec = t4.cell(0,0).paragraphs[0]
    p_rec.add_run("Recursos:\n").bold = True
    p_rec.add_run(str(datos.get("RECURSOS", "")))
    
    p_neae = t4.cell(1,0).paragraphs[0]
    p_neae.add_run("Adaptaciones para NEAE:\n").bold = True
    p_neae.add_run(str(datos.get("NEAE", "")))

    doc.add_paragraph()

    # --- TABLA 5: LISTA DE COTEJO ---
    add_table_header("Tabla 5: Lista de Cotejo - Instrumento de Evaluación")
    t5 = doc.add_table(rows=6, cols=5)
    t5.style = 'Table Grid'
    encabezados_cot = ["No.", "Criterios de Evaluación", "L", "EP", "NA"]
    
    # Encabezados con sombreado
    for i, text in enumerate(encabezados_cot):
        cell = t5.cell(0,i)
        cell.text = text
        p = cell.paragraphs[0]
        p.runs[0].bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade_cell(cell, "E2E8F0")
        
    criterios = datos.get("COTEJO", [])
    # Asegurar que haya al menos 5 criterios
    while len(criterios) < 5: criterios.append("Criterio pendiente de definir")
    
    # Anchos para la tabla de cotejo
    anchos_cotejo = [Inches(0.5), Inches(5.5), Inches(0.8), Inches(0.8), Inches(0.8)]
    
    for i in range(5):
        row_cells = t5.rows[i+1].cells
        row_cells[0].text = str(i+1)
        row_cells[1].text = str(criterios[i])
        row_cells[2].text = "" # L
        row_cells[3].text = "" # EP
        row_cells[4].text = "" # NA
        
        for idx, width in enumerate(anchos_cotejo):
            row_cells[idx].width = width
            row_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_leyenda = doc.add_paragraph()
    run_ley = p_leyenda.add_run("(Leyenda: L = Logrado, EP = En Proceso, NA = Necesita Apoyo)")
    run_ley.italic = True
    run_ley.font.size = Pt(9)

    # --- FIRMAS ---
    doc.add_paragraph("\n\n")
    t_firmas = doc.add_table(rows=2, cols=3)
    
    t_firmas.cell(0,0).text = "_________________________"
    t_firmas.cell(0,1).text = "_________________________"
    t_firmas.cell(0,2).text = "_________________________"
    
    t_firmas.cell(1,0).text = "Director/a de Centro"
    t_firmas.cell(1,1).text = "Coordinador/a Módulos Formativos ETP"
    t_firmas.cell(1,2).text = "Docente ETP"
    
    for row in t_firmas.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if "Docente ETP" in cell.text:
                cell.paragraphs[0].runs[0].bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- CONFIGURACIÓN CENTRALIZADA ---
api_key_usuario = st.session_state.get("api_key_global", "")
proveedor_ia = st.session_state.get("proveedor_ia_global", "Google Gemini")
modelo_seleccionado = st.session_state.get("modelo_global", "gemini-2.5-flash")

with st.sidebar:
    st.markdown("##### ⚡ Plan Diario ETP")
    if not api_key_usuario:
        st.error("🔒 Configura tu API Key en la página de Inicio")
    else:
        st.success(f"✅ {proveedor_ia} · {modelo_seleccionado}")

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Planificación de Clase Diaria ETP</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Diseño de momentos pedagógicos y matriz de evaluación en 50 minutos</div>', unsafe_allow_html=True)

# --- FORMULARIO ---
with st.form("form_pladiario", clear_on_submit=False):
    
    st.markdown('<div class="section-title">🏫 1. Datos Generales e Institucionales</div>', unsafe_allow_html=True)
    col_inst1, col_inst2, col_inst3 = st.columns(3)
    with col_inst1:
        docente = st.text_input("Docente", value="Ing. Bernardo Antonio Hernández Batista")
        regional = st.text_input("Regional", value="06 (La Vega/Espaillat)")
        grado = st.text_input("Grado y Sección", value="Nivel de Secundaria")
    with col_inst2:
        cedula = st.text_input("Cédula", placeholder="---")
        distrito = st.text_input("Distrito", value="06 (Moca)")
        fecha = st.date_input("Fecha de Ejecución")
    with col_inst3:
        centro = st.text_input("Centro Educativo", value="Politécnico")
        estrategias = st.text_input("Estrategias Base", value="Estudio de casos, Trabajo colaborativo")
        
    st.markdown('<div class="section-title">🎯 2. Parámetros Curriculares de la Clase</div>', unsafe_allow_html=True)
    col_curr1, col_curr2 = st.columns(2)
    with col_curr1:
        modulo = st.text_area("Módulo Formativo (MF)", height=100, placeholder="Ej: MF_358_3: Impuestos al consumo...")
        ra = st.text_area("Resultado de Aprendizaje (RA)", height=100)
        tema = st.text_input("Tema Específico de la clase hoy", placeholder="Ej: ITBIS Ley 254-06. Origen. Base Legal.")
    with col_curr2:
        ce = st.text_area("Criterio de Evaluación (CE)", height=100)
        ec = st.text_area("Elemento de Capacidad (EC)", height=100)
        
    st.markdown('<div class="section-title">👥 3. Perfil Sociocognitivo del Grupo</div>', unsafe_allow_html=True)
    caracteristicas = st.text_area("Características del grupo de estudiantes y NEAE", placeholder="Ej: Grupo muy visual y kinestésico. Tienen buena disposición colaborativa pero se distraen rápido. Hay 2 estudiantes con dislexia y 1 con TDAH leve.", height=70)
    
    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Generar Planificación Diaria")

# --- LÓGICA CORE ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la página de Inicio (barra lateral).")
    elif not modulo or not ra or not ce or not ec or not tema or not caracteristicas:
        st.warning("📝 Por favor, completa todos los campos curriculares y las características del grupo.")
    else:
        with st.spinner(f'🧠 Diseñando estructura pedagógica adaptada (50 min) en formato JSON con {modelo_seleccionado}...'):
            try:
                prompt_maestro = f"""Actúa como experto en planificación educativa y diseño curricular de la ETP (MINERD).

Objetivo: Diseñar una "Planificación de Clase Diaria" de 50 minutos para estudiantes de {grado} de secundaria (Bachillerato Técnico). 
Características Específicas del Grupo: {caracteristicas}
(Ajusta el nivel de las actividades, la estrategia y las adaptaciones NEAE basándote ESTRICTAMENTE en este perfil sociocognitivo).

Reglas Estrictas:
1. Coherencia: Los contenidos deben derivarse del EC y CE.
2. Gestión de Tiempo: Inicio (8 min), Desarrollo (32 min) y Cierre (10 min).
3. Formato de Salida: Genera ÚNICAMENTE un objeto JSON válido, sin bloques de código markdown, con la estructura exacta detallada abajo.
4. TEXTO PLANO: No utilices formato Markdown (como ** o *) en los valores del JSON. Usa texto plano.

INSUMOS CURRICULARES:
- Módulo Formativo (MF): {modulo}
- Resultado de Aprendizaje (RA): {ra}
- Criterio de Evaluación (CE): {ce}
- Elemento de Capacidad (EC): {ec}
- Tema Específico: {tema}

FORMATO DE SALIDA ESTRICTO (JSON NATIVO):
{{
  "COMPONENTES": {{
    "CONCEPTUALES": "[Redacta los conceptos clave]",
    "PROCEDIMENTALES": "[Redacta los procedimientos paso a paso]",
    "ACTITUDINALES": "[Redacta las actitudes a fomentar]"
  }},
  "ACTIVIDAD": {{
    "ENUNCIADO": "[Enunciado de la actividad práctica/estudio de caso]",
    "INTENCION": "[Intención educativa orientada al perfil]"
  }},
  "INICIO": {{
    "FASE1": "[Fase 1 Motivación/Reflexión (Ej. 2 min) - ADAPTADO AL GRUPO]",
    "FASE2": "[Fase 2 Saberes Previos (Ej. 4 min)]",
    "FASE3": "[Fase 3 Intención Educativa (Ej. 2 min)]"
  }},
  "DESARROLLO": {{
    "FASE1": "[Fase 1 Análisis/Introducción técnica (Ej. 10 min)]",
    "FASE2": "[Fase 2 Aplicación colaborativa (Ej. 15 min) - ADAPTADO AL GRUPO]",
    "FASE3": "[Fase 3 Socialización (Ej. 7 min)]"
  }},
  "CIERRE": {{
    "FASE1": "[Fase 1 Consolidación interactiva (Ej. 5 min)]",
    "FASE2": "[Fase 2 Metacognición con preguntas (Ej. 5 min)]"
  }},
  "RECURSOS": "[Lista de recursos didácticos]",
  "NEAE": "[Párrafo detallado de Adaptaciones para NEAE, respondiendo a las características ingresadas]",
  "COTEJO": [
    "[Criterio 1 - Técnico]",
    "[Criterio 2 - Técnico]",
    "[Criterio 3 - Procedimental]",
    "[Criterio 4 - Procedimental]",
    "[Criterio 5 - Actitudinal/Colaborativo]"
  ]
}}
"""
                # Llamada a la API
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_con_reintento(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    respuesta_ia = solicitar_openai_con_reintento(api_key_usuario, modelo_seleccionado, prompt_maestro)

                # Parseo JSON Directo
                try:
                    datos = json.loads(respuesta_ia)
                except json.JSONDecodeError as e:
                    st.error("❌ Error grave: La IA no devolvió un formato JSON válido.")
                    with st.expander("🔍 Ver respuesta cruda (Depuración)"):
                        st.text(respuesta_ia)
                        st.text(str(e))
                    st.stop()

                # Empaquetar datos del formulario
                datos_formulario = {
                    "docente": docente, "cedula": cedula, "regional": regional, "distrito": distrito,
                    "centro": centro, "grado": grado, "modulo": modulo, "estrategias": estrategias,
                    "fecha": fecha, "ra": ra, "ce": ce, "ec": ec
                }

                # Generar Documento Word
                buffer_docx = generar_documento_plandiario(datos, datos_formulario)
                
                st.success("✅ ¡Planificación diaria estructurada con éxito en JSON nativo!")
                
                st.download_button(
                    label="📥 Descargar Planificación de Clase Diaria (.docx)",
                    data=buffer_docx,
                    file_name=f"Plan_Diario_{fecha.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary" 
                )
                
            except ResourceExhausted:
                st.error("❌ Se alcanzó el límite de API Gemini. Espera unos momentos antes de reintentar.")
            except OpenAIRateLimitError:
                st.error("❌ Se alcanzó el límite de API OpenAI. Espera unos momentos antes de reintentar.")
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")