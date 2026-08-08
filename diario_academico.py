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
            max_output_tokens=4000, 
            temperature=0.2,
            response_mime_type="application/json"
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
        response_format={"type": "json_object"}
    )
    return response.choices[0].message.content

# --- FUNCIÓN MODULAR PARA GENERAR WORD (ESQUEMA 2026) ---
def generar_documento_diario_2026(datos, form_data):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)

    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    def shade_cell(cell, color):
        shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
        cell._tc.get_or_add_tcPr().append(shd)

    # ENCABEZADO OFICIAL
    p_centro = doc.add_paragraph()
    p_centro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_centro = p_centro.add_run(f"{form_data['centro']}\n")
    run_centro.bold = True
    run_centro.font.size = Pt(12)
    run_eslogan = p_centro.add_run(f"“{form_data['eslogan']}”")
    run_eslogan.italic = True
    run_eslogan.font.size = Pt(10)

    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_tit = p_titulo.add_run("ESQUEMA DE PLANIFICACIÓN DIARIA")
    run_tit.bold = True
    run_tit.font.size = Pt(14)

    # TABLA 1: DATOS BÁSICOS (Fecha, Área, Docente, Grado)
    t1 = doc.add_table(rows=2, cols=4)
    t1.style = 'Table Grid'
    headers_t1 = ["Fecha", "Área", "Docente", "Grado y sección"]
    valores_t1 = [form_data['fecha'], form_data['area'], form_data['docente'], form_data['grado']]
    for i, txt in enumerate(headers_t1):
        cell = t1.cell(0, i)
        cell.text = txt
        cell.paragraphs[0].runs[0].bold = True
        shade_cell(cell, "E2E8F0")
    for i, txt in enumerate(valores_t1):
        t1.cell(1, i).text = txt

    doc.add_paragraph()

    # TABLA 2: DETALLES PEDAGÓGICOS
    t2 = doc.add_table(rows=4, cols=2)
    t2.style = 'Table Grid'
    detalles = [
        ("Estrategias de enseñanza - aprendizaje", form_data['estrategias']),
        ("Tema:", form_data['tema']),
        ("Intención pedagógica del día:", datos.get("INTENCION", form_data['intencion'])),
        ("Indicador de logro:", datos.get("INDICADOR", form_data['indicador']))
    ]
    for i, (label, val) in enumerate(detalles):
        cell_label = t2.cell(i, 0)
        cell_label.text = label
        cell_label.paragraphs[0].runs[0].bold = True
        shade_cell(cell_label, "F1F5F9")
        t2.cell(i, 1).text = val
        # Hacer la primera columna más estrecha
        t2.cell(i, 0).width = Inches(2.5)
        t2.cell(i, 1).width = Inches(5.5)

    doc.add_paragraph()

    # TABLA 3: MATRIZ DE MOMENTOS PEDAGÓGICOS (4 Columnas)
    t3 = doc.add_table(rows=4, cols=4) # 1 encabezado + 3 momentos
    t3.style = 'Table Grid'
    
    # Encabezados
    headers_t3 = ["Competencias específicas del grado", "Momentos", "Actividades/duración", "Recursos"]
    for i, txt in enumerate(headers_t3):
        cell = t3.cell(0, i)
        cell.text = txt
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade_cell(cell, "DBEAFE")

    # Fila 1: INICIO
    t3.cell(1, 1).text = "Inicio"
    t3.cell(1, 1).paragraphs[0].runs[0].bold = True
    t3.cell(1, 2).text = datos.get("INICIO", {}).get("ACTIVIDADES", "")
    t3.cell(1, 3).text = datos.get("INICIO", {}).get("RECURSOS", "")

    # Fila 2: DESARROLLO (Contiene sub-elementos)
    t3.cell(2, 1).text = "Desarrollo"
    t3.cell(2, 1).paragraphs[0].runs[0].bold = True
    
    # Formatear el Desarrollo con sus sub-temas en la misma celda
    des_data = datos.get("DESARROLLO", {})
    p_des = t3.cell(2, 2).paragraphs[0]
    p_des.add_run("Procedimientos: ").bold = True
    p_des.add_run(des_data.get("PROCEDIMIENTOS", "") + "\n")
    p_des.add_run("Actividad: ").bold = True
    p_des.add_run(des_data.get("ACTIVIDAD", "") + "\n")
    p_des.add_run("Estrategias: ").bold = True
    p_des.add_run(des_data.get("ESTRATEGIAS", ""))
    
    t3.cell(2, 3).text = des_data.get("RECURSOS", "")

    # Fila 3: CIERRE (Contiene Indagación y Metacognición)
    t3.cell(3, 1).text = "Cierre"
    t3.cell(3, 1).paragraphs[0].runs[0].bold = True
    
    cie_data = datos.get("CIERRE", {})
    p_cie = t3.cell(3, 2).paragraphs[0]
    p_cie.add_run("Indagación Dialógica o Cuestionamiento: ").bold = True
    p_cie.add_run(cie_data.get("INDAGACION", "") + "\n")
    p_cie.add_run("Metacognición: ").bold = True
    p_cie.add_run(cie_data.get("METACOGNICION", ""))
    
    t3.cell(3, 3).text = cie_data.get("RECURSOS", "")

    # Combinar verticalmente la columna "Competencias"
    comp_cell = t3.cell(1, 0)
    comp_cell.merge(t3.cell(3, 0))
    comp_cell.text = datos.get("COMPETENCIAS", "Competencias específicas del grado redactadas por la IA.")
    comp_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # TABLA 4: RECUPERACIÓN PEDAGÓGICA
    t4 = doc.add_table(rows=1, cols=1)
    t4.style = 'Table Grid'
    cell_rec = t4.cell(0, 0)
    p_rec = cell_rec.paragraphs[0]
    p_rec.add_run("Actividades de recuperación pedagógica: ").bold = True
    p_rec.add_run(datos.get("RECUPERACION", ""))

    # Guardar en buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- CONFIGURACIÓN CENTRALIZADA ---
api_key_usuario = st.session_state.get("api_key_global", "")
proveedor_ia = st.session_state.get("proveedor_ia_global", "Google Gemini")
modelo_seleccionado = st.session_state.get("modelo_global", "gemini-2.5-flash")

with st.sidebar:
    st.markdown("##### ⚡ Plan Diario 2026")
    if not api_key_usuario:
        st.error("🔒 Configura tu API Key en la página de Inicio")
    else:
        st.success(f"✅ {proveedor_ia} · {modelo_seleccionado}")

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Esquema de Planificación Diaria 2026</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Ministerio de Educación de la República Dominicana (MINERD)</div>', unsafe_allow_html=True)

# --- FORMULARIO ---
with st.form("form_plandiario2026", clear_on_submit=False):
    
    st.markdown('<div class="section-title">🏫 1. Datos Generales</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        centro = st.text_input("Nombre del Centro Educativo", value="Politécnico Salesiano Arquídes Calderón")
        docente = st.text_input("Docente", value="Ing. Bernardo Antonio Hernández Batista")
        area = st.text_input("Área / Módulo", placeholder="Ej: Contabilidad / Impuestos al Consumo")
    with col2:
        eslogan = st.text_input("Eslogan del Centro", value="Formando Honrados Ciudadanos y Buenos Cristianos")
        grado = st.text_input("Grado y sección", value="5to de Bachillerato, Sección A")
        fecha = st.date_input("Fecha de Ejecución")

    st.markdown('<div class="section-title">📝 2. Datos Curriculares de la Clase</div>', unsafe_allow_html=True)
    estrategias = st.text_input("Estrategias de enseñanza - aprendizaje globales", value="Estudio de casos, Trabajo colaborativo")
    tema = st.text_input("Tema de la clase", placeholder="Ej: ITBIS Ley 254-06. Origen. Base Legal.")
    intencion = st.text_area("Intención pedagógica del día", height=70, placeholder="Ej: Comprender el origen legal del ITBIS para aplicarlo en facturas prácticas.")
    indicador = st.text_area("Indicador de logro", height=70, placeholder="Ej: Calcula el ITBIS correctamente en documentos comerciales.")
    
    st.markdown('<div class="section-title">👥 3. Perfil Sociocognitivo (Opcional, para adaptar IA)</div>', unsafe_allow_html=True)
    perfil_grupo = st.text_area("Características del grupo y NEAE", placeholder="Ej: Grupo kinestésico, 2 estudiantes con TDAH. Requiere ejemplos visuales.")

    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Generar Planificación Diaria 2026")

# --- LÓGICA CORE ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la página de Inicio.")
    elif not tema or not intencion or not indicador:
        st.warning("📝 Por favor, completa al menos el Tema, la Intención Pedagógica y el Indicador de Logro.")
    else:
        with st.spinner(f'🧠 Diseñando estructura pedagógica 2026 con {modelo_seleccionado}...'):
            try:
                prompt_maestro = f"""Actúa como experto en planificación educativa del MINERD (República Dominicana) para el esquema 2026.
                Diseña la estructura metodológica para una clase de 50 minutos.

                INSUMOS:
                - Tema: {tema}
                - Intención pedagógica: {intencion}
                - Indicador de logro: {indicador}
                - Estrategias globales: {estrategias}
                - Perfil del grupo: {perfil_grupo if perfil_grupo else "Grupo estándar de bachillerato técnico"}

                REGLAS:
                1. Redacta 1 o 2 Competencias específicas del grado muy precisas.
                2. Para el INICIO, detalla actividades y recursos (8 min).
                3. Para el DESARROLLO, separa claramente Procedimientos, Actividad y Estrategias, junto con recursos (32 min).
                4. Para el CIERRE, redacta una Indagación Dialógica/Cuestionamiento y una actividad de Metacognición, con recursos (10 min).
                5. Redacta Actividades de recuperación pedagógica para estudiantes que necesiten apoyo.
                6. FORMATO: Devuelve ÚNICAMENTE un JSON válido en texto plano (sin markdown). Estructura exacta:

                {{
                  "COMPETENCIAS": "[Competencias específicas]",
                  "INTENCION": "{intencion}",
                  "INDICADOR": "{indicador}",
                  "INICIO": {{
                    "ACTIVIDADES": "[Actividades de inicio con duración]",
                    "RECURSOS": "[Recursos para el inicio]"
                  }},
                  "DESARROLLO": {{
                    "PROCEDIMIENTOS": "[Pasos a seguir]",
                    "ACTIVIDAD": "[Actividad principal]",
                    "ESTRATEGIAS": "[Estrategias aplicadas]",
                    "RECURSOS": "[Recursos para el desarrollo]"
                  }},
                  "CIERRE": {{
                    "INDAGACION": "[Preguntas o cuestionamiento]",
                    "METACOGNICION": "[Actividad de reflexión]",
                    "RECURSOS": "[Recursos para el cierre]"
                  }},
                  "RECUPERACION": "[Actividades de recuperación pedagógica]"
                }}
                """
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_con_reintento(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    respuesta_ia = solicitar_openai_con_reintento(api_key_usuario, modelo_seleccionado, prompt_maestro)

                try:
                    datos = json.loads(respuesta_ia)
                except json.JSONDecodeError as e:
                    st.error("❌ Error: La IA no devolvió un formato JSON válido.")
                    with st.expander("🔍 Ver respuesta cruda"):
                        st.text(respuesta_ia)
                        st.text(str(e))
                    st.stop()

                datos_formulario = {
                    "centro": centro, "eslogan": eslogan, "fecha": fecha.strftime('%d/%m/%Y'),
                    "area": area, "docente": docente, "grado": grado, 
                    "estrategias": estrategias, "tema": tema, "intencion": intencion, "indicador": indicador
                }

                buffer_docx = generar_documento_diario_2026(datos, datos_formulario)
                
                st.success("✅ ¡Planificación Diaria 2026 estructurada con éxito!")
                
                st.download_button(
                    label="📥 Descargar Planificación Diaria 2026 (.docx)",
                    data=buffer_docx,
                    file_name=f"Plan_Diario_2026_{fecha.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary" 
                )
                
            except ResourceExhausted:
                st.error("❌ Límite de API Gemini alcanzado. Espera unos momentos.")
            except OpenAIRateLimitError:
                st.error("❌ Límite de API OpenAI alcanzado. Espera unos momentos.")
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")