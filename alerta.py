import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import json
import re
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
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=8192, 
            temperature=0.2,
            response_mime_type="application/json"
        )
    )
    return respuesta.text

def solicitar_openai_json(api_key, modelo, prompt):
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=modelo,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        response_format={"type": "json_object"}
    )
    return response.choices[0].message.content

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("⚡ Núcleo de Procesamiento")
    proveedor_ia = st.selectbox("Motor Analítico:", ["Google Gemini", "OpenAI (ChatGPT)"], key="prov_alerta")
    
    if proveedor_ia == "Google Gemini":
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"], key="mod_alerta")
    else:
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gpt-4o-mini", "gpt-3.5-turbo", "gpt-4o"], key="mod_alerta2")
        
    api_key_usuario = st.text_input("Clave de Autenticación (API Key):", type="password", key="api_alerta")

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Sistema de Alerta Temprana y Reforzamiento</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Identificación de brechas formativas y planes de recuperación personalizados ETP</div>', unsafe_allow_html=True)

# --- FORMULARIO ---
with st.form("form_alerta", clear_on_submit=False):
    
    st.markdown('<div class="section-title">🏫 1. Datos del Contexto e Identificación</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        docente = st.text_input("Nombre del Docente", value="Ing. Bernardo Antonio Hernández Batista")
        asignatura = st.text_input("Módulo / Asignatura", placeholder="Ej: Sistemas Operativos / Redes LAN")
    with col2:
        politecnico = st.text_input("Centro Educativo", value="Politécnico Salesiano Arquides Calderón")
        seccion = st.text_input("Sección / Grado", placeholder="Ej: 6to de Informática")
        
    st.markdown('<div class="section-title">⚠️ 2. Registro de Estudiantes y Brechas Detectadas</div>', unsafe_allow_html=True)
    competencia_evaluada = st.text_input("Resultado de Aprendizaje (R.A.) o Competencia Evaluada", placeholder="Ej: R.A.1 Configurar los parámetros de red local...")
    
    estudiantes_apoyo = st.text_area(
        "Estudiantes Identificados 'En Proceso' o 'Necesitan Apoyo' (Nombres y dificultad principal)", 
        height=120, 
        placeholder="Ej:\n1. Carlos Pérez - Dificultad en el direccionamiento IP estático.\n2. María Gómez - Confusión en la identificación de topologías físicas."
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Generar Plan de Mejora Automático (Word)")

# --- LÓGICA CORE ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la barra lateral.")
    elif not asignatura or not estudiantes_apoyo or not competencia_evaluada:
        st.warning("📝 Por favor, completa la asignatura, la competencia y el listado de estudiantes.")
    else:
        with st.spinner(f'🧠 Analizando brechas de aprendizaje y redactando plan de recuperación con {modelo_seleccionado}...'):
            try:
                prompt_maestro = f"""Actúa como un Coordinador Pedagógico Expertos y Especialista en la Educación Técnico Profesional (ETP) del MINERD. 

CONTEXTO:
- Competencia / R.A. Evaluado: {competencia_evaluada}
- Estudiantes con brechas detectadas: 
{estudiantes_apoyo}

OBJETIVO:
Diseñar un Sistema de Alerta Temprana y un Plan de Reforzamiento / Recuperación Académica personalizado que atienda de manera directa las debilidades específicas reportadas en cada estudiante.

REGLAS ESTRICTAS:
1. RIGOR PEDAGÓGICO: No busques excluir; el plan debe orientarse a consolidar las competencias mediante estrategias prácticas y contextualizadas de nivel técnico.
2. ESTRATEGIAS ESPECÍFICAS: Proponga acciones concretas de nivelación para los temas donde los alumnos mostraron deficiencias.
3. ACTIVIDAD DE RECUPERACIÓN: Diseña una microactividad práctica de nivelación que los estudiantes puedan realizar de forma tutorial o guiada.

FORMATO DE SALIDA ESTRICTO (JSON NATIVO OBLIGATORIO):
Devuelve un objeto JSON válido con la estructura exacta:
{{
  "DIAGNOSTICO_GENERAL": "Breve análisis de las brechas detectadas en el grupo sobre esta competencia...",
  "PLAN_ACCION_ESTUDIANTES": [
    {{
      "ESTUDIANTE": "Nombre del estudiante",
      "BRECHA_DETECTADA": "Resumen de la dificultad técnica",
      "ESTRATEGIA_REFORZAMIENTO": "Acción tutorial o correctiva recomendada"
    }}
  ],
  "ACTIVIDAD_RECUPERACION_GRUPAL": {{
    "TITULO": "Título de la actividad de nivelación",
    "DESCRIPCION": "Descripción metodológica de la práctica de recuperación...",
    "PASOS": [
      "Paso 1...",
      "Paso 2...",
      "Paso 3..."
    ]
  }},
  "CRITERIO_REVALUACION": "Cómo se comprobará que el estudiante superó la brecha (indicador de logro final)"
}}
"""
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_json(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    respuesta_ia = solicitar_openai_json(api_key_usuario, modelo_seleccionado, prompt_maestro)

                datos = json.loads(respuesta_ia)

                diagnostico = datos.get("DIAGNOSTICO_GENERAL", "")
                plan_estudiantes = datos.get("PLAN_ACCION_ESTUDIANTES", [])
                actividad_recu = datos.get("ACTIVIDAD_RECUPERACION_GRUPAL", {})
                criterio_reval = datos.get("CRITERIO_REVALUACION", "")

                # --- CONSTRUCCIÓN DEL DOCUMENTO WORD ---
                doc = Document()
                doc.styles['Normal'].font.name = 'Calibri'
                doc.styles['Normal'].font.size = Pt(11)

                sections = doc.sections
                for section in sections:
                    section.left_margin = Inches(0.75)
                    section.right_margin = Inches(0.75)

                # Encabezado institucional
                p_encabezado = doc.add_paragraph()
                p_encabezado.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_encabezado.add_run(f"{politecnico}\n").bold = True
                p_encabezado.add_run("Sistema de Alerta Temprana y Plan de Reforzamiento Académico (ETP)\n").bold = True
                
                doc.add_paragraph(f"Docente: {docente} | Módulo: {asignatura} | Sección: {seccion}")
                doc.add_paragraph(f"Competencia / R.A. Analizado: {competencia_evaluada}")
                doc.add_paragraph("_" * 70)

                # 1. Diagnóstico
                doc.add_heading("📊 1. Diagnóstico de Alerta Temprana", level=1)
                p_diag = doc.add_paragraph(diagnostico)
                p_diag.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                doc.add_paragraph()

                # 2. Plan Personalizado por Estudiante (Tabla)
                doc.add_heading("🎯 2. Plan de Mejora Personalizado por Estudiante", level=1)
                doc.add_paragraph("Estrategias de intervención enfocadas en los estudiantes identificados 'En Proceso' o 'Necesitan Apoyo':")

                tabla_plan = doc.add_table(rows=1, cols=3)
                tabla_plan.style = 'Table Grid'
                
                def shade_cell(cell, color):
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
                    cell._tc.get_or_add_tcPr().append(shd)

                hdr_cells = tabla_plan.rows[0].cells
                headers = ["Estudiante", "Brecha / Dificultad Detectada", "Estrategia de Reforzamiento"]
                for i, h_text in enumerate(headers):
                    p = hdr_cells[i].paragraphs[0]
                    run = p.add_run(h_text)
                    run.bold = True
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    shade_cell(hdr_cells[i], "E2E8F0")

                for est in plan_estudiantes:
                    row_cells = tabla_plan.add_row().cells
                    row_cells[0].text = str(est.get("ESTUDIANTE", ""))
                    row_cells[0].paragraphs[0].runs[0].bold = True
                    row_cells[1].text = str(est.get("BRECHA_DETECTADA", ""))
                    row_cells[2].text = str(est.get("ESTRATEGIA_REFORZAMIENTO", ""))

                doc.add_paragraph()

                # 3. Actividad de Recuperación Grupal/Tutorial
                doc.add_heading(f"⚙️ 3. Actividad de Recuperación: {actividad_recu.get('TITULO', 'Nivelación')}", level=1)
                p_act = doc.add_paragraph(str(actividad_recu.get("DESCRIPCION", "")))
                p_act.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                doc.add_heading("Pasos Metodológicos de la Tutoría:", level=3)
                for paso in actividad_recu.get("PASOS", []):
                    p_paso = doc.add_paragraph(str(paso), style='List Bullet')
                    p_paso.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                doc.add_paragraph()

                # 4. Criterio de Reevaluación
                doc.add_heading("📋 4. Criterio de Cierre y Reevaluación", level=1)
                p_reval = doc.add_paragraph(criterio_reval)
                p_reval.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.success("✅ ¡Plan de Alerta Temprana y Reforzamiento generado con éxito!")
                
                st.download_button(
                    label="📥 Descargar Plan de Mejora (.docx)",
                    data=buffer,
                    file_name=f"Plan_Mejora_{asignatura[:10]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary" 
                )
                
            except ResourceExhausted:
                st.error("❌ Se alcanzó el límite de API.")
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")