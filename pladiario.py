import streamlit as st
import google.generativeai as genai
from openai import OpenAI
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

# --- REINTENTO DE API ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_con_reintento(api_key, modelo, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(max_output_tokens=6000, temperature=0.2)
    )
    return respuesta.text

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("⚡ Núcleo de Procesamiento")
    proveedor_ia = st.selectbox("Motor Analítico:", ["Google Gemini", "OpenAI (ChatGPT)"], key="prov_plad")
    
    if proveedor_ia == "Google Gemini":
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"], key="mod_plad")
    else:
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gpt-4o-mini", "gpt-3.5-turbo", "gpt-4o"], key="mod_plad2")
        
    api_key_usuario = st.text_input("Clave de Autenticación (API Key):", type="password", key="api_plad")

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
        st.error("🔒 Debes ingresar tu API Key en la barra lateral.")
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
                respuesta_ia = ""
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_con_reintento(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    client = OpenAI(api_key=api_key_usuario)
                    response = client.chat.completions.create(
                        model=modelo_seleccionado,
                        messages=[{"role": "user", "content": prompt_maestro}],
                        temperature=0.2,
                        max_tokens=4000
                    )
                    respuesta_ia = response.choices[0].message.content

                # Limpieza de Markdowns
                respuesta_limpia = respuesta_ia.strip()
                if respuesta_limpia.startswith("```json"): respuesta_limpia = respuesta_limpia[7:]
                elif respuesta_limpia.startswith("```"): respuesta_limpia = respuesta_limpia[3:]
                if respuesta_limpia.endswith("```"): respuesta_limpia = respuesta_limpia[:-3]
                respuesta_limpia = respuesta_limpia.strip()

                try:
                    datos = json.loads(respuesta_limpia)
                except json.JSONDecodeError as e:
                    st.error("❌ Error grave: La IA no devolvió un formato JSON válido.")
                    with st.expander("🔍 Ver respuesta cruda (Depuración)"):
                        st.text(respuesta_limpia)
                        st.text(str(e))
                    st.stop()

                conceptuales = datos.get("COMPONENTES", {}).get("CONCEPTUALES", "")
                procedimentales = datos.get("COMPONENTES", {}).get("PROCEDIMENTALES", "")
                actitudinales = datos.get("COMPONENTES", {}).get("ACTITUDINALES", "")
                
                enunciado = datos.get("ACTIVIDAD", {}).get("ENUNCIADO", "")
                intencion = datos.get("ACTIVIDAD", {}).get("INTENCION", "")
                
                inicio_f1 = datos.get("INICIO", {}).get("FASE1", "")
                inicio_f2 = datos.get("INICIO", {}).get("FASE2", "")
                inicio_f3 = datos.get("INICIO", {}).get("FASE3", "")
                
                desarrollo_f1 = datos.get("DESARROLLO", {}).get("FASE1", "")
                desarrollo_f2 = datos.get("DESARROLLO", {}).get("FASE2", "")
                desarrollo_f3 = datos.get("DESARROLLO", {}).get("FASE3", "")
                
                cierre_f1 = datos.get("CIERRE", {}).get("FASE1", "")
                cierre_f2 = datos.get("CIERRE", {}).get("FASE2", "")
                
                recursos = datos.get("RECURSOS", "")
                neae = datos.get("NEAE", "")
                
                criterios = datos.get("COTEJO", [])
                while len(criterios) < 5: criterios.append("Criterio pendiente de definir")

                # --- CONSTRUCCIÓN DEL DOCUMENTO WORD ---
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

                def add_table_header(text):
                    p = doc.add_paragraph()
                    run = p.add_run(text)
                    run.bold = True
                    run.font.color.rgb = RGBColor(37, 99, 235) 

                # --- TABLA 1: DATOS GENERALES ---
                add_table_header("Tabla 1: Datos Generales")
                t1 = doc.add_table(rows=6, cols=2)
                t1.style = 'Table Grid'
                t1.cell(0,0).text = f"Nombre de Docente: {docente}"
                t1.cell(0,1).text = f"Cédula: {cedula}"
                t1.cell(1,0).text = f"Regional: {regional}"
                t1.cell(1,1).text = f"Distrito: {distrito}"
                t1.cell(2,0).text = f"Centro Educativo: {centro}"
                t1.cell(2,1).text = "Modalidad: Técnico Profesional (ETP)"
                t1.cell(3,0).text = "Nivel / Subsistema: Secundaria"
                t1.cell(3,1).text = f"Grado y Sección: {grado}"
                t1.cell(4,0).text = f"Módulo Formativo: {modulo}"
                t1.cell(4,1).text = "Tiempo Estimado: 50 minutos"
                t1.cell(5,0).text = f"Estrategias: {estrategias}"
                t1.cell(5,1).text = f"Fecha: {fecha.strftime('%d/%m/%Y')}"
                for row in t1.rows:
                    for cell in row.cells:
                        cell.paragraphs[0].runs[0].bold = True

                doc.add_paragraph()

                # --- TABLA 2: DATOS CURRICULARES ---
                add_table_header("Tabla 2: Datos Curriculares")
                t2 = doc.add_table(rows=6, cols=3)
                t2.style = 'Table Grid'
                
                def set_merged_row(row_idx, label, text):
                    cell = t2.cell(row_idx, 0)
                    cell.merge(t2.cell(row_idx, 1)).merge(t2.cell(row_idx, 2))
                    p = cell.paragraphs[0]
                    p.add_run(label).bold = True
                    p.add_run(f" {text}")

                set_merged_row(0, "Resultado de Aprendizaje (RA):", str(ra))
                set_merged_row(1, "Criterio de Evaluación (CE):", str(ce))
                set_merged_row(2, "Elemento de Capacidad (EC):", str(ec))
                
                t2.cell(3,0).text = f"Conceptuales:\n{conceptuales}"
                t2.cell(3,1).text = f"Procedimentales:\n{procedimentales}"
                t2.cell(3,2).text = f"Actitudinales:\n{actitudinales}"
                
                set_merged_row(4, "Enunciado de la Actividad:", str(enunciado))
                set_merged_row(5, "Intención Educativa:", str(intencion))

                doc.add_paragraph()

                # --- TABLA 3: MOMENTOS PEDAGÓGICOS ---
                add_table_header("Tabla 3: Momentos Pedagógicos y Tiempo")
                t3 = doc.add_table(rows=3, cols=1)
                t3.style = 'Table Grid'
                
                p_inicio = t3.cell(0,0).paragraphs[0]
                p_inicio.add_run("INICIO (8 min):\n").bold = True
                p_inicio.add_run(f"• Fase 1 (Motivación): {inicio_f1}\n• Fase 2 (Saberes Previos): {inicio_f2}\n• Fase 3 (Intención): {inicio_f3}")
                
                p_desarrollo = t3.cell(1,0).paragraphs[0]
                p_desarrollo.add_run("DESARROLLO (32 min):\n").bold = True
                p_desarrollo.add_run(f"• Fase 1 (Análisis/Introducción): {desarrollo_f1}\n• Fase 2 (Aplicación colaborativa): {desarrollo_f2}\n• Fase 3 (Socialización): {desarrollo_f3}")
                
                p_cierre = t3.cell(2,0).paragraphs[0]
                p_cierre.add_run("CIERRE (10 min):\n").bold = True
                p_cierre.add_run(f"• Fase 1 (Consolidación): {cierre_f1}\n• Fase 2 (Metacognición): {cierre_f2}")

                doc.add_paragraph()

                # --- TABLA 4: RECURSOS Y NEAE ---
                add_table_header("Tabla 4: Recursos y Adaptaciones NEAE")
                t4 = doc.add_table(rows=2, cols=1)
                t4.style = 'Table Grid'
                p_rec = t4.cell(0,0).paragraphs[0]
                p_rec.add_run("Recursos:\n").bold = True
                p_rec.add_run(str(recursos))
                
                p_neae = t4.cell(1,0).paragraphs[0]
                p_neae.add_run("Adaptaciones para NEAE:\n").bold = True
                p_neae.add_run(str(neae))

                doc.add_paragraph()

                # --- TABLA 5: LISTA DE COTEJO ---
                add_table_header("Tabla 5: Lista de Cotejo - Instrumento de Evaluación")
                t5 = doc.add_table(rows=6, cols=5)
                t5.style = 'Table Grid'
                encabezados_cot = ["No.", "Criterios de Evaluación", "L", "EP", "NA"]
                for i, text in enumerate(encabezados_cot):
                    t5.cell(0,i).text = text
                    t5.cell(0,i).paragraphs[0].runs[0].bold = True
                
                for i in range(5):
                    t5.cell(i+1, 0).text = str(i+1)
                    t5.cell(i+1, 1).text = str(criterios[i]) if i < len(criterios) else ""
                
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
                
                st.success("✅ ¡Planificación diaria estructurada con éxito en JSON nativo!")
                
                st.download_button(
                    label="📥 Descargar Planificación de Clase Diaria (.docx)",
                    data=buffer,
                    file_name=f"Plan_Diario_{fecha.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary" 
                )
                
            except ResourceExhausted:
                st.error("❌ Se alcanzó el límite de API. Espera unos momentos antes de reintentar.")
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")