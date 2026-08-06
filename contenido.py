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

# --- LLAMADAS A API CON JSON FORZADO ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_json(api_key, modelo, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=8192, 
            temperature=0.1,
            response_mime_type="application/json"
        )
    )
    return respuesta.text

def solicitar_openai_json(api_key, modelo, prompt):
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=modelo,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.1,
        response_format={"type": "json_object"}
    )
    return response.choices[0].message.content

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("⚡ Núcleo de Procesamiento")
    proveedor_ia = st.selectbox("Motor Analítico:", ["Google Gemini", "OpenAI (ChatGPT)"], key="prov_cont")
    
    if proveedor_ia == "Google Gemini":
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gemini-3.5-flash","gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"], key="mod_cont")
    else:
        modelo_seleccionado = st.selectbox("Versión de Red Neuronal:", ["gpt-4o-mini", "gpt-3.5-turbo", "gpt-4o"], key="mod_cont2")
        
    api_key_usuario = st.text_input("Clave de Autenticación (API Key):", type="password", key="api_cont")

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Generador de Contenidos y Actividades ETP</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Desarrollo didáctico, glosario, simuladores con enlaces y rúbricas avanzadas</div>', unsafe_allow_html=True)

# --- FORMULARIO ---
with st.form("form_contenido", clear_on_submit=False):
    
    st.markdown('<div class="section-title">🏫 1. Datos Institucionales</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        docente = st.text_input("Nombre del Docente", value="Ing. Bernardo Antonio Hernández Batista")
        asignatura = st.text_input("Módulo / Asignatura", placeholder="Ej: Ofimática, Sistema Operativo, Redes LAN")
    with col2:
        politecnico = st.text_input("Centro Educativo", value="Politécnico Salesiano Arquides Calderón")
        fecha = st.date_input("Fecha de Aplicación")
        
    st.markdown('<div class="section-title">📚 2. Base Pedagógica y Recursos</div>', unsafe_allow_html=True)
    contenido = st.text_area("Contenido a Desarrollar", height=100, placeholder="Ej: Configuración de subredes IP y Enrutamiento estático.")
    actividad = st.text_area("Actividad Práctica de Clase", height=100, placeholder="Ej: Los estudiantes simularán una red de área local conectando routers y switches.")
    
    st.markdown('<div class="section-title">📋 3. Estrategia de Evaluación</div>', unsafe_allow_html=True)
    col_eval1, col_eval2 = st.columns([3, 1])
    with col_eval1:
        instrumento = st.selectbox("Técnica / Instrumento de Evaluación", [
            "Lista de Cotejo Avanzada (Indicadores de Logro)",
            "Escala Estimativa con Descriptores de Desempeño",
            "Guía de Observación Metodológica",
            "Rúbrica Analítica de Competencias ETP",
            "Registro de Desempeño Técnico"
        ])
    with col_eval2:
        valor_puntos = st.number_input("Valor (Puntos)", min_value=1, max_value=100, value=100, help="Puntuación total de la actividad")
    
    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Generar Material e Instrumento Académico (Word)")

# --- LÓGICA CORE ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la barra lateral.")
    elif not asignatura or not contenido or not actividad:
        st.warning("📝 Por favor, completa la asignatura, el contenido y la actividad.")
    else:
        with st.spinner(f'🧠 Desarrollando contenido académico, enlaces de simuladores e instrumentos avanzados con {modelo_seleccionado}...'):
            try:
                prompt_maestro = f"""Actúa como un Catedrático Universitario de Alto Nivel y Especialista Curricular en la Educación Técnico Profesional (ETP). 

INSUMOS:
- Contenido a desarrollar: {contenido}
- Actividad propuesta: {actividad}
- Instrumento seleccionado: {instrumento}
- Valor total de la actividad: {valor_puntos} puntos.

REGLAS ESTRICTAS DE CALIDAD ACADÉMICA (CERO ALUCINACIONES):
1. DESARROLLO DE CONTENIDO: Crea un texto académico y técnico profesional (mínimo 3 párrafos robustos). Cero alucinaciones, basado en estándares reales de la industria.
2. REPOSITORIO DE SIMULADORES CON ENLACES REALES: Recomienda 3 herramientas digitales, simuladores de código abierto, laboratorios online o infografías técnicas específicas para este tema. CADA simulador DEBE incluir obligatoriamente su URL o enlace web oficial real y funcional (ej. https://www.netacad.com/, https://www.virtualbox.org/, https://phet.colorado.edu/, https://www.tinkercad.com/, etc.).
3. WEBGRAFÍA / BIBLIOGRAFÍA: Proporciona fuentes reales o referencias de donde provienen los conceptos.
4. GLOSARIO: Extrae al menos 5 términos técnicos del contenido y defínelos con precisión.
5. DESARROLLO DE ACTIVIDAD: Desglosa la actividad en pasos metodológicos (mínimo 4 pasos).
6. INSTRUMENTO DE EVALUACIÓN MEJORADO: Genera exactamente 5 criterios de evaluación de nivel experto. Cada criterio debe contener un título claro y un **indicador de desempeño observable** detallado. Distribuye los {valor_puntos} puntos exactos entre los 5 criterios (la suma debe dar {valor_puntos}).

FORMATO DE SALIDA ESTRICTO (JSON NATIVO OBLIGATORIO):
Devuelve un objeto JSON válido con la estructura exacta:
{{
  "TITULO": "Título académico y atractivo",
  "CONTENIDO_TEORICO": [
    "Párrafo 1...",
    "Párrafo 2...",
    "Párrafo 3..."
  ],
  "SIMULADORES_RECURSOS": [
    {{
      "TIPO": "Simulador Virtual / Laboratorio Online / Infografía",
      "NOMBRE": "Nombre de la herramienta",
      "DESCRIPCION": "Cómo se aplica este recurso digital específico.",
      "URL": "https://enlace-oficial-real.com"
    }}
  ],
  "WEBGRAFIA": [
    "Autor/Organización (Año). Título del recurso o sitio web..."
  ],
  "GLOSARIO": [
    {{"TERMINO": "Término 1", "DEFINICION": "Definición técnica..."}}
  ],
  "PASOS_ACTIVIDAD": [
    "Paso 1: [Descripción]"
  ],
  "CRITERIOS_EVALUACION": [
    {{
      "CRITERIO": "Nombre breve del criterio",
      "INDICADOR": "Descripción detallada del comportamiento o competencia observable que se evaluará.",
      "PUNTOS": 20
    }}
  ]
}}
"""
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_json(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    respuesta_ia = solicitar_openai_json(api_key_usuario, modelo_seleccionado, prompt_maestro)

                datos = json.loads(respuesta_ia)

                titulo_ia = datos.get("TITULO", "Desarrollo de Contenido")
                parrafos = datos.get("CONTENIDO_TEORICO", [])
                simuladores = datos.get("SIMULADORES_RECURSOS", [])
                webgrafia = datos.get("WEBGRAFIA", [])
                glosario = datos.get("GLOSARIO", [])
                pasos = datos.get("PASOS_ACTIVIDAD", [])
                criterios = datos.get("CRITERIOS_EVALUACION", [])

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
                p_encabezado.add_run("Material de Apoyo, Simuladores Multimedia y Guía Didáctica (ETP)\n").bold = True
                
                doc.add_paragraph(f"Docente: {docente} | Módulo: {asignatura} | Fecha: {fecha.strftime('%d/%m/%Y')}")
                doc.add_paragraph("Nombre del Estudiante: _________________________________________________ | Sección: _______")
                doc.add_paragraph("_" * 70)

                # 1. Teoría Académica
                doc.add_heading(f"📚 {titulo_ia}", level=1)
                for p_texto in parrafos:
                    p = doc.add_paragraph(str(p_texto))
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                doc.add_paragraph()

                # 2. Simuladores y Repositorio Multimedia con Enlaces
                if simuladores:
                    doc.add_heading("🌐 Repositorio de Simuladores y Recursos Multimedia", level=2)
                    doc.add_paragraph("Herramientas digitales, laboratorios online y simuladores recomendados con acceso web verificado:")
                    for sim in simuladores:
                        p_sim = doc.add_paragraph(style='List Bullet')
                        p_sim.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p_sim.add_run(f"[{sim.get('TIPO', 'Herramienta')}] ").bold = True
                        p_sim.add_run(f"{sim.get('NOMBRE', '')}: ").bold = True
                        p_sim.add_run(str(sim.get('DESCRIPCION', '')) + " ")
                        
                        # Inserción limpia del enlace web
                        url_val = sim.get('URL', '')
                        if url_val:
                            p_sim.add_run(f"🔗 [Enlace de Acceso: {url_val}]").italic = True

                doc.add_paragraph()

                # 3. Webgrafía
                if webgrafia:
                    doc.add_heading("🔗 Fuentes y Referencias (Webgrafía)", level=2)
                    for ref in webgrafia:
                        p_ref = doc.add_paragraph(str(ref), style='List Bullet')
                        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                doc.add_paragraph()

                # 4. Glosario
                if glosario:
                    doc.add_heading("🔑 Glosario de Palabras Claves", level=2)
                    for item in glosario:
                        p_glos = doc.add_paragraph(style='List Bullet')
                        p_glos.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p_glos.add_run(str(item.get("TERMINO", "")) + ": ").bold = True
                        p_glos.add_run(str(item.get("DEFINICION", "")))
                        
                doc.add_paragraph()

                # 5. Actividad Práctica
                doc.add_heading("⚙️ Desarrollo Metodológico de la Actividad", level=1)
                for paso in pasos:
                    p_paso = doc.add_paragraph(str(paso), style='List Bullet')
                    p_paso.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                doc.add_paragraph()

                # 6. Instrumento de Evaluación Mejorado con Indicadores
                doc.add_heading(f"📋 Instrumento de Evaluación Avanzado: {instrumento.split('(')[0].strip()}", level=1)
                doc.add_paragraph("Criterios evaluativos detallados con indicadores de desempeño observables para el registro docente.")
                
                columnas = ["No.", "Criterio e Indicador de Desempeño Observable", "Valor", "Calificación / Escala"]

                tabla_eval = doc.add_table(rows=1, cols=len(columnas))
                tabla_eval.style = 'Table Grid'
                
                def shade_cell(cell, color):
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
                    cell._tc.get_or_add_tcPr().append(shd)

                hdr_cells = tabla_eval.rows[0].cells
                for i, col_name in enumerate(columnas):
                    p = hdr_cells[i].paragraphs[0]
                    run = p.add_run(col_name)
                    run.bold = True
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    shade_cell(hdr_cells[i], "E2E8F0")

                for idx, crit_data in enumerate(criterios):
                    row_cells = tabla_eval.add_row().cells
                    row_cells[0].text = str(idx + 1)
                    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Formato mejorado unificando el criterio con su indicador descriptivo
                    texto_criterio = f"• Criterio: {str(crit_data.get('CRITERIO', ''))}\n• Indicador: {str(crit_data.get('INDICADOR', ''))}"
                    row_cells[1].text = texto_criterio
                    
                    row_cells[2].text = f"{crit_data.get('PUNTOS', 0)} pts"
                    row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    row_cells[3].text = "" # Celda abierta para anotación del maestro

                # Fila de Totales
                row_totales = tabla_eval.add_row().cells
                row_totales[1].text = "PUNTUACIÓN TOTAL:"
                row_totales[1].paragraphs[0].runs[0].bold = True
                row_totales[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                row_totales[2].text = f"{valor_puntos} pts"
                row_totales[2].paragraphs[0].runs[0].bold = True
                row_totales[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.success("✅ ¡Material Académico con Enlaces de Simuladores e Instrumentos Avanzados generado con éxito!")
                
                st.download_button(
                    label="📥 Descargar Documento Académico (.docx)",
                    data=buffer,
                    file_name=f"Material_Simuladores_Avanzado_{asignatura[:10]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary" 
                )
                
            except ResourceExhausted:
                st.error("❌ Se alcanzó el límite de API.")
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")