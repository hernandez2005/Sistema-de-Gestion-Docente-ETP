import streamlit as st
import google.generativeai as genai
from openai import OpenAI, RateLimitError as OpenAIRateLimitError
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from io import BytesIO
import PyPDF2
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

# --- FUNCIONES DE TEXTO Y JSON ROBUSTO ---
MARKER_NL = "<<NL>>"
MARKER_DQ = "<<DQ>>"
MARKER_TAB = "<<TAB>>"

def decodificar_marcadores(obj):
    if isinstance(obj, str):
        return obj.replace(MARKER_NL, "\n").replace(MARKER_DQ, '"').replace(MARKER_TAB, "\t")
    if isinstance(obj, dict):
        return {decodificar_marcadores(k): decodificar_marcadores(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [decodificar_marcadores(item) for item in obj]
    return obj

def reparar_json_truncado(texto):
    in_string = False
    escape_next = False
    llaves = corchetes = 0
    last_safe_pos = 0
    for i, char in enumerate(texto):
        if escape_next:
            escape_next = False
            continue
        if in_string:
            if char == "\\": escape_next = True
            elif char == '"':
                in_string = False
                last_safe_pos = i + 1
            continue
        if char == '"': in_string = True
        elif char == "{": llaves += 1; last_safe_pos = i + 1
        elif char == "}": llaves -= 1; last_safe_pos = i + 1
        elif char == "[": corchetes += 1; last_safe_pos = i + 1
        elif char == "]": corchetes -= 1; last_safe_pos = i + 1
        elif char in (",", ":", " ", "\n", "\r", "\t"): last_safe_pos = i + 1

    reparado = texto[:last_safe_pos]
    if in_string: reparado += '"'
    reparado = reparado.rstrip()
    if reparado.endswith(","): reparado = reparado[:-1]
    reparado += "]" * max(corchetes, 0) + "}" * max(llaves, 0)
    return reparado

def parsear_json_robusto(respuesta):
    if not respuesta or not respuesta.strip(): raise ValueError("La IA devolvió una respuesta vacía.")
    texto = respuesta.strip()
    if texto.startswith("```json"): texto = texto[7:]
    elif texto.startswith("```"): texto = texto[3:]
    if texto.endswith("```"): texto = texto[:-3]
    texto = texto.strip()
    try: return json.loads(texto, strict=False)
    except json.JSONDecodeError: pass
    match = re.search(r"(\{[\s\S]*\})", texto)
    if match:
        try: return json.loads(match.group(1), strict=False)
        except json.JSONDecodeError: pass
    json_start = texto.find("{")
    if json_start >= 0:
        cuerpo = texto[json_start:]
        try: return json.loads(reparar_json_truncado(cuerpo), strict=False)
        except json.JSONDecodeError: pass
    raise ValueError(f"JSON irrecuperable. Inicio: {texto[:400]}...")

# --- EXTRACTOR DE TEXTO (PDF Y WORD) ---
def extraer_texto_documento(archivo):
    if archivo.name.lower().endswith(".pdf"):
        pdf_reader = PyPDF2.PdfReader(archivo)
        return "".join([pagina.extract_text() or "" for pagina in pdf_reader.pages])
    elif archivo.name.lower().endswith(".docx"):
        doc = Document(archivo)
        texto = []
        for para in doc.paragraphs:
            texto.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    texto.append(cell.text)
        return "\n".join(texto)
    return ""

# --- FUNCIONES DE API CON REINTENTOS ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_con_reintento(api_key, modelo, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=8192, temperature=0.2, response_mime_type="application/json"
        )
    )
    return respuesta.text

@retry(retry=retry_if_exception_type(OpenAIRateLimitError), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_openai_con_reintento(api_key, modelo, prompt):
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=modelo, messages=[{"role": "user", "content": prompt}],
        temperature=0.2, max_tokens=8192, response_format={"type": "json_object"}
    )
    return response.choices[0].message.content

# --- FUNCIÓN MODULAR PARA GENERAR WORD (ESQUEMA 2026 OFICIAL) ---
def generar_documento_diario_2026(datos, form_data):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)

    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

    def shade_cell(cell, color):
        shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
        cell._tc.get_or_add_tcPr().append(shd)

    def fijar_anchos_columna(tabla, anchos_pulgadas):
        tabla.autofit = False
        for row in tabla.rows:
            for idx, ancho in enumerate(anchos_pulgadas):
                if idx < len(row.cells):
                    row.cells[idx].width = Inches(ancho)

    # ENCABEZADO OFICIAL (Membrete)
    p_centro = doc.add_paragraph()
    p_centro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_centro = p_centro.add_run(f"{form_data['centro']}\n")
    run_centro.bold = True
    run_centro.font.size = Pt(12)
    
    if form_data['eslogan']:
        run_eslogan = p_centro.add_run(f"“{form_data['eslogan']}”\n")
        run_eslogan.italic = True
        run_eslogan.font.size = Pt(10)
        
    if form_data['ubicacion']:
        p_centro.add_run(f"{form_data['ubicacion']}\n").font.size = Pt(9)
    if form_data['telefono']:
        p_centro.add_run(f"Teléfono: {form_data['telefono']}\n").font.size = Pt(9)
    if form_data['email']:
        p_centro.add_run(f"E-mail: {form_data['email']}").font.size = Pt(9)

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
    fijar_anchos_columna(t1, [1.5, 2.0, 2.0, 2.0])

    doc.add_paragraph()

    # TABLA 2: DETALLES PEDAGÓGICOS (Ampliada con Comp. Fundamental)
    t2 = doc.add_table(rows=6, cols=2)
    t2.style = 'Table Grid'
    detalles = [
        ("Estrategias de enseñanza - aprendizaje:", form_data['estrategias']),
        ("Tema:", form_data['tema']),
        ("Competencia Fundamental:", datos.get("COMPETENCIA_FUNDAMENTAL", "No especificada en la malla.")),
        ("Competencias específicas del grado:", datos.get("COMPETENCIAS", "No especificada en la malla.")),
        ("Intención pedagógica del día:", datos.get("INTENCION", form_data['intencion'])),
        ("Indicador de logro:", datos.get("INDICADOR", "No especificado en la malla."))
    ]
    for i, (label, val) in enumerate(detalles):
        cell_label = t2.cell(i, 0)
        cell_label.text = label
        cell_label.paragraphs[0].runs[0].bold = True
        shade_cell(cell_label, "F1F5F9")
        t2.cell(i, 1).text = val
    fijar_anchos_columna(t2, [2.5, 5.5])

    doc.add_paragraph()

    # TABLA 3: MATRIZ DE MOMENTOS PEDAGÓGICOS (4 Columnas)
    t3 = doc.add_table(rows=4, cols=4)
    t3.style = 'Table Grid'
    
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
    comp_cell.merge(t3.cell(2, 0)).merge(t3.cell(3, 0))
    comp_cell.text = datos.get("COMPETENCIAS", "Competencias específicas del grado extraídas de la malla.")
    comp_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    fijar_anchos_columna(t3, [1.5, 1.0, 3.5, 2.0])

    doc.add_paragraph()

    # TABLA 4: RECUPERACIÓN PEDAGÓGICA
    t4 = doc.add_table(rows=1, cols=1)
    t4.style = 'Table Grid'
    cell_rec = t4.cell(0, 0)
    p_rec = cell_rec.paragraphs[0]
    p_rec.add_run("Actividades de recuperación pedagógica: ").bold = True
    p_rec.add_run(datos.get("RECUPERACION", ""))
    fijar_anchos_columna(t4, [8.0])

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- CONFIGURACIÓN CENTRALIZADA ---
api_key_usuario = st.session_state.get("api_key_global", "")
proveedor_ia = st.session_state.get("proveedor_ia_global", "Google Gemini")
modelo_seleccionado = st.session_state.get("modelo_global", "gemini-2.5-flash")

with st.sidebar:
    st.markdown("##### ⚡ Plan Diario")
    if not api_key_usuario:
        st.error("🔒 Configura tu API Key en la página de Inicio")
    else:
        st.success(f"✅ {proveedor_ia} · {modelo_seleccionado}")

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Esquema de Planificación Diaria</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Ministerio de Educación de la República Dominicana (MINERD)</div>', unsafe_allow_html=True)

# --- FORMULARIO ---
with st.form("form_plandiario2026", clear_on_submit=False):
    
    st.markdown('<div class="section-title">🏫 1. Datos de Identificación Institucional</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        centro = st.text_input("Nombre del Centro Educativo", placeholder="Ej: Politécnico Salesiano Arquides Calderón")
        ubicacion = st.text_input("Ubicación", placeholder="Ej: Moca, Provincia Espaillat")
        telefono = st.text_input("Teléfono", placeholder="Ej: 809-823-3322")
    with col2:
        eslogan = st.text_input("Eslogan del Centro", placeholder="Ej: Formando Honrados Ciudadanos y Buenos Cristianos")
        email = st.text_input("E-mail", placeholder="Ej: politecnicoacmoca@gmail.com")
        docente = st.text_input("Docente", placeholder="Nombre del docente")
        
    col3, col4 = st.columns(2)
    with col3:
        area = st.selectbox("Área Curricular", [
            "Matemática", "Lengua Española", "Ciencias Sociales", "Ciencias de la Naturaleza", 
            "Inglés", "Francés", "Educación Física", "Educación Artística", "Formación Integral Humana y Religiosa"
        ])
    with col4:
        grado = st.text_input("Grado y sección", placeholder="Ej: 5to de Bachillerato, Sección A")
        fecha = st.date_input("Fecha de Ejecución")

    st.markdown('<div class="section-title">📄 2. Malla Curricular Oficial (Obligatoria)</div>', unsafe_allow_html=True)
    archivo_malla = st.file_uploader("Carga el documento de la Malla Curricular (PDF o Word)", type=["pdf", "docx"], help="La IA extraerá automáticamente las competencias e indicadores desde este documento, evitando invenciones.")

    st.markdown('<div class="section-title">📝 3. Datos Curriculares de la Clase</div>', unsafe_allow_html=True)
    estrategias = st.text_input("Estrategias de enseñanza - aprendizaje", value="Recuperación de Experiencias Previas, Sociodrama o Dramatización, Expositiva de Conocimientos Elaborados y/o Acumulados, Indagación Dialógica o Cuestionamiento, Descubrimiento e Indagación")
    tema = st.text_input("Tema de la clase", placeholder="Ej: ITBIS Ley 254-06. Origen. Base Legal.")
    intencion = st.text_area("Intención pedagógica del día", height=70, placeholder="Ej: Comprender el origen legal del ITBIS para aplicarlo en facturas prácticas.")
    
    st.markdown('<div class="section-title">👥 4. Perfil Sociocognitivo</div>', unsafe_allow_html=True)
    perfil_grupo = st.text_area("Características del grupo y NEAE (Opcional)", placeholder="Ej: Grupo kinestésico, 2 estudiantes con TDAH. Requiere ejemplos visuales.")

    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Generar Planificación Diaria 2026")

# --- LÓGICA CORE ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la página de Inicio.")
    elif not archivo_malla:
        st.warning("⚠️ Debes cargar la Malla Curricular para extraer las competencias e indicadores sin alucinaciones.")
    elif not tema or not intencion:
        st.warning("📝 Por favor, completa el Tema y la Intención Pedagógica.")
    else:
        with st.spinner(f'🧠 Leyendo malla curricular y diseñando planificación con {modelo_seleccionado}...'):
            respuesta_ia = None
            try:
                # 1. Extraer texto de la malla
                texto_malla = extraer_texto_documento(archivo_malla)
                if len(texto_malla) > 60000: texto_malla = texto_malla[:60000]

                if len(texto_malla.strip()) < 50:
                    st.error("❌ No se pudo extraer texto del archivo. Si es un PDF escaneado, intenta pasar el texto a Word.")
                    st.stop()

                prompt_maestro = f"""Actúa como experto en planificación educativa del MINERD (República Dominicana) para el esquema oficial 2026.
Diseña la estructura metodológica para una clase de 50 minutos del área de {area}.

TEMA DE LA CLASE: {tema}
INTENCIÓN PEDAGÓGICA: {intencion}
PERFIL DEL GRUPO: {perfil_grupo if perfil_grupo else "Grupo estándar de secundaria"}

MALLA CURRICULAR OFICIAL (DOCUMENTO FUENTE):
{texto_malla}

REGLA CRÍTICA DE EXTRACCIÓN (CERO ALUCINACIONES):
1. Busca en la MALLA CURRICULAR la Competencia Fundamental, la Competencia Específica del grado y el Indicador de Logro que mejor se adapten al TEMA de la clase.
2. Transcríbelos TEXTUALMENTE en el JSON. NO inventes competencias ni indicadores que no estén en el documento.
3. Si el documento no contiene algún elemento, escribe "No encontrado en la malla" en ese campo.

REGLAS DE PLANIFICACIÓN:
1. Para el INICIO (8 min), detalla actividades y recursos.
2. Para el DESARROLLO (32 min), separa claramente en campos distintos: Procedimientos, Actividad y Estrategias, junto con recursos.
3. Para el CIERRE (10 min), redacta una Indagación Dialógica/Cuestionamiento y una actividad de Metacognición, con recursos.
4. Redacta Actividades de recuperación pedagógica para estudiantes que necesiten apoyo.
5. FORMATO: Devuelve ÚNICAMENTE un JSON válido en texto plano (sin markdown). 
6. Si necesitas salto de línea en un texto, usa la etiqueta: {MARKER_NL}

FORMATO DE SALIDA ESTRICTO (JSON NATIVO):
{{
  "COMPETENCIA_FUNDAMENTAL": "[Extraída textualmente de la malla]",
  "COMPETENCIAS": "[Competencias específicas extraídas textualmente]",
  "INDICADOR": "[Indicador de logro extraído textualmente]",
  "INTENCION": "{intencion}",
  "INICIO": {{
    "ACTIVIDADES": "[Actividades de inicio con duración (8 min)]",
    "RECURSOS": "[Recursos para el inicio]"
  }},
  "DESARROLLO": {{
    "PROCEDIMIENTOS": "[Pasos a seguir por el docente (10 min)]",
    "ACTIVIDAD": "[Actividad principal del estudiante (15 min)]",
    "ESTRATEGIAS": "[Estrategias aplicadas (7 min)]",
    "RECURSOS": "[Recursos para el desarrollo]"
  }},
  "CIERRE": {{
    "INDAGACION": "[Preguntas o cuestionamiento (5 min)]",
    "METACOGNICION": "[Actividad de reflexión (5 min)]",
    "RECURSOS": "[Recursos para el cierre]"
  }},
  "RECUPERACION": "[Actividades de recuperación pedagógica]"
}}
"""
                if proveedor_ia == "Google Gemini":
                    respuesta_ia = solicitar_gemini_con_reintento(api_key_usuario, modelo_seleccionado, prompt_maestro)
                else:
                    respuesta_ia = solicitar_openai_con_reintento(api_key_usuario, modelo_seleccionado, prompt_maestro)

                datos = parsear_json_robusto(respuesta_ia)
                datos = decodificar_marcadores(datos)

                datos_formulario = {
                    "centro": centro, "eslogan": eslogan, "ubicacion": ubicacion, 
                    "telefono": telefono, "email": email, "fecha": fecha.strftime('%d/%m/%Y'),
                    "area": area, "docente": docente, "grado": grado, 
                    "estrategias": estrategias, "tema": tema, "intencion": intencion
                }

                buffer_docx = generar_documento_diario_2026(datos, datos_formulario)
                
                st.success("✅ ¡Planificación Diaria 2026 generada con extracción automática de malla!")
                
                st.download_button(
                    label="📥 Descargar Planificación Diaria 2026 (.docx)",
                    data=buffer_docx,
                    file_name=f"Plan_Diario_2026_{area}_{fecha.strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary" 
                )
                
            except ResourceExhausted:
                st.error("❌ Límite de API Gemini alcanzado. Espera unos momentos.")
            except OpenAIRateLimitError:
                st.error("❌ Límite de API OpenAI alcanzado. Espera unos momentos.")
            except ValueError as ve:
                st.error(f"⚠️ Error de procesamiento de JSON: {ve}")
                if respuesta_ia:
                    with st.expander("🔍 Ver respuesta cruda de la IA (Depuración)"):
                        st.text(respuesta_ia[:3000])
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")