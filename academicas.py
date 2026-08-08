import streamlit as st
import google.generativeai as genai
from openai import OpenAI
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

# --- FUNCIÓN DE LIMPIEZA JSON ---
def extraer_json_seguro(texto_ia):
    if not texto_ia or not texto_ia.strip():
        return {}, True, ""
    texto = texto_ia.strip()
    if texto.startswith("```json"): texto = texto[7:]
    elif texto.startswith("```"): texto = texto[3:]
    if texto.endswith("```"): texto = texto[:-3]
    texto = texto.strip()
    try:
        return json.loads(texto, strict=False), False, texto
    except json.JSONDecodeError:
        pass
    match = re.search(r'(\{.*\})', texto, re.DOTALL)
    if match:
        bloque = match.group(1)
        try:
            return json.loads(bloque, strict=False), False, bloque
        except json.JSONDecodeError:
            pass
    else:
        bloque = texto
    truncado = True
    reparado = bloque
    comillas = len(re.findall(r'(?<!\\)"', reparado))
    if comillas % 2 != 0: reparado += '"'
    reparado += ']' * max(0, reparado.count('[') - reparado.count(']'))
    reparado += '}' * max(0, reparado.count('{') - reparado.count('}'))
    reparado = re.sub(r',\s*([}\]])', r'\1', reparado)
    try:
        return json.loads(reparado, strict=False), truncado, reparado
    except json.JSONDecodeError:
        pass
    reparado2 = re.sub(r'[\n\r]', ' ', bloque)
    comillas2 = len(re.findall(r'(?<!\\)"', reparado2))
    if comillas2 % 2 != 0: reparado2 += '"'
    reparado2 += ']' * max(0, reparado2.count('[') - reparado2.count(']'))
    reparado2 += '}' * max(0, reparado2.count('{') - reparado2.count('}'))
    reparado2 = re.sub(r',\s*([}\]])', r'\1', reparado2)
    try:
        return json.loads(reparado2, strict=False), truncado, reparado2
    except json.JSONDecodeError:
        return {}, True, reparado

# --- LLAMADAS A API ---
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_json(api_key, modelo, prompt, max_tokens=16384):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            max_output_tokens=max_tokens, temperature=0.15,
            response_mime_type="application/json"
        )
    )
    corte = False
    try:
        fr = respuesta.candidates[0].finish_reason
        corte = (str(fr).upper().find("MAX_TOKENS") != -1) or fr == 2
    except Exception: pass
    try: texto = respuesta.text
    except Exception: texto = ""
    return texto, corte

def solicitar_openai_json(api_key, modelo, prompt, max_tokens=16384):
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=modelo, messages=[{"role": "user", "content": prompt}],
        temperature=0.15, max_tokens=max_tokens, response_format={"type": "json_object"}
    )
    corte = response.choices[0].finish_reason == "length"
    return response.choices[0].message.content, corte

def solicitar_con_reintento(proveedor, api_key, modelo, prompt, max_tokens=16384, tope=32000):
    texto, corte = (solicitar_gemini_json if proveedor == "Google Gemini" else solicitar_openai_json)(api_key, modelo, prompt, max_tokens)
    reintentado = False
    if corte and max_tokens < tope:
        nuevos = min(max_tokens * 2, tope)
        texto2, corte2 = (solicitar_gemini_json if proveedor == "Google Gemini" else solicitar_openai_json)(api_key, modelo, prompt, nuevos)
        reintentado = True
        if len(texto2 or "") >= len(texto or ""):
            texto, corte = texto2, corte2
    return texto, corte, reintentado

# --- CONFIGURACIÓN CENTRALIZADA (desde main.py) ---
api_key_usuario = st.session_state.get("api_key_global", "")
proveedor_ia = st.session_state.get("proveedor_ia_global", "Google Gemini")
modelo_seleccionado = st.session_state.get("modelo_global", "gemini-2.5-flash")

with st.sidebar:
    st.markdown("##### ⚡ Plan de Unidad Académica")
    if not api_key_usuario:
        st.error("🔒 Configura tu API Key en la página de Inicio")
    else:
        st.success(f"✅ {proveedor_ia} · {modelo_seleccionado}")

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Planificación de Unidad de Aprendizaje</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Diseño de unidad para áreas académicas — Nivel Secundario MINERD</div>', unsafe_allow_html=True)

# --- FORMULARIO ---
with st.form("form_academicas", clear_on_submit=False):

    st.markdown('<div class="section-title">📄 1. Fuente Curricular (Opcional)</div>', unsafe_allow_html=True)
    archivo_pdf = st.file_uploader("Cargue el documento curricular (PDF) del área", type=["pdf"], help="Ancla la planificación al documento oficial. Elimina alucinaciones.")

    st.markdown('<div class="section-title">🏫 2. Datos Institucionales</div>', unsafe_allow_html=True)
    col_i1, col_i2, col_i3 = st.columns(3)
    with col_i1:
        docente = st.text_input("Docente", value="Ing. Bernardo Antonio Hernández Batista")
        centro = st.text_input("Centro Educativo", value="Politécnico Salesiano Arquides Calderón")
    with col_i2:
        area = st.selectbox("Área Curricular", [
            "Ciencias Sociales", "Lengua Española", "Matemática",
            "Ciencias de la Naturaleza", "Inglés", "Francés",
            "Educación Artística", "Educación Física",
            "Formación Integral Humana y Religiosa"
        ])
        grado = st.text_input("Grado", value="2do de Secundaria")
    with col_i3:
        seccion = st.text_input("Sección", value="B")
        ano_escolar = st.text_input("Año Escolar", value="2026-2027")

    st.markdown('<div class="section-title">🎯 3. Parámetros de la Unidad</div>', unsafe_allow_html=True)
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        titulo_unidad = st.text_input("Título de la Unidad", placeholder="Ej: Riesgos naturales en nuestra comunidad")
        duracion_unidad = st.text_input("Duración total", placeholder="Ej: 5 semanas / 15 sesiones")
        cantidad_sesiones = st.number_input("Cantidad de sesiones a planificar", min_value=1, max_value=20, value=5)
    with col_u2:
        competencia = st.text_area("Competencia / Competencia Fundamental", height=68, placeholder="Pega la competencia del área curricular...")
        indicadores_logro = st.text_area("Indicadores de Logro", height=68, placeholder="Pega los indicadores de logro de la unidad...")

    st.markdown('📚 4. Situación de Aprendizaje y Contenidos</div>', unsafe_allow_html=True)
    situacion_aprendizaje = st.text_area(
        "Situación de Aprendizaje (Contexto motivador)", 
        height=100,
        placeholder="Ej: La Defensa Civil del municipio ha solicitado a los centros educativos elaborar un mapa de riesgos naturales de la comunidad..."
    )
    contenidos_decl = st.text_area(
        "Contenidos declarados en el currículo (resumen)",
        height=80,
        placeholder="Ej: Riesgos naturales. Prevención y mitigación. Mapas de riesgo. Organismos de respuesta."
    )

    st.markdown('<div class="section-title">👥 5. Perfil del Grupo</div>', unsafe_allow_html=True)
    perfil_grupo = st.text_area(
        "Características del grupo / NEAE", 
        height=68,
        placeholder="Ej: Grupo de 35 estudiantes, predominio visual-kinestésico. 2 con necesidades educativas especiales."
    )

    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("⚙️ Generar Plan de Unidad Académica (Word)")

# --- LÓGICA CORE ---
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la página de Inicio (barra lateral).")
    elif not titulo_unidad or not situacion_aprendizaje:
        st.warning("📝 Completa el título de la unidad y la situación de aprendizaje.")
    else:
        with st.spinner(f'🧠 Diseñando plan de unidad académica con {modelo_seleccionado}...'):
            try:
                # --- PDF ---
                texto_curriculo = ""
                if archivo_pdf:
                    pdf_reader = PyPDF2.PdfReader(archivo_pdf)
                    texto_curriculo = "".join([pagina.extract_text() for pagina in pdf_reader.pages])
                    if len(texto_curriculo) > 80000: texto_curriculo = texto_curriculo[:80000]

                anclaje = ""
                if texto_curriculo:
                    anclaje = f"""
DOCUMENTO CURRICULAR OFICIAL CARGADO (BASE ÚNICA):
{texto_curriculo}

REGLA: Toda la planificación debe anclarse EXCLUSIVAMENTE a este documento.
"""
                else:
                    anclaje = "Genera la planificación basándote en los insumos proporcionados y los estándares curriculares del MINERD para el nivel de secundaria."

                prompt_maestro = f"""Actúa como un Doctor en Educación, Experto en Diseño Instruccional y Asesor Curricular del MINERD de la República Dominicana, especialista en áreas académicas de nivel secundario.

INSUMOS:
- Área Curricular: {area}
- Grado: {grado}
- Título de la Unidad: {titulo_unidad}
- Duración: {duracion_unidad}
- Cantidad de sesiones a planificar: {cantidad_sesiones}
- Competencia: {competencia}
- Indicadores de Logro: {indicadores_logro}
- Situación de Aprendizaje: {situacion_aprendizaje}
- Contenidos curriculares: {contenidos_decl}
- Perfil del grupo: {perfil_grupo}

{anclaje}

REGLAS DE DISEÑO:

1. COMPETENCIAS E INDICADORES: Desglosa la competencia en los indicadores específicos que se trabajarán en esta unidad. Si se proporcionaron, úsalos y compleméntalos.

2. CONTENIDOS (3 tipos): Organiza los contenidos en: Conceptuales (saber), Procedimentales (saber hacer) y Actitudinales (ser). Extraídos del currículo.

3. SITUACIÓN DE APRENDIZAJE: Contextualiza la situación como eje integrador de toda la unidad. Justifica su pertinencia.

4. SECUENCIA DIDÁCTICA ({cantidad_sesiones} sesiones): Para CADA sesión define:
   - Número y título de la sesión
   - Fase de la unidad (Inicio / Desarrollo / Cierre de la unidad)
   - Contenidos específicos de la sesión
   - Momentos didácticos (Inicio, Desarrollo, Cierre) con actividades concretas y tiempos realistas que sumen ~50 min
   - Recursos de la sesión
   - Evaluación de la sesión (técnica e instrumento)
   - Indicador de logro que se aborda

5. RECURSOS GENERALES: Lista consolidada de recursos para toda la unidad.

6. EVALUACIÓN INTEGRAL: Define la evaluación formativa y sumativa de la unidad con instrumentos variados.

7. ADAPTACIONES NEAE: Adaptaciones específicas según el perfil del grupo.

8. PRODUCTO FINAL: Describa el producto o desempeño final esperado del estudiante al concluir la unidad.

REGLA CRÍTICA DE JSON: NO uses saltos de línea literales (\\n) dentro de los valores. Une con espacios. NO uses comillas dobles internas.

FORMATO DE SALIDA ESTRICTO (JSON NATIVO OBLIGATORIO):
{{
  "COMPETENCIAS_INDICADORES": {{
    "COMPETENCIA": "Competencia fundamental del área",
    "INDICADORES": ["Indicador 1...", "Indicador 2...", "Indicador 3..."]
  }},
  "CONTENIDOS": {{
    "CONCEPTUALES": ["Contenido conceptual 1...", "Contenido conceptual 2..."],
    "PROCEDIMENTALES": ["Contenido procedimental 1...", "Contenido procedimental 2..."],
    "ACTITUDINALES": ["Contenido actitudinal 1...", "Contenido actitudinal 2..."]
  }},
  "SITUACION_APRENDIZAJE": {{
    "ENUNCIADO": "Contexto integrador de la unidad...",
    "PERTINENCIA": "Justificación de por qué esta situación es significativa para los estudiantes"
  }},
  "SECUENCIA_DIDACTICA": [
    {{
      "NUMERO": 1,
      "TITULO_SESION": "Título motivador",
      "FASE_UNIDAD": "Inicio de la unidad",
      "CONTENIDO_SESION": "Contenido específico abordado",
      "INDICADOR_SESION": "Indicador de logro trabajado",
      "MOMENTOS": {{
        "INICIO_TIEMPO": "10 min",
        "INICIO_ACTIVIDAD": "Descripción de la actividad de inicio",
        "DESARROLLO_TIEMPO": "30 min",
        "DESARROLLO_ACTIVIDAD": "Descripción de la actividad central",
        "CIERRE_TIEMPO": "10 min",
        "CIERRE_ACTIVIDAD": "Descripción del cierre y metacognición"
      }},
      "RECURSOS": "Recursos de esta sesión",
      "EVALUACION": "Técnica e instrumento de evaluación de la sesión"
    }}
  ],
  "RECURSOS_GENERALES": "Lista consolidada de recursos para toda la unidad",
  "EVALUACION_INTEGRAL": {{
    "FORMATIVA": "Evaluación formativa: técnicas e instrumentos durante la unidad",
    "SUMATIVA": "Evaluación sumativa: instrumento final y criterios",
    "CRITERIOS": ["Criterio 1...", "Criterio 2...", "Criterio 3..."]
  }},
  "ADAPTACIONES_NEAE": "Adaptaciones específicas según el perfil del grupo",
  "PRODUCTO_FINAL": "Producto o desempeño final esperado al concluir la unidad"
}}
"""
                respuesta_ia, corte_tokens, se_reintento = solicitar_con_reintento(
                    proveedor_ia, api_key_usuario, modelo_seleccionado, prompt_maestro
                )

                datos, truncado, texto_limpio = extraer_json_seguro(respuesta_ia)

                if not datos:
                    st.error("❌ La IA no devolvió un JSON válido. Sube el límite de tokens en la barra lateral y reintent.")
                    st.stop()

                if truncado:
                    st.warning("⚠️ JSON reparado parcialmente. Algunos campos pueden estar incompletos.")

                if se_reintento and not truncado:
                    st.info("ℹ️ Reintento automático con más tokens tuvo éxito.")

                # --- Extracción ---
                comp_ind = datos.get("COMPETENCIAS_INDICADORES", {})
                contenidos = datos.get("CONTENIDOS", {})
                situacion = datos.get("SITUACION_APRENDIZAJE", {})
                secuencia = datos.get("SECUENCIA_DIDACTICA", [])
                recursos_gen = datos.get("RECURSOS_GENERALES", "")
                eval_integral = datos.get("EVALUACION_INTEGRAL", {})
                adapt_neae = datos.get("ADAPTACIONES_NEAE", "")
                producto_final = datos.get("PRODUCTO_FINAL", "")

                # ═══════════════════════════════════════════════
                # CONSTRUCCIÓN DEL WORD
                # ═══════════════════════════════════════════════
                doc = Document()
                doc.styles['Normal'].font.name = 'Calibri'
                doc.styles['Normal'].font.size = Pt(11)

                for section in doc.sections:
                    section.left_margin = Inches(0.75)
                    section.right_margin = Inches(0.75)

                def shade_cell(cell, color):
                    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color))
                    cell._tc.get_or_add_tcPr().append(shd)

                # ── I. ENCABEZADO ──
                p_enc = doc.add_paragraph()
                p_enc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_enc.add_run("MINISTERIO DE EDUCACIÓN DE LA REPÚBLICA DOMINICANA\n").bold = True
                p_enc.add_run("PLANIFICACIÓN DE UNIDAD DE APRENDIZAJE — NIVEL SECUNDARIO\n").bold = True

                # ── II. DATOS GENERALES ──
                doc.add_heading("Datos Generales", level=2)
                t_datos = doc.add_table(rows=5, cols=4)
                t_datos.style = 'Table Grid'
                datos_grid = [
                    ("Centro Educativo", centro, "Año Escolar", ano_escolar),
                    ("Docente", docente, "Grado", grado),
                    ("Área Curricular", area, "Sección", seccion),
                    ("Unidad", titulo_unidad, "Duración", duracion_unidad),
                    ("Sesiones", str(cantidad_sesiones), "", ""),
                ]
                for r, (l1, v1, l2, v2) in enumerate(datos_grid):
                    for c, (lbl, val) in enumerate([(l1, v1), (l2, v2)]):
                        if lbl:
                            t_datos.cell(r, c*2).text = lbl
                            t_datos.cell(r, c*2).paragraphs[0].runs[0].bold = True
                            shade_cell(t_datos.cell(r, c*2), "F1F5F9")
                        t_datos.cell(r, c*2+1).text = val
                doc.add_paragraph()

                # ── III. COMPETENCIAS E INDICADORES ──
                doc.add_heading("Competencias e Indicadores de Logro", level=2)
                if comp_ind.get("COMPETENCIA"):
                    p_comp = doc.add_paragraph()
                    p_comp.add_run("Competencia Fundamental: ").bold = True
                    p_comp.add_run(str(comp_ind.get("COMPETENCIA", "")))
                indicadores = comp_ind.get("INDICADORES", [])
                if indicadores:
                    doc.add_paragraph("Indicadores de Logro:")
                    for ind in indicadores:
                        doc.add_paragraph(str(ind), style='List Bullet')
                doc.add_paragraph()

                # ── IV. CONTENIDOS ──
                doc.add_heading("Contenidos", level=2)
                t_cont = doc.add_table(rows=1, cols=3)
                t_cont.style = 'Table Grid'
                hdr_c = t_cont.rows[0].cells
                for i, txt in enumerate(["Conceptuales", "Procedimentales", "Actitudinales"]):
                    hdr_c[i].text = txt
                    hdr_c[i].paragraphs[0].runs[0].bold = True
                    shade_cell(hdr_c[i], "DBEAFE")
                conceptuales = contenidos.get("CONCEPTUALES", [])
                procedimentales = contenidos.get("PROCEDIMENTALES", [])
                actitudinales = contenidos.get("ACTITUDINALES", [])
                max_filas = max(len(conceptuales), len(procedimentales), len(actitudinales), 1)
                for i in range(max_filas):
                    row_c = t_cont.add_row().cells
                    row_c[0].text = str(conceptuales[i]) if i < len(conceptuales) else ""
                    row_c[1].text = str(procedimentales[i]) if i < len(procedimentales) else ""
                    row_c[2].text = str(actitudinales[i]) if i < len(actitudinales) else ""
                doc.add_paragraph()

                # ── V. SITUACIÓN DE APRENDIZAJE ──
                doc.add_heading("Situación de Aprendizaje", level=2)
                p_sit = doc.add_paragraph()
                p_sit.add_run("Contexto integrador: ").bold = True
                p_sit.add_run(str(situacion.get("ENUNCIADO", "")))
                p_sit.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                if situacion.get("PERTINENCIA"):
                    p_pert = doc.add_paragraph()
                    p_pert.add_run("Pertinencia: ").bold = True
                    p_pert.add_run(str(situacion.get("PERTINENCIA", "")))
                    p_pert.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                doc.add_paragraph()

                # ── VI. SECUENCIA DIDÁCTICA ──
                doc.add_heading("Secuencia Didáctica", level=2)
                for sesion in secuencia:
                    num = sesion.get("NUMERO", "")
                    titulo_s = sesion.get("TITULO_SESION", "")
                    fase = sesion.get("FASE_UNIDAD", "")

                    doc.add_heading(f"Sesión {num}: {titulo_s}", level=3)

                    # Info de la sesión
                    t_s = doc.add_table(rows=3, cols=2)
                    t_s.style = 'Table Grid'
                    t_s.cell(0, 0).text = "Fase de la unidad"
                    t_s.cell(0, 0).paragraphs[0].runs[0].bold = True
                    shade_cell(t_s.cell(0, 0), "F1F5F9")
                    t_s.cell(0, 1).text = str(fase)
                    t_s.cell(1, 0).text = "Contenido"
                    t_s.cell(1, 0).paragraphs[0].runs[0].bold = True
                    shade_cell(t_s.cell(1, 0), "F1F5F9")
                    t_s.cell(1, 1).text = str(sesion.get("CONTENIDO_SESION", ""))
                    t_s.cell(2, 0).text = "Indicador de logro"
                    t_s.cell(2, 0).paragraphs[0].runs[0].bold = True
                    shade_cell(t_s.cell(2, 0), "F1F5F9")
                    t_s.cell(2, 1).text = str(sesion.get("INDICADOR_SESION", ""))

                    doc.add_paragraph()

                    # Momentos didácticos
                    momentos = sesion.get("MOMENTOS", {})
                    t_m = doc.add_table(rows=4, cols=2)
                    t_m.style = 'Table Grid'
                    hdr_m = t_m.rows[0].cells
                    hdr_m[0].text = "Momento"
                    hdr_m[1].text = "Actividades"
                    hdr_m[0].paragraphs[0].runs[0].bold = True
                    hdr_m[1].paragraphs[0].runs[0].bold = True
                    shade_cell(hdr_m[0], "E2E8F0")
                    shade_cell(hdr_m[1], "E2E8F0")

                    t_m.cell(1, 0).text = f"Inicio\n({momentos.get('INICIO_TIEMPO', '')})"
                    t_m.cell(1, 0).paragraphs[0].runs[0].bold = True
                    t_m.cell(1, 1).text = str(momentos.get("INICIO_ACTIVIDAD", ""))

                    t_m.cell(2, 0).text = f"Desarrollo\n({momentos.get('DESARROLLO_TIEMPO', '')})"
                    t_m.cell(2, 0).paragraphs[0].runs[0].bold = True
                    t_m.cell(2, 1).text = str(momentos.get("DESARROLLO_ACTIVIDAD", ""))

                    t_m.cell(3, 0).text = f"Cierre\n({momentos.get('CIERRE_TIEMPO', '')})"
                    t_m.cell(3, 0).paragraphs[0].runs[0].bold = True
                    t_m.cell(3, 1).text = str(momentos.get("CIERRE_ACTIVIDAD", ""))

                    doc.add_paragraph()

                    p_rec_s = doc.add_paragraph()
                    p_rec_s.add_run("Recursos: ").bold = True
                    p_rec_s.add_run(str(sesion.get("RECURSOS", "")))

                    p_eval_s = doc.add_paragraph()
                    p_eval_s.add_run("Evaluación: ").bold = True
                    p_eval_s.add_run(str(sesion.get("EVALUACION", "")))

                    doc.add_paragraph("_" * 50)
                    doc.add_paragraph()

                # ── VII. RECURSOS GENERALES ──
                if recursos_gen:
                    doc.add_heading("Recursos Generales de la Unidad", level=2)
                    doc.add_paragraph(str(recursos_gen))
                    doc.add_paragraph()

                # ── VIII. EVALUACIÓN INTEGRAL ──
                doc.add_heading("Evaluación Integral de la Unidad", level=2)
                if eval_integral.get("FORMATIVA"):
                    p_f = doc.add_paragraph()
                    p_f.add_run("Evaluación Formativa: ").bold = True
                    p_f.add_run(str(eval_integral.get("FORMATIVA", "")))
                    p_f.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                if eval_integral.get("SUMATIVA"):
                    p_s = doc.add_paragraph()
                    p_s.add_run("Evaluación Sumativa: ").bold = True
                    p_s.add_run(str(eval_integral.get("SUMATIVA", "")))
                    p_s.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                criterios_eval = eval_integral.get("CRITERIOS", [])
                if criterios_eval:
                    doc.add_paragraph("Criterios de Evaluación:")
                    for crit in criterios_eval:
                        doc.add_paragraph(str(crit), style='List Bullet')
                doc.add_paragraph()

                # ── IX. ADAPTACIONES NEAE ──
                doc.add_heading("Adaptaciones para NEAE", level=2)
                doc.add_paragraph(str(adapt_neae))
                doc.add_paragraph()

                # ── X. PRODUCTO FINAL ──
                if producto_final:
                    doc.add_heading("Producto Final de la Unidad", level=2)
                    p_prod = doc.add_paragraph(str(producto_final))
                    p_prod.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    doc.add_paragraph()

                # ── Firmas ──
                doc.add_paragraph("\n\n")
                t_firmas = doc.add_table(rows=2, cols=2)
                t_firmas.cell(0, 0).text = "__________________________"
                t_firmas.cell(0, 1).text = "__________________________"
                t_firmas.cell(1, 0).text = "Docente"
                t_firmas.cell(1, 1).text = "Coordinador/a Pedagógico"
                for row in t_firmas.rows:
                    for cell in row.cells:
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                col_m1, col_m2, col_m3 = st.columns(3)
                with col_m1: st.metric("📚 Sesiones planificadas", len(secuencia))
                with col_m2: st.metric("🎯 Indicadores", len(indicadores))
                with col_m3: st.metric("📋 Criterios eval.", len(criterios_eval))

                st.success(f"✅ ¡Plan de Unidad Académica generado con éxito! ({len(secuencia)} sesiones)")

                st.download_button(
                    label="📥 Descargar Plan de Unidad (.docx)",
                    data=buffer,
                    file_name=f"Plan_Unidad_{area}_{titulo_unidad[:15]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

            except ResourceExhausted:
                st.error("❌ Se alcanzó el límite de API.")
            except Exception as e:
                st.error(f"⚠️ Error de procesamiento: {e}")