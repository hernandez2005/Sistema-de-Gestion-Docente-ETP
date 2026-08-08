import streamlit as st
import google.generativeai as genai
from openai import OpenAI
import re
import base64
import unicodedata
import datetime
import streamlit.components.v1 as components
from google.api_core.exceptions import ResourceExhausted
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

st.set_page_config(page_title="Fábrica de Simuladores ETP", page_icon="⚡", layout="wide")

# =========================================================================
# ESTILOS CSS DE LA PLATAFORMA
# =========================================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #F8FAFC; color: #0F172A; }
    .main-header { font-size: 2.4rem; font-weight: 800; color: #1E293B; text-align: center; margin-bottom: 5px; line-height: 1.2; letter-spacing: -0.02em; }
    .sub-header { text-align: center; color: #475569; font-size: 1.15rem; font-weight: 400; margin-bottom: 40px; }
    [data-testid="stForm"] { background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05); padding: 35px; margin-bottom: 25px; }
    .section-title { color: #0F172A; font-weight: 700; font-size: 1.3rem; border-bottom: 2px solid #E2E8F0; padding-bottom: 10px; margin-top: 25px; margin-bottom: 20px; }
    div.stButton > button:first-child { background-color: #0F172A !important; color: #FFFFFF !important; border: none !important; border-radius: 6px !important; font-weight: 600 !important; padding: 12px 24px !important; width: 100%; transition: all 0.2s ease; }
    div.stButton > button:first-child:hover { background-color: #334155 !important; transform: translateY(-1px); }
    .hist-card { background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 14px 16px; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# =========================================================================
# TEMAS VISUALES (design tokens que se inyectan literalmente en el prompt)
# =========================================================================
TEMAS_VISUALES = {
    "Corporativo Moderno": """
- Fondo general: gradiente sutil de #F1F5F9 a #E2E8F0.
- Tarjetas: blancas (#FFFFFF), border-radius 12px, box-shadow 0 10px 25px rgba(0,0,0,0.08).
- Color de acento primario: #2563EB (azul). Botones primarios con ese color y hover #1D4ED8.
- Texto principal: #0F172A. Texto secundario: #64748B.
- Tipografía: system-ui, -apple-system, sans-serif.""",
    "Oscuro Tech (Dark Mode)": """
- Fondo general: #0B1120 con un leve gradiente radial a #111827.
- Tarjetas: #1E293B, border-radius 14px, borde 1px solid #334155, sombra 0 10px 30px rgba(0,0,0,0.4).
- Color de acento primario: #22D3EE (cian) para bordes activos, botones y resultados destacados.
- Texto principal: #E2E8F0. Texto secundario: #94A3B8.
- Añade un ligero efecto glow (box-shadow con el color de acento) en inputs enfocados (:focus).""",
    "Educativo Vibrante": """
- Fondo general: gradiente diagonal de #FFF7ED a #FEF3C7 muy suave.
- Tarjetas: blancas con border-radius 16px y sombra suave, bordes superiores de 4px con color de acento.
- Color de acento primario: #F97316 (naranja) combinado con #10B981 (verde) para estados de éxito.
- Texto principal: #1F2937. Usa emojis o iconos SVG simples para reforzar contexto educativo sin saturar.
- Tipografía con buen tamaño (mínimo 16px) pensada para estudiantes.""",
    "Minimalista Suizo": """
- Fondo general: blanco puro (#FFFFFF) o gris muy claro (#FAFAFA).
- Tarjetas: bordes finos 1px solid #E5E5E5, sin sombras pronunciadas, border-radius 4px.
- Color de acento primario: #111111 (negro) o un único color de acento (#DC2626 rojo) usado con moderación.
- Tipografía protagonista: jerarquía tipográfica clara (títulos grandes en negrita, cuerpo ligero).
- Mucho espacio en blanco (whitespace), grid limpio, cero decoración innecesaria."""
}

NIVELES_COMPLEJIDAD = {
    "Básico (cálculo directo)": "El simulador debe resolver una lógica directa y sencilla. Prioriza la claridad sobre la cantidad de features. Un único flujo principal, sin pasos intermedios.",
    "Intermedio (con validaciones)": "El simulador debe incluir validación de inputs en tiempo real, mensajes de error claros, y un botón de 'Reiniciar' que limpie el formulario y los resultados.",
    "Avanzado (multi-paso / interactivo)": "El simulador debe sentirse como una mini-aplicación: puede tener pasos o pestañas internas, actualización de resultados en tiempo real (evento 'input', no solo 'click' cuando aplique), animaciones de transición entre estados, y un resumen final destacado."
}

# =========================================================================
# UTILIDADES DE SANEAMIENTO / REPARACIÓN
# =========================================================================
def sanear_nombre_archivo(texto):
    """Convierte un texto libre en un nombre de archivo seguro (sin acentos ni caracteres especiales)."""
    if not texto:
        texto = "simulador"
    texto_normalizado = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    texto_limpio = re.sub(r'[^A-Za-z0-9_\- ]', '', texto_normalizado).strip()
    texto_limpio = re.sub(r'\s+', '_', texto_limpio)
    return texto_limpio[:60] if texto_limpio else "simulador"


def _balancear_etiqueta(codigo, tag):
    """Cierra una etiqueta (style/script/head/body/html) que quedó abierta por
    un corte de tokens. Evita que un <style> o <script> sin cerrar se trague
    el resto del documento."""
    abiertas = len(re.findall(rf'<{tag}\b[^>]*>', codigo, flags=re.IGNORECASE))
    cerradas = len(re.findall(rf'</{tag}\s*>', codigo, flags=re.IGNORECASE))
    faltantes = abiertas - cerradas
    if faltantes > 0:
        codigo += f'</{tag}>' * faltantes
    return codigo


def extraer_html_ui_pro(texto_ia):
    """Extrae el HTML de estructura (fase 1), repara truncamientos.
    Devuelve (codigo_html, fue_truncado)."""
    if not texto_ia:
        return "<h2 style='color:red;'>Error: La IA no generó código (respuesta vacía).</h2>", False

    codigo = re.sub(r'```(?:html|xml)?', '', texto_ia, flags=re.IGNORECASE).replace('```', '').strip()

    inicio = re.search(r'<!DOCTYPE html>|<html\b', codigo, flags=re.IGNORECASE)
    if inicio:
        codigo = codigo[inicio.start():]

    match = re.search(r'(<!DOCTYPE html>.*</html>|<html.*</html>)', codigo, flags=re.DOTALL | re.IGNORECASE)
    truncado = False

    if match:
        codigo = match.group(1)
    elif "<html" in codigo.lower() or "<!doctype" in codigo.lower():
        truncado = True
        codigo = _balancear_etiqueta(codigo, 'style')
        codigo = _balancear_etiqueta(codigo, 'script')
        if '</head>' not in codigo.lower() and '<head' in codigo.lower():
            codigo += '</head>'
        if '<body' not in codigo.lower():
            codigo += '<body><div style="padding:2rem;color:#b91c1c;">⚠️ La respuesta de la IA se cortó antes de generar el contenido visible.</div></body>'
        elif '</body>' not in codigo.lower():
            codigo += '</body>'
        if '</html>' not in codigo.lower():
            codigo += '</html>'

    if "<body" not in codigo.lower():
        codigo = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>body {{ font-family: system-ui, -apple-system, sans-serif; background-color: #f8fafc; color: #0f172a; padding: 2rem; }}</style>
</head>
<body>
    {codigo}
</body>
</html>"""

    if "viewport" not in codigo.lower():
        codigo = re.sub(
            r'(<head[^>]*>)',
            r'\1\n<meta name="viewport" content="width=device-width, initial-scale=1.0">',
            codigo, count=1, flags=re.IGNORECASE
        )

    return codigo, truncado


def extraer_script(texto_ia):
    """Extrae el bloque <script> de la respuesta de la fase de lógica (fase 2).
    Devuelve (script_html, fue_truncado)."""
    if not texto_ia:
        return "<script>console.error('La IA no devolvió lógica JavaScript.');</script>", False

    codigo = re.sub(r'```(?:javascript|js|html)?', '', texto_ia, flags=re.IGNORECASE).replace('```', '').strip()

    match = re.search(r'(<script[^>]*>.*</script>)', codigo, flags=re.DOTALL | re.IGNORECASE)
    truncado = False

    if match:
        codigo = match.group(1)
    elif '<script' in codigo.lower():
        truncado = True
        codigo = _balancear_etiqueta(codigo, 'script')
    elif codigo:
        codigo = f"<script>\n{codigo}\n</script>"
    else:
        codigo = "<script>console.error('La IA devolvió una respuesta vacía en la fase de lógica.');</script>"

    return codigo, truncado


def insertar_script(codigo_html, script_bloque):
    """Inserta el bloque de lógica justo antes de </body>, eliminando primero
    cualquier <script> preexistente en la estructura (por seguridad).

    IMPORTANTE: el reemplazo se pasa como función (lambda), no como string.
    Si se pasa como string, re.sub intenta interpretar backslashes tipo
    \\1, \\g<...> dentro del reemplazo — y el JavaScript generado por la IA
    casi siempre contiene expresiones regulares propias (\\s, \\d, \\w, etc.)
    que Python confunde con referencias de grupo inválidas, provocando
    'bad escape \\s'. Una función como reemplazo inserta el texto tal cual,
    sin ningún procesamiento de escapes."""
    sin_scripts = re.sub(r'<script[^>]*>.*?</script>', '', codigo_html, flags=re.DOTALL | re.IGNORECASE)
    if re.search(r'</body>', sin_scripts, flags=re.IGNORECASE):
        return re.sub(r'</body>', lambda m: script_bloque + '\n</body>', sin_scripts, count=1, flags=re.IGNORECASE)
    elif re.search(r'</html>', sin_scripts, flags=re.IGNORECASE):
        return re.sub(r'</html>', lambda m: script_bloque + '\n</html>', sin_scripts, count=1, flags=re.IGNORECASE)
    return sin_scripts + script_bloque


def generar_data_uri(codigo_html):
    b64 = base64.b64encode(codigo_html.encode('utf-8')).decode('utf-8')
    return f"data:text/html;base64,{b64}"


# =========================================================================
# LLAMADAS A API (DETERMINISTAS, CON DETECCIÓN DE TRUNCAMIENTO)
# =========================================================================
@retry(retry=retry_if_exception_type(ResourceExhausted), wait=wait_exponential(multiplier=2, min=4, max=20), stop=stop_after_attempt(5), reraise=True)
def solicitar_gemini_html(api_key, modelo, prompt, max_tokens=8192):
    """Devuelve (texto, corte_por_limite_de_tokens)."""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(modelo)
    respuesta = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(max_output_tokens=max_tokens, temperature=0.15)
    )
    corte_tokens = False
    try:
        finish_reason = respuesta.candidates[0].finish_reason
        corte_tokens = (str(finish_reason).upper().find("MAX_TOKENS") != -1) or finish_reason == 2
    except Exception:
        pass
    try:
        texto = respuesta.text
    except Exception:
        texto = ""
    return texto, corte_tokens


def solicitar_openai_html(api_key, modelo, prompt, max_tokens=8192):
    """Devuelve (texto, corte_por_limite_de_tokens)."""
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=modelo,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.15,
        max_tokens=max_tokens
    )
    corte_tokens = response.choices[0].finish_reason == "length"
    return response.choices[0].message.content, corte_tokens


def solicitar_html(proveedor, api_key, modelo, prompt, max_tokens=8192):
    if proveedor == "Google Gemini":
        return solicitar_gemini_html(api_key, modelo, prompt, max_tokens)
    return solicitar_openai_html(api_key, modelo, prompt, max_tokens)


def solicitar_con_reintento(proveedor, api_key, modelo, prompt, max_tokens, tope=32000):
    """Igual que solicitar_html, pero si detecta un corte por límite de tokens
    reintenta UNA vez automáticamente con el doble de presupuesto.
    Devuelve (texto, sigue_truncado, se_reintento)."""
    texto, truncado = solicitar_html(proveedor, api_key, modelo, prompt, max_tokens)
    reintentado = False
    if truncado and max_tokens < tope:
        max_tokens_2 = min(max_tokens * 2, tope)
        texto2, truncado2 = solicitar_html(proveedor, api_key, modelo, prompt, max_tokens_2)
        reintentado = True
        if len(texto2 or "") >= len(texto or ""):
            texto, truncado = texto2, truncado2
    return texto, truncado, reintentado


# =========================================================================
# CONSTRUCCIÓN DE PROMPTS (DOS FASES: ESTRUCTURA → LÓGICA)
# =========================================================================
def construir_prompt_estructura(politecnico, docente, modulo, tema, descripcion, tema_visual, nivel):
    tokens_diseno = TEMAS_VISUALES[tema_visual]
    guia_nivel = NIVELES_COMPLEJIDAD[nivel]

    return f"""Actúa como un Diseñador UI/UX Senior especializado en simuladores educativos.

ESTA ES LA FASE 1 DE 2: SOLO ESTRUCTURA VISUAL (HTML + CSS). NO escribas JavaScript todavía, la lógica se generará en un segundo paso por separado.

CONTEXTO EDUCATIVO:
- Institución: {politecnico}
- Docente: {docente}
- Asignatura: {modulo}
- Tema: {tema}

REQUERIMIENTOS FUNCIONALES QUE EL SIMULADOR DEBERÁ CUMPLIR (úsalos solo para decidir qué inputs, botones y contenedores de resultado necesitas crear; NO los implementes en JS ahora):
{descripcion}

NIVEL DE COMPLEJIDAD ESPERADO (afecta cuántos campos/pasos/controles debes crear):
{guia_nivel}

SISTEMA DE DISEÑO A APLICAR (tema visual: "{tema_visual}"):
{tokens_diseno}
- Escala de espaciado consistente (múltiplos de 4px u 8px).
- Al menos 2 niveles de jerarquía tipográfica claros.
- Estados :hover / :focus con transición suave (transition: all 0.2s ease).
- Layout responsivo con CSS Grid o Flexbox y media query para max-width: 600px.

REGLAS INQUEBRANTABLES:
1. NO uses la etiqueta <script> ni JavaScript de ningún tipo en esta fase. Solo HTML + CSS.
2. Cero DEPENDENCIAS EXTERNAS: prohibido Tailwind, Bootstrap, Google Fonts o cualquier CDN. Si necesitas iconos, dibújalos como SVG inline.
3. NO uses <form>, NO botones type="submit", NO alert()/confirm()/prompt().
4. Cada input, select, botón, y cada contenedor donde luego se inyectarán resultados o errores DEBE tener un atributo id único, descriptivo, en minúsculas con guiones. Ejemplos: id="input-ip", id="btn-calcular", id="resultado", id="error", id="btn-reiniciar".
5. Cada botón que dispare una acción debe incluir el atributo onclick="" (vacío, se completará en la fase 2). Ejemplo: <button id="btn-calcular" onclick="">Calcular</button>.
6. Justo después de abrir <body>, agrega un comentario HTML con el contrato de ids exactamente en este formato (uno por línea):
<!-- CONTRATO_JS
id-del-elemento: descripción breve de su función
id-del-elemento-2: descripción breve de su función
-->
   Incluye en ese contrato TODOS los ids funcionales: inputs, selects, botones y contenedores de resultado/error.
7. Accesibilidad: <label for="..."> asociado a cada input, buen contraste, texto de al menos 14px.
8. Todo el texto visible debe estar en español neutro.
9. Devuelve ÚNICA Y EXCLUSIVAMENTE el HTML crudo empezando por <!DOCTYPE html>, sin explicaciones antes o después, sin comentarios de markdown, y SIN etiqueta <script>.

CÓDIGO HTML (solo estructura, sin JavaScript):
"""


def construir_prompt_logica(html_estructura, descripcion, nivel):
    guia_nivel = NIVELES_COMPLEJIDAD[nivel]

    return f"""Actúa como un Desarrollador JavaScript Senior.

ESTA ES LA FASE 2 DE 2: SOLO LÓGICA. Te doy la estructura HTML/CSS ya terminada de un simulador educativo. Tu única tarea es escribir el JavaScript que lo haga funcionar por completo, enganchándote a los ids que YA EXISTEN en ese HTML (revisa el comentario CONTRATO_JS dentro del <body>).

ESTRUCTURA HTML YA EXISTENTE (no la repitas ni la modifiques, solo léela para saber qué ids usar):
{html_estructura}

REQUERIMIENTOS FUNCIONALES QUE TU JAVASCRIPT DEBE CUMPLIR POR COMPLETO:
{descripcion}

NIVEL DE COMPLEJIDAD ESPERADO:
{guia_nivel}

REGLAS INQUEBRANTABLES:
1. Usa EXACTAMENTE los ids que ya existen en el HTML dado (los del comentario CONTRATO_JS). No inventes ids nuevos ni asumas que existen otros.
2. No dependas del atributo onclick="" vacío de los botones: dentro de tu script, usa document.getElementById('id-del-boton').addEventListener('click', funcion) para engancharte a cada botón.
3. Valida los inputs antes de calcular (vacíos, tipo incorrecto, fuera de rango); si son inválidos, muestra el mensaje en el contenedor de error correspondiente y NO ejecutes el cálculo. Nunca uses alert(), confirm() ni prompt().
4. Implementa el 100% de la lógica solicitada arriba, sin dejar funciones a medias ni placeholders tipo "// TODO".
5. Envuelve todo tu código en un listener DOMContentLoaded o en una IIFE, para asegurar que el DOM ya existe.
6. Todo tu código va dentro de UNA sola etiqueta <script>...</script>. Nada de HTML, nada de CSS, nada de explicaciones ni comentarios de markdown antes o después.
7. Prioriza que la lógica esté completa y funcional por encima de comentarios extensos o nombres de variable elaborados: sé conciso.

DEVUELVE ÚNICAMENTE:
<script>
...tu código JavaScript completo...
</script>
"""


def construir_prompt_refinamiento(codigo_actual, instrucciones_ajuste):
    return f"""Eres el mismo Diseñador/Desarrollador Senior que generó el siguiente simulador web (Single-File HTML). El usuario quiere AJUSTES puntuales sin romper lo que ya funciona.

CÓDIGO HTML ACTUAL:
{codigo_actual}

AJUSTES SOLICITADOS POR EL USUARIO:
{instrucciones_ajuste}

REGLAS:
1. Conserva toda la funcionalidad existente que el usuario no pidió cambiar; no borres lógica que ya funcionaba.
2. Sigue respetando: cero dependencias externas, cero <form>, cero alert(), todo el CSS en un único <style> y todo el JS en un único <script>.
3. Si el ajuste implica lógica nueva, impleméntala por completo (nada de placeholders).
4. Devuelve ÚNICA Y EXCLUSIVAMENTE el código HTML completo y actualizado, empezando por <!DOCTYPE html>, sin explicaciones ni comentarios de markdown.

CÓDIGO HTML ACTUALIZADO:
"""


# =========================================================================
# ESTADO DE SESIÓN
# =========================================================================
if "codigo_actual" not in st.session_state:
    st.session_state.codigo_actual = None
if "tema_actual" not in st.session_state:
    st.session_state.tema_actual = None
if "historial" not in st.session_state:
    st.session_state.historial = []

# =========================================================================
# CONFIGURACIÓN CENTRALIZADA (desde main.py)
# =========================================================================
api_key_usuario = st.session_state.get("api_key_global", "")
proveedor_ia = st.session_state.get("proveedor_ia_global", "Google Gemini")
modelo_seleccionado = st.session_state.get("modelo_global", "gemini-2.5-flash")

# =========================================================================
# PANEL LATERAL (contenido específico de esta página)
# =========================================================================
with st.sidebar:
    st.markdown("##### ⚡ Fábrica de Simuladores")
    if not api_key_usuario:
        st.error("🔒 Configura tu API Key en la página de Inicio")
    else:
        st.success(f"✅ {proveedor_ia} · {modelo_seleccionado}")

    st.markdown("---")
    st.subheader("🎨 Presentación")
    tema_visual = st.selectbox("Tema visual del simulador:", list(TEMAS_VISUALES.keys()), key="tema_visual_sim")
    nivel_complejidad = st.selectbox("Nivel de complejidad:", list(NIVELES_COMPLEJIDAD.keys()), index=1, key="nivel_sim")
    altura_preview = st.slider("Altura de la vista previa (px):", 400, 1200, 800, step=50)

    with st.expander("🛠️ Opciones avanzadas"):
        st.caption("El simulador se genera en 2 llamadas a la IA: primero la estructura (HTML+CSS), luego la lógica (JavaScript). Esto reduce muchísimo el riesgo de que la lógica quede a medias.")
        max_tokens_sim = st.slider(
            "Límite de tokens por llamada:", 4096, 32000, 10000, step=1024,
            help="Presupuesto de tokens para CADA una de las dos llamadas (estructura y lógica). Si el simulador sigue incompleto, súbelo. Ten en cuenta que algunos modelos (ej. gpt-3.5-turbo) tienen un tope propio más bajo."
        )
        modo_debug = st.checkbox(
            "🐛 Modo depuración (ver respuestas crudas de la IA)",
            help="Muestra el texto exacto que devolvió la IA en cada fase, antes de procesarlo."
        )

    if st.session_state.historial:
        st.markdown("---")
        st.subheader("🗂️ Historial de esta sesión")
        for i, item in enumerate(reversed(st.session_state.historial[-8:])):
            with st.container():
                st.markdown(f"<div class='hist-card'><b>{item['tema']}</b><br><small>{item['hora']}</small></div>", unsafe_allow_html=True)
                st.download_button(
                    "Descargar",
                    data=item['codigo'].encode('utf-8'),
                    file_name=f"Simulador_{sanear_nombre_archivo(item['tema'])}.html",
                    mime="text/html",
                    key=f"hist_dl_{i}_{item['hora']}"
                )

# =========================================================================
# ENCABEZADO
# =========================================================================
st.markdown('<div class="main-header">Fábrica de Simuladores Web Interactivos</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Aplicaciones web modernas, elegantes y 100% funcionales para la ETP</div>', unsafe_allow_html=True)

# =========================================================================
# FORMULARIO PRINCIPAL
# =========================================================================
with st.form("form_simulador", clear_on_submit=False):

    st.markdown('<div class="section-title">🏫 1. Datos Institucionales</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        politecnico = st.text_input("Politécnico", value="Politécnico Salesiano Arquides Calderón")
        docente = st.text_input("Docente", value="Ing. Bernardo Antonio Hernández Batista")
    with col2:
        modulo = st.text_input("Módulo / Asignatura", placeholder="Ej: Redes LAN o Sistemas Operativos")
        tema = st.text_input("Tema a Simular", placeholder="Ej: Calculadora de Subredes IPv4")

    st.markdown('<div class="section-title">⚙️ 2. Especificación de Experiencia de Usuario y Lógica</div>', unsafe_allow_html=True)
    descripcion = st.text_area(
        "Detalla qué debe hacer el simulador (Cálculos, inputs, resultados esperados):",
        height=150,
        placeholder="Ej: El usuario ingresa la IP y la máscara. Al hacer clic, debe calcular de inmediato la IP de Red, Broadcast y cantidad de hosts utilizables, presentándolo en una tabla."
    )

    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button("🚀 Generar Simulador de Alta Gama")

# =========================================================================
# LÓGICA DE GENERACIÓN INICIAL (DOS FASES: ESTRUCTURA → LÓGICA)
# =========================================================================
if submit_button:
    if not api_key_usuario:
        st.error("🔒 Debes ingresar tu API Key en la página de Inicio (barra lateral).")
    elif not modulo or not tema or not descripcion:
        st.warning("📝 Completa la asignatura, el tema y la descripción exacta.")
    elif len(descripcion.strip()) < 25:
        st.warning("📝 La descripción es muy corta. Detalla mejor los cálculos, inputs y resultados esperados para obtener un simulador de mayor calidad.")
    else:
        try:
            # --- FASE 1: ESTRUCTURA (HTML + CSS) ---
            with st.spinner(f'🧱 Paso 1/2: diseñando la estructura visual con {modelo_seleccionado}...'):
                prompt_estructura = construir_prompt_estructura(
                    politecnico, docente, modulo, tema, descripcion, tema_visual, nivel_complejidad
                )
                resp_estructura, corte_estructura, reintento_1 = solicitar_con_reintento(
                    proveedor_ia, api_key_usuario, modelo_seleccionado, prompt_estructura, max_tokens_sim
                )
                html_estructura, truncado_estructura = extraer_html_ui_pro(resp_estructura)

            # --- FASE 2: LÓGICA (JavaScript) ---
            with st.spinner('⚙️ Paso 2/2: programando la lógica JavaScript...'):
                prompt_logica = construir_prompt_logica(html_estructura, descripcion, nivel_complejidad)
                resp_logica, corte_logica, reintento_2 = solicitar_con_reintento(
                    proveedor_ia, api_key_usuario, modelo_seleccionado, prompt_logica, max_tokens_sim
                )
                script_bloque, truncado_logica = extraer_script(resp_logica)

            codigo_html = insertar_script(html_estructura, script_bloque)

            if modo_debug:
                with st.expander("🐛 Fase 1 — respuesta cruda (estructura)"):
                    st.text_area("Estructura:", resp_estructura or "(vacío)", height=200, key="debug_estructura")
                with st.expander("🐛 Fase 2 — respuesta cruda (lógica)"):
                    st.text_area("Lógica:", resp_logica or "(vacío)", height=200, key="debug_logica")

            st.session_state.codigo_actual = codigo_html
            st.session_state.tema_actual = tema
            st.session_state.historial.append({
                "tema": tema,
                "codigo": codigo_html,
                "hora": datetime.datetime.now().strftime("%H:%M:%S")
            })

            hubo_corte = corte_estructura or truncado_estructura or corte_logica or truncado_logica
            if hubo_corte:
                fase = "estructura" if (corte_estructura or truncado_estructura) else "lógica"
                st.warning(f"⚠️ La fase de {fase} se cortó por límite de tokens incluso tras el reintento automático. Sube el 'Límite de tokens por llamada' en Opciones avanzadas o simplifica la descripción, y vuelve a generar.")
            else:
                nota_reintento = " (se reintentó automáticamente con más tokens y salió bien)" if (reintento_1 or reintento_2) else ""
                st.success(f"✅ ¡Simulador UI Pro compilado en 2 fases! Estructura y lógica ensambladas correctamente{nota_reintento}.")

        except ResourceExhausted:
            st.error("❌ Cuota de API de Gemini alcanzada. Intenta nuevamente en unos minutos.")
        except Exception as e:
            st.error(f"⚠️ Error de procesamiento: {e}")

# =========================================================================
# VISTA PREVIA + REFINAMIENTO + CÓDIGO (si ya hay un simulador generado)
# =========================================================================
if st.session_state.codigo_actual:
    codigo_html = st.session_state.codigo_actual
    tema_actual = st.session_state.tema_actual or "simulador"

    tab_preview, tab_refinar, tab_codigo = st.tabs(["🖥️ Vista Previa", "🔄 Solicitar Ajustes", "🔍 Código Fuente"])

    with tab_preview:
        with st.container(border=True):
            components.html(codigo_html, height=altura_preview, scrolling=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            st.download_button(
                label="📥 Descargar Simulador (.html)",
                data=codigo_html.encode('utf-8'),
                file_name=f"Simulador_{sanear_nombre_archivo(tema_actual)}.html",
                mime="text/html",
                type="primary"
            )
        with col_btn2:
            st.markdown(
                f'<a href="{generar_data_uri(codigo_html)}" target="_blank" '
                f'style="display:block;text-align:center;background:#0F172A;color:#fff;'
                f'padding:12px 24px;border-radius:6px;text-decoration:none;font-weight:600;">'
                f'🔗 Abrir en pestaña nueva</a>',
                unsafe_allow_html=True
            )

    with tab_refinar:
        st.caption("Pide cambios puntuales (colores, textos, validaciones, campos adicionales) sin regenerar todo desde cero.")
        with st.form("form_refinamiento", clear_on_submit=True):
            ajuste = st.text_area(
                "¿Qué quieres ajustar?",
                placeholder="Ej: Cambia el color de acento a verde, agrega un botón de 'Reiniciar' y valida que la IP tenga formato correcto.",
                height=100
            )
            submit_ajuste = st.form_submit_button("✨ Aplicar Ajustes")

        if submit_ajuste:
            if not api_key_usuario:
                st.error("🔒 Debes ingresar tu API Key en la página de Inicio (barra lateral).")
            elif not ajuste or len(ajuste.strip()) < 5:
                st.warning("📝 Describe con un poco más de detalle el ajuste que necesitas.")
            else:
                with st.spinner("🔧 Aplicando ajustes al simulador..."):
                    try:
                        prompt_ajuste = construir_prompt_refinamiento(codigo_html, ajuste)
                        respuesta_ajuste, corte_tokens, reintentado = solicitar_con_reintento(
                            proveedor_ia, api_key_usuario, modelo_seleccionado, prompt_ajuste, max_tokens_sim
                        )
                        codigo_actualizado, truncado = extraer_html_ui_pro(respuesta_ajuste)

                        if modo_debug:
                            with st.expander("🐛 Respuesta cruda del ajuste"):
                                st.text_area("Ajuste:", respuesta_ajuste or "(vacío)", height=200, key="debug_ajuste")

                        st.session_state.codigo_actual = codigo_actualizado
                        st.session_state.historial.append({
                            "tema": f"{tema_actual} (ajustado)",
                            "codigo": codigo_actualizado,
                            "hora": datetime.datetime.now().strftime("%H:%M:%S")
                        })
                        if corte_tokens or truncado:
                            st.warning("⚠️ El ajuste se cortó por el límite de tokens incluso tras el reintento automático. Sube el límite en Opciones avanzadas y vuelve a intentar.")
                        else:
                            st.success("✅ Ajustes aplicados. Revisa la pestaña 'Vista Previa'.")
                        st.rerun()

                    except ResourceExhausted:
                        st.error("❌ Cuota de API de Gemini alcanzada. Intenta nuevamente en unos minutos.")
                    except Exception as e:
                        st.error(f"⚠️ Error de procesamiento: {e}")

    with tab_codigo:
        st.code(codigo_html, language='html')