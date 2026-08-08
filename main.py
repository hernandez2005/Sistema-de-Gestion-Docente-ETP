import streamlit as st
import os
import base64 as _b64

# --- 1. CONFIGURACIÓN GLOBAL ---
st.set_page_config(
    page_title="Sistema de Gestión Docente - ETP",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. INICIALIZACIÓN DE SESIÓN ---
_SESSION_DEFAULTS = {
    "api_key_global": "",
    "proveedor_ia_global": "Google Gemini",
    "modelo_global": "gemini-2.5-flash",
    "modelo_custom_text": "",
    "usar_modelo_custom": False,
}

for key, default in _SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default

# --- 3. IMAGEN EN EL TOPE DEL SIDEBAR (CSS ::before) ---
_img_path = "Gemini_Generated_Image_5ck0tc5ck0tc5ck0.png"
if os.path.exists(_img_path):
    with open(_img_path, "rb") as _f:
        _img_b64 = _b64.b64encode(_f.read()).decode()
    st.markdown(f"""
    <style>
        [data-testid="stSidebarContent"]::before {{
            content: '';
            display: block;
            height: 85px;
            background-image: url("data:image/png;base64,{_img_b64}");
            background-size: contain;
            background-repeat: no-repeat;
            background-position: center;
            margin: 8px 12px 4px 12px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }}
    </style>
    """, unsafe_allow_html=True)

# --- 4. CSS DINÁMICO (color API Key) ---
_color_input = "#DBEAFE" if not st.session_state.api_key_global else "#D1FAE5"
_borde_input = "#3B82F6" if not st.session_state.api_key_global else "#10B981"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    .block-container {{ padding-top: 1.5rem; padding-bottom: 0.5rem; max-width: 1200px; }}

    .hero-box {{
        background: linear-gradient(135deg, #0F172A 0%, #1E3A5F 40%, #0F172A 100%);
        border-radius: 18px; padding: 2.8rem 2rem 2rem 2rem; margin-bottom: 1.5rem;
        box-shadow: 0 25px 50px rgba(0,0,0,0.18); text-align: center;
        position: relative; overflow: hidden;
    }}
    .hero-box::before {{
        content: ''; position: absolute; top: -50%; left: -50%; width: 200%; height: 200%;
        background: radial-gradient(circle, rgba(59,130,246,0.08) 0%, transparent 60%);
        animation: heroPulse 8s ease-in-out infinite;
    }}
    @keyframes heroPulse {{ 0%,100%{{transform:scale(1)}} 50%{{transform:scale(1.05)}} }}
    .hero-title {{ font-size: 2.5rem; font-weight: 800; color: #FFF; letter-spacing: -0.03em; line-height: 1.15; position: relative; }}
    .hero-sub {{ font-size: 1.1rem; color: #94A3B8; font-weight: 400; margin-top: 0.5rem; margin-bottom: 1.2rem; position: relative; }}
    .hero-badge {{
        display: inline-block; background: rgba(59,130,246,0.15); color: #60A5FA;
        padding: 5px 14px; border-radius: 20px; font-size: 0.82rem; font-weight: 600;
        border: 1px solid rgba(59,130,246,0.25); position: relative;
    }}

    .stat-card {{
        background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px;
        padding: 1rem 1.2rem; box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        text-align: center; min-height: 90px; display: flex; flex-direction: column;
        justify-content: center; align-items: center;
    }}
    .stat-label {{ font-size: 0.78rem; color: #64748B; font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 4px; }}
    .stat-value {{ font-size: 1.05rem; color: #0F172A; font-weight: 700; }}
    .stat-ok {{ border-left: 4px solid #10B981; }}
    .stat-warn {{ border-left: 4px solid #F59E0B; }}
    .stat-err {{ border-left: 4px solid #EF4444; }}

    .tool-card {{
        background: #FFFFFF; border: 2px solid #E2E8F0; border-radius: 12px;
        padding: 1.1rem 1.2rem 0.6rem 1.2rem; transition: all 0.25s ease;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    }}
    .tool-card:hover {{
        border-color: #3B82F6; box-shadow: 0 8px 24px rgba(59,130,246,0.13);
        transform: translateY(-3px);
    }}
    .tool-icon {{ font-size: 1.8rem; margin-bottom: 6px; }}
    .tool-name {{ font-size: 1rem; font-weight: 700; color: #0F172A; margin-bottom: 3px; }}
    .tool-tag {{ display: inline-block; font-size: 0.7rem; font-weight: 600; padding: 2px 8px; border-radius: 10px; margin-bottom: 6px; }}
    .tag-etp {{ background: #DBEAFE; color: #1D4ED8; }}
    .tag-acad {{ background: #FEF3C7; color: #92400E; }}
    .tool-desc {{ font-size: 0.82rem; color: #64748B; line-height: 1.45; margin-bottom: 8px; }}

    .info-box {{
        background: linear-gradient(90deg, #F0F9FF 0%, #E0F2FE 100%);
        border-left: 4px solid #0EA5E9; padding: 14px 20px; border-radius: 8px;
        color: #0C4A6E; font-size: 0.9rem; line-height: 1.55; margin-bottom: 1rem;
    }}

    .footer-main {{ text-align: center; margin-top: 2.5rem; padding-top: 1.5rem; border-top: 1px solid #E2E8F0; }}
    .social-btn {{
        text-decoration: none !important; color: #FFF !important; padding: 8px 20px;
        border-radius: 8px; margin: 0 6px; font-weight: 600; font-size: 0.88rem;
        display: inline-block; transition: opacity 0.2s; box-shadow: 0 2px 6px rgba(0,0,0,0.12);
    }}
    .social-btn:hover {{ opacity: 0.85; }}
    .btn-ig {{ background: linear-gradient(135deg, #833AB4, #FD1D1D, #FCB045); }}
    .btn-tk {{ background: #000; }}

    .sidebar-header {{ font-size: 1.1rem; font-weight: 700; color: #0F172A; margin-bottom: 4px; }}
    .sidebar-caption {{ font-size: 0.78rem; color: #94A3B8; margin-bottom: 12px; }}

    .key-box {{
        background: {_color_input}; border: 2px solid {_borde_input}; border-radius: 10px;
        padding: 12px 14px; margin-bottom: 8px; text-align: center;
        font-weight: 600; font-size: 0.9rem; transition: all 0.3s ease;
    }}
    .key-empty {{ color: #1E40AF; }}
    .key-ok {{ color: #065F46; }}
</style>
""", unsafe_allow_html=True)

# --- 5. BARRA LATERAL ---
with st.sidebar:

    st.markdown('<p class="sidebar-header">⚙️ Configuración de IA</p>', unsafe_allow_html=True)
    st.markdown('<p class="sidebar-caption">Configura una vez · Se comparte en todas las herramientas</p>', unsafe_allow_html=True)

    # Proveedor
    proveedor = st.selectbox(
        "🧠 Proveedor de IA",
        ["Google Gemini", "OpenAI (ChatGPT)"],
        index=["Google Gemini", "OpenAI (ChatGPT)"].index(st.session_state.proveedor_ia_global),
        key="prov_global_nav"
    )
    st.session_state.proveedor_ia_global = proveedor

    # Modelos
    if proveedor == "Google Gemini":
        modelos_preset = ["gemini-3.5-flash","gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
    else:
        modelos_preset = ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-3.5-turbo"]

    usar_custom = st.checkbox(
        "✏️ Modelo personalizado",
        value=st.session_state.usar_modelo_custom,
        key="usar_custom_nav",
        help="Escribe el nombre exacto de cualquier modelo nuevo."
    )
    st.session_state.usar_modelo_custom = usar_custom

    if usar_custom:
        modelo = st.text_input(
            "Nombre exacto del modelo:",
            value=st.session_state.modelo_custom_text,
            placeholder="Ej: gemini-3.0-flash · gpt-5",
            key="modelo_custom_nav"
        )
        st.session_state.modelo_custom_text = modelo
        st.session_state.modelo_global = modelo if modelo.strip() else st.session_state.modelo_global
    else:
        idx = modelos_preset.index(st.session_state.modelo_global) if st.session_state.modelo_global in modelos_preset else 0
        modelo = st.selectbox("🤖 Modelo", modelos_preset, index=idx, key="modelo_global_nav")
        st.session_state.modelo_global = modelo

    st.markdown('<hr style="border:0;border-top:1px solid #E2E8F0;margin:1rem 0;">', unsafe_allow_html=True)

    # API Key con indicador dinámico
    if not st.session_state.api_key_global:
        st.markdown('<div class="key-box key-empty">🔑 Ingresa tu API Key abajo</div>', unsafe_allow_html=True)
    else:
        if st.session_state.api_key_global.startswith("AIza"):
            st.markdown('<div class="key-box key-ok">✅ Google Gemini — Clave activa</div>', unsafe_allow_html=True)
        elif st.session_state.api_key_global.startswith("sk-"):
            st.markdown('<div class="key-box key-ok">✅ OpenAI — Clave activa</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="key-box key-empty">⚠️ Formato reconocido</div>', unsafe_allow_html=True)

    api_key = st.text_input(
        "API Key:",
        type="password",
        value=st.session_state.api_key_global,
        key="api_key_global_nav",
        help="Se comparte automáticamente con TODAS las herramientas."
    )
    st.session_state.api_key_global = api_key

    with st.expander("🔗 ¿Cómo obtengo una API Key?"):
        st.markdown("""
        **Google Gemini (Gratis):**
        1. Ve a [Google AI Studio](https://aistudio.google.com/)
        2. Inicia sesión → **"Get API Key"**

        **OpenAI (De pago):**
        1. Ve a [platform.openai.com](https://platform.openai.com/api-keys)
        2. Genera una nueva clave
        """)

    st.markdown('<hr style="border:0;border-top:1px solid #E2E8F0;margin:1rem 0;">', unsafe_allow_html=True)

    # Sobre el Desarrollador
    st.markdown("#### 👨‍💻 Ing. Bernardo Hernández")
    st.markdown("""
    *Innovación y Eficiencia en la Gestión Técnico-Profesional*

    Desarrollado desde la experiencia real en la coordinación académica y la docencia tecnológica, este Programa de Gestión Educativa nace para resolver los desafíos cotidianos de la administración escolar moderna. No se trata solo de un software de registro, sino de un ecosistema digital diseñado específicamente para optimizar el seguimiento de módulos formativos y agilizar los procesos administrativos en instituciones de nivel medio y superior.

    El sistema destaca por entender exactamente cómo funciona la educación técnica hoy en día, fusionando una arquitectura de software robusta con una interfaz intuitiva pensada para el docente.
    """)
    st.markdown("[📸 Instagram](https://instagram.com/El_Profe_Hernandez) · [🎵 TikTok](https://tiktok.com/@El_Profe_Hernandez)", unsafe_allow_html=True)


# --- 6. PÁGINA DE INICIO ---
def pagina_inicio():

    # Hero
    st.markdown("""
    <div class="hero-box">
        <div class="hero-title">Sistema de Gestión Docente ETP</div>
        <div class="hero-sub">Automatización pedagógica con Inteligencia Artificial — Alineado al MINERD</div>
        <div class="hero-badge">🚀 Potenciado por IA &nbsp;·&nbsp; 🏫 Modalidad Técnico Profesional</div>
    </div>
    """, unsafe_allow_html=True)

    # Estado del Sistema
    st.markdown("#### 📊 Estado del Sistema")
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        if st.session_state.api_key_global:
            st.markdown('<div class="stat-card stat-ok"><div class="stat-label">API Key</div><div class="stat-value">✅ Configurada</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="stat-card stat-err"><div class="stat-label">API Key</div><div class="stat-value">🔒 Pendiente</div></div>', unsafe_allow_html=True)
    with col_s2:
        st.markdown(f'<div class="stat-card"><div class="stat-label">Proveedor</div><div class="stat-value">{st.session_state.proveedor_ia_global}</div></div>', unsafe_allow_html=True)
    with col_s3:
        st.markdown(f'<div class="stat-card"><div class="stat-label">Modelo</div><div class="stat-value" style="font-size:0.9rem">{st.session_state.modelo_global}</div></div>', unsafe_allow_html=True)

    if not st.session_state.api_key_global:
        st.markdown("""
        <div class="info-box">
            <b>⚠️ Acción requerida:</b> Ingresa tu <b>API Key</b> en la barra lateral (sección <i>⚙️ Configuración de IA</i>).
        </div>
        """, unsafe_allow_html=True)

    # Herramientas ETP
    st.markdown("#### 🔧 ETP — Talleres y Módulos Formativos")

    herramientas_etp = [
        {"icon": "📊", "name": "Ponderación RA", "page": "ponderacionra.py", "desc": "Distribución porcentual y temporal de Resultados de Aprendizaje a partir del PDF curricular."},
        {"icon": "🚀", "name": "Planificación Modular", "page": "planifiacionra.py", "desc": "Matriz de planificación por RA con Elementos de Capacidad, actividades e instrumentos."},
        {"icon": "📅", "name": "Plan Diario ETP", "page": "pladiario.py", "desc": "Planificación de clase diaria en 50 minutos: momentos pedagógicos y lista de cotejo."},
        {"icon": "📚", "name": "Generador de Contenidos", "page": "contenido.py", "desc": "Contenido anclado, progresión Bloom, actividades diferenciadas, rúbrica multinivel."},
        {"icon": "💻", "name": "Fábrica de Simuladores", "page": "simuladores.py", "desc": "Simuladores web interactivos personalizados con diseño UI/UX profesional."},
        {"icon": "📝", "name": "Banco de Ítems", "page": "bancoitems.py", "desc": "Pruebas diversificadas: opción múltiple, C/I, completar, apareamiento, clasificación, casos."},
        {"icon": "🚨", "name": "Alerta Temprana", "page": "alerta.py", "desc": "Diagnóstico de severidad, plan de recuperación multinivel, seguimiento semanal."},
        {"icon": "⭐", "name": "Feedback del Portal", "page": "feedback.py", "desc": "Valoración y comentarios sobre la experiencia de uso del portal."}
    ]

    cols_etp = st.columns(3)
    for i, t in enumerate(herramientas_etp):
        with cols_etp[i % 3]:
            st.markdown(f"""
            <div class="tool-card">
                <div class="tool-icon">{t['icon']}</div>
                <div class="tool-name">{t['name']}</div>
                <div class="tool-tag tag-etp">ETP</div>
                <div class="tool-desc">{t['desc']}</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("▶ Abrir", key=f"btn_{t['page']}", use_container_width=True):
                st.switch_page(t['page'])

    # Herramientas Académicas
    st.markdown("#### 📚 Áreas Académicas")

    herramientas_acad = [
        {"icon": "📖", "name": "Plan de Unidad (Académicas)", "page": "academicas.py", "desc": "Planificación de unidades de aprendizaje para áreas académicas generales."},
        {"icon": "🗓️", "name": "Plan Diario (Académicas)", "page": "diario_academico.py", "desc": "Planificación diaria con esquema oficial MINERD para áreas académicas."},
    ]

    cols_acad = st.columns(2)
    for i, t in enumerate(herramientas_acad):
        with cols_acad[i % 2]:
            st.markdown(f"""
            <div class="tool-card">
                <div class="tool-icon">{t['icon']}</div>
                <div class="tool-name">{t['name']}</div>
                <div class="tool-tag tag-acad">ACAD</div>
                <div class="tool-desc">{t['desc']}</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("▶ Abrir", key=f"btn_{t['page']}", use_container_width=True):
                st.switch_page(t['page'])

    # Nota Pedagógica
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div class="info-box">
        <b>📌 Nota de Adaptación Pedagógica:</b> Si bien el motor de IA proporciona una base
        estructural robusta y altamente estandarizada, estas generaciones poseen un carácter global.
        Es <b>fundamental</b> que el maestro revise, contextualice y adapte cada propuesta a la realidad de su aula.
    </div>
    """, unsafe_allow_html=True)

    # Manual
    with st.expander("📖 Manual de Ayuda y Uso Rápido", expanded=False):
        st.markdown("""
        ### 🚀 Inicio Rápido

        **Paso 1 — Configura tu API Key (solo una vez):**
        - En la barra lateral, selecciona proveedor y modelo.
        - Si Google/OpenAI lanzan un **nuevo modelo**, activa **"✏️ Modelo personalizado"**.
        - Pega tu **API Key**. Se compartirá con TODAS las herramientas automáticamente.

        **Paso 2 — Selecciona una herramienta:**
        - Usa el menú de navegación en la barra lateral, o haz clic en **▶ Abrir** en las tarjetas.

        ---

        ### 🔧 Herramientas ETP

        | Herramienta | Función |
        |---|---|
        | **Ponderación RA** | PDF → distribución de % y semanas por cada R.A. |
        | **Planificación Modular** | PDF + UC + RA → matriz con EC, actividades e instrumentos |
        | **Plan Diario ETP** | Perfil del grupo → momentos pedagógicos (50 min) + lista de cotejo |
        | **Generador de Contenidos** | Tema → contenido anclado (PDF), Bloom, actividades diferenciadas, rúbrica L/EP/NA |
        | **Fábrica de Simuladores** | Descripción → simulador web HTML interactivo completo |
        | **Banco de Ítems** | PDF curricular → 9 tipos de ítems diversificados + solucionario |
        | **Alerta Temprana** | Estudiantes con brechas → diagnóstico de severidad + recuperación multinivel |

        ### 📚 Herramientas Académicas

        | Herramienta | Función |
        |---|---|
        | **Plan de Unidad** | Unidades para áreas académicas (Sociales, Español, Matemática...) |
        | **Plan Diario Académico** | Esquema MINERD con Inicio/Desarrollo/Cierre + Indagación + Metacognición |
        """)

    # Footer
    st.markdown("""
    <div class="footer-main">
        <div style="font-size:1rem;font-weight:700;color:#0F172A;">Ing. Bernardo Hernández</div>
        <div style="font-size:0.85rem;color:#64748B;margin-bottom:10px;">Coordinador de Módulos Formativos — ETP</div>
        <a href="https://instagram.com/El_Profe_Hernandez" target="_blank" class="social-btn btn-ig">📷 Instagram</a>
        <a href="https://tiktok.com/@El_Profe_Hernandez" target="_blank" class="social-btn btn-tk">🎵 TikTok</a>
    </div>
    """, unsafe_allow_html=True)


# --- 7. NAVEGACIÓN ---
inicio = st.Page(pagina_inicio, title="Inicio", icon="🏠", default=True)

pagina_ponderacion = st.Page("ponderacionra.py",  title="Ponderación RA",              icon="📊")
pagina_planificacion = st.Page("planifiacionra.py", title="Planificación Modular",       icon="🚀")
pagina_pladiario = st.Page("pladiario.py",        title="Plan Diario ETP",             icon="📅")
pagina_contenido = st.Page("contenido.py",        title="Generador de Contenidos",     icon="📚")
pagina_simuladores = st.Page("simuladores.py",    title="Fábrica de Simuladores",      icon="💻")
pagina_banco = st.Page("bancoitems.py",           title="Banco de Ítems",             icon="📝")
pagina_alerta = st.Page("alerta.py",              title="Recuperación R.A o Pedagógica",            icon="🚨")
pagina_feedback = st.Page("feedback.py",          title="Feedback del Portal",            icon="⭐")

pagina_academicas = st.Page("academicas.py",       title="Plan de Unidad (Académicas)", icon="📖")
pagina_diario_acad = st.Page("diario_academico.py", title="Plan Diario (Académicas)",  icon="🗓️")

menu = st.navigation({
    "🏠 Principal": [inicio],
    "🔧 ETP — Talleres y Módulos": [
        pagina_ponderacion,
        pagina_planificacion,
        pagina_pladiario,
        pagina_contenido,
        pagina_simuladores,
        pagina_banco,
        pagina_alerta,
        pagina_feedback
    ],
    "📚 Áreas Académicas": [
        pagina_academicas,
        pagina_diario_acad,
    ],
})

menu.run()