import streamlit as st
import os

# --- 1. CONFIGURACIÓN GLOBAL DE LA APLICACIÓN ---
st.set_page_config(
    page_title="Sistema de Gestión Docente - ETP",
    page_icon="🏫",
    layout="wide"
)

# --- ESTILOS CSS PARA HACER LA INTERFAZ COMPACTA Y SIN SCROLL ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    .block-container {
        padding-top: 2rem;
        padding-bottom: 1rem;
    }
    
    .intro-text {
        font-size: 1.05rem;
        color: #475569;
        text-align: center;
        max-width: 900px;
        margin: 10px auto 15px auto;
        line-height: 1.5;
    }
    
    .warning-box {
        background-color: #F8FAFC;
        border-left: 4px solid #3B82F6;
        padding: 12px 20px;
        border-radius: 6px;
        max-width: 900px;
        margin: 0 auto 15px auto;
        color: #334155;
        font-size: 0.95rem;
        line-height: 1.5;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: justify;
    }
    
    .creator-text {
        text-align: center;
        font-size: 1.1rem;
        font-weight: 600;
        color: #1E293B;
        margin-top: 10px;
        margin-bottom: 10px;
    }
    
    .creator-subtext {
        text-align: center;
        font-size: 0.9rem;
        color: #64748B;
        margin-bottom: 15px;
    }
    
    .social-container {
        text-align: center;
        margin-bottom: 10px;
    }
    
    .btn-social {
        text-decoration: none !important;
        color: white !important;
        padding: 8px 24px;
        border-radius: 6px;
        margin: 0 8px;
        font-weight: 500;
        font-size: 0.95rem;
        transition: opacity 0.2s ease;
        display: inline-block;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .btn-ig { background-color: #E1306C; } 
    .btn-tk { background-color: #000000; } 
    .btn-social:hover { opacity: 0.8; }
</style>
""", unsafe_allow_html=True)

# --- 2. FUNCIÓN DE LA PÁGINA DE INICIO (PORTADA) ---
def pagina_inicio():
    
    # Sistema de seguridad para la imagen (Evita que la app colapse en la nube si falta el archivo)
    col_vacia1, col_img, col_vacia2 = st.columns([1, 1.2, 1])
    with col_img:
        nombre_imagen = "Gemini_Generated_Image_5ck0tc5ck0tc5ck0.png"
        if os.path.exists(nombre_imagen):
            st.image(nombre_imagen, use_container_width=True)
    
    # 2.2 Introducción de la aplicación
    st.markdown("""
    <div class="intro-text">
        <b>Bienvenido al entorno de automatización pedagógica.</b><br>
        Esta plataforma está diseñada exclusivamente para optimizar el trabajo de los docentes de la Educación Técnico Profesional (ETP). 
        Utilizando Inteligencia Artificial, podrás procesar Resultados de Aprendizaje (R.A), generar matrices modulares 
        y estructurar planes diarios alineados con los estándares del MINERD, ahorrando tiempo valioso.
    </div>
    """, unsafe_allow_html=True)
    
    # 2.3 MANUAL DE AYUDA INTEGRADO
    with st.expander("📖 Manual de Ayuda y Uso de la Plataforma", expanded=False):
        st.markdown("""
        ### ¿Cómo utilizar este sistema?
        
        **1. Autenticación (API Key)**
        * Ve a la barra lateral en cualquiera de las herramientas e ingresa tu clave de Google AI Studio (empieza con `AIza` o `AQ`) https://aistudio.google.com/.
        
        **2. Ponderación RA (Sábana de Porcentajes)**
        * Sube el PDF del diseño curricular (solo páginas del módulo). El sistema evaluará los contenidos y distribuirá las semanas y los porcentajes automáticamente.
        
        **3. Planificación Modular (Matriz)**
        * Sube tu PDF, pega la Unidad de Competencia y el R.A. La IA generará los Elementos de Capacidad y desglosará las actividades y criterios.
        
        **4. Plan Diario (50 min)**
        * Completa el perfil de tu grupo. La herramienta diseñará los momentos pedagógicos (Inicio, Desarrollo, Cierre) y generará una lista de cotejo.
        
        **5. Generador de Contenidos**
        * Introduce el tema a impartir. Obtendrás el desarrollo teórico para los estudiantes, el paso a paso de la actividad y la rúbrica de evaluación lista para imprimir.
        """)
    
    # 2.4 Nota de Recordatorio Pedagógico
    st.markdown("""
    <div class="warning-box">
        <b>📌 Nota de Adaptación Pedagógica:</b> Si bien el motor de Inteligencia Artificial proporciona una base estructural robusta y altamente estandarizada, 
        estas generaciones algorítmicas poseen un carácter global. Es fundamental que el maestro revise, contextualice y adapte 
        cada propuesta a la realidad viva de su aula, tomando en consideración las necesidades de sus estudiantes, los recursos disponibles 
        y el contexto particular de su centro educativo.
    </div>
    """, unsafe_allow_html=True)
    
    # 2.5 Créditos institucionales y Redes Sociales
    st.markdown('<div class="creator-text">Desarrollado y Arquitectado por: Ing. Bernardo Hernández</div>', unsafe_allow_html=True)
    st.markdown('<div class="creator-subtext">Coordinador de Módulos Formativos ETP</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="social-container">
        <a href="https://instagram.com/El_Profe_Hernandez" target="_blank" class="btn-social btn-ig">📷 Instagram</a>
        <a href="https://tiktok.com/@El_Profe_Hernandez" target="_blank" class="btn-social btn-tk">🎵 TikTok</a>
    </div>
    """, unsafe_allow_html=True)

# --- 3. DEFINICIÓN DE PÁGINAS Y MENÚ DE NAVEGACIÓN ---

inicio = st.Page(pagina_inicio, title="Inicio", icon="🏠", default=True)

# SE CORRIGIÓ EL ERROR TIPOGRÁFICO DE PLANIFICACIÓN (planificacionra.py)
pagina_ponderacion = st.Page("ponderacionra.py", title="Ponderación RA", icon="📊")
pagina_planificacion = st.Page("planifiacionra.py", title="Planificación Modular", icon="🚀")
pagina_pladiario = st.Page("pladiario.py", title="Plan Diario", icon="📅")
pagina_contenido = st.Page("contenido.py", title="Generador de Contenidos", icon="📚")
pagina_simuladores = st.Page("simuladores.py", title="Fábrica de Simuladores", icon="💻")
pagina_banco = st.Page("bancoitems.py", title="Banco de Ítems y Pruebas", icon="📝")
pagina_alerta = st.Page("alerta.py", title="Alerta Temprana y Reforzamiento", icon="🚨")

# Materias Académicas
pagina_academicas = st.Page("academicas.py", title="Plan de Unidad (Académicas)", icon="📖")
pagina_diario_acad = st.Page("diario_academico.py", title="Plan Diario (Académicas)", icon="🗓️")

menu = st.navigation({
    "Principal": [inicio],
    "ETP - Talleres y Módulos": [
        pagina_ponderacion, 
        pagina_planificacion, 
        pagina_pladiario,
        pagina_contenido,
        pagina_simuladores,
        pagina_banco,
        pagina_alerta
    ],
    "Áreas Académicas": [
        pagina_academicas,
        pagina_diario_acad
    ]
})

menu.run()

# --- 4. SECCIÓN MOTIVACIONAL Y REDES (BARRA LATERAL) ---
with st.sidebar:
    st.markdown("---")
    st.markdown("### 👨‍💻 Sobre el Desarrollador")
    st.markdown(
        "**¡Conectemos y sigamos innovando!** 💡\n\n"
        "Si esta plataforma te ayuda a optimizar tu tiempo de coordinación y docencia, te invito a apoyarme "
        "siguiéndome en mis redes. Allí comparto más recursos para docentes, programación y tecnología educativa."
    )
    st.markdown("[📸 **Instagram: @El_Profe_Hernandez**](https://instagram.com/El_Profe_Hernandez)", unsafe_allow_html=True)
    st.markdown("[🎵 **TikTok: @El_Profe_Hernandez**](https://tiktok.com/@El_Profe_Hernandez)", unsafe_allow_html=True)