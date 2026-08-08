import streamlit as st
import json
import os
from datetime import datetime

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

# --- ENCABEZADO ---
st.markdown('<div class="main-header">Valoración y Comentarios del Portal</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Tu opinión nos ayuda a optimizar las herramientas pedagógicas para la ETP</div>', unsafe_allow_html=True)

# --- FORMULARIO DE VALORACIÓN ---
with st.form("form_feedback", clear_on_submit=True):
    st.markdown('<div class="section-title">⭐ Déjanos tu Experiencia</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        docente_nombre = st.text_input("Nombre del Docente (Opcional)", placeholder="Ej: Ing. Bernardo Hernández")
    with col2:
        estrellas = st.slider("Valoración general del portal", min_value=1, max_value=5, value=5, format="%d ⭐")
        
    comentario = st.text_area(
        "Comentarios, sugerencias o reportes de mejora:", 
        height=120,
        placeholder="Ej: La herramienta de ponderación ahorra muchísimo tiempo. Sería excelente añadir..."
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    submit_feedback = st.form_submit_button("📤 Enviar Valoración")

if submit_feedback:
    if not comentario.strip():
        st.warning("⚠️ Por favor, escribe un breve comentario antes de enviar.")
    else:
        # Estructura del registro
        nuevo_registro = {
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "docente": docente_nombre.strip() if docente_nombre.strip() else "Anónimo",
            "valoracion": estrellas,
            "comentario": comentario.strip()
        }
        
        # Guardar en un archivo JSON local persistente
        archivo_json = "feedback_portal.json"
        registros_existentes = []
        
        if os.path.exists(archivo_json):
            try:
                with open(archivo_json, "r", encoding="utf-8") as f:
                    registros_existentes = json.load(f)
            except Exception:
                registros_existentes = []
                
        registros_existentes.append(nuevo_registro)
        
        try:
            with open(archivo_json, "w", encoding="utf-8") as f:
                json.dump(registros_existentes, f, ensure_ascii=False, indent=4)
            st.success("🎉 ¡Muchas gracias! Tu comentario ha sido registrado correctamente en el sistema.")
        except Exception as e:
            st.error(f"⚠️ Error al guardar la valoración: {e}")

# --- PANEL DE VISUALIZACIÓN PARA EL ADMINISTRADOR ---
with st.expander("🔒 Panel de Administración (Ver valoraciones recibidas)"):
    archivo_json = "feedback_portal.json"
    if os.path.exists(archivo_json):
        try:
            with open(archivo_json, "r", encoding="utf-8") as f:
                datos = json.load(f)
                
            if datos:
                st.markdown(f"**Total de valoraciones recibidas:** {len(datos)}")
                promedio = sum(item['valoracion'] for item in datos) / len(datos)
                st.markdown(f"**Valoración promedio:** {promedio:.1f} ⭐")
                st.markdown("---")
                
                for idx, item in enumerate(reversed(datos)):
                    st.markdown(f"**{item['docente']}** — {item['valoracion']} ⭐ *({item['fecha']})*")
                    st.write(f'"{item["comentario"]}"')
                    if idx < len(datos) - 1:
                        st.divider()
            else:
                st.info("Aún no hay valoraciones registradas.")
        except Exception:
            st.info("No se pudieron leer los registros de valoración.")
    else:
        st.info("Aún no hay archivo de valoraciones creado.")