import io
import random
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from twilio.rest import Client

# 1. Configuración de la interfaz
st.set_page_config(
    page_title="Formateador APA 7 Automático",
    page_icon="🎓",
    layout="centered"
)

st.title("🎓 Generador & Formateador APA 7")
st.markdown(
    "Pega el texto de tu trabajo o tu lista de fuentes desordenadas. "
    "La herramienta organizará las citas y referencias en **APA 7ma edición** "
    "y te generará el documento **Word (.docx)** listo para entregar."
)

# 2. Claves desde los secretos de Streamlit
api_key = st.secrets.get("GEMINI_API_KEY", "")
twilio_sid = st.secrets.get("TWILIO_ACCOUNT_SID", "")
twilio_token = st.secrets.get("TWILIO_AUTH_TOKEN", "")

# 3. Formulario principal
tipo_trabajo = st.selectbox(
    "¿Qué deseas procesar hoy?",
    [
        "Referencias bibliográficas (Lista final en orden alfabético)",
        "Párrafo/Texto con citas parentéticas y narrativas",
        "Ensayo o resumen completo"
    ]
)

texto_usuario = st.text_area(
    "Pega tu texto, enlaces o fuentes aquí:",
    height=200,
    placeholder="Ejemplo:\n- Sampieri metodología 2014 McGrawHill\n- https://scielo.org/articulo-ejemplo\n- Vygotsky teoría sociocultural..."
)

# 4. Generación de Word con normas APA 7
def generar_word_apa(texto_procesado, titulo="Trabajo Académico en Formato APA"):
    doc = Document()

    # Márgenes: 2.54 cm (1 pulgada)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Tipografía: Times New Roman 12, interlineado doble
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.line_spacing = 2.0
    style.paragraph_format.space_after = Pt(0)

    # Título centrado
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_titulo = p_titulo.add_run(titulo)
    run_titulo.bold = True

    lineas = texto_procesado.split("\n")
    es_seccion_referencias = False

    for linea in lineas:
        linea_limpia = linea.strip()
        if not linea_limpia:
            continue

        if "referencias" in linea_limpia.lower() and len(linea_limpia) < 25:
            es_seccion_referencias = True
            p_ref = doc.add_paragraph()
            p_ref.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_ref = p_ref.add_run("Referencias")
            run_ref.bold = True
            continue

        p = doc.add_paragraph()
        if es_seccion_referencias:
            # Sangría francesa reglamentaria
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run(linea_limpia)
        else:
            p.paragraph_format.first_line_indent = Inches(0.5)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run(linea_limpia)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# 5. Procesar con Gemini (con respaldo automático anti-saturación)
if st.button("⚡ Procesar en Formato APA 7", type="primary"):
    if not api_key:
        st.error("⚠️ Falta configurar la GEMINI_API_KEY en los Secrets.")
    elif not texto_usuario.strip():
        st.warning("⚠️ Debes pegar texto o fuentes en el cuadro superior.")
    else:
        with st.spinner("Organizando fuentes y aplicando normas APA 7ma edición..."):
            instrucciones = (
                "Eres un experto metodólogo universitario y revisor de estilo en normas APA 7ma edición.\n"
                f"El usuario seleccionó: {tipo_trabajo}.\n\n"
                "Instrucciones estrictas:\n"
                "1. Transforma el contenido al formato oficial APA 7ma edición.\n"
                "2. Si son referencias, ordénalas alfabéticamente por el apellido del autor. "
                "Estructura: Apellido, Inicial. (Año). Título en cursiva. Editorial/Revista, DOI o URL.\n"
                "3. Si es texto, verifica y corrige las citas parentéticas (Apellido, Año).\n"
                "4. No agregues introducciones, saludos ni comentarios. Solo entrega el contenido final listo, "
                "encabezado por 'Referencias' si corresponde.\n\n"
                "Contenido a procesar:\n"
                f"{texto_usuario}"
            )

            # Lista de modelos oficiales en orden de velocidad y disponibilidad
            modelos_disponibles = [
                "gemini-3.5-flash-lite",
                "gemini-3.8-flash",
                "gemini-3.1-flash-lite"
            ]

            cliente = genai.Client(api_key=api_key)
            respuesta_obtenida = None
            ultimo_error = None

            for modelo in modelos_disponibles:
                try:
                    response = cliente.models.generate_content(
                        model=modelo,
                        contents=instrucciones
                    )
                    respuesta_obtenida = response.text
                    break # Si respondió con éxito, salimos del ciclo de inmediato
                except Exception as err:
                    ultimo_error = err
                    continue # Si ese modelo tiene alta demanda, salta al siguiente

            if respuesta_obtenida:
                st.session_state["resultado_apa"] = respuesta_obtenida
                st.success("¡Contenido formateado exitosamente!")
            else:
                st.error(f"No fue posible procesar en este momento: {str(ultimo_error)}")
# 6. Vista previa y entrega de código vía WhatsApp
if "resultado_apa" in st.session_state:
    st.markdown("---")
    st.subheader("📄 Vista Previa del Resultado")
    st.text_area("Contenido generado:", value=st.session_state["resultado_apa"], height=180)

    st.markdown("---")
    st.subheader("📥 Descargar Documento Word (.docx)")
    st.info("Para recibir tu código de acceso para la descarga oficial en Word, ingresa tu número con código de país (ejemplo: +1809... o +1829...):")

    if "codigo_generado" not in st.session_state:
        st.session_state["codigo_generado"] = str(random.randint(100000, 999999))

    col_tel, col_btn = st.columns([2, 1])
    with col_tel:
        telefono_destino = st.text_input("Tu WhatsApp:", placeholder="+18090000000")
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        boton_enviar_ws = st.button("📲 Recibir Código")

    if boton_enviar_ws:
        if not twilio_sid or not twilio_token:
            st.error("⚠️ Faltan las credenciales de Twilio en los Secrets.")
        elif not telefono_destino.strip():
            st.warning("⚠️ Ingresa tu identificador o número de WhatsApp.")
        else:
            try:
                twilio_client = Client(twilio_sid, twilio_token)
                codigo = st.session_state["codigo_generado"]

                mensaje = twilio_client.messages.create(
                    from_="whatsapp:+14155238886",
                    to=f"whatsapp:{telefono_destino.strip()}",
                    body=f"🎓 Formateador APA: Tu código de descarga de 6 dígitos es: {codigo}"
                )
                st.success("✅ ¡Código enviado a tu WhatsApp! Revisa tu iPhone.")
            except Exception as err:
                st.error(f"Error de envío: {str(err)}")

    # Verificación del código
    codigo_input = st.text_input("Introduce el código de 6 dígitos que recibiste:", type="password")

    if codigo_input and codigo_input.strip() == st.session_state.get("codigo_generado"):
        st.success("¡Código verificado con éxito! Tu descarga está desbloqueada:")
        archivo_word = generar_word_apa(st.session_state["resultado_apa"])
        st.download_button(
            label="⬇️ Descargar archivo Word (.docx)",
            data=archivo_word,
            file_name="Trabajo_Formato_APA7.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    elif codigo_input:
        st.error("Código incorrecto. Verifica el mensaje en tu WhatsApp.")
