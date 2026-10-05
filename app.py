import io
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai

# 1. Configuración de la interfaz
st.set_page_config(
    page_title="Formateador APA 7 Automático",
    page_icon="🎓",
    layout="centered"
)

st.title("🎓 Generador & Formateador APA 7")
st.markdown(
    "Sube tu archivo o pega tu trabajo. La herramienta aplicará las normas "
    "**APA 7ma edición** conservando tu contenido y te generará el documento **Word (.docx)** listo para entregar."
)

# 2. Clave desde los secretos de Streamlit
api_key = st.secrets.get("GEMINI_API_KEY", "")

# 3. Formulario principal
tipo_trabajo = st.selectbox(
    "¿Qué deseas procesar hoy?",
    [
        "Ensayo o trabajo académico completo (Conserva todo el texto + Citas + Referencias)",
        "Referencias bibliográficas (Solo lista final en orden alfabético)",
        "Párrafo/Texto corto con citas parentéticas y narrativas"
    ]
)

# Opción para subir archivo .docx / .txt
archivo_subido = st.file_uploader("📂 Sube tu archivo (.docx o .txt) [Opcional]:", type=["docx", "txt"])

texto_extraido_archivo = ""
if archivo_subido is not None:
    if archivo_subido.name.endswith(".docx"):
        doc_in = Document(archivo_subido)
        texto_extraido_archivo = "\n".join([p.text for p in doc_in.paragraphs if p.text.strip()])
    else:
        texto_extraido_archivo = archivo_subido.read().decode("utf-8")

texto_usuario = st.text_area(
    "O pega tu texto / fuentes aquí:",
    value=texto_extraido_archivo if texto_extraido_archivo else "",
    height=220,
    placeholder="Pega el trabajo completo o la lista de fuentes..."
)

# 4. Generación de Word con normas APA 7
def generar_word_apa(texto_procesado, titulo="Trabajo Académico en Formato APA"):
    doc = Document()

    # Márgenes reglamentarios: 2.54 cm (1 pulgada)
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
            # Sangría francesa (0.5 pulgadas)
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run(linea_limpia)
        else:
            # Sangría de primera línea (0.5 pulgadas)
            p.paragraph_format.first_line_indent = Inches(0.5)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run(linea_limpia)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# 5. Procesamiento con Gemini 3.8 Flash
if st.button("⚡ Procesar en Formato APA 7", type="primary"):
    if not api_key:
        st.error("⚠️ Falta configurar la GEMINI_API_KEY en los Secrets.")
    elif not texto_usuario.strip():
        st.warning("⚠️ Debes subir un archivo o pegar texto en el cuadro superior.")
    else:
        with st.spinner("Procesando documento bajo normas APA 7ma edición..."):
            try:
                # Instrucción dinámica según el tipo de trabajo seleccionado
                if "Referencias bibliográficas" in tipo_trabajo:
                    instrucciones = (
                        "Eres un experto metodólogo universitario en normas APA 7ma edición.\n"
                        "Tu tarea es procesar ÚNICAMENTE una lista de referencias bibliográficas.\n"
                        "1. Organízalas alfabéticamente por apellido del autor.\n"
                        "2. Aplica la estructura: Apellido, Inicial. (Año). Título. Editorial/Revista, DOI o URL.\n"
                        "3. Encabeza con el título 'Referencias'. No agregues saludos ni explicaciones.\n\n"
                        f"Contenido:\n{texto_usuario}"
                    )
                else:
                    instrucciones = (
                        "Eres un experto metodólogo universitario y revisor de estilo en normas APA 7ma edición.\n"
                        "REGLA CRÍTICA Y OBLIGATORIA: NO RESUMAS NI OMITAS PÁRRAFOS. Conserva la TOTALIDAD del contenido original íntegro.\n\n"
                        "Instrucciones de formato:\n"
                        "1. Mantén todo el texto, desarrollo y argumentos del usuario.\n"
                        "2. Corrige las citas en el texto para que cumplan estrictamente APA 7 (Apellido, Año).\n"
                        "3. Estructura los títulos y subtítulos según los niveles APA.\n"
                        "4. Al final del trabajo, genera la sección 'Referencias' con las fuentes citadas ordenadas alfabéticamente.\n"
                        "5. No incluyas introducciones tuyas, comentarios ni notas al usuario. Devuelve el trabajo completo listo.\n\n"
                        f"Contenido a formatear:\n{texto_usuario}"
                    )

                cliente = genai.Client(api_key=api_key)
                response = cliente.models.generate_content(
                    modelos_disponibles = [
                "gemini-3.8-flash",
                "gemini-3.5-flash-lite",
                "gemini-3.1-pro",
                "gemini-3.1-flash-lite"
            ]

                    contents=instrucciones
                )

                if response.text:
                    st.session_state["resultado_apa"] = response.text
                    st.success("¡Documento formateado exitosamente!")
                    st.rerun()

            except Exception as err:
                st.error(f"Error al procesar: {str(err)}")

# 6. Vista previa y entrega de código vía WhatsApp directo
if "resultado_apa" in st.session_state:
    st.markdown("---")
    st.subheader("📄 Vista Previa del Resultado")
    st.text_area("Contenido generado:", value=st.session_state["resultado_apa"], height=250)

    st.markdown("---")
    st.subheader("📥 Descargar Documento Word (.docx)")
    st.info("Para obtener tu código de descarga oficial en Word, solicítalo directamente a través de WhatsApp:")

    CODIGO_ACCESO_MAESTRO = "APA2026"

    mensaje_wa = "Hola, he generado mi trabajo en formato APA 7 y deseo mi código de acceso para descargar el Word."
    url_whatsapp = f"https://wa.me/18297631349?text={mensaje_wa.replace(' ', '%20')}"

    st.link_button("📲 Solicitar mi código por WhatsApp", url_whatsapp, type="primary")

    st.markdown("<br>", unsafe_allow_html=True)
    codigo_input = st.text_input("Introduce el código recibido:", type="password", placeholder="Ingresa el código aquí...")

    if codigo_input:
        if codigo_input.strip() == CODIGO_ACCESO_MAESTRO:
            st.success("✅ ¡Código verificado con éxito! Tu descarga está desbloqueada:")
            archivo_word = generar_word_apa(st.session_state["resultado_apa"])
            st.download_button(
                label="⬇️ Descargar archivo Word (.docx)",
                data=archivo_word,
                file_name="Trabajo_Completo_APA7.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("Código incorrecto. Verifica el mensaje recibido en WhatsApp.")
