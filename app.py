import streamlit as st
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(
    page_title="Mapa de Procesos SEO & ASO",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ──────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Rubik:wght@300;400;500;600;700&display=swap');

/* Reset streamlit defaults */
html, body, [class*="css"] { font-family: 'Rubik', sans-serif !important; }
#MainMenu, footer, header, [data-testid="stHeader"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
.block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; max-width: 1100px !important; }
[data-testid="stSidebar"] { background-color: #111 !important; }
[data-testid="stSidebar"] > div { background-color: #111 !important; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: #ccc; }

/* ── Password screen ── */
.login-center {
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; min-height: 65vh; padding: 20px;
}
.login-card {
    background: #141414; border: 1px solid #2a2a2a; border-radius: 16px;
    padding: 48px 40px; max-width: 380px; width: 100%; text-align: center;
    box-shadow: 0 20px 60px rgba(0,0,0,0.5);
}
.login-card .lock { font-size: 3rem; margin-bottom: 16px; display: block; }
.login-card h2 { font-size: 1.5rem; font-weight: 700; color: #fff; margin: 0 0 6px; }
.login-card p { color: #777; font-size: 0.9rem; margin: 0 0 24px; }
/* Force all containers on login page to have no visible background */
[data-testid="stVerticalBlock"] { background: transparent !important; }
[data-testid="column"] { background: transparent !important; }

/* ── Section titles ── */
.sec-title {
    text-align: center; padding: 40px 0 8px;
}
.sec-title h1 {
    font-size: 2.2rem; font-weight: 700; color: #fff;
    margin: 0 0 4px; letter-spacing: -0.5px;
}
.sec-title span {
    font-size: 0.8rem; color: #666; text-transform: uppercase;
    letter-spacing: 3px; font-weight: 300;
}

/* ── Cards ── */
.card-box {
    background: #141414; border: 1px solid #222; border-radius: 12px;
    padding: 28px 24px; height: 100%; transition: all 0.25s ease;
    cursor: default;
}
.card-box:hover { border-color: #444; background: #1a1a1a; transform: translateY(-3px); box-shadow: 0 8px 30px rgba(0,0,0,0.3); }
.card-box .icon { font-size: 2rem; margin-bottom: 10px; display: block; }
.card-box .title { font-size: 1.1rem; font-weight: 600; color: #fff; margin-bottom: 8px; }
.card-box .desc { font-size: 0.85rem; color: #888; line-height: 1.55; }

/* ── Page titles ── */
.page-head {
    padding: 10px 0 24px; border-bottom: 1px solid #1a1a1a; margin-bottom: 28px;
}
.page-head h1 { font-size: 2rem; font-weight: 700; color: #fff; margin: 0; }

/* ── Process flow ── */
.flow-box {
    background: #0d0d0d; border: 1px solid #222; border-radius: 8px;
    padding: 20px 24px; font-family: 'Courier New', monospace;
    color: #6ee7a0; font-size: 0.82rem; line-height: 1.7;
    white-space: pre-wrap; word-wrap: break-word; margin: 16px 0;
    overflow-x: auto;
}

/* ── Code block ── */
.code-box {
    background: #0d0d0d; border: 1px solid #222; border-radius: 6px;
    padding: 12px 16px; font-family: 'Courier New', monospace;
    color: #6ee7a0; font-size: 0.85rem; margin: 10px 0;
}

/* ── Info cards in detail pages ── */
.info-card {
    background: #111; border: 1px solid #1e1e1e; border-radius: 10px;
    padding: 24px; margin-bottom: 16px;
}
.info-card h3 { color: #fff; font-size: 1.05rem; margin: 0 0 12px; font-weight: 600; }
.info-card p { color: #aaa; line-height: 1.75; margin-bottom: 10px; font-size: 0.92rem; }
.info-card ul { margin: 8px 0 8px 20px; padding: 0; }
.info-card li { color: #aaa; line-height: 1.6; margin-bottom: 6px; font-size: 0.92rem; }
.info-card strong { color: #ddd; }
.info-card a { color: #4a9eff; text-decoration: none; }
.info-card a:hover { text-decoration: underline; }

/* ── Banner SVG ── */
.svg-banner {
    width: 100%; overflow-x: auto; background: #111; border: 1px solid #1e1e1e;
    border-radius: 10px; margin-bottom: 24px; padding: 8px;
}

/* ── Streamlit button overrides ── */
.stButton > button {
    border-radius: 8px !important; font-family: 'Rubik', sans-serif !important;
    font-weight: 500 !important; transition: all 0.2s !important;
}

/* ── Expander overrides ── */
[data-testid="stExpander"] {
    background: #111 !important; border: 1px solid #1e1e1e !important;
    border-radius: 10px !important; margin-bottom: 12px !important;
}
[data-testid="stExpander"] summary span { font-weight: 600 !important; font-size: 1rem !important; }

/* ── Divider ── */
hr { border-color: #1a1a1a !important; }

/* ── Sidebar nav buttons ── */
[data-testid="stSidebar"] .stButton > button {
    text-align: left !important; justify-content: flex-start !important;
    background: transparent !important; border: none !important;
    color: #aaa !important; padding: 6px 12px !important;
    font-size: 0.9rem !important; border-radius: 6px !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: #1a1a1a !important; color: #fff !important;
}
</style>
""", unsafe_allow_html=True)

# ── Session state ────────────────────────────────
if "auth" not in st.session_state:
    st.session_state.auth = False
if "page" not in st.session_state:
    st.session_state.page = "index"


def go(p):
    st.session_state.page = p


# ── Word generator ───────────────────────────────
def make_word(title, sections):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    h = doc.add_heading(title, level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for sec_title, blocks in sections:
        doc.add_heading(sec_title, level=1)
        for b in blocks:
            t = b["type"]
            if t == "p":
                doc.add_paragraph(b["text"])
            elif t == "h2":
                doc.add_heading(b["text"], level=2)
            elif t == "h3":
                doc.add_heading(b["text"], level=3)
            elif t == "li":
                doc.add_paragraph(b["text"], style="List Bullet")
            elif t == "code":
                p = doc.add_paragraph()
                r = p.add_run(b["text"])
                r.font.name = "Consolas"
                r.font.size = Pt(9)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def word_btn(title, sections, filename):
    w = make_word(title, sections)
    st.download_button("📥 Descargar Word", w, filename,
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=False)


# ══════════════════════════════════════════════════
# PASSWORD
# ══════════════════════════════════════════════════
def page_login():
    # Extra CSS to center the Streamlit widgets on this page
    st.markdown("""<style>
    .block-container { max-width: 400px !important; margin: auto !important; }
    .block-container { display: flex; flex-direction: column; justify-content: center; min-height: 80vh; }
    .stTextInput > div > div > input {
        background: #1a1a1a !important; border: 1px solid #333 !important;
        border-radius: 8px !important; color: #fff !important;
        padding: 12px 16px !important; font-size: 1rem !important;
    }
    .stTextInput > div > div > input:focus { border-color: #4a9eff !important; }
    </style>""", unsafe_allow_html=True)

    st.markdown("""<div class="login-center"><div class="login-card">
        <span class="lock">🔒</span>
        <h2>Acceso privado</h2>
        <p>Ingresa la contraseña para continuar</p>
    </div></div>""", unsafe_allow_html=True)

    pw = st.text_input("Contraseña", type="password", label_visibility="collapsed",
                       placeholder="Contraseña...")
    if st.button("**Entrar**", type="primary", use_container_width=True):
        if pw.strip() == "SEOLAM2026":
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")


# ══════════════════════════════════════════════════
# INDEX
# ══════════════════════════════════════════════════
def page_index():
    # SEO
    st.markdown('<div class="sec-title"><h1>Mapa de Procesos SEO</h1><span>Technical SEO</span></div>', unsafe_allow_html=True)

    seo = [
        ("🔍", "Redirecciones 404", "Gestión de errores 404, redirecciones 301 y preservación de autoridad SEO.", "auditoria"),
        ("🔗", "Estructura de URLs", "URLs amigables, jerarquía lógica y canonicalización de páginas.", "urls"),
        ("📱", "Mobile Optimization", "Responsive design, velocidad en móvil y Core Web Vitals.", "mobile"),
        ("🚀", "Rendimiento", "Velocidad, compresión de imágenes, caching y optimización de recursos.", "rendimiento"),
        ("🏷️", "Metadata & Schema", "Meta tags, structured data y schema markup.", "metadata"),
        ("📑", "Indexación", "Sitemaps, robots.txt, crawl budget y estrategias de indexación.", "indexacion"),
    ]

    for row_start in range(0, len(seo), 3):
        cols = st.columns(3, gap="medium")
        for i, col in enumerate(cols):
            idx = row_start + i
            if idx < len(seo):
                icon, title, desc, pid = seo[idx]
                with col:
                    st.markdown(f"""<div class="card-box">
                        <span class="icon">{icon}</span>
                        <div class="title">{title}</div>
                        <div class="desc">{desc}</div>
                    </div>""", unsafe_allow_html=True)
                    if st.button("Ver proceso →", key=f"b_{pid}", use_container_width=True):
                        go(pid)
                        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    st.divider()

    # ASO
    st.markdown('<div class="sec-title"><h1>Mapa de Procesos ASO</h1><span>App Store Optimization</span></div>', unsafe_allow_html=True)

    aso = [
        ("⚙️", "ASO Optimizations", "Metadatos de apps, keywords, icono y visibilidad en tiendas.", "aso_opt"),
        ("📊", "In App Events", "Configuración de eventos in-app para analytics y conversión.", "in_app"),
        ("🧪", "AB Test", "Diseño, ejecución y análisis de pruebas A/B.", "ab_test"),
    ]

    cols = st.columns(3, gap="medium")
    for i, (icon, title, desc, pid) in enumerate(aso):
        with cols[i]:
            st.markdown(f"""<div class="card-box">
                <span class="icon">{icon}</span>
                <div class="title">{title}</div>
                <div class="desc">{desc}</div>
            </div>""", unsafe_allow_html=True)
            if st.button("Ver proceso →", key=f"b_{pid}", use_container_width=True):
                go(pid)
                st.rerun()


# ══════════════════════════════════════════════════
# DETAIL PAGE HELPERS
# ══════════════════════════════════════════════════
def back():
    if st.button("← Volver al mapa"):
        go("index")
        st.rerun()


def page_header(icon, title):
    back()
    st.markdown(f'<div class="page-head"><h1>{icon} {title}</h1></div>', unsafe_allow_html=True)


def info(html):
    st.markdown(f'<div class="info-card">{html}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════
# REDIRECCIONES 404
# ══════════════════════════════════════════════════
def page_auditoria():
    page_header("🔍", "Redirecciones 404")

    # SVG banner
    st.markdown("""<div class="svg-banner">
    <svg viewBox="0 0 1850 350" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="xMidYMid meet" style="width:100%;min-width:900px;">
        <rect width="1850" height="350" fill="#111"/>
        <defs>
            <marker id="a1" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto"><polygon points="0 0,10 3,0 6" fill="#555"/></marker>
            <marker id="a2" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto"><polygon points="0 0,10 3,0 6" fill="#ff6b6b"/></marker>
            <marker id="a3" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto"><polygon points="0 0,10 3,0 6" fill="#4ade80"/></marker>
        </defs>
        <circle cx="50" cy="175" r="22" fill="#1a1a1a" stroke="#4a9eff" stroke-width="2"/><text x="50" y="180" font-family="Rubik,sans-serif" font-size="11" font-weight="600" fill="#4a9eff" text-anchor="middle">Inicio</text>
        <line x1="72" y1="175" x2="112" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <rect x="112" y="155" width="115" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="170" y="180" font-family="Rubik,sans-serif" font-size="11" fill="#ccc" text-anchor="middle">Identificar 404</text>
        <line x1="227" y1="175" x2="267" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <rect x="267" y="155" width="115" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="325" y="180" font-family="Rubik,sans-serif" font-size="11" fill="#ccc" text-anchor="middle">Revisar GSC</text>
        <line x1="382" y1="175" x2="422" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <polygon points="470,175 520,145 570,175 520,205" fill="#1a1a1a" stroke="#eab308" stroke-width="2"/><text x="520" y="180" font-family="Rubik,sans-serif" font-size="10" font-weight="600" fill="#eab308" text-anchor="middle">¿Válido?</text>
        <line x1="520" y1="205" x2="520" y2="260" stroke="#ff6b6b" stroke-width="2" marker-end="url(#a2)"/><text x="530" y="238" font-family="Rubik,sans-serif" font-size="10" fill="#ff6b6b" font-weight="600">No</text>
        <rect x="440" y="262" width="160" height="32" rx="6" fill="#1a1a1a" stroke="#ff6b6b" stroke-width="1.5"/><text x="520" y="283" font-family="Rubik,sans-serif" font-size="10" fill="#ff6b6b" text-anchor="middle">Descartar</text>
        <line x1="570" y1="175" x2="610" y2="175" stroke="#4ade80" stroke-width="2" marker-end="url(#a3)"/><text x="586" y="165" font-family="Rubik,sans-serif" font-size="10" fill="#4ade80" font-weight="600">Sí</text>
        <rect x="610" y="155" width="110" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="665" y="180" font-family="Rubik,sans-serif" font-size="11" fill="#ccc" text-anchor="middle">Crear Ticket</text>
        <line x1="720" y1="175" x2="760" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <rect x="760" y="155" width="110" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="815" y="180" font-family="Rubik,sans-serif" font-size="11" fill="#ccc" text-anchor="middle">Definir Tipo</text>
        <line x1="870" y1="175" x2="910" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <rect x="910" y="155" width="100" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="960" y="180" font-family="Rubik,sans-serif" font-size="11" fill="#ccc" text-anchor="middle">Enviar</text>
        <line x1="1010" y1="175" x2="1050" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <rect x="1050" y="155" width="115" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="1108" y="176" font-family="Rubik,sans-serif" font-size="10" fill="#ccc" text-anchor="middle">Esperar</text><text x="1108" y="190" font-family="Rubik,sans-serif" font-size="8" fill="#666" text-anchor="middle">(1-2 días)</text>
        <line x1="1165" y1="175" x2="1205" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <rect x="1205" y="155" width="120" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="1265" y="180" font-family="Rubik,sans-serif" font-size="11" fill="#ccc" text-anchor="middle">Implementar</text>
        <line x1="1325" y1="175" x2="1365" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <rect x="1365" y="155" width="100" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="1415" y="180" font-family="Rubik,sans-serif" font-size="11" fill="#ccc" text-anchor="middle">Validar</text>
        <line x1="1465" y1="175" x2="1505" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <rect x="1505" y="155" width="100" height="40" rx="6" fill="#1a1a1a" stroke="#333" stroke-width="1.5"/><text x="1555" y="180" font-family="Rubik,sans-serif" font-size="11" fill="#ccc" text-anchor="middle">Cerrar</text>
        <line x1="1605" y1="175" x2="1645" y2="175" stroke="#555" stroke-width="2" marker-end="url(#a1)"/>
        <circle cx="1685" cy="175" r="22" fill="#1a1a1a" stroke="#4ade80" stroke-width="2"/><text x="1685" y="180" font-family="Rubik,sans-serif" font-size="11" font-weight="600" fill="#4ade80" text-anchor="middle">Fin</text>
    </svg></div>""", unsafe_allow_html=True)

    with st.expander("📋 Información General 404", expanded=True):
        info("""
        <h3>¿Qué es un Error 404?</h3>
        <p>Un error 404 (Not Found) ocurre cuando un usuario intenta acceder a una página que no existe. Gestionar correctamente estos errores es crucial para mantener buena experiencia de usuario y SEO.</p>
        <h3>¿Por qué importan?</h3>
        <ul>
            <li><strong>Experiencia del Usuario:</strong> Un 404 sin redirección crea frustración.</li>
            <li><strong>Pérdida de Autoridad SEO:</strong> Los enlaces rotos disipan link juice.</li>
            <li><strong>Rastreo Ineficiente:</strong> Google desperdicia crawl budget en URLs 404.</li>
            <li><strong>Impacto en Conversiones:</strong> Los usuarios abandonan sitios con errores.</li>
        </ul>
        <h3>Tipos de Redirecciones</h3>
        <ul>
            <li><strong>301 (Permanente):</strong> Transfiere 90-99% de autoridad. La opción preferida.</li>
            <li><strong>302 (Temporal):</strong> Para cambios temporales.</li>
            <li><strong>307 (Temporal):</strong> Similar a 302, más estricto con método HTTP.</li>
            <li><strong>Meta Refresh / JavaScript:</strong> No recomendadas para SEO.</li>
        </ul>""")

    with st.expander("🔄 Proceso de Redirección"):
        info("<h3>Objetivo</h3><p>Gestionar solicitudes de redirecciones 404 para asegurar buena experiencia de usuario y minimizar impactos SEO.</p>")

        st.markdown("""<div class="flow-box">Inicio
   ↓
Identificación de URL 404
   ↓
Validar si la redirección tiene sentido SEO
   ↓
Revisar información en Google Search Console
   ↓
¿La redirección es válida?
   ├── No → Descartar solicitud / reevaluar destino
   └── Sí → Ir al template de Confluence
          ↓
Completar ticket de redirección
          ↓
Definir tipo: Directa o Dinámica
          ↓
Enviar ticket → Esperar 1-2 días hábiles
          ↓
Implementación → Validar → Cerrar
          ↓
Fin</div>""", unsafe_allow_html=True)

        info("""
        <h3>URL importante</h3>
        <p><a href="https://confluence.tools.3stripes.net/spaces/CDN/pages/1613803193/CDN+Service+Catalog" target="_blank">📎 CDN Service Catalog — Confluence</a></p>
        <p>Dependiendo del objetivo (eliminar o crear redirección), ver la opción correspondiente al hacer scroll down.</p>

        <h3>Tipos de Redirección</h3>
        <p><strong>Directa:</strong> URL específica → URL específica</p>
        <p><strong>Dinámica:</strong> Basada en patrones o reglas</p>

        <h3>QA y Validación Final</h3>
        <ul>
            <li>Redirección 301 correcta y destino funcional</li>
            <li>Sin cadenas de redirección ni loops</li>
            <li>Correcto en desktop y mobile</li>
            <li>Herramientas: Redirect Path, DevTools, Screaming Frog, GSC</li>
        </ul>""")

    with st.expander("👥 Contactos"):
        st.warning("En caso **URGENTE**, contactar por Teams:")
        info("""<ul>
            <li><strong>Geison Garzón:</strong> Geison.Garzon@externals.adidas.com</li>
            <li><strong>Oksana Tretiak:</strong> Oksana.Tretiak@externals.adidas.com</li>
        </ul>""")

    word_btn("Redirecciones 404", [
        ("Información General", [
            {"type": "p", "text": "Un error 404 ocurre cuando un usuario accede a una página inexistente."},
            {"type": "li", "text": "Experiencia del Usuario: frustración."},
            {"type": "li", "text": "Pérdida de Autoridad SEO: disipa link juice."},
            {"type": "li", "text": "Rastreo Ineficiente: desperdicia crawl budget."},
        ]),
        ("Proceso", [
            {"type": "li", "text": "1. Validación → 2. Crear Ticket → 3. Definir Tipo → 4. Enviar → 5. Validar → 6. Cerrar"},
        ]),
        ("Contactos", [
            {"type": "li", "text": "Geison Garzón: Geison.Garzon@externals.adidas.com"},
            {"type": "li", "text": "Oksana Tretiak: Oksana.Tretiak@externals.adidas.com"},
        ]),
    ], "Redirecciones_404.docx")


# ══════════════════════════════════════════════════
# URLS
# ══════════════════════════════════════════════════
def page_urls():
    page_header("🔗", "Estructura de URLs")
    with st.expander("📋 Principios y Buenas Prácticas", expanded=True):
        info("""
        <h3>Importancia</h3>
        <p>Las URLs bien estructuradas ayudan a los motores de búsqueda a entender la jerarquía del sitio y mejoran la experiencia del usuario.</p>
        <h3>Principios de una URL Ideal</h3>
        <ul>
            <li><strong>Claridad:</strong> Descriptivas y fáciles de entender.</li>
            <li><strong>Brevedad:</strong> Menos de 75 caracteres.</li>
            <li><strong>Palabras Clave:</strong> Términos relevantes del contenido.</li>
            <li><strong>Guiones (-):</strong> Separar palabras con guiones, no guiones bajos.</li>
            <li><strong>Minúsculas:</strong> Siempre en minúsculas.</li>
            <li><strong>Sin parámetros innecesarios.</strong></li>
        </ul>
        <h3>Ejemplos</h3>
        <ul>
            <li>✅ ejemplo.com/blog/seo-tecnico-guia-completa</li>
            <li>✅ ejemplo.com/servicios/consultoria-seo</li>
            <li>❌ ejemplo.com/p?id=123&cat=seo</li>
            <li>❌ ejemplo.com/blog_post_seo_tecnico</li>
        </ul>
        <h3>Canonicalización</h3>
        <p>Implementar etiquetas canónicas para evitar contenido duplicado. Indica a los buscadores la versión preferida de cada página.</p>""")
    word_btn("Estructura de URLs", [("Principios", [{"type": "li", "text": "Claridad, Brevedad, Keywords, Guiones, Minúsculas"}])], "Estructura_URLs.docx")


# ══════════════════════════════════════════════════
# MOBILE
# ══════════════════════════════════════════════════
def page_mobile():
    page_header("📱", "Mobile Optimization")
    with st.expander("📋 Mobile-First & Core Web Vitals", expanded=True):
        info("""
        <h3>Mobile-First</h3>
        <p>Google indexa principalmente la versión móvil de tu sitio. La optimización móvil es esencial.</p>
        <h3>Core Web Vitals</h3>
        <ul>
            <li><strong>LCP (Largest Contentful Paint):</strong> Velocidad de carga → &lt; 2.5s</li>
            <li><strong>FID (First Input Delay):</strong> Interactividad → &lt; 100ms</li>
            <li><strong>CLS (Cumulative Layout Shift):</strong> Estabilidad visual → &lt; 0.1</li>
        </ul>
        <h3>Checklist</h3>
        <ul>
            <li>Diseño responsive para todos los tamaños</li>
            <li>Texto legible sin zoom</li>
            <li>Imágenes optimizadas para conexiones lentas</li>
            <li>Sin intersticiales intrusivos</li>
            <li>Botones con tamaño adecuado para tocar</li>
            <li>Minimizar redirecciones y código innecesario</li>
        </ul>
        <p><strong>Testing:</strong> Google Mobile-Friendly Test, Lighthouse, Chrome DevTools.</p>""")
    word_btn("Mobile Optimization", [("Core Web Vitals", [{"type": "li", "text": "LCP < 2.5s, FID < 100ms, CLS < 0.1"}])], "Mobile_Optimization.docx")


# ══════════════════════════════════════════════════
# RENDIMIENTO
# ══════════════════════════════════════════════════
def page_rendimiento():
    page_header("🚀", "Rendimiento")
    with st.expander("📋 Velocidad como Factor de Ranking", expanded=True):
        info("""
        <h3>Estrategias de Optimización</h3>
        <ul>
            <li><strong>Compresión de Imágenes:</strong> Formatos WebP.</li>
            <li><strong>Minificación:</strong> Reducir CSS, JS y HTML.</li>
            <li><strong>Lazy Loading:</strong> Cargar contenido al ser visible.</li>
            <li><strong>Caching:</strong> Cachés de navegador y servidor.</li>
            <li><strong>CDN:</strong> Servidores cercanos a los usuarios.</li>
            <li><strong>Code Splitting:</strong> Dividir JavaScript.</li>
            <li><strong>Eliminar Render-Blocking:</strong> Optimizar recursos críticos.</li>
        </ul>
        <h3>Herramientas</h3>
        <ul>
            <li>Google PageSpeed Insights</li>
            <li>Google Lighthouse</li>
            <li>WebPageTest / GTmetrix / Pingdom</li>
        </ul>
        <p><strong>Impacto:</strong> Cada segundo de delay puede disminuir conversiones entre 4-7%.</p>""")
    word_btn("Rendimiento", [("Estrategias", [{"type": "li", "text": "Compresión, Minificación, Lazy Loading, Caching, CDN"}])], "Rendimiento.docx")


# ══════════════════════════════════════════════════
# METADATA
# ══════════════════════════════════════════════════
def page_metadata():
    page_header("🏷️", "Metadata & Schema")
    with st.expander("📋 Meta Tags & Datos Estructurados", expanded=True):
        info("""
        <h3>Meta Tags Esenciales</h3>
        <ul>
            <li><strong>Title Tag:</strong> 60-70 caracteres con keywords principales.</li>
            <li><strong>Meta Description:</strong> 150-160 caracteres, atractivo para CTR.</li>
            <li><strong>Meta Robots:</strong> Controla rastreo e indexación.</li>
            <li><strong>Open Graph:</strong> Optimiza compartir en redes sociales.</li>
            <li><strong>Twitter Cards:</strong> Específicas para Twitter/X.</li>
        </ul>
        <h3>Schema Markup</h3>
        <p>Datos estructurados que ayudan a los buscadores a entender el contenido. Puede resultar en rich snippets.</p>
        <ul>
            <li>Organization, Product, Article, FAQ, LocalBusiness, Review</li>
        </ul>
        <p><strong>JSON-LD</strong> es el formato recomendado por Google. Más fácil de implementar que Microdata.</p>""")
    word_btn("Metadata & Schema", [("Meta Tags", [{"type": "li", "text": "Title, Description, Robots, Open Graph, Twitter Cards, Schema"}])], "Metadata_Schema.docx")


# ══════════════════════════════════════════════════
# INDEXACIÓN
# ══════════════════════════════════════════════════
def page_indexacion():
    page_header("📑", "Indexación")
    with st.expander("📋 Robots.txt, Sitemap & Crawl Budget", expanded=True):
        info("""
        <h3>¿Qué es la Indexación?</h3>
        <p>Proceso mediante el cual los motores de búsqueda descubren, rastrean y añaden páginas a su base de datos.</p>
        <h3>Robots.txt</h3>
        <ul>
            <li>Define carpetas/archivos a evitar</li>
            <li>Especifica ubicación del sitemap XML</li>
            <li>Controla crawl delay</li>
            <li>Previene indexación de contenido duplicado</li>
        </ul>
        <h3>Sitemap XML</h3>
        <ul>
            <li>URLs principales con metadatos</li>
            <li>Fecha de última modificación</li>
            <li>Prioridad relativa y frequency de actualización</li>
        </ul>
        <h3>Crawl Budget</h3>
        <ul>
            <li>Eliminar páginas duplicadas o de bajo valor</li>
            <li>Bloquear parámetros de sesión innecesarios</li>
            <li>Usar 301s en lugar de contenido duplicado</li>
            <li>Monitorear rastreo en Google Search Console</li>
        </ul>""")
    word_btn("Indexación", [("Componentes", [{"type": "li", "text": "Robots.txt, Sitemap XML, Crawl Budget"}])], "Indexacion.docx")


# ══════════════════════════════════════════════════
# ASO OPTIMIZATIONS
# ══════════════════════════════════════════════════
def page_aso_opt():
    page_header("⚙️", "ASO Optimizations")
    with st.expander("📋 Información General ASO", expanded=True):
        info("""
        <h3>¿Qué es ASO?</h3>
        <p>Optimización de apps móviles para mejorar visibilidad en App Store y Google Play.</p>
        <h3>¿Por qué es importante?</h3>
        <ul>
            <li><strong>Mayor Visibilidad:</strong> Mejora ranking en búsquedas.</li>
            <li><strong>Más Descargas:</strong> Sin costo de adquisición.</li>
            <li><strong>Mejor Conversión:</strong> Listing optimizado convierte más.</li>
            <li><strong>Competitividad:</strong> En mercados saturados marca la diferencia.</li>
        </ul>
        <h3>Elementos Clave</h3>
        <ul>
            <li><strong>Nombre:</strong> Palabras clave relevantes y memorable.</li>
            <li><strong>Keywords:</strong> Alto volumen, baja competencia.</li>
            <li><strong>Descripción:</strong> Orientada a beneficios.</li>
            <li><strong>Icono:</strong> Atractivo y distinguible.</li>
            <li><strong>Capturas:</strong> Funciones principales.</li>
            <li><strong>Reseñas:</strong> Influyen en algoritmo.</li>
        </ul>""")

    with st.expander("🔄 Proceso de Optimización"):
        info("""
        <h3>Etapas</h3>
        <ul>
            <li><strong>1. Keywords:</strong> Analizar competidores con App Annie, Sensor Tower, Mobile Action.</li>
            <li><strong>2. Metadatos:</strong> Nombre, subtitle, descripción, keywords.</li>
            <li><strong>3. Diseño Visual:</strong> Icono, capturas, preview video.</li>
            <li><strong>4. Config Tiendas:</strong> Categoría, precio, compatibilidad.</li>
            <li><strong>5. Testing:</strong> Búsquedas, listing, ranking.</li>
            <li><strong>6. Monitoreo:</strong> Rankings, conversión, reseñas, ajustes.</li>
        </ul>""")

    with st.expander("👥 Contactos"):
        info("""<ul>
            <li><strong>Ana López</strong> — ASO Manager</li>
            <li><strong>David Pérez</strong> — App Marketing Specialist</li>
            <li><strong>Elena Rodríguez</strong> — UX/UI Designer</li>
        </ul>""")
    word_btn("ASO Optimizations", [("Proceso", [{"type": "li", "text": "Keywords → Metadatos → Visual → Config → Testing → Monitoreo"}])], "ASO_Optimizations.docx")


# ══════════════════════════════════════════════════
# IN APP EVENTS
# ══════════════════════════════════════════════════
def page_in_app():
    page_header("📊", "In App Events")
    with st.expander("📋 Información General", expanded=True):
        info("""
        <h3>¿Qué son?</h3>
        <p>Acciones que los usuarios realizan dentro de una app móvil, rastreadas mediante herramientas de analytics. Generan datos cruciales para entender comportamiento, optimizar experiencia y medir campañas.</p>
        <h3>¿Por qué importan?</h3>
        <ul>
            <li><strong>Medición de Conversiones:</strong> Rastrean acciones clave.</li>
            <li><strong>Análisis de Comportamiento:</strong> Insight sobre interacciones.</li>
            <li><strong>Optimización:</strong> Identificar puntos débiles.</li>
            <li><strong>Retargeting:</strong> Campañas dirigidas.</li>
            <li><strong>ROI:</strong> Demostrar valor de inversión.</li>
        </ul>
        <h3>Tipos</h3>
        <ul>
            <li><strong>Custom Events:</strong> Personalizados del negocio.</li>
            <li><strong>Standard Events:</strong> Purchase, AddToCart, Signup.</li>
            <li><strong>Funnel Events:</strong> Flujo de conversión.</li>
            <li><strong>Revenue Events:</strong> Monetización e ingresos.</li>
        </ul>""")

    with st.expander("🔄 Proceso de Configuración"):
        info("""
        <h3>Objetivo</h3>
        <p>Organizar, planificar y activar In App Events para LAM coordinando con Global ASO y Content.</p>

        <h3>1. Validación de slots con Global ASO</h3>
        <p>Llamadas recurrentes con <strong>Federica Bello</strong> de Global ASO para conocer slots disponibles. LAM es pionero en implementación — generalmente tenemos varios slots.</p>

        <h3>2. Cronograma con Content</h3>
        <p>Reunión con <strong>Alexa Peralta</strong> de Content para construir el cronograma:</p>
        <p><a href="https://adidasgroup-my.sharepoint.com/:x:/r/personal/jose_rubertone_adidas_com/_layouts/15/Doc.aspx?sourcedoc=%7BBD2C4A94-E8A9-42EB-8093-87C9DB56043F%7D&file=LAM%20-%20ASO%20-%20COMMERCIAR%20EVENTS%20-%20Organizer.xlsx&action=default&mobileredirect=true" target="_blank">📎 LAM - ASO - COMMERCIAR EVENTS - Organizer</a></p>

        <h3>3. Push a Content por país</h3>
        <p>Seguimiento con Content responsable por país. Materiales idealmente <strong>7 días antes</strong> del go live.</p>""")

    with st.expander("🛠️ Configuración en Herramienta"):
        info("""
        <h3>Checklist inicial</h3>
        <ul>
            <li>✅ Slot confirmado con Global ASO</li>
            <li>✅ Fecha go live y terminación en cronograma</li>
            <li>✅ Materiales finales de Content</li>
            <li>✅ Textos, imágenes, país, fechas verificados</li>
        </ul>
        <h3>Materiales necesarios</h3>
        <ul>
            <li>Nombre del evento</li>
            <li>Copy aprobado</li>
            <li>Assets visuales finales</li>
            <li>País(es) de activación</li>
            <li>Fecha/hora inicio y fin</li>
        </ul>
        <h3>Validación final</h3>
        <ul>
            <li>Revisar vista previa antes de enviar</li>
            <li>Confirmar fechas vs cronograma</li>
            <li>Compartir confirmación con ASO y Content</li>
        </ul>""")

    with st.expander("👥 Contactos"):
        info("""<ul>
            <li><strong>Federica Bello</strong> — Global ASO. Confirma slots.</li>
            <li><strong>Alexa Peralta</strong> — Content. Cronograma.</li>
            <li><strong>Content por país</strong> — Materiales 7 días antes del go live.</li>
        </ul>""")
    word_btn("In App Events", [
        ("Proceso", [
            {"type": "li", "text": "1. Validación slots con Federica Bello (Global ASO)"},
            {"type": "li", "text": "2. Cronograma con Alexa Peralta (Content)"},
            {"type": "li", "text": "3. Push a Content por país — 7 días antes del go live"},
        ]),
    ], "In_App_Events.docx")


# ══════════════════════════════════════════════════
# AB TEST
# ══════════════════════════════════════════════════
def page_ab_test():
    page_header("🧪", "AB Test")
    with st.expander("📋 Información General", expanded=True):
        info("""
        <h3>¿Qué es AB Testing?</h3>
        <p>Metodología que compara dos versiones para determinar cuál performa mejor mediante división aleatoria del tráfico.</p>
        <h3>¿Por qué es importante?</h3>
        <ul>
            <li><strong>Decisiones Basadas en Datos:</strong> Elimina especulación.</li>
            <li><strong>Mejora Continua:</strong> Iterar constantemente.</li>
            <li><strong>Reducción de Riesgo:</strong> Valida antes de implementar.</li>
            <li><strong>Incremento de Conversiones:</strong> Mejoras acumulativas.</li>
        </ul>
        <h3>Elementos Comunes</h3>
        <ul>
            <li><strong>Copy:</strong> Títulos, botones, descripciones.</li>
            <li><strong>Design:</strong> Colores, layout, tipografía.</li>
            <li><strong>Visuales:</strong> Imágenes, iconos, animaciones.</li>
            <li><strong>User Flow:</strong> Orden de pasos.</li>
            <li><strong>Precios:</strong> Valores, planes, descuentos.</li>
        </ul>""")

    with st.expander("🔄 Proceso de AB Testing"):
        info("""
        <h3>Etapas</h3>
        <ul>
            <li><strong>1. Hipótesis:</strong> "Si cambio X, entonces Z mejorará". Priorizar por impacto.</li>
            <li><strong>2. Diseño:</strong> Control (A) vs Variante (B). Definir KPI, muestra, duración (mín. 2 semanas).</li>
            <li><strong>3. Config Técnica:</strong> Optimizely / Apptimize / Firebase. Distribución 50/50.</li>
            <li><strong>4. Ejecución:</strong> Monitorear diario. NO detener prematuramente.</li>
            <li><strong>5. Análisis:</strong> Significancia estadística (p-value &lt; 0.05). Documentar insights.</li>
            <li><strong>6. Decisión:</strong> B gana → implementar 100%. Empate → evaluar. A gana → mantener o nueva variante.</li>
        </ul>""")

    with st.expander("👥 Contactos"):
        info("""<ul>
            <li><strong>María Flores</strong> — Product Manager</li>
            <li><strong>Roberto Jiménez</strong> — UX Researcher</li>
            <li><strong>Camila Soto</strong> — Growth Analyst</li>
        </ul>""")
    word_btn("AB Test", [("Proceso", [
        {"type": "li", "text": "Hipótesis → Diseño → Config → Ejecución → Análisis → Decisión"},
    ])], "AB_Test.docx")


# ══════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════
PAGES = {
    "index": page_index,
    "auditoria": page_auditoria,
    "urls": page_urls,
    "mobile": page_mobile,
    "rendimiento": page_rendimiento,
    "metadata": page_metadata,
    "indexacion": page_indexacion,
    "aso_opt": page_aso_opt,
    "in_app": page_in_app,
    "ab_test": page_ab_test,
}

if not st.session_state.auth:
    page_login()
else:
    with st.sidebar:
        st.markdown("### 🗺️ Navegación")
        st.divider()
        st.caption("SEO")
        for label, pid in [("🏠 Inicio", "index"), ("🔍 Redirecciones 404", "auditoria"),
                           ("🔗 Estructura de URLs", "urls"), ("📱 Mobile", "mobile"),
                           ("🚀 Rendimiento", "rendimiento"), ("🏷️ Metadata & Schema", "metadata"),
                           ("📑 Indexación", "indexacion")]:
            if st.button(label, key=f"s_{pid}", use_container_width=True):
                go(pid); st.rerun()
        st.divider()
        st.caption("ASO")
        for label, pid in [("⚙️ ASO Optimizations", "aso_opt"), ("📊 In App Events", "in_app"),
                           ("🧪 AB Test", "ab_test")]:
            if st.button(label, key=f"s_{pid}", use_container_width=True):
                go(pid); st.rerun()
        st.divider()
        if st.button("🚪 Cerrar sesión", key="logout", use_container_width=True):
            st.session_state.auth = False
            st.session_state.page = "index"
            st.rerun()

    PAGES.get(st.session_state.page, page_index)()
