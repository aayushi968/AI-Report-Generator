import streamlit as st
import requests
import re
import io
import time
from fpdf import FPDF
from docx import Document

st.set_page_config(page_title="Reportify AI", layout="wide", page_icon="📄", initial_sidebar_state="expanded")

if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = True

dark = st.session_state.dark_mode

if dark:
    bg         = "#0d0f18"
    sb_bg      = "#12141f"
    card       = "#181b2e"
    card2      = "#1e2138"
    border     = "#272b45"
    inp_bg     = "#1e2138"
    text       = "#f0f0f5"
    sub        = "#7b7f9e"
    accent     = "#6366f1"
    acc2       = "#818cf8"
    ban_bg     = "linear-gradient(135deg,#1a1744 0%,#2d2a6e 60%,#1a1744 100%)"
    ban_bdr    = "#3730a3"
    ban_h1     = "#e0e7ff"
    ban_p      = "#a5b4fc"
    shadow     = "0 4px 20px rgba(0,0,0,0.45)"
    tog_bg     = "rgba(255,255,255,0.08)"
    tog_bdr    = "rgba(255,255,255,0.15)"
    tog_icon   = "☀️"
else:
    bg         = "#f0f2ff"
    sb_bg      = "#ffffff"
    card       = "#ffffff"
    card2      = "#f5f7ff"
    border     = "#e0e4f5"
    inp_bg     = "#f5f7ff"
    text       = "#111827"
    sub        = "#6b7280"
    accent     = "#4f46e5"
    acc2       = "#6366f1"
    ban_bg     = "linear-gradient(135deg,#4338ca 0%,#7c3aed 60%,#4338ca 100%)"
    ban_bdr    = "#4f46e5"
    ban_h1     = "#ffffff"
    ban_p      = "#e0e7ff"
    shadow     = "0 4px 20px rgba(79,70,229,0.12)"
    tog_bg     = "rgba(79,70,229,0.1)"
    tog_bdr    = "rgba(79,70,229,0.25)"
    tog_icon   = "🌙"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@500;600;700;800&family=Source+Serif+4:opsz,wght@8..60,400;8..60,500;8..60,600&display=swap');
:root {{
    --font-ui: 'Manrope', 'Segoe UI', sans-serif;
    --font-copy: 'Source Serif 4', Georgia, serif;
}}
*, *::before, *::after {{ box-sizing: border-box; }}
.stApp, .stButton button, .stDownloadButton button, .stTextInput input, .stTextArea textarea, .stSelectbox div {{
    font-family: var(--font-ui) !important;
}}
p, li {{ font-size: 1rem; }}

/* ── Hide Streamlit chrome ── */
#MainMenu, header[data-testid="stHeader"], footer {{ display: none !important; }}
.stApp {{ background: {bg}; color: {text}; }}
.block-container {{ padding-top: 0.7rem !important; padding-bottom: 2rem !important; }}

/* ── Sidebar ── */
[data-testid="stSidebar"] {{
    min-width: 292px !important; max-width: 292px !important;
    background: {sb_bg} !important;
    border-right: 1px solid {border};
    width: 292px !important;
    transform: translateX(0) !important;
    margin-left: 0 !important;
    visibility: visible !important;
}}
[data-testid="stSidebar"] > div {{
    transform: translateX(0) !important;
    margin-left: 0 !important;
    width: 100% !important;
}}
[data-testid="stSidebarHeader"] {{
    height: 0 !important;
    min-height: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
    border: 0 !important;
}}
[data-testid="stSidebarContent"] {{
    padding: 0.38rem 0.8rem 0.55rem !important;
    height: 100dvh !important;
    overflow-y: auto !important;
    overflow-x: hidden !important;
    position: relative !important;
}}
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div {{ color: {text} !important; }}
[data-testid="collapsedControl"], button[kind="header"] {{ display: none !important; }}

.sidebar-head {{ margin-bottom: 0.05rem; }}
.sidebar-head p {{ margin: 0 !important; }}

[data-testid="stSidebarContent"] .element-container {{
    margin-bottom: 0.28rem !important;
}}
[data-testid="stSidebarContent"] [data-testid="stMarkdownContainer"] p {{
    margin-bottom: 0 !important;
}}

/* ── Inputs ── */
.stTextInput input, .stTextArea textarea {{
    background: {inp_bg} !important; color: {text} !important;
    border: 1.5px solid {border} !important; border-radius: 8px !important;
    padding: 0.66rem 0.9rem !important; font-size: 1.05rem !important;
    line-height: 1.45 !important;
    transition: border-color 0.2s, box-shadow 0.2s;
}}
.stTextArea textarea {{ min-height: 78px !important; }}
.stTextInput input:focus, .stTextArea textarea:focus {{
    border-color: {accent} !important;
    box-shadow: 0 0 0 3px {accent}20 !important;
}}
.stSelectbox > div > div {{
    background: {inp_bg} !important; color: {text} !important;
    border: 1.5px solid {border} !important; border-radius: 8px !important;
    font-size: 1.05rem !important;
    min-height: 44px !important;
}}
/* Remove default label margin */
.stTextInput > label, .stSelectbox > label, .stTextArea > label {{ display: none !important; }}
div[data-baseweb="select"] {{ margin-top: 0 !important; }}

/* ── All buttons base ── */
div.stButton > button, div.stDownloadButton > button {{
    border-radius: 8px !important; font-weight: 600 !important;
    font-size: 0.98rem !important; transition: all 0.18s !important;
    border: none !important; cursor: pointer !important;
}}

/* ── Generate button ── */
div.stButton > button {{
    background: linear-gradient(135deg, #4f46e5, #7c3aed) !important;
    color: white !important; width: 100% !important;
    padding: 0.6rem 1rem !important;
    box-shadow: 0 3px 12px {accent}55 !important;
}}
div.stButton > button:hover {{
    opacity: 0.88 !important; transform: translateY(-1px) !important;
    box-shadow: 0 5px 16px {accent}77 !important;
}}

/* ── Main-section theme toggle (inside banner area) ── */
.st-key-main_tog {{
    display: flex;
    justify-content: flex-end;
    margin-top: 0.55rem;
    margin-bottom: -2.65rem;
    padding-right: 0.9rem;
    position: relative;
    z-index: 40;
}}
.st-key-main_tog [data-testid="stButton"] {{
    margin: 0 !important;
}}
.st-key-main_tog [data-testid="stButton"] > button {{
    background: rgba(255,255,255,0.14) !important;
    color: {ban_h1} !important;
    border: 1px solid rgba(255,255,255,0.22) !important;
    border-radius: 8px !important;
    width: 2.05rem !important;
    min-width: 2.05rem !important;
    height: 2.05rem !important;
    padding: 0 !important;
    line-height: 1 !important;
    box-shadow: none !important;
    font-size: 0.95rem !important;
}}
.st-key-main_tog [data-testid="stButton"] > button:hover {{
    background: rgba(255,255,255,0.22) !important;
    transform: none !important;
    box-shadow: none !important;
}}

/* ── Regenerate button ── */
.regen > div > button {{
    background: {card2} !important; color: {acc2} !important;
    border: 1px solid {border} !important;
    padding: 0.45rem 0.95rem !important;
    width: auto !important; font-size: 0.9rem !important;
    box-shadow: none !important;
}}
.regen > div > button:hover {{
    background: {accent}18 !important; border-color: {accent} !important;
    transform: none !important;
}}

/* ── Download buttons ── */
div.stDownloadButton > button {{
    width: 100% !important; padding: 0.55rem 1rem !important;
    color: white !important;
}}
div.stDownloadButton:first-of-type > button {{
    background: linear-gradient(135deg,#dc2626,#b91c1c) !important;
    box-shadow: 0 3px 10px #dc262644 !important;
}}
div.stDownloadButton:last-of-type > button {{
    background: linear-gradient(135deg,#2563eb,#1d4ed8) !important;
    box-shadow: 0 3px 10px #2563eb44 !important;
}}
div.stDownloadButton > button:hover {{
    opacity: 0.9 !important; transform: translateY(-1px) !important;
}}

/* ── Banner ── */
.banner {{
    background: {ban_bg};
    border-radius: 14px; padding: 1.85rem 2.2rem;
    margin-bottom: 1.2rem; border: 1px solid {ban_bdr};
    box-shadow: {shadow}; position: relative; overflow: hidden;
}}
.banner::after {{
    content: ''; position: absolute;
    width: 220px; height: 220px; border-radius: 50%;
    background: rgba(255,255,255,0.04);
    top: -60px; right: -40px;
}}
.banner h1 {{
    color: {ban_h1}; font-size: 2.05rem; font-weight: 800;
    margin: 0; letter-spacing: -0.03em; line-height: 1.2;
}}
.banner p {{ color: {ban_p}; margin: 0.4rem 0 0; font-size: 1.04rem; opacity: 0.9; }}
.ban-badge {{
    display: inline-flex; align-items: center; gap: 0.3rem;
    background: rgba(255,255,255,0.13); color: {ban_p};
    border-radius: 20px; padding: 0.15rem 0.65rem;
    font-size: 0.76rem; font-weight: 700; letter-spacing: 0.07em;
    text-transform: uppercase; margin-bottom: 0.55rem;
}}

/* ── Steps ── */
.steps {{
    display: flex; align-items: center; gap: 0.4rem;
    margin-bottom: 1.1rem;
}}
.step {{
    display: flex; align-items: center; gap: 0.35rem;
    background: {card}; border: 1px solid {border};
    border-radius: 20px; padding: 0.3rem 0.85rem;
    font-size: 1rem; font-weight: 700; color: {sub};
    flex: 1; justify-content: center;
}}
.step.on {{ background: {accent}18; border-color: {accent}; color: {acc2}; }}
.snum {{
    width: 17px; height: 17px; border-radius: 50%;
    background: {border}; color: {sub};
    display: flex; align-items: center; justify-content: center;
    font-size: 0.72rem; font-weight: 700; flex-shrink: 0;
}}
.step.on .snum {{ background: {accent}; color: white; }}
.sarr {{ color: {border}; font-size: 0.75rem; flex-shrink: 0; }}

/* ── Info cards ── */
.icard {{
    background: {card}; border: 1px solid {border};
    border-radius: 12px; padding: 1.1rem 1.2rem;
    box-shadow: {shadow}; height: 100%;
}}
.icard-title {{
    font-size: 0.9rem; font-weight: 700; letter-spacing: 0.08em;
    text-transform: uppercase; color: {accent}; margin-bottom: 0.65rem;
}}
.icard li, .icard p {{ color: {sub}; font-size: 1.08rem; line-height: 1.75; margin: 0.2rem 0; }}
.icard b {{ color: {text}; }}

/* ── Empty state ── */
.empty-state {{
    text-align: center; padding: 2.5rem 1rem;
    color: {sub};
}}
.empty-state .icon {{ font-size: 2.8rem; margin-bottom: 0.6rem; }}
.empty-state h3 {{ color: {text}; font-size: 1.6rem; font-weight: 800; margin: 0 0 0.45rem; }}
.empty-state p {{ font-size: 1.2rem; margin: 0; line-height: 1.5; }}

/* ── Report box ── */
.report-box {{
    background: {card}; border: 1px solid {border};
    border-radius: 12px; padding: 2rem 2.2rem;
    line-height: 1.85; color: {text};
    box-shadow: {shadow}; font-size: 1.06rem;
    font-family: var(--font-copy) !important;
}}

/* ── Sidebar labels ── */
.slabel {{
    font-size: 0.8rem; font-weight: 700; letter-spacing: 0.09em;
    text-transform: uppercase; color: {accent};
    margin: 0.35rem 0 0.14rem;
}}

/* ── Tip ── */
.tip {{
    background: {accent}12; border: 1px solid {accent}30;
    border-radius: 6px; padding: 0.36rem 0.58rem;
    font-size: 0.92rem; color: {acc2}; margin-top: 0.3rem;
}}

/* Keep generate button visible while sidebar content scrolls */
.sb-generate {{
    position: sticky;
    bottom: 0;
    z-index: 12;
    background: {sb_bg};
    padding-top: 0.5rem;
    padding-bottom: max(0.25rem, env(safe-area-inset-bottom));
    border-top: 1px solid {border};
    box-shadow: 0 -8px 18px rgba(0,0,0,0.12);
}}
.sb-generate div.stButton > button {{
    font-size: 1.04rem !important;
    min-height: 46px !important;
    padding: 0.7rem 0.95rem !important;
}}

/* ── Export row ── */
.exp-label {{
    font-size: 0.86rem; font-weight: 700; letter-spacing: 0.09em;
    text-transform: uppercase; color: {sub}; margin-bottom: 0.4rem;
    margin-top: 1.2rem;
}}

/* Better readability for markdown text rendered inside report */
.report-box h1, .report-box h2, .report-box h3, .report-box h4 {{
    font-family: var(--font-ui) !important;
    line-height: 1.35;
    margin-top: 1rem;
    margin-bottom: 0.5rem;
}}
.report-box p, .report-box li {{
    font-size: 1.1rem;
    line-height: 1.85;
}}

/* Responsive sizing for tablets and phones */
@media (max-width: 1024px) {{
    [data-testid="stSidebar"] {{
        min-width: 282px !important;
        max-width: 282px !important;
    }}
    .banner h1 {{ font-size: 1.9rem; }}
    .banner p {{ font-size: 1.05rem; }}
    .report-box {{ font-size: 1rem; padding: 1.5rem 1.3rem; }}
    .step {{ font-size: 0.95rem; }}
    .icard li, .icard p {{ font-size: 1rem; }}
    .empty-state h3 {{ font-size: 1.35rem; }}
    .empty-state p {{ font-size: 1.08rem; }}
}}

@media (max-width: 768px) {{
    [data-testid="stSidebar"],
    [data-testid="stSidebar"][aria-expanded="false"],
    [data-testid="stSidebar"][aria-expanded="false"] > div {{
        min-width: 86vw !important;
        max-width: 86vw !important;
        width: 86vw !important;
    }}
    .block-container {{ padding-top: 0.5rem !important; padding-bottom: 1.2rem !important; }}
    [data-testid="stSidebarContent"] {{
        padding: 0.28rem 0.62rem 0.45rem !important;
    }}
    .banner {{ padding: 1.15rem 0.95rem; border-radius: 12px; }}
    .banner h1 {{ font-size: 1.45rem; }}
    .banner p {{ font-size: 1rem; }}
    .steps {{ gap: 0.3rem; flex-wrap: wrap; }}
    .step {{ font-size: 0.92rem; padding: 0.35rem 0.5rem; min-width: 31%; }}
    .sarr {{ display: none; }}
    .icard {{ padding: 1rem 0.95rem; }}
    .icard li, .icard p {{ font-size: 1rem; }}
    .empty-state h3 {{ font-size: 1.3rem; }}
    .empty-state p {{ font-size: 1.05rem; }}
    .report-box {{ padding: 1.15rem 1rem; font-size: 1rem; line-height: 1.75; }}
    div.stButton > button, div.stDownloadButton > button {{
        min-height: 44px !important;
        font-size: 1rem !important;
    }}
    .stTextInput input, .stTextArea textarea, .stSelectbox > div > div {{
        font-size: 1rem !important;
    }}
    .tip {{ font-size: 0.9rem; }}
    .sb-generate {{
        padding-top: 0.4rem;
        padding-bottom: max(0.22rem, env(safe-area-inset-bottom));
    }}
    .st-key-main_tog {{
        margin-bottom: -2.45rem;
        padding-right: 0.65rem;
    }}
    .st-key-main_tog [data-testid="stButton"] > button {{
        width: 1.9rem !important;
        min-width: 1.9rem !important;
        height: 1.9rem !important;
        font-size: 0.9rem !important;
    }}
}}

hr {{ border-color: {border}; margin: 0.55rem 0; }}
.stSpinner > div {{ border-top-color: {accent} !important; }}
[data-testid="stSidebar"] hr {{ margin: 0.34rem 0 0.48rem !important; }}
</style>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────
with st.sidebar:
    # Header
    st.markdown('<div class="sidebar-head">', unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:1.22rem;font-weight:800;color:{text};line-height:1.3;'>⚙️ Reportify AI</div>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    st.markdown(f'<div class="slabel">📌 Topic</div>', unsafe_allow_html=True)
    report_topic = st.text_input("t", "Impact of AI on Society", label_visibility="collapsed")

    st.markdown(f'<div class="slabel">📏 Length</div>', unsafe_allow_html=True)
    report_length = st.selectbox("l", ["Short", "Medium", "Long"], label_visibility="collapsed")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown(f'<div class="slabel">✏️ Custom Structure <span style="font-weight:400;text-transform:none;letter-spacing:0;font-size:0.68rem;color:{sub};">(optional)</span></div>', unsafe_allow_html=True)
    custom_prompt = st.text_area("c", label_visibility="collapsed",
        placeholder="e.g. Include case studies and Future Scope\ne.g. Add Executive Summary and Recommendations",
        height=76)
    st.markdown(f'<div class="tip">💡 Leave empty for default 4-section structure</div>', unsafe_allow_html=True)

    st.markdown('<div class="sb-generate">', unsafe_allow_html=True)
    generate_clicked = st.button("🚀 Generate Report")
    st.markdown('</div>', unsafe_allow_html=True)

# ── Banner ────────────────────────────────────────────────
if st.button(f"{tog_icon}", key="main_tog", help=f"Switch to {'Light' if dark else 'Dark'} mode"):
    st.session_state.dark_mode = not st.session_state.dark_mode
    st.rerun()

st.markdown(f"""
<div class="banner">
    <div class="ban-badge">✦ AI Powered</div>
    <h1>📄 Reportify AI</h1>
    <p>Create structured, professional reports on any topic in seconds</p>
</div>
""", unsafe_allow_html=True)

# ── Steps ─────────────────────────────────────────────────
has_report = 'report_content' in st.session_state
s = ["on" if not has_report else "", "on" if has_report else "", "on" if has_report else ""]
st.markdown(f"""
<div class="steps">
    <div class="step {s[0]}"><span class="snum">1</span>Configure</div>
    <span class="sarr">›</span>
    <div class="step {s[1]}"><span class="snum">2</span>View Report</div>
    <span class="sarr">›</span>
    <div class="step {s[2]}"><span class="snum">3</span>Export</div>
</div>
""", unsafe_allow_html=True)

# ── Info cards / Empty state ──────────────────────────────
if not has_report:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
        <div class="icard">
            <div class="icard-title">📋 Default Structure</div>
            <ol style="padding-left:1rem;margin:0;">
                <li><b>Introduction</b> — Overview & significance</li>
                <li><b>Key Areas</b> — Core discussion</li>
                <li><b>Challenges</b> — Risks & limitations</li>
                <li><b>Conclusion</b> — Summary & outlook</li>
            </ol>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="icard">
            <div class="icard-title">💡 Sample Prompts</div>
            <ul style="padding-left:1rem;margin:0;">
                <li>"Include real-world case studies"</li>
                <li>"Use formal academic tone"</li>
                <li>"Add Executive Summary & Recommendations"</li>
                <li>"Structure as: Abstract, Analysis, Findings"</li>
            </ul>
        </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div class="empty-state">
        <div class="icon">📝</div>
        <h3>Ready to generate your report</h3>
        <p>Enter a topic in the sidebar and click <b>Generate Report</b></p>
    </div>""", unsafe_allow_html=True)

# ── API ───────────────────────────────────────────────────
def generate_report(topic, length, prompt=''):
    API_KEY = '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'
    API_URL = 'https://openrouter.ai/api/v1/chat/completions'
    payload = {
        "model": "google/gemma-3-27b-it:free",
        "messages": [{"role": "user", "content": f"Generate a {length} report on {topic}. {prompt}"}]
    }
    headers = {'Authorization': f'Bearer {API_KEY}', 'Content-Type': 'application/json'}
    try:
        r = requests.post(API_URL, json=payload, headers=headers, timeout=60)
        r.raise_for_status()
        if not r.text.strip():
            st.error("Empty response from API.")
            return None
        return r.json().get('choices', [{}])[0].get('message', {}).get('content')
    except requests.exceptions.Timeout:
        st.error("Request timed out.")
    except requests.exceptions.ConnectionError:
        st.error("Connection failed.")
    except requests.exceptions.HTTPError as e:
        if r.status_code == 429:
            with st.spinner("Rate limit hit. Retrying in 30 seconds..."):
                time.sleep(30)
            return generate_report(topic, length, prompt)
        st.error(f"HTTP Error: {e}")
    except Exception as e:
        st.error(f"Error: {e}")
    return None

# ── Markdown Parser ───────────────────────────────────────
def parse_md(content):
    out = []
    for line in content.split('\n'):
        if re.match(r'^#{1,6}\s+', line):
            out.append(('h', re.sub(r'^#{1,6}\s+', '', line).strip()))
        elif re.match(r'^[\*\-]\s+', line):
            out.append(('b', re.sub(r'^[\*\-]\s+', '', line).strip()))
        elif line.strip() == '':
            out.append(('e', ''))
        else:
            out.append(('t', line))
    return out

def ct(t):
    return t.encode('latin-1', 'replace').decode('latin-1')

# ── PDF ───────────────────────────────────────────────────
def gen_pdf(topic, content):
    pdf = FPDF(); pdf.set_margins(20,20,20); pdf.add_page()
    pw = pdf.w - 40
    pdf.set_font("Arial",'B',16); pdf.multi_cell(pw,10,ct(f"Report on: {topic}")); pdf.ln(5)
    for k, t in parse_md(content):
        if k=='e': pdf.ln(3)
        elif k=='h':
            pdf.ln(2); pdf.set_font("Arial",'B',13); pdf.multi_cell(pw,8,ct(t)); pdf.ln(1)
        elif k=='b':
            pdf.set_font("Arial",'',11); pdf.set_x(25)
            pdf.multi_cell(pw-5,6,ct("- "+re.sub(r'\*+','',t).strip()))
        else:
            for p in re.split(r'(\*\*.*?\*\*)',t):
                if p.startswith('**') and p.endswith('**'):
                    pdf.set_font("Arial",'B',11); pdf.write(6,ct(p[2:-2]))
                else:
                    pdf.set_font("Arial",'',11); pdf.write(6,ct(p))
            pdf.ln(6)
    return bytes(pdf.output())

# ── Word ──────────────────────────────────────────────────
def gen_word(topic, content):
    doc = Document(); doc.add_heading(f"Report on: {topic}",0)
    for k, t in parse_md(content):
        if k=='e': continue
        elif k=='h': doc.add_heading(t,level=2)
        elif k=='b': doc.add_paragraph(style='List Bullet').add_run(re.sub(r'\*+','',t).strip())
        else:
            p = doc.add_paragraph()
            for part in re.split(r'(\*\*.*?\*\*)',t):
                if part.startswith('**') and part.endswith('**'):
                    p.add_run(part[2:-2]).bold = True
                else:
                    p.add_run(part)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

# ── Generate ──────────────────────────────────────────────
if generate_clicked:
    wait_msg = st.info("⏳ Generating report, please wait...")
    with st.spinner("✨ Generating your report. Please wait..."):
        result = generate_report(report_topic, report_length, custom_prompt)
    if result:
        st.session_state['report_content'] = result
        st.session_state['report_topic'] = report_topic
        st.rerun()
    wait_msg.empty()

# ── Display ───────────────────────────────────────────────
if has_report:
    content = st.session_state['report_content']
    topic   = st.session_state['report_topic']

    rc1, rc2 = st.columns([6, 1])
    with rc1:
        st.markdown(f"<div style='font-size:1.28rem;font-weight:700;color:{text};margin-bottom:0.45rem;'>📝 {topic}</div>", unsafe_allow_html=True)
    with rc2:
        st.markdown('<div class="regen">', unsafe_allow_html=True)
        if st.button("🔄 Regenerate"):
            with st.spinner("✨ Regenerating..."):
                result = generate_report(topic, report_length, custom_prompt)
            if result:
                st.session_state['report_content'] = result
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(f'<div class="report-box">{content}</div>', unsafe_allow_html=True)

    st.markdown(f'<div class="exp-label">📥 Export Report</div>', unsafe_allow_html=True)
    e1, e2 = st.columns(2)
    with e1:
        st.download_button("⬇ Download as PDF", data=gen_pdf(topic, content),
                           file_name=f"{topic}.pdf", mime="application/pdf")
    with e2:
        st.download_button("⬇ Download as Word", data=gen_word(topic, content),
                           file_name=f"{topic}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
