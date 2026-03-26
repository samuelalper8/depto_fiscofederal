"""
ConPrev — Gerador EFD-Reinf  ·  SaaS Premium (v5.0)
=============================================================
UI Glassmorphism, Fontes Space Grotesk/Inter, Animações CSS, 
Tabelas Word Estilizadas via XML, PDF Export e Lógica Sem DARF.
"""
import streamlit as st
import io
import os
import subprocess
import tempfile
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple

from openpyxl import load_workbook
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ── Configuração da Página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="ConPrev — EFD-Reinf",
    page_icon="🌌", layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Dicionários de Dados (Omitidos parcialmente para brevidade) ───────────────
CLIENTES: Dict[str, Dict[str, str]] = {
    "Legislativo - Eldorado": {"UF": "MS", "CNPJ": "70.524.376/0001-80"},
    "Município - Uirapuru": {"UF": "GO", "CNPJ": "37.622.164/0001-60"},
    # ... (Mantenha o seu dicionário completo de clientes aqui) ...
}

RESPONSAVEIS: Dict[str, str] = {
    "Wênia Rodrigues": "1024", "Aline Moreno": "1021",
    "Gustavo Nogueira": "1023", "Rafael Reis": "1022", "Samuel Almeida": "1020"
}

# ── CSS Premium (Glassmorphism, Glow, Fade-in, Space Grotesk) ─────────────────
_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Space+Grotesk:wght@500;700&display=swap');

/* Fundo com Iluminação Ambiente Radial */
.stApp {
    background: radial-gradient(circle at 15% 50%, rgba(26, 111, 175, 0.1), transparent 25%),
                radial-gradient(circle at 85% 30%, rgba(242, 159, 5, 0.08), transparent 25%),
                #0B1E33 !important;
}

html, body, p, span, div, label, li {
    font-family: 'Inter', sans-serif !important;
    color: #dce8f2;
}

h1, h2, h3, h4, h5, h6 {
    font-family: 'Space Grotesk', sans-serif !important;
    letter-spacing: -0.5px;
}

/* Animação Global de Entrada (Fade + Slide) */
@keyframes fadeSlideUp {
    0% { opacity: 0; transform: translateY(20px); }
    100% { opacity: 1; transform: translateY(0); }
}
.block-container {
    animation: fadeSlideUp 0.8s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    padding-top: 2rem !important;
}

/* Glassmorphism nas Caixas de Inputs */
.stTextInput>div>div>input, .stDateInput>div>div>input, .stNumberInput>div>div>input, [data-baseweb="select"]>div {
    background: rgba(255, 255, 255, 0.03) !important;
    backdrop-filter: blur(10px) !important;
    -webkit-backdrop-filter: blur(10px) !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    border-radius: 12px !important;
    color: #fff !important;
    transition: all 0.3s ease !important;
}

/* Efeito Glow no Foco/Hover */
.stTextInput>div>div>input:focus, .stDateInput>div>div>input:focus, [data-baseweb="select"]>div:focus-within {
    border-color: rgba(45, 143, 212, 0.5) !important;
    box-shadow: 0 0 20px rgba(45, 143, 212, 0.2) !important;
    background: rgba(255, 255, 255, 0.06) !important;
}

/* Labels */
.stTextInput>label, .stSelectbox>label, .stDateInput>label, .stNumberInput>label {
    color: #7a95ad !important;
    font-size: 11px !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 1px;
}

/* Botão Primário Premium */
.stButton>button[kind="primary"] {
    background: linear-gradient(135deg, #F29F05, #d78904) !important;
    color: #0B1E33 !important;
    font-weight: 700 !important;
    font-family: 'Space Grotesk', sans-serif !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 12px 28px !important;
    box-shadow: 0 4px 15px rgba(242, 159, 5, 0.2) !important;
    transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
}
.stButton>button[kind="primary"]:hover {
    transform: translateY(-2px) scale(1.02) !important;
    box-shadow: 0 8px 25px rgba(242, 159, 5, 0.4) !important;
}

/* Menus de topo ocultos para imersão */
#MainMenu, footer, [data-testid="stDecoration"], [data-testid="stToolbar"] { display: none !important; }
</style>
"""
st.markdown(_CSS, unsafe_allow_html=True)

# ── Funções Auxiliares ────────────────────────────────────────────────────────
def set_cell_background(cell, fill_color: str):
    """Injeta XML para colorir o fundo de uma célula da tabela do Word."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

def safe_float(value: Any) -> float:
    if value is None: return 0.0
    try: return float(value)
    except (ValueError, TypeError): return 0.0

def _brl_fmt(valor: Any) -> str:
    return f"R$ {safe_float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

# ── Construção da Tabela Word Premium ─────────────────────────────────────────
def criar_tabela_reinf(doc: Document, dados_nfs: List[Dict[str, Any]]) -> Any:
    headers = ['Órgão', 'CNPJ Tomador', 'Nº NF', 'CNPJ Prestador', 'Total Contrib. Prev.', 'Compensação']
    
    if not dados_nfs:
        table = doc.add_table(rows=2, cols=6)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            set_cell_background(cell, "1c3f60") # Azul Escuro
            p = cell.paragraphs[0]
            r = p.add_run(h)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(10)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        row_msg = table.rows[1].cells
        row_msg[0].text = "Nenhuma retenção de INSS declarada na EFD-REINF"
        row_msg[0].merge(row_msg[5]) 
        row_msg[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return table
        
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Cabeçalho Premium
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_background(cell, "1c3f60")
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(10)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    for nf in dados_nfs:
        row = table.add_row().cells
        row[0].text = str(nf.get('Órgão', ''))
        row[1].text = str(nf.get('CNPJ Tomador', ''))
        row[2].text = str(nf.get('Nº NF', ''))
        row[3].text = str(nf.get('CNPJ Prestador', ''))
        row[4].text = _brl_fmt(nf.get('Total Contrib. Prev.'))
        row[5].text = _brl_fmt(nf.get('Compensação'))
        for cell in row:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs: run.font.size = Pt(10)

    # Linha Totalizadora Premium
    total_contrib = sum(safe_float(nf.get('Total Contrib. Prev.')) for nf in dados_nfs)
    total_compensacao = sum(safe_float(nf.get('Compensação')) for nf in dados_nfs)
    
    t_row = table.add_row().cells
    for cell in t_row: set_cell_background(cell, "f0f0f0") # Cinza claro
    
    t_row[3].text = "Total Geral"
    t_row[4].text = _brl_fmt(total_contrib)
    t_row[5].text = _brl_fmt(total_compensacao)
    for i in [3, 4, 5]:
        p = t_row[i].paragraphs[0]
        if not p.runs: p.add_run(t_row[i].text)
        for run in p.runs: 
            run.bold = True
            run.font.size = Pt(10)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    return table

def replace_everywhere(doc: Document, old: str, new: str) -> None:
    def repl(par):
        if old in par.text:
            for run in par.runs:
                if old in run.text: run.text = run.text.replace(old, new)
            if old in par.text: par.text = par.text.replace(old, new)

    for p in doc.paragraphs: repl(p)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs: repl(p)

def converter_para_pdf(docx_bytes: bytes) -> Optional[bytes]:
    """Usa o LibreOffice em Nuvem para converter DOCX para PDF de forma robusta."""
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "temp.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
            
        try:
            # Comando Linux Headless para converter para PDF
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path], 
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            pdf_path = os.path.join(tmpdir, "temp.pdf")
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    return f.read()
        except FileNotFoundError:
            st.warning("⚠️ LibreOffice não instalado no servidor. O recurso de PDF está indisponível.")
        except Exception as e:
            st.error(f"Erro ao converter PDF: {e}")
    return None

# ── App Principal ─────────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    # TELA DE LOGIN GLASSMORPHISM
    _, col, _ = st.columns([1.4, 1, 1.4])
    with col:
        st.markdown("""
        <div style="background: rgba(255,255,255,0.02); backdrop-filter: blur(20px); border: 1px solid rgba(255,255,255,0.05); padding: 40px; border-radius: 24px; text-align:center; margin-top: 10vh; box-shadow: 0 20px 40px rgba(0,0,0,0.3);">
          <div style="width:70px;height:70px;background:linear-gradient(135deg,#F29F05,#d78904);border-radius:20px;display:inline-flex;align-items:center;justify-content:center;font-size:32px;box-shadow:0 10px 30px rgba(242,159,5,.4);margin-bottom:20px">🌌</div>
          <h2 style="font-size:28px;font-weight:800;color:#fff;margin:0 0 8px; font-family:'Space Grotesk', sans-serif;">ConPrev</h2>
          <p style="font-size:12px;color:#7a95ad;letter-spacing:2px;text-transform:uppercase;margin:0 0 30px 0;">EFD-Reinf &middot; Sistema de Retenções</p>
        </div>""", unsafe_allow_html=True)
        pwd = st.text_input("Credencial de Acesso", type="password", placeholder="••••••••", label_visibility="collapsed")
        if st.button("Acessar Plataforma", type="primary", use_container_width=True):
            if pwd == st.secrets.get("APP_PASSWORD", "conprev2026"):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Credencial inválida.")
else:
    # DASHBOARD
    st.markdown("""
    <div style="display:flex;align-items:center;gap:16px;padding:10px 0 20px;">
      <div style="width:48px;height:48px;background:linear-gradient(135deg,#F29F05,#d78904);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:24px;box-shadow:0 8px 20px rgba(242,159,5,.4)">📄</div>
      <div>
        <div style="font-size:22px;font-weight:800;color:#fff;font-family:'Space Grotesk', sans-serif;">Folha de Rosto <span style="font-weight:400;color:#7a95ad;">EFD-Reinf</span></div>
      </div>
    </div>""", unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["⚙️ Gerador Oficial (Word & PDF)", "📝 Lançador Auxiliar de Tabela"])
    
    with tab1:
        st.markdown("<br>", unsafe_allow_html=True)
        colL, colR = st.columns([1, 1], gap="large")
        
        with colL:
            st.markdown("<h4 style='color:#fff; font-size:16px;'>1. Configurações do Ato</h4>", unsafe_allow_html=True)
            cliente_sel = st.selectbox("Cliente", list(CLIENTES.keys()))
            num_ato = st.text_input("Nº do Ato", value="001/2026")
            resp_sel = st.selectbox("Responsável", list(RESPONSAVEIS.keys()))
            competencia = st.text_input("Competência", value="02/2026")
            vencimento = st.text_input("Vencimento", value="20/03/2026")
            
            tipo_darf = st.radio("Tipo de Documento", ["Reinf", "Avulso", "Sem DARF"], horizontal=True)
            
        with colR:
            st.markdown("<h4 style='color:#fff; font-size:16px;'>2. Base de Dados</h4>", unsafe_allow_html=True)
            houve_retencao = st.checkbox("✅ Houve retenções a declarar?", value=True)
            
            arq_excel = None
            if houve_retencao:
                arq_excel = st.file_uploader("Upload da Planilha Excel (.xlsx)", type=["xlsx"])
            else:
                st.info("ℹ️ Declaração sem movimento. A grelha de lançamentos será omitida.")

            can_run = bool(arq_excel) if houve_retencao else True
            
            if st.button("Gerar Documentos Finais", type="primary", use_container_width=True, disabled=not can_run):
                with st.spinner("Compilando Documento..."):
                    # Extrai dados (se houver)
                    dados_nfs = []
                    if houve_retencao and arq_excel:
                        wb = load_workbook(io.BytesIO(arq_excel.getvalue()), data_only=True)
                        aba = [n for n in wb.sheetnames if n.lower().startswith("valores")][0]
                        ws = wb[aba]
                        headers = [str(c.value) for c in ws[1]]
                        dados_nfs = [dict(zip(headers, r)) for r in ws.iter_rows(min_row=2, values_only=True) if not all(c is None for c in r)]

                    # Lógica Sem DARF / Reinf / Avulso
                    chk_reinf = "☒" if tipo_darf == "Reinf" else "☐"
                    chk_avulso = "☒" if tipo_darf == "Avulso" else "☐"
                    
                    uf = CLIENTES[cliente_sel].get("UF", "")
                    
                    contexto = {
                        '{{numero_ato}}': num_ato,
                        '{{data_emissao}}': datetime.now().strftime('%d/%m/%Y'),
                        '{{municipio_uf}}': f"{cliente_sel} / {uf}" if uf else cliente_sel,
                        '{{competencia}}': competencia,
                        '{{vencimento}}': vencimento,
                        '{{responsavel}}': resp_sel,
                        '{{ramal}}': RESPONSAVEIS[resp_sel],
                        '{{check_reinf}}': chk_reinf,
                        '{{check_avulso}}': chk_avulso
                    }
                    
                    try:
                        with open("Modelo_Folha_Rosto.docx", "rb") as f:
                            doc = Document(io.BytesIO(f.read()))
                        
                        for k, v in contexto.items(): replace_everywhere(doc, k, v)
                        
                        tabela = criar_tabela_reinf(doc, dados_nfs)
                        target = next((p for p in doc.paragraphs if "{{TABELA_NOTAS}}" in p.text), None)
                        if target:
                            target._p.addnext(tabela._tbl)
                            target.text = ""
                        
                        buf_docx = io.BytesIO()
                        doc.save(buf_docx)
                        bytes_docx = buf_docx.getvalue()
                        
                        st.success("✅ Documento compilado com sucesso!")
                        
                        dl1, dl2 = st.columns(2)
                        with dl1:
                            st.download_button("📥 Baixar WORD (.docx)", data=bytes_docx, file_name=f"Folha Rosto - {cliente_sel}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                        
                        # Tenta converter para PDF
                        with st.spinner("Gerando PDF..."):
                            bytes_pdf = converter_para_pdf(bytes_docx)
                            with dl2:
                                if bytes_pdf:
                                    st.download_button("📥 Baixar PDF (.pdf)", data=bytes_pdf, file_name=f"Folha Rosto - {cliente_sel}.pdf", mime="application/pdf", use_container_width=True)
                                else:
                                    st.button("🚫 PDF Indisponível (Sem LibreOffice)", disabled=True, use_container_width=True)
                                    
                    except Exception as e:
                        st.error(f"❌ Erro crítico: {e}")
