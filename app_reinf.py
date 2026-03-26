"""
ConPrev — Gerador EFD-Reinf  ·  Interface Web SaaS (v4.0)
=============================================================
Processamento 100% em memória, Injeção Dinâmica de Tabelas via XML (Sem quebrar layouts),
Replace seguro com preservação de formatação e Lógica "Sem Movimento".
"""
import streamlit as st
import io
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple

from openpyxl import load_workbook
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# ── Configuração da Página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="ConPrev — EFD-Reinf",
    page_icon="📄", layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Dados Estáticos ───────────────────────────────────────────────────────────
CLIENTES: Dict[str, Dict[str, str]] = {
    "Legislativo - Eldorado": {"UF": "MS", "CNPJ": "70.524.376/0001-80"},
    "Município - Uirapuru": {"UF": "GO", "CNPJ": "37.622.164/0001-60"},
    "Município - Santa Maria do Tocantins": {"UF": "TO", "CNPJ": "37.421.039/0001-92"},
    "Município - Lajeado": {"UF": "TO", "CNPJ": "37.420.650/0001-04"},
    "Município - Jaú do Tocantins": {"UF": "TO", "CNPJ": "37.344.413/0001-01"},
    "Município - Alcinópolis": {"UF": "MS", "CNPJ": "37.226.651/0001-04"},
    "Município - Teresina de Goiás": {"UF": "GO", "CNPJ": "25.105.339/0001-83"},
    "Município - Goianorte": {"UF": "TO", "CNPJ": "25.086.612/0001-70"},
    "Município - Palmeiras do Tocantins": {"UF": "TO", "CNPJ": "25.064.056/0001-30"},
    "Município - Maurilândia do Tocantins": {"UF": "TO", "CNPJ": "25.064.015/0001-44"},
    "Município - São Valério": {"UF": "TO", "CNPJ": "25.043.449/0001-68"},
    "Legislativo - Rio Verde": {"UF": "GO", "CNPJ": "25.040.627/0001-05"},
    "Município - Perolândia": {"UF": "GO", "CNPJ": "24.859.324/0001-48"},
    "Município - Rio Quente": {"UF": "GO", "CNPJ": "24.852.675/0001-27"},
    "Município - Sonora": {"UF": "MS", "CNPJ": "24.651.234/0001-67"},
    "Município - Chapadão do Sul": {"UF": "MS", "CNPJ": "24.651.200/0001-72"},
    "Município - Japorã": {"UF": "MS", "CNPJ": "15.905.342/0001-28"},
    "Autarquia - SAAE de Jaraguari": {"UF": "MS", "CNPJ": "15.435.936/0001-12"},
    "RPPS - Baliza": {"UF": "GO", "CNPJ": "11.329.148/0001-90"},
    "RPPS - Piranhas": {"UF": "GO", "CNPJ": "07.578.154/0001-04"},
    "RPPS - Serranópolis": {"UF": "GO", "CNPJ": "05.433.433/0001-54"},
    "RPPS - Itaberaí": {"UF": "GO", "CNPJ": "05.370.217/0001-07"},
    "Autarquia - Goiatuba IAG": {"UF": "GO", "CNPJ": "05.098.663/0001-04"},
    "RPPS - Trindade": {"UF": "GO", "CNPJ": "05.015.173/0001-05"},
    "RPPS - Barro Alto": {"UF": "GO", "CNPJ": "05.004.744/0001-06"},
    "RPPS - Crixás": {"UF": "GO", "CNPJ": "04.739.716/0001-66"},
    "Legislativo - Ceres": {"UF": "GO", "CNPJ": "04.340.201/0001-99"},
    "RPPS - Sonora": {"UF": "MS", "CNPJ": "04.318.288/0001-06"},
    "Município - Sete Quedas": {"UF": "MS", "CNPJ": "03.889.011/0001-62"},
    "Município - Tacuru": {"UF": "MS", "CNPJ": "03.888.989/0001-00"},
    "Município - Iguatemi": {"UF": "MS", "CNPJ": "03.568.318/0001-61"},
    "Município - Coxim": {"UF": "MS", "CNPJ": "03.510.211/0001-62"},
    "Município - Jaraguari": {"UF": "MS", "CNPJ": "03.501.533/0001-45"},
    "Município - Anastácio": {"UF": "MS", "CNPJ": "03.452.307/0001-11"},
    "Município - Brejinho de Nazaré": {"UF": "TO", "CNPJ": "02.884.153/0001-74"},
    "Município - Pilar de Goiás": {"UF": "GO", "CNPJ": "02.647.303/0001-26"},
    "Município - São Francisco de Goiás": {"UF": "GO", "CNPJ": "02.468.437/0001-80"},
    "Município - Itaberaí": {"UF": "GO", "CNPJ": "02.451.938/0001-53"},
    "Município - Peixe": {"UF": "TO", "CNPJ": "02.396.166/0001-02"},
    "Município - Crixás": {"UF": "GO", "CNPJ": "02.382.067/0001-63"},
    "Município - Barro Alto": {"UF": "GO", "CNPJ": "02.355.675/0001-89"},
    "Legislativo - Itapaci": {"UF": "GO", "CNPJ": "02.353.368/0001-69"},
    "Município - Córrego do Ouro": {"UF": "GO", "CNPJ": "02.321.115/0001-03"},
    "Município - São Luís de Montes Belos": {"UF": "GO", "CNPJ": "02.320.406/0001-87"},
    "Município - Goiás": {"UF": "GO", "CNPJ": "02.295.772/0001-23"},
    "Legislativo - Perolândia": {"UF": "GO", "CNPJ": "02.254.179/0001-39"},
    "Legislativo - Jaraguari": {"UF": "MS", "CNPJ": "02.210.819/0001-09"},
    "Município - Pedro Afonso": {"UF": "TO", "CNPJ": "02.070.589/0001-20"},
    "Município - Guaraí": {"UF": "TO", "CNPJ": "02.070.548/0001-33"},
    "Município - Paranaiguara": {"UF": "GO", "CNPJ": "02.056.745/0001-06"},
    "Município - Natividade": {"UF": "TO", "CNPJ": "01.809.474/0001-41"},
    "Município - Montes Claros de Goiás": {"UF": "GO", "CNPJ": "01.767.722/0001-39"},
    "Município - Brazabrantes": {"UF": "GO", "CNPJ": "01.756.741/0001-60"},
    "Município - Goiatuba": {"UF": "GO", "CNPJ": "01.753.722/0001-80"},
    "Legislativo - São Luís de Montes Belos": {"UF": "GO", "CNPJ": "01.725.501/0001-06"},
    "Município - Aguiarnópolis": {"UF": "TO", "CNPJ": "01.634.074/0001-42"},
    "Município - Novo Gama": {"UF": "GO", "CNPJ": "01.629.276/0001-04"},
    "Município - Santa Rita do Tocantins": {"UF": "TO", "CNPJ": "01.613.127/0001-49"},
    "Município - Bandeirantes do Tocantins": {"UF": "TO", "CNPJ": "01.612.819/0001-72"},
    "Município - Barra do Ouro": {"UF": "TO", "CNPJ": "01.612.818/0001-28"},
    "Autarquia - Goiatuba FESG": {"UF": "GO", "CNPJ": "01.494.665/0001-61"},
    "Município - Amaralina": {"UF": "GO", "CNPJ": "01.492.098/0001-04"},
    "Legislativo - Peixe": {"UF": "TO", "CNPJ": "01.447.812/0001-42"},
    "Município - Buriti Alegre": {"UF": "GO", "CNPJ": "01.345.909/0001-44"},
    "Município - Serranópolis": {"UF": "GO", "CNPJ": "01.343.086/0001-18"},
    "Município - Rianápolis": {"UF": "GO", "CNPJ": "01.300.094/0001-87"},
    "Município - Jaraguá": {"UF": "GO", "CNPJ": "01.223.916/0001-73"},
    "Município - Trindade": {"UF": "GO", "CNPJ": "01.217.538/0001-15"},
    "Município - Pium": {"UF": "TO", "CNPJ": "01.189.497/0001-09"},
    "Município - Piranhas": {"UF": "GO", "CNPJ": "01.168.145/0001-69"},
    "Município - Caiapônia": {"UF": "GO", "CNPJ": "01.164.946/0001-56"},
    "Município - Almas": {"UF": "TO", "CNPJ": "01.138.551/0001-89"},
    "Município - Cristalina": {"UF": "GO", "CNPJ": "01.138.122/0001-01"},
    "Município - Itapaci": {"UF": "GO", "CNPJ": "01.134.808/0001-24"},
    "Município - Ceres": {"UF": "GO", "CNPJ": "01.131.713/0001-57"},
    "Município - Corumbá de Goiás": {"UF": "GO", "CNPJ": "01.118.850/0001-51"},
    "Município - Hidrolina": {"UF": "GO", "CNPJ": "01.067.230/0001-30"},
    "Município - Cristalândia": {"UF": "TO", "CNPJ": "01.067.156/0001-52"},
    "Município - Baliza": {"UF": "GO", "CNPJ": "01.067.131/0001-59"},
    "Município - Bela Vista de Goiás": {"UF": "GO", "CNPJ": "01.005.917/0001-41"},
    "Legislativo - Costa Rica": {"UF": "MS", "CNPJ": "00.991.547/0001-04"},
    "Legislativo - Catalão": {"UF": "GO", "CNPJ": "00.833.942/0001-50"},
    "Município - Paraíso do Tocantins": {"UF": "TO", "CNPJ": "00.299.180/0001-54"},
    "Município - Campinaçu": {"UF": "GO", "CNPJ": "00.145.789/0001-79"},
    "Município - Silvanópolis": {"UF": "TO", "CNPJ": "00.114.819/0001-80"},
    "Município - Palmeirópolis": {"UF": "TO", "CNPJ": "00.007.401/0001-73"}
}

RESPONSAVEIS: Dict[str, str] = {
    "Wênia Rodrigues": "1024",
    "Aline Moreno": "1021",
    "Gustavo Nogueira": "1023",
    "Rafael Reis": "1022",
    "Samuel Almeida": "1020"
}

# ── CSS System ────────────────────────────────────────────────────────────────
_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500&display=swap');

:root{
  --navy:#0B1E33; --navy2:#0f2540; --navy3:#1c3f60; --navy4:#0d2035;
  --blue:#1a6faf; --sky:#2d8fd4; --sky-dim:rgba(45,143,212,.12);
  --amber:#F29F05; --amber2:#d78904; --amber-dim:rgba(242,159,5,.12);
  --red:#d63b3b; --green:#2a9c6b; --green-dim:rgba(42,156,107,.12);
  --text:#dce8f2; --text2:#b0c4d8; --muted:#7a95ad;
  --card:rgba(255,255,255,.035); --border:rgba(255,255,255,.07); --border2:rgba(255,255,255,.13);
  --radius:12px;
}
.stApp,[data-testid="stAppViewContainer"],[data-testid="stMain"],section[data-testid="stMain"]{background:var(--navy)!important}
[data-testid="stHeader"]{background:var(--navy2)!important;border-bottom:1px solid var(--border)!important}
html,body,.stApp,.stMarkdown,p,span,div,label,li{font-family:'Sora',sans-serif!important;color:var(--text)}

/* Inputs */
.stTextInput>div>div>input, .stDateInput>div>div>input, .stNumberInput>div>div>input{background:rgba(255,255,255,.05)!important; border:1px solid var(--border2)!important;border-radius:var(--radius)!important;color:var(--text)!important;font-size:14px!important;}
.stTextInput>div>div>input:focus, .stDateInput>div>div>input:focus, .stNumberInput>div>div>input:focus{border-color:var(--sky)!important; box-shadow:0 0 0 3px rgba(45,143,212,.18)!important}
.stTextInput>label, .stSelectbox>label, .stDateInput>label, .stNumberInput>label{color:var(--muted)!important;font-size:11px!important;font-weight:600!important;text-transform:uppercase;letter-spacing:1px}

/* Selectbox & Checkbox */
[data-baseweb="select"]>div{background:rgba(255,255,255,.05)!important;border:1px solid var(--border2)!important;border-radius:var(--radius)!important;color:var(--text)!important}
[data-baseweb="menu"]{background:var(--navy4)!important;border:1px solid var(--border2)!important;}
.stCheckbox>label{color:var(--text2)!important;font-size:13px!important;cursor:pointer}

/* Buttons */
.stButton>button[kind="primary"],button[data-testid="baseButton-primary"]{background:linear-gradient(135deg,var(--amber),var(--amber2))!important;color:#0B1E33!important;font-weight:700!important;border:none!important;border-radius:var(--radius)!important;padding:12px 28px!important;box-shadow:0 4px 16px rgba(242,159,5,.3)!important;}
.stButton>button[kind="primary"]:hover{transform:translateY(-1px)!important}
.stDownloadButton>button{background:var(--green-dim)!important;color:#4dd8a0!important;border:1px solid rgba(42,156,107,.35)!important;border-radius:var(--radius)!important;font-weight:700!important;padding:13px 24px!important;}

/* Code/Pre blocks (For Email Area) */
.stCodeBlock pre{background:var(--navy4)!important;border:1px solid var(--border)!important;border-radius:8px!important;}

/* File Uploader */
[data-testid="stFileUploader"]{background:rgba(255,255,255,.025)!important;border:1.5px dashed rgba(45,143,212,.3)!important;border-radius:var(--radius)!important;}

#MainMenu,footer,[data-testid="stDecoration"],[data-testid="stToolbar"]{display:none!important}
.block-container{padding-top:1.4rem!important;padding-bottom:2rem!important;max-width:1100px!important}
</style>
"""
st.markdown(_CSS, unsafe_allow_html=True)

# ── Gerenciamento de Estado ───────────────────────────────────────────────────
ss = st.session_state
ss.setdefault("authenticated", False)

# ── Lógica de Negócios & Arquitetura Word ─────────────────────────────────────
def get_datas_padrao() -> Tuple[str, str, str, datetime]:
    hoje = datetime.now()
    primeiro_dia = hoje.replace(day=1)
    ultimo_dia_mes_ant = primeiro_dia - timedelta(days=1)
    meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    
    comp_folha = f"{meses_pt[ultimo_dia_mes_ant.month - 1]}/{ultimo_dia_mes_ant.year}"
    comp_email = f"{ultimo_dia_mes_ant.month:02d}/{ultimo_dia_mes_ant.year}"
    
    venc_dt = datetime(hoje.year, hoje.month, 20)
    venc_str = venc_dt.strftime("%d/%m/%Y")
    return comp_folha, venc_str, comp_email, venc_dt

def safe_float(value: Any) -> float:
    if value is None: return 0.0
    try: return float(value)
    except (ValueError, TypeError): return 0.0

def _brl_fmt(valor: Any) -> str:
    val_float = safe_float(valor)
    return f"R$ {val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def read_excel_data(file_bytes: bytes) -> Optional[List[Dict[str, Any]]]:
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet_name = next((name for name in wb.sheetnames if name.strip().lower().startswith("valores")), None)
        if not sheet_name:
            st.error('Aba "Valores" não encontrada na planilha Excel.')
            return None
        sheet = wb[sheet_name]
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in sheet[1]]
        return [dict(zip(headers, row)) for row in sheet.iter_rows(min_row=2, values_only=True) if not all(cell is None for cell in row)]
    except Exception as e:
        st.error(f"Erro ao ler Excel: {e}")
        return None

# --- NOVAS FUNÇÕES DE ALTA PERFORMANCE (XML & SUBSTITUIÇÃO) ---
def replace_everywhere(doc: Document, old: str, new: str) -> None:
    """Substitui texto cirurgicamente em parágrafos, runs, tabelas e cabeçalhos preservando a formatação."""
    def repl(par):
        if old in par.text:
            # Tenta substituir de forma isolada preservando o 'run' (negrito, itálico)
            for run in par.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
            # Fallback seguro: se a palavra tiver sido quebrada em múltiplos runs pelo motor do Word
            if old in par.text:
                par.text = par.text.replace(old, new)

    for p in doc.paragraphs: repl(p)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs: repl(p)
    for s in doc.sections:
        for h in [s.header, s.first_page_header, s.footer, s.first_page_footer]:
            if h:
                for p in h.paragraphs: repl(p)

def mover_tabela_para_placeholder(doc: Document, table: Any, placeholder_text: str) -> bool:
    """Encontra a tag mágica e injeta o XML da tabela perfeitamente na sua posição."""
    target_p = None
    for p in doc.paragraphs:
        if placeholder_text in p.text:
            target_p = p
            break
            
    if target_p:
        target_p._p.addnext(table._tbl)
        target_p.text = "" # Apaga o texto {{TABELA_NOTAS}} após inserir a tabela
        return True
    return False

def criar_tabela_reinf(doc: Document, dados_nfs: List[Dict[str, Any]]) -> Any:
    """Constrói a tabela de 6 colunas de forma programática (Com e Sem Movimento)."""
    headers = ['Órgão', 'CNPJ Tomador', 'Nº NF', 'CNPJ Prestador', 'Total Contrib. Prev.', 'Compensação']
    
    # --- CENÁRIO: SEM MOVIMENTO ---
    if not dados_nfs:
        table = doc.add_table(rows=2, cols=6)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Formata cabeçalho
        for i, h in enumerate(headers):
            p = table.rows[0].cells[i].paragraphs[0]
            r = p.add_run(h)
            r.font.bold = True
            r.font.size = Pt(10)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        # Linha de Mensagem (Células Mescladas)
        row_msg = table.rows[1].cells
        row_msg[0].text = "Nenhuma retenção de INSS declarada na EFD-REINF"
        row_msg[0].merge(row_msg[5]) # Mescla as 6 colunas num único bloco
        row_msg[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Linha de Totalizadores (Zerada)
        t_row = table.add_row().cells
        t_row[3].text = "Total Geral"
        t_row[4].text = "R$ 0,00"
        t_row[5].text = "R$ 0,00"
        for i in [3, 4, 5]:
            p = t_row[i].paragraphs[0]
            if not p.runs: p.add_run(t_row[i].text)
            for run in p.runs: run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        return table
        
    # --- CENÁRIO: COM MOVIMENTO ---
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Formata cabeçalho
    for i, h in enumerate(headers):
        p = table.rows[0].cells[i].paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(10)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    # Popula dados reais
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

    # Calculadores Totais
    total_contrib = sum(safe_float(nf.get('Total Contrib. Prev.')) for nf in dados_nfs)
    total_compensacao = sum(safe_float(nf.get('Compensação')) for nf in dados_nfs)
    
    t_row = table.add_row().cells
    t_row[3].text = "Total Geral"
    t_row[4].text = _brl_fmt(total_contrib)
    t_row[5].text = _brl_fmt(total_compensacao)
    for i in [3, 4, 5]:
        p = t_row[i].paragraphs[0]
        # Aplica negrito final
        if not p.runs: p.add_run(t_row[i].text)
        for run in p.runs: 
            run.bold = True
            run.font.size = Pt(10)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    return table

def processar_word(template_bytes: bytes, contexto: Dict[str, str], dados_nfs: List[Dict[str, Any]]) -> io.BytesIO:
    doc = Document(io.BytesIO(template_bytes))
    
    # 1. Substituição Profunda e Segura em todos os elementos do Word
    for key, value in contexto.items():
        replace_everywhere(doc, key, value)

    # 2. Construção e Injeção do XML da Tabela
    tabela_xml = criar_tabela_reinf(doc, dados_nfs)
    sucesso = mover_tabela_para_placeholder(doc, tabela_xml, "{{TABELA_NOTAS}}")
    
    if not sucesso:
        raise ValueError("A tag {{TABELA_NOTAS}} não foi encontrada no ficheiro Word. Adicione esta tag no local onde deseja que os lançamentos apareçam.")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ── UI Components ─────────────────────────────────────────────────────────────
def _section(title: str, icon: str="", accent: str="#F29F05") -> None:
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:10px;padding:13px 18px 11px;background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.07);border-left:3px solid {accent};border-radius:10px;margin-bottom:15px">
      <span style="font-size:15px">{icon}</span>
      <span style="font-size:11.5px;font-weight:700;color:#b0c4d8;text-transform:uppercase;letter-spacing:1.2px">{title}</span>
    </div>""", unsafe_allow_html=True)

def render_login() -> None:
    _, col, _ = st.columns([1.4, 1, 1.4])
    with col:
        st.markdown("""
        <div style="text-align:center;margin:56px 0 32px">
          <div style="width:62px;height:62px;background:linear-gradient(145deg,#F29F05,#d78904);border-radius:16px;display:inline-flex;align-items:center;justify-content:center;font-size:28px;box-shadow:0 8px 32px rgba(242,159,5,.45);margin-bottom:16px">📄</div>
          <h2 style="font-size:24px;font-weight:800;color:#dce8f2;margin:0 0 6px">ConPrev</h2>
          <p style="font-size:11px;color:#7a95ad;letter-spacing:1.4px;text-transform:uppercase;margin:0">EFD-Reinf &middot; Acesso Restrito</p>
        </div>""", unsafe_allow_html=True)
        
        pwd = st.text_input("Senha de acesso", type="password", placeholder="••••••••")
        
        if st.button("Entrar", type="primary", use_container_width=True):
            senha_oficial = st.secrets.get("APP_PASSWORD", None)
            if not senha_oficial:
                st.error("⚠️ Infraestrutura: A variável 'APP_PASSWORD' não foi configurada nos Secrets.")
                return

            if pwd == senha_oficial:
                ss.authenticated = True
                st.rerun()
            else:
                st.error("⚠️ Senha incorreta. Acesso negado.")

def render_header():
    left, right = st.columns([5,1])
    with left:
        st.markdown("""
        <div style="display:flex;align-items:center;gap:14px;padding:6px 0 16px">
          <div style="width:42px;height:42px;flex-shrink:0;background:linear-gradient(145deg,#F29F05,#d78904);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:20px;box-shadow:0 4px 14px rgba(242,159,5,.4)">📄</div>
          <div>
            <div style="font-size:18px;font-weight:800;color:#dce8f2;line-height:1.2">ConPrev <span style="font-weight:400;color:#7a95ad;font-size:14px;margin-left:6px">Folha de Rosto EFD-Reinf</span></div>
            <div style="font-size:11px;color:#7a95ad;margin-top:2px">Automação de Documentos &nbsp;·&nbsp; Fisco Federal</div>
          </div>
        </div>""", unsafe_allow_html=True)
    with right:
        if st.button("↩ Sair", key="logout_btn"):
            ss.authenticated = False; st.rerun()

# ── Views (Tabs) ──────────────────────────────────────────────────────────────
def render_app():
    render_header()
    
    comp_folha, venc_str_padrao, comp_email, venc_dt_padrao = get_datas_padrao()
    
    tab1, tab2, tab3 = st.tabs(["📝 1. Lançador de Notas", "⚙️ 2. Gerador Folha Rosto", "✉️ 3. E-mail Padrão"])
    
    # TAB 1: LANÇADOR
    with tab1:
        st.markdown("<br>", unsafe_allow_html=True)
        _section("Edição Rápida da Planilha Excel", "✏️", accent="#2d8fd4")
        st.markdown("<p style='font-size:13px; color:var(--muted); margin-bottom: 20px;'>Copie e cole dados do seu sistema diretamente na tabela abaixo.</p>", unsafe_allow_html=True)
        
        cols = ["Órgão", "CNPJ Tomador", "Nº NF", "CNPJ Prestador", "Total Contrib. Prev.", "Compensação"]
        df_base = pd.DataFrame(columns=cols)
        for _ in range(5): df_base.loc[len(df_base)] = [None]*6
        
        df_editado = st.data_editor(df_base, num_rows="dynamic", use_container_width=True)
        
        if st.button("📥 Baixar Planilha (.xlsx)", type="primary"):
            df_limpo = df_editado.dropna(how="all")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_limpo.to_excel(writer, sheet_name="Valores", index=False)
            
            st.download_button(
                label="Baixar Arquivo Preenchido",
                data=output.getvalue(),
                file_name="EFD-REINF_Dados_Preenchidos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # TAB 2: GERADOR DE FOLHA DE ROSTO
    with tab2:
        st.markdown("<br>", unsafe_allow_html=True)
        col_left, col_right = st.columns([1, 1], gap="large")
        
        with col_left:
            _section("Configurações do Ato", "⚙️")
            cliente_sel = st.selectbox("Selecione o Cliente", list(CLIENTES.keys()))
            
            c_ato1, c_ato2 = st.columns([2, 1])
            with c_ato1:
                num_ato_int = st.number_input("Nº Inicial do Ato", min_value=1, value=1, step=1)
            with c_ato2:
                ano_ato = st.text_input("Ano", value=str(datetime.now().year))
            num_ato = f"{num_ato_int:03d}/{ano_ato}"
            
            resp_sel = st.selectbox("Responsável", list(RESPONSAVEIS.keys()))
            competencia = st.text_input("Competência", value=comp_folha)
            vencimento = st.text_input("Vencimento", value=venc_str_padrao)
            
            tipo_darf = st.radio("Tipo de DARF Emitido", ["Reinf (Via informações declaradas)", "Avulso (Excepcionalmente)"])
            is_reinf = "Reinf" in tipo_darf
            
        with col_right:
            _section("Lançamentos & Documento Oficial", "📤", accent="#2a9c6b")
            
            with st.expander("ℹ️ Instruções de Formatação (Modelo Word)", expanded=False):
                st.markdown("""
                **O seu ficheiro `Modelo_Folha_Rosto.docx` deve ter as seguintes tags mágicas:**
                - `{{numero_ato}}`, `{{data_emissao}}`, `{{municipio_uf}}`
                - `{{competencia}}`, `{{vencimento}}`, `{{responsavel}}`, `{{ramal}}`
                - `{{check_reinf}}`, `{{check_avulso}}`
                - E o mais importante: Escreva **`{{TABELA_NOTAS}}`** isolado numa linha onde deseja que a grelha de lançamentos/valores seja inserida automaticamente.
                """)
            
            houve_retencao = st.checkbox("✅ Houve retenções a declarar neste mês?", value=True)
            
            arq_excel = None
            if houve_retencao:
                arq_excel = st.file_uploader("Upload da Planilha de Lançamentos (.xlsx)", type=["xlsx"])
            else:
                st.info("ℹ️ Você marcou que não há retenções. O sistema omitirá a grelha de lançamentos e injetará a indicação oficial 'Sem Movimento'.")

            st.markdown("<br><br>", unsafe_allow_html=True)
            can_run = bool(arq_excel) if houve_retencao else True
            
            if st.button("Gerar Folha de Rosto Oficial", type="primary", use_container_width=True, disabled=not can_run):
                with st.spinner("Compilando documento a partir do modelo local..."):
                    
                    dados_nfs = []
                    if houve_retencao and arq_excel:
                        dados_nfs = read_excel_data(arq_excel.getvalue()) or []

                    uf = CLIENTES[cliente_sel].get("UF", "")
                    municipio_uf = f"{cliente_sel} / {uf}" if uf else cliente_sel
                    
                    contexto = {
                        '{{numero_ato}}': num_ato,
                        '{{data_emissao}}': datetime.now().strftime('%d/%m/%Y'),
                        '{{municipio_uf}}': municipio_uf,
                        '{{competencia}}': competencia,
                        '{{vencimento}}': vencimento,
                        '{{responsavel}}': resp_sel,
                        '{{ramal}}': RESPONSAVEIS[resp_sel],
                        '{{check_reinf}}': "☒" if is_reinf else "☐",
                        '{{check_avulso}}': "☐" if is_reinf else "☒"
                    }
                    try:
                        with open("Modelo_Folha_Rosto.docx", "rb") as f:
                            template_bytes = f.read()
                            
                        buf = processar_word(template_bytes, contexto, dados_nfs)
                        st.success("✅ Documento gerado com sucesso!")
                        st.download_button(
                            label="📥 Baixar Folha de Rosto Final (.docx)",
                            data=buf,
                            file_name=f"Folha de Rosto - {cliente_sel} - {competencia.replace('/','-')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    except FileNotFoundError:
                        st.error("❌ O ficheiro 'Modelo_Folha_Rosto.docx' não foi encontrado na mesma pasta do código. Por favor, coloque-o lá.")
                    except ValueError as ve:
                        st.error(f"❌ {ve}")
                    except Exception as e:
                        st.error(f"❌ Ocorreu um erro no processamento do Word: {e}")

    # TAB 3: GERADOR DE E-MAIL
    with tab3:
        st.markdown("<br>", unsafe_allow_html=True)
        _section("Gerador de Texto Padrão para E-mail", "✉️", accent="#e8a020")
        
        e_col1, e_col2 = st.columns(2)
        with e_col1:
            em_cliente = st.selectbox("Cliente (Para o Assunto)", list(CLIENTES.keys()), key="em_cliente")
            em_comp = st.text_input("Competência (Ex: 03/2026)", value=comp_email, key="em_comp")
            em_retencao = st.checkbox("Houve retenção nesta competência?", value=True, key="em_retencao")
        with e_col2:
            em_venc_dt = st.date_input("Data de Vencimento", value=venc_dt_padrao, format="DD/MM/YYYY", key="em_venc_dt")
            em_valor = st.text_input("Valor Bruto do Documento (Ex: 68.103,60)", value="0,00", key="em_valor", disabled=not em_retencao)

        if st.button("Gerar Texto do E-mail", type="primary"):
            dias_semana_pt = {0: "segunda-feira", 1: "terça-feira", 2: "quarta-feira", 3: "quinta-feira", 4: "sexta-feira", 5: "sábado", 6: "domingo"}
            dia_semana_str = dias_semana_pt[em_venc_dt.weekday()]
            cliente_formatado = em_cliente.upper().replace(" - ", " DE ")
            venc_str_final = em_venc_dt.strftime("%d/%m/%Y")
            
            assunto = f"Emissão DARF - EFD-REINF | {em_comp} - {cliente_formatado}"
            
            if em_retencao:
                corpo = f"""Prezados,

Encaminhamos em anexo o DARF referente à Contribuição Previdenciária (INSS) retida da nota fiscal de pagamento efetuado à pessoa jurídica, com competência {em_comp} e vencimento em {venc_str_final} ({dia_semana_str}).

Informamos que a Receita Federal não admite a dedução de valores globais sem a devida comprovação analítica. Conforme o caput do art. 116 da IN RFB nº 2.110/2022, os valores de materiais só não integram a base de cálculo da retenção "desde que comprovados". A simples menção de um valor global de material no corpo da nota, desacompanhada de um Boletim de Medição analítico que comprove os insumos efetivamente aplicados, configura omissão de materialidade.

Nesses casos, a legislação obriga o Município, na condição de responsável tributário, a aplicar a retenção de 11% sobre 100% do valor bruto do documento fiscal (R$ {em_valor}).

Observação: Solicitamos, por gentileza, a conferência do valor do DARF em conformidade com a nota fiscal e a confirmação do recebimento deste e-mail."""
            else:
                corpo = f"""Prezados,

Informamos que, após a análise das movimentações de pagamentos efetuados à pessoa jurídica na competência {em_comp}, constatamos que:

Nenhuma retenção de INSS foi declarada na EFD-REINF.

Consequentemente, não há emissão de DARF para o vencimento de {venc_str_final} ({dia_semana_str}) em relação a essas contribuições.

Qualquer dúvida ou necessidade de verificação, continuamos à disposição. Solicitamos, por gentileza, a confirmação do recebimento deste e-mail."""

            st.markdown("##### 📌 Assunto do E-mail")
            st.code(assunto, language="text")
            
            st.markdown("##### 📝 Corpo do E-mail")
            st.code(corpo, language="text")

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not ss.authenticated:
        render_login()
    else:
        render_app()
