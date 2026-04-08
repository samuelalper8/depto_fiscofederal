"""
ConPrev — Gerador EFD-Reinf  ·  SaaS Premium (v10.7 - Versão Definitiva Sem Cortes)
=============================================================
Ajuste de Largura de Colunas no WORD/PDF (Sem quebra de linha em CNPJ),
Paleta de Verde ConPrev, IA Gemini com Fallback 404, Coluna Index Oculta.
Código 100% integral.
"""
import streamlit as st
import io
import os
import json
import subprocess
import tempfile
import re
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple
from collections import defaultdict

from openpyxl import load_workbook
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:
    import google.generativeai as genai
    IA_DISPONIVEL = True
except ImportError:
    IA_DISPONIVEL = False

# ── Configuração da Página ────────────────────────────────────────────────────
st.set_page_config(page_title="ConPrev — EFD-Reinf", page_icon="🌌", layout="wide", initial_sidebar_state="collapsed")

if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if "dark_mode" not in st.session_state: st.session_state["dark_mode"] = False
if "ia_dados_importados" not in st.session_state: st.session_state["ia_dados_importados"] = []

ARQUIVO_CLIENTES = "clientes.json"
ARQUIVO_LANCAMENTOS = "lancamentos.json"

CLIENTES_PADRAO: Dict[str, Dict[str, str]] = {
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
    "Wênia Rodrigues": "1024", "Aline Moreno": "1021",
    "Gustavo Nogueira": "1023", "Rafael Reis": "1022", "Samuel Almeida": "1020"
}

def carregar_clientes() -> dict:
    if os.path.exists(ARQUIVO_CLIENTES):
        try:
            with open(ARQUIVO_CLIENTES, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    with open(ARQUIVO_CLIENTES, "w", encoding="utf-8") as f: json.dump(CLIENTES_PADRAO, f, ensure_ascii=False, indent=4)
    return CLIENTES_PADRAO

def salvar_novo_cliente(nome: str, uf: str, cnpj: str):
    clientes = carregar_clientes()
    clientes[nome] = {"UF": uf.upper(), "CNPJ": cnpj}
    with open(ARQUIVO_CLIENTES, "w", encoding="utf-8") as f: json.dump(clientes, f, ensure_ascii=False, indent=4)

def carregar_lancamentos() -> dict:
    if os.path.exists(ARQUIVO_LANCAMENTOS):
        try:
            with open(ARQUIVO_LANCAMENTOS, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return {}

def salvar_lancamentos(cliente: str, competencia: str, dados: list, modo: str = "sobrepor"):
    db = carregar_lancamentos()
    if cliente not in db: db[cliente] = {}
    
    if modo == "adicionar" and competencia in db[cliente]:
        db[cliente][competencia].extend(dados)
    else:
        db[cliente][competencia] = dados
        
    with open(ARQUIVO_LANCAMENTOS, "w", encoding="utf-8") as f: json.dump(db, f, ensure_ascii=False, indent=4)

def injetar_css():
    is_dark = st.session_state["dark_mode"]
    bg_color = "#0B1E33" if is_dark else "#F8FAFC"
    text_color = "#dce8f2" if is_dark else "#2D3748"
    heading_color = "#fff" if is_dark else "#1A365D"
    glass_bg = "rgba(255, 255, 255, 0.03)" if is_dark else "rgba(255, 255, 255, 0.8)"
    glass_border = "rgba(255, 255, 255, 0.08)" if is_dark else "rgba(0, 0, 0, 0.08)"
    label_color = "#7a95ad" if is_dark else "#4A5568"
    card_bg = "rgba(255,255,255,0.02)" if is_dark else "#FFFFFF"
    shadow = "rgba(0,0,0,0.3)" if is_dark else "rgba(0,0,0,0.05)"

    css = f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Space+Grotesk:wght@500;700&display=swap');
    .stApp {{ background: {bg_color} !important; }}
    
    html, body, p, div, label, li {{ font-family: 'Inter', sans-serif !important; color: {text_color}; }}
    h1, h2, h3, h4, h5, h6 {{ font-family: 'Space Grotesk', sans-serif !important; color: {heading_color} !important; }}
    span.material-symbols-rounded {{ font-family: 'Material Symbols Rounded' !important; }}
    
    .stTextInput>div>div>input, .stDateInput>div>div>input, .stNumberInput>div>div>input, [data-baseweb="select"]>div {{ background: {glass_bg} !important; border: 1px solid {glass_border} !important; border-radius: 8px !important; color: {text_color} !important; }}
    .stTextInput>label, .stSelectbox>label, .stDateInput>label, .stNumberInput>label {{ color: {label_color} !important; font-size: 12px !important; font-weight: 600 !important; text-transform: uppercase; }}
    .custom-card {{ background: {card_bg}; border: 1px solid {glass_border}; padding: 30px; border-radius: 16px; box-shadow: 0 10px 30px {shadow}; }}
    .stButton>button[kind="primary"] {{ background: #F29F05 !important; color: #fff !important; font-weight: bold !important; border-radius: 8px !important; border: none !important; }}
    #MainMenu, footer, [data-testid="stDecoration"], [data-testid="stToolbar"] {{ display: none !important; }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

injetar_css()

# ── Motor de Matemática Blindado e Máscara de CNPJ ────────────────────────────
def safe_float(value: Any) -> float:
    """Transforma qualquer loucura digitada num float perfeito."""
    if pd.isna(value) or value is None or value == "": return 0.0
    if isinstance(value, (int, float)): return float(value)
    
    val_str = str(value).upper().strip()
    val_str = val_str.replace('R$', '').replace('R', '').replace('$', '').strip()
    
    if '.' in val_str and ',' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
        
    val_str = re.sub(r'[^\d\.\-]', '', val_str)
    try: return float(val_str)
    except ValueError: return 0.0

def _brl_fmt(valor: Any) -> str:
    return f"R$ {safe_float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_cnpj(cnpj_val: Any) -> str:
    """Formata CNPJ inserindo pontos e traços, mesmo se digitado ou colado sem nada."""
    if pd.isna(cnpj_val) or not cnpj_val: return ""
    cnpj_str = str(cnpj_val).strip()
    
    if cnpj_str.endswith('.0'): cnpj_str = cnpj_str[:-2]
    
    digits = re.sub(r'\D', '', cnpj_str)
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    return str(cnpj_val) 

# ── Cérebro de IA (Gemini Vision com Fallback de Resiliência) ────────────────
def extrair_dados_ia_gemini(uploaded_file, api_key: str) -> Optional[Dict[str, Any]]:
    if not IA_DISPONIVEL: return None
    genai.configure(api_key=api_key)
    
    prompt = """
    Analise o documento fiscal. Extraia:
    {"Órgão": "Nome do Tomador", "CNPJ Tomador": "Apenas números", "Nº NF": "Número da Nota", "CNPJ Prestador": "Apenas números", "Total Contrib. Prev.": 0.0}
    Devolva APENAS um JSON válido.
    """
    try:
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
            
        sample_file = genai.upload_file(path=tmp_path)
        
        # 🟢 CADEIA DE FALLBACK CONTRA ERRO 404 🟢
        modelos = ['gemini-1.5-flash-latest', 'gemini-1.5-flash', 'gemini-1.0-pro', 'gemini-pro']
        response = None
        for nome in modelos:
            try:
                model = genai.GenerativeModel(nome)
                response = model.generate_content([prompt, sample_file])
                break
            except Exception:
                continue
                
        genai.delete_file(sample_file.name)
        os.remove(tmp_path)
        
        if not response:
            st.error("O Google recusou a requisição em todos os modelos. Tente novamente em breve.")
            return None
            
        txt_limpo = response.text.replace('```json', '').replace('```', '').strip()
        dados = json.loads(txt_limpo)
        
        dados["CNPJ Tomador"] = formatar_cnpj(dados.get("CNPJ Tomador", ""))
        dados["CNPJ Prestador"] = formatar_cnpj(dados.get("CNPJ Prestador", ""))
        
        v_inss = safe_float(dados.get("Total Contrib. Prev.", 0.0))
        dados["Total Contrib. Prev."] = f"R$ {v_inss:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if v_inss > 0 else ""
        dados["Compensação"] = ""
        
        return dados
    except Exception as e:
        st.error(f"Erro de IA (Possível Limite de Quota). Aguarde 1 min.")
        return None

def get_datas_padrao() -> Tuple[str, str, str, datetime]:
    hoje = datetime.now()
    primeiro_dia = hoje.replace(day=1)
    ultimo_dia_mes_ant = primeiro_dia - timedelta(days=1)
    comp_folha = f"{ultimo_dia_mes_ant.strftime('%m/%Y')}"
    return comp_folha, datetime(hoje.year, hoje.month, 20).strftime("%d/%m/%Y")

def set_cell_background(cell, fill_color: str):
    """Aplica cor de fundo hexadecimal à célula do Word."""
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill_color)
    cell._tc.get_or_add_tcPr().append(shd)

def fix_cell_width(row, widths):
    """Força a largura exata em centímetros para cada célula da linha."""
    for i, w in enumerate(widths):
        row.cells[i].width = w

# ── Gerador do Word/PDF (Agrupamento, Paleta Verde ConPrev e Larguras Exatas) ────────────────────
def criar_tabela_reinf(doc: Document, dados_nfs: List[Dict[str, Any]]) -> Any:
    hex_verde_conprev = "A9D08E" 

    headers = ['Órgão', 'CNPJ Tomador', 'Nº NF', 'CNPJ Prestador', 'Total Contrib. Prev.', 'Compensação']
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 🟢 DESLIGA O AUTOFIT E FORÇA AS LARGURAS EM CENTÍMETROS 🟢
    table.autofit = False
    table.allow_autofit = False
    
    # Larguras otimizadas: Órgão(1.5cm), Tomador(4.2cm), NF(1.5cm), Prestador(4.2cm), Prev(3.0cm), Comp(2.6cm)
    col_widths = [Cm(1.5), Cm(4.2), Cm(1.5), Cm(4.2), Cm(3.0), Cm(2.6)]
    for i, w in enumerate(col_widths): table.columns[i].width = w
    
    fix_cell_width(table.rows[0], col_widths)

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_background(cell, hex_verde_conprev)
        p = cell.paragraphs[0]; r = p.add_run(h); r.font.bold = True; r.font.size = Pt(10); p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not dados_nfs:
        row = table.add_row()
        fix_cell_width(row, col_widths)
        row.cells[0].text = "Nenhuma retenção de INSS declarada na EFD-REINF"; row.cells[0].merge(row.cells[5]); row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return table

    grupos = defaultdict(lambda: defaultdict(list))
    for nf in dados_nfs: grupos[str(nf.get('Órgão', ''))][str(nf.get('CNPJ Prestador', ''))].append(nf)

    total_geral_contrib = total_geral_comp = 0.0

    for orgao, prestadores in grupos.items():
        sub_org_contrib = sub_org_comp = 0.0
        for prestador, nfs in prestadores.items():
            sub_prest_contrib = sub_prest_comp = 0.0
            for nf in nfs:
                row = table.add_row()
                fix_cell_width(row, col_widths)
                
                row.cells[0].text = orgao; row.cells[1].text = str(nf.get('CNPJ Tomador', '')); row.cells[2].text = str(nf.get('Nº NF', ''))
                row.cells[3].text = prestador
                
                v_c = safe_float(nf.get('Total Contrib. Prev.', 0))
                v_cp = safe_float(nf.get('Compensação', 0))
                
                row.cells[4].text = _brl_fmt(v_c)
                row.cells[5].text = _brl_fmt(v_cp)
                
                sub_prest_contrib += v_c; sub_prest_comp += v_cp
                for c in row.cells: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            st_row = table.add_row()
            fix_cell_width(st_row, col_widths)
            set_cell_background(st_row.cells[0], hex_verde_conprev)
            
            st_row.cells[0].text = f"Subtotal - CNPJ {prestador}"; st_row.cells[0].merge(st_row.cells[3]); st_row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            st_row.cells[4].text = _brl_fmt(sub_prest_contrib); st_row.cells[5].text = _brl_fmt(sub_prest_comp)
            for idx in [0,4,5]:
                for r in st_row.cells[idx].paragraphs[0].runs: r.bold = True
                if idx != 0: st_row.cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            sub_org_contrib += sub_prest_contrib; sub_org_comp += sub_prest_comp

        org_row = table.add_row()
        fix_cell_width(org_row, col_widths)
        set_cell_background(org_row.cells[0], hex_verde_conprev)
        
        org_row.cells[0].text = f"Subtotal do Órgão ({orgao})"; org_row.cells[0].merge(org_row.cells[3]); org_row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        org_row.cells[4].text = _brl_fmt(sub_org_contrib); org_row.cells[5].text = _brl_fmt(sub_org_comp)
        for idx in [0,4,5]:
            for r in org_row.cells[idx].paragraphs[0].runs: r.bold = True
            if idx != 0: org_row.cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        total_geral_contrib += sub_org_contrib; total_geral_comp += sub_org_comp

    t_row = table.add_row()
    fix_cell_width(t_row, col_widths)
    set_cell_background(t_row.cells[0], hex_verde_conprev)
    
    t_row.cells[0].text = "TOTAL GERAL"; t_row.cells[0].merge(t_row.cells[3]); t_row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    t_row.cells[4].text = _brl_fmt(total_geral_contrib); t_row.cells[5].text = _brl_fmt(total_geral_comp)
    for idx in [0,4,5]:
        for r in t_row.cells[idx].paragraphs[0].runs: r.bold = True
        if idx != 0: t_row.cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
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
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "temp.docx")
        with open(docx_path, "wb") as f: f.write(docx_bytes)
        try:
            subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            pdf_path = os.path.join(tmpdir, "temp.pdf")
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f: return f.read()
        except Exception: return None
    return None

# ── UI Components Principais ──────────────────────────────────────────────────
def render_login() -> None:
    _, col, _ = st.columns([1.4, 1, 1.4])
    with col:
        st.markdown("""<style> [data-testid="stSidebar"] {display: none;} </style>""", unsafe_allow_html=True)
        st.markdown("""
        <div class="custom-card" style="margin-top: 15vh; text-align: center;">
          <div style="width:70px;height:70px;background:linear-gradient(135deg,#F29F05,#d78904);border-radius:20px;display:inline-flex;align-items:center;justify-content:center;font-size:32px;box-shadow:0 10px 30px rgba(242,159,5,.4);margin-bottom:20px">🌌</div>
          <h2 style="font-size:28px;font-weight:800;margin:0 0 8px; font-family:'Space Grotesk', sans-serif;">ConPrev</h2>
          <p style="font-size:12px;letter-spacing:2px;text-transform:uppercase;margin:0 0 30px 0; opacity: 0.7;">EFD-Reinf &middot; Sistema de Retenções</p>
        </div>""", unsafe_allow_html=True)
        
        pwd = st.text_input("Credencial de Acesso", type="password", placeholder="••••••••", label_visibility="collapsed")
        
        if st.button("Acessar Plataforma", type="primary", use_container_width=True):
            senha_oficial = st.secrets.get("APP_PASSWORD", None)
            if not senha_oficial:
                st.error("⚠️ Infraestrutura: A variável 'APP_PASSWORD' não foi configurada nos Secrets.")
                return

            if pwd == senha_oficial:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("⚠️ Senha incorreta. Acesso negado.")

def render_header():
    left, mid, right = st.columns([6, 2, 1])
    with left:
        st.markdown("""
        <div style="display:flex;align-items:center;gap:14px;padding:6px 0 16px">
          <div style="width:42px;height:42px;flex-shrink:0;background:linear-gradient(145deg,#F29F05,#d78904);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:20px;box-shadow:0 4px 14px rgba(242,159,5,.4)">📄</div>
          <div>
            <div style="font-size:22px;font-weight:800;line-height:1.2;font-family:'Space Grotesk', sans-serif;">Folha de Rosto <span style="font-weight:400;opacity:0.7;font-size:18px;margin-left:6px">EFD-Reinf</span></div>
            <div style="font-size:11px;opacity:0.6;margin-top:2px">Automação de Documentos &nbsp;·&nbsp; Fisco Federal</div>
          </div>
        </div>""", unsafe_allow_html=True)
    with mid:
        st.markdown("<br>", unsafe_allow_html=True)
        st.session_state["dark_mode"] = st.toggle("🌙 Modo Escuro", value=st.session_state["dark_mode"])
    with right:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("↩ Sair", key="logout_btn", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()

def render_app():
    render_header()
    comp_folha, venc_str_padrao = get_datas_padrao()
    tab1, tab2, tab3 = st.tabs(["📝 1. Lançador de Notas (Nuvem/IA)", "⚙️ 2. Gerador Oficial (Word/PDF)", "🏢 3. Gestão de Clientes"])
    
    clientes_bd = carregar_clientes()
    lancamentos_bd = carregar_lancamentos()
    
    with tab1:
        st.markdown("<br><h4 style='margin-bottom:15px;'>🤖 Importação Inteligente (IA Vision)</h4>", unsafe_allow_html=True)
        chave_gemini = st.secrets.get("GEMINI_API_KEY", None)
        
        with st.expander("📂 Importar Notas Fiscais (Fotos ou PDFs)", expanded=False):
            arquivos_ia = st.file_uploader("Arraste fotos/PDFs de notas aqui", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
            if st.button("✨ Extrair dados com IA", type="primary") and arquivos_ia and chave_gemini:
                novos = []
                barra = st.progress(0)
                for i, arq in enumerate(arquivos_ia):
                    st.toast(f"Lendo: {arq.name}...", icon='👁️')
                    res = extrair_dados_ia_gemini(arq, chave_gemini)
                    if res: novos.append(res)
                    barra.progress((i + 1) / len(arquivos_ia))
                if novos: 
                    st.session_state["ia_dados_importados"] = novos
                    st.success(f"🎉 {len(novos)} notas extraídas com sucesso! Revise os dados na tabela abaixo.")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("#### 📋 Tabela de Conferência")
        
        c1, c2 = st.columns(2)
        with c1: cliente_t1 = st.selectbox("Selecione o Cliente", list(clientes_bd.keys()))
        with c2: comp_t1 = st.text_input("Competência", value=comp_folha)

        dados_existentes = lancamentos_bd.get(cliente_t1, {}).get(comp_t1, [])
        dados_tabela = dados_existentes + st.session_state.get("ia_dados_importados", [])
        
        cols = ["Órgão", "CNPJ Tomador", "Nº NF", "CNPJ Prestador", "Total Contrib. Prev.", "Compensação"]
        
        if dados_tabela:
            df_base = pd.DataFrame(dados_tabela)
            for c in cols:
                if c not in df_base.columns: df_base[c] = None
            df_base = df_base[cols]
        else:
            df_base = pd.DataFrame(columns=cols)
            for _ in range(5): df_base.loc[len(df_base)] = [None]*6
            
        df_base = df_base.reset_index(drop=True)
        
        # 🟢 CONTROLE DE LARGURA DAS COLUNAS WEB 🟢
        col_config = {
            "Órgão": st.column_config.TextColumn("Órgão", width=80),
            "CNPJ Tomador": st.column_config.TextColumn("CNPJ Tomador", width=180),
            "Nº NF": st.column_config.TextColumn("Nº NF", width=80),
            "CNPJ Prestador": st.column_config.TextColumn("CNPJ Prestador", width=180),
            "Total Contrib. Prev.": st.column_config.TextColumn("Total Contrib. Prev.", width=130),
            "Compensação": st.column_config.TextColumn("Compensação", width=100)
        }

        df_editado = st.data_editor(
            df_base, 
            num_rows="dynamic", 
            use_container_width=True, 
            hide_index=True, 
            column_config=col_config
        )

        st.markdown("<br><div style='background:rgba(42, 156, 107, 0.1); border: 1px solid rgba(42, 156, 107, 0.3); padding: 15px; border-radius: 10px;'>", unsafe_allow_html=True)
        st.markdown("##### 💾 Opções de Salvamento na Nuvem")
        
        colA, colB = st.columns([1.5, 1])
        with colA:
            modo_salvar = st.radio("Como deseja salvar os dados para esta competência?", 
                                   ["🔄 Substituir todos os dados que já estão na nuvem", 
                                    "➕ Adicionar estas notas aos dados que já estão lá"], 
                                   horizontal=False)
                                   
            if st.button("Salvar Tabela no Sistema", type="primary", use_container_width=True):
                df_limpo = df_editado.dropna(how="all").where(pd.notnull(df_editado), None)
                
                # 🟢 MÁSCARA AUTOMÁTICA DE CNPJ ANTES DE SALVAR 🟢
                if "CNPJ Tomador" in df_limpo.columns:
                    df_limpo['CNPJ Tomador'] = df_limpo['CNPJ Tomador'].apply(formatar_cnpj)
                if "CNPJ Prestador" in df_limpo.columns:
                    df_limpo['CNPJ Prestador'] = df_limpo['CNPJ Prestador'].apply(formatar_cnpj)
                
                modo = "sobrepor" if "Substituir" in modo_salvar else "adicionar"
                salvar_lancamentos(cliente_t1, comp_t1, df_limpo.to_dict("records"), modo)
                st.session_state["ia_dados_importados"] = []
                
                msg = "A nuvem foi atualizada!" if modo == "sobrepor" else "Novas notas foram adicionadas à nuvem!"
                st.success(f"✅ Sucesso! {msg} Os dados do cliente {cliente_t1} estão protegidos.")
                st.rerun() 
                
        with colB:
            st.markdown("<br>", unsafe_allow_html=True)
            df_export = df_editado.dropna(how="all")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer: df_export.to_excel(writer, sheet_name="Valores", index=False)
            st.download_button("📥 Baixar Planilha Manual (.xlsx)", data=output.getvalue(), file_name="Lançamentos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with tab2:
        st.markdown("<br>", unsafe_allow_html=True)
        colL, colR = st.columns([1, 1], gap="large")
        with colL:
            st.markdown("#### ⚙️ Configurações do Ato")
            cliente_sel = st.selectbox("Selecione o Cliente (Gerador)", list(clientes_bd.keys()))
            c_ato1, c_ato2 = st.columns([2, 1])
            with c_ato1: num_ato_int = st.number_input("Nº Inicial", min_value=1, value=1, step=1)
            with c_ato2: ano_ato = st.text_input("Ano", value=str(datetime.now().year))
            num_ato = f"{num_ato_int:03d}/{ano_ato}"
            resp_sel = st.selectbox("Responsável", list(RESPONSAVEIS.keys()))
            competencia = st.text_input("Competência (Gerador)", value=comp_folha)
            vencimento = st.text_input("Vencimento", value="20/03/2026")
            tipo_darf = st.radio("Tipo de Documento", ["Reinf", "Avulso", "Sem DARF"], horizontal=True)
            
        with colR:
            st.markdown("#### 📤 Fechamento e Exportação")
            houve_retencao = st.checkbox("✅ Houve retenções a declarar?", value=True)
            
            can_run = True
            dados_nfs = []
            if houve_retencao:
                fonte_dados = st.radio("Fonte dos Dados:", ["☁️ Nuvem (Lançamentos Salvos da IA/Manual)", "📂 Upload Antigo (.xlsx)"], horizontal=True)
                
                if "Nuvem" in fonte_dados:
                    dados_nfs = lancamentos_bd.get(cliente_sel, {}).get(competencia, [])
                    if dados_nfs:
                        st.success(f"✅ {len(dados_nfs)} notas carregadas automaticamente da nuvem.")
                    else:
                        st.warning("⚠️ Nenhum lançamento encontrado na nuvem para este cliente. Volte na aba 1 para salvar.")
                        can_run = False
                else:
                    arq_excel = st.file_uploader("Upload da Planilha Excel (.xlsx)", type=["xlsx"])
                    can_run = bool(arq_excel)
                    if can_run:
                        wb = load_workbook(io.BytesIO(arq_excel.getvalue()), data_only=True)
                        aba = [n for n in wb.sheetnames if n.lower().startswith("valores")][0]
                        ws = wb[aba]
                        headers = [str(c.value) for c in ws[1]]
                        dados_nfs = [dict(zip(headers, r)) for r in ws.iter_rows(min_row=2, values_only=True) if not all(c is None for c in r)]

            if st.button("Gerar PDF e Word Oficiais", type="primary", use_container_width=True, disabled=not can_run):
                with st.spinner("Construindo documento com larguras fixas e paleta verde..."):
                    uf = clientes_bd[cliente_sel].get("UF", "")
                    contexto = {
                        '{{numero_ato}}': num_ato, '{{data_emissao}}': datetime.now().strftime('%d/%m/%Y'),
                        '{{municipio_uf}}': f"{cliente_sel} / {uf}" if uf else cliente_sel,
                        '{{competencia}}': competencia, '{{vencimento}}': vencimento,
                        '{{responsavel}}': resp_sel, '{{ramal}}': RESPONSAVEIS[resp_sel],
                        '{{check_reinf}}': "☒" if tipo_darf == "Reinf" else "☐",
                        '{{check_avulso}}': "☒" if tipo_darf == "Avulso" else "☐"
                    }
                    
                    try:
                        with open("Modelo_Folha_Rosto.docx", "rb") as f: doc = Document(io.BytesIO(f.read()))
                        for p in doc.paragraphs:
                            for k, v in contexto.items():
                                if k in p.text:
                                    for r in p.runs:
                                        if k in r.text: r.text = r.text.replace(k, v)
                                    if k in p.text: p.text = p.text.replace(k, v)
                        for t in doc.tables:
                            for r in t.rows:
                                for c in r.cells:
                                    for p in c.paragraphs:
                                        for k, v in contexto.items():
                                            if k in p.text: p.text = p.text.replace(k, v)
                        
                        tabela = criar_tabela_reinf(doc, dados_nfs)
                        target = next((p for p in doc.paragraphs if "{{TABELA_NOTAS}}" in p.text), None)
                        if target: target._p.addnext(tabela._tbl); target.text = ""
                        
                        buf_docx = io.BytesIO(); doc.save(buf_docx); bytes_docx = buf_docx.getvalue()
                        st.toast('Documento compilado com sucesso!', icon='🎉')
                        
                        dl1, dl2 = st.columns(2)
                        with dl1: st.download_button("📥 Baixar WORD", data=bytes_docx, file_name=f"Folha - {cliente_sel}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                        with st.spinner("Convertendo para PDF..."):
                            with tempfile.TemporaryDirectory() as tmpdir:
                                dp = os.path.join(tmpdir, "t.docx"); pp = os.path.join(tmpdir, "t.pdf")
                                with open(dp, "wb") as f: f.write(bytes_docx)
                                try:
                                    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, dp], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                                    with open(pp, "rb") as f: 
                                        with dl2: st.download_button("📥 Baixar PDF", data=f.read(), file_name=f"Folha - {cliente_sel}.pdf", mime="application/pdf", use_container_width=True)
                                except Exception:
                                    with dl2: st.button("🚫 PDF Indisponível (Instale LibreOffice)", disabled=True, use_container_width=True)
                    except Exception as e: st.error(f"❌ Erro crítico no Word: {e}")

    with tab3:
        st.markdown("<br><div class='custom-card'><h4>🏢 Novo Cliente</h4>", unsafe_allow_html=True)
        with st.form("form_novo_cliente", clear_on_submit=True):
            cc1, cc2, cc3 = st.columns([2, 1, 1])
            with cc1: novo_nome = st.text_input("Nome")
            with cc2: nova_uf = st.text_input("UF", max_chars=2)
            with cc3: novo_cnpj = st.text_input("CNPJ")
            if st.form_submit_button("Salvar Base de Dados", type="primary", use_container_width=True):
                if novo_nome and nova_uf and novo_cnpj:
                    salvar_novo_cliente(novo_nome, nova_uf, novo_cnpj); st.rerun()

if __name__ == "__main__":
    if not st.session_state["authenticated"]: render_login()
    else: render_app()
