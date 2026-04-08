"""
ConPrev — Gerador EFD-Reinf  ·  SaaS Premium (v10.5)
=============================================================
Ajuste de Largura de Colunas, Máscara Automática de CNPJ Implacável,
Cálculo de PDF blindado, Coluna Index Oculta e UX Glassmorphism.
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
from docx.shared import Pt, RGBColor
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
    "Município - Uirapuru": {"UF": "GO", "CNPJ": "37.622.164/0001-60"}
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
    html, body, p, span, div, label, li {{ font-family: 'Inter', sans-serif !important; color: {text_color}; }}
    h1, h2, h3, h4, h5, h6 {{ font-family: 'Space Grotesk', sans-serif !important; color: {heading_color} !important; }}
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
    
    # Se o Pandas leu o CNPJ colado como um float (ex: 11743228000198.0)
    if cnpj_str.endswith('.0'): cnpj_str = cnpj_str[:-2]
    
    digits = re.sub(r'\D', '', cnpj_str)
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    return str(cnpj_val) # Retorna original se for incompleto

# ── Cérebro de IA (Gemini Vision Otimizado contra Limites 429) ────────────────
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
        model = genai.GenerativeModel('gemini-1.5-flash') 
        
        response = model.generate_content([prompt, sample_file])
        genai.delete_file(sample_file.name)
        os.remove(tmp_path)
        
        txt_limpo = response.text.replace('```json', '').replace('```', '').strip()
        dados = json.loads(txt_limpo)
        
        dados["CNPJ Tomador"] = formatar_cnpj(dados.get("CNPJ Tomador", ""))
        dados["CNPJ Prestador"] = formatar_cnpj(dados.get("CNPJ Prestador", ""))
        
        v_inss = safe_float(dados.get("Total Contrib. Prev.", 0.0))
        dados["Total Contrib. Prev."] = f"R$ {v_inss:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if v_inss > 0 else ""
        dados["Compensação"] = ""
        
        return dados
    except Exception as e:
        st.error(f"Erro de IA (Possível Limite de Quota): Aguarde 1 minuto e tente novamente.")
        return None

def get_datas_padrao() -> Tuple[str, str, str, datetime]:
    hoje = datetime.now()
    primeiro_dia = hoje.replace(day=1)
    ultimo_dia_mes_ant = primeiro_dia - timedelta(days=1)
    comp_folha = f"{ultimo_dia_mes_ant.strftime('%m/%Y')}"
    return comp_folha, datetime(hoje.year, hoje.month, 20).strftime("%d/%m/%Y")

def set_cell_background(cell, fill_color: str):
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill_color)
    cell._tc.get_or_add_tcPr().append(shd)

# ── Gerador do Word/PDF com Agrupamento e Cálculo Perfeito ────────────────────
def criar_tabela_reinf(doc: Document, dados_nfs: List[Dict[str, Any]]) -> Any:
    headers = ['Órgão', 'CNPJ Tomador', 'Nº NF', 'CNPJ Prestador', 'Total Contrib. Prev.', 'Compensação']
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'; table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]; set_cell_background(cell, "D9D9D9")
        p = cell.paragraphs[0]; r = p.add_run(h); r.font.bold = True; r.font.size = Pt(10); p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not dados_nfs:
        row = table.add_row().cells
        row[0].text = "Nenhuma retenção de INSS declarada na EFD-REINF"; row[0].merge(row[5]); row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return table

    grupos = defaultdict(lambda: defaultdict(list))
    for nf in dados_nfs: grupos[str(nf.get('Órgão', ''))][str(nf.get('CNPJ Prestador', ''))].append(nf)

    total_geral_contrib = total_geral_comp = 0.0

    for orgao, prestadores in grupos.items():
        sub_org_contrib = sub_org_comp = 0.0
        for prestador, nfs in prestadores.items():
            sub_prest_contrib = sub_prest_comp = 0.0
            for nf in nfs:
                row = table.add_row().cells
                row[0].text = orgao; row[1].text = str(nf.get('CNPJ Tomador', '')); row[2].text = str(nf.get('Nº NF', ''))
                row[3].text = prestador
                
                v_c = safe_float(nf.get('Total Contrib. Prev.', 0))
                v_cp = safe_float(nf.get('Compensação', 0))
                
                row[4].text = _brl_fmt(v_c)
                row[5].text = _brl_fmt(v_cp)
                
                sub_prest_contrib += v_c; sub_prest_comp += v_cp
                for c in row: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            st_row = table.add_row().cells; set_cell_background(st_row[0], "FDFDFD")
            st_row[0].text = f"Subtotal - CNPJ {prestador}"; st_row[0].merge(st_row[3]); st_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            st_row[4].text = _brl_fmt(sub_prest_contrib); st_row[5].text = _brl_fmt(sub_prest_comp)
            for idx in [0,4,5]:
                for r in st_row[idx].paragraphs[0].runs: r.bold = True
                if idx != 0: st_row[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            sub_org_contrib += sub_prest_contrib; sub_org_comp += sub_prest_comp

        org_row = table.add_row().cells; set_cell_background(org_row[0], "F2F2F2")
        org_row[0].text = f"Subtotal do Órgão ({orgao})"; org_row[0].merge(org_row[3]); org_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        org_row[4].text = _brl_fmt(sub_org_contrib); org_row[5].text = _brl_fmt(sub_org_comp)
        for idx in [0,4,5]:
            for r in org_row[idx].paragraphs[0].runs: r.bold = True
            if idx != 0: org_row[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        total_geral_contrib += sub_org_contrib; total_geral_comp += sub_org_comp

    t_row = table.add_row().cells; set_cell_background(t_row[0], "EAEAEA")
    t_row[0].text = "TOTAL GERAL"; t_row[0].merge(t_row[3]); t_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    t_row[4].text = _brl_fmt(total_geral_contrib); t_row[5].text = _brl_fmt(total_geral_comp)
    for idx in [0,4,5]:
        for r in t_row[idx].paragraphs[0].runs: r.bold = True
        if idx != 0: t_row[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
    return table

def render_app():
    comp_folha, venc_str = get_datas_padrao()
    clientes_bd = carregar_clientes()
    lancamentos_bd = carregar_lancamentos()

    st.markdown("""
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 20px;">
        <h2 style="margin:0;">📄 EFD-Reinf <span style="font-weight:400; opacity:0.6; font-size:18px;">Automação Fiscal</span></h2>
    </div>
    """, unsafe_allow_html=True)
    
    st.session_state["dark_mode"] = st.toggle("🌙 Modo Escuro", value=st.session_state["dark_mode"])
    
    tab1, tab2, tab3 = st.tabs(["📝 1. Lançador de Notas (Nuvem/IA)", "⚙️ 2. Gerador Oficial (Word/PDF)", "🏢 3. Gestão de Clientes"])

    with tab1:
        st.markdown("<br><h4 style='margin-bottom:15px;'>🤖 Importação Inteligente (IA Vision)</h4>", unsafe_allow_html=True)
        chave_gemini = st.secrets.get("GEMINI_API_KEY", None)
        
        with st.expander("Clique aqui para enviar fotos ou PDFs das notas fiscais", expanded=False):
            arquivos_ia = st.file_uploader("Arraste fotos/PDFs de notas aqui", type=["pdf", "png", "jpg"], accept_multiple_files=True)
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
            
        # Limpa índices de importações anteriores
        df_base = df_base.reset_index(drop=True)
        
        # 🟢 CONTROLE DE LARGURA DAS COLUNAS 🟢
        col_config = {
            "Órgão": st.column_config.TextColumn("Órgão", width="small"),
            "CNPJ Tomador": st.column_config.TextColumn("CNPJ Tomador", width="large"),
            "Nº NF": st.column_config.TextColumn("Nº NF", width="small"),
            "CNPJ Prestador": st.column_config.TextColumn("CNPJ Prestador", width="large"),
            "Total Contrib. Prev.": st.column_config.TextColumn("Total Contrib. Prev.", width="medium"),
            "Compensação": st.column_config.TextColumn("Compensação", width="small")
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
                st.rerun() # Atualiza a tabela na tela imediatamente com os CNPJs formatados
                
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
                dados_nfs = lancamentos_bd.get(cliente_sel, {}).get(competencia, [])
                if dados_nfs:
                    st.success(f"✅ {len(dados_nfs)} notas carregadas automaticamente da nuvem para o documento final.")
                else:
                    st.warning("⚠️ Nenhum lançamento encontrado na nuvem para este cliente e competência. Volte na aba 1 para salvar.")
                    can_run = False

            if st.button("Gerar PDF e Word Oficiais", type="primary", use_container_width=True, disabled=not can_run):
                with st.spinner("Construindo documento com subtotais..."):
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
                        st.toast('Documento compilado!', icon='🎉')
                        
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
                                    with dl2: st.button("🚫 PDF Indisponível", disabled=True, use_container_width=True)
                    except Exception as e: st.error(f"❌ Erro: {e}")

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
