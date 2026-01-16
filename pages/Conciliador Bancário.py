import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import datetime
import xlsxwriter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
from PIL import Image
import fitz  # PyMuPDF

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal Financeiro - Conciliação", layout="wide")

# --- CSS PARA MELHORAR O VISUAL ---
st.markdown("""
<style>
    .block-container { padding-top: 2rem !important; }
    div.stButton > button {
        background-color: #262730;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        height: 3em;
        width: 100%;
    }
    .big-label { font-size: 20px !important; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES AUXILIARES E DE PROCESSAMENTO
# ==============================================================================
CURRENT_YEAR = str(datetime.datetime.now().year)

def limpar_documento_pdf(doc_str):
    if not doc_str: return ""
    apenas_digitos = re.sub(r'\D', '', str(doc_str))
    return apenas_digitos[-6:] if len(apenas_digitos) > 6 else apenas_digitos

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == "-": return "-"
    return f"{valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

def parse_br_date(date_val):
    if pd.isna(date_val): return pd.NaT
    try:
        if isinstance(date_val, str): date_val = date_val.split()[0]
        return pd.to_datetime(date_val, dayfirst=True, errors='coerce')
    except: return pd.to_datetime(date_val, errors='coerce')

def processar_pdf(file_bytes):
    rows_debitos = []
    rows_devolucoes = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                linhas_dict = {}
                for w in page.extract_words(x_tolerance=2, y_tolerance=2):
                    top = round(w['top'], 1)
                    linhas_dict.setdefault(top, []).append(w)
                
                for top in sorted(linhas_dict.keys()):
                    linha_words = linhas_dict[top]
                    texto_linha = " ".join([w['text'] for w in linha_words])
                    match_data = re.search(r'^(\d{2}/\d{2}(?:/\d{4})?)', texto_linha)
                    if not match_data: continue 
                    
                    data_str = match_data.group(1)
                    if len(data_str) == 5: data_str = f"{data_str}/{CURRENT_YEAR}"
                    match_valor = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})\s?([DC])', texto_linha)
                    
                    if match_valor:
                        valor_float = float(match_valor.group(1).replace('.', '').replace(',', '.'))
                        tipo = match_valor.group(2)
                        texto_limpo = texto_linha.replace(match_data.group(0), "", 1).strip().replace(match_valor.group(0), "").strip()
                        
                        entry = {
                            "Data": data_str, 
                            "Histórico": texto_limpo, 
                            "Documento": "", 
                            "Valor_Extrato": valor_float, 
                            "coords": (page_idx, linha_words[0]['x0'], top, linha_words[-1]['x1'], top+10)
                        }

                        if tipo == 'D':
                            for t in reversed(texto_limpo.split()):
                                if t.replace('.', '').replace('-', '').isdigit() and len(t.replace('.', '')) >= 4:
                                    entry["Documento"] = limpar_documento_pdf(t)
                                    break
                            rows_debitos.append(entry)
                        elif tipo == 'C' and any(x in texto_limpo.upper() for x in ["DEVOL", "TED DEV"]):
                            rows_devolucoes.append(entry)
        
        df_debitos = pd.DataFrame(rows_debitos)
        df_debitos['Data_dt'] = pd.to_datetime(df_debitos['Data'], format='%d/%m/%Y', errors='coerce')
        
        # Agrupamento de Tarifas Bancárias (Mantido conforme solicitado)
        mask_t = df_debitos['Histórico'].str.contains("13113|TARIFA", case=False, na=False)
        if any(mask_t):
            df_t = df_debitos[mask_t].copy()
            df_o = df_debitos[~mask_t].copy()
            df_t_agg = df_t.groupby('Data_dt').agg({'Valor_Extrato': 'sum', 'Data': 'first'}).reset_index()
            df_t_agg['Documento'] = "Tarifas Bancárias"; df_t_agg['Histórico'] = "Tarifas Bancárias do Dia"
            df_debitos = pd.concat([df_o, df_t_agg], ignore_index=True)
            
        return df_debitos, rows_debitos + rows_devolucoes
    except: return pd.DataFrame(), []

def processar_excel_detalhado(file_bytes, df_pdf_ref, is_csv=False):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='latin1', sep=None, engine='python') if is_csv else pd.read_excel(io.BytesIO(file_bytes), header=None)
        # Ajuste de colunas baseado na estrutura GovBr padrão
        df = df.iloc[:, [4, 5, 8, 25, 26, 27]].copy()
        df.columns = ['Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
        df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x, str) else float(x))
        df['Data_dt'] = df['Data'].apply(parse_br_date); df = df.dropna(subset=['Data_dt'])
        df['Data'] = df['Data_dt'].dt.strftime('%d/%m/%Y')
        
        lookup = {dt: {str(doc).lstrip('0'): doc for doc in g['Documento'].unique()} for dt, g in df_pdf_ref.groupby('Data')}
        def find_doc(row):
            txt, dt = str(row['Info_AB']).upper(), row['Data']
            if dt in lookup:
                for n in re.findall(r'\d+', txt):
                    if n.lstrip('0') in lookup[dt]: return lookup[dt][n.lstrip('0')]
            return "NÃO LOCALIZADO"
            
        df['Documento'] = df.apply(find_doc, axis=1)
        df['Descricao_Excel'] = df['Info_AA']
        # Filtramos apenas os Créditos (C), que representam as saídas na contabilidade
        return df[df['DC'] == 'C'].reset_index(drop=True)
    except: return pd.DataFrame()

# ==============================================================================
# 2. CORE: LÓGICA DE CONCILIAÇÃO HÍBRIDA
# ==============================================================================
def executar_conciliacao_inteligente(df_pdf, df_excel):
    res = []; idx_p_u = set(); idx_e_u = set()
    
    # Etapa 1: Consolidação de Grupos (FUNDEB, PASEP, RETENÇÃO, DEDUÇÃO)
    # Estas categorias são somadas de ambos os lados para bater o total do dia
    categorias = ["FUNDEB", "PASEP", "RETEN", "DEDU"]
    for cat in categorias:
        mask_p = df_pdf['Histórico'].str.contains(cat, case=False, na=False)
        mask_e = df_excel['Descricao_Excel'].str.contains(cat, case=False, na=False)
        
        for (dt, doc), group_p in df_pdf[mask_p].groupby(['Data', 'Documento']):
            idxs_p = group_p.index
            # No Excel, pegamos tudo daquela categoria no mesmo dia
            idxs_e = df_excel[mask_e & (df_excel['Data'] == dt)].index
            idxs_e = [i for i in idxs_e if i not in idx_e_u]
            
            soma_p = df_pdf.loc[idxs_p, 'Valor_Extrato'].sum()
            soma_e = df_excel.loc[idxs_e, 'Valor_Razao'].sum() if idxs_e else 0.0
            
            if soma_p > 0:
                res.append({
                    'Data': dt, 
                    'Histórico': f"Consolidado {cat.replace('RETEN', 'Retenções')}", 
                    'Documento': doc, 
                    'Valor_Extrato': soma_p, 
                    'Valor_Razao': soma_e, 
                    'Diferença': round(soma_p - soma_e, 2)
                })
                idx_p_u.update(idxs_p)
                idx_e_u.update(idxs_e)

    # Etapa 2: Match Individual (Saúde, Transferências, etc)
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p in idx_p_u: continue
        
        # Procura por valor exato no mesmo dia
        match = df_excel[(df_excel['Data'] == row_p['Data']) & 
                         (abs(df_excel['Valor_Razao'] - row_p['Valor_Extrato']) < 0.05) & 
                         (~df_excel.index.isin(idx_e_u))]
        
        if not match.empty:
            idx_e = match.index[0]
            res.append({
                'Data': row_p['Data'], 
                'Histórico': row_p['Histórico'], 
                'Documento': row_p['Documento'] if row_p['Documento'] else match.loc[idx_e, 'Documento'], 
                'Valor_Extrato': row_p['Valor_Extrato'], 
                'Valor_Razao': match.loc[idx_e, 'Valor_Razao'], 
                'Diferença': 0.0
            })
            idx_p_u.add(idx_p)
            idx_e_u.add(idx_e)

    # Etapa 3: Tratamento de Pendências (Sobras)
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p not in idx_p_u:
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'], 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': 0.0, 'Diferença': row_p['Valor_Extrato']})
    
    for idx_e, row_e in df_excel.iterrows():
        if idx_e not in idx_e_u:
            res.append({'Data': row_e['Data'], 'Histórico': "Pendente no Razão", 'Documento': row_e['Documento'], 'Valor_Extrato': 0.0, 'Valor_Razao': row_e['Valor_Razao'], 'Diferença': -row_e['Valor_Razao']})

    df_final = pd.DataFrame(res)
    if not df_final.empty:
        df_final['dt_temp'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y', errors='coerce')
        df_final = df_final.sort_values(by=['dt_temp', 'Diferença']).drop(columns=['dt_temp'])
    return df_final

# ==============================================================================
# 3. EXPORTAÇÃO E DOWNLOADS
# ==============================================================================
def gerar_pdf_relatorio(df_f, titulo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()
    elements.append(Paragraph(f"Relatório de Conciliação: {titulo}", styles['Title']))
    
    data = [['Data', 'Documento', 'Extrato', 'Razão', 'Dif.']]
    for _, r in df_f.iterrows():
        data.append([r['Data'], str(r['Documento']), formatar_moeda_br(r['Valor_Extrato']), formatar_moeda_br(r['Valor_Razao']), formatar_moeda_br(r['Diferença'])])
    
    t = Table(data, colWidths=[25*mm, 50*mm, 35*mm, 35*mm, 35*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (2,0), (-1,-1), 'RIGHT')
    ]))
    elements.append(t)
    doc.build(elements)
    return buffer.getvalue()

def gerar_excel_relatorio(df_f):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_f.to_excel(writer, index=False, sheet_name='Conciliação')
    return output.getvalue()

# ==============================================================================
# 4. INTERFACE STREAMLIT
# ==============================================================================
st.title("🏦 Conciliador Bancário Automático")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    up_pdf = st.file_uploader("Carregar Extrato PDF", type="pdf")
with col2:
    up_xlsx = st.file_uploader("Carregar Razão Excel/CSV", type=["xlsx", "csv"])

if st.button("EXECUTAR CONCILIAÇÃO"):
    if up_pdf and up_xlsx:
        with st.spinner("Processando dados..."):
            pdf_bytes = up_pdf.read()
            excel_bytes = up_xlsx.read()
            
            df_p, coords_raw = processar_pdf(pdf_bytes)
            df_e = processar_excel_detalhado(excel_bytes, df_p, is_csv=up_xlsx.name.endswith('csv'))
            
            if df_p.empty or df_e.empty:
                st.error("Não foi possível processar os ficheiros. Verifique os formatos.")
            else:
                df_f = executar_conciliacao_inteligente(df_p, df_e)
                
                # Exibição dos resultados
                st.subheader("Resultado da Conciliação")
                st.dataframe(df_f.style.format({'Valor_Extrato': '{:.2f}', 'Valor_Razao': '{:.2f}', 'Diferença': '{:.2f}'}), use_container_width=True)
                
                # Botões de Download
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("Baixar Relatório PDF", gerar_pdf_relatorio(df_f, up_pdf.name), "Relatorio.pdf", "application/pdf")
                with col2:
                    st.download_button("Baixar Relatório Excel", gerar_excel_relatorio(df_f), "Conciliacao.xlsx", "application/vnd.ms-excel")
    else:
        st.warning("Por favor, carregue ambos os ficheiros.")
