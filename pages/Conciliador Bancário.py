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
import fitz

# --- CONFIGURAÇÃO DA PÁGINA ---
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
try:
    icon_image = Image.open(icon_path)
    st.set_page_config(page_title="Portal Financeiro", page_icon=icon_image, layout="wide")
except:
    st.set_page_config(page_title="Portal Financeiro", layout="wide")

# --- CSS PERSONALIZADO ---
st.markdown("""
<style>
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
    div.stButton > button {
        background-color: rgb(38, 39, 48) !important;
        color: white !important;
        font-weight: bold !important;
        border: 1px solid rgb(60, 60, 60);
        border-radius: 5px;
        font-size: 16px;
        transition: 0.3s;
    }
    div.stButton > button:hover { background-color: rgb(20, 20, 25) !important; border-color: white; }
    .big-label { font-size: 24px !important; font-weight: 600 !important; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE PROCESSAMENTO
# ==============================================================================
CURRENT_YEAR = str(datetime.datetime.now().year)

def limpar_documento_pdf(doc_str):
    if not doc_str: return ""
    apenas_digitos = re.sub(r'\D', '', str(doc_str))
    if not apenas_digitos: return ""
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
                words = page.extract_words(x_tolerance=2, y_tolerance=2)
                linhas_dict = {}
                for w in words:
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
                        valor_bruto = match_valor.group(1)
                        tipo = match_valor.group(2)
                        valor_float = float(valor_bruto.replace('.', '').replace(',', '.'))
                        
                        coord_box = None
                        for w in linha_words:
                            if valor_bruto in w['text']:
                                coord_box = (page_idx, w['x0'], w['top'], w['x1'], w['bottom'])
                                break

                        texto_sem_data = texto_linha.replace(match_data.group(0), "", 1).strip()
                        texto_sem_valor = texto_sem_data.replace(match_valor.group(0), "").strip()
                        
                        entry = {
                            "Data": data_str, "Histórico": texto_sem_valor.strip(),
                            "Documento": "", "Valor_Extrato": valor_float, "coords": coord_box
                        }

                        if tipo == 'D':
                            tokens = texto_sem_valor.split()
                            if tokens:
                                for t in reversed(tokens):
                                    limpo = t.replace('.', '').replace('-', '')
                                    if limpo.isdigit() and len(limpo) >= 4:
                                        entry["Documento"] = limpar_documento_pdf(t)
                                        break
                            rows_debitos.append(entry)
                        elif tipo == 'C':
                            hist_upper = texto_sem_valor.upper()
                            if any(x in hist_upper for x in ["TED DEVOLVIDA", "DEVOLUCAO DE TED", "TED DEVOL"]):
                                rows_devolucoes.append(entry)
                                
        df_debitos = pd.DataFrame(rows_debitos)
        coords_referencia = rows_debitos + rows_devolucoes
        
    except:
        return pd.DataFrame(), []

    if not rows_devolucoes == [] and not df_debitos.empty:
        idx_rem = []
        for r_dev in rows_devolucoes:
            m = df_debitos[(df_debitos['Data'] == r_dev['Data']) & (abs(df_debitos['Valor_Extrato'] - r_dev['Valor_Extrato']) < 0.01) & (~df_debitos.index.isin(idx_rem))]
            if not m.empty: idx_rem.append(m.index[0])
        df_debitos = df_debitos.drop(idx_rem).reset_index(drop=True)

    termos_excluir = "SALDO|S A L D O|Resgate|BB-APLIC C\.PRZ-APL\.AUT|1\.972"
    df = df_debitos[~df_debitos['Histórico'].astype(str).str.contains(termos_excluir, case=False, na=False)].copy()
    
    df['Data_dt'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    mask_13113 = df['Histórico'].astype(str).str.contains("13113", na=False)
    if any(mask_13113):
        df_t = df[mask_13113].copy(); df_o = df[~mask_13113].copy()
        df_t_agg = df_t.groupby('Data_dt').agg({'Valor_Extrato': 'sum', 'Data': 'first'}).reset_index()
        df_t_agg['Documento'] = "Tarifas Bancárias"; df_t_agg['Histórico'] = "Tarifas Bancárias do Dia"
        df = pd.concat([df_o, df_t_agg], ignore_index=True)
    
    return df, coords_referencia

def processar_excel_detalhado(file_bytes, df_pdf_ref, is_csv=False):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='latin1', sep=None, engine='python') if is_csv else pd.read_excel(io.BytesIO(file_bytes), header=None)
        
        try: df = df.iloc[:, [4, 5, 8, 25, 26, 27]].copy()
        except: df = df.iloc[:, [4, 5, 8, -4, -2, -1]].copy()
        
        df.columns = ['Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
        df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x, str) else float(x))
        
        mask_pagto = df['Info_Z'].astype(str).str.contains("Pagamento", case=False, na=False)
        mask_transf_std = (df['Info_Z'].astype(str).str.contains("TRANSFERENCIA ENTRE CONTAS DE MESMA UG", case=False, na=False)) & (df['DC'].str.strip().str.upper() == 'C')
        mask_codes_z = df['Info_Z'].astype(str).str.contains(r"266|264|268", case=False, regex=True, na=False)
        cond_250_z = df['Info_Z'].astype(str).str.contains("250", case=False, na=False)
        cond_ab_text = df['Info_AB'].astype(str).str.contains("transferência financeira concedida|repasse financeiro concedido", case=False, na=False)
        mask_250_restrict = cond_250_z & cond_ab_text
        mask_aa_ded = df['Info_AA'].astype(str).str.contains(r"Ded\.", case=False, regex=True, na=False)
        
        df_filtered = df[mask_pagto | mask_transf_std | mask_codes_z | mask_250_restrict | mask_aa_ded].copy()
        
        df_final = df_filtered.copy()
        df_final['Data_dt'] = df_final['Data'].apply(parse_br_date)
        df_final = df_final.dropna(subset=['Data_dt'])
        df_final['Data'] = df_final['Data_dt'].dt.strftime('%d/%m/%Y')
        
        lookup = {dt: {str(doc).lstrip('0'): doc for doc in g['Documento'].unique()} for dt, g in df_pdf_ref.groupby('Data')}
        lookup_ded = {}
        if not df_pdf_ref.empty:
            mask_pdf_ded = df_pdf_ref['Histórico'].astype(str).str.contains(r"Dedução|Ded\.|FUNDEB|PASEP", case=False, regex=True, na=False)
            for idx, row in df_pdf_ref[mask_pdf_ded].iterrows():
                if row['Data'] not in lookup_ded: lookup_ded[row['Data']] = row['Documento']
        
        def find_doc(row):
            txt, dt, info_aa = str(row['Info_AB']).upper(), row['Data'], str(row['Info_AA']).upper()
            if "DED." in info_aa and dt in lookup_ded: return lookup_ded[dt]
            if dt not in lookup: return "S/D"
            if "TARIFA" in txt and "Tarifas Bancárias" in lookup[dt].values(): return "Tarifas Bancárias"
            for n in re.findall(r'\d+', txt):
                if n.lstrip('0') in lookup[dt]: return lookup[dt][n.lstrip('0')]
            return "NÃO LOCALIZADO"
            
        df_final['Documento'] = df_final.apply(find_doc, axis=1)
        df_final['Desc_AA'] = df_final['Info_AA'] # Mantém para identificar FUNDEB
        return df_final.reset_index(drop=True)[['Data', 'Documento', 'Valor_Razao', 'Desc_AA']]
    except: return pd.DataFrame()

def executar_conciliacao_inteligente(df_pdf, df_excel):
    """
    LOGICA CORRIGIDA:
    1. MATCH DIRETO (1-para-1) para PASEP e outros.
    2. AGRUPAMENTO EXCLUSIVO FUNDEB (Soma Excel vs Soma PDF para 'Ded. FUNDEB').
    """
    res = []
    idx_p_u = set()
    idx_e_u = set()

    # --- PASSO 1: MATCH EXATO (PASEP e outros itens individuais) ---
    for idx_p, row_p in df_pdf.iterrows():
        # Ignora FUNDEB neste passo para não casar errado
        if "FUNDEB" in str(row_p['Histórico']).upper(): continue
        
        cand = df_excel[
            (df_excel['Data'] == row_p['Data']) & 
            (df_excel['Documento'] == row_p['Documento']) & 
            (~df_excel.index.isin(idx_e_u))
        ]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'], 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': m.loc[idx_e]['Valor_Razao'], 'Diferença': 0.0})
            idx_p_u.add(idx_p); idx_e_u.add(idx_e)

    # --- PASSO 2: AGRUPAMENTO ESPECIAL FUNDEB ---
    # Pegamos tudo o que é FUNDEB em ambos os lados
    df_p_fundeb = df_pdf[df_pdf['Histórico'].astype(str).str.contains("FUNDEB", case=False, na=False) & (~df_pdf.index.isin(idx_p_u))]
    df_e_fundeb = df_excel[df_excel['Desc_AA'].astype(str).str.contains("FUNDEB", case=False, na=False) & (~df_excel.index.isin(idx_e_u))]

    if not df_p_fundeb.empty:
        # Agrupamos por Data e Documento
        for (dt, doc), group_p in df_p_fundeb.groupby(['Data', 'Documento']):
            val_p_total = group_p['Valor_Extrato'].sum()
            
            group_e = df_e_fundeb[(df_e_fundeb['Data'] == dt) & (df_e_fundeb['Documento'] == doc)]
            val_e_total = group_e['Valor_Razao'].sum()
            
            # Reconciliamos o bloco de FUNDEB do dia
            res.append({
                'Data': dt, 
                'Histórico': f"Agrupado: {group_p['Histórico'].iloc[0]}", 
                'Documento': doc, 
                'Valor_Extrato': val_p_total, 
                'Valor_Razao': val_e_total, 
                'Diferença': val_p_total - val_e_total
            })
            idx_p_u.update(group_p.index)
            idx_e_u.update(group_e.index)

    # --- PASSO 3: MATCH FLEXÍVEL (Docs Dif) ---
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p in idx_p_u: continue
        cand = df_excel[(df_excel['Data'] == row_p['Data']) & (~df_excel.index.isin(idx_e_u))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': "Docs dif.", 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': m.loc[idx_e]['Valor_Razao'], 'Diferença': 0.0})
            idx_p_u.add(idx_p); idx_e_u.add(idx_e)

    # --- PASSO 4: SOBRAS (O que não conciliou) ---
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p not in idx_p_u:
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'], 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': 0.0, 'Diferença': row_p['Valor_Extrato']})
    
    excel_sobra = df_excel[~df_excel.index.isin(idx_e_u)]
    if not excel_sobra.empty:
        for _, row_e in excel_sobra.groupby(['Data', 'Documento'])['Valor_Razao'].sum().reset_index().iterrows():
            res.append({'Data': row_e['Data'], 'Histórico': "Não Conciliado (Razão)", 'Documento': row_e['Documento'], 'Valor_Extrato': 0.0, 'Valor_Razao': row_e['Valor_Razao'], 'Diferença': -row_e['Valor_Razao']})

    df_f = pd.DataFrame(res)
    df_f['dt'] = pd.to_datetime(df_f['Data'], format='%d/%m/%Y', errors='coerce')
    return df_f.sort_values(by=['dt', 'Documento']).drop(columns=['dt'])

# ==============================================================================
# 2. GERAÇÃO DE SAÍDAS (Mantido original)
# ==============================================================================
def gerar_pdf_final(df_f, titulo_completo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm, title=titulo_completo)
    story, styles = [], getSampleStyleSheet()
    story.append(Paragraph("Relatório de Conciliação Bancária", styles["Title"]))
    story.append(Spacer(1, 15))
    data = [['Data', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']]
    for _, r in df_f.iterrows():
        data.append([r['Data'], str(r['Documento']), formatar_moeda_br(r['Valor_Extrato']), formatar_moeda_br(r['Valor_Razao']), formatar_moeda_br(r['Diferença'])])
    t = Table(data, colWidths=[25*mm, 65*mm, 33*mm, 33*mm, 33*mm])
    t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.black), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('ALIGN', (2,0), (-1,-1), 'RIGHT')]))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

def gerar_excel_final(df_f):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_f.to_excel(writer, sheet_name='Conciliacao', index=False)
    return output.getvalue()

# ==============================================================================
# 3. INTERFACE
# ==============================================================================
st.markdown("<h1 style='text-align: center;'>Portal Financeiro - Conciliação</h1>", unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1: up_pdf = st.file_uploader("Extrato PDF", type="pdf")
with c2: up_xlsx = st.file_uploader("Razão Excel", type=["xlsx", "csv"])

if st.button("PROCESSAR"):
    if up_pdf and up_xlsx:
        df_p, _ = processar_pdf(up_pdf.read())
        df_e = processar_excel_detalhado(up_xlsx.read(), df_p, is_csv=up_xlsx.name.endswith('csv'))
        df_f = executar_conciliacao_inteligente(df_p, df_e)
        st.dataframe(df_f, use_container_width=True)
        st.download_button("Baixar PDF", gerar_pdf_final(df_f, "Conciliação"), "Relatorio.pdf")
        st.download_button("Baixar Excel", gerar_excel_final(df_f), "Relatorio.xlsx")
