import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
from PIL import Image
import fitz  # Requer pymupdf no requirements.txt

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

def limpar_lancamento(val):
    """Remove o .0 e formata como string limpa"""
    try:
        if pd.isna(val): return ""
        return str(int(float(val)))
    except:
        return str(val)

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
                        entry = {"Data": data_str, "Histórico": texto_sem_valor.strip(), "Documento": "", "Valor_Extrato": valor_float, "coords": coord_box}
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
                            if any(x in texto_sem_valor.upper() for x in ["TED DEVOLVIDA", "DEVOLUCAO"]):
                                rows_devolucoes.append(entry)
        df_p = pd.DataFrame(rows_debitos)
        coords_ref = rows_debitos + rows_devolucoes
        
        termos_excluir = "SALDO|S A L D O|Resgate|BB-APLIC C\.PRZ-APL\.AUT|1\.972"
        df_p = df_p[~df_p['Histórico'].astype(str).str.contains(termos_excluir, case=False, na=False)].copy()
        
        df_p['Data_dt'] = pd.to_datetime(df_p['Data'], format='%d/%m/%Y', errors='coerce')
        mask_t = df_p['Histórico'].str.contains("13113", na=False)
        if any(mask_t):
            df_t = df_p[mask_t].copy(); df_o = df_p[~mask_t].copy()
            df_t_agg = df_t.groupby('Data_dt').agg({'Valor_Extrato': 'sum', 'Data': 'first'}).reset_index()
            df_t_agg['Documento'] = "Tarifas Bancárias"; df_t_agg['Histórico'] = "Tarifas Bancárias do Dia"
            df_p = pd.concat([df_o, df_t_agg], ignore_index=True)
        return df_p, coords_ref
    except: return pd.DataFrame(), []

def processar_excel_detalhado(file_bytes, df_pdf_ref, is_csv=False):
    try:
        df_raw = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='latin1', sep=None, engine='python') if is_csv else pd.read_excel(io.BytesIO(file_bytes), header=None)
        df = df_raw.iloc[:, [1, 4, 5, 8, 25, 26, 27]].copy()
        df.columns = ['Lancamento', 'Data', 'DC', 'Valor_Razao', 'LCP', 'Info_AA', 'Historico']
        
        mask_pagto = df['LCP'].astype(str).str.contains("Pagamento", case=False, na=False)
        mask_transf = (df['LCP'].astype(str).str.contains("TRANSFERENCIA ENTRE CONTAS DE MESMA UG", case=False, na=False)) & (df['DC'].str.strip().str.upper() == 'C')
        df = df[mask_pagto | mask_transf].copy()
        
        df['Data_dt'] = df['Data'].apply(parse_br_date); df = df.dropna(subset=['Data_dt'])
        df['Data_str'] = df['Data_dt'].dt.strftime('%d/%m/%Y')
        df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x, str) else float(x))
        
        lookup = {dt: {str(doc).lstrip('0'): doc for doc in g['Documento'].unique()} for dt, g in df_pdf_ref.groupby('Data')}
        def find_doc(row):
            txt, dt = str(row['Historico']).upper(), row['Data_str']
            if dt not in lookup: return "S/D"
            if "TARIFA" in txt and "Tarifas Bancárias" in lookup[dt].values(): return "Tarifas Bancárias"
            for n in re.findall(r'\d+', txt):
                if n.lstrip('0') in lookup[dt]: return lookup[dt][n.lstrip('0')]
            return "NÃO LOCALIZADO"
        
        df['Documento'] = df.apply(find_doc, axis=1)
        df['Lancamento_Limpo'] = df['Lancamento'].apply(limpar_lancamento)
        return df
    except: return pd.DataFrame()

def executar_conciliacao_inteligente(df_pdf, df_excel):
    res, idx_p_u, idx_e_u = [], set(), set()
    
    # 1. Match Exato (1-para-1)
    for idx_p, row_p in df_pdf.iterrows():
        cand = df_excel[(df_excel['Data_str'] == row_p['Data']) & (df_excel['Documento'] == row_p['Documento']) & (~df_excel.index.isin(idx_e_u))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'], 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': m.loc[idx_e]['Valor_Razao'], 'Diferença': 0.0})
            idx_p_u.add(idx_p); idx_e_u.add(idx_e)
            
    # 2. Match por Valor (Ignorando Doc se valor for único no dia)
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p in idx_p_u: continue
        cand = df_excel[(df_excel['Data_str'] == row_p['Data']) & (~df_excel.index.isin(idx_e_u))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': "Docs dif.", 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': m.loc[idx_e]['Valor_Razao'], 'Diferença': 0.0})
            idx_p_u.add(idx_p); idx_e_u.add(idx_e)

    # 3. Lógica de Agrupamento e identificação de divergências reais
    df_e_s = df_excel[~df_excel.index.isin(idx_e_u)].copy()
    df_p_s = df_pdf[~df_pdf.index.isin(idx_p_u)].copy()
    
    # Agrupamos sobras por Data e Documento
    g_e = df_e_s.groupby(['Data_str', 'Documento'])
    g_p = df_p_s.groupby(['Data', 'Documento'])
    
    docs_restantes = set(df_e_s[['Data_str', 'Documento']].apply(tuple, axis=1)) | set(df_p_s[['Data', 'Documento']].apply(tuple, axis=1))
    
    for dt, doc in docs_restantes:
        v_ext = df_p_s[(df_p_s['Data'] == dt) & (df_p_s['Documento'] == doc)]['Valor_Extrato'].sum()
        v_raz = df_e_s[(df_e_s['Data_str'] == dt) & (df_e_s['Documento'] == doc)]['Valor_Razao'].sum()
        
        diff = v_ext - v_raz
        hist = df_p_s[(df_p_s['Data'] == dt) & (df_p_s['Documento'] == doc)]['Histórico'].iloc[0] if v_ext > 0 else "S/H"
        
        res.append({'Data': dt, 'Histórico': hist, 'Documento': doc, 'Valor_Extrato': v_ext, 'Valor_Razao': v_raz, 'Diferença': diff})
        
        # Se a soma bateu perfeitamente, marcamos os lançamentos do Razão como 'usados'
        if abs(diff) < 0.01:
            indices_grupo = df_e_s[(df_e_s['Data_str'] == dt) & (df_e_s['Documento'] == doc)].index
            idx_e_u.update(indices_grupo)
    
    df_f = pd.DataFrame(res)
    df_f['dt'] = pd.to_datetime(df_f['Data'], format='%d/%m/%Y', errors='coerce')
    return df_f.sort_values(by=['dt', 'Documento']).drop(columns=['dt']), idx_e_u

# ==============================================================================
# 2. GERAÇÃO DE SAÍDAS (PDFs)
# ==============================================================================

def gerar_pdf_final(df_f, titulo_completo):
    buffer = io.BytesIO()
    # MODO PAISAGEM
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm, title=titulo_completo)
    story = []
    styles = getSampleStyleSheet()
    story.append(Paragraph("Relatório de Conciliação Bancária", styles["Title"]))
    story.append(Paragraph(f"<b>Conta:</b> {titulo_completo.replace('Conciliação ', '')}", ParagraphStyle(name='C', alignment=1)))
    story.append(Spacer(1, 15))
    
    headers = ['Data', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']
    data = [headers]
    for _, r in df_f.iterrows():
        diff = r['Diferença']
        data.append([r['Data'], str(r['Documento']), formatar_moeda_br(r['Valor_Extrato']), formatar_moeda_br(r['Valor_Razao']), formatar_moeda_br(diff) if abs(diff) >= 0.01 else "-"] )
    data.append(['TOTAL', '', formatar_moeda_br(df_f['Valor_Extrato'].sum()), formatar_moeda_br(df_f['Valor_Razao'].sum()), formatar_moeda_br(df_f['Diferença'].sum())])
    
    # Ajuste de larguras para Paisagem (Total ~277mm)
    t = Table(data, colWidths=[35*mm, 100*mm, 45*mm, 45*mm, 45*mm])
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.black), ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (1,-1), 'CENTER'), ('ALIGN', (2,0), (-1,-1), 'RIGHT'), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey), ('SPAN', (0,-1), (1,-1))
    ]))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

def gerar_pdf_razao_divergente(df_excel, idx_usados, nome_pdf):
    buffer = io.BytesIO()
    # MODO PAISAGEM
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm, title="Divergências do Razão")
    story = []
    styles = getSampleStyleSheet()
    story.append(Paragraph("Lançamentos do Razão com Divergência", styles["Title"]))
    story.append(Paragraph(f"<b>Origem:</b> {nome_pdf}", ParagraphStyle(name='C', alignment=1)))
    story.append(Spacer(1, 15))
    
    df_div = df_excel[~df_excel.index.isin(idx_usados)].copy()
    headers = ['Lançamento', 'Data', 'Valor', 'LCP', 'Histórico']
    data = [headers]
    
    total_div = 0
    for _, r in df_div.iterrows():
        total_div += r['Valor_Razao']
        data.append([
            r['Lancamento_Limpo'], r['Data_str'], formatar_moeda_br(r['Valor_Razao']), 
            str(r['LCP']), Paragraph(str(r['Historico']), styles['Normal'])
        ])
    data.append(['TOTAL', '', formatar_moeda_br(total_div), '', ''])
    
    # Colunas B(1), E(4), I(8), Z(25), AB(27) -> Proporção Paisagem
    t = Table(data, colWidths=[30*mm, 35*mm, 40*mm, 50*mm, 120*mm])
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.black), ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (1,-1), 'CENTER'), ('ALIGN', (2,0), (2,-1), 'RIGHT'), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey), ('SPAN', (0,-1), (1,-1))
    ]))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

def gerar_extrato_marcado(pdf_bytes, df_f, coords_ref, nome_original):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    meta = doc.metadata; meta["title"] = f"{nome_original} Marcado"; doc.set_metadata(meta)
    divergencias = df_f[abs(df_f['Diferença']) >= 0.01]
    for _, erro in divergencias.iterrows():
        for item in coords_ref:
            if item['Data'] == erro['Data'] and abs(item['Valor_Extrato'] - erro['Valor_Extrato']) < 0.01:
                if item['coords']:
                    pno, x0, top, x1, bottom = item['coords']
                    page = doc[pno]; rect = fitz.Rect(x0-2, top-2, x1+2, bottom+2)
                    annot = page.add_highlight_annot(rect); annot.set_colors(stroke=[1, 1, 0]); annot.update()
    return doc.tobytes()

# ==============================================================================
# 3. INTERFACE STREAMLIT
# ==============================================================================
st.markdown("<h1 style='text-align: center;'>Conciliador Bancário (Banco x GovBr)</h1>", unsafe_allow_html=True)
st.markdown("---")

c1, c2 = st.columns(2)
with c1: 
    st.markdown('<p class="big-label">Selecione o Extrato Bancário em PDF</p>', unsafe_allow_html=True)
    up_pdf = st.file_uploader("", type="pdf", key="up_pdf", label_visibility="collapsed")
with c2: 
    st.markdown('<p class="big-label">Selecione o Razão da Contabilidade em Excel</p>', unsafe_allow_html=True)
    up_xlsx = st.file_uploader("", type=["xlsx", "csv"], key="up_xlsx", label_visibility="collapsed")

if st.button("PROCESSAR CONCILIAÇÃO", use_container_width=True):
    if up_pdf and up_xlsx:
        with st.spinner("Processando..."):
            pdf_b, xlsx_b = up_pdf.read(), up_xlsx.read()
            df_p, coords_ref = processar_pdf(pdf_b)
            df_e_comp = processar_excel_detalhado(xlsx_b, df_p, up_xlsx.name.endswith('csv'))
            
            if df_p.empty or df_e_comp.empty: st.error("Erro no processamento."); st.stop()
            
            df_f, idx_usados_excel = executar_conciliacao_inteligente(df_p, df_e_comp)
            
            # Exibição na tela
            st.dataframe(df_f.style.format(subset=['Valor_Extrato', 'Valor_Razao', 'Diferença'], formatter="{:,.2f}"), use_container_width=True)
            
            nome_base = os.path.splitext(up_pdf.name)[0]
            relatorio_pdf = gerar_pdf_final(df_f, f"Conciliação {nome_base}")
            extrato_marcado = gerar_extrato_marcado(pdf_b, df_f, coords_ref, nome_base)
            razao_div_pdf = gerar_pdf_razao_divergente(df_e_comp, idx_usados_excel, up_xlsx.name)
            
            st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
            st.download_button("BAIXAR RELATÓRIO PDF", relatorio_pdf, f"Conciliação {nome_base}.pdf", "application/pdf", use_container_width=True)
            st.download_button("BAIXAR EXTRATO COM MARCAÇÕES", extrato_marcado, f"{nome_base} Marcado.pdf", "application/pdf", use_container_width=True)
            st.download_button("BAIXAR LANÇAMENTOS DO RAZÃO CONTÁBIL", razao_div_pdf, f"Lançamentos Divergentes Razão.pdf", "application/pdf", use_container_width=True)
    else:
        st.warning("⚠️ Selecione os dois arquivos primeiro.")
