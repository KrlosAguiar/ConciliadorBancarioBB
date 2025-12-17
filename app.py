import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Conciliador Bancário - Banco do Brasil", layout="wide")

# --- SENHA ---
SENHA_MESTRA = "cliente123"

def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if st.session_state["password_correct"]: return True
    st.title("🔐 Acesso Restrito")
    password = st.text_input("Digite a chave de acesso:", type="password")
    if st.button("Entrar"):
        if password == SENHA_MESTRA:
            st.session_state["password_correct"] = True
            st.rerun()
        else: st.error("Chave incorreta!")
    return False

# ==============================================================================
# 1. FUNÇÕES ORIGINAIS (INTEGRALMENTE CONSERVADAS)
# ==============================================================================
CURRENT_YEAR = str(datetime.datetime.now().year)

def limpar_documento_pdf(doc_str):
    if not doc_str: return ""
    apenas_digitos = re.sub(r'\D', '', str(doc_str))
    if not apenas_digitos: return ""
    if len(apenas_digitos) > 6:
        return apenas_digitos[-6:]
    return apenas_digitos

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == "-": return "-"
    return f"{valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

def parse_br_date(date_val):
    if pd.isna(date_val): return pd.NaT
    try:
        if isinstance(date_val, str):
            date_val = date_val.split()[0]
        return pd.to_datetime(date_val, dayfirst=True, errors='coerce')
    except:
        return pd.to_datetime(date_val, errors='coerce')

def processar_pdf(file_bytes):
    rows_debitos = []
    rows_devolucoes = [] 
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
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
                        texto_sem_data = texto_linha.replace(match_data.group(0), "", 1).strip()
                        texto_sem_valor = texto_sem_data.replace(match_valor.group(0), "").strip()
                        if tipo == 'D':
                            tokens = texto_sem_valor.split()
                            doc_cand = ""
                            if tokens:
                                for t in reversed(tokens):
                                    limpo = t.replace('.', '').replace('-', '')
                                    if limpo.isdigit() and len(limpo) >= 4:
                                        doc_cand = t
                                        break
                            rows_debitos.append({
                                "Data": data_str, "Histórico": texto_sem_valor.strip(),
                                "Documento": limpar_documento_pdf(doc_cand), "Valor_Extrato": valor_float
                            })
                        elif tipo == 'C':
                            hist_upper = texto_sem_valor.upper()
                            if any(x in hist_upper for x in ["TED DEVOLVIDA", "DEVOLUCAO DE TED", "TED DEVOL"]):
                                rows_devolucoes.append({"Data": data_str, "Valor_Extrato": valor_float})
    except: return pd.DataFrame()
    df_debitos = pd.DataFrame(rows_debitos)
    if not df_devolucoes.empty and not df_debitos.empty:
        idx_rem = []
        for _, r_dev in df_devolucoes.iterrows():
            m = df_debitos[(df_debitos['Data'] == r_dev['Data']) & (abs(df_debitos['Valor_Extrato'] - r_dev['Valor_Extrato']) < 0.01) & (~df_debitos.index.isin(idx_rem))]
            if not m.empty: idx_rem.append(m.index[0])
        df_debitos = df_debitos.drop(idx_rem).reset_index(drop=True)
    df = df_debitos[~df_debitos['Histórico'].astype(str).str.contains("SALDO|S A L D O|Resgate|BB-APLIC", case=False, na=False)].copy()
    df['Data_dt'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    mask_13113 = df['Histórico'].astype(str).str.contains("13113", na=False)
    if any(mask_13113):
        df_t = df[mask_13113].copy(); df_o = df[~mask_13113].copy()
        df_t_agg = df_t.groupby('Data_dt').agg({'Valor_Extrato': 'sum', 'Data': 'first'}).reset_index()
        df_t_agg['Documento'] = "Tarifas Bancárias"; df_t_agg['Histórico'] = "Tarifas Bancárias do Dia"
        df = pd.concat([df_o, df_t_agg], ignore_index=True)
    return df

def processar_excel_detalhado(file_bytes, df_pdf_ref, is_csv=False):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='latin1', sep=None, engine='python') if is_csv else pd.read_excel(io.BytesIO(file_bytes), header=None)
        try: df = df.iloc[:, [4, 5, 8, 25, 26, 27]].copy()
        except: df = df.iloc[:, [4, 5, 8, -4, -2, -1]].copy()
        df.columns = ['Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
        mask_pagto = df['Info_Z'].astype(str).str.contains("Pagamento", case=False, na=False)
        mask_transf = (df['Info_Z'].astype(str).str.contains("TRANSFERENCIA", case=False, na=False)) & (df['DC'].str.strip().str.upper() == 'C')
        df = df[mask_pagto | mask_transf].copy()
        df['Data_dt'] = df['Data'].apply(parse_br_date); df = df.dropna(subset=['Data_dt'])
        df['Data'] = df['Data_dt'].dt.strftime('%d/%m/%Y')
        df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x, str) else float(x))
        lookup = {dt: {str(doc).lstrip('0'): doc for doc in g['Documento'].unique()} for dt, g in df_pdf_ref.groupby('Data')}
        def find_doc(row):
            txt, dt = str(row['Info_AB']).upper(), row['Data']
            if dt not in lookup: return "S/D"
            if "TARIFA" in txt and "Tarifas Bancárias" in lookup[dt].values(): return "Tarifas Bancárias"
            for n in re.findall(r'\d+', txt):
                if n.lstrip('0') in lookup[dt]: return lookup[dt][n.lstrip('0')]
            return "NÃO LOCALIZADO"
        df['Documento'] = df.apply(find_doc, axis=1)
        return df[['Data', 'Documento', 'Valor_Razao']]
    except: return pd.DataFrame()

def executar_conciliacao_inteligente(df_pdf, df_excel):
    res, idx_p_u, idx_e_u = [], set(), set()
    for idx_p, row_p in df_pdf.iterrows():
        cand = df_excel[(df_excel['Data'] == row_p['Data']) & (df_excel['Documento'] == row_p['Documento']) & (~df_excel.index.isin(idx_e_u))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'], 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': m.loc[idx_e]['Valor_Razao'], 'Diferença': 0.0})
            idx_p_u.add(idx_p); idx_e_u.add(idx_e)
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p in idx_p_u: continue
        cand = df_excel[(df_excel['Data'] == row_p['Data']) & (~df_excel.index.isin(idx_e_u))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': "Docs dif.", 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': m.loc[idx_e]['Valor_Razao'], 'Diferença': 0.0})
            idx_p_u.add(idx_p); idx_e_u.add(idx_e)
    df_e_s = df_excel[~df_excel.index.isin(idx_e_u)].groupby(['Data', 'Documento'])['Valor_Razao'].sum().reset_index()
    df_p_s = df_pdf[~df_pdf.index.isin(idx_p_u)].groupby(['Data', 'Documento', 'Histórico'])['Valor_Extrato'].sum().reset_index()
    df_m = pd.merge(df_p_s, df_e_s, on=['Data', 'Documento'], how='outer').fillna(0)
    for _, row in df_m.iterrows():
        res.append({'Data': row['Data'], 'Histórico': row.get('Histórico', 'S/H'), 'Documento': row['Documento'], 'Valor_Extrato': row['Valor_Extrato'], 'Valor_Razao': row['Valor_Razao'], 'Diferença': row['Valor_Extrato'] - row['Valor_Razao']})
    df_f = pd.DataFrame(res)
    df_f['dt'] = pd.to_datetime(df_f['Data'], format='%d/%m/%Y', errors='coerce')
    return df_f.sort_values(by=['dt', 'Documento']).drop(columns=['dt'])

# ==============================================================================
# 2. GERAÇÃO PDF (HISTÓRICO REMOVIDO CONFORME PEDIDO)
# ==============================================================================
def gerar_pdf_final(df_f, nome_orig):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm)
    story = []
    styles = getSampleStyleSheet()
    story.append(Paragraph("Relatório de Conciliação Bancária", styles["Title"]))
    story.append(Paragraph(f"<b>Conta:</b> {os.path.splitext(nome_orig)[0]}", ParagraphStyle(name='C', alignment=1)))
    story.append(Spacer(1, 15))
    
    headers = ['Data', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']
    data = [headers]
    for _, r in df_f.iterrows():
        diff = r['Diferença']
        data.append([r['Data'], str(r['Documento']), formatar_moeda_br(r['Valor_Extrato']), formatar_moeda_br(r['Valor_Razao']), formatar_moeda_br(diff) if abs(diff) >= 0.01 else "-"] )
    
    data.append(['TOTAL', '', formatar_moeda_br(df_f['Valor_Extrato'].sum()), formatar_moeda_br(df_f['Valor_Razao'].sum()), formatar_moeda_br(df_f['Diferença'].sum())])
    
    t = Table(data, colWidths=[25*mm, 65*mm, 33*mm, 33*mm, 33*mm])
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (0,-1), 'CENTER'), ('ALIGN', (1,0), (1,-1), 'CENTER'), ('ALIGN', (2,0), (-1,-1), 'RIGHT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'), ('SPAN', (0,-1), (1,-1))
    ]))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

# ==============================================================================
# 3. INTERFACE (STREAMLIT COM FUNDO BRANCO E HISTÓRICO)
# ==============================================================================
if check_password():
    st.title("🏦 Conciliador Bancário - Banco do Brasil")
    c1, c2 = st.columns(2)
    with c1: up_pdf = st.file_uploader("1. Extrato (PDF)", type="pdf")
    with c2: up_xlsx = st.file_uploader("2. Razão (Excel/CSV)", type=["xlsx", "csv"])

    if st.button("🚀 Processar Conciliação", use_container_width=True):
        if up_pdf and up_xlsx:
            with st.spinner("Processando..."):
                df_p = processar_pdf(up_pdf.read())
                df_e = processar_excel_detalhado(up_xlsx.read(), df_p, is_csv=up_xlsx.name.endswith('csv'))
                if df_p.empty or df_e.empty: st.error("Erro nos arquivos."); st.stop()
                df_f = executar_conciliacao_inteligente(df_p, df_e)
                
                # HTML com Fundo Branco, Histórico e Totais em Negrito
                html = """
                <div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>
                <table style='width:100%; border-collapse: collapse; color: black; background-color: white;'>
                    <tr style='background-color: #00008B; color: white;'>
                        <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Data</th>
                        <th style='padding: 8px; text-align: left; border: 1px solid #000;'>Histórico</th>
                        <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Documento</th>
                        <th style='padding: 8px; text-align: right; border: 1px solid #000;'>Vlr. Extrato</th>
                        <th style='padding: 8px; text-align: right; border: 1px solid #000;'>Vlr. Razão</th>
                        <th style='padding: 8px; text-align: right; border: 1px solid #000;'>Diferença</th>
                    </tr>"""
                for _, r in df_f.iterrows():
                    d_c = "red" if abs(r['Diferença']) >= 0.01 else "black"
                    html += f"""
                    <tr style='border: 1px solid #000;'> 
                        <td style='text-align: center; border: 1px solid #000;'>{r['Data']}</td> 
                        <td style='text-align: left; border: 1px solid #000;'>{r['Histórico']}</td> 
                        <td style='text-align: center; border: 1px solid #000;'>{r['Documento']}</td> 
                        <td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(r['Valor_Extrato'])}</td> 
                        <td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(r['Valor_Razao'])}</td> 
                        <td style='text-align: right; color: {d_c}; border: 1px solid #000;'>{formatar_moeda_br(r['Diferença']) if abs(r['Diferença']) >= 0.01 else '-'}</td> 
                    </tr>"""
                
                # Linha de Totais: Fundo Branco, Fonte Negrito
                html += f"""
                    <tr style='font-weight: bold; background-color: white; border: 1px solid #000;'> 
                        <td colspan='3' style='padding: 10px; text-align: center; border: 1px solid #000;'>TOTAL</td>
                        <td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_f['Valor_Extrato'].sum())}</td>
                        <td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_f['Valor_Razao'].sum())}</td>
                        <td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_f['Diferença'].sum())}</td> 
                    </tr> </table> </div>"""
                
                st.markdown(html, unsafe_allow_html=True)
                pdf_data = gerar_pdf_final(df_f, up_pdf.name)
                st.download_button("📥 Baixar Relatório PDF", pdf_data, "Relatorio.pdf", "application/pdf", use_container_width=True)
        else:
            st.warning("⚠️ Selecione os dois arquivos primeiro.")
