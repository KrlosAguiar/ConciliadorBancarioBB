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
# SUAS FUNÇÕES ORIGINAIS (MANTIDAS 100%)
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

def processar_pdf(file_stream):
    rows_debitos = []
    rows_devolucoes = [] 
    with pdfplumber.open(file_stream) as pdf:
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
    
    df_debitos = pd.DataFrame(rows_debitos)
    if not rows_devolucoes == [] and not df_debitos.empty:
        df_dev = pd.DataFrame(rows_devolucoes)
        idx_rem = []
        for _, row_dev in df_dev.iterrows():
            m = df_debitos[(df_debitos['Data'] == row_dev['Data']) & (abs(df_debitos['Valor_Extrato'] - row_dev['Valor_Extrato']) < 0.01) & (~df_debitos.index.isin(idx_rem))]
            if not m.empty: idx_rem.append(m.index[0])
        df_debitos = df_debitos.drop(idx_rem).reset_index(drop=True)

    termos_excluir = "SALDO|S A L D O|Resgate|BB-APLIC C\.PRZ-APL\.AUT"
    df = df_debitos[~df_debitos['Histórico'].str.contains(termos_excluir, case=False, na=False)].copy()
    
    df['Data_dt'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    mask_13113 = df['Histórico'].str.contains("13113", na=False)
    df_tarifas = df[mask_13113].copy()
    df_outros = df[~mask_13113].copy()
    if not df_tarifas.empty:
        df_t_agg = df_tarifas.groupby('Data_dt').agg({'Valor_Extrato':'sum','Data':'first'}).reset_index()
        df_t_agg['Documento'] = "Tarifas Bancárias"
        df_t_agg['Histórico'] = "Tarifas Bancárias do Dia"
        df = pd.concat([df_outros, df_t_agg], ignore_index=True)
    return df

def processar_excel_detalhado(file_stream, df_pdf_ref, is_csv=False):
    if is_csv: df = pd.read_csv(file_stream, header=None, encoding='latin1', sep=None, engine='python')
    else: df = pd.read_excel(file_stream, header=None)
    try: df = df.iloc[:, [4, 5, 8, 25, 26, 27]].copy()
    except: df = df.iloc[:, [4, 5, 8, -4, -2, -1]].copy()
    df.columns = ['Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
    mask_pagto = df['Info_Z'].astype(str).str.contains("Pagamento", case=False, na=False)
    mask_transf = (df['Info_Z'].astype(str).str.contains("TRANSFERENCIA", case=False, na=False)) & (df['DC'].str.strip().str.upper() == 'C')
    df = df[mask_pagto | mask_transf].copy()
    df['Data_dt'] = df['Data'].apply(parse_br_date)
    df = df.dropna(subset=['Data_dt'])
    df['Data'] = df['Data_dt'].dt.strftime('%d/%m/%Y')
    df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.','').replace(',','.')) if isinstance(x, str) else float(x))
    
    # Lógica de match inteligente de documento (sua original)
    lookup_pdf = {}
    for dt, group in df_pdf_ref.groupby('Data'):
        lookup_pdf[dt] = {str(d).lstrip('0'): d for d in group['Documento'].unique() if d != "Tarifas Bancárias"}
        if "Tarifas Bancárias" in group['Documento'].values: lookup_pdf[dt]["TARIFA"] = "Tarifas Bancárias"

    def find_doc(row):
        txt = str(row['Info_AB']).upper()
        dt = row['Data']
        if dt not in lookup_pdf: return "S/D"
        if "TARIFA" in txt and "TARIFA" in lookup_pdf[dt]: return lookup_pdf[dt]["TARIFA"]
        nums = re.findall(r'\d+', txt)
        for n in nums:
            nl = n.lstrip('0')
            if nl in lookup_pdf[dt]: return lookup_pdf[dt][nl]
        return "NÃO LOCALIZADO"

    df['Documento'] = df.apply(find_doc, axis=1)
    return df[['Data', 'Documento', 'Valor_Razao']]

def executar_conciliacao_inteligente(df_pdf, df_excel):
    # Sua lógica de 3 níveis de Match (Exato, Valor, Agrupado)
    df_pdf = df_pdf.copy()
    df_excel = df_excel.copy()
    res = []
    idx_p_usados = set()
    idx_e_usados = set()
    # 1. Exato
    for idx_p, row_p in df_pdf.iterrows():
        cand = df_excel[(df_excel['Data'] == row_p['Data']) & (df_excel['Documento'] == row_p['Documento']) & (~df_excel.index.isin(idx_e_usados))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'], 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': m.loc[idx_e]['Valor_Razao'], 'Diferença': 0.0, 'Obs': ''})
            idx_p_usados.add(idx_p); idx_e_usados.add(idx_e)
    # 2. Valor
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p in idx_p_usados: continue
        cand = df_excel[(df_excel['Data'] == row_p['Data']) & (~df_excel.index.isin(idx_e_usados))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            res.append({'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': "Docs dif.", 'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': m.loc[idx_e]['Valor_Razao'], 'Diferença': 0.0, 'Obs': 'Doc Diferente'})
            idx_p_usados.add(idx_p); idx_e_usados.add(idx_e)
    # 3. Sobras
    df_e_s = df_excel[~df_excel.index.isin(idx_e_usados)].groupby(['Data', 'Documento'])['Valor_Razao'].sum().reset_index()
    df_p_s = df_pdf[~df_pdf.index.isin(idx_p_usados)].groupby(['Data', 'Documento', 'Histórico'])['Valor_Extrato'].sum().reset_index()
    df_m = pd.merge(df_p_s, df_e_s, on=['Data', 'Documento'], how='outer').fillna(0)
    for _, row in df_m.iterrows():
        res.append({'Data': row['Data'], 'Histórico': row.get('Histórico', 'S/H'), 'Documento': row['Documento'], 'Valor_Extrato': row['Valor_Extrato'], 'Valor_Razao': row['Valor_Razao'], 'Diferença': row['Valor_Extrato'] - row['Valor_Razao'], 'Obs': 'Agrupado'})
    
    df_f = pd.DataFrame(res)
    df_f['dt'] = pd.to_datetime(df_f['Data'], format='%d/%m/%Y', errors='coerce')
    return df_f.sort_values('dt').drop(columns=['dt'])

def gerar_pdf_reportlab(df_final, nome_conta):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm)
    styles = getSampleStyleSheet()
    story = []
    
    # Estilos de alinhamento
    style_center = ParagraphStyle(name='Center', parent=styles['Normal'], alignment=1)
    story.append(Paragraph("Relatório de Conciliação Bancária", styles["Title"]))
    story.append(Paragraph(f"<b>Conta:</b> {nome_conta}", style_center))
    story.append(Spacer(1, 12))

    # Tabela formatada (Sua original)
    headers = ['Data', 'Histórico', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']
    data = [headers]
    for _, row in df_final.iterrows():
        diff = row['Diferença']
        data.append([
            row['Data'], 
            Paragraph(str(row['Histórico']), styles['Normal']), # Histórico alinhado a esquerda
            row['Documento'],
            formatar_moeda_br(row['Valor_Extrato']),
            formatar_moeda_br(row['Valor_Razao']),
            formatar_moeda_br(diff) if abs(diff) >= 0.01 else "-"
        ])
    
    # Totais
    data.append(['TOTAL', '', '', formatar_moeda_br(df_final['Valor_Extrato'].sum()), formatar_moeda_br(df_final['Valor_Razao'].sum()), formatar_moeda_br(df_final['Diferença'].sum())])

    t = Table(data, colWidths=[22*mm, 55*mm, 25*mm, 28*mm, 28*mm, 28*mm])
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),   # Data Centro
        ('ALIGN', (1,0), (1,-1), 'LEFT'),     # Histórico Esquerda
        ('ALIGN', (2,0), (2,-1), 'CENTER'),   # Doc Centro
        ('ALIGN', (3,0), (-1,-1), 'RIGHT'),   # Valores Direita
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
    ]))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

# ==============================================================================
# INTERFACE STREAMLIT (SUBSTITUI O IPYWIDGETS)
# ==============================================================================
if check_password():
    st.title("🏦 Conciliador Bancário - Banco do Brasil")
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1: up_pdf = st.file_uploader("1. Extrato (PDF)", type="pdf")
    with col2: up_xlsx = st.file_uploader("2. Razão (Excel/CSV)", type=["xlsx", "csv"])

    # O BOTÃO DE PROCESSAR QUE VOCÊ PEDIU
    if st.button("🚀 Processar Conciliação", use_container_width=True):
        if not up_pdf or not up_xlsx:
            st.warning("⚠️ Por favor, faça o upload dos dois arquivos.")
        else:
            with st.spinner("⚙️ Processando..."):
                # Lógica idêntica ao seu on_proc_click
                df_p = processar_pdf(up_pdf)
                df_e = processar_excel_detalhado(up_xlsx, df_p, is_csv=up_xlsx.name.endswith('csv'))
                
                if df_p.empty or df_e.empty:
                    st.error("❌ Verifique se os arquivos estão corretos (Filtro Z não encontrou dados).")
                else:
                    df_f = executar_conciliacao_inteligente(df_p, df_e)
                    
                    st.success("✅ Conciliação Concluída!")
                    
                    # Exibição na tela com alinhamentos pedidos
                    st.dataframe(df_f.style.format({
                        'Valor_Extrato': 'R$ {:,.2f}',
                        'Valor_Razao': 'R$ {:,.2f}',
                        'Diferença': 'R$ {:,.2f}'
                    }), use_container_width=True)

                    # Geração do PDF original ReportLab
                    pdf_data = gerar_pdf_reportlab(df_f, up_pdf.name)
                    st.download_button(
                        label="📥 Baixar Relatório PDF Profissional",
                        data=pdf_data,
                        file_name=f"Relatorio_{up_pdf.name.split('.')[0]}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
