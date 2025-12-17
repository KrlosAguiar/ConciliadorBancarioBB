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
# 1. SUAS FUNÇÕES ORIGINAIS (IDÊNTICAS AO SEU SCRIPT)
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
    if pd.isna(valor): return "-"
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
                            documento_candidato = ""
                            if tokens:
                                for t in reversed(tokens):
                                    limpo = t.replace('.', '').replace('-', '')
                                    if limpo.isdigit() and len(limpo) >= 4:
                                        documento_candidato = t
                                        break
                            rows_debitos.append({
                                "Data": data_str,
                                "Histórico": texto_sem_valor.strip(),
                                "Documento": limpar_documento_pdf(documento_candidato),
                                "Valor_Extrato": valor_float
                            })
                        elif tipo == 'C':
                            hist_upper = texto_sem_valor.upper()
                            if "TED DEVOLVIDA" in hist_upper or "DEVOLUCAO DE TED" in hist_upper or "TED DEVOL" in hist_upper:
                                rows_devolucoes.append({"Data": data_str, "Valor_Extrato": valor_float})
    except Exception as e:
        return pd.DataFrame()

    df_debitos = pd.DataFrame(rows_debitos)
    df_devolucoes = pd.DataFrame(rows_devolucoes)
    if df_debitos.empty: return df_debitos
    if not df_devolucoes.empty:
        indices_para_remover = []
        for _, row_dev in df_devolucoes.iterrows():
            matches = df_debitos[(df_debitos['Data'] == row_dev['Data']) & (abs(df_debitos['Valor_Extrato'] - row_dev['Valor_Extrato']) < 0.01) & (~df_debitos.index.isin(indices_para_remover))]
            if not matches.empty: indices_para_remover.append(matches.index[0])
        if indices_para_remover: df_debitos = df_debitos.drop(indices_para_remover).reset_index(drop=True)

    termos_excluir = "SALDO|S A L D O|Resgate|BB-APLIC C\.PRZ-APL\.AUT"
    filtro_excecao = df_debitos['Histórico'].astype(str).str.strip().str.contains(termos_excluir, case=False, regex=True)
    df = df_debitos[~filtro_excecao].copy()
    df['Data_dt'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    mask_13113 = df['Histórico'].astype(str).str.contains("13113", na=False)
    df_tarifas = df[mask_13113].copy()
    df_outros = df[~mask_13113].copy()
    if not df_tarifas.empty:
        df_tarifas_agg = df_tarifas.groupby('Data_dt').agg({'Valor_Extrato': 'sum', 'Data': 'first'}).reset_index()
        df_tarifas_agg['Documento'] = "Tarifas Bancárias"
        df_tarifas_agg['Histórico'] = "Tarifas Bancárias do Dia"
        df = pd.concat([df_outros, df_tarifas_agg], ignore_index=True)
    return df

def processar_excel_detalhado(file_bytes, df_pdf_ref, is_csv=False):
    try:
        if is_csv:
            df = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='latin1', sep=None, engine='python')
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None)
        try:
            df = df.iloc[:, [4, 5, 8, 25, 26, 27]].copy()
        except IndexError:
            df = df.iloc[:, [4, 5, 8, -4, -2, -1]].copy()
        df.columns = ['Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
        mask_pagto = df['Info_Z'].astype(str).str.contains("Pagamento", case=False, na=False)
        mask_transf_nome = df['Info_Z'].astype(str).str.contains("TRANSFERENCIA ENTRE CONTAS DE MESMA UG", case=False, na=False)
        mask_credito = df['DC'].astype(str).str.strip().str.upper() == 'C'
        df = df[mask_pagto | (mask_transf_nome & mask_credito)].copy()
        df['Data_dt'] = df['Data'].apply(parse_br_date)
        df = df.dropna(subset=['Data_dt'])
        df['Data'] = df['Data_dt'].dt.strftime('%d/%m/%Y')
        df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x, str) else float(x))
        
        lookup_pdf = {}
        for dt, group in df_pdf_ref.groupby('Data'):
            lookup_pdf[dt] = {str(doc).lstrip('0'): doc for doc in group['Documento'].unique() if doc != "Tarifas Bancárias"}
            if "Tarifas Bancárias" in group['Documento'].values: lookup_pdf[dt]["TARIFA"] = "Tarifas Bancárias"
        
        def encontrar_doc_inteligente(row):
            texto = str(row['Info_AB']).upper()
            dt = row['Data']
            if dt not in lookup_pdf: return "S/D"
            if "TARIFA" in texto and "TARIFA" in lookup_pdf[dt]: return lookup_pdf[dt]["TARIFA"]
            nums = re.findall(r'\d+', texto)
            for num in nums:
                nl = num.lstrip('0')
                if nl in lookup_pdf[dt]: return lookup_pdf[dt][nl]
            return "NÃO LOCALIZADO"
        df['Documento'] = df.apply(encontrar_doc_inteligente, axis=1)
        return df[['Data', 'Documento', 'Valor_Razao']].reset_index(drop=True)
    except: return pd.DataFrame()

def executar_conciliacao_inteligente(df_pdf, df_excel):
    df_pdf = df_pdf.copy()
    df_excel = df_excel.copy()
    resultados_finais = []
    indices_pdf_usados = set()
    indices_excel_usados = set()
    for idx_pdf, row_pdf in df_pdf.iterrows():
        candidatos = df_excel[(df_excel['Data'] == row_pdf['Data']) & (df_excel['Documento'] == row_pdf['Documento']) & (~df_excel.index.isin(indices_excel_usados))]
        match = candidatos[abs(candidatos['Valor_Razao'] - row_pdf['Valor_Extrato']) < 0.01]
        if not match.empty:
            idx_excel = match.index[0]
            resultados_finais.append({'Data': row_pdf['Data'], 'Histórico': row_pdf['Histórico'], 'Documento': row_pdf['Documento'], 'Valor_Extrato': row_pdf['Valor_Extrato'], 'Valor_Razao': match.loc[idx_excel]['Valor_Razao'], 'Diferença': 0.0, 'Obs': ''})
            indices_pdf_usados.add(idx_pdf); indices_excel_usados.add(idx_excel)
    for idx_pdf, row_pdf in df_pdf.iterrows():
        if idx_pdf in indices_pdf_usados: continue
        candidatos = df_excel[(df_excel['Data'] == row_pdf['Data']) & (~df_excel.index.isin(indices_excel_usados))]
        match = candidatos[abs(candidatos['Valor_Razao'] - row_pdf['Valor_Extrato']) < 0.01]
        if not match.empty:
            idx_excel = match.index[0]
            resultados_finais.append({'Data': row_pdf['Data'], 'Histórico': row_pdf['Histórico'], 'Documento': "Docs dif.", 'Valor_Extrato': row_pdf['Valor_Extrato'], 'Valor_Razao': match.loc[idx_excel]['Valor_Razao'], 'Diferença': 0.0, 'Obs': 'Doc Diferente'})
            indices_pdf_usados.add(idx_pdf); indices_excel_usados.add(idx_excel)
    df_e_s = df_excel[~df_excel.index.isin(indices_excel_usados)].groupby(['Data', 'Documento'])['Valor_Razao'].sum().reset_index()
    df_p_s = df_pdf[~df_pdf.index.isin(indices_pdf_usados)].groupby(['Data', 'Documento', 'Histórico'])['Valor_Extrato'].sum().reset_index()
    df_sobras = pd.merge(df_p_s, df_e_s, on=['Data', 'Documento'], how='outer').fillna(0)
    for _, row in df_sobras.iterrows():
        resultados_finais.append({'Data': row['Data'], 'Histórico': row.get('Histórico', 'S/H'), 'Documento': row['Documento'], 'Valor_Extrato': row['Valor_Extrato'], 'Valor_Razao': row['Valor_Razao'], 'Diferença': row['Valor_Extrato'] - row['Valor_Razao'], 'Obs': 'Agrupado'})
    df_final = pd.DataFrame(resultados_finais)
    df_final['dt'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y', errors='coerce')
    return df_final.sort_values(by=['dt', 'Documento']).drop(columns=['dt'])

def gerar_pdf_final(df_final, nome_conta_original):
    buffer = io.BytesIO()
    nome_conta_limpo = os.path.splitext(nome_conta_original)[0]
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm, title=f"Conciliação {nome_conta_limpo}")
    styles = getSampleStyleSheet()
    story = []
    style_centered = ParagraphStyle(name='Centered', parent=styles['Normal'], alignment=1)
    story.append(Paragraph("Relatório de Conciliação Bancária", styles["Title"]))
    story.append(Paragraph(f"<b>Conta:</b> {nome_conta_limpo}", style_centered))
    story.append(Spacer(1, 24))
    headers = ['Data', 'Histórico', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']
    data_list = [headers]
    for _, row in df_final.iterrows():
        diff = row['Diferença']
        data_list.append([row['Data'], Paragraph(str(row['Histórico']), styles['Normal']), str(row['Documento']), formatar_moeda_br(row['Valor_Extrato']), formatar_moeda_br(row['Valor_Razao']), formatar_moeda_br(diff) if abs(diff) >= 0.01 else "-"])
    data_list.append(['TOTAL', '', '', formatar_moeda_br(df_final['Valor_Extrato'].sum()), formatar_moeda_br(df_final['Valor_Razao'].sum()), formatar_moeda_br(df_final['Diferença'].sum())])
    t = Table(data_list, colWidths=[22*mm, 55*mm, 25*mm, 28*mm, 28*mm, 28*mm])
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),   # Data Centro
        ('ALIGN', (1,0), (1,-1), 'LEFT'),     # Histórico Esquerda
        ('ALIGN', (2,0), (2,-1), 'CENTER'),   # Documento Centro
        ('ALIGN', (3,0), (-1,-1), 'RIGHT'),   # Valores Direita
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
        ('SPAN', (0,-1), (2,-1)),
    ]))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

# ==============================================================================
# 2. INTERFACE (STREAMLIT ADAPTADO)
# ==============================================================================
if check_password():
    st.title("🏦 Conciliador Bancário - Banco do Brasil")
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1: up_pdf = st.file_uploader("1. Extrato (PDF)", type="pdf")
    with col2: up_xlsx = st.file_uploader("2. Razão (Excel/CSV)", type=["xlsx", "csv"])

    if st.button("🚀 Processar Conciliação", use_container_width=True):
        if up_pdf and up_xlsx:
            with st.spinner("⚙️ Processando..."):
                df_p = processar_pdf(up_pdf.read())
                df_e = processar_excel_detalhado(up_xlsx.read(), df_p, is_csv=up_xlsx.name.endswith('csv'))
                
                if df_p.empty: st.error("❌ Erro no PDF."); st.stop()
                if df_e.empty: st.error("❌ Erro no Excel (Verifique Filtro Z)."); st.stop()
                
                df_f = executar_conciliacao_inteligente(df_p, df_e)
                
                st.success("✅ RELATÓRIO PRONTO!")
                
                # --- EXIBIÇÃO EM TELA IGUAL AO PDF ---
                # Usando HTML para garantir os alinhamentos exatos pedidos
                def formatar_tabela_html(df):
                    html = """<table style='width:100%; border-collapse: collapse; font-family: sans-serif;'>
                        <tr style='background-color: #00008B; color: white;'>
                            <th style='padding: 8px; text-align : center;'>Data</th>
                            <th style='padding: 8px; text-align : left;'>Histórico</th>
                            <th style='padding: 8px; text-align : center;'>Documento</th>
                            <th style='padding: 8px; text-align : right;'>Vlr. Extrato</th>
                            <th style='padding: 8px; text-align : right;'>Vlr. Razão</th>
                            <th style='padding: 8px; text-align : right;'>Diferença</th>
                        </tr>"""
                    for _, row in df.iterrows():
                        diff_color = "red" if abs(row['Diferença']) >= 0.01 else "black"
                        diff_text = formatar_moeda_br(row['Diferença']) if abs(row['Diferença']) >= 0.01 else "-"
                        html += f"""<tr style='border-bottom: 1px solid #ddd;'>
                            <td style='padding: 8px; text-align: center;'>{row['Data']}</td>
                            <td style='padding: 8px; text-align: left;'>{row['Histórico']}</td>
                            <td style='padding: 8px; text-align: center;'>{row['Documento']}</td>
                            <td style='padding: 8px; text-align: right;'>{formatar_moeda_br(row['Valor_Extrato'])}</td>
                            <td style='padding: 8px; text-align: right;'>{formatar_moeda_br(row['Valor_Razao'])}</td>
                            <td style='padding: 8px; text-align: right; color: {diff_color};'>{diff_text}</td>
                        </tr>"""
                    # Linha de Total
                    html += f"""<tr style='background-color: #f2f2f2; font-weight: bold;'>
                        <td colspan='3' style='padding: 8px; text-align: center;'>TOTAL</td>
                        <td style='padding: 8px; text-align: right;'>{formatar_moeda_br(df['Valor_Extrato'].sum())}</td>
                        <td style='padding: 8px; text-align: right;'>{formatar_moeda_br(df['Valor_Razao'].sum())}</td>
                        <td style='padding: 8px; text-align: right;'>{formatar_moeda_br(df['Diferença'].sum())}</td>
                    </tr></table>"""
                    return html

                st.markdown(formatar_tabela_html(df_f), unsafe_allow_html=True)

                # --- DOWNLOAD DO PDF ---
                pdf_data = gerar_pdf_final(df_f, up_pdf.name)
                st.download_button(label="📥 Baixar Relatório PDF", data=pdf_data, file_name=f"Conciliacao_{up_pdf.name.split('.')[0]}.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.warning("⚠️ Selecione os dois arquivos primeiro.")
