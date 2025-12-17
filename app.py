import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Conciliador Bancário Pro", page_icon="🏦", layout="wide")

# --- SISTEMA DE LOGIN (SENHA) ---
def check_password():
    """Retorna True se o usuário acertar a senha."""
    
    # --- DEFINA A SENHA AQUI ---
    SENHA_MESTRA = "cliente123" 
    # ---------------------------

    def password_entered():
        if st.session_state["password"] == SENHA_MESTRA:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("### 🔒 Área Restrita")
        st.text_input("Digite sua Chave de Acesso:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("### 🔒 Área Restrita")
        st.text_input("Digite sua Chave de Acesso:", type="password", on_change=password_entered, key="password")
        st.error("🚫 Chave incorreta.")
        return False
    else:
        return True

# ==============================================================================
# 1. FUNÇÕES AUXILIARES
# ==============================================================================

def limpar_documento_pdf(doc_str):
    if not doc_str: return ""
    apenas_digitos = re.sub(r'\D', '', str(doc_str))
    if not apenas_digitos: return ""
    if len(apenas_digitos) > 6:
        return apenas_digitos[-6:]
    return apenas_digitos

def formatar_moeda_br(valor):
    if pd.isna(valor): return "-"
    # Formata para padrão brasileiro: 1.000,00
    return f"{valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

def parse_br_date(date_val):
    if pd.isna(date_val): return pd.NaT
    try:
        if isinstance(date_val, str):
            date_val = date_val.split()[0]
        return pd.to_datetime(date_val, dayfirst=True, errors='coerce')
    except:
        return pd.to_datetime(date_val, errors='coerce')

# ==============================================================================
# 2. PROCESSAMENTO DO PDF
# ==============================================================================

def processar_pdf(file_path):
    rows_debitos = []
    rows_devolucoes = [] 
    current_year = str(datetime.datetime.now().year)
    
    try:
        with pdfplumber.open(file_path) as pdf:
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
                    if len(data_str) == 5: data_str = f"{data_str}/{current_year}"
                    
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
                                rows_devolucoes.append({
                                    "Data": data_str,
                                    "Valor_Extrato": valor_float
                                })
                    else:
                        continue 

    except Exception as e:
        st.error(f"Erro ao ler PDF: {e}")
        return pd.DataFrame()

    df_debitos = pd.DataFrame(rows_debitos)
    df_devolucoes = pd.DataFrame(rows_devolucoes)
    
    if df_debitos.empty: return df_debitos

    if not df_devolucoes.empty:
        indices_para_remover = []
        for _, row_dev in df_devolucoes.iterrows():
            matches = df_debitos[
                (df_debitos['Data'] == row_dev['Data']) &
                (abs(df_debitos['Valor_Extrato'] - row_dev['Valor_Extrato']) < 0.01) &
                (~df_debitos.index.isin(indices_para_remover))
            ]
            if not matches.empty:
                indices_para_remover.append(matches.index[0])
        if indices_para_remover:
            df_debitos = df_debitos.drop(indices_para_remover).reset_index(drop=True)

    termos_excluir = "SALDO|S A L D O|Resgate|BB-APLIC C\.PRZ-APL\.AUT"
    filtro_excecao = df_debitos['Histórico'].astype(str).str.strip().str.contains(termos_excluir, case=False, regex=True)
    df = df_debitos[~filtro_excecao].copy()

    df['Data_dt'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    mask_13113 = df['Histórico'].astype(str).str.contains("13113", na=False)
    
    df_tarifas = df[mask_13113].copy()
    df_outros = df[~mask_13113].copy()

    if not df_tarifas.empty:
        df_tarifas_agg = df_tarifas.groupby('Data_dt').agg({
            'Valor_Extrato': 'sum',
            'Data': 'first'
        }).reset_index()
        df_tarifas_agg['Documento'] = "Tarifas Bancárias"
        df_tarifas_agg['Histórico'] = "Tarifas Bancárias do Dia"
        df = pd.concat([df_outros, df_tarifas_agg], ignore_index=True)
    
    return df.drop(columns=['Data_dt', 'Histórico'], errors='ignore')

# ==============================================================================
# 3. PROCESSAMENTO DO EXCEL
# ==============================================================================

def processar_excel_detalhado(file_path, df_pdf_ref):
    try:
        if file_path.endswith('.csv'):
             df = pd.read_csv(file_path, header=None, encoding='latin1', sep=None, engine='python')
        else:
             df = pd.read_excel(file_path, header=None)

        try:
            df = df.iloc[:, [4, 5, 8, 25, 26, 27]].copy()
        except IndexError:
            df = df.iloc[:, [4, 5, 8, -4, -2, -1]].copy()

        df.columns = ['Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
        
        # Filtro Coluna Z: Pagamento OU (Transferência E Crédito)
        mask_pagto = df['Info_Z'].astype(str).str.contains("Pagamento", case=False, na=False)
        mask_transf_nome = df['Info_Z'].astype(str).str.contains("TRANSFERENCIA ENTRE CONTAS DE MESMA UG", case=False, na=False)
        mask_credito = df['DC'].astype(str).str.strip().str.upper() == 'C'
        mask_transf_final = mask_transf_nome & mask_credito
        
        df = df[mask_pagto | mask_transf_final].copy()
        
        if df.empty:
            return pd.DataFrame()

        df['Data_dt'] = df['Data'].apply(parse_br_date)
        df = df.dropna(subset=['Data_dt'])
        df['Data'] = df['Data_dt'].dt.strftime('%d/%m/%Y')
        
        def clean_val(x):
            try:
                if isinstance(x, (int, float)): return float(x)
                return float(str(x).replace('.', '').replace(',', '.'))
            except: return 0.0
        df['Valor_Razao'] = df['Valor_Razao'].apply(clean_val)
        
        termos_estorno = ["Est Pagto", "Est Pgto Ext"]
        mask_estorno = df['Info_AA'].astype(str).str.contains('|'.join(termos_estorno), case=False, na=False)
        df_estornos = df[mask_estorno].copy()
        df_normais = df[~mask_estorno].copy()
        
        indices_para_remover = []
        for valor, grupo_estorno in df_estornos.groupby('Valor_Razao'):
            matches_normais = df_normais[df_normais['Valor_Razao'] == valor]
            pares = min(len(grupo_estorno), len(matches_normais))
            if pares > 0:
                indices_para_remover.extend(matches_normais.head(pares).index.tolist())
        
        indices_para_remover.extend(df_estornos.index.tolist())
        df = df.drop(indices_para_remover).copy()
        
        lookup_pdf = {}
        for dt, group in df_pdf_ref.groupby('Data'):
            lookup_pdf[dt] = {}
            for doc_orig in group['Documento'].unique():
                doc_str = str(doc_orig)
                if doc_str == "Tarifas Bancárias":
                    lookup_pdf[dt]["TARIFA"] = doc_orig
                    continue
                doc_limpo = doc_str.lstrip('0')
                if doc_limpo: lookup_pdf[dt][doc_limpo] = doc_orig
        
        def encontrar_doc_inteligente(row):
            texto = str(row['Info_AB']).upper()
            dt = row['Data']
            if dt not in lookup_pdf: return "S/D"
            mapa_dia = lookup_pdf[dt]
            if "TARIFA" in texto and "TARIFA" in mapa_dia:
                return mapa_dia["TARIFA"]
            numeros_no_texto = re.findall(r'\d+', texto)
            for num in numeros_no_texto:
                num_limpo = num.lstrip('0')
                if num_limpo in mapa_dia:
                    return mapa_dia[num_limpo]
            return "NÃO LOCALIZADO"

        df['Documento'] = df.apply(encontrar_doc_inteligente, axis=1)
        return df[['Data', 'Documento', 'Valor_Razao']].reset_index(drop=True)

    except Exception as e:
        return pd.DataFrame()

# ==============================================================================
# 4. CONCILIAÇÃO INTELIGENTE
# ==============================================================================

def executar_conciliacao_inteligente(df_pdf, df_excel):
    df_pdf = df_pdf.copy()
    df_excel = df_excel.copy()
    resultados_finais = []
    
    indices_pdf_usados = set()
    indices_excel_usados = set()
    
    # 1. MATCH EXATO
    for idx_pdf, row_pdf in df_pdf.iterrows():
        if idx_pdf in indices_pdf_usados: continue
        candidatos = df_excel[
            (df_excel['Data'] == row_pdf['Data']) &
            (df_excel['Documento'] == row_pdf['Documento']) &
            (~df_excel.index.isin(indices_excel_usados))
        ]
        match = candidatos[abs(candidatos['Valor_Razao'] - row_pdf['Valor_Extrato']) < 0.01]
        
        if not match.empty:
            idx_excel = match.index[0]
            resultados_finais.append({
                'Data': row_pdf['Data'], 'Documento': row_pdf['Documento'],
                'Valor_Extrato': row_pdf['Valor_Extrato'], 'Valor_Razao': match.loc[idx_excel]['Valor_Razao'],
                'Diferença': 0.0, 'Obs': ''
            })
            indices_pdf_usados.add(idx_pdf)
            indices_excel_usados.add(idx_excel)

    # 2. MATCH POR VALOR
    for idx_pdf, row_pdf in df_pdf.iterrows():
        if idx_pdf in indices_pdf_usados: continue
        candidatos = df_excel[
            (df_excel['Data'] == row_pdf['Data']) &
            (~df_excel.index.isin(indices_excel_usados))
        ]
        match = candidatos[abs(candidatos['Valor_Razao'] - row_pdf['Valor_Extrato']) < 0.01]
        
        if not match.empty:
            idx_excel = match.index[0]
            resultados_finais.append({
                'Data': row_pdf['Data'], 'Documento': "Docs diferentes ou não encontrado",
                'Valor_Extrato': row_pdf['Valor_Extrato'], 'Valor_Razao': match.loc[idx_excel]['Valor_Razao'],
                'Diferença': 0.0, 'Obs': 'Doc Diferente'
            })
            indices_pdf_usados.add(idx_pdf)
            indices_excel_usados.add(idx_excel)

    # 3. AGRUPAMENTO
    df_excel_sobra = df_excel[~df_excel.index.isin(indices_excel_usados)].copy()
    df_pdf_sobra = df_pdf[~df_pdf.index.isin(indices_pdf_usados)].copy()
    
    if not df_excel_sobra.empty:
        g_excel = df_excel_sobra.groupby(['Data', 'Documento'])['Valor_Razao'].sum().reset_index()
    else: g_excel = pd.DataFrame(columns=['Data', 'Documento', 'Valor_Razao'])
    
    if not df_pdf_sobra.empty:
        g_pdf = df_pdf_sobra.groupby(['Data', 'Documento'])['Valor_Extrato'].sum().reset_index()
    else: g_pdf = pd.DataFrame(columns=['Data', 'Documento', 'Valor_Extrato'])

    df_sobras = pd.merge(g_pdf, g_excel, on=['Data', 'Documento'], how='outer')
    
    for _, row in df_sobras.iterrows():
        ve = row['Valor_Extrato'] if pd.notna(row['Valor_Extrato']) else 0.0
        vr = row['Valor_Razao'] if pd.notna(row['Valor_Razao']) else 0.0
        resultados_finais.append({
            'Data': row['Data'], 'Documento': row['Documento'],
            'Valor_Extrato': ve, 'Valor_Razao': vr,
            'Diferença': ve - vr, 'Obs': 'Agrupado'
        })
    
    df_final = pd.DataFrame(resultados_finais)
    if not df_final.empty:
        df_final['Data_dt'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y', errors='coerce')
        df_final = df_final.sort_values(by=['Data_dt', 'Documento'])
        
    return df_final

# ==============================================================================
# 5. GERAÇÃO PDF
# ==============================================================================

def gerar_pdf_final(df_final, output_filename, nome_conta_original):
    nome_conta_limpo = os.path.splitext(nome_conta_original)[0]
    titulo_doc = f"Conciliação {nome_conta_limpo}"

    doc = SimpleDocTemplate(
        output_filename, 
        pagesize=A4,
        rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm,
        title=titulo_doc
    )
    styles = getSampleStyleSheet()
    Story = []
    style_centered = ParagraphStyle(name='Centered', parent=styles['Normal'], alignment=1)
    
    Story.append(Paragraph("Relatório de Conciliação Bancária", styles["Title"]))
    Story.append(Spacer(1, 12))
    Story.append(Paragraph(f"<b>Conta:</b> {nome_conta_limpo}", style_centered))
    Story.append(Spacer(1, 12))
    
    if not df_final.empty:
        Story.append(Paragraph(f"<b>Período:</b> {df_final['Data'].iloc[0]} a {df_final['Data'].iloc[-1]}", style_centered))
    Story.append(Spacer(1, 24))
    
    headers = ['Data', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']
    data_list = [headers]
    
    t_ext = df_final['Valor_Extrato'].sum()
    t_raz = df_final['Valor_Razao'].sum()
    t_dif = df_final['Diferença'].sum()
    
    for _, row in df_final.iterrows():
        diff = row['Diferença']
        val_diff = formatar_moeda_br(diff)
        if abs(diff) < 0.01: val_diff = "-"
        line = [row['Data'], str(row['Documento']), formatar_moeda_br(row['Valor_Extrato']), formatar_moeda_br(row['Valor_Razao']), val_diff]
        data_list.append(line)
        
    data_list.append(['TOTAL', '', formatar_moeda_br(t_ext), formatar_moeda_br(t_raz), formatar_moeda_br(t_dif)])
    
    cw = [22*mm, 68*mm, 32*mm, 32*mm, 32*mm]
    t = Table(data_list, colWidths=cw)
    t_style = TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (1, 1), (1, -1), 'LEFT'),  
        ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
        ('SPAN', (0, -1), (1, -1)),
    ])
    
    for i, row_data in enumerate(data_list[1:-1], start=1):
        diff_val = df_final.iloc[i-1]['Diferença']
        if abs(diff_val) >= 0.01: t_style.add('TEXTCOLOR', (4, i), (4, i), colors.red)

    t.setStyle(t_style)
    Story.append(t)
    doc.build(Story)
    return True

# ==============================================================================
# 6. FUNÇÃO DE VISUALIZAÇÃO (IGUAL AO PDF)
# ==============================================================================

def preparar_tabela_visual(df_final):
    """
    Transforma o DataFrame de dados no formato EXATO do PDF para exibição na tela.
    (Formata moeda, adiciona Total, renomeia colunas).
    """
    # 1. Calcula totais
    t_ext = df_final['Valor_Extrato'].sum()
    t_raz = df_final['Valor_Razao'].sum()
    t_dif = df_final['Diferença'].sum()
    
    # 2. Cria cópia para formatação visual (strings)
    df_vis = df_final.copy()
    
    # 3. Formata colunas de valor
    df_vis['Valor_Extrato'] = df_vis['Valor_Extrato'].apply(formatar_moeda_br)
    df_vis['Valor_Razao'] = df_vis['Valor_Razao'].apply(formatar_moeda_br)
    
    # 4. Formata Diferença (com hífen se zero)
    df_vis['Diferença'] = df_vis['Diferença'].apply(lambda x: "-" if abs(x) < 0.01 else formatar_moeda_br(x))
    
    # 5. Cria linha de TOTAL
    row_total = pd.DataFrame([{
        'Data': 'TOTAL',
        'Documento': '',
        'Valor_Extrato': formatar_moeda_br(t_ext),
        'Valor_Razao': formatar_moeda_br(t_raz),
        'Diferença': formatar_moeda_br(t_dif)
    }])
    
    # 6. Seleciona e ordena colunas (Igual ao PDF)
    # Obs: Removemos colunas auxiliares que não vão para o PDF
    df_vis = df_vis[['Data', 'Documento', 'Valor_Extrato', 'Valor_Razao', 'Diferença']]
    
    # 7. Adiciona Total ao final
    df_vis = pd.concat([df_vis, row_total], ignore_index=True)
    
    # 8. Renomeia Cabeçalhos
    df_vis.columns = ['Data', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']
    
    return df_vis

# ==============================================================================
# 7. FUNÇÃO PRINCIPAL DO APLICATIVO WEB
# ==============================================================================
def main():
    st.title("🏦 Conciliador Bancário Profissional")
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("1. Extrato (PDF)", type="pdf")
    with col2:
        xlsx_file = st.file_uploader("2. Razão (Excel/CSV)", type=["xlsx", "csv"])

    if st.button("🚀 Processar Conciliação", type="primary"):
        if pdf_file and xlsx_file:
            with st.spinner("Analisando dados e cruzando informações..."):
                try:
                    # Salva arquivos temporários
                    with open("temp.pdf", "wb") as f: f.write(pdf_file.getbuffer())
                    
                    ext = "xlsx" if xlsx_file.name.endswith('xlsx') else "csv"
                    with open(f"temp.{ext}", "wb") as f: f.write(xlsx_file.getbuffer())

                    # Processamento
                    df_pdf = processar_pdf("temp.pdf")
                    df_excel = processar_excel_detalhado(f"temp.{ext}", df_pdf)
                    
                    if df_excel.empty:
                        st.error("Erro no Excel: Verifique se existem pagamentos na Coluna Z ou transferências válidas.")
                    elif df_pdf.empty:
                         st.error("Erro no PDF: Não foi possível ler os lançamentos.")
                    else:
                        df_final = executar_conciliacao_inteligente(df_pdf, df_excel)
                        
                        # GERA PDF (para download)
                        nome_conta = pdf_file.name
                        nome_limpo = os.path.splitext(nome_conta)[0]
                        pdf_out = f"Conciliacao_{nome_limpo}.pdf"
                        gerar_pdf_final(df_final, pdf_out, nome_conta)

                        # --- EXIBIÇÃO NA TELA (IGUAL AO PDF) ---
                        st.success("✅ Relatório Gerado com Sucesso!")
                        
                        # Prepara e exibe tabela visual com Total
                        df_visual = preparar_tabela_visual(df_final)
                        st.dataframe(df_visual, use_container_width=True, hide_index=True)

                        # --- DOWNLOAD ---
                        with open(pdf_out, "rb") as f:
                            st.download_button(
                                label="📥 Baixar PDF Final",
                                data=f,
                                file_name=pdf_out,
                                mime="application/pdf"
                            )

                except Exception as e:
                    st.error(f"Ocorreu um erro técnico: {e}")
        else:
            st.warning("Por favor, faça upload dos dois arquivos.")

# --- EXECUÇÃO ---
if __name__ == "__main__":
    if check_password():
        main()