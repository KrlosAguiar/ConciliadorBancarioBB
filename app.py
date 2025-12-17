import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Conciliador Bancário - Banco do Brasil", layout="wide")

# --- ESTILO CSS PARA ALINHAMENTOS ---
st.markdown("""
    <style>
    .reportview-container .main .block-container { padding-top: 2rem; }
    th { background-color: #1B5E20 !important; color: white !important; text-align: center !important; }
    td:nth-child(1), td:nth-child(3) { text-align: center !important; } /* Data e Documento */
    td:nth-child(2) { text-align: left !important; }   /* Histórico */
    td:nth-child(4), td:nth-child(5), td:nth-child(6) { text-align: right !important; } /* Valores */
    </style>
""", unsafe_allow_html=True)

# --- SEGURANÇA ---
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
        else:
            st.error("Chave incorreta!")
    return False

# ==============================================================================
# LÓGICA ORIGINAL ADAPTADA
# ==============================================================================

CURRENT_YEAR = str(datetime.datetime.now().year)

def limpar_documento_pdf(doc_str):
    if not doc_str: return ""
    apenas_digitos = re.sub(r'\D', '', str(doc_str))
    if not apenas_digitos: return ""
    return apenas_digitos[-6:] if len(apenas_digitos) > 6 else apenas_digitos

def parse_br_date(date_val):
    if pd.isna(date_val): return pd.NaT
    try:
        return pd.to_datetime(date_val, dayfirst=True, errors='coerce')
    except:
        return pd.to_datetime(date_val, errors='coerce')

def processar_pdf(file_bytes):
    rows_debitos = []
    rows_devolucoes = [] 
    
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
                        if any(x in hist_upper for x in ["TED DEVOLVIDA", "DEVOLUCAO DE TED", "TED DEVOL"]):
                            rows_devolucoes.append({"Data": data_str, "Valor_Extrato": valor_float})

    df_debitos = pd.DataFrame(rows_debitos)
    if df_debitos.empty: return df_debitos

    # Filtro de Devoluções (Lógica Original)
    if rows_devolucoes:
        df_dev = pd.DataFrame(rows_devolucoes)
        idx_rem = []
        for _, r in df_dev.iterrows():
            m = df_debitos[(df_debitos['Data']==r['Data']) & (abs(df_debitos['Valor_Extrato']-r['Valor_Extrato'])<0.01) & (~df_debitos.index.isin(idx_rem))]
            if not m.empty: idx_rem.append(m.index[0])
        df_debitos = df_debitos.drop(idx_rem).reset_index(drop=True)

    termos_excluir = "SALDO|S A L D O|Resgate|BB-APLIC C\.PRZ-APL\.AUT"
    df = df_debitos[~df_debitos['Histórico'].str.contains(termos_excluir, case=False, regex=True)].copy()
    
    # Agrupamento de Tarifas (Lógica Original)
    df['Data_dt'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    mask_13113 = df['Histórico'].str.contains("13113", na=False)
    df_tarifas = df[mask_13113].copy()
    df_outros = df[~mask_13113].copy()

    if not df_tarifas.empty:
        df_t_agg = df_tarifas.groupby('Data_dt').agg({'Valor_Extrato':'sum', 'Data':'first'}).reset_index()
        df_t_agg['Documento'] = "Tarifas"
        df_t_agg['Histórico'] = "Tarifas Bancárias do Dia"
        df = pd.concat([df_outros, df_t_agg], ignore_index=True)
    
    return df.sort_values('Data_dt').drop(columns=['Data_dt'])

def processar_excel(file_bytes, df_pdf_ref, is_csv=False):
    if is_csv:
        df = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='latin1', sep=None, engine='python')
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    
    # Seleção de Colunas (Filtro Z Original)
    df = df.iloc[:, [4, 5, 8, 25, 26, 27]].copy()
    df.columns = ['Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
    
    mask_pagto = df['Info_Z'].astype(str).str.contains("Pagamento", case=False, na=False)
    mask_transf = (df['Info_Z'].astype(str).str.contains("TRANSFERENCIA", case=False, na=False)) & (df['DC'].str.strip().str.upper() == 'C')
    df = df[mask_pagto | mask_transf].copy()
    
    df['Data_dt'] = df['Data'].apply(parse_br_date)
    df = df.dropna(subset=['Data_dt'])
    df['Data'] = df['Data_dt'].dt.strftime('%d/%m/%Y')
    df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.','').replace(',','.')) if isinstance(x, str) else float(x))
    
    return df[['Data', 'Valor_Razao', 'Info_AB']]

# ==============================================================================
# INTERFACE STREAMLIT
# ==============================================================================

if check_password():
    st.title("🏦 Conciliador Bancário - Banco do Brasil")
    
    up_pdf = st.file_uploader("1. Extrato (PDF)", type="pdf")
    up_xlsx = st.file_uploader("2. Razão (Excel/CSV)", type=["xlsx", "csv"])

    if up_pdf and up_xlsx:
        df_p = processar_pdf(up_pdf.read())
        df_e = processar_excel(up_xlsx.read(), df_p, is_csv=up_xlsx.name.endswith('csv'))
        
        st.subheader("📋 Relatório de Conciliação")
        
        # Formatação de Moeda e Exibição
        df_p['Valor_Extrato'] = df_p['Valor_Extrato'].map('R$ {:,.2f}'.format)
        
        # Reorganizando Colunas para o padrão pedido: Data, Histórico, Documento, Valor
        df_display = df_p[['Data', 'Histórico', 'Documento', 'Valor_Extrato']]
        
        # Exibição com alinhamento
        st.table(df_display)

        # Botão para Download
        output = io.BytesIO()
        df_p.to_excel(output, index=False)
        st.download_button("📥 Baixar Resultado (Excel)", output.getvalue(), "conciliacao.xlsx")
