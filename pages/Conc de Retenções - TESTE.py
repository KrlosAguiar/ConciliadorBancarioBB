import streamlit as st
import pandas as pd
import io
import os
import re
import unicodedata
from datetime import datetime
from PIL import Image

# Bibliotecas para geração do PDF (ReportLab)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ==============================================================================
# 0. CONFIGURAÇÃO DA PÁGINA E CSS
# ==============================================================================

icon_filename = "Barcarena.png"
current_dir = os.getcwd()
possible_paths = [
    os.path.join(current_dir, icon_filename),
    os.path.join(os.path.dirname(__file__), icon_filename),
    os.path.join(current_dir, "pages", icon_filename),
    os.path.join("..", icon_filename)
]
icon_image = None
for p in possible_paths:
    if os.path.exists(p):
        try:
            icon_image = Image.open(p)
            break
        except: pass

try:
    if icon_image:
        st.set_page_config(page_title="Conciliador de Retenções", page_icon=icon_image, layout="wide")
    else:
        st.set_page_config(page_title="Conciliador de Retenções", layout="wide")
except: pass

st.markdown("""
<style>
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
    
    /* CSS DOS BOTÕES (Mantido igual) */
    div.stButton > button, 
    div[data-testid="stForm"] button {
        background-color: rgb(38, 39, 48) !important;
        color: white !important;
        font-weight: bold !important;
        border: 1px solid rgb(60, 60, 60);
        border-radius: 5px;
        font-size: 16px;
        transition: 0.3s;
        height: 50px; 
        margin-top: 10px;
    }
    div.stButton > button:hover, 
    div[data-testid="stForm"] button:hover {
        background-color: rgb(20, 20, 25) !important;
        border-color: white;
    }

    .big-label { font-size: 24px !important; font-weight: 600 !important; margin-bottom: 10px; }
    
    /* Cards de Resumo (Mantido igual) */
    .metric-card {
        background-color: #f8f9fa;
        border-left: 5px solid #ff4b4b;
        padding: 15px;
        border-radius: 5px;
        color: black;
        border: 1px solid #ddd;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin-bottom: 10px;
    }
    .metric-card-green { border-left: 5px solid #28a745; }
    .metric-card-orange { border-left: 5px solid #ffc107; }
    .metric-card-blue { border-left: 5px solid #007bff; }
    .metric-card-dark { border-left: 5px solid #343a40; }
    .metric-value { font-size: 22px; font-weight: bold; }
    .metric-label { font-size: 13px; color: #555; text-transform: uppercase; letter-spacing: 0.5px; }

    /* --- TRUQUE PARA AUMENTAR A TABELA --- */
    
    div[data-testid="stDataEditor"] {
        /* Aumenta o tamanho visual de tudo na tabela em 30% */
        zoom: 1.3 !important;
        
        /* Tenta inverter as cores para Fundo Claro / Texto Escuro */
        /* Se as cores verde/vermelho ficarem estranhas, remova a linha abaixo */
        filter: invert(1) hue-rotate(180deg) !important; 
        
        /* Margem para não colar nos lados quando der zoom */
        margin-top: 10px;
        margin-bottom: 10px;
    }

    /* Ajuste para quando você clica para digitar (Input) */
    div[data-testid="stDataEditor"] input {
        font-size: 18px !important;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. LISTAS FIXAS DE OPÇÕES E CONFIGURAÇÕES
# ==============================================================================

LISTA_UGS = [
    "0 - PMB",
    "3 - FMS",
    "4 - FMAS",
    "5 - FMDCA",
    "6 - FME",
    "7 - FMMA",
    "10 - ARSEP",
    "11 - FMDPI",
    "12 - SEMER",
    "9999 - CONSOLIDADO"
]

LISTA_CONTAS = [
    "7812 - Salario Maternidade",
    "7814 - PENSÃO ALIMENTICIA",
    "7815 - UNIMED Belem Coop. de Trabalho Medico",
    "7816 - UNIODONTO - Coop. de Trab. Odontologico",
    "7817 - ODONTOPREV",
    "7819 - Sindicato dos Trabalhadores em Educação",
    "7821 - Sind. dos Agentes de Vig. de Barcarena",
    "7824 - Emp. Consignado BANPARA",
    "7826 - Emp. Cons. CAIXA ECONOMICA FEDERAL",
    "7827 - Emp. Cons. BANCO DO BRASIL",
    "7828 - Emp. Consignado SANTANDER",
    "7831 - A.M.P.E Barcarena - Di",
    "7832 - Desc. Autorizado PSDB 3%",
    "7837 - Desc. Aut. ASPEB",
    "7845 - IRRF DE SERVIÇOS DE TERCEIROS PJ",
    "7846 - IRRF DE SERV. DA ADM. DIR. E INDIRETA",
    "7847 - IRPF - Imposto de Renda da Pessoa Fisica",
    "7852 - ISS de Pessoa Juridica Retido na Fonte",
    "7853 - ISS De Pessoa Fisica Retido na Fonte",
    "7857 - INSS - Pessoa Fisica",
    "7858 - INSS FOPAG EFETIVOS",
    "7859 - INSS FOPAG TEMPORARIOS E COMISSIONADOS",
    "7864 - SALARIO FAMILIA",
    "7865 - INSS - Pessoa Juridica",
    "8926 - GARANTIA DE SAÚDE",
    "8931 - Emp. Consignado BRADESCO",
    "9032 - SEPUB -SINDICATO DOS SERVIDORES CIVIS DO PARÁ E MUNICIPIOS",
    "9159 - Emp. Cons.Kardbank",
    "9160 - Emp. Cons. Fydigital",
    "9210 - Emp. Consignado HBI - Scd"
]

MAPA_MESES = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARCO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}

# ==============================================================================
# 2. FUNÇÕES DE PROCESSAMENTO
# ==============================================================================

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == "": return "R$ 0,00"
    try:
        val_float = float(valor)
        return f"R$ {val_float:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    except:
        return "R$ 0,00"

def converter_moeda_input(texto):
    if not texto: return 0.0
    try:
        if isinstance(texto, (int, float)): return float(texto)
        texto = str(texto).replace('R$', '').replace(' ', '')
        return float(texto.replace('.', '').replace(',', '.'))
    except:
        return 0.0

def formatar_data(dt):
    if pd.isna(dt): return "-"
    try:
        return dt.strftime("%d/%m/%Y")
    except:
        return str(dt)

def limpar_nome_arquivo(texto):
    nfkd_form = unicodedata.normalize('NFKD', str(texto))
    texto_sem_acento = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    texto_limpo = re.sub(r'[\\/*?:"<>|]', '_', texto_sem_acento)
    return texto_limpo.strip()

def normalizar_texto(texto):
    if not isinstance(texto, str): return ""
    nfkd_form = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).upper()

def verificar_compatibilidade_mes(hist_pagamento, data_retencao):
    if pd.isna(data_retencao):
        return True 
    hist_norm = normalizar_texto(hist_pagamento)
    mes_retencao_num = data_retencao.month
    mes_retencao_nome = MAPA_MESES[mes_retencao_num]
    
    meses_encontrados = []
    for mes_num, mes_nome in MAPA_MESES.items():
        if mes_nome in hist_norm:
            meses_encontrados.append(mes_nome)
            
    if not meses_encontrados:
        return True 
    if mes_retencao_nome in meses_encontrados:
        return True
    return False

@st.cache_data(show_spinner=False)
def carregar_dados(file):
    try:
        df = pd.read_excel(file, header=None)
    except:
        file.seek(0)
        try:
            df = pd.read_csv(file, sep=None, engine='python', encoding='latin1', header=None, on_bad_lines='skip')
        except:
            df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8', header=None, on_bad_lines='skip')
    
    min_cols = 35 
    if df.shape[1] < min_cols:
        for i in range(df.shape[1], min_cols):
            df[i] = pd.NA

    if 5 in df.columns:
        df[5] = df[5].astype(str).str.strip().str.upper()
        df = df[df[5].isin(['C', 'D'])]
    
    col_valor_idx = 8 
    def converter_valor(val):
        try:
            if isinstance(val, str):
                return float(val.replace('.', '').replace(',', '.'))
            return float(val)
        except:
            return 0.0

    if col_valor_idx in df.columns:
        df[col_valor_idx] = df[col_valor_idx].apply(converter_valor)
    
    col_data_idx = 4
    if col_data_idx in df.columns:
        df['Data_Dt'] = pd.to_datetime(df[col_data_idx], dayfirst=True, errors='coerce')

    return df

def identificar_colunas_dinamicas(df):
    mapa = {'empenho': 14, 'tipo': 19, 'hist': 21}
    coluna_ab_tem_dados = df[27].head(50).notna().sum() > 0
    if coluna_ab_tem_dados: mapa['hist'] = 27 
    else: mapa['hist'] = 21

    for idx, row in df.head(50).iterrows():
        for c in range(12, 16): 
            val = str(row[c]).strip()
            if re.match(r'^\d{4}/\d+$', val):
                mapa['empenho'] = c
                break
        start_tipo = mapa['empenho'] + 1
        for c in range(start_tipo, start_tipo + 10):
            val = str(row[c]).upper()
            if any(k in val for k in ["LIQUIDAÇÃO", "PAGAMENTO", "LANÇAMENTO", "RETENÇÃO", "ESTORNO"]):
                mapa['tipo'] = c
                break
    return mapa

def sanitizar_historico(val):
    if pd.isna(val) or str(val).lower().strip() == 'nan':
        return ""
    return str(val).strip()

def inserir_subtotais_diarios(df):
    """
    Insere linhas de subtotal diário para os itens com Status 'Retido s/ Pagto'.
    """
    if df.empty: return df
    
    df_ret = df[df['Status'] == 'Retido s/ Pagto'].copy()
    df_others = df[df['Status'] != 'Retido s/ Pagto'].copy()
    
    if df_ret.empty:
        return df
    
    # Ordena por data (o campo _dt_sort é datetime e deve existir aqui)
    df_ret = df_ret.sort_values(by=['_dt_sort'])
    
    new_rows = []
    df_ret['temp_date'] = df_ret['_dt_sort'].dt.date
    datas_unicas = df_ret['temp_date'].dropna().unique()
    
    for dt in datas_unicas:
        df_dia = df_ret[df_ret['temp_date'] == dt]
        
        for _, row in df_dia.iterrows():
            new_rows.append(row.drop('temp_date'))
            
        subtotal = df_dia['Vlr Retido'].sum()
        
        # Cria linha de subtotal
        row_sub = df_dia.iloc[0].drop('temp_date').copy()
        row_sub['Empenho'] = "TOTAL"
        row_sub['Data Emp'] = formatar_data(pd.to_datetime(dt))
        row_sub['Vlr Retido'] = subtotal
        row_sub['Vlr Pago'] = 0.0 
        row_sub['Dif'] = 0.0 
        row_sub['Data Pag'] = "-"
        row_sub['Histórico'] = "-"
        row_sub['Status'] = "SUBTOTAL"
        
        new_rows.append(row_sub)
        
    df_ret_sem_data = df_ret[df_ret['temp_date'].isna()]
    if not df_ret_sem_data.empty:
        for _, row in df_ret_sem_data.iterrows():
            new_rows.append(row.drop('temp_date'))

    df_ret_final = pd.DataFrame(new_rows)
    # Reconcatena
    return pd.concat([df_ret_final, df_others], ignore_index=True)

def preparar_dados_resumo_superior(df):
    """
    Prepara os dados para o novo relatório resumido.
    1. Subtotais diários de 'Retido s/ Pagto'.
    2. Valores individuais de 'Pago s/ Retenção'.
    3. Subtotais diários de 'Conciliado'.
    """
    if df.empty: return pd.DataFrame()
    
    # 1. Subtotais Diários de Retido s/ Pagto
    df_ret = df[df['Status'] == 'Retido s/ Pagto'].copy()
    rows_ret = []
    if not df_ret.empty:
        df_ret['temp_date'] = df_ret['_dt_sort'].dt.date
        dates = df_ret['temp_date'].unique()
        for dt in dates:
            d = df_ret[df_ret['temp_date'] == dt]
            # CALCULA A SOMA DE DIFERENÇA (Para itens retidos s/ pagto, Dif = Vlr Retido)
            soma_dif = d['Dif'].sum()
            
            rows_ret.append({
                "Empenho": "Retenção", 
                "Data": formatar_data(pd.to_datetime(dt)),
                "Vlr Retido": d['Vlr Retido'].sum(),
                "Vlr Pago": 0.0,
                "Dif": soma_dif, # Usa a soma real das diferenças
                "Histórico": "-",
                "Status": "Retido s/ Pagto",
                "_dt_sort": pd.to_datetime(dt),
                "_sort_order": 1
            })
            
    # 2. Itens individuais de Pago s/ Retenção (como no relatório atual)
    df_pag = df[df['Status'] == 'Pago s/ Retenção'].copy()
    rows_pag = []
    for _, row in df_pag.iterrows():
        rows_pag.append({
            "Empenho": row['Empenho'],
            "Data": row['Data Pag'], 
            "Vlr Retido": 0.0,
            "Vlr Pago": row['Vlr Pago'],
            "Dif": row['Dif'],
            "Histórico": row['Histórico'],
            "Status": "Pago s/ Retenção",
            "_dt_sort": row['_dt_sort'],
            "_sort_order": 2
        })

    # 3. Subtotais Diários de Conciliados
    df_conc = df[df['Status'] == 'Conciliado'].copy()
    rows_conc = []
    if not df_conc.empty:
        df_conc['temp_date'] = df_conc['_dt_sort'].dt.date
        dates = df_conc['temp_date'].unique()
        for dt in dates:
            d = df_conc[df_conc['temp_date'] == dt]
            # CALCULA A SOMA DE DIFERENÇA
            soma_dif = d['Dif'].sum()
            
            rows_conc.append({
                "Empenho": "Conciliado", 
                "Data": formatar_data(pd.to_datetime(dt)),
                "Vlr Retido": d['Vlr Retido'].sum(),
                "Vlr Pago": d['Vlr Pago'].sum(),
                "Dif": soma_dif, # Usa a soma real das diferenças
                "Histórico": "-",
                "Status": "Conciliado", 
                "_dt_sort": pd.to_datetime(dt),
                "_sort_order": 3
            })

    all_rows = rows_ret + rows_pag + rows_conc
    if not all_rows: return pd.DataFrame()
    
    df_final = pd.DataFrame(all_rows)
    # Ordena PRIMEIRO pelo grupo (Retido -> Pago -> Conciliado) e DEPOIS pela Data
    df_final = df_final.sort_values(by=['_sort_order', '_dt_sort'])
    return df_final

def processar_conciliacao(df, ug_sel, conta_sel, saldo_anterior_val):
    cod_ug = ug_sel.split(' - ')[0].strip()
    cod_conta = conta_sel.split(' - ')[0].strip()

    cols_map = identificar_colunas_dinamicas(df)
    c_ug, c_status, c_data, c_dc, c_conta, c_valor = 0, 2, 4, 5, 6, 8
    c_empenho, c_tipo, c_hist = cols_map['empenho'], cols_map['tipo'], cols_map['hist']

    colunas_para_preencher = [c_ug, c_status, c_data, c_conta, c_empenho, c_tipo]
    for col in colunas_para_preencher:
        if col < df.shape[1]:
            df[col] = df[col].ffill()

    if cod_ug == '9999':
        mask_ug = pd.Series(True, index=df.index)
    else:
        mask_ug = df[c_ug].astype(str).str.split('.').str[0] == str(cod_ug)
        
    mask_conta = df[c_conta].astype(str).str.startswith(str(cod_conta))
    
    df_base = df[mask_ug & mask_conta].copy()
    
    resumo = {
        "ret_pendente": 0, "val_ret_pendente": 0.0,
        "pag_sobra": 0,    "val_pag_sobra": 0.0,
        "ok": 0,           "val_ok": 0.0,
        "tot_ret": 0.0, "tot_pag": 0.0, "saldo": saldo_anterior_val
    }

    if df_base.empty: return pd.DataFrame(), resumo

    df_base['Tipo_Norm'] = df_base[c_tipo].astype(str).str.strip()
    df_base['Status_Lanc'] = df_base[c_status].astype(str).str.strip()

    mask_ret = (df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'C')
    df_ret = df_base[mask_ret].copy()
    
    mask_estorno_ret = (df_base[c_dc] == 'D') & (
        df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False) | 
        df_base['Status_Lanc'].str.contains("Estorno", case=False)
    )
    df_est_ret = df_base[mask_estorno_ret].copy()
    
    mask_pag = (df_base['Tipo_Norm'].str.contains("Pagamento", case=False)) & (df_base[c_dc] == 'D')
    df_pag = df_base[mask_pag].copy()

    condicao_credito = (df_base[c_dc] == 'C')
    condicao_nome_estorno = (
        df_base['Status_Lanc'].str.contains("Estorno", case=False) | 
        df_base['Tipo_Norm'].str.contains("Estorno", case=False) | 
        df_base[c_hist].astype(str).str.contains("Estorno", case=False)
    )
    
    mask_estorno_pag = condicao_credito & condicao_nome_estorno
    df_est_pag = df_base[mask_estorno_pag].copy()

    idx_ret_cancel = set()
    for _, r_est in df_est_ret.iterrows():
        v, e = r_est[c_valor], r_est[c_empenho]
        cand = df_ret[(df_ret[c_empenho] == e) & (abs(df_ret[c_valor] - v) < 0.01) & (~df_ret.index.isin(idx_ret_cancel))]
        if not cand.empty: idx_ret_cancel.add(cand.index[0])
    df_ret_limpa = df_ret[~df_ret.index.isin(idx_ret_cancel)]

    idx_pag_cancel = set()
    for _, r_est in df_est_pag.iterrows():
        v, e = r_est[c_valor], r_est[c_empenho]
        cand = df_pag[(df_pag[c_empenho] == e) & (abs(df_pag[c_valor] - v) < 0.01) & (~df_pag.index.isin(idx_pag_cancel))]
        if not cand.empty: idx_pag_cancel.add(cand.index[0])
    
    df_pag_limpa = df_pag[~df_pag.index.isin(idx_pag_cancel)]

    resultados = []
    idx_pag_usado = set()
    
    for _, r in df_ret_limpa.iterrows():
        val = r[c_valor]
        dt_retencao = r['Data_Dt']
        
        condicao_valor = (df_pag_limpa[c_valor] == val)
        condicao_usado = (~df_pag_limpa.index.isin(idx_pag_usado))
        
        if pd.notna(dt_retencao):
            condicao_data = (df_pag_limpa['Data_Dt'] >= dt_retencao) | (df_pag_limpa['Data_Dt'].isna())
            cand = df_pag_limpa[condicao_valor & condicao_data & condicao_usado]
        else:
            cand = df_pag_limpa[condicao_valor & condicao_usado]
        
        val_pago, dt_pag_str, match, sort = 0.0, "-", False, 0
        dt_pag_sort = pd.NaT 
        hist_final = sanitizar_historico(r[c_hist])
        
        r_pag = None
        match_encontrado = False
        
        if not cand.empty:
            for idx_c, row_c in cand.iterrows():
                hist_candidato = sanitizar_historico(row_c[c_hist])
                if verificar_compatibilidade_mes(hist_candidato, dt_retencao):
                    r_pag = row_c
                    match_encontrado = True
                    break
            
            if match_encontrado:
                val_pago = r_pag[c_valor]
                dt_real_pag = r_pag[c_data]
                dt_pag_sort = r['Data_Dt']
                if pd.notna(dt_real_pag):
                    dt_pag_str = formatar_data(dt_real_pag)
                
                hist_pag = sanitizar_historico(r_pag[c_hist])
                if hist_pag: 
                    hist_final = hist_pag
                
                idx_pag_usado.add(r_pag.name)
                match, sort = True, 2
                resumo["ok"] += 1
                resumo["val_ok"] += val_pago
            else:
                resumo["ret_pendente"] += 1
                resumo["val_ret_pendente"] += val
        else:
            resumo["ret_pendente"] += 1
            resumo["val_ret_pendente"] += val
            
        resultados.append({
            "Empenho": r[c_empenho], 
            "Data Emp": formatar_data(r[c_data]), 
            "Vlr Retido": val, 
            "Vlr Pago": val_pago,
            "Dif": val - val_pago, 
            "Data Pag": dt_pag_str, 
            "Histórico": hist_final, 
            "_sort": sort,
            "_dt_sort": r['Data_Dt'], 
            "Status": "Conciliado" if match else "Retido s/ Pagto"
        })

    for _, r in df_pag_limpa[~df_pag_limpa.index.isin(idx_pag_usado)].iterrows():
        resumo["pag_sobra"] += 1
        resumo["val_pag_sobra"] += r[c_valor]
        
        resultados.append({
            "Empenho": r[c_empenho], 
            "Data Emp": "-", 
            "Vlr Retido": 0.0, 
            "Vlr Pago": r[c_valor],
            "Dif": 0.0 - r[c_valor], 
            "Data Pag": formatar_data(r[c_data]), 
            "Histórico": sanitizar_historico(r[c_hist]), 
            "_sort": 1, 
            "_dt_sort": r['Data_Dt'], 
            "Status": "Pago s/ Retenção"
        })

    if not resultados: return pd.DataFrame(), resumo

    # Retorna o DataFrame contendo _dt_sort
    df_res = pd.DataFrame(resultados).sort_values(by=['_sort', '_dt_sort'])
    
    # Resumo com dados brutos
    resumo["tot_ret"] = df_res["Vlr Retido"].sum()
    resumo["tot_pag"] = df_res["Vlr Pago"].sum()
    diferenca_tabela = df_res["Dif"].sum()
    resumo["saldo"] = diferenca_tabela + saldo_anterior_val
    
    return df_res, resumo

def gerar_excel(df, resumo, saldo_anterior, ug, conta):
    out = io.BytesIO()
    # Remove colunas auxiliares
    df_exp = df.drop(columns=['_sort', '_dt_sort', 'Status'], errors='ignore')
    
    start_row_table = 8 
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df_exp.to_excel(writer, sheet_name='Conciliacao', index=False, startrow=start_row_table)
        wb = writer.book
        ws = writer.sheets['Conciliacao']
        
        fmt_head = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_head_filter = wb.add_format({'bold': True, 'bg_color': '#E6E6E6', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
        fmt_center = wb.add_format({'align': 'center', 'valign': 'vcenter'})
        fmt_money_center = wb.add_format({'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
        fmt_green = wb.add_format({'font_color': '#006400', 'bold': True, 'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
        fmt_red = wb.add_format({'font_color': '#FF0000', 'bold': True, 'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
        fmt_hist = wb.add_format({'text_wrap': True, 'valign': 'vcenter', 'font_size': 10, 'align': 'left'})
        
        fmt_tot_label = wb.add_format({'bold': True, 'bg_color': '#E6E6E6', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
        fmt_tot_val_green = wb.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'font_color': '#006400', 'valign': 'vcenter'})
        fmt_tot_val_red = wb.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'font_color': '#FF0000', 'valign': 'vcenter'})
        fmt_tot_val = wb.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'valign': 'vcenter'})
        fmt_card_label = wb.add_format({'bold': True, 'bg_color': '#f0f0f0', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
        fmt_card_header_center = wb.add_format({'bold': True, 'bg_color': '#f0f0f0', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_card_qtd = wb.add_format({'bold': True, 'num_format': '0', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_card_money = wb.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'valign': 'vcenter'})

        ws.merge_range('A1:G1', f"UG: {ug}  |  CONTA: {conta}", fmt_head_filter)
        ws.merge_range('A3:C3', 'RESUMO POR SITUAÇÃO (CARDS)', fmt_head)
        ws.write(3, 0, "CATEGORIA", fmt_card_label)
        ws.write(3, 1, "QTD", fmt_card_header_center)
        ws.write(3, 2, "VALOR", fmt_card_header_center)
        cat_names = ["Retido s/ Pagto (Pendente)", "Pago s/ Retenção (Sobra)", "Conciliados (OK)"]
        ws.write(4, 0, cat_names[0], fmt_card_label)
        ws.write(4, 1, resumo['ret_pendente'], fmt_card_qtd)
        ws.write(4, 2, resumo['val_ret_pendente'], fmt_card_money)
        ws.write(5, 0, cat_names[1], fmt_card_label)
        ws.write(5, 1, resumo['pag_sobra'], fmt_card_qtd)
        ws.write(5, 2, resumo['val_pag_sobra'], fmt_card_money)
        ws.write(6, 0, cat_names[2], fmt_card_label)
        ws.write(6, 1, resumo['ok'], fmt_card_qtd)
        ws.write(6, 2, resumo['val_ok'], fmt_card_money)
        
        ws.merge_range('E3:F3', 'RESUMO FINANCEIRO (TOTAIS)', fmt_head)
        ws.write(3, 4, "SALDO ANTERIOR", fmt_tot_label)
        ws.write(3, 5, saldo_anterior, fmt_tot_val)
        ws.write(4, 4, "TOTAL RETIDO", fmt_tot_label)
        ws.write(4, 5, resumo['tot_ret'], fmt_tot_val_green)
        ws.write(5, 4, "TOTAL PAGO", fmt_tot_label)
        ws.write(5, 5, resumo['tot_pag'], fmt_tot_val_red)
        
        fmt_saldo_final = fmt_tot_val_red if resumo['saldo'] > 0.01 else (fmt_tot_val_green if resumo['saldo'] < -0.01 else fmt_tot_val)
        ws.write(6, 4, "SALDO A PAGAR", fmt_tot_label)
        ws.write(6, 5, resumo['saldo'], fmt_saldo_final)

        for i, col in enumerate(df_exp.columns):
            ws.write(start_row_table, i, col, fmt_head)
            if i == 6: ws.set_column(i, i, 50, fmt_hist) 
            elif i in [2, 3, 4]: ws.set_column(i, i, 18, fmt_money_center)
            else: ws.set_column(i, i, 15, fmt_center)
        
        first_row = start_row_table + 1
        last_row = start_row_table + len(df_exp)
        
        ws.conditional_format(first_row, 4, last_row, 4, {'type': 'cell', 'criteria': '>', 'value': 0.001, 'format': fmt_red})
        ws.conditional_format(first_row, 4, last_row, 4, {'type': 'cell', 'criteria': '<', 'value': -0.001, 'format': fmt_green})
        
    return out.getvalue()

def gerar_pdf(df_f, ug, conta, resumo, saldo_anterior):
    buffer = io.BytesIO()
    doc_title = f"Conciliação de Retenções - UG {ug} - Retenção {conta}"
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm, title=doc_title
    )
    story = []
    styles = getSampleStyleSheet()
    style_hist = ParagraphStyle(name='SmallHist', parent=styles['Normal'], fontSize=6, leading=7, alignment=TA_LEFT)
    story.append(Paragraph("Relatório de Conciliação de Retenções", styles["Title"]))
    filtro_text = f"<b>UG:</b> {ug}  |  <b>CONTA:</b> {conta}"
    story.append(Paragraph(filtro_text, ParagraphStyle(name='C', alignment=1, spaceAfter=10)))
    bg_red = colors.Color(1, 0.9, 0.9)
    bg_org = colors.Color(1, 0.95, 0.8)
    bg_grn = colors.Color(0.9, 1, 0.9)
    bg_blu = colors.Color(0.9, 0.95, 1)
    
    data_resumo = [
        ["PENDENTES (RETIDO S/ PGTO)", "SOBRAS (PAGO S/ RETENÇÃO)", "CONCILIADOS"],
        [f"{resumo['ret_pendente']} itens", f"{resumo['pag_sobra']} itens", f"{resumo['ok']} itens"],
        [formatar_moeda_br(resumo['val_ret_pendente']), formatar_moeda_br(resumo['val_pag_sobra']), formatar_moeda_br(resumo['val_ok'])]
    ]
    t_res = Table(data_resumo, colWidths=[63*mm, 63*mm, 63*mm])
    t_res.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (-1,0), 9), ('FONTSIZE', (0,1), (-1,2), 11),
        ('BACKGROUND', (0,0), (0,-1), bg_red), ('BACKGROUND', (1,0), (1,-1), bg_org), ('BACKGROUND', (2,0), (2,-1), bg_grn),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_res)
    story.append(Spacer(1, 3*mm))

    data_totais = [
        ["SALDO ANTERIOR", "TOTAL RETIDO", "TOTAL PAGO", "SALDO A PAGAR"],
        [formatar_moeda_br(saldo_anterior), formatar_moeda_br(resumo['tot_ret']), formatar_moeda_br(resumo['tot_pag']), formatar_moeda_br(resumo['saldo'])]
    ]
    t_tot = Table(data_totais, colWidths=[47*mm, 47*mm, 47*mm, 48*mm])
    
    cor_saldo_final = colors.red if resumo['saldo'] > 0.01 else (colors.darkgreen if resumo['saldo'] < -0.01 else colors.black)

    t_tot.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'), ('BACKGROUND', (0,0), (-1,-1), bg_blu),
        ('TEXTCOLOR', (1,1), (1,1), colors.darkgreen), ('TEXTCOLOR', (2,1), (2,1), colors.red),
        ('TEXTCOLOR', (3,1), (3,1), cor_saldo_final),
    ]))
    story.append(t_tot)
    story.append(Spacer(1, 8*mm))
    
    headers = ['Empenho', 'Data', 'Vlr. Retido', 'Vlr. Pago', 'Diferença', 'Histórico', 'Status']
    data = [headers]
    
    table_style = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), 
        ('TEXTCOLOR', (0,0), (-1,0), colors.black), ('ALIGN', (0,0), (-1,0), 'CENTER'), ('ALIGN', (0,1), (-1,-1), 'CENTER'),
        ('ALIGN', (2,1), (4,-1), 'RIGHT'), ('ALIGN', (5,1), (5,-1), 'LEFT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (-1,-1), 7),
    ]
    
    bg_subtotal = colors.Color(0.9, 0.9, 0.9) # Cinza claro

    for i, (_, r) in enumerate(df_f.iterrows()):
        row_idx = i + 1
        
        if r['Status'] == 'SUBTOTAL':
            # Linha de Subtotal
            row_data = [
                "TOTAL", 
                r['Data Emp'], 
                formatar_moeda_br(r['Vlr Retido']), 
                "-", 
                "-", 
                "-", 
                "-"
            ]
            data.append(row_data)
            
            # Estilo da linha de Subtotal
            table_style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), bg_subtotal))
            table_style.append(('FONTNAME', (0, row_idx), (-1, row_idx), 'Helvetica-Bold'))
            table_style.append(('TEXTCOLOR', (0, row_idx), (-1, row_idx), colors.black))
            
        else:
            # Linha Normal
            dif = r['Dif']
            data_emp_pdf = str(r['Data Pag']) if r['Status'] == "Pago s/ Retenção" else str(r['Data Emp'])
            row_data = [
                str(r['Empenho']), data_emp_pdf, formatar_moeda_br(r['Vlr Retido']), formatar_moeda_br(r['Vlr Pago']),
                formatar_moeda_br(dif) if abs(dif) >= 0.01 else "-", Paragraph(str(r['Histórico']), style_hist), str(r['Status'])
            ]
            data.append(row_data)
            
            if abs(dif) >= 0.01:
                cor_fonte = colors.red if dif > 0 else colors.darkgreen
                table_style.append(('TEXTCOLOR', (4, row_idx), (4, row_idx), cor_fonte))
                table_style.append(('FONTNAME', (4, row_idx), (4, row_idx), 'Helvetica-Bold'))

    data.append(['TOTAIS PERÍODO', '', formatar_moeda_br(resumo['tot_ret']), formatar_moeda_br(resumo['tot_pag']), formatar_moeda_br(resumo['saldo'] - saldo_anterior), '', ''])
    last_row_idx = len(data) - 1
    table_style.append(('BACKGROUND', (0, last_row_idx), (-1, last_row_idx), colors.lightgrey))
    table_style.append(('FONTNAME', (0, last_row_idx), (-1, last_row_idx), 'Helvetica-Bold'))
    table_style.append(('SPAN', (0, last_row_idx), (1, last_row_idx)))

    col_widths = [20*mm, 16*mm, 25*mm, 25*mm, 25*mm, 59*mm, 19*mm]
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle(table_style))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

def gerar_excel_geral(df_resumo, ug):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df_resumo.to_excel(writer, sheet_name='Geral', index=False, startrow=2)
        wb = writer.book
        ws = writer.sheets['Geral']
        
        fmt_title = wb.add_format({'bold': True, 'bg_color': '#E6E6E6', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        fmt_head = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_money = wb.add_format({'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'valign': 'vcenter'})
        fmt_text = wb.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
        fmt_green = wb.add_format({'font_color': '#006400', 'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'valign': 'vcenter'})
        fmt_red = wb.add_format({'font_color': '#FF0000', 'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'valign': 'vcenter'})
        
        ws.merge_range('A1:E1', f"Relatório Geral de Retenções | UG: {ug}", fmt_title)
        
        for i, col in enumerate(df_resumo.columns):
            ws.write(2, i, col, fmt_head)
            
        for row_num, row_data in enumerate(df_resumo.values):
            excel_row = row_num + 3
            for col_num, cell_data in enumerate(row_data):
                if col_num == 0:
                    ws.write(excel_row, col_num, cell_data, fmt_text)
                else:
                    style = fmt_money
                    if col_num == 4:
                        if cell_data > 0.01: style = fmt_red
                        elif cell_data < -0.01: style = fmt_green
                    ws.write(excel_row, col_num, cell_data, style)

        for i, col in enumerate(df_resumo.columns):
            max_len = len(str(col))
            for val in df_resumo[col]:
                val_len = len(str(val))
                if val_len > max_len: max_len = val_len
            ws.set_column(i, i, max_len + 2)
    return out.getvalue()

def gerar_pdf_geral(df_resumo, ug):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm, title=f"Relatório Geral - {ug}")
    story = []
    styles = getSampleStyleSheet()
    
    story.append(Paragraph(f"Relatório Geral de Retenções | UG: {ug}", styles["Title"]))
    story.append(Spacer(1, 5*mm))
    
    style_cell = ParagraphStyle(name='Cell', parent=styles['Normal'], fontSize=8, leading=9)

    data = []
    data.append([
        "Conta De Retenção", "Saldo Anterior", "Retido Período", "Pago Período", "Saldo A Pagar"
    ])

    for _, row in df_resumo.iterrows():
        r_data = [
            Paragraph(str(row['Conta De Retenção']), style_cell),
            formatar_moeda_br(row['Saldo Anterior']),
            formatar_moeda_br(row['Retido Período']),
            formatar_moeda_br(row['Pago Período']),
            formatar_moeda_br(row['Saldo A Pagar'])
        ]
        data.append(r_data)
        
    table_style = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (0,1), (0,-1), 'LEFT'),
        ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]
    
    for i in range(1, len(data)):
        val_saldo = df_resumo.iloc[i-1]['Saldo A Pagar']
        if val_saldo > 0.01: color = colors.red
        elif val_saldo < -0.01: color = colors.darkgreen
        else: color = colors.black
        table_style.append(('TEXTCOLOR', (4, i), (4, i), color))
        table_style.append(('FONTNAME', (4, i), (4, i), 'Helvetica-Bold'))

    t = Table(data, colWidths=[85*mm, 26.25*mm, 26.25*mm, 26.25*mm, 26.25*mm])
    t.setStyle(TableStyle(table_style))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

def gerar_tabela_html_geral(df_resultado):
    html = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>"
    html += "<table style='width:100%; border-collapse: collapse; color: black !important; background-color: white !important; table-layout: fixed;'>"
    html += "<tr style='background-color: black; color: white !important;'>"
    html += "<th style='padding: 8px; border: 1px solid #000; text-align: left; width: 40%;'>Conta de Retenção</th>"
    html += "<th style='padding: 8px; border: 1px solid #000; text-align: right; width: 15%;'>Saldo Anterior</th>"
    html += "<th style='padding: 8px; border: 1px solid #000; text-align: right; width: 15%;'>Retido Período</th>"
    html += "<th style='padding: 8px; border: 1px solid #000; text-align: right; width: 15%;'>Pago Período</th>"
    html += "<th style='padding: 8px; border: 1px solid #000; text-align: right; width: 15%;'>Saldo a Pagar</th>"
    html += "</tr>"
    
    for _, r in df_resultado.iterrows():
        saldo = r['Saldo A Pagar']
        style_saldo = "color: red; font-weight: bold;" if saldo > 0.01 else ("color: darkgreen; font-weight: bold;" if saldo < -0.01 else "color: black;")
        
        html += "<tr style='background-color: white;'>"
        html += f"<td style='border: 1px solid #000; text-align: left; color: black;'>{r['Conta De Retenção']}</td>"
        html += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Saldo Anterior'])}</td>"
        html += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Retido Período'])}</td>"
        html += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Pago Período'])}</td>"
        html += f"<td style='border: 1px solid #000; text-align: right; {style_saldo}'>{formatar_moeda_br(saldo)}</td>"
        html += "</tr>"
    html += "</table></div>"
    return html

# ==============================================================================
# 2. INTERFACE GRÁFICA (AJUSTADA)
# ==============================================================================

st.markdown("<h1 style='text-align: center;'>Conciliação de Retenções</h1>", unsafe_allow_html=True)
st.markdown("---")

c_top_1, c_top_2 = st.columns(2)

with c_top_1: 
    st.markdown('<p class="big-label">1. Selecione o Razão da Contabilidade (.xlsx)</p>', unsafe_allow_html=True)
    arquivo = st.file_uploader("", type=["xlsx", "csv"], key="up_razao", label_visibility="collapsed")

with c_top_2:
    st.markdown('<p class="big-label">2. Filtros e Saldos</p>', unsafe_allow_html=True)
    placeholder_filtros = st.empty()

    if 'modo_conciliacao' not in st.session_state: st.session_state['modo_conciliacao'] = 'individual'
    
    # Flag para controlar execução
    if 'executar_individual' not in st.session_state: st.session_state['executar_individual'] = False

    if 'df_saldos_geral' not in st.session_state:
        st.session_state['df_saldos_geral'] = pd.DataFrame({
            "CONTA DE RETENÇÃO": LISTA_CONTAS,
            "SALDO ANTERIOR": 0.0
        })

if arquivo:
    df_dados = carregar_dados(arquivo)
    
    if not df_dados.empty:
        opcoes_ug = LISTA_UGS
        opcoes_conta = LISTA_CONTAS
        
        with placeholder_filtros.container():
            r1_col1, r1_col2 = st.columns([1, 3]) 
            with r1_col1: 
                ug_sel = st.selectbox("UG", opcoes_ug)
            with r1_col2: 
                conta_sel = st.selectbox("Conta de Retenção", opcoes_conta)
            
            val_anterior_str = st.text_input("Saldo Anterior", value="0,00", help="Digite o saldo acumulado de períodos anteriores.")
            
            st.markdown("<br>", unsafe_allow_html=True)
        
        # BOTÕES LADO A LADO FORA DO PLACEHOLDER
        c_btn_geral, c_btn_indiv = st.columns(2)
        
        with c_btn_geral:
            if st.button("PROCESSAR CONCILIAÇÃO GERAL", use_container_width=True):
                st.session_state['modo_conciliacao'] = 'geral'
                st.session_state['executar_individual'] = False
                pass
        
        with c_btn_indiv:
            if st.button("PROCESSAR CONCILIAÇÃO INDIVIDUAL", use_container_width=True):
                st.session_state['modo_conciliacao'] = 'individual'
                st.session_state['executar_individual'] = True

        st.markdown("---")
        
        # MODO INDIVIDUAL
        if st.session_state['modo_conciliacao'] == 'individual':
            if st.session_state.get('executar_individual'):
                saldo_ant_float = converter_moeda_input(val_anterior_str)
                
                with st.spinner("Processando..."):
                    # Retorna dados BRUTOS
                    df_res, resumo = processar_conciliacao(df_dados, ug_sel, conta_sel, saldo_ant_float)
                    
                    # Cria versão para o NOVO RELATÓRIO DE CIMA (Resumido)
                    df_resumo_superior = preparar_dados_resumo_superior(df_res)
                    
                    # Cria versão para o RELATÓRIO DE BAIXO (Detalhado - já existente)
                    df_res_visual = inserir_subtotais_diarios(df_res)
                
                if not df_res.empty:
                    c_k1, c_k2, c_k3 = st.columns(3)
                    with c_k1: st.markdown(f"""<div class="metric-card"><div class="metric-label">Retido s/ Pgto</div><div class="metric-value" style="color: #ff4b4b;">{resumo['ret_pendente']}</div></div>""", unsafe_allow_html=True)
                    with c_k2: st.markdown(f"""<div class="metric-card metric-card-orange"><div class="metric-label">Pago s/ Retenção</div><div class="metric-value" style="color: #ffc107;">{resumo['pag_sobra']}</div></div>""", unsafe_allow_html=True)
                    with c_k3: st.markdown(f"""<div class="metric-card metric-card-green"><div class="metric-label">Conciliados</div><div class="metric-value" style="color: #28a745;">{resumo['ok']}</div></div>""", unsafe_allow_html=True)
                    
                    v1, v2, v3 = st.columns(3)
                    with v1: st.markdown(f"""<div class="metric-card"><div class="metric-label">Total Retido s/ Pgto</div><div class="metric-value" style="color: #ff4b4b;">{formatar_moeda_br(resumo['val_ret_pendente'])}</div></div>""", unsafe_allow_html=True)
                    with v2: st.markdown(f"""<div class="metric-card metric-card-orange"><div class="metric-label">Total Pago s/ Retenção</div><div class="metric-value" style="color: #ffc107;">{formatar_moeda_br(resumo['val_pag_sobra'])}</div></div>""", unsafe_allow_html=True)
                    with v3: st.markdown(f"""<div class="metric-card metric-card-green"><div class="metric-label">Total Retido e Pago</div><div class="metric-value" style="color: #28a745;">{formatar_moeda_br(resumo['val_ok'])}</div></div>""", unsafe_allow_html=True)

                    f1, f2, f3, f4 = st.columns(4)
                    cor_saldo = "#ff4b4b" if resumo['saldo'] > 0.01 else ("#28a745" if resumo['saldo'] < -0.01 else "#343a40")
                    
                    with f1: st.markdown(f"""<div class="metric-card metric-card-dark"><div class="metric-label">Saldo Anterior</div><div class="metric-value" style="color: #343a40;">{formatar_moeda_br(saldo_ant_float)}</div></div>""", unsafe_allow_html=True)
                    with f2: st.markdown(f"""<div class="metric-card metric-card-blue"><div class="metric-label">Total Retido (Período)</div><div class="metric-value" style="color: #004085;">{formatar_moeda_br(resumo['tot_ret'])}</div></div>""", unsafe_allow_html=True)
                    with f3: st.markdown(f"""<div class="metric-card metric-card-blue"><div class="metric-label">Total Pago (Período)</div><div class="metric-value" style="color: #004085;">{formatar_moeda_br(resumo['tot_pag'])}</div></div>""", unsafe_allow_html=True)
                    with f4: st.markdown(f"""<div class="metric-card metric-card-dark"><div class="metric-label">Saldo a Pagar</div><div class="metric-value" style="color: {cor_saldo};">{formatar_moeda_br(resumo['saldo'])}</div></div>""", unsafe_allow_html=True)

                    # --- 1. RELATÓRIO RESUMIDO (ATUALIZADO) ---
                    st.markdown("### RESUMO DE RETENÇÕES E PAGAMENTOS")
                    html_resumo = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>"
                    html_resumo += "<table style='width:100%; border-collapse: collapse; color: black !important; background-color: white !important; table-layout: fixed;'>"
                    html_resumo += "<tr style='background-color: black; color: white !important;'>"
                    html_resumo += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Empenho</th>"
                    html_resumo += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Data</th>"
                    html_resumo += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Vlr Retido</th>"
                    html_resumo += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Vlr Pago</th>"
                    html_resumo += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Diferença</th>"
                    html_resumo += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 34%;'>Histórico</th>"
                    html_resumo += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Status</th>"
                    html_resumo += "</tr>"
                    
                    if not df_resumo_superior.empty:
                        for _, r in df_resumo_superior.iterrows():
                            # Se for Subtotal de "Retenção" (Pendente)
                            if r['Empenho'] == "Retenção":
                                html_resumo += "<tr style='background-color: white;'>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black;'>Retenção</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Data']}</td>"
                                # Valor Retido PRETO
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Retido'])}</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Pago'])}</td>"
                                # Diferença VERMELHA E NEGRITO
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; color: red; font-weight: bold;'>{formatar_moeda_br(r['Dif'])}</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black;'>-</td>"
                                # Status sem negrito e com fonte 12px
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black; font-size: 12px; font-weight: normal !important;'>{r['Status']}</td>"
                                html_resumo += "</tr>"

                            # Se for Subtotal de "Conciliado"
                            elif r['Empenho'] == "Conciliado":
                                html_resumo += "<tr style='background-color: white;'>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black;'>Conciliado</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Data']}</td>"
                                # Valor Preto
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Retido'])}</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Pago'])}</td>"
                                
                                # Diferença com cor condicional padrão
                                dif_conc = r['Dif']
                                style_dif_conc = "color: red; font-weight: bold;" if dif_conc > 0.01 else ("color: darkgreen; font-weight: bold;" if dif_conc < -0.01 else "color: black;")
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; {style_dif_conc}'>{formatar_moeda_br(dif_conc)}</td>"
                                
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black;'>-</td>"
                                # Status sem negrito e com fonte 12px
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black; font-size: 12px; font-weight: normal !important;'>{r['Status']}</td>"
                                html_resumo += "</tr>"

                            # Linha normal (Pago s/ Retenção)
                            else:
                                dif = r['Dif']
                                style_dif = "color: red; font-weight: bold;" if dif > 0.01 else ("color: darkgreen; font-weight: bold;" if dif < -0.01 else "color: black;")
                                html_resumo += "<tr style='background-color: white;'>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Empenho']}</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Data']}</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Retido'])}</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Pago'])}</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: right; {style_dif}'>{formatar_moeda_br(dif)}</td>"
                                html_resumo += f"<td style='border: 1px solid #000; text-align: left; color: black; font-size: 11px; word-wrap: break-word; white-space: normal;'>{r['Histórico']}</td>"
                                # Status sem negrito e com fonte 12px
                                html_resumo += f"<td style='border: 1px solid #000; text-align: center; color: black; font-size: 12px; font-weight: normal !important;'>{r['Status']}</td>"
                                html_resumo += "</tr>"
                    else:
                         html_resumo += "<tr><td colspan='7' style='text-align:center; padding:10px; border:1px solid #000;'>Nenhum dado para o resumo.</td></tr>"
                         
                    html_resumo += "</table></div>"
                    st.markdown(html_resumo, unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)

                    # --- 2. RELATÓRIO DETALHADO (EXISTENTE) ---
                    st.markdown("### RELATÓRIO DETALHADO DE RETENÇÕES")
                    html = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>"
                    html += "<table style='width:100%; border-collapse: collapse; color: black !important; background-color: white !important; table-layout: fixed;'>"
                    html += "<tr style='background-color: black; color: white !important;'>"
                    html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Empenho</th>"
                    html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Data</th>"
                    html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Vlr Retido</th>"
                    html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Vlr Pago</th>"
                    html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Diferença</th>"
                    html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 34%;'>Histórico</th>"
                    html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Status</th>"
                    html += "</tr>"
                    
                    for _, r in df_res_visual.iterrows():
                        if r['Status'] == 'SUBTOTAL':
                            html += "<tr style='background-color: #E6E6E6; font-weight: bold;'>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>TOTAL</td>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Data Emp']}</td>"
                            html += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Retido'])}</td>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>-</td>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>-</td>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>-</td>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>-</td>"
                            html += "</tr>"
                        else:
                            dif = r['Dif']
                            style_dif = "color: red; font-weight: bold;" if dif > 0.01 else ("color: darkgreen; font-weight: bold;" if dif < -0.01 else "color: black;")
                            data_exibicao = r['Data Pag'] if r['Status'] == "Pago s/ Retenção" else r['Data Emp']
                            html += "<tr style='background-color: white;'>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Empenho']}</td>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{data_exibicao}</td>"
                            html += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Retido'])}</td>"
                            html += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Pago'])}</td>"
                            html += f"<td style='border: 1px solid #000; text-align: right; {style_dif}'>{formatar_moeda_br(dif)}</td>"
                            html += f"<td style='border: 1px solid #000; text-align: left; color: black; font-size: 11px; word-wrap: break-word; white-space: normal;'>{r['Histórico']}</td>"
                            html += f"<td style='border: 1px solid #000; text-align: center; color: black; font-size: 12px;'>{r['Status']}</td>"
                            html += "</tr>"
                    
                    html += f"<tr style='font-weight: bold; background-color: lightgrey; color: black;'>"
                    html += "<td colspan='2' style='padding: 10px; text-align: center; border: 1px solid #000;'>TOTAL PERÍODO</td>"
                    html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(resumo['tot_ret'])}</td>"
                    html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(resumo['tot_pag'])}</td>"
                    html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(resumo['saldo'] - saldo_ant_float)}</td>"
                    html += "<td colspan='2' style='border: 1px solid #000;'></td></tr>"
                    html += "</table></div>"
                    st.markdown(html, unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    ug_limpa = limpar_nome_arquivo(str(ug_sel))
                    conta_limpa = limpar_nome_arquivo(str(conta_sel).split(' ')[0])
                    nome_base = f"Conciliacao_Retencoes_UG_{ug_limpa}_Retencao_{conta_limpa}"
                    
                    # Excel usa df_res (SEM subtotais)
                    excel_bytes = gerar_excel(df_res, resumo, saldo_ant_float, ug_sel, conta_sel)
                    st.download_button("BAIXAR RELATÓRIO EM EXCEL", excel_bytes, f"{nome_base}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    
                    # PDF usa df_res_visual (COM subtotais do relatório detalhado)
                    pdf_bytes = gerar_pdf(df_res_visual, ug_sel, conta_sel, resumo, saldo_ant_float)
                    st.download_button("BAIXAR RELATÓRIO EM PDF", pdf_bytes, f"{nome_base}.pdf", "application/pdf", use_container_width=True)
                else:
                    st.warning("Nenhum dado encontrado.")

# MODO GERAL (SUBSTITUA O BLOCO ANTIGO POR ESTE)
        elif st.session_state['modo_conciliacao'] == 'geral':
            st.markdown("### Conciliação Geral (Múltiplas Contas)")
            st.info("Insira o Saldo Anterior para cada conta abaixo. Os valores só serão processados ao clicar em CONCILIAR.")
            
            # CSS PARA FORÇAR O VISUAL BRANCO E GRANDE NOS INPUTS
            st.markdown("""
            <style>
                /* Aumenta a fonte e muda as cores dos inputs numéricos */
                div[data-testid="stNumberInput"] input {
                    background-color: white !important;
                    color: black !important;
                    font-size: 22px !important; /* Fonte bem grande */
                    font-weight: bold !important;
                    border: 2px solid #ccc !important;
                    border-radius: 5px;
                    height: 50px;
                }
                /* Botões de + e - dentro do input ficam pretos */
                div[data-testid="stNumberInput"] button {
                    color: black !important;
                }
                /* Tira o rótulo padrão pequeno para usarmos o nosso personalizado */
                label[data-testid="stWidgetLabel"] {
                    display: none;
                }
            </style>
            """, unsafe_allow_html=True)

            # --- INÍCIO DO FORMULÁRIO (ISSO GARANTE QUE OS DADOS NÃO SUMAM) ---
            with st.form("form_conciliacao_geral"):
                
                # Cabeçalho da "Tabela Falsa"
                c_h1, c_h2 = st.columns([3, 1])
                c_h1.markdown("**CONTA DE RETENÇÃO**", unsafe_allow_html=True)
                c_h2.markdown("**SALDO ANTERIOR**", unsafe_allow_html=True)
                st.markdown("---")

                df_editor = st.session_state['df_saldos_geral'].copy()
                
                # Dicionário para guardar temporariamente as referências dos inputs
                inputs_temp = {}

                # CRIA UMA LINHA PARA CADA CONTA
                for index, row in df_editor.iterrows():
                    c_nome, c_valor = st.columns([3, 1])
                    
                    with c_nome:
                        # Exibe o nome da conta com letra maior e alinhado
                        st.markdown(f"<div style='padding-top: 15px; font-size: 18px;'>{row['CONTA DE RETENÇÃO']}</div>", unsafe_allow_html=True)
                    
                    with c_valor:
                        # Input numérico GRANDE e BRANCO
                        # A 'key' única garante que o Streamlit não perca o valor
                        chave_unica = f"input_saldo_{index}"
                        val = st.number_input(
                            "Saldo", 
                            value=float(row['SALDO ANTERIOR']),
                            key=chave_unica,
                            format="%.2f",
                            label_visibility="collapsed"
                        )
                        # Guardamos o valor digitado no dicionário usando o índice
                        inputs_temp[index] = val
                    
                    # Uma linha fina para separar visualmente
                    st.markdown("<hr style='margin: 5px 0; border-color: #444;'>", unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                submit_btn = st.form_submit_button("CONCILIAR TODOS OS ITENS", use_container_width=True)

            # --- LÓGICA APÓS CLICAR NO BOTÃO ---
            if submit_btn:
                # 1. Atualiza o DataFrame na memória com o que foi coletado no loop
                # Isso "Salva" os dados igual você fazia na tabela
                for idx, valor_digitado in inputs_temp.items():
                    df_editor.at[idx, 'SALDO ANTERIOR'] = valor_digitado
                
                st.session_state['df_saldos_geral'] = df_editor
                
                # DAQUI PARA BAIXO É O PROCESSAMENTO NORMAL
                resultados_gerais = []
                progresso = st.progress(0)
                total_contas = len(df_editor)
                
                for idx, row in df_editor.iterrows():
                    conta = row['CONTA DE RETENÇÃO']
                    saldo_ant = row['SALDO ANTERIOR']
                    _, resumo = processar_conciliacao(df_dados, ug_sel, conta, saldo_ant)
                    resultados_gerais.append({
                        "Conta De Retenção": conta,
                        "Saldo Anterior": saldo_ant,
                        "Retido Período": resumo['tot_ret'],
                        "Pago Período": resumo['tot_pag'],
                        "Saldo A Pagar": resumo['saldo']
                    })
                    progresso.progress((idx + 1) / total_contas)
                
                df_resultado_geral = pd.DataFrame(resultados_gerais)
                st.success("Conciliação Geral concluída!")
                
                html_geral = gerar_tabela_html_geral(df_resultado_geral)
                st.markdown(html_geral, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                
                ug_limpa = limpar_nome_arquivo(str(ug_sel))
                nome_base_geral = f"Relatorio_Geral_Retencoes_UG_{ug_limpa}"
                
                excel_bytes_geral = gerar_excel_geral(df_resultado_geral, ug_sel)
                st.download_button("BAIXAR RELATÓRIO GERAL (XLSX)", excel_bytes_geral, f"{nome_base_geral}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                
                pdf_bytes_geral = gerar_pdf_geral(df_resultado_geral, ug_sel)
                st.download_button("BAIXAR RELATÓRIO GERAL (PDF)", pdf_bytes_geral, f"{nome_base_geral}.pdf", "application/pdf", use_container_width=True)
    else:
        st.warning("Nenhum dado encontrado para os filtros selecionados.")
