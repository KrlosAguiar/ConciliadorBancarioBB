import streamlit as st
import pandas as pd
import io
import os
from PIL import Image

# --- CONFIGURAÇÃO DA PÁGINA (Deve ser a primeira chamada Streamlit) ---
st.set_page_config(
    page_title="Conciliação de Retenções",
    layout="wide"
)

# --- CONFIGURAÇÃO DE ÍCONE (Ajuste para estrutura de pages) ---
# Tenta localizar a imagem na raiz ou na pasta atual
icon_filename = "Barcarena.png"
possible_paths = [
    os.path.join(os.getcwd(), icon_filename),           # Raiz
    os.path.join(os.path.dirname(__file__), icon_filename), # Mesma pasta do script
    os.path.join("..", icon_filename)                   # Pasta pai
]

icon_image = None
for p in possible_paths:
    if os.path.exists(p):
        try:
            icon_image = Image.open(p)
            # Reaplica config com ícone se achar (algumas versões do st permitem, outras ignoram a segunda chamada)
            # Mas como set_page_config só pode ser chamado uma vez, definimos o ícone na barra lateral
            st.sidebar.image(icon_image, width=150)
            break
        except:
            pass

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
        width: 100%;
    }
    div.stButton > button:hover { background-color: rgb(20, 20, 25) !important; border-color: white; }
    .big-label { font-size: 18px !important; font-weight: 600 !important; margin-bottom: 5px; color: #333; }
    
    /* Cards de Resumo */
    .metric-card {
        background-color: #f0f2f6;
        border-left: 5px solid #ff4b4b;
        padding: 15px;
        border-radius: 5px;
        color: black;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .metric-card-green { border-left: 5px solid #28a745; }
    .metric-card-orange { border-left: 5px solid #ffc107; }
    .metric-value { font-size: 24px; font-weight: bold; }
    .metric-label { font-size: 14px; color: #555; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE PROCESSAMENTO
# ==============================================================================

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == "": return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

@st.cache_data(show_spinner=False)
def carregar_e_tratar_dados(file):
    try:
        # Tenta ler como Excel primeiro
        df = pd.read_excel(file)
    except:
        # Fallback para CSV
        file.seek(0)
        try:
            df = pd.read_csv(file, sep=None, engine='python', encoding='latin1', on_bad_lines='skip')
        except:
            df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8', on_bad_lines='skip')

    # Tratamento de Colunas Mescladas (Ffill)
    df = df.ffill(axis=0)
    df = df.dropna(how='all')

    # Conversão de Valores (Coluna 8 - índice 8)
    # Assumindo estrutura do Razão Geral padrão
    col_valor_idx = 8 
    col_nome_valor = df.columns[col_valor_idx]

    def converter_valor(val):
        try:
            if isinstance(val, str):
                return float(val.replace('.', '').replace(',', '.'))
            return float(val)
        except:
            return 0.0

    df[col_nome_valor] = df[col_nome_valor].apply(converter_valor)
    return df

def executar_conciliacao_core(df, ug_selecionada, conta_texto):
    # Extrai o código numérico da conta (ex: "7852" de "7852 - ISS...")
    codigo_conta = conta_texto.split(' - ')[0].strip()
    if not codigo_conta.isdigit(): 
        codigo_conta = conta_texto.split(' ')[0]

    # Mapeamento de Colunas (baseado na análise anterior)
    c_ug = df.columns[0]
    c_data = df.columns[4]
    c_dc = df.columns[5]
    c_conta = df.columns[6]
    c_valor = df.columns[8]
    c_empenho = df.columns[14]
    c_tipo = df.columns[19]
    c_hist = df.columns[27]

    # Filtros Iniciais
    mask_ug = df[c_ug].astype(str) == str(ug_selecionada)
    mask_conta = df[c_conta].astype(str).str.startswith(str(codigo_conta))
    
    df_base = df[mask_ug & mask_conta].copy()
    
    if df_base.empty:
        return pd.DataFrame(), {}

    df_base['Tipo_Norm'] = df_base[c_tipo].astype(str).str.strip()

    # --- Separação dos Grupos ---
    # 1. Retenções (Crédito)
    mask_ret = (df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'C')
    df_retencoes = df_base[mask_ret].copy()
    
    # 2. Estornos de Retenção (Débito)
    mask_estorno = (df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'D')
    df_estornos = df_base[mask_estorno].copy()
    
    # 3. Pagamentos (Débito)
    mask_pag = (df_base['Tipo_Norm'].str.contains("Pagamento de Documento Extra", case=False)) & (df_base[c_dc] == 'D')
    df_pagamentos = df_base[mask_pag].copy()

    # --- Limpeza de Estornos ---
    indices_retencoes_canceladas = set()
    for _, row_est in df_estornos.iterrows():
        val_est = row_est[c_valor]
        emp_est = row_est[c_empenho]
        # Busca retenção compatível para remover
        candidatos = df_retencoes[
            (df_retencoes[c_empenho] == emp_est) &
            (df_retencoes[c_valor] == val_est) &
            (~df_retencoes.index.isin(indices_retencoes_canceladas))
        ]
        if not candidatos.empty:
            indices_retencoes_canceladas.add(candidatos.index[0])
    
    df_retencoes_limpas = df_retencoes[~df_retencoes.index.isin(indices_retencoes_canceladas)]

    # --- Lógica de Match ---
    resultados = []
    indices_pagos_usados = set()
    
    # Resumo Contadores
    resumo = {
        "retido_nao_pago": 0,
        "pago_nao_retido": 0,
        "conciliado": 0,
        "total_retido": 0.0,
        "total_pago": 0.0,
        "saldo": 0.0
    }

    # Passo A: Retenções X Pagamentos
    for _, row_ret in df_retencoes_limpas.iterrows():
        val_ret = row_ret[c_valor]
        match = False
        
        cand_pag = df_pagamentos[
            (df_pagamentos[c_valor] == val_ret) &
            (~df_pagamentos.index.isin(indices_pagos_usados))
        ]
        
        val_pago = 0.0
        data_pag = "Pendente"
        hist_pag = "-"
        status_sort = 0 # 0=Prioridade (Retido n/ Pago)
        
        if not cand_pag.empty:
            row_pag = cand_pag.iloc[0]
            val_pago = row_pag[c_valor]
            data_pag = row_pag[c_data]
            hist_pag = row_pag[c_hist]
            indices_pagos_usados.add(row_pag.name)
            match = True
            status_sort = 2 # Final da lista (Conciliado)
            resumo["conciliado"] += 1
        else:
            resumo["retido_nao_pago"] += 1
        
        diff = val_ret - val_pago
        
        resultados.append({
            "Empenho": row_ret[c_empenho],
            "Data Empenho": row_ret[c_data],
            "Valor Retido": val_ret,
            "Valor Pago": val_pago,
            "Diferença": diff,
            "Data Pagamento": data_pag,
            "Histórico Retenção": row_ret[c_hist],
            "Histórico Pagamento": hist_pag,
            "_sort": status_sort,
            "Status": "Conciliado" if match else "Pendente Pagto"
        })

    # Passo B: Pagamentos Sobras (Pagos e não retidos)
    df_pagamentos_sobras = df_pagamentos[~df_pagamentos.index.isin(indices_pagos_usados)]
    
    for _, row_pag in df_pagamentos_sobras.iterrows():
        val_pago = row_pag[c_valor]
        resumo["pago_nao_retido"] += 1
        
        resultados.append({
            "Empenho": row_pag[c_empenho],
            "Data Empenho": "-", 
            "Valor Retido": 0.0,
            "Valor Pago": val_pago,
            "Diferença": 0.0 - val_pago,
            "Data Pagamento": row_pag[c_data],
            "Histórico Retenção": "-",
            "Histórico Pagamento": row_pag[c_hist],
            "_sort": 1, # Meio da lista
            "Status": "Pago s/ Retenção"
        })

    if not resultados:
        return pd.DataFrame(), resumo

    df_res = pd.DataFrame(resultados)
    # Ordenação: Status -> Data Empenho -> Data Pagamento
    df_res = df_res.sort_values(by=['_sort', 'Data Empenho', 'Data Pagamento'])
    
    # Preencher totais
    resumo["total_retido"] = df_res["Valor Retido"].sum()
    resumo["total_pago"] = df_res["Valor Pago"].sum()
    resumo["saldo"] = df_res["Diferença"].sum()

    return df_res, resumo

def gerar_excel_ajustado(df_final):
    output = io.BytesIO()
    
    # Remove colunas auxiliares antes de exportar
    df_export = df_final.drop(columns=['_sort', 'Status'])
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, sheet_name='Conciliacao', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Conciliacao']
        
        # Formatos
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
        fmt_red = workbook.add_format({'font_color': '#FF0000', 'num_format': '#,##0.00'})
        fmt_total = workbook.add_format({'bold': True, 'bg_color': '#E6E6E6', 'num_format': '#,##0.00', 'border': 1})
        
        # Aplicar cabeçalho
        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, fmt_header)
        
        # --- AJUSTE AUTOMÁTICO DE LARGURA DAS COLUNAS ---
        for i, col in enumerate(df_export.columns):
            # Tamanho base do cabeçalho
            max_len = len(str(col)) + 2
            
            # Varre as primeiras 50 linhas para estimar tamanho
            column_data = df_export[col].head(50).astype(str)
            if not column_data.empty:
                max_data_len = column_data.map(len).max()
                if max_data_len > max_len:
                    max_len = max_data_len
            
            # Limites visuais
            max_len = min(max_len, 60)
            max_len = max(max_len, 12)
            
            worksheet.set_column(i, i, max_len)
            
        # Formatação de Moeda nas colunas C, D, E (índices 2, 3, 4)
        worksheet.set_column('C:E', 18, fmt_currency)
        
        # Condicional para Diferença (Vermelho se != 0)
        worksheet.conditional_format(1, 4, len(df_export), 4, 
            {'type': 'cell', 'criteria': '!=', 'value': 0, 'format': fmt_red})
            
        # Linha de Total
        last_row = len(df_export) + 1
        worksheet.write(last_row, 0, "TOTAL GERAL", fmt_total)
        worksheet.write(last_row, 2, df_export['Valor Retido'].sum(), fmt_total)
        worksheet.write(last_row, 3, df_export['Valor Pago'].sum(), fmt_total)
        worksheet.write(last_row, 4, df_export['Diferença'].sum(), fmt_total)
        
    return output.getvalue()

# ==============================================================================
# 2. INTERFACE
# ==============================================================================
st.markdown("<h1 style='text-align: center;'>Sistema de Conciliação de Retenções</h1>", unsafe_allow_html=True)
st.markdown("---")

st.markdown('<p class="big-label">1. Upload do Arquivo (Razão Geral .xlsx ou .csv)</p>', unsafe_allow_html=True)
arquivo_upload = st.file_uploader("", type=["xlsx", "csv"], label_visibility="collapsed")

if arquivo_upload:
    with st.spinner("Lendo arquivo..."):
        df_dados = carregar_e_tratar_dados(arquivo_upload)
    
    if not df_dados.empty:
        # Colunas Filtros
        c1, c2 = st.columns(2)
        
        # Filtro UG (Coluna 0)
        ugs = sorted(df_dados[df_dados.columns[0]].astype(str).unique().tolist())
        with c1:
            st.markdown('<p class="big-label">Unidade Gestora (UG)</p>', unsafe_allow_html=True)
            ug_selecionada = st.selectbox("Selecione a UG", options=ugs, label_visibility="collapsed")
            
        # Filtro Conta (Coluna 6)
        contas_raw = sorted(df_dados[df_dados.columns[6]].astype(str).unique().tolist())
        # Tenta colocar a opção padrão no topo e remover duplicatas na visualização se necessário
        opcoes_conta = ["7852 - ISS Pessoa Jurídica (Padrão)"] + [c for c in contas_raw if "7852" not in str(c)]
        
        with c2:
            st.markdown('<p class="big-label">Conta de Retenção</p>', unsafe_allow_html=True)
            conta_selecionada = st.selectbox("Selecione a Conta", options=opcoes_conta, label_visibility="collapsed")
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("PROCESSAR CONCILIAÇÃO"):
            with st.spinner("Cruzando dados, removendo estornos e ordenando..."):
                df_resultado, resumo = executar_conciliacao_core(df_dados, ug_selecionada, conta_selecionada)
                
            if not df_resultado.empty:
                # --- EXIBIÇÃO DE RESUMO (CARDS) ---
                k1, k2, k3 = st.columns(3)
                with k1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Retido e Não Pago</div>
                        <div class="metric-value" style="color: #ff4b4b;">{resumo['retido_nao_pago']}</div>
                        <small>Itens Pendentes</small>
                    </div>
                    """, unsafe_allow_html=True)
                with k2:
                    st.markdown(f"""
                    <div class="metric-card metric-card-orange">
                        <div class="metric-label">Pago sem Retenção</div>
                        <div class="metric-value" style="color: #ffc107;">{resumo['pago_nao_retido']}</div>
                         <small>Sobras / Anteriores</small>
                    </div>
                    """, unsafe_allow_html=True)
                with k3:
                    st.markdown(f"""
                    <div class="metric-card metric-card-green">
                        <div class="metric-label">Conciliados</div>
                        <div class="metric-value" style="color: #28a745;">{resumo['conciliado']}</div>
                         <small>Baixados com Sucesso</small>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # --- TABELA HTML ---
                html = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd; overflow-x: auto;'>"
                html += "<table style='width:100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px;'>"
                
                # Cabeçalho
                html += "<tr style='background-color: #333; color: white; text-align: center;'>"
                cols_view = ["Empenho", "Data Empenho", "Valor Retido", "Valor Pago", "Diferença", "Histórico Retenção", "Status"]
                for c in cols_view:
                    html += f"<th style='padding: 10px; border: 1px solid #555;'>{c}</th>"
                html += "</tr>"
                
                for _, row in df_resultado.iterrows():
                    # Cores baseadas no Status
                    bg_color = "white"
                    status_color = "black"
                    diff = row['Diferença']
                    
                    if row['Status'] == "Pendente Pagto":
                        bg_color = "#fff0f0" # Vermelho claro
                        status_color = "#d9534f"
                    elif row['Status'] == "Pago s/ Retenção":
                        bg_color = "#fffbe6" # Amarelo claro
                        status_color = "#f0ad4e"
                    
                    html += f"<tr style='background-color: {bg_color}; color: #333; border-bottom: 1px solid #ddd;'>"
                    html += f"<td style='padding: 8px; text-align: center;'>{row['Empenho']}</td>"
                    html += f"<td style='padding: 8px; text-align: center;'>{row['Data Empenho']}</td>"
                    html += f"<td style='padding: 8px; text-align: right;'>{formatar_moeda_br(row['Valor Retido'])}</td>"
                    html += f"<td style='padding: 8px; text-align: right;'>{formatar_moeda_br(row['Valor Pago'])}</td>"
                    
                    # Coluna Diferença (Vermelho se != 0)
                    cor_diff = "red" if abs(diff) > 0.01 else "green"
                    html += f"<td style='padding: 8px; text-align: right; font-weight: bold; color: {cor_diff};'>{formatar_moeda_br(diff)}</td>"
                    
                    # Histórico (Truncar se muito longo)
                    hist = str(row['Histórico Retenção'])
                    if len(hist) > 50: hist = hist[:47] + "..."
                    html += f"<td style='padding: 8px; text-align: left; font-size: 12px;'>{hist}</td>"
                    
                    html += f"<td style='padding: 8px; text-align: center; font-weight: bold; color: {status_color};'>{row['Status']}</td>"
                    html += "</tr>"
                
                # Linha de Total
                html += f"<tr style='background-color: #e6e6e6; font-weight: bold;'>"
                html += "<td colspan='2' style='padding: 10px; text-align: center;'>TOTAL GERAL</td>"
                html += f"<td style='text-align: right; padding: 10px;'>{formatar_moeda_br(resumo['total_retido'])}</td>"
                html += f"<td style='text-align: right; padding: 10px;'>{formatar_moeda_br(resumo['total_pago'])}</td>"
                html += f"<td style='text-align: right; padding: 10px; color: {'red' if abs(resumo['saldo']) > 0.01 else 'black'}'>{formatar_moeda_br(resumo['saldo'])}</td>"
                html += "<td colspan='2'></td></tr>"
                
                html += "</table></div>"
                st.markdown(html, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Botão Download
                excel_data = gerar_excel_ajustado(df_resultado)
                st.download_button(
                    label="BAIXAR RELATÓRIO EM EXCEL",
                    data=excel_data,
                    file_name="Conciliacao_Retencoes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("Nenhum registro encontrado com os filtros selecionados.")
    else:
        st.error("Erro ao ler o arquivo. Verifique se é um Razão Geral válido.")
