import streamlit as st
import pandas as pd
import io
import os
from PIL import Image

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
        st.set_page_config(page_title="Portal Financeiro - Retenções", page_icon=icon_image, layout="wide")
    else:
        st.set_page_config(page_title="Portal Financeiro - Retenções", layout="wide")
except: pass

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
    
    /* Cards de Resumo */
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
    .metric-sub { font-size: 11px; color: #888; margin-top: 5px;}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE PROCESSAMENTO
# ==============================================================================

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == "": return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

@st.cache_data(show_spinner=False)
def carregar_dados(file):
    try:
        df = pd.read_excel(file)
    except:
        file.seek(0)
        try:
            df = pd.read_csv(file, sep=None, engine='python', encoding='latin1', on_bad_lines='skip')
        except:
            df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8', on_bad_lines='skip')
    
    df = df.ffill(axis=0)
    df = df.dropna(how='all')
    
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

def processar_conciliacao(df, ug_sel, conta_sel):
    cod_conta = conta_sel.split(' - ')[0].strip()
    if not cod_conta.isdigit(): cod_conta = conta_sel.split(' ')[0]

    c_ug = df.columns[0]
    c_dc = df.columns[5]
    c_conta = df.columns[6]
    c_valor = df.columns[8]
    c_empenho = df.columns[14]
    c_tipo = df.columns[19]
    c_hist = df.columns[27]
    c_data = df.columns[4]

    # Filtros
    mask_ug = df[c_ug].astype(str) == str(ug_sel)
    mask_conta = df[c_conta].astype(str).str.startswith(str(cod_conta))
    df_base = df[mask_ug & mask_conta].copy()
    
    if df_base.empty: return pd.DataFrame(), {}

    df_base['Tipo_Norm'] = df_base[c_tipo].astype(str).str.strip()

    # Separação
    mask_ret = (df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'C')
    df_ret = df_base[mask_ret].copy()
    
    mask_estorno = (df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'D')
    df_est = df_base[mask_estorno].copy()
    
    mask_pag = (df_base['Tipo_Norm'].str.contains("Pagamento de Documento Extra", case=False)) & (df_base[c_dc] == 'D')
    df_pag = df_base[mask_pag].copy()

    # Limpeza Estornos
    idx_cancel = set()
    for _, r_est in df_est.iterrows():
        v = r_est[c_valor]
        e = r_est[c_empenho]
        cand = df_ret[(df_ret[c_empenho] == e) & (df_ret[c_valor] == v) & (~df_ret.index.isin(idx_cancel))]
        if not cand.empty: idx_cancel.add(cand.index[0])
    
    df_ret_limpa = df_ret[~df_ret.index.isin(idx_cancel)]

    # Conciliação
    resultados = []
    idx_pag_usado = set()
    
    # Inicializa Resumo
    resumo = {
        "ret_pendente": 0, "val_ret_pendente": 0.0,
        "pag_sobra": 0,    "val_pag_sobra": 0.0,
        "ok": 0,           "val_ok": 0.0,
        "tot_ret": 0.0, "tot_pag": 0.0, "saldo": 0.0
    }

    # Loop Retenções
    for _, r in df_ret_limpa.iterrows():
        val = r[c_valor]
        cand = df_pag[(df_pag[c_valor] == val) & (~df_pag.index.isin(idx_pag_usado))]
        
        val_pago = 0.0
        dt_pag = "-"
        match = False
        sort = 0 
        
        if not cand.empty:
            r_pag = cand.iloc[0]
            val_pago = r_pag[c_valor]
            dt_pag = r_pag[c_data]
            idx_pag_usado.add(r_pag.name)
            match = True
            sort = 2
            
            resumo["ok"] += 1
            resumo["val_ok"] += val_pago
        else:
            resumo["ret_pendente"] += 1
            resumo["val_ret_pendente"] += val
            
        resultados.append({
            "Empenho": r[c_empenho], "Data Emp": r[c_data],
            "Vlr Retido": val, "Vlr Pago": val_pago,
            "Dif": val - val_pago, "Data Pag": dt_pag,
            "Histórico": r[c_hist], "_sort": sort,
            "Status": "Conciliado" if match else "Retido s/ Pagto"
        })

    # Loop Sobras
    for _, r in df_pag[~df_pag.index.isin(idx_pag_usado)].iterrows():
        resumo["pag_sobra"] += 1
        resumo["val_pag_sobra"] += r[c_valor]
        
        resultados.append({
            "Empenho": r[c_empenho], "Data Emp": "-",
            "Vlr Retido": 0.0, "Vlr Pago": r[c_valor],
            "Dif": 0.0 - r[c_valor], "Data Pag": r[c_data],
            "Histórico": r[c_hist], "_sort": 1,
            "Status": "Pago s/ Retenção"
        })

    if not resultados: return pd.DataFrame(), resumo

    df_res = pd.DataFrame(resultados).sort_values(by=['_sort', 'Data Emp', 'Data Pag'])
    
    # Totais Gerais
    resumo["tot_ret"] = df_res["Vlr Retido"].sum()
    resumo["tot_pag"] = df_res["Vlr Pago"].sum()
    resumo["saldo"] = df_res["Dif"].sum()
    
    return df_res, resumo

def gerar_excel(df):
    out = io.BytesIO()
    df_exp = df.drop(columns=['_sort', 'Status'])
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df_exp.to_excel(writer, sheet_name='Conciliacao', index=False)
        wb = writer.book
        ws = writer.sheets['Conciliacao']
        
        fmt_head = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center'})
        fmt_money = wb.add_format({'num_format': '#,##0.00'})
        fmt_red = wb.add_format({'font_color': '#FF0000', 'num_format': '#,##0.00'})
        fmt_tot = wb.add_format({'bold': True, 'bg_color': '#E6E6E6', 'num_format': '#,##0.00', 'border': 1})

        for i, col in enumerate(df_exp.columns):
            ws.write(0, i, col, fmt_head)
            ws.set_column(i, i, 15)

        ws.set_column('C:E', 18, fmt_money)
        ws.conditional_format(1, 4, len(df_exp), 4, {'type': 'cell', 'criteria': '!=', 'value': 0, 'format': fmt_red})
        
        l = len(df_exp) + 1
        ws.write(l, 0, "TOTAL", fmt_tot)
        ws.write(l, 2, df_exp["Vlr Retido"].sum(), fmt_tot)
        ws.write(l, 3, df_exp["Vlr Pago"].sum(), fmt_tot)
        ws.write(l, 4, df_exp["Dif"].sum(), fmt_tot)
        
    return out.getvalue()

# ==============================================================================
# 2. INTERFACE GRÁFICA
# ==============================================================================

st.markdown("<h1 style='text-align: center;'>Conciliação de Retenções (Contábil)</h1>", unsafe_allow_html=True)
st.markdown("---")

c1, c2 = st.columns(2)

with c1: 
    st.markdown('<p class="big-label">1. Upload Razão Geral (.xlsx)</p>', unsafe_allow_html=True)
    arquivo = st.file_uploader("", type=["xlsx", "csv"], key="up_razao", label_visibility="collapsed")

with c2:
    st.markdown('<p class="big-label">2. Configuração dos Filtros</p>', unsafe_allow_html=True)
    placeholder_filtros = st.empty()

if arquivo:
    df_dados = carregar_dados(arquivo)
    
    if not df_dados.empty:
        ugs = sorted(df_dados[df_dados.columns[0]].astype(str).unique().tolist())
        contas = sorted(df_dados[df_dados.columns[6]].astype(str).unique().tolist())
        opcoes_conta = ["7852 - ISS Pessoa Jurídica (Padrão)"] + [c for c in contas if "7852" not in str(c)]
        
        with placeholder_filtros.container():
            col_a, col_b = st.columns(2)
            with col_a: ug_sel = st.selectbox("Selecione a UG", ugs)
            with col_b: conta_sel = st.selectbox("Conta de Retenção", opcoes_conta)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # BOTÃO COM LARGURA TOTAL
        if st.button("PROCESSAR CONCILIAÇÃO", use_container_width=True):
            with st.spinner("Processando..."):
                df_res, resumo = processar_conciliacao(df_dados, ug_sel, conta_sel)
            
            if not df_res.empty:
                
                # --- LINHA 1: CONTAGEM DE ITENS ---
                k1, k2, k3 = st.columns(3)
                with k1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Retido e Não Pago</div>
                        <div class="metric-value" style="color: #ff4b4b;">{resumo['ret_pendente']}</div>
                        <div class="metric-sub">Lançamentos Pendentes</div>
                    </div>
                    """, unsafe_allow_html=True)
                with k2:
                    st.markdown(f"""
                    <div class="metric-card metric-card-orange">
                        <div class="metric-label">Pago sem Retenção</div>
                        <div class="metric-value" style="color: #ffc107;">{resumo['pag_sobra']}</div>
                        <div class="metric-sub">Lançamentos de Sobra</div>
                    </div>
                    """, unsafe_allow_html=True)
                with k3:
                    st.markdown(f"""
                    <div class="metric-card metric-card-green">
                        <div class="metric-label">Conciliados</div>
                        <div class="metric-value" style="color: #28a745;">{resumo['ok']}</div>
                        <div class="metric-sub">Lançamentos Baixados</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # --- LINHA 2: VALORES ESPECÍFICOS ---
                v1, v2, v3 = st.columns(3)
                with v1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Total Retido s/ Pgto</div>
                        <div class="metric-value" style="color: #ff4b4b;">{formatar_moeda_br(resumo['val_ret_pendente'])}</div>
                        <div class="metric-sub">Montante Pendente</div>
                    </div>
                    """, unsafe_allow_html=True)
                with v2:
                    st.markdown(f"""
                    <div class="metric-card metric-card-orange">
                        <div class="metric-label">Total Pago s/ Retenção</div>
                        <div class="metric-value" style="color: #ffc107;">{formatar_moeda_br(resumo['val_pag_sobra'])}</div>
                         <div class="metric-sub">Montante Descasado</div>
                    </div>
                    """, unsafe_allow_html=True)
                with v3:
                    st.markdown(f"""
                    <div class="metric-card metric-card-green">
                        <div class="metric-label">Total Retido e Pago</div>
                        <div class="metric-value" style="color: #28a745;">{formatar_moeda_br(resumo['val_ok'])}</div>
                         <div class="metric-sub">Montante Conciliado</div>
                    </div>
                    """, unsafe_allow_html=True)

                # --- LINHA 3: TOTAIS GERAIS ---
                f1, f2, f3 = st.columns(3)
                cor_saldo = "#ff4b4b" if resumo['saldo'] > 0.01 else ("#28a745" if resumo['saldo'] == 0 else "#007bff")
                
                with f1:
                    st.markdown(f"""
                    <div class="metric-card metric-card-blue">
                        <div class="metric-label">Total Retido (Geral)</div>
                        <div class="metric-value" style="color: #004085;">{formatar_moeda_br(resumo['tot_ret'])}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with f2:
                    st.markdown(f"""
                    <div class="metric-card metric-card-blue">
                        <div class="metric-label">Total Pago (Geral)</div>
                        <div class="metric-value" style="color: #004085;">{formatar_moeda_br(resumo['tot_pag'])}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with f3:
                    st.markdown(f"""
                    <div class="metric-card metric-card-dark">
                        <div class="metric-label">Saldo (Diferença)</div>
                        <div class="metric-value" style="color: {cor_saldo};">{formatar_moeda_br(resumo['saldo'])}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # --- TABELA HTML (AJUSTE VISUAL: LARGURAS + QUEBRA DE LINHA NO HISTÓRICO) ---
                html = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>"
                html += "<table style='width:100%; border-collapse: collapse; color: black !important; background-color: white !important; table-layout: fixed;'>"
                
                html += "<tr style='background-color: black; color: white !important;'>"
                
                # Definição de Larguras (Total 100%)
                html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Empenho</th>"
                html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Data</th>"
                html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Vlr Retido</th>"
                html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Vlr Pago</th>"
                html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 12%;'>Diferença</th>"
                html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 34%;'>Histórico</th>"
                html += "<th style='padding: 8px; border: 1px solid #000; text-align: center; width: 10%;'>Status</th>"
                html += "</tr>"
                
                for _, r in df_res.iterrows():
                    dif = r['Dif']
                    style_dif = "color: red; font-weight: bold;" if abs(dif) > 0.01 else "color: black;"
                    
                    html += "<tr style='background-color: white;'>"
                    html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Empenho']}</td>"
                    html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Data Emp']}</td>"
                    html += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Retido'])}</td>"
                    html += f"<td style='border: 1px solid #000; text-align: right; color: black;'>{formatar_moeda_br(r['Vlr Pago'])}</td>"
                    html += f"<td style='border: 1px solid #000; text-align: right; {style_dif}'>{formatar_moeda_br(dif)}</td>"
                    
                    # HISTÓRICO: SEM TRUNCAMENTO, FONTE MENOR, QUEBRA DE LINHA
                    hist = str(r['Histórico'])
                    html += f"<td style='border: 1px solid #000; text-align: left; color: black; font-size: 11px; word-wrap: break-word; white-space: normal;'>{hist}</td>"
                    
                    html += f"<td style='border: 1px solid #000; text-align: center; color: black; font-size: 12px;'>{r['Status']}</td>"
                    html += "</tr>"
                
                html += f"<tr style='font-weight: bold; background-color: lightgrey; color: black;'>"
                html += "<td colspan='2' style='padding: 10px; text-align: center; border: 1px solid #000;'>TOTAL</td>"
                html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(resumo['tot_ret'])}</td>"
                html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(resumo['tot_pag'])}</td>"
                html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(resumo['saldo'])}</td>"
                html += "<td colspan='2' style='border: 1px solid #000;'></td></tr>"
                
                html += "</table></div>"
                st.markdown(html, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                excel_bytes = gerar_excel(df_res)
                st.download_button(
                    label="BAIXAR RELATÓRIO EM EXCEL",
                    data=excel_bytes,
                    file_name="Conciliacao_Retencoes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.warning("Nenhum dado encontrado para os filtros selecionados.")
