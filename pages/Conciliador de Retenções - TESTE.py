import streamlit as st
import pandas as pd
import io
import os
import re
from datetime import datetime
from PIL import Image

# Bibliotecas para geração do PDF (ReportLab)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm

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
    div.stButton > button {
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
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE PROCESSAMENTO
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
        texto = str(texto).replace('R$', '').replace(' ', '')
        return float(texto.replace('.', '').replace(',', '.'))
    except:
        return 0.0

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
    
    df = df.ffill(axis=0)
    
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

def processar_conciliacao(df, ug_sel, conta_sel, saldo_anterior_val):
    cod_conta = conta_sel.split(' - ')[0].strip()
    if not cod_conta.isdigit(): cod_conta = conta_sel.split(' ')[0]

    cols_map = identificar_colunas_dinamicas(df)
    c_ug, c_data, c_dc, c_conta, c_valor = 0, 4, 5, 6, 8
    c_empenho, c_tipo, c_hist = cols_map['empenho'], cols_map['tipo'], cols_map['hist']

    mask_ug = df[c_ug].astype(str) == str(ug_sel)
    mask_conta = df[c_conta].astype(str).str.startswith(str(cod_conta))
    df_base = df[mask_ug & mask_conta].copy()
    
    if df_base.empty: return pd.DataFrame(), {}

    df_base['Tipo_Norm'] = df_base[c_tipo].astype(str).str.strip()

    # Separação
    mask_ret = (df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'C')
    df_ret = df_base[mask_ret].copy()
    
    mask_estorno_ret = (df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'D')
    df_est_ret = df_base[mask_estorno_ret].copy()
    
    mask_pag = (df_base['Tipo_Norm'].str.contains("Pagamento de Documento Extra", case=False)) & (df_base[c_dc] == 'D')
    df_pag = df_base[mask_pag].copy()

    mask_estorno_pag = ((df_base[c_dc] == 'C') & (df_base['Tipo_Norm'].str.contains("Estorno", case=False) | df_base[c_hist].astype(str).str.contains("Estorno", case=False)))
    df_est_pag = df_base[mask_estorno_pag].copy()

    # Limpeza
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

    # Conciliação
    resultados = []
    idx_pag_usado = set()
    
    resumo = {
        "ret_pendente": 0, "val_ret_pendente": 0.0,
        "pag_sobra": 0,    "val_pag_sobra": 0.0,
        "ok": 0,           "val_ok": 0.0,
        "tot_ret": 0.0, "tot_pag": 0.0, "saldo": 0.0
    }

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
        
        val_pago, dt_pag, match, sort = 0.0, "-", False, 0
        
        if not cand.empty:
            r_pag = cand.iloc[0]
            val_pago, dt_pag = r_pag[c_valor], r_pag[c_data]
            idx_pag_usado.add(r_pag.name)
            match, sort = True, 2
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

    for _, r in df_pag_limpa[~df_pag_limpa.index.isin(idx_pag_usado)].iterrows():
        resumo["pag_sobra"] += 1
        resumo["val_pag_sobra"] += r[c_valor]
        resultados.append({
            "Empenho": r[c_empenho], "Data Emp": r[c_data],
            "Vlr Retido": 0.0, "Vlr Pago": r[c_valor],
            "Dif": 0.0 - r[c_valor], "Data Pag": r[c_data],
            "Histórico": r[c_hist], "_sort": 1,
            "Status": "Pago s/ Retenção"
        })

    if not resultados: return pd.DataFrame(), resumo

    df_res = pd.DataFrame(resultados).sort_values(by=['_sort', 'Data Emp', 'Data Pag'])
    
    # Totais
    resumo["tot_ret"] = df_res["Vlr Retido"].sum()
    resumo["tot_pag"] = df_res["Vlr Pago"].sum()
    diferenca_tabela = df_res["Dif"].sum()
    resumo["saldo"] = diferenca_tabela + saldo_anterior_val
    
    return df_res, resumo

def gerar_excel(df, resumo, saldo_anterior):
    out = io.BytesIO()
    df_exp = df.drop(columns=['_sort', 'Status'])
    
    # Início da tabela de dados (deixando espaço para o resumo e totais no topo)
    # Linhas: 0=Título, 1=HeaderCards, 2-4=DadosCards, 5=Espaço, 
    # 6=HeaderTotais, 7=DadosTotais, 8=Espaço, 9=HeaderTabela
    start_row_table = 9
    
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df_exp.to_excel(writer, sheet_name='Conciliacao', index=False, startrow=start_row_table)
        wb = writer.book
        ws = writer.sheets['Conciliacao']
        
        # Formatos
        fmt_head = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center'})
        fmt_money = wb.add_format({'num_format': '#,##0.00'})
        fmt_green = wb.add_format({'font_color': '#006400', 'bold': True, 'num_format': '#,##0.00'}) # Verde Escuro
        fmt_red = wb.add_format({'font_color': '#FF0000', 'bold': True, 'num_format': '#,##0.00'})   # Vermelho
        fmt_tot_head = wb.add_format({'bold': True, 'bg_color': '#E6E6E6', 'border': 1, 'align': 'center'})
        fmt_tot_val = wb.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'center'})
        fmt_card_label = wb.add_format({'bold': True, 'bg_color': '#f0f0f0', 'border': 1, 'align': 'left'})
        fmt_card_val = wb.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right'})
        
        # --- 1. QUADRO DE RESUMO (CARDS) ---
        ws.merge_range('A1:C1', 'RESUMO POR SITUAÇÃO (CARDS)', fmt_head)
        ws.write(1, 0, "CATEGORIA", fmt_card_label)
        ws.write(1, 1, "QTD", fmt_card_label)
        ws.write(1, 2, "VALOR", fmt_card_label)
        
        ws.write(2, 0, "Retido s/ Pagto (Pendente)", fmt_card_label)
        ws.write(2, 1, resumo['ret_pendente'], fmt_card_val)
        ws.write(2, 2, resumo['val_ret_pendente'], fmt_money)
        
        ws.write(3, 0, "Pago s/ Retenção (Sobra)", fmt_card_label)
        ws.write(3, 1, resumo['pag_sobra'], fmt_card_val)
        ws.write(3, 2, resumo['val_pag_sobra'], fmt_money)
        
        ws.write(4, 0, "Conciliados (OK)", fmt_card_label)
        ws.write(4, 1, resumo['ok'], fmt_card_val)
        ws.write(4, 2, resumo['val_ok'], fmt_money)

        # --- 2. QUADRO DE TOTAIS E SALDOS (AGORA NO TOPO) ---
        ws.merge_range('E1:H1', 'RESUMO FINANCEIRO (TOTAIS)', fmt_head)
        totais_headers = ["TOTAL RETIDO", "TOTAL PAGO", "SALDO ANTERIOR", "SALDO A PAGAR"]
        totais_values = [resumo['tot_ret'], resumo['tot_pag'], saldo_anterior, resumo['saldo']]
        
        for i, h in enumerate(totais_headers):
            ws.write(1, 4 + i, h, fmt_tot_head) # Começa na coluna E (index 4)
            ws.write(2, 4 + i, totais_values[i], fmt_tot_val)

        # --- 3. FORMATAÇÃO DA TABELA E AUTOAJUSTE ---
        # Formatar cabeçalho
        for i, col in enumerate(df_exp.columns):
            ws.write(start_row_table, i, col, fmt_head)
            
            # AUTOAJUSTE DE COLUNAS
            # Calcula o tamanho máximo do conteúdo na coluna
            max_len = len(str(col)) # Começa com tamanho do header
            for val in df_exp[col]:
                val_len = len(str(val))
                if val_len > max_len: max_len = val_len
            
            # Limite máximo para histórico não ficar gigante
            if max_len > 60: max_len = 60
            if max_len < 12: max_len = 12 # Mínimo
            
            ws.set_column(i, i, max_len + 2) # +2 para respiro

        # Formatar colunas de valor
        ws.set_column('C:E', 18, fmt_money)
        
        # --- 4. FORMATAÇÃO CONDICIONAL DA DIFERENÇA ---
        # Range da coluna "Dif" (Coluna E -> index 4)
        first_row = start_row_table + 1
        last_row = start_row_table + len(df_exp)
        
        # Regra 1: Maior que 0 (Retido s/ Pgto) -> VERDE ESCURO
        ws.conditional_format(first_row, 4, last_row, 4, {
            'type': 'cell',
            'criteria': '>',
            'value': 0.001,
            'format': fmt_green
        })
        
        # Regra 2: Menor que 0 (Pago s/ Retenção) -> VERMELHO
        ws.conditional_format(first_row, 4, last_row, 4, {
            'type': 'cell',
            'criteria': '<',
            'value': -0.001,
            'format': fmt_red
        })
        
    return out.getvalue()

def gerar_pdf(df_f, titulo_conta, resumo, saldo_anterior):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    story = []
    styles = getSampleStyleSheet()
    
    story.append(Paragraph("Relatório de Conciliação de Retenções", styles["Title"]))
    story.append(Paragraph(f"<b>Filtro:</b> {titulo_conta}", ParagraphStyle(name='C', alignment=1, spaceAfter=10)))
    
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
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('FONTSIZE', (0,1), (-1,2), 11),
        ('BACKGROUND', (0,0), (0,-1), bg_red),
        ('BACKGROUND', (1,0), (1,-1), bg_org),
        ('BACKGROUND', (2,0), (2,-1), bg_grn),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_res)
    story.append(Spacer(1, 3*mm))

    data_totais = [
        ["TOTAL RETIDO", "TOTAL PAGO", "SALDO ANTERIOR", "SALDO A PAGAR"],
        [
            formatar_moeda_br(resumo['tot_ret']), 
            formatar_moeda_br(resumo['tot_pag']), 
            formatar_moeda_br(saldo_anterior),
            formatar_moeda_br(resumo['saldo'])
        ]
    ]
    t_tot = Table(data_totais, colWidths=[47*mm, 47*mm, 47*mm, 48*mm])
    t_tot.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('BACKGROUND', (0,0), (-1,-1), bg_blu),
        ('TEXTCOLOR', (3,1), (3,1), colors.red if resumo['saldo'] > 0.01 else (colors.blue if resumo['saldo'] < -0.01 else colors.green)),
    ]))
    story.append(t_tot)
    story.append(Spacer(1, 8*mm))
    
    headers = ['Empenho', 'Data', 'Vlr. Retido', 'Vlr. Pago', 'Diferença', 'Histórico', 'Status']
    data = [headers]
    
    table_style = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.black),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (2,0), (4,-1), 'RIGHT'),
        ('ALIGN', (5,0), (5,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 7),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]
    
    for i, (_, r) in enumerate(df_f.iterrows()):
        dif = r['Dif']
        hist = str(r['Histórico'])
        if len(hist) > 60: hist = hist[:57] + "..."
        
        row_data = [
            str(r['Empenho']), str(r['Data Emp']),
            formatar_moeda_br(r['Vlr Retido']), formatar_moeda_br(r['Vlr Pago']),
            formatar_moeda_br(dif) if abs(dif) >= 0.01 else "-",
            hist, str(r['Status'])
        ]
        data.append(row_data)
        
        if abs(dif) >= 0.01:
            # Cor verde para dif positiva (retido s/ pagto), vermelha para negativa
            cor_fonte = colors.darkgreen if dif > 0 else colors.red
            table_style.append(('TEXTCOLOR', (4, i+1), (4, i+1), cor_fonte))
            table_style.append(('FONTNAME', (4, i+1), (4, i+1), 'Helvetica-Bold'))

    # Totais no rodapé do PDF também, mas sem Saldo Anterior/A Pagar pois já está no topo
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

# ==============================================================================
# 2. INTERFACE GRÁFICA
# ==============================================================================

st.markdown("<h1 style='text-align: center;'>Conciliação de Retenções</h1>", unsafe_allow_html=True)
st.markdown("---")

c_top_1, c_top_2 = st.columns(2)

with c_top_1: 
    st.markdown('<p class="big-label">1. Upload Razão Contábil</p>', unsafe_allow_html=True)
    arquivo = st.file_uploader("", type=["xlsx", "csv"], key="up_razao", label_visibility="collapsed")

with c_top_2:
    st.markdown('<p class="big-label">2. Filtros e Saldos</p>', unsafe_allow_html=True)
    placeholder_filtros = st.empty()

if arquivo:
    df_dados = carregar_dados(arquivo)
    
    if not df_dados.empty:
        ugs = sorted(df_dados[0].astype(str).unique().tolist())
        contas = sorted(df_dados[6].astype(str).unique().tolist())
        opcoes_conta = ["7852 - ISS Pessoa Jurídica (Padrão)"] + [c for c in contas if "7852" not in str(c)]
        
        with placeholder_filtros.container():
            r1_col1, r1_col2 = st.columns([1, 3]) 
            with r1_col1: 
                ug_sel = st.selectbox("UG", ugs)
            with r1_col2: 
                conta_sel = st.selectbox("Conta de Retenção", opcoes_conta)
            
            val_anterior_str = st.text_input("Saldo Anterior", value="0,00", help="Digite o saldo acumulado de períodos anteriores.")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("PROCESSAR CONCILIAÇÃO", use_container_width=True):
            saldo_ant_float = converter_moeda_input(val_anterior_str)
            
            with st.spinner("Processando..."):
                df_res, resumo = processar_conciliacao(df_dados, ug_sel, conta_sel, saldo_ant_float)
            
            if not df_res.empty:
                
                c_k1, c_k2, c_k3 = st.columns(3)
                with c_k1:
                    st.markdown(f"""<div class="metric-card"><div class="metric-label">Retido e Não Pago</div><div class="metric-value" style="color: #ff4b4b;">{resumo['ret_pendente']}</div></div>""", unsafe_allow_html=True)
                with c_k2:
                    st.markdown(f"""<div class="metric-card metric-card-orange"><div class="metric-label">Pago sem Retenção</div><div class="metric-value" style="color: #ffc107;">{resumo['pag_sobra']}</div></div>""", unsafe_allow_html=True)
                with c_k3:
                    st.markdown(f"""<div class="metric-card metric-card-green"><div class="metric-label">Conciliados</div><div class="metric-value" style="color: #28a745;">{resumo['ok']}</div></div>""", unsafe_allow_html=True)
                
                v1, v2, v3 = st.columns(3)
                with v1:
                    st.markdown(f"""<div class="metric-card"><div class="metric-label">Total Retido s/ Pgto</div><div class="metric-value" style="color: #ff4b4b;">{formatar_moeda_br(resumo['val_ret_pendente'])}</div></div>""", unsafe_allow_html=True)
                with v2:
                    st.markdown(f"""<div class="metric-card metric-card-orange"><div class="metric-label">Total Pago s/ Retenção</div><div class="metric-value" style="color: #ffc107;">{formatar_moeda_br(resumo['val_pag_sobra'])}</div></div>""", unsafe_allow_html=True)
                with v3:
                    st.markdown(f"""<div class="metric-card metric-card-green"><div class="metric-label">Total Retido e Pago</div><div class="metric-value" style="color: #28a745;">{formatar_moeda_br(resumo['val_ok'])}</div></div>""", unsafe_allow_html=True)

                f1, f2, f3, f4 = st.columns(4)
                cor_saldo = "#ff4b4b" if resumo['saldo'] > 0.01 else ("#28a745" if resumo['saldo'] == 0 else "#007bff")
                
                with f1:
                    st.markdown(f"""<div class="metric-card metric-card-blue"><div class="metric-label">Total Retido (Período)</div><div class="metric-value" style="color: #004085;">{formatar_moeda_br(resumo['tot_ret'])}</div></div>""", unsafe_allow_html=True)
                with f2:
                    st.markdown(f"""<div class="metric-card metric-card-blue"><div class="metric-label">Total Pago (Período)</div><div class="metric-value" style="color: #004085;">{formatar_moeda_br(resumo['tot_pag'])}</div></div>""", unsafe_allow_html=True)
                with f3:
                    st.markdown(f"""<div class="metric-card metric-card-dark"><div class="metric-label">Saldo Anterior</div><div class="metric-value" style="color: #343a40;">{formatar_moeda_br(saldo_ant_float)}</div></div>""", unsafe_allow_html=True)
                with f4:
                    st.markdown(f"""<div class="metric-card metric-card-dark"><div class="metric-label">Saldo a Pagar</div><div class="metric-value" style="color: {cor_saldo};">{formatar_moeda_br(resumo['saldo'])}</div></div>""", unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
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
                
                for _, r in df_res.iterrows():
                    dif = r['Dif']
                    # Cor condicional na tela também, para consistência com o pedido
                    style_dif = "color: darkgreen; font-weight: bold;" if dif > 0.01 else ("color: red; font-weight: bold;" if dif < -0.01 else "color: black;")
                    
                    html += "<tr style='background-color: white;'>"
                    html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Empenho']}</td>"
                    html += f"<td style='border: 1px solid #000; text-align: center; color: black;'>{r['Data Emp']}</td>"
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
                
                excel_bytes = gerar_excel(df_res, resumo, saldo_ant_float)
                st.download_button(
                    label="BAIXAR RELATÓRIO EM EXCEL",
                    data=excel_bytes,
                    file_name="Conciliacao_Retencoes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                pdf_bytes = gerar_pdf(df_res, f"{conta_sel} (UG: {ug_sel})", resumo, saldo_ant_float)
                st.download_button(
                    label="BAIXAR RELATÓRIO EM PDF",
                    data=pdf_bytes,
                    file_name="Relatorio_Retencoes.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.warning("Nenhum dado encontrado para os filtros selecionados.")
