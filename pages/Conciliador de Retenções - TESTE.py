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
# 1. LISTAS FIXAS DE OPÇÕES
# ==============================================================================

LISTA_UGS = [
    "0 - PMB", "3 - FMS", "4 - FMAS", "5 - FMDCA", "6 - FME", "7 - FMMA",
    "10 - ARSEP", "11 - FMDPI", "12 - SEMER", "9999 - CONSOLIDADO"
]

LISTA_CONTAS = [
    "7812 - Salario Maternidade", "7814 - PENSÃO ALIMENTICIA", 
    "7815 - UNIMED Belem Coop. de Trabalho Medico", "7816 - UNIODONTO - Coop. de Trab. Odontologico",
    "7817 - ODONTOPREV", "7819 - Sindicato dos Trabalhadores em Educação", 
    "7821 - Sind. dos Agentes de Vig. de Barcarena", "7824 - Emp. Consignado BANPARA",
    "7826 - Emp. Cons. CAIXA ECONOMICA FEDERAL", "7827 - Emp. Cons. BANCO DO BRASIL",
    "7828 - Emp. Consignado SANTANDER", "7831 - A.M.P.E Barcarena - Di",
    "7832 - Desc. Autorizado PSDB 3%", "7837 - Desc. Aut. ASPEB",
    "7845 - IRRF DE SERVIÇOS DE TERCEIROS PJ", "7846 - IRRF DE SERV. DA ADM. DIR. E INDIRETA",
    "7847 - IRPF - Imposto de Renda da Pessoa Fisica", "7852 - ISS de Pessoa Juridica Retido na Fonte",
    "7853 - ISS De Pessoa Fisica Retido na Fonte", "7857 - INSS - Pessoa Fisica",
    "7858 - INSS FOPAG EFETIVOS", "7859 - INSS FOPAG TEMPORARIOS E COMISSIONADOS",
    "7864 - SALARIO FAMILIA", "7865 - INSS - Pessoa Juridica", "8926 - GARANTIA DE SAÚDE",
    "8931 - Emp. Consignado BRADESCO", "9032 - SEPUB -SINDICATO DOS SERVIDORES CIVIS DO PARÁ E MUNICIPIOS",
    "9159 - Emp. Cons.Kardbank", "9160 - Emp. Cons. Fydigital", "9210 - Emp. Consignado HBI - Scd"
]

# ==============================================================================
# 2. FUNÇÕES DE PROCESSAMENTO
# ==============================================================================

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == "": return "R$ 0,00"
    try:
        val_float = float(valor)
        return f"R$ {val_float:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    except: return "R$ 0,00"

def converter_moeda_input(texto):
    if not texto: return 0.0
    try:
        texto = str(texto).replace('R$', '').replace(' ', '')
        return float(texto.replace('.', '').replace(',', '.'))
    except: return 0.0

def formatar_data(dt):
    if pd.isna(dt): return "-"
    try: return dt.strftime("%d/%m/%Y")
    except: return str(dt)

def limpar_nome_arquivo(texto):
    nfkd_form = unicodedata.normalize('NFKD', str(texto))
    texto_sem_acento = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return re.sub(r'[\\/*?:"<>|]', '_', texto_sem_acento).strip()

@st.cache_data(show_spinner=False)
def carregar_dados(file):
    try:
        df = pd.read_excel(file, header=None)
    except:
        file.seek(0)
        try: df = pd.read_csv(file, sep=None, engine='python', encoding='latin1', header=None, on_bad_lines='skip')
        except: df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8', header=None, on_bad_lines='skip')
    
    if df.shape[1] < 35:
        for i in range(df.shape[1], 35): df[i] = pd.NA
    if 5 in df.columns:
        df[5] = df[5].astype(str).str.strip().str.upper()
        df = df[df[5].isin(['C', 'D'])]
    if 8 in df.columns:
        df[8] = df[8].apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x, str) else float(x or 0))
    if 4 in df.columns:
        df['Data_Dt'] = pd.to_datetime(df[4], dayfirst=True, errors='coerce')
    return df

def identificar_colunas_dinamicas(df):
    mapa = {'empenho': 14, 'tipo': 19, 'hist': 21}
    coluna_ab_tem_dados = df[27].head(50).notna().sum() > 0
    mapa['hist'] = 27 if coluna_ab_tem_dados else 21
    for idx, row in df.head(50).iterrows():
        for c in range(12, 16):
            if re.match(r'^\d{4}/\d+$', str(row[c])):
                mapa['empenho'] = c
                break
        for c in range(mapa['empenho'] + 1, mapa['empenho'] + 10):
            if any(k in str(row[c]).upper() for k in ["LIQUIDAÇÃO", "PAGAMENTO", "LANÇAMENTO", "RETENÇÃO", "ESTORNO"]):
                mapa['tipo'] = c
                break
    return mapa

def processar_conciliacao(df, ug_sel, conta_sel, saldo_anterior_val):
    cod_ug = ug_sel.split(' - ')[0].strip()
    cod_conta = conta_sel.split(' - ')[0].strip()
    cols_map = identificar_colunas_dinamicas(df)
    c_ug, c_data, c_dc, c_conta, c_valor = 0, 4, 5, 6, 8
    c_empenho, c_tipo, c_hist = cols_map['empenho'], cols_map['tipo'], cols_map['hist']

    for col in [c_ug, c_data, c_conta, c_empenho, c_tipo]: df[col] = df[col].ffill()

    if cod_ug == '9999': 
        mask_ug = pd.Series(True, index=df.index)
    else: 
        mask_ug = df[c_ug].apply(lambda x: str(x)[:-2] if str(x).endswith('.0') else str(x)) == str(cod_ug)
        
    mask_conta = df[c_conta].astype(str).str.startswith(str(cod_conta))
    df_base = df[mask_ug & mask_conta].copy()
    if df_base.empty: return pd.DataFrame(), {}

    df_base['Tipo_Norm'] = df_base[c_tipo].astype(str).str.strip()
    df_ret = df_base[(df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'C')].copy()
    df_est_ret = df_base[(df_base['Tipo_Norm'].str.contains("Retenção Empenho", case=False)) & (df_base[c_dc] == 'D')].copy()
    df_pag = df_base[(df_base['Tipo_Norm'].str.contains("Pagamento de Documento Extra", case=False)) & (df_base[c_dc] == 'D')].copy()
    df_est_pag = df_base[(df_base[c_dc] == 'C') & (df_base['Tipo_Norm'].str.contains("Estorno", case=False) | df_base[c_hist].astype(str).str.contains("Estorno", case=False))].copy()

    def cancelar_estornos(df_orig, df_est):
        cancelados = set()
        for _, r in df_est.iterrows():
            cand = df_orig[(df_orig[c_empenho] == r[c_empenho]) & (abs(df_orig[c_valor] - r[c_valor]) < 0.01) & (~df_orig.index.isin(cancelados))]
            if not cand.empty: cancelados.add(cand.index[0])
        return df_orig[~df_orig.index.isin(cancelados)]

    df_ret_limpa = cancelar_estornos(df_ret, df_est_ret)
    df_pag_limpa = cancelar_estornos(df_pag, df_est_pag)

    resultados, idx_pag_usado = [], set()
    resumo = {"ret_pendente": 0, "val_ret_pendente": 0.0, "pag_sobra": 0, "val_pag_sobra": 0.0, "ok": 0, "val_ok": 0.0, "tot_ret": 0.0, "tot_pag": 0.0, "saldo": 0.0}

    for _, r in df_ret_limpa.iterrows():
        val = r[c_valor]
        cand = df_pag_limpa[(df_pag_limpa[c_valor] == val) & (~df_pag_limpa.index.isin(idx_pag_usado)) & ((df_pag_limpa['Data_Dt'] >= r['Data_Dt']) | (df_pag_limpa['Data_Dt'].isna()))]
        val_pago, dt_pag_str, match, sort = 0.0, "-", False, 0
        hist_final = str(r[c_hist]).strip()
        if not cand.empty:
            r_pag = cand.iloc[0]
            val_pago, dt_pag_str, match, sort = r_pag[c_valor], formatar_data(r_pag[c_data]), True, 2
            idx_pag_usado.add(r_pag.name)
            resumo["ok"] += 1; resumo["val_ok"] += val_pago
            if str(r_pag[c_hist]).strip(): hist_final = str(r_pag[c_hist]).strip()
        else: resumo["ret_pendente"] += 1; resumo["val_ret_pendente"] += val
            
        resultados.append({"Empenho": r[c_empenho], "Data Emp": formatar_data(r[c_data]), "Vlr Retido": val, "Vlr Pago": val_pago, "Dif": val - val_pago, "Data Pag": dt_pag_str, "Histórico": hist_final, "_sort": sort, "_dt_sort": r['Data_Dt'], "Status": "Conciliado" if match else "Retido s/ Pagto"})

    for _, r in df_pag_limpa[~df_pag_limpa.index.isin(idx_pag_usado)].iterrows():
        resumo["pag_sobra"] += 1; resumo["val_pag_sobra"] += r[c_valor]
        # AJUSTE SOLICITADO: Para 'Pago s/ Retenção', exibe a data do pagamento em Data Emp
        resultados.append({"Empenho": r[c_empenho], "Data Emp": formatar_data(r[c_data]), "Vlr Retido": 0.0, "Vlr Pago": r[c_valor], "Dif": -r[c_valor], "Data Pag": formatar_data(r[c_data]), "Histórico": str(r[c_hist]).strip(), "_sort": 1, "_dt_sort": r['Data_Dt'], "Status": "Pago s/ Retenção"})

    if not resultados: return pd.DataFrame(), resumo
    df_final = pd.DataFrame(resultados).sort_values(by=['_sort', '_dt_sort']).drop(columns=['_dt_sort'])
    resumo.update({"tot_ret": df_final["Vlr Retido"].sum(), "tot_pag": df_final["Vlr Pago"].sum(), "saldo": df_final["Dif"].sum() + saldo_anterior_val})
    return df_final, resumo

def gerar_excel(df, resumo, saldo_anterior, ug, conta):
    out = io.BytesIO()
    df_exp = df.drop(columns=['_sort', 'Status'])
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df_exp.to_excel(writer, sheet_name='Conciliacao', index=False, startrow=8)
        wb, ws = writer.book, writer.sheets['Conciliacao']
        fmt_h = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_f = wb.add_format({'bold': True, 'bg_color': '#E6E6E6', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
        fmt_m = wb.add_format({'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
        fmt_r = wb.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'font_color': '#FF0000', 'valign': 'vcenter'})
        fmt_g = wb.add_format({'bold': True, 'num_format': '#,##0.00', 'border': 1, 'align': 'right', 'font_color': '#006400', 'valign': 'vcenter'})
        fmt_card_h = wb.add_format({'bold': True, 'bg_color': '#f0f0f0', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        
        ws.merge_range('A1:G1', f"UG: {ug}  |  CONTA: {conta}", fmt_f)
        ws.merge_range('A3:C3', 'RESUMO POR SITUAÇÃO (CARDS)', fmt_h)
        ws.write(3, 0, "CATEGORIA", fmt_card_h); ws.write(3, 1, "QTD", fmt_card_h); ws.write(3, 2, "VALOR", fmt_card_h)
        for i, (l, k, v) in enumerate([("Retido s/ Pagto (Pendente)", 'ret_pendente', 'val_ret_pendente'), ("Pago s/ Retenção (Sobra)", 'pag_sobra', 'val_pag_sobra'), ("Conciliados (OK)", 'ok', 'val_ok')]):
            ws.write(i+4, 0, l); ws.write(i+4, 1, resumo[k]); ws.write(i+4, 2, resumo[v], fmt_m)
        
        ws.merge_range('E3:F3', 'RESUMO FINANCEIRO (TOTAIS)', fmt_h)
        ws.write(3, 4, "SALDO ANTERIOR"); ws.write(3, 5, saldo_anterior, fmt_m)
        ws.write(4, 4, "TOTAL RETIDO"); ws.write(4, 5, resumo['tot_ret'], fmt_g)
        ws.write(5, 4, "TOTAL PAGO"); ws.write(5, 5, resumo['tot_pag'], fmt_r)
        
        cor_saldo = fmt_r if resumo['saldo'] > 0.001 else (fmt_g if resumo['saldo'] < -0.001 else fmt_m)
        ws.write(6, 4, "SALDO A PAGAR"); ws.write(6, 5, resumo['saldo'], cor_saldo)
        
        for i, col in enumerate(df_exp.columns):
            ws.write(8, i, col, fmt_h)
            ws.set_column(i, i, 50 if i==6 else (18 if i in [2,3,4] else 15), fmt_m if i in [2,3,4] else fmt_h)
    return out.getvalue()

def gerar_pdf(df_f, ug, conta, resumo, saldo_anterior):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm, title="Relatório")
    styles = getSampleStyleSheet()
    s_hist = ParagraphStyle(name='S', fontSize=6, leading=7, alignment=TA_LEFT)
    story = [Paragraph("Relatório de Conciliação de Retenções", styles["Title"]), Paragraph(f"<b>UG:</b> {ug}  |  <b>CONTA:</b> {conta}", ParagraphStyle(name='C', alignment=1, spaceAfter=10))]
    
    t_res = Table([["PENDENTES", "SOBRAS", "CONCILIADOS"], [f"{resumo['ret_pendente']} itens", f"{resumo['pag_sobra']} itens", f"{resumo['ok']} itens"], [formatar_moeda_br(resumo['val_ret_pendente']), formatar_moeda_br(resumo['val_pag_sobra']), formatar_moeda_br(resumo['val_ok'])]], colWidths=[63*mm]*3)
    t_res.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BACKGROUND', (0,0), (0,-1), colors.Color(1,0.9,0.9)), ('BACKGROUND', (1,0), (1,-1), colors.Color(1,0.95,0.8)), ('BACKGROUND', (2,0), (2,-1), colors.Color(0.9,1,0.9))]))
    story.append(t_res); story.append(Spacer(1, 3*mm))

    cor_s = colors.red if resumo['saldo'] > 0.01 else (colors.darkgreen if resumo['saldo'] < -0.01 else colors.black)
    t_tot = Table([["SALDO ANTERIOR", "TOTAL RETIDO", "TOTAL PAGO", "SALDO A PAGAR"], [formatar_moeda_br(saldo_anterior), formatar_moeda_br(resumo['tot_ret']), formatar_moeda_br(resumo['tot_pag']), formatar_moeda_br(resumo['saldo'])]], colWidths=[47*mm, 47*mm, 47*mm, 48*mm])
    t_tot.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'), ('BACKGROUND', (0,0), (-1,-1), colors.Color(0.9,0.95,1)), ('TEXTCOLOR', (1,1), (1,1), colors.darkgreen), ('TEXTCOLOR', (2,1), (2,1), colors.red), ('TEXTCOLOR', (3,1), (3,1), cor_s)]))
    story.append(t_tot); story.append(Spacer(1, 8*mm))
    
    data = [['Empenho', 'Data', 'Vlr. Retido', 'Vlr. Pago', 'Diferença', 'Histórico', 'Status']]
    for _, r in df_f.iterrows():
        data.append([str(r['Empenho']), str(r['Data Emp']), formatar_moeda_br(r['Vlr Retido']), formatar_moeda_br(r['Vlr Pago']), formatar_moeda_br(r['Dif']) if abs(r['Dif'])>=0.01 else "-", Paragraph(str(r['Histórico']), s_hist), str(r['Status'])])
    
    t = Table(data, colWidths=[17*mm, 14*mm, 25*mm, 25*mm, 25*mm, 59*mm, 24*mm], repeatRows=1)
    t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('FONTSIZE', (0,0), (-1,-1), 7)]))
    story.append(t); doc.build(story)
    return buffer.getvalue()

# ==============================================================================
# 2. INTERFACE GRÁFICA (RESTORED)
# ==============================================================================

st.markdown("<h1 style='text-align: center;'>Conciliação de Retenções</h1>", unsafe_allow_html=True)
st.markdown("---")
c1, c2 = st.columns(2)
with c1: 
    st.markdown('<p class="big-label">1. Upload Razão Contábil</p>', unsafe_allow_html=True)
    arquivo = st.file_uploader("", type=["xlsx", "csv"], key="up_razao", label_visibility="collapsed")
with c2:
    st.markdown('<p class="big-label">2. Filtros e Saldos</p>', unsafe_allow_html=True)
    ug_sel = st.selectbox("UG", LISTA_UGS)
    conta_sel = st.selectbox("Conta", LISTA_CONTAS)
    v_ant = st.text_input("Saldo Anterior", value="0,00")

if arquivo and st.button("PROCESSAR CONCILIAÇÃO", use_container_width=True):
    df_d = carregar_dados(arquivo)
    df_r, res = processar_conciliacao(df_d, ug_sel, conta_sel, converter_moeda_input(v_ant))
    if not df_r.empty:
        # Exibição dos Cards
        ck1, ck2, ck3 = st.columns(3)
        with ck1: st.markdown(f'<div class="metric-card"><div class="metric-label">Retido s/ Pgto</div><div class="metric-value" style="color:#ff4b4b;">{res["ret_pendente"]}</div></div>', unsafe_allow_html=True)
        with ck2: st.markdown(f'<div class="metric-card metric-card-orange"><div class="metric-label">Pago s/ Retenção</div><div class="metric-value" style="color:#ffc107;">{res["pag_sobra"]}</div></div>', unsafe_allow_html=True)
        with ck3: st.markdown(f'<div class="metric-card metric-card-green"><div class="metric-label">Conciliados</div><div class="metric-value" style="color:#28a745;">{res["ok"]}</div></div>', unsafe_allow_html=True)
        
        # Tabela HTML
        html = "<div style='background-color:white; padding:15px; border-radius:5px; border:1px solid #ddd;'>"
        html += "<table style='width:100%; border-collapse:collapse; color:black; background-color:white; table-layout:fixed;'>"
        html += "<tr style='background-color:black; color:white;'><th>Empenho</th><th>Data</th><th>Vlr Retido</th><th>Vlr Pago</th><th>Dif</th><th>Histórico</th><th>Status</th></tr>"
        for _, r in df_r.iterrows():
            sd = "color:darkgreen; font-weight:bold;" if r['Dif'] > 0.01 else ("color:red; font-weight:bold;" if r['Dif'] < -0.01 else "color:black;")
            html += f"<tr><td>{r['Empenho']}</td><td>{r['Data Emp']}</td><td style='text-align:right;'>{formatar_moeda_br(r['Vlr Retido'])}</td><td style='text-align:right;'>{formatar_moeda_br(r['Vlr Pago'])}</td><td style='text-align:right; {sd}'>{formatar_moeda_br(r['Dif'])}</td><td style='font-size:10px;'>{r['Histórico']}</td><td style='text-align:center;'>{r['Status']}</td></tr>"
        html += "</table></div>"
        st.markdown(html, unsafe_allow_html=True)
        
        st.download_button("BAIXAR EXCEL", gerar_excel(df_r, res, converter_moeda_input(v_ant), ug_sel, conta_sel), f"Conciliacao_{limpar_nome_arquivo(ug_sel)}.xlsx", use_container_width=True)
        st.download_button("BAIXAR PDF", gerar_pdf(df_r, ug_sel, conta_sel, res, converter_moeda_input(v_ant)), f"Conciliacao_{limpar_nome_arquivo(ug_sel)}.pdf", use_container_width=True)
    else: st.warning("Nenhum dado encontrado.")
