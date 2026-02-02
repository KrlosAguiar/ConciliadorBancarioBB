import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import textwrap
from datetime import datetime

# Bibliotecas para geração do PDF (ReportLab)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ==============================================================================
# 0. CONFIGURAÇÃO DA PÁGINA E CSS
# ==============================================================================

st.set_page_config(page_title="Conciliador de Receitas", layout="wide")

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
        width: 100% !important;
    }
    div.stButton > button:hover {
        background-color: rgb(20, 20, 25) !important;
        border-color: white;
    }

    .metric-card {
        background-color: #f8f9fa;
        border-left: 5px solid #ccc;
        padding: 15px;
        border-radius: 5px;
        color: black;
        border: 1px solid #ddd;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin-bottom: 15px;
        height: 100%;
    }
    .metric-card-green { border-left: 5px solid #28a745; }
    .metric-card-red { border-left: 5px solid #dc3545; }
    
    .metric-title { font-size: 13px; color: #000; text-transform: uppercase; font-weight: bold; margin-bottom: 8px; min-height: 35px; border-bottom: 2px solid #eee; padding-bottom: 5px;}
    .metric-row { display: flex; justify-content: space-between; font-size: 15px; margin-bottom: 4px; }
    .metric-val { font-weight: bold; }
    
    .metric-status { margin-top: 10px; padding: 5px; text-align: center; border-radius: 4px; font-weight: bold; font-size: 14px; }
    
    table { width: 100%; border-collapse: collapse; }
    th { background-color: black; color: white; padding: 10px; text-align: center; }
    td { padding: 8px; border-bottom: 1px solid #ddd; color: black; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE SUPORTE
# ==============================================================================

def formatar_moeda(valor):
    try:
        return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "R$ 0,00"

def limpar_numero(valor_str):
    if not valor_str: return 0.0
    try:
        limpo = re.sub(r'[^\d,]', '', str(valor_str)).replace(',', '.')
        return float(limpo)
    except:
        return 0.0

def obter_mes_referencia_extrato(texto_pdf):
    mapa_meses = {'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'ABRIL': '04', 'MAIO': '05', 'JUNHO': '06', 'JULHO': '07', 'AGOSTO': '08', 'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12'}
    texto_limpo = texto_pdf.replace('"', ' ').replace('\n', ' ')
    match = re.search(r'Mês:.*?([A-Za-zç]+)\s*/\s*(\d{4})', texto_limpo, re.IGNORECASE)
    if match:
        nome_mes, ano = match.groups()
        mes_num = mapa_meses.get(nome_mes.upper())
        if mes_num: return f"{mes_num}/{ano}"
    return None

def encontrar_arquivo_no_upload(lista_arquivos, termo_nome):
    if not lista_arquivos: return None
    for arq in lista_arquivos:
        if termo_nome in arq.name: return arq
    return None

# ==============================================================================
# 2. EXTRAÇÃO
# ==============================================================================

def extrair_bb(arquivo_bytes):
    total = 0.0
    try:
        with pdfplumber.open(arquivo_bytes) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text() or ""
                for linha in texto.split('\n'):
                    if "14109" in linha or "14020" in linha:
                        matches = re.findall(r'([\d\.]+,\d{2})', linha)
                        if matches: total += limpar_numero(matches[0])
    except: return None
    return total

def extrair_banpara(arquivo_bytes):
    total = 0.0
    termo = "REPAS ARRE PREF"
    try:
        with pdfplumber.open(arquivo_bytes) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text() or ""
                for linha in texto.split('\n'):
                    if termo in linha:
                        matches = re.findall(r'([\d\.]+,\d{2})', linha)
                        if matches: total += limpar_numero(matches[-1])
    except: return None
    return total

def extrair_caixa(arquivo_bytes):
    total = 0.0
    termos_regex = r"ARR\s+CCV\s+DH|ARR\s+CV\s+INT|ARR\s+DH\s+AG"
    try:
        with pdfplumber.open(arquivo_bytes) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                texto_completo += (pagina.extract_text() or "") + "\n"
            texto_limpo = texto_completo.replace('"', ' ') 
            mes_ref = obter_mes_referencia_extrato(texto_completo)
            regex = r'(\d{2}/\d{2}/\d{4}).*?(' + termos_regex + r')[^\d]+(\d{1,3}(?:\.\d{3})*,\d{2})\s*([CD])'
            matches = re.findall(regex, texto_limpo, re.IGNORECASE)
            for data_str, desc, valor_str, tipo in matches:
                mes_lancamento = data_str[3:] 
                if mes_ref is None or mes_lancamento == mes_ref:
                    valor_num = limpar_numero(valor_str)
                    if tipo.upper() == 'D': total -= valor_num
                    else: total += valor_num
    except: return None
    return total

# ==============================================================================
# 3. PDF
# ==============================================================================

def criar_tabela_card_pdf(titulo, lbl1, v1, lbl2, v2, larguras_colunas=[12*mm, 36*mm]):
    ok = round(v1, 2) == round(v2, 2)
    dif = abs(v1 - v2)
    cor_status = colors.HexColor("#28a745") if ok else colors.HexColor("#dc3545")
    bg_status = colors.HexColor("#e8f5e9") if ok else colors.HexColor("#fbe9eb")
    texto_status = "CONCILIADO" if ok else "NÃO CONCILIADO"
    if not ok: texto_status += f"\n(Dif: {formatar_moeda(dif)})"

    style_titulo = ParagraphStyle('T', fontSize=7, leading=8, textColor=colors.black, fontName='Helvetica-Bold', alignment=TA_LEFT)
    style_lbl = ParagraphStyle('L', fontSize=8, leading=9, textColor=colors.black, fontName='Helvetica', alignment=TA_LEFT)
    style_val = ParagraphStyle('V', fontSize=8, leading=9, textColor=colors.black, fontName='Helvetica-Bold', alignment=TA_RIGHT)
    style_sts = ParagraphStyle('S', fontSize=8, leading=9, textColor=cor_status, fontName='Helvetica-Bold', alignment=TA_CENTER)

    data = [
        [Paragraph(titulo.upper(), style_titulo), ''],
        [Paragraph(str(lbl1), style_lbl), Paragraph(formatar_moeda(v1), style_val)],
        [Paragraph(str(lbl2), style_lbl), Paragraph(formatar_moeda(v2), style_val)],
        [Paragraph(texto_status, style_sts), '']
    ]

    t = Table(data, colWidths=larguras_colunas, rowHeights=[10*mm, 6*mm, 6*mm, 10*mm])
    
    t.setStyle(TableStyle([
        ('SPAN', (0,0), (1,0)), ('SPAN', (0,3), (1,3)),
        ('BottomPadding', (0,0), (-1,0), 3),
        ('LINEBELOW', (0,0), (-1,0), 0.5, colors.lightgrey),
        ('BACKGROUND', (0,3), (-1,3), bg_status),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
        ('ROUNDEDCORNERS', [3, 3, 3, 3]),
    ]))
    return t

def gerar_pdf_completo(dados_cards_1, dados_cards_2, df_receitas):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=5*mm, leftMargin=5*mm, topMargin=10*mm, bottomMargin=10*mm, title="Relatório Conciliação")
    story = []
    styles = getSampleStyleSheet()

    story.append(Paragraph("RELATÓRIO DE CONCILIAÇÃO CONTÁBIL", styles["Title"]))
    story.append(Spacer(1, 8*mm))

    # --- 1. CARDS ---
    story.append(Paragraph("<b>1. SITUAÇÃO DAS TRANSFERÊNCIAS</b>", styles["Heading3"]))
    story.append(Spacer(1, 2*mm))

    # LINHA 1
    row1_tables = []
    width_l1 = [13*mm, 32*mm]
    for item in dados_cards_1:
        card = criar_tabela_card_pdf(item['titulo'], item['l1'], item['v1'], item['l2'], item['v2'], larguras_colunas=width_l1)
        row1_tables.append(card)
    
    master_table_1 = Table([row1_tables], colWidths=[49*mm]*4)
    master_table_1.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(master_table_1)
    story.append(Spacer(1, 4*mm))

    # LINHA 2
    row2_tables = []
    width_l2 = [25*mm, 35*mm]
    for item in dados_cards_2:
        card = criar_tabela_card_pdf(item['titulo'], item['l1'], item['v1'], item['l2'], item['v2'], larguras_colunas=width_l2)
        row2_tables.append(card)
    
    master_table_2 = Table([row2_tables], colWidths=[65*mm]*3)
    master_table_2.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(master_table_2)
    story.append(Spacer(1, 8*mm))

    # --- 2. TABELA ---
    story.append(Paragraph("<b>2. CONCILIAÇÃO DE RECEITAS PRÓPRIAS</b>", styles["Heading3"]))
    story.append(Spacer(1, 3*mm))

    data = [["Conta", "Valor Contábil", "Extrato Bancário", "Diferença"]]
    total_cont = 0.0
    total_ext = 0.0

    for _, row in df_receitas.iterrows():
        dif = row['Diferença']
        val_dif_str = formatar_moeda(dif)
        if abs(dif) < 0.01: val_dif_str = "OK"

        val_ext_str = formatar_moeda(row['Valor Extrato'])
        if row['Status'] == "Sem PDF": val_ext_str = "(Falta PDF)"

        data.append([str(row['Conta']), formatar_moeda(row['Valor Contábil']), val_ext_str, val_dif_str])
        total_cont += row['Valor Contábil']
        total_ext += row['Valor Extrato']

    data.append(["TOTAL GERAL", formatar_moeda(total_cont), formatar_moeda(total_ext), "-"])

    t = Table(data, colWidths=[60*mm, 45*mm, 45*mm, 40*mm])
    t_style = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.black),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (1,1), (2,-1), 'RIGHT'), 
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        
        # Estilo para deixar a COLUNA CONTA (índice 0) em Negrito nas linhas de dados
        ('FONTNAME', (0,1), (0,-2), 'Helvetica-Bold'),

        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
        ('TEXTCOLOR', (0,-1), (-1,-1), colors.black),
    ]

    for i in range(1, len(data)-1):
        val_dif = df_receitas.iloc[i-1]['Diferença']
        cor = colors.red if abs(val_dif) > 0.01 else colors.darkgreen
        t_style.append(('TEXTCOLOR', (3, i), (3, i), cor))
        t_style.append(('FONTNAME', (3, i), (3, i), 'Helvetica-Bold'))
        if df_receitas.iloc[i-1]['Status'] == "Sem PDF":
             t_style.append(('TEXTCOLOR', (2, i), (2, i), colors.grey))

    t.setStyle(TableStyle(t_style))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

# ==============================================================================
# 4. APLICAÇÃO
# ==============================================================================

st.markdown("<h1 style='text-align: center;'>Painel de Conciliação Contábil</h1>", unsafe_allow_html=True)
st.markdown("---")

c1, c2 = st.columns(2)
with c1:
    st.markdown("### 1. Razão Contábil (.xlsx)")
    arquivo_excel = st.file_uploader("", type=["xlsx"], key="up_excel", label_visibility="collapsed")
with c2:
    st.markdown("### 2. Extratos Bancários (.pdf)")
    arquivos_pdf = st.file_uploader("", type=["pdf"], accept_multiple_files=True, key="up_pdf", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if arquivo_excel:
    if st.button("INICIAR CONFERÊNCIA", use_container_width=True):
        with st.spinner("Processando dados e gerando relatórios..."):
            try:
                # --- LEITURA EXCEL ---
                df = pd.read_excel(arquivo_excel, skiprows=6, dtype=str)
                mask = df['UG'].str.contains("Totalizadores", case=False, na=False)
                if mask.any(): df = df.iloc[:mask.idxmax()].copy()
                df['Valor'] = df['Valor'].fillna('0')
                df['Valor'] = df['Valor'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce').fillna(0.0)
                df['Conta'] = df['Conta'].astype(str).str.replace(r'\.0$', '', regex=True)

                def get_vals(col, v1, v2):
                    f1 = df[col].astype(str).str.startswith(str(v1), na=False)
                    f2 = df[col].astype(str).str.startswith(str(v2), na=False)
                    return df.loc[f1, 'Valor'].sum(), df.loc[f2, 'Valor'].sum()

                # --- DADOS ---
                dados_c1 = []
                for tit, col, c1, c2 in [("TRANSF. CONST. - FMS", 'LCP', 264, 265), ("TRANSF. CONST. - FME", 'LCP', 266, 267), ("TRANSF. CONST. - FMAS", 'LCP', 268, 269), ("TRANSF. PARA ARSEP", 'LCP', 270, 271)]:
                    v1, v2 = get_vals(col, c1, c2)
                    dados_c1.append({'titulo': tit, 'l1': str(c1), 'v1': v1, 'l2': str(c2), 'v2': v2})

                dados_c2 = []
                v1, v2 = get_vals('LCP', 250, 251)
                dados_c2.append({'titulo': "TRANSFERÊNCIAS ENTRE UGS", 'l1': "250", 'v1': v1, 'l2': "251", 'v2': v2})
                v1, v2 = get_vals('LCP', 258, 259)
                dados_c2.append({'titulo': "RENDIMENTOS DE APLICAÇÃO", 'l1': "258", 'v1': v1, 'l2': "259", 'v2': v2})
                
                f1 = df['Fato Contábil'].astype(str).str.startswith("Transferência Financeira Concedida", na=False)
                f2 = df['Fato Contábil'].astype(str).str.startswith("Transferência Financeira Recebida", na=False)
                t1_duo = df.loc[f1, 'Valor'].sum()
                t2_duo = df.loc[f2, 'Valor'].sum()
                dados_c2.append({'titulo': "TRANSF. DUODÉCIMO", 'l1': "Concedida", 'v1': t1_duo, 'l2': "Recebida", 'v2': t2_duo})

                # --- RENDER TELA ---
                def render_card(d):
                    ok = round(d['v1'], 2) == round(d['v2'], 2)
                    dif = abs(d['v1'] - d['v2'])
                    cls_cor = "metric-card-green" if ok else "metric-card-red"
                    txt_st = "CONCILIADO" if ok else "NÃO CONCILIADO"
                    bg_st = "#e8f5e9" if ok else "#fbe9eb"
                    col_st = "#28a745" if ok else "#dc3545"
                    
                    dif_html = ""
                    if not ok: dif_html = f"<div style='font-size:11px; margin-top:2px; color:#c00'>(Dif: {formatar_moeda(dif)})</div>"

                    return textwrap.dedent(f"""<div class="metric-card {cls_cor}"><div class="metric-title">{d['titulo']}</div><div class="metric-row"><span>{d['l1']}</span><span class="metric-val">{formatar_moeda(d['v1'])}</span></div><div class="metric-row" style="border-bottom:1px dashed #ccc; padding-bottom:3px; margin-bottom:3px;"><span>{d['l2']}</span><span class="metric-val">{formatar_moeda(d['v2'])}</span></div><div class="metric-status" style="background:{bg_st}; color:{col_st};">{txt_st}{dif_html}</div></div>""").strip()

                st.markdown("### 1. Situação das Transferências")
                cols1 = st.columns(4)
                for i, d in enumerate(dados_c1):
                    with cols1[i]: st.markdown(render_card(d), unsafe_allow_html=True)
                
                cols2 = st.columns(3)
                for i, d in enumerate(dados_c2):
                    with cols2[i]: st.markdown(render_card(d), unsafe_allow_html=True)

                st.markdown("---")

                # --- RECEITAS ---
                # MAPA COM NOME DE EXIBIÇÃO CORRIGIDO
                mapa = {
                    '8346': {'key': '105628', 'f': extrair_bb, 'nome': '105628-X'},
                    '8416': {'key': '112005', 'f': extrair_bb, 'nome': '112005-0'},
                    '8364': {'key': '126022', 'f': extrair_bb, 'nome': '126022-7'},
                    '9150': {'key': '78101', 'f': extrair_bb, 'nome': '78101-0'},
                    '9130': {'key': '575230061', 'f': extrair_caixa, 'nome': '575230061-0'},
                    '8241': {'key': '538298', 'f': extrair_banpara, 'nome': '538298-0'}
                }
                
                ks = list(mapa.keys())
                df_res = df[(df['Fato Contábil']=='Arrecadação da Receita') & (df['Conta'].isin(ks))].groupby('Conta')['Valor'].sum().reset_index()
                df_final = pd.merge(pd.DataFrame({'Conta': ks}), df_res, on='Conta', how='left').fillna(0)

                recs = []
                for _, r in df_final.iterrows():
                    c_original, vc = r['Conta'], r['Valor']
                    cfg = mapa.get(c_original)
                    
                    # Usa o nome mapeado para exibição
                    c_display = cfg['nome']
                    
                    ve, stt = 0.0, "Sem PDF"
                    arq = encontrar_arquivo_no_upload(arquivos_pdf, cfg['key'])
                    if arq:
                        arq.seek(0)
                        val = cfg['f'](arq)
                        if val is not None: ve, stt = val, "OK"
                        else: stt = "Erro Leitura"
                    recs.append({"Conta": c_display, "PDF Ref": cfg['key'], "Valor Contábil": vc, "Valor Extrato": ve, "Diferença": vc - ve, "Status": stt})
                
                df_rec = pd.DataFrame(recs)
                tot_c = df_rec['Valor Contábil'].sum()
                tot_e = df_rec['Valor Extrato'].sum()

                st.markdown("### 2. Conciliação de Receitas Próprias")
                
                rows_html = ""
                for _, r in df_rec.iterrows():
                    dif = r['Diferença']
                    style_dif = "color:#28a745;font-weight:bold;" if abs(dif)<0.01 else "color:#dc3545;font-weight:bold;"
                    txt_dif = "CONCILIADO" if abs(dif)<0.01 else formatar_moeda(dif)
                    ve_str = formatar_moeda(r['Valor Extrato'])
                    if r['Status'] == "Sem PDF": 
                        ve_str = "<span style='color:#999;font-style:italic;'>(Falta PDF)</span>"
                        if r['Valor Contábil'] > 0: txt_dif = "PENDENTE"
                    
                    # AJUSTE: Conta em Negrito, Valor Contábil Normal
                    rows_html += f"<tr style='border-bottom:1px solid #eee;'><td style='font-weight:bold;'>{r['Conta']}</td><td style='color:#666;font-size:12px;'>{r['PDF Ref']}</td><td style='text-align:right;'>{formatar_moeda(r['Valor Contábil'])}</td><td style='text-align:right;'>{ve_str}</td><td style='text-align:center;{style_dif}'>{txt_dif}</td></tr>"

                tbl_html = textwrap.dedent(f"""<div style='background-color:white;padding:15px;border-radius:5px;border:1px solid #ddd;'><table style='width:100%;border-collapse:collapse;color:black !important;'><tr style='background-color:black;color:white;'><th>Conta</th><th>Ref. PDF</th><th style='text-align:right;'>Valor Contábil</th><th style='text-align:right;'>Valor Extrato</th><th style='text-align:center;'>Diferença</th></tr>{rows_html}<tr style='background-color:#f0f0f0;border-top:2px solid black;'><td colspan='2'><b>TOTAL GERAL</b></td><td style='text-align:right;'><b>{formatar_moeda(tot_c)}</b></td><td style='text-align:right;'><b>{formatar_moeda(tot_e)}</b></td><td></td></tr></table></div>""").strip()
                
                st.markdown(tbl_html, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)

                pdf_bytes = gerar_pdf_completo(dados_c1, dados_c2, df_rec)
                st.download_button("BAIXAR RELATÓRIO PDF COMPLETO", pdf_bytes, "Relatorio_Conciliacao_Completo.pdf", "application/pdf", use_container_width=True)

            except Exception as e:
                st.error(f"Ocorreu um erro no processamento: {e}")
else:
    st.info("Aguardando upload do Razão Contábil (.xlsx) para habilitar o início.")
