import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import datetime
import xlsxwriter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
from PIL import Image
import fitz

# --- CONFIGURAÇÃO DA PÁGINA ---
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
try:
    icon_image = Image.open(icon_path)
    st.set_page_config(page_title="Portal Financeiro", page_icon=icon_image, layout="wide")
except:
    st.set_page_config(page_title="Portal Financeiro", layout="wide")

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
    }
    div.stButton > button:hover { background-color: rgb(20, 20, 25) !important; border-color: white; }
    .big-label { font-size: 24px !important; font-weight: 600 !important; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE PROCESSAMENTO
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

def processar_pdf(file_bytes):
    rows_debitos = []
    rows_devolucoes = []
    
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
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
                        
                        coord_box = None
                        for w in linha_words:
                            if valor_bruto in w['text']:
                                coord_box = (page_idx, w['x0'], w['top'], w['x1'], w['bottom'])
                                break

                        texto_sem_data = texto_linha.replace(match_data.group(0), "", 1).strip()
                        texto_sem_valor = texto_sem_data.replace(match_valor.group(0), "").strip()
                        
                        entry = {
                            "Data": data_str, "Histórico": texto_sem_valor.strip(),
                            "Documento": "", "Valor_Extrato": valor_float, "coords": coord_box
                        }

                        if tipo == 'D':
                            tokens = texto_sem_valor.split()
                            if tokens:
                                for t in reversed(tokens):
                                    limpo = t.replace('.', '').replace('-', '')
                                    if limpo.isdigit() and len(limpo) >= 4:
                                        entry["Documento"] = limpar_documento_pdf(t)
                                        break
                            rows_debitos.append(entry)
                        elif tipo == 'C':
                            hist_upper = texto_sem_valor.upper()
                            if any(x in hist_upper for x in ["TED DEVOLVIDA", "DEVOLUCAO DE TED", "TED DEVOL"]):
                                rows_devolucoes.append(entry)
                                
        df_debitos = pd.DataFrame(rows_debitos)
        coords_referencia = rows_debitos + rows_devolucoes
        
    except:
        return pd.DataFrame(), []

    if not rows_devolucoes == [] and not df_debitos.empty:
        idx_rem = []
        for r_dev in rows_devolucoes:
            m = df_debitos[(df_debitos['Data'] == r_dev['Data']) & (abs(df_debitos['Valor_Extrato'] - r_dev['Valor_Extrato']) < 0.01) & (~df_debitos.index.isin(idx_rem))]
            if not m.empty: idx_rem.append(m.index[0])
        df_debitos = df_debitos.drop(idx_rem).reset_index(drop=True)

    termos_excluir = "SALDO|S A L D O|Resgate|BB-APLIC C\.PRZ-APL\.AUT|1\.972"
    df = df_debitos[~df_debitos['Histórico'].astype(str).str.contains(termos_excluir, case=False, na=False)].copy()
    
    df['Data_dt'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    mask_13113 = df['Histórico'].astype(str).str.contains("13113", na=False)
    if any(mask_13113):
        df_t = df[mask_13113].copy(); df_o = df[~mask_13113].copy()
        df_t_agg = df_t.groupby('Data_dt').agg({'Valor_Extrato': 'sum', 'Data': 'first'}).reset_index()
        df_t_agg['Documento'] = "Tarifas Bancárias"; df_t_agg['Histórico'] = "Tarifas Bancárias do Dia"
        df = pd.concat([df_o, df_t_agg], ignore_index=True)
    
    return df, coords_referencia

def processar_excel_detalhado(file_bytes, df_pdf_ref, is_csv=False):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='latin1', sep=None, engine='python') if is_csv else pd.read_excel(io.BytesIO(file_bytes), header=None)
        
        try: df = df.iloc[:, [4, 5, 8, 25, 26, 27]].copy()
        except: df = df.iloc[:, [4, 5, 8, -4, -2, -1]].copy()
        
        df.columns = ['Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
        df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x, str) else float(x))
        
        mask_pagto = df['Info_Z'].astype(str).str.contains("Pagamento", case=False, na=False)
        mask_transf_std = (df['Info_Z'].astype(str).str.contains("TRANSFERENCIA ENTRE CONTAS DE MESMA UG", case=False, na=False)) & (df['DC'].str.strip().str.upper() == 'C')
        mask_codes_z = df['Info_Z'].astype(str).str.contains(r"266|264|268", case=False, regex=True, na=False)
        cond_250_z = df['Info_Z'].astype(str).str.contains("250", case=False, na=False)
        cond_ab_text = df['Info_AB'].astype(str).str.contains("transferência financeira concedida|repasse financeiro concedido", case=False, na=False)
        mask_250_restrict = cond_250_z & cond_ab_text
        mask_aa_ded = df['Info_AA'].astype(str).str.contains(r"Ded\.", case=False, regex=True, na=False)
        
        df_filtered = df[mask_pagto | mask_transf_std | mask_codes_z | mask_250_restrict | mask_aa_ded].copy()
        
        termos_estorno = r"Est Pgto Ext|Est Pagto"
        mask_eh_estorno = df_filtered['Info_AA'].astype(str).str.contains(termos_estorno, case=False, regex=True, na=False)
        
        df_estornos = df_filtered[mask_eh_estorno].copy()
        df_validos = df_filtered[~mask_eh_estorno].copy()
        
        indices_para_remover = []
        indices_usados_validos = set()
        
        for idx_est, row_est in df_estornos.iterrows():
            valor_est = row_est['Valor_Razao']
            candidatos = df_validos[(abs(df_validos['Valor_Razao'] - valor_est) < 0.01) & (~df_validos.index.isin(indices_usados_validos))]
            if not candidatos.empty:
                idx_par = candidatos.index[0]
                indices_para_remover.append(idx_par)
                indices_usados_validos.add(idx_par)
                indices_para_remover.append(idx_est)
            else:
                indices_para_remover.append(idx_est)

        df_final = df_filtered.drop(indices_para_remover, errors='ignore').copy()
        
        df_final['Data_dt'] = df_final['Data'].apply(parse_br_date)
        df_final = df_final.dropna(subset=['Data_dt'])
        df_final['Data'] = df_final['Data_dt'].dt.strftime('%d/%m/%Y')
        
        lookup = {dt: {str(doc).lstrip('0'): doc for doc in g['Documento'].unique()} for dt, g in df_pdf_ref.groupby('Data')}
        lookup_ded = {}
        if not df_pdf_ref.empty:
            mask_pdf_ded = df_pdf_ref['Histórico'].astype(str).str.contains(r"Dedução|Ded\.|FUNDEB|PASEP", case=False, regex=True, na=False)
            for idx, row in df_pdf_ref[mask_pdf_ded].iterrows():
                if row['Data'] not in lookup_ded:
                    lookup_ded[row['Data']] = row['Documento']
        
        def find_doc(row):
            txt = str(row['Info_AB']).upper()
            dt = row['Data']
            info_aa = str(row['Info_AA']).upper()
            if "DED." in info_aa and dt in lookup_ded: return lookup_ded[dt]
            if dt not in lookup: return "S/D"
            if "TARIFA" in txt and "Tarifas Bancárias" in lookup[dt].values(): return "Tarifas Bancárias"
            for n in re.findall(r'\d+', txt):
                if n.lstrip('0') in lookup[dt]: return lookup[dt][n.lstrip('0')]
            return "NÃO LOCALIZADO"
            
        df_final['Documento'] = df_final.apply(find_doc, axis=1)
        
        # --- CORREÇÃO IMPORTANTE: Resetar índice para evitar KeyError posterior ---
        return df_final.reset_index(drop=True)[['Data', 'Documento', 'Valor_Razao']]
        
    except Exception as e:
        return pd.DataFrame()

def executar_conciliacao_inteligente(df_pdf, df_excel):
    """
    Algoritmo de Conciliação em 3 Etapas (Hierárquico):
    1. Match Exato (Data + Doc + Valor): Para itens perfeitos (ex: PASEP).
    2. Match Grupo (Data + Doc + Soma): Para itens fragmentados (ex: FUNDEB).
    3. Match Flexível (Data + Valor): Para itens onde o Doc não bate (ex: Transferências).
    """
    res = []
    idx_p_u = set()
    idx_e_u = set()

    # 1. MATCH EXATO (Prioridade Máxima)
    for idx_p, row_p in df_pdf.iterrows():
        cand = df_excel[
            (df_excel['Data'] == row_p['Data']) & 
            (df_excel['Documento'] == row_p['Documento']) & 
            (~df_excel.index.isin(idx_e_u))
        ]
        match_valor = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not match_valor.empty:
            idx_e = match_valor.index[0]
            val_excel = match_valor.loc[idx_e]['Valor_Razao']
            res.append({
                'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'],
                'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': val_excel, 'Diferença': 0.0
            })
            idx_p_u.add(idx_p)
            idx_e_u.add(idx_e)

    # 2. MATCH POR GRUPO (Soma) - Para FUNDEB
    excel_restante = df_excel[~df_excel.index.isin(idx_e_u)].copy()
    if not excel_restante.empty:
        # Cria mapa de índices para evitar erros de busca
        mapa_grupos = {k: v.index.tolist() for k, v in excel_restante.groupby(['Data', 'Documento'])}
        
        for idx_p, row_p in df_pdf.iterrows():
            if idx_p in idx_p_u: continue
            
            chave = (row_p['Data'], row_p['Documento'])
            if chave in mapa_grupos:
                indices_grupo = mapa_grupos[chave]
                indices_disponiveis = [i for i in indices_grupo if i not in idx_e_u]
                
                if not indices_disponiveis: continue
                
                soma_excel = df_excel.loc[indices_disponiveis, 'Valor_Razao'].sum()
                
                if abs(soma_excel - row_p['Valor_Extrato']) < 1.00:
                    res.append({
                        'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'],
                        'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': soma_excel,
                        'Diferença': row_p['Valor_Extrato'] - soma_excel
                    })
                    idx_p_u.add(idx_p)
                    idx_e_u.update(indices_disponiveis)

    # 3. MATCH FLEXÍVEL (Apenas Valor) - Para Transferências com Doc diferente
    # Varre o que sobrou do PDF
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p in idx_p_u: continue
        
        # Procura no Excel items com MESMA DATA e MESMO VALOR, ignorando Documento
        cand_flex = df_excel[
            (df_excel['Data'] == row_p['Data']) &
            (abs(df_excel['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01) &
            (~df_excel.index.isin(idx_e_u))
        ]
        
        if not cand_flex.empty:
            idx_e = cand_flex.index[0]
            val_excel = cand_flex.loc[idx_e]['Valor_Razao']
            res.append({
                'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'],
                'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': val_excel, 'Diferença': 0.0
            })
            idx_p_u.add(idx_p)
            idx_e_u.add(idx_e)

    # 4. SOBRAS
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p not in idx_p_u:
             res.append({
                'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'],
                'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': 0.0, 'Diferença': row_p['Valor_Extrato']
            })
             
    excel_sobra = df_excel[~df_excel.index.isin(idx_e_u)]
    if not excel_sobra.empty:
        agg_sobra = excel_sobra.groupby(['Data', 'Documento'])['Valor_Razao'].sum().reset_index()
        for _, row_e in agg_sobra.iterrows():
             res.append({
                'Data': row_e['Data'], 'Histórico': "Não Conciliado (Razão)", 'Documento': row_e['Documento'],
                'Valor_Extrato': 0.0, 'Valor_Razao': row_e['Valor_Razao'], 'Diferença': -row_e['Valor_Razao']
            })

    df_f = pd.DataFrame(res)
    if df_f.empty: return pd.DataFrame(columns=['Data', 'Histórico', 'Documento', 'Valor_Extrato', 'Valor_Razao', 'Diferença'])
    
    df_f['dt'] = pd.to_datetime(df_f['Data'], format='%d/%m/%Y', errors='coerce')
    return df_f.sort_values(by=['dt', 'Documento']).drop(columns=['dt'])

# ==============================================================================
# 2. GERAÇÃO DE SAÍDAS
# ==============================================================================
def gerar_pdf_final(df_f, titulo_completo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm, title=titulo_completo)
    story = []
    styles = getSampleStyleSheet()
    story.append(Paragraph("Relatório de Conciliação Bancária", styles["Title"]))
    nome_conta_interno = titulo_completo.replace("Conciliação ", "")
    story.append(Paragraph(f"<b>Conta:</b> {nome_conta_interno}", ParagraphStyle(name='C', alignment=1)))
    story.append(Spacer(1, 15))
    headers = ['Data', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']
    data = [headers]
    table_style = [('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.black), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('ALIGN', (0,0), (0,-1), 'CENTER'), ('ALIGN', (1,0), (1,-1), 'CENTER'), ('ALIGN', (2,0), (-1,-1), 'RIGHT'), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey), ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'), ('SPAN', (0,-1), (1,-1))]
    for i, (_, r) in enumerate(df_f.iterrows()):
        diff = r['Diferença']
        data.append([r['Data'], str(r['Documento']), formatar_moeda_br(r['Valor_Extrato']), formatar_moeda_br(r['Valor_Razao']), formatar_moeda_br(diff) if abs(diff) >= 0.01 else "-"] )
        if abs(diff) >= 0.01:
            table_style.append(('TEXTCOLOR', (4, i+1), (4, i+1), colors.red))
            table_style.append(('FONTNAME', (4, i+1), (4, i+1), 'Helvetica-Bold'))
    data.append(['TOTAL', '', formatar_moeda_br(df_f['Valor_Extrato'].sum()), formatar_moeda_br(df_f['Valor_Razao'].sum()), formatar_moeda_br(df_f['Diferença'].sum())])
    t = Table(data, colWidths=[25*mm, 65*mm, 33*mm, 33*mm, 33*mm])
    t.setStyle(TableStyle(table_style))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

def gerar_excel_final(df_f):
    output = io.BytesIO()
    df_export = df_f.copy()
    df_export['Diferença'] = df_export['Diferença'].round(2)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, sheet_name='Conciliacao', index=False)
        workbook = writer.book; worksheet = writer.sheets['Conciliacao']
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
        fmt_red_bold = workbook.add_format({'font_color': '#FF0000', 'bold': True, 'num_format': '#,##0.00'})
        fmt_total = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'num_format': '#,##0.00', 'border': 1})
        fmt_total_label = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center'})
        for col_num, value in enumerate(df_export.columns.values): worksheet.write(0, col_num, value, fmt_header)
        worksheet.set_column('A:A', 12); worksheet.set_column('B:B', 40); worksheet.set_column('C:C', 15); worksheet.set_column('D:E', 18, fmt_currency)
        worksheet.set_column('F:F', 18, fmt_currency)
        worksheet.conditional_format(1, 5, len(df_export), 5, {'type': 'cell', 'criteria': '!=', 'value': 0, 'format': fmt_red_bold})
        last_row = len(df_export) + 1
        worksheet.merge_range(last_row, 0, last_row, 2, "TOTAL", fmt_total_label)
        worksheet.write(last_row, 3, df_export['Valor_Extrato'].sum(), fmt_total)
        worksheet.write(last_row, 4, df_export['Valor_Razao'].sum(), fmt_total)
        worksheet.write(last_row, 5, df_export['Diferença'].sum(), fmt_total)
    return output.getvalue()

def gerar_extrato_marcado(pdf_bytes, df_f, coords_referencia, nome_original):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    meta = doc.metadata; meta["title"] = f"{nome_original} Marcado"; doc.set_metadata(meta)
    divergencias = df_f[abs(df_f['Diferença']) >= 0.01]
    for _, erro in divergencias.iterrows():
        for item in coords_referencia:
            if item['Data'] == erro['Data'] and abs(item['Valor_Extrato'] - erro['Valor_Extrato']) < 0.01:
                if item['coords']:
                    pno, x0, top, x1, bottom = item['coords']
                    page = doc[pno]; rect = fitz.Rect(x0 - 2, top - 2, x1 + 2, bottom + 2)
                    annot = page.add_highlight_annot(rect); annot.set_colors(stroke=[1, 1, 0]); annot.update()
    return doc.tobytes()

# ==============================================================================
# 3. INTERFACE
# ==============================================================================
st.markdown("<h1 style='text-align: center;'>Conciliador Bancário (Banco x GovBr)</h1>", unsafe_allow_html=True)
st.markdown("---")
c1, c2 = st.columns(2)
with c1: 
    st.markdown('<p class="big-label">Selecione o Extrato Bancário em PDF</p>', unsafe_allow_html=True)
    up_pdf = st.file_uploader("", type="pdf", key="up_pdf", label_visibility="collapsed")
with c2: 
    st.markdown('<p class="big-label">Selecione o Razão da Contabilidade em Excel</p>', unsafe_allow_html=True)
    up_xlsx = st.file_uploader("", type=["xlsx", "csv"], key="up_xlsx", label_visibility="collapsed")

if st.button("PROCESSAR CONCILIAÇÃO", use_container_width=True):
    if up_pdf and up_xlsx:
        with st.spinner("Processando..."):
            pdf_bytes = up_pdf.read()
            xlsx_bytes = up_xlsx.read()
            df_p, coords_ref = processar_pdf(pdf_bytes)
            df_e = processar_excel_detalhado(xlsx_bytes, df_p, is_csv=up_xlsx.name.endswith('csv'))
            if df_p.empty or df_e.empty: st.error("Erro no processamento."); st.stop()
            df_f = executar_conciliacao_inteligente(df_p, df_e)
            
            html = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>"
            html += "<table style='width:100%; border-collapse: collapse; color: black !important; background-color: white !important;'>"
            html += "<tr style='background-color: black; color: white !important;'><th style='padding: 8px; border: 1px solid #000;'>Data</th><th style='padding: 8px; border: 1px solid #000;'>Histórico</th><th style='padding: 8px; border: 1px solid #000;'>Documento</th><th style='padding: 8px; border: 1px solid #000;'>Vlr. Extrato</th><th style='padding: 8px; border: 1px solid #000;'>Vlr. Razão</th><th style='padding: 8px; border: 1px solid #000;'>Diferença</th></tr>"
            for _, r in df_f.iterrows():
                estilo_dif = "color: red; font-weight: bold;" if abs(r['Diferença']) >= 0.01 else "color: black;"
                html += f"<tr style='background-color: white;'><td style='text-align: center; border: 1px solid #000; color: black;'>{r['Data']}</td><td style='text-align: left; border: 1px solid #000; color: black;'>{r['Histórico']}</td><td style='text-align: center; border: 1px solid #000; color: black;'>{r['Documento']}</td><td style='text-align: right; border: 1px solid #000; color: black;'>{formatar_moeda_br(r['Valor_Extrato'])}</td><td style='text-align: right; border: 1px solid #000; color: black;'>{formatar_moeda_br(r['Valor_Razao'])}</td><td style='text-align: right; border: 1px solid #000; {estilo_dif}'>{formatar_moeda_br(r['Diferença']) if abs(r['Diferença']) >= 0.01 else '-'}</td></tr>"
            html += f"<tr style='font-weight: bold; background-color: lightgrey; color: black;'><td colspan='3' style='padding: 10px; text-align: center; border: 1px solid #000;'>TOTAL</td><td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_f['Valor_Extrato'].sum())}</td><td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_f['Valor_Razao'].sum())}</td><td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_f['Diferença'].sum())}</td></tr></table></div>"
            st.markdown(html, unsafe_allow_html=True)
            
            nome_base = os.path.splitext(up_pdf.name)[0]
            st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
            st.download_button("BAIXAR RELATÓRIO DE CONCILIAÇÃO EM PDF", gerar_pdf_final(df_f, f"Conciliação {nome_base}"), f"Conciliacao_{nome_base}.pdf", "application/pdf", use_container_width=True)
            st.download_button("GERAR RELATÓRIO EM EXCEL", gerar_excel_final(df_f), f"Conciliacao_{nome_base}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            st.download_button("BAIXAR EXTRATO BANCÁRIO COM MARCAÇÕES", gerar_extrato_marcado(pdf_bytes, df_f, coords_ref, nome_base), f"{nome_base}_Marcado.pdf", "application/pdf", use_container_width=True)
    else:
        st.warning("⚠️ Selecione os dois arquivos primeiro.")
