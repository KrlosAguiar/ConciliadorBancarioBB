import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import datetime
import xlsxwriter  # Obrigatório estar no requirements.txt
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth # Necessário para calcular largura do texto
from PIL import Image
import fitz  # Requer pymupdf no requirements.txt

# --- CONFIGURAÇÃO DA PÁGINA ---
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
try:
    icon_image = Image.open(icon_path)
    st.set_page_config(page_title="Conciliador Bancário", page_icon=icon_image, layout="wide")
except:
    st.set_page_config(page_title="Conciliador Bancário", layout="wide")

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
    if isinstance(valor, str): return valor
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

    # --- TAGGING ---
    if not df_debitos.empty:
        def refinar_doc(r):
            doc = str(r['Documento'])
            hist = str(r['Histórico']).upper()
            if "FUNDEB" in hist: return f"{doc}-FUNDEB"
            if "PASEP" in hist: return f"{doc}-PASEP"
            if "RETENÇÃO RFB" in hist or "RETENCAO RFB" in hist: return f"{doc}-RFB"
            return doc
        df_debitos['Documento'] = df_debitos.apply(refinar_doc, axis=1)

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
        
        # --- AJUSTE INTELIGENTE: Detecção da Coluna de Valor (Coluna I vs J) ---
        col_valor = 8  # Padrão: Coluna I (índice 8)
        try:
            n_val_8 = df.iloc[:, 8].dropna().astype(str).str.strip().replace('', pd.NA).count()
            n_val_9 = df.iloc[:, 9].dropna().astype(str).str.strip().replace('', pd.NA).count()
            if n_val_8 == 0 and n_val_9 > 0: col_valor = 9
        except: pass
        
        try: df = df.iloc[:, [1, 4, 5, col_valor, 25, 26, 27]].copy()
        except: df = df.iloc[:, [1, 4, 5, col_valor, -4, -2, -1]].copy()
        
        df.columns = ['Lancamento', 'Data', 'DC', 'Valor_Razao', 'Info_Z', 'Info_AA', 'Info_AB']
        df['Valor_Razao'] = df['Valor_Razao'].apply(lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x, str) else float(x))
        df['Lancamento'] = df['Lancamento'].astype(str).str.replace(r'\.0$', '', regex=True)
        
        mask_pagto = df['Info_Z'].astype(str).str.contains("Pagamento", case=False, na=False)
        mask_transf_std = (df['Info_Z'].astype(str).str.contains("TRANSFERENCIA ENTRE CONTAS DE MESMA UG", case=False, na=False)) & (df['DC'].str.strip().str.upper() == 'C')
        mask_codes_z = df['Info_Z'].astype(str).str.contains(r"266|264|268|262", case=False, regex=True, na=False)
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
            candidatos = df_validos[
                (abs(df_validos['Valor_Razao'] - valor_est) < 0.01) & 
                (~df_validos.index.isin(indices_usados_validos))
            ]
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
        lookup_fundeb = {}
        lookup_pasep = {}
        lookup_rfb = {}
        lookup_ded_geral = {}
        
        if not df_pdf_ref.empty:
            for _, row in df_pdf_ref.iterrows():
                dt = row['Data']
                doc = row['Documento']
                hist = str(row['Histórico']).upper()
                if "FUNDEB" in hist: lookup_fundeb[dt] = doc
                if "PASEP" in hist: lookup_pasep[dt] = doc
                if "RETENÇÃO RFB" in hist or "RETENCAO RFB" in hist: lookup_rfb[dt] = doc
                if "DEDUÇÃO" in hist or "DED." in hist:
                    if dt not in lookup_ded_geral: lookup_ded_geral[dt] = doc
        
        def find_doc(row):
            dt = row['Data']
            info_aa = str(row['Info_AA']).upper()
            info_ab = str(row['Info_AB']).upper()
            txt_ab = info_ab
            if "DED.FUNDEB" in info_aa: return lookup_fundeb.get(dt, "NÃO LOCALIZADO")
            if "PASEP" in info_ab: return lookup_pasep.get(dt, "NÃO LOCALIZADO")
            if any(t in info_ab for t in ["PARCELAMENTO SIMPLIFICADO", "PARCELAMENTO SIMPLICADO", "PARCELAMENTO EXCEPCIONAL"]):
                return lookup_rfb.get(dt, "NÃO LOCALIZADO")
            if "DED." in info_aa:
                if dt in lookup_ded_geral: return lookup_ded_geral[dt]
            if dt not in lookup: return "S/D"
            if "TARIFA" in txt_ab and "Tarifas Bancárias" in lookup[dt].values(): return "Tarifas Bancárias"
            for n in re.findall(r'\d+', txt_ab):
                if n.lstrip('0') in lookup[dt]: return lookup[dt][n.lstrip('0')]
            return "NÃO LOCALIZADO"
            
        df_final['Documento'] = df_final.apply(find_doc, axis=1)
        return df_final[['Data', 'Documento', 'Valor_Razao', 'Lancamento']]
        
    except Exception as e:
        return pd.DataFrame()

def executar_conciliacao_inteligente(df_pdf, df_excel):
    res = []
    idx_p_u, idx_e_u = set(), set()
    
    # 1. MATCH EXATO
    for idx_p, row_p in df_pdf.iterrows():
        cand = df_excel[(df_excel['Data'] == row_p['Data']) & (df_excel['Documento'] == row_p['Documento']) & (~df_excel.index.isin(idx_e_u))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            row_e = m.loc[idx_e]
            res.append({
                'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': row_p['Documento'], 
                'Lancamento': row_e['Lancamento'], # Pega do Excel
                'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': row_e['Valor_Razao'], 
                'Diferença': 0.0, 'Tipo': 'Mestre',
                'Sort_Data': row_p['Data'], 'Sort_Doc': row_p['Documento'], 'Order_Idx': 0
            })
            idx_p_u.add(idx_p); idx_e_u.add(idx_e)
            
    # 2. MATCH POR VALOR
    for idx_p, row_p in df_pdf.iterrows():
        if idx_p in idx_p_u: continue
        cand = df_excel[(df_excel['Data'] == row_p['Data']) & (~df_excel.index.isin(idx_e_u))]
        m = cand[abs(cand['Valor_Razao'] - row_p['Valor_Extrato']) < 0.01]
        if not m.empty:
            idx_e = m.index[0]
            row_e = m.loc[idx_e]
            res.append({
                'Data': row_p['Data'], 'Histórico': row_p['Histórico'], 'Documento': "Docs dif.", 
                'Lancamento': row_e['Lancamento'],
                'Valor_Extrato': row_p['Valor_Extrato'], 'Valor_Razao': row_e['Valor_Razao'], 
                'Diferença': 0.0, 'Tipo': 'Mestre',
                'Sort_Data': row_p['Data'], 'Sort_Doc': "Docs dif.", 'Order_Idx': 0
            })
            idx_p_u.add(idx_p); idx_e_u.add(idx_e)
            
    # 3. CONSOLIDAÇÃO DE SOBRAS (COM DETALHAMENTO)
    df_p_sobras = df_pdf[~df_pdf.index.isin(idx_p_u)].copy()
    df_e_sobras = df_excel[~df_excel.index.isin(idx_e_u)].copy()

    # Agrupamento PDF
    if not df_p_sobras.empty:
        df_p_s = df_p_sobras.groupby(['Data', 'Documento']).agg({
            'Valor_Extrato': ['sum', list],
            'Histórico': lambda x: ' | '.join(x.unique())
        }).reset_index()
        df_p_s.columns = ['Data', 'Documento', 'Valor_Extrato_Sum', 'Valor_Extrato_List', 'Histórico']
    else:
        df_p_s = pd.DataFrame(columns=['Data', 'Documento', 'Valor_Extrato_Sum', 'Valor_Extrato_List', 'Histórico'])

    # Agrupamento Excel (Guardando Lançamentos e Valores individuais)
    if not df_e_sobras.empty:
        df_e_s = df_e_sobras.groupby(['Data', 'Documento']).agg({
            'Valor_Razao': ['sum', list],
            'Lancamento': list
        }).reset_index()
        df_e_s.columns = ['Data', 'Documento', 'Valor_Razao_Sum', 'Valor_Razao_List', 'Lancamento_List']
    else:
        df_e_s = pd.DataFrame(columns=['Data', 'Documento', 'Valor_Razao_Sum', 'Valor_Razao_List', 'Lancamento_List'])

    df_m = pd.merge(df_p_s, df_e_s, on=['Data', 'Documento'], how='outer')
    
    for _, row in df_m.iterrows():
        lst_ext = row['Valor_Extrato_List'] if isinstance(row['Valor_Extrato_List'], list) else []
        lst_raz = row['Valor_Razao_List'] if isinstance(row['Valor_Razao_List'], list) else []
        lst_lanc = row['Lancamento_List'] if isinstance(row['Lancamento_List'], list) else []
        
        v_ext_sum = row['Valor_Extrato_Sum'] if pd.notna(row['Valor_Extrato_Sum']) else 0.0
        v_raz_sum = row['Valor_Razao_Sum'] if pd.notna(row['Valor_Razao_Sum']) else 0.0
        
        hist = row['Histórico'] if pd.notna(row['Histórico']) else "S/H"
        
        is_grouped = (len(lst_lanc) > 1)
        lanc_mestre = lst_lanc[0] if len(lst_lanc) == 1 else ("-" if is_grouped else "-")
        
        res.append({
            'Data': row['Data'], 'Histórico': hist, 'Documento': row['Documento'],
            'Lancamento': lanc_mestre,
            'Valor_Extrato': v_ext_sum, 'Valor_Razao': v_raz_sum, 
            'Diferença': v_ext_sum - v_raz_sum,
            'Tipo': 'Mestre',
            'Sort_Data': row['Data'], 'Sort_Doc': row['Documento'], 'Order_Idx': 0
        })
        
        if len(lst_ext) > 1 or len(lst_raz) > 1:
            order_idx = 1
            if len(lst_ext) > 1:
                for val in lst_ext:
                    res.append({
                        'Data': '', 'Histórico': '', 'Documento': '',
                        'Lancamento': '-',
                        'Valor_Extrato': val, 'Valor_Razao': '-',
                        'Diferença': 0.0, 'Tipo': 'Detalhe',
                        'Sort_Data': row['Data'], 'Sort_Doc': row['Documento'], 'Order_Idx': order_idx
                    })
                    order_idx += 1
            
            if len(lst_raz) > 1:
                for val, lanc in zip(lst_raz, lst_lanc):
                    res.append({
                        'Data': '', 'Histórico': '', 'Documento': '',
                        'Lancamento': lanc,
                        'Valor_Extrato': '-', 'Valor_Razao': val,
                        'Diferença': 0.0, 'Tipo': 'Detalhe',
                        'Sort_Data': row['Data'], 'Sort_Doc': row['Documento'], 'Order_Idx': order_idx
                    })
                    order_idx += 1

    df_f = pd.DataFrame(res)
    df_f['dt_sort'] = pd.to_datetime(df_f['Sort_Data'], format='%d/%m/%Y', errors='coerce')
    
    df_sorted = df_f.sort_values(by=['dt_sort', 'Sort_Doc', 'Order_Idx']).drop(columns=['Sort_Data', 'Sort_Doc', 'Order_Idx', 'dt_sort'])
    
    return df_sorted

# ==============================================================================
# 2. GERAÇÃO DE SAÍDAS (PDF, EXCEL E MARCAÇÃO)
# ==============================================================================

def calcular_tamanho_fonte(text, font_name, max_width_pt, start_size=8, min_size=4):
    """Calcula o maior tamanho de fonte que faz o texto caber em uma linha, respeitando um mínimo."""
    size = start_size
    text = str(text)
    while size > min_size:
        if stringWidth(text, font_name, size) <= max_width_pt:
            return size
        size -= 0.5
    return min_size

def gerar_pdf_final(df_f, titulo_completo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=5*mm, leftMargin=5*mm, topMargin=15*mm, bottomMargin=15*mm, title=titulo_completo)
    story = []
    styles = getSampleStyleSheet()
    story.append(Paragraph("Relatório de Conciliação Bancária", styles["Title"]))
    nome_conta_interno = titulo_completo.replace("Conciliação ", "")
    story.append(Paragraph(f"<b>Conta:</b> {nome_conta_interno}", ParagraphStyle(name='C', alignment=1)))
    story.append(Spacer(1, 15))
    
    # NOVAS COLUNAS: Inclusão de Histórico e ajustes de largura
    # A4 Width = 210mm. Margins 5+5=10mm. Usable = 200mm.
    # Data(18) + Lanc(14) + Hist(84) + Doc(18) + Ext(22) + Raz(22) + Dif(22) = 200mm
    headers = ['Data', 'Lanc.', 'Histórico', 'Documento', 'Vlr. Extrato', 'Vlr. Razão', 'Diferença']
    
    col_widths = [18*mm, 14*mm, 84*mm, 18*mm, 22*mm, 22*mm, 22*mm]
    max_hist_width_pt = (84*mm) / 0.352778 # Conversão mm para points (aprox 238 pts)
    
    data = [headers]
    
    # Estilos Base
    table_style = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.black),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'), # Center All Default
        ('ALIGN', (2,0), (2,-1), 'LEFT'),    # Histórico Left
        ('ALIGN', (4,0), (-1,-1), 'RIGHT'),  # Values Right
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
    ]

    row_idx = 1
    for _, r in df_f.iterrows():
        diff = r['Diferença']
        tipo = r['Tipo']
        hist_text = str(r['Histórico'])
        
        # Cálculo dinâmico da fonte para o Histórico
        font_size_hist = calcular_tamanho_fonte(hist_text, 'Helvetica', max_hist_width_pt - 6) # -6 padding
        
        # Criação do Parágrafo para o Histórico com o tamanho calculado
        # style_hist = ParagraphStyle(name=f'H_{row_idx}', fontName='Helvetica', fontSize=font_size_hist, leading=font_size_hist+2)
        p_hist = Paragraph(f'<font size={font_size_hist}>{hist_text}</font>', styles['Normal'])

        linha = [
            r['Data'], 
            str(r['Lancamento']), 
            p_hist, # Objeto Paragraph
            str(r['Documento']), 
            formatar_moeda_br(r['Valor_Extrato']), 
            formatar_moeda_br(r['Valor_Razao']), 
            formatar_moeda_br(diff) if abs(diff) >= 0.01 else "-"
        ]
        data.append(linha)
        
        if tipo == 'Detalhe':
            # Fundo Cinza Claro, Fonte MENOR, Fonte PRETA
            table_style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.whitesmoke))
            table_style.append(('FONTSIZE', (0, row_idx), (-1, row_idx), 7))
            table_style.append(('TEXTCOLOR', (0, row_idx), (-1, row_idx), colors.black))
            table_style.append(('TOPPADDING', (0, row_idx), (-1, row_idx), 0))
            table_style.append(('BOTTOMPADDING', (0, row_idx), (-1, row_idx), 0))
        else:
            # Mestre
            table_style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.lightgrey))
            table_style.append(('FONTNAME', (0, row_idx), (-1, row_idx), 'Helvetica-Bold'))
            table_style.append(('FONTSIZE', (0, row_idx), (-1, row_idx), 8))
            
            if abs(diff) >= 0.01:
                table_style.append(('TEXTCOLOR', (6, row_idx), (6, row_idx), colors.red))
        
        row_idx += 1

    data.append(['TOTAL', '', '', '',
                 formatar_moeda_br(df_f[df_f['Tipo']=='Mestre']['Valor_Extrato'].sum()), 
                 formatar_moeda_br(df_f[df_f['Tipo']=='Mestre']['Valor_Razao'].sum()), 
                 formatar_moeda_br(df_f[df_f['Tipo']=='Mestre']['Diferença'].sum())])
    
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle(table_style))
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

def gerar_excel_final(df_f):
    output = io.BytesIO()
    
    df_export = df_f.copy()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, sheet_name='Conciliacao', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Conciliacao']
        
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
        fmt_red_bold = workbook.add_format({'font_color': '#FF0000', 'bold': True, 'num_format': '#,##0.00'})
        
        # Estilo para Detalhes: Cor Preta e Borda 1
        fmt_detalhe = workbook.add_format({'font_color': '#000000', 'bg_color': '#F2F2F2', 'italic': True, 'num_format': '#,##0.00', 'font_size': 9, 'border': 1})
        fmt_detalhe_txt = workbook.add_format({'font_color': '#000000', 'bg_color': '#F2F2F2', 'italic': True, 'font_size': 9, 'border': 1})
        
        fmt_total = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'num_format': '#,##0.00', 'border': 1})
        fmt_total_label = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center'})

        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, fmt_header)

        worksheet.set_column('A:A', 12) # Data
        worksheet.set_column('B:B', 30) # Historico
        worksheet.set_column('C:C', 15) # Documento
        worksheet.set_column('D:D', 12) # Lancamento
        worksheet.set_column('E:F', 18, fmt_currency) # Valores
        worksheet.set_column('G:G', 18, fmt_currency) # Diferenca
        
        for i, row in df_export.iterrows():
            excel_row = i + 1
            if row['Tipo'] == 'Detalhe':
                worksheet.set_row(excel_row, 10)
                worksheet.write(excel_row, 0, row['Data'], fmt_detalhe_txt)
                worksheet.write(excel_row, 1, row['Histórico'], fmt_detalhe_txt)
                worksheet.write(excel_row, 2, row['Documento'], fmt_detalhe_txt)
                worksheet.write(excel_row, 3, row['Lancamento'], fmt_detalhe_txt)
                
                val_ext = row['Valor_Extrato']
                val_raz = row['Valor_Razao']
                
                if val_ext == '-': worksheet.write(excel_row, 4, '-', fmt_detalhe)
                else: worksheet.write(excel_row, 4, val_ext, fmt_detalhe)
                
                if val_raz == '-': worksheet.write(excel_row, 5, '-', fmt_detalhe)
                else: worksheet.write(excel_row, 5, val_raz, fmt_detalhe)
                
                worksheet.write(excel_row, 6, '', fmt_detalhe)
            else:
                if abs(row['Diferença']) >= 0.01:
                     worksheet.write(excel_row, 6, row['Diferença'], fmt_red_bold)

        df_mestre = df_export[df_export['Tipo'] == 'Mestre']
        last_row = len(df_export) + 1
        worksheet.merge_range(last_row, 0, last_row, 3, "TOTAL", fmt_total_label)
        worksheet.write(last_row, 4, df_mestre['Valor_Extrato'].sum(), fmt_total)
        worksheet.write(last_row, 5, df_mestre['Valor_Razao'].sum(), fmt_total)
        worksheet.write(last_row, 6, df_mestre['Diferença'].sum(), fmt_total)

    return output.getvalue()

def gerar_extrato_marcado(pdf_bytes, df_f, coords_referencia, nome_original):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    meta = doc.metadata
    meta["title"] = f"{nome_original} Marcado"
    doc.set_metadata(meta)
    
    divergencias = df_f[(df_f['Tipo'] == 'Mestre') & (abs(df_f['Diferença']) >= 0.01)]
    
    for _, erro in divergencias.iterrows():
        for item in coords_referencia:
            if item['Data'] == erro['Data'] and abs(item['Valor_Extrato'] - erro['Valor_Extrato']) < 0.01:
                if item['coords']:
                    pno, x0, top, x1, bottom = item['coords']
                    page = doc[pno]
                    rect = fitz.Rect(x0 - 2, top - 2, x1 + 2, bottom + 2)
                    annot = page.add_highlight_annot(rect)
                    annot.set_colors(stroke=[1, 1, 0])
                    annot.update()
    
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
            
            # --- TABELA HTML (VISUAL) ---
            html = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>"
            html += "<table style='width:100%; border-collapse: collapse; color: black !important; background-color: white !important; font-family: Arial, sans-serif;'>"
            html += "<tr style='background-color: black; color: white !important;'>"
            html += "<th style='padding: 8px; border: 1px solid #000;'>Data</th>"
            html += "<th style='padding: 8px; border: 1px solid #000;'>Lançamento</th>"
            html += "<th style='padding: 8px; border: 1px solid #000;'>Histórico</th>"
            html += "<th style='padding: 8px; border: 1px solid #000;'>Documento</th>"
            html += "<th style='padding: 8px; border: 1px solid #000;'>Vlr. Extrato</th>"
            html += "<th style='padding: 8px; border: 1px solid #000;'>Vlr. Razão</th>"
            html += "<th style='padding: 8px; border: 1px solid #000;'>Diferença</th></tr>"
            
            for _, r in df_f.iterrows():
                if r['Tipo'] == 'Detalhe':
                    # Linha de Detalhe: Cor preta (#000) e Borda Preta Sólida (#000)
                    style_row = "background-color: #f2f2f2; color: #000; font-size: 11px; line-height: 1.0;"
                    style_cell = "padding: 2px 8px; border: 1px solid #000;"
                    html += f"<tr style='{style_row}'>"
                    html += f"<td style='{style_cell} text-align: center;'></td>" # Data Vazia
                    html += f"<td style='{style_cell} text-align: center;'>{r['Lancamento']}</td>"
                    html += f"<td style='{style_cell}'></td>" # Historico Vazio
                    html += f"<td style='{style_cell}'></td>" # Documento Vazio
                    html += f"<td style='{style_cell} text-align: right;'>{formatar_moeda_br(r['Valor_Extrato'])}</td>"
                    html += f"<td style='{style_cell} text-align: right;'>{formatar_moeda_br(r['Valor_Razao'])}</td>"
                    html += f"<td style='{style_cell}'></td></tr>"
                else:
                    estilo_dif = "color: red; font-weight: bold;" if abs(r['Diferença']) >= 0.01 else "color: black;"
                    html += f"<tr style='background-color: white;'>"
                    html += f"<td style='text-align: center; border: 1px solid #000; color: black;'>{r['Data']}</td>"
                    html += f"<td style='text-align: center; border: 1px solid #000; color: black;'>{r['Lancamento']}</td>"
                    html += f"<td style='text-align: left; border: 1px solid #000; color: black;'>{r['Histórico']}</td>"
                    html += f"<td style='text-align: center; border: 1px solid #000; color: black;'>{r['Documento']}</td>"
                    html += f"<td style='text-align: right; border: 1px solid #000; color: black;'>{formatar_moeda_br(r['Valor_Extrato'])}</td>"
                    html += f"<td style='text-align: right; border: 1px solid #000; color: black;'>{formatar_moeda_br(r['Valor_Razao'])}</td>"
                    html += f"<td style='text-align: right; border: 1px solid #000; {estilo_dif}'>{formatar_moeda_br(r['Diferença']) if abs(r['Diferença']) >= 0.01 else '-'}</td></tr>"
            
            df_mestre = df_f[df_f['Tipo'] == 'Mestre']
            html += f"<tr style='font-weight: bold; background-color: lightgrey; color: black;'><td colspan='4' style='padding: 10px; text-align: center; border: 1px solid #000;'>TOTAL</td>"
            html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_mestre['Valor_Extrato'].sum())}</td>"
            html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_mestre['Valor_Razao'].sum())}</td>"
            html += f"<td style='text-align: right; border: 1px solid #000;'>{formatar_moeda_br(df_mestre['Diferença'].sum())}</td></tr></table></div>"
            st.markdown(html, unsafe_allow_html=True)
            
            nome_base = os.path.splitext(up_pdf.name)[0]
            
            pdf_bytes_final = gerar_pdf_final(df_f, f"Conciliação {nome_base}")
            excel_bytes_final = gerar_excel_final(df_f)
            pdf_marcado_final = gerar_extrato_marcado(pdf_bytes, df_f, coords_ref, nome_base)
            
            st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
            
            st.download_button(
                label="BAIXAR RELATÓRIO DE CONCILIAÇÃO EM PDF",
                data=pdf_bytes_final,
                file_name=f"Conciliacao_{nome_base}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            
            st.download_button(
                label="GERAR RELATÓRIO EM EXCEL",
                data=excel_bytes_final,
                file_name=f"Conciliacao_{nome_base}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.download_button(
                label="BAIXAR EXTRATO BANCÁRIO COM MARCAÇÕES",
                data=pdf_marcado_final,
                file_name=f"{nome_base}_Marcado.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    else:
        st.warning("⚠️ Selecione os dois arquivos primeiro.")
