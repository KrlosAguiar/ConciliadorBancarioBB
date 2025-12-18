import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import datetime
import unicodedata
from ftfy import fix_text

# ReportLab Imports para PDF
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm

# ==============================================================================
# CONFIGURAÇÃO INICIAL
# ==============================================================================
st.set_page_config(page_title="Extração de Tarifas Bancárias", layout="wide")

# ==============================================================================
# 1. DESIGN GERAL DA PÁGINA
# ==============================================================================

def aplicar_estilo_global():
    """Aplica o CSS personalizado."""
    st.markdown("""
    <style>
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }
        /* Botões com largura total e cor personalizada */
        div.stButton > button {
            background-color: rgb(38, 39, 48) !important;
            color: white !important;
            font-weight: bold !important;
            border: 1px solid rgb(60, 60, 60);
            border-radius: 5px;
            font-size: 16px;
            transition: 0.3s;
            width: 100%; /* Garante 100% da largura do container pai */
        }
        div.stButton > button:hover {
            background-color: rgb(20, 20, 25) !important;
            border-color: white;
        }
        .big-label {
            font-size: 24px !important;
            font-weight: 600 !important;
            margin-bottom: 10px;
        }
    </style>
    """, unsafe_allow_html=True)

def renderizar_cabecalho(titulo):
    st.markdown(f"<h1 style='text-align: center;'>{titulo}</h1>", unsafe_allow_html=True)
    st.markdown("---")

def renderizar_label_uploader(texto):
    st.markdown(f'<p class="big-label">{texto}</p>', unsafe_allow_html=True)

def renderizar_espacador_botao():
    st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)

# ==============================================================================
# 2. DESIGN DA TABELA E PDF
# ==============================================================================

class DesignTabelaHTML:
    """Templates HTML para exibição na tela."""
    CONTAINER_OPEN = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>"
    CONTAINER_CLOSE = "</div>"
    TABLE_OPEN = "<table style='width:100%; border-collapse: collapse; color: black !important; background-color: white !important; font-family: sans-serif;'>"
    TABLE_CLOSE = "</table>"
    
    # Cabeçalho
    HEADER_HTML = """
    <tr style='background-color: #00008B; color: white !important;'>
        <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Data</th>
        <th style='padding: 8px; text-align: left; border: 1px solid #000;'>Histórico</th>
        <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Documento</th>
        <th style='padding: 8px; text-align: right; border: 1px solid #000;'>Valor</th>
    </tr>
    """

class DesignRelatorioPDF:
    """Configurações visuais para a geração do PDF."""
    
    PAGE_SIZE = A4
    MARGIN_RIGHT = 10 * mm
    MARGIN_LEFT = 10 * mm
    MARGIN_TOP = 15 * mm
    MARGIN_BOTTOM = 15 * mm
    
    @staticmethod
    def get_table_style(has_data=True):
        style_cmds = [
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.darkblue), # Cabeçalho Azul
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),     # Texto Cabeçalho Branco
            ('ALIGN', (0,0), (0,-1), 'CENTER'),     # Data
            ('ALIGN', (1,0), (1,-1), 'LEFT'),       # Histórico
            ('ALIGN', (2,0), (2,-1), 'CENTER'),     # Documento
            ('ALIGN', (3,0), (-1,-1), 'RIGHT'),     # Valor
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('BOTTOMPADDING', (0,0), (-1,0), 8),
            ('TEXTCOLOR', (0,1), (-1,-1), colors.black), # Dados pretos
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), # Alinhamento vertical
        ]
        return TableStyle(style_cmds)

# ==============================================================================
# 3. UTILITÁRIOS DE PROCESSAMENTO
# ==============================================================================

def clean_text_general(s):
    if pd.isna(s): return s
    if not isinstance(s, str): s = str(s)
    try: s = fix_text(s)
    except: pass
    s = unicodedata.normalize('NFC', s)
    return re.sub(r'\s+', ' ', s).strip()

def format_currency_br(val):
    if pd.isna(val): return "-"
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==============================================================================
# 4. MOTORES DE EXTRAÇÃO
# ==============================================================================

def processar_bb(file_bytes):
    TARGET_LOTE_BB = "13113"
    VALOR_RE = re.compile(r'^\(?-?\d{1,3}(?:\.\d{3})*,\d{2}\)?$')

    def br_to_float(s):
        if pd.isna(s): return None
        s = str(s).strip()
        if not s: return None
        is_neg = s.startswith("(") and s.endswith(")")
        if is_neg: s = s[1:-1].strip()
        s = re.sub(r'[^\d\.,\-]', '', s)
        if s.startswith('-'): is_neg = True; s = s[1:]
        s = s.replace('.', '').replace(',', '.')
        try:
            val = float(s)
            return -val if is_neg else val
        except: return None

    all_rows = []
    header = None
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if not tables: continue
                for t in tables:
                    rows = [[("" if c is None else str(c).strip()) for c in row] for row in t]
                    if header is None:
                        header = rows[0]
                        all_rows.extend(rows[1:])
                    else:
                        if rows[0] == header: all_rows.extend(rows[1:])
                        else: all_rows.extend(rows)
    except: return pd.DataFrame()

    if not header: return pd.DataFrame()
    df = pd.DataFrame(all_rows, columns=header)
    
    for c in df.select_dtypes(include=['object']).columns:
        df[c] = df[c].apply(clean_text_general)
    
    for col in df.columns:
        norm = col.lower()
        if "lote" in norm: df = df.rename(columns={col:"lote"})
        if "data" in norm: df = df.rename(columns={col:"dt_balancete"})
        if "valor" in norm: df = df.rename(columns={col:"valor"})
        if "hist" in norm: df = df.rename(columns={col:"historico"})
    
    df = df.loc[:, ~df.columns.duplicated()]

    if "dt_balancete" not in df.columns:
        for col in df.columns:
            if any(re.match(r'\d{2}/\d{2}/\d{4}', str(v)) for v in df[col].head(10)):
                df = df.rename(columns={col:"dt_balancete"}); break
    if "valor" not in df.columns:
        for col in df.columns:
            if any(VALOR_RE.match(str(v).strip()) for v in df[col].head(10)):
                df = df.rename(columns={col:"valor"}); break
    
    df = df.loc[:, ~df.columns.duplicated()]
    if "lote" not in df.columns: df["lote"] = None

    df["lote_norm"] = df["lote"].apply(lambda v: re.sub(r'\D+','', str(v)) if v else "")
    df["valor_num"] = df.get("valor").apply(br_to_float)
    df["dt_obj"] = pd.to_datetime(df.get("dt_balancete"), dayfirst=True, errors="coerce")

    df_final = df[df["lote_norm"] == TARGET_LOTE_BB].copy()
    df_final = df_final.rename(columns={"historico": "Histórico", "lote": "Documento", "valor_num": "Valor"})
    return df_final[["dt_obj", "Histórico", "Documento", "Valor"]]

def processar_caixa(file_bytes):
    FILTER_TERMS = ["DEB TARIFA", "DEB ARREC", "TAR ARREC", "DOC/TED PE", "TAR CC ATV"]
    DATE_RE = re.compile(r'^\s*(\d{2}/\d{2}/\d{4})\b')
    VALOR_RE = re.compile(r'\(?-?\d{1,3}(?:[.\u00A0]\d{3})*,\d{2}\)?')
    NOISE_RE = re.compile(r'_{3,}|(^|\s)(CAIXA|SAC|OUVIDORIA|AL[ÔO] CAIXA|GERENCIADOR)(\s|$)', re.I)

    def br_to_float(s):
        if not s: return None
        t = str(s).strip()
        neg = False
        if t.startswith('(') and t.endswith(')'): neg = True; t = t[1:-1].strip()
        t = re.sub(r'[^\d,\.\u00A0-]', '', t)
        t = t.replace('\u00A0','').replace('.', '').replace(',', '.')
        try: v = float(t); return -v if neg else v
        except: return None

    rows = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for ln in text.splitlines():
                    s = clean_text_general(ln)
                    if not s or NOISE_RE.search(s): continue
                    m = DATE_RE.match(s)
                    if m:
                        date_str = m.group(1)
                        rest = s[m.end():].strip()
                        val_matches = list(VALOR_RE.finditer(rest))
                        valor_text = None
                        if len(val_matches) == 1:
                            valor_text = val_matches[0].group(0)
                            before_val = rest[:val_matches[0].start()].strip()
                        elif len(val_matches) >= 2:
                            valor_text = val_matches[-2].group(0)
                            before_val = rest[:val_matches[-2].start()].strip()
                        else:
                            before_val = rest
                        nr_doc = None
                        for tok in rest.split():
                            if re.fullmatch(r'\d+', tok): nr_doc = tok; break
                        if nr_doc:
                            idx = rest.find(nr_doc)
                            hist_zone = rest[idx+len(nr_doc):].strip()
                            if valor_text:
                                vpos = hist_zone.rfind(valor_text)
                                hist = hist_zone[:vpos].strip() if vpos != -1 else hist_zone
                            else: hist = hist_zone
                        else: hist = before_val
                        rows.append({"dt_str": date_str, "Documento": nr_doc if nr_doc else "", "Histórico": clean_text_general(hist), "Valor": br_to_float(valor_text)})
                    else:
                        if rows:
                            extra = clean_text_general(ln)
                            if len(extra) > 3 and not re.match(r'^\d', extra):
                                rows[-1]["Histórico"] += " " + extra
    except: return pd.DataFrame()

    df = pd.DataFrame(rows)
    if df.empty: return pd.DataFrame()

    def is_target(s):
        ss = str(s).upper()
        return any(t in ss for t in FILTER_TERMS)

    df_filtered = df[df["Histórico"].apply(is_target)].copy()
    df_filtered["dt_obj"] = pd.to_datetime(df_filtered["dt_str"], dayfirst=True, errors="coerce")
    return df_filtered[["dt_obj", "Histórico", "Documento", "Valor"]]

def processar_banpara(file_bytes):
    TARGET_TERMS = ["TAR ELET TRIB ARREC", "TAF ARREC TRIB", "MANUT CONTA ATIVA PJ", "TED PESSOAL", "TRANSF.RECURSO(P)", "PCT SERV MAXIEMP"]
    DATE_FULL = re.compile(r'^\s*\d{2}/\d{2}/\d{4}\b')
    DATE_DM = re.compile(r'^\s*(\d{2}/\d{2})\b')
    VALOR_RE = re.compile(r'\(?-?\d{1,3}(?:[.\u00A0]\d{3})*,\d{2}\)?')

    def br_to_float(s):
        if not s: return None
        t = str(s).strip()
        if t.startswith('(') and t.endswith(')'): t = t[1:-1].strip()
        t = re.sub(r'[^\d,\.\u00A0-]', '', t)
        t = t.replace('\u00A0','').replace('.', '').replace(',', '.')
        if t in ("", "-", "."): return None
        try: return abs(float(t))
        except: return None

    transactions = []
    last_tx = None
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            cur_year = str(datetime.datetime.now().year)
            for page in pdf.pages:
                text = page.extract_text() or ""
                for ln in text.splitlines():
                    s = clean_text_general(ln)
                    if not s: continue
                    if DATE_FULL.match(s): continue
                    m = DATE_DM.match(s)
                    if m:
                        dt = f"{m.group(1)}/{cur_year}"
                        rest = s[m.end():].strip()
                        vals = list(VALOR_RE.finditer(rest))
                        v_txt = None
                        if len(vals) >= 1:
                            idx = -2 if len(vals) >= 2 else 0
                            v_txt = vals[idx].group(0)
                            search = rest[:vals[idx].start()]
                        else: search = rest
                        doc_m = re.search(r'\b(\d{1,6})\b', search)
                        doc = doc_m.group(1) if doc_m else ""
                        desc = search[:doc_m.start()].strip() if doc_m else search
                        tx = {"dt_str": dt, "Histórico": clean_text_general(desc), "Documento": doc, "Valor": br_to_float(v_txt)}
                        transactions.append(tx)
                        last_tx = tx
                    else:
                        if last_tx:
                            ex = clean_text_general(s)
                            if "SALDO" not in ex.upper() and "PÁGINA" not in ex.upper():
                                last_tx["Histórico"] += " " + ex
    except: return pd.DataFrame()

    df = pd.DataFrame(transactions)
    if df.empty: return pd.DataFrame()

    def is_target(s):
        ss = str(s).upper()
        return any(t in ss for t in TARGET_TERMS)

    df_f = df[df["Histórico"].apply(is_target)].copy()
    if "Valor" in df_f.columns: df_f["Valor"] = df_f["Valor"].abs()
    df_f["dt_obj"] = pd.to_datetime(df_f["dt_str"], dayfirst=True, errors="coerce")
    return df_f[["dt_obj", "Histórico", "Documento", "Valor"]]

# ==============================================================================
# 5. LÓGICA DE GERAÇÃO DE RELATÓRIO (HTML E PDF)
# ==============================================================================

def preparar_dados_relatorio(df):
    if df.empty: return [], 0
    df = df.sort_values(by="dt_obj")
    report_data = []
    grand_total = 0
    
    for date, group in df.groupby("dt_obj"):
        daily_total = 0
        date_str = date.strftime("%d/%m/%Y")
        
        for _, row in group.iterrows():
            val = row["Valor"] if pd.notnull(row["Valor"]) else 0
            daily_total += val
            grand_total += val
            
            raw_doc = row["Documento"]
            if isinstance(raw_doc, pd.Series): raw_doc = raw_doc.iloc[0]
            doc_val = str(raw_doc) if pd.notna(raw_doc) and str(raw_doc).strip() != "" else ""

            report_data.append({
                "Data": date_str,
                "Histórico": row["Histórico"],
                "Documento": doc_val,
                "Valor": val,
                "IsTotal": False,
                "IsGrandTotal": False
            })
            
        report_data.append({
            "Data": date_str,
            "Histórico": "Total do Dia",
            "Documento": "-",
            "Valor": daily_total,
            "IsTotal": True,
            "IsGrandTotal": False
        })
    
    report_data.append({
        "Data": "",
        "Histórico": "TOTAL GERAL",
        "Documento": "-",
        "Valor": grand_total,
        "IsTotal": True,
        "IsGrandTotal": True
    })
    
    return report_data, grand_total

def gerar_html_tabela(report_data):
    html = DesignTabelaHTML.CONTAINER_OPEN + DesignTabelaHTML.TABLE_OPEN
    html += DesignTabelaHTML.HEADER_HTML
    
    for row in report_data:
        v_fmt = format_currency_br(row['Valor'])
        style = "background-color: white; color: black;"
        doc_align = "center"
        hist_style = "text-align: left;"
        
        if row['IsTotal']:
            style = "background-color: #f0f0f0; font-weight: bold; border-top: 1px solid #ccc; color: black;"
            if row['IsGrandTotal']:
                style = "background-color: #d1d9e6; color: black; font-weight: bold; border-top: 2px solid #000;"
        
        html += f"<tr style='{style}'>"
        html += f"<td style='padding: 8px; text-align: center; border: 1px solid #ddd;'>{row['Data']}</td>"
        html += f"<td style='padding: 8px; {hist_style} border: 1px solid #ddd;'>{row['Histórico']}</td>"
        html += f"<td style='padding: 8px; text-align: {doc_align}; border: 1px solid #ddd;'>{row['Documento']}</td>"
        html += f"<td style='padding: 8px; text-align: right; border: 1px solid #ddd;'>{v_fmt}</td>"
        html += "</tr>"

    html += DesignTabelaHTML.TABLE_CLOSE + DesignTabelaHTML.CONTAINER_CLOSE
    return html

def gerar_pdf_bytes(report_data, titulo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=A4, 
        rightMargin=DesignRelatorioPDF.MARGIN_RIGHT,
        leftMargin=DesignRelatorioPDF.MARGIN_LEFT,
        topMargin=DesignRelatorioPDF.MARGIN_TOP,
        bottomMargin=DesignRelatorioPDF.MARGIN_BOTTOM,
        title=titulo
    )
    
    elements = []
    styles = getSampleStyleSheet()
    
    elements.append(Paragraph(titulo, styles['Title']))
    elements.append(Spacer(1, 10*mm))
    
    table_data = [['Data', 'Histórico', 'Documento', 'Valor']]
    
    for row in report_data:
        v_fmt = format_currency_br(row['Valor'])
        hist_txt = str(row['Histórico']).replace("<", "&lt;").replace(">", "&gt;")
        table_data.append([row['Data'], Paragraph(hist_txt, styles['Normal']), row['Documento'], v_fmt])

    t = Table(table_data, colWidths=[25*mm, 90*mm, 35*mm, 35*mm], repeatRows=1)
    
    ts = DesignRelatorioPDF.get_table_style()
    
    for idx, row in enumerate(report_data):
        row_idx = idx + 1
        if row['IsTotal']:
            ts.add('BACKGROUND', (0, row_idx), (-1, row_idx), colors.lightgrey)
            ts.add('FONTNAME', (0, row_idx), (-1, row_idx), 'Helvetica-Bold')
            ts.add('LINEABOVE', (0, row_idx), (-1, row_idx), 1, colors.black)
            ts.add('TEXTCOLOR', (0, row_idx), (-1, row_idx), colors.black)
            
            if row['IsGrandTotal']:
                ts.add('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor('#d1d9e6'))
                # AUMENTAR FONTE DO VALOR PARA O DOBRO (9 * 2 = 18)
                # A coluna Valor é a índice 3
                ts.add('FONTSIZE', (3, row_idx), (3, row_idx), 18)
                # Opcional: Ajustar padding se necessário
                ts.add('TOPPADDING', (3, row_idx), (3, row_idx), 6)
                ts.add('BOTTOMPADDING', (3, row_idx), (3, row_idx), 6)

    t.setStyle(ts)
    elements.append(t)
    
    doc.build(elements)
    return buffer.getvalue()

# ==============================================================================
# 6. APP PRINCIPAL
# ==============================================================================

aplicar_estilo_global()
renderizar_cabecalho("Extração de Tarifas Bancárias")

col1, col2 = st.columns(2)
with col1:
    renderizar_label_uploader("Selecione o Banco")
    banco_option = st.selectbox("", ["Banco do Brasil", "Caixa Econômica", "BANPARÁ"], label_visibility="collapsed")

with col2:
    renderizar_label_uploader("Upload do Extrato (PDF)")
    uploaded_file = st.file_uploader("", type="pdf", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)
# Botão Iniciar com largura total
if st.button("INICIAR PROCESSAMENTO", use_container_width=True):
    if uploaded_file:
        with st.spinner("Processando dados..."):
            file_bytes = uploaded_file.read()
            df_result = pd.DataFrame()
            
            if banco_option == "Banco do Brasil":
                df_result = processar_bb(file_bytes)
            elif banco_option == "Caixa Econômica":
                df_result = processar_caixa(file_bytes)
            elif banco_option == "BANPARÁ":
                df_result = processar_banpara(file_bytes)
            
            if not df_result.empty:
                report_data, _ = preparar_dados_relatorio(df_result)
                
                # HTML na tela
                html_table = gerar_html_tabela(report_data)
                st.markdown(html_table, unsafe_allow_html=True)
                
                # Nome do arquivo
                nome_arquivo_original = os.path.splitext(uploaded_file.name)[0]
                nome_pdf_final = f"Tarifas Bancárias {nome_arquivo_original}"
                
                # Gera PDF
                pdf_bytes = gerar_pdf_bytes(report_data, nome_pdf_final)
                
                renderizar_espacador_botao()
                # Botão Download com largura total
                st.download_button(
                    label="BAIXAR RELATÓRIO PDF",
                    data=pdf_bytes,
                    file_name=f"{nome_pdf_final}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.warning(f"Nenhum dado encontrado para **{banco_option}** com os filtros configurados.")
    else:
        st.warning("⚠️ Por favor, faça o upload do arquivo PDF antes de iniciar.")
