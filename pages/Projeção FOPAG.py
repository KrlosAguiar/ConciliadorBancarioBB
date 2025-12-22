import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import xml.etree.ElementTree as ET
import re
import io
import os
import datetime
from PIL import Image

# ReportLab Imports
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
# Define o caminho para a imagem que está na raiz do projeto
# Isso garante que funcione tanto localmente quanto no servidor
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
icon_image = Image.open(icon_path)

st.set_page_config(
    page_title="Projeção de Folha",
    page_icon=icon_image,
    layout="wide"
)

# ==============================================================================
# 1. DESIGN E ESTILIZAÇÃO (CSS)
# ==============================================================================
def aplicar_estilo_global():
    st.markdown("""
    <style>
        .block-container { padding-top: 2rem !important; }
        
        /* Ajuste para que todos os botões Streamlit ocupem 100% da largura */
        div.stButton > button, div.stDownloadButton > button {
            background-color: rgb(38, 39, 48) !important;
            color: white !important;
            font-weight: bold !important;
            border-radius: 5px;
            width: 100% !important; 
            height: 50px;
            transition: 0.3s;
            border: 1px solid rgb(60, 60, 60);
        }
        
        div.stButton > button:hover, div.stDownloadButton > button:hover { 
            background-color: rgb(20, 20, 25) !important; 
            border-color: white; 
        }

        .big-label { font-size: 20px !important; font-weight: 600 !important; margin-bottom: 10px; }
        .footer-info { text-align: center; color: #666; font-size: 12px; margin-top: 50px; }
    </style>
    """, unsafe_allow_html=True)

def renderizar_cabecalho(titulo):
    st.markdown(f"<h1 style='text-align: center;'>{titulo}</h1>", unsafe_allow_html=True)
    st.markdown("---")

# ==============================================================================
# 2. FUNÇÕES DE APOIO (LÓGICA PRESERVADA)
# ==============================================================================

def read_ods_streamlit(file_bytes):
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        with z.open('content.xml') as f:
            root = ET.parse(f).getroot()
    ns = {'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
          'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
          'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'}
    data = []
    for sheet in root.findall('.//table:table', ns)[:1]:
        for row in sheet.findall('.//table:table-row', ns):
            row_data = []
            for cell in row.findall('.//table:table-cell', ns):
                repeat = cell.get('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-repeated')
                val = cell.get('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value')
                if val is None:
                    text_nodes = cell.findall('.//text:p', ns)
                    val = " ".join([ET.tostring(t, method='text', encoding='unicode') for t in text_nodes]) if text_nodes else ""
                for _ in range(int(repeat) if repeat else 1):
                    row_data.append(val.strip())
            data.append(row_data)
    max_len = max(len(r) for r in data) if data else 0
    return pd.DataFrame([r + [""] * (max_len - len(r)) for r in data])

def format_currency_br(val):
    try:
        v = float(val)
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

def to_num(val):
    if val is None or str(val).strip() in ["", "-", "None"]: return 0.0
    s = str(val).strip()
    try:
        if ',' in s: s = s.replace('.', '').replace(',', '.')
        return float(s)
    except: return 0.0

def gerar_pdf_folha(df, decorridos, restantes, titulo_pdf):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=15, leftMargin=15, topMargin=20, bottomMargin=20)
    doc.title = titulo_pdf
    elements = []
    styles = getSampleStyleSheet()

    header_style = ParagraphStyle('HBold', parent=styles['Normal'], alignment=1, fontSize=12, leading=14, fontName='Helvetica-Bold')
    col_header_style = ParagraphStyle('ColH', parent=styles['Normal'], alignment=1, fontSize=8, fontName='Helvetica-Bold')
    small_text_style = ParagraphStyle('Small', parent=styles['Normal'], fontSize=7, leading=8)
    val_style = ParagraphStyle('Value', parent=styles['Normal'], alignment=1, fontSize=9, leading=10)

    elements.append(Paragraph("<b>PROJEÇÃO DE FOLHA DE PAGAMENTO</b>", styles['Title']))
    elements.append(Paragraph(f"<b>FOLHAS DECORRIDAS: {decorridos} | FOLHAS RESTANTES: {restantes}</b>", header_style))
    elements.append(Spacer(1, 10*mm))

    data_table = [[
        Paragraph("Órgão/Unidade", col_header_style), Paragraph("Cod", col_header_style), 
        Paragraph("Despesa", col_header_style), Paragraph("Liquidado", col_header_style), 
        Paragraph("Saldo", col_header_style), Paragraph("Média", col_header_style), 
        Paragraph("Projeção", col_header_style), Paragraph("Suplementar", col_header_style)
    ]]

    for _, row in df.iterrows():
        suple_val = row['Suplementar']
        color = "red" if suple_val < 0 else "black"
        suple_style = ParagraphStyle('Suple', parent=val_style, textColor=color, fontName='Helvetica-Bold' if suple_val < 0 else 'Helvetica')

        data_table.append([
            Paragraph(str(row['Órgão']), small_text_style), 
            Paragraph(str(row['Código']), val_style), 
            Paragraph(str(row['Despesa']), small_text_style),
            Paragraph(format_currency_br(row['Liquidado']), val_style), 
            Paragraph(format_currency_br(row['Saldo']), val_style),
            Paragraph(format_currency_br(row['Média']), val_style), 
            Paragraph(format_currency_br(row['Projeção']), val_style), 
            Paragraph(format_currency_br(row['Suplementar']), suple_style)
        ])

    widths = [2.2*inch, 0.4*inch, 2.2*inch, 1.22*inch, 1.22*inch, 1.22*inch, 1.22*inch, 1.22*inch]
    t = Table(data_table, repeatRows=1, colWidths=widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f2f2f2")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))

    elements.append(t)
    doc.build(elements)
    return buffer.getvalue()

# ==============================================================================
# 3. EXECUÇÃO DA INTERFACE
# ==============================================================================

aplicar_estilo_global()
renderizar_cabecalho("Projeção de Folha de Pagamento")

st.markdown('<p class="big-label">Selecione o arquivo no formato .ods</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type=["ods"], label_visibility="collapsed")

if uploaded_file:
    # ADICIONADO use_container_width=True PARA IGUALAR AO BOTÃO DE DOWNLOAD
    if st.button("INICIAR PROCESSAMENTO", use_container_width=True):
        with st.spinner("Extraindo e calculando dados..."):
            file_bytes = uploaded_file.read()
            df_raw = read_ods_streamlit(file_bytes)
            
            df_limpo = df_raw.drop([0, 1, 2, 3, 4, 7]).reset_index(drop=True)
            
            if df_limpo.iloc[1].astype(str).str.contains("Julho", case=False).any():
                merged = []
                for i in range(0, len(df_limpo), 2):
                    if i + 1 < len(df_limpo):
                        merged.append(pd.concat([df_limpo.iloc[i], df_limpo.iloc[i+1].iloc[4:13]], ignore_index=True))
                df_res = pd.DataFrame(merged)
            else:
                df_res = df_limpo

            header = df_res.iloc[0].astype(str).str.strip().tolist()
            meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            meses_encontrados = [m for m in meses_lista if any(m.lower() == h.lower() for h in header)]
            ultimo_mes = meses_encontrados[-1] if meses_encontrados else "Processamento"
            
            decorridos = len(meses_encontrados)
            restantes = 13 - decorridos

            df_calc = df_res.iloc[1:].copy()
            idx_total = header.index("Total") if "Total" in header else 6
            idx_saldo = header.index("Saldo") if "Saldo" in header else 18
            
            col_total = df_calc.iloc[:, idx_total].apply(to_num)
            col_saldo = df_calc.iloc[:, idx_saldo].apply(to_num)

            df_calc['Liquidado'] = abs(col_saldo - col_total)
            df_calc['Média'] = df_calc['Liquidado'] / decorridos
            df_calc['Projeção'] = df_calc['Média'] * restantes
            df_calc['Suplementar'] = col_saldo - df_calc['Projeção']

            df_calc.iloc[:, 0] = df_calc.iloc[:, 0].replace("", np.nan).ffill()
            df_calc['Órgão'] = df_calc.iloc[:, 0].apply(lambda x: re.sub(r'^\d+\s*', '', str(x)))
            df_calc['Código'] = df_calc.iloc[:, 1]
            df_calc['Despesa'] = df_calc.iloc[:, 2]
            df_calc['Saldo'] = col_saldo

            st.success(f"Cálculo concluído: {decorridos} meses decorridos. Último mês: {ultimo_mes}")
            
            df_view = df_calc[['Órgão', 'Código', 'Despesa', 'Liquidado', 'Saldo', 'Média', 'Projeção', 'Suplementar']].copy()
            df_view_fmt = df_view.copy()
            for col in ['Liquidado', 'Saldo', 'Média', 'Projeção', 'Suplementar']:
                df_view_fmt[col] = df_view_fmt[col].apply(format_currency_br)
            
            st.dataframe(df_view_fmt, use_container_width=True)

            nome_original = os.path.splitext(uploaded_file.name)[0]
            titulo_final = f"Projeção {nome_original} {ultimo_mes}"
            pdf_bytes = gerar_pdf_folha(df_view, decorridos, restantes, titulo_final)
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 BAIXAR RELATÓRIO PDF",
                data=pdf_bytes,
                file_name=f"{titulo_final}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
