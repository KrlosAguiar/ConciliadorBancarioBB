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
from reportlab.lib.units import mm, inch

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
icon_image = Image.open(icon_path)

st.set_page_config(
    page_title="Projeção de Folha",
    page_icon=icon_image,
    layout="wide"
)

# --- CSS PERSONALIZADO ---
st.markdown("""
<style>
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
    div.stButton > button, div.stDownloadButton > button {
        background-color: rgb(38, 39, 48) !important;
        color: white !important;
        font-weight: bold !important;
        border: 1px solid rgb(60, 60, 60);
        border-radius: 5px;
        font-size: 16px;
        height: 50px;
        width: 100%;
    }
    div.stButton > button:hover { background-color: rgb(20, 20, 25) !important; border-color: white; }
    .big-label { font-size: 24px !important; font-weight: 600 !important; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE PROCESSAMENTO
# ==============================================================================

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == "-": return "-"
    try:
        v = float(valor)
        return f"{v:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    except:
        return str(valor)

def to_num(val):
    if val is None or str(val).strip() in ["", "-", "None"]: return 0.0
    s = str(val).strip()
    try:
        if ',' in s: s = s.replace('.', '').replace(',', '.')
        return float(s)
    except: return 0.0

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

# ==============================================================================
# 2. GERAÇÃO DO PDF
# ==============================================================================

def gerar_pdf_final(df_f, decorridos, restantes, titulo_completo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm, title=titulo_completo)
    
    story = []
    styles = getSampleStyleSheet()
    story.append(Paragraph("Projeção de Folha de Pagamento", styles["Title"]))
    params = f"<b>Parâmetros:</b> {decorridos} Meses Decorridos | {restantes} Meses Restantes"
    story.append(Paragraph(params, ParagraphStyle(name='C', alignment=1, fontSize=10)))
    story.append(Spacer(1, 15))
    
    small_text = ParagraphStyle('Small', parent=styles['Normal'], fontSize=7, leading=8)
    val_text = ParagraphStyle('Val', parent=styles['Normal'], fontSize=8, alignment=2)

    headers = ['Órgão/Unidade', 'Cod', 'Despesa', 'Liquidado', 'Saldo', 'Média', 'Projeção', 'Suplementar']
    data = [headers]
    
    for _, r in df_f.iterrows():
        # Ajuste de tamanho se for negativo na coluna Suplementar
        is_neg = r['Suplementar'] < 0
        f_size = 10 if is_neg else 8
        sup_style = ParagraphStyle('Sup', parent=val_text, textColor=colors.red if is_neg else colors.black, 
                                   fontSize=f_size, fontName='Helvetica-Bold' if is_neg else 'Helvetica')
        
        data.append([
            Paragraph(str(r['Órgão']), small_text),
            str(r['Código']),
            Paragraph(str(r['Despesa']), small_text),
            formatar_moeda_br(r['Liquidado']),
            formatar_moeda_br(r['Saldo']),
            formatar_moeda_br(r['Média']),
            formatar_moeda_br(r['Projeção']),
            Paragraph(formatar_moeda_br(r['Suplementar']), sup_style)
        ])
    
    t = Table(data, colWidths=[55*mm, 15*mm, 55*mm, 30*mm, 30*mm, 30*mm, 30*mm, 32*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.black),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,0), 'CENTER'), # Centraliza títulos
        ('ALIGN', (0,1), (2,-1), 'LEFT'),
        ('ALIGN', (3,1), (-1,-1), 'RIGHT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    
    story.append(t)
    doc.build(story)
    return buffer.getvalue()

# ==============================================================================
# 3. INTERFACE
# ==============================================================================



st.markdown("<h1 style='text-align: center;'>Projeção de Folha de Pagamento</h1>", unsafe_allow_html=True)
st.markdown("---")

st.markdown('<p class="big-label">Selecione o arquivo no formato .ods</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type=["ods"], label_visibility="collapsed")

if uploaded_file:
    if st.button("INICIAR PROCESSAMENTO", use_container_width=True):
        with st.spinner("Processando..."):
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

            header_row = df_res.iloc[0].astype(str).str.strip().tolist()
            meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            meses_encontrados = [m for m in meses_lista if any(m.lower() == h.lower() for h in header_row)]
            ultimo_mes = meses_encontrados[-1] if meses_encontrados else "Processamento"
            decorridos = len(meses_encontrados)
            restantes = 13 - decorridos

            # Cálculos
            df_calc = df_res.iloc[1:].copy()
            idx_total = header_row.index("Total") if "Total" in header_row else 6
            idx_saldo = header_row.index("Saldo") if "Saldo" in header_row else 18
            
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
            df_calc['Saldo_Val'] = col_saldo

            df_f = df_calc[['Órgão', 'Código', 'Despesa', 'Liquidado', 'Saldo_Val', 'Média', 'Projeção', 'Suplementar']].rename(columns={'Saldo_Val': 'Saldo'})

            # --- EXIBIÇÃO EM TELA (HTML STYLE CORRIGIDO) ---
            html = """
            <div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>
            <table style='width:100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px;'>
                <tr style='background-color: black; color: white !important;'>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Órgão/Unidade</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Cod</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Despesa</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Liquidado</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Saldo</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Média</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Projeção</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000;'>Suplementar</th>
                </tr>"""
            
            for _, r in df_f.iterrows():
                s_color = "red" if r['Suplementar'] < 0 else "black"
                html += f"""
                <tr style='background-color: white;'>
                    <td style='padding: 5px; border: 1px solid #000; color: black !important;'>{r['Órgão']}</td>
                    <td style='text-align: center; border: 1px solid #000; color: black !important;'>{r['Código']}</td>
                    <td style='padding: 5px; border: 1px solid #000; color: black !important;'>{r['Despesa']}</td>
                    <td style='text-align: right; border: 1px solid #000; color: black !important;'>{formatar_moeda_br(r['Liquidado'])}</td>
                    <td style='text-align: right; border: 1px solid #000; color: black !important;'>{formatar_moeda_br(r['Saldo'])}</td>
                    <td style='text-align: right; border: 1px solid #000; color: black !important;'>{formatar_moeda_br(r['Média'])}</td>
                    <td style='text-align: right; border: 1px solid #000; color: black !important;'>{formatar_moeda_br(r['Projeção'])}</td>
                    <td style='text-align: right; color: {s_color} !important; border: 1px solid #000; font-weight: bold;'>{formatar_moeda_br(r['Suplementar'])}</td>
                </tr>"""
            html += "</table></div>"
            
            st.markdown(html, unsafe_allow_html=True)

            nome_limpo = os.path.splitext(uploaded_file.name)[0]
            titulo_final = f"Projeção {nome_limpo} {ultimo_mes}"
            pdf_data = gerar_pdf_final(df_f, decorridos, restantes, titulo_final)
            
            st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
            st.download_button(
                label="BAIXAR RELATÓRIO PDF",
                data=pdf_data,
                file_name=f"{titulo_final}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
