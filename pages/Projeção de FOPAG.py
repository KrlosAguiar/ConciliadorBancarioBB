import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import os
from PIL import Image

# ReportLab Imports
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ==============================================================================
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
try:
    icon_image = Image.open(icon_path)
except:
    icon_image = None # Evita quebrar se a imagem não for encontrada

st.set_page_config(
    page_title="Projeção de FOPAG",
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
    
    /* Regra global para garantir que tabelas HTML fiquem pretas */
    .tabela-fopag td { color: black !important; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE APOIO
# ==============================================================================

def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == "-": return "-"
    try:
        v = float(valor)
        return f"{v:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    except:
        return str(valor)

def to_num(val):
    if val is None or str(val).strip() in ["", "-", "None", "nan"]: return 0.0
    s = str(val).strip()
    try:
        if ',' in s: s = s.replace('.', '').replace(',', '.')
        return float(s)
    except: return 0.0

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
    val_text_std = ParagraphStyle('ValStd', parent=styles['Normal'], fontSize=8, alignment=2) # Direita
    val_text_center = ParagraphStyle('ValCenter', parent=styles['Normal'], fontSize=8, alignment=1) # Centro

    headers = ['Órgão/Unidade', 'Cod', 'Despesa', 'Liquidado', 'Saldo', 'Média', 'Projeção', 'Suplementar']
    data = [headers]
    
    for _, r in df_f.iterrows():
        # Destaque condicional APENAS para a coluna Suplementar
        is_neg = r['Suplementar'] < 0
        sup_font = 10 if is_neg else 8
        sup_style = ParagraphStyle('Sup', parent=val_text_std, 
                                   textColor=colors.red if is_neg else colors.black, 
                                   fontSize=sup_font, 
                                   fontName='Helvetica-Bold' if is_neg else 'Helvetica')
        
        data.append([
            Paragraph(str(r['Órgão']), small_text),
            Paragraph(str(r['Código']), val_text_center), # Código Centralizado
            Paragraph(str(r['Despesa']), small_text),
            Paragraph(formatar_moeda_br(r['Liquidado']), val_text_std),
            Paragraph(formatar_moeda_br(r['Saldo']), val_text_std),
            Paragraph(formatar_moeda_br(r['Média']), val_text_std),
            Paragraph(formatar_moeda_br(r['Projeção']), val_text_std),
            Paragraph(formatar_moeda_br(r['Suplementar']), sup_style)
        ])
    
    t = Table(data, colWidths=[55*mm, 15*mm, 55*mm, 30*mm, 30*mm, 30*mm, 30*mm, 32*mm], repeatRows=1)
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.black),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,0), 'CENTER'), # Títulos centralizados
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
    col_btn1, col_btn2 = st.columns(2)
    modo_processamento = None
    
    with col_btn1:
        if st.button("INICIAR PROCESSAMENTO (até Março)", use_container_width=True):
            modo_processamento = "1_linha"
            
    with col_btn2:
        if st.button("INICIAR PROCESSAMENTO (Abril a Dezembro)", use_container_width=True):
            modo_processamento = "2_linhas"

    if modo_processamento:
        with st.spinner("A processar os dados..."):
            file_bytes = uploaded_file.read()
            
            # 1. Leitura robusta do Pandas (Resolve o agrupamento das células e necessita da odfpy)
            try:
                df_raw = pd.read_excel(io.BytesIO(file_bytes), engine="odf", header=None)
            except Exception as e:
                st.error("Erro na leitura. Certifique-se de que a biblioteca 'odfpy' está instalada no seu ambiente (pip install odfpy).")
                st.stop()
                
            # 2. Busca Dinâmica do Cabeçalho
            idx_cabecalho = 0
            for i in range(min(20, len(df_raw))):
                linha_str = " ".join(df_raw.iloc[i].astype(str).str.upper())
                # Procura a linha que contém as palavras-chave da sua tabela
                if "DESPESA" in linha_str and "SALDO" in linha_str:
                    idx_cabecalho = i
                    break
                    
            # 3. Corta o dataframe exatamente no cabeçalho
            df_limpo = df_raw.iloc[idx_cabecalho:].reset_index(drop=True)
            
            # --- ROTEAMENTO DA LÓGICA BASEADO NO BOTÃO ---
            if modo_processamento == "2_linhas":
                # Lógica Antiga (Separar cabeçalho dos dados e mesclar linhas)
                df_cab = df_limpo.iloc[[0]]
                df_dados = df_limpo.iloc[1:].reset_index(drop=True)
                merged = [df_cab]
                
                for i in range(0, len(df_dados), 2):
                    if i + 1 < len(df_dados):
                        row_c = df_dados.iloc[i].copy()
                        row_baixo = df_dados.iloc[i+1]
                        
                        # Preenche as colunas vazias com a linha de baixo
                        for c_idx in range(len(row_c)):
                            v_c = str(row_c.iloc[c_idx]).strip()
                            if v_c in ["nan", "None", ""]:
                                row_c.iloc[c_idx] = row_baixo.iloc[c_idx]
                                
                        merged.append(row_c.to_frame().T)
                df_res = pd.concat(merged, ignore_index=True)
            else:
                # Formato novo (1 linha direta)
                df_res = df_limpo

            # --- LÓGICA COMPARTILHADA (Meses, Cálculos e Limpeza) ---
            header_row = df_res.iloc[0].astype(str).str.strip().tolist()
            meses_lista = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            meses_encontrados = [m for m in meses_lista if any(m.lower() == h.lower() for h in header_row)]
            ultimo_mes = meses_encontrados[-1] if meses_encontrados else "Processamento"
            decorridos = len(meses_encontrados)
            restantes = 13 - decorridos if decorridos > 0 else 12

            df_calc = df_res.iloc[1:].copy()
            
            # Limpa qualquer linha que tenha vindo totalmente vazia
            df_calc = df_calc.dropna(how='all')

            idx_total = header_row.index("Total") if "Total" in header_row else 6
            idx_saldo = header_row.index("Saldo") if "Saldo" in header_row else 18
            
            col_total = df_calc.iloc[:, idx_total].apply(to_num)
            col_saldo = df_calc.iloc[:, idx_saldo].apply(to_num)

            df_calc['Liquidado'] = abs(col_saldo - col_total)
            df_calc['Média'] = df_calc['Liquidado'] / decorridos if decorridos > 0 else 0
            df_calc['Projeção'] = df_calc['Média'] * restantes
            df_calc['Suplementar'] = col_saldo - df_calc['Projeção']

            # Ajuste do preenchimento do Órgão (ffill) de forma segura
            df_calc.iloc[:, 0] = df_calc.iloc[:, 0].replace(["", "nan", "None"], np.nan).ffill()
            df_calc['Órgão'] = df_calc.iloc[:, 0].apply(lambda x: re.sub(r'^\d+\s*', '', str(x)))
            df_calc['Código'] = df_calc.iloc[:, 1]
            
            # Filtro para ignorar linhas de lixo: só passa o que tem código preenchido
            df_calc = df_calc[df_calc['Código'].astype(str).str.strip() != "nan"]
            df_calc = df_calc[df_calc['Código'].astype(str).str.strip() != "None"]
            df_calc = df_calc[df_calc['Código'].astype(str).str.strip() != ""]

            df_calc['Despesa'] = df_calc.iloc[:, 2]
            df_calc['Saldo_Val'] = col_saldo

            df_f = df_calc[['Órgão', 'Código', 'Despesa', 'Liquidado', 'Saldo_Val', 'Média', 'Projeção', 'Suplementar']].rename(columns={'Saldo_Val': 'Saldo'})

            # --- EXIBIÇÃO EM TELA (CSS INLINE PARA PRETO) ---
            html = f"""
            <div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>
            <table class='tabela-fopag' style='width:100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px; color: black !important; background-color: white;'>
                <tr style='background-color: black; color: white !important;'>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000; color: white !important;'>Órgão/Unidade</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000; color: white !important;'>Cod</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000; color: white !important;'>Despesa</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000; color: white !important;'>Liquidado</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000; color: white !important;'>Saldo</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000; color: white !important;'>Média</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000; color: white !important;'>Projeção</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #000; color: white !important;'>Suplementar</th>
                </tr>"""
            
            for _, r in df_f.iterrows():
                # Destaque Visual APENAS na coluna Suplementar (Vermelho/Negrito se Negativo)
                is_neg = r['Suplementar'] < 0
                s_color = "red" if is_neg else "black"
                s_weight = "bold" if is_neg else "normal"
                
                html += f"""
                <tr style='background-color: white;'>
                    <td style='padding: 5px; border: 1px solid #000; color: black !important; text-align: left;'>{r['Órgão']}</td>
                    <td style='text-align: center; border: 1px solid #000; color: black !important;'>{r['Código']}</td>
                    <td style='padding: 5px; border: 1px solid #000; color: black !important; text-align: left;'>{r['Despesa']}</td>
                    <td style='text-align: right; border: 1px solid #000; color: black !important; padding-right: 5px;'>{formatar_moeda_br(r['Liquidado'])}</td>
                    <td style='text-align: right; border: 1px solid #000; color: black !important; padding-right: 5px;'>{formatar_moeda_br(r['Saldo'])}</td>
                    <td style='text-align: right; border: 1px solid #000; color: black !important; padding-right: 5px;'>{formatar_moeda_br(r['Média'])}</td>
                    <td style='text-align: right; border: 1px solid #000; color: black !important; padding-right: 5px;'>{formatar_moeda_br(r['Projeção'])}</td>
                    <td style='text-align: right; color: {s_color} !important; border: 1px solid #000; font-weight: {s_weight}; padding-right: 5px;'>{formatar_moeda_br(r['Suplementar'])}</td>
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
