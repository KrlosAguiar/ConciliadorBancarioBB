import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
from datetime import datetime

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

st.set_page_config(page_title="Conciliador de Receitas", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
    
    /* CSS PARA BOTÕES */
    div.stButton > button {
        background-color: rgb(38, 39, 48) !important;
        color: white !important;
        font-weight: bold !important;
        border: 1px solid rgb(60, 60, 60);
        border-radius: 5px;
        transition: 0.3s;
        height: 50px; 
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: rgb(20, 20, 25) !important;
        border-color: white;
    }

    /* CARDS DE RESUMO */
    .metric-card {
        background-color: #f8f9fa;
        border-left: 5px solid #ff4b4b;
        padding: 15px;
        border-radius: 5px;
        color: black;
        border: 1px solid #ddd;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin-bottom: 15px;
        height: 100%; /* Altura uniforme */
    }
    .metric-card-green { border-left: 5px solid #28a745; }
    .metric-card-red { border-left: 5px solid #dc3545; }
    
    .metric-title { font-size: 14px; color: #555; text-transform: uppercase; font-weight: bold; margin-bottom: 5px; min-height: 40px;}
    .metric-value { font-size: 16px; font-weight: bold; display: flex; justify-content: space-between; border-bottom: 1px dashed #ccc; padding: 3px 0;}
    .metric-status { margin-top: 10px; padding: 5px; text-align: center; border-radius: 4px; font-weight: bold; font-size: 14px; }
    
    /* TABELAS */
    table { width: 100%; border-collapse: collapse; }
    th { background-color: black; color: white; padding: 10px; text-align: center; }
    td { padding: 8px; border-bottom: 1px solid #ddd; color: black; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES DE SUPORTE E EXTRAÇÃO
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
    mapa_meses = {
        'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'ABRIL': '04',
        'MAIO': '05', 'JUNHO': '06', 'JULHO': '07', 'AGOSTO': '08',
        'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12'
    }
    texto_limpo = texto_pdf.replace('"', ' ').replace('\n', ' ')
    match = re.search(r'Mês:.*?([A-Za-zç]+)\s*/\s*(\d{4})', texto_limpo, re.IGNORECASE)
    if match:
        nome_mes, ano = match.groups()
        mes_num = mapa_meses.get(nome_mes.upper())
        if mes_num:
            return f"{mes_num}/{ano}"
    return None

# --- Funções de Extração (Lógica Original Melhorada) ---

def extrair_bb(arquivo_bytes):
    total = 0.0
    try:
        with pdfplumber.open(arquivo_bytes) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text() or ""
                for linha in texto.split('\n'):
                    if "14109" in linha or "14020" in linha:
                        matches = re.findall(r'([\d\.]+,\d{2})', linha)
                        if matches:
                            total += limpar_numero(matches[0])
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
                        if matches:
                            total += limpar_numero(matches[-1])
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

            # Limpeza crucial para o arquivo da Caixa
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

def encontrar_arquivo_no_upload(lista_arquivos, termo_nome):
    """Encontra o arquivo PDF na lista de uploads pelo nome."""
    if not lista_arquivos: return None
    for arq in lista_arquivos:
        if termo_nome in arq.name:
            return arq
    return None

# ==============================================================================
# 2. FUNÇÕES DE EXIBIÇÃO (HTML & PDF)
# ==============================================================================

def render_card_html(titulo, label_v1, val_v1, label_v2, val_v2):
    ok = round(val_v1, 2) == round(val_v2, 2)
    dif = abs(val_v1 - val_v2)
    classe_cor = "metric-card-green" if ok else "metric-card-red"
    status_txt = "CONCILIADO" if ok else "NÃO CONCILIADO"
    status_bg = "#e8f5e9" if ok else "#fbe9eb"
    status_color = "#28a745" if ok else "#dc3545"
    dif_html = f"<br><small style='color:#c00'>(Dif: {formatar_moeda(dif)})</small>" if not ok else ""

    html = f"""
    <div class="metric-card {classe_cor}">
        <div class="metric-title">{titulo}</div>
        <div class="metric-value">
            <span>{label_v1}</span>
            <span>{formatar_moeda(val_v1)}</span>
        </div>
        <div class="metric-value" style="border-bottom:none;">
            <span>{label_v2}</span>
            <span>{formatar_moeda(val_v2)}</span>
        </div>
        <div class="metric-status" style="background:{status_bg}; color:{status_color};">
            {status_txt}
            {dif_html}
        </div>
    </div>
    """
    return html

def gerar_pdf_relatorio(df_receitas, data_cards):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm, title="Relatório Conciliação")
    story = []
    styles = getSampleStyleSheet()
    
    # Título
    story.append(Paragraph("RELATÓRIO DE CONCILIAÇÃO CONTÁBIL", styles["Title"]))
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph(f"Data de Emissão: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 10*mm))

    # --- Tabela de Receitas ---
    story.append(Paragraph("<b>1. CONCILIAÇÃO DE RECEITAS PRÓPRIAS</b>", styles["Heading2"]))
    story.append(Spacer(1, 3*mm))

    data = [["Conta", "Valor Contábil", "Extrato Bancário", "Diferença"]]
    
    total_contabil = 0
    
    for _, row in df_receitas.iterrows():
        dif = row['Diferença']
        style_dif = colors.red if abs(dif) > 0.01 else colors.darkgreen
        
        # Formatação condicional para o PDF
        val_dif_str = formatar_moeda(dif)
        if abs(dif) < 0.01: val_dif_str = "OK"

        data.append([
            str(row['Conta']),
            formatar_moeda(row['Valor Contábil']),
            formatar_moeda(row['Valor Extrato']),
            val_dif_str
        ])
        total_contabil += row['Valor Contábil']

    # Linha de total
    data.append(["TOTAL GERAL", formatar_moeda(total_contabil), "-", "-"])

    t = Table(data, colWidths=[60*mm, 40*mm, 40*mm, 40*mm])
    
    # Estilo da Tabela
    t_style = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.black),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (1,1), (-1,-1), 'RIGHT'), # Valores à direita
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'), # Total Bold
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
    ]

    # Colorir a coluna de diferença linha a linha
    for i in range(1, len(data)-1):
        val_dif = df_receitas.iloc[i-1]['Diferença']
        cor = colors.red if abs(val_dif) > 0.01 else colors.darkgreen
        t_style.append(('TEXTCOLOR', (3, i), (3, i), cor))
        t_style.append(('FONTNAME', (3, i), (3, i), 'Helvetica-Bold'))

    t.setStyle(TableStyle(t_style))
    story.append(t)
    
    doc.build(story)
    return buffer.getvalue()

# ==============================================================================
# 3. LÓGICA PRINCIPAL (STREAMLIT)
# ==============================================================================

st.markdown("<h1 style='text-align: center;'>Painel de Conciliação Contábil</h1>", unsafe_allow_html=True)
st.markdown("---")

col_up1, col_up2 = st.columns(2)

with col_up1:
    st.markdown("### 1. Razão Contábil (.xlsx)")
    arquivo_excel = st.file_uploader("Carregar Excel", type=["xlsx"], label_visibility="collapsed")

with col_up2:
    st.markdown("### 2. Extratos Bancários (.pdf)")
    arquivos_pdf = st.file_uploader("Carregar PDFs (Opcional)", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed")

if arquivo_excel:
    try:
        # --- PROCESSAMENTO EXCEL ---
        df = pd.read_excel(arquivo_excel, skiprows=6, dtype=str)
        
        # Limpeza Inicial
        mask = df['UG'].str.contains("Totalizadores", case=False, na=False)
        if mask.any(): df = df.iloc[:mask.idxmax()].copy()

        df['Valor'] = df['Valor'].fillna('0')
        df['Valor'] = df['Valor'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce').fillna(0.0)
        df['Conta'] = df['Conta'].astype(str).str.replace(r'\.0$', '', regex=True)

        st.success("Dados Contábeis Processados com Sucesso!")
        st.markdown("---")

        # --- SEÇÃO 1: CARDS DE CONCILIAÇÃO ---
        st.subheader("Situação das Transferências")
        
        # Função auxiliar para calcular totais do card
        def calc_card(col_filtro, v1, v2):
            f1 = df[col_filtro].astype(str).str.startswith(str(v1), na=False)
            f2 = df[col_filtro].astype(str).str.startswith(str(v2), na=False)
            t1 = df.loc[f1, 'Valor'].sum()
            t2 = df.loc[f2, 'Valor'].sum()
            return t1, t2

        # LINHA 1 (4 Cards)
        row1 = st.columns(4)
        configs_r1 = [
            ("TRANSF. CONST. - FMS", 264, 265),
            ("TRANSF. CONST. - FME", 266, 267),
            ("TRANSF. CONST. - FMAS", 268, 269),
            ("TRANSF. PARA ARSEP", 270, 271)
        ]
        
        for idx, (titulo, c1, c2) in enumerate(configs_r1):
            t1, t2 = calc_card('LCP', c1, c2)
            with row1[idx]:
                st.markdown(render_card_html(titulo, str(c1), t1, str(c2), t2), unsafe_allow_html=True)

        # LINHA 2 (3 Cards)
        row2 = st.columns(3)
        configs_r2 = [
            ("TRANSFERÊNCIAS ENTRE UGS", 'LCP', 250, 251),
            ("RENDIMENTOS DE APLICAÇÃO", 'LCP', 258, 259),
            ("TRANSF. DUODÉCIMO", 'Fato Contábil', 'Transferência Financeira Concedida', 'Transferência Financeira Recebida')
        ]

        # 1. Entre UGs
        t1_ugs, t2_ugs = calc_card('LCP', 250, 251)
        with row2[0]:
            st.markdown(render_card_html("TRANSFERÊNCIAS ENTRE UGS", "250", t1_ugs, "251", t2_ugs), unsafe_allow_html=True)

        # 2. Rendimentos
        t1_rend, t2_rend = calc_card('LCP', 258, 259)
        with row2[1]:
            st.markdown(render_card_html("RENDIMENTOS DE APLICAÇÃO", "258", t1_rend, "259", t2_rend), unsafe_allow_html=True)

        # 3. Duodécimo (Filtro por Fato Contábil, lógica ligeiramente diferente)
        f1_duo = df['Fato Contábil'].astype(str).str.startswith("Transferência Financeira Concedida", na=False)
        f2_duo = df['Fato Contábil'].astype(str).str.startswith("Transferência Financeira Recebida", na=False)
        t1_duo = df.loc[f1_duo, 'Valor'].sum()
        t2_duo = df.loc[f2_duo, 'Valor'].sum()
        with row2[2]:
            st.markdown(render_card_html("TRANSF. DUODÉCIMO", "Concedida", t1_duo, "Recebida", t2_duo), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # --- SEÇÃO 2: RECEITAS PRÓPRIAS (TABELA) ---
        st.subheader("Receitas Próprias (Contábil x Bancário)")

        mapa_receitas = {
            '8346': {'pdf_key': '105628',    'func': extrair_bb},
            '8416': {'pdf_key': '112005',    'func': extrair_bb},
            '8364': {'pdf_key': '126022',    'func': extrair_bb},
            '9150': {'pdf_key': '78101',     'func': extrair_bb},
            '9130': {'pdf_key': '575230061', 'func': extrair_caixa},
            '8241': {'pdf_key': '538298',    'func': extrair_banpara}
        }

        # 1. Agrega valores do Excel
        contas_interesse = list(mapa_receitas.keys())
        df_resumo = df[(df['Fato Contábil']=='Arrecadação da Receita') & (df['Conta'].isin(contas_interesse))].groupby('Conta')['Valor'].sum().reset_index()
        
        # Garante que todas as contas apareçam mesmo se não houver no excel
        df_base = pd.DataFrame({'Conta': contas_interesse})
        df_final = pd.merge(df_base, df_resumo, on='Conta', how='left').fillna(0)
        
        # 2. Processa PDFs (se houver)
        resultados = []
        for _, row in df_final.iterrows():
            conta = row['Conta']
            val_contabil = row['Valor']
            
            cfg = mapa_receitas.get(conta)
            val_extrato = 0.0
            msg_status = "Sem PDF"
            
            # Procura o arquivo correspondente na lista de uploads
            arquivo_encontrado = encontrar_arquivo_no_upload(arquivos_pdf, cfg['pdf_key'])
            
            if arquivo_encontrado:
                # Reinicia ponteiro do arquivo para leitura
                arquivo_encontrado.seek(0)
                try:
                    v = cfg['func'](arquivo_encontrado)
                    if v is not None:
                        val_extrato = v
                        msg_status = "OK"
                    else:
                        msg_status = "Erro Leitura"
                except:
                    msg_status = "Erro Geral"
            
            dif = val_contabil - val_extrato
            resultados.append({
                "Conta": conta,
                "PDF Ref": cfg['pdf_key'],
                "Valor Contábil": val_contabil,
                "Valor Extrato": val_extrato,
                "Diferença": dif,
                "Status": msg_status
            })

        df_tabela = pd.DataFrame(resultados)
        
        # --- RENDERIZAÇÃO DA TABELA HTML ---
        html_table = """
        <div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>
        <table style='width:100%; border-collapse: collapse; color: black !important;'>
            <tr style='background-color: black; color: white;'>
                <th>Conta</th>
                <th>PDF Ref</th>
                <th style='text-align: right;'>Valor Contábil</th>
                <th style='text-align: right;'>Valor Extrato</th>
                <th style='text-align: center;'>Diferença</th>
            </tr>
        """
        
        total_g_contabil = 0.0
        
        for _, r in df_tabela.iterrows():
            total_g_contabil += r['Valor Contábil']
            
            # Estilo da diferença
            dif = r['Diferença']
            if abs(dif) < 0.01:
                style_dif = "color: #28a745; font-weight: bold;"
                txt_dif = "CONCILIADO"
            else:
                style_dif = "color: #dc3545; font-weight: bold;"
                txt_dif = formatar_moeda(dif)

            # Estilo se faltar PDF
            style_row = ""
            val_ext_str = formatar_moeda(r['Valor Extrato'])
            if r['Status'] == "Sem PDF":
                val_ext_str = "<span style='color:#999; font-style:italic;'>Falta PDF</span>"
                if r['Valor Contábil'] > 0: txt_dif = "PENDENTE"
            
            html_table += f"""
            <tr style='border-bottom: 1px solid #eee;'>
                <td>{r['Conta']}</td>
                <td style='color:#555; font-size:12px;'>{r['PDF Ref']}</td>
                <td style='text-align: right; font-weight:bold;'>{formatar_moeda(r['Valor Contábil'])}</td>
                <td style='text-align: right;'>{val_ext_str}</td>
                <td style='text-align: center; {style_dif}'>{txt_dif}</td>
            </tr>
            """
            
        html_table += f"""
            <tr style='background-color: #f0f0f0; border-top: 2px solid black;'>
                <td colspan='2'><b>TOTAL GERAL</b></td>
                <td style='text-align: right;'><b>{formatar_moeda(total_g_contabil)}</b></td>
                <td colspan='2'></td>
            </tr>
        </table></div>
        """
        
        st.markdown(html_table, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # --- DOWNLOAD PDF ---
        pdf_bytes = gerar_pdf_relatorio(df_tabela, None)
        st.download_button(
            label="BAIXAR RELATÓRIO PDF",
            data=pdf_bytes,
            file_name="Relatorio_Conciliacao.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo Excel: {e}")
else:
    st.info("Por favor, faça o upload do arquivo 'Razão da Contabilidade GERAL Dez.xlsx' para iniciar.")
