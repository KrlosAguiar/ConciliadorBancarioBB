import streamlit as st
import pandas as pd
import pdfplumber
import xlsxwriter
import rarfile
import zipfile
import re
import io
import os
import shutil
import tempfile
import unicodedata
from PIL import Image

# --- IMPORTAÇÕES PARA PDF (REPORTLAB) ---
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.units import mm

# Configuração do executável UNRAR
rarfile.UNRAR_TOOL = "unrar"

# ==============================================================================
# 0. CONFIGURAÇÃO DA PÁGINA E CSS
# ==============================================================================
icon_path = "Barcarena.png"
try:
    icon_image = Image.open(icon_path)
    st.set_page_config(page_title="Conciliador de Saldos Bancários", page_icon=icon_image, layout="wide")
except:
    st.set_page_config(page_title="Conciliador de Saldos Bancários", layout="wide")

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
    
    /* --- AJUSTE CSS DOS CARDS --- */
    .metric-card-base {
        background-color: white !important; /* Força fundo branco */
        padding: 15px;
        border-radius: 8px;
        color: black;
        border: 1px solid #ddd;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 15px;
        text-align: center !important; /* Força centralização */
        height: 100%;
    }
    
    /* Variação Vermelha (Padrão para pendências) */
    .metric-card-red {
        border-left: 8px solid #ff4b4b !important;
    }

    /* Variação Verde (Sem pendências) */
    .metric-card-green { 
        border-left: 8px solid #28a745 !important;
    }
    
    .metric-ug-title { font-size: 14px; font-weight: bold; color: #555; margin-bottom: 5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .metric-value { font-size: 26px; font-weight: bold; margin: 5px 0; }
    .metric-label { font-size: 11px; color: #777; text-transform: uppercase; }
    
    /* --- ESTILO DA TABELA HTML EM TELA --- */
    .section-title {
        background-color: black !important;
        color: white !important;
        padding: 10px 15px;
        margin-bottom: 0;
        font-weight: bold;
        border-top-left-radius: 5px;
        border-top-right-radius: 5px;
    }

    .preview-table-container {
        border: 1px solid #000;
        border-top: none;
        margin-bottom: 20px;
        overflow-x: auto;
    }

    .preview-table {
        width: 100%;
        border-collapse: collapse;
        color: black;
        background-color: white;
        font-family: Arial, sans-serif;
    }
    .preview-table th {
        background-color: #f0f0f0; /* Cinza claro no header interno */
        color: black;
        padding: 10px;
        border: 1px solid #ccc;
        text-align: center;
        font-weight: bold;
        font-size: 12px;
    }
    .preview-table td {
        padding: 8px;
        border: 1px solid #ccc;
        text-align: center;
        font-size: 12px;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. FUNÇÕES UTILITÁRIAS
# ==============================================================================

def formatar_moeda(valor):
    if pd.isna(valor): return "0,00"
    try:
        valor_float = float(valor)
    except: return "0,00"
    return f"{valor_float:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

def limpar_numero(valor):
    if pd.isna(valor) or valor == '': return 0.0
    if isinstance(valor, (float, int)): return float(valor)
    valor_str = str(valor).strip().upper()
    valor_str = re.sub(r'[R\$\sA-Z"]', '', valor_str)
    try:
        if ',' in valor_str:
            valor_str = valor_str.replace('.', '').replace(',', '.')
        return float(valor_str)
    except ValueError:
        return 0.0

def extrair_digitos(texto):
    return "".join(filter(str.isdigit, str(texto)))

def limpar_conta_excel(conta_raw):
    texto = str(conta_raw).strip()
    match = re.match(r'^(001|104|037|237|341)\s+(.*)', texto)
    if match:
        return extrair_digitos(match.group(2))
    return extrair_digitos(texto)

def identificar_banco(texto):
    texto_lower = texto.lower()
    if "banparanet" in texto_lower or "banco do estado do para" in texto_lower: return "BANPARA"
    elif "itaú" in texto_lower or "itau" in texto_lower: return "ITAU"
    elif "santander" in texto_lower: return "SANTANDER"
    elif "ouvidoria bb" in texto_lower or "banco do brasil" in texto_lower or "bb.com.br" in texto_lower: return "BB"
    elif "caixa economica" in texto_lower or "caixa.gov.br" in texto_lower or "sac caixa" in texto_lower: return "CAIXA"
    return "DESCONHECIDO"

# ==============================================================================
# 2. MOTOR DE LEITURA DE PDF (V35)
# ==============================================================================

def encontrar_saldo_pdf(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                t = pagina.extract_text()
                if t: texto_completo += "\n" + t

            if not texto_completo.strip(): return 0.0, "Imagem"

            banco = identificar_banco(texto_completo)
            saldo = 0.0
            nome_arquivo = os.path.basename(caminho_pdf).lower()
            eh_aplicacao = "aplic" in nome_arquivo

            if banco == "ITAU":
                matches_invest = re.findall(r"(?:Saldo Líquido|TOTAL LIQUIDO P/RESGATE).*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                matches_mov = re.findall(r"\d{2}/\d{2}\s+SALDO\s+.*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                if matches_invest: saldo = limpar_numero(matches_invest[-1])
                elif matches_mov: saldo = limpar_numero(matches_mov[-1])

            elif banco == "SANTANDER":
                linhas = texto_completo.split('\n')
                meses_encontrados = []
                padrao_mes = r"^(janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+\d{4}"
                for linha in linhas:
                    if re.match(padrao_mes, linha.strip(), re.IGNORECASE): meses_encontrados.append(linha)
                
                linha_alvo = ""
                if len(meses_encontrados) >= 3: linha_alvo = meses_encontrados[1] 
                elif len(meses_encontrados) == 2: linha_alvo = meses_encontrados[-1] 
                elif len(meses_encontrados) == 1: linha_alvo = meses_encontrados[0]

                if linha_alvo:
                    valores = re.findall(r"([\d\.]+,\d{2})", linha_alvo)
                    if len(valores) >= 8: saldo = limpar_numero(valores[7])
                    elif valores: saldo = limpar_numero(valores[-1])
                else:
                    matches = re.findall(r"Saldo Bruto Final.*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                    if matches: saldo = limpar_numero(matches[-1])

            elif banco == "BB":
                texto_sem_aspas = texto_completo.replace('"', '').replace("'", "")
                saldo_encontrado = False
                if not eh_aplicacao:
                    matches_saldo_espacado = re.findall(r"S\s+A\s+L\s+D\s+O.*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                    if matches_saldo_espacado:
                        saldo = limpar_numero(matches_saldo_espacado[-1])
                        saldo_encontrado = True

                if not saldo_encontrado:
                    matches_resumo = re.findall(r"SALDO ATUAL[\s\n]*=[\s\n]*([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                    matches_invest = re.findall(r"SALDO ATUAL\s+([\d\.]+,\d{2})", texto_sem_aspas, re.IGNORECASE)

                    if matches_resumo: saldo = sum(limpar_numero(v) for v in matches_resumo)
                    elif matches_invest: saldo = limpar_numero(matches_invest[-1])
                    else:
                        linhas = texto_completo.split('\n')
                        saldo_encontrado_bb = None
                        padrao_data = r"^\s*\d{2}/\d{2}/\d{4}"
                        for linha in linhas:
                            if re.match(padrao_data, linha):
                                valores = re.findall(r"([\d\.]+,\d{2})[CD]?", linha)
                                if valores: saldo_encontrado_bb = limpar_numero(valores[-1])
                        if saldo_encontrado_bb is not None: saldo = saldo_encontrado_bb
                        else:
                            matches_res = re.findall(r"SALDO ATUAL.*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                            if matches_res: saldo = limpar_numero(matches_res[-1])

            elif banco == "BANPARA":
                texto_limpo = texto_completo.upper().replace("Ã", "A").replace("Ç", "C")
                if "NAO EXISTEM LANCAMENTOS NO PERIODO" in texto_limpo:
                    matches_vazio = re.findall(r"Saldo Conta Corrente.*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE | re.DOTALL)
                    if matches_vazio: saldo = limpar_numero(matches_vazio[0])
                    else: saldo = 0.0
                elif "aplic" in nome_arquivo:
                    matches = re.findall(r"(?:SALDO PARA SAQUE|SALDO TOTAL|SALDO ATUAL|SALDO LÍQUIDO).*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                    if matches: saldo = limpar_numero(matches[-1])
                else:
                    linhas = texto_completo.split('\n')
                    saldo_encontrado_banpara = None
                    padrao_data = r"^\s*\d{2}/\d{2}"
                    for linha in linhas:
                        if re.match(padrao_data, linha):
                            valores = re.findall(r"([\d\.]+,\d{2})", linha)
                            if valores: saldo_encontrado_banpara = limpar_numero(valores[-1])
                    if saldo_encontrado_banpara is not None: saldo = saldo_encontrado_banpara
                    else:
                        matches = re.findall(r"(?:SALDO PARA SAQUE|SALDO TOTAL|SALDO ATUAL).*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                        if matches: saldo = limpar_numero(matches[-1])

            elif banco == "CAIXA":
                if "aplic" not in nome_arquivo:
                    linhas = texto_completo.split('\n')
                    saldo_encontrado_caixa = None
                    padrao_data = r"^\s*\d{2}/\d{2}/\d{4}"
                    for linha in linhas:
                        if re.match(padrao_data, linha):
                            valores = re.findall(r"([\d\.]+,\d{2})[CD]?", linha)
                            if valores: saldo_encontrado_caixa = limpar_numero(valores[-1])
                    if saldo_encontrado_caixa is not None: saldo = saldo_encontrado_caixa
                    else:
                        matches_dia = re.findall(r"SALDO DIA.*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                        if matches_dia: saldo = limpar_numero(matches_dia[-1])
                else:
                    matches_bruto = re.findall(r"SALDO BRUTO.*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                    if matches_bruto: saldo = limpar_numero(matches_bruto[-1])
                    else:
                        matches_dia = re.findall(r"SALDO DIA.*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                        if matches_dia: saldo = limpar_numero(matches_dia[-1])

            else:
                matches = re.findall(r"(?:Saldo Final|Total Disponível).*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                if matches: saldo = limpar_numero(matches[-1])

            return saldo, banco
    except:
        return 0.0, "Erro"

# ==============================================================================
# 3. LÓGICA DE CONSOLIDAÇÃO E CONFRONTO (CORRIGIDA E ROBUSTA)
# ==============================================================================

def ler_planilha_e_consolidar(file_obj):
    import unicodedata
    
    # Normalização para busca segura (ignora maiúsculas e acentos)
    def normalizar(txt):
        if pd.isna(txt): return ""
        txt_str = str(txt).lower().strip()
        return "".join(c for c in unicodedata.normalize('NFD', txt_str) if unicodedata.category(c) != 'Mn')

    # 1. TENTATIVA DE LEITURA (EXCEL OU TEXTO/CSV)
    df_raw = pd.DataFrame()
    try:
        # Tenta como Excel nativo
        df_raw = pd.read_excel(file_obj, header=None, engine='openpyxl', dtype=object)
    except:
        try:
            # Se falhar, tenta como CSV/HTML (comum em exportações de sistemas)
            file_obj.seek(0)
            df_raw = pd.read_csv(file_obj, header=None, sep=None, engine='python', encoding='latin-1', dtype=object)
        except Exception as e:
            st.error(f"Erro crítico ao ler o arquivo: {e}")
            return {}

    idx_header = -1
    col_map = {} 
    
    # 2. BUSCA DO CABEÇALHO INTELIGENTE (Detecta formato Antigo ou Novo)
    for i, row in df_raw.head(30).iterrows():
        # Cria linha normalizada
        row_norm = [normalizar(x) for x in row.values]
        row_str = " ".join(row_norm)
        
        # Procura "Conta" E ("Saldo Atual" OU "Saldo Contábil")
        if "conta" in row_str and ("saldo atual" in row_str or "saldo contabil" in row_str):
            idx_header = i
            
            # Mapeamento Dinâmico das Colunas
            for col_idx, val in enumerate(row_norm):
                if val == "conta": 
                    col_map['CODIGO'] = col_idx
                elif "descri" in val: 
                    col_map['DESCRICAO'] = col_idx
                elif "banco" in val: 
                    col_map['CONTA_BANCO'] = col_idx
                elif "saldo atual" in val or "saldo contabil" in val:
                    col_map['RAZAO'] = col_idx
            break
            
    if idx_header == -1 or not col_map:
        st.error("Formato do relatório desconhecido. Não foi possível localizar as colunas 'Conta', 'Banco' e 'Saldo'.")
        return {}

    dados_consolidados = {} 
    grupo_atual = "MOVIMENTO"

    for index, row in df_raw.iloc[idx_header+1:].iterrows():
        # Verifica se o índice da coluna existe na linha atual
        if max(col_map.values()) >= len(row): continue

        linha_texto = " ".join([normalizar(x) for x in row.values if pd.notna(x)])
        
        if "conta movimento" in linha_texto: grupo_atual = "MOVIMENTO"; continue
        elif "conta aplicacao" in linha_texto: grupo_atual = "APLICACAO"; continue
        
        try:
            conta_raw = str(row[col_map['CONTA_BANCO']]).strip()
            codigo = str(row[col_map['CODIGO']]).strip()
            
            # Filtros de validação da linha
            if not codigo or codigo.lower() == 'nan' or not conta_raw or conta_raw.lower() == 'nan': continue
            if "banco" in conta_raw.lower(): continue # Evita ler repetição de cabeçalho

            numeros_conta_match = limpar_conta_excel(conta_raw)
            numeros_conta_full = extrair_digitos(conta_raw)

            if codigo.isdigit() and len(conta_raw) > 3:
                descricao = str(row[col_map['DESCRICAO']]).strip()
                valor_final = limpar_numero(row[col_map['RAZAO']])
                chave = (numeros_conta_full, grupo_atual)

                if chave in dados_consolidados:
                    dados_consolidados[chave]['RAZÃO'] += valor_final
                    if descricao not in dados_consolidados[chave]['DESCRICAO']:
                         dados_consolidados[chave]['DESCRICAO'] += f" / {descricao}"
                else:
                    dados_consolidados[chave] = {
                        "CÓDIGO": codigo,
                        "DESCRIÇÃO": descricao,
                        "CONTA": conta_raw,
                        "RAZÃO": valor_final,
                        "GRUPO": grupo_atual,
                        "EXTRATO": 0.0,
                        "ARQUIVO_ORIGEM": "",
                        "TEM_PDF": False,
                        "MATCH_KEY": numeros_conta_match,
                        "UG": "N/D"
                    }
        except: continue
    return dados_consolidados

def processar_confronto(pasta_extratos, dados_dict):
    chaves_existentes = list(dados_dict.keys())
    arquivos_aplicacao = []
    arquivos_movimento = []

    for root, dirs, files in os.walk(pasta_extratos):
        if "__MACOSX" in root: continue
        nome_ug = os.path.basename(root)
        if root == pasta_extratos: nome_ug = "Raiz"

        for file in files:
            if file.lower().endswith('.pdf'):
                item = {
                    'caminho': os.path.join(root, file),
                    'nome': file,
                    'ug': nome_ug,
                    'numeros': extrair_digitos(file),
                    'processado': False
                }
                if "aplic" in file.lower(): arquivos_aplicacao.append(item)
                else: arquivos_movimento.append(item)

    def match_pdf(pdf_list, grupo_alvo):
        for pdf in pdf_list:
            saldo_extrato, _ = encontrar_saldo_pdf(pdf['caminho'])
            pdf['saldo'] = saldo_extrato
            for (conta_excel, grupo_excel) in chaves_existentes:
                if grupo_excel == grupo_alvo:
                    match_key = dados_dict[(conta_excel, grupo_excel)]['MATCH_KEY']
                    if len(pdf['numeros']) >= 4 and (pdf['numeros'] in match_key or match_key in pdf['numeros']):
                        dados_dict[(conta_excel, grupo_excel)]['EXTRATO'] = pdf['saldo']
                        dados_dict[(conta_excel, grupo_excel)]['ARQUIVO_ORIGEM'] = pdf['nome']
                        dados_dict[(conta_excel, grupo_excel)]['UG'] = pdf['ug']
                        dados_dict[(conta_excel, grupo_excel)]['TEM_PDF'] = True
                        pdf['processado'] = True
                        break
    
    match_pdf(arquivos_aplicacao, "APLICACAO")
    match_pdf(arquivos_movimento, "MOVIMENTO")

    # REPESCAGEM RIGIDA
    for pdf in arquivos_aplicacao + arquivos_movimento:
        if not pdf['processado']:
            grupo_pdf = "APLICACAO" if "aplic" in pdf['nome'].lower() else "MOVIMENTO"
            for (conta_excel, grupo_excel) in chaves_existentes:
                if grupo_excel == grupo_pdf and not dados_dict[(conta_excel, grupo_excel)]['TEM_PDF']:
                    match_key = dados_dict[(conta_excel, grupo_excel)]['MATCH_KEY']
                    if len(pdf['numeros']) >= 4 and (pdf['numeros'] in match_key or match_key in pdf['numeros']):
                        dados_dict[(conta_excel, grupo_excel)]['EXTRATO'] = pdf['saldo']
                        dados_dict[(conta_excel, grupo_excel)]['ARQUIVO_ORIGEM'] = pdf['nome'] + " (Repescagem)"
                        dados_dict[(conta_excel, grupo_excel)]['UG'] = pdf['ug']
                        dados_dict[(conta_excel, grupo_excel)]['TEM_PDF'] = True
                        pdf['processado'] = True
                        break

    lista_final = []
    for chave, dados in dados_dict.items():
        dados['DIFERENÇA'] = round(dados['RAZÃO'] - dados['EXTRATO'], 2)
        lista_final.append(dados)
    
    for pdf in arquivos_aplicacao + arquivos_movimento:
        if not pdf['processado']:
            grupo_pdf = "APLICACAO" if "aplic" in pdf['nome'].lower() else "MOVIMENTO"
            lista_final.append({
                "UG": pdf['ug'], "CÓDIGO": "N/A", "DESCRIÇÃO": "PDF sem conta no Excel",
                "CONTA": pdf['nome'], "RAZÃO": 0.0, "GRUPO": grupo_pdf,
                "EXTRATO": pdf['saldo'], "DIFERENÇA": round(0.0 - pdf['saldo'], 2), "ARQUIVO_ORIGEM": pdf['nome']
            })

    return pd.DataFrame(lista_final)

# ==============================================================================
# 4. FUNÇÕES DE GERAÇÃO DE RELATÓRIOS (EXCEL E PDF)
# ==============================================================================

def aplicar_estilo_excel(ws, wb, df, start_row, cols):
    fmt_moeda = wb.add_format({'num_format': '#,##0.00'})
    fmt_red = wb.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE', 'num_format': '#,##0.00'})
    fmt_header = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    
    for i, col in enumerate(cols):
        ws.write(start_row, i, col, fmt_header)
    
    try:
        idx_dif = cols.index("DIFERENÇA")
        letra = chr(65 + idx_dif) 
        ws.conditional_format(f'{letra}{start_row+2}:{letra}{start_row+1+len(df)}', 
                              {'type': 'cell', 'criteria': 'not between', 'minimum': -0.009, 'maximum': 0.009, 'format': fmt_red})
    except: pass
    
    for i, col in enumerate(cols):
        max_len = max(df[col].astype(str).map(len).max() if not df[col].empty else 0, len(col))
        largura_final = min(max_len + 3, 60) 
        if col in ["RAZÃO", "EXTRATO", "DIFERENÇA"]:
            ws.set_column(i, i, width=largura_final, cell_format=fmt_moeda)
        else:
            ws.set_column(i, i, width=largura_final)

def gerar_pdf_conciliacao(df_final):
    buffer = io.BytesIO()
    # Usa paisagem (landscape) para caber melhor as colunas
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=10*mm, leftMargin=10*mm, topMargin=15*mm, bottomMargin=15*mm)
    story = []
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    title_style.alignment = 1 # Centralizado

    # --- TÍTULO PRINCIPAL ---
    story.append(Paragraph("Relatório de Conciliação de Saldos Bancários", title_style))
    story.append(Spacer(1, 10*mm))

    # --- SEÇÃO DE CARDS (RESUMO POR UG) ---
    st_card_title = ParagraphStyle(name='CardTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, alignment=1, textColor=colors.darkgray)
    st_card_value = ParagraphStyle(name='CardValue', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=16, alignment=1, spaceBefore=4, spaceAfter=4)
    st_card_label = ParagraphStyle(name='CardLabel', parent=styles['Normal'], fontName='Helvetica', fontSize=8, alignment=1, textColor=colors.gray)

    ugs_unicas = sorted(df_final['UG'].unique())
    card_data_matrix = []
    row_cards = []

    for i, ug in enumerate(ugs_unicas):
        df_ug = df_final[df_final['UG'] == ug]
        pendencias = len(df_ug[abs(df_ug['DIFERENÇA']) > 0.01])
        
        cor_valor = colors.red if pendencias > 0 else colors.green
        cor_borda = colors.red if pendencias > 0 else colors.green
        
        # Cria uma tabela interna para o conteúdo do card
        sub_data = [
            [Paragraph(ug, st_card_title)],
            [Paragraph(str(pendencias), ParagraphStyle(name='V', parent=st_card_value, textColor=cor_valor))],
            [Paragraph("PENDÊNCIAS", st_card_label)]
        ]
        sub_table = Table(sub_data, colWidths=[48*mm])
        # Estilo do card: Borda esquerda grossa e colorida, resto cinza fino
        sub_table.setStyle(TableStyle([
            ('BOX', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('LINEBEFORE', (0,0), (0,-1), 6, cor_borda),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('BACKGROUND', (0,0), (-1,-1), colors.white)
        ]))
        row_cards.append(sub_table)

        # Agrupa em linhas de 5 cards para o PDF paisagem
        if len(row_cards) == 5 or i == len(ugs_unicas) - 1:
            # Preenche com células vazias se a última linha tiver menos de 5
            while len(row_cards) < 5: row_cards.append("") 
            card_data_matrix.append(row_cards)
            row_cards = []

    if card_data_matrix:
        # Tabela contêiner para os cards
        t_cards = Table(card_data_matrix, colWidths=[50*mm]*5, hAlign='CENTER')
        t_cards.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 5),
            ('RIGHTPADDING', (0,0), (-1,-1), 5),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 15), # Espaço após cada linha de cards
        ]))
        story.append(t_cards)
        story.append(Spacer(1, 5*mm))

    # --- FUNÇÃO PARA GERAR TABELAS DE DADOS ---
    def add_data_table(df_part, titulo_secao):
        if df_part.empty: return
        
        # Cabeçalho da Seção (Fundo Preto, Texto Branco)
        header_style = ParagraphStyle(name='SectionHeader', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, alignment=0, textColor=colors.white, backColor=colors.black, padding=8, borderPadding=6)
        story.append(KeepTogether(Paragraph(titulo_secao.upper(), header_style)))

        # Dados da Tabela
        cols_pdf = ["UG", "CÓDIGO", "DESCRIÇÃO", "CONTA", "RAZÃO", "EXTRATO", "DIFERENÇA"]
        data = [cols_pdf]
        
        ts = [
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('ALIGN', (4,1), (-1,-1), 'RIGHT'), # Alinha valores à direita
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]

        for i, row in df_part.iterrows():
            dif = row['DIFERENÇA']
            row_data = [
                str(row['UG']),
                str(row['CÓDIGO']),
                Paragraph(str(row['DESCRIÇÃO']), ParagraphStyle(name='DescTiny', fontSize=7)), # Descrição menor
                str(row['CONTA']),
                formatar_moeda(row['RAZÃO']),
                formatar_moeda(row['EXTRATO']),
                formatar_moeda(dif)
            ]
            data.append(row_data)
            
            # Destaca diferença em vermelho
            if abs(dif) > 0.01:
                ts.append(('TEXTCOLOR', (6, i+1), (6, i+1), colors.red))
                ts.append(('FONTNAME', (6, i+1), (6, i+1), 'Helvetica-Bold'))

        # Larguras das colunas para A4 Paisagem (~277mm úteis)
        col_widths = [35*mm, 18*mm, 85*mm, 35*mm, 35*mm, 35*mm, 35*mm]
        t_data = Table(data, colWidths=col_widths, repeatRows=1)
        t_data.setStyle(TableStyle(ts))
        story.append(t_data)
        story.append(Spacer(1, 10*mm))

    # Divide os dados
    df_app = df_final[df_final['GRUPO'] == 'APLICACAO'].copy()
    df_mov = df_final[df_final['GRUPO'] == 'MOVIMENTO'].copy()

    # Adiciona as tabelas ao PDF
    add_data_table(df_app, "Contas Aplicação")
    add_data_table(df_mov, "Contas Movimento")

    doc.build(story)
    return buffer.getvalue()

# ==============================================================================
# 5. INTERFACE STREAMLIT
# ==============================================================================

st.markdown("<h1 style='text-align: center;'>Conciliador de Saldos Bancários</h1>", unsafe_allow_html=True)
st.markdown("---")

c1, c2 = st.columns(2)
with c1: 
    st.markdown('<p class="big-label">Selecione o arquivo .zip/.rar com os extratos bancários</p>', unsafe_allow_html=True)
    up_extratos = st.file_uploader("", type=["zip", "rar"], key="up_pdf", label_visibility="collapsed")
with c2: 
    st.markdown('<p class="big-label">Selecione o relatório de saldos bancários</p>', unsafe_allow_html=True)
    up_planilha = st.file_uploader("", type=["xlsx", "csv"], key="up_xlsx", label_visibility="collapsed")

if st.button("PROCESSAR CONCILIAÇÃO DE SALDOS BANCÁRIOS", use_container_width=True):
    if up_extratos and up_planilha:
        with st.spinner("Processando..."):
            
            temp_dir = tempfile.mkdtemp()
            try:
                # --- Preparação ---
                path_zip = os.path.join(temp_dir, up_extratos.name)
                with open(path_zip, "wb") as f:
                    f.write(up_extratos.getbuffer())
                
                try:
                    if up_extratos.name.lower().endswith('.rar'):
                        with rarfile.RarFile(path_zip, 'r') as r: r.extractall(temp_dir)
                    else:
                        with zipfile.ZipFile(path_zip, 'r') as z: z.extractall(temp_dir)
                except Exception as e:
                    st.error(f"Erro ao descompactar: {e}. Verifique se o arquivo não está corrompido.")
                    shutil.rmtree(temp_dir)
                    st.stop()

                # --- Leitura da Planilha ---
                dados_excel = ler_planilha_e_consolidar(up_planilha)
                if not dados_excel:
                    # st.error já foi chamado dentro da função se falhar
                    shutil.rmtree(temp_dir)
                    st.stop()

                # --- Processamento ---
                df_final = processar_confronto(temp_dir, dados_excel)

                if not df_final.empty:
                    # Ordenação e Limpeza
                    df_final['is_nd'] = df_final['UG'] == 'N/D'
                    df_final = df_final.sort_values(by=['is_nd', 'UG', 'CÓDIGO']).drop(columns=['is_nd'])

                    cols_view = ["UG", "CÓDIGO", "DESCRIÇÃO", "CONTA", "RAZÃO", "EXTRATO", "DIFERENÇA", "ARQUIVO_ORIGEM"]
                    cols_validas = [c for c in cols_view if c in df_final.columns]
                    
                    # CORREÇÃO: df_view é apenas para visualização. df_final (com GRUPO) é usado para downloads.
                    df_view = df_final[cols_validas].copy()

                    # ==========================================================
                    # EXIBIÇÃO DE CARDS EM TELA
                    # ==========================================================
                    st.markdown("### Resumo de Pendências por UG")
                    
                    ugs_unicas = sorted(df_view['UG'].unique())
                    
                    for i in range(0, len(ugs_unicas), 4):
                        cols = st.columns(4)
                        for j in range(4):
                            if i + j < len(ugs_unicas):
                                ug_atual = ugs_unicas[i+j]
                                df_ug = df_view[df_view['UG'] == ug_atual]
                                pendencias = len(df_ug[abs(df_ug['DIFERENÇA']) > 0.01])
                                
                                # Usa classe base + classe de cor da borda
                                classe_cor = "metric-card-green" if pendencias == 0 else "metric-card-red"
                                cor_texto = "#28a745" if pendencias == 0 else "#ff4b4b"
                                
                                html_card = f"""
                                <div class="metric-card-base {classe_cor}">
                                    <div class="metric-ug-title" title="{ug_atual}">{ug_atual}</div>
                                    <div class="metric-value" style="color: {cor_texto};">{pendencias}</div>
                                    <div class="metric-label">Pendências</div>
                                </div>
                                """
                                with cols[j]:
                                    st.markdown(html_card, unsafe_allow_html=True)

                    st.markdown("---")

                    # ==========================================================
                    # TABELA HTML EM TELA (Com Coluna Arquivo Origem)
                    # ==========================================================
                    # ATENÇÃO: Usando df_final para filtrar por GRUPO corretamente
                    df_app_view = df_view[df_final['GRUPO'] == 'APLICACAO']
                    df_mov_view = df_view[df_final['GRUPO'] == 'MOVIMENTO']
                    
                    def gerar_tabela_html(df_input, titulo):
                        if df_input.empty: return ""
                        # Título com fundo preto e texto branco
                        html = f"<h4 class='section-title'>{titulo.upper()}</h4>"
                        html += "<div class='preview-table-container'>"
                        html += "<table class='preview-table'>"
                        # Adicionada coluna ARQUIVO ORIGEM no cabeçalho
                        html += "<thead><tr><th>UG</th><th>CÓDIGO</th><th>DESCRIÇÃO</th><th>CONTA</th><th>RAZÃO</th><th>EXTRATO</th><th>DIFERENÇA</th><th>ARQUIVO ORIGEM</th></tr></thead><tbody>"
                        
                        for _, row in df_input.iterrows():
                            dif = row['DIFERENÇA']
                            style_dif = "color: red; font-weight: bold;" if abs(dif) > 0.01 else "color: black;"
                            html += "<tr>"
                            html += f"<td>{row['UG']}</td><td>{row['CÓDIGO']}</td><td style='text-align: left;'>{row['DESCRIÇÃO']}</td><td>{row['CONTA']}</td>"
                            html += f"<td style='text-align: right;'>{formatar_moeda(row['RAZÃO'])}</td><td style='text-align: right;'>{formatar_moeda(row['EXTRATO'])}</td>"
                            html += f"<td style='text-align: right; {style_dif}'>{formatar_moeda(dif)}</td>"
                            # Adicionada coluna ARQUIVO ORIGEM na linha
                            html += f"<td style='font-size: 11px; font-style: italic; color: #555;'>{row['ARQUIVO_ORIGEM']}</td></tr>"
                        html += "</tbody></table></div><br>"
                        return html

                    st.markdown(gerar_tabela_html(df_app_view, "Contas Aplicação"), unsafe_allow_html=True)
                    st.markdown(gerar_tabela_html(df_mov_view, "Contas Movimento"), unsafe_allow_html=True)

                    # ==========================================================
                    # GERAÇÃO DOS ARQUIVOS PARA DOWNLOAD (Excel e PDF)
                    # ==========================================================
                    
                    # Filtrando do df_final (que TEM a coluna GRUPO)
                    df_app_final = df_final[df_final['GRUPO'] == 'APLICACAO']
                    df_mov_final = df_final[df_final['GRUPO'] == 'MOVIMENTO']

                    # 1. Excel (Mantém lógica original mas com colunas filtradas na escrita)
                    output_excel = io.BytesIO()
                    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                        wb = writer.book
                        ws = wb.add_worksheet('Conciliacao')
                        writer.sheets['Conciliacao'] = ws
                        row = 0
                        if not df_app_final.empty:
                            ws.write(row, 0, "--- CONTAS APLICAÇÃO ---", wb.add_format({'bold':True, 'font_color':'blue'}))
                            row+=1
                            # Exporta apenas as colunas visuais
                            df_app_final[cols_validas].to_excel(writer, sheet_name='Conciliacao', startrow=row, index=False)
                            aplicar_estilo_excel(ws, wb, df_app_final[cols_validas], row, cols_validas)
                            row += len(df_app_final) + 2
                        row += 1
                        if not df_mov_final.empty:
                            ws.write(row, 0, "--- CONTAS MOVIMENTO ---", wb.add_format({'bold':True, 'font_color':'blue'}))
                            row+=1
                            df_mov_final[cols_validas].to_excel(writer, sheet_name='Conciliacao', startrow=row, index=False)
                            aplicar_estilo_excel(ws, wb, df_mov_final[cols_validas], row, cols_validas)
                    output_excel.seek(0)

                    # 2. PDF (Usa df_final completo para ter acesso ao GRUPO dentro da função)
                    pdf_bytes = gerar_pdf_conciliacao(df_final)

                    st.success("Processamento concluído com sucesso!")

                    col_d1, col_d2 = st.columns(2)
                    with col_d1:
                        st.download_button(
                            label="BAIXAR RELATÓRIO EM EXCEL",
                            data=output_excel,
                            file_name="Relatorio_Conciliacao_Saldos.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    with col_d2:
                         st.download_button(
                            label="BAIXAR RELATÓRIO EM PDF",
                            data=pdf_bytes,
                            file_name="Relatorio_Conciliacao_Saldos.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

                else:
                    st.warning("O processamento não gerou dados. Verifique os arquivos.")
            
            except Exception as e:
                st.error(f"Erro fatal: {e}")
            finally:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
    else:
        st.warning("⚠️ Selecione o arquivo ZIP/RAR e a Planilha Excel primeiro.")
