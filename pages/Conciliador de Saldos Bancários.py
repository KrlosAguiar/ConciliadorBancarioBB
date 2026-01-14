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
from PIL import Image

# Configuração do executável UNRAR (Necessário para o Cloud Run/Linux)
rarfile.UNRAR_TOOL = "unrar"

# --- CONFIGURAÇÃO DA PÁGINA (VISUAL IDÊNTICO AO EXEMPLO) ---
icon_path = "Barcarena.png" # Certifique-se de ter essa imagem na pasta ou remova o try/except
try:
    icon_image = Image.open(icon_path)
    st.set_page_config(page_title="Conciliador de Saldos Bancários", page_icon=icon_image, layout="wide")
except:
    st.set_page_config(page_title="Conciliador de Saldos Bancários", layout="wide")

# --- CSS PERSONALIZADO (MANTIDO DO SEU EXEMPLO) ---
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
    
    /* Estilo para a tabela de preview */
    .preview-table {
        width: 100%;
        border-collapse: collapse;
        color: black;
        background-color: white;
    }
    .preview-table th {
        background-color: black;
        color: white;
        padding: 8px;
        border: 1px solid #000;
        text-align: center;
    }
    .preview-table td {
        padding: 8px;
        border: 1px solid #000;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. FUNÇÕES DA LÓGICA V32 (MOTOR DE PROCESSAMENTO)
# ==============================================================================

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

            if banco == "ITAU":
                matches = re.findall(r"(?:Saldo Líquido|TOTAL LIQUIDO P/RESGATE).*?([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)
                if matches: saldo = limpar_numero(matches[-1])

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
                matches_invest = re.findall(r"SALDO ATUAL.*?([\d\.]+,\d{2})", texto_sem_aspas, re.IGNORECASE | re.DOTALL)
                matches_resumo = re.findall(r"SALDO ATUAL[\s\n]*=[\s\n]*([\d\.]+,\d{2})", texto_completo, re.IGNORECASE)

                if matches_resumo: saldo = sum(limpar_numero(v) for v in matches_resumo)
                elif matches_invest: saldo = limpar_numero(matches_invest[-1])
                else:
                    linhas = texto_completo.split('\n')
                    saldo_encontrado_bb = None
                    padrao_data = r"^\s*\d{2}/\d{2}/\d{4}"
                    padrao_saldo_final = r"S\s+A\s+L\s+D\s+O"
                    for linha in linhas:
                        if re.search(padrao_saldo_final, linha, re.IGNORECASE):
                             valores = re.findall(r"([\d\.]+,\d{2})[CD]?", linha)
                             if valores: saldo_encontrado_bb = limpar_numero(valores[-1])
                        elif re.match(padrao_data, linha):
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

def ler_planilha_e_consolidar(file_obj):
    # Aceita objeto de arquivo (BytesIO ou TempFile)
    try:
        if file_obj.name.endswith('.csv'):
            df_raw = pd.read_csv(file_obj, header=None, dtype=object)
        else:
            df_raw = pd.read_excel(file_obj, header=None, engine='openpyxl', dtype=object)
    except Exception as e:
        st.error(f"Erro ao ler planilha: {e}")
        return {}

    idx_header = -1
    col_map = {'CODIGO': 0, 'DESCRICAO': 2, 'CONTA_BANCO': 8, 'RAZAO': 10}
    
    for i, row in df_raw.head(20).iterrows():
        row_str = " ".join([str(x) for x in row.values]).lower()
        if "conta" in row_str and "saldo atual" in row_str:
            idx_header = i
            for col_idx, val in enumerate(row.values):
                val_str = str(val).lower().strip()
                if val_str == "conta": col_map['CODIGO'] = col_idx
                elif "descri" in val_str: col_map['DESCRICAO'] = col_idx
                elif "banco" in val_str: col_map['CONTA_BANCO'] = col_idx
                elif "saldo atual" in val_str: col_map['RAZAO'] = col_idx
            break
    
    if idx_header == -1: idx_header = 7
    dados_consolidados = {} 
    grupo_atual = "MOVIMENTO"

    for index, row in df_raw.iloc[idx_header+1:].iterrows():
        linha_texto = " ".join([str(x) for x in row.values if pd.notna(x)]).lower()
        if "conta movimento" in linha_texto: grupo_atual = "MOVIMENTO"; continue
        elif "conta aplicação" in linha_texto or "conta aplicacao" in linha_texto: grupo_atual = "APLICACAO"; continue
        
        try:
            if max(col_map.values()) >= len(row): continue
            conta_raw = str(row[col_map['CONTA_BANCO']]).strip()
            codigo = str(row[col_map['CODIGO']]).strip()
            
            numeros_conta_match = limpar_conta_excel(conta_raw)
            numeros_conta_full = extrair_digitos(conta_raw)

            if codigo.isdigit() and len(conta_raw) > 3 and "banco" not in conta_raw.lower():
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
                        "MATCH_KEY": numeros_conta_match
                    }
        except: continue
    return dados_consolidados

def processar_confronto(pasta_extratos, dados_dict):
    chaves_existentes = list(dados_dict.keys())
    arquivos_aplicacao = []
    arquivos_movimento = []

    for root, dirs, files in os.walk(pasta_extratos):
        nome_ug = os.path.basename(root)
        if "__MACOSX" in root: continue
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

    # 1. APLICAÇÃO (Sem misturar)
    for pdf in arquivos_aplicacao:
        saldo_extrato, _ = encontrar_saldo_pdf(pdf['caminho'])
        pdf['saldo'] = saldo_extrato
        for (conta_excel, grupo_excel) in chaves_existentes:
            if grupo_excel == "APLICACAO":
                match_key = dados_dict[(conta_excel, grupo_excel)]['MATCH_KEY']
                if len(pdf['numeros']) >= 4 and (pdf['numeros'] in match_key or match_key in pdf['numeros']):
                    dados_dict[(conta_excel, grupo_excel)]['EXTRATO'] = pdf['saldo']
                    dados_dict[(conta_excel, grupo_excel)]['ARQUIVO_ORIGEM'] = pdf['nome']
                    dados_dict[(conta_excel, grupo_excel)]['UG'] = pdf['ug']
                    dados_dict[(conta_excel, grupo_excel)]['TEM_PDF'] = True
                    pdf['processado'] = True
                    break

    # 2. MOVIMENTO (Sem misturar)
    for pdf in arquivos_movimento:
        saldo_extrato, _ = encontrar_saldo_pdf(pdf['caminho'])
        pdf['saldo'] = saldo_extrato
        for (conta_excel, grupo_excel) in chaves_existentes:
            if grupo_excel == "MOVIMENTO":
                match_key = dados_dict[(conta_excel, grupo_excel)]['MATCH_KEY']
                if len(pdf['numeros']) >= 4 and (pdf['numeros'] in match_key or match_key in pdf['numeros']):
                    dados_dict[(conta_excel, grupo_excel)]['EXTRATO'] = pdf['saldo']
                    dados_dict[(conta_excel, grupo_excel)]['ARQUIVO_ORIGEM'] = pdf['nome']
                    dados_dict[(conta_excel, grupo_excel)]['UG'] = pdf['ug']
                    dados_dict[(conta_excel, grupo_excel)]['TEM_PDF'] = True
                    pdf['processado'] = True
                    break

    # 3. REPESCAGEM RIGIDA
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
        if not dados.get('TEM_PDF'): dados['UG'] = 'N/D'
        dados['DIFERENÇA'] = dados['RAZÃO'] - dados['EXTRATO']
        lista_final.append(dados)
    
    for pdf in arquivos_aplicacao + arquivos_movimento:
        if not pdf['processado']:
            grupo_pdf = "APLICACAO" if "aplic" in pdf['nome'].lower() else "MOVIMENTO"
            lista_final.append({
                "UG": pdf['ug'], "CÓDIGO": "N/A", "DESCRIÇÃO": "PDF sem conta no Excel",
                "CONTA": pdf['nome'], "RAZÃO": 0.0, "GRUPO": grupo_pdf,
                "EXTRATO": pdf['saldo'], "DIFERENÇA": 0.0 - pdf['saldo'], "ARQUIVO_ORIGEM": pdf['nome']
            })

    return pd.DataFrame(lista_final)

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
                              {'type': 'cell', 'criteria': '!=', 'value': 0, 'format': fmt_red})
    except: pass
    for i, col in enumerate(cols):
        max_len = max(df[col].astype(str).map(len).max() if not df[col].empty else 0, len(col))
        largura_final = min(max_len + 3, 60) 
        if col in ["RAZÃO", "EXTRATO", "DIFERENÇA"]:
            ws.set_column(i, i, width=largura_final, cell_format=fmt_moeda)
        else:
            ws.set_column(i, i, width=largura_final)

# ==============================================================================
# 3. INTERFACE STREAMLIT
# ==============================================================================

st.markdown("<h1 style='text-align: center;'>Conciliador Bancário V32 (Banco x GovBr)</h1>", unsafe_allow_html=True)
st.markdown("---")

c1, c2 = st.columns(2)
with c1: 
    st.markdown('<p class="big-label">Upload Compactado (.zip/.rar) com PDFs</p>', unsafe_allow_html=True)
    up_extratos = st.file_uploader("", type=["zip", "rar"], key="up_pdf", label_visibility="collapsed")
with c2: 
    st.markdown('<p class="big-label">Upload Planilha Comparação (.xlsx/.csv)</p>', unsafe_allow_html=True)
    up_planilha = st.file_uploader("", type=["xlsx", "csv"], key="up_xlsx", label_visibility="collapsed")

if st.button("PROCESSAR CONCILIAÇÃO", use_container_width=True):
    if up_extratos and up_planilha:
        with st.spinner("Descompactando arquivos e analisando PDFs..."):
            
            # --- 1. Preparação do Ambiente Temporário ---
            temp_dir = tempfile.mkdtemp()
            try:
                # --- 2. Salva e Extrai o Arquivo Compactado ---
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

                # --- 3. Processa a Planilha ---
                dados_excel = ler_planilha_e_consolidar(up_planilha)
                if not dados_excel:
                    st.error("Erro ao ler a planilha Excel. Verifique o formato.")
                    shutil.rmtree(temp_dir)
                    st.stop()

                # --- 4. Executa o Cruzamento (V32) ---
                df_final = processar_confronto(temp_dir, dados_excel)

                # --- 5. Gera Saída Excel ---
                if not df_final.empty:
                    
                    # Ordenação para o Excel
                    def ordenar(df):
                        df['is_nd'] = df['UG'] == 'N/D'
                        df_sorted = df.sort_values(by=['is_nd', 'UG', 'CÓDIGO'])
                        return df_sorted.drop(columns=['is_nd'])

                    cols_view = ["UG", "CÓDIGO", "DESCRIÇÃO", "CONTA", "RAZÃO", "EXTRATO", "DIFERENÇA", "ARQUIVO_ORIGEM"]
                    cols_validas = [c for c in cols_view if c in df_final.columns]

                    df_app = df_final[df_final['GRUPO'] == 'APLICACAO'][cols_validas]
                    df_app = ordenar(df_app)
                    df_mov = df_final[df_final['GRUPO'] == 'MOVIMENTO'][cols_validas]
                    df_mov = ordenar(df_mov)

                    # Gera o Excel em memória
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        wb = writer.book
                        ws = wb.add_worksheet('Conciliacao')
                        writer.sheets['Conciliacao'] = ws
                        
                        row = 0
                        if not df_app.empty:
                            ws.write(row, 0, "--- CONTAS APLICAÇÃO ---", wb.add_format({'bold':True, 'font_color':'blue'}))
                            row+=1
                            df_app.to_excel(writer, sheet_name='Conciliacao', startrow=row, index=False)
                            aplicar_estilo_excel(ws, wb, df_app, row, cols_validas)
                            row += len(df_app) + 2
                        
                        row += 1
                        if not df_mov.empty:
                            ws.write(row, 0, "--- CONTAS MOVIMENTO ---", wb.add_format({'bold':True, 'font_color':'blue'}))
                            row+=1
                            df_mov.to_excel(writer, sheet_name='Conciliacao', startrow=row, index=False)
                            aplicar_estilo_excel(ws, wb, df_mov, row, cols_validas)

                    output.seek(0)
                    
                    # --- 6. Exibe Preview Simplificado ---
                    st.success("Conciliação Finalizada com Sucesso!")
                    
                    # Mostra um resumo em HTML (Estilo Visual Solicitado)
                    html = "<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>"
                    html += "<h4 style='color:black; margin-top:0;'>Resumo do Processamento</h4>"
                    html += f"<p style='color:black;'>Total de Contas Analisadas: <b>{len(df_final)}</b></p>"
                    html += f"<p style='color:black;'>Total de Diferenças Encontradas: <b style='color:red;'>{len(df_final[df_final['DIFERENÇA'] != 0])}</b></p>"
                    html += "</div>"
                    st.markdown(html, unsafe_allow_html=True)
                    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

                    # --- 7. Botão de Download ---
                    st.download_button(
                        label="BAIXAR RELATÓRIO EXCEL V32",
                        data=output,
                        file_name="Relatorio_Conciliacao_V32_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("O processamento não gerou dados. Verifique os arquivos.")
            
            except Exception as e:
                st.error(f"Erro fatal: {e}")
            finally:
                # Limpa arquivos temporários
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
    else:
        st.warning("⚠️ Selecione o arquivo ZIP/RAR e a Planilha Excel primeiro.")
