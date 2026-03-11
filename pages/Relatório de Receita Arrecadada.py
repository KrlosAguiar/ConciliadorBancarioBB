import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
from PIL import Image

# --- CONFIGURAÇÃO DA PÁGINA ---
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
try:
    icon_image = Image.open(icon_path)
    st.set_page_config(page_title="Extrator Contábil", page_icon=icon_image, layout="wide")
except FileNotFoundError:
    st.set_page_config(page_title="Extrator Contábil", layout="wide")

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

def limpar_conta(conta_raw):
    """Limpa zeros à esquerda de contas no formato XX/000123/X"""
    partes = conta_raw.split('/')
    if len(partes) == 3:
        meio = partes[1].strip()
        meio_limpo = re.sub(r'^0+(\d)', r'\1', meio)
        return meio_limpo
    return conta_raw.strip()

def extrair_relatorio_inteligente(file_bytes):
    dados = []
    
    # Variáveis de contexto
    banco_atual = None
    conta_atual = None
    data_pag_atual = None
    data_cred_atual = None
    
    texto_completo = []
    
    # 1. Junta todo o texto do PDF usando io.BytesIO para o Streamlit
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo.append(texto)
                
    texto_total = "\n".join(texto_completo)
    
    # LIMPEZA ESTRUTURAL
    texto_total = texto_total.replace('"', '')
    texto_total = re.sub(r'(\d{1,3}(?:\.\d{3})*,\d{2})(\d{5,20})', r'\1\n\2', texto_total)
    
    linhas_originais = texto_total.split('\n')
    linhas_processadas = []
    
    # ===============================================================
    # PASSO 1: RECONSTRUÇÃO INTELIGENTE SEM FRONTEIRAS DE PÁGINA
    # ===============================================================
    for lin in linhas_originais:
        lin = lin.strip()
        lin = re.sub(r'^,\s*', '', lin) 
        
        if not lin:
            continue
            
        # Filtra Lixo de Cabeçalho/Rodapé para que não interrompa a costura das linhas
        lin_lower = lin.lower()
        if (lin_lower in ['eitura', 'munici', 'barcarena'] or
            lin_lower.startswith('página:') or
            lin_lower.startswith('hora:') or
            (lin_lower.startswith('data:') and not lin_lower.startswith('data pag') and not lin_lower.startswith('data cr')) or
            lin_lower.startswith('prefeitura municipal') or
            lin_lower.startswith('secretaria municipal') or
            lin_lower.startswith('e-mail:') or
            lin_lower.startswith('rodovia pa') or
            lin_lower.startswith('relatório contábil') or
            lin_lower.startswith('filtros:') or
            re.match(r'^\d{1,4}$', lin) or
            re.match(r'^\d{2}/\d{2}/\d{4}$', lin) or
            re.match(r'^\d{2}:\d{2}$', lin)):
            continue
            
        # Regra de Identificação de Início de Bloco
        is_novo_bloco = (
            lin.startswith('Banco:') or
            lin.startswith('Conta:') or
            lin.startswith('Data Pagamento:') or
            lin.startswith('Data Crédito:') or
            lin.lower().startswith('total') or
            bool(re.match(r'^\d{5,20}(?!,\d{2})', lin)) 
        )
        
        if is_novo_bloco:
            linhas_processadas.append(lin)
        else:
            if linhas_processadas:
                linhas_processadas[-1] += " " + lin
            else:
                linhas_processadas.append(lin)
                
    # ===============================================================
    # PASSO 2: EXTRAÇÃO DOS DADOS LIMPOS
    # ===============================================================
    for linha in linhas_processadas:
        linha = re.sub(r'\s+', ' ', linha).strip()
        
        if linha.startswith('Banco:'):
            if 'Conta:' in linha:
                banco_part, conta_part = linha.split('Conta:', 1)
                banco_atual = banco_part.replace('Banco:', '').strip()
                conta_atual = limpar_conta(conta_part.strip())
            else:
                banco_atual = linha.replace('Banco:', '').strip()
                
        elif linha.startswith('Conta:'):
            conta_atual = limpar_conta(linha.replace('Conta:', '').strip())
            
        elif linha.startswith('Data Pagamento:'):
            if 'Data Crédito:' in linha:
                dp_part, dc_part = linha.split('Data Crédito:', 1)
                m_pag = re.search(r'(\d{2}/\d{2}/\d{4})', dp_part)
                if m_pag: data_pag_atual = m_pag.group(1)
                m_cred = re.search(r'(\d{2}/\d{2}/\d{4})', dc_part)
                if m_cred: data_cred_atual = m_cred.group(1)
            else:
                m_pag = re.search(r'(\d{2}/\d{2}/\d{4})', linha)
                if m_pag: data_pag_atual = m_pag.group(1)
                
        elif linha.startswith('Data Crédito:'):
            m_cred = re.search(r'(\d{2}/\d{2}/\d{4})', linha)
            if m_cred: data_cred_atual = m_cred.group(1)
                
        elif linha.lower().startswith('total'):
            continue
            
        elif re.match(r'^\d{5,20}', linha):
            # MOTOR DE EXTRAÇÃO COM BUSCA REVERSA
            m_codigo = re.match(r'^(\d{5,20})', linha)
            if m_codigo:
                codigo = m_codigo.group(1)
                resto = linha[m_codigo.end():].strip()
                
                # Busca valores no formato monetário que sobraram na linha
                matches_valores = list(re.finditer(r'(?<!\d)(\d{1,3}(?:\.\d{3})*,\d{2})(?!\d)', resto))
                
                if matches_valores:
                    # Captura sempre o último valor encontrado na linha
                    m_val = matches_valores[-1]
                    valor_str = m_val.group(1)
                    
                    # A descrição é tudo que fica entre o Código e o Valor Monetário
                    desc = resto[:m_val.start()].strip('- ,')
                    
                    dados.append({
                        'Banco': banco_atual,
                        'Conta': conta_atual,
                        'Data Pagamento': data_pag_atual,
                        'Data Crédito': data_cred_atual,
                        'Código': codigo,
                        'Descrição': desc,
                        'Valor (R$)': valor_str
                    })
                else:
                    # =========================================================
                    # FALLBACK MAGNÉTICO: RESGATA O VALOR SOBREPOSTO AO TEXTO
                    # =========================================================
                    idx_virgula = resto.rfind(',')
                    if idx_virgula != -1:
                        centavos = ""
                        idx_fim_centavos = idx_virgula
                        
                        # Caça os 2 próximos números ignorando letras soltas
                        for i in range(idx_virgula + 1, len(resto)):
                            if resto[i].isdigit():
                                centavos += resto[i]
                                if len(centavos) == 2:
                                    idx_fim_centavos = i
                                    break
                        
                        if len(centavos) == 2:
                            # Isola o último bloco de palavras para não roubar números de outras partes da descrição
                            idx_espaco = resto.rfind(' ', 0, idx_virgula)
                            if idx_espaco == -1: idx_espaco = 0
                                
                            inteiros = ""
                            # Sugando a parte inteira
                            for i in range(idx_espaco, idx_virgula):
                                if resto[i].isdigit() or resto[i] == '.':
                                    inteiros += resto[i]
                            
                            inteiros = inteiros.lstrip('.')
                            if inteiros == "": inteiros = "0"
                            
                            valor_str = f"{inteiros},{centavos}"
                            
                            # Limpa a descrição removendo APENAS os dígitos e pontos do bloco final misturado
                            bloco_final = resto[idx_espaco:idx_fim_centavos+1]
                            bloco_limpo = re.sub(r'[\d,\.]', '', bloco_final)
                            
                            desc_raw = resto[:idx_espaco] + " " + bloco_limpo + resto[idx_fim_centavos+1:]
                            desc = re.sub(r'\s+', ' ', desc_raw).strip('- ,')
                            
                            dados.append({
                                'Banco': banco_atual,
                                'Conta': conta_atual,
                                'Data Pagamento': data_pag_atual,
                                'Data Crédito': data_cred_atual,
                                'Código': codigo,
                                'Descrição': desc,
                                'Valor (R$)': valor_str
                            })

    df = pd.DataFrame(dados)
    if not df.empty:
        df['Código'] = df['Código'].astype(str) # Garante que o excel não corte zeros do código
    return df

def gerar_excel_simples(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Lançamentos')
        workbook = writer.book
        worksheet = writer.sheets['Lançamentos']
        
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, fmt_header)
            
        worksheet.set_column('A:B', 25)
        worksheet.set_column('C:D', 15)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 40)
        worksheet.set_column('G:G', 15)
        
    return output.getvalue()

# ==============================================================================
# 2. INTERFACE DO APLICATIVO
# ==============================================================================
st.markdown("<h1 style='text-align: center;'>Conversor de Relatório de Receitas Arrecadadas</h1>", unsafe_allow_html=True)
st.markdown("---")

st.markdown('<p class="big-label" style="text-align: center;">Selecione o Relatório Contábil em PDF</p>', unsafe_allow_html=True)
c1, c2, c3 = st.columns([1, 2, 1])
with c2:
    up_pdf = st.file_uploader("", type="pdf", key="up_pdf", label_visibility="collapsed")

st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

if st.button("PROCESSAR A EXTRAÇÃO DE DADOS", use_container_width=True):
    if up_pdf:
        with st.spinner("Lendo documento e extraíndo dados..."):
            pdf_bytes = up_pdf.read()
            df_resultado = extrair_relatorio_inteligente(pdf_bytes)
            
            if not df_resultado.empty:
                st.success(f"✅ Extração concluída! Foram encontrados {len(df_resultado)} registros.")
                
                # Exibe a tabela na tela do Streamlit
                st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                
                # Prepara o botão de download
                excel_data = gerar_excel_simples(df_resultado)
                nome_base = os.path.splitext(up_pdf.name)[0]
                
                st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
                
                st.download_button(
                    label="BAIXAR RELATÓRIO CONVERTIDO EM EXCEL",
                    data=excel_data,
                    file_name=f"{nome_base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.error("❌ Nenhum dado financeiro capturado. Verifique se o arquivo PDF contém a estrutura contábil esperada.")
    else:
        st.warning("⚠️ Por favor, selecione um arquivo PDF antes de iniciar o processamento.")
