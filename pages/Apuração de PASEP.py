import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import base64
import xlsxwriter
from PIL import Image
from xhtml2pdf import pisa

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E CSS (DESIGN DE REFERÊNCIA)
# ==============================================================================
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
try:
    icon_image = Image.open(icon_path)
    st.set_page_config(page_title="Apuração do PASEP", page_icon=icon_image, layout="wide")
except:
    st.set_page_config(page_title="Apuração do PASEP", layout="wide")

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
    .big-label { font-size: 20px !important; font-weight: 600 !important; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# REGRAS DE NEGÓCIO E LISTAS DE CONTAS
# ==============================================================================
CODIGOS_PRINCIPAIS = [
    '1.1.0.0.00.0.0.00.00.00', '1.2.0.0.00.0.0.00.00.00', '1.3.0.0.00.0.0.00.00.00', 
    '1.6.0.0.00.0.0.00.00.00', '1.7.0.0.00.0.0.00.00.00', '1.9.0.0.00.0.0.00.00.00', 
    '2.0.0.0.00.0.0.00.00.00', '2.3.0.0.00.0.0.00.00.00', 'DEDUCAO_FUNDEB'
]

CODIGOS_EDUCACAO = [
    '1.3.2.1.01.0.1.01.00.00', '1.3.2.1.01.0.1.02.00.00', '1.3.2.1.01.0.1.04.00.00', 
    '1.7.1.4.00.0.0.00.00.00', '1.7.5.1.00.0.0.00.00.00', '1.7.1.5.00.0.0.00.00.00', 
    '1.7.2.4.51.0.0.00.00.00', '1.3.2.1.01.1.1.01.01.00', '1.3.2.1.01.1.1.02.01.00',
    '1.3.2.1.01.1.1.03.01.00', '1.3.2.1.01.1.1.03.02.00', '1.3.2.1.01.1.1.03.03.00', 
    '1.3.2.1.01.1.1.03.04.00', '1.3.2.1.01.2.1.00.00.00'
]

CODIGOS_SAUDE = [
    '1.3.2.1.01.0.1.05.00.00', '1.7.1.3.00.0.0.00.00.00', '1.7.2.3.00.0.0.00.00.00',
    '1.3.2.1.01.1.1.04.01.00', '1.3.2.1.01.1.1.04.02.00', '1.3.2.1.01.1.1.04.03.00', 
    '1.3.2.1.01.1.1.04.04.00',
    '1.1.2.1.50.0.1.00.00.00', '1.1.2.1.50.0.2.00.00.00', '1.1.2.1.50.0.3.00.00.00',
    '1.1.2.1.50.0.4.00.00.00'
]

CODIGOS_PMB = [
    '1.3.2.1.01.1.1.05.01.00', '1.3.2.1.01.1.1.05.02.00'
]

ALVO_TAXAS, ALVO_APLIC = '1.1.2.1.01.0.1.00.00.00', '1.3.2.1.01.0.1.09.20.00'

# ==============================================================================
# FUNÇÕES DE EXTRAÇÃO E PROCESSAMENTO
# ==============================================================================
def is_filho(pai, filho):
    pai_str = str(pai).strip()
    filho_str = str(filho).strip()
    if pai_str == filho_str: return False
    
    p = pai_str.split('.')
    c = filho_str.split('.')
    if len(p) != len(c): return False
    
    last_idx = -1
    for i in range(len(p)-1, -1, -1):
        if p[i] not in ['0', '00', '000']:
            last_idx = i
            break
            
    for i in range(last_idx + 1):
        if p[i] != c[i]:
            return False
    return True

def categorizar_coluna(codigo):
    codigo_limpo = str(codigo).strip()
    for c in CODIGOS_PMB:
        if codigo_limpo == c.strip() or is_filho(c.strip(), codigo_limpo): return 'PMB'
    for c in CODIGOS_SAUDE: 
        if codigo_limpo == c.strip() or is_filho(c.strip(), codigo_limpo): return 'SAÚDE'
    for c in CODIGOS_EDUCACAO: 
        if codigo_limpo == c.strip() or is_filho(c.strip(), codigo_limpo): return 'EDUCAÇÃO'
    return 'PMB'

def formatar_para_br(valor):
    if pd.isna(valor): return ""
    if isinstance(valor, (int, float)): return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return str(valor).strip()

def formatar_valor(valor, is_red=False):
    if valor == 0 or pd.isna(valor): return "-" 
    texto = f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    if is_red or valor < 0: return f"<span style='color: red;'>{texto}</span>"
    return texto

def extrair_pasep_pdf(pdf_bytes):
    dados_extraidos = []
    origem_atual = "DESCONHECIDO"
    data_atual = ""

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            
            lines = text.split('\n')
            for i, line in enumerate(lines):
                line_upper = line.strip().upper()
                if not line_upper: continue

                if "VALOR DISTRIBUIDO" in line_upper or ("DATA" in line_upper and "PARCELA" in line_upper):
                    ignorar = ["PAGINA", "PÁGINA", "DEMONSTRATIVO", "BARCARENA", "SISBB", "BENEFICIARIO", "BENEFICIÁRIO", "BANCO DO BRASIL", "SISTEMA", "CNPJ"]
                    for j in range(i-1, -1, -1):
                        prev_line = lines[j].strip().upper()
                        if prev_line and not any(ign in prev_line for ign in ignorar):
                            origem_atual = prev_line
                            if "FPM" in origem_atual and "MUNICIPIOS" in origem_atual: origem_atual = "FPM - FUNDO DE PARTICIPACAO DOS MUNICIPIOS"
                            elif "ITR" in origem_atual and "RURAL" in origem_atual: origem_atual = "ITR - IMPOSTO TERRITORIAL RURAL"
                            elif "SIMPLES" in origem_atual: origem_atual = "SIMPLES NACIONAL"
                            elif "FUNDEB" in origem_atual: origem_atual = "FUNDEB"
                            elif "CFM" in origem_atual: origem_atual = "CFM - COMPENSACAO FINANCEIRA"
                            break
                
                date_match = re.search(r'^(\d{2}\.\d{2}\.\d{4})', line_upper)
                if date_match:
                    data_atual = date_match.group(1)
                elif "TOTAL POR PARCELA" in line_upper:
                    data_atual = "TOTAL POR PARCELA / NATUREZA"
                    
                if "RETENCAO PASEP" in line_upper:
                    valor_encontrado = None
                    valores_linha = re.findall(r'(\d{1,3}(?:\.\d{3})*,\d{2})[CD]?', line_upper)
                    if valores_linha:
                        valor_encontrado = valores_linha[-1] 
                    
                    if not valor_encontrado and i + 1 < len(lines):
                        valores_prox = re.findall(r'(\d{1,3}(?:\.\d{3})*,\d{2})[CD]?', lines[i+1].upper())
                        if valores_prox:
                            valor_encontrado = valores_prox[-1]
                    
                    if valor_encontrado:
                        dados_extraidos.append({
                            'ORIGEM': origem_atual, 'DATA': data_atual,
                            'PARCELA': 'RETENCAO PASEP', 'VALOR': valor_encontrado
                        })

    df_pasep = pd.DataFrame(dados_extraidos)
    soma_total = 0.0

    if not df_pasep.empty:
        df_pasep = df_pasep.drop_duplicates().reset_index(drop=True)
        def limpar_valor(v):
            try: return float(str(v).replace('.', '').replace(',', '.'))
            except: return 0.0

        mask_totais = df_pasep['DATA'] == 'TOTAL POR PARCELA / NATUREZA'
        soma_total = df_pasep[mask_totais]['VALOR'].apply(limpar_valor).sum()
        
        if soma_total == 0:
            soma_total = df_pasep['VALOR'].apply(limpar_valor).sum()

        df_pasep.loc[len(df_pasep)] = {
            'ORIGEM': 'TOTAL GERAL', 'DATA': '',
            'PARCELA': 'SOMA DAS RETENÇÕES',
            'VALOR': f"{soma_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        }

    return df_pasep, soma_total

def parse_balancete_exato_v2(lines):
    data = []
    code_pattern = r'^(\d+(?:\.\d+){2,})'
    value_pattern = r'(\d{1,3}(?:\.\d{3})*,\d{2})'
    current_code, current_desc, in_deducao = None, [], False

    for line in lines:
        line_str = line.strip()
        if not line_str: continue

        upper_line = line_str.upper()
        if ("DEDUÇÃO" in upper_line or "DEDUCAO" in upper_line) and "FUNDEB" in upper_line: in_deducao = True
        elif "SUBTOTAL" in upper_line or "SUB TOTAL" in upper_line:
            if in_deducao: in_deducao = False 

        if "Unidade Gestora:" in line_str or ("Receita" in line_str and "Descrição" in line_str) or line_str.startswith("Balancete da Receita") or ("PAGE" in upper_line): continue

        matches = re.findall(value_pattern, line_str)
        code_match = re.search(code_pattern, line_str)

        if len(matches) >= 4:
            values = matches[-4:]
            first_val_idx = line_str.find(values[0])
            text_part = line_str[:first_val_idx].strip()

            if code_match:
                code = code_match.group(1)
                desc = text_part[len(code):].strip() if text_part.startswith(code) else text_part
                data.append({'Receita': code, 'Descrição': desc, 'Arrecadado Mês': values[1], 'Is_Deducao': in_deducao})
                current_code, current_desc = None, []
            else:
                if current_code:
                    full_desc = " ".join(current_desc + [text_part.strip()]).strip()
                    data.append({'Receita': current_code, 'Descrição': full_desc, 'Arrecadado Mês': values[1], 'Is_Deducao': in_deducao})
                    current_code, current_desc = None, []
        elif code_match:
            current_code = code_match.group(1)
            raw_desc = line_str[len(current_code):].strip() if line_str.startswith(current_code) else line_str.strip()
            current_desc = [raw_desc]
        else:
            if current_code: current_desc.append(line_str.strip())

    return pd.DataFrame(data)

# ==============================================================================
# INTERFACE STREAMLIT
# ==============================================================================
st.markdown("<h1 style='text-align: center;'>Relatório de Apuração do PASEP</h1>", unsafe_allow_html=True)
st.markdown("---")

st.markdown("### Informe os valores referentes à ARSEP")
col_arsep1, col_arsep2 = st.columns(2)
with col_arsep1:
    val_taxas_arsep = st.number_input("Valor Total das Taxas", min_value=0.0, step=0.01, format="%.2f")
with col_arsep2:
    val_aplic_arsep = st.number_input("Valor do Rendimento de Aplicação Financeira", min_value=0.0, step=0.01, format="%.2f")

st.markdown("<br>", unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1: 
    st.markdown('<p class="big-label">Demonstrativo DAF-PASEP (PDF)</p>', unsafe_allow_html=True)
    upload_pdf = st.file_uploader(" ", type="pdf", key="up_pdf", label_visibility="collapsed")
with c2: 
    st.markdown('<p class="big-label">Balancete da Receita (.xlsx)</p>', unsafe_allow_html=True)
    upload_xlsx = st.file_uploader(" ", type=["xlsx"], key="up_xlsx", label_visibility="collapsed")

if st.button("PROCESSAR RELATÓRIO", use_container_width=True):
    if upload_xlsx and upload_pdf:
        with st.spinner("Extraindo e processando dados..."):
            
            # 1. PROCESSAR PASEP (PDF)
            pdf_bytes = upload_pdf.read()
            df_pasep, total_pasep_retido = extrair_pasep_pdf(pdf_bytes)

            # 2. PROCESSAR BALANCETE (EXCEL)
            xlsx_bytes = upload_xlsx.read()
            df_raw = pd.read_excel(io.BytesIO(xlsx_bytes), header=None).dropna(how='all') 
            linhas_texto = [" ".join([formatar_para_br(val) for val in row if pd.notnull(val) and str(val).strip() != '']) for index, row in df_raw.iterrows()]
            df_resultado = parse_balancete_exato_v2(linhas_texto)
            
            df = df_resultado.copy()
            if df.empty:
                st.error("Não foi possível extrair dados do Balancete. Verifique o arquivo Excel.")
                st.stop()

            col_rec = next((c for c in df.columns if 'Receita' in c), 'Receita')
            col_desc = next((c for c in df.columns if 'Descrição' in c), 'Descrição')
            col_val = next((c for c in df.columns if 'Arrecadado Mês' in c), 'Arrecadado Mês')
            if 'Is_Deducao' not in df.columns: df['Is_Deducao'] = False
            
            df[col_val] = df[col_val].apply(lambda x: float(str(x).replace('R$', '').replace('.', '').replace(',', '.')) if isinstance(x, str) and ',' in str(x) else float(x) if isinstance(x, (int, float)) else 0.0)

            def get_folhas_agrupadas(codigo_pai):
                cands = df[df['Is_Deducao'] == True].copy() if codigo_pai == 'DEDUCAO_FUNDEB' else df[(df[col_rec].str.startswith(".".join(codigo_pai.split('.')[:2]))) & (df['Is_Deducao'] == False)].copy()
                lista_cands = cands[col_rec].tolist()
                folhas_raw = []
                for _, row in cands.iterrows():
                    cod = row[col_rec]
                    if codigo_pai != 'DEDUCAO_FUNDEB' and cod != codigo_pai and not is_filho(codigo_pai, cod): continue
                    if row[col_val] == 0: continue
                    if any(is_filho(cod, outro) for outro in lista_cands): continue
                    folhas_raw.append(row)

                agrupamento = {}
                for row in folhas_raw:
                    chave = row[col_rec]
                    if chave in agrupamento: agrupamento[chave]['valor'] += row[col_val]
                    else: agrupamento[chave] = {'codigo': chave, 'descricao': row[col_desc], 'valor': row[col_val], 'is_deducao': row.get('Is_Deducao', False)}
                return list(agrupamento.values())

            excel_data = []
            html_main = []
            grand_totals = {'PMB': 0.0, 'EDUCAÇÃO': 0.0, 'SAÚDE': 0.0, 'ARSEP': 0.0}

            style_table = "width:100%; border-collapse: collapse; font-family: Arial, sans-serif; color: black; font-size: 11px;"
            style_th = "text-align:left; padding: 6px; border-bottom: 2px solid black; background-color: #f2f2f2;"
            
            html_main.append(f'''<table style="{style_table}"><thead><tr>
                <th style="{style_th} width:16%; text-align: left;">Receita</th>
                <th style="{style_th} width:36%; text-align: left;">Descrição</th>
                <th style="{style_th} width:12%; text-align:right;">PMB</th>
                <th style="{style_th} width:12%; text-align:right;">EDUCAÇÃO</th>
                <th style="{style_th} width:12%; text-align:right;">SAÚDE</th>
                <th style="{style_th} width:12%; text-align:right;">ARSEP</th>
            </tr></thead><tbody>''')

            taxas_aplicadas, aplic_aplicadas = False, False

            for pai in CODIGOS_PRINCIPAIS:
                filhos = sorted(get_folhas_agrupadas(pai), key=lambda x: x['codigo'])
                parent_totals = {'PMB': 0.0, 'EDUCAÇÃO': 0.0, 'SAÚDE': 0.0, 'ARSEP': 0.0}
                filhos_render_data = []
                vermelho = (pai in ['DEDUCAO_FUNDEB', '2.3.0.0.00.0.0.00.00.00'])

                for f in filhos:
                    v_filho = f['valor']
                    if v_filho == 0: continue
                    leaf_vals = {'PMB': 0.0, 'EDUCAÇÃO': 0.0, 'SAÚDE': 0.0, 'ARSEP': 0.0}
                    if f['is_deducao']:
                        leaf_vals['EDUCAÇÃO'] = v_filho
                        parent_totals['EDUCAÇÃO'] += v_filho
                    else:
                        col_dest = categorizar_coluna(f['codigo'])
                        leaf_vals[col_dest] = v_filho
                        parent_totals[col_dest] += v_filho
                    
                    filhos_render_data.append({'codigo': f['codigo'], 'descricao': f['descricao'], 'vals': leaf_vals, 'is_adjust': False})

                    # ARSEP Injections
                    if f['codigo'] == ALVO_TAXAS and val_taxas_arsep > 0:
                        parent_totals['PMB'] -= val_taxas_arsep; parent_totals['ARSEP'] += val_taxas_arsep
                        filhos_render_data.append({'codigo': '↳ Dedução', 'descricao': 'Repasse ARSEP (Taxas)', 'vals': {'PMB': -val_taxas_arsep, 'EDUCAÇÃO': 0, 'SAÚDE': 0, 'ARSEP': val_taxas_arsep}, 'is_adjust': True})
                        taxas_aplicadas = True

                    if f['codigo'] == ALVO_APLIC and val_aplic_arsep > 0:
                        parent_totals['PMB'] -= val_aplic_arsep; parent_totals['ARSEP'] += val_aplic_arsep
                        filhos_render_data.append({'codigo': '↳ Dedução', 'descricao': 'Repasse ARSEP (Aplic. Fin.)', 'vals': {'PMB': -val_aplic_arsep, 'EDUCAÇÃO': 0, 'SAÚDE': 0, 'ARSEP': val_aplic_arsep}, 'is_adjust': True})
                        aplic_aplicadas = True

                if pai == '1.1.0.0.00.0.0.00.00.00' and val_taxas_arsep > 0 and not taxas_aplicadas:
                    parent_totals['PMB'] -= val_taxas_arsep; parent_totals['ARSEP'] += val_taxas_arsep
                    filhos_render_data.append({'codigo': '↳ Dedução', 'descricao': 'Repasse ARSEP (Taxas)', 'vals': {'PMB': -val_taxas_arsep, 'EDUCAÇÃO': 0, 'SAÚDE': 0, 'ARSEP': val_taxas_arsep}, 'is_adjust': True})
                if pai == '1.3.0.0.00.0.0.00.00.00' and val_aplic_arsep > 0 and not aplic_aplicadas:
                    parent_totals['PMB'] -= val_aplic_arsep; parent_totals['ARSEP'] += val_aplic_arsep
                    filhos_render_data.append({'codigo': '↳ Dedução', 'descricao': 'Repasse ARSEP (Aplic. Fin.)', 'vals': {'PMB': -val_aplic_arsep, 'EDUCAÇÃO': 0, 'SAÚDE': 0, 'ARSEP': val_aplic_arsep}, 'is_adjust': True})

                for col in grand_totals:
                    if pai in ['DEDUCAO_FUNDEB', '2.3.0.0.00.0.0.00.00.00']: 
                        grand_totals[col] -= parent_totals[col]
                    else: 
                        grand_totals[col] += parent_totals[col]

                desc_pai_excel = "(-) Dedução de Receita (FUNDEB)" if pai == 'DEDUCAO_FUNDEB' else (df[(df[col_rec] == pai) & (df['Is_Deducao'] == False)].iloc[0][col_desc] if not df[(df[col_rec] == pai) & (df['Is_Deducao'] == False)].empty else "Não encontrado")
                codigo_display_excel = "" if pai == 'DEDUCAO_FUNDEB' else pai
                
                desc_pai_html = desc_pai_excel if desc_pai_excel else "&nbsp;"
                codigo_display_html = codigo_display_excel if codigo_display_excel else "&nbsp;"

                excel_data.append({'Receita': codigo_display_excel, 'Descrição': desc_pai_excel, 'PMB': parent_totals['PMB'] * (-1 if vermelho else 1), 'EDUCAÇÃO': parent_totals['EDUCAÇÃO'] * (-1 if vermelho else 1), 'SAÚDE': parent_totals['SAÚDE'] * (-1 if vermelho else 1), 'ARSEP': parent_totals['ARSEP'] * (-1 if vermelho else 1)})
                html_main.append(f'<tr style="border-top: 1px solid black; background-color: #f9f9f9;"><td style="padding: 6px; font-weight: bold; font-size: 12px; text-align: left;">{codigo_display_html}</td><td style="padding: 6px; font-weight: bold; font-size: 12px; text-align: left;">{desc_pai_html}</td><td style="padding: 6px; font-weight: bold; font-size: 12px; text-align: right;">{formatar_valor(parent_totals["PMB"], is_red=vermelho)}</td><td style="padding: 6px; font-weight: bold; font-size: 12px; text-align: right;">{formatar_valor(parent_totals["EDUCAÇÃO"], is_red=vermelho)}</td><td style="padding: 6px; font-weight: bold; font-size: 12px; text-align: right;">{formatar_valor(parent_totals["SAÚDE"], is_red=vermelho)}</td><td style="padding: 6px; font-weight: bold; font-size: 12px; text-align: right;">{formatar_valor(parent_totals["ARSEP"], is_red=vermelho)}</td></tr>')

                for f in filhos_render_data:
                    excel_data.append({'Receita': f['codigo'], 'Descrição': f['descricao'], 'PMB': f['vals']['PMB'] * (-1 if vermelho else 1), 'EDUCAÇÃO': f['vals']['EDUCAÇÃO'] * (-1 if vermelho else 1), 'SAÚDE': f['vals']['SAÚDE'] * (-1 if vermelho else 1), 'ARSEP': f['vals']['ARSEP'] * (-1 if vermelho else 1)})
                    bg_c = "#fff8f8" if f.get('is_adjust') else "transparent"
                    c_html = f["codigo"] if f["codigo"] else "&nbsp;"
                    d_html = f["descricao"] if f["descricao"] else "&nbsp;"
                    html_main.append(f'<tr style="border-bottom: 1px solid #ccc; background-color: {bg_c};"><td style="padding: 4px 4px 4px 10px; text-align: left;">{c_html}</td><td style="padding: 4px; text-align: left;">{d_html}</td><td style="padding: 4px; text-align: right;">{formatar_valor(f["vals"]["PMB"], is_red=vermelho)}</td><td style="padding: 4px; text-align: right;">{formatar_valor(f["vals"]["EDUCAÇÃO"], is_red=vermelho)}</td><td style="padding: 4px; text-align: right;">{formatar_valor(f["vals"]["SAÚDE"], is_red=vermelho)}</td><td style="padding: 4px; text-align: right;">{formatar_valor(f["vals"]["ARSEP"], is_red=vermelho)}</td></tr>')

            val_1_pmb, val_1_edu, val_1_sau, val_1_ars = grand_totals['PMB']*0.01, grand_totals['EDUCAÇÃO']*0.01, grand_totals['SAÚDE']*0.01, grand_totals['ARSEP']*0.01
            pagar_pmb, pagar_edu, pagar_sau, pagar_ars = val_1_pmb - total_pasep_retido, val_1_edu, val_1_sau, val_1_ars

            totais_linhas = [
                ("TOTAL GERAL", grand_totals['PMB'], grand_totals['EDUCAÇÃO'], grand_totals['SAÚDE'], grand_totals['ARSEP']),
                ("PASEP (1%)", val_1_pmb, val_1_edu, val_1_sau, val_1_ars),
                ("PASEP RETIDO", total_pasep_retido, None, None, None),
                ("PASEP A PAGAR", pagar_pmb, pagar_edu, pagar_sau, pagar_ars)
            ]

            for nome, v_pmb, v_edu, v_sau, v_ars in totais_linhas:
                excel_data.append({'Receita': '', 'Descrição': nome, 'PMB': v_pmb, 'EDUCAÇÃO': v_edu if v_edu is not None else '-', 'SAÚDE': v_sau if v_sau is not None else '-', 'ARSEP': v_ars if v_ars is not None else '-'})
                text_color = "color: red;" if nome == "PASEP RETIDO" else "color: black;"
                tr_st = 'border-top: 2px solid black; background-color: #e0e0e0; font-weight: bold; font-size: 14px;'
                
                td_pmb = formatar_valor(v_pmb) if v_pmb is not None else "-"
                td_edu = formatar_valor(v_edu) if v_edu is not None else "-"
                td_sau = formatar_valor(v_sau) if v_sau is not None else "-"
                td_ars = formatar_valor(v_ars) if v_ars is not None else "-"
                
                html_main.append(f'<tr style="{tr_st}">\
                    <td colspan="2" style="padding: 4px 8px; text-align: left; {text_color}">{nome}</td>\
                    <td style="padding: 4px 8px; text-align: right; {text_color}">{td_pmb}</td>\
                    <td style="padding: 4px 8px; text-align: right; {text_color}">{td_edu}</td>\
                    <td style="padding: 4px 8px; text-align: right; {text_color}">{td_sau}</td>\
                    <td style="padding: 4px 8px; text-align: right; {text_color}">{td_ars}</td>\
                </tr>')

            html_main.append("</tbody></table>")
            html_tela = "".join(html_main)

            # --- GERAÇÃO EXCEL ---
            df_export = pd.DataFrame(excel_data)
            excel_io = io.BytesIO()
            with pd.ExcelWriter(excel_io, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Relatório Apuração do PASEP')
                workbook = writer.book
                worksheet = writer.sheets['Relatório Apuração do PASEP']
                
                formato_normal_moeda = workbook.add_format({'num_format': '#,##0.00'})
                formato_pai = workbook.add_format({'bold': True, 'size': 12})
                formato_pai_moeda = workbook.add_format({'bold': True, 'size': 12, 'num_format': '#,##0.00'})
                
                worksheet.set_column('A:A', 23)
                worksheet.set_column('B:B', 50)
                worksheet.set_column('C:F', 18, formato_normal_moeda)

                for row_num in range(len(df_export)):
                    row_data = df_export.iloc[row_num]
                    rec = str(row_data['Receita']).strip()
                    desc = str(row_data['Descrição']).strip()
                    is_pai = (rec in CODIGOS_PRINCIPAIS or desc in ["TOTAL GERAL", "PASEP (1%)", "PASEP RETIDO", "PASEP A PAGAR", "(-) Dedução de Receita (FUNDEB)"])
                    
                    if is_pai:
                        worksheet.set_row(row_num + 1, 18) 
                        worksheet.write(row_num + 1, 0, rec, formato_pai)
                        worksheet.write(row_num + 1, 1, desc, formato_pai)
                        for col_idx, col_name in enumerate(['PMB', 'EDUCAÇÃO', 'SAÚDE', 'ARSEP'], start=2):
                            val = row_data[col_name]
                            if val == '-' or pd.isna(val): worksheet.write(row_num + 1, col_idx, '-', formato_pai)
                            else: worksheet.write(row_num + 1, col_idx, float(val), formato_pai_moeda)

            excel_bytes_final = excel_io.getvalue()

            # --- GERAÇÃO PDF ---
            html_pasep = f'<div style="page-break-before: always;"></div><h2>Demonstrativo de Retenções PASEP</h2><table style="{style_table}"><thead><tr><th style="{style_th} width:40%;">ORIGEM</th><th style="{style_th} width:20%;">DATA</th><th style="{style_th} width:25%;">PARCELA</th><th style="{style_th} width:15%; text-align:right;">VALOR</th></tr></thead><tbody>'
            if not df_pasep.empty:
                for _, row in df_pasep.iterrows():
                    is_total = row['ORIGEM'] == 'TOTAL GERAL' or row['DATA'] == 'TOTAL POR PARCELA / NATUREZA'
                    bg = "#d3d3d3" if row['ORIGEM'] == 'TOTAL GERAL' else "transparent"
                    fw = "bold" if is_total else "normal"
                    fs = "12px" if is_total else "11px"
                    origem_v = row["ORIGEM"] if row["ORIGEM"] else "&nbsp;"
                    data_v = row["DATA"] if row["DATA"] else "&nbsp;"
                    parcela_v = row["PARCELA"] if row["PARCELA"] else "&nbsp;"
                    valor_v = row["VALOR"] if row["VALOR"] else "&nbsp;"
                    html_pasep += f'<tr style="background-color:{bg}; font-weight:{fw}; font-size:{fs};"><td style="padding:6px; border-bottom:1px solid #ccc;">{origem_v}</td><td style="padding:6px; border-bottom:1px solid #ccc;">{data_v}</td><td style="padding:6px; border-bottom:1px solid #ccc;">{parcela_v}</td><td style="padding:6px; border-bottom:1px solid #ccc; text-align:right;">{valor_v}</td></tr>'
            html_pasep += "</tbody></table>"

            full_html_pdf = f"""
            <html>
                <head>
                    <meta charset="utf-8">
                    <style>
                        @page {{ size: A4 landscape; margin: 1.0cm; }}
                        h2 {{ font-family: Arial; color: #333; margin-top: 15px; padding-bottom: 5px; border-bottom: 2px solid #333; font-size: 16px; }}
                        td {{ word-wrap: break-word; }}
                    </style>
                </head>
                <body>
                    <h2>Relatório de Apuração do PASEP</h2>
                    {html_tela}
                    {html_pasep}
                </body>
            </html>
            """
            pdf_io = io.BytesIO()
            pisa.CreatePDF(io.StringIO(full_html_pdf), dest=pdf_io)
            pdf_bytes_final = pdf_io.getvalue()

            # --- RENDERIZAÇÃO NA TELA ---
            st.success(f"✅ Extração de PASEP concluída (Retenção Total: R$ {total_pasep_retido:,.2f}).")
            
            st.markdown(f"<div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd; overflow-x: auto;'>{html_tela}</div>", unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                st.download_button(
                    label="GERAR RELATÓRIO EM EXCEL",
                    data=excel_bytes_final,
                    file_name="Relatorio_Apuracao_do_PASEP.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_btn2:
                st.download_button(
                    label="GERAR RELATÓRIO EM PDF",
                    data=pdf_bytes_final,
                    file_name="Relatorio_Apuracao_do_PASEP.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
    else:
        st.warning("⚠️ Selecione os dois arquivos primeiro (PDF do DAF PASEP e Excel do Balancete).")
