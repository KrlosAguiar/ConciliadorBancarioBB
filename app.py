import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Conciliador Bancário - Banco do Brasil", layout="wide")

# --- SEGURANÇA ---
SENHA_MESTRA = "cliente123"

def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    
    if st.session_state["password_correct"]:
        return True

    st.title("Acesso Restrito")
    password = st.text_input("Digite a chave de acesso:", type="password")
    if st.button("Entrar"):
        if password == SENHA_MESTRA:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("Chave incorreta!")
    return False

if check_password():
    # --- TÍTULO ATUALIZADO ---
    st.title("🏦 Conciliador Bancário - Banco do Brasil")
    st.markdown("---")

    # --- UPLOAD DOS ARQUIVOS ---
    col1, col2 = st.columns(2)
    with col1:
        extrato_pdf = st.file_uploader("Upload do Extrato (PDF do BB)", type=["pdf"])
    with col2:
        extrato_excel = st.file_uploader("Upload do Relatório (Excel/CSV)", type=["xlsx", "csv"])

    if extrato_pdf and extrato_excel:
        try:
            # 1. PROCESSAR PDF (Extraindo Histórico Completo)
            lista_pdf = []
            with pdfplumber.open(extrato_pdf) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        for row in table:
                            # Filtra linhas que parecem ter datas (ex: 01/01/2023)
                            if row[0] and re.match(r'\d{2}/\d{2}/\d{4}', str(row[0])):
                                # Captura: Data, Histórico, Documento e Valor
                                data = row[0]
                                historico = str(row[1]).strip() if row[1] else ""
                                documento = str(row[2]).strip() if row[2] else ""
                                valor_str = str(row[3]).replace('.', '').replace(',', '.') if row[3] else "0"
                                try:
                                    valor = float(valor_str)
                                except:
                                    valor = 0.0
                                
                                lista_pdf.append({
                                    "Data": data,
                                    "Histórico": historico,
                                    "Documento": documento,
                                    "Valor_PDF": valor
                                })
            
            df_pdf = pd.DataFrame(lista_pdf)

            # 2. PROCESSAR EXCEL
            if extrato_excel.name.endswith('.csv'):
                df_excel = pd.read_csv(extrato_excel)
            else:
                df_excel = pd.read_excel(extrato_excel)
            
            # (Aqui você deve garantir que as colunas do seu Excel batam com o PDF para o merge)
            # Exemplo genérico de cruzamento por valor e data:
            df_final = pd.merge(df_pdf, df_excel, left_on="Valor_PDF", right_on=df_excel.columns[0], how="left")

            st.subheader("📊 Relatório de Conciliação")

            # --- FORMATAÇÃO DAS COLUNAS SOLICITADA ---
            # Aplicando estilos CSS para alinhamento
            def alinhar_tabela(df):
                return df.style.format({
                    "Valor_PDF": "R$ {:.2f}"
                }).set_properties(subset=["Data", "Documento"], **{'text-align': 'center'})\
                  .set_properties(subset=["Histórico"], **{'text-align': 'left'})\
                  .set_properties(subset=["Valor_PDF"], **{'text-align': 'right'})

            st.table(alinhar_tabela(df_pdf.head(20))) # Exibindo as primeiras 20 linhas formatadas

            # --- BOTÃO DE DOWNLOAD ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_pdf.to_excel(writer, index=False, sheet_name='Conciliado')
            
            st.download_button(
                label="📥 Baixar Relatório Completo (Excel)",
                data=output.getvalue(),
                file_name="conciliacao_bb.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro ao processar arquivos: {e}")
