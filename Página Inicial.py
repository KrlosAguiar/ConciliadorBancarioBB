import streamlit as st
from PIL import Image
import os
import base64

# Configuração de Caminho e Ícone
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
icon_image = Image.open(icon_path)

st.set_page_config(
    page_title="DECON - Barcarena/PA",
    page_icon=icon_image,
    layout="wide"
)

# Função para converter imagem para base64 (necessário para HTML inline)
def get_base64_image(img_path):
    with open(img_path, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode()

img_base64 = get_base64_image(icon_path)

# --- CABEÇALHO CUSTOMIZADO ---
# Usamos HTML para garantir cantos quadrados, tamanho reduzido e distância exata
st.markdown(
    f"""
    <div style="display: flex; align-items: center;">
        <img src="data:image/png;base64,{img_base64}" 
             style="width: 40px; border-radius: 0px;">
        <span style="margin-left: 15px;">
            <h1 style="margin: 0;">Departamento de Contabilidade</h1>
        </span>
    </div>
    """, 
    unsafe_allow_html=True
)

st.markdown("---")

# --- CORPO DA PÁGINA ---
st.markdown("""
### Bem-vindo ao sistema centralizado de ferramentas contábeis.

&nbsp;
&nbsp;

Utilize o menu lateral à esquerda para navegar entre os módulos disponíveis:

&nbsp;
&nbsp;

* **Conciliador Bancário:** Cruza os dados do Extrato Bancário com o Razão da Contabilidade.

* **Conciliador de Saldos Bancários:** Cruza os saldos de vários extratos com os saldos do relatório da GovBr.

* **Tarifas Bancárias:** Extrai os lançamentos das tarifas do Extrato Bancário para empenho, liquidação e pagamento.

* **Projeção de FOPAG:** Calcula a projeção da Folha de Pagamento para os meses restantes.

* **(Em breve) Novos Módulos:** Outras ferramentas serão adicionadas aqui.

&nbsp;
&nbsp;

---
**Status do Sistema:** ✅ Online
""")



