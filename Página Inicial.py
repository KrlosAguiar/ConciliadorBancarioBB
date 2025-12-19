import streamlit as st
from PIL import Image
import os

# Configuração de Caminho e Ícone
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
icon_image = Image.open(icon_path)

st.set_page_config(
    page_title="DECON - Barcarena/PA",
    page_icon=icon_image,
    layout="wide"
)

# --- CABEÇALHO COM ÍCONE ---
# Criamos duas colunas: uma pequena para o logo e uma grande para o título
col_logo, col_titulo = st.columns([1, 8])

with col_logo:
    st.image(icon_image, width=80) # Ajuste a largura conforme necessário

with col_titulo:
    st.title("Departamento de Contabilidade - Barcarena/PA")

st.markdown("---")

# --- CORPO DA PÁGINA COM ESPAÇAMENTO DUPLO ---
st.markdown("""
### Bem-vindo ao sistema centralizado de ferramentas contábeis.

&nbsp;
&nbsp;

Utilize o menu lateral à esquerda para navegar entre os módulos disponíveis:

&nbsp;
&nbsp;

* **Conciliador Bancário:** Cruza os dados do Extrato Bancário com o Razão da Contabilidade.

&nbsp;

* **Tarifas Bancárias:** Extrai os lançamentos das tarifas do Extrato Bancário e emite um relatório pronto para empenho, liquidação e pagamento.

&nbsp;

* **(Em breve) Novos Módulos:** Outras ferramentas serão adicionadas aqui.

&nbsp;
&nbsp;

---
**Status do Sistema:** ✅ Online
""")
