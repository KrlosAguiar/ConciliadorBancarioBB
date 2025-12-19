import streamlit as st
from PIL import Image
import os

# Configuração de Caminho e Ícone para a Aba do Navegador
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
icon_image = Image.open(icon_path)

st.set_page_config(
    page_title="DECON - Barcarena/PA",
    page_icon=icon_image,
    layout="wide"
)

# --- CABEÇALHO ---
# Insere a imagem acima do título (o padrão do st.image já é alinhado à esquerda)
st.image(icon_image, width=120) 

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

* **Tarifas Bancárias:** Extrai os lançamentos das tarifas do Extrato Bancário e emite um relatório pronto para empenho, liquidação e pagamento.

* **(Em breve) Novos Módulos:** Outras ferramentas serão adicionadas aqui.

&nbsp;
&nbsp;

---
**Status do Sistema:** ✅ Online
""")
