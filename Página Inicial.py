import streamlit as st
from PIL import Image
import os

# Define o caminho para a imagem que está na raiz do projeto
# Isso garante que funcione tanto localmente quanto no servidor
icon_path = os.path.join(os.getcwd(), "Barcarena.png")
icon_image = Image.open(icon_path)

st.set_page_config(
    page_title="DECON - Barcarena/PA",
    page_icon=icon_image, # Aqui aplicamos o seu ícone customizado
    layout="wide"
)

st.title("Departamento de Contabilidade - Barcarena/PA")
st.markdown("---")

st.markdown("""
### Bem-vindo ao sistema centralizado de ferramentas contábeis.

<br>

Utilize o menu lateral à esquerda para navegar entre os módulos disponíveis:

<br>

- **Conciliador Bancário:** Cruza os dados do Extrato Bancário com o Razão da Contabilidade.

- **Tarifas Bancárias:** Extrai os lançamentos das tarifas do Extrato Bancário e emite um relatório pronto para empenho, liquidação e pagamento.

- **(Em breve) Novos Módulos:** Outras ferramentas serão adicionadas aqui.

<br>

---
**Status do Sistema:** ✅ Online
""")

# Dica: Se quiser que o login seja feito AQUI e valha para tudo,
# você pode mover a função check_password para cá no futuro.










